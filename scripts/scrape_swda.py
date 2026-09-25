#!/usr/bin/env python3
"""
SkillsFuture for Business (GoBusiness) Course Directory Scraper
=================================================================
Scrapes course listings + "About this course" + "What you'll learn"
from https://skillsfuture.gobusiness.gov.sg/course-directory/search
and saves results to an Excel file.

This replaces the old MySkillsFuture (myskillsfuture.gov.sg) scraper,
which targeted a different site that has since moved. The site's
markup uses generated/hashed CSS class names that change between
deploys, so this version deliberately avoids relying on them:
  - Cards are found via the one stable anchor: links matching
    /course-directory/courses/{course-code}
  - All card fields (title, provider, rating, fees, etc.) are
    extracted by regex over the card's plain innerText, not by
    CSS selectors.
  - Pagination controls are located by nearby "X of Y" page-count
    text plus several fallback selector patterns for the "next" button.

If GoBusiness changes their page layout enough that innerText field
order changes, only the regex patterns in `parse_card_text()` need
updating -- the card-discovery logic itself should keep working.

Usage:
    python skillsfuture_scraper.py --keyword "data analytics" --pages 3
    python skillsfuture_scraper.py --keyword "HR management" --pages 0   # 0 = all pages
    python skillsfuture_scraper.py --keyword "python" --pages 5 --no-details

Requirements:
    pip install playwright openpyxl
    playwright install chromium
"""

import asyncio
import argparse
import re
import sys
import urllib.parse
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Windows fix ───────────────────────────────────────────────
if sys.platform == "win32":
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

# ── Config ────────────────────────────────────────────────────
SITE_BASE = "https://skillsfuture.gobusiness.gov.sg"


def build_search_url(keyword: str, page: int, page_size: int = 10) -> str:
    """
    Build a search URL using the site's actual query-param pagination:
        ?search_query=<kw>&page_size=10&page=<N>&sort_by=relevance
    Note: page is 0-indexed on this site (page=0 is the first page).
    """
    return (
        f"{SITE_BASE}/course-directory/search"
        f"?search_query={urllib.parse.quote(keyword)}"
        f"&page_size={page_size}"
        f"&page={page}"
        f"&sort_by=relevance"
    )

COURSE_LINK_RE = re.compile(r"^/course-directory/courses/[A-Za-z0-9\-]+$")

CHROMIUM_ARGS = [
    "--no-sandbox",
    "--disable-setuid-sandbox",
    "--disable-dev-shm-usage",
    "--disable-gpu",
    "--no-first-run",
]

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)

# ── Excel styling ─────────────────────────────────────────────
_thin       = Side(style="thin", color="C0C0C0")
HEADER_FILL = PatternFill("solid", fgColor="1A4E8C")
ALT_FILL    = PatternFill("solid", fgColor="EBF0FA")
HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
BODY_FONT   = Font(name="Calibri", size=10)
TITLE_FONT  = Font(name="Calibri", bold=True, size=13, color="1A4E8C")
BORDER      = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
WRAP        = Alignment(wrap_text=True, vertical="top")
CENTER      = Alignment(horizontal="center", vertical="top")

BASE_COLUMNS = [
    ("Course Title",               52),
    ("Training Provider",          32),
    ("Course URL",                 42),
    ("Full Course Fee",            18),
    ("After Subsidy Fee",          20),
    ("After SFEC Fee",             18),
    ("Star Rating",                12),
    ("No. of Ratings",             14),
    ("Upcoming Course Date",       20),
]

DETAIL_COLUMNS = [
    ("About this course",          52),
    ("What you'll learn",          52),
]

TRAILING_COLUMNS = [
    ("Scraped At",                 26),
]


def get_columns(include_details: bool) -> list[tuple[str, int]]:
    cols = list(BASE_COLUMNS)
    if include_details:
        cols += DETAIL_COLUMNS
    cols += TRAILING_COLUMNS
    return cols

# ─────────────────────────────────────────────────────────────
# PARSING
# ─────────────────────────────────────────────────────────────

def parse_card_text(text: str, provider_hint: str, href: str) -> dict | None:
    """
    Parse the plain innerText of a course card into structured fields.
    Example raw text block:

        Industry Supported
        NTUC LEARNINGHUB PTE. LTD.
        CSA Certificate of Cloud Security Knowledge (CCSK) Plus (SF)
        4.5 (30 ratings)
        SkillsFuture Enterprise Credit (SFEC) - 24 hours | 2 - 3 Days - Full Time
        Full course fee
        S$2,150.00
        After subsidy
        From S$645.00
        After SFEC
        From S$64.50
        All fees are exclusive of GST. ...
    """
    norm = re.sub(r"[ \t]+", " ", text).strip()
    lines = [l.strip() for l in norm.split("\n") if l.strip()]
    if not lines:
        return None

    # Provider: prefer the hint scraped separately (more reliable),
    # else fall back to first non-badge line.
    provider = provider_hint or ""
    if not provider:
        for l in lines:
            if l.lower() != "industry supported" and not l.lower().startswith("s$"):
                provider = l
                break

    # Rating: "4.5 (30 ratings)" or "No rating"
    star_val = None
    num_ratings = 0
    m = re.search(r"(\d\.\d)\s*\((\d+)\s*ratings?\)", norm, re.I)
    if m:
        star_val = float(m.group(1))
        num_ratings = int(m.group(2))

    # Upcoming course date. Actual observed format:
    #   "Upcoming course date : 9 July 2026"
    # (colon may or may not have a leading space; label case-insensitive)
    upcoming = "None listed"
    m = re.search(
        r"Upcoming\s+course\s+date\s*:\s*([0-9]{1,2}\s+\w+\s+[0-9]{4})",
        norm, re.I
    )
    if m:
        upcoming = m.group(1).strip()

    # Fees
    def fee_after(label_pattern):
        m = re.search(label_pattern + r"\s*(From\s*)?(S\$[\d,.]+)", norm, re.I)
        return m.group(2) if m else "N/A"

    full_fee     = fee_after(r"Full course fee")
    subsidy_fee  = fee_after(r"After subsidy")
    sfec_fee     = fee_after(r"After SFEC")

    # Title: the anchor's own text is the most reliable source and is
    # passed in separately by the caller when available; as a parsing
    # fallback here, take the line right before the rating line (or,
    # if no rating found, the last non-fee, non-provider line before
    # the "SkillsFuture Enterprise Credit" line).
    title = ""
    rating_idx = None
    for i, l in enumerate(lines):
        if re.match(r"^\d\.\d\s*\(\d+\s*ratings?\)$", l, re.I) or l.lower() == "no rating":
            rating_idx = i
            break
    if rating_idx is not None and rating_idx >= 1:
        candidates = [l for l in lines[:rating_idx] if l.lower() != "industry supported"]
        if provider and candidates and candidates[0] == provider:
            candidates = candidates[1:]
        if candidates:
            title = candidates[-1]

    if not title:
        # last-resort fallback
        candidates = [l for l in lines if l.lower() != "industry supported" and l != provider]
        title = candidates[0] if candidates else ""

    return {
        "Course Title":         title,
        "Training Provider":    provider,
        "Course URL":           href,
        "Full Course Fee":      full_fee,
        "After Subsidy Fee":    subsidy_fee,
        "After SFEC Fee":       sfec_fee,
        "Star Rating":          star_val,
        "No. of Ratings":       num_ratings,
        "Upcoming Course Date": upcoming,
        "Scraped At":           datetime.now(timezone.utc).isoformat(),
    }


# ─────────────────────────────────────────────────────────────
# SCRAPING
# ─────────────────────────────────────────────────────────────

async def scrape_search_page(page) -> list[dict]:
    """Extract course cards from the current search results page."""
    await page.wait_for_selector('a[href*="/course-directory/courses/"]', timeout=60_000)
    await asyncio.sleep(1.5)

    raw_cards = await page.evaluate("""
        () => {
            const seen = new Set();
            const results = [];
            const anchors = Array.from(
                document.querySelectorAll('a[href*="/course-directory/courses/"]')
            );
            for (const a of anchors) {
                const href = a.getAttribute('href') || '';
                if (!href) continue;

                // Walk up to find the enclosing card: the smallest
                // ancestor whose innerText also contains a fee line,
                // which distinguishes the full card from the bare link.
                let node = a;
                let card = a;
                for (let i = 0; i < 8 && node; i++) {
                    node = node.parentElement;
                    if (!node) break;
                    const t = node.innerText || '';
                    if (/full course fee/i.test(t)) {
                        card = node;
                        break;
                    }
                }

                const key = href;
                if (seen.has(key)) continue;
                seen.add(key);

                // Try to grab provider name and title separately if
                // structured spans/elements are present (best effort;
                // falls back to full-text regex parsing on Python side).
                let providerHint = '';
                const providerEl = card.querySelector('[class*="provider" i]');
                if (providerEl) providerHint = providerEl.innerText.trim();

                results.push({
                    href,
                    text: card.innerText || '',
                    providerHint,
                    anchorText: a.innerText || ''
                });
            }
            return results;
        }
    """)

    rows = []
    for c in raw_cards:
        href = c["href"]
        if href.startswith("/"):
            full_url = SITE_BASE + href
        else:
            full_url = href

        row = parse_card_text(c["text"], c.get("providerHint", ""), full_url)
        if not row:
            continue
        rows.append(row)
    return rows


async def scrape_course_details(page, course_url: str) -> tuple[str, str]:
    """
    Visit a course detail page and extract:
      - About this course
      - What you'll learn
    Returns (about, learn).
    """
    if not course_url or not course_url.startswith("http"):
        return "", ""

    try:
        await page.goto(course_url, wait_until="domcontentloaded", timeout=60_000)
        await page.wait_for_load_state("networkidle")
        await asyncio.sleep(2.0)

        # Expand "Show more" if present
        try:
            await page.locator("text=Show more").first.click(timeout=3000)
            await asyncio.sleep(0.6)
        except Exception:
            pass

        # Scroll to trigger lazy-load
        await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        await asyncio.sleep(0.8)
        await page.evaluate("window.scrollTo(0, 0)")
        await asyncio.sleep(0.8)

        details = await page.evaluate("""
            () => {
                const root = document.querySelector("main") || document.body;
                const norm = (s) => (s || "").replace(/\\r\\n/g, "\\n").replace(/[ \\t]+/g, " ").trim();
                const clean = (s) => norm(s).replace(/\\n\\s*\\n/g, "\\n\\n");

                const RE_AFTER_ABOUT = /\\s*(?:what\\s+you(?:'|\\u2019)?ll\\s+learn|minimum\\s+entry|skills\\s+you(?:'|\\u2019)?ll\\s+pick|need\\s+more\\s+course\\s+info)\\b/i;
                const RE_AFTER_LEARN = /\\s*(?:minimum\\s+entry\\s+requirement|skills\\s+you(?:'|\\u2019)?ll\\s+pick\\s+up|need\\s+more\\s+course\\s+info|fee\\s+details|you\\s+may\\s+also\\s+like|all\\s+fees\\s+are)\\b/i;

                const findHeadingFlexible = (pattern) => {
                    const headings = Array.from(root.querySelectorAll("h1,h2,h3,h4,h5,h6"));
                    let el = headings.find(h => pattern.test(norm(h.innerText || "").toLowerCase()));
                    if (el) return el;
                    const wide = Array.from(root.querySelectorAll(
                        '[role="heading"],button,.accordion-button,.accordion-header,.nav-link,a,strong,p,span'
                    ));
                    return wide.find(node => {
                        const t = norm(node.innerText || "");
                        return t.length <= 200 && pattern.test(t.toLowerCase());
                    });
                };

                const collectAfterHeading = (heading, boundaryRegex) => {
                    if (!heading) return "";
                    const parts = [];
                    let node = heading.nextElementSibling;
                    while (node) {
                        const tag = (node.tagName || "").toLowerCase();
                        if (/^h[1-6]$/.test(tag)) {
                            const ht = norm(node.innerText || "").toLowerCase();
                            if (/what\\s+you(?:'|\\u2019)?ll\\s+learn/.test(ht) || /minimum\\s+entry/.test(ht) ||
                                /skills\\s+you(?:'|\\u2019)?ll/.test(ht) || /fee\\s+details/.test(ht)) break;
                        }
                        let nodeText = norm(node.innerText || "");
                        const cut = nodeText.search(boundaryRegex);
                        if (cut >= 0) {
                            nodeText = nodeText.slice(0, cut).trim();
                            if (nodeText) parts.push(nodeText);
                            break;
                        }
                        if (nodeText) parts.push(nodeText);
                        node = node.nextElementSibling;
                    }
                    return parts.join("\\n\\n");
                };

                const trimAbout = (raw) => {
                    let t = norm(raw);
                    let cut = t.search(RE_AFTER_ABOUT);
                    if (cut >= 0) t = t.slice(0, cut);
                    return clean(t);
                };
                const trimLearn = (raw) => {
                    let t = norm(raw);
                    let cut = t.search(RE_AFTER_LEARN);
                    if (cut >= 0) t = t.slice(0, cut);
                    return clean(t);
                };

                let about = "";
                let learn = "";

                const hAbout = findHeadingFlexible(/about\\s+this\\s+course|about\\s+course|course\\s+overview/i);
                if (hAbout) about = trimAbout(collectAfterHeading(hAbout, RE_AFTER_ABOUT));

                const hLearn = findHeadingFlexible(
                    /what\\s+you(?:'|\\u2019)?ll\\s+learn|what\\s+you\\s+will\\s+learn|learning\\s+outcomes|course\\s+outcomes/i
                );
                if (hLearn) learn = trimLearn(collectAfterHeading(hLearn, RE_AFTER_LEARN));

                const mainText = norm(root.innerText || "");
                if (!about) {
                    const m = mainText.match(
                        /about\\s+this\\s+course\\s*:?\\s*([\\s\\S]+?)(?=\\s*(?:what\\s+you(?:'|\\u2019)?ll\\s+learn|what\\s+you\\s+will\\s+learn)\\b)/i
                    );
                    if (m) about = trimAbout(m[1]);
                }
                if (!learn) {
                    const m = mainText.match(
                        /what\\s+you(?:'|\\u2019)?ll\\s+learn\\s*:?\\s*([\\s\\S]+?)(?=\\s*(?:minimum\\s+entry|you\\s+may\\s+also\\s+like|fee\\s+details)\\b)/i
                    );
                    if (m) learn = trimLearn(m[1]);
                }

                about = about.replace(/^\\s*(?:about\\s+this\\s+course|about\\s+course|course\\s+overview)\\s*:?\\s*/i, "").trim().slice(0, 8000);
                learn = learn.replace(/^\\s*(?:what\\s+you\\s+will\\s+learn\\s+from\\s+this\\s+course|what\\s+you(?:'|\\u2019)?ll\\s+learn|learning\\s+outcomes|course\\s+outcomes)\\s*:?\\s*/i, "").trim().slice(0, 8000);

                return { about, learn };
            }
        """)

        return details.get("about", ""), details.get("learn", "")

    except Exception as e:
        print(f"    ⚠ Detail scrape failed: {str(e)[:80]}")
        return "", ""



async def run_scraper(keyword: str, max_pages: int, scrape_details: bool) -> list[dict]:
    from playwright.async_api import async_playwright

    all_courses: list[dict] = []
    seen_keys: set[str] = set()
    scrape_all = max_pages <= 0
    limit = 999999 if scrape_all else max_pages
    PAGE_SIZE = 10

    print(f"\n🔍 Searching: '{keyword or '(all courses)'}'")
    print(f"   Pages: {'all' if scrape_all else max_pages}  |  Detail scrape: {scrape_details}\n")

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True, args=CHROMIUM_ARGS)
        context = await browser.new_context(
            user_agent=USER_AGENT,
            viewport={"width": 1280, "height": 900},
        )
        page = await context.new_page()

        # ── Phase 1: Collect all search-result cards, page by page ──
        page_index = 0  # site's pagination is 0-indexed
        page_num_display = 0
        for page_num_display in range(1, limit + 1):
            url = build_search_url(keyword, page_index, PAGE_SIZE)
            label = f"all ({page_num_display})" if scrape_all else f"{page_num_display}/{max_pages}"
            print(f"   📄 Scraping page {label}…", end=" ", flush=True)

            try:
                await page.goto(url, wait_until="domcontentloaded", timeout=120_000)
                await asyncio.sleep(3)
                page_courses = await scrape_search_page(page)
            except Exception as e:
                print(f"FAILED ({str(e)[:60]})")
                break

            new = [c for c in page_courses if c["Course URL"] not in seen_keys]
            seen_keys.update(c["Course URL"] for c in new)
            all_courses.extend(new)
            print(f"  {len(new)} new  (total {len(all_courses)})")

            # Stop once a page returns no cards at all (past the last page)
            if not page_courses:
                break
            # Stop once a page returns fewer than a full page's worth
            # (last page of results) -- but still counts what it found
            if len(page_courses) < PAGE_SIZE:
                break

            page_index += 1

        print(f"\n   ✅ {len(all_courses)} courses collected from {page_num_display} page(s).\n")

        # ── Phase 2: Scrape detail pages (About + What you'll learn) ──
        if scrape_details and all_courses:
            print(f"   🔎 Fetching detail pages for {len(all_courses)} courses…\n")
            for i, course in enumerate(all_courses):
                title = course["Course Title"]
                url_c = course["Course URL"]
                print(f"   [{i+1}/{len(all_courses)}] {title[:55]}…")
                about, learn = await scrape_course_details(page, url_c)
                course["About this course"] = about
                course["What you'll learn"] = learn
                if about or learn:
                    print(f"        About: {len(about)} chars  |  Learn: {len(learn)} chars")
                else:
                    print("        (no detail text found)")

        await browser.close()

    return all_courses


# ─────────────────────────────────────────────────────────────
# EXCEL OUTPUT
# ─────────────────────────────────────────────────────────────

def build_workbook(courses: list[dict], keyword: str, include_details: bool) -> Workbook:
    columns = get_columns(include_details)

    wb = Workbook()
    ws = wb.active
    ws.title = "SkillsFuture Courses"

    ws.merge_cells(f"A1:{get_column_letter(len(columns))}1")
    ws["A1"] = (
        f'SkillsFuture for Business: "{keyword or "all courses"}"  |  '
        f"Exported {datetime.now().strftime('%d %b %Y %H:%M')}"
    )
    ws["A1"].font = TITLE_FONT
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 28

    for col_idx, (header, width) in enumerate(columns, 1):
        cell = ws.cell(row=2, column=col_idx, value=header)
        cell.font      = HEADER_FONT
        cell.fill      = HEADER_FILL
        cell.alignment = CENTER
        cell.border    = BORDER
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[2].height = 22

    for row_idx, course in enumerate(courses, 3):
        fill = ALT_FILL if row_idx % 2 == 0 else None
        for col_idx, (key, _) in enumerate(columns, 1):
            val  = course.get(key, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font      = BODY_FONT
            cell.alignment = WRAP
            cell.border    = BORDER
            if fill:
                cell.fill = fill
        ws.row_dimensions[row_idx].height = 60 if include_details else 22

    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(len(columns))}2"
    return wb


def save_excel(courses: list[dict], keyword: str, include_details: bool, out_dir: str = ".") -> Path:
    safe_name = re.sub(r'[\\/*?:"<>|]', "_", keyword or "all").replace(" ", "_").lower()
    path = Path(out_dir) / f"skillsfuture_{safe_name}.xlsx"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = build_workbook(courses, keyword, include_details)
    wb.save(path)
    return path


# ─────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Scrape SkillsFuture for Business (GoBusiness) course directory to Excel."
    )
    parser.add_argument(
        "--keyword", "-k",
        default="",
        help='Search keyword, e.g. "data analytics" (leave blank for all courses)'
    )
    parser.add_argument(
        "--pages", "-p",
        type=int,
        default=3,
        help="Number of search result pages to scrape (0 = all pages, default: 3)"
    )
    parser.add_argument(
        "--no-details",
        action="store_true",
        help="Skip scraping individual course pages (faster but no About/Learn text)"
    )
    parser.add_argument(
        "--out", "-o",
        default=".",
        help="Output directory for the Excel file (default: current directory)"
    )
    args = parser.parse_args()

    courses = asyncio.run(
        run_scraper(
            keyword=args.keyword,
            max_pages=args.pages,
            scrape_details=not args.no_details,
        )
    )

    if not courses:
        print("\n⚠  No courses found. Check your keyword or try again later.")
        sys.exit(1)

    out_path = save_excel(courses, args.keyword, not args.no_details, args.out)
    print(f"\n✅ Saved {len(courses)} courses → {out_path}\n")


if __name__ == "__main__":
    main()
