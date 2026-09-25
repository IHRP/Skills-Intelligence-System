# scripts/load_courses.py
import os, re, sys
from datetime import datetime, timezone
import pandas as pd
from supabase import create_client

xlsx = sys.argv[1]
sb = create_client(os.environ["SUPABASE_URL"], os.environ["SUPABASE_SERVICE_KEY"])
run_started = datetime.now(timezone.utc).isoformat()

# header is on row 2 (row 1 is the merged title banner)
df = pd.read_excel(xlsx, header=1)
df = df[df["Course Title"].notna()]

if len(df) < 50:
    sys.exit(f"Only {len(df)} rows - aborting, SWDA layout may have changed")

def money(v):
    if not isinstance(v, str) or not v.startswith("S$"):
        return None
    return float(v.replace("S$", "").replace(",", ""))

def course_ref(url):
    m = re.search(r"/courses/([A-Za-z0-9\-]+)", str(url))
    return m.group(1) if m else None

def when(v):
    try:
        return datetime.strptime(str(v).strip(), "%d %B %Y").date().isoformat()
    except Exception:
        return None            # covers "None listed"

df["course_ref"] = df["Course URL"].map(course_ref)
df = df[df["course_ref"].notna()].drop_duplicates("course_ref")

# --- providers first ---
provs = sorted({p.strip() for p in df["Training Provider"].dropna() if p.strip()})
sb.table("provider").upsert([{"provider_name": p} for p in provs],
                            on_conflict="provider_name").execute()
pmap = {r["provider_name"]: r["provider_id"] for r in
        sb.table("provider").select("provider_id, provider_name").execute().data}

# --- courses ---
payload = []
for _, r in df.iterrows():
    prov = str(r["Training Provider"]).strip()
    payload.append({
        "course_ref":           r["course_ref"],
        "course_title":         str(r["Course Title"]).strip(),
        "provider_id":          pmap.get(prov),
        "course_url":           str(r["Course URL"]),
        "full_course_fee":      money(r.get("Full Course Fee")),
        "after_subsidy_fee":    money(r.get("After Subsidy Fee")),
        "after_sfec_fee":       money(r.get("After SFEC Fee")),
        "star_rating":          None if pd.isna(r.get("Star Rating")) else float(r["Star Rating"]),
        "no_of_ratings":        int(r.get("No. of Ratings") or 0),
        "upcoming_course_date": when(r.get("Upcoming Course Date")),
        "about_this_course":    str(r.get("About this course") or "").strip(),
        "what_youll_learn":     str(r.get("What you'll learn") or "").strip(),
        "last_seen_at":         run_started,
        "is_active":            True,
    })

payload = [p for p in payload if p["provider_id"] and p["about_this_course"]]

for i in range(0, len(payload), 500):
    sb.table("course").upsert(payload[i:i+500], on_conflict="course_ref").execute()

gone = sb.table("course").update({"is_active": False}) \
         .lt("last_seen_at", run_started).eq("is_active", True).execute()

print(f"Upserted {len(payload)} · retired {len(gone.data)}")
