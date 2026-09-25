# scripts/load_courses.py
import os, csv, sys
from datetime import datetime, timezone
from supabase import create_client

sb = create_client(os.environ["SUPABASE_URL"],
                   os.environ["SUPABASE_SERVICE_KEY"])

run_started = datetime.now(timezone.utc).isoformat()

rows = list(csv.DictReader(open("data/swda_courses.csv", encoding="utf-8")))
if len(rows) < 100:                       # sanity guard
    sys.exit(f"Only {len(rows)} rows scraped - aborting, SWDA layout may have changed")

# --- providers first (courses reference them) ---
providers = sorted({r["training_provider"].strip() for r in rows if r["training_provider"]})
sb.table("provider").upsert(
    [{"provider_name": p} for p in providers],
    on_conflict="provider_name").execute()

pmap = {p["provider_name"]: p["provider_id"]
        for p in sb.table("provider").select("provider_id, provider_name")
                   .execute().data}

# --- courses, in chunks ---
payload = [{
    "course_ref":           r["course_ref"],
    "course_title":         r["course_title"],
    "provider_id":          pmap[r["training_provider"].strip()],
    "course_url":           r["course_url"],
    "full_course_fee":      r["full_course_fee"] or None,
    "star_rating":          r["star_rating"] or None,
    "no_of_ratings":        r["no_of_ratings"] or None,
    "upcoming_course_date": r["upcoming_course_date"] or None,
    "about_this_course":    r["about_this_course"],
    "last_seen_at":         run_started,
    "is_active":            True,
} for r in rows]

for i in range(0, len(payload), 500):
    sb.table("course").upsert(payload[i:i+500], on_conflict="course_ref").execute()

# --- retire anything SWDA no longer lists ---
gone = sb.table("course").update({"is_active": False}) \
         .lt("last_seen_at", run_started).eq("is_active", True).execute()

print(f"Upserted {len(payload)} courses · retired {len(gone.data)}")
