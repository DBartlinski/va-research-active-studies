"""
crossref_nih.py
Fetches VA Research projects from NIH RePORTER and matches them against
Book1.xlsx by PI last name + title similarity score.

Outputs nih_matches.json — a dict keyed by Project_ID containing:
  { url, matched (bool), score, appl_id }

Run this once (takes ~2-3 minutes); then re-run generate_site.py to rebuild
index.html with NIH Reporter links.  Re-run whenever data is refreshed.
"""

import requests
import pandas as pd
import json
import time
from difflib import SequenceMatcher
from urllib.parse import quote
from collections import defaultdict

EXCEL_PATH   = 'Book1.xlsx'
OUTPUT_JSON  = 'nih_matches.json'
NIH_API      = 'https://api.reporter.nih.gov/v2/projects/search'
MATCH_THRESH = 0.62          # title-similarity threshold (0-1)
FY_RANGE     = [2021, 2022, 2023, 2024, 2025, 2026]   # fiscal years to pull

# ── helpers ───────────────────────────────────────────────────────────────

def title_sim(a, b):
    """Normalised title similarity (0–1)."""
    a = ' '.join(str(a).lower().split())
    b = ' '.join(str(b).lower().split())
    return SequenceMatcher(None, a, b).ratio()


def extract_last(contact_pi_name):
    """'BROWN, SHELDON D' → 'brown'"""
    raw = str(contact_pi_name or '').strip()
    return raw.split(',')[0].strip().lower() if raw else ''


# ── load local data ───────────────────────────────────────────────────────

print("Loading Book1.xlsx …")
df = pd.read_excel(EXCEL_PATH)
df = df.drop(columns=['Unnamed: 11'], errors='ignore')
total_local = len(df)
print(f"  {total_local:,} studies to cross-reference\n")

# ── fetch NIH Reporter ───────────────────────────────────────────────────

def fetch_nih_va_projects(fiscal_years):
    all_projects = []
    offset = 0
    limit  = 500
    total  = None

    while total is None or offset < total:
        payload = {
            "criteria": {
                "agencies": ["VA"],
                "fiscal_years": fiscal_years
            },
            "include_fields": [
                "ApplId", "ProjectTitle", "ContactPiName",
                "PrincipalInvestigators", "FiscalYear",
                "ProjectStartDate", "ProjectEndDate",
                "ActivityCode", "ProjectNum", "CoreProjectNum"
            ],
            "limit":      limit,
            "offset":     offset,
            "sort_field": "fiscal_year",
            "sort_order": "desc"
        }

        try:
            r = requests.post(NIH_API, json=payload, timeout=45)
            r.raise_for_status()
            data = r.json()
        except Exception as e:
            print(f"  [!] API error at offset {offset}: {e}")
            break

        results = data.get('results', [])
        if not results:
            break

        all_projects.extend(results)

        if total is None:
            total = data.get('meta', {}).get('total', 0)
            print(f"  NIH Reporter: {total:,} VA projects found for FY {min(fiscal_years)}–{max(fiscal_years)}")

        offset += limit
        pct = min(100, int(len(all_projects) / max(total, 1) * 100))
        print(f"  Fetched {len(all_projects):,} / {total:,}  ({pct}%)")
        time.sleep(0.35)   # polite rate limiting

    return all_projects


print(f"Fetching VA projects from NIH Reporter for FY {min(FY_RANGE)}–{max(FY_RANGE)} …")
nih_projects = fetch_nih_va_projects(FY_RANGE)
print(f"\nRetrieved {len(nih_projects):,} NIH Reporter records\n")

# ── build PI-last-name index ─────────────────────────────────────────────
# Most-recent fiscal year first (already sorted desc), so first match per
# PI+title will be the most recent active appl_id.

nih_by_pi = defaultdict(list)
for p in nih_projects:
    last = extract_last(p.get('contact_pi_name', ''))
    if last:
        nih_by_pi[last].append(p)

    # Also index by each co-PI's last name as a secondary lookup
    for pi in (p.get('principal_investigators') or []):
        pi_last = str(pi.get('last_name', '') or '').strip().lower()
        if pi_last and pi_last != last:
            nih_by_pi[pi_last].append(p)

# ── match ─────────────────────────────────────────────────────────────────

print("Matching studies …")
results    = {}
matched    = 0
fallback   = 0

for _, row in df.iterrows():
    pid       = str(row.get('Project_ID', ''))
    pi_last   = str(row.get('PI_Lname', '') or '').strip().lower()
    our_title = str(row.get('Project_Title', '') or '').strip()

    best_score   = 0.0
    best_appl_id = None
    best_nih_num = None

    candidates = nih_by_pi.get(pi_last, [])
    for p in candidates:
        score = title_sim(our_title, p.get('project_title', ''))
        if score > best_score:
            best_score   = score
            best_appl_id = p.get('appl_id')
            best_nih_num = p.get('project_num') or p.get('core_project_num')

    if best_score >= MATCH_THRESH and best_appl_id:
        url = f"https://reporter.nih.gov/project-details/{best_appl_id}"
        matched += 1
        is_matched = True
    else:
        # Fallback: pre-seeded NIH Reporter search with PI name + first 8 title words
        pi_fname = str(row.get('PI_Fname', '') or '').strip()
        pi_lname = str(row.get('PI_Lname', '') or '').strip()
        q = f"{pi_lname} {pi_fname} {' '.join(our_title.split()[:6])}"
        url = f"https://reporter.nih.gov/search?term={quote(q.strip())}"
        fallback    += 1
        is_matched   = False

    results[pid] = {
        'url':     url,
        'matched': is_matched,
        'score':   round(best_score, 3),
        'appl_id': best_appl_id,
        'nih_num': best_nih_num,
    }

pct_matched = matched / total_local * 100 if total_local else 0
print(f"\nResults:")
print(f"  Directly matched to NIH Reporter project: {matched:,}  ({pct_matched:.1f}%)")
print(f"  Fallback search URL:                      {fallback:,}  ({100-pct_matched:.1f}%)")

with open(OUTPUT_JSON, 'w') as f:
    json.dump(results, f, indent=2)

print(f"\nSaved → {OUTPUT_JSON}")
print("Now run  generate_site.py  to rebuild index.html with NIH Reporter links.")
