# VA Research Active Studies

A transparent, searchable directory of active research studies funded by the Department of Veterans Affairs (VA) Office of Research & Development.

## 📊 Dataset

This site displays **1,843 active VA research studies** across 12 research programs:

- Mental Health & Brain Research
- Clinical Trials & Cooperative Studies
- Gulf War Illness Research
- Health Systems & Policy Research
- General Medical Research
- Military Exposures Research
- Million Veteran Program (MVP)
- Pain & Opioid Research
- Cancer & Oncology Research
- Rehabilitation & Recovery Research
- Suicide Prevention Research
- Traumatic Brain Injury Research

## 🔍 Features

- **Category filtering** by research program
- **Award type filtering** (Merit Review, Career Development, Pilot, Cooperative)
- **Status tracking** (Active, Ending Soon, Completed)
- **Location search** by state
- **Free-text search** by title, PI name, or institution
- **Direct links** to 89% of studies in NIH Reporter
- **Responsive design** for desktop and mobile

## 📁 Files

- **`index.html`** — Standalone website (no server needed)
- **`generate_site.py`** — Regenerates HTML from Excel data
- **`crossref_nih.py`** — Matches studies to NIH Reporter and fetches URLs
- **`nih_matches.json`** — Cached NIH Reporter crossref results
- **`Book1.xlsx`** — Source data (not in this repo for privacy)

## 🔄 Updating the Data

```bash
# When your Excel file is updated:
python generate_site.py

# When you want fresh NIH Reporter links (monthly):
python crossref_nih.py
python generate_site.py

# Then push to GitHub:
git add index.html nih_matches.json
git commit -m "Data refresh: [DATE]"
git push
```

## 🛠 Requirements

- Python 3.8+
- `pandas`, `openpyxl`, `requests` (auto-installed in `.venv`)

## 📖 Setup

```bash
# Initial setup
python -m venv .venv
.venv\Scripts\activate
pip install pandas openpyxl requests

# Generate the site
python generate_site.py
```

## 📜 Data Source

Data is compiled from internal VA Research tracking systems and published for transparency. Fiscal Year 2026.

---

**Created:** June 2026 | **Updated:** Regularly

