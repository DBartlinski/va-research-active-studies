"""
generate_site.py
Reads Book1.xlsx and produces a self-contained index.html for public display
of VA Research Active Studies.
"""

import pandas as pd
import json
import os

# ── Load & clean ────────────────────────────────────────────────────────────
df = pd.read_excel(r'Book1.xlsx')
df = df.drop(columns=['Unnamed: 11'], errors='ignore')
df['Project_Start_Date'] = pd.to_datetime(df['Project_Start_Date'], errors='coerce')
df['Project_End_Date']   = pd.to_datetime(df['Project_End_Date'],   errors='coerce')

df['start_fmt'] = df['Project_Start_Date'].dt.strftime('%b %Y').fillna('')
df['end_fmt']   = df['Project_End_Date'].dt.strftime('%b %Y').fillna('')

# ── Status (reference date = today) ─────────────────────────────────────────
today = pd.Timestamp('2026-06-11')
soon  = pd.Timestamp('2027-06-11')
def get_status(d):
    if pd.isna(d):  return 'Active'
    if d <  today:  return 'Completed'
    if d <= soon:   return 'Ending Soon'
    return 'Active'
df['status'] = df['Project_End_Date'].apply(get_status)

# ── PI full name ─────────────────────────────────────────────────────────────
mi = df['PI_MI'].fillna('').str.strip()
df['PI_Name'] = (
    df['PI_Fname'].fillna('').str.strip() + ' ' +
    mi.apply(lambda x: x + '.' if x else '') + ' ' +
    df['PI_Lname'].fillna('').str.strip()
).str.replace(r'\s+', ' ', regex=True).str.strip()

# ── Mappings ─────────────────────────────────────────────────────────────────
CATEGORY_MAP = {
    'Brain, Behavioral and Mental Health Research':        'Mental Health & Brain Research',
    'Cooperative Studies':                                 'Clinical Trials & Cooperative Studies',
    'Gulf War Research':                                   'Gulf War Illness Research',
    'Health Systems Research':                             'Health Systems & Policy Research',
    'Medical Health Research':                             'General Medical Research',
    'Military Exposures Research':                         'Military Exposures Research',
    'Million Veteran Program':                             'Million Veteran Program (MVP)',
    'Pain and Opioid Use Research':                        'Pain & Opioid Research',
    'Precision Oncology Research':                         'Cancer & Oncology Research',
    'Rehabilitation Research, Development and Translation':'Rehabilitation & Recovery Research',
    'Suicide Prevention Research':                         'Suicide Prevention Research',
    'Traumatic Brain Injury Research':                     'Traumatic Brain Injury Research',
}

AWARD_LABEL = {
    'CD': 'Career Development',
    'CN': 'Cooperative Network',
    'MR': 'Merit Review',
    'MS': 'Multi-Site Merit',
    'PI': 'Pilot Investigator',
    'PP': 'Pilot Project',
    'RC': 'Research Career Scientist',
    'RE': 'Research Enhancement',
    'RP': 'Research Project',
    'SD': 'Senior Development',
}

AWARD_GROUP = {
    'MR': 'Research Grant', 'MS': 'Research Grant', 'RP': 'Research Grant',
    'CD': 'Career Award',   'RC': 'Career Award',   'RE': 'Career Award', 'SD': 'Career Award',
    'PP': 'Pilot',          'PI': 'Pilot',
    'CN': 'Cooperative',
}

df['category']    = df['Program_Name'].map(CATEGORY_MAP).fillna(df['Program_Name'])
df['award_label'] = df['Award_Code'].map(AWARD_LABEL).fillna(df['Award_Code'])
df['award_group'] = df['Award_Code'].map(AWARD_GROUP).fillna('Other')

# ── Load NIH Reporter matches (if crossref_nih.py has been run) ──────────────
nih_matches = {}
if os.path.exists('nih_matches.json'):
    with open('nih_matches.json') as f:
        nih_matches = json.load(f)
    nih_count = sum(1 for v in nih_matches.values() if v.get('matched'))
    print(f'Loaded nih_matches.json — {nih_count:,} direct NIH Reporter links')
else:
    print('nih_matches.json not found — run crossref_nih.py to add NIH Reporter links')

# ── Build JSON records ───────────────────────────────────────────────────────
records = []
for _, r in df.iterrows():
    pid = str(r.get('Project_ID', ''))
    nih = nih_matches.get(pid, {})
    records.append({
        'id':       pid,
        'num':      str(r.get('Project_Number', '')),
        'title':    str(r.get('Project_Title', '') or ''),
        'start':    str(r.get('start_fmt', '')),
        'end':      str(r.get('end_fmt', '')),
        'cat':      str(r.get('category', '')),
        'code':     str(r.get('Award_Code', '')),
        'award':    str(r.get('award_label', '')),
        'group':    str(r.get('award_group', '')),
        'pi':       str(r.get('PI_Name', '')),
        'city':     str(r.get('Med_Center_City', '') or ''),
        'state':    str(r.get('Med_Center_State', '') or '').strip(),
        'status':   str(r.get('status', '')),
        'nih_url':  nih.get('url', ''),
        'nih_ok':   nih.get('matched', False),
    })

data_js    = json.dumps(records, ensure_ascii=False)
total      = len(records)
states_js  = json.dumps(sorted(df['Med_Center_State'].dropna().str.strip().unique().tolist()))
cats_js    = json.dumps(sorted(df['category'].dropna().unique().tolist()))

# ── Category colours (one per program) ──────────────────────────────────────
CAT_COLORS = {
    'Mental Health & Brain Research':               '#1565C0',
    'Clinical Trials & Cooperative Studies':        '#00695C',
    'Gulf War Illness Research':                    '#6D4C41',
    'Health Systems & Policy Research':             '#0277BD',
    'General Medical Research':                     '#2E7D32',
    'Military Exposures Research':                  '#4E342E',
    'Million Veteran Program (MVP)':                '#6A1B9A',
    'Pain & Opioid Research':                       '#B71C1C',
    'Cancer & Oncology Research':                   '#BF360C',
    'Rehabilitation & Recovery Research':            '#00695C',
    'Suicide Prevention Research':                  '#4527A0',
    'Traumatic Brain Injury Research':              '#37474F',
}
cat_colors_js = json.dumps(CAT_COLORS)

# ── HTML ─────────────────────────────────────────────────────────────────────
html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>VA Research — Active Studies</title>
<style>
/* ── Reset & base ─────────────────────────────────────────────── */
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;
  background:#f4f6f9;color:#212121;min-height:100vh}}
a{{color:inherit;text-decoration:none}}

/* ── Header ───────────────────────────────────────────────────── */
header{{background:#003f72;color:#fff;padding:0 24px}}
.hdr-inner{{max-width:1400px;margin:0 auto;display:flex;align-items:center;
  justify-content:space-between;padding:16px 0}}
.hdr-brand{{display:flex;align-items:center;gap:14px}}
.hdr-logo{{width:52px;height:52px;background:#fff;border-radius:6px;
  display:flex;align-items:center;justify-content:center}}
.hdr-logo svg{{width:40px;height:40px}}
.hdr-title{{font-size:1.4rem;font-weight:700;letter-spacing:-.3px}}
.hdr-sub{{font-size:.85rem;opacity:.8;margin-top:2px}}
.hdr-badge{{background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.3);
  border-radius:20px;padding:6px 16px;font-size:.9rem;white-space:nowrap}}
.hdr-badge span{{font-weight:700;font-size:1.1rem}}

/* ── Page wrapper ─────────────────────────────────────────────── */
.page{{max-width:1400px;margin:0 auto;padding:24px 24px 48px}}

/* ── Search ───────────────────────────────────────────────────── */
.search-wrap{{position:relative;margin-bottom:16px}}
.search-wrap svg{{position:absolute;left:14px;top:50%;transform:translateY(-50%);
  color:#666;pointer-events:none}}
#searchInput{{width:100%;padding:13px 16px 13px 44px;border:2px solid #d0d7de;
  border-radius:8px;font-size:1rem;background:#fff;transition:border-color .2s}}
#searchInput:focus{{outline:none;border-color:#003f72;box-shadow:0 0 0 3px rgba(0,63,114,.12)}}

/* ── Secondary filters ────────────────────────────────────────── */
.filter-row{{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:20px;align-items:center}}
.filter-row select{{padding:9px 36px 9px 12px;border:2px solid #d0d7de;border-radius:8px;
  font-size:.9rem;background:#fff;cursor:pointer;appearance:none;
  background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6'%3E%3Cpath d='M0 0l5 6 5-6z' fill='%23555'/%3E%3C/svg%3E");
  background-repeat:no-repeat;background-position:right 10px center;min-width:170px;
  transition:border-color .2s}}
.filter-row select:focus{{outline:none;border-color:#003f72}}
#clearBtn{{padding:9px 18px;border:2px solid #003f72;border-radius:8px;
  background:#fff;color:#003f72;font-size:.9rem;font-weight:600;cursor:pointer;
  transition:all .2s;margin-left:auto}}
#clearBtn:hover{{background:#003f72;color:#fff}}

/* ── Category nav ─────────────────────────────────────────────── */
.cat-nav{{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:20px}}
.cat-pill{{padding:8px 16px;border-radius:20px;border:2px solid #d0d7de;
  background:#fff;font-size:.85rem;font-weight:500;cursor:pointer;
  transition:all .15s;display:flex;align-items:center;gap:6px;white-space:nowrap}}
.cat-pill:hover{{border-color:#003f72;color:#003f72}}
.cat-pill.active{{border-color:currentColor;color:#fff}}
.cat-pill .pill-count{{background:rgba(255,255,255,.25);border-radius:10px;
  padding:1px 7px;font-size:.75rem;font-weight:700}}
.cat-pill:not(.active) .pill-count{{background:#f0f0f0;color:#555}}

/* ── Results bar ──────────────────────────────────────────────── */
.results-bar{{display:flex;align-items:center;justify-content:space-between;
  margin-bottom:16px;color:#555;font-size:.9rem}}
.sort-wrap{{display:flex;align-items:center;gap:8px}}
.sort-wrap select{{padding:6px 28px 6px 10px;border:1px solid #d0d7de;border-radius:6px;
  font-size:.85rem;background:#fff;appearance:none;
  background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6'%3E%3Cpath d='M0 0l5 6 5-6z' fill='%23555'/%3E%3C/svg%3E");
  background-repeat:no-repeat;background-position:right 8px center;cursor:pointer}}

/* ── Grid ─────────────────────────────────────────────────────── */
.grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));gap:16px}}

/* ── Card ─────────────────────────────────────────────────────── */
.card{{background:#fff;border-radius:10px;box-shadow:0 1px 4px rgba(0,0,0,.08);
  border:1px solid #e8ecf0;overflow:hidden;display:flex;flex-direction:column;
  transition:box-shadow .2s,transform .15s}}
.card:hover{{box-shadow:0 4px 16px rgba(0,0,0,.12);transform:translateY(-2px)}}
.card-accent{{height:5px;flex-shrink:0}}
.card-body{{padding:16px;flex:1;display:flex;flex-direction:column;gap:10px}}
.card-badges{{display:flex;gap:6px;flex-wrap:wrap}}
.badge{{display:inline-flex;align-items:center;padding:3px 10px;border-radius:12px;
  font-size:.75rem;font-weight:600;letter-spacing:.2px}}
.badge-award{{background:#e8f0fe;color:#1a56db}}
.badge-status-active{{background:#e6f4ea;color:#1e6e2d}}
.badge-status-soon{{background:#fff3e0;color:#b45309}}
.badge-status-completed{{background:#f1f1f1;color:#666}}
.card-title{{font-size:.95rem;font-weight:600;color:#1a1a1a;line-height:1.4;
  display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden}}
.card-meta{{display:flex;flex-direction:column;gap:4px;margin-top:auto}}
.meta-row{{display:flex;align-items:center;gap:6px;font-size:.82rem;color:#555}}
.meta-row svg{{flex-shrink:0;color:#888}}
.meta-row strong{{color:#222}}
.card-cat{{font-size:.75rem;font-weight:600;padding:2px 0;letter-spacing:.1px}}
.card-nih{{margin-top:8px;padding-top:10px;border-top:1px solid #f0f0f0}}
.nih-link{{display:inline-flex;align-items:center;gap:5px;font-size:.8rem;
  font-weight:600;color:#003f72;padding:5px 10px;border:1.5px solid #003f72;
  border-radius:6px;transition:all .15s;background:#fff}}
.nih-link:hover{{background:#003f72;color:#fff}}
.nih-link-search{{color:#555;border-color:#bbb}}
.nih-link-search:hover{{background:#555;color:#fff;border-color:#555}}

/* ── Empty state ──────────────────────────────────────────────── */
.empty{{text-align:center;padding:80px 20px;color:#888}}
.empty svg{{margin-bottom:16px;color:#ccc}}
.empty h3{{font-size:1.1rem;margin-bottom:8px;color:#555}}

/* ── Pagination ───────────────────────────────────────────────── */
.pagination{{display:flex;justify-content:center;align-items:center;
  gap:6px;margin-top:32px;flex-wrap:wrap}}
.pg-btn{{padding:8px 14px;border:2px solid #d0d7de;border-radius:8px;
  background:#fff;font-size:.875rem;cursor:pointer;transition:all .15s;
  font-weight:500;min-width:40px;text-align:center}}
.pg-btn:hover:not(:disabled){{border-color:#003f72;color:#003f72}}
.pg-btn.active{{background:#003f72;border-color:#003f72;color:#fff}}
.pg-btn:disabled{{opacity:.4;cursor:default}}
.pg-info{{font-size:.85rem;color:#666;padding:0 8px}}

/* ── Responsive ───────────────────────────────────────────────── */
@media(max-width:640px){{
  .hdr-inner{{flex-direction:column;gap:12px;text-align:center}}
  .filter-row{{flex-direction:column}}
  .filter-row select{{min-width:100%}}
  #clearBtn{{margin-left:0;width:100%}}
}}
</style>
</head>
<body>

<header>
  <div class="hdr-inner">
    <div class="hdr-brand">
      <div class="hdr-logo">
        <svg viewBox="0 0 40 40" fill="none" xmlns="http://www.w3.org/2000/svg">
          <rect width="40" height="40" rx="4" fill="#003f72"/>
          <text x="50%" y="58%" text-anchor="middle" dominant-baseline="middle"
            fill="#fff" font-family="Arial,sans-serif" font-size="16" font-weight="bold">VA</text>
        </svg>
      </div>
      <div>
        <div class="hdr-title">VA Research Active Studies</div>
        <div class="hdr-sub">Office of Research &amp; Development — Fiscal Year 2026</div>
      </div>
    </div>
    <div class="hdr-badge"><span>{total:,}</span> Active Studies</div>
  </div>
</header>

<div class="page">

  <!-- Search -->
  <div class="search-wrap">
    <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2"
         viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
    <input type="search" id="searchInput"
           placeholder="Search by study title, principal investigator, or location&hellip;">
  </div>

  <!-- Secondary filters -->
  <div class="filter-row">
    <select id="groupSelect">
      <option value="">All Award Types</option>
      <option value="Research Grant">Research Grant (MR, MS, RP)</option>
      <option value="Career Award">Career Award (CD, RC, RE, SD)</option>
      <option value="Pilot">Pilot / Exploratory (PI, PP)</option>
      <option value="Cooperative">Cooperative Network (CN)</option>
    </select>
    <select id="statusSelect">
      <option value="">All Statuses</option>
      <option value="Active">Active</option>
      <option value="Ending Soon">Ending Soon (within 12 months)</option>
      <option value="Completed">Completed</option>
    </select>
    <select id="stateSelect">
      <option value="">All States</option>
    </select>
    <select id="sortSelect">
      <option value="title">Sort: Title A–Z</option>
      <option value="end_asc">Sort: End Date (Soonest)</option>
      <option value="end_desc">Sort: End Date (Latest)</option>
      <option value="pi">Sort: PI Name</option>
      <option value="state">Sort: State</option>
    </select>
    <button id="clearBtn">Clear All Filters</button>
  </div>

  <!-- Category pills -->
  <div class="cat-nav" id="catNav"></div>

  <!-- Results bar -->
  <div class="results-bar">
    <span id="resultCount"></span>
  </div>

  <!-- Grid -->
  <div class="grid" id="studiesGrid"></div>

  <!-- Pagination -->
  <div class="pagination" id="pagination"></div>

</div>

<script>
const STUDIES    = {data_js};
const CAT_COLORS = {cat_colors_js};
const STATES     = {states_js};
const ALL_CATS   = {cats_js};
const PAGE_SIZE  = 24;

let activeCat  = '';
let curPage    = 1;
let filtered   = STUDIES.slice();

// ── Populate state dropdown ────────────────────────────────────────────────
const stateSelect = document.getElementById('stateSelect');
STATES.forEach(s => {{
  const o = document.createElement('option');
  o.value = o.textContent = s;
  stateSelect.appendChild(o);
}});

// ── Category counts (from full dataset) ──────────────────────────────────
const catCounts = {{}};
STUDIES.forEach(s => {{
  catCounts[s.cat] = (catCounts[s.cat] || 0) + 1;
}});

// ── Build category pills ──────────────────────────────────────────────────
function buildCatNav() {{
  const nav = document.getElementById('catNav');
  nav.innerHTML = '';

  // "All" pill
  const all = makePill('All Studies', STUDIES.length, '', activeCat === '');
  nav.appendChild(all);

  ALL_CATS.forEach(cat => {{
    const pill = makePill(cat, catCounts[cat] || 0, cat, activeCat === cat);
    nav.appendChild(pill);
  }});
}}

function makePill(label, count, cat, isActive) {{
  const btn = document.createElement('button');
  btn.className = 'cat-pill' + (isActive ? ' active' : '');
  const color = CAT_COLORS[cat] || '#003f72';
  if (isActive) {{
    btn.style.background = color;
    btn.style.borderColor = color;
    btn.style.color = '#fff';
  }} else {{
    btn.style.color = color;
    btn.style.borderColor = '#d0d7de';
  }}
  btn.innerHTML = `${{label}} <span class="pill-count">${{count.toLocaleString()}}</span>`;
  btn.addEventListener('click', () => {{
    activeCat = cat;
    buildCatNav();
    applyFilters();
  }});
  return btn;
}}

// ── Filter & sort ─────────────────────────────────────────────────────────
function applyFilters() {{
  const q      = document.getElementById('searchInput').value.trim().toLowerCase();
  const group  = document.getElementById('groupSelect').value;
  const status = document.getElementById('statusSelect').value;
  const state  = stateSelect.value;
  const sort   = document.getElementById('sortSelect').value;

  filtered = STUDIES.filter(s => {{
    if (activeCat && s.cat !== activeCat) return false;
    if (group  && s.group  !== group)  return false;
    if (status && s.status !== status) return false;
    if (state  && s.state  !== state)  return false;
    if (q) {{
      const hay = (s.title + ' ' + s.pi + ' ' + s.city + ' ' + s.state).toLowerCase();
      if (!hay.includes(q)) return false;
    }}
    return true;
  }});

  // Sort
  filtered.sort((a, b) => {{
    switch (sort) {{
      case 'title':    return a.title.localeCompare(b.title);
      case 'end_asc':  return (a.end || 'zzz').localeCompare(b.end || 'zzz');
      case 'end_desc': return (b.end || '').localeCompare(a.end || '');
      case 'pi':       return a.pi.localeCompare(b.pi);
      case 'state':    return a.state.localeCompare(b.state);
      default:         return 0;
    }}
  }});

  curPage = 1;
  render();
}}

// ── Render ────────────────────────────────────────────────────────────────
function render() {{
  const grid = document.getElementById('studiesGrid');
  const start = (curPage - 1) * PAGE_SIZE;
  const page  = filtered.slice(start, start + PAGE_SIZE);

  document.getElementById('resultCount').textContent =
    filtered.length === STUDIES.length
      ? `Showing all ${{STUDIES.length.toLocaleString()}} studies`
      : `Showing ${{filtered.length.toLocaleString()}} of ${{STUDIES.length.toLocaleString()}} studies`;

  if (page.length === 0) {{
    grid.innerHTML = `
      <div class="empty" style="grid-column:1/-1">
        <svg width="48" height="48" fill="none" stroke="currentColor" stroke-width="1.5" viewBox="0 0 24 24">
          <circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/>
        </svg>
        <h3>No studies match your filters</h3>
        <p>Try adjusting your search or clearing some filters.</p>
      </div>`;
    renderPagination();
    return;
  }}

  grid.innerHTML = page.map(cardHTML).join('');
  renderPagination();
}}

function cardHTML(s) {{
  const color  = CAT_COLORS[s.cat] || '#003f72';
  const loc    = [s.city, s.state].filter(Boolean).join(', ');
  const dates  = [s.start, s.end].filter(Boolean).join(' – ');
  const statusClass = s.status === 'Active'      ? 'badge-status-active'
                    : s.status === 'Ending Soon'  ? 'badge-status-soon'
                    : 'badge-status-completed';

  return `
  <div class="card">
    <div class="card-accent" style="background:${{color}}"></div>
    <div class="card-body">
      <div class="card-badges">
        <span class="badge badge-award">${{esc(s.award)}}</span>
        <span class="badge ${{statusClass}}">${{esc(s.status)}}</span>
      </div>
      <div class="card-cat" style="color:${{color}}">${{esc(s.cat)}}</div>
      <div class="card-title">${{esc(s.title)}}</div>
      <div class="card-meta">
        ${{s.pi ? `<div class="meta-row">
          <svg width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
            <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/>
            <circle cx="12" cy="7" r="4"/>
          </svg>
          <span>${{esc(s.pi)}}</span>
        </div>` : ''}}
        ${{loc ? `<div class="meta-row">
          <svg width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
            <path d="M12 2C8.13 2 5 5.13 5 9c0 5.25 7 13 7 13s7-7.75 7-13c0-3.87-3.13-7-7-7z"/>
            <circle cx="12" cy="9" r="2.5"/>
          </svg>
          <span>${{esc(loc)}}</span>
        </div>` : ''}}
        ${{dates ? `<div class="meta-row">
          <svg width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
            <rect x="3" y="4" width="18" height="18" rx="2" ry="2"/>
            <line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/>
            <line x1="3" y1="10" x2="21" y2="10"/>
          </svg>
          <span>${{esc(dates)}}</span>
        </div>` : ''}}
        <div class="meta-row">
          <svg width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
            <polyline points="14 2 14 8 20 8"/>
          </svg>
          <span style="color:#888">Award: <strong>${{esc(s.code)}}</strong> &nbsp;|&nbsp; Project #${{esc(s.num)}}</span>
        </div>
      </div>
      ${{s.nih_url ? `
      <div class="card-nih">
        <a class="nih-link${{s.nih_ok ? '' : ' nih-link-search'}}" href="${{esc(s.nih_url)}}" target="_blank" rel="noopener noreferrer">
          <svg width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
            <path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/>
            <polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/>
          </svg>
          ${{s.nih_ok ? 'View on NIH RePORTER' : 'Search NIH RePORTER'}}
        </a>
      </div>` : ''}}
    </div>
  </div>`;
}}

function esc(str) {{
  return String(str || '')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;');
}}

// ── Pagination ────────────────────────────────────────────────────────────
function renderPagination() {{
  const totalPages = Math.ceil(filtered.length / PAGE_SIZE);
  const pag = document.getElementById('pagination');
  if (totalPages <= 1) {{ pag.innerHTML = ''; return; }}

  let html = '';
  html += `<button class="pg-btn" id="pg-prev" ${{curPage===1?'disabled':''}}>&#8592; Prev</button>`;

  // Show window of pages
  const delta = 2;
  const pages = [];
  for (let i = 1; i <= totalPages; i++) {{
    if (i===1 || i===totalPages || (i>=curPage-delta && i<=curPage+delta)) pages.push(i);
    else if (pages[pages.length-1] !== '...') pages.push('...');
  }}
  pages.forEach(p => {{
    if (p === '...') html += `<span class="pg-info">…</span>`;
    else html += `<button class="pg-btn${{p===curPage?' active':''}}" data-page="${{p}}">${{p}}</button>`;
  }});

  html += `<button class="pg-btn" id="pg-next" ${{curPage===totalPages?'disabled':''}}>Next &#8594;</button>`;
  pag.innerHTML = html;

  pag.querySelector('#pg-prev').addEventListener('click', () => {{ curPage--; render(); window.scrollTo(0,0); }});
  pag.querySelector('#pg-next').addEventListener('click', () => {{ curPage++; render(); window.scrollTo(0,0); }});
  pag.querySelectorAll('[data-page]').forEach(btn => {{
    btn.addEventListener('click', () => {{ curPage=+btn.dataset.page; render(); window.scrollTo(0,0); }});
  }});
}}

// ── Event listeners ───────────────────────────────────────────────────────
document.getElementById('searchInput').addEventListener('input', applyFilters);
document.getElementById('groupSelect').addEventListener('change', applyFilters);
document.getElementById('statusSelect').addEventListener('change', applyFilters);
stateSelect.addEventListener('change', applyFilters);
document.getElementById('sortSelect').addEventListener('change', applyFilters);
document.getElementById('clearBtn').addEventListener('click', () => {{
  document.getElementById('searchInput').value = '';
  document.getElementById('groupSelect').value = '';
  document.getElementById('statusSelect').value = '';
  stateSelect.value = '';
  document.getElementById('sortSelect').value = 'title';
  activeCat = '';
  buildCatNav();
  applyFilters();
}});

// ── Init ──────────────────────────────────────────────────────────────────
buildCatNav();
applyFilters();
</script>
</body>
</html>
"""

out_path = r'index.html'
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(html)

print(f"Done! index.html created with {total:,} studies.")
