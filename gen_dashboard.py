import json, openpyxl
from collections import defaultdict

import os as _os
_DIR       = _os.path.dirname(_os.path.abspath(__file__))
GAMES_PATH = _os.path.join(_DIR, 'games_data.json')
EXCEL_PATH = r'C:\Users\shlom\OneDrive\Documents\כדורגל שישי\כדורגל שישי בילו 2026.xlsx'
# fallback for CI builds (GitHub Actions), where the OneDrive path doesn't exist
if not _os.path.exists(EXCEL_PATH):
    EXCEL_PATH = _os.path.join(_DIR, 'soccer.xlsx')
OUT_PATH   = _os.path.join(_DIR, 'dashboard.html')

MERGE_MAP = {
    "אסף ד":            "אסף",
    "אסף ש":            "אסף",
    'אסף ד"ר':          "אסף",
    "אסף דר":           "אסף",
    "קצל":              "קצ'ל",
    "בבצוק":            "בבצ'וק",
    "באבצוק":           "בבצ'וק",
    "יוני?":            "יוני",
    "יוני חבר של רן":   "יוני",
    "תומר":             "תומר סגל",
    "אסף חדש":          "אסף בן ארי",
    "אסף חבר של פוגל":  "אסף בן ארי",
    "זלציקי":           "זליצקי",
    "קליטו":            "קרליטו",
}

SKIP_NAMES = {"עצמי", "שער עצמי"}

MIN_GAMES_THRESHOLD = 20  # players with fewer games are treated as guests and excluded

def normalize(name):
    if not name: return None
    name = name.strip()
    if name in SKIP_NAMES: return None
    return MERGE_MAP.get(name, name)

def extract_year(date_str):
    if not date_str: return 'unknown'
    for part in date_str.replace('-', '/').split('/'):
        if len(part) == 4 and part.isdigit():
            return part
    return 'unknown'

_PLAYER_HDR = {'תוויות שורה', 'שחקן', 'שם', 'Row Labels'}
_MVP_COLS   = {'סכום של MVP', 'MVP'}
_WG_COLS    = {'סכום של שער ניצחון', 'שער ניצחון'}

def read_excel_bonus():
    try:
        wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    except Exception as e:
        print(f'Warning: could not read Excel bonus data: {e}')
        return [], {}
    by_year  = []
    totals   = defaultdict(lambda: {'mvp': 0, 'wg': 0})
    for yr in range(2017, 2027):
        sheet = str(yr)
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        col_p = col_mvp = col_wg = None
        past_header = False
        for row in ws.iter_rows(values_only=True):
            if not past_header:
                for ci, cell in enumerate(row):
                    if isinstance(cell, str) and cell.strip() in _PLAYER_HDR:
                        col_p = ci
                        for ci2, val in enumerate(row):
                            if not isinstance(val, str): continue
                            v = val.strip()
                            if v in _MVP_COLS:  col_mvp = ci2
                            elif v in _WG_COLS: col_wg  = ci2
                        past_header = True
                        break
                continue
            if col_p is None: continue
            pc = row[col_p] if col_p < len(row) else None
            if not pc or not isinstance(pc, str): continue
            player = normalize(pc.strip())
            if not player: continue
            def safe_int(col):
                if col is None or col >= len(row): return 0
                v = row[col]
                try: return int(v) if v is not None else 0
                except: return 0
            m, w = safe_int(col_mvp), safe_int(col_wg)
            totals[player]['mvp'] += m
            totals[player]['wg']  += w
            if m > 0 or w > 0:
                by_year.append({'name': player, 'yr': sheet, 'mvp': m, 'wg': w})
    return by_year, {k: dict(v) for k, v in totals.items()}

def compute_stats(games):
    player_stats  = defaultdict(lambda: {'gm':0,'w':0,'l':0,'d':0,'g':0,'a':0})
    year_stats    = defaultdict(lambda: defaultdict(lambda: {'gm':0,'w':0,'l':0,'d':0,'g':0,'a':0}))
    pair_stats    = defaultdict(lambda: {'w':0,'l':0,'t':0,'g':0,'a':0})
    rival_stats   = defaultdict(lambda: {'fw':0,'t':0,'d':0,'p1g':0,'p2g':0,'p1a':0,'p2a':0})

    for game in games:
        sA = game.get('scoreA', 0)
        sB = game.get('scoreB', 0)
        tA = [normalize(p['name']) for p in game.get('teamA', []) if normalize(p['name'])]
        tB = [normalize(p['name']) for p in game.get('teamB', []) if normalize(p['name'])]
        tAdata = {normalize(p['name']): p for p in game.get('teamA', []) if normalize(p['name'])}
        tBdata = {normalize(p['name']): p for p in game.get('teamB', []) if normalize(p['name'])}
        yr = extract_year(game.get('date', ''))

        if sA > sB:   ra, rb = 'w', 'l'
        elif sA < sB: ra, rb = 'l', 'w'
        else:         ra, rb = 'd', 'd'

        for name in tA:
            pd = tAdata[name]
            player_stats[name]['gm'] += 1
            player_stats[name][ra]   += 1
            player_stats[name]['g']  += pd.get('goals', 0)
            player_stats[name]['a']  += pd.get('assists', 0)
            if yr != 'unknown':
                year_stats[yr][name]['gm'] += 1
                year_stats[yr][name][ra]   += 1
                year_stats[yr][name]['g']  += pd.get('goals', 0)
                year_stats[yr][name]['a']  += pd.get('assists', 0)

        for name in tB:
            pd = tBdata[name]
            player_stats[name]['gm'] += 1
            player_stats[name][rb]   += 1
            player_stats[name]['g']  += pd.get('goals', 0)
            player_stats[name]['a']  += pd.get('assists', 0)
            if yr != 'unknown':
                year_stats[yr][name]['gm'] += 1
                year_stats[yr][name][rb]   += 1
                year_stats[yr][name]['g']  += pd.get('goals', 0)
                year_stats[yr][name]['a']  += pd.get('assists', 0)

        for i, n1 in enumerate(tA):
            for n2 in tA[i+1:]:
                k = tuple(sorted([n1, n2]))
                pair_stats[k]['t'] += 1
                pair_stats[k]['g'] += tAdata[n1].get('goals',0) + tAdata[n2].get('goals',0)
                pair_stats[k]['a'] += tAdata[n1].get('assists',0) + tAdata[n2].get('assists',0)
                if ra == 'w': pair_stats[k]['w'] += 1
                elif ra == 'l': pair_stats[k]['l'] += 1

        for i, n1 in enumerate(tB):
            for n2 in tB[i+1:]:
                k = tuple(sorted([n1, n2]))
                pair_stats[k]['t'] += 1
                pair_stats[k]['g'] += tBdata[n1].get('goals',0) + tBdata[n2].get('goals',0)
                pair_stats[k]['a'] += tBdata[n1].get('assists',0) + tBdata[n2].get('assists',0)
                if rb == 'w': pair_stats[k]['w'] += 1
                elif rb == 'l': pair_stats[k]['l'] += 1

        for pa in tA:
            for pb in tB:
                k = tuple(sorted([pa, pb]))
                first_is_a = (k[0] == pa)
                rival_stats[k]['t'] += 1
                pad = tAdata[pa]; pbd = tBdata[pb]
                if first_is_a:
                    rival_stats[k]['p1g'] += pad.get('goals', 0)
                    rival_stats[k]['p1a'] += pad.get('assists', 0)
                    rival_stats[k]['p2g'] += pbd.get('goals', 0)
                    rival_stats[k]['p2a'] += pbd.get('assists', 0)
                    if ra == 'w':   rival_stats[k]['fw'] += 1
                    elif ra == 'd': rival_stats[k]['d']  += 1
                else:
                    rival_stats[k]['p1g'] += pbd.get('goals', 0)
                    rival_stats[k]['p1a'] += pbd.get('assists', 0)
                    rival_stats[k]['p2g'] += pad.get('goals', 0)
                    rival_stats[k]['p2a'] += pad.get('assists', 0)
                    if rb == 'w':   rival_stats[k]['fw'] += 1
                    elif rb == 'd': rival_stats[k]['d']  += 1

    players = [{'name': n, **v} for n, v in player_stats.items()]
    by_year = [{'yr': yr, 'name': nm, **v}
               for yr, yd in year_stats.items()
               for nm, v in yd.items()]
    pairs   = [{'p1': k[0], 'p2': k[1], **v} for k, v in pair_stats.items()]
    rivals  = [{'p1': k[0], 'p2': k[1], **v} for k, v in rival_stats.items()]

    return {'players': players, 'byYear': by_year, 'pairs': pairs, 'rivals': rivals}

def read_games_bonus(games):
    by_year = []
    totals  = defaultdict(lambda: {'mvp': 0, 'wg': 0})
    for g in games:
        yr  = extract_year(g.get('date', ''))
        mvp = normalize(g.get('mvp') or '')
        wg  = normalize(g.get('wg')  or '')
        if mvp:
            totals[mvp]['mvp'] += 1
            by_year.append({'name': mvp, 'yr': yr, 'mvp': 1, 'wg': 0})
        if wg:
            totals[wg]['wg'] += 1
            by_year.append({'name': wg, 'yr': yr, 'mvp': 0, 'wg': 1})
    return by_year, {k: dict(v) for k, v in totals.items()}

def generate():
    with open(GAMES_PATH, encoding='utf-8-sig') as f:
        games = json.load(f)

    for g in games:
        g['teamA'] = [p for p in g.get('teamA', []) if normalize(p.get('name',''))]
        g['teamB'] = [p for p in g.get('teamB', []) if normalize(p.get('name',''))]
        for p in g['teamA']: p['name'] = normalize(p['name'])
        for p in g['teamB']: p['name'] = normalize(p['name'])

    stats = compute_stats(games)

    keep = {p['name'] for p in stats['players'] if p['gm'] >= MIN_GAMES_THRESHOLD}
    stats['players'] = [p for p in stats['players'] if p['name'] in keep]
    stats['byYear']  = [e for e in stats['byYear']  if e['name'] in keep]
    stats['pairs']   = [p for p in stats['pairs']   if p['p1'] in keep and p['p2'] in keep]
    stats['rivals']  = [r for r in stats['rivals']  if r['p1'] in keep and r['p2'] in keep]

    by_year_excel, totals_excel = read_excel_bonus()
    by_year_games, totals_games = read_games_bonus(games)

    merged = defaultdict(lambda: {'mvp': 0, 'wg': 0})
    for name, v in totals_excel.items():
        merged[name]['mvp'] += v['mvp']; merged[name]['wg'] += v['wg']
    for name, v in totals_games.items():
        merged[name]['mvp'] += v['mvp']; merged[name]['wg'] += v['wg']

    stats['bonus']       = [{'name': k, 'mvp': v['mvp'], 'wg': v['wg']} for k, v in merged.items()]
    stats['bonusByYear'] = by_year_excel + by_year_games

    # Deterministic ordering — Python set/dict iteration order varies between runs,
    # which otherwise produces spurious diffs and merge conflicts with the CI auto-build.
    stats['players']     = sorted(stats['players'],     key=lambda p: (-p['gm'], p['name']))
    stats['byYear']      = sorted(stats['byYear'],      key=lambda e: (e['yr'], e['name']))
    stats['pairs']       = sorted(stats['pairs'],       key=lambda p: (p['p1'], p['p2']))
    stats['rivals']      = sorted(stats['rivals'],      key=lambda r: (r['p1'], r['p2']))
    stats['bonus']       = sorted(stats['bonus'],       key=lambda b: b['name'])
    stats['bonusByYear'] = sorted(stats['bonusByYear'], key=lambda b: (b['name'], b['yr']))

    stats_json = json.dumps(stats, ensure_ascii=False)
    games_json = json.dumps(games, ensure_ascii=False)

    html = HTML_TEMPLATE.replace('STATS_PLACEHOLDER', stats_json).replace('GAMES_PLACEHOLDER', games_json)
    with open(OUT_PATH, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'Dashboard saved: {OUT_PATH}')
    print(f'Players: {len(stats["players"])} | Games: {len(games)} | Pairs: {len(stats["pairs"])} | Rivals: {len(stats["rivals"])}')
    mvp_total = sum(v['mvp'] for v in merged.values())
    wg_total  = sum(v['wg']  for v in merged.values())
    print(f'Bonus data: {len(merged)} players | {mvp_total} MVP awards | {wg_total} winning goals')

HTML_TEMPLATE = r'''<!DOCTYPE html>
<html dir="rtl" lang="he">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>כדורגל שישי - דשבורד</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:Arial,sans-serif;background:#0f172a;color:#e2e8f0;direction:rtl;min-width:0}
h1{text-align:center;padding:20px;font-size:1.8rem;background:linear-gradient(135deg,#1e3a8a,#065f46);color:#fbbf24}
.subtitle{text-align:center;color:#94a3b8;font-size:.85rem;padding:6px 0 14px;background:linear-gradient(135deg,#1e3a8a,#065f46)}
nav{display:flex;flex-wrap:wrap;gap:4px;padding:8px 14px;background:#1e293b;border-bottom:2px solid #334155}
nav button{padding:6px 13px;border:none;border-radius:6px;cursor:pointer;font-size:.85rem;background:#334155;color:#94a3b8;transition:all .2s}
nav button.active{background:#fbbf24;color:#0f172a;font-weight:bold}
nav button:hover:not(.active){background:#475569;color:#e2e8f0}
.tab{display:none;padding:14px}
.tab.active{display:block}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.grid3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px}
.grid4{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}
.card{background:#1e293b;border-radius:10px;padding:12px;border:1px solid #334155}
.card h3{color:#fbbf24;margin-bottom:8px;font-size:.88rem;border-bottom:1px solid #334155;padding-bottom:6px}
.tbl-wrap{overflow-x:auto;max-height:450px;overflow-y:auto}
table{width:100%;border-collapse:collapse;font-size:.78rem;white-space:nowrap}
th{background:#334155;color:#fbbf24;padding:6px 7px;text-align:center;cursor:pointer;user-select:none;position:sticky;top:0;z-index:1}
th:hover{background:#475569}
td{padding:5px 7px;border-bottom:1px solid #1e293b;text-align:center}
td:first-child{text-align:right;font-weight:bold;white-space:normal;min-width:80px}
tr:hover td{background:#273548}
tr:nth-child(even) td{background:#162030}
.bar-wrap{background:#334155;border-radius:3px;height:5px;margin-top:2px;min-width:40px}
.bar{height:5px;border-radius:3px;transition:width .3s}
.bar-gold{background:linear-gradient(90deg,#fbbf24,#f59e0b)}
.bar-green{background:linear-gradient(90deg,#10b981,#059669)}
.bar-red{background:linear-gradient(90deg,#ef4444,#dc2626)}
.bar-blue{background:linear-gradient(90deg,#3b82f6,#2563eb)}
select,input[type=text]{background:#334155;color:#e2e8f0;border:1px solid #475569;border-radius:6px;padding:6px 10px;font-size:.85rem}
select:focus,input:focus{outline:none;border-color:#fbbf24}
.hero-card{background:#162030;border:1px solid #334155;border-radius:10px;padding:12px}
.vs-wrap{display:flex;align-items:center;justify-content:space-around;padding:14px;background:#0f172a;border-radius:8px;margin-top:8px}
.vs-player{text-align:center;min-width:110px}
.vs-name{font-size:1rem;font-weight:bold;color:#fbbf24;margin-bottom:4px;cursor:pointer}
.vs-name:hover{text-decoration:underline}
.vs-wins{font-size:2rem;font-weight:bold}
.vs-label{font-size:.7rem;color:#64748b;margin-top:2px}
.vs-mid{font-size:1.2rem;color:#475569;text-align:center}
.chip{display:inline-block;padding:2px 7px;border-radius:12px;font-size:.68rem;margin:1px}
.chip-green{background:#064e3b;color:#6ee7b7}
.chip-red{background:#7f1d1d;color:#fca5a5}
.chip-blue{background:#1e3a8a;color:#93c5fd}
.chip-orange{background:#78350f;color:#fcd34d}
.chart-wrap{position:relative;height:220px}
.pair-row{display:flex;justify-content:space-between;align-items:center;padding:5px 0;border-bottom:1px solid #334155;font-size:.8rem}
.trend-up{color:#10b981;font-size:1.1rem;font-weight:bold}
.trend-down{color:#ef4444;font-size:1.1rem;font-weight:bold}
.trend-same{color:#fbbf24;font-size:1.1rem}
/* clickable player names */
.pl{cursor:pointer;border-bottom:1px dotted #475569;transition:color .15s}
.pl:hover{color:#fbbf24;border-bottom-color:#fbbf24}
/* streak badges */
.sk{font-size:.68rem;font-weight:bold;padding:2px 6px;border-radius:8px}
.sk-w{background:#064e3b;color:#6ee7b7}
.sk-l{background:#7f1d1d;color:#fca5a5}
.sk-d{background:#1e3a8a;color:#93c5fd}
/* player profile modal */
.modal-ov{display:none;position:fixed;inset:0;background:rgba(0,0,0,.78);z-index:1000;overflow-y:auto;padding:16px}
.modal-ov.open{display:flex;align-items:flex-start;justify-content:center}
.modal-box{background:#1e293b;border-radius:14px;border:1px solid #475569;width:100%;max-width:740px;padding:22px 20px 20px;position:relative;margin:auto}
.modal-close{position:absolute;top:12px;left:12px;background:#334155;border:none;color:#94a3b8;font-size:.82rem;cursor:pointer;padding:4px 10px;border-radius:6px}
.modal-close:hover{background:#475569;color:#e2e8f0}
.modal-chart{position:relative;height:190px;margin:10px 0}
.stat-box{text-align:center;background:#0f172a;border-radius:8px;padding:8px 4px}
.stat-box .sv{font-size:1.1rem;font-weight:bold}
.stat-box .sl{font-size:.63rem;color:#64748b;margin-top:2px}
/* profile compact stat grid — stays multi-column even on phones */
.pstats{display:grid;grid-template-columns:repeat(4,1fr);gap:6px}
@media(max-width:600px){.pstats{grid-template-columns:repeat(3,1fr)}}
/* visitor widget */
.vis-btn{position:fixed;bottom:14px;left:14px;background:#1e3a8a;color:#93c5fd;border:1px solid #3b82f6;border-radius:20px;padding:5px 12px;font-size:.75rem;cursor:pointer;z-index:500;transition:background .2s;white-space:nowrap}
.vis-btn:hover{background:#2d4fa3}
.vis-panel{display:none;position:fixed;bottom:50px;left:14px;background:#1e293b;border:1px solid #334155;border-radius:10px;padding:14px;z-index:600;min-width:220px;box-shadow:0 4px 20px rgba(0,0,0,.6)}
.vis-panel.open{display:block}
@media(max-width:600px){
  .grid2,.grid3,.grid4{grid-template-columns:1fr}
  .modal-box{padding:14px 12px 14px}
  h1{font-size:1.3rem;padding:14px}
}
@media(max-width:900px) and (min-width:601px){.grid3,.grid4{grid-template-columns:1fr 1fr}}
</style>
</head>
<body>
<h1>⚽ כדורגל שישי</h1>
<div class="subtitle" id="mainSubtitle">טוען נתונים...</div>
<nav>
  <button class="active" onclick="showTab('overview',this)">🏠 סקירה</button>
  <button onclick="showTab('lastgame',this)">📋 מחזור אחרון</button>
  <button onclick="showTab('leaderboard',this)">🏆 מצטיינים</button>
  <button onclick="showTab('records',this);buildRecords()">🏅 שיאים</button>
  <button onclick="showTab('me',this);buildMe()">👤 הדף שלי</button>
  <button onclick="showTab('h2h',this)">⚔️ ראש בראש</button>
  <button onclick="showTab('kosher',this)">💪 כושר</button>
</nav>

<!-- OVERVIEW -->
<div id="tab-overview" class="tab active">
  <div class="card" style="margin-bottom:12px">
    <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;border-bottom:1px solid #334155;padding-bottom:6px;margin-bottom:8px">
      <h3 id="curYearTitle" style="border:none;padding:0;margin:0">🗓️ שנה נוכחית</h3>
      <select id="ovYearSel" onchange="renderYearHeroes(this.value)"></select>
    </div>
    <div class="grid4" id="curYearHeroes"></div>
  </div>
  <div class="grid4" style="margin-bottom:12px" id="heroes"></div>
  <div class="grid2">
    <div class="card">
      <h3>📊 משחקים לפי שנה</h3>
      <div class="chart-wrap"><canvas id="chartYears"></canvas></div>
    </div>
    <div class="card">
      <h3>⚽ מבקיעים מובילים</h3>
      <div class="chart-wrap"><canvas id="chartScorers"></canvas></div>
    </div>
  </div>
  <div class="grid2" style="margin-top:12px">
    <div class="card">
      <h3>📈 מובילי נקודות</h3>
      <div class="chart-wrap"><canvas id="chartPoints"></canvas></div>
    </div>
    <div class="card">
      <h3>🎯 אחוזי ניצחון מובילים</h3>
      <div class="chart-wrap"><canvas id="chartWinPct"></canvas></div>
    </div>
  </div>
  <div class="grid2" style="margin-top:12px">
    <div class="card">
      <h3>🅰️ מלכי הבישולים</h3>
      <div class="chart-wrap"><canvas id="chartAssists"></canvas></div>
    </div>
    <div class="card">
      <h3>⚔️ יריבויות הכי נפוצות</h3>
      <div id="topRivalsOverview" style="max-height:220px;overflow-y:auto"></div>
    </div>
  </div>
</div>

<!-- LEADERBOARD -->
<div id="tab-leaderboard" class="tab">
  <div class="card">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;flex-wrap:wrap;gap:8px">
      <h3 style="border:none;padding:0;margin:0">🏆 טבלת מצטיינים</h3>
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <input type="text" id="searchPlayer" placeholder="חיפוש שחקן..." oninput="filterTable()" style="width:140px">
        <select id="lbYear" onchange="filterTable()"></select>
        <select id="minGames" onchange="filterTable()">
          <option value="1">כולם</option>
          <option value="20">מינ׳ 20</option>
          <option value="50">מינ׳ 50</option>
          <option value="100" selected>מינ׳ 100</option>
        </select>
      </div>
    </div>
    <div class="tbl-wrap">
      <table>
        <thead><tr>
          <th onclick="sortTable('name')">#שחקן</th>
          <th onclick="sortTable('gm')">משח׳</th>
          <th onclick="sortTable('w')">נצח׳</th>
          <th onclick="sortTable('winpct')">%נצח</th>
          <th>רצף</th>
          <th onclick="sortTable('pts')">נק׳</th>
          <th onclick="sortTable('ppg')">נק/מ</th>
          <th onclick="sortTable('g')">שע׳</th>
          <th onclick="sortTable('gpg')">ש/מ</th>
          <th onclick="sortTable('a')">בישול</th>
          <th onclick="sortTable('apg')">ב/מ</th>
          <th onclick="sortTable('contrib')">תרומה/מ</th>
          <th onclick="sortTable('mvp_n')" style="color:#fbbf24">🏅 MVP</th>
          <th onclick="sortTable('wg_n')" style="color:#10b981">⚡ שניצ׳</th>
        </tr></thead>
        <tbody id="lbBody"></tbody>
      </table>
    </div>
  </div>
  <div class="card" style="margin-top:12px">
    <h3>🏅 מלכי MVP כל הזמנים (2017+)</h3>
    <div class="tbl-wrap" id="mvpLeaderboard"></div>
  </div>
</div>

<!-- RECORDS -->
<div id="tab-records" class="tab">
  <div class="card" style="margin-bottom:12px">
    <h3 style="border:none;padding:0;margin:0">🏅 שיאי כל הזמנים</h3>
    <div style="font-size:.72rem;color:#475569;margin-top:4px">השיאים הגדולים ב-15 שנות כדורגל שישי — לחיצה על שם פותחת דף שחקן</div>
  </div>
  <div id="recordsBody"></div>
</div>

<!-- KOSHER -->
<div id="tab-kosher" class="tab">
  <div class="card">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;flex-wrap:wrap;gap:8px">
      <h3 style="border:none;padding:0;margin:0">💪 דירוג כושר שחקנים</h3>
      <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center">
        <span style="font-size:.75rem;color:#64748b">מינ׳ משחקים:</span>
        <select id="kosherMinGames" onchange="buildKosher()">
          <option value="1">כולם</option>
          <option value="5" selected>5+</option>
          <option value="10">10+</option>
          <option value="20">20+</option>
        </select>
        <span style="font-size:.72rem;color:#475569">
          <span class="trend-up">↑</span> שיפור &nbsp;
          <span class="trend-same">→</span> יציב &nbsp;
          <span class="trend-down">↓</span> ירידה
        </span>
      </div>
    </div>
    <div style="font-size:.72rem;color:#475569;margin-bottom:8px;padding:6px 10px;background:#0f172a;border-radius:6px" id="kosherFormula"></div>
    <div class="tbl-wrap" id="kosherTable"></div>
  </div>
</div>

<!-- H2H -->
<div id="tab-h2h" class="tab">
  <div class="card">
    <h3>⚔️ השוואה ישירה</h3>
    <div style="display:flex;gap:10px;margin-bottom:14px;align-items:center;flex-wrap:wrap">
      <select id="h2hA" onchange="calcH2H()" style="flex:1;min-width:130px"></select>
      <div style="font-weight:bold;color:#fbbf24;font-size:1.1rem">VS</div>
      <select id="h2hB" onchange="calcH2H()" style="flex:1;min-width:130px"></select>
    </div>
    <div id="h2hResult"></div>
  </div>
</div>

<!-- LAST GAME -->
<div id="tab-lastgame" class="tab">
  <div class="card" style="margin-bottom:12px">
    <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap">
      <h3 style="margin:0">📋 בחר מחזור</h3>
      <select id="gameDateSel" onchange="showSelectedGame()" style="background:#1e293b;color:#e2e8f0;border:1px solid #334155;border-radius:6px;padding:6px 10px;font-size:.95rem"></select>
    </div>
  </div>
  <div id="lastGameDisplay"></div>
</div>

<!-- MY PAGE -->
<div id="tab-me" class="tab">
  <div class="card" style="margin-bottom:12px">
    <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap">
      <h3 style="border:none;padding:0;margin:0">👤 הדף האישי שלי</h3>
      <select id="meSelect" onchange="renderMe()" style="min-width:160px"></select>
    </div>
  </div>
  <div id="meContent"><div class="card" style="color:#64748b">בחר שחקן למעלה — או לחץ על כפתור 👤 בפינה כדי שנזכור אותך.</div></div>
</div>

<!-- PLAYER PROFILE MODAL -->
<div id="playerModal" class="modal-ov" onclick="modalBgClick(event)">
  <div class="modal-box">
    <button class="modal-close" onclick="closeModal()">✕ סגור</button>
    <div id="modalContent"></div>
  </div>
</div>

<!-- ENTRY GATE -->
<div id="gateOverlay" style="display:none;position:fixed;inset:0;background:#0f172a;z-index:300;overflow-y:auto">
  <div style="max-width:420px;margin:0 auto;padding:40px 20px;text-align:center">
    <div style="font-size:3rem;margin-bottom:8px">⚽</div>
    <h2 style="color:#fbbf24;font-size:1.5rem;margin-bottom:6px">כדורגל שישי</h2>
    <div style="color:#94a3b8;font-size:.9rem;margin-bottom:24px">מי אתה? בחר את שמך כדי להיכנס</div>
    <select id="gateSelect" onchange="document.getElementById('gateBtn').disabled=!this.value"
      style="width:100%;padding:12px;font-size:1.05rem;margin-bottom:14px;text-align:center">
    </select>
    <button id="gateBtn" disabled onclick="enterGate()"
      style="width:100%;padding:13px;font-size:1.05rem;font-weight:bold;background:#fbbf24;color:#0f172a;border:none;border-radius:10px;cursor:pointer">
      ⚽ כניסה לדשבורד
    </button>
    <div style="color:#475569;font-size:.72rem;margin-top:14px">הבחירה נשמרת בדפדפן — תתבקש רק פעם אחת</div>
  </div>
</div>
<style>
#gateBtn:disabled{background:#334155;color:#64748b;cursor:not-allowed}
</style>

<!-- VISITOR WIDGET -->
<button class="vis-btn" id="visBtn" onclick="toggleVisPanel()">
  👤 <span id="visLabel">מי אתה?</span>
</button>
<div class="vis-panel" id="visPanel">
  <div style="font-size:.82rem;color:#94a3b8;margin-bottom:8px">בחר את שמך — ייזכר בדפדפן:</div>
  <select id="visSelect" style="width:100%;margin-bottom:8px"></select>
  <button onclick="saveVisitor()" style="width:100%;padding:6px;background:#1e3a8a;color:#93c5fd;border:1px solid #3b82f6;border-radius:6px;cursor:pointer;font-size:.82rem">✓ שמור</button>
  <div style="font-size:.68rem;color:#475569;margin-top:8px;line-height:1.4">
    כדי לראות מי נכנס — הגדר Webhook ב-<code style="color:#94a3b8">WEBHOOK_URL</code> בקוד.
    <a href="https://docs.google.com/spreadsheets" target="_blank" style="color:#3b82f6">הוראות: Google Sheets + Apps Script</a>
  </div>
</div>

<script>
const STATS = STATS_PLACEHOLDER;
const GAMES = GAMES_PLACEHOLDER;

// ── Visitor tracking webhook ─────────────────────────────────────────────────
// 1. Open script.google.com → New project → paste the following code:
//    function doGet(e){
//      var sheet = SpreadsheetApp.openById('YOUR_SHEET_ID').getActiveSheet();
//      sheet.appendRow([new Date(), e.parameter.v, e.parameter.ua]);
//      return ContentService.createTextOutput('ok');
//    }
// 2. Deploy → Web app → Anyone can access → copy the URL
// 3. Paste it below:
const WEBHOOK_URL = 'https://script.google.com/macros/s/AKfycbyPufHBwouLd74f2oWQBeZCnnZmSQ3qmvnW1784VKeEy5PHhb6vL7tSseMiu6R7pBIV/exec';
// ─────────────────────────────────────────────────────────────────────────────

const BONUS = {};
(STATS.bonus||[]).forEach(b => BONUS[b.name] = b);
const BONUS_BY_YEAR = {};
(STATS.bonusByYear||[]).forEach(b => {
  if (!BONUS_BY_YEAR[b.name]) BONUS_BY_YEAR[b.name] = {};
  if (!BONUS_BY_YEAR[b.name][b.yr]) BONUS_BY_YEAR[b.name][b.yr] = {mvp:0,wg:0};
  BONUS_BY_YEAR[b.name][b.yr].mvp += b.mvp||0;
  BONUS_BY_YEAR[b.name][b.yr].wg  += b.wg||0;
});

// Dynamic kosher years: last 3 years that have data
const ALL_YRS = [...new Set(STATS.byYear.map(e=>e.yr))].filter(y=>y&&y!=='unknown').sort();

// ── Date parsing (M/D/YYYY — American format, as written by entry.html) ──────
function parseDate(d) {
  if (!d) return 0;
  const p = d.replace(/-/g,'/').split('/');
  if (p.length === 3) {
    const [a,b,c] = [+p[0],+p[1],+p[2]];
    // M/D/YYYY (e.g. 8/12/2011 = Aug 12) or YYYY/M/D
    return c > 1000 ? new Date(c,a-1,b).getTime() : new Date(a,b-1,c).getTime();
  }
  return new Date(d).getTime();
}

// ── Streak computation (all players, sorted by date) ─────────────────────────
// run = consecutive games matching a condition; tracks best with date range
function _pushRun(c, key, ok, dt) {
  if (!ok) { c['cur'+key] = 0; return; }
  if (!c['cur'+key]) c['_f'+key] = dt;
  c['cur'+key] = (c['cur'+key]||0) + 1;
  if (c['cur'+key] > (c['best'+key]||0)) {
    c['best'+key] = c['cur'+key];
    c['best'+key+'From'] = c['_f'+key];
    c['best'+key+'To']   = dt;
  }
}
function computeStreaks(games) {
  const sorted = [...games].sort((a,b) => parseDate(a.date) - parseDate(b.date));
  const s = {};
  sorted.forEach(g => {
    const sA = g.scoreA??0, sB = g.scoreB??0;
    const rA = sA>sB?'W':sA<sB?'L':'D';
    const rB = sB>sA?'W':sB<sA?'L':'D';
    const upd = (name, r, dt) => {
      if (!s[name]) s[name] = {type:null, count:0};
      const c = s[name];
      if (c.type === r) c.count++;
      else { c.type = r; c.count = 1; }
      _pushRun(c, 'W',  r==='W', dt);  // win streak
      _pushRun(c, 'U',  r!=='L', dt);  // unbeaten (W+D)
      _pushRun(c, 'NW', r!=='W', dt);  // winless (L+D)
    };
    (g.teamA||[]).forEach(p => { if(p.name) upd(p.name, rA, g.date); });
    (g.teamB||[]).forEach(p => { if(p.name) upd(p.name, rB, g.date); });
  });
  return s;
}
const STREAKS = computeStreaks(GAMES);
const _yearStreaksCache = {};
function yearStreaks(yr) {
  if (!_yearStreaksCache[yr])
    _yearStreaksCache[yr] = computeStreaks(GAMES.filter(g => (g.date||'').endsWith('/'+yr)));
  return _yearStreaksCache[yr];
}

// ── Per-player chronological game log (for recent form) ──────────────────────
function playerGameLog(name) {
  const log = [];
  GAMES.forEach(g => {
    const inA = (g.teamA||[]).some(p=>p.name===name);
    const inB = (g.teamB||[]).some(p=>p.name===name);
    if (!inA && !inB) return;
    const sA = g.scoreA??0, sB = g.scoreB??0;
    const my = inA?sA:sB, opp = inA?sB:sA;
    const res = my>opp?'W':my<opp?'L':'D';
    const pdat = (inA?g.teamA:g.teamB).find(p=>p.name===name) || {};
    log.push({date:g.date, res, my, opp, g:pdat.goals||0, a:pdat.assists||0});
  });
  log.sort((x,y) => parseDate(x.date)-parseDate(y.date));
  return log;
}

// Percentile rank of each player across the qualifying pool (FBref-style scouting bars).
const PCT_POOL = STATS.players.map(p => ({
  winpct: p.gm ? p.w/p.gm : 0,
  gpg:    p.gm ? p.g/p.gm : 0,
  apg:    p.gm ? p.a/p.gm : 0,
  cpg:    p.gm ? (p.g+p.a)/p.gm : 0,
  gm:     p.gm,
  mvp:    (BONUS[p.name]||{}).mvp || 0,
  bestW:  (STREAKS[p.name]||{}).bestW || 0,
}));
function pctRank(key, val) {
  const n = PCT_POOL.length;
  if (n <= 1) return 100;
  const below = PCT_POOL.filter(x => x[key] < val).length;
  return Math.round(100 * below / (n - 1));
}

// ── Utils ────────────────────────────────────────────────────────────────────
function pct(w,t)  { return t>0 ? Math.round(100*w/t) : 0; }
// M/D/YYYY -> D.M.YY for display
function fmtD(d) {
  if (!d) return '?';
  const [m,dd,y] = d.split('/');
  return `${dd}.${m}.${String(y).slice(-2)}`;
}
// "from – to (N ימים)" display line for a streak record
function streakRange(from, to) {
  const days = Math.round((parseDate(to)-parseDate(from)) / 86400000);
  return `${fmtD(from)} – ${fmtD(to)} (${days} ימים)`;
}
function r2(v)     { return isFinite(v) ? Math.round(v*100)/100 : 0; }
function showTab(name, btn) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('nav button').forEach(b => b.classList.remove('active'));
  document.getElementById('tab-'+name).classList.add('active');
  if (btn) btn.classList.add('active');
}
// Clickable player name
function pl(name) {
  return `<span class="pl" onclick="openProfile(${JSON.stringify(name)})">${name}</span>`;
}
function skHtml(name) {
  const s = STREAKS[name];
  if (!s || !s.type) return '';
  const cls = s.type==='W'?'sk-w':s.type==='L'?'sk-l':'sk-d';
  const lbl = s.type==='W'?'נצח':s.type==='L'?'הפס':'תיקו';
  return `<span class="sk ${cls}">${s.count}${lbl}</span>`;
}

// ── Overview ─────────────────────────────────────────────────────────────────
function renderYearHeroes(yr) {
  const latest = ALL_YRS[ALL_YRS.length-1];
  const sel = document.getElementById('ovYearSel');
  if (sel.value !== yr) sel.value = yr;
  document.getElementById('curYearTitle').textContent =
    `🗓️ הישגי ${yr}` + (yr===latest ? ' — שנה נוכחית' : '');

  const data    = STATS.byYear.filter(e => e.yr===yr);
  const scorers = [...data].sort((a,b)=>b.g-a.g).filter(p=>p.g>0).slice(0,3);
  const assists = [...data].sort((a,b)=>b.a-a.a).filter(p=>p.a>0).slice(0,3);
  const mvp = [...(STATS.bonusByYear||[])].filter(b=>b.yr===yr&&b.mvp>0).sort((a,b)=>b.mvp-a.mvp).slice(0,3);
  const wg  = [...(STATS.bonusByYear||[])].filter(b=>b.yr===yr&&b.wg>0).sort((a,b)=>b.wg-a.wg).slice(0,3);
  const row = (arr, fmt) => arr.length ? arr.map((p,i)=>`
    <div style="display:flex;justify-content:space-between;font-size:.76rem;margin-bottom:3px">
      <span>${i===0?'🥇':i===1?'🥈':'🥉'} ${pl(p.name)}</span>
      <b style="color:#fbbf24">${fmt(p)}</b>
    </div>`).join('') : '<div style="color:#475569;font-size:.75rem">אין נתונים</div>';

  // per-year streak records
  const ys = yearStreaks(yr);
  const known = new Set(STATS.players.map(p=>p.name));
  const topStreak = (key, lbl) => {
    const list = Object.entries(ys)
      .filter(([n,s]) => known.has(n) && (s['best'+key]||0) > 0)
      .sort((a,b) => b[1]['best'+key]-a[1]['best'+key]).slice(0,3);
    return list.length ? list.map(([n,s],i)=>`
      <div style="font-size:.76rem;margin-bottom:4px">
        <div style="display:flex;justify-content:space-between">
          <span>${i===0?'🥇':i===1?'🥈':'🥉'} ${pl(n)}</span>
          <b style="color:#fbbf24">${s['best'+key]} ${lbl}</b>
        </div>
        <div style="font-size:.64rem;color:#64748b;text-align:left">${streakRange(s['best'+key+'From'], s['best'+key+'To'])}</div>
      </div>`).join('') : '<div style="color:#475569;font-size:.75rem">אין נתונים</div>';
  };

  document.getElementById('curYearHeroes').innerHTML = [
    {icon:'⚽', label:`מלך שערים ${yr}`, content:row(scorers,p=>`${p.g} שע׳`)},
    {icon:'🅰️', label:`מלך בישולים ${yr}`, content:row(assists,p=>`${p.a} בישול`)},
    {icon:'🏅', label:`MVP ${yr}`, content:mvp.length?row(mvp,p=>`${p.mvp} MVP`):'<div style="color:#475569;font-size:.75rem">'+(yr===latest?'טרם נקבע':'אין נתונים')+'</div>'},
    {icon:'⚡', label:`שניצ׳ ${yr}`, content:wg.length?row(wg,p=>`${p.wg} שניצ׳`):'<div style="color:#475569;font-size:.75rem">'+(yr===latest?'טרם נקבע':'אין נתונים')+'</div>'},
    {icon:'🔥', label:`רצף ניצחונות ${yr}`, content:topStreak('W','ברצף')},
    {icon:'🛡️', label:`רצף ללא הפסד ${yr}`, content:topStreak('U','ללא הפסד')},
    {icon:'🥶', label:`רצף ללא ניצחון ${yr}`, content:topStreak('NW','ללא ניצחון')},
  ].map(h=>`
    <div class="hero-card">
      <div style="display:flex;align-items:center;gap:6px;margin-bottom:8px">
        <span style="font-size:1.5rem">${h.icon}</span>
        <div style="font-size:.72rem;color:#64748b;font-weight:bold">${h.label}</div>
      </div>
      ${h.content}
    </div>`).join('');
}

function buildOverview() {
  const ps = STATS.players.map(p => ({...p,
    pts: (p.w||0)*2+(p.d||0),
    ppg: p.gm>0 ? ((p.w||0)*2+(p.d||0))/p.gm : 0,
    gpg: p.gm>0 ? p.g/p.gm : 0
  }));
  // threshold = top third by games, min 20
  const sortedGm = [...ps].sort((a,b)=>b.gm-a.gm);
  const thresh = Math.max(20, sortedGm[Math.floor(sortedGm.length/3)]?.gm||20);
  const active = ps.filter(p => p.gm >= thresh);

  const totalGames = GAMES.length;
  const years = [...new Set(GAMES.map(g=>g.date?.match(/(\d{4})/)?.[1]).filter(Boolean))].sort();
  const curYr = years[years.length-1];
  document.getElementById('mainSubtitle').textContent =
    `${years[0]}–${curYr} | ${totalGames} משחקים | ${STATS.players.length} שחקנים`;

  const top3 = (arr, key, fmt) => arr.sort((a,b)=>b[key]-a[key]).slice(0,3).map((p,i)=>`
    <div style="display:flex;justify-content:space-between;font-size:.76rem;margin-bottom:3px">
      <span>${i===0?'🥇':i===1?'🥈':'🥉'} ${pl(p.name)}</span>
      <b style="color:#fbbf24">${fmt(p[key])}</b>
    </div>`).join('');

  // year heroes (with year selector, defaults to current year)
  document.getElementById('ovYearSel').innerHTML =
    [...ALL_YRS].reverse().map(y=>`<option value="${y}">${y}</option>`).join('');
  renderYearHeroes(curYr);

  document.getElementById('heroes').innerHTML = [
    {icon:'🏆', label:'מובילי נקודות',       content:top3([...active],'pts',v=>v+' נק\'')},
    {icon:'⚽', label:'מלכי השערים',          content:top3([...active],'g',v=>v+' ש\'')},
    {icon:'🎯', label:'אחוז ניצחון',          content:[...active].sort((a,b)=>(b.w/b.gm)-(a.w/a.gm)).slice(0,3).map((p,i)=>`
      <div style="display:flex;justify-content:space-between;font-size:.76rem;margin-bottom:3px">
        <span>${i===0?'🥇':i===1?'🥈':'🥉'} ${pl(p.name)}</span>
        <b style="color:#fbbf24">${pct(p.w,p.gm)}%</b>
      </div>`).join('')},
    {icon:'📈', label:'ממוצע נקודות / משחק',  content:top3([...active],'ppg',v=>r2(v))},
    {icon:'🔥', label:'רצף הניצחונות הארוך בהיסטוריה', content:(()=> {
      const names=new Set(STATS.players.map(p=>p.name));
      return Object.entries(STREAKS).filter(([n])=>names.has(n))
        .sort((a,b)=>b[1].bestW-a[1].bestW).slice(0,3).map(([n,s],i)=>`
        <div style="font-size:.76rem;margin-bottom:4px">
          <div style="display:flex;justify-content:space-between">
            <span>${i===0?'🥇':i===1?'🥈':'🥉'} ${pl(n)}</span>
            <b style="color:#fbbf24">${s.bestW} נצח׳ ברצף</b>
          </div>
          <div style="font-size:.64rem;color:#64748b;text-align:left">${streakRange(s.bestWFrom, s.bestWTo)}</div>
        </div>`).join('');
    })()},
    {icon:'🛡️', label:'הרצף הארוך ביותר ללא הפסד', content:(()=> {
      const names=new Set(STATS.players.map(p=>p.name));
      return Object.entries(STREAKS).filter(([n])=>names.has(n))
        .sort((a,b)=>b[1].bestU-a[1].bestU).slice(0,3).map(([n,s],i)=>`
        <div style="font-size:.76rem;margin-bottom:4px">
          <div style="display:flex;justify-content:space-between">
            <span>${i===0?'🥇':i===1?'🥈':'🥉'} ${pl(n)}</span>
            <b style="color:#fbbf24">${s.bestU} ללא הפסד</b>
          </div>
          <div style="font-size:.64rem;color:#64748b;text-align:left">${streakRange(s.bestUFrom, s.bestUTo)}</div>
        </div>`).join('');
    })()},
    {icon:'🥶', label:'הרצף הארוך ביותר ללא ניצחון', content:(()=> {
      const names=new Set(STATS.players.map(p=>p.name));
      return Object.entries(STREAKS).filter(([n])=>names.has(n))
        .sort((a,b)=>b[1].bestNW-a[1].bestNW).slice(0,3).map(([n,s],i)=>`
        <div style="font-size:.76rem;margin-bottom:4px">
          <div style="display:flex;justify-content:space-between">
            <span>${i===0?'🥇':i===1?'🥈':'🥉'} ${pl(n)}</span>
            <b style="color:#94a3b8">${s.bestNW} ללא ניצחון</b>
          </div>
          <div style="font-size:.64rem;color:#64748b;text-align:left">${streakRange(s.bestNWFrom, s.bestNWTo)}</div>
        </div>`).join('');
    })()},
  ].map(h=>`
    <div class="hero-card">
      <div style="display:flex;align-items:center;gap:6px;margin-bottom:8px">
        <span style="font-size:1.5rem">${h.icon}</span>
        <div style="font-size:.72rem;color:#64748b;font-weight:bold">${h.label}</div>
      </div>
      ${h.content}
    </div>`).join('');

  // Charts
  const gByYear = {};
  GAMES.forEach(g => { const m=g.date?.match(/(\d{4})/); if(m) gByYear[m[1]]=(gByYear[m[1]]||0)+1; });
  const yrs = Object.keys(gByYear).sort();
  newChart('chartYears',  'bar', yrs, yrs.map(y=>gByYear[y]), 'משחקים','#3b82f6');
  newChart('chartScorers','bar', active.sort((a,b)=>b.g-a.g).slice(0,10).map(p=>p.name),
                                 active.slice(0,10).map(p=>p.g), 'שערים','#fbbf24');
  const topPts = [...active].sort((a,b)=>b.pts-a.pts).slice(0,10);
  newChart('chartPoints', 'bar', topPts.map(p=>p.name), topPts.map(p=>p.pts), 'נקודות','#10b981');
  const topW = [...active].sort((a,b)=>(b.w/b.gm)-(a.w/a.gm)).slice(0,10);
  newChart('chartWinPct', 'bar', topW.map(p=>p.name), topW.map(p=>pct(p.w,p.gm)), '% ניצחון','#3b82f6', 100);
  const topA = [...active].sort((a,b)=>b.a-a.a).slice(0,10);
  newChart('chartAssists','bar', topA.map(p=>p.name), topA.map(p=>p.a), 'בישולים','#8b5cf6');

  // Top rivals
  const rivals = [...STATS.rivals].filter(r=>r.t>=15).sort((a,b)=>b.t-a.t).slice(0,10);
  document.getElementById('topRivalsOverview').innerHTML = rivals.length ? rivals.map(r => {
    const aW=r.fw, bW=r.t-r.fw-r.d;
    const col = aW>bW?'#10b981':aW<bW?'#ef4444':'#fbbf24';
    return `<div class="pair-row">
      <span style="flex:1">${pl(r.p1)} <span style="color:#475569;font-size:.7rem">vs</span> ${pl(r.p2)}</span>
      <span>
        <span class="chip chip-blue">${r.t} מ׳</span>
        <span class="chip" style="background:#0f172a;color:${col}">${aW}:${bW}</span>
      </span>
    </div>`;
  }).join('') : '<div style="color:#475569;padding:8px">אין מספיק נתונים</div>';
}

// ── Records (Hall of Fame) ───────────────────────────────────────────────────
function buildRecords() {
  const KNOWN = new Set(STATS.players.map(p=>p.name));
  // game-level scans
  let bigWin=null, highScore=null, soloG=null, soloA=null;
  GAMES.forEach(g => {
    const sA=g.scoreA??0, sB=g.scoreB??0, tot=sA+sB, mar=Math.abs(sA-sB);
    if (mar>0 && (!bigWin || mar>bigWin.mar)) bigWin={mar, hi:Math.max(sA,sB), lo:Math.min(sA,sB), date:g.date};
    if (!highScore || tot>highScore.tot) highScore={tot, sA, sB, date:g.date};
    [...(g.teamA||[]), ...(g.teamB||[])].forEach(p => {
      if (!p.name || !KNOWN.has(p.name)) return;
      if ((p.goals||0)   && (!soloG || p.goals>soloG.v))   soloG={v:p.goals,   name:p.name, date:g.date};
      if ((p.assists||0) && (!soloA || p.assists>soloA.v)) soloA={v:p.assists, name:p.name, date:g.date};
    });
  });
  // career leaders (players already filtered to MIN_GAMES_THRESHOLD)
  const ps  = STATS.players;
  const top = (arr,key) => arr.length ? [...arr].sort((a,b)=>b[key]-a[key])[0] : null;
  const mostGames = top(ps,'gm'), mostG = top(ps,'g'), mostA = top(ps,'a');
  const winp = [...ps].filter(p=>p.gm>=50).sort((a,b)=>(b.w/b.gm)-(a.w/a.gm))[0] || null;
  // bonus leaders
  const bonusArr = Object.values(BONUS);
  const mostMvp = [...bonusArr].filter(b=>b.mvp>0).sort((a,b)=>b.mvp-a.mvp)[0] || null;
  const mostWg  = [...bonusArr].filter(b=>b.wg>0).sort((a,b)=>b.wg-a.wg)[0]  || null;
  // longest win streak ever
  const bw = Object.entries(STREAKS).filter(([n])=>KNOWN.has(n)).sort((a,b)=>b[1].bestW-a[1].bestW)[0] || null;

  const card = (icon, label, holder, big, sub) => `
    <div class="hero-card">
      <div style="display:flex;align-items:center;gap:6px;margin-bottom:6px">
        <span style="font-size:1.5rem">${icon}</span>
        <div style="font-size:.72rem;color:#64748b;font-weight:bold">${label}</div>
      </div>
      <div style="font-size:1.4rem;font-weight:bold;color:#fbbf24;line-height:1.1">${big}</div>
      <div style="font-size:.82rem;margin-top:3px">${holder}</div>
      ${sub?`<div style="font-size:.66rem;color:#64748b;margin-top:2px">${sub}</div>`:''}
    </div>`;

  const cards = [];
  if (mostGames) cards.push(card('🎖️','הכי הרבה משחקים (ותק)', pl(mostGames.name), mostGames.gm, 'משחקים בקריירה'));
  if (mostG)     cards.push(card('⚽','מלך השערים — כל הזמנים', pl(mostG.name), mostG.g, 'שערים בקריירה'));
  if (mostA)     cards.push(card('🅰️','מלך הבישולים — כל הזמנים', pl(mostA.name), mostA.a, 'בישולים בקריירה'));
  if (winp)      cards.push(card('🎯','אחוז הניצחון הגבוה ביותר', pl(winp.name), pct(winp.w,winp.gm)+'%', `${winp.w}/${winp.gm} · מינ׳ 50 משחקים`));
  if (bw)        cards.push(card('🔥','רצף הניצחונות הארוך אי-פעם', pl(bw[0]), bw[1].bestW, streakRange(bw[1].bestWFrom, bw[1].bestWTo)));
  if (soloG)     cards.push(card('💥','הכי הרבה שערים במשחק יחיד', pl(soloG.name), soloG.v, fmtD(soloG.date)));
  if (soloA)     cards.push(card('✨','הכי הרבה בישולים במשחק יחיד', pl(soloA.name), soloA.v, fmtD(soloA.date)));
  if (mostMvp)   cards.push(card('🏅','הכי הרבה MVP', pl(mostMvp.name), mostMvp.mvp, 'תארי שחקן המשחק'));
  if (mostWg)    cards.push(card('⚡','הכי הרבה שערי ניצחון', pl(mostWg.name), mostWg.wg, 'שערים מכריעים'));
  if (bigWin)    cards.push(card('🚀','הניצחון הגדול בהיסטוריה', `הפרש ${bigWin.mar} שערים`, `${bigWin.hi}:${bigWin.lo}`, fmtD(bigWin.date)));
  if (highScore) cards.push(card('🎆','המשחק עתיר השערים', `${highScore.tot} שערים`, `${highScore.sA}:${highScore.sB}`, fmtD(highScore.date)));

  document.getElementById('recordsBody').innerHTML =
    `<div class="grid4">${cards.join('')}</div>`;
}

function newChart(id, type, labels, data, label, color, max) {
  new Chart(document.getElementById(id), {type, data:{labels, datasets:[{label, data, backgroundColor:color, borderRadius:4}]},
    options:{plugins:{legend:{display:false}}, scales:{
      y:{max, grid:{color:'#334155'}, ticks:{color:'#94a3b8'}},
      x:{grid:{color:'#334155'}, ticks:{color:'#94a3b8', font:{size:9}}}
    }, maintainAspectRatio:false}});
}

// ── Leaderboard ───────────────────────────────────────────────────────────────
let lbSortKey='w', lbSortDir=-1;
function buildLeaderboard() {
  const names = [...STATS.players].sort((a,b)=>a.name.localeCompare(b.name,'he')).map(p=>p.name);
  document.getElementById('h2hA').innerHTML = names.map(n=>`<option>${n}</option>`).join('');
  document.getElementById('h2hB').innerHTML = names.map(n=>`<option>${n}</option>`).join('');
  document.getElementById('h2hB').selectedIndex = 1;
  // year selector: all-time + each year (newest first)
  document.getElementById('lbYear').innerHTML =
    ['<option value="all">כל הזמנים</option>',
     ...[...ALL_YRS].reverse().map(y=>`<option value="${y}">${y}</option>`)].join('');
  filterTable();
}
function filterTable() {
  const q   = (document.getElementById('searchPlayer').value||'').trim();
  const yr  = document.getElementById('lbYear').value||'all';
  let min   = parseInt(document.getElementById('minGames').value)||1;
  if (yr!=='all' && min>50) min = 10;  // yearly: cap threshold so the table isn't empty
  const base = yr==='all'
    ? STATS.players
    : STATS.byYear.filter(e=>e.yr===yr);
  const rows = base
    .filter(p => p.gm>=min && (q===''||p.name.includes(q)))
    .map(p => ({...p,
      pts:    (p.w||0)*2+(p.d||0),
      ppg:    p.gm>0?((p.w||0)*2+(p.d||0))/p.gm:0,
      winpct: p.gm>0?p.w/p.gm:0,
      gpg:    p.gm>0?p.g/p.gm:0,
      apg:    p.gm>0?p.a/p.gm:0,
      contrib:p.gm>0?(p.g+p.a)/p.gm:0,
      mvp_n:  yr==='all' ? ((BONUS[p.name]||{}).mvp||0) : (((BONUS_BY_YEAR[p.name]||{})[yr]||{}).mvp||0),
      wg_n:   yr==='all' ? ((BONUS[p.name]||{}).wg||0)  : (((BONUS_BY_YEAR[p.name]||{})[yr]||{}).wg||0),
    }));
  rows.sort((a,b) => {
    // dir=-1 means descending (best first)
    if (lbSortKey==='w') {
      let d=a.w-b.w; if(d!==0) return lbSortDir*d;
      d=a.pts-b.pts; if(d!==0) return lbSortDir*d;
      return lbSortDir*(a.g-b.g);
    }
    return lbSortDir*((a[lbSortKey]??0)-(b[lbSortKey]??0));
  });
  document.getElementById('lbBody').innerHTML = rows.map((p,i) => `
    <tr>
      <td>${i+1}. ${pl(p.name)}</td>
      <td>${p.gm}</td>
      <td style="font-weight:bold;color:#10b981">${p.w}</td>
      <td>${pct(p.w,p.gm)}%</td>
      <td>${skHtml(p.name)}</td>
      <td style="color:#fbbf24">${p.pts}</td>
      <td>${r2(p.ppg)}</td>
      <td>${p.g}</td>
      <td>${r2(p.gpg)}</td>
      <td>${p.a}</td>
      <td>${r2(p.apg)}</td>
      <td>${r2(p.contrib)}</td>
      <td style="color:#fbbf24;font-weight:bold">${p.mvp_n||'-'}</td>
      <td style="color:#10b981;font-weight:bold">${p.wg_n||'-'}</td>
    </tr>`).join('');
}
function sortTable(key) {
  if (lbSortKey===key) lbSortDir*=-1; else { lbSortKey=key; lbSortDir=-1; }
  filterTable();
}

// ── MVP leaderboard ───────────────────────────────────────────────────────────
function buildMVPLeaderboard() {
  const rows = [...(STATS.bonus||[])].filter(b=>b.mvp>0||b.wg>0).sort((a,b)=>b.mvp-a.mvp||b.wg-a.wg);
  if (!rows.length) { document.getElementById('mvpLeaderboard').innerHTML='<p style="color:#64748b;padding:10px">אין נתונים</p>'; return; }
  document.getElementById('mvpLeaderboard').innerHTML = `
    <table><thead><tr>
      <th>#</th><th style="text-align:right">שחקן</th>
      <th style="color:#fbbf24">🏅 MVP</th>
      <th style="color:#10b981">⚡ שניצ׳</th>
    </tr></thead>
    <tbody>${rows.map((b,i)=>{
      const medal = i===0?'🥇':i===1?'🥈':i===2?'🥉':'';
      return `<tr>
        <td>${medal||i+1}</td>
        <td style="text-align:right;font-weight:bold">${pl(b.name)}</td>
        <td style="color:#fbbf24;font-weight:bold">${b.mvp||'-'}</td>
        <td style="color:#10b981;font-weight:bold">${b.wg||'-'}</td>
      </tr>`;
    }).join('')}</tbody></table>`;
}

// ── Kosher (day-window based: last 180 days ×2 + previous 180 days ×1) ───────
let ksrSortKey='rating', ksrSortDir=-1, _ksrRows=[];
const KSR_DAY = 24*60*60*1000;
function winScore(wd) {
  if (!wd || wd.gm<1) return null;
  return (wd.w/wd.gm)*40 + (wd.g/wd.gm)*20 + (wd.a/wd.gm)*10 + (wd.mvp/wd.gm)*30 + (wd.wg/wd.gm)*20;
}
function buildKosher() {
  const min = parseInt(document.getElementById('kosherMinGames').value)||5;
  const ref = Math.max(...GAMES.map(g=>parseDate(g.date)));  // date of latest game
  const cutA = ref - 180*KSR_DAY;   // window A: last 180 days
  const cutB = ref - 360*KSR_DAY;   // window B: 180 days before that
  const cutActive = ref - 730*KSR_DAY;  // eligibility: played within the last 730 days
  document.getElementById('kosherFormula').textContent =
    'מוצגים רק מי ששיחקו ב-730 הימים האחרונים | ניקוד = ממוצע משוקלל: 180 הימים האחרונים ×4 + 180 שלפניהם ×2 + עד 730 ימים ×1 | נוסחה: %נצח×40 + ש/מ×20 + ב/מ×10 + MVP/מ×30 + שניצ׳/מ×20';

  // aggregate per player per window from raw games
  // A = last 180d, B = 180-360d, C = 360-730d (all scored, with decreasing weight)
  const agg = {};  // name -> {A:{gm,w,g,a,mvp,wg}, B:{...}, C:{...}}
  const blank = () => ({gm:0,w:0,g:0,a:0,mvp:0,wg:0});
  GAMES.forEach(g => {
    const t = parseDate(g.date);
    if (t < cutActive) return;
    const wnd = t >= cutA ? 'A' : t >= cutB ? 'B' : 'C';
    const sA=g.scoreA??0, sB=g.scoreB??0;
    const res = team => team==='A' ? (sA>sB?'w':sA<sB?'l':'d') : (sB>sA?'w':sB<sA?'l':'d');
    [['A', g.teamA||[]], ['B', g.teamB||[]]].forEach(([team, list]) => {
      list.forEach(p => {
        if (!p.name) return;
        if (!agg[p.name]) agg[p.name] = {A:blank(), B:blank(), C:blank()};
        const w = agg[p.name][wnd];
        w.gm++; if (res(team)==='w') w.w++;
        w.g += p.goals||0; w.a += p.assists||0;
      });
    });
    if (g.mvp && agg[g.mvp]) agg[g.mvp][wnd].mvp++;
    if (g.wg  && agg[g.wg])  agg[g.wg][wnd].wg++;
  });

  _ksrRows = STATS.players
    .filter(p => agg[p.name] && (agg[p.name].A.gm + agg[p.name].B.gm + agg[p.name].C.gm) >= min)
    .map(p => {
      const {A, B, C} = agg[p.name];
      const scA = winScore(A), scB = winScore(B), scC = winScore(C);
      let wSum=0, wTot=0;
      if (scA!==null) { wSum+=scA*4; wTot+=4; }
      if (scB!==null) { wSum+=scB*2; wTot+=2; }
      if (scC!==null) { wSum+=scC*1; wTot+=1; }
      const rating = wTot>0 ? wSum/wTot : 0;
      let trend='same';
      if (scA!==null&&scB!==null) { if(scA>scB*1.05) trend='up'; else if(scA<scB*0.95) trend='down'; }
      else if (scA!==null) trend='up';
      else if (scB!==null) trend='down';
      const gm=A.gm+B.gm+C.gm, w=A.w+B.w+C.w, g=A.g+B.g+C.g, a=A.a+B.a+C.a;
      return {...p, rating, trend, recentGm:gm, recentW:w,
        gmA:A.gm, gmB:B.gm, gmC:C.gm,
        wp:gm>0?w/gm:0, gp:gm>0?g/gm:0, ap:gm>0?a/gm:0,
        mvp_n:A.mvp+B.mvp+C.mvp, wg_n:A.wg+B.wg+C.wg};
    });
  renderKosherTable();
}
function sortKosher(key) {
  if (ksrSortKey===key) ksrSortDir*=-1; else { ksrSortKey=key; ksrSortDir=-1; }
  renderKosherTable();
}
function renderKosherTable() {
  const rows = [..._ksrRows].sort((a,b) => {
    // dir=-1 means descending (most in-form first)
    const va=a[ksrSortKey]??0, vb=b[ksrSortKey]??0;
    if (typeof va==='string') return ksrSortDir*(va<vb?1:va>vb?-1:0);
    return ksrSortDir*(va-vb);
  });
  const maxR = Math.max(...rows.map(r=>r.rating), 0.01);
  const ti   = t => t==='up'?'<span class="trend-up">↑</span>':t==='down'?'<span class="trend-down">↓</span>':'<span class="trend-same">→</span>';
  const thK  = (key,label) => `<th onclick="sortKosher('${key}')" style="${ksrSortKey===key?'color:#fff':''}">${label}${ksrSortKey===key?(ksrSortDir===-1?' ▼':' ▲'):''}</th>`;
  document.getElementById('kosherTable').innerHTML = `
    <table><thead><tr>
      <th>#</th>
      <th onclick="sortKosher('name')" style="text-align:right;${ksrSortKey==='name'?'color:#fff':''}">שחקן${ksrSortKey==='name'?(ksrSortDir===-1?' ▼':' ▲'):''}</th>
      ${thK('recentGm', 'מ׳ סה"כ')}
      ${thK('gmA', '180 אחרונים')}
      ${thK('gmB', '180 שלפני')}
      ${thK('gmC', '360-730')}
      <th>מגמה</th>
      ${thK('wp','%נצח')}
      ${thK('gp','ש/מ')}
      ${thK('ap','ב/מ')}
      ${thK('mvp_n','MVP')}
      ${thK('wg_n','שניצ׳')}
      <th>רצף</th>
      ${thK('rating','ניקוד ★')}
    </tr></thead>
    <tbody>${rows.map((p,i)=>`
      <tr>
        <td>${i+1}</td>
        <td style="text-align:right;font-weight:bold">${pl(p.name)}</td>
        <td>${p.recentGm}</td>
        <td style="color:#94a3b8">${p.gmA||'-'}</td>
        <td style="color:#475569">${p.gmB||'-'}</td>
        <td style="color:#475569">${p.gmC||'-'}</td>
        <td style="text-align:center">${ti(p.trend)}</td>
        <td>${pct(p.recentW,p.recentGm)}%</td>
        <td>${r2(p.gp)}</td>
        <td>${r2(p.ap)}</td>
        <td style="color:#fbbf24;font-weight:bold">${p.mvp_n||'-'}</td>
        <td style="color:#10b981;font-weight:bold">${p.wg_n||'-'}</td>
        <td>${skHtml(p.name)}</td>
        <td>
          <b style="color:#fbbf24;font-size:.9rem">${r2(p.rating)}</b>
          <div class="bar-wrap"><div class="bar bar-gold" style="width:${Math.round(100*p.rating/maxR)}%"></div></div>
        </td>
      </tr>`).join('')}
    </tbody></table>`;
}

// ── H2H ───────────────────────────────────────────────────────────────────────
function calcH2H() {
  const pA=document.getElementById('h2hA').value, pB=document.getElementById('h2hB').value;
  if (pA===pB) { document.getElementById('h2hResult').innerHTML='<p style="color:#ef4444;padding:12px">בחר שני שחקנים שונים</p>'; return; }
  const sorted=[pA,pB].sort();
  const rv=STATS.rivals.find(r=>r.p1===sorted[0]&&r.p2===sorted[1]);
  const aIsFirst=sorted[0]===pA;
  const aW=rv?(aIsFirst?rv.fw:rv.t-rv.fw-rv.d):0;
  const bW=rv?(aIsFirst?rv.t-rv.fw-rv.d:rv.fw):0;
  const draws=rv?rv.d:0, total=rv?rv.t:0;
  const aG=rv?(aIsFirst?rv.p1g:rv.p2g):0, aAs=rv?(aIsFirst?rv.p1a:rv.p2a):0;
  const bG=rv?(aIsFirst?rv.p2g:rv.p1g):0, bAs=rv?(aIsFirst?rv.p2a:rv.p1a):0;
  const sA=STATS.players.find(p=>p.name===pA), sB=STATS.players.find(p=>p.name===pB);
  const pair=STATS.pairs.find(p=>p.p1===sorted[0]&&p.p2===sorted[1]);

  if (!rv) {
    document.getElementById('h2hResult').innerHTML=`<div style="text-align:center;padding:20px;color:#64748b">
      ⚠️ לא נמצאו מפגשים ישירים בין ${pA} ל-${pB}</div>`;
    return;
  }
  let html=`
    <div class="vs-wrap">
      <div class="vs-player">
        <div class="vs-name" onclick="openProfile(${JSON.stringify(pA)})">${pA}</div>
        <div class="vs-wins" style="color:#3b82f6">${aW}</div>
        <div class="vs-label">ניצחונות מולו</div>
        <div style="margin-top:6px;font-size:.76rem;color:#94a3b8">${aG} ש׳ | ${aAs} ב׳</div>
      </div>
      <div class="vs-mid">⚔️<br>
        <span style="font-size:.78rem;color:#64748b">${total} מפגשים</span><br>
        <span style="font-size:.72rem;color:#475569">${draws} תיקו</span>
      </div>
      <div class="vs-player">
        <div class="vs-name" onclick="openProfile(${JSON.stringify(pB)})">${pB}</div>
        <div class="vs-wins" style="color:#ef4444">${bW}</div>
        <div class="vs-label">ניצחונות מולו</div>
        <div style="margin-top:6px;font-size:.76rem;color:#94a3b8">${bG} ש׳ | ${bAs} ב׳</div>
      </div>
    </div>`;
  if (pair) html+=`<div style="text-align:center;margin:10px 0;font-size:.84rem;color:#94a3b8">
    🤝 כשמשחקים <b style="color:#fbbf24">ביחד</b>: ${pair.t} משחקים —
    <b style="color:#10b981">${pair.w} ניצחונות</b> |
    <span style="color:#ef4444">${pair.l} הפסדים</span> |
    ${pct(pair.w,pair.t)}% ניצחון</div>`;
  if (sA&&sB) html+=`<div class="grid2" style="margin-top:10px;gap:8px">
    ${miniCard(sA)}${miniCard(sB)}</div>`;
  document.getElementById('h2hResult').innerHTML=html;
}
function miniCard(p) {
  const pts=(p.w||0)*2+(p.d||0), b=BONUS[p.name]||{mvp:0,wg:0};
  return `<div class="card" style="font-size:.8rem">
    <div style="font-size:.95rem;font-weight:bold;color:#fbbf24;margin-bottom:6px;cursor:pointer" onclick="openProfile(${JSON.stringify(p.name)})">${p.name}</div>
    <div>משחקים: <b>${p.gm}</b> | ניצח: <b>${p.w}</b> (${pct(p.w,p.gm)}%)</div>
    <div>נקודות: <b>${pts}</b> | שערים: <b>${p.g}</b> | בישולים: <b>${p.a}</b></div>
    ${b.mvp?`<div>MVP: <b style="color:#fbbf24">${b.mvp}</b></div>`:''}
    ${b.wg?`<div>שניצ׳: <b style="color:#10b981">${b.wg}</b></div>`:''}
  </div>`;
}

// ── Last game ─────────────────────────────────────────────────────────────────
function buildLastGame() {
  const dates = [...new Set(GAMES.map(g=>g.date))]
    .sort((a,b) => parseDate(b) - parseDate(a));  // newest first
  document.getElementById('gameDateSel').innerHTML = dates.map(d=>`<option value="${d}">${d}</option>`).join('');
  showSelectedGame();
}
function showSelectedGame() {
  const date = document.getElementById('gameDateSel').value;
  const game = [...GAMES].reverse().find(g => g.date===date);
  const el   = document.getElementById('lastGameDisplay');
  if (!game) { el.innerHTML='<div class="card">לא נמצא משחק לתאריך זה</div>'; return; }
  el.innerHTML = `<div class="card" style="margin-bottom:12px;text-align:center">
    <h2 style="margin:0">📅 ${game.date}</h2>
  </div>` + renderGameCard(game);
}
function renderGameCard(game) {
  const sA=game.scoreA??'?', sB=game.scoreB??'?';
  const winner=sA>sB?'A':sB>sA?'B':'X';
  const teamRows = players => (players||[]).map(p => {
    const g=p.goals||0, a=p.assists||0;
    const chips=[];
    if(g) chips.push(`<span class="chip chip-green">${g}⚽</span>`);
    if(a) chips.push(`<span class="chip chip-blue">${a}🅰️</span>`);
    return `<div style="display:flex;justify-content:space-between;align-items:center;padding:4px 0;border-bottom:1px solid #1e293b">
      <span>${pl(p.name)}</span><span>${chips.join(' ')}</span></div>`;
  }).join('');
  const colA=winner==='A'?'#10b981':winner==='B'?'#ef4444':'#fbbf24';
  const colB=winner==='B'?'#10b981':winner==='A'?'#ef4444':'#fbbf24';
  return `<div class="grid2" style="gap:12px">
    <div class="card" style="border-top:3px solid ${colA}">
      <h3 style="color:${colA};text-align:center;margin-bottom:8px">קבוצה א׳ — ${sA}</h3>
      ${teamRows(game.teamA)}
    </div>
    <div class="card" style="border-top:3px solid ${colB}">
      <h3 style="color:${colB};text-align:center;margin-bottom:8px">קבוצה ב׳ — ${sB}</h3>
      ${teamRows(game.teamB)}
    </div>
  </div>
  ${game.mvp?`<div style="margin-top:10px;text-align:center"><span class="chip chip-orange">🏅 MVP: ${pl(game.mvp)}</span></div>`:''}
  ${game.wg ?`<div style="margin-top:4px;text-align:center"><span class="chip chip-green">🥅 שער ניצחון: ${pl(game.wg)}</span></div>`:''}`;
}

// Head-to-head rivalry stats for one player, limited to games within `days` of today.
// Rolling window (uses today's date) so old rivalries drop off over time.
function rivalsWithin(name, days) {
  const cut = Date.now() - days*86400000;
  const acc = {};
  GAMES.forEach(g => {
    if (parseDate(g.date) < cut) return;
    const inA=(g.teamA||[]).some(p=>p.name===name);
    const inB=(g.teamB||[]).some(p=>p.name===name);
    if (!inA && !inB) return;
    const sA=g.scoreA??0, sB=g.scoreB??0;
    const myWin  = inA ? sA>sB : sB>sA;
    const myLoss = inA ? sA<sB : sB<sA;
    (inA ? (g.teamB||[]) : (g.teamA||[])).forEach(o => {
      if (!o.name) return;
      if (!acc[o.name]) acc[o.name] = {other:o.name, t:0, w:0, l:0};
      acc[o.name].t++; if (myWin) acc[o.name].w++; else if (myLoss) acc[o.name].l++;
    });
  });
  return Object.values(acc).map(x => ({...x, wp:x.t?x.w/x.t:0}));
}

// ── Player Profile Modal ──────────────────────────────────────────────────────
let _modalChart = null, _meChart = null;
// Returns the full profile HTML for a player (used by modal + my-page tab)
function profileBody(name, chartId) {
  const p = STATS.players.find(x => x.name===name);
  if (!p) return null;
  const pts = (p.w||0)*2+(p.d||0);
  const b   = BONUS[name]||{mvp:0,wg:0};
  const sk  = STREAKS[name];
  const yrs = STATS.byYear.filter(e=>e.name===name).sort((a,b_)=>a.yr.localeCompare(b_.yr));

  const myPairs = STATS.pairs
    .filter(x => x.p1===name||x.p2===name)
    .map(x => ({other:x.p1===name?x.p2:x.p1, ...x}))
    .sort((a,b_) => b_.t-a.t).slice(0,8);

  const myRivals = STATS.rivals
    .filter(x => x.p1===name||x.p2===name)
    .map(x => {
      const is1 = x.p1===name;
      return {other:is1?x.p2:x.p1, t:x.t, d:x.d,
        w:is1?x.fw:x.t-x.fw-x.d, l:is1?x.t-x.fw-x.d:x.fw};
    })
    .sort((a,b_) => b_.t-a.t).slice(0,8);

  // ── Chemistry & nemesis (min shared games for meaning) ──
  const PAIR_MIN = 20, RIVAL_MIN = 20;
  const pairPool = STATS.pairs
    .filter(x => (x.p1===name||x.p2===name) && x.t>=PAIR_MIN)
    .map(x => ({other:x.p1===name?x.p2:x.p1, t:x.t, w:x.w, wp:x.t?x.w/x.t:0}));
  const bestMate  = pairPool.length ? [...pairPool].sort((a,b_)=>b_.wp-a.wp||b_.t-a.t)[0] : null;
  const worstMate = pairPool.length ? [...pairPool].sort((a,b_)=>a.wp-b_.wp||b_.t-a.t)[0] : null;
  // rivalries: last 5 years only, ≥20 shared games in that window (rolls forward over time)
  const rivalPool = rivalsWithin(name, 5*365).filter(x => x.t>=RIVAL_MIN);
  const nemesis   = rivalPool.length ? [...rivalPool].sort((a,b_)=>a.wp-b_.wp||b_.t-a.t)[0] : null;
  const favVictim = rivalPool.length ? [...rivalPool].sort((a,b_)=>b_.wp-a.wp||b_.t-a.t)[0] : null;
  const chemTile = (lbl, icon, o, extra, col) => o ? `
    <div class="stat-box" style="text-align:right;padding:8px 10px">
      <div style="font-size:.64rem;color:#64748b;margin-bottom:2px">${icon} ${lbl}</div>
      <div style="font-size:.9rem;font-weight:bold;color:${col}">${pl(o.other)}</div>
      <div style="font-size:.64rem;color:#94a3b8;margin-top:1px">${extra(o)}</div>
    </div>` : '';
  const chemHtml = (bestMate||worstMate||nemesis||favVictim) ? `
    <div style="font-size:.78rem;color:#64748b;font-weight:bold;margin:14px 0 6px">🧪 כימיה ויריבות <span style="font-weight:normal;color:#475569">(${PAIR_MIN}+ משחקים)</span></div>
    <div class="grid4" style="gap:6px;margin-bottom:6px">
      ${chemTile('חבר מנצח','🤝',bestMate,o=>`${pct(o.w,o.t)}% · ${o.t}מ׳`,'#10b981')}
      ${chemTile('חבר ביש מזל','😬',worstMate,o=>`${pct(o.w,o.t)}% · ${o.t}מ׳`,'#ef4444')}
      ${chemTile('נמסיס (5 ש׳)','😈',nemesis,o=>`${o.w}-${o.l} · ${pct(o.w,o.t)}%`,'#ef4444')}
      ${chemTile('קורבן אהוב (5 ש׳)','🎯',favVictim,o=>`${o.w}-${o.l} · ${pct(o.w,o.t)}%`,'#10b981')}
    </div>` : '';

  // ── Recent form heatmap (last 20, newest first — Sofascore style) ──
  const _log = playerGameLog(name);
  const recent = _log.slice(-20).reverse();
  const rW=recent.filter(x=>x.res==='W').length, rD=recent.filter(x=>x.res==='D').length, rL=recent.filter(x=>x.res==='L').length;
  const cellCol = {W:'#10b981', D:'#64748b', L:'#ef4444'};
  const cellLbl = {W:'נ', D:'ת', L:'ה'};
  const formCells = recent.map(x => {
    const extra = (x.g?` · ${x.g}⚽`:'') + (x.a?` · ${x.a}🅰️`:'');
    return `<div title="${fmtD(x.date)} · ${x.my}:${x.opp}${extra}"
      style="width:24px;height:24px;border-radius:5px;background:${cellCol[x.res]};
      display:flex;align-items:center;justify-content:center;color:#fff;font-size:.72rem;font-weight:bold;flex:0 0 auto">${cellLbl[x.res]}</div>`;
  }).join('');
  const formHtml = recent.length ? `
    <div class="card" style="padding:10px;margin-bottom:12px">
      <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:6px;margin-bottom:8px">
        <span style="font-size:.78rem;color:#64748b;font-weight:bold">🔥 פורם — ${recent.length} משחקים אחרונים</span>
        <span style="font-size:.72rem"><b style="color:#10b981">${rW}</b> נצח׳ · <b style="color:#94a3b8">${rD}</b> תיקו · <b style="color:#ef4444">${rL}</b> הפס׳</span>
      </div>
      <div style="display:flex;gap:4px;flex-wrap:wrap;direction:rtl">${formCells}</div>
    </div>` : '';

  // ── Scouting percentiles (FBref style) — where this player ranks in the pool ──
  const scoutRows = [
    {l:'אחוז ניצחון',   v:pct(p.w,p.gm)+'%', pr:pctRank('winpct', p.gm?p.w/p.gm:0)},
    {l:'שערים למשחק',   v:r2(p.g/p.gm),      pr:pctRank('gpg',    p.gm?p.g/p.gm:0)},
    {l:'בישולים למשחק', v:r2(p.a/p.gm),      pr:pctRank('apg',    p.gm?p.a/p.gm:0)},
    {l:'תרומה למשחק',   v:r2((p.g+p.a)/p.gm),pr:pctRank('cpg',    p.gm?(p.g+p.a)/p.gm:0)},
    {l:'ותק (משחקים)',  v:p.gm,              pr:pctRank('gm',     p.gm)},
    {l:'MVP',           v:b.mvp||0,          pr:pctRank('mvp',    b.mvp||0)},
  ];
  const prColor = pr => pr>=80?'#10b981':pr>=50?'#3b82f6':pr>=25?'#f59e0b':'#ef4444';
  const scoutHtml = `
    <div class="card" style="padding:10px;margin-bottom:12px">
      <div style="font-size:.78rem;color:#64748b;font-weight:bold;margin-bottom:8px">📊 דירוג מול הקבוצה <span style="font-weight:normal;color:#475569">(אחוזון מבין ${PCT_POOL.length} השחקנים)</span></div>
      ${scoutRows.map(r => `
        <div style="display:flex;align-items:center;gap:8px;margin:6px 0">
          <div style="width:88px;font-size:.72rem;color:#94a3b8;text-align:right">${r.l}</div>
          <div style="flex:1;height:14px;background:#0f172a;border-radius:4px;overflow:hidden">
            <div style="width:${r.pr}%;height:100%;background:${prColor(r.pr)};border-radius:4px"></div>
          </div>
          <div style="width:30px;font-size:.68rem;color:#64748b;text-align:left">${r.v}</div>
          <div style="width:38px;font-size:.74rem;font-weight:bold;color:${prColor(r.pr)};text-align:left">${r.pr}%</div>
        </div>`).join('')}
    </div>`;

  const skBadge = sk&&sk.type ? (() => {
    const cls=sk.type==='W'?'sk-w':sk.type==='L'?'sk-l':'sk-d';
    const lbl=sk.type==='W'?'ניצחונות ברצף':sk.type==='L'?'הפסדים ברצף':'תיקו';
    const best=sk.type==='W'&&sk.bestW>sk.count?` | שיא: ${sk.bestW}`:'';
    return `<span class="sk ${cls}" style="font-size:.78rem;padding:3px 9px">${sk.count} ${lbl}${best}</span>`;
  })() : '';

  const sBox = (label, val, color) =>
    `<div class="stat-box"><div class="sv" style="color:${color}">${val}</div><div class="sl">${label}</div></div>`;

  const html = `
    <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:14px;padding-bottom:12px;border-bottom:1px solid #334155">
      <h2 style="color:#fbbf24;font-size:1.35rem;margin:0">${name}</h2>
      <span style="background:#0f172a;color:#94a3b8;padding:3px 10px;border-radius:12px;font-size:.78rem">${p.gm} משחקים</span>
      ${skBadge}
    </div>
    <div class="pstats" style="margin-bottom:14px">
      ${sBox('ניצחונות', p.w, '#10b981')}
      ${sBox('הפסדים',   p.l, '#ef4444')}
      ${sBox('תיקו',     p.d, '#fbbf24')}
      ${sBox('%ניצחון', pct(p.w,p.gm)+'%', '#3b82f6')}
      ${sBox('שערים',   p.g, '#fbbf24')}
      ${sBox('בישולים', p.a, '#8b5cf6')}
      ${sBox('MVP',   b.mvp||'-', '#f59e0b')}
      ${sBox("שניצ'",  b.wg||'-',  '#10b981')}
      ${sBox('שיא רצף ניצחונות', (sk&&sk.bestW) ? sk.bestW+`<div style="font-size:.58rem;color:#64748b;font-weight:normal">${streakRange(sk.bestWFrom, sk.bestWTo)}</div>` : '-', '#fb923c')}
      ${sBox('שיא ללא הפסד',     (sk&&sk.bestU) ? sk.bestU+`<div style="font-size:.58rem;color:#64748b;font-weight:normal">${streakRange(sk.bestUFrom, sk.bestUTo)}</div>` : '-', '#22d3ee')}
      ${sBox('שיא ללא ניצחון',   (sk&&sk.bestNW) ? sk.bestNW+`<div style="font-size:.58rem;color:#64748b;font-weight:normal">${streakRange(sk.bestNWFrom, sk.bestNWTo)}</div>` : '-', '#94a3b8')}
    </div>
    ${formHtml}
    ${scoutHtml}
    ${yrs.length>1 ? `
    <div class="card" style="padding:10px;margin-bottom:12px">
      <div style="font-size:.76rem;color:#64748b;margin-bottom:4px">📈 מגמה לפי שנה</div>
      <div class="modal-chart"><canvas id="${chartId}"></canvas></div>
    </div>` : ''}
    ${chemHtml}
    <div class="grid2" style="gap:10px">
      <div>
        <div style="font-size:.78rem;color:#64748b;font-weight:bold;margin-bottom:6px">🤝 שותפים הכי נפוצים</div>
        ${myPairs.length ? myPairs.map(x=>`
          <div class="pair-row">
            <span>${pl(x.other)}</span>
            <span>
              <span class="chip chip-blue">${x.t}מ׳</span>
              <span class="chip chip-green">${x.w}W</span>
              <span class="chip chip-orange">${pct(x.w,x.t)}%</span>
            </span>
          </div>`).join('') : '<div style="color:#475569;font-size:.75rem;padding:6px">אין נתונים</div>'}
      </div>
      <div>
        <div style="font-size:.78rem;color:#64748b;font-weight:bold;margin-bottom:6px">⚔️ יריבים מובילים</div>
        ${myRivals.length ? myRivals.map(x=>`
          <div class="pair-row">
            <span>${pl(x.other)}</span>
            <span>
              <span class="chip chip-blue">${x.t}מ׳</span>
              <span class="chip ${x.w>=x.l?'chip-green':'chip-red'}">${x.w}W / ${x.l}L</span>
            </span>
          </div>`).join('') : '<div style="color:#475569;font-size:.75rem;padding:6px">אין נתונים</div>'}
      </div>
    </div>`;
  return {html, hasChart: yrs.length>1, yrs};
}

function makeTrendChart(chartId, yrs) {
  const ctx = document.getElementById(chartId);
  if (!ctx) return null;
  const winPct = yrs.map(y=>y.gm>0?Math.round(100*y.w/y.gm):0);
  return new Chart(ctx, {
    data:{
      labels: yrs.map(y=>y.yr),
      datasets:[
        // games per year — subtle bars (context, "how many games")
        {type:'bar', label:'משחקים', data:yrs.map(y=>y.gm),
          backgroundColor:'rgba(59,130,246,.35)', borderRadius:3,
          yAxisID:'y2', order:2},
        // win% — bold line on top (the main story)
        {type:'line', label:'% ניצחון', data:winPct,
          borderColor:'#fbbf24', backgroundColor:'rgba(251,191,36,.10)',
          borderWidth:3, tension:.3, yAxisID:'y', fill:true,
          pointRadius:4, pointBackgroundColor:'#fbbf24', order:1},
      ]
    },
    options:{
      plugins:{
        legend:{labels:{color:'#94a3b8',font:{size:10},usePointStyle:true}},
        tooltip:{callbacks:{label:c=> c.dataset.label==='% ניצחון'
          ? `% ניצחון: ${c.raw}%` : `משחקים: ${c.raw}`}}
      },
      scales:{
        y:{max:100,min:0,position:'right',grid:{color:'#334155'},
           ticks:{color:'#fbbf24',font:{size:9},callback:v=>v+'%'},
           title:{display:true,text:'% ניצחון',color:'#fbbf24',font:{size:9}}},
        y2:{min:0,position:'left',grid:{display:false},
           ticks:{color:'#60a5fa',font:{size:9}},
           title:{display:true,text:'משחקים',color:'#60a5fa',font:{size:9}}},
        x:{grid:{color:'#334155'},ticks:{color:'#94a3b8',font:{size:9}}}
      },
      maintainAspectRatio:false
    }
  });
}

function openProfile(name) {
  const prof = profileBody(name, 'mdlChart');
  if (!prof) return;
  document.getElementById('modalContent').innerHTML = prof.html;
  document.getElementById('playerModal').classList.add('open');
  document.body.style.overflow = 'hidden';
  if (prof.hasChart) {
    if (_modalChart) { _modalChart.destroy(); _modalChart=null; }
    setTimeout(() => { _modalChart = makeTrendChart('mdlChart', prof.yrs); }, 50);
  }
}

// ── My Page ───────────────────────────────────────────────────────────────────
let _meBuilt = false;
function buildMe() {
  if (!_meBuilt) {
    const names = [...STATS.players].sort((a,b)=>a.name.localeCompare(b.name,'he')).map(p=>p.name);
    document.getElementById('meSelect').innerHTML =
      ['<option value="">-- בחר שחקן --</option>', ...names.map(n=>`<option value="${n}">${n}</option>`)].join('');
    _meBuilt = true;
  }
  const sel = document.getElementById('meSelect');
  const visitor = localStorage.getItem('soccer_visitor');
  if (!sel.value && visitor && STATS.players.some(p=>p.name===visitor)) sel.value = visitor;
  renderMe();
}
function renderMe() {
  const name = document.getElementById('meSelect').value;
  const el = document.getElementById('meContent');
  if (!name) {
    el.innerHTML = '<div class="card" style="color:#64748b">בחר שחקן למעלה — או לחץ על כפתור 👤 בפינה כדי שנזכור אותך.</div>';
    return;
  }
  const prof = profileBody(name, 'meChart');
  if (!prof) { el.innerHTML = '<div class="card" style="color:#64748b">לא נמצאו נתונים</div>'; return; }
  el.innerHTML = `<div class="card">${prof.html}</div>`;
  if (prof.hasChart) {
    if (_meChart) { _meChart.destroy(); _meChart=null; }
    setTimeout(() => { _meChart = makeTrendChart('meChart', prof.yrs); }, 50);
  }
}
function closeModal() {
  document.getElementById('playerModal').classList.remove('open');
  document.body.style.overflow = '';
}
function modalBgClick(e) {
  if (e.target===document.getElementById('playerModal')) closeModal();
}
document.addEventListener('keydown', e => { if(e.key==='Escape') closeModal(); });

// ── Visitor tracking ──────────────────────────────────────────────────────────
function initVisitor() {
  // ── One-time visitor reset: force these names to re-select ──
  // To evict someone again later: change the names AND bump VISITOR_RESET.
  const VISITOR_RESET = 'r1';
  const VISITOR_RESET_NAMES = ['מיקי'];
  if (localStorage.getItem('soccer_visitor_reset') !== VISITOR_RESET) {
    const cur = localStorage.getItem('soccer_visitor');
    if (cur && VISITOR_RESET_NAMES.includes(cur)) localStorage.removeItem('soccer_visitor');
    localStorage.setItem('soccer_visitor_reset', VISITOR_RESET);
  }
  const names = [...STATS.players].sort((a,b)=>a.name.localeCompare(b.name,'he')).map(p=>p.name);
  const sel   = document.getElementById('visSelect');
  sel.innerHTML = ['<option value="">-- בחר --</option>',
    '<option value="אורח">אורח/ת</option>',
    ...names.map(n=>`<option value="${n}">${n}</option>`)].join('');
  const stored = localStorage.getItem('soccer_visitor');
  if (stored) {
    document.getElementById('visLabel').textContent = stored;
    sel.value = stored;
    logVisit(stored);
  } else {
    // first visit: block the dashboard until a name is chosen
    document.getElementById('gateSelect').innerHTML = sel.innerHTML;
    document.getElementById('gateOverlay').style.display = 'block';
    document.body.style.overflow = 'hidden';
  }
}
function enterGate() {
  const v = document.getElementById('gateSelect').value;
  if (!v) return;
  localStorage.setItem('soccer_visitor', v);
  document.getElementById('visLabel').textContent = v;
  document.getElementById('visSelect').value = v;
  document.getElementById('gateOverlay').style.display = 'none';
  document.body.style.overflow = '';
  logVisit(v);
}
function saveVisitor() {
  const v = document.getElementById('visSelect').value;
  if (!v) return;
  localStorage.setItem('soccer_visitor', v);
  document.getElementById('visLabel').textContent = v;
  toggleVisPanel();
  logVisit(v);
}
function logVisit(name) {
  if (!WEBHOOK_URL) return;
  const ua = encodeURIComponent(navigator.userAgent.slice(0,80));
  fetch(`${WEBHOOK_URL}?v=${encodeURIComponent(name)}&ua=${ua}`, {mode:'no-cors'}).catch(()=>{});
}
function toggleVisPanel() {
  document.getElementById('visPanel').classList.toggle('open');
}
document.addEventListener('click', e => {
  const panel = document.getElementById('visPanel');
  if (panel.classList.contains('open')
      && !panel.contains(e.target)
      && !document.getElementById('visBtn').contains(e.target)) {
    panel.classList.remove('open');
  }
});

// ── Init ──────────────────────────────────────────────────────────────────────
buildOverview();
buildLeaderboard();
buildMVPLeaderboard();
buildKosher();
buildLastGame();
calcH2H();
initVisitor();
</script>
</body>
</html>
'''

if __name__ == '__main__':
    generate()
