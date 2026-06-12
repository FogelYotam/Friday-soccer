import json, openpyxl
from collections import defaultdict

import os as _os
_DIR       = _os.path.dirname(_os.path.abspath(__file__))
GAMES_PATH = _os.path.join(_DIR, 'games_data.json')
EXCEL_PATH = r'C:\Users\shlom\OneDrive\Documents\כדורגל שישי\כדורגל שישי בילו 2026.xlsx'
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

MIN_GAMES_THRESHOLD = 10  # players with fewer games are treated as guests and excluded

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
  <button onclick="showTab('kosher',this)">💪 כושר</button>
  <button onclick="showTab('h2h',this)">⚔️ ראש בראש</button>
</nav>

<!-- OVERVIEW -->
<div id="tab-overview" class="tab active">
  <div class="card" style="margin-bottom:12px">
    <h3 id="curYearTitle">🗓️ שנה נוכחית</h3>
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
        <select id="minGames" onchange="filterTable()">
          <option value="1">כולם</option>
          <option value="10">מינ׳ 10</option>
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

<!-- PLAYER PROFILE MODAL -->
<div id="playerModal" class="modal-ov" onclick="modalBgClick(event)">
  <div class="modal-box">
    <button class="modal-close" onclick="closeModal()">✕ סגור</button>
    <div id="modalContent"></div>
  </div>
</div>

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
const WEBHOOK_URL = '';  // e.g. 'https://script.google.com/macros/s/ABC.../exec'
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
const KSR_YEARS = ALL_YRS.slice(-3);
const KSR_WEIGHTS = {};
KSR_YEARS.forEach((y,i) => KSR_WEIGHTS[y] = Math.pow(2, i)); // 1,2,4

// ── Date parsing (DD/MM/YYYY Israeli format) ─────────────────────────────────
function parseDate(d) {
  if (!d) return 0;
  const p = d.replace(/-/g,'/').split('/');
  if (p.length === 3) {
    const [a,b,c] = [+p[0],+p[1],+p[2]];
    return c > 1000 ? new Date(c,b-1,a).getTime() : new Date(a,b-1,c).getTime();
  }
  return new Date(d).getTime();
}

// ── Streak computation (all players, sorted by date) ─────────────────────────
const STREAKS = (() => {
  const sorted = [...GAMES].sort((a,b) => parseDate(a.date) - parseDate(b.date));
  const s = {};
  sorted.forEach(g => {
    const sA = g.scoreA??0, sB = g.scoreB??0;
    const rA = sA>sB?'W':sA<sB?'L':'D';
    const rB = sB>sA?'W':sB<sA?'L':'D';
    const upd = (name, r) => {
      if (!s[name]) s[name] = {type:null, count:0, bestW:0};
      const c = s[name];
      if (c.type === r) { c.count++; if (r==='W' && c.count>c.bestW) c.bestW=c.count; }
      else              { c.type=r; c.count=1; if (r==='W') c.bestW=Math.max(c.bestW,1); }
    };
    (g.teamA||[]).forEach(p => { if(p.name) upd(p.name, rA); });
    (g.teamB||[]).forEach(p => { if(p.name) upd(p.name, rB); });
  });
  return s;
})();

// ── Utils ────────────────────────────────────────────────────────────────────
function pct(w,t)  { return t>0 ? Math.round(100*w/t) : 0; }
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

  // current year
  document.getElementById('curYearTitle').textContent = `🗓️ הישגי ${curYr} — שנה נוכחית`;
  const cyrData   = STATS.byYear.filter(e => e.yr===curYr);
  const cyrScorers = [...cyrData].sort((a,b)=>b.g-a.g).slice(0,3);
  const cyrAssists = [...cyrData].sort((a,b)=>b.a-a.a).slice(0,3);
  const cyrMVP = [...(STATS.bonusByYear||[])].filter(b=>b.yr===curYr&&b.mvp>0).sort((a,b)=>b.mvp-a.mvp).slice(0,3);
  const cyrWG  = [...(STATS.bonusByYear||[])].filter(b=>b.yr===curYr&&b.wg>0).sort((a,b)=>b.wg-a.wg).slice(0,3);
  const cyrRow = (arr, fmt) => arr.length ? arr.map((p,i)=>`
    <div style="display:flex;justify-content:space-between;font-size:.76rem;margin-bottom:3px">
      <span>${i===0?'🥇':i===1?'🥈':'🥉'} ${pl(p.name)}</span>
      <b style="color:#fbbf24">${fmt(p)}</b>
    </div>`).join('') : '<div style="color:#475569;font-size:.75rem">אין נתונים</div>';

  document.getElementById('curYearHeroes').innerHTML = [
    {icon:'⚽', label:`מלך שערים ${curYr}`, content:cyrRow(cyrScorers,p=>`${p.g} שע׳`)},
    {icon:'🅰️', label:`מלך בישולים ${curYr}`, content:cyrRow(cyrAssists,p=>`${p.a} בישול`)},
    {icon:'🏅', label:`MVP ${curYr}`, content:cyrMVP.length?cyrRow(cyrMVP,p=>`${p.mvp} MVP`):'<div style="color:#475569;font-size:.75rem">טרם נקבע</div>'},
    {icon:'⚡', label:`שניצ׳ ${curYr}`, content:cyrWG.length?cyrRow(cyrWG,p=>`${p.wg} שניצ׳`):'<div style="color:#475569;font-size:.75rem">טרם נקבע</div>'},
  ].map(h=>`
    <div class="hero-card">
      <div style="display:flex;align-items:center;gap:6px;margin-bottom:8px">
        <span style="font-size:1.5rem">${h.icon}</span>
        <div style="font-size:.72rem;color:#64748b;font-weight:bold">${h.label}</div>
      </div>
      ${h.content}
    </div>`).join('');

  document.getElementById('heroes').innerHTML = [
    {icon:'🏆', label:'מובילי נקודות',       content:top3([...active],'pts',v=>v+' נק\'')},
    {icon:'⚽', label:'מלכי השערים',          content:top3([...active],'g',v=>v+' ש\'')},
    {icon:'🎯', label:'אחוז ניצחון',          content:[...active].sort((a,b)=>(b.w/b.gm)-(a.w/a.gm)).slice(0,3).map((p,i)=>`
      <div style="display:flex;justify-content:space-between;font-size:.76rem;margin-bottom:3px">
        <span>${i===0?'🥇':i===1?'🥈':'🥉'} ${pl(p.name)}</span>
        <b style="color:#fbbf24">${pct(p.w,p.gm)}%</b>
      </div>`).join('')},
    {icon:'📈', label:'ממוצע נקודות / משחק',  content:top3([...active],'ppg',v=>r2(v))},
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
  filterTable();
}
function filterTable() {
  const q   = (document.getElementById('searchPlayer').value||'').trim();
  const min = parseInt(document.getElementById('minGames').value)||1;
  const rows = STATS.players
    .filter(p => p.gm>=min && (q===''||p.name.includes(q)))
    .map(p => ({...p,
      pts:    (p.w||0)*2+(p.d||0),
      ppg:    p.gm>0?((p.w||0)*2+(p.d||0))/p.gm:0,
      winpct: p.gm>0?p.w/p.gm:0,
      gpg:    p.gm>0?p.g/p.gm:0,
      apg:    p.gm>0?p.a/p.gm:0,
      contrib:p.gm>0?(p.g+p.a)/p.gm:0,
      mvp_n:  (BONUS[p.name]||{}).mvp||0,
      wg_n:   (BONUS[p.name]||{}).wg||0,
    }));
  rows.sort((a,b) => {
    if (lbSortKey==='w') {
      let d=b.w-a.w; if(d!==0) return lbSortDir*d;
      d=b.pts-a.pts; if(d!==0) return lbSortDir*d;
      return lbSortDir*(b.g-a.g);
    }
    return lbSortDir*(b[lbSortKey]-a[lbSortKey]);
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

// ── Kosher ────────────────────────────────────────────────────────────────────
let ksrSortKey='rating', ksrSortDir=-1, _ksrRows=[];
function yrScore(yd, bonus) {
  if (!yd || yd.gm<1) return null;
  const b = bonus||{mvp:0,wg:0};
  return (yd.w/yd.gm)*40 + (yd.g/yd.gm)*20 + (yd.a/yd.gm)*10 + (b.mvp/yd.gm)*30 + (b.wg/yd.gm)*20;
}
function buildKosher() {
  const min = parseInt(document.getElementById('kosherMinGames').value)||5;
  const ksrRange = KSR_YEARS.length>=2
    ? `${KSR_YEARS[0].slice(-2)}-${KSR_YEARS[KSR_YEARS.length-1].slice(-2)}`
    : KSR_YEARS[0]||'';
  document.getElementById('kosherFormula').textContent =
    `ניקוד = ממוצע משוקלל שנים ${KSR_YEARS.join('/')} | משקלות: ${KSR_YEARS.map((y,i)=>`${y}×${KSR_WEIGHTS[y]}`).join(' · ')} | נוסחה: %נצח×40 + ש/מ×20 + ב/מ×10 + MVP/מ×30 + שניצ׳/מ×20`;

  const byYearMap = {};
  STATS.byYear.forEach(e => {
    if (!byYearMap[e.name]) byYearMap[e.name] = {};
    byYearMap[e.name][e.yr] = e;
  });

  _ksrRows = STATS.players
    .filter(p => {
      const yrs = byYearMap[p.name]||{};
      return KSR_YEARS.reduce((s,y) => s+(yrs[y]?yrs[y].gm:0), 0) >= min;
    })
    .map(p => {
      const yrs = byYearMap[p.name]||{};
      const bonusYrs = BONUS_BY_YEAR[p.name]||{};
      let wSum=0,wTot=0,rGm=0,rW=0,rG=0,rA=0,rMvp=0,rWg=0;
      KSR_YEARS.forEach(y => {
        const yd=yrs[y]; if(!yd||yd.gm<1) return;
        const bns=bonusYrs[y]||{mvp:0,wg:0};
        const sc=yrScore(yd,bns), wt=KSR_WEIGHTS[y]||1;
        wSum+=sc*wt; wTot+=wt;
        rGm+=yd.gm; rW+=yd.w; rG+=yd.g; rA+=yd.a;
        rMvp+=(bns.mvp||0); rWg+=(bns.wg||0);
      });
      const rating = wTot>0 ? wSum/wTot : 0;
      const lastY=KSR_YEARS[KSR_YEARS.length-1], prevY=KSR_YEARS[KSR_YEARS.length-2];
      const sc2=yrScore(yrs[lastY],bonusYrs[lastY]), sc1=yrScore(yrs[prevY],bonusYrs[prevY]);
      let trend='same';
      if (sc2!==null&&sc1!==null) { if(sc2>sc1*1.05) trend='up'; else if(sc2<sc1*0.95) trend='down'; }
      else if (sc2!==null&&sc1===null) trend='up';
      else if (sc2===null&&sc1!==null) trend='down';
      const gmByYr = {};
      KSR_YEARS.forEach(y => gmByYr['yr'+y] = yrs[y]?yrs[y].gm:0);
      return {...p, rating, trend, recentGm:rGm, recentW:rW,
        wp:rGm>0?rW/rGm:0, gp:rGm>0?rG/rGm:0, ap:rGm>0?rA/rGm:0,
        mvp_n:rMvp, wg_n:rWg, ...gmByYr};
    });
  renderKosherTable();
}
function sortKosher(key) {
  if (ksrSortKey===key) ksrSortDir*=-1; else { ksrSortKey=key; ksrSortDir=-1; }
  renderKosherTable();
}
function renderKosherTable() {
  const rows = [..._ksrRows].sort((a,b) => {
    const va=a[ksrSortKey]??0, vb=b[ksrSortKey]??0;
    if (typeof va==='string') return ksrSortDir*(va<vb?-1:va>vb?1:0);
    return ksrSortDir*(vb-va);
  });
  const maxR = Math.max(...rows.map(r=>r.rating), 0.01);
  const ti   = t => t==='up'?'<span class="trend-up">↑</span>':t==='down'?'<span class="trend-down">↓</span>':'<span class="trend-same">→</span>';
  const thK  = (key,label) => `<th onclick="sortKosher('${key}')" style="${ksrSortKey===key?'color:#fff':''}">${label}${ksrSortKey===key?(ksrSortDir===-1?' ▼':' ▲'):''}</th>`;
  document.getElementById('kosherTable').innerHTML = `
    <table><thead><tr>
      <th>#</th>
      <th onclick="sortKosher('name')" style="text-align:right;${ksrSortKey==='name'?'color:#fff':''}">שחקן${ksrSortKey==='name'?(ksrSortDir===-1?' ▼':' ▲'):''}</th>
      ${thK('recentGm', 'מ׳ '+KSR_YEARS.map(y=>y.slice(-2)).join('-'))}
      ${KSR_YEARS.map(y=>thK('yr'+y, y)).join('')}
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
        ${KSR_YEARS.map(y=>`<td style="color:#475569">${p['yr'+y]||'-'}</td>`).join('')}
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

// ── Player Profile Modal ──────────────────────────────────────────────────────
let _modalChart = null;
function openProfile(name) {
  const p = STATS.players.find(x => x.name===name);
  if (!p) return;
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

  const skBadge = sk&&sk.type ? (() => {
    const cls=sk.type==='W'?'sk-w':sk.type==='L'?'sk-l':'sk-d';
    const lbl=sk.type==='W'?'ניצחונות ברצף':sk.type==='L'?'הפסדים ברצף':'תיקו';
    const best=sk.type==='W'&&sk.bestW>sk.count?` | שיא: ${sk.bestW}`:'';
    return `<span class="sk ${cls}" style="font-size:.78rem;padding:3px 9px">${sk.count} ${lbl}${best}</span>`;
  })() : '';

  const sBox = (label, val, color) =>
    `<div class="stat-box"><div class="sv" style="color:${color}">${val}</div><div class="sl">${label}</div></div>`;

  const chartId = 'mdlChart';
  document.getElementById('modalContent').innerHTML = `
    <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:14px;padding-bottom:12px;border-bottom:1px solid #334155">
      <h2 style="color:#fbbf24;font-size:1.35rem;margin:0">${name}</h2>
      <span style="background:#0f172a;color:#94a3b8;padding:3px 10px;border-radius:12px;font-size:.78rem">${p.gm} משחקים</span>
      ${skBadge}
    </div>
    <div class="grid4" style="gap:6px;margin-bottom:14px">
      ${sBox('ניצחונות', p.w, '#10b981')}
      ${sBox('הפסדים',   p.l, '#ef4444')}
      ${sBox('תיקו',     p.d, '#fbbf24')}
      ${sBox('%ניצחון', pct(p.w,p.gm)+'%', '#3b82f6')}
      ${sBox('שערים',   p.g, '#fbbf24')}
      ${sBox('בישולים', p.a, '#8b5cf6')}
      ${sBox('MVP',   b.mvp||'-', '#f59e0b')}
      ${sBox("שניצ'",  b.wg||'-',  '#10b981')}
    </div>
    ${yrs.length>1 ? `
    <div class="card" style="padding:10px;margin-bottom:12px">
      <div style="font-size:.76rem;color:#64748b;margin-bottom:4px">📈 מגמה לפי שנה</div>
      <div class="modal-chart"><canvas id="${chartId}"></canvas></div>
    </div>` : ''}
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

  document.getElementById('playerModal').classList.add('open');
  document.body.style.overflow = 'hidden';

  if (yrs.length > 1) {
    if (_modalChart) { _modalChart.destroy(); _modalChart=null; }
    setTimeout(() => {
      const ctx = document.getElementById(chartId);
      if (!ctx) return;
      _modalChart = new Chart(ctx, {
        type:'line',
        data:{
          labels: yrs.map(y=>y.yr),
          datasets:[
            {label:'%ניצ׳', data:yrs.map(y=>y.gm>0?Math.round(100*y.w/y.gm):0),
              borderColor:'#fbbf24', backgroundColor:'rgba(251,191,36,.12)',
              tension:.3, yAxisID:'y', fill:true, pointRadius:4},
            {label:'משחקים', data:yrs.map(y=>y.gm),
              borderColor:'#3b82f6', backgroundColor:'rgba(59,130,246,.08)',
              tension:.3, yAxisID:'y2', pointRadius:3},
          ]
        },
        options:{
          plugins:{legend:{labels:{color:'#94a3b8',font:{size:10}}}},
          scales:{
            y:{max:100,min:0,grid:{color:'#334155'},ticks:{color:'#94a3b8',font:{size:9}},
               title:{display:true,text:"% ניצ'",color:'#94a3b8',font:{size:9}}},
            y2:{position:'left',grid:{display:false},ticks:{color:'#475569',font:{size:9}},
               title:{display:true,text:"מ'",color:'#475569',font:{size:9}}},
            x:{grid:{color:'#334155'},ticks:{color:'#94a3b8',font:{size:9}}}
          },
          maintainAspectRatio:false
        }
      });
    }, 50);
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
  }
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
