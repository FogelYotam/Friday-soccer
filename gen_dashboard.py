import json, openpyxl
from collections import defaultdict

GAMES_PATH = r'C:\Users\shlom\OneDrive\Documents\SOCCER\games_data.json'
EXCEL_PATH = r'C:\Users\shlom\OneDrive\Documents\כדורגל שישי\כדורגל שישי בילו 2026.xlsx'
OUT_PATH   = r'C:\Users\shlom\OneDrive\Documents\SOCCER\dashboard.html'

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
    """Read MVP and winning-goal counts per player per year from Excel annual sheets (2017+).
    Returns (by_year_list, totals_dict):
      by_year_list: [{name, yr, mvp, wg}, ...]
      totals_dict:  {name: {mvp, wg}}
    """
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
    pair_stats    = defaultdict(lambda: {'w':0,'l':0,'t':0,'g':0,'a':0})   # same team
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

        # Partners (same team) — track wins/losses and combined goals+assists
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

        # Rivals / H2H (opposite teams)
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
    """Read MVP and winning-goal from games_data.json entries."""
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

    # Normalize player names in raw games
    for g in games:
        g['teamA'] = [p for p in g.get('teamA', []) if normalize(p.get('name',''))]
        g['teamB'] = [p for p in g.get('teamB', []) if normalize(p.get('name',''))]
        for p in g['teamA']: p['name'] = normalize(p['name'])
        for p in g['teamB']: p['name'] = normalize(p['name'])

    stats = compute_stats(games)

    # Exclude players with fewer than MIN_GAMES_THRESHOLD total games (guests)
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
.hero-icon{font-size:1.6rem}
.vs-wrap{display:flex;align-items:center;justify-content:space-around;padding:14px;background:#0f172a;border-radius:8px;margin-top:8px}
.vs-player{text-align:center;min-width:110px}
.vs-name{font-size:1rem;font-weight:bold;color:#fbbf24;margin-bottom:4px}
.vs-wins{font-size:2rem;font-weight:bold}
.vs-label{font-size:.7rem;color:#64748b;margin-top:2px}
.vs-mid{font-size:1.2rem;color:#475569;text-align:center}
.fun-fact{background:#0f172a;border-right:3px solid #fbbf24;padding:8px 12px;margin:5px 0;border-radius:4px}
.fun-fact .ff-label{font-size:.72rem;color:#64748b}
.fun-fact .ff-val{font-size:.9rem;color:#fbbf24;font-weight:bold;margin-top:2px}
.chip{display:inline-block;padding:2px 7px;border-radius:12px;font-size:.68rem;margin:1px}
.chip-green{background:#064e3b;color:#6ee7b7}
.chip-red{background:#7f1d1d;color:#fca5a5}
.chip-blue{background:#1e3a8a;color:#93c5fd}
.chip-orange{background:#78350f;color:#fcd34d}
.chart-wrap{position:relative;height:220px}
.pair-row{display:flex;justify-content:space-between;align-items:center;padding:5px 0;border-bottom:1px solid #334155;font-size:.8rem}
.pair-names{color:#e2e8f0;font-weight:bold;flex:1;padding-left:8px}
.h2h-stats{margin-top:8px;display:grid;grid-template-columns:1fr 1fr;gap:8px;font-size:.8rem}
.h2h-stat-val{font-size:1.1rem;font-weight:bold;color:#fbbf24}
.h2h-stat-label{font-size:.68rem;color:#64748b}
.trend-up{color:#10b981;font-size:1.1rem;font-weight:bold}
.trend-down{color:#ef4444;font-size:1.1rem;font-weight:bold}
.trend-same{color:#fbbf24;font-size:1.1rem}
@media(max-width:900px){.grid2,.grid3,.grid4{grid-template-columns:1fr}}
</style>
</head>
<body>
<h1>⚽ כדורגל שישי</h1>
<div class="subtitle" id="mainSubtitle">טוען נתונים...</div>
<nav>
  <button class="active" onclick="showTab('overview',this)">🏠 סקירה</button>
  <button onclick="showTab('leaderboard',this)">🏆 מצטיינים</button>
  <button onclick="showTab('kosher',this)">💪 כושר</button>
  <button onclick="showTab('h2h',this)">⚔️ ראש בראש</button>
  <button onclick="showTab('partners',this)">🤝 שותפויות</button>
  <button onclick="showTab('years',this)">📅 לפי שנה</button>
  <button onclick="showTab('facts',this)">✨ עובדות</button>
</nav>

<!-- OVERVIEW TAB -->
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
      <h3>⚽ מבקיעים מובילים (100+ משחקים)</h3>
      <div class="chart-wrap"><canvas id="chartScorers"></canvas></div>
    </div>
  </div>
  <div class="grid2" style="margin-top:12px">
    <div class="card">
      <h3>📈 מובילי נקודות (100+ משחקים)</h3>
      <div class="chart-wrap"><canvas id="chartPoints"></canvas></div>
    </div>
    <div class="card">
      <h3>🎯 אחוזי ניצחון מובילים (100+ משחקים)</h3>
      <div class="chart-wrap"><canvas id="chartWinPct"></canvas></div>
    </div>
  </div>
  <div class="card" style="margin-top:12px">
    <h3>🅰️ מלכי הבישולים (100+ משחקים)</h3>
    <div class="chart-wrap"><canvas id="chartAssists"></canvas></div>
  </div>
</div>

<!-- LEADERBOARD TAB -->
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
        <thead>
          <tr>
            <th onclick="sortTable('name')">#שחקן</th>
            <th onclick="sortTable('gm')">משח׳</th>
            <th onclick="sortTable('w')">נצח׳</th>
            <th onclick="sortTable('winpct')">%נצח</th>
            <th onclick="sortTable('pts')">נק׳</th>
            <th onclick="sortTable('ppg')">נק/מ</th>
            <th onclick="sortTable('g')">שע׳</th>
            <th onclick="sortTable('gpg')">ש/מ</th>
            <th onclick="sortTable('a')">בישול</th>
            <th onclick="sortTable('apg')">ב/מ</th>
            <th onclick="sortTable('contrib')">תרומה/מ</th>
            <th onclick="sortTable('mvp_n')" style="color:#fbbf24">🏅 MVP</th>
            <th onclick="sortTable('wg_n')" style="color:#10b981">⚡ שניצ׳</th>
          </tr>
        </thead>
        <tbody id="lbBody"></tbody>
      </table>
    </div>
  </div>
  <div class="card" style="margin-top:12px">
    <h3>🏅 מלכי MVP כל הזמנים (2017+)</h3>
    <div class="tbl-wrap" id="mvpLeaderboard"></div>
  </div>
</div>

<!-- KOSHER (RATING) TAB -->
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
    <div style="font-size:.72rem;color:#475569;margin-bottom:8px;padding:6px 10px;background:#0f172a;border-radius:6px">
      ניקוד = ממוצע משוקלל לשנים 2024-2026 בלבד &nbsp;|&nbsp;
      משקלות: 2024×1 · 2025×2 · 2026×4 &nbsp;|&nbsp;
      נוסחת שנה: %נצח×40 + ש/מ×20 + ב/מ×10 + MVP/מ×30 + שניצ׳/מ×20
    </div>
    <div class="tbl-wrap" id="kosherTable"></div>
  </div>
</div>

<!-- H2H TAB -->
<div id="tab-h2h" class="tab">
  <div class="card">
    <h3>⚔️ השוואה ישירה</h3>
    <div style="display:flex;gap:10px;margin-bottom:14px;align-items:center;flex-wrap:wrap">
      <select id="h2hA" onchange="calcH2H()" style="flex:1;min-width:120px"></select>
      <div style="font-weight:bold;color:#fbbf24;font-size:1.1rem">VS</div>
      <select id="h2hB" onchange="calcH2H()" style="flex:1;min-width:120px"></select>
    </div>
    <div id="h2hResult"></div>
  </div>
</div>

<!-- PARTNERS TAB -->
<div id="tab-partners" class="tab">
  <div class="grid2" style="margin-bottom:12px">
    <div class="card">
      <h3>🏆 הזוגות המנצחים ביותר (מינ׳ 15 משחקים)</h3>
      <div id="bestPairsWins"></div>
    </div>
    <div class="card">
      <h3>💔 הזוגות המפסידים ביותר (מינ׳ 15 משחקים)</h3>
      <div id="worstPairsLosses"></div>
    </div>
  </div>
  <div class="card">
    <h3>🔍 שותפויות לפי שחקן</h3>
    <div style="margin-bottom:10px">
      <select id="partnerSearch" onchange="showPartnerStats()" style="min-width:180px"></select>
    </div>
    <div id="partnerResult"></div>
  </div>
</div>

<!-- YEARS TAB -->
<div id="tab-years" class="tab">
  <div class="card">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;flex-wrap:wrap;gap:8px">
      <h3 style="border:none;padding:0;margin:0">📅 סטטיסטיקה לפי שנה</h3>
      <div style="display:flex;gap:8px">
        <select id="yearMinGames" onchange="showYearStats()">
          <option value="1">כולם</option>
          <option value="5">מינ׳ 5</option>
          <option value="10" selected>מינ׳ 10</option>
        </select>
        <select id="yearSelect" onchange="showYearStats()"></select>
      </div>
    </div>
    <div id="yearTable" class="tbl-wrap"></div>
  </div>
</div>

<!-- FACTS TAB -->
<div id="tab-facts" class="tab">
  <div class="grid2">
    <div class="card"><div id="factsLeft"></div></div>
    <div class="card"><div id="factsRight"></div></div>
  </div>
  <div class="card" style="margin-top:12px">
    <h3>🔥 יריבויות הכי נפוצות</h3>
    <div id="rivalsList" class="tbl-wrap"></div>
  </div>
</div>

<script>
const STATS = STATS_PLACEHOLDER;
const GAMES = GAMES_PLACEHOLDER;

// BONUS[name] = {mvp, wg}  (totals)
const BONUS = {};
(STATS.bonus || []).forEach(b => BONUS[b.name] = b);
// BONUS_BY_YEAR[name][yr] = {mvp, wg}
const BONUS_BY_YEAR = {};
(STATS.bonusByYear || []).forEach(b => {
  if(!BONUS_BY_YEAR[b.name]) BONUS_BY_YEAR[b.name]={};
  BONUS_BY_YEAR[b.name][b.yr]={mvp:b.mvp,wg:b.wg};
});
// Year weights for כושר rating (only 2024+)
const KSR_YEARS   = ['2024','2025','2026'];
const KSR_WEIGHTS = {'2024':1,'2025':2,'2026':4};

// ============ UTILS ============
function pct(w,t){ return t>0 ? Math.round(100*w/t) : 0; }
function r2(v){ return isFinite(v)? Math.round(v*100)/100 : 0; }
function showTab(name, btn){
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  document.querySelectorAll('nav button').forEach(b=>b.classList.remove('active'));
  document.getElementById('tab-'+name).classList.add('active');
  if(btn) btn.classList.add('active');
}

// ============ ANALYTICS HELPERS ============
function getTopLineups(limit=3){
  const counts={};
  GAMES.forEach(g=>{
    let winner=null;
    if(g.scoreA>g.scoreB) winner=g.teamA;
    else if(g.scoreB>g.scoreA) winner=g.teamB;
    if(winner){
      const key=winner.map(p=>p.name).sort().join(' + ');
      counts[key]=(counts[key]||0)+1;
    }
  });
  return Object.entries(counts).sort((a,b)=>b[1]-a[1]).slice(0,limit);
}

function getTopPairs(metric='w', limit=3, min=15){
  let list=[...STATS.pairs].filter(p=>p.t>=min);
  if(metric==='pct') list.sort((a,b)=>(b.w/b.t)-(a.w/a.t));
  else if(metric==='l') list.sort((a,b)=>b.l-a.l);
  else list.sort((a,b)=>b.w-a.w);
  return list.slice(0,limit);
}

// ============ OVERVIEW ============
function buildOverview(){
  const ps=STATS.players.map(p=>({...p,
    pts:(p.w||0)*2+(p.d||0),
    ppg:p.gm>0?((p.w||0)*2+(p.d||0))/p.gm:0,
    gpg:p.gm>0?p.g/p.gm:0,
    apg:p.gm>0?p.a/p.gm:0
  }));
  const active=ps.filter(p=>p.gm>=100);
  const totalGames=GAMES.length;
  const years=[...new Set(GAMES.map(g=>g.date?.match(/(\d{4})/)?.[1]).filter(Boolean))].sort();
  const curYr=years[years.length-1];
  document.getElementById('mainSubtitle').textContent=
    `${years[0]}–${curYr} | ${totalGames} משחקים | ${STATS.players.length} שחקנים`;

  const top3html=(arr,key,fmt)=>arr.sort((a,b)=>b[key]-a[key]).slice(0,3).map((p,i)=>`
    <div style="display:flex;justify-content:space-between;font-size:.76rem;margin-bottom:3px">
      <span>${i===0?'🥇':i===1?'🥈':'🥉'} ${p.name}</span>
      <b style="color:#fbbf24">${fmt(p[key])}</b>
    </div>`).join('');

  // ---- Current year heroes ----
  document.getElementById('curYearTitle').textContent=`🗓️ הישגי ${curYr} — שנה נוכחית`;
  const cyrData=STATS.byYear.filter(e=>e.yr===curYr);
  const cyrScorers=[...cyrData].sort((a,b)=>b.g-a.g).slice(0,3);
  const cyrAssists=[...cyrData].sort((a,b)=>b.a-a.a).slice(0,3);
  const cyrWins=[...cyrData].sort((a,b)=>(b.w/Math.max(b.gm,1))-(a.w/Math.max(a.gm,1))).filter(p=>p.gm>=3).slice(0,3);
  const cyrMVP=[...(STATS.bonusByYear||[])].filter(b=>b.yr===curYr&&b.mvp>0).sort((a,b)=>b.mvp-a.mvp).slice(0,3);
  const cyrWG=[...(STATS.bonusByYear||[])].filter(b=>b.yr===curYr&&b.wg>0).sort((a,b)=>b.wg-a.wg).slice(0,3);

  const cyrRow=(arr,fmt)=>arr.length?arr.map((p,i)=>`
    <div style="display:flex;justify-content:space-between;font-size:.76rem;margin-bottom:3px">
      <span>${i===0?'🥇':i===1?'🥈':'🥉'} ${p.name}</span>
      <b style="color:#fbbf24">${fmt(p)}</b>
    </div>`).join(''):'<div style="color:#475569;font-size:.75rem">אין נתונים</div>';

  const cyrHeroes=[
    {icon:'⚽',label:`מלך שערים ${curYr}`,content:cyrRow(cyrScorers,p=>`${p.g} שע׳`)},
    {icon:'🅰️',label:`מלך בישולים ${curYr}`,content:cyrRow(cyrAssists,p=>`${p.a} בישול`)},
    {icon:'🏅',label:`MVP ${curYr}`,content:cyrMVP.length?cyrRow(cyrMVP,p=>`${p.mvp} MVP`):'<div style="color:#475569;font-size:.75rem">טרם נקבע</div>'},
    {icon:'⚡',label:`שניצ׳ ${curYr}`,content:cyrWG.length?cyrRow(cyrWG,p=>`${p.wg} שניצ׳`):'<div style="color:#475569;font-size:.75rem">טרם נקבע</div>'},
  ];
  document.getElementById('curYearHeroes').innerHTML=cyrHeroes.map(h=>`
    <div class="hero-card">
      <div style="display:flex;align-items:center;gap:6px;margin-bottom:8px">
        <span style="font-size:1.5rem">${h.icon}</span>
        <div style="font-size:.72rem;color:#64748b;font-weight:bold">${h.label}</div>
      </div>
      ${h.content}
    </div>`).join('');

  // ---- All-time heroes ----
  const heroes=[
    {icon:'🏆',label:'מובילי נקודות',content:top3html([...active],'pts',v=>v+' נק\'')},
    {icon:'⚽',label:'מלכי השערים',content:top3html([...active],'g',v=>v+' ש\'')},
    {icon:'🎯',label:'אחוז ניצחון',content:[...active].sort((a,b)=>(b.w/b.gm)-(a.w/a.gm)).slice(0,3).map((p,i)=>`
      <div style="display:flex;justify-content:space-between;font-size:.76rem;margin-bottom:3px">
        <span>${i===0?'🥇':i===1?'🥈':'🥉'} ${p.name}</span>
        <b style="color:#fbbf24">${pct(p.w,p.gm)}%</b>
      </div>`).join('')},
    {icon:'📈',label:'ממוצע נקודות / משחק',content:top3html([...active],'ppg',v=>r2(v))}
  ];

  document.getElementById('heroes').innerHTML=heroes.map(h=>`
    <div class="hero-card">
      <div style="display:flex;align-items:center;gap:6px;margin-bottom:8px">
        <span style="font-size:1.5rem">${h.icon}</span>
        <div style="font-size:.72rem;color:#64748b;font-weight:bold">${h.label}</div>
      </div>
      ${h.content}
    </div>`).join('');

  // Charts
  const gByYear={};
  GAMES.forEach(g=>{const m=g.date?.match(/(\d{4})/);if(m)gByYear[m[1]]=(gByYear[m[1]]||0)+1;});
  const yrs=Object.keys(gByYear).sort();
  newChart('chartYears','bar',yrs,yrs.map(y=>gByYear[y]),'משחקים','#3b82f6');

  const top10=[...active].sort((a,b)=>b.g-a.g).slice(0,10);
  newChart('chartScorers','bar',top10.map(p=>p.name),top10.map(p=>p.g),'שערים','#fbbf24');

  const topPts=[...active].sort((a,b)=>b.pts-a.pts).slice(0,10);
  newChart('chartPoints','bar',topPts.map(p=>p.name),topPts.map(p=>p.pts),'נקודות','#10b981');

  const topW=[...active].sort((a,b)=>(b.w/b.gm)-(a.w/a.gm)).slice(0,10);
  newChart('chartWinPct','bar',topW.map(p=>p.name),topW.map(p=>pct(p.w,p.gm)),'% ניצחון','#3b82f6',100);

  const topA=[...active].sort((a,b)=>b.a-a.a).slice(0,10);
  newChart('chartAssists','bar',topA.map(p=>p.name),topA.map(p=>p.a),'בישולים','#8b5cf6');
}

function newChart(id,type,labels,data,label,color,max){
  new Chart(document.getElementById(id),{type,data:{labels,datasets:[{label,data,backgroundColor:color,borderRadius:4}]},
    options:{plugins:{legend:{display:false}},scales:{
      y:{max,grid:{color:'#334155'},ticks:{color:'#94a3b8'}},
      x:{grid:{color:'#334155'},ticks:{color:'#94a3b8',font:{size:9}}}
    },maintainAspectRatio:false}});
}

// ============ LEADERBOARD ============
let lbSortKey='w', lbSortDir=-1;
let yrSortKey='w', yrSortDir=-1;
let ksrSortKey='rating', ksrSortDir=-1;
let _ksrRows=[], _ksrLatestYr='';
function buildLeaderboard(){
  const names=[...STATS.players].sort((a,b)=>b.gm-a.gm).map(p=>p.name);
  document.getElementById('h2hA').innerHTML=names.map(n=>`<option>${n}</option>`).join('');
  document.getElementById('h2hB').innerHTML=names.map(n=>`<option>${n}</option>`).join('');
  document.getElementById('h2hB').selectedIndex=1;
  document.getElementById('partnerSearch').innerHTML=names.map(n=>`<option>${n}</option>`).join('');
  filterTable();
}
function filterTable(){
  const q=(document.getElementById('searchPlayer').value||'').trim();
  const min=parseInt(document.getElementById('minGames').value)||1;
  const rows=STATS.players
    .filter(p=>p.gm>=min&&(q===''||p.name.includes(q)))
    .map(p=>({...p,
      pts:(p.w||0)*2+(p.d||0),
      ppg:p.gm>0?((p.w||0)*2+(p.d||0))/p.gm:0,
      winpct:p.gm>0?p.w/p.gm:0,
      gpg:p.gm>0?p.g/p.gm:0,
      apg:p.gm>0?p.a/p.gm:0,
      contrib:p.gm>0?(p.g+p.a)/p.gm:0,
      mvp_n:(BONUS[p.name]||{}).mvp||0,
      wg_n:(BONUS[p.name]||{}).wg||0
    }));
  rows.sort((a,b)=>{
    if(lbSortKey==='w'){
      let d=b.w-a.w; if(d!==0) return lbSortDir*d;
      d=b.pts-a.pts; if(d!==0) return lbSortDir*d;
      return lbSortDir*(b.g-a.g);
    }
    return lbSortDir*(b[lbSortKey]-a[lbSortKey]);
  });
  document.getElementById('lbBody').innerHTML=rows.map((p,i)=>`
    <tr>
      <td>${i+1}. ${p.name}</td>
      <td>${p.gm}</td>
      <td style="font-weight:bold;color:#10b981">${p.w}</td>
      <td>${pct(p.w,p.gm)}%</td>
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
function sortTable(key){
  if(lbSortKey===key) lbSortDir*=-1; else{lbSortKey=key;lbSortDir=-1;}
  filterTable();
}

// ============ KOSHER (RATING) ============
// score for a single year's stats + bonus: win%×40 + g/gm×20 + a/gm×10 + mvp/gm×30 + wg/gm×20
function yrScore(yd, bonus){
  if(!yd||yd.gm<1) return null;
  const b=bonus||{mvp:0,wg:0};
  return (yd.w/yd.gm)*40 + (yd.g/yd.gm)*20 + (yd.a/yd.gm)*10
        + (b.mvp/yd.gm)*30 + (b.wg/yd.gm)*20;
}
function buildKosher(){
  const min=parseInt(document.getElementById('kosherMinGames').value)||5;
  const byYearMap={};
  STATS.byYear.forEach(e=>{
    if(!byYearMap[e.name]) byYearMap[e.name]={};
    byYearMap[e.name][e.yr]=e;
  });

  _ksrLatestYr=KSR_YEARS[KSR_YEARS.length-1];

  _ksrRows=STATS.players
    .filter(p=>{
      // must have played in at least one כושר year
      const yrs=byYearMap[p.name]||{};
      const recentGm=KSR_YEARS.reduce((s,y)=>s+(yrs[y]?yrs[y].gm:0),0);
      return recentGm>=min;
    })
    .map(p=>{
      const yrs=byYearMap[p.name]||{};
      const bonusYrs=BONUS_BY_YEAR[p.name]||{};

      // weighted rating over כושר years
      let wSum=0, wTot=0;
      let recentGm=0,recentW=0,recentG=0,recentA=0,recentMvp=0,recentWg=0;
      KSR_YEARS.forEach(y=>{
        const yd=yrs[y]; if(!yd||yd.gm<1) return;
        const bns=bonusYrs[y]||{mvp:0,wg:0};
        const sc=yrScore(yd,bns);
        const wt=KSR_WEIGHTS[y]||1;
        wSum+=sc*wt; wTot+=wt;
        recentGm+=yd.gm; recentW+=yd.w;
        recentG+=yd.g;   recentA+=yd.a;
        recentMvp+=(bns.mvp||0); recentWg+=(bns.wg||0);
      });
      const rating=wTot>0?wSum/wTot:0;

      // trend: compare 2026 vs 2025 year scores
      const sc26=yrScore(yrs['2026'],bonusYrs['2026']);
      const sc25=yrScore(yrs['2025'],bonusYrs['2025']);
      let trend='same';
      if(sc26!==null&&sc25!==null){
        if(sc26>sc25*1.05) trend='up';
        else if(sc26<sc25*0.95) trend='down';
      } else if(sc26!==null&&sc25===null){
        trend='up';
      } else if(sc26===null&&sc25!==null){
        trend='down';
      }

      const latestYd=yrs['2026']||yrs['2025']||yrs['2024'];
      return {...p,
        rating, trend,
        recentGm, recentW,
        wp: recentGm>0?recentW/recentGm:0,
        gp: recentGm>0?recentG/recentGm:0,
        ap: recentGm>0?recentA/recentGm:0,
        mvp_n: recentMvp,
        wg_n:  recentWg,
        latestGm: latestYd?latestYd.gm:0,
        yr26gm: yrs['2026']?yrs['2026'].gm:0,
        yr25gm: yrs['2025']?yrs['2025'].gm:0,
        yr24gm: yrs['2024']?yrs['2024'].gm:0,
      };
    });
  renderKosherTable();
}
function sortKosher(key){
  if(ksrSortKey===key) ksrSortDir*=-1; else{ksrSortKey=key;ksrSortDir=-1;}
  renderKosherTable();
}
function renderKosherTable(){
  const rows=[..._ksrRows].sort((a,b)=>{
    const va=a[ksrSortKey]??0, vb=b[ksrSortKey]??0;
    if(typeof va==='string') return ksrSortDir*(va<vb?-1:va>vb?1:0);
    return ksrSortDir*(vb-va);
  });
  const maxR=Math.max(...rows.map(r=>r.rating),0.01);
  const trendIcon=t=>t==='up'?'<span class="trend-up">↑</span>':t==='down'?'<span class="trend-down">↓</span>':'<span class="trend-same">→</span>';
  const thK=(key,label)=>`<th onclick="sortKosher('${key}')" style="${ksrSortKey===key?'color:#fff':''}">${label}${ksrSortKey===key?(ksrSortDir===-1?' ▼':' ▲'):''}</th>`;

  document.getElementById('kosherTable').innerHTML=`
    <table>
      <thead><tr>
        <th>#</th>
        <th onclick="sortKosher('name')" style="text-align:right;${ksrSortKey==='name'?'color:#fff':''}">שחקן${ksrSortKey==='name'?(ksrSortDir===-1?' ▼':' ▲'):''}</th>
        ${thK('recentGm',"מ' 24-26")}
        ${thK('yr26gm','2026')}
        ${thK('yr25gm','2025')}
        ${thK('yr24gm','2024')}
        <th>מגמה</th>
        ${thK('wp','%נצח')}
        ${thK('gp','ש/מ')}
        ${thK('ap','ב/מ')}
        ${thK('mvp_n','MVP')}
        ${thK('wg_n','שניצ׳')}
        ${thK('rating','ניקוד ★')}
      </tr></thead>
      <tbody>${rows.map((p,i)=>`
        <tr>
          <td>${i+1}</td>
          <td style="text-align:right;font-weight:bold">${p.name}</td>
          <td>${p.recentGm}</td>
          <td style="color:#64748b">${p.yr26gm||'-'}</td>
          <td style="color:#475569">${p.yr25gm||'-'}</td>
          <td style="color:#475569">${p.yr24gm||'-'}</td>
          <td style="text-align:center">${trendIcon(p.trend)}</td>
          <td>${pct(p.recentW,p.recentGm)}%</td>
          <td>${r2(p.gp)}</td>
          <td>${r2(p.ap)}</td>
          <td style="color:#fbbf24;font-weight:bold">${p.mvp_n||'-'}</td>
          <td style="color:#10b981;font-weight:bold">${p.wg_n||'-'}</td>
          <td>
            <b style="color:#fbbf24;font-size:.9rem">${r2(p.rating)}</b>
            <div class="bar-wrap"><div class="bar bar-gold" style="width:${Math.round(100*p.rating/maxR)}%"></div></div>
          </td>
        </tr>`).join('')}
      </tbody>
    </table>`;
}

// ============ MVP LEADERBOARD ============
function buildMVPLeaderboard(){
  const rows=[...(STATS.bonus||[])].filter(b=>b.mvp>0||b.wg>0)
    .sort((a,b)=>b.mvp-a.mvp||b.wg-a.wg);
  if(!rows.length){document.getElementById('mvpLeaderboard').innerHTML='<p style="color:#64748b;padding:10px">אין נתונים</p>';return;}
  document.getElementById('mvpLeaderboard').innerHTML=`
    <table>
      <thead><tr>
        <th>#</th>
        <th style="text-align:right">שחקן</th>
        <th style="color:#fbbf24">🏅 MVP</th>
        <th style="color:#10b981">⚡ שניצ׳</th>
      </tr></thead>
      <tbody>${rows.map((b,i)=>{
        const medal=i===0?'🥇':i===1?'🥈':i===2?'🥉':'';
        return `<tr>
          <td>${medal||i+1}</td>
          <td style="text-align:right;font-weight:bold">${b.name}</td>
          <td style="color:#fbbf24;font-weight:bold">${b.mvp||'-'}</td>
          <td style="color:#10b981;font-weight:bold">${b.wg||'-'}</td>
        </tr>`;
      }).join('')}</tbody>
    </table>`;
}

// ============ HEAD TO HEAD ============
function calcH2H(){
  const pA=document.getElementById('h2hA').value;
  const pB=document.getElementById('h2hB').value;
  if(pA===pB){document.getElementById('h2hResult').innerHTML='<p style="color:#ef4444;padding:12px">בחר שני שחקנים שונים</p>';return;}
  const sorted=[pA,pB].sort();
  const rv=STATS.rivals.find(r=>r.p1===sorted[0]&&r.p2===sorted[1]);

  const aIsFirst=sorted[0]===pA;
  const aWins=rv ? (aIsFirst ? rv.fw : rv.t-rv.fw-rv.d) : 0;
  const bWins=rv ? (aIsFirst ? rv.t-rv.fw-rv.d : rv.fw) : 0;
  const draws=rv ? rv.d : 0;
  const total=rv ? rv.t : 0;

  const aG=rv ? (aIsFirst ? rv.p1g : rv.p2g) : 0;
  const aAs=rv ? (aIsFirst ? rv.p1a : rv.p2a) : 0;
  const bG=rv ? (aIsFirst ? rv.p2g : rv.p1g) : 0;
  const bAs=rv ? (aIsFirst ? rv.p2a : rv.p1a) : 0;

  const sA=STATS.players.find(p=>p.name===pA);
  const sB=STATS.players.find(p=>p.name===pB);
  const pair=STATS.pairs.find(p=>p.p1===sorted[0]&&p.p2===sorted[1]);

  let html=`
    <div class="vs-wrap">
      <div class="vs-player">
        <div class="vs-name">${pA}</div>
        <div class="vs-wins" style="color:#3b82f6">${aWins}</div>
        <div class="vs-label">ניצחונות מולו</div>
        <div style="margin-top:6px;font-size:.76rem;color:#94a3b8">${aG} ש׳ | ${aAs} ב׳</div>
      </div>
      <div class="vs-mid">
        ⚔️<br>
        <span style="font-size:.78rem;color:#64748b">${total} מפגשים ישירים</span><br>
        <span style="font-size:.72rem;color:#475569">${draws} תיקו</span>
      </div>
      <div class="vs-player">
        <div class="vs-name">${pB}</div>
        <div class="vs-wins" style="color:#ef4444">${bWins}</div>
        <div class="vs-label">ניצחונות מולו</div>
        <div style="margin-top:6px;font-size:.76rem;color:#94a3b8">${bG} ש׳ | ${bAs} ב׳</div>
      </div>
    </div>`;

  if(pair){
    html+=`<div style="text-align:center;margin:10px 0;font-size:.84rem;color:#94a3b8">
      🤝 כשמשחקים <b style="color:#fbbf24">ביחד</b>: ${pair.t} משחקים —
      <b style="color:#10b981">${pair.w} ניצחונות</b> |
      <span style="color:#ef4444">${pair.l} הפסדים</span> |
      ${pct(pair.w,pair.t)}% ניצחון
    </div>`;
  }

  if(!rv){
    html=`<div style="text-align:center;padding:20px;color:#64748b">
      ⚠️ לא נמצאו מפגשים ישירים בין ${pA} ל-${pB}
    </div>`;
  }

  if(sA&&sB){
    html+=`<div class="grid2" style="margin-top:10px;gap:8px">
      ${playerMiniCard(sA)} ${playerMiniCard(sB)}
    </div>`;
  }
  document.getElementById('h2hResult').innerHTML=html;
}
function playerMiniCard(p){
  const pts=(p.w||0)*2+(p.d||0);
  const b=BONUS[p.name]||{mvp:0,wg:0};
  return `<div class="card" style="font-size:.8rem">
    <div style="font-size:.95rem;font-weight:bold;color:#fbbf24;margin-bottom:6px">${p.name}</div>
    <div>משחקים כלל: <b>${p.gm}</b></div>
    <div>ניצחונות: <b>${p.w}</b> (${pct(p.w,p.gm)}%)</div>
    <div>נקודות: <b>${pts}</b> (${r2(pts/p.gm)}/מ׳)</div>
    <div>שערים: <b>${p.g}</b> (${r2(p.g/p.gm)}/מ׳)</div>
    <div>בישולים: <b>${p.a}</b> (${r2(p.a/p.gm)}/מ׳)</div>
    ${b.mvp?`<div>MVP: <b style="color:#fbbf24">${b.mvp}</b></div>`:''}
    ${b.wg?`<div>שניצ׳: <b style="color:#10b981">${b.wg}</b></div>`:''}
  </div>`;
}

// ============ PARTNERS ============
function buildPartners(){
  const topW=getTopPairs('w',15,15);
  const topL=getTopPairs('l',15,15);

  document.getElementById('bestPairsWins').innerHTML=topW.map((p,i)=>`
    <div class="pair-row">
      <span class="pair-names">${i+1}. ${p.p1} + ${p.p2}</span>
      <span>
        <span class="chip chip-green">${p.w} נצח׳</span>
        <span class="chip chip-orange">${pct(p.w,p.t)}%</span>
        <span class="chip chip-blue">${p.t} מ׳</span>
        <span class="chip" style="background:#1a2a1a;color:#6ee7b7">${p.g||0} ש׳</span>
      </span>
    </div>`).join('');

  document.getElementById('worstPairsLosses').innerHTML=topL.map((p,i)=>`
    <div class="pair-row">
      <span class="pair-names">${i+1}. ${p.p1} + ${p.p2}</span>
      <span>
        <span class="chip chip-red">${p.l} הפס׳</span>
        <span class="chip chip-orange">${pct(p.l,p.t)}% הפסד</span>
        <span class="chip chip-blue">${p.t} מ׳</span>
      </span>
    </div>`).join('');
}

function showPartnerStats(){
  const pName=document.getElementById('partnerSearch').value;
  const myPairs=[...STATS.pairs]
    .filter(p=>p.p1===pName||p.p2===pName)
    .map(p=>{const other=p.p1===pName?p.p2:p.p1; return {...p,other};})
    .sort((a,b)=>(b.t>0&&a.t>0)?(b.w/b.t)-(a.w/a.t):b.t-a.t);  // sort by win%
  if(!myPairs.length){document.getElementById('partnerResult').innerHTML='<p style="color:#64748b;padding:10px">אין נתונים</p>';return;}

  document.getElementById('partnerResult').innerHTML=`
    <div class="tbl-wrap" style="max-height:380px">
      <table>
        <thead><tr>
          <th>שותף</th>
          <th>משחקים</th>
          <th style="color:#10b981">ניצחונות</th>
          <th style="color:#ef4444">הפסדים</th>
          <th>תיקו</th>
          <th>% ניצחון</th>
          <th>שערים</th>
          <th>בישולים</th>
        </tr></thead>
        <tbody>${myPairs.map(p=>`
          <tr>
            <td>${p.other}</td>
            <td>${p.t}</td>
            <td style="color:#10b981;font-weight:bold">${p.w}</td>
            <td style="color:#ef4444;font-weight:bold">${p.l}</td>
            <td>${p.t-p.w-p.l}</td>
            <td>
              ${pct(p.w,p.t)}%
              <div class="bar-wrap"><div class="bar bar-green" style="width:${pct(p.w,p.t)}%"></div></div>
            </td>
            <td style="color:#fbbf24">${p.g||0}</td>
            <td style="color:#8b5cf6">${p.a||0}</td>
          </tr>`).join('')}
        </tbody>
      </table>
    </div>`;
}

// ============ YEARS ============
function buildYears(){
  const yrs=[...new Set(STATS.byYear.map(e=>e.yr))].filter(y=>y&&y!=='unknown').sort();
  const sel=document.getElementById('yearSelect');
  sel.innerHTML=yrs.map(y=>`<option value="${y}">${y}</option>`).join('');
  sel.value=yrs[yrs.length-1];
  showYearStats();
}
function showYearStats(){
  yrSortKey='w'; yrSortDir=-1;
  renderYearTable();
}
function sortYearTable(key){
  if(yrSortKey===key) yrSortDir*=-1; else{yrSortKey=key;yrSortDir=-1;}
  renderYearTable();
}
function renderYearTable(){
  const yr=document.getElementById('yearSelect').value;
  const min=parseInt(document.getElementById('yearMinGames').value)||1;
  const rows=STATS.byYear
    .filter(e=>e.yr===yr&&e.gm>=min)
    .map(p=>({...p,
      pts:(p.w||0)*2+(p.d||0),
      ppg:p.gm>0?((p.w||0)*2+(p.d||0))/p.gm:0,
      winpct:p.gm>0?p.w/p.gm:0,
      gpg:p.gm>0?p.g/p.gm:0,
      apg:p.gm>0?p.a/p.gm:0
    }))
    .sort((a,b)=>{
      if(yrSortKey==='name') return yrSortDir*(a.name<b.name?-1:a.name>b.name?1:0);
      return yrSortDir*(b[yrSortKey]-a[yrSortKey]);
    });
  const maxPts=Math.max(...rows.map(r=>r.pts),1);
  const maxG=Math.max(...rows.map(r=>r.g),1);
  const yrTh=(key,label)=>`<th onclick="sortYearTable('${key}')" style="${yrSortKey===key?'color:#fff':''}">${label}${yrSortKey===key?(yrSortDir===-1?' ▼':' ▲'):''}</th>`;
  document.getElementById('yearTable').innerHTML=`
    <table>
      <thead><tr>
        <th>#</th>
        ${yrTh('name','שחקן')}
        ${yrTh('gm','משח׳')}
        ${yrTh('w','נצח׳')}
        ${yrTh('winpct','%נצח')}
        ${yrTh('pts','נק׳')}
        ${yrTh('ppg','נק/מ')}
        ${yrTh('g','שע׳')}
        ${yrTh('a','בישול')}
      </tr></thead>
      <tbody>${rows.map((p,i)=>`
        <tr>
          <td>${i+1}</td><td>${p.name}</td><td>${p.gm}</td>
          <td style="font-weight:bold;color:#10b981">${p.w}</td>
          <td>${pct(p.w,p.gm)}%</td>
          <td style="color:#fbbf24">${p.pts}
            <div class="bar-wrap"><div class="bar bar-green" style="width:${Math.round(100*p.pts/maxPts)}%"></div></div>
          </td>
          <td>${r2(p.ppg)}</td>
          <td>${p.g}
            <div class="bar-wrap"><div class="bar bar-gold" style="width:${Math.round(100*p.g/maxG)}%"></div></div>
          </td>
          <td>${p.a}</td>
        </tr>`).join('')}
      </tbody>
    </table>`;
}

// ============ FUN FACTS ============
function buildFacts(){
  const topGames=[...GAMES].sort((a,b)=>(b.scoreA+b.scoreB)-(a.scoreA+a.scoreB)).slice(0,3);
  const bigWins=[...GAMES].sort((a,b)=>Math.abs(b.scoreA-b.scoreB)-Math.abs(a.scoreA-a.scoreB)).slice(0,3);
  const bestW=getTopPairs('w',3,15);
  const bestPct=getTopPairs('pct',3,15);
  const worstL=getTopPairs('l',3,15);
  const topLineups=getTopLineups(3);

  const ff=(label,val)=>`<div class="fun-fact"><div class="ff-label">${label}</div><div class="ff-val">${val}</div></div>`;

  document.getElementById('factsLeft').innerHTML=`
    <h3>⚽ משחקים עם הכי הרבה שערים</h3>
    ${topGames.map((g,i)=>ff(`#${i+1} — ${g.date}`,`${g.scoreA}:${g.scoreB} (סה״כ ${g.scoreA+g.scoreB})`)).join('')}
    <h3 style="margin-top:12px">🏆 הזוגות עם הכי הרבה ניצחונות</h3>
    ${bestW.map((p,i)=>ff(`#${i+1} — ${p.t} משחקים יחד`,`${p.p1} + ${p.p2} — ${p.w} ניצחונות`)).join('')}
    <h3 style="margin-top:12px">🎯 הזוגות היעילים ביותר (%)</h3>
    ${bestPct.map((p,i)=>ff(`#${i+1} — ${p.t} משחקים יחד`,`${p.p1} + ${p.p2} — ${pct(p.w,p.t)}%`)).join('')}
  `;

  document.getElementById('factsRight').innerHTML=`
    <h3>💥 ניצחונות בפער הגדול ביותר</h3>
    ${bigWins.map((g,i)=>ff(`#${i+1} — ${g.date}`,`${g.scoreA}:${g.scoreB} (פער ${Math.abs(g.scoreA-g.scoreB)})`)).join('')}
    <h3 style="margin-top:12px">🛡️ ההרכב המנצח המדויק</h3>
    ${topLineups.map((l,i)=>ff(`#${i+1} — ${l[1]} ניצחונות`,`<span style="font-size:.75rem;line-height:1.5">${l[0]}</span>`)).join('')}
    <h3 style="margin-top:12px">💔 הזוגות עם הכי הרבה הפסדים</h3>
    ${worstL.map((p,i)=>ff(`#${i+1} — ${p.t} משחקים יחד`,`${p.p1} + ${p.p2} — ${p.l} הפסדים`)).join('')}
  `;

  const rivals=[...STATS.rivals].filter(r=>r.t>=20).sort((a,b)=>b.t-a.t).slice(0,15);
  document.getElementById('rivalsList').innerHTML=`
    <table>
      <thead><tr>
        <th>שחקן א׳</th>
        <th>מפגשים</th>
        <th>נצח׳</th>
        <th style="color:#64748b">VS</th>
        <th>נצח׳</th>
        <th>שחקן ב׳</th>
        <th>תיקו</th>
        <th>דומיננטיות</th>
      </tr></thead>
      <tbody>${rivals.map(r=>{
        const aW=r.fw, bW=r.t-r.fw-r.d, dom=Math.abs(pct(aW,r.t)-pct(bW,r.t));
        return `<tr>
          <td>${r.p1}</td>
          <td style="color:#fbbf24;font-weight:bold">${r.t}</td>
          <td style="color:${aW>bW?'#10b981':'#ef4444'};font-weight:bold">${aW}</td>
          <td style="color:#475569">⚔️</td>
          <td style="color:${bW>aW?'#10b981':'#ef4444'};font-weight:bold">${bW}</td>
          <td>${r.p2}</td>
          <td style="color:#64748b">${r.d}</td>
          <td><span class="chip chip-orange">${dom}% הפרש</span></td>
        </tr>`;
      }).join('')}</tbody>
    </table>`;
}

// ============ INIT ============
buildOverview();
buildLeaderboard();
buildMVPLeaderboard();
buildKosher();
buildPartners();
buildYears();
buildFacts();
calcH2H();
showPartnerStats();
</script>
</body>
</html>
'''

if __name__ == '__main__':
    generate()
