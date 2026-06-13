# Friday Soccer Dashboard

## מה זה
דשבורד סטטיסטיקות לקבוצת כדורגל שישי שבועית — 600+ משחקים, ~40 שחקנים.

- דשבורד חי: https://fogelyotam.github.io/Friday-soccer/
- GitHub: https://github.com/FogelYotam/Friday-soccer
- בעלים: FogelYotam (yotamfogel@gmail.com)

---

## קבצים מרכזיים

| קובץ | תפקיד |
|------|--------|
| `games_data.json` | מקור האמת היחיד — כל המשחקים |
| `gen_dashboard.py` | בונה dashboard.html מתוך games_data.json |
| `dashboard.html` | הדשבורד הנבנה (לא לערוך ידנית) |
| `index.html` | עותק של dashboard.html — מה שמוגש ב-GitHub Pages |
| `entry.html` | הזנת משחקים מהטלפון דרך GitHub API |
| `match_entry.html` | הזנת משחקים מהמחשב (פותח קובץ מקומי) |
| `update.bat` | עדכון ידני מהמחשב |
| `.github/workflows/build.yml` | בניה אוטומטית בכל push של games_data.json |

---

## מבנה games_data.json

```json
[
  {
    "date": "5/29/2026",
    "teamA": [{"name": "עידן", "goals": 2, "assists": 1}],
    "teamB": [{"name": "פוגל", "goals": 0, "assists": 0}],
    "scoreA": 3,
    "scoreB": 2,
    "mvp": "עידן",
    "wg": "ירון"
  }
]
```

- פורמט תאריך: `M/D/YYYY` (ללא אפס מוביל)
- `mvp` ו-`wg` (שער ניצחון) — אופציונליים

---

## זרימת עבודה

### מהטלפון (אוטומטי)
טלפון שומר ← `entry.html` מבצע PUT ל-GitHub API ← GitHub Actions מריץ `gen_dashboard.py` ← מעדכן `index.html` ← Pages חי

### מהמחשב (ידני)
1. פתח `match_entry.html` ← טען `games_data.json` ← הזן משחק ← שמור (יורד ל-Downloads)
2. הרץ `update.bat` ← מעתיק מ-Downloads ← בונה ← push

---

## gen_dashboard.py — נקודות חשובות

- `MERGE_MAP` — מילון לנרמול שמות (שמות כפולים/שגויים → שם קנוני)
- `MIN_GAMES_THRESHOLD = 10` — שחקנים עם פחות לא מופיעים בלידרבורד הראשי
- `EXCEL_PATH` — קובץ Excel היסטורי לנתוני MVP/WG מלפני המעבר ל-JSON (לא חובה קיים)
- HTML_TEMPLATE בתוך הקובץ — לשנות כרטיסיות/עיצוב כאן, לא ב-dashboard.html
- לאחר כל שינוי: `python gen_dashboard.py` ואז `cp dashboard.html index.html`

### כרטיסיות הדשבורד (בסדר הניווט)
סקירה → מחזור אחרון → לפי שנה → מצטיינים → כושר → ראש בראש → שותפויות → עובדות

---

## שחקנים — הערות

- **לאון** (וותיק, 300+ משחקים) ≠ **ליאון** (שחקן חדש מ-2026, שחקן נפרד)
- שמות שנרשמים בצורות שונות מטופלים דרך `MERGE_MAP` ב-`gen_dashboard.py`
- שחקן חדש שמופיע בפעם הראשונה — `entry.html` ו-`match_entry.html` מדגישים אותו בצהוב

---

## entry.html (טלפון)

- טוקן GitHub נשמר ב-`localStorage` כ-`gh_token`
- יוצר טוקן ב: https://github.com/settings/tokens (הרשאה: `repo`)
- כתובת: https://fogelyotam.github.io/Friday-soccer/entry.html
- SHA נטען מחדש לפני כל commit (מונע קונפליקטים בשמירות מרובות)

---

## GitHub Actions

`.github/workflows/build.yml` — מופעל בכל push שמשנה `games_data.json` או `gen_dashboard.py`:
1. `python gen_dashboard.py`
2. `cp dashboard.html index.html`
3. `git commit && git push`

אם יש קונפליקט בין push מקומי ל-Actions: `git pull` ואז בנה מחדש.
