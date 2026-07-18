# Friday Soccer Dashboard

## מה זה
דשבורד אינטראקטיבי לקבוצת כדורגל שישי שבועית — 15 שנה, 600+ משחקים, ~40 שחקנים.

- דשבורד חי: https://fogelyotam.github.io/Friday-soccer/
- GitHub: https://github.com/FogelYotam/Friday-soccer
- בעלים: FogelYotam (yotamfogel@gmail.com)

---

## קבצים מרכזיים

| קובץ | תפקיד |
|------|--------|
| `games_data.json` | מקור האמת היחיד — כל המשחקים |
| `gen_dashboard.py` | בונה dashboard.html מ-JSON + Excel |
| `dashboard.html` | הדשבורד הנבנה (לא לערוך ידנית) |
| `index.html` | עותק של dashboard.html — GitHub Pages |
| `entry.html` | הזנת משחקים מהטלפון דרך GitHub API |
| `match_entry.html` | הזנת משחקים מהמחשב (קובץ מקומי) |
| `update.bat` / `update.ps1` | עדכון ידני מהמחשב |
| `soccer.xlsx` | נתוני Excel היסטוריים (MVP/WG לפני JSON) |
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
`entry.html` → PUT ל-GitHub API → GitHub Actions → `gen_dashboard.py` → `index.html` → Pages חי (~2 דקות)

### מהמחשב (ידני)
1. פתח `match_entry.html` → טען `games_data.json` → הזן משחק → שמור (יורד ל-Downloads)
2. הרץ `update.bat` → מעתיק מ-Downloads → בונה → push

---

## כרטיסיות הדשבורד (סדר הניווט)

| כרטיסייה | תוכן |
|-----------|-------|
| 🏠 סקירה | גיבורי השנה (בורר שנה), שיאי כל-הזמנים, גרף משחקים, רצפי שיא |
| 📋 מחזור אחרון | בורר מחזורים, הרכב מלא + תוצאות, MVP ושער ניצחון |
| 🏆 מצטיינים | טבלה מלאה, מיון לפי ניצחונות, בורר שנה + סף משחקים |
| 💪 כושר | דירוג כושר משוקלל (ראה נוסחה למטה) |
| ⚔️ ראש בראש | השוואה ישירה בין שני שחקנים |
| 👤 הדף שלי | דף אישי — סטטיסטיקות, גרף מגמה, שותפים, יריבים, שיאי רצף |

---

## דירוג כוח (טאב "כושר")

מבוסס **ELO** (שחמט-style, `ELO` ב-JS): כל שחקן מתחיל ב-1500, ומתעדכן אחרי כל משחק לפי התוצאה מול הציפייה (ממוצע ELO של היריבים מול השותפים), K=28.

- **דירוג כוח** (`eloForm`) = **בסיס + כושר − היעדרות**:
  - *בסיס* — ה-ELO שהשחקן צבר עד לפני 6 חודשים (רמת הקריירה שלו, "כמה הוא שווה").
  - *כושר* — כמה עלה/ירד ב-6 החודשים האחרונים.
  - *היעדרות* — `ELO_ABSENCE_PER_MONTH` (5) נק׳ לכל חודש היעדרות אחרי חודש חסד, עד `ELO_ABSENCE_MAX` (60).
  - זכאות: מעל `KSR_MIN_CAREER_GAMES` (100) משחקי קריירה **וגם** לפחות משחק אחד ב-`KSR_ACTIVE_YEARS` (2024–2026).
- **עמדה** (`ROLE`) — מוסקת מהנתונים (אין שדה עמדה): שוער = פחות מ-0.25 שערים *וגם* בישולים למשחק; אחרת התקפה/הגנה לפי תרומה למשחק מול חציון הקבוצה.
- מבני עזר: `ELO_SERIES[name]` (סדרת דירוגים לכל משחק), `ELO_MOVES` (game→{name:{d,r}}), `ELO_UPSETS` (הפתעות).
- טבלת "ELO כל הזמנים" (`buildElo`) — הדירוג הנוכחי + שיא, כל ההיסטוריה.
- הסיכום לוואטסאפ (`gameSummaryText`) מציג את ה-ELO של כל שחקן אחרי המשחק.

---

## gen_dashboard.py — נקודות חשובות

- **`MERGE_MAP`** — נרמול שמות שגויים/כפולים → שם קנוני
- **`SKIP_NAMES`** — `{"עצמי", "שער עצמי"}` — מסוננים מהסטטיסטיקות
- **`MIN_GAMES_THRESHOLD = 20`** — פחות מ-20 משחקים = נחשב אורח, לא מופיע בסטטיסטיקות (37 שחקנים עומדים בסף)
- **`EXCEL_PATH`** — קובץ Excel היסטורי ל-MVP/WG; אם לא קיים, נופל ל-`soccer.xlsx` (ל-CI)
- **`HTML_TEMPLATE`** — כאן לשנות עיצוב/כרטיסיות, לא ב-`dashboard.html`
- **פורמט תאריכים ב-`games_data.json`: `M/D/YYYY`** (אמריקאי, כמו `entry.html`). `parseDate` ב-JS קורא כך.
- **בנייה דטרמיניסטית** — כל הרשימות ב-`stats` ממוינות לפני השמירה, כדי שבנייה מקומית ו-CI יפיקו פלט זהה (מונע קונפליקטים).
- אחרי כל שינוי: `python gen_dashboard.py` ואז `cp dashboard.html index.html`

---

## מעקב מבקרים

- כפתור "מי אתה?" פלוטינג — מבקר בוחר שמו, נשמר ב-`localStorage`
- כניסה נשלחת לـ Google Sheet דרך `WEBHOOK_URL` ב-`gen_dashboard.py`
- ה-webhook כבר מוגדר בקוד (Apps Script)

---

## entry.html (טלפון)

- טוקן GitHub נשמר ב-`localStorage` כ-`gh_token`
- יוצרים טוקן ב: https://github.com/settings/tokens (הרשאה: `repo`)
- כתובת: https://fogelyotam.github.io/Friday-soccer/entry.html
- SHA נטען מחדש לפני כל commit — מונע קונפליקטים בשמירות מרובות

---

## שחקנים — הערות

- **לאון** (וותיק, 300+ משחקים) ≠ **ליאון** (שחקן חדש מ-2026, אדם אחר)
- שמות כפולים: ראה `MERGE_MAP` בראש `gen_dashboard.py`
- שחקן חדש (לא הופיע בעבר) — מודגש בצהוב בטפסי ההזנה

---

## GitHub Actions

`.github/workflows/build.yml` — מופעל בכל push שמשנה `games_data.json` או `gen_dashboard.py`:
1. `python gen_dashboard.py`
2. `cp dashboard.html index.html`
3. `git commit && git push`

אם יש קונפליקט בין push מקומי ל-Actions: `git pull` → בנה מחדש → push.
