$env:PYTHONUTF8 = "1"
$soccerDir  = "C:\Users\shlom\OneDrive\Documents\SOCCER"
$dataFile   = "$soccerDir\games_data.json"
$downloads  = "$env:USERPROFILE\Downloads\games_data.json"

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  כדורגל שישי - עדכון דשבורד" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Step 1: pull latest from GitHub (phone commits)
Write-Host "מושך עדכונים מ-GitHub..." -ForegroundColor Yellow
git pull --quiet

Write-Host ""

# Step 2: detect new file in Downloads
if (Test-Path $downloads) {
    $dlTime  = (Get-Item $downloads).LastWriteTime
    $curTime = (Get-Item $dataFile).LastWriteTime
    if ($dlTime -gt $curTime) {
        Write-Host "נמצא קובץ חדש ב-Downloads — מעתיק..." -ForegroundColor Yellow
        Copy-Item $downloads $dataFile -Force
        Remove-Item $downloads
        Write-Host "הועתק בהצלחה." -ForegroundColor Green
    } else {
        Write-Host "הקובץ ב-Downloads ישן יותר — מדלג." -ForegroundColor Gray
    }
} else {
    Write-Host "לא נמצא קובץ חדש ב-Downloads." -ForegroundColor Gray
}

Write-Host ""

# Step 2: build dashboard
Write-Host "בונה דשבורד..." -ForegroundColor Yellow
python "$soccerDir\gen_dashboard.py"
if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "שגיאה בבניית הדשבורד!" -ForegroundColor Red
    pause; exit 1
}

# Step 3: copy to index.html
Copy-Item "$soccerDir\dashboard.html" "$soccerDir\index.html" -Force

# Step 4: get last game date for commit message
$lastDate = python -c "import json; g=json.load(open('$soccerDir\games_data.json','r',encoding='utf-8')); print(g[-1]['date'])"

Write-Host ""

# Step 5: git commit + push
Set-Location $soccerDir
git add index.html games_data.json
git diff --cached --quiet
if ($LASTEXITCODE -ne 0) {
    git commit -m "Update dashboard - $lastDate"
    git push
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "  הדשבורד עודכן! משחק: $lastDate" -ForegroundColor Green
    Write-Host "  https://fogelyotam.github.io/Friday-soccer/" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
} else {
    Write-Host "אין שינויים לדחוף." -ForegroundColor Gray
}

Write-Host ""
pause
