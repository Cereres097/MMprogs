$u = "https://github.com"
$p = "$env:TEMP\DocStats.exe"

# Удаляем старый файл, если он есть (теперь без ошибок)
if (Test-Path "$p") { Remove-Item "$p" -Force -ErrorAction SilentlyContinue }

Write-Host "Загрузка программы..." -ForegroundColor Cyan
Invoke-WebRequest -Uri $u -OutFile "$p"

# Проверка: если файл скачался слишком маленьким (меньше 1МБ), значит это не EXE
if ((Get-Item "$p").Length -lt 1MB) {
    Write-Host "Ошибка: Файл скачался некорректно. Проверьте ссылку в релизе!" -ForegroundColor Red
    Remove-Item "$p"
    exit
}

Write-Host "Запуск..." -ForegroundColor Green
& "$p"

# Удаляем после закрытия
if (Test-Path "$p") { Remove-Item "$p" -Force -ErrorAction SilentlyContinue }
