$url = "https://github.com"
$localPath = Join-Path $env:TEMP "DocStats.exe"

# 1. Чистим старые копии
if (Test-Path -Path $localPath) { 
    Remove-Item -Path $localPath -Force -ErrorAction SilentlyContinue 
}

Write-Host "Загрузка DocStats.exe..." -ForegroundColor Cyan

# 2. Скачиваем файл
try {
    Invoke-WebRequest -Uri $url -OutFile $localPath -ErrorAction Stop
} catch {
    Write-Host "Ошибка: Не удалось скачать файл. Проверьте интернет." -ForegroundColor Red
    return
}

# 3. Проверка: если скачался текст вместо программы
if ((Get-Item $localPath).Length -lt 1000000) {
    Write-Host "Ошибка: Скачанный файл слишком мал. Ссылка на GitHub может быть неверной." -ForegroundColor Red
    Remove-Item $localPath -Force
    return
}

Write-Host "Запуск программы..." -ForegroundColor Green

# 4. Запуск в обход проблем с путями
Start-Process -FilePath $localPath -Wait

# 5. Очистка после закрытия
if (Test-Path -Path $localPath) { 
    Remove-Item -Path $localPath -Force -ErrorAction SilentlyContinue 
}
