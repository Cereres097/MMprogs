# Устанавливаем пути в кавычках для защиты от пробелов
$url = "https://github.com"
$path = "$env:TEMP\DocStats.exe"

# 1. Полная очистка перед стартом
if (Test-Path "$path") { 
    Remove-Item "$path" -Force -ErrorAction SilentlyContinue 
}

Write-Host "Загрузка DocStats (около 70 МБ)..." -ForegroundColor Cyan

# 2. Скачивание
try {
    Invoke-WebRequest -Uri $url -OutFile "$path" -ErrorAction Stop
} catch {
    Write-Host "Ошибка при скачивании. Проверьте интернет или ссылку на GitHub." -ForegroundColor Red
    exit
}

# 3. Проверка: если файл меньше 1 МБ, значит скачалась ошибка, а не программа
if ((Get-Item "$path").Length -lt 1048576) {
    Write-Host "Критическая ошибка: Скачанный файл поврежден или ссылка ведет на пустую страницу." -ForegroundColor Red
    Remove-Item "$path" -Force
    exit
}

Write-Host "Запуск программы..." -ForegroundColor Green

# 4. Запуск (используем кавычки и амперсанд)
& "$path"

# 5. Очистка после закрытия
if (Test-Path "$path") { 
    Remove-Item "$path" -Force -ErrorAction SilentlyContinue 
}
