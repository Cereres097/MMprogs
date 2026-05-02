[Net.ServicePointManager]::SecurityProtocol = "Tls12, Tls11, Tls"; \



# Указываем пути в кавычках для безопасности
$url = "https://github.com/Cereres097/MMprogs/releases/download/DocStats/DocStats.exe"
$localPath = "$env:TEMP\DocStats.exe"

# 1. Удаляем старый файл, если он остался (в кавычках!)
if (Test-Path "$localPath") { 
    Remove-Item -Path "$localPath" -Force -ErrorAction SilentlyContinue 
}

Write-Host "Загрузка DocStats (70 MB)..." -ForegroundColor Cyan

# 2. Скачивание с принудительным игнорированием кэша
try {
    $webClient = New-Object System.Net.WebClient
    $webClient.DownloadFile($url, "$localPath")
} catch {
    Write-Host "Ошибка скачивания: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# 3. Проверка размера (должен быть > 1 МБ)
if (Test-Path "$localPath") {
    $size = (Get-Item "$localPath").Length
    if ($size -lt 1000000) {
        Write-Host "Ошибка: Скачано всего $size байт. Это не программа, а ошибка GitHub." -ForegroundColor Red
        Write-Host "Попробуйте зайти в настройки репозитория и убедиться, что он Public." -ForegroundColor Yellow
        Remove-Item "$localPath" -Force -ErrorAction SilentlyContinue
        return
    }
} else {
    Write-Host "Ошибка: Файл не был создан." -ForegroundColor Red
    return
}

Write-Host "Запуск программы..." -ForegroundColor Green

# 4. Запуск процесса (в кавычках)
Start-Process -FilePath "$localPath" -Wait

# 5. Итоговая очистка
if (Test-Path "$localPath") { 
    Remove-Item -Path "$localPath" -Force -ErrorAction SilentlyContinue 
}
