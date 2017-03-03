Sub AutoExec()
'
' AutoExec Macro
'
'
　
' Set environment path
Dim fso As Object
'US = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%Homepath%")
US = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%Homeshare%")
Path = US & "\My Documents\WindowsPowerShell"
SubDir = US & "\My Documents"
PSFile = Path & "\profile.ps1"
　
' Create directory and file
Set fso = CreateObject("Scripting.FileSystemObject")
Dim oFile As Object
　
If fso.FolderExists(Path) Then
    Else
    Set objFolder = fso.CreateFolder(SubDir + "\WindowsPowerShell")
End If
　
If fso.FileExists(PSFile) Then
    Else
    Set oFile = fso.CreateTextFile(PSFile, True, True)
    oFile.WriteLine "<Insert ps1 code>"
    oFile.Close
    Set fso = Nothing
    Set oFile = Nothing
End If
　
' Execute powershell in background
Set objShell = CreateObject("WScript.Shell")
'objShell.Run ("C:\Windows\system32\WindowsPowerShell\1.0\powershell.exe -File") & Path, 0, False
　
　
End Sub
