Dim destination1, destination2

Set objShell = CreateObject("WScript.Shell")
startupFolder = objShell.SpecialFolders("Startup")
destination1 = startupFolder & "\WebHelperService.exe"

appDataFolder = objShell.ExpandEnvironmentStrings("%LOCALAPPDATA%")
folderPath = appDataFolder & "\WindowsBackgroundProcess"
destination2 = folderPath & "\WebServiceHelper.dll"

' Create a reference to the FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Stop all instances of the exe from running before trying to delete them
Set colProcesses = GetObject("Winmgmts:").ExecQuery("Select * from Win32_Process where Name = 'WebHelperService.exe'")
For Each objProcess in colProcesses
    objProcess.Terminate()
Next

' Stop all instances of the rundll32.exe
Set colProcesses = GetObject("Winmgmts:").ExecQuery("Select * from Win32_Process where Name = 'rundll32.exe'")
For Each objProcess in colProcesses
    objProcess.Terminate()
Next

' Check if the files exist and delete them
If objFSO.FileExists(destination1) Then
    objFSO.DeleteFile(destination1)
End If

If objFSO.FileExists(destination2) Then
    objFSO.DeleteFile(destination2)
End If

' Delete the folder if it's empty
If objFSO.FolderExists(folderPath) Then
    If objFSO.GetFolder(folderPath).Files.Count = 0 Then
        objFSO.DeleteFolder(folderPath)
    End If
End If

strKeyPath = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\"
strValueName = "MyProgram"
Set objWshReg= CreateObject("WScript.Shell")

' Check if the registry value exists and delete it
On Error Resume Next
objWshReg.RegRead strKeyPath & strValueName

If Err.Number = 0 Then
    objWshReg.RegDelete strKeyPath & strValueName
End If

' Delete the scheduled task
Set objShell = CreateObject("WScript.Shell")

' The command to delete the scheduled task
strCommand = "schtasks /Delete /TN ""MyProgram"" /F"

' Run the command
objShell.Run strCommand, 0, True