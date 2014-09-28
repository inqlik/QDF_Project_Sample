Const qwReloadCommand = "cmd /C ""C:\Program Files\QlikView\Qv.Exe"" /R "


Set objFSO = CreateObject("Scripting.FileSystemObject")
rootPath = objFSO.GetParentFolderName(objFSO.GetParentFolderName(objFSO.GetParentFolderName(objFSO.GetParentFolderName(Wscript.ScriptFullName))))
projectName = objFSO.GetFolder(rootPath).Name

'You may set sourceFileName manually'
sourceFileName = rootPath & "\1.App\1.Application\" & projectName & ".qvw"

backupDirName = rootPath & "\1.App\9.Misc"
If not objFSO.FolderExists(backupDirName) Then
  objFSO.CreateFolder(backupDirName)
End If
backupDirName = backupDirName & "\AppBackup"
If not objFSO.FolderExists(backupDirName) Then
  objFSO.CreateFolder(backupDirName)
End If

If not objFSO.FileExists(sourceFileName) Then
    Wscript.Echo "Source file does not exist: " & sourceFileName
    WScript.Quit 1
End If
sourceBaseName = objFSO.GetFileName(sourceFileName)
backupFile = backupDirName & "\" & sourceBaseName
if objFSO.FileExists(backupFile) Then
  if objFSO.GetFile(backupFile).DateLastModified >= objFSO.GetFile(sourceFileName).DateLastModified Then
    WScript.Echo "Backup is fresh. Skipping file " & sourceFileName
    WScript.Quit
  End if
end if
WSCript.Echo "Creating backup for " & sourceFileName
objFSO.CopyFile sourceFileName, backupFile
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run qwReloadCommand & backupFile , 0, True

dateStr = FormatDateTime(Now())
dateStr = Replace(dateStr,".","_")
dateStr = Replace(dateStr,":","_")
dateStr = Replace(dateStr," ","__")
backupWithTimestamp = Replace(backupFile,".qvw","_" & dateStr & ".qvw")
objFSO.CopyFile backupFile, backupWithTimestamp
WScript.Echo "Created backup: " & backupWithTimestamp