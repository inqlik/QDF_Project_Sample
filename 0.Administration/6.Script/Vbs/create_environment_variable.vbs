Const ForReading = 1 
Set objFSO = CreateObject("Scripting.FileSystemObject")
rootPath = objFSO.GetParentFolderName(objFSO.GetParentFolderName(objFSO.GetParentFolderName(objFSO.GetParentFolderName(Wscript.ScriptFullName))))

Set objNet = CreateObject("WScript.Network")
strCompName = objNet.ComputerName

envVar = "PROD"
if strCompName = "INQLIK" then
  envVar = "DEV"
end if

Dim file
path = rootPath & "\0.Administration\3.Include\1.BaseVariable\generated_environment_descriptor.qvs"
set file = objFSO.CreateTextFile(path)
file.Write "LET vU.Environment = '" & envVar & "';"
file.close
WScript.echo "Created file: " & path



