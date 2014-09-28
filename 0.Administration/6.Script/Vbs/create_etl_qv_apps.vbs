Dim objFSO
Dim objQV
Set objQV=CreateObject("QlikTech.QlikView")
Set objFSO = CreateObject("Scripting.FileSystemObject")

rootPath = objFSO.GetParentFolderName(objFSO.GetParentFolderName(objFSO.GetParentFolderName(objFSO.GetParentFolderName(Wscript.ScriptFullName))))

 
CreateQvApps "2.Transformation"
CreateQvApps "3.Source"

objFSO.CopyFile "_NewFileTemplate.qvw" , rootPath & "\1.App\4.Mart\DataMart.qvw"

Set FSO = Nothing
 
Sub CreateQvApps(targetDir)
    targetPath = rootPath & "\" & targetDir
    if targetDir = "1.App" then
        qvwFile = targetPath & "\4.Mart\_NewFileTemplate.qvw"
    else
        qvwFile = targetPath & "\1.Application\_NewFileTemplate.qvw"
    end if

    Set objFolder = objFSO.GetFolder(targetPath & "\3.Include\6.Custom\")

    Set colFiles = objFolder.Files
    For Each objFile in colFiles
    	if InStr(objFile.Name,".qvs") > 0 then
	        Dim f
	    	Set f = objFSO.OpenTextFile (objFile)
	    	line = f.ReadLine
	        directive = InStr(line,"1.Application")
	    	if directive > 0 then
		    	targetFile = targetPath	& "\1.Application\" & Replace(objFile.Name,".qvs",".qvw")
			    if not objFSO.FileExists(targetFile) then
			        if not objFSO.FileExists("_NewFileTemplate.qvw") then
			            WScript.Echo "Cannot find file _NewFileTemplate.qvw"
			            WScript.Exit 1
			        end if
			        WScript.Echo targetFile
			        objFSO.CopyFile "_NewFileTemplate.qvw" , targetFile
			    end if     
	    	end if
	    	f.close
    	end if
    Next

End Sub
