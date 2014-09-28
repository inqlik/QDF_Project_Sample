Const ForReading = 1 

Set objFSO = CreateObject("Scripting.FileSystemObject") 

rootPath = objFSO.GetParentFolderName(objFSO.GetParentFolderName(objFSO.GetParentFolderName(objFSO.GetParentFolderName(Wscript.ScriptFullName))))
Set objTextFile = objFSO.OpenTextFile(rootPath & "\0.Administration\3.Include\1.BaseVariable\ContainerMap.csv", ForReading) 

function createProjectContent()
	strOutput = "" 
	strOutput = strOutput & "{" & vbCrLf
	strOutput = strOutput & """folders"":" & vbCrLf
	strOutput = strOutput & "  [" & vbCrLf

	firstLine = True

	strOutput = strOutput & "    {" & vbCrLf
	strOutput = strOutput & "      ""path"": """ & rootPath & "\1.App\3.Include\1.BaseVariable""," & vbCrLf
	strOutput = strOutput & "      ""name"": ""App Variables""" & vbCrLf
	strOutput = strOutput & "    }," & vbCrLf


	Do Until objTextFile.AtEndOfStream 
	    strNextLine = objTextFile.Readline 
	    if firstLine then
	    	firstLine = False
	    else
		    arrServiceList = Split(strNextLine , ",")
		    If (UBound(arrServiceList) = 3) Then
		    	containerName = arrServiceList(0)
		    	if containerName <> "" and containerName <> "Admin" and containerName <> "Shared" and containerName <> "Sys" then
		    		dirName = arrServiceList(1)
					strOutput = strOutput & "    {" & vbCrLf
					strOutput = strOutput & "      ""path"": """ & rootPath & "\" & dirName & "\3.Include\6.Custom""," & vbCrLf
					strOutput = strOutput & "      ""name"": """ & containerName & " scripts""" & vbCrLf
					strOutput = strOutput & "    }," & vbCrLf
					strOutput = strOutput & "    {" & vbCrLf
					strOutput = strOutput & "      ""path"": """ & rootPath & "\" & dirName & "\2.QVD""," & vbCrLf
					strOutput = strOutput & "      ""name"": """ & containerName & " QVDS""" & vbCrLf
					strOutput = strOutput & "    }," & vbCrLf
					if containerName = "Trans" then 
						strOutput = strOutput & "    {" & vbCrLf
						strOutput = strOutput & "      ""path"": """ & rootPath & "\" & dirName & "\3.Include\4.Sub""," & vbCrLf
						strOutput = strOutput & "      ""name"": """ & containerName & " subroutines""" & vbCrLf
						strOutput = strOutput & "    }," & vbCrLf
					end if		
				end if        	
		    End If
	    end if
	Loop 
	strOutput = strOutput & "    {" & vbCrLf
	strOutput = strOutput & "      ""path"": """ & rootPath & "\99.Shared_Folders\3.Include\1.BaseVariable""," & vbCrLf
	strOutput = strOutput & "      ""name"": ""Shared Variables""" & vbCrLf
	strOutput = strOutput & "    }," & vbCrLf

	strOutput = strOutput & "    {" & vbCrLf
	strOutput = strOutput & "      ""path"": """ & rootPath & "\0.Administration\6.Script\Vbs""," & vbCrLf
	strOutput = strOutput & "      ""name"": ""VBS Administration scripts""" & vbCrLf
	strOutput = strOutput & "    }" & vbCrLf

	strOutput = strOutput & "  ]" & vbCrLf
	strOutput = strOutput & "}" & vbCrLf

	strOutput = Replace(strOutput,"\","~~~~~")
	createProjectContent = Replace(strOutput,"~~~~~","\\")
End Function
Function createProject(userName)
	qvProjectName = objFSO.getFolder(rootPath).Name
	projectsRootPath = objFSO.BuildPath(rootPath, "0.Administration\9.Misc")
	if not objFSO.FolderExists(projectsRootPath) then
    WScript.Echo projectsRootPath & " created"
		objFSO.CreateFolder(projectsRootPath)
	end if
	projectsRootPath = objFSO.BuildPath(projectsRootPath, "\" & "SublimeProjects")
	if not objFSO.FolderExists(projectsRootPath) then
    WScript.Echo projectsRootPath & " created"
		objFSO.CreateFolder(projectsRootPath)
	end if
	if userName <> "" then 
		projectsRootPath = objFSO.BuildPath(projectsRootPath, "\" & userName)
		if not objFSO.FolderExists(projectsRootPath) then
			objFSO.CreateFolder(projectsRootPath)
		end if
	end if
	fileName = objFSO.BuildPath(projectsRootPath, qvProjectName & ".sublime-project")
	set objOutFile = objFSO.CreateTextFile(fileName,True)
    objOutFile.Write(content)
	objOutFile.Close
	WScript.Echo fileName & " created"
End Function
Dim content
content = createProjectContent()
'WScript.Echo content
createProject "vts"
createProject "gla"
createProject "osk"
createProject "apn"

'WScript.Echo res