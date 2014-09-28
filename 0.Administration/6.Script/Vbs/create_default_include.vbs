Dim objFSO
Dim objQV
Set objQV=CreateObject("QlikTech.QlikView")
Set objFSO = CreateObject("Scripting.FileSystemObject")

rootPath = objFSO.GetParentFolderName(objFSO.GetParentFolderName(objFSO.GetParentFolderName(objFSO.GetParentFolderName(Wscript.ScriptFullName))))

 
CreateVarFile "2.Transformation"
CreateVarFile "3.Source"
CreateVarFile "1.App"


Set FSO = Nothing
 
Sub CreateVarFile(targetDir)
    targetPath = rootPath & "\" & targetDir
    if targetDir = "1.App" then
        qvwFile = targetPath & "\4.Mart\_NewFileTemplate.qvw"
    else
        qvwFile = targetPath & "\1.Application\_NewFileTemplate.qvw"
    end if 
    varFile = targetPath & "\3.Include\6.Custom\default_include.qvs"
''    WScript.Echo qvwFile & " " & varFile
    if not objFSO.FileExists(qvwFile) then
        if not objFSO.FileExists("_NewFileTemplate.qvw") then
            WScript.Echo "Cannot find file _NewFileTemplate.qvw"
            WScript.Exit 1
        end if
        objFSO.CopyFile "_NewFileTemplate.qvw" ,qvwFile
    end if     
    Dim folder, file
    Dim objSource
    Dim objVars, varcontent, objTempVar, varname, i
    Dim str, fl
    Set objSource = objQV.OpenDoc(qvwFile)
    set objVars = objSource.GetVariableDescriptions
    for i = 0 to objVars.Count - 1
        set objTempVar = objVars.Item(i)
        varname=Trim(objTempVar.Name)
        objSource.RemoveVariable(varname)
    next 'end of loop
    objSource.Reload    
    set objVars = objSource.GetVariableDescriptions
''    Set fl = objFSO.OpenTextFile(varFile, 2, True, -1) 
    Set fl = CreateObject("ADODB.Stream")
    fl.Open
    fl.CharSet = "utf-8"
    for i = 0 to objVars.Count - 1
        set objTempVar = objVars.Item(i)
        varname=Trim(objTempVar.Name)
        Set objSourceVar=objSource.Variables(varname)
        varcontent=objSourceVar.GetRawContent
        fl.WriteText "LET " & varname & " = '" & varcontent & "';" & vbNewline
        objSource.RemoveVariable(varname)
    next 'end of loop

    fl.writeText vbNewline
    fl.writeText "$(must_include=$(vG.SubPath)InQlik.qvs);"
    fl.SaveToFile varFile, 2
    Set fl = Nothing
    WScript.Echo varFile & " created"  
    objSource.Save
    objSource.CloseDoc                  
End Sub

Sub Save2File (sText, sFile)
    Dim oStream
    Set oStream = CreateObject("ADODB.Stream")
    With oStream
        .Open
        .CharSet = "utf-8"
        .WriteText sText
        .SaveToFile sFile, 2
    End With
    Set oStream = Nothing
End Sub