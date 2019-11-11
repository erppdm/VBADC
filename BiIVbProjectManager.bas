' V0.1.0
' Copyright (c) 2019 Ulf-Dirk Stockburger

Option Explicit

' If yes, the macro will be terminated without feedback if an error occurs
Private Const beQuiet = True
' The headline in the message window
Private Const errMsgHeader = "BiIVbManager"
' The VBE project name
Private Const vbProject = "BiIVbProject"
' The directory with the INI file
Private Const iniFolder = "ini"
' Section to import
Private Const iniImportModuleSection = "Import"
'Key with the full file name of the module to import
Private Const iniImportModuleKey = "importMod"

' <Loads the modules specified in the ini file>
Function ProjectLoader() As Boolean
    On Error GoTo errFunc
    Dim projectFound As Boolean
    Dim project As Object
    
    If Not LoadLibraries Then ShowErrMsg "Error loading 'Microsoft Visual Basic for Applications Extensibility 5.3'."
    
    For Each project In Application.VBE.VBProjects
        'Debug.Print project.Name
        If project.name = vbProject Then
            projectFound = True
            If Not ImportVbeModules(project) Then ShowErrMsg "Error loading the macro modules"
            Exit For
        End If
    Next
    
    If projectFound Then If Not beQuiet Then MsgBox "All modules imported."
errFunc:
    If Err.Number <> 0 Then On Error GoTo 0: Exit Function
exitFunc:
    ProjectLoader = True
End Function
' </Loads the modules specified in the ini file>

'<Exports the modules specified in the ini file>
Function ProjectExporter() As Boolean
    On Error GoTo errFunc
    Dim i%
    Dim iniFileName$
    Dim exportModule$
    Dim moduleName$
    Dim moduleExportet As Boolean
    Dim project As Object
    Dim module As Object
    Dim ini As BiIClassIni
    
    If Not LoadLibraries Then ShowErrMsg "Error loading 'Microsoft Visual Basic for Applications Extensibility 5.3'."
    
    Set ini = New BiIClassIni
    
    For Each project In Application.VBE.VBProjects
        If project.name = vbProject Then Exit For
    Next
       
    iniFileName = GetIniFileName(project.fileName)
    If iniFileName = vbNullString Then
        ShowErrMsg "The name of the INI file cannot be retrieved."
    End If
    If Not FileExists(iniFileName) Then ShowErrMsg "INI file '" & iniFileName & "' not found."
    
    i = 0
    Do
        exportModule = ini.ReadKey(iniFileName, iniImportModuleSection, iniImportModuleKey & i)
        If exportModule <> vbNullString Then
            moduleName = GetFilenameWoExtension(exportModule)
            For Each module In project.VBComponents
                If module.name = moduleName Then
                    project.VBComponents(moduleName).Export exportModule
                    moduleExportet = True
                    Exit For
                End If
            Next
        End If
        i = i + 1
    Loop While exportModule <> vbNullString
    
    If moduleExportet Then
        If Not beQuiet Then MsgBox "All modules exportet."
    End If
errFunc:
    If Err.Number <> 0 Then On Error GoTo 0: Exit Function
exitFunc:
    ProjectExporter = True
End Function
'</Exports the modules specified in the ini file>

'<Loads 'Microsoft Visual Basic for Applications Extensibility 5.3'>
Private Function LoadLibraries() As Boolean
    On Error GoTo errFunc
    
    Dim reference As Object
    Dim vbProj As Object
    Dim vbideExists As Boolean

    Set vbProj = Application.VBE.ActiveVBProject

    For Each reference In vbProj.References
        If reference.name = "VBIDE" Then vbideExists = True: Exit For
    Next
    
    If Not vbideExists Then Application.VBE.ActiveVBProject.References.AddFromGuid GUID:="{0002E157-0000-0000-C000-000000000046}", Major:=5, Minor:=3
    'Application.VBE.ActiveVBProject.References.AddFromFile "C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"

errFunc:
    Set vbProj = Nothing
    If Err.Number <> 0 Then On Error GoTo 0: Exit Function
exitFunc:
    LoadLibraries = True
End Function
'<Loads 'Microsoft Visual Basic for Applications Extensibility 5.3'>

'<Imports the modules specified in the ini file>
Private Function ImportVbeModules(ByRef project As Object) As Boolean
    On Error GoTo errFunc
    Dim i%
    Dim iniFileName$
    Dim importModule$
    Dim VBComp As Object
    Dim ini As BiIClassIni
    
    Set ini = New BiIClassIni
    
    iniFileName = GetIniFileName(project.fileName)
    If iniFileName = vbNullString Then
        ShowErrMsg "The name of the INI file cannot be retrieved."
    End If
    If Not FileExists(iniFileName) Then
        ShowErrMsg "INI file '" & iniFileName & "' not found."
    End If
    
    i = 0
    Do
        importModule = ini.ReadKey(iniFileName, iniImportModuleSection, iniImportModuleKey & i)
        If importModule = vbNullString Then
            If i = 0 Then ShowErrMsg "No module to import found."
            GoTo errFunc
        End If
        
        If Not RemoveModuleIfExists(project, GetFilenameWoExtension(importModule)) Then
            ShowErrMsg "Module '" & GetFilenameWoExtension(importModule) & "' could not be deleted."
        End If
        
        Set VBComp = Nothing
        Set VBComp = project.VBComponents.Import(importModule)
        VBComp.name = GetFilenameWoExtension(importModule)
        
        i = i + 1
    Loop While importModule <> vbNullString
    
errFunc:
    Set project = Nothing
    Set VBComp = Nothing
    Set ini = Nothing
    Set VBComp = Nothing
    If Err.Number <> 0 Then On Error GoTo 0: Exit Function
exitFunc:
    ImportVbeModules = True
End Function
'</Imports the modules specified in the ini file>

'<Deletes an existing module>
Private Function RemoveModuleIfExists(ByRef project As Object, ByVal moduleName$) As Boolean
    On Error GoTo errFunc
    
    Dim module As Object
    
    Set module = project.VBComponents(moduleName)
    module.name = module.name & "_del"
    project.VBComponents.Remove module
errFunc:
    Set module = Nothing
    If Err.Number <> 0 Then
        Dim errNo&: errNo = Err.Number
        On Error GoTo 0
        If errNo <> 9 Then Exit Function
    End If
exitFunc:
    RemoveModuleIfExists = True
End Function
'</Deletes an existing module>

'<Gets the name of the INI file>
Private Function GetIniFileName(projectFileName$) As String
    On Error GoTo errFunc
    
    GetIniFileName = GetPath(projectFileName) & "\" & iniFolder & "\" & GetFilenameWoExtension(projectFileName) & ".ini"

errFunc:
    If Err.Number <> 0 Then On Error GoTo 0: Exit Function
exitFunc:
End Function
'</Gets the name of the INI file>

'<Extracts the path from the full filename>
Private Function GetPath(fileName As String) As String
    On Error GoTo errFunc
    
    If VBA.InStr(1, fileName, "\", vbTextCompare) > 1 And VBA.Len(fileName) > 0 Then
        GetPath = VBA.Left(fileName, VBA.InStrRev(fileName, "\") - 1)
    End If

errFunc:
    If Err.Number <> 0 Then On Error GoTo 0: Exit Function
exitFunc:
End Function
'</Extracts the path from the full filename>

'<Extracts the filename from the full filename>
Private Function GetFilename(fileName As String) As String
    On Error GoTo errFunc
    
    If VBA.Right(fileName, 1) <> "\" And VBA.Len(fileName) > 0 Then
        GetFilename = VBA.Right(fileName, VBA.Len(fileName) - VBA.InStrRev(fileName, "\"))
    End If

errFunc:
    If Err.Number <> 0 Then On Error GoTo 0: Exit Function
exitFunc:
End Function
'</Extracts the filename from the full filename>

'<Extracts the filename without extension from the full filename>
Private Function GetFilenameWoExtension(fileName As String) As String
    On Error GoTo errFunc
    
    Dim tmpFileName$
    
    If VBA.Right(fileName, 1) <> "\" And VBA.Len(fileName) > 0 Then
        tmpFileName = VBA.Right(fileName, VBA.Len(fileName) - VBA.InStrRev(fileName, "\"))
        If (InStrRev(tmpFileName, ".") - 1) <> -1 Then
            GetFilenameWoExtension = VBA.Left(tmpFileName, VBA.InStrRev(tmpFileName, ".") - 1)
        End If
    End If
    
errFunc:
    If Err.Number <> 0 Then On Error GoTo 0: Exit Function
exitFunc:
End Function
'</Extracts the filename without extension from the full filename>

'<Checks whether the specified file exists>
Private Function FileExists(fileName$) As Boolean
    On Error GoTo errFunc
   
    If Not VBA.dir(fileName, vbDirectory) = vbNullString Then
        FileExists = True
    End If

errFunc:
    If Err.Number <> 0 Then On Error GoTo 0: Exit Function
exitFunc:
End Function
'</Checks whether the specified file exists>

'<Error handler>
Private Function ShowErrMsg(msg$) As Boolean
    On Error GoTo errFunc
    
    If Not beQuiet Then
        MsgBox msg, vbCritical, errMsgHeader
    Else
        ' <Whatever needs to be done>
    End If
    
errFunc:
    If Err.Number <> 0 Then On Error GoTo 0
exitFunc:
    ShowErrMsg = True
    End
End Function
'</Error handler>
