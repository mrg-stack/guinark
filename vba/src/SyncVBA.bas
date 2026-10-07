Option Explicit

' One-time setup in Excel:
' 1) Windows: Developer -> Macro Security -> Trust access to the VBA project object model
' 2) Mac: Excel -> Settings/Preferences -> Security (or Security & Privacy) -> enable VBA project access if available
' 2) Import this module into your workbook once
'
' Workflow:
' - Run ExportVbaProjectToFiles before editing with AI in VS Code
' - Edit files under /vba/src
' - Run ImportVbaFilesToProject to load changes back into the workbook

Private Const vbext_ct_StdModule As Long = 1
Private Const vbext_ct_ClassModule As Long = 2
Private Const vbext_ct_MSForm As Long = 3
Private Const vbext_ct_Document As Long = 100
Private Const SYNC_MODULE_NAME As String = "SyncVBA"
' Local project root, used when ThisWorkbook.Path is a SharePoint/OneDrive URL.
Private Const MANUAL_PROJECT_ROOT As String = "/Users/guin/Developer/guinark"
' Optional: set full folder path to VBA files to bypass all auto-detection.
Private Const MANUAL_SYNC_FOLDER As String = ""

Public Sub ExportVbaProjectToFiles()
    On Error GoTo ExportErr

    Dim exportPath As String

    If Not CanAccessVBProject() Then
        ShowVBProjectAccessHelp "export", Err.Number, Err.Description
        Exit Sub
    End If

    exportPath = GetExportPath()

    EnsureFolderExists exportPath
    ExportComponents exportPath

    MsgBox "VBA exported to: " & exportPath, vbInformation
    Exit Sub

ExportErr:
    ShowVBProjectAccessHelp "export", Err.Number, Err.Description
End Sub

Public Sub ImportVbaFilesToProject()
    On Error GoTo ImportErr

    Dim importPath As String

    If Not CanAccessVBProject() Then
        ShowVBProjectAccessHelp "import", Err.Number, Err.Description
        Exit Sub
    End If

    importPath = GetExportPath()

    If Dir(importPath, vbDirectory) = vbNullString Then
        MsgBox "Folder not found: " & importPath, vbExclamation
        Exit Sub
    End If

    ' Defensive cleanup for stale legacy modules that can break compilation
    ' before normal import replacement happens.
    RemoveKnownBrokenModules

    RemoveImportedComponents
    ImportComponents importPath

    MsgBox "SUCCESS! VBA imported from: " & importPath, vbInformation
    Exit Sub

ImportErr:
    ShowVBProjectAccessHelp "import", Err.Number, Err.Description
End Sub

Public Sub RepairVbaProjectForImport()
    On Error GoTo RepairErr

    If Not CanAccessVBProject() Then
        ShowVBProjectAccessHelp "repair", Err.Number, Err.Description
        Exit Sub
    End If

    RemoveKnownBrokenModules
    RemoveImportedComponents

    MsgBox "VBA project cleanup completed. You can run ImportVbaFilesToProject now.", vbInformation
    Exit Sub

RepairErr:
    ShowVBProjectAccessHelp "repair", Err.Number, Err.Description
End Sub

Private Function GetExportPath() As String
    Dim basePath As String
    Dim srcPath As String
    Dim vbaPath As String

    If Len(Trim$(MANUAL_SYNC_FOLDER)) > 0 Then
        GetExportPath = NormalizePath(MANUAL_SYNC_FOLDER)
        Exit Function
    End If

    basePath = ResolveBasePath()
    If basePath = vbNullString Then
        GetExportPath = vbNullString
        Exit Function
    End If

    vbaPath = BuildPath(basePath, "vba")
    srcPath = BuildPath(vbaPath, "src")

    If Dir(srcPath, vbDirectory) <> vbNullString Then
        GetExportPath = srcPath
    Else
        GetExportPath = vbaPath
    End If
End Function

Private Function ResolveBasePath() As String
    Dim workbookPath As String
    workbookPath = NormalizePath(ThisWorkbook.Path)

    If Len(Trim$(MANUAL_PROJECT_ROOT)) > 0 Then
        ResolveBasePath = NormalizePath(MANUAL_PROJECT_ROOT)
        Exit Function
    End If

    If IsWebPath(workbookPath) Then
        MsgBox "Workbook path is a web URL (SharePoint/OneDrive)." & vbCrLf & vbCrLf & _
               "Set MANUAL_PROJECT_ROOT in SyncVBA.bas to your local project folder, then run again.", _
               vbExclamation, "SyncVBA Path Setup Required"
        ResolveBasePath = vbNullString
        Exit Function
    End If

    ResolveBasePath = workbookPath
End Function

Private Function BuildPath(ByVal basePath As String, ByVal childName As String) As String
    BuildPath = NormalizePath(basePath) & "/" & childName
End Function

Private Function NormalizePath(ByVal pathValue As String) As String
    Dim p As String
    p = Replace(pathValue, "\", "/")

    Do While Right$(p, 1) = "/"
        p = Left$(p, Len(p) - 1)
    Loop

    NormalizePath = p
End Function

Private Function IsWebPath(ByVal pathValue As String) As Boolean
    Dim p As String
    p = LCase$(Trim$(pathValue))

    IsWebPath = (Left$(p, 7) = "http://" Or Left$(p, 8) = "https://")
End Function

Private Sub EnsureFolderExists(ByVal folderPath As String)
    Dim parentPath As String
    Dim sepPos As Long

    If Dir(folderPath, vbDirectory) <> vbNullString Then
        Exit Sub
    End If

    sepPos = InStrRev(folderPath, Application.PathSeparator)
    If sepPos > 0 Then
        parentPath = Left$(folderPath, sepPos - 1)
        If Len(parentPath) > 0 And Dir(parentPath, vbDirectory) = vbNullString Then
            EnsureFolderExists parentPath
        End If
    End If

    MkDir folderPath
End Sub

Private Sub ExportComponents(ByVal folderPath As String)
    Dim vbComp As Object
    Dim fileName As String
    Dim filePath As String

    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case CLng(vbComp.Type)
            Case vbext_ct_StdModule
                fileName = vbComp.Name & ".bas"
            Case vbext_ct_ClassModule
                fileName = vbComp.Name & ".cls"
            Case vbext_ct_MSForm
                fileName = vbComp.Name & ".frm"
            Case Else
                ' Skip worksheet/workbook document modules and unknown component types.
                fileName = vbNullString
        End Select

        If fileName <> vbNullString Then
            filePath = folderPath & Application.PathSeparator & fileName
            If Dir(filePath) <> vbNullString Then
                Kill filePath
            End If
            vbComp.Export filePath
        End If
    Next vbComp
End Sub

Private Sub RemoveImportedComponents()
    Dim vbComp As Object
    Dim toDelete As Collection
    Dim i As Long

    Set toDelete = New Collection

    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case CLng(vbComp.Type)
            Case vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_MSForm
                If StrComp(vbComp.Name, SYNC_MODULE_NAME, vbTextCompare) <> 0 Then
                    toDelete.Add vbComp
                End If
            Case vbext_ct_Document
                ' Keep ThisWorkbook and Sheet modules in place.
            Case Else
                ' Keep unknown types.
        End Select
    Next vbComp

    For i = toDelete.Count To 1 Step -1
        ThisWorkbook.VBProject.VBComponents.Remove toDelete(i)
    Next i
End Sub

Private Sub RemoveKnownBrokenModules()
    Dim namesToRemove As Variant
    Dim i As Long

    namesToRemove = Array( _
        "scheduelcolorevents", _
        "schedulecolorevents", _
        "clsScheduleColorEvents", _
        "ScheduleColorEvents", _
        "ScheduelColorEvents" _
    )

    For i = LBound(namesToRemove) To UBound(namesToRemove)
        RemoveComponentIfExists CStr(namesToRemove(i))
    Next i

    RemoveMalformedClassHeaderStdModules
End Sub

Private Sub RemoveComponentIfExists(ByVal componentName As String)
    Dim vbComp As Object

    On Error Resume Next
    Set vbComp = ThisWorkbook.VBProject.VBComponents(componentName)
    On Error GoTo 0

    If Not vbComp Is Nothing Then
        ThisWorkbook.VBProject.VBComponents.Remove vbComp
    End If
End Sub

Private Sub RemoveMalformedClassHeaderStdModules()
    Dim vbComp As Object
    Dim toDelete As Collection
    Dim i As Long
    Dim firstLine As String

    Set toDelete = New Collection

    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If CLng(vbComp.Type) = vbext_ct_StdModule Then
            If vbComp.CodeModule.CountOfLines > 0 Then
                firstLine = Trim$(CStr(vbComp.CodeModule.Lines(1, 1)))
                If UCase$(firstLine) = "VERSION 1.0 CLASS" Then
                    toDelete.Add vbComp
                End If
            End If
        End If
    Next vbComp

    For i = toDelete.Count To 1 Step -1
        ThisWorkbook.VBProject.VBComponents.Remove toDelete(i)
    Next i
End Sub

Private Sub ImportComponents(ByVal folderPath As String)
    ImportByPattern folderPath, "*.bas"
    ImportByPattern folderPath, "*.cls"
    ImportByPattern folderPath, "*.frm"
End Sub

Private Sub ImportByPattern(ByVal folderPath As String, ByVal pattern As String)
    Dim fileName As String
    Dim fullPath As String

    fileName = Dir(folderPath & Application.PathSeparator & pattern)

    Do While fileName <> vbNullString
        fullPath = folderPath & Application.PathSeparator & fileName
        If StrComp(fileName, SYNC_MODULE_NAME & ".bas", vbTextCompare) <> 0 Then
            ThisWorkbook.VBProject.VBComponents.Import fullPath
        End If
        fileName = Dir
    Loop
End Sub

Private Function CanAccessVBProject() As Boolean
    On Error GoTo AccessDenied

    Dim componentCount As Long
    componentCount = ThisWorkbook.VBProject.VBComponents.Count

    CanAccessVBProject = (componentCount >= 0)
    Exit Function

AccessDenied:
    CanAccessVBProject = False
End Function

Private Sub ShowVBProjectAccessHelp(ByVal actionName As String, ByVal errNum As Long, ByVal errDesc As String)
    Dim msg As String

    msg = "Cannot " & actionName & " VBA modules. " & vbCrLf & vbCrLf
    msg = msg & "Error: " & CStr(errNum) & " - " & errDesc & vbCrLf & vbCrLf
    msg = msg & "Try this:" & vbCrLf
    msg = msg & "1) In Excel settings/security, enable VBA project access if available." & vbCrLf
    msg = msg & "2) Save workbook locally (outside OneDrive/iCloud) and retry." & vbCrLf
    msg = msg & "3) Ensure workbook VBA project is not password protected." & vbCrLf
    msg = msg & "4) Grant Excel file access in macOS Privacy if prompted."

    MsgBox msg, vbExclamation, "VBA Sync Access Error"
End Sub
