Attribute VB_Name = "modOverblik"
Option Explicit

Public Function RebuildOverviewSafely() As String
    On Error GoTo Failed
    GatherSchedulesToOverblik
    ThisWorkbook.Save
    RebuildOverviewSafely = "Overview rebuilt and saved."
    Exit Function
Failed:
    RebuildOverviewSafely = "ERROR " & Err.Number & ": " & Err.Description
End Function

Public Sub GatherSchedulesToOverblik()
    Dim wb As Workbook, overview As Worksheet, semester As Worksheet, firstSemester As Worksheet
    Dim sem As Long, blockRow As Long, isBigChecked As Boolean, oldAlerts As Boolean
    Dim oldEvents As Boolean, oldScreen As Boolean, oldCalculation As XlCalculation, failure As String, stage As String
    On Error GoTo Failed
    Set wb = ThisWorkbook
    oldAlerts = Application.DisplayAlerts
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldCalculation = Application.Calculation
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    isBigChecked = (wb.Worksheets("Data").Range("D43").Value2 = True)
    For sem = 1 To 7
        If SheetExists(sem & ". Sem", wb) Then
            Set firstSemester = wb.Worksheets(sem & ". Sem")
            Exit For
        End If
    Next sem
    If firstSemester Is Nothing Then Err.Raise 5, , "Generate the semester sheets before building Overblik."
    stage = "Preparing references"
    PrepareOverviewReferences
    If SheetExists("__NewOverblik", wb) Then Err.Raise 5, , "Remove the unfinished __NewOverblik sheet before rebuilding."
    Set overview = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    overview.Name = "__NewOverblik"
    stage = "Building header"
    BuildOverviewHeader firstSemester, overview
    blockRow = 6
    For sem = 1 To 7
        If SheetExists(sem & ". Sem", wb) Then
            Set semester = wb.Worksheets(sem & ". Sem")
            stage = "Copying " & semester.Name
            If IsBigSemesterLayout(sem, isBigChecked) Then
                CopyOverviewBlock semester, overview, 7, blockRow, semester.Name & " A"
                blockRow = blockRow + 16
                CopyOverviewBlock semester, overview, 28, blockRow, semester.Name & " B"
            Else
                CopyOverviewBlock semester, overview, 7, blockRow, semester.Name
            End If
            blockRow = blockRow + 16
        End If
    Next sem
    Application.DisplayAlerts = False
    If SheetExists("Overblik", wb) Then wb.Worksheets("Overblik").Delete
    overview.Name = "Overblik"
    Application.DisplayAlerts = oldAlerts
    stage = "Ordering workbook tabs"
    OrderWorkbookTabs
    overview.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitRow = 0
    ActiveWindow.SplitColumn = 0
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    overview.Range("C6").Select
    ActiveWindow.FreezePanes = True
    overview.Calculate
    ProtectOverblik overview
    Application.CutCopyMode = False
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalculation
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Exit Sub
Failed:
    failure = "Overblik could not be generated (" & stage & "): " & Err.Description
    On Error GoTo CleanupFailed
    If Not overview Is Nothing Then
        If overview.Name = "__NewOverblik" Then
            Application.DisplayAlerts = False
            overview.Delete
        End If
    End If
    Application.CutCopyMode = False
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalculation
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    Err.Raise 5, "GatherSchedulesToOverblik", failure
CleanupFailed:
    failure = failure & "; cleanup failed: " & Err.Description
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalculation
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Err.Raise 5, "GatherSchedulesToOverblik", failure
End Sub

Public Sub ProtectOverblik(ByVal overview As Worksheet)
    overview.Unprotect
    overview.Cells.Validation.Delete
    overview.Cells.Locked = True
    overview.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    overview.EnableSelection = xlNoRestrictions
End Sub

Private Sub BuildOverviewHeader(ByVal semester As Worksheet, ByVal overview As Worksheet)
    Dim col As Long, source As Range, target As Range, template As Worksheet, row As Long
    Set template = ThisWorkbook.Worksheets("TemplateOverblikHeader")
    template.Range("A1:DC4").Copy Destination:=overview.Range("A1")
    For row = 1 To 4
        overview.Rows(row).RowHeight = template.Rows(row).RowHeight
    Next row
    overview.Columns("A").ColumnWidth = template.Columns("A").ColumnWidth
    overview.Columns("B").ColumnWidth = template.Columns("B").ColumnWidth
    For col = 3 To 107
        overview.Columns(col).ColumnWidth = template.Columns(col).ColumnWidth
        Set source = semester.Cells(4, col + 2)
        Set target = overview.Cells(3, col)
        target.Formula2 = "='" & semester.Name & "'!" & source.Address
        Set source = semester.Cells(5, col + 2)
        Set target = overview.Cells(4, col)
        target.Formula2 = "='" & semester.Name & "'!" & source.Address
        If (col - 3) Mod 5 = 0 Then
            overview.Cells(2, col).Formula2 = "=""Uge ""&WEEKNUM(" & target.Address & ",21)"
        End If
    Next col
    overview.Rows(5).RowHeight = semester.Rows(6).RowHeight
    overview.Range("A5:DC180").Interior.Color = semester.Range("E6").Interior.Color
End Sub

Private Sub CopyOverviewBlock(ByVal semester As Worksheet, ByVal overview As Worksheet, _
                              ByVal sourceRow As Long, ByVal targetRow As Long, ByVal label As String)
    Dim source As Range, target As Range, row As Long, col As Long, formulas(1 To 15, 1 To 105) As Variant
    Dim sourceCell As Range, labelRange As Range, stepName As String
    On Error GoTo Failed
    Set source = semester.Range("E" & sourceRow & ":DE" & sourceRow + 14)
    Set target = overview.Range("C" & targetRow & ":DC" & targetRow + 14)
    semester.Calculate
    stepName = "Formatting semester label"
    Set labelRange = overview.Range("A" & targetRow & ":A" & targetRow + 14)
    labelRange.UnMerge
    ThisWorkbook.Worksheets("TemplateOverblikHeader").Range("A6:A20").Copy
    labelRange.PasteSpecial Paste:=xlPasteFormats
    labelRange.Merge
    With labelRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Orientation = 90
    End With
    stepName = "Copying range"
    source.Copy Destination:=target
    semester.Range("B" & sourceRow & ":B" & sourceRow + 14).Copy _
        Destination:=overview.Cells(targetRow, 2)
    semester.Range("E6:DE6").Copy _
        Destination:=overview.Range("C" & targetRow - 1 & ":DC" & targetRow - 1)
    labelRange.Cells(1, 1).Value2 = label
    stepName = "Building linked values"
    For row = 1 To 15
        overview.Rows(targetRow + row - 1).RowHeight = semester.Rows(sourceRow + row - 1).RowHeight
        overview.Cells(targetRow + row - 1, 2).Value2 = semester.Cells(sourceRow + row - 1, 2).Value2
        For col = 1 To 105
            Set sourceCell = source.Cells(row, col)
            formulas(row, col) = "=IF('" & semester.Name & "'!" & sourceCell.Address & _
                "="""","""",'" & semester.Name & "'!" & sourceCell.Address & ")"
        Next col
    Next row
    target.Formula2 = formulas
    Exit Sub
Failed:
    Err.Raise 5, "CopyOverviewBlock", stepName & ": " & Err.Description
End Sub

Public Function VerifyGeneratedOverblik() As String
    Dim overview As Worksheet, semester As Worksheet, row As Long, offset As Long, col As Long
    Dim label As String, sourceRow As Long, checked As Long, sourceValues As Variant, targetValues As Variant
    Dim labelRange As Range
    Set overview = ThisWorkbook.Worksheets("Overblik")
    overview.Calculate
    For row = 6 To overview.Cells(overview.Rows.Count, 1).End(xlUp).Row Step 16
        label = CStr(overview.Cells(row, 1).Value2)
        Set labelRange = overview.Range("A" & row & ":A" & row + 14)
        If Not labelRange.MergeCells Then _
            Err.Raise 5, , "Semester label is not merged at " & overview.Cells(row, 1).Address
        If overview.Cells(row, 1).Orientation <> 90 Then _
            Err.Raise 5, , "Semester label is not vertical at " & overview.Cells(row, 1).Address
        Set semester = ThisWorkbook.Worksheets(Trim$(Replace(Replace(label, " A", ""), " B", "")))
        sourceRow = 7
        If Right$(label, 1) = "B" Then sourceRow = 28
        sourceValues = semester.Range("E" & sourceRow & ":DE" & sourceRow + 14).Value2
        targetValues = overview.Range("C" & row & ":DC" & row + 14).Value2
        For offset = 0 To 14
            For col = 3 To 107
                If targetValues(offset + 1, col - 2) <> sourceValues(offset + 1, col - 2) Then _
                    Err.Raise 5, , "Overview value mismatch at " & overview.Cells(row + offset, col).Address
                checked = checked + 1
            Next col
        Next offset
    Next row
    VerifyGeneratedOverblik = "PASS: " & checked & " overview values match their semester source cells."
End Function
