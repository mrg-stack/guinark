Attribute VB_Name = "modSemesters"

Option Explicit

Public Function VerifySemesterGeneration() As String
    On Error GoTo Failed
    Generate_Semesters
    VerifySemesterGeneration = VerifyGeneratedOverblik()
    Exit Function
Failed:
    VerifySemesterGeneration = "ERROR " & Err.Number & ": " & Err.Description
End Function

'=============================================================
' Create 1_7.sem sheets from TemplateBig / TemplateSmall
'=============================================================
Public Sub MakeSemesters()
    Dim wb As Workbook
    Dim templateName As String
    Dim wsNew As Worksheet
    Dim i As Long
    Dim isBigChecked As Boolean
    Dim checkCell As String
    
    Set wb = ThisWorkbook
    Dim wsData As Worksheet
    Set wsData = wb.Worksheets("Data")

    ' Unprotect workbook structure so Copy lands in this workbook
    On Error Resume Next
    wb.Unprotect
    On Error GoTo 0

    ' Read checkbox once from Data!D43
    isBigChecked = (wsData.Range("D43").Value = True)
    
    For i = 1 To 7
        checkCell = "D" & (47 + 2 * (i - 1))
        If wsData.Range(checkCell).Value = True Then
            If IsBigSemesterLayout(i, isBigChecked) Then
                templateName = "TemplateBig"
            Else
                templateName = "TemplateSmall"
            End If
            
            Set wsNew = CopySemesterTemplate(wb, templateName)
            If Not SheetExists(i & ". Sem", wb) Then
                wsNew.Name = i & ". Sem"
                wsNew.Visible = xlSheetVisible
            Else
                Application.DisplayAlerts = False
                wsNew.Delete
                Application.DisplayAlerts = True
            End If
            
        End If
    Next i
End Sub

Private Function CopySemesterTemplate(ByVal wb As Workbook, ByVal templateName As String) As Worksheet
    Dim names() As String, sheet As Worksheet, i As Long, oldCount As Long, found As Boolean
    oldCount = wb.Worksheets.Count
    ReDim names(1 To oldCount)
    i = 0
    For Each sheet In wb.Worksheets
        i = i + 1
        names(i) = sheet.Name
    Next sheet
    wb.Worksheets(templateName).Copy After:=wb.Worksheets(templateName)
    If wb.Worksheets.Count <> oldCount + 1 Then _
        Err.Raise 5, "CopySemesterTemplate", "The template copy did not create exactly one worksheet."
    For Each sheet In wb.Worksheets
        found = False
        For i = 1 To oldCount
            If sheet.Name = names(i) Then
                found = True
                Exit For
            End If
        Next i
        If Not found Then
            Set CopySemesterTemplate = sheet
            Exit Function
        End If
    Next sheet
    Err.Raise 5, "CopySemesterTemplate", "The copied semester worksheet could not be identified."
End Function

'=============================================================
' Main entry point: creates semesters and fills 1. & 2. sem
'=============================================================
Public Sub Generate_Semesters()
    Dim wb As Workbook
    Dim isBigChecked As Boolean
    Dim semNumber As Long
    Dim checkCell As String
    Dim isBigLayout As Boolean
    Dim oldEvents As Boolean, oldScreen As Boolean, oldCalculation As XlCalculation
    Dim failure As String, stage As String
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldCalculation = Application.Calculation
    On Error GoTo Failed
    Application.EnableEvents = False
    
    Set wb = ThisWorkbook
    Dim wsData As Worksheet
    Set wsData = wb.Worksheets("Data")
    
    ' Read checkbox once (Data!D43)
    isBigChecked = (wsData.Range("D43").Value = True)
    
    ' Create all semesters (only checked ones)
    stage = "Copying semester templates"
    MakeSemesters
    
    ' Fill semesters that are checked
    For semNumber = 1 To 7
        checkCell = "D" & (47 + 2 * (semNumber - 1))
        If wsData.Range(checkCell).Value = True Then
            stage = "Filling semester " & semNumber
            FillSemester semNumber, isBigChecked
            FillSemesterTitles semNumber, isBigChecked
            FillSemesterConditionalFormat semNumber, isBigChecked
            ApplyPlottedHoursConditionalFormat semNumber, isBigChecked
            isBigLayout = IsBigSemesterLayout(semNumber, isBigChecked)
            If SheetExists(semNumber & ". Sem", wb) Then
                ApplyNoTeachingDaysToSheet wb.Worksheets(semNumber & ". Sem"), isBigLayout
                With wb.Worksheets(semNumber & ". Sem")
                    .Activate
                    ActiveWindow.FreezePanes = False
                    ActiveWindow.SplitRow = 0
                    ActiveWindow.SplitColumn = 0
                    ActiveWindow.ScrollRow = 1
                    ActiveWindow.ScrollColumn = 1
                    .Range("E1").Select
                    ActiveWindow.FreezePanes = True
                End With
            End If
        End If
    Next semNumber

    stage = "Building Overblik"
    GatherSchedulesToOverblik
    Application.Calculation = oldCalculation
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Exit Sub
Failed:
    failure = stage & ": " & Err.Description
    Application.Calculation = oldCalculation
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Err.Raise 5, "Generate_Semesters", failure
End Sub
'=============================================================
' GENERIC: Fill any semester with dropdown lists & formulas
'=============================================================
Private Sub FillSemester(semNumber As Long, isBigChecked As Boolean)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsData As Worksheet
    Dim rngList As Range
    Dim validRanges As String, localListRange As String, formulaC As String, dataSourceRows As String
    
    ' Configuration based on semester number
    Dim cfLegendRange As String, sem As String
    cfLegendRange = Chr(64 + semNumber + 5) & "3:" & Chr(64 + semNumber + 5) & "40"
    sem = semNumber & ". Sem"
    
    Set wb = ThisWorkbook
    Set wsData = wb.Worksheets("Data")
    
    If Not SheetExists(sem, wb) Then Exit Sub
    Set ws = wb.Worksheets(sem)
    
    ' Determine validation ranges and list range based on checkbox and semester
    ' When checkbox TRUE: Sem1=big, Sem2=small
    ' When checkbox FALSE: Sem1=small, Sem2=big
    ' Data source is always semester-specific
    Select Case semNumber
        Case 1
            dataSourceRows = "'Ressourcer'!$B$7:$X$10"
        Case 2
            dataSourceRows = "'Ressourcer'!$B$12:$X$15"
        Case 3
            dataSourceRows = "'Ressourcer'!$B$17:$X$20"
        Case 4
            dataSourceRows = "'Ressourcer'!$B$23:$X$25"
        Case Else
            dataSourceRows = "" ' No data for other semesters
    End Select
    If dataSourceRows = "" Then Exit Sub
    
    Dim isBigLayout As Boolean
    isBigLayout = IsBigSemesterLayout(semNumber, isBigChecked)
    
    If isBigLayout Then
        ' BIG layout
        validRanges = "E7:DE15,E27:DE35"
        localListRange = "B45:B59"
    Else
        ' SMALL layout
        validRanges = "E7:DE10,E12:DE15"
        localListRange = "B24:B38"
    End If
    
    ' Apply dropdown validation for data values
    Set rngList = wsData.Range(cfLegendRange)
    ApplyListValidation ws, validRanges, rngList
    
    ' Apply dropdown validation for lookup values
    Set rngList = ws.Range(localListRange)
    ' Use same layout logic for dropdown ranges
    If isBigLayout Then
        ' Keep row 36 aligned with the topic row pattern (like row 15 in top block).
        ApplyListValidation ws, "E16:DE21,E37:DE42", rngList
    Else
        ApplyListValidation ws, "E16:DE21", rngList
    End If
    
    ' Set formulas for student lookup (B column)
    ws.Range(localListRange).FormulaLocal = _
        "=IFERROR(INDEX(FILTER('Ressourcer'!$B$1:$X$1;BYCOL(" & dataSourceRows & ";LAMBDA(c;COUNTIF(c;""<>"" )>0)));ROWS($B$2:B2));"""")"
    
    ' Set formulas for role mapping (C column) - sum multiple hours per teacher
    formulaC = "=IFERROR(" & _
            "INDEX(" & _
                "MAP(" & _
                    "FILTER(COLUMN('Ressourcer'!$B$1:$X$1);" & _
                           "BYCOL(" & dataSourceRows & ";LAMBDA(c;COUNTIF(c;""<>"" )>0))" & _
                    ");" & _
                    "LAMBDA(col;" & _
                        "SUM(" & _
                            "INDEX(" & dataSourceRows & ";;col-COLUMN(" & dataSourceRows & ")+1)" & _
                        ")" & _
                    ")" & _
                ");" & _
                "ROWS($B$2:B2)" & _
            ")/2;"""")"
    ws.Range(localListRange).Offset(0, 1).FormulaLocal = formulaC
    
    ' Set formulas for schedule counting (D column) - count filled slots per teacher
    Dim dRange As Range
    Dim formulaD As String
    If isBigLayout Then
        Set dRange = ws.Range("E45:DE59")
    Else
        Set dRange = ws.Range("E24:DE38")
    End If
    
    ' Unprotect sheet if protected (assume no password)
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0
    
    ' Turn off calculation and screen updating to prevent freeze
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    Dim cell As Range
    Dim colLetter As String

    ' Build D-column formulas cell-by-cell so each column references its own
    ' schedule block (top only for small, top+bottom for big).
    For Each cell In dRange
        ' Clear any existing content first
        On Error Resume Next
        cell.ClearContents
        If Err.Number <> 0 Then
            Err.Clear
            cell.Value = "Cannot modify cell"
            GoTo NextCell
        End If
        
        ' Get column letter using column number
        colLetter = ColumnLetter(cell.Column)
        
        ' Build formula for this row and column
        If isBigLayout Then
            formulaD = "=IF($B" & cell.Row & "="""";"""";IF(COUNTIF(" & colLetter & "$16:" & colLetter & "$21;$B" & cell.Row & ")>0;COUNTA(" & colLetter & "$7:" & colLetter & "$15);0)+IF(COUNTIF(" & colLetter & "$37:" & colLetter & "$42;$B" & cell.Row & ")>0;COUNTA(" & colLetter & "$27:" & colLetter & "$35);0))"
        Else
            formulaD = "=IF($B" & cell.Row & "="""";"""";IF(COUNTIF(" & colLetter & "$16:" & colLetter & "$21;$B" & cell.Row & ")>0;COUNTA(" & colLetter & "$7:" & colLetter & "$15);0))"
        End If
        
        cell.FormulaLocal = formulaD
        
        If Err.Number <> 0 Then
            ' If formula fails, set a simple message
            cell.Value = "Formula failed"
            Err.Clear
        End If
        
NextCell:
        On Error GoTo 0
    Next cell
    
    ' Turn calculation and screen updating back on
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Dim semCodeA As String
    Dim semCodeB As String

    semCodeA = GetSemesterTeamCode(semNumber, 1)
    semCodeB = GetSemesterTeamCode(semNumber, 2)

    If semCodeA = "" Then
        semCodeA = CStr(wsData.Range("C" & (47 + (semNumber - 1) * 2)).Value)
    End If

    If isBigLayout Then
        ' This semester uses BIG layout
        ws.Range("B2").Value = semCodeA
        If semCodeB <> "" Then
            ws.Range("B23").Value = semCodeB
        Else
            ws.Range("B23").Value = semCodeA
        End If
        SetSemesterRoomHeader ws.Range("B3"), semNumber, " A"
        SetSemesterRoomHeader ws.Range("B24"), semNumber, " B"
    Else
        ' This semester uses SMALL layout
        ws.Range("B2").Value = semCodeA
        SetSemesterRoomHeader ws.Range("B3"), semNumber, ""
        ' Do not set B23 for small layout
    End If
End Sub

Private Function SemesterRoomLabel(ByVal semNumber As Long, ByVal groupSuffix As String) As String
    Dim roomValue As Variant, room As String
    roomValue = ThisWorkbook.Worksheets("Data").Cells(47 + (semNumber - 1) * 2, 7).Value2
    If IsError(roomValue) Then Err.Raise 5, "SemesterRoomLabel", "Invalid room entry for semester " & semNumber
    room = Trim$(CStr(roomValue))
    SemesterRoomLabel = semNumber & ". Sem" & groupSuffix
    If room <> "" Then SemesterRoomLabel = SemesterRoomLabel & " / " & room
End Function

Private Sub SetSemesterRoomHeader(ByVal target As Range, ByVal semNumber As Long, ByVal groupSuffix As String)
    target.Value2 = SemesterRoomLabel(semNumber, groupSuffix)
    target.MergeArea.Font.Size = 22
End Sub

Public Sub RefreshSemesterRoomHeaders()
    Dim semNumber As Long, ws As Worksheet, isBigChecked As Boolean
    isBigChecked = (ThisWorkbook.Worksheets("Data").Range("D43").Value2 = True)
    For semNumber = 1 To 7
        If SheetExists(semNumber & ". Sem", ThisWorkbook) Then
            Set ws = ThisWorkbook.Worksheets(semNumber & ". Sem")
            If IsBigSemesterLayout(semNumber, isBigChecked) Then
                SetSemesterRoomHeader ws.Range("B3"), semNumber, " A"
                SetSemesterRoomHeader ws.Range("B24"), semNumber, " B"
            Else
                SetSemesterRoomHeader ws.Range("B3"), semNumber, ""
            End If
        End If
    Next semNumber
End Sub

Private Function GetSemesterTeamCode(ByVal semNumber As Long, ByVal teamIndex As Long) As String
    Dim wsData As Worksheet
    Dim baseRow As Long
    Dim rowNumber As Long
    Dim r As Long
    Dim currentSem As Long
    Dim semValA As Variant
    Dim semValB As Variant
    Dim codeVal As String
    Dim hitCount As Long

    Set wsData = ThisWorkbook.Worksheets("Data")

    If teamIndex < 1 Or teamIndex > 2 Then Exit Function

    ' Primary mapping: Data table starts at row 47 (row 46 is header).
    ' Two rows per semester:
    '   Sem 1 -> rows 47-48
    '   Sem 2 -> rows 49-50
    '   Sem 3 -> rows 51-52
    '   ...
    baseRow = 47 + (semNumber - 1) * 2
    rowNumber = baseRow + (teamIndex - 1)
    If rowNumber >= 46 And rowNumber <= 60 Then
        codeVal = GetClassCodeFromRow(wsData, rowNumber)
        If codeVal <> "" Then
            GetSemesterTeamCode = codeVal
            Exit Function
        End If
    End If

    ' Fallback mapping: support marker-based tables where semester appears in A or B.
    For r = 46 To 60
        semValA = wsData.Cells(r, 1).Value
        semValB = wsData.Cells(r, 2).Value

        If IsNumeric(semValA) Then
            currentSem = CLng(semValA)
        ElseIf IsNumeric(semValB) Then
            currentSem = CLng(semValB)
        End If

        If currentSem = semNumber Then
            codeVal = GetClassCodeFromRow(wsData, r)
            If codeVal <> "" Then
                hitCount = hitCount + 1
                If hitCount = teamIndex Then
                    GetSemesterTeamCode = codeVal
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

Private Function GetClassCodeFromRow(ByVal wsData As Worksheet, ByVal rowNumber As Long) As String
    Dim codeVal As String

    codeVal = CleanCellText(wsData.Cells(rowNumber, 3).Value) ' C
    If IsLikelyClassCode(codeVal) Then
        GetClassCodeFromRow = codeVal
        Exit Function
    End If

    codeVal = CleanCellText(wsData.Cells(rowNumber, 2).Value) ' B
    If IsLikelyClassCode(codeVal) Then
        GetClassCodeFromRow = codeVal
        Exit Function
    End If

    codeVal = CleanCellText(wsData.Cells(rowNumber, 4).Value) ' D
    If IsLikelyClassCode(codeVal) Then
        GetClassCodeFromRow = codeVal
    End If
End Function

Private Function CleanCellText(ByVal value As Variant) As String
    CleanCellText = Trim$(CStr(value))
End Function

Private Function IsLikelyClassCode(ByVal valueText As String) As Boolean
    If valueText = "" Then Exit Function
    If UCase$(valueText) = "TRUE" Or UCase$(valueText) = "FALSE" Then Exit Function
    If IsNumeric(valueText) Then Exit Function

    IsLikelyClassCode = True
End Function

'=============================================================
' DEPRECATED _ These are now consolidated into FillSemester
'=============================================================
' (Keeping stubs for backwards compatibility if needed)
Public Sub FillFirstBig()
    FillSemester 1, True
End Sub

Public Sub FillSecondSmall()
    FillSemester 2, True
End Sub

Public Sub FillSecondBig()
    FillSemester 2, False
End Sub

Public Sub FillFirstSmall()
    FillSemester 1, False
End Sub

'=============================================================
' GENERIC: Fill semester titles from Data table + copy if needed
'=============================================================
Private Sub FillSemesterTitles(semNumber As Long, isBigChecked As Boolean)
    Dim wb As Workbook, wsData As Worksheet, wsPlan As Worksheet
    Dim firstRow As Long, lastRow As Long, r As Long, colRow As Long, i As Long
    Dim code As String, title As String, weeks As Long
    Dim codeCol As Long, titleCol As Long, weeksCol As Long
    Dim startRow As Long, sem As String
    Dim cfLegendRange As String, codePrefix As String
    Dim isBigLayout As Boolean
    
    ' Configuration
    Select Case semNumber
        Case 1
            sem = "1. Sem"
            cfLegendRange = "F3:F40"
            codePrefix = "1."
        Case 2
            sem = "2. Sem"
            cfLegendRange = "G3:G40"
            codePrefix = "2."
        Case Else
            sem = semNumber & ". Sem"
            cfLegendRange = Chr(64 + semNumber + 5) & "3:" & Chr(64 + semNumber + 5) & "40"
            codePrefix = semNumber & "."
    End Select
    
    Set wb = ThisWorkbook
    Set wsData = wb.Worksheets("Data")
    
    On Error Resume Next
    Set wsPlan = wb.Worksheets(sem)
    On Error GoTo 0
    
    If wsPlan Is Nothing Then Exit Sub

    isBigLayout = IsBigSemesterLayout(semNumber, isBigChecked)
    
    firstRow = 63
    lastRow = 86
    
    ' Get column numbers
    codeCol = 2   ' B
    titleCol = 3  ' C
    weeksCol = 4  ' D
    
    ' Clear title row
    wsPlan.Range("E3:DE3").ClearContents
    wsPlan.Range("E3:DE3").Interior.ColorIndex = xlNone
    
    colRow = 5  ' E
    
    ' Find first matching code (doesn't assume x.0, just grab first x.*)
    startRow = 0
    For r = firstRow To lastRow
        code = Trim$(CStr(wsData.Cells(r, codeCol).Value))
        If Left$(code, Len(codePrefix)) = codePrefix Then
            startRow = r
            Exit For
        End If
    Next r
    
    If startRow = 0 Then
        ' No code for this semester, skip silently
        Exit Sub
    End If
    
    ' Fill titles
    For r = startRow To lastRow
        code = Trim$(CStr(wsData.Cells(r, codeCol).Value))
        
        If code = "" Then Exit For
        If Left$(code, Len(codePrefix)) <> codePrefix Then Exit For
        
        title = CStr(wsData.Cells(r, titleCol).Value)
        
        If IsNumeric(wsData.Cells(r, weeksCol).Value) Then
            weeks = CLng(wsData.Cells(r, weeksCol).Value)
        Else
            weeks = 0
        End If
        
        If weeks > 0 Then
            For i = 1 To weeks
                With wsPlan.Cells(3, colRow)
                    .Value = code & " " & title
                    .Interior.Color = wsData.Cells(r, titleCol).Interior.Color
                End With
                colRow = colRow + 5
            Next i
        End If
    Next r
    
    ' Clear row 24 first so small semesters do not keep old title/formula content
    ' Only clear for big semesters; small semesters need formulas in row 24
    If isBigLayout Then
        wsPlan.Range("E24:DE24").ClearContents
        wsPlan.Range("E24:DE24").Interior.ColorIndex = xlNone
    End If

    ' Copy to row 24 only for BIG semesters
    If isBigLayout Then
        wsPlan.Range("E3:DE3").Copy Destination:=wsPlan.Range("E24:DE24")
    End If
End Sub

'=============================================================
' GENERIC: Apply conditional formatting to semester
'=============================================================
Private Sub FillSemesterConditionalFormat(semNumber As Long, isBigChecked As Boolean)
    Dim wb As Workbook, wsData As Worksheet, wsPlan As Worksheet
    Dim c As Range, fc As FormatCondition
    Dim sem As String, cfLegendRange As String
    Dim isBigLayout As Boolean
    
    ' Configuration
    Select Case semNumber
        Case 1
            sem = "1. Sem"
            cfLegendRange = "F3:F40"
        Case 2
            sem = "2. Sem"
            cfLegendRange = "G3:G40"
        Case Else
            sem = semNumber & ". Sem"
            cfLegendRange = Chr(64 + semNumber + 5) & "3:" & Chr(64 + semNumber + 5) & "40"
    End Select
    
    ' Determine if this semester uses big layout
    isBigLayout = IsBigSemesterLayout(semNumber, isBigChecked)
    
    Set wb = ThisWorkbook
    Set wsData = wb.Worksheets("Data")
    If Not SheetExists(sem, wb) Then Exit Sub
    Set wsPlan = wb.Worksheets(sem)
    
    ' Always apply CF to top ranges
    With wsPlan.Range("E7:DE10,E12:DE15")
        .FormatConditions.Delete
        For Each c In wsData.Range(cfLegendRange)
            If Trim$(CStr(c.Value)) <> "" Then
                Set fc = .FormatConditions.Add( _
                    Type:=xlTextString, _
                    String:=CStr(c.Value), _
                    TextOperator:=xlContains)
                fc.Interior.Color = c.Interior.Color
                fc.Font.Color = vbBlack
            End If
        Next c
    End With
    
    ' Apply CF to bottom ranges if this semester uses big layout
    If isBigLayout Then
        With wsPlan.Range("E27:DE35")
            .FormatConditions.Delete
            For Each c In wsData.Range(cfLegendRange)
                If Trim$(CStr(c.Value)) <> "" Then
                    Set fc = .FormatConditions.Add( _
                        Type:=xlTextString, _
                        String:=CStr(c.Value), _
                        TextOperator:=xlContains)
                    fc.Interior.Color = c.Interior.Color
                    fc.Font.Color = vbBlack
                End If
            Next c
        End With
    End If
End Sub

Private Sub ApplyPlottedHoursConditionalFormat(semNumber As Long, isBigChecked As Boolean)
    Dim wsPlan As Worksheet
    Dim sem As String
    Dim isBigLayout As Boolean
    Dim plottedRange As Range
    Dim firstRow As Long

    sem = semNumber & ". Sem"
    If Not SheetExists(sem, ThisWorkbook) Then Exit Sub

    Set wsPlan = ThisWorkbook.Worksheets(sem)
    isBigLayout = IsBigSemesterLayout(semNumber, isBigChecked)

    If isBigLayout Then
        Set plottedRange = wsPlan.Range("D45:D59")
        firstRow = 45
    Else
        Set plottedRange = wsPlan.Range("D24:D38")
        firstRow = 24
    End If

    plottedRange.FormatConditions.Delete

    Dim fcGreen As FormatCondition, fcRed As FormatCondition

    ' Green when plotted hours are within +/- 10 of allocated hours.
    Set fcGreen = plottedRange.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=ABS($D" & firstRow & "-$C" & firstRow & ")<=10")
    fcGreen.Interior.Color = RGB(198, 239, 206)
    fcGreen.Font.Color = RGB(0, 97, 0)
    fcGreen.StopIfTrue = True

    ' Red when plotted hours differ from allocated hours by more than 10.
    Set fcRed = plottedRange.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=ABS($D" & firstRow & "-$C" & firstRow & ")>10")
    fcRed.Interior.Color = RGB(255, 199, 206)
    fcRed.Font.Color = RGB(156, 0, 6)
End Sub

'=============================================================
' DEPRECATED _ Kept for backwards compatibility
'=============================================================
Public Sub FillFirstTitles()
    Dim isBigChecked As Boolean
    isBigChecked = (ThisWorkbook.Worksheets("Data").Range("D43").Value = True)
    FillSemesterTitles 1, isBigChecked
End Sub

Public Sub FillCndFirst()
    Dim isBigChecked As Boolean
    isBigChecked = (ThisWorkbook.Worksheets("Data").Range("D43").Value = True)
    FillSemesterConditionalFormat 1, isBigChecked
End Sub

'=============================================================
' Helper function to convert column number to letter
'=============================================================
Private Function ColumnLetter(col As Long) As String
    Dim s As String
    Do While col > 0
        col = col - 1
        s = Chr(65 + (col Mod 26)) & s
        col = col \ 26
    Loop
    ColumnLetter = s
End Function


'=============================================================
' Mark no-teaching days in generated semester sheets
'=============================================================
Public Sub MarkNoTeachingDays()
    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim semNumber As Long
    Dim checkCell As String
    Dim isBigChecked As Boolean

    Set wb = ThisWorkbook
    Set wsData = wb.Worksheets("Data")
    isBigChecked = (wsData.Range("D43").Value = True)

    For semNumber = 1 To 7
        checkCell = "D" & (47 + 2 * (semNumber - 1))
        If wsData.Range(checkCell).Value = True Then
            If SheetExists(semNumber & ". Sem", wb) Then
                ApplyNoTeachingDaysToSheet wb.Worksheets(semNumber & ". Sem"), IsBigSemesterLayout(semNumber, isBigChecked)
            End If
        End If
    Next semNumber
End Sub

Private Sub ApplyNoTeachingDaysToSheet(wsTarget As Worksheet, isBigLayout As Boolean)
    Dim blockedDates As Range

    Set blockedDates = ThisWorkbook.Worksheets("Data").Range("M3:M35")

    ApplyBlockedDateFill wsTarget.Range("E5:DE5"), wsTarget.Range("E7:DE15"), blockedDates, isBigLayout, wsTarget.Range("E28:DE36")
End Sub

Private Sub ApplyBlockedDateFill(headerRange As Range, fillRange As Range, blockedDates As Range, _
                                 Optional applyToBottomBlock As Boolean = False, _
                                 Optional bottomFillRange As Range = Nothing)
    Dim headerCell As Range
    Dim blockedCell As Range
    Dim relativeCol As Long
    Dim headerDateKey As String
    Dim blockedDateKey As String

    For Each headerCell In headerRange.Cells
        headerDateKey = DateKeyFromCell(headerCell)
        If headerDateKey <> "" Then
            For Each blockedCell In blockedDates.Cells
                blockedDateKey = DateKeyFromValue(blockedCell.Value)
                If blockedDateKey <> "" And blockedDateKey = headerDateKey Then
                    relativeCol = headerCell.Column - fillRange.Column + 1
                    fillRange.Columns(relativeCol).Interior.Color = RGB(192, 0, 0)
                    If applyToBottomBlock Then
                        bottomFillRange.Columns(relativeCol).Interior.Color = RGB(192, 0, 0)
                    End If
                    Exit For
                End If
            Next blockedCell
        End If
    Next headerCell
End Sub

Private Function DateKeyFromCell(sourceCell As Range) As String
    DateKeyFromCell = DateKeyFromValue(sourceCell.Value)

    If DateKeyFromCell = "" Then
        DateKeyFromCell = DateKeyFromValue(sourceCell.Text)
    End If
End Function

Private Function DateKeyFromValue(sourceValue As Variant) As String
    On Error Resume Next

    If IsDate(sourceValue) Then
        DateKeyFromValue = Format$(CDate(sourceValue), "yyyymmdd")
    Else
        DateKeyFromValue = ""
    End If

    On Error GoTo 0
End Function
