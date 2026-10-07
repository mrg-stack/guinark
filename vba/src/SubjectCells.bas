Attribute VB_Name = "SubjectCells"

Option Explicit

Sub ColorMatchesFromD()
    Dim ws As Worksheet
    Dim rngSource As Range
    Dim rngTarget As Range
    Dim cSrc As Range
    Dim cTgt As Range
    
    ' Use the active sheet _ change to ThisWorkbook.Worksheets("Data") if you want to fix it
    Set ws = ActiveSheet
    
    ' Source values + colours (originals in column D)
    Set rngSource = ws.Range("D3:D40")
    
    ' Target area to recolour
    Set rngTarget = ws.Range("F3:J40")
    
    ' Loop source cells
    For Each cSrc In rngSource
        If Trim$(CStr(cSrc.Value)) <> "" Then
            ' For each source value, scan the target range
            For Each cTgt In rngTarget
                If CStr(cTgt.Value) = CStr(cSrc.Value) Then
                    cTgt.Interior.Color = cSrc.Interior.Color
                End If
            Next cTgt
        End If
    Next cSrc
End Sub

Public Sub ApplyDataCodeColorsToSchedule()
    ApplyDataCodeColorsToWorksheet ActiveSheet
End Sub

Public Sub ApplyDataCodeConditionalFormattingToSchedule()
    Dim wsData As Worksheet
    Dim wsTarget As Worksheet
    Dim sourceCodes As Range
    Dim targetCells As Range
    Dim c As Range
    Dim codeText As String
    Dim fc As FormatCondition

    On Error GoTo CFFail

    Set wsData = ThisWorkbook.Worksheets("Data")
        Set wsTarget = ActiveSheet
        Set sourceCodes = wsData.Range("D3:D40")
        Set targetCells = GetScheduleTargetRange(wsTarget)

        If targetCells Is Nothing Then
            MsgBox "No schedule blocks found on " & wsTarget.Name & ".", vbExclamation
            Exit Sub
        End If

    SuppressUnlockedFormulaWarningsForRange targetCells

    ' Reset previous color rules, then create one rule per code from Data.
    targetCells.FormatConditions.Delete
    targetCells.Interior.Pattern = xlNone

    For Each c In sourceCodes.Cells
        codeText = Trim$(CStr(c.Value))
        If codeText <> vbNullString Then
            Set fc = targetCells.FormatConditions.Add( _
                        Type:=xlTextString, _
                        String:=codeText, _
                        TextOperator:=xlContains)
            fc.Interior.Color = c.Interior.Color
            fc.StopIfTrue = False
        End If
    Next c

    MsgBox "Conditional formatting linked to Data codes on " & wsTarget.Name & ".", vbInformation
    Exit Sub

CFFail:
    MsgBox "Could not apply conditional formatting: " & Err.Description, vbExclamation
End Sub

Public Sub SuppressUnlockedFormulaWarningsOnActiveSchedule()
    On Error GoTo WarnFail

        Dim targetCells As Range
        Set targetCells = GetScheduleTargetRange(ActiveSheet)
        If targetCells Is Nothing Then
            MsgBox "No schedule blocks found on " & ActiveSheet.Name & ".", vbExclamation
            Exit Sub
        End If

        SuppressUnlockedFormulaWarningsForRange targetCells
        MsgBox "Unlocked-formula warnings disabled for " & ActiveSheet.Name & " " & targetCells.Address(False, False) & ".", vbInformation
    Exit Sub

WarnFail:
    MsgBox "Could not suppress warnings: " & Err.Description, vbExclamation
End Sub

Public Sub DisableUnlockedFormulaWarningsGlobally()
    On Error GoTo GlobalFail

    Application.ErrorCheckingOptions.UnlockedFormulaCells = False
    MsgBox "Excel unlocked-formula warnings disabled globally.", vbInformation
    Exit Sub

GlobalFail:
    MsgBox "Could not update global warning settings: " & Err.Description, vbExclamation
End Sub

Private Sub SuppressUnlockedFormulaWarningsForRange(ByVal targetCells As Range)
    On Error Resume Next
    targetCells.Errors(xlUnlockedFormulaCells).Ignore = True
    On Error GoTo 0
End Sub

Public Sub ApplyDataCodeColorsToCalculatedSheet(ByVal Sh As Object)
    On Error GoTo SafeExit

    If TypeName(Sh) <> "Worksheet" Then Exit Sub
    If StrComp(CStr(Sh.Name), "Data", vbTextCompare) = 0 Then Exit Sub

    ApplyDataCodeColorsToWorksheet Sh

SafeExit:
End Sub

Public Sub InstallScheduleAutoColoringHook()
    Dim vbComp As Object
    Dim codeModule As Object
    Dim allCode As String
    Dim hookCode As String
    Dim insertAtLine As Long

    On Error GoTo InstallErr

    ' Use workbook codename instead of literal "ThisWorkbook".
    Set vbComp = ThisWorkbook.VBProject.VBComponents(ThisWorkbook.CodeName)
    Set codeModule = vbComp.CodeModule

    If codeModule.CountOfLines > 0 Then
        allCode = codeModule.Lines(1, codeModule.CountOfLines)
    Else
        allCode = vbNullString
    End If

    If InStr(1, allCode, "Private Sub Workbook_SheetCalculate(ByVal Sh As Object)", vbTextCompare) > 0 Then
        MsgBox "Workbook calculate hook already installed.", vbInformation
        Exit Sub
    End If

    hookCode = "Private Sub Workbook_SheetCalculate(ByVal Sh As Object)" & vbCrLf & _
               "    ApplyDataCodeColorsToCalculatedSheet Sh" & vbCrLf & _
               "End Sub" & vbCrLf

    insertAtLine = codeModule.CountOfLines + 1
    If insertAtLine < 1 Then insertAtLine = 1
    codeModule.InsertLines insertAtLine, hookCode

    MsgBox "Workbook calculate hook installed.", vbInformation
    Exit Sub

InstallErr:
    MsgBox "Could not install workbook hook: " & Err.Description, vbExclamation
End Sub

Public Sub ApplyDataCodeColorsToWorksheet(ByVal wsTarget As Worksheet)
    Dim wsData As Worksheet
    Dim sourceCodes As Range
    Dim targetCells As Range
    Dim colorByCode As Object
    Dim c As Range
    Dim key As String

    On Error GoTo SafeExit

    Set wsData = ThisWorkbook.Worksheets("Data")

    ' Data!D3:D40 contains code labels with the desired fill colors.
    Set sourceCodes = wsData.Range("D3:D40")
        Set targetCells = GetScheduleTargetRange(wsTarget)
        If targetCells Is Nothing Then Exit Sub

    Set colorByCode = BuildColorMap(sourceCodes)

    For Each c In targetCells.Cells
        key = UCase$(Trim$(CStr(c.Value)))
        If key = vbNullString Then
            c.Interior.Pattern = xlNone
        ElseIf colorByCode.Exists(key) Then
            c.Interior.Color = CLng(colorByCode(key))
        End If
    Next c

SafeExit:
End Sub

    Private Function GetScheduleTargetRange(ByVal ws As Worksheet) As Range
        Dim firstRow As Long
        Dim lastLabelRow As Long
        Dim lastBlockBottom As Long

        firstRow = 6
        lastLabelRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Nothing meaningful in column A below the header area.
        If lastLabelRow < firstRow Then Exit Function

        ' Each schedule block label sits in column A and the block spans 15 rows.
        ' Example: A6 covers rows 6:20, A22 covers rows 22:36, etc.
        lastBlockBottom = lastLabelRow + 14
        If lastBlockBottom < firstRow Then Exit Function

        Set GetScheduleTargetRange = ws.Range("C" & firstRow & ":DC" & lastBlockBottom)
    End Function

Private Function BuildColorMap(ByVal codeRange As Range) As Object
    Dim dict As Object
    Dim c As Range
    Dim key As String

    Set dict = CreateObject("Scripting.Dictionary")

    For Each c In codeRange.Cells
        key = UCase$(Trim$(CStr(c.Value)))
        If key <> vbNullString Then
            If Not dict.Exists(key) Then
                dict.Add key, c.Interior.Color
            End If
        End If
    Next c

    Set BuildColorMap = dict
End Function

