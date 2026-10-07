Attribute VB_Name = "Delete"
Option Explicit

Public Sub DeleteAllSemesters()
    Dim wb As Workbook
    Dim i As Long
    Dim semName As String
    Dim prevDisplayAlerts As Boolean

    On Error GoTo Failed
    
    Set wb = ThisWorkbook  ' workbook that holds your code

    prevDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False  ' avoid Excel "Are you sure?" prompts
    
    For i = 1 To 7
        semName = i & ". Sem"
        If SheetExists(semName, wb) Then
            wb.Worksheets(semName).Delete
        End If
    Next i
    PrepareOverviewReferences
    If SheetExists("Overblik", wb) Then wb.Worksheets("Overblik").Delete

    Application.DisplayAlerts = prevDisplayAlerts
    Exit Sub
Failed:
    Application.DisplayAlerts = prevDisplayAlerts
    Err.Raise Err.Number, "DeleteAllSemesters", Err.Description
End Sub

Public Sub PrepareOverviewReferences()
    Dim ws As Worksheet, cell As Range, formula As String
    If Not SheetExists("Dobbeltbooking", ThisWorkbook) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets("Dobbeltbooking")
    For Each cell In ws.UsedRange
        If cell.HasFormula Then
            formula = cell.Formula2
            If InStr(1, formula, "INDIRECT(""'Overblik'!", vbTextCompare) = 0 Then
                formula = Replace(formula, "'Overblik'!$C$4:$DC$4", "INDIRECT(""'Overblik'!$C$4:$DC$4"")")
                formula = Replace(formula, "'Overblik'!$C$6:$DC$1000", "INDIRECT(""'Overblik'!$C$6:$DC$1000"")")
                formula = Replace(formula, "Overblik!$C$4:$DC$4", "INDIRECT(""'Overblik'!$C$4:$DC$4"")")
                formula = Replace(formula, "Overblik!$C$6:$DC$1000", "INDIRECT(""'Overblik'!$C$6:$DC$1000"")")
            End If
            If formula <> cell.Formula2 Then cell.Formula2 = formula
        End If
    Next cell
End Sub
