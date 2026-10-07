Attribute VB_Name = "modUtils"
Option Explicit

'================= SHEET UTILITIES ============================
Public Function SheetExists(sName As String, Optional wb As Workbook) As Boolean
    ' Checks if a worksheet exists in the provided workbook.
    ' If no workbook is supplied, ThisWorkbook is used.
    Dim ws As Worksheet
    
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    On Error Resume Next
    Set ws = wb.Worksheets(sName)
    SheetExists = Not ws Is Nothing
    Set ws = Nothing
    On Error GoTo 0
End Function

Public Sub OrderWorkbookTabs()
    Dim wb As Workbook, previous As Worksheet, ws As Worksheet, names As Variant, name As Variant
    Set wb = ThisWorkbook
    names = Array("Ressourcer", "Data", "Overblik", "Dobbeltbooking", _
                  "1. Sem", "2. Sem", "3. Sem", "4. Sem", "5. Sem", "6. Sem", "7. Sem")
    For Each name In names
        If SheetExists(CStr(name), wb) Then
            Set ws = wb.Worksheets(CStr(name))
            If previous Is Nothing Then
                If ws.Index <> 1 Then ws.Move Before:=wb.Worksheets(1)
            Else
                If ws.Index <> previous.Index + 1 Then ws.Move After:=previous
            End If
            Set previous = ws
        End If
    Next name
End Sub

'================= DATA VALIDATION HELPER =====================
Public Sub ApplyListValidation(targetWs As Worksheet, targetAddress As String, sourceRange As Range)
    ' Applies list validation using a workbook-scoped source range.
    Dim formulaList As String
    
    'Build the list formula: ='SheetName'!$F$3:$F$40
    formulaList = "='" & sourceRange.Parent.Name & "'!" & sourceRange.Address
    
    With targetWs.Range(targetAddress)
        On Error Resume Next
        .Validation.Delete
        On Error GoTo 0
        
        .Validation.Add Type:=xlValidateList, _
                        AlertStyle:=xlValidAlertStop, _
                        Operator:=xlBetween, _
                        Formula1:=formulaList
        .Validation.IgnoreBlank = True
        .Validation.InCellDropdown = True
    End With
End Sub

'================= SEMESTER LAYOUT RULES ======================
Public Function IsBigSemesterLayout(ByVal semNumber As Long, ByVal isBigChecked As Boolean) As Boolean
    ' Central rule used across modules:
    ' - TRUE in Data!D43 => odd semesters use TemplateBig.
    ' - FALSE in Data!D43 => even semesters use TemplateBig.
    If isBigChecked Then
        IsBigSemesterLayout = (semNumber Mod 2 = 1)
    Else
        IsBigSemesterLayout = (semNumber Mod 2 = 0)
    End If
End Function



