Attribute VB_Name = "modResourceLayout"
Option Explicit

Public Function ApplyResourceSubjectColors() As String
    Dim resource As Worksheet, data As Worksheet, target As Range, legend As Range
    Dim subject As String, matched As Long
    Set resource = ThisWorkbook.Worksheets("Ressourcer")
    Set data = ThisWorkbook.Worksheets("Data")
    For Each target In resource.Range("A1:A67").Cells
        subject = Trim$(CStr(target.Value2))
        If subject <> "" Then
            For Each legend In data.Range("C3:C40").Cells
                If StrComp(subject, Trim$(CStr(legend.Value2)), vbTextCompare) = 0 Then
                    target.Interior.Pattern = xlSolid
                    target.Interior.Color = legend.Offset(0, 1).Interior.Color
                    matched = matched + 1
                    Exit For
                End If
            Next legend
        End If
    Next target
    If matched = 0 Then Err.Raise 5, "ApplyResourceSubjectColors", "No resource subjects matched the Data legend."
    ApplyResourceSubjectColors = "PASS: " & matched & " resource subject fills match Data."
End Function

Public Function AuditResourceDependencies() As String
    Dim component As Object, code As Object, i As Long, text As String, result As String
    For Each component In ThisWorkbook.VBProject.VBComponents
        If component.Name <> "modResourceLayout" Then
            Set code = component.CodeModule
            For i = 1 To code.CountOfLines
                text = code.Lines(i, 1)
                If InStr(text, "Ressourcer") > 0 Or InStr(text, "'26B'!") > 0 Then
                    result = result & component.Name & ":" & i & ": " & text & vbLf
                End If
            Next i
        End If
    Next component
    AuditResourceDependencies = result
End Function

Public Function VerifyResourceLookups() As String
    Dim testSheet As Worksheet, resource As Worksheet, ranges As Variant, source As String
    Dim sem As Long, col As Long, count As Long, row As Long, hours As Double
    Dim teacher As String, oldEvents As Boolean, oldAlerts As Boolean, failure As String
    On Error GoTo Failed
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    Application.EnableEvents = False
    Set resource = ThisWorkbook.Worksheets("Ressourcer")
    Set testSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    testSheet.Name = "__ResourceVerification"
    ranges = Array("B7:X10", "B12:X15", "B17:X20", "B23:X25")
    For sem = 0 To 3
        testSheet.Range("B2:C24").ClearContents
        source = "'Ressourcer'!" & resource.Range(ranges(sem)).Address
        For row = 2 To 24
            testSheet.Cells(row, 2).FormulaLocal = _
                "=IFERROR(INDEX(FILTER('Ressourcer'!$B$1:$X$1;BYCOL(" & source & _
                ";LAMBDA(c;COUNTIF(c;""<>"" )>0)));ROWS($B$2:B" & row & "));"""")"
            testSheet.Cells(row, 3).FormulaLocal = _
                "=IFERROR(INDEX(MAP(FILTER(COLUMN('Ressourcer'!$B$1:$X$1);BYCOL(" & source & _
                ";LAMBDA(c;COUNTIF(c;""<>"" )>0)));LAMBDA(col;SUM(INDEX(" & source & _
                ";;col-COLUMN(" & source & ")+1))));ROWS($B$2:B" & row & "))/2;"""")"
        Next row
        testSheet.Calculate
        count = 0
        For col = 2 To 24
            If Application.CountA(resource.Range(ranges(sem)).Columns(col - 1)) > 0 Then
                count = count + 1
                teacher = CStr(resource.Cells(1, col).Value2)
                hours = Application.Sum(resource.Range(ranges(sem)).Columns(col - 1)) / 2
                If IsError(testSheet.Cells(count + 1, 2).Value) Or IsError(testSheet.Cells(count + 1, 3).Value) Then
                    Err.Raise 5, , "Teacher lookup calculation failed for semester " & sem + 1
                End If
                If testSheet.Cells(count + 1, 2).Value2 <> teacher Or testSheet.Cells(count + 1, 3).Value2 <> hours Then
                    Err.Raise 5, , "Teacher/hour lookup mismatch for semester " & sem + 1 & ", teacher " & teacher
                End If
            End If
        Next col
        If count > 15 Then Err.Raise 5, , "Semester " & sem + 1 & " exceeds the existing 15-teacher list capacity."
        If count < 23 Then
            If testSheet.Cells(count + 2, 2).Value2 <> "" Then Err.Raise 5, , "Teacher list contains an unexpected extra entry."
        End If
    Next sem
    Application.DisplayAlerts = False
    testSheet.Delete
    Set testSheet = Nothing
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    VerifyResourceLookups = "PASS: all four semester teacher lists and allocated hours match the resource data; no list overflow."
    Exit Function
Failed:
    failure = "ERROR " & Err.Number & ": " & Err.Description
    On Error GoTo CleanupFailed
    If Not testSheet Is Nothing Then
        Application.DisplayAlerts = False
        testSheet.Delete
    End If
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    VerifyResourceLookups = failure
    Exit Function
CleanupFailed:
    VerifyResourceLookups = failure & "; cleanup failed: " & Err.Description
End Function
