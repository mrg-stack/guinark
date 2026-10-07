Attribute VB_Name = "modInput"

Option Explicit

' ==========================================
' Global sheet name configuration
' ==========================================
Public Const SH_FERIE As String = "Ferie"
Public Const SH_DATA  As String = "Data"
Public Const SH_SEM1  As String = "1. Sem"

' ==========================================
' 1. Sem _ title row / forl�b titles
' ==========================================

' The forl�b definition table on Data sheet
Public Const SEM1_TITLE_DATA_RANGE As String = "B63:D86"

' Where titles for 1. Sem appear (row 3)
Public Const SEM1_TITLE_ROW3 As String = "E3:DE3"

' Where titles may be duplicated (row 24)
Public Const SEM1_TITLE_ROW24 As String = "E24:DE24"

' Checkbox controlling row24 copy (on Data sheet)
Public Const SEM1_TITLE_COPY_CHECKBOX As String = "D43"

' ==========================================
' 1. Sem _ conditional formatting config
' ==========================================

' CF always applies here
Public Const SEM1_CF_TOP_RANGE As String = "E7:DE10,E12:DE15"

' CF only applies here when checkbox TRUE
Public Const SEM1_CF_BOTTOM_RANGE As String = "E28:DE31,E33:DE36"

' Legend range (text + colour) on Data
Public Const SEM1_CF_LEGEND_RANGE As String = "F3:F40"

' ==========================================
' Semester populate table (Data!B46:D60)
' ==========================================
' Layout:
'   Col B: Sem (1..7, first row per sem, second row empty)
'   Col C: Hold (e.g. �KIT-GBG-F26A)
'   Col D: Populate? (TRUE/FALSE, checkboxes)

Public Const SEM_POP_FIRST_ROW As Long = 46
Public Const SEM_POP_LAST_ROW  As Long = 59  ' covers 1_7 with 2 rows each
Public Const SEM_POP_COL_SEM   As Long = 2   ' column B
Public Const SEM_POP_COL_POP   As Long = 4   ' column D

' Must be in module declaration section (before any procedures).
Public Type SemesterRanges
    sem As String                    ' Sheet name
    titleDataRange As String        ' Data range for titles (e.g., "B63:D86")
    titleRow As Long                ' Row number where titles go
    titleStartCol As String         ' Starting column (e.g., "E")
    titleCopyRow As Long            ' Row to copy titles to (usually 24)
    copyCheckbox As String          ' Checkbox controlling copy (e.g., "D43")
    cfTopRange As String            ' CF always applies here
    cfBottomRange As String         ' CF applies here when checkbox TRUE
    cfLegendRange As String         ' Legend for CF
    dataCodeCol As String           ' Column with codes (e.g., "B")
    dataTitleCol As String          ' Column with titles (e.g., "C")
    dataWeeksCol As String          ' Column with weeks (e.g., "D")
    dataValueCol As String          ' Column with CF values (e.g., "F" or "G")
    codePrefix As String            ' Code prefix to match (e.g., "1.", "2.")
End Type

' Returns TRUE if the given semester has ANY Populate?=TRUE in Data!B46:D59
Public Function ShouldPopulateSem(semNumber As Long) As Boolean
    Dim ws As Worksheet
    Dim r As Long
    Dim semVal As Variant
    Dim lastSem As Long
    
    Set ws = ThisWorkbook.Worksheets(SH_DATA)
    
    For r = SEM_POP_FIRST_ROW To SEM_POP_LAST_ROW
        
        semVal = ws.Cells(r, SEM_POP_COL_SEM).Value
        
        ' If there is a number in column B, update current semester
        If IsNumeric(semVal) Then
            lastSem = CLng(semVal)
        End If
        
        ' If this row belongs to the requested semester
        If lastSem = semNumber Then
            ' If Populate? column is TRUE, this semester should be populated
            If ws.Cells(r, SEM_POP_COL_POP).Value = True Then
                ShouldPopulateSem = True
                Exit Function
            End If
        End If
    
    Next r
End Function

Public Function GetSemesterRanges(semNumber As Long) As SemesterRanges
    Dim cfg As SemesterRanges
    
    Select Case semNumber
        Case 1
            cfg.sem = "1. Sem"
            cfg.titleDataRange = "B63:D86"
            cfg.titleRow = 3
            cfg.titleStartCol = "E"
            cfg.titleCopyRow = 24
            cfg.copyCheckbox = "D43"
            cfg.cfTopRange = SEM1_CF_TOP_RANGE
            cfg.cfBottomRange = SEM1_CF_BOTTOM_RANGE
            cfg.cfLegendRange = "F3:F40"
            cfg.dataCodeCol = "B"
            cfg.dataTitleCol = "C"
            cfg.dataWeeksCol = "D"
            cfg.dataValueCol = "F"
            cfg.codePrefix = "1."
        
        Case 2
            cfg.sem = "2. Sem"
            cfg.titleDataRange = "B63:D86"
            cfg.titleRow = 3
            cfg.titleStartCol = "E"
            cfg.titleCopyRow = 24
            cfg.copyCheckbox = "D43"
            cfg.cfTopRange = SEM1_CF_TOP_RANGE
            cfg.cfBottomRange = SEM1_CF_BOTTOM_RANGE
            cfg.cfLegendRange = "G3:G40"
            cfg.dataCodeCol = "B"
            cfg.dataTitleCol = "C"
            cfg.dataWeeksCol = "D"
            cfg.dataValueCol = "G"
            cfg.codePrefix = "2."
        
        Case Else
            MsgBox "Unsupported semester: " & semNumber, vbCritical
    End Select
    
    GetSemesterRanges = cfg
End Function


