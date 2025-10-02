Option Explicit

' ========= INPUTS SHEET EVENT HANDLER MODULE =========
' This module handles worksheet events specific to the Inputs sheet
' Manages automatic parameter highlighting when Case Type changes

' ========= CONFIGURATION CONSTANTS =========
Private Const SHEET_INPUTS As String = "Inputs"
Private Const TABLE_INPUT As String = "tbInput"
Private Const CASE_TYPE_PARAMETER As String = "Case Type"
Private Const INPUTS_FIRST_COL As Long = 5   ' "E" — first valve column in Inputs


' ========= EVENT HANDLERS =========
' Event handler for Worksheet_SelectionChange event in Inputs sheet
' This should be called from the Inputs worksheet's Worksheet_SelectionChange event
Public Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_INPUTS)
    
    Dim lo As ListObject
    Set lo = ws.ListObjects(TABLE_INPUT)
    
    If lo Is Nothing Then Exit Sub
    
    ' Find the Case Type row number
    Dim caseTypeRow As Long
    caseTypeRow = FindCaseTypeRow(lo)
    
    If caseTypeRow = 0 Then Exit Sub
    
    ' Define WatchRange: Case Type row from INPUTS_FIRST_COL to end of table
    Dim lastCol As Long
    lastCol = lo.Range.Column + lo.ListColumns.Count - 1
    
    Dim WatchRange As Range
    Set WatchRange = ws.Range(ws.Cells(caseTypeRow, INPUTS_FIRST_COL), ws.Cells(caseTypeRow, lastCol))
    
    ' Check if Target intersects with WatchRange
    If Not Application.Intersect(Target, WatchRange) Is Nothing Then
        ' Process each cell in the intersection
        Dim cell As Range
        For Each cell In Application.Intersect(Target, WatchRange)
            Call UIManager.HighlightParametersByCase(ws, cell.Column, cell.value)
        Next cell
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in Inputs_Worksheet_SelectionChange: " & Err.Number & " - " & Err.Description
End Sub

' ========= HELPER FUNCTIONS =========

' Get the column index for a given column name in a table
Private Function GetTableColumnIndex(lo As ListObject, columnName As String) As Long
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        If lo.ListColumns(i).Name = columnName Then
            GetTableColumnIndex = i
            Exit Function
        End If
    Next i
    GetTableColumnIndex = 0
End Function

' Find the actual row number of the Case Type parameter in the worksheet
Private Function FindCaseTypeRow(lo As ListObject) As Long
    On Error GoTo ErrorHandler
    
    FindCaseTypeRow = 0
    
    Dim parameterColIndex As Long
    parameterColIndex = GetTableColumnIndex(lo, "Parameter")
    
    If parameterColIndex = 0 Then Exit Function
    
    Dim i As Long
    For i = 1 To lo.ListRows.Count
        Dim parameterValue As String
        parameterValue = lo.ListColumns(parameterColIndex).DataBodyRange.Cells(i, 1).value
        
        If Trim(parameterValue) = CASE_TYPE_PARAMETER Then
            FindCaseTypeRow = lo.Range.row + i
            Exit Function
        End If
    Next i
    
    Exit Function
    
ErrorHandler:
    FindCaseTypeRow = 0
End Function
