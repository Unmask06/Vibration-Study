Attribute VB_Name = "TableValidationManager"
Option Explicit

' ========= TABLE VALIDATION MANAGER =========
' This module handles data validation for Excel tables
' Specifically designed to work with valve list tables

Private Const DATA_SHEET As String = "Data"
Private Const CASE_TYPE_VALIDATION_RANGE As String = "B3:B5"

' Add data validation to a specific column in a table
Public Sub AddDataValidationToTableColumn(tableName As String, sheetName As String, _
                                         columnName As String, validationRange As String, _
                                         validationSheet As String)
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    Dim lo As ListObject
    Set lo = ws.ListObjects(tableName)
    
    If lo Is Nothing Then
        MsgBox "Table '" & tableName & "' not found on sheet '" & sheetName & "'.", vbExclamation
        Exit Sub
    End If
    
    ' Find the column index
    Dim colIndex As Long
    colIndex = GetTableColumnIndex(lo, columnName)
    
    If colIndex = 0 Then
        MsgBox "Column '" & columnName & "' not found in table '" & tableName & "'.", vbExclamation
        Exit Sub
    End If
    
    ' Get the data column range (excluding header)
    Dim dataRange As Range
    Set dataRange = lo.ListColumns(colIndex).DataBodyRange
    
    ' If there's no data yet, create a reference to the first data cell
    If dataRange Is Nothing Then
        Set dataRange = lo.ListColumns(colIndex).Range.Cells(2, 1)
    End If
    
    ' Create validation formula
    Dim validationFormula As String
    validationFormula = "=" & validationSheet & "!" & validationRange
    
    ' Apply data validation
    With dataRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:=validationFormula
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .InputTitle = "Select " & columnName
        .InputMessage = "Please select a value from the dropdown list."
        .ShowError = True
        .ErrorTitle = "Invalid Entry"
        .ErrorMessage = "Please select a value from the dropdown list."
    End With
End Sub

' Setup data validation for the CaseType column in tbValveList
Public Sub SetupCaseTypeValidation()
    AddDataValidationToTableColumn "tbValveList", "ValveList", "CaseType", _
                                  CASE_TYPE_VALIDATION_RANGE, DATA_SHEET
End Sub

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

' Event handler for when new rows are added to a table
Public Sub OnTableRowAdded(tableName As String, sheetName As String)
    Select Case tableName
        Case "tbValveList"
            Call SetupCaseTypeValidation
        ' Add other table handlers here as needed
    End Select
End Sub

' Refresh all table validations
Public Sub RefreshAllTableValidations()
    Call SetupCaseTypeValidation
    ' Add other validation refresh calls here
End Sub

' Apply validation to a specific range
Public Sub ApplyValidationToRange(targetRange As Range, validationRange As String, _
                                 validationSheet As String, Optional title As String = "", _
                                 Optional message As String = "")
    
    Dim validationFormula As String
    validationFormula = "=" & validationSheet & "!" & validationRange
    
    With targetRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:=validationFormula
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        If title <> "" Then .InputTitle = title
        If message <> "" Then .InputMessage = message
        .ShowError = True
        .ErrorTitle = "Invalid Entry"
        .ErrorMessage = "Please select a value from the dropdown list."
    End With
End Sub
