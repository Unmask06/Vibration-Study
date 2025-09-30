Attribute VB_Name = "TableEventHandler"
Option Explicit

' ========= TABLE EVENT HANDLER MODULE =========
' This module handles events for Excel tables, particularly the tbValveList table

' Monitor for table changes and apply validation when needed
Public Sub MonitorTableChanges()
    ' This would typically be called from worksheet events
    ' For now, we'll provide a manual trigger
    Call RefreshTableValidations
End Sub

' Refresh validations for all monitored tables
Public Sub RefreshTableValidations()
    On Error GoTo ErrorHandler
    
    ' Refresh validation for tbValveList
    Call SetupValveListValidation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error refreshing table validations: " & Err.Description, vbExclamation
End Sub

' Setup validation specifically for the valve list table
Private Sub SetupValveListValidation()
    Dim ws As Worksheet
    Dim lo As ListObject
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("ValveList")
    Set lo = ws.ListObjects("tbValveList")
    On Error GoTo 0
    
    If lo Is Nothing Then
        MsgBox "Table 'tbValveList' not found on sheet 'ValveList'.", vbExclamation
        Exit Sub
    End If
    
    ' Setup CaseType validation
    Call SetupCaseTypeValidationForTable(lo)
End Sub

' Setup CaseType validation for the entire table
Private Sub SetupCaseTypeValidationForTable(lo As ListObject)
    On Error GoTo ErrorHandler
    
    ' Find CaseType column
    Dim caseTypeCol As ListColumn
    Set caseTypeCol = Nothing
    
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        If lo.ListColumns(i).Name = "CaseType" Then
            Set caseTypeCol = lo.ListColumns(i)
            Exit For
        End If
    Next i
    
    If caseTypeCol Is Nothing Then
        ' Column doesn't exist, try to add it
        MsgBox "CaseType column not found in tbValveList table.", vbInformation
        Exit Sub
    End If
    
    ' Get the validation range
    Dim validationRange As Range
    Set validationRange = caseTypeCol.DataBodyRange
    
    ' If no data rows exist yet, get the first potential data cell
    If validationRange Is Nothing Then
        Set validationRange = caseTypeCol.Range.Cells(2, 1)
    End If
    
    ' Apply validation
    Call ApplyValidationToRange(validationRange, "B3:B5", "Data", _
                               "Case Type Selection", _
                               "Select the type of calculation case from the available options.")
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error setting up CaseType validation: " & Err.Description, vbExclamation
End Sub

' Event procedure that can be called when a new row is added
Public Sub OnTableRowAdded(sheetName As String, tableName As String)
    If sheetName = "ValveList" And tableName = "tbValveList" Then
        ' Small delay to ensure the row is fully created
        Application.Wait (Now + TimeValue("0:00:01"))
        Call SetupValveListValidation
    End If
End Sub

' Manual trigger to setup validation (can be called from a button or menu)
Public Sub SetupTableValidations()
    Call RefreshTableValidations
    MsgBox "Table validations have been applied successfully.", vbInformation
End Sub

' Helper function to apply validation to a range
Private Sub ApplyValidationToRange(targetRange As Range, validationRange As String, _
                                  validationSheet As String, Optional title As String = "", _
                                  Optional message As String = "")
    
    On Error GoTo ErrorHandler
    
    Dim validationFormula As String
    validationFormula = "=" & validationSheet & "!" & validationRange
    
    ' Clear existing validation first
    targetRange.Validation.Delete
    
    ' Apply new validation
    With targetRange.Validation
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
    
    Exit Sub
    
ErrorHandler:
    ' If validation fails, continue without error to avoid breaking other functionality
    Debug.Print "Warning: Could not apply validation to range " & targetRange.Address & ": " & Err.Description
End Sub
