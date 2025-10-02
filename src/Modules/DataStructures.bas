Attribute VB_Name = "DataStructures"
Option Explicit

' ========= PRIVATE CONSTANTS =========
Private Const REF_SHEET As String = "Ref"
Private Const REQ_INPUT_TABLE As String = "tbRequiredInput"
Private Const DATA_SHEET As String = "Data"
' Named ranges for validation dropdowns
Private Const CASE_LIST_RANGE As String = "CaseList"
Private Const VALVE_LIST_RANGE As String = "ValveList"
Private Const SUPPORT_LIST_RANGE As String = "SupportList"

' ========= DATA STRUCTURES MODULE =========
' This module defines common data structures used across the application

' Structure to hold all valve input parameters
Public Type ValveInputs
    ' Basic fluid properties
    rho As Double                    ' Fluid density
    gamma As Double                  ' Ratio of Specific Heat Capacities
    c0 As Double                     ' Speed of sound
    Mw As Double                     ' Molecular Weight
    r As Double                      ' Universal Gas Constant
    Te As Double                     ' Upstream Temperature
    Pv As Double                     ' Vapour Pressure
    Kbulk As Double                  ' Fluid Bulk Modulus
    
    ' Pipe properties
    Dext_mm As Double               ' External Main Line Diameter (mm)
    Dint_mm As Double               ' Internal Main Line Diameter (mm)
    T_mm As Double                  ' Main line Wall Thickness (mm)
    Tsch40 As Double                ' Main line Wall Thickness for Schedule 40 Piping (mm)
    Em As Double                    ' Young's Modulus of the main line material
    Lup As Double                   ' Upstream Pipe Length
    
    ' Operating conditions
    P1 As Double                    ' Upstream Static Pressure
    dP As Double                    ' Static Pressure drop
    v As Double                     ' Steady State Fluid Velocity
    W As Double                     ' Mass Flow Rate
    Pshut As Double                 ' Pump head at zero flow
    
    ' Valve properties
    Tclose As Double                ' Valve Closing Time
    valveType As String             ' Valve Type
    supportType As String           ' Pipe Support Type
    
    ' Identification
    tag As String                   ' Valve tag
    caseType As String              ' Case type for calculation
End Type

' Structure for calculation results
Public Type CalculationResult
    Ppeak As Double
    Fmax As Double
    Flim As Double
    LOF As Double
    FlagText As String
End Type

' ========= PARAMETER MANAGEMENT (Types must be before any procedures) =========
' Holds a fast index for parameter names to row numbers
Private Type ParameterIndex
    Dictionary As Object
    lastRow As Long
    IsInitialized As Boolean
End Type

' Module-level cache of the parameter index
Private paramIdx As ParameterIndex

' Function to create a default ValveInputs structure
Public Function CreateDefaultValveInputs() As ValveInputs
    Dim inputs As ValveInputs
    ' Initialize with default/zero values
    ' All numeric fields are automatically initialized to 0
    ' String fields are automatically initialized to empty string
    CreateDefaultValveInputs = inputs
End Function

' Function to validate ValveInputs structure
Public Function ValidateValveInputs(inputs As ValveInputs) As String
    Dim errors As String
    
    ' Check for required positive values
    If inputs.rho <= 0 Then errors = errors & "Fluid density must be positive. "
    If inputs.Dext_mm <= 0 Then errors = errors & "External diameter must be positive. "
    If inputs.Dint_mm <= 0 Then errors = errors & "Internal diameter must be positive. "
    If inputs.T_mm <= 0 Then errors = errors & "Wall thickness must be positive. "
    
    ' Check for logical relationships
    If inputs.Dint_mm >= inputs.Dext_mm Then
        errors = errors & "Internal diameter must be less than external diameter. "
    End If
    
    If inputs.T_mm >= inputs.Dext_mm / 2 Then
        errors = errors & "Wall thickness too large for given external diameter. "
    End If
    
    ' Check for required string values
    If Trim$(inputs.tag) = "" Then errors = errors & "Valve tag cannot be empty. "
    If Trim$(inputs.caseType) = "" Then errors = errors & "Case type cannot be empty. "
    
    ValidateValveInputs = Trim$(errors)
End Function

' Function to copy ValveInputs structure
Public Function CopyValveInputs(source As ValveInputs) As ValveInputs
    Dim Target As ValveInputs
    
    ' Copy all fields
    Target.rho = source.rho
    Target.gamma = source.gamma
    Target.c0 = source.c0
    Target.Mw = source.Mw
    Target.r = source.r
    Target.Te = source.Te
    Target.Pv = source.Pv
    Target.Kbulk = source.Kbulk
    Target.Dext_mm = source.Dext_mm
    Target.Dint_mm = source.Dint_mm
    Target.T_mm = source.T_mm
    Target.Tsch40 = source.Tsch40
    Target.Em = source.Em
    Target.Lup = source.Lup
    Target.P1 = source.P1
    Target.dP = source.dP
    Target.v = source.v
    Target.W = source.W
    Target.Pshut = source.Pshut
    Target.Tclose = source.Tclose
    Target.valveType = source.valveType
    Target.supportType = source.supportType
    Target.tag = source.tag
    Target.caseType = source.caseType
    
    CopyValveInputs = Target
End Function

' ========= PARAMETER MANAGEMENT =========

' Initialize or refresh the parameter index for a given worksheet
Public Sub InitializeParameterIndex(ws As Worksheet)
    Set paramIdx.Dictionary = CreateObject("Scripting.Dictionary")
    paramIdx.lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    paramIdx.IsInitialized = False
    
    Dim r As Long
    For r = 3 To paramIdx.lastRow
        Dim paramName As String
        paramName = Trim$(ws.Cells(r, 1).value)
        If paramName <> "" Then
            paramIdx.Dictionary(paramName) = r
        End If
    Next r
    
    paramIdx.IsInitialized = True
End Sub

' Get the row number for a specific parameter name
Public Function GetParameterRow(parameterName As String) As Long
    If Not paramIdx.IsInitialized Then
        Err.Raise 9999, "DataStructures", "Parameter index not initialized. Call InitializeParameterIndex first."
    End If
    
    If paramIdx.Dictionary.Exists(parameterName) Then
        GetParameterRow = paramIdx.Dictionary(parameterName)
    Else
        GetParameterRow = 0 ' Parameter not found
    End If
End Function

' Get parameter value from a specific column
Public Function GetParameterValue(ws As Worksheet, parameterName As String, columnIndex As Long) As Variant
    Dim rowIndex As Long
    rowIndex = GetParameterRow(parameterName)
    
    If rowIndex = 0 Then
        GetParameterValue = Empty
        Exit Function
    End If
    
    GetParameterValue = ws.Cells(rowIndex, columnIndex).value
End Function

' Get parameter value as Double with null protection
Public Function GetParameterDouble(ws As Worksheet, parameterName As String, columnIndex As Long) As Double
    Dim value As Variant
    value = GetParameterValue(ws, parameterName, columnIndex)
    GetParameterDouble = NzD(value)
End Function

' Get parameter value as String with null protection
Public Function GetParameterString(ws As Worksheet, parameterName As String, columnIndex As Long) As String
    Dim value As Variant
    value = GetParameterValue(ws, parameterName, columnIndex)
    GetParameterString = NzS(value)
End Function

' Check if parameter exists in the index
Public Function ParameterExists(parameterName As String) As Boolean
    If Not paramIdx.IsInitialized Then
        ParameterExists = False
        Exit Function
    End If
    
    ParameterExists = paramIdx.Dictionary.Exists(parameterName)
End Function

' Get all parameter names
' Returns an array of parameter names; returns an uninitialized dynamic array if empty.
' Usage:
'   Dim names() As String
'   names = GetAllParameterNames()
'   If (Not Not names) = 0 Then ' empty array
'       ' handle no parameters
'   Else
'       ' use LBound/UBound on names
'   End If
Public Function GetAllParameterNames() As String()
    If Not paramIdx.IsInitialized Then
        Err.Raise 9999, "DataStructures", "Parameter index not initialized. Call InitializeParameterIndex first."
    End If

    If paramIdx.Dictionary.Count = 0 Then
        Dim emptyResult() As String ' uninitialized dynamic array (valid empty return)
        GetAllParameterNames = emptyResult
        Exit Function
    End If

    Dim keys As Variant
    keys = paramIdx.Dictionary.keys

    Dim result() As String
    ReDim result(0 To paramIdx.Dictionary.Count - 1)

    Dim i As Long
    For i = 0 To paramIdx.Dictionary.Count - 1
        result(i) = CStr(keys(i))
    Next i

    GetAllParameterNames = result
End Function

' Clean up the parameter index
Public Sub ClearParameterIndex()
    Set paramIdx.Dictionary = Nothing
    paramIdx.lastRow = 0
    paramIdx.IsInitialized = False
End Sub

' ========= REQUIRED INPUT VALIDATION =========

' Check if a parameter is required for a specific case type
Public Function IsParameterRequired(parameterName As String, caseType As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(REF_SHEET)
    
    Dim lo As ListObject
    Set lo = ws.ListObjects(REQ_INPUT_TABLE)
    
    If lo Is Nothing Then
        IsParameterRequired = False
        Exit Function
    End If
    
    ' Find the parameter row
    Dim lr As ListRow
    For Each lr In lo.ListRows
        If Trim$(lr.Range.Cells(1, 1).value) = parameterName Then
            ' Find the case type column and get the value
            Dim colIndex As Long
            Select Case LCase$(Trim$(caseType))
                Case "valve closure"
                    colIndex = 4 ' Valve Closure column
                Case "valve opening (liquid/multiphase)"
                    colIndex = 5 ' Valve Opening (Liquid/Multiphase) column
                Case "valve opening (dry gas)"
                    colIndex = 6 ' Valve Opening (Dry Gas) column
                Case Else
                    IsParameterRequired = False
                    Exit Function
            End Select
            
            IsParameterRequired = (lr.Range.Cells(1, colIndex).value = 1)
            Exit Function
        End If
    Next lr
    
    IsParameterRequired = False
    Exit Function
    
ErrorHandler:
    IsParameterRequired = False
End Function

' Get all required parameters for a specific case type
Public Function GetRequiredParameters(caseType As String) As Collection
    Set GetRequiredParameters = New Collection
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(REF_SHEET)
    
    Dim lo As ListObject
    Set lo = ws.ListObjects(REQ_INPUT_TABLE)
    
    If lo Is Nothing Then Exit Function
    
    ' Determine case type column
    Dim colIndex As Long
    Select Case LCase$(Trim$(caseType))
        Case "valve closure"
            colIndex = 4 ' Valve Closure column
        Case "valve opening (liquid/multiphase)"
            colIndex = 5 ' Valve Opening (Liquid/Multiphase) column
        Case "valve opening (dry gas)"
            colIndex = 6 ' Valve Opening (Dry Gas) column
        Case Else
            Exit Function
    End Select
    
    ' Collect required parameters
    Dim lr As ListRow
    For Each lr In lo.ListRows
        If lr.Range.Cells(1, colIndex).value = 1 Then
            GetRequiredParameters.Add Trim$(lr.Range.Cells(1, 1).value)
        End If
    Next lr
    
    Exit Function
    
ErrorHandler:
    ' Return empty collection on error
End Function

' Validate all required inputs for a specific case type and valve column
Public Function ValidateRequiredInputs(ws As Worksheet, caseType As String, columnIndex As Long) As String
    Dim errors As String
    Dim requiredParams As Collection
    Set requiredParams = GetRequiredParameters(caseType)
    
    If requiredParams.Count = 0 Then
        ValidateRequiredInputs = "No validation rules found for case type: " & caseType
        Exit Function
    End If
    
    Dim param As Variant
    For Each param In requiredParams
        Dim value As Variant
        value = GetParameterValue(ws, CStr(param), columnIndex)
        
        If IsEmpty(value) Or Trim$(CStr(value)) = "" Or value = 0 Then
            errors = errors & "Required parameter missing: " & param & vbCrLf
        End If
    Next param
    
    ValidateRequiredInputs = Trim$(errors)
End Function

' ========= UNIT CONVERSION FUNCTIONS =========

' Convert pressure from barg to Pa (absolute pressure in Pascals)
Public Function BargToPa(ByVal pressureBarg As Double) As Double
    ' Convert barg to absolute bar, then to Pa
    ' 1 bar = 100000 Pa, Atmospheric pressure = 1.01325 bar
    BargToPa = (pressureBarg + 1.01325) * 100000#
End Function

' Convert pressure from Pa to barg
Public Function PaToBarg(ByVal pressurePa As Double) As Double
    ' Convert Pa to absolute bar, then subtract atmospheric pressure
    PaToBarg = (pressurePa / 100000#) - 1.01325
End Function

' Get parameter value as Double with automatic barg to Pa conversion for pressure parameters
Public Function GetPressureParameterPa(ws As Worksheet, parameterName As String, columnIndex As Long) As Double
    Dim value As Variant
    value = GetParameterValue(ws, parameterName, columnIndex)
    Dim bargValue As Double: bargValue = NzD(value)
    GetPressureParameterPa = BargToPa(bargValue)
End Function

' Update pressure parameter units from Pa to barg in column C
Public Sub UpdatePressureUnitsToBarG(ws As Worksheet)
    If Not paramIdx.IsInitialized Then
        InitializeParameterIndex ws
    End If
    
    ' List of pressure parameters to update
    Dim pressureParams As Variant
    pressureParams = Array("Upstream Static Pressure", "Pump head at zero flow", "Vapour Pressure", "Static Pressure drop")
    
    Dim i As Long
    For i = 0 To UBound(pressureParams)
        Dim paramName As String: paramName = pressureParams(i)
        Dim rowIndex As Long: rowIndex = GetParameterRow(paramName)
        
        If rowIndex > 0 Then
            ' Update the unit in column C from "Pa" to "barg"
            Dim currentUnit As String: currentUnit = Trim$(ws.Cells(rowIndex, 3).value)
            If LCase$(currentUnit) = "pa" Or LCase$(currentUnit) = "pascal" Then
                ws.Cells(rowIndex, 3).value = "barg"
                ' Optional: Add a comment to indicate the change
                ws.Cells(rowIndex, 3).AddComment "Unit changed from " & currentUnit & " to barg - values now expected in bar gauge"
            End If
        End If
    Next i
    
    MsgBox "Updated pressure parameter units to barg in column C for:" & vbCrLf & _
           "• Upstream Static Pressure" & vbCrLf & _
           "• Pump head at zero flow" & vbCrLf & _
           "• Vapour Pressure" & vbCrLf & _
           "• Static Pressure drop", vbInformation, "Unit Conversion Complete"
End Sub

' ========= TABLE VALIDATION FUNCTIONS =========

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
    ' If validationSheet is empty, treat validationRange as a named range
    If validationSheet = "" Then
        validationFormula = "=" & validationRange
    Else
        validationFormula = "=" & validationSheet & "!" & validationRange
    End If
    
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
                                  CASE_LIST_RANGE, ""
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
    ' If validationSheet is empty, treat validationRange as a named range
    If validationSheet = "" Then
        validationFormula = "=" & validationRange
    Else
        validationFormula = "=" & validationSheet & "!" & validationRange
    End If
    
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

' ========= HELPER FUNCTIONS =========

' NzD: Null/Empty/Error-safe CDbl conversion.
' Returns 0 for Null/Empty/Error/Non-numeric values.
' Usage:
'   Dim x As Double: x = NzD(Cells(2, 1).Value)
Private Function NzD(v) As Double
    On Error GoTo ErrorHandler
    If IsError(v) Then NzD = 0#: Exit Function
    If IsNull(v) Then NzD = 0#: Exit Function
    If Len(Trim$(v & "")) = 0 Then NzD = 0#: Exit Function
    If IsNumeric(v) Then
        NzD = CDbl(v)
    Else
        NzD = 0#
    End If
    Exit Function
ErrorHandler:
    NzD = 0#
End Function

' NzS: Null/Error-safe CStr conversion.
' Usage:
'   Dim s As String: s = NzS(Cells(3, 1).Value)
Private Function NzS(v) As String
    If IsError(v) Then Exit Function
    NzS = CStr(v)
End Function
