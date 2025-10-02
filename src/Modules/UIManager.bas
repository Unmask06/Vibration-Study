Attribute VB_Name = "UIManager"
Option Explicit

' ========= UI MANAGER MODULE =========
' This module handles UI interactions and worksheet management
' Manages valve list generation, input highlighting, and user interface

' ========= CONFIGURATION CONSTANTS =========
Private Const INPUTS_FIRST_COL As Long = 5   ' "E" — first valve column in Inputs
Private Const SHEET_VALVELIST As String = "ValveList"
Private Const SHEET_INPUTS As String = "Inputs"
Private Const SHEET_RESULTS As String = "Results"
Private Const SHEET_DATA As String = "Data"
Private Const SHEET_REF As String = "Ref"
Private Const TABLE_VALVES As String = "tbValveList"

' Colors for parameter highlighting
Private Const COLOR_REQUIRED As Long = 65535      ' Light yellow (RGB(255, 255, 0))
Private Const COLOR_NOT_REQUIRED As Long = 12632256 ' Light grey (RGB(192, 192, 192))

' Row labels in Inputs (must match the left-most "Parameter" text)
Private Const ROW_PIPE_SUPPORT As String = "Pipe Support Type"
Private Const ROW_VALVE_TYPE As String = "Valve Type"

' ========= MAIN PUBLIC FUNCTIONS =========

' Initialize required modules for valve list operations
Public Sub InitializeModules()
    Dim wsIn As Worksheet: Set wsIn = Sheets(SHEET_INPUTS)
    Call DataStructures.InitializeParameterIndex(wsIn)
    Call DataStructures.RefreshAllTableValidations
End Sub

' Generate Inputs sheet columns directly from the Excel Table tbValveList (Tag, CaseType)
Public Sub Generate_Inputs_From_tbValveList()
    ' Initialize modules
    Call InitializeModules
    
    Dim lo As ListObject
    On Error Resume Next
    Set lo = Sheets(SHEET_VALVELIST).ListObjects(TABLE_VALVES)
    On Error GoTo 0
    If lo Is Nothing Then
        MsgBox "ListObject '" & TABLE_VALVES & "' not found on sheet '" & SHEET_VALVELIST & "'.", vbExclamation
        Exit Sub
    End If
    
    Dim wsIn As Worksheet: Set wsIn = Sheets(SHEET_INPUTS)
    Dim lastCol As Long: lastCol = wsIn.Cells(2, wsIn.Columns.Count).End(xlToLeft).Column
    If lastCol >= INPUTS_FIRST_COL Then
        wsIn.Range(wsIn.Cells(2, INPUTS_FIRST_COL), wsIn.Cells(wsIn.Rows.Count, lastCol)).Clear
    End If
    
    Dim lr As ListRow, c As Long: c = INPUTS_FIRST_COL
    Dim added As Long: added = 0
    For Each lr In lo.ListRows
        Dim tag As String: tag = Trim$(NzS(lr.Range.Cells(1, 1).value))  ' first column = Tag
        Dim caseType As String: caseType = Trim$(NzS(lr.Range.Cells(1, 2).value))  ' second column = CaseType
        If tag <> "" Then
            ' Add column header
            wsIn.Cells(2, c).value = tag
            With wsIn.Cells(2, c)
                .Interior.Color = RGB(31, 78, 120)
                .Font.Color = vbWhite
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
            End With
            
            ' Set Case Type in the first row for this valve
            Call SetParameterFromTable(wsIn, "Case Type", c, caseType)
            
            ' Set Pipe Support Type and Valve Type from table
            Call SetParameterFromTable(wsIn, ROW_PIPE_SUPPORT, c, NzS(lr.Range.Cells(1, 4).value))  ' 4th column = SupportType
            Call SetParameterFromTable(wsIn, ROW_VALVE_TYPE, c, NzS(lr.Range.Cells(1, 3).value))    ' 3rd column = ValveType
            
            ' Apply parameter highlighting based on case type
            Call HighlightParametersByCase(wsIn, c, caseType)
            
            c = c + 1
            added = added + 1
        End If
    Next lr
    
    If added > 0 Then
        MsgBox "Generated " & added & " valve column(s) in Inputs sheet from tbValveList.", vbInformation
    Else
        MsgBox "No valid valve tags found in tbValveList.", vbExclamation
    End If
End Sub

' Run calculations for all valve columns
Public Sub RunCalculations()
    Call InitializeModules
    
    Dim wsIn As Worksheet, wsRes As Worksheet
    Set wsIn = Sheets(SHEET_INPUTS)
    Set wsRes = Sheets(SHEET_RESULTS)
    
    ' Find valve columns (starting from column E)
    Dim lastCol As Long: lastCol = wsIn.Cells(2, wsIn.Columns.Count).End(xlToLeft).Column
    If lastCol < INPUTS_FIRST_COL Then
        MsgBox "No valve columns found in Inputs sheet.", vbExclamation
        Exit Sub
    End If
    
    ' Clear previous results
    wsRes.Range("B2:Z1000").Clear
    
    ' Process each valve column
    Dim c As Long, resultCol As Long: resultCol = 2
    For c = INPUTS_FIRST_COL To lastCol
        Dim tag As String: tag = Trim$(wsIn.Cells(2, c).value)
        If tag <> "" Then
            ' Read inputs for this valve
            Dim inputs As ValveInputs
            inputs = ReadValveInputsFromColumn(wsIn, c)
            inputs.tag = tag
            
            ' Get case type from ValveList table
            inputs.caseType = GetCaseTypeFromTable(tag)
            
            ' Validate required inputs before calculation
            Dim validationErrors As String
            validationErrors = DataStructures.ValidateRequiredInputs(wsIn, inputs.caseType, c)
            
            If validationErrors <> "" Then
                MsgBox "Validation errors for valve " & tag & ":" & vbCrLf & validationErrors, vbExclamation, "Input Validation Failed"
                GoTo NextValve ' Skip this valve and continue with the next one
            End If
            
            ' Perform calculation
            Dim result As CalculationResult
            result = CalculationEngine.CalculateByCase(inputs.caseType, inputs)
            
            ' Write results
            Call WriteCalculationResults(wsRes, resultCol, tag, result)
            resultCol = resultCol + 1
            
NextValve:
        End If
    Next c
    
    MsgBox "Calculations completed for " & (resultCol - 2) & " valve(s).", vbInformation
End Sub

' ========= PARAMETER HIGHLIGHTING FUNCTIONS =========

' Highlight parameters based on case type requirements
Public Sub HighlightParametersByCase(ws As Worksheet, columnIndex As Long, caseType As String)
    ' Get all parameter names from the worksheet
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    Dim r As Long
    For r = 3 To lastRow ' Start from row 3 (assuming row 1 is header, row 2 is valve tags)
        Dim paramName As String
        paramName = Trim$(ws.Cells(r, 1).value)
        
        If paramName <> "" And paramName <> "Case Type" Then
            Dim isRequired As Boolean
            isRequired = DataStructures.IsParameterRequired(paramName, caseType)
            
            Dim targetCell As Range
            Set targetCell = ws.Cells(r, columnIndex)
            
            If isRequired Then
                ' Highlight required parameters with light yellow
                targetCell.Interior.Color = COLOR_REQUIRED
                targetCell.Locked = False
                targetCell.Font.Color = vbBlack
            Else
                ' Grey out and lock non-required parameters
                targetCell.Interior.Color = COLOR_NOT_REQUIRED
                targetCell.Locked = True
                targetCell.Font.Color = RGB(128, 128, 128)
                targetCell.value = "" ' Clear any existing values for non-required parameters
            End If
        End If
    Next r
End Sub

' Refresh highlighting for all valve columns when case types change
Public Sub RefreshParameterHighlighting()
    Call InitializeModules
    
    Dim wsIn As Worksheet
    Set wsIn = Sheets(SHEET_INPUTS)
    
    ' Find valve columns (starting from column E)
    Dim lastCol As Long: lastCol = wsIn.Cells(2, wsIn.Columns.Count).End(xlToLeft).Column
    If lastCol < INPUTS_FIRST_COL Then
        MsgBox "No valve columns found in Inputs sheet.", vbExclamation
        Exit Sub
    End If
    
    ' Process each valve column
    Dim c As Long
    For c = INPUTS_FIRST_COL To lastCol
        Dim tag As String: tag = Trim$(wsIn.Cells(2, c).value)
        If tag <> "" Then
            ' Get case type for this valve
            Dim caseType As String
            caseType = GetCaseTypeFromTable(tag)
            
            ' Apply highlighting
            Call HighlightParametersByCase(wsIn, c, caseType)
        End If
    Next c
    
    MsgBox "Parameter highlighting refreshed for all valve columns.", vbInformation
End Sub

' ========= HELPER FUNCTIONS =========

' Set parameter value from table data
Private Sub SetParameterFromTable(ws As Worksheet, parameterName As String, columnIndex As Long, value As String)
    Dim rowIndex As Long
    rowIndex = DataStructures.GetParameterRow(parameterName)
    
    If rowIndex > 0 Then
        ws.Cells(rowIndex, columnIndex).value = value
    End If
End Sub

' Setup dropdowns for valve column with case type consideration
Private Sub SetupValveColumnDropdowns(ws As Worksheet, columnIndex As Long, valveType As String, supportType As String, Optional caseType As String = "")
    ' Find parameter rows
    Dim lastParamRow As Long
    lastParamRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    Dim vtRow As Long, stRow As Long, ctRow As Long
    Dim r As Long
    For r = 3 To lastParamRow
        Dim paramName As String: paramName = Trim$(ws.Cells(r, "A").value)
        If paramName = ROW_VALVE_TYPE Then vtRow = r
        If paramName = ROW_PIPE_SUPPORT Then stRow = r
        If paramName = "Case Type" Then ctRow = r
    Next r
    
    ' Set case type if provided
    If ctRow > 0 And caseType <> "" Then
        ws.Cells(ctRow, columnIndex).value = caseType
        ' Add validation using named range
        Call DataStructures.ApplyValidationToRange(ws.Cells(ctRow, columnIndex), "CaseList", "", "Case Type", "Select case type")
    End If
    
    ' Set valve type
    If vtRow > 0 Then
        ws.Cells(vtRow, columnIndex).value = valveType
        ' Add validation using named range
        Call DataStructures.ApplyValidationToRange(ws.Cells(vtRow, columnIndex), "ValveList", "", "Valve Type", "Select valve type")
    End If
    
    ' Set support type
    If stRow > 0 Then
        ws.Cells(stRow, columnIndex).value = supportType
        ' Add validation using named range
        Call DataStructures.ApplyValidationToRange(ws.Cells(stRow, columnIndex), "SupportList", "", "Support Type", "Select support type")
    End If
End Sub

' Read valve inputs from a specific column
Private Function ReadValveInputsFromColumn(ws As Worksheet, columnIndex As Long) As ValveInputs
    Dim inputs As ValveInputs
    
    ' Read all parameters using DataStructures module
    inputs.rho = DataStructures.GetParameterDouble(ws, "Fluid density", columnIndex)
    inputs.gamma = DataStructures.GetParameterDouble(ws, "Ratio of Specific Heat Capacities (Cp/Cv)", columnIndex)
    inputs.c0 = DataStructures.GetParameterDouble(ws, "Speed of sound", columnIndex)
    inputs.Mw = DataStructures.GetParameterDouble(ws, "Molecular Weight", columnIndex)
    inputs.r = DataStructures.GetParameterDouble(ws, "Universal Gas Constant", columnIndex)
    inputs.Te = DataStructures.GetParameterDouble(ws, "Upstream Temperature", columnIndex)
    inputs.Pv = DataStructures.GetParameterDouble(ws, "Vapour Pressure", columnIndex)
    inputs.Kbulk = DataStructures.GetParameterDouble(ws, "Fluid Bulk Modulus", columnIndex)
    
    inputs.Dext_mm = DataStructures.GetParameterDouble(ws, "External Main Line Diameter", columnIndex)
    inputs.Dint_mm = DataStructures.GetParameterDouble(ws, "Internal Main Line Diameter", columnIndex)
    inputs.T_mm = DataStructures.GetParameterDouble(ws, "Main line Wall Thickness", columnIndex)
    inputs.Tsch40 = DataStructures.GetParameterDouble(ws, "Main line Wall Thickness for SCH 40", columnIndex)
    inputs.Em = DataStructures.GetParameterDouble(ws, "Young's Modulus of main line material", columnIndex)
    inputs.Lup = DataStructures.GetParameterDouble(ws, "Upstream Pipe Length", columnIndex)
    
    inputs.P1 = DataStructures.GetParameterDouble(ws, "Upstream Static Pressure", columnIndex)
    inputs.dP = DataStructures.GetParameterDouble(ws, "Static Pressure drop", columnIndex)
    inputs.v = DataStructures.GetParameterDouble(ws, "Steady State Fluid Velocity", columnIndex)
    inputs.W = DataStructures.GetParameterDouble(ws, "Mass Flow Rate", columnIndex)
    inputs.Pshut = DataStructures.GetParameterDouble(ws, "Pump head at zero flow", columnIndex)
    
    inputs.Tclose = DataStructures.GetParameterDouble(ws, "Valve Closing Time", columnIndex)
    inputs.valveType = DataStructures.GetParameterString(ws, ROW_VALVE_TYPE, columnIndex)
    inputs.supportType = DataStructures.GetParameterString(ws, ROW_PIPE_SUPPORT, columnIndex)
    
    ReadValveInputsFromColumn = inputs
End Function

' Get case type from the ValveList table
Private Function GetCaseTypeFromTable(tag As String) As String
    Dim ws As Worksheet
    Set ws = Sheets(SHEET_VALVELIST)
    
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(TABLE_VALVES)
    On Error GoTo 0
    
    If lo Is Nothing Then
        GetCaseTypeFromTable = "Valve Closure"  ' default
        Exit Function
    End If
    
    ' Search for the tag in the table
    Dim lr As ListRow
    For Each lr In lo.ListRows
        If Trim$(NzS(lr.Range.Cells(1, 1).value)) = tag Then
            GetCaseTypeFromTable = Trim$(NzS(lr.Range.Cells(1, 2).value))  ' 2nd column = CaseType
            Exit Function
        End If
    Next lr
    
    GetCaseTypeFromTable = "Valve Closure"  ' default if not found
End Function

' Write calculation results to the Results sheet
Private Sub WriteCalculationResults(ws As Worksheet, columnIndex As Long, tag As String, result As CalculationResult)
    ' Write header
    ws.Cells(1, columnIndex).value = tag
    ws.Cells(1, columnIndex).Font.Bold = True
    
    ' Write results
    ws.Cells(2, columnIndex).value = result.Ppeak
    ws.Cells(3, columnIndex).value = result.Fmax
    ws.Cells(4, columnIndex).value = result.Flim
    ws.Cells(5, columnIndex).value = result.LOF
    ws.Cells(6, columnIndex).value = result.FlagText
    
    ' Add labels if this is the first column
    If columnIndex = 2 Then
        ws.Cells(2, 1).value = "Ppeak (Pa)"
        ws.Cells(3, 1).value = "Fmax (kN)"
        ws.Cells(4, 1).value = "Flim (kN)"
        ws.Cells(5, 1).value = "LOF"
        ws.Cells(6, 1).value = "Flag"
    End If
End Sub

' ========= UTILITY FUNCTIONS =========

' Clear all valve data from Inputs sheet
Public Sub ClearValveData()
    Dim ws As Worksheet: Set ws = Sheets(SHEET_INPUTS)
    Dim lastCol As Long: lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    
    If lastCol >= INPUTS_FIRST_COL Then
        ws.Range(ws.Cells(2, INPUTS_FIRST_COL), ws.Cells(ws.Rows.Count, lastCol)).Clear
        MsgBox "Cleared all valve data from Inputs sheet.", vbInformation
    Else
        MsgBox "No valve data found to clear.", vbInformation
    End If
End Sub

' Refresh table validations and parameter highlighting
Public Sub RefreshValidations()
    Call DataStructures.RefreshAllTableValidations
    Call RefreshParameterHighlighting
    MsgBox "Table validations and parameter highlighting refreshed.", vbInformation
End Sub

' Null protection for String values
Private Function NzS(v) As String
    If IsError(v) Then Exit Function
    NzS = CStr(v)
End Function
