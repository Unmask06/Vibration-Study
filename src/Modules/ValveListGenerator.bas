Attribute VB_Name = "ValveListGenerator"
Option Explicit

' ========= VALVE LIST GENERATOR MODULE =========
' This module handles valve list generation and management
' Consolidates functionality from T28_TableDriven and T28_UI_Calc modules

' ========= CONFIGURATION CONSTANTS =========
Private Const INPUTS_FIRST_COL As Long = 5   ' "E" — first valve column in Inputs
Private Const SHEET_VALVELIST As String = "ValveList"
Private Const SHEET_INPUTS As String = "Inputs"
Private Const SHEET_RESULTS As String = "Results"
Private Const SHEET_DATA As String = "Data"
Private Const TABLE_VALVES As String = "tbValveList"

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
        Dim tag As String: tag = Trim$(NzS(lr.Range.Cells(1, 1).Value))  ' first column = Tag
        If tag <> "" Then
            ' Add column header
            wsIn.Cells(2, c).Value = tag
            With wsIn.Cells(2, c)
                .Interior.Color = RGB(31, 78, 120)
                .Font.Color = vbWhite
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
            End With
            
            ' Set Pipe Support Type and Valve Type from table
            Call SetParameterFromTable(wsIn, ROW_PIPE_SUPPORT, c, NzS(lr.Range.Cells(1, 4).Value))  ' 4th column = SupportType
            Call SetParameterFromTable(wsIn, ROW_VALVE_TYPE, c, NzS(lr.Range.Cells(1, 3).Value))    ' 3rd column = ValveType
            
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

' Alternative generation method from ValveList sheet (legacy support)
Public Sub Generate_Inputs_From_ValveList()
    Dim wsVL As Worksheet, wsIn As Worksheet
    Set wsVL = Sheets(SHEET_VALVELIST)
    Set wsIn = Sheets(SHEET_INPUTS)
    
    ' Read tags from worksheet
    Dim lastRow As Long: lastRow = wsVL.Cells(wsVL.Rows.Count, "A").End(xlUp).Row
    Dim tags As Collection: Set tags = New Collection
    Dim ct As Collection: Set ct = New Collection
    Dim vt As Collection: Set vt = New Collection
    Dim st As Collection: Set st = New Collection
    
    Dim r As Long
    For r = 3 To lastRow
        If Trim$(wsVL.Cells(r, "A").Value) <> "" Then
            tags.Add wsVL.Cells(r, "A").Value
            ct.Add wsVL.Cells(r, "B").Value
            vt.Add wsVL.Cells(r, "C").Value
            st.Add wsVL.Cells(r, "D").Value
        End If
    Next r
    
    If tags.Count = 0 Then
        MsgBox "No valve tags found.", vbExclamation
        Exit Sub
    End If
    
    ' Clear previous valve columns beyond D (first 4 columns are meta)
    Dim lastCol As Long: lastCol = wsIn.Cells(2, wsIn.Columns.Count).End(xlToLeft).Column
    If lastCol > 4 Then wsIn.Range(wsIn.Cells(2, 5), wsIn.Cells(wsIn.Rows.Count, lastCol)).Clear
    
    ' Write header columns for each tag
    Dim i As Long, c As Long: c = 5  ' start column for first valve
    For i = 1 To tags.Count
        wsIn.Cells(2, c).Value = tags(i)
        wsIn.Cells(2, c).Interior.Color = RGB(31, 78, 120)
        wsIn.Cells(2, c).Font.Color = vbWhite
        wsIn.Cells(2, c).Font.Bold = True
        wsIn.Cells(2, c).HorizontalAlignment = xlCenter
        
        ' Setup dropdowns and default values
        Call SetupValveColumnDropdowns(wsIn, c, vt(i), st(i))
        
        c = c + 1
    Next i
    
    MsgBox "Generated " & tags.Count & " valve column(s) in Inputs sheet.", vbInformation
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
        Dim tag As String: tag = Trim$(wsIn.Cells(2, c).Value)
        If tag <> "" Then
            ' Read inputs for this valve
            Dim inputs As ValveInputs
            inputs = ReadValveInputsFromColumn(wsIn, c)
            inputs.tag = tag
            
            ' Get case type from ValveList table
            inputs.casetype = GetCaseTypeFromTable(tag)
            
            ' Perform calculation
            Dim result As CalculationResult
            result = CalculationEngine.CalculateByCase(inputs.casetype, inputs)
            
            ' Write results
            Call WriteCalculationResults(wsRes, resultCol, tag, result)
            resultCol = resultCol + 1
        End If
    Next c
    
    MsgBox "Calculations completed for " & (resultCol - 2) & " valve(s).", vbInformation
End Sub

' ========= HELPER FUNCTIONS =========

' Set parameter value from table data
Private Sub SetParameterFromTable(ws As Worksheet, parameterName As String, columnIndex As Long, value As String)
    Dim rowIndex As Long
    rowIndex = DataStructures.GetParameterRow(parameterName)
    
    If rowIndex > 0 Then
        ws.Cells(rowIndex, columnIndex).Value = value
    End If
End Sub

' Setup dropdowns for valve column
Private Sub SetupValveColumnDropdowns(ws As Worksheet, columnIndex As Long, valveType As String, supportType As String)
    ' Find parameter rows
    Dim lastParamRow As Long
    lastParamRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim vtRow As Long, stRow As Long
    Dim r As Long
    For r = 3 To lastParamRow
        Dim paramName As String: paramName = Trim$(ws.Cells(r, "A").Value)
        If paramName = ROW_VALVE_TYPE Then vtRow = r
        If paramName = ROW_PIPE_SUPPORT Then stRow = r
    Next r
    
    ' Set valve type
    If vtRow > 0 Then
        ws.Cells(vtRow, columnIndex).Value = valveType
        ' Add validation from Data sheet
        Call DataStructures.ApplyValidationToRange(ws.Cells(vtRow, columnIndex), "D3:D10", SHEET_DATA, "Valve Type", "Select valve type")
    End If
    
    ' Set support type
    If stRow > 0 Then
        ws.Cells(stRow, columnIndex).Value = supportType
        ' Add validation from Data sheet
        Call DataStructures.ApplyValidationToRange(ws.Cells(stRow, columnIndex), "E3:E6", SHEET_DATA, "Support Type", "Select support type")
    End If
End Sub

' Read valve inputs from a specific column
Private Function ReadValveInputsFromColumn(ws As Worksheet, columnIndex As Long) As ValveInputs
    Dim inputs As ValveInputs
    
    ' Read all parameters using DataStructures module
    inputs.rho = DataStructures.GetParameterDouble(ws, "Fluid density", columnIndex)
    inputs.gamma = DataStructures.GetParameterDouble(ws, "Ratio of Specific Heat Capacities", columnIndex)
    inputs.c0 = DataStructures.GetParameterDouble(ws, "Speed of sound", columnIndex)
    inputs.Mw = DataStructures.GetParameterDouble(ws, "Molecular Weight", columnIndex)
    inputs.r = DataStructures.GetParameterDouble(ws, "Universal Gas Constant", columnIndex)
    inputs.Te = DataStructures.GetParameterDouble(ws, "Upstream Temperature", columnIndex)
    inputs.Pv = DataStructures.GetParameterDouble(ws, "Vapour Pressure", columnIndex)
    inputs.Kbulk = DataStructures.GetParameterDouble(ws, "Fluid Bulk Modulus", columnIndex)
    
    inputs.Dext_mm = DataStructures.GetParameterDouble(ws, "External Main Line Diameter", columnIndex)
    inputs.Dint_mm = DataStructures.GetParameterDouble(ws, "Internal Main Line Diameter", columnIndex)
    inputs.T_mm = DataStructures.GetParameterDouble(ws, "Main line Wall Thickness", columnIndex)
    inputs.Tsch40 = DataStructures.GetParameterDouble(ws, "Main line Wall Thickness for Schedule 40 Piping", columnIndex)
    inputs.Em = DataStructures.GetParameterDouble(ws, "Young's Modulus of the main line material", columnIndex)
    inputs.Lup = DataStructures.GetParameterDouble(ws, "Upstream Pipe Length", columnIndex)
    
    inputs.P1 = DataStructures.GetParameterDouble(ws, "Upstream Static Pressure", columnIndex)
    inputs.dP = DataStructures.GetParameterDouble(ws, "Static Pressure drop", columnIndex)
    inputs.v = DataStructures.GetParameterDouble(ws, "Steady State Fluid Velocity", columnIndex)
    inputs.W = DataStructures.GetParameterDouble(ws, "Mass Flow Rate", columnIndex)
    inputs.Pshut = DataStructures.GetParameterDouble(ws, "Pump head at zero flow", columnIndex)
    
    inputs.Tclose = DataStructures.GetParameterDouble(ws, "Valve Closing Time", columnIndex)
    inputs.valvetype = DataStructures.GetParameterString(ws, ROW_VALVE_TYPE, columnIndex)
    inputs.supporttype = DataStructures.GetParameterString(ws, ROW_PIPE_SUPPORT, columnIndex)
    
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
        GetCaseTypeFromTable = "liqclose"  ' default
        Exit Function
    End If
    
    ' Search for the tag in the table
    Dim lr As ListRow
    For Each lr In lo.ListRows
        If Trim$(NzS(lr.Range.Cells(1, 1).Value)) = tag Then
            GetCaseTypeFromTable = Trim$(NzS(lr.Range.Cells(1, 2).Value))  ' 2nd column = CaseType
            Exit Function
        End If
    Next lr
    
    GetCaseTypeFromTable = "liqclose"  ' default if not found
End Function

' Write calculation results to the Results sheet
Private Sub WriteCalculationResults(ws As Worksheet, columnIndex As Long, tag As String, result As CalculationResult)
    ' Write header
    ws.Cells(1, columnIndex).Value = tag
    ws.Cells(1, columnIndex).Font.Bold = True
    
    ' Write results
    ws.Cells(2, columnIndex).Value = result.Ppeak
    ws.Cells(3, columnIndex).Value = result.Fmax
    ws.Cells(4, columnIndex).Value = result.Flim
    ws.Cells(5, columnIndex).Value = result.LOF
    ws.Cells(6, columnIndex).Value = result.FlagText
    
    ' Add labels if this is the first column
    If columnIndex = 2 Then
        ws.Cells(2, 1).Value = "Ppeak (Pa)"
        ws.Cells(3, 1).Value = "Fmax (kN)"
        ws.Cells(4, 1).Value = "Flim (kN)"
        ws.Cells(5, 1).Value = "LOF"
        ws.Cells(6, 1).Value = "Flag"
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

' Refresh table validations
Public Sub RefreshValidations()
    Call DataStructures.RefreshAllTableValidations
    MsgBox "Table validations refreshed.", vbInformation
End Sub

' Null protection for String values
Private Function NzS(v) As String
    If IsError(v) Then Exit Function
    NzS = CStr(v)
End Function