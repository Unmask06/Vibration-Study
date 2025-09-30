Attribute VB_Name = "T28_TableDriven_New"
Option Explicit

' ========= CONFIG =========
Private Const INPUTS_FIRST_COL As Long = 5   ' "E" — first valve column in Inputs
Private Const SHEET_VALVELIST As String = "ValveList"
Private Const SHEET_INPUTS As String = "Inputs"
Private Const SHEET_RESULTS As String = "Results"
Private Const TABLE_VALVES As String = "tbValveList"

' Row labels in Inputs (must match the left-most "Parameter" text)
Private Const ROW_PIPE_SUPPORT As String = "Pipe Support Type"
Private Const ROW_VALVE_TYPE As String = "Valve Type"

' Initialize the application modules
Private Sub InitializeModules()
    Dim wsIn As Worksheet: Set wsIn = Sheets(SHEET_INPUTS)
    Call ParameterManager.InitializeParameterIndex(wsIn)
    Call TableEventHandler.SetupTableValidations
End Sub

' ========= PUBLIC MACROS =========
' Generates Inputs sheet columns directly from the Excel Table tbValveList (Tag, CaseType)
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
        If tag <> "" Then
            ' Add column header
            wsIn.Cells(2, c).value = tag
            With wsIn.Cells(2, c)
                .HorizontalAlignment = xlCenter
                .Font.Bold = True
                .Interior.Color = RGB(31, 78, 120)
                .Font.Color = vbWhite
            End With
            
            added = added + 1
            c = c + 1
        End If
    Next lr
    
    MsgBox "Inputs generated for " & added & " valve(s) from tbValveList.", vbInformation
End Sub

' Calculates all valves listed in tbValveList, reading ValveType and SupportType from Inputs rows per valve column.
Public Sub Calculate_All_From_tbValveList()
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
    Dim wsRes As Worksheet: Set wsRes = Sheets(SHEET_RESULTS)
    
    ' Clear results (keep headers)
    Dim resLast As Long: resLast = wsRes.Cells(wsRes.Rows.Count, "A").End(xlUp).row
    If resLast >= 3 Then wsRes.Rows("3:" & resLast).ClearContents
    
    ' For each ListRow in the table, find its matching column in Inputs by Tag (header in row 2)
    Dim lr As ListRow, resRow As Long: resRow = 3
    For Each lr In lo.ListRows
        Dim tag As String: tag = Trim$(NzS(lr.Range.Cells(1, 1).value))      ' Col1: Tag
        Dim casetype As String: casetype = Trim$(NzS(lr.Range.Cells(1, 2).value)) ' Col2: CaseType
        If tag = "" Then GoTo NextValve
        
        Dim c As Long: c = FindValveColumn(wsIn, tag)
        If c = 0 Then GoTo NextValve
        
        ' Build inputs structure using the new modular approach
        Dim inputs As ValveInputs
        inputs = BuildValveInputs(wsIn, c, tag, casetype)
        
        ' Calculate wave speed if needed
        If inputs.c0 <= 0 Then
            inputs.c0 = CalculationEngine.CalculateWaveSpeed(inputs)
        End If
        
        ' Perform calculation using the modular calculation engine
        Dim result As CalculationResult
        result = CalculationEngine.CalculateByCase(casetype, inputs)
        
        ' Write results
        With wsRes
            .Cells(resRow, 1).value = tag
            .Cells(resRow, 2).value = casetype
            .Cells(resRow, 3).value = inputs.valvetype
            .Cells(resRow, 4).value = inputs.supporttype
            .Cells(resRow, 5).value = result.Ppeak
            .Cells(resRow, 6).value = result.Fmax
            .Cells(resRow, 7).value = result.Flim
            .Cells(resRow, 8).value = result.LOF
            .Cells(resRow, 9).value = result.FlagText
        End With
        
        resRow = resRow + 1
NextValve:
    Next lr
    
    MsgBox "Calculation complete for " & (resRow - 3) & " valve(s).", vbInformation
End Sub

' ========= HELPERS =========
' Build ValveInputs structure from worksheet data
Private Function BuildValveInputs(wsIn As Worksheet, columnIndex As Long, tag As String, casetype As String) As ValveInputs
    Dim inputs As ValveInputs
    
    ' Basic identification
    inputs.tag = tag
    inputs.casetype = casetype
    
    ' Read all parameters using the ParameterManager
    inputs.rho = ParameterManager.GetParameterDouble(wsIn, "Fluid density", columnIndex)
    inputs.gamma = ParameterManager.GetParameterDouble(wsIn, "Ratio of Specific Heat Capacities", columnIndex)
    inputs.c0 = ParameterManager.GetParameterDouble(wsIn, "Speed of sound", columnIndex)
    inputs.Dext_mm = ParameterManager.GetParameterDouble(wsIn, "External Main Line Diameter", columnIndex)
    inputs.Dint_mm = ParameterManager.GetParameterDouble(wsIn, "Internal Main Line Diameter", columnIndex)
    inputs.Em = ParameterManager.GetParameterDouble(wsIn, "Young's Modulus of the main line material", columnIndex)
    inputs.Kbulk = ParameterManager.GetParameterDouble(wsIn, "Fluid Bulk Modulus", columnIndex)
    inputs.Lup = ParameterManager.GetParameterDouble(wsIn, "Upstream Pipe Length", columnIndex)
    inputs.Mw = ParameterManager.GetParameterDouble(wsIn, "Molecular Weight", columnIndex)
    ' Pressure parameters - convert from barg input to Pa for calculations
    inputs.P1 = ParameterManager.GetPressureParameterPa(wsIn, "Upstream Static Pressure", columnIndex)
    inputs.Pshut = ParameterManager.GetPressureParameterPa(wsIn, "Pump head at zero flow", columnIndex)
    inputs.Pv = ParameterManager.GetPressureParameterPa(wsIn, "Vapour Pressure", columnIndex)
    inputs.dP = ParameterManager.GetPressureParameterPa(wsIn, "Static Pressure drop", columnIndex)
    inputs.r = ParameterManager.GetParameterDouble(wsIn, "Universal Gas Constant", columnIndex)
    inputs.T_mm = ParameterManager.GetParameterDouble(wsIn, "Main line Wall Thickness", columnIndex)
    inputs.Tclose = ParameterManager.GetParameterDouble(wsIn, "Valve Closing Time", columnIndex)
    inputs.Te = ParameterManager.GetParameterDouble(wsIn, "Upstream Temperature", columnIndex)
    inputs.v = ParameterManager.GetParameterDouble(wsIn, "Steady State Fluid Velocity", columnIndex)
    inputs.W = ParameterManager.GetParameterDouble(wsIn, "Mass Flow Rate", columnIndex)
    inputs.Tsch40 = ParameterManager.GetParameterDouble(wsIn, "Main line Wall Thickness for Schedule 40 Piping", columnIndex)
    inputs.valvetype = ParameterManager.GetParameterString(wsIn, ROW_VALVE_TYPE, columnIndex)
    inputs.supporttype = ParameterManager.GetParameterString(wsIn, ROW_PIPE_SUPPORT, columnIndex)
    
    BuildValveInputs = inputs
End Function

Private Function FindValveColumn(wsIn As Worksheet, ByVal tag As String) As Long
    Dim lastCol As Long: lastCol = wsIn.Cells(2, wsIn.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = INPUTS_FIRST_COL To lastCol
        If Trim$(wsIn.Cells(2, c).value) = tag Then
            FindValveColumn = c
            Exit Function
        End If
    Next c
    FindValveColumn = 0
End Function

' Helper functions for backward compatibility
Private Function NzD(v) As Double
    If IsError(v) Then NzD = 0#: Exit Function
    If Len(Trim$(v & "")) = 0 Then NzD = 0#: Exit Function
    NzD = CDbl(v)
End Function

Private Function NzS(v) As String
    If IsError(v) Then Exit Function
    NzS = CStr(v)
End Function

' Public function to setup data validation for tbValveList
Public Sub Setup_ValveList_DataValidation()
    Call TableEventHandler.SetupTableValidations
    MsgBox "Data validation has been set up for the valve list table.", vbInformation
End Sub
