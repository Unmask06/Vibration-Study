Attribute VB_Name = "ParameterManager"
Option Explicit

' ========= PARAMETER MANAGER MODULE =========
' This module handles parameter row indexing and lookup operations
' for the table-driven valve calculations

Private Type ParameterIndex
    Dictionary As Object
    lastRow As Long
    IsInitialized As Boolean
End Type

Private paramIdx As ParameterIndex

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
        Err.Raise 9999, "ParameterManager", "Parameter index not initialized. Call InitializeParameterIndex first."
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
Public Function GetAllParameterNames() As String()
    If Not paramIdx.IsInitialized Then
        Err.Raise 9999, "ParameterManager", "Parameter index not initialized. Call InitializeParameterIndex first."
    End If
    
    Dim keys As Variant
    keys = paramIdx.Dictionary.keys
    
    Dim result() As String
    ReDim result(0 To UBound(keys))
    
    Dim i As Long
    For i = 0 To UBound(keys)
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

' Helper functions (moved from main module for better modularity)
Private Function NzD(v) As Double
    If IsError(v) Then NzD = 0#: Exit Function
    If Len(Trim$(v & "")) = 0 Then NzD = 0#: Exit Function
    NzD = CDbl(v)
End Function

Private Function NzS(v) As String
    If IsError(v) Then Exit Function
    NzS = CStr(v)
End Function

' ========= UNIT CONVERSION FUNCTIONS =========
' Convert pressure from barg to Pa (absolute pressure in Pascals)
' barg (bar gauge) = bar above atmospheric pressure
' Pa = (barg + 1.01325) × 100000 (converting to absolute pressure in Pa)
Public Function BargToPa(ByVal pressureBarg As Double) As Double
    ' Convert barg to absolute bar, then to Pa
    ' 1 bar = 100000 Pa
    ' Atmospheric pressure = 1.01325 bar
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

' ========= WORKSHEET HELPER FUNCTIONS =========
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
