Attribute VB_Name = "DataStructures"
Option Explicit

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
    valvetype As String             ' Valve Type
    supporttype As String           ' Pipe Support Type
    
    ' Identification
    tag As String                   ' Valve tag
    casetype As String              ' Case type for calculation
End Type

' Structure for validation settings
Public Type ValidationSettings
    Sheet As String
    Range As String
    title As String
    message As String
End Type

' Structure for table configuration
Public Type TableConfig
    sheetName As String
    tableName As String
    Columns As Object               ' Dictionary of column names and their settings
End Type

' Enumeration for calculation case types
Public Enum CaseTypes
    LiquidClose = 1
    GasOpenRapid = 2
    LiquidOpen = 3
End Enum

' Enumeration for support types
Public Enum SupportTypes
    Anchor = 1
    Guide = 2
    Sliding = 3
    None = 4
End Enum

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
    If Trim$(inputs.casetype) = "" Then errors = errors & "Case type cannot be empty. "
    
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
    Target.valvetype = source.valvetype
    Target.supporttype = source.supporttype
    Target.tag = source.tag
    Target.casetype = source.casetype
    
    CopyValveInputs = Target
End Function
