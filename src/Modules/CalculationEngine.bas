Attribute VB_Name = "CalculationEngine"
Option Explicit

' ========= CALCULATION ENGINE MODULE =========
' This module contains all calculation methods organized by case type

' Calculation result structure
Public Type CalculationResult
    Ppeak As Double
    Fmax As Double
    Flim As Double
    LOF As Double
    FlagText As String
End Type

' Main calculation dispatcher based on case type - pass values directly
Public Function CalculateByCase(casetype As String, rho As Double, gamma As Double, c0 As Double, _
                               Mw As Double, R As Double, Te As Double, Pv As Double, Kbulk As Double, _
                               Dext_mm As Double, Dint_mm As Double, T_mm As Double, Tsch40 As Double, _
                               Em As Double, Lup As Double, P1 As Double, dP As Double, v As Double, _
                               W As Double, Pshut As Double, Tclose As Double, valvetype As String, _
                               supporttype As String) As CalculationResult
    Dim result As CalculationResult
    
    ' Validate inputs first
    Dim validationErrors As String
    validationErrors = ValidateCalculationInputs(casetype, rho, gamma, c0, Mw, R, Te, Pv, Kbulk, _
                                                Dext_mm, Dint_mm, T_mm, Tsch40, Em, Lup, P1, dP, v, _
                                                W, Pshut, Tclose, valvetype, supporttype)
    
    If validationErrors <> "" Then
        result.FlagText = "Validation errors: " & validationErrors
        result.Ppeak = 0#
        result.Fmax = 0#
        result.Flim = 0#
        result.LOF = 0#
        CalculateByCase = result
        Exit Function
    End If
    
    ' Calculate common values first
    result.Flim = CalculateFlim(T_mm, Tsch40, Dext_mm, Dint_mm, supporttype)
    
    Select Case LCase(Trim(casetype))
        Case "liqclose"
            result = CalculateLiquidClose(rho, c0, v, Dint_mm, P1, Lup, result.Flim)
        Case "gasopenrapid"
            result = CalculateGasOpenRapid(W, gamma, R, Te, Mw, result.Flim)
        Case "liqopen"
            result = CalculateLiquidOpen(W, dP, rho, result.Flim)
        Case Else
            result.FlagText = "Unknown CaseType: " & casetype
    End Select
    
    CalculateByCase = result
End Function

' Calculate liquid closure case - pass values directly
Private Function CalculateLiquidClose(rho As Double, c0 As Double, v As Double, _
                                     Dint_mm As Double, Pi As Double, Lup As Double, _
                                     Flim As Double) As CalculationResult
    Dim result As CalculationResult
    result.Flim = Flim
    
    ' Check if upstream pipe length requires detailed surge analysis
    If Lup > 100# Then
        result.LOF = 1#
        result.FlagText = "Lup>100 m → Detailed surge analysis"
        result.Ppeak = 0#
        result.Fmax = 0#
    Else
        ' Calculate peak pressure using Joukowsky equation: ΔP = ρ * c₀ * v
        Dim pressureSurge As Double
        pressureSurge = rho * c0 * v
        result.Ppeak = pressureSurge
        
        ' Calculate pipe cross-sectional area (convert mm to m²)
        Dim pipeArea As Double
        pipeArea = WorksheetFunction.Pi() * (Dint_mm / 1000#) ^ 2 / 4#
        
        ' Calculate maximum force: F = P * A (convert Pa·m² to kN)
        result.Fmax = result.Ppeak * pipeArea / 1000#

        If result.Fmax < 1 Then
            result.Fmax = 0#
            result.FlagText = "Negative force adjusted to zero"
            result.LOF = 0#
        End If

        If result.Fmax > 1 Then
            ' Calculate surge pressure using the provided formula
            ' Ω = ρ * v * L_up * φ / P_i
            Dim Omega As Double
            Dim phi As Double: phi = 1# ' Assuming φ = 1 for standard calculation
            Omega = rho * v * Lup * phi / Pi
            
            ' P_surge = P_i * (Ω/2 + Ω² * sqrt(1/4 + 1/Ω²))
            Dim surgeFactor As Double
            surgeFactor = (1 / 2) + (Sqr((1 / 4) + (1 / (Omega ^ 2))))
            Dim P_Surge As Double
            P_Surge = (Pi * Omega ^ 2) * surgeFactor
            
            ' if valve downstream of pump then Ptotal = Psurge+Pshutin-Pi
            'if not then Ptotal = Psurge+Pshutin
            ' Update peak pressure with surge component
            result.Fmax = result.Ppeak * pipeArea / 1000#
        End If
        
        ' Calculate Level of Force (LOF) ratio
        result.LOF = SafeDiv(result.Fmax, result.Flim)
        result.FlagText = ""
    End If
    
    CalculateLiquidClose = result
End Function

' Calculate gas opening rapid case - pass values directly
Private Function CalculateGasOpenRapid(W As Double, gamma As Double, R As Double, _
                                      Tc As Double, Mw As Double, Flim As Double) As CalculationResult
    Dim result As CalculationResult
    result.Flim = Flim
    
    ' Calculate peak force using dry gas rapid opening formula
    ' Fmax = (W/1000) * sqrt((2*γ*R*Tc)/((γ+1)*Mw))
    
    If W > 0 And gamma > 0 And R > 0 And Tc > 0 And Mw > 0 Then
        Dim numerator As Double
        Dim denominator As Double
        
        numerator = 2# * gamma * R * Tc
        denominator = (gamma + 1#) * Mw
        
        If denominator > 0 Then
            result.Fmax = (W / 1000#) * Sqr(numerator / denominator)
            result.LOF = SafeDiv(result.Fmax, result.Flim)
            result.FlagText = ""
        Else
            result.Fmax = 0#
            result.LOF = 0#
            result.FlagText = "Invalid gas properties for calculation"
        End If
    Else
        result.Fmax = 0#
        result.LOF = 0#
        result.FlagText = "Missing required gas properties (W, γ, R, Tc, Mw)"
    End If
    
    CalculateGasOpenRapid = result
End Function

' Calculate liquid opening case - pass values directly
Private Function CalculateLiquidOpen(W As Double, dP As Double, rho As Double, _
                                    Flim As Double) As CalculationResult
    Dim result As CalculationResult
    result.Flim = Flim
    
    ' Calculate peak force using liquid/multiphase opening formula
    ' Fmax = (1/1.58) * W * sqrt(ΔP/ρ)
    
    If W > 0 And dP > 0 And rho > 0 Then
        result.Fmax = (1# / 1.58) * W * Sqr(dP / rho)
        result.LOF = SafeDiv(result.Fmax, result.Flim)
        result.FlagText = ""
    Else
        result.Fmax = 0#
        result.LOF = 0#
        result.FlagText = "Missing required parameters (W, ΔP, ρ) for liquid opening calculation"
    End If
    
    CalculateLiquidOpen = result
End Function

' Calculate force limit (Flim) - pass values directly
Private Function CalculateFlim(T_mm As Double, Tsch40 As Double, Dext_mm As Double, _
                              Dint_mm As Double, supporttype As String) As Double
    Dim Psi As Double
    If Tsch40 > 0 Then Psi = T_mm / Tsch40
    
    Dim theta As Double
    theta = ThetaFromSupport(supporttype)
    
    CalculateFlim = Flim_kN_EI(Psi, Dext_mm / 1000#, theta, Dint_mm / 1000#)
End Function

' Calculate wave speed if not provided - pass values directly
Public Function CalculateWaveSpeed(c0 As Double, rho As Double, Dext_mm As Double, _
                                  T_mm As Double, Em As Double, Kbulk As Double) As Double
    ' Use provided speed of sound if available
    If c0 > 0 Then
        CalculateWaveSpeed = c0
        Exit Function
    End If
    
    ' If not provided, calculate from other parameters
    If rho > 0 And Dext_mm > 0 And T_mm > 0 And Em > 0 And Kbulk > 0 Then
        CalculateWaveSpeed = WaveSpeed_EI(rho, Dext_mm, T_mm, Em, Kbulk)
    Else
        CalculateWaveSpeed = 0
    End If
End Function

' ========= HELPER FUNCTIONS =========

Private Function ThetaFromSupport(ByVal support As String) As Double
    If InStr(1, support, "Anchor", vbTextCompare) > 0 Then ThetaFromSupport = 4#: Exit Function
    If InStr(1, support, "Guide", vbTextCompare) > 0 Then ThetaFromSupport = 2#: Exit Function
    If InStr(1, support, "Sliding", vbTextCompare) > 0 Then ThetaFromSupport = 1#: Exit Function
    If InStr(1, support, "None", vbTextCompare) > 0 Then ThetaFromSupport = 0.5: Exit Function
    ThetaFromSupport = 1#  ' default
End Function

Private Function WaveSpeed_EI(ByVal rho As Double, ByVal Dext_mm As Double, _
                              ByVal T_mm As Double, ByVal Em As Double, ByVal Kbulk As Double) As Double
    If rho <= 0 Or Dext_mm <= 0 Or T_mm <= 0 Or Em <= 0 Or Kbulk <= 0 Then Exit Function
    Dim term1 As Double, term2 As Double
    term1 = 1# / Kbulk
    term2 = Dext_mm / (T_mm * Em * 1000#)
    Dim denom As Double: denom = rho * (term1 + term2)
    If denom > 0 Then WaveSpeed_EI = 1# / denom
End Function

Private Function Flim_kN_EI(ByVal Psi As Double, ByVal Dext_m As Double, _
                            ByVal theta As Double, ByVal Dint_m As Double) As Double
    Dim poly As Double
    poly = 16.8 * Psi ^ 3 - 1.81 * Psi ^ 2 + 525# * Psi + 25.3
    Dim area As Double: area = WorksheetFunction.Pi() * Dint_m ^ 2 / 4#
    Flim_kN_EI = poly * Dext_m * theta * area / 1000000000#
End Function

Private Function SafeDiv(ByVal num As Double, ByVal den As Double) As Double
    If den <= 0 Then SafeDiv = 0# Else SafeDiv = num / den
End Function

' ========= UTILITY FUNCTIONS =========

' Validate required parameters for calculation - pass values directly
Public Function ValidateCalculationInputs(casetype As String, rho As Double, gamma As Double, c0 As Double, _
                                         Mw As Double, R As Double, Te As Double, Pv As Double, Kbulk As Double, _
                                         Dext_mm As Double, Dint_mm As Double, T_mm As Double, Tsch40 As Double, _
                                         Em As Double, Lup As Double, P1 As Double, dP As Double, v As Double, _
                                         W As Double, Pshut As Double, Tclose As Double, valvetype As String, _
                                         supporttype As String) As String
    Dim errors As String
    
    ' Common validations for all case types
    If rho <= 0 Then errors = errors & "Fluid density must be positive. "
    If Dext_mm <= 0 Then errors = errors & "External diameter must be positive. "
    If Dint_mm <= 0 Then errors = errors & "Internal diameter must be positive. "
    If T_mm <= 0 Then errors = errors & "Wall thickness must be positive. "
    
    ' Check logical relationships
    If Dint_mm >= Dext_mm Then errors = errors & "Internal diameter must be less than external diameter. "
    If T_mm >= Dext_mm / 2 Then errors = errors & "Wall thickness too large for given external diameter. "
    
    ' Case-specific validations
    Select Case LCase(Trim(casetype))
        Case "liqclose"
            If c0 <= 0 Then errors = errors & "Speed of sound required for liquid closure. "
            If v <= 0 Then errors = errors & "Fluid velocity required for liquid closure. "
            
        Case "gasopenrapid"
            If W <= 0 Then errors = errors & "Mass flow rate required for gas opening. "
            If gamma <= 0 Then errors = errors & "Specific heat ratio required for gas opening. "
            
        Case "liqopen"
            If W <= 0 Then errors = errors & "Mass flow rate required for liquid opening. "
            If dP <= 0 Then errors = errors & "Pressure drop required for liquid opening. "
    End Select
    
    ValidateCalculationInputs = Trim$(errors)
End Function

' Get basic project information
Public Function GetProjectInfo() As String
    GetProjectInfo = "Vibration Study Calculator - " & _
                    "Workbook: " & ThisWorkbook.Name & " - " & _
                    "Date: " & Format(Date, "yyyy-mm-dd")
End Function

' ========= EXAMPLE USAGE =========
' Example function showing how to calculate using direct parameter values
Public Sub ExampleDirectCalculation()
    ' Method 1: Read from specific cells and pass values directly
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Input") ' Adjust sheet name as needed
    
    ' Read values from any cells you want (user can select any format)
    Dim rho As Double: rho = ws.Range("D3").Value        ' Fluid density from D3
    Dim gamma As Double: gamma = ws.Range("D4").Value    ' Gamma from D4
    Dim c0 As Double: c0 = ws.Range("D5").Value          ' Speed of sound from D5
    Dim Mw As Double: Mw = ws.Range("D6").Value          ' Molecular weight from D6
    Dim R As Double: R = ws.Range("D7").Value            ' Gas constant from D7
    Dim Te As Double: Te = ws.Range("D8").Value          ' Temperature from D8
    Dim Pv As Double: Pv = 0                            ' Vapour pressure (optional)
    Dim Kbulk As Double: Kbulk = ws.Range("D10").Value  ' Bulk modulus from D10
    Dim Dext_mm As Double: Dext_mm = ws.Range("D11").Value    ' External diameter
    Dim Dint_mm As Double: Dint_mm = ws.Range("D12").Value    ' Internal diameter
    Dim T_mm As Double: T_mm = ws.Range("D13").Value          ' Wall thickness
    Dim Tsch40 As Double: Tsch40 = ws.Range("D14").Value      ' Schedule 40 thickness
    Dim Em As Double: Em = ws.Range("D15").Value              ' Young's modulus
    Dim Lup As Double: Lup = ws.Range("D18").Value            ' Upstream pipe length
    Dim P1 As Double: P1 = ws.Range("D19").Value              ' Upstream pressure
    Dim dP As Double: dP = ws.Range("D20").Value              ' Pressure drop
    Dim v As Double: v = ws.Range("D21").Value                ' Fluid velocity
    Dim W As Double: W = ws.Range("D22").Value                ' Mass flow rate
    Dim Pshut As Double: Pshut = 0                            ' Pump shutoff head (optional)
    Dim Tclose As Double: Tclose = 0                          ' Valve closing time (optional)
    Dim valvetype As String: valvetype = "Ball"               ' Valve type
    Dim supporttype As String: supporttype = "Anchor"        ' Support type
    Dim casetype As String: casetype = "liqclose"             ' Case type
    
    ' Perform calculation with direct values
    Dim result As CalculationResult
    result = CalculateByCase(casetype, rho, gamma, c0, Mw, R, Te, Pv, Kbulk, _
                           Dext_mm, Dint_mm, T_mm, Tsch40, Em, Lup, P1, dP, v, _
                           W, Pshut, Tclose, valvetype, supporttype)
    
    ' Display results
    If result.FlagText <> "" Then
        MsgBox "Calculation completed with notes: " & result.FlagText, vbInformation
    Else
        MsgBox "Results:" & vbCrLf & _
               "Peak Pressure: " & Format(result.Ppeak, "#,##0") & " Pa" & vbCrLf & _
               "Max Force: " & Format(result.Fmax, "#,##0.0") & " kN" & vbCrLf & _
               "Force Limit: " & Format(result.Flim, "#,##0.0") & " kN" & vbCrLf & _
               "Level of Force: " & Format(result.LOF, "0.00"), vbInformation
    End If
End Sub

' Method 2: Even simpler - hardcode values for testing
Public Sub ExampleHardcodedCalculation()
    Dim result As CalculationResult
    result = CalculateByCase("liqclose", _
                           1000, 1.4, 1500, 18, 8314, 298, 0, 2000000000#, _
                           100, 80, 10, 8, 200000000000#, 50, 5000000, 500000, 2.5, _
                           10, 0, 0, "Ball", "Anchor")
    
    MsgBox "Test Result - Max Force: " & result.Fmax & " kN, LOF: " & result.LOF
End Sub

' Method 3: Let user select cells interactively
Public Sub ExampleInteractiveCalculation()
    On Error GoTo ErrorHandler
    
    ' Ask user to select cells for key parameters
    Dim rho As Double: rho = Val(Application.InputBox("Select cell for Fluid Density (kg/m³):", Type:=8).Value)
    Dim c0 As Double: c0 = Val(Application.InputBox("Select cell for Speed of Sound (m/s):", Type:=8).Value)
    Dim v As Double: v = Val(Application.InputBox("Select cell for Fluid Velocity (m/s):", Type:=8).Value)
    Dim Dint_mm As Double: Dint_mm = Val(Application.InputBox("Select cell for Internal Diameter (mm):", Type:=8).Value)
    Dim P1 As Double: P1 = Val(Application.InputBox("Select cell for Upstream Pressure (Pa):", Type:=8).Value)
    Dim Lup As Double: Lup = Val(Application.InputBox("Select cell for Upstream Pipe Length (m):", Type:=8).Value)
    
    ' Use default values for other parameters
    Dim result As CalculationResult
    result = CalculateByCase("liqclose", rho, 1.4, c0, 18, 8314, 298, 0, 2000000000#, _
                           Dint_mm + 20, Dint_mm, 10, 8, 200000000000#, Lup, P1, 500000, v, _
                           10, 0, 0, "Ball", "Anchor")
    
    MsgBox "Interactive Result:" & vbCrLf & _
           "Max Force: " & Format(result.Fmax, "#,##0.0") & " kN" & vbCrLf & _
           "Level of Force: " & Format(result.LOF, "0.00")
    Exit Sub
    
ErrorHandler:
    MsgBox "Calculation cancelled or invalid cell selected.", vbExclamation
End Sub

' ========= WRAPPER FUNCTIONS FOR BACKWARD COMPATIBILITY =========

' Wrapper function for existing code that uses worksheet/column approach
Public Function CalculateByCaseFromSheet(casetype As String, ws As Worksheet, columnIndex As Long) As CalculationResult
    ' Read all parameters from the worksheet using direct cell references
    ' Note: Adjust row numbers to match your Excel layout
    Dim rho As Double: rho = Val(ws.Cells(3, columnIndex).Value)       ' Fluid density
    Dim gamma As Double: gamma = Val(ws.Cells(4, columnIndex).Value)   ' Ratio of Specific Heat Capacities
    Dim c0 As Double: c0 = Val(ws.Cells(5, columnIndex).Value)         ' Speed of sound
    Dim Mw As Double: Mw = Val(ws.Cells(6, columnIndex).Value)         ' Molecular Weight
    Dim R As Double: R = Val(ws.Cells(7, columnIndex).Value)           ' Universal Gas Constant
    Dim Te As Double: Te = Val(ws.Cells(8, columnIndex).Value)         ' Upstream Temperature
    Dim Pv As Double: Pv = Val(ws.Cells(9, columnIndex).Value)         ' Vapour Pressure
    Dim Kbulk As Double: Kbulk = Val(ws.Cells(10, columnIndex).Value)  ' Fluid Bulk Modulus
    Dim Dext_mm As Double: Dext_mm = Val(ws.Cells(11, columnIndex).Value) ' External Main Line Diameter
    Dim Dint_mm As Double: Dint_mm = Val(ws.Cells(12, columnIndex).Value) ' Internal Main Line Diameter
    Dim T_mm As Double: T_mm = Val(ws.Cells(13, columnIndex).Value)    ' Main line Wall Thickness
    Dim Tsch40 As Double: Tsch40 = Val(ws.Cells(14, columnIndex).Value) ' Schedule 40 Wall Thickness
    Dim Em As Double: Em = Val(ws.Cells(15, columnIndex).Value)        ' Young's Modulus
    Dim Lup As Double: Lup = Val(ws.Cells(18, columnIndex).Value)      ' Upstream Pipe Length
    Dim P1 As Double: P1 = Val(ws.Cells(19, columnIndex).Value)        ' Upstream Static Pressure
    Dim dP As Double: dP = Val(ws.Cells(20, columnIndex).Value)        ' Static Pressure drop
    Dim v As Double: v = Val(ws.Cells(21, columnIndex).Value)          ' Steady State Fluid Velocity
    Dim W As Double: W = Val(ws.Cells(22, columnIndex).Value)          ' Mass Flow Rate
    Dim Pshut As Double: Pshut = Val(ws.Cells(23, columnIndex).Value)  ' Pump head at zero flow
    Dim Tclose As Double: Tclose = Val(ws.Cells(24, columnIndex).Value) ' Valve Closing Time
    Dim valvetype As String: valvetype = Trim$(ws.Cells(25, columnIndex).Value) ' Valve Type
    Dim supporttype As String: supporttype = Trim$(ws.Cells(26, columnIndex).Value) ' Pipe Support Type
    
    ' Call the main calculation function with all parameters
    CalculateByCaseFromSheet = CalculateByCase(casetype, rho, gamma, c0, Mw, R, Te, Pv, Kbulk, _
                                             Dext_mm, Dint_mm, T_mm, Tsch40, Em, Lup, P1, dP, v, _
                                             W, Pshut, Tclose, valvetype, supporttype)
End Function

' Helper function - calculate from any range of cells (user-friendly)
Public Function CalculateFromRange(casetype As String, parameterRange As Range) As CalculationResult
    ' Assumes parameterRange contains values in order as listed in parameter reference
    ' Example: =CalculateFromRange("liqclose", A1:A23)
    
    If parameterRange.Cells.Count < 20 Then
        Dim result As CalculationResult
        result.FlagText = "Parameter range must contain at least 20 values"
        CalculateFromRange = result
        Exit Function
    End If
    
    ' Read parameters from range (1-based indexing)
    Dim rho As Double: rho = Val(parameterRange.Cells(1).Value)         ' Fluid density
    Dim gamma As Double: gamma = Val(parameterRange.Cells(2).Value)     ' Ratio of Specific Heat Capacities
    Dim c0 As Double: c0 = Val(parameterRange.Cells(3).Value)           ' Speed of sound
    Dim Mw As Double: Mw = Val(parameterRange.Cells(4).Value)           ' Molecular Weight
    Dim R As Double: R = Val(parameterRange.Cells(5).Value)             ' Universal Gas Constant
    Dim Te As Double: Te = Val(parameterRange.Cells(6).Value)           ' Upstream Temperature
    Dim Pv As Double: Pv = 0: If parameterRange.Cells.Count > 6 Then Pv = Val(parameterRange.Cells(7).Value)
    Dim Kbulk As Double: Kbulk = 0: If parameterRange.Cells.Count > 7 Then Kbulk = Val(parameterRange.Cells(8).Value)
    Dim Dext_mm As Double: Dext_mm = Val(parameterRange.Cells(9).Value) ' External Main Line Diameter
    Dim Dint_mm As Double: Dint_mm = Val(parameterRange.Cells(10).Value) ' Internal Main Line Diameter
    Dim T_mm As Double: T_mm = Val(parameterRange.Cells(11).Value)      ' Main line Wall Thickness
    Dim Tsch40 As Double: Tsch40 = Val(parameterRange.Cells(12).Value)  ' Schedule 40 Wall Thickness
    Dim Em As Double: Em = Val(parameterRange.Cells(13).Value)          ' Young's Modulus
    Dim Lup As Double: Lup = Val(parameterRange.Cells(14).Value)        ' Upstream Pipe Length
    Dim P1 As Double: P1 = Val(parameterRange.Cells(15).Value)          ' Upstream Static Pressure
    Dim dP As Double: dP = Val(parameterRange.Cells(16).Value)          ' Static Pressure drop
    Dim v As Double: v = Val(parameterRange.Cells(17).Value)            ' Steady State Fluid Velocity
    Dim W As Double: W = Val(parameterRange.Cells(18).Value)            ' Mass Flow Rate
    Dim Pshut As Double: Pshut = 0: If parameterRange.Cells.Count > 18 Then Pshut = Val(parameterRange.Cells(19).Value)
    Dim Tclose As Double: Tclose = 0: If parameterRange.Cells.Count > 19 Then Tclose = Val(parameterRange.Cells(20).Value)
    Dim valvetype As String: valvetype = "Ball": If parameterRange.Cells.Count > 20 Then valvetype = Trim$(parameterRange.Cells(21).Value)
    Dim supporttype As String: supporttype = "Anchor": If parameterRange.Cells.Count > 21 Then supporttype = Trim$(parameterRange.Cells(22).Value)
    
    ' Call the main calculation function
    CalculateFromRange = CalculateByCase(casetype, rho, gamma, c0, Mw, R, Te, Pv, Kbulk, _
                                       Dext_mm, Dint_mm, T_mm, Tsch40, Em, Lup, P1, dP, v, _
                                       W, Pshut, Tclose, valvetype, supporttype)
End Function

' ========= PARAMETER REFERENCE =========
' The main calculation function now accepts all parameters directly:
'
' CalculateByCase(casetype, rho, gamma, c0, Mw, R, Te, Pv, Kbulk,
'                 Dext_mm, Dint_mm, T_mm, Tsch40, Em, Lup, P1, dP, v,
'                 W, Pshut, Tclose, valvetype, supporttype)
'
' Where:
' casetype    = "liqclose", "gasopenrapid", or "liqopen"
' rho         = Fluid density (kg/m³)
' gamma       = Ratio of Specific Heat Capacities
' c0          = Speed of sound (m/s)
' Mw          = Molecular Weight (kg/kmol)
' R           = Universal Gas Constant (J/kmol·K)
' Te          = Upstream Temperature (K)
' Pv          = Vapour Pressure (Pa) - optional
' Kbulk       = Fluid Bulk Modulus (Pa)
' Dext_mm     = External Main Line Diameter (mm)
' Dint_mm     = Internal Main Line Diameter (mm)
' T_mm        = Main line Wall Thickness (mm)
' Tsch40      = Schedule 40 Wall Thickness (mm)
' Em          = Young's Modulus (Pa)
' Lup         = Upstream Pipe Length (m)
' P1          = Upstream Static Pressure (Pa)
' dP          = Static Pressure drop (Pa)
' v           = Steady State Fluid Velocity (m/s)
' W           = Mass Flow Rate (kg/s)
' Pshut       = Pump head at zero flow (Pa) - optional
' Tclose      = Valve Closing Time (s) - optional
' valvetype   = Valve Type string
' supporttype = Pipe Support Type ("Anchor", "Guide", "Sliding", "None")
'
' BENEFITS:
' ✓ Works with ANY Excel format/layout
' ✓ User can select any cells they want
' ✓ No dependencies on specific row/column positions
' ✓ Can be called from anywhere with direct values
' ✓ Easy to integrate with other systems
