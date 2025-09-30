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

' Main calculation dispatcher based on case type
Public Function CalculateByCase(casetype As String, inputs As ValveInputs) As CalculationResult
    Dim result As CalculationResult
    
    ' Calculate common values first
    result.Flim = CalculateFlim(inputs)
    
    Select Case LCase$(Trim$(casetype))
        Case "liqclose"
            result = CalculateLiquidClose(inputs, result.Flim)
        Case "gasopenrapid"
            result = CalculateGasOpenRapid(inputs, result.Flim)
        Case "liqopen"
            result = CalculateLiquidOpen(inputs, result.Flim)
        Case Else
            result.FlagText = "Unknown CaseType: " & casetype
    End Select
    
    CalculateByCase = result
End Function

' Calculate liquid closure case
Private Function CalculateLiquidClose(inputs As ValveInputs, Flim As Double) As CalculationResult
    Dim result As CalculationResult
    result.Flim = Flim
    
    If inputs.Lup > 100# Then
        result.LOF = 1#
        result.FlagText = "Lup>100 m → Detailed surge analysis"
    Else
        result.Ppeak = inputs.rho * inputs.c0 * inputs.v  ' Joukowsky equation
        result.Fmax = ForceFromPressure_kN(result.Ppeak, inputs.Dint_mm / 1000#)
        result.LOF = SafeDiv(result.Fmax, result.Flim)
    End If
    
    CalculateLiquidClose = result
End Function

' Calculate gas opening rapid case
Private Function CalculateGasOpenRapid(inputs As ValveInputs, Flim As Double) As CalculationResult
    Dim result As CalculationResult
    result.Flim = Flim
    
    ' TODO: Implement EI gas-opening calculation
    result.Fmax = 0#
    result.LOF = SafeDiv(result.Fmax, result.Flim)
    result.FlagText = "Gas opening calculation not yet implemented"
    
    CalculateGasOpenRapid = result
End Function

' Calculate liquid opening case
Private Function CalculateLiquidOpen(inputs As ValveInputs, Flim As Double) As CalculationResult
    Dim result As CalculationResult
    result.Flim = Flim
    
    ' Calculate delta P (pressure difference)
    Dim deltaP As Double
    deltaP = Abs(inputs.P2 - inputs.P1)  ' Assuming P2 and P1 are upstream and downstream pressures
    
    ' Calculate peak force using T2.8 formula: Fmax = (1/1.58) * W * sqrt(deltaP/rho)
    ' Where W is the weight factor (assumed to be related to pipe characteristics)
    Dim W As Double
    ' For liquid opening, W could be related to pipe diameter and fluid properties
    ' This may need adjustment based on complete T2.8 specification
    W = WorksheetFunction.Pi() * (inputs.Dint_mm / 1000#) ^ 2 / 4# * inputs.rho  ' Basic weight estimation
    
    If inputs.rho > 0 And deltaP >= 0 Then
        result.Fmax = (1# / 1.58) * W * Sqr(deltaP / inputs.rho)
        result.Fmax = result.Fmax / 1000#  ' Convert to kN if needed
        result.LOF = SafeDiv(result.Fmax, result.Flim)
        result.FlagText = "Liquid opening calculated using T2.8 formula"
    Else
        result.Fmax = 0#
        result.LOF = 0#
        result.FlagText = "Invalid inputs for liquid opening calculation (rho <= 0 or deltaP < 0)"
    End If
    
    CalculateLiquidOpen = result
End Function

' Calculate force limit (Flim)
Private Function CalculateFlim(inputs As ValveInputs) As Double
    Dim Psi As Double
    If inputs.Tsch40 > 0 Then Psi = inputs.T_mm / inputs.Tsch40
    
    Dim theta As Double
    theta = ThetaFromSupport(inputs.supporttype)
    
    CalculateFlim = Flim_kN_EI(Psi, inputs.Dext_mm / 1000#, theta, inputs.Dint_mm / 1000#)
End Function

' Calculate wave speed if not provided
Public Function CalculateWaveSpeed(inputs As ValveInputs) As Double
    If inputs.c0 > 0 Then
        CalculateWaveSpeed = inputs.c0
        Exit Function
    End If
    
    If inputs.rho > 0 And inputs.Dext_mm > 0 And inputs.T_mm > 0 And _
       inputs.Em > 0 And inputs.Kbulk > 0 Then
        CalculateWaveSpeed = WaveSpeed_EI(inputs.rho, inputs.Dext_mm, _
                                        inputs.T_mm, inputs.Em, inputs.Kbulk)
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
    term2 = (Dext_mm / T_mm) / Em * (1000# / rho)
    Dim denom As Double: denom = term1 + term2
    If denom > 0 Then WaveSpeed_EI = 1# / denom
End Function

Private Function Flim_kN_EI(ByVal Psi As Double, ByVal Dext_m As Double, _
                            ByVal theta As Double, ByVal Dint_m As Double) As Double
    Dim poly As Double
    poly = 16.8 * Psi ^ 3 - 1.81 * Psi ^ 2 + 525# * Psi + 25.3
    Dim area As Double: area = WorksheetFunction.Pi() * Dint_m ^ 2 / 4#
    Flim_kN_EI = poly * Dext_m * theta * area / 1000000000#
End Function

Private Function ForceFromPressure_kN(ByVal P As Double, ByVal Dint_m As Double) As Double
    Dim area As Double: area = WorksheetFunction.Pi() * Dint_m ^ 2 / 4#
    ForceFromPressure_kN = P * area / 1000#
End Function

Private Function SafeDiv(ByVal num As Double, ByVal den As Double) As Double
    If den <= 0 Then SafeDiv = 0# Else SafeDiv = num / den
End Function
