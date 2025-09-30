Attribute VB_Name = "T28_UI_Calc"

Option Explicit

' === Public entry points ===
Sub Generate_Inputs_From_ValveList()
    Dim wsVL As Worksheet, wsIn As Worksheet, wsData As Worksheet
    Set wsVL = Sheets("ValveList")
    Set wsIn = Sheets("Inputs")
    Set wsData = Sheets("Data")
    
    ' Read tags
    Dim lastRow As Long: lastRow = wsVL.Cells(wsVL.Rows.Count, "A").End(xlUp).row
    Dim tags As Collection: Set tags = New Collection
    Dim ct As Collection: Set ct = New Collection
    Dim vt As Collection: Set vt = New Collection
    Dim st As Collection: Set st = New Collection
    
    Dim r As Long
    For r = 3 To lastRow
        If Trim$(wsVL.Cells(r, "A").value) <> "" Then
            tags.Add wsVL.Cells(r, "A").value
            ct.Add wsVL.Cells(r, "B").value
            vt.Add wsVL.Cells(r, "C").value
            st.Add wsVL.Cells(r, "D").value
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
        wsIn.Cells(2, c).value = tags(i)
        wsIn.Cells(2, c).Interior.Color = RGB(31, 78, 120)
        wsIn.Cells(2, c).Font.Color = vbWhite
        wsIn.Cells(2, c).Font.Bold = True
        wsIn.Cells(2, c).HorizontalAlignment = xlCenter
        
        ' Add dropdowns for "Pipe Support Type" and "Valve Type" rows
        Dim lastParamRow As Long
        lastParamRow = wsIn.Cells(wsIn.Rows.Count, "A").End(xlUp).row
        
        Dim rSupport As Long, rValveType As Long
        rSupport = FindRow(wsIn, "Pipe Support Type")
        rValveType = FindRow(wsIn, "Valve Type")
        
        If rSupport > 0 Then
            AddListValidation wsIn, c, rSupport, Sheets("Data").Range("B13:B16")
        End If
        If rValveType > 0 Then
            AddListValidation wsIn, c, rValveType, Sheets("Data").Range("B7:B11")
        End If
        
        c = c + 1
    Next i
    
    MsgBox "Inputs generated for " & tags.Count & " valves.", vbInformation
End Sub

Sub Calculate_All()
    Dim wsIn As Worksheet: Set wsIn = Sheets("Inputs")
    Dim wsVL As Worksheet: Set wsVL = Sheets("ValveList")
    Dim wsRes As Worksheet: Set wsRes = Sheets("Results")
    
    ' Map parameter row indices once
    Dim idx As Object: Set idx = CreateObject("Scripting.Dictionary")
    Dim lastParamRow As Long: lastParamRow = wsIn.Cells(wsIn.Rows.Count, "A").End(xlUp).row
    Dim r As Long
    For r = 3 To lastParamRow
        Dim key As String: key = Trim$(wsIn.Cells(r, "A").value)
        If key <> "" Then idx(key) = r
    Next r
    
    ' Find first valve column
    Dim firstCol As Long: firstCol = 5
    Dim lastCol As Long: lastCol = wsIn.Cells(2, wsIn.Columns.Count).End(xlToLeft).Column
    If lastCol < firstCol Then
        MsgBox "No valve columns found. Generate Inputs first.", vbExclamation
        Exit Sub
    End If
    
    ' Clear previous results (keep headers)
    Dim resLastRow As Long: resLastRow = wsRes.Cells(wsRes.Rows.Count, "A").End(xlUp).row
    If resLastRow >= 3 Then wsRes.Rows("3:" & resLastRow).ClearContents
    
    Dim c As Long, resRow As Long: resRow = 3
    For c = firstCol To lastCol
        Dim tag As String: tag = wsIn.Cells(2, c).value
        If Trim$(tag) = "" Then GoTo NextC
        
        ' Read from ValveList row for metadata
        Dim rowVL As Long: rowVL = FindRowInCol(wsVL, tag, 1)
        Dim casetype As String, valvetype As String, supporttype As String
        If rowVL > 0 Then
            casetype = wsVL.Cells(rowVL, "B").value
            valvetype = wsVL.Cells(rowVL, "C").value
            supporttype = wsVL.Cells(rowVL, "D").value
        End If
        
        ' Inputs
        Dim rho As Double: rho = NzD(wsIn.Cells(idx("Fluid density"), c).value)
        Dim gamma As Double: gamma = NzD(wsIn.Cells(idx("Ratio of Specific Heat Capacities"), c).value)
        Dim c0 As Double: c0 = NzD(wsIn.Cells(idx("Speed of sound"), c).value)
        Dim Dext_mm As Double: Dext_mm = NzD(wsIn.Cells(idx("External Main Line Diameter"), c).value)
        Dim Dint_mm As Double: Dint_mm = NzD(wsIn.Cells(idx("Internal Main Line Diameter"), c).value)
        Dim Em As Double: Em = NzD(wsIn.Cells(idx("Young’s Modulus of the main line material"), c).value)
        Dim Kbulk As Double: Kbulk = NzD(wsIn.Cells(idx("Fluid Bulk Modulus"), c).value)
        Dim Lup As Double: Lup = NzD(wsIn.Cells(idx("Upstream Pipe Length"), c).value)
        Dim Mw As Double: Mw = NzD(wsIn.Cells(idx("Molecular Weight"), c).value)
        Dim P1 As Double: P1 = NzD(wsIn.Cells(idx("Upstream Static Pressure"), c).value)
        Dim Pshut As Double: Pshut = NzD(wsIn.Cells(idx("Pump head at zero flow"), c).value)
        Dim Pv As Double: Pv = NzD(wsIn.Cells(idx("Vapour Pressure"), c).value)
        Dim dP As Double: dP = NzD(wsIn.Cells(idx("Static Pressure drop"), c).value)
        Dim r As Double: r = NzD(wsIn.Cells(idx("Universal Gas Constant"), c).value)
        Dim T_mm As Double: T_mm = NzD(wsIn.Cells(idx("Main line Wall Thickness"), c).value)
        Dim Tclose As Double: Tclose = NzD(wsIn.Cells(idx("Valve Closing Time"), c).value)
        Dim Te As Double: Te = NzD(wsIn.Cells(idx("Upstream Temperature"), c).value)
        Dim v As Double: v = NzD(wsIn.Cells(idx("Steady State Fluid Velocity"), c).value)
        Dim W As Double: W = NzD(wsIn.Cells(idx("Mass Flow Rate"), c).value)
        Dim Tsch40 As Double: Tsch40 = NzD(wsIn.Cells(idx("Main line Wall Thickness for Schedule 40 Piping"), c).value)
        
        ' Derived: Psi & theta
        Dim Psi As Double: If Tsch40 > 0 Then Psi = T_mm / Tsch40
        Dim theta As Double: theta = ThetaFromSupport(supporttype)
        
        ' Compute c if blank
        If c0 <= 0 And rho > 0 And Dext_mm > 0 And T_mm > 0 And Em > 0 And Kbulk > 0 Then
            c0 = WaveSpeed_EI(rho, Dext_mm, T_mm, Em, Kbulk)
        End If
        
        ' Compute branch
        Dim Ppeak As Double, Fmax As Double, Flim As Double, LOF As Double, flagTxt As String
        Flim = Flim_kN_EI(Psi, Dext_mm / 1000#, theta, Dint_mm / 1000#)
        
        Select Case LCase$(casetype)
            Case "liqclose"
                If Lup > 100# Then
                    LOF = 1#: flagTxt = "Lup>100 m → Detailed surge analysis"
                Else
                    Ppeak = rho * c0 * v   ' Joukowsky
                    Fmax = ForceFromPressure_kN(Ppeak, Dint_mm / 1000#)
                    LOF = SafeDiv(Fmax, Flim)
                End If
            Case "gasopenrapid"
                Fmax = 0#  ' TODO: insert EI gas opening relation
                LOF = SafeDiv(Fmax, Flim)
            Case "liqopen"
                Fmax = 0#  ' TODO: insert EI liquid opening relation
                LOF = SafeDiv(Fmax, Flim)
            Case Else
                flagTxt = "Unknown CaseType"
        End Select
        
        ' Write results row
        wsRes.Cells(resRow, 1).value = tag
        wsRes.Cells(resRow, 2).value = casetype
        wsRes.Cells(resRow, 3).value = valvetype
        wsRes.Cells(resRow, 4).value = supporttype
        wsRes.Cells(resRow, 5).value = Ppeak
        wsRes.Cells(resRow, 6).value = Fmax
        wsRes.Cells(resRow, 7).value = Flim
        wsRes.Cells(resRow, 8).value = LOF
        wsRes.Cells(resRow, 9).value = flagTxt
        
        resRow = resRow + 1
NextC:
    Next c
    
    MsgBox "Calculation complete.", vbInformation
End Sub

' === Helpers ===
Private Function FindRow(ws As Worksheet, ByVal key As String) As Long
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    Dim r As Long
    For r = 1 To lastRow
        If Trim$(ws.Cells(r, "A").value) = key Then
            FindRow = r: Exit Function
        End If
    Next r
    FindRow = 0
End Function

Private Function FindRowInCol(ws As Worksheet, ByVal key As String, ByVal col As Long) As Long
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).row
    Dim r As Long
    For r = 1 To lastRow
        If Trim$(ws.Cells(r, col).value) = key Then
            FindRowInCol = r: Exit Function
        End If
    Next r
    FindRowInCol = 0
End Function

Private Sub AddListValidation(ws As Worksheet, ByVal col As Long, ByVal row As Long, rng As Range)
    Dim dv As Validation
    On Error Resume Next
    ws.Cells(row, col).Validation.Delete
    On Error GoTo 0
    ws.Cells(row, col).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="=" & rng.Address(True, True, xlA1, True)
End Sub

Private Function ThetaFromSupport(ByVal support As String) As Double
    If InStr(1, support, "Anchor", vbTextCompare) > 0 Then ThetaFromSupport = 4#
    If InStr(1, support, "Guide", vbTextCompare) > 0 Then ThetaFromSupport = 2#
    If InStr(1, support, "Sliding", vbTextCompare) > 0 Then ThetaFromSupport = 1#
    If InStr(1, support, "None", vbTextCompare) > 0 Then ThetaFromSupport = 0.5
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

Private Function NzD(v) As Double
    If IsError(v) Then Exit Function
    If Len(Trim$(v & "")) = 0 Then Exit Function
    NzD = CDbl(v)
End Function
