Attribute VB_Name = "ModMain"
Option Explicit

' ModMain - Example standard module for VBA Import/Export Starter
' This is a demonstration module to show the workflow

Public Sub HelloWorld()
    ' Simple demonstration procedure
    Debug.Print "Hello from ModMain module!"
    Debug.Print "Current time: " & Now()
    Debug.Print "VBA Import/Export workflow is working correctly."
End Sub

Public Sub TestCalculation()
    ' Example calculation procedure
    Dim result As Double
    Dim i As Long
    
    result = 0
    For i = 1 To 10
        result = result + i
    Next i
    
    Debug.Print "Sum of 1 to 10: " & result
    MsgBox "Calculation complete. Sum of 1 to 10 = " & result, vbInformation, "ModMain Test"
End Sub

Public Function GetProjectInfo() As String
    ' Returns basic project information
    GetProjectInfo = "VBA Import/Export Starter - " & _
                    "Workbook: " & ThisWorkbook.Name & " - " & _
                    "Date: " & Format(Date, "yyyy-mm-dd")
End Function

' ========= PRESSURE UNIT CONVERSION =========
Public Sub ConvertPressureUnitsToBarg()
    ' Convert pressure parameter units from Pa to barg in the Inputs worksheet
    ' This updates column C units and prepares the system to accept barg input values
    Dim wsIn As Worksheet
    
    On Error GoTo ErrorHandler
    Set wsIn = Sheets("Inputs")
    
    ' Update the units in the worksheet
    Call ParameterManager.UpdatePressureUnitsToBarG(wsIn)
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error converting pressure units: " & Err.Description, vbCritical, "Conversion Error"
End Sub
