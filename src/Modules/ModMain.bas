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
