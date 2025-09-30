Option Explicit

' ========= WORKSHEET EVENT HANDLER FOR VALVELIST SHEET =========
' This code should be placed in the ValveList worksheet's code module
' To use: Copy this code to the worksheet object in the VBA editor

Private Sub Worksheet_Change(ByVal Target As Range)
    ' Handle changes to the tbValveList table
    HandleTableChanges Target
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Optional: Could be used for real-time validation setup
End Sub

Private Sub HandleTableChanges(ByVal Target As Range)
    On Error GoTo ErrorHandler
    
    ' Check if the change occurred in the tbValveList table
    Dim lo As ListObject
    Dim i As Long
    
    ' Find if any ListObject contains the changed range
    For i = 1 To Me.ListObjects.Count
        Set lo = Me.ListObjects(i)
        If lo.Name = "tbValveList" Then
            ' Check if the target intersects with the table
            If Not Intersect(Target, lo.Range) Is Nothing Then
                ' Small delay to ensure row is fully created
                Application.EnableEvents = False
                Application.Wait (Now + TimeValue("0:00:01"))
                Application.EnableEvents = True
                
                ' Setup validation for the table
                Call TableEventHandler.OnTableRowAdded("ValveList", "tbValveList")
                Exit For
            End If
        End If
    Next i
    
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    ' Continue silently to avoid breaking normal operations
End Sub