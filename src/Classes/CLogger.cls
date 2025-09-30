VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' CLogger - Example class module for VBA Import/Export Starter
' Demonstrates a simple logging utility class

Private pPrefix As String
Private pTimestamp As Boolean

' ============================================================================
' PROPERTIES
' ============================================================================

Public Property Let Prefix(ByVal newPrefix As String)
    ' Sets the prefix for log messages
    pPrefix = newPrefix
End Property

Public Property Get Prefix() As String
    ' Gets the current prefix for log messages
    Prefix = pPrefix
End Property

Public Property Let IncludeTimestamp(ByVal includeTime As Boolean)
    ' Sets whether to include timestamps in log messages
    pTimestamp = includeTime
End Property

Public Property Get IncludeTimestamp() As Boolean
    ' Gets whether timestamps are included in log messages
    IncludeTimestamp = pTimestamp
End Property

' ============================================================================
' PUBLIC METHODS
' ============================================================================

Public Sub Info(ByVal msg As String)
    ' Logs an information message
    Debug.Print BuildLogMessage("INFO", msg)
End Sub

Public Sub Warning(ByVal msg As String)
    ' Logs a warning message
    Debug.Print BuildLogMessage("WARN", msg)
End Sub

Public Sub ErrorMsg(ByVal msg As String)
    ' Logs an error message
    Debug.Print BuildLogMessage("ERROR", msg)
End Sub

Public Sub Debug(ByVal msg As String)
    ' Logs a debug message
    Debug.Print BuildLogMessage("DEBUG", msg)
End Sub

' ============================================================================
' PRIVATE METHODS
' ============================================================================

Private Function BuildLogMessage(ByVal level As String, ByVal msg As String) As String
    ' Constructs the formatted log message
    
    Dim logMsg As String
    
    ' Add timestamp if enabled
    If pTimestamp Then
        logMsg = Format(Now, "yyyy-mm-dd hh:mm:ss") & " "
    End If
    
    ' Add level
    logMsg = logMsg & "[" & level & "] "
    
    ' Add prefix if set
    If Len(pPrefix) > 0 Then
        logMsg = logMsg & pPrefix & ": "
    End If
    
    ' Add the actual message
    logMsg = logMsg & msg
    
    BuildLogMessage = logMsg
End Function

' ============================================================================
' CLASS EVENTS
' ============================================================================

Private Sub Class_Initialize()
    ' Default initialization
    pPrefix = ""
    pTimestamp = True
End Sub
