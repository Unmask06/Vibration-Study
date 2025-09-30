VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmAbout 
   Caption         =   "About VBA Import/Export Starter"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "FrmAbout.frx:0000"
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' FrmAbout - Example UserForm for VBA Import/Export Starter
' Demonstrates a simple about dialog

Private Sub UserForm_Initialize()
    ' Initialize the form when it loads
    Me.lblTitle.Caption = "VBA Import/Export Starter"
    Me.lblVersion.Caption = "Version 1.0"
    Me.lblDescription.Caption = "A utility for seamless VBA development between VS Code and Excel VBE."
    Me.lblCopyright.Caption = "Created with GitHub Copilot"
End Sub

Private Sub cmdOK_Click()
    ' Close the form when OK is clicked
    Unload Me
End Sub

Private Sub cmdShowLog_Click()
    ' Demonstrate the CLogger class
    Dim logger As CLogger
    Set logger = New CLogger
    
    logger.Prefix = "FrmAbout"
    logger.Info "About dialog opened"
    logger.Info "Demonstrating CLogger functionality"
    
    MsgBox "Check the Immediate Window (Ctrl+G) for log messages.", vbInformation, "Logger Demo"
End Sub