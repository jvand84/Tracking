VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmIdleWarning 
   Caption         =   "Idle Warning"
   ClientHeight    =   1905
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4410
   OleObjectBlob   =   "frmIdleWarning.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmIdleWarning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStayOpen_Click()
    ' User is still there – reset activity and close the warning
    ThisWorkbook.LastActivity = Now
    WarningShown = False
    Unload Me
End Sub

Private Sub cmdCloseNow_Click()
    ' User explicitly wants to close now
    WarningShown = False
    Unload Me
    AutoCloseWorkbook
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' If they hit the X, treat it like "Stay Open"
    If CloseMode = vbFormControlMenu Then
        ThisWorkbook.LastActivity = Now
        WarningShown = False
    End If
End Sub
