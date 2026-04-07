Attribute VB_Name = "modShutdown"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" ( _
        ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
#Else
    Private Declare Function GetForegroundWindow Lib "user32" () As Long
    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" ( _
        ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
#End If

Private mIdlePaused As Boolean   'internal pause state

' === CONFIGURATION ===
Public Const TimeOutMinutes As Double = 5          ' idle time before auto-close
Public Const WarnBeforeMinutes As Double = 1       ' warning period before close
Public Const IdleCheckIntervalSeconds As Long = 30 ' how often to check (seconds)

' === STATE VARIABLES ===
Public IdleCheckTime As Date
Public IsMacroBusy As Boolean
Public WarnDeadline As Date
Public WarningShown As Boolean
Public DevMode As Boolean


Public Sub ToggleDevMode()
    DevMode = Not DevMode
    If DevMode Then
        PauseIdleCheck True
        Application.StatusBar = "DEV MODE: Idle auto-close PAUSED"
    Else
        Application.StatusBar = False
        PauseIdleCheck False
    End If
End Sub

'==========================================================
' Entry point – called from Workbook_Open
'==========================================================
Public Sub StartIdleCheck()
    ScheduleNextIdleCheck
End Sub

Private Sub ScheduleNextIdleCheck()
    IdleCheckTime = Now + TimeSerial(0, 0, IdleCheckIntervalSeconds)
    Application.OnTime EarliestTime:=IdleCheckTime, Procedure:="IdleCheck"
End Sub


'==========================================================
' Periodic idle check – non-blocking (no MsgBox)
'==========================================================
Public Sub IdleCheck()
    On Error GoTo ExitPoint

    If DevMode Then GoTo ExitPoint

    Dim idleMinutes As Double
    idleMinutes = (Now - ThisWorkbook.LastActivity) * 1440#   ' days ? minutes

    ' Dev-friendly: pause while VBE is the active window
    If IsVBEActiveWindow() Then
        PauseIdleCheck True
        GoTo ExitPoint
    ElseIf mIdlePaused Then
        ' If we were paused but VBE is no longer active, resume
        PauseIdleCheck False
    End If


    ' 1) Don't close while you've flagged a long macro,
    '    or if events are disabled (common during heavy VBA work)
    If IsMacroBusy Or Application.EnableEvents = False Then
        ScheduleNextIdleCheck
        GoTo ExitPoint
    End If

    ' 2) Hard timeout – close regardless of warning state
    If idleMinutes >= TimeOutMinutes Then
        AutoCloseWorkbook
        GoTo ExitPoint
    End If



    ' 3) Warning window / countdown
    If idleMinutes >= TimeOutMinutes - WarnBeforeMinutes Then

        ' Show warning form only once
        If Not WarningShown Then
            WarnDeadline = Now + WarnBeforeMinutes / 1440#

            On Error Resume Next
            With frmIdleWarning
                .lblText.Caption = _
                    "This workbook will auto-save and close in " & _
                    Format$(WarnBeforeMinutes * 60 - (idleMinutes - (TimeOutMinutes - WarnBeforeMinutes)) * 60, "0") & _
                    " seconds due to inactivity." & vbCrLf & vbCrLf & _
                    "Click 'Stay Open' to continue working, or 'Close Now' to exit immediately."
                .Show vbModeless
            End With
            On Error GoTo ExitPoint

            WarningShown = True
        End If

        ' If the deadline has passed, close even if the form is still sitting there
        If Now >= WarnDeadline Then
            AutoCloseWorkbook
            GoTo ExitPoint
        End If
    End If

    ' 4) Not yet timed out – keep checking
    ScheduleNextIdleCheck

ExitPoint:
    Exit Sub
End Sub


'==========================================================
' Auto-save, log, and close
'==========================================================
Public Sub AutoCloseWorkbook()
    On Error Resume Next

    ' Cancel any pending timer first
    Application.OnTime EarliestTime:=IdleCheckTime, Procedure:="IdleCheck", Schedule:=False

    ' Tidy up warning form
    If WarningShown Then
        Unload frmIdleWarning
        WarningShown = False
    End If

    ' Log who left it open
    LogIdleClose

    ' Save & close
    ThisWorkbook.Save
    ThisWorkbook.Close SaveChanges:=True
End Sub


'==========================================================
' Log sheet: who left it open & when
'==========================================================
Public Sub LogIdleClose()
    On Error GoTo ExitPoint

    Const LOG_SHEET As String = "IdleLog"
    Dim ws As Worksheet
    Dim NextRow As Long

    ' Get or create log sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.name = LOG_SHEET
        ws.Range("A1:D1").Value = Array("Timestamp", "Windows User", "Excel UserName", "Workbook")
    End If

    NextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1

    ws.Cells(NextRow, "A").Value = Now
    ws.Cells(NextRow, "B").Value = Environ$("USERNAME")
    ws.Cells(NextRow, "C").Value = Application.UserName
    ws.Cells(NextRow, "D").Value = ThisWorkbook.FullName

ExitPoint:
    On Error Resume Next
    If Not ws Is Nothing Then ws.Visible = xlSheetVeryHidden
End Sub


'==========================================================
' OPTIONAL: helpers to wrap long-running macros
'==========================================================
Public Sub BeginLongMacro()
    IsMacroBusy = True
End Sub

Public Sub EndLongMacro()
    IsMacroBusy = False
End Sub



Private Function IsVBEActiveWindow() As Boolean
    On Error GoTo SafeExit

    Dim hWndFg As LongPtr
    Dim className As String * 256
    Dim n As Long

    hWndFg = GetForegroundWindow()
    If hWndFg = 0 Then GoTo SafeExit

    n = GetClassName(hWndFg, className, 255)
    If n <= 0 Then GoTo SafeExit

    ' VBE main window class is typically "wndclass_desked_gsk"
    IsVBEActiveWindow = (InStr(1, Left$(className, n), "wndclass_desked_gsk", vbTextCompare) > 0)

SafeExit:
End Function

Public Sub PauseIdleCheck(Optional ByVal paused As Boolean = True)
    On Error Resume Next
    mIdlePaused = paused

    'If pausing, cancel any pending tick
    If mIdlePaused Then
        Application.OnTime EarliestTime:=IdleCheckTime, Procedure:="IdleCheck", Schedule:=False
    Else
        'resume cleanly
        WarningShown = False
        ScheduleNextIdleCheck
    End If
End Sub

