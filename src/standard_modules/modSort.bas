Attribute VB_Name = "modSort"


Public Sub ToggleSortFromButton(tblName As String, colname As String)

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim btn As Button
    Dim btnName As String
    Dim btnText As String
    Dim asc As Boolean
    Dim shGuard As TSheetGuardState
    
    On Error GoTo ErrHandler
    
    AppGuard_Begin
    
    'Resolve worksheet containing the table
    Set ws = GetWorksheetOfTable(ThisWorkbook, tblName)
    
    If ws Is Nothing Then
        Err.Raise vbObjectError + 5000, "ToggleSortFromButton", _
                  "Table '" & tblName & "' not found."
    End If
    
    'Resolve table
    Set tbl = ws.ListObjects(tblName)
    
    'Ensure macro was called from a button
    If TypeName(Application.Caller) <> "String" Then
        MsgBox "This macro must be run from a button.", vbExclamation
        GoTo CleanExit
    End If
    
    btnName = Application.Caller
    Set btn = ws.Buttons(btnName)
    
    'Temporarily release sheet protection
    shGuard = SheetGuard_Begin(ws)
    
    btnText = btn.Caption
    
    If btnText = "Sort Asc" Then
        asc = True
        btn.Caption = "Sort Desc"
    Else
        asc = False
        btn.Caption = "Sort Asc"
    End If
    
    SafeSortTable tbl, colname, asc, password

CleanExit:

    SheetGuard_End ws, shGuard, pwd
    AppGuard_End
    Exit Sub

ErrHandler:

    MsgBox Err.Description, vbExclamation
    Resume CleanExit

End Sub
