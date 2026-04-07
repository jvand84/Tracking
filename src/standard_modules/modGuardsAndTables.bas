Attribute VB_Name = "modGuardsAndTables"
Option Explicit

'============================================================
' modGuardsAndTables
'
' Guard + Table Helper Module Template (paste-ready)
'============================================================

'--------------------------------
' Sheet protection guard (optional)
'--------------------------------
Public Type TSheetGuardState
    wasProtected As Boolean
    sheetName As String
End Type

'-----------------------------
' Application Guard (AppGuard)
'-----------------------------
Public Type TAppGuardState
    Calc As XlCalculation
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    DisplayStatusBar As Boolean
    StatusBarText As Variant
    Cursor As XlMousePointer
End Type

Private mAppSaved As TAppGuardState
Private mAppHasSaved As Boolean
Private mAppGuardDepth As Long

' Begin an application guard.
Public Sub AppGuard_Begin(Optional ByVal showStatus As Boolean = False, _
                         Optional ByVal statusText As String = vbNullString, _
                         Optional ByVal setCalcManual As Boolean = True)

    On Error GoTo ErrorHandle

    mAppGuardDepth = mAppGuardDepth + 1

    ' Only capture state on the outermost guard.
    If mAppGuardDepth = 1 Then
        mAppSaved.Calc = Application.Calculation
        mAppSaved.ScreenUpdating = Application.ScreenUpdating
        mAppSaved.EnableEvents = Application.EnableEvents
        mAppSaved.DisplayStatusBar = Application.DisplayStatusBar
        mAppSaved.StatusBarText = Application.StatusBar
        mAppSaved.Cursor = Application.Cursor

        mAppHasSaved = True

        ' Apply "safe performance mode"
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayStatusBar = True
        Application.Cursor = xlWait

        If setCalcManual Then
            Application.Calculation = xlCalculationManual
        End If
    End If

    If showStatus Then
        If Len(statusText) > 0 Then Application.StatusBar = statusText
    End If

    Exit Sub

ErrorHandle:
    ' If AppGuard_Begin fails, do not leave Excel broken.
    On Error Resume Next
    mAppGuardDepth = mAppGuardDepth - 1
    If mAppGuardDepth <= 0 Then
        mAppGuardDepth = 0
        If mAppHasSaved Then AppGuard_End
    End If
End Sub

' End an application guard.
Public Sub AppGuard_End(Optional ByVal clearStatus As Boolean = True)
    On Error GoTo ErrorHandle

    If mAppGuardDepth <= 0 Then Exit Sub

    mAppGuardDepth = mAppGuardDepth - 1

    ' Only restore on the outermost end.
    If mAppGuardDepth = 0 Then
        If mAppHasSaved Then
            Application.Calculation = mAppSaved.Calc
            Application.ScreenUpdating = mAppSaved.ScreenUpdating
            Application.EnableEvents = mAppSaved.EnableEvents
            Application.DisplayStatusBar = mAppSaved.DisplayStatusBar
            Application.Cursor = mAppSaved.Cursor

            If clearStatus Then
                Application.StatusBar = False
            Else
                Application.StatusBar = mAppSaved.StatusBarText
            End If
        End If

        mAppHasSaved = False
    End If

    Exit Sub

ErrorHandle:
    ' Last-ditch safety restore attempt
    On Error Resume Next
    mAppGuardDepth = 0
    If mAppHasSaved Then
        Application.Calculation = mAppSaved.Calc
        Application.ScreenUpdating = mAppSaved.ScreenUpdating
        Application.EnableEvents = mAppSaved.EnableEvents
        Application.DisplayStatusBar = mAppSaved.DisplayStatusBar
        Application.Cursor = mAppSaved.Cursor
        Application.StatusBar = False
    End If
    mAppHasSaved = False
End Sub

' Unprotects a sheet only if it was protected.
Public Function SheetGuard_Begin(ByVal ws As Worksheet, Optional ByVal password As String = vbNullString) As TSheetGuardState

    Dim st As TSheetGuardState
    On Error GoTo ErrorHandle

    st.sheetName = ws.name
    st.wasProtected = ws.ProtectContents

    If st.wasProtected Then
        ws.Unprotect password:=pwd
    End If

    SheetGuard_Begin = st
    Exit Function

ErrorHandle:
    Err.Raise Err.Number, "SheetGuard_Begin(" & ws.name & ")", Err.Description
End Function



' Always re-protect the sheet at the end.
Public Sub SheetGuard_End(ByVal ws As Worksheet, _
                          ByRef st As TSheetGuardState, _
                          Optional ByVal password As String = vbNullString)

    On Error GoTo ErrorHandle

    If ws Is Nothing Then Exit Sub

    ' Optional safety check to ensure we are ending guard
    ' on the same sheet we started on.
    If Len(st.sheetName) > 0 Then
        If ws.name <> st.sheetName Then Exit Sub
    End If

    ' Always protect, even if already protected.
    ws.Protect password:=password, _
               UserInterfaceOnly:=True, _
               AllowFiltering:=True, _
               AllowSorting:=True, _
               AllowUsingPivotTables:=True, _
               AllowFormattingCells:=False, _
               AllowFormattingColumns:=True, _
               AllowFormattingRows:=False, _
               AllowInsertingRows:=True, _
               AllowDeletingRows:=False, _
               AllowInsertingColumns:=False, _
               AllowDeletingColumns:=False

    Exit Sub

ErrorHandle:
    Err.Raise Err.Number, "SheetGuard_End(" & ws.name & ")", Err.Description
End Sub

'-----------------------------
' Table / ListObject Helpers
'-----------------------------

Public Function GetWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error GoTo ErrorHandle
    Set GetWorksheet = wb.Worksheets(sheetName)
    Exit Function
ErrorHandle:
    Err.Raise 5, "GetWorksheet", "Worksheet not found: '" & sheetName & "' in workbook '" & wb.name & "'"
End Function

Public Function GetTable(ByVal ws As Worksheet, ByVal tableName As String) As ListObject
    On Error GoTo ErrorHandle
    Set GetTable = ws.ListObjects(tableName)
    Exit Function
ErrorHandle:
    Err.Raise 5, "GetTable", "Table not found: '" & tableName & "' on sheet '" & ws.name & "'"
End Function

Public Function GetWorksheetOfTable(ByVal wb As Workbook, ByVal tableName As String) As Worksheet
    Dim ws As Worksheet
    On Error GoTo ErrorHandle

    For Each ws In wb.Worksheets
        If HasTable(ws, tableName) Then
            Set GetWorksheetOfTable = ws
            Exit Function
        End If
    Next ws

    Err.Raise 5, "GetWorksheetOfTable", "Table '" & tableName & "' not found in workbook '" & wb.name & "'"
    Exit Function

ErrorHandle:
    Err.Raise Err.Number, "GetWorksheetOfTable", Err.Description
End Function

Public Function HasTable(ByVal ws As Worksheet, ByVal tableName As String) As Boolean
    On Error Resume Next
    Dim lo As ListObject
    Set lo = ws.ListObjects(tableName)
    HasTable = Not (lo Is Nothing)
End Function

Public Function GetTableColIndex(ByVal lo As ListObject, ByVal headerName As String) As Long
    Dim i As Long, h As String, Target As String
    Target = NormalizeHeader(headerName)

    For i = 1 To lo.ListColumns.Count
        h = NormalizeHeader(lo.ListColumns(i).name)
        If h = Target Then
            GetTableColIndex = i
            Exit Function
        End If
    Next i

    Err.Raise 5, "GetTableColIndex", "Column not found: '" & headerName & "' in table '" & lo.name & "'"
End Function

Public Function GetTableDataColRange(ByVal lo As ListObject, ByVal headerName As String) As Range
    Dim idx As Long
    idx = GetTableColIndex(lo, headerName)

    If lo.DataBodyRange Is Nothing Then Exit Function
    Set GetTableDataColRange = lo.ListColumns(idx).DataBodyRange
End Function

' Clear a table to header-only (keeps formatting and total row).
' Deterministic: deletes all ListRows safely.
Public Sub ClearTableToHeaderOnly(ByVal lo As ListObject)
    On Error GoTo ErrorHandle

    If lo.ListRows.Count = 0 Then Exit Sub

    Do While lo.ListRows.Count > 0
        lo.ListRows(lo.ListRows.Count).Delete
    Loop

    Exit Sub

ErrorHandle:
    Err.Raise Err.Number, "ClearTableToHeaderOnly(" & lo.name & ")", Err.Description
End Sub

Public Sub ClearTableRowsToHeaderOnly(ByVal lo As ListObject)
    On Error GoTo ErrorHandle

    Do While lo.ListRows.Count > 0
        lo.ListRows(lo.ListRows.Count).Delete
    Loop

    Exit Sub

ErrorHandle:
    Err.Raise Err.Number, "ClearTableRowsToHeaderOnly(" & lo.name & ")", Err.Description
End Sub

Public Sub ResizeListObjectRowsExact(ByVal lo As ListObject, ByVal nRows As Long)
    Dim cur As Long
    On Error GoTo ErrorHandle

    If nRows < 0 Then Err.Raise 5, "ResizeListObjectRowsExact", "nRows cannot be negative."

    cur = lo.ListRows.Count

    Do While cur < nRows
        lo.ListRows.Add
        cur = cur + 1
    Loop

    Do While cur > nRows
        lo.ListRows(cur).Delete
        cur = cur - 1
    Loop

    Exit Sub

ErrorHandle:
    Err.Raise Err.Number, "ResizeListObjectRowsExact(" & lo.name & ")", Err.Description
End Sub

Public Function IsCalculatedColumnFast(ByVal lo As ListObject, ByVal colIndex As Long) As Boolean
    On Error GoTo ErrorHandle

    If lo.ListRows.Count = 0 Then
        IsCalculatedColumnFast = False
        Exit Function
    End If

    Dim r As Range
    Set r = lo.ListColumns(colIndex).DataBodyRange.Cells(1, 1)

    IsCalculatedColumnFast = (Len(r.Formula) > 0 And Left$(r.Formula, 1) = "=")
    Exit Function

ErrorHandle:
    IsCalculatedColumnFast = False
End Function

Public Function NormalizeHeader(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    NormalizeHeader = LCase$(t)
End Function

Public Function TableToArray(ByVal lo As ListObject) As Variant
    On Error GoTo ErrorHandle

    If lo.DataBodyRange Is Nothing Then
        TableToArray = Empty
        Exit Function
    End If

    TableToArray = lo.DataBodyRange.Value2
    Exit Function

ErrorHandle:
    Err.Raise Err.Number, "TableToArray(" & lo.name & ")", Err.Description
End Function

Public Sub ArrayToTable(ByVal lo As ListObject, ByVal data As Variant, _
                        Optional ByVal skipCalculatedColumns As Boolean = True)

    Dim r As Long, c As Long, nR As Long, nC As Long
    Dim writeArr As Variant
    Dim colIsCalc() As Boolean

    On Error GoTo ErrorHandle

    If IsEmpty(data) Then
        ClearTableRowsToHeaderOnly lo
        Exit Sub
    End If

    nR = UBound(data, 1)
    nC = UBound(data, 2)

    ResizeListObjectRowsExact lo, nR

    ReDim colIsCalc(1 To lo.ListColumns.Count)
    If skipCalculatedColumns Then
        For c = 1 To lo.ListColumns.Count
            colIsCalc(c) = IsCalculatedColumnFast(lo, c)
        Next c
    End If

    ReDim writeArr(1 To nR, 1 To lo.ListColumns.Count)

    For r = 1 To nR
        For c = 1 To lo.ListColumns.Count
            If c <= nC Then
                If skipCalculatedColumns And colIsCalc(c) Then
                    ' Leave formula column untouched
                Else
                    writeArr(r, c) = data(r, c)
                End If
            End If
        Next c
    Next r

    lo.DataBodyRange.Value2 = writeArr
    Exit Sub

ErrorHandle:
    Err.Raise Err.Number, "ArrayToTable(" & lo.name & ")", Err.Description
End Sub

'-------------------------------------------------------
' FindListObjectByName
'
' Purpose:
'   Searches all worksheets in the supplied workbook for
'   a ListObject with the specified name.
'
' Inputs:
'   - wb: Target workbook
'   - tableName: ListObject name to find
'
' Outputs:
'   - Returns matching ListObject if found
'   - Returns Nothing if not found
'
' Notes:
'   - NOT finding a table is NOT an error condition
'   - Only raises error for invalid inputs or runtime issues
'-------------------------------------------------------
Public Function FindListObjectByName(ByVal wb As Workbook, _
                                     ByVal tableName As String) As ListObject

    '-------------------------------------------------------
    ' Variable Declarations
    '-------------------------------------------------------
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim searchName As String

    On Error GoTo ErrHandler

    '-------------------------------------------------------
    ' Validation
    '-------------------------------------------------------
    If wb Is Nothing Then
        Err.Raise vbObjectError + 3030, _
                  "FindListObjectByName", _
                  "Workbook reference is Nothing."
    End If

    searchName = Trim$(tableName)
    If Len(searchName) = 0 Then Exit Function

    '-------------------------------------------------------
    ' Search workbook tables
    '-------------------------------------------------------
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.name, searchName, vbTextCompare) = 0 Then
                Set FindListObjectByName = lo
                Exit Function
            End If
        Next lo
    Next ws

    '-------------------------------------------------------
    ' Not found ? return Nothing (expected outcome)
    '-------------------------------------------------------
    Exit Function

ErrHandler:
    ' Only real errors reach here
    Err.Raise Err.Number, "FindListObjectByName", Err.Description

End Function


