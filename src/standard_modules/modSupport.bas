Attribute VB_Name = "modSupport"
Option Explicit

'================================================================================
' modSupport
'
' Purpose:
'   Single “toolbox” module used across Tracking Schedule / Cost Track Migration /
'   Manning / Forms workflows.
'
' Design Principles:
'   - Safe getters (return Nothing / 0 / Empty instead of throwing)
'   - Header-agnostic table access (CleanKey-normalised matching)
'   - Centralised App state guard (performance + reliability)
'   - Centralised error logging (never throws from logger)
'
' Notes:
'   - Keep this as ONE module (as requested), but segregate into regions.
'   - Avoid global On Error Resume Next; only use it on known Excel tantrum lines.
'================================================================================

'================================================================================
' SECTION 2 — TEXT / KEY NORMALISATION (FOUNDATION)
'================================================================================

'------------------------------------------------------------
' NzStr
' Safe string conversion: Error/Empty/Null -> ""
'------------------------------------------------------------
Public Function NzStr(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
        NzStr = vbNullString
    Else
        NzStr = CStr(v)
    End If
End Function

'------------------------------------------------------------
' NormalizeLineBreaks
' Standardises CR/LF variants to vbLf
'------------------------------------------------------------
Public Function NormalizeLineBreaks(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    NormalizeLineBreaks = s
End Function

'------------------------------------------------------------
' CleanKey
' Normalises keys for reliable comparisons and lookups.
' Handles NBSP, tabs, CR/LF, double spaces, leading/trailing whitespace.
'------------------------------------------------------------
Public Function CleanKey(ByVal v As Variant) As String
    Dim s As String

    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
        CleanKey = vbNullString
        Exit Function
    End If

    s = CStr(v)

    s = Replace(s, Chr$(160), " ")  ' NBSP -> space
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")

    s = Trim$(s)
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    CleanKey = s
End Function

'------------------------------------------------------------
' NormalizeHeader
' Semantic alias for CleanKey (kept because you reference it elsewhere)
'------------------------------------------------------------
Public Function NormalizeHeader(ByVal v As Variant) As String
    NormalizeHeader = CleanKey(v)
End Function


'================================================================================
' SECTION 3 — WORKSHEET / TABLE SAFE GETTERS
'================================================================================

'------------------------------------------------------------
' GetWorksheetSafe
' Returns a worksheet reference or Nothing if not found.
' Case-insensitive. No direct index lookup.
'------------------------------------------------------------
Public Function GetWorksheetSafe(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    If Len(sheetName) = 0 Then Exit Function

    For Each ws In wb.Worksheets
        If StrComp(ws.name, sheetName, vbTextCompare) = 0 Then
            Set GetWorksheetSafe = ws
            Exit Function
        End If
    Next ws
End Function

'------------------------------------------------------------
' GetListObjectSafe
' Returns a ListObject from a given sheet or Nothing if not found.
'------------------------------------------------------------
Public Function GetListObjectSafe(ByVal ws As Worksheet, ByVal tableName As String) As ListObject
    Dim lo As ListObject

    If ws Is Nothing Then Exit Function
    If Len(tableName) = 0 Then Exit Function

    For Each lo In ws.ListObjects
        If StrComp(lo.name, tableName, vbTextCompare) = 0 Then
            Set GetListObjectSafe = lo
            Exit Function
        End If
    Next lo
End Function

'------------------------------------------------------------
' GetTableByName
' Finds a ListObject by name anywhere in a workbook (default ThisWorkbook).
' Returns Nothing if not found.
'------------------------------------------------------------
Public Function GetTableByName(ByVal tableName As String, Optional ByVal wb As Workbook) As ListObject
    Dim ws As Worksheet, lo As ListObject

    On Error GoTo SafeExit
    If wb Is Nothing Then Set wb = ThisWorkbook
    If Len(tableName) = 0 Then Exit Function

    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.name, tableName, vbTextCompare) = 0 Then
                Set GetTableByName = lo
                Exit Function
            End If
        Next lo
    Next ws

SafeExit:
End Function




'================================================================================
' SECTION 4 — TABLE HEADER / COLUMN HELPERS
'================================================================================

'------------------------------------------------------------
' TableColIndexByHeaderClean
' Returns column index by header using CleanKey matching. 0 if not found.
'------------------------------------------------------------
Public Function TableColIndexByHeaderClean(ByVal lo As ListObject, ByVal headerName As String) As Long
    Dim c As Long, needle As String, hay As String
    If lo Is Nothing Then Exit Function

    needle = CleanKey(headerName)
    For c = 1 To lo.ListColumns.Count
        hay = CleanKey(lo.ListColumns(c).name)
        If StrComp(hay, needle, vbTextCompare) = 0 Then
            TableColIndexByHeaderClean = c
            Exit Function
        End If
    Next c
End Function

'------------------------------------------------------------
' HeaderIndex
' Alias/variant of TableColIndexByHeaderClean (kept for readability).
'------------------------------------------------------------
Public Function HeaderIndex(ByVal lo As ListObject, ByVal headerName As String) As Long
    HeaderIndex = TableColIndexByHeaderClean(lo, headerName)
End Function

'------------------------------------------------------------
' TableHasData
' Reliable check for any data rows.
'------------------------------------------------------------
Public Function TableHasData(ByVal lo As ListObject) As Boolean
    On Error GoTo SafeExit
    If lo Is Nothing Then Exit Function
    TableHasData = (lo.ListRows.Count > 0)
    Exit Function
SafeExit:
    TableHasData = False
End Function

'------------------------------------------------------------
' GetTableDataRowIndexFromTarget
' Returns 1-based DataBody row index if Target is inside lo.DataBodyRange, else 0.
'------------------------------------------------------------
Public Function GetTableDataRowIndexFromTarget(ByVal lo As ListObject, ByVal Target As Range) As Long
    On Error GoTo SafeExit
    If lo Is Nothing Then GoTo SafeExit
    If lo.DataBodyRange Is Nothing Then GoTo SafeExit
    If Target Is Nothing Then GoTo SafeExit
    If Intersect(Target, lo.DataBodyRange) Is Nothing Then GoTo SafeExit

    GetTableDataRowIndexFromTarget = Target.Row - lo.DataBodyRange.Row + 1
    Exit Function

SafeExit:
    GetTableDataRowIndexFromTarget = 0
End Function

'------------------------------------------------------------
' TblValueSafe
' Returns a cell value from table by header + 1-based data row.
' Returns defaultValue (or Empty) if invalid/missing.
'------------------------------------------------------------
Public Function TblValueSafe(ByVal lo As ListObject, ByVal dataRow As Long, ByVal headerName As String, _
                            Optional ByVal defaultValue As Variant) As Variant
    Dim lc As ListColumn

    On Error GoTo Nope
    If lo Is Nothing Then GoTo Nope
    If lo.ListRows.Count = 0 Then GoTo Nope
    If dataRow < 1 Or dataRow > lo.ListRows.Count Then GoTo Nope

    Set lc = lo.ListColumns(headerName)
    TblValueSafe = lc.DataBodyRange.Cells(dataRow, 1).Value
    Exit Function

Nope:
    If IsMissing(defaultValue) Then
        TblValueSafe = Empty
    Else
        TblValueSafe = defaultValue
    End If
End Function

'------------------------------------------------------------
' FindTableRowIndexByValue
' Finds 1..n row index where column header equals findValue (string compare).
' Returns 0 if not found.
'------------------------------------------------------------
Public Function FindTableRowIndexByValue( _
    ByVal lo As ListObject, _
    ByVal headerName As String, _
    ByVal findValue As Variant, _
    Optional ByVal matchCase As Boolean = False _
) As Long

    Dim rng As Range, v As Variant, i As Long
    Dim needle As String, hay As String

    On Error GoTo SafeExit
    If lo Is Nothing Then GoTo SafeExit
    If lo.DataBodyRange Is Nothing Then GoTo SafeExit

    Set rng = lo.ListColumns(headerName).DataBodyRange
    If rng Is Nothing Then GoTo SafeExit

    needle = CStr(findValue)

    For i = 1 To rng.Rows.Count
        v = rng.Cells(i, 1).Value
        hay = CStr(v)

        If matchCase Then
            If hay = needle Then
                FindTableRowIndexByValue = i
                Exit Function
            End If
        Else
            If StrComp(hay, needle, vbTextCompare) = 0 Then
                FindTableRowIndexByValue = i
                Exit Function
            End If
        End If
    Next i

SafeExit:
    FindTableRowIndexByValue = 0
End Function

'------------------------------------------------------------
' FindTableRowIndexByValue_Clean
' Array-based search using CleanKey on both sides.
' Faster and more robust for “dirty” keys.
'------------------------------------------------------------
Public Function FindTableRowIndexByValue_Clean(ByVal lo As ListObject, ByVal headerName As String, ByVal findValue As String) As Long
    Dim col As Long, a As Variant, r As Long, n As Long, needle As String

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    col = TableColIndexByHeaderClean(lo, headerName)
    If col = 0 Then Exit Function

    a = lo.DataBodyRange.Value2
    n = UBound(a, 1)
    needle = CleanKey(findValue)

    For r = 1 To n
        If StrComp(CleanKey(a(r, col)), needle, vbTextCompare) = 0 Then
            FindTableRowIndexByValue_Clean = r
            Exit Function
        End If
    Next r
End Function

'------------------------------------------------------------
' IsCalculatedColumnFast
' True if the column appears to be formula-driven (calculated column).
' (Fast heuristic: looks at first data cell formula.)
'------------------------------------------------------------
Public Function IsCalculatedColumnFast(ByVal lc As ListColumn) As Boolean
    On Error GoTo SafeExit
    If lc Is Nothing Then Exit Function
    If lc.DataBodyRange Is Nothing Then Exit Function

    IsCalculatedColumnFast = (Left$(lc.DataBodyRange.Cells(1, 1).Formula, 1) = "=")
    Exit Function

SafeExit:
    IsCalculatedColumnFast = False
End Function


'================================================================================
' SECTION 5 — TABLE MUTATION HELPERS (CLEAR / RESIZE / FILTER)
'================================================================================





'------------------------------------------------------------
' TryShowAllData
' Attempts to clear filters without throwing 1004 tantrums.
' Works for both sheet-level and table-level filter mode.
'------------------------------------------------------------
Public Sub TryShowAllData(ByVal lo As ListObject)
    Dim ws As Worksheet

    On Error GoTo SafeExit
    If lo Is Nothing Then Exit Sub

    Set ws = lo.Parent
    If ws Is Nothing Then Exit Sub

    ' If either the sheet or the table is in filter mode, attempt to clear.
    If ws.FilterMode Then
        On Error Resume Next
        lo.AutoFilter.ShowAllData
        On Error GoTo SafeExit
    ElseIf Not lo.AutoFilter Is Nothing Then
        If lo.AutoFilter.FilterMode Then
            On Error Resume Next
            lo.AutoFilter.ShowAllData
            On Error GoTo SafeExit
        End If
    End If

SafeExit:
End Sub



'================================================================================
' SECTION 6 — FORM / RANGE HELPERS
'================================================================================

'------------------------------------------------------------
' FindLabelCell
' Finds a label cell in a given column (exact whole-cell match).
' Returns Nothing if not found.
'------------------------------------------------------------
Public Function FindLabelCell(ByVal ws As Worksheet, ByVal labelText As String, ByVal labelCol As Long, _
                              Optional ByVal firstRow As Long = 1, Optional ByVal lastRow As Long = 0) As Range
    Dim rng As Range
    If ws Is Nothing Then Exit Function

    If lastRow = 0 Then
        lastRow = ws.Cells(ws.Rows.Count, labelCol).End(xlUp).Row
        If lastRow < firstRow Then lastRow = firstRow
    End If

    Set rng = ws.Range(ws.Cells(firstRow, labelCol), ws.Cells(lastRow, labelCol))
    Set FindLabelCell = rng.Find(What:=labelText, LookIn:=xlValues, LookAt:=xlWhole, matchCase:=False)
End Function

'------------------------------------------------------------
' SetValueByLabel
' Finds label in labelCol and writes valueToWrite to valueCol on the same row.
' Returns True if written.
'------------------------------------------------------------
Public Function SetValueByLabel(ByVal ws As Worksheet, ByVal labelText As String, ByVal valueToWrite As Variant, _
                                ByVal labelCol As Long, ByVal valueCol As Long, _
                                Optional ByVal firstRow As Long = 1, Optional ByVal lastRow As Long = 0) As Boolean
    Dim hit As Range
    Set hit = FindLabelCell(ws, labelText, labelCol, firstRow, lastRow)

    If Not hit Is Nothing Then
        ws.Cells(hit.Row, valueCol).Value = valueToWrite
        SetValueByLabel = True
    Else
        SetValueByLabel = False
    End If
End Function

'------------------------------------------------------------
' GetFormDocketNumberFromLabel
' Locates a label on a form and reads the corresponding value cell.
'------------------------------------------------------------
Public Function GetFormDocketNumberFromLabel( _
    ByVal wsForm As Worksheet, _
    ByVal docketLabelText As String, _
    ByVal labelCol As Long, _
    ByVal valueCol As Long _
) As String

    Dim hit As Range
    If wsForm Is Nothing Then Exit Function

    Set hit = FindLabelCell(wsForm, docketLabelText, labelCol)
    If hit Is Nothing Then Exit Function

    GetFormDocketNumberFromLabel = CleanKey(wsForm.Cells(hit.Row, valueCol).Value)
End Function

'------------------------------------------------------------
' BuildLabelRowMap
' Builds a dictionary: CleanKey(label) -> row number (first occurrence only)
'------------------------------------------------------------
Public Function BuildLabelRowMap(ByVal ws As Worksheet, ByVal labelCol As Long) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim lastRow As Long, r As Long, key As String

    d.CompareMode = 1 'TextCompare
    If ws Is Nothing Then
        Set BuildLabelRowMap = d
        Exit Function
    End If

    lastRow = ws.Cells(ws.Rows.Count, labelCol).End(xlUp).Row

    For r = 1 To lastRow
        key = CleanKey(ws.Cells(r, labelCol).Value)
        If Len(key) > 0 Then
            If Not d.Exists(key) Then d.Add key, r
        End If
    Next r

    Set BuildLabelRowMap = d
End Function

'------------------------------------------------------------
' WriteFormValue
' Writes value to row mapped by labelRowMap (built by BuildLabelRowMap).
'------------------------------------------------------------
Public Sub WriteFormValue(ByVal ws As Worksheet, ByVal labelRowMap As Object, ByVal labelText As String, ByVal valueCol As Long, ByVal v As Variant)
    Dim k As String
    If ws Is Nothing Then Exit Sub
    If labelRowMap Is Nothing Then Exit Sub

    k = CleanKey(labelText)
    If labelRowMap.Exists(k) Then
        ws.Cells(labelRowMap(k), valueCol).Value = v
    End If
End Sub


'================================================================================
' SECTION 7 — DATE HELPERS
'================================================================================

'------------------------------------------------------------
' FormatAsLongDate
' Formats a value as "Wednesday, 14 January 2026" if it is a date.
'------------------------------------------------------------
Public Function FormatAsLongDate(ByVal v As Variant) As String
    If IsDate(v) Then
        FormatAsLongDate = Format$(CDate(v), "dddd, d mmmm yyyy")
    Else
        FormatAsLongDate = CStr(v)
    End If
End Function

'================================================================================
' SECTION 8 — ERROR LOGGING
'================================================================================

'------------------------------------------------------------
' LogError
' Centralised error logging for all procedures.
' Safe to call inside error handlers (logger must never throw).
'
' Sheet:  "ErrorLog"
' Columns:
'   A Timestamp | B Procedure | C ErrNo | D Description | E Context
'   F User      | G Workbook  | H Machine | I ExcelVersion
'------------------------------------------------------------
Public Sub LogError(ByVal procName As String, ByVal errNumber As Long, ByVal errDescription As String, _
                    Optional ByVal context As String = vbNullString)

    Dim ws As Worksheet
    Dim NextRow As Long

    On Error GoTo SilentFail

    Set ws = GetWorksheetSafe(ThisWorkbook, "ErrorLog")

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.name = "ErrorLog"
        ws.Range("A1:I1").Value = Array( _
            "Timestamp", "Procedure", "Error Number", "Error Description", _
            "Context", "User", "Workbook", "Machine", "ExcelVersion")
    End If

    NextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(NextRow, 1).Value = Now
    ws.Cells(NextRow, 2).Value = procName
    ws.Cells(NextRow, 3).Value = errNumber
    ws.Cells(NextRow, 4).Value = errDescription
    ws.Cells(NextRow, 5).Value = context
    ws.Cells(NextRow, 6).Value = Environ$("Username")
    ws.Cells(NextRow, 7).Value = ThisWorkbook.name
    ws.Cells(NextRow, 8).Value = Environ$("ComputerName")
    ws.Cells(NextRow, 9).Value = Application.Version

SilentFail:
    'Intentionally empty: logging must never interrupt execution.
End Sub

Public Function IsVariant2DArray(ByVal v As Variant) As Boolean
    On Error GoTo No
    If IsArray(v) Then
        Dim lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long
        lb1 = LBound(v, 1): ub1 = UBound(v, 1)
        lb2 = LBound(v, 2): ub2 = UBound(v, 2)
        IsVariant2DArray = True
        Exit Function
    End If
No:
    IsVariant2DArray = False
End Function


'-------------------------------------------------------
' InsertRowsInTable
' Description:
' Inserts blank rows into a target ListObject table.
' Existing table contents are preserved using an array-
' based read / resize / write approach for speed.
'
' Placement rules:
' - If vTopRow is omitted / Empty / Nothing / invalid:
'       append to bottom of table
' - If vTopRow is a valid numeric table row:
'       insert ABOVE that table-relative row
' - If vTopRow is a Range inside the target table body:
'       insert ABOVE that table-relative row
' - If vTopRow is a Range outside the target table:
'       append to bottom of table
'
' Inputs:
' vListObject - Name of the target ListObject
' vTopRow     - Optional insert reference
'               Can be:
'               * omitted / empty / nothing
'               * numeric table-relative row number
'               * Range reference
' vNum        - Number of rows to insert
'               If <= 0, prompts user (default 10)
' passwd      - Optional worksheet password
'
' Outputs:
' - Resizes the table
' - Inserts blank rows into the target position
' - Restores calculated-column formulas after rebuild
'
' Dependencies:
' - FindListObjectByName
' - AppGuard_Begin / AppGuard_End
' - SheetGuard_Begin / SheetGuard_End
' - TryShowAllData
' - IsVarMissingOrEmptyOrNothing
' - pwd (project-level password variable/constant)
' - TSheetGuardState
'
' Assumptions:
' - The table exists and has a valid header row
' - Helper procedures and guard framework already exist
' - Calculated columns should continue through inserted rows
'-------------------------------------------------------
Public Sub InsertRowsInTable( _
    ByVal vListObject As String, _
    Optional ByVal vTopRow As Variant, _
    Optional ByVal vNum As Long = 0, _
    Optional ByVal passwd As String = vbNullString)

    Dim lo As ListObject
    Dim ws As Worksheet
    Dim shGuard As TSheetGuardState
    Dim inputVal As Variant

    Dim oldArr As Variant
    Dim outArr() As Variant

    Dim rowCount As Long
    Dim colCount As Long
    Dim insertPos As Long
    Dim newTotalRows As Long

    Dim headerTopLeft As Range
    Dim newRange As Range
    Dim dbNew As Range

    Dim r As Long
    Dim c As Long

    On Error GoTo CleanFail
    AppGuard_Begin

    '-------------------------------------------------------
    ' Resolve target table and parent worksheet
    '-------------------------------------------------------
    Set lo = FindListObjectByName(ThisWorkbook, vListObject)
    If lo Is Nothing Then
        Err.Raise 5, "InsertRowsInTable", "Table not found: " & vListObject
    End If

    Set ws = lo.Parent

    If Len(passwd) = 0 Then
        passwd = pwd
    End If

    '-------------------------------------------------------
    ' Resolve requested number of rows to insert
    ' If not provided, prompt user with a default of 10
    '-------------------------------------------------------
    If vNum <= 0 Then
        inputVal = Application.InputBox( _
                        Prompt:="How many blank rows would you like to add?", _
                        Title:="Insert Rows (" & lo.name & ")", _
                        Default:=10, _
                        Type:=1)

        If inputVal = False Then GoTo CleanExit
        If CLng(inputVal) <= 0 Then GoTo CleanExit

        vNum = CLng(inputVal)
    End If

    If vNum < 1 Then GoTo CleanExit

    '-------------------------------------------------------
    ' Temporarily unprotect worksheet only if required
    '-------------------------------------------------------
    shGuard = SheetGuard_Begin(ws, passwd)

    '-------------------------------------------------------
    ' Remove active filters before structural changes
    '-------------------------------------------------------
    TryShowAllData lo

    '-------------------------------------------------------
    ' Capture current table dimensions and body data
    ' BEFORE resizing the table
    '-------------------------------------------------------
    colCount = lo.ListColumns.Count

    If lo.DataBodyRange Is Nothing Then
        rowCount = 0
    Else
        rowCount = lo.DataBodyRange.Rows.Count

        ' Preserve both formulas and literal values
        oldArr = lo.DataBodyRange.Formula
    End If

    '-------------------------------------------------------
    ' Resolve insert position
    '
    ' If the passed reference/selection is outside the
    ' target table, rows are appended to the bottom.
    '-------------------------------------------------------
    insertPos = ResolveInsertPosition(lo, vTopRow, rowCount)

    If insertPos < 1 Then insertPos = 1
    If insertPos > rowCount + 1 Then insertPos = rowCount + 1

    '-------------------------------------------------------
    ' Build output array:
    ' 1. Copy rows before insertion point
    ' 2. Insert blank rows
    ' 3. Copy rows after insertion point
    '-------------------------------------------------------
    newTotalRows = rowCount + vNum
    ReDim outArr(1 To newTotalRows, 1 To colCount)

    ' Copy rows before insertion point
    If rowCount > 0 And insertPos > 1 Then
        For r = 1 To insertPos - 1
            For c = 1 To colCount
                outArr(r, c) = oldArr(r, c)
            Next c
        Next r
    End If

    ' Insert blank rows
    For r = insertPos To insertPos + vNum - 1
        For c = 1 To colCount
            outArr(r, c) = vbNullString
        Next c
    Next r

    ' Copy rows after insertion point
    If rowCount > 0 And insertPos <= rowCount Then
        For r = insertPos To rowCount
            For c = 1 To colCount
                outArr(r + vNum, c) = oldArr(r, c)
            Next c
        Next r
    End If

    '-------------------------------------------------------
    ' Resize table once only
    '-------------------------------------------------------
    Set headerTopLeft = lo.HeaderRowRange.Cells(1, 1)
    Set newRange = headerTopLeft.Resize(newTotalRows + 1, colCount)
    lo.Resize newRange

    '-------------------------------------------------------
    ' Get resized data body range
    '-------------------------------------------------------
    Set dbNew = lo.DataBodyRange

    '-------------------------------------------------------
    ' Write rebuilt table body in a single assignment
    '
    ' Using .Formula preserves both literal values and any
    ' existing formulas captured from the original table.
    '-------------------------------------------------------
    dbNew.Formula = outArr

    '-------------------------------------------------------
    ' Force calculated/formula columns to refill down the
    ' full resized table body.
    '
    ' This is more reliable than only filling the inserted
    ' block because ListObject resize + array writes can
    ' break Excel's native calculated-column propagation.
    '-------------------------------------------------------
    FillDownCalculatedColumns lo

CleanExit:
    '-------------------------------------------------------
    ' Always restore worksheet protection state and app state
    '-------------------------------------------------------
    On Error Resume Next
    SheetGuard_End ws, shGuard, passwd
    AppGuard_End
    On Error GoTo 0
    Exit Sub

CleanFail:
    MsgBox "InsertRowsInTable failed (" & vListObject & "): " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

'-------------------------------------------------------
' ResolveInsertPosition
' Description:
' Converts the optional vTopRow input into a safe,
' bounded, table-relative insert position.
'
' Rules:
' - Missing / Empty / Nothing / invalid => append
' - Numeric => use as table-relative row
' - Range inside table body => use that row
' - Range outside target table => append
'
' Inputs:
' lo       - Target ListObject
' vTopRow  - Variant insert reference
' rowCount - Current number of table body rows
'
' Outputs:
' Returns a valid insert position between 1 and rowCount+1
'
' Assumptions:
' - rowCount is already resolved from lo.DataBodyRange
'-------------------------------------------------------
Private Function ResolveInsertPosition( _
    ByVal lo As ListObject, _
    ByVal vTopRow As Variant, _
    ByVal rowCount As Long) As Long

    Dim rngRef As Range
    Dim rngHit As Range
    Dim resultPos As Long

    ' Default behaviour = append to bottom
    resultPos = rowCount + 1

    '-------------------------------------------------------
    ' Missing / Empty / Nothing => append
    '-------------------------------------------------------
    If IsVarMissingOrEmptyOrNothing(vTopRow) Then
        ResolveInsertPosition = resultPos
        Exit Function
    End If

    '-------------------------------------------------------
    ' If a Range was passed:
    ' - If inside target DataBodyRange => insert above that row
    ' - If outside target table => append
    '-------------------------------------------------------
    If IsObject(vTopRow) Then
        On Error Resume Next
        Set rngRef = vTopRow
        On Error GoTo 0

        If Not rngRef Is Nothing Then
            If Not lo.DataBodyRange Is Nothing Then
                Set rngHit = Intersect(rngRef.Cells(1, 1), lo.DataBodyRange)

                If Not rngHit Is Nothing Then
                    resultPos = rngHit.Row - lo.DataBodyRange.Row + 1
                Else
                    resultPos = rowCount + 1
                End If
            Else
                ' Empty table: any outside selection still appends,
                ' which for an empty table is position 1
                resultPos = 1
            End If

            ResolveInsertPosition = resultPos
            Exit Function
        End If
    End If

    '-------------------------------------------------------
    ' Numeric or numeric-like value => use as table row
    ' Invalid values fall back to append
    '-------------------------------------------------------
    If IsNumeric(vTopRow) Then
        resultPos = CLng(vTopRow)
    Else
        resultPos = rowCount + 1
    End If

    If resultPos < 1 Then resultPos = 1
    If resultPos > rowCount + 1 Then resultPos = rowCount + 1

    ResolveInsertPosition = resultPos
End Function

'-------------------------------------------------------
' ColumnHasFormula
' Description:
' Safely determines whether the first data row in a given
' column contains a formula.
'
' Inputs:
' db  - DataBodyRange of the resized table
' colIndex - 1-based column index
'
' Outputs:
' True if the first body cell in that column has a formula
'
' Assumptions:
' - db is not Nothing
' - colIndex is within db column bounds
'-------------------------------------------------------
Private Function ColumnHasFormula( _
    ByVal db As Range, _
    ByVal colIndex As Long) As Boolean

    On Error GoTo SafeExit

    If db Is Nothing Then Exit Function
    If db.Rows.Count < 1 Then Exit Function
    If colIndex < 1 Or colIndex > db.Columns.Count Then Exit Function

    ColumnHasFormula = db.Cells(1, colIndex).HasFormula

SafeExit:
End Function

'============================================================
' DeleteSelectedRowsFromTable
'
' FAST: deletes selected rows from a specified ListObject by
' rebuilding the table in-memory and resizing once.
'
' FIXES:
'   1. Preserves formulas by using .Formula instead of .Value
'   2. Clears ALL stale data below the resized table within the
'      table column span, not just the original body tail
'============================================================

Public Sub DeleteSelectedRowsFromTable( _
    ByVal vListObject As String, _
    Optional ByVal passwd As String = vbNullString)

    Dim lo As ListObject
    Dim ws As Worksheet
    Dim sel As Range
    Dim deleteRows As Range
    Dim resp As VbMsgBoxResult

    Dim db As Range
    Dim firstDataRow As Long
    Dim lastDataRow As Long

    Dim dataArr As Variant
    Dim outArr() As Variant
    Dim keepCount As Long
    Dim r As Long, c As Long

    Dim delMask() As Boolean
    Dim area As Range
    Dim rr As Range
    Dim absRow As Long
    Dim idx As Long

    Dim headerTopLeft As Range
    Dim newRange As Range
    Dim newTotalRows As Long
    Dim colCount As Long
    Dim rowCount As Long

    ' --- original body bounds ---
    Dim oldBodyTopLeft As Range
    Dim oldRowCount As Long

    ' --- orphan cleanup bounds ---
    Dim firstOrphanRow As Long
    Dim lastOrphanRow As Long
    Dim orphan As Range

    Dim shGuard As TSheetGuardState

    On Error GoTo CleanFail
    AppGuard_Begin

    '----------------------------------------
    ' Resolve table
    '----------------------------------------
    Set lo = FindListObjectByName(ThisWorkbook, vListObject)
    If lo Is Nothing Then Err.Raise 5, , "Table not found: " & vListObject

    Set ws = lo.Parent
    If lo.DataBodyRange Is Nothing Then GoTo CleanExit

    Set sel = Selection
    If sel Is Nothing Then GoTo CleanExit

    If Not sel.Worksheet Is ws Then
        MsgBox "Selection is not on the same sheet as table '" & lo.name & "'.", vbExclamation
        GoTo CleanExit
    End If

    If Len(passwd) = 0 Then passwd = pwd

    '----------------------------------------
    ' Unprotect only if needed
    '----------------------------------------
    shGuard = SheetGuard_Begin(ws, passwd)

    ' Remove filters before structural change
    TryShowAllData lo

    Set db = lo.DataBodyRange
    Set deleteRows = Intersect(sel.EntireRow, db.EntireRow)

    If deleteRows Is Nothing Then
        MsgBox "Selection is not within table '" & lo.name & "'.", vbExclamation
        GoTo CleanExit
    End If

    resp = MsgBox("Delete " & deleteRows.Rows.Count & _
                  " selected row(s) from table '" & lo.name & "'?", _
                  vbYesNo + vbQuestion, "Confirm Delete")
    If resp <> vbYes Then GoTo CleanExit

    '----------------------------------------
    ' Setup bounds + arrays
    '----------------------------------------
    firstDataRow = db.Row
    rowCount = db.Rows.Count
    colCount = db.Columns.Count
    lastDataRow = firstDataRow + rowCount - 1

    Set oldBodyTopLeft = db.Cells(1, 1)
    oldRowCount = rowCount

    ReDim delMask(1 To rowCount)

    '----------------------------------------
    ' Build delete mask
    '----------------------------------------
    For Each area In deleteRows.Areas
        For Each rr In area.Rows
            absRow = rr.Row
            If absRow >= firstDataRow And absRow <= lastDataRow Then
                idx = absRow - firstDataRow + 1
                delMask(idx) = True
            End If
        Next rr
    Next area

    '----------------------------------------
    ' Read table once (preserve formulas)
    '----------------------------------------
    dataArr = db.Formula

    ReDim outArr(1 To rowCount, 1 To colCount)

    keepCount = 0
    For r = 1 To rowCount
        If Not delMask(r) Then
            keepCount = keepCount + 1
            For c = 1 To colCount
                outArr(keepCount, c) = dataArr(r, c)
            Next c
        End If
    Next r

    '----------------------------------------
    ' Write back + resize once
    '----------------------------------------
    Set headerTopLeft = lo.HeaderRowRange.Cells(1, 1)

    If keepCount = 0 Then
        ' No rows left: resize to header only
        Set newRange = headerTopLeft.Resize(1, colCount)
        lo.Resize newRange
    Else
        ' Write kept rows to top of existing body range
        oldBodyTopLeft.Resize(keepCount, colCount).Formula = outArr

        ' Resize table down
        newTotalRows = keepCount
        Set newRange = headerTopLeft.Resize(newTotalRows + 1, colCount)
        lo.Resize newRange
    End If

    '----------------------------------------
    ' Clear ALL stale cells below resized table
    ' within the table column span
    '----------------------------------------
    If keepCount = 0 Then
        firstOrphanRow = oldBodyTopLeft.Row
    Else
        firstOrphanRow = oldBodyTopLeft.Row + keepCount
    End If

    lastOrphanRow = LastUsedRowInColumnSpan(ws, oldBodyTopLeft.Column, oldBodyTopLeft.Column + colCount - 1)
    If lastOrphanRow < (oldBodyTopLeft.Row + oldRowCount - 1) Then
        lastOrphanRow = oldBodyTopLeft.Row + oldRowCount - 1
    End If

    If lastOrphanRow >= firstOrphanRow Then
        Set orphan = ws.Range(ws.Cells(firstOrphanRow, oldBodyTopLeft.Column), _
                              ws.Cells(lastOrphanRow, oldBodyTopLeft.Column + colCount - 1))
        orphan.ClearContents
        orphan.ClearFormats
    End If

CleanExit:
    SheetGuard_End ws, shGuard, passwd
    AppGuard_End
    Exit Sub

CleanFail:
    MsgBox "DeleteSelectedRowsFromTable failed: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub


Private Function LastUsedRowInColumnSpan( _
    ByVal ws As Worksheet, _
    ByVal firstCol As Long, _
    ByVal lastCol As Long) As Long

    Dim c As Long
    Dim lastRow As Long
    Dim tmp As Long

    lastRow = 0

    For c = firstCol To lastCol
        tmp = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If tmp > lastRow Then lastRow = tmp
    Next c

    LastUsedRowInColumnSpan = lastRow
End Function

Public Function IsVarMissingOrEmptyOrNothing(ByVal v As Variant) As Boolean
    If IsMissing(v) Then
        IsVarMissingOrEmptyOrNothing = True
        Exit Function
    End If

    If IsEmpty(v) Then
        IsVarMissingOrEmptyOrNothing = True
        Exit Function
    End If

    ' Only legal to use "Is Nothing" if v is an object variant
    If IsObject(v) Then
        IsVarMissingOrEmptyOrNothing = (v Is Nothing)
        Exit Function
    End If

    IsVarMissingOrEmptyOrNothing = False
End Function

Public Sub SafeSortTable( _
        ByVal tbl As ListObject, _
        ByVal columnName As String, _
        Optional ByVal ascending As Boolean = True, _
        Optional ByVal password As String = "")

    Dim ws As Worksheet
    Dim sortOrder As XlSortOrder
    
    On Error GoTo CleanFail
    
    Set ws = tbl.Parent
    
    If ascending Then
        sortOrder = xlAscending
    Else
        sortOrder = xlDescending
    End If
    
    'Temporarily unprotect
    'ws.Unprotect password
    
    With tbl.Sort
        .SortFields.Clear
        
        .SortFields.Add _
            key:=tbl.ListColumns(columnName).DataBodyRange, _
            SortOn:=xlSortOnValues, _
            Order:=sortOrder, _
            DataOption:=xlSortNormal
            
        .header = xlYes
        .matchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

CleanExit:

    Exit Sub

CleanFail:

    Debug.Print "SafeSortTable Error: " & Err.Description
    Resume CleanExit

End Sub


'-------------------------------------------------------
' Sort2DArrayByTwoKeys
'
' Purpose:
'   Sort a 2D variant array in-place using:
'   - Primary column
'   - Secondary column
'
' Arguments:
'   arr             : 2D 1-based array
'   primaryCol      : primary sort column index
'   secondaryCol    : secondary sort column index
'   primaryAsc      : True for ascending primary sort
'   secondaryAsc    : True for ascending secondary sort
'
' Notes:
'   - Designed for relatively small UI arrays such as
'     UserForm ListBox data
'   - Sorts rows in-place
'   - Compares values as text
'-------------------------------------------------------
Public Sub Sort2DArrayByTwoKeys(ByRef arr As Variant, _
                                ByVal primaryCol As Long, _
                                ByVal secondaryCol As Long, _
                                Optional ByVal primaryAsc As Boolean = True, _
                                Optional ByVal secondaryAsc As Boolean = True)

    On Error GoTo ErrHandler

    '-------------------------------------------------------
    ' Variable Declarations
    '-------------------------------------------------------
    Dim r1 As Long
    Dim r2 As Long
    Dim c As Long
    Dim colCount As Long
    Dim rowCount As Long
    Dim tmp As Variant
    Dim doSwap As Boolean
    Dim cmpPrimary As Long
    Dim cmpSecondary As Long

    '-------------------------------------------------------
    ' Validation
    '-------------------------------------------------------
    If Not IsArray(arr) Then Exit Sub

    rowCount = UBound(arr, 1)
    colCount = UBound(arr, 2)

    If rowCount <= 1 Then Exit Sub
    If primaryCol < LBound(arr, 2) Or primaryCol > colCount Then Exit Sub
    If secondaryCol < LBound(arr, 2) Or secondaryCol > colCount Then Exit Sub

    ReDim tmp(1 To colCount)

    '-------------------------------------------------------
    ' Main Processing Logic
    '-------------------------------------------------------
    For r1 = 1 To rowCount - 1
        For r2 = r1 + 1 To rowCount

            cmpPrimary = CompareSortValues(arr(r1, primaryCol), arr(r2, primaryCol), primaryAsc)

            If cmpPrimary = 0 Then
                cmpSecondary = CompareSortValues(arr(r1, secondaryCol), arr(r2, secondaryCol), secondaryAsc)
                doSwap = (cmpSecondary > 0)
            Else
                doSwap = (cmpPrimary > 0)
            End If

            If doSwap Then
                For c = 1 To colCount
                    tmp(c) = arr(r1, c)
                    arr(r1, c) = arr(r2, c)
                    arr(r2, c) = tmp(c)
                Next c
            End If
        Next r2
    Next r1

CleanExit:
    Exit Sub

ErrHandler:
    SafeLogError "Sort2DArrayByTwoKeys", Err.Number, Err.Description
    Err.Raise Err.Number, "Sort2DArrayByTwoKeys", Err.Description

End Sub

'-------------------------------------------------------
' CompareSortValues
'
' Purpose:
'   Compares two values for sorting.
'
' Returns:
'   -1 if v1 < v2
'    0 if v1 = v2
'    1 if v1 > v2
'
' Notes:
'   - Comparison is text-based and case-insensitive
'   - Null, Empty, and Error values are treated as ""
'-------------------------------------------------------
Public Function CompareSortValues(ByVal v1 As Variant, _
                                  ByVal v2 As Variant, _
                                  Optional ByVal ascending As Boolean = True) As Long

    On Error GoTo ErrHandler

    Dim s1 As String
    Dim s2 As String
    Dim result As Long

    s1 = UCase$(NzText(v1))
    s2 = UCase$(NzText(v2))

    If s1 < s2 Then
        result = -1
    ElseIf s1 > s2 Then
        result = 1
    Else
        result = 0
    End If

    If ascending Then
        CompareSortValues = result
    Else
        CompareSortValues = -result
    End If

CleanExit:
    Exit Function

ErrHandler:
    CompareSortValues = 0
    Resume CleanExit

End Function

'-------------------------------------------------------
' FillDownCalculatedColumns
' Description:
'   Forces Excel table formula columns to extend into all
'   rows of the table by copying the formula from the first
'   existing formula cell in each column down through the
'   full DataBodyRange.
'
' Inputs:
'   - lo: Target ListObject
'
' Outputs:
'   - Any formula-bearing table columns are filled down
'
' Assumptions:
'   - Formula pattern in each calculated column is already
'     correct in at least one existing data row
'-------------------------------------------------------
Public Sub FillDownCalculatedColumns(ByVal lo As ListObject)

    Dim lc As ListColumn
    Dim rngCol As Range
    Dim firstFormula As String
    Dim c As Range

    On Error GoTo SafeExit

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    For Each lc In lo.ListColumns
        Set rngCol = Nothing
        firstFormula = vbNullString

        On Error Resume Next
        Set rngCol = lc.DataBodyRange
        On Error GoTo SafeExit

        If Not rngCol Is Nothing Then

            '-----------------------------------------------
            ' Find the first formula in the column
            '-----------------------------------------------
            For Each c In rngCol.Cells
                If c.HasFormula Then
                    firstFormula = c.Formula
                    Exit For
                End If
            Next c

            '-----------------------------------------------
            ' If a formula exists in this table column,
            ' force it down the full column
            '-----------------------------------------------
            If Len(firstFormula) > 0 Then
                rngCol.Formula = firstFormula
            End If
        End If
    Next lc

SafeExit:
End Sub

'====================================================================================
' TableColIndexRequired
'
' Returns the ListObject column index for a required header.
' Raises an error if the header is not found.
'====================================================================================
Public Function TableColIndexRequired(ByVal lo As ListObject, ByVal headerName As String) As Long

    Dim i As Long
    Dim hdr As String

    For i = 1 To lo.ListColumns.Count
        hdr = Trim$(CStr(lo.ListColumns(i).name))
        If StrComp(hdr, headerName, vbTextCompare) = 0 Then
            TableColIndexRequired = i
            Exit Function
        End If
    Next i

    Err.Raise 5, "TableColIndexRequired", _
              "Required column '" & headerName & "' was not found in table '" & lo.name & "'."

End Function

'====================================================================================
' TableColIndexAny
'
' Returns the first matching ListObject column index from a list of candidate headers.
' Returns 0 when none are found.
'====================================================================================
Public Function TableColIndexAny(ByVal lo As ListObject, ParamArray headerNames() As Variant) As Long

    Dim i As Long
    Dim j As Long
    Dim hdr As String
    Dim want As String

    For i = 1 To lo.ListColumns.Count
        hdr = Trim$(CStr(lo.ListColumns(i).name))

        For j = LBound(headerNames) To UBound(headerNames)
            want = Trim$(CStr(headerNames(j)))
            If StrComp(hdr, want, vbTextCompare) = 0 Then
                TableColIndexAny = i
                Exit Function
            End If
        Next j
    Next i

    TableColIndexAny = 0

End Function

'====================================================================================
' IsTableDataRowVisible
'
' Returns True when the specified table data row is visible.
' relRow is 1-based relative to DataBodyRange.
'
' This excludes:
'   - filtered-out rows
'   - manually hidden rows
'====================================================================================
Public Function IsTableDataRowVisible(ByVal lo As ListObject, ByVal relRow As Long) As Boolean

    Dim r As Range

    On Error GoTo CleanFail

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If relRow < 1 Or relRow > lo.DataBodyRange.Rows.Count Then Exit Function

    Set r = lo.DataBodyRange.Rows(relRow)

    If r.EntireRow.Hidden Then Exit Function
    If r.RowHeight <= 0 Then Exit Function

    IsTableDataRowVisible = True
    Exit Function

CleanFail:
    IsTableDataRowVisible = False

End Function

'====================================================================================
' AppendLine
'
' Appends a line of text to a string buffer cleanly.
'====================================================================================
Public Sub AppendLine(ByRef buffer As String, ByVal txt As String)

    If Len(txt) = 0 Then Exit Sub

    If Len(buffer) > 0 Then
        buffer = buffer & vbCrLf & txt
    Else
        buffer = txt
    End If

End Sub

'====================================================================================
' TryGetNumber
'
' Attempts to coerce a value to Double.
' Returns True only if the value is non-blank, non-error, and numeric.
'====================================================================================
Public Function TryGetNumber(ByVal v As Variant, ByRef outVal As Double) As Boolean

    If IsError(v) Then Exit Function
    If Len(Trim$(CStr(v))) = 0 Then Exit Function
    If Not IsNumeric(v) Then Exit Function

    outVal = CDbl(v)
    TryGetNumber = True

End Function

'====================================================================================
' TryGetPositiveNumber
'
' Same as TryGetNumber, but requires the value to be > 0.
'====================================================================================
Public Function TryGetPositiveNumber(ByVal v As Variant, ByRef outVal As Double) As Boolean

    If Not TryGetNumber(v, outVal) Then Exit Function
    If outVal <= 0 Then Exit Function

    TryGetPositiveNumber = True

End Function

'-------------------------------------------------------
' CoerceDateForDisplay
' Description:
'   Converts a source value into display-ready date text.
'   Handles true Excel dates, numeric serial dates, text
'   dates, blanks, and invalid values safely.
'
' Inputs:
'   vValue      - Source value to assess
'   dateFormat  - Desired output format
'
' Outputs:
'   Returns a formatted date string where valid
'   Returns vbNullString for blank values
'   Returns original text/value where not a valid date
'
' Assumptions:
'   - Excel serial dates are valid when numeric > 0
'   - Output is intended for UI/list display rather than
'     for date arithmetic
'-------------------------------------------------------
Public Function CoerceDateForDisplay( _
    ByVal vValue As Variant, _
    Optional ByVal dateFormat As String = "dd/mm/yyyy") As String

    On Error GoTo Fallback

    If IsError(vValue) Then
        CoerceDateForDisplay = vbNullString
        Exit Function
    End If

    If IsEmpty(vValue) Or IsNull(vValue) Then
        CoerceDateForDisplay = vbNullString
        Exit Function
    End If

    If VarType(vValue) = vbString Then
        If Len(Trim$(CStr(vValue))) = 0 Then
            CoerceDateForDisplay = vbNullString
            Exit Function
        End If

        If IsDate(vValue) Then
            CoerceDateForDisplay = Format$(CDate(vValue), dateFormat)
        Else
            CoerceDateForDisplay = CStr(vValue)
        End If
        Exit Function
    End If

    If IsDate(vValue) Then
        CoerceDateForDisplay = Format$(CDate(vValue), dateFormat)
        Exit Function
    End If

    If IsNumeric(vValue) Then
        If CDbl(vValue) > 0 Then
            CoerceDateForDisplay = Format$(CDate(CDbl(vValue)), dateFormat)
        Else
            CoerceDateForDisplay = vbNullString
        End If
        Exit Function
    End If

Fallback:
    CoerceDateForDisplay = CStr(vValue)
End Function

