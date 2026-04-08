Attribute VB_Name = "ModVersion"
Option Explicit

Public Const pwd As String = "Walz25!"
'====================================================================================
' BACKTEST / DRY-RUN ENTRYPOINT
'====================================================================================
Public Sub Backtest_UpdateTablesFromActiveWorkbook_Tracking()
    ' Runs a full dry-run: no writes, prints actions to Immediate window (Ctrl+G).
    UpdateTablesFromActiveWorkbook_Tracking WhatIf:=True
End Sub

Public Sub Run_UpdateTablesFromActiveWorkbook_Tracking()
    ' Runs a full dry-run: no writes, prints actions to Immediate window (Ctrl+G).
    UpdateTablesFromActiveWorkbook_Tracking WhatIf:=False, passwd:=pwd
End Sub

Public Sub UpdateTablesFromActiveWorkbook_Tracking( _
    Optional ByVal WhatIf As Boolean = False, _
    Optional ByVal passwd As String = vbNullString)

    Dim procName As String: procName = "UpdateTablesFromActiveWorkbook_Tracking"

    Dim tWbk As Workbook, sWbk As Workbook
    Dim loDriver As ListObject
    Dim r As ListRow

    Dim wsName As String, tblName As String
    Dim doUpdate As Boolean

    Dim frm As frmSelect
    Dim selectedWbName As String

    Dim sWs As Worksheet, tWs As Worksheet
    Dim sourceTbl As ListObject, targetTbl As ListObject
    Dim wsDriver As Worksheet

    Dim srcRows As Long, tgtRows As Long
    Dim srcCols As Long, tgtCols As Long

    Dim srcDict As Object
    Dim hdrArr As Variant
    Dim c As Long
    Dim keyName As String

    Dim tgtNorm As String
    Dim srcIndex As Long

    Dim srcData As Variant
    Dim calcCol() As Boolean
    Dim mapSrcCol() As Long

    Dim reprotect As TSheetGuardState
    Dim guardActive As Boolean

    Dim ErrString As String

    Dim cWS As Long, cTbl As Long, cUpd As Long

    Dim okTables As Object
    Dim failTables As Object
    Dim tableKey As String
    Dim tableFailed As Boolean
    Dim didClear As Boolean

    Dim okList As String, failList As String
    Dim summaryMsg As String, errMsg As String
    Dim k As Variant

    Dim missingPolicy As Long
    Dim resp As VbMsgBoxResult
    Dim perResp As VbMsgBoxResult

    Dim hadHardFail As Boolean
    Dim errNum As Long
    Dim errDesc As String

    Dim j As Long
    Dim segStart As Long, segEnd As Long, segW As Long
    Dim outArr As Variant
    Dim rr As Long, cc As Long
    Dim tBody As Range
    Dim colPtr As Long
    Dim isBreak As Boolean

    ' Debug state
    Dim dbgTargetColName As String
    Dim dbgMapVal As Long
    Dim dbgWriteAddr As String

    On Error GoTo HardFail

    missingPolicy = IIf(WhatIf, 2, -1)
    hadHardFail = False
    guardActive = False

    Set tWbk = ThisWorkbook

    AppGuard_Begin xlCalculationManual, IIf(WhatIf, "DRY-RUN: Updating Tracking tables…", "Updating Tracking tables…")

    Set okTables = CreateObject("Scripting.Dictionary")
    okTables.CompareMode = vbTextCompare

    Set failTables = CreateObject("Scripting.Dictionary")
    failTables.CompareMode = vbTextCompare

    Set frm = New frmSelect
    frm.FrmType = 1
    frm.LoadCombo
    frm.Show
    selectedWbName = frm.SelectedWorkbookName
    Unload frm
    Set frm = Nothing

    If Len(selectedWbName) = 0 Then
        MsgBox "No workbook selected. Exiting.", vbExclamation
        GoTo CleanExit
    End If

    Set sWbk = Nothing
    On Error Resume Next
    Set sWbk = Application.Workbooks(selectedWbName)
    On Error GoTo HardFail

    If sWbk Is Nothing Then
        MsgBox "Workbook '" & selectedWbName & "' is not open / not found.", vbExclamation
        GoTo CleanExit
    End If

    If Not WhatIf Then
        If MsgBox("Transfer data from '" & sWbk.Name & "' to this Tracking workbook?", _
                  vbYesNo + vbQuestion, "Tracking Convert") = vbNo Then
            GoTo CleanExit
        End If
    End If

    If WhatIf Then
        Debug.Print "WhatIf: Would copy Intro!B2:B6 from [" & sWbk.Name & "] to [" & tWbk.Name & "]"
    Else
        On Error Resume Next
        tWbk.Worksheets("Intro").Range("B2:B6").Value2 = sWbk.Worksheets("Intro").Range("B2:B6").Value2
        On Error GoTo HardFail
    End If

    Set loDriver = Nothing
    Set wsDriver = Nothing

    On Error Resume Next
    Set wsDriver = GetWorksheetOfTable(tWbk, "tbl_Tables")
    If Not wsDriver Is Nothing Then Set loDriver = wsDriver.ListObjects("tbl_Tables")
    On Error GoTo HardFail

    If loDriver Is Nothing Then
        ErrString = ErrString & vbNewLine & "Driver table 'tbl_Tables' not found in target workbook."
        GoTo Finalize
    End If

    cWS = GetTableColIndex(loDriver, "Worksheet")
    cTbl = GetTableColIndex(loDriver, "Table")
    cUpd = GetTableColIndex(loDriver, "Update")

    If cWS = 0 Or cTbl = 0 Or cUpd = 0 Then
        ErrString = ErrString & vbNewLine & "tbl_Tables must contain headers: Worksheet, Table, Update."
        GoTo Finalize
    End If

    If loDriver.ListRows.Count = 0 Then GoTo Finalize

    For Each r In loDriver.ListRows

        tableFailed = False
        didClear = False
        guardActive = False
        dbgTargetColName = vbNullString
        dbgMapVal = 0
        dbgWriteAddr = vbNullString

        wsName = Trim$(CStr(r.Range(1, cWS).Value2))
        tblName = Trim$(CStr(r.Range(1, cTbl).Value2))
        doUpdate = myCoerceBool(r.Range(1, cUpd).Value2)

        If Not doUpdate Then GoTo NextRow
        If Len(wsName) = 0 Or Len(tblName) = 0 Then GoTo NextRow

        tableKey = wsName & "|" & tblName

        Set tWs = Nothing
        On Error Resume Next
        Set tWs = tWbk.Worksheets(wsName)
        On Error GoTo HardFail

        If tWs Is Nothing Then
            tableFailed = True
            ErrString = ErrString & vbNewLine & "Sheet missing: [" & wsName & "] (TARGET)."
            If Not failTables.Exists(tableKey) Then failTables.Add tableKey, "Target sheet missing"
            GoTo NextRow
        End If

        Set targetTbl = Nothing
        On Error Resume Next
        Set targetTbl = tWs.ListObjects(tblName)
        On Error GoTo HardFail

        If targetTbl Is Nothing Then
            tableFailed = True
            ErrString = ErrString & vbNewLine & "No table '" & tblName & "' in TARGET sheet '" & wsName & "'."
            If Not failTables.Exists(tableKey) Then failTables.Add tableKey, "Target table missing"
            GoTo NextRow
        End If

        Set sWs = Nothing
        On Error Resume Next
        Set sWs = sWbk.Worksheets(wsName)
        On Error GoTo HardFail

        If sWs Is Nothing Then
            tableFailed = True

            HandleMissingSource missingPolicy, WhatIf, tWs, targetTbl, passwd, _
                                "SOURCE sheet not found for:" & vbNewLine & wsName & " :: " & tblName, _
                                didClear, resp, perResp

            If resp = vbCancel Or perResp = vbCancel Then GoTo CleanExit

            If didClear Then
                ErrString = ErrString & vbNewLine & "Not found (missing SOURCE sheet): " & wsName & " :: " & tblName & " -> TARGET cleared."
                If Not failTables.Exists(tableKey) Then failTables.Add tableKey, "Missing source sheet (target cleared)"
            Else
                ErrString = ErrString & vbNewLine & "Not found (missing SOURCE sheet): " & wsName & " :: " & tblName & " -> TARGET left unchanged."
                If Not failTables.Exists(tableKey) Then failTables.Add tableKey, "Missing source sheet (target NOT cleared)"
            End If

            GoTo NextRow
        End If

        Set sourceTbl = Nothing
        On Error Resume Next
        Set sourceTbl = sWs.ListObjects(tblName)
        On Error GoTo HardFail

        If sourceTbl Is Nothing Then
            tableFailed = True

            HandleMissingSource missingPolicy, WhatIf, tWs, targetTbl, passwd, _
                                "SOURCE table not found for:" & vbNewLine & wsName & " :: " & tblName, _
                                didClear, resp, perResp

            If resp = vbCancel Or perResp = vbCancel Then GoTo CleanExit

            If didClear Then
                ErrString = ErrString & vbNewLine & "Not found (missing SOURCE table): " & wsName & " :: " & tblName & " -> TARGET cleared."
                If Not failTables.Exists(tableKey) Then failTables.Add tableKey, "Missing source table (target cleared)"
            Else
                ErrString = ErrString & vbNewLine & "Not found (missing SOURCE table): " & wsName & " :: " & tblName & " -> TARGET left unchanged."
                If Not failTables.Exists(tableKey) Then failTables.Add tableKey, "Missing source table (target NOT cleared)"
            End If

            GoTo NextRow
        End If

        On Error Resume Next
        If sourceTbl.ShowAutoFilter Then
            If sourceTbl.AutoFilter.FilterMode Then sourceTbl.AutoFilter.ShowAllData
        End If
        If targetTbl.ShowAutoFilter Then
            If targetTbl.AutoFilter.FilterMode Then targetTbl.AutoFilter.ShowAllData
        End If
        sourceTbl.Range.Rows.Hidden = False
        targetTbl.Range.Rows.Hidden = False
        If targetTbl.ShowTotals Then targetTbl.ShowTotals = False
        If sourceTbl.ShowTotals Then sourceTbl.ShowTotals = False
        On Error GoTo HardFail

        '==========================================================
        ' SAFE ROW COUNTS - DO NOT TRUST DATA TO EXIST
        '==========================================================
        If sourceTbl.DataBodyRange Is Nothing Then
            srcRows = 0
        Else
            srcRows = sourceTbl.DataBodyRange.Rows.Count
        End If

        If targetTbl.DataBodyRange Is Nothing Then
            tgtRows = 0
        Else
            tgtRows = targetTbl.DataBodyRange.Rows.Count
        End If

        Debug.Print String(100, "=")
        Debug.Print "TABLE START: " & wsName & " :: " & tblName
        Debug.Print "Source rows:   " & srcRows
        Debug.Print "Target rows before resize:   " & tgtRows
        Debug.Print "Source totals:              " & sourceTbl.ShowTotals
        Debug.Print "Target totals:              " & targetTbl.ShowTotals

        reprotect = SheetGuard_Begin(tWs, passwd)
        guardActive = True

        '==========================================================
        ' SOURCE EMPTY -> CLEAR TARGET TO HEADER ONLY / SKIP COPY
        '==========================================================
        If srcRows = 0 Then
            If WhatIf Then
                Debug.Print "WhatIf: " & wsName & " :: " & tblName & " -> Source has 0 rows. Would clear TARGET to header only."
            Else
                ClearListObjectData targetTbl
            End If

            If Not failTables.Exists(tableKey) Then failTables.Add tableKey, "Source empty (target cleared)"
            GoTo AfterCopy
        End If

        Set srcDict = CreateObject("Scripting.Dictionary")
        srcDict.CompareMode = vbTextCompare

        hdrArr = sourceTbl.HeaderRowRange.Value2
        If IsArray(hdrArr) Then
            For c = LBound(hdrArr, 2) To UBound(hdrArr, 2)
                keyName = NormalizeHeader(CStr(hdrArr(1, c)))
                If Len(keyName) > 0 Then
                    If Not srcDict.Exists(keyName) Then
                        srcDict.Add keyName, c
                    Else
                        tableFailed = True
                        ErrString = ErrString & vbNewLine & _
                                    "Duplicate SOURCE header after normalize: [" & keyName & "] in '" & wsName & " :: " & tblName & "'."
                    End If
                End If
            Next c
        Else
            keyName = NormalizeHeader(CStr(hdrArr))
            If Len(keyName) > 0 Then srcDict(keyName) = 1
        End If

        If WhatIf Then
            Debug.Print "WhatIf: " & wsName & " :: " & tblName & " -> Would resize TARGET to " & srcRows & " rows."
        Else
            ResizeListObjectExact targetTbl, srcRows
        End If

        If targetTbl.DataBodyRange Is Nothing Then
            tgtRows = 0
        Else
            tgtRows = targetTbl.DataBodyRange.Rows.Count
        End If

        Debug.Print "Target rows after resize: " & tgtRows
        Debug.Print "Target range after resize: " & targetTbl.Range.Address(0, 0)
        If targetTbl.DataBodyRange Is Nothing Then
            Debug.Print "Target body after resize: Nothing"
        Else
            Debug.Print "Target body after resize: " & targetTbl.DataBodyRange.Address(0, 0)
        End If

        If tgtRows <> srcRows Then
            tableFailed = True
            ErrString = ErrString & vbNewLine & wsName & " :: " & tblName & _
                        ": Target rows (" & tgtRows & ") <> Source rows (" & srcRows & ")."
            If Not failTables.Exists(tableKey) Then failTables.Add tableKey, "Row mismatch after resize"
            GoTo AfterCopy
        End If

        '==========================================================
        ' SAFE: ONLY READ SOURCE DATA WHEN BODY EXISTS
        '==========================================================
        If sourceTbl.DataBodyRange Is Nothing Then
            tableFailed = True
            ErrString = ErrString & vbNewLine & wsName & " :: " & tblName & ": Source DataBodyRange missing despite srcRows > 0."
            If Not failTables.Exists(tableKey) Then failTables.Add tableKey, "Source DataBodyRange missing"
            GoTo AfterCopy
        End If

        srcData = sourceTbl.DataBodyRange.Value2
        srcData = myEnsure2D(srcData)

        srcCols = sourceTbl.ListColumns.Count
        tgtCols = targetTbl.ListColumns.Count

        Debug.Print "Source cols: " & srcCols
        Debug.Print "Target cols: " & tgtCols
        Debug.Print "srcData row bounds: " & LBound(srcData, 1) & " to " & UBound(srcData, 1)
        Debug.Print "srcData col bounds: " & LBound(srcData, 2) & " to " & UBound(srcData, 2)

        ReDim calcCol(1 To tgtCols) As Boolean
        ReDim mapSrcCol(1 To tgtCols) As Long

        For j = 1 To tgtCols
            calcCol(j) = IsCalculatedColumnFast(targetTbl, j)

            If calcCol(j) Then
                mapSrcCol(j) = -1
            Else
                tgtNorm = NormalizeHeader(targetTbl.ListColumns(j).Name)
                If srcDict.Exists(tgtNorm) Then
                    mapSrcCol(j) = CLng(srcDict(tgtNorm))
                Else
                    mapSrcCol(j) = 0
                    ErrString = ErrString & vbNewLine & _
                                "Missing in Source: [" & targetTbl.ListColumns(j).Name & "] -> [" & tgtNorm & "] in '" & wsName & " :: " & tblName & "'."
                End If
            End If
        Next j

        Debug.Print String(80, "-")
        Debug.Print "COLUMN MAP SNAPSHOT: " & wsName & " :: " & tblName
        For j = 1 To tgtCols
            Debug.Print j, targetTbl.ListColumns(j).Name, "calc=" & calcCol(j), "map=" & mapSrcCol(j)
        Next j

        Set tBody = Nothing
        On Error Resume Next
        Set tBody = targetTbl.DataBodyRange
        On Error GoTo HardFail

        If tBody Is Nothing Then
            tableFailed = True
            ErrString = ErrString & vbNewLine & wsName & " :: " & tblName & ": Target DataBodyRange missing after resize."
            If Not failTables.Exists(tableKey) Then failTables.Add tableKey, "DataBodyRange missing"
            GoTo AfterCopy
        End If

        segStart = 0

        For j = 1 To tgtCols + 1

            If j > tgtCols Then
                isBreak = True
            Else
                isBreak = calcCol(j)
            End If

            If segStart = 0 Then
                If j <= tgtCols Then
                    If Not calcCol(j) Then segStart = j
                End If
            Else
                If isBreak Then
                    segEnd = j - 1
                    segW = segEnd - segStart + 1

                    If WhatIf Then
                        Debug.Print "WhatIf: " & wsName & " :: " & tblName & _
                                    " -> Would write TARGET cols " & segStart & " to " & segEnd & " in one batch."
                    Else
                        On Error GoTo SegmentFail

                        Debug.Print String(80, "-")
                        Debug.Print "BUILD SEGMENT"
                        Debug.Print "Table:", wsName, tblName
                        Debug.Print "segStart:", segStart
                        Debug.Print "segEnd:", segEnd
                        Debug.Print "segW:", segW
                        Debug.Print "srcRows:", srcRows
                        Debug.Print "srcCols:", srcCols
                        Debug.Print "tBody rows:", tBody.Rows.Count
                        Debug.Print "tBody cols:", tBody.Columns.Count
                        Debug.Print "srcData row bounds:", LBound(srcData, 1), "to", UBound(srcData, 1)
                        Debug.Print "srcData col bounds:", LBound(srcData, 2), "to", UBound(srcData, 2)

                        ReDim outArr(1 To srcRows, 1 To segW)

                        For rr = 1 To srcRows
                            cc = 1
                            For colPtr = segStart To segEnd

                                dbgTargetColName = vbNullString
                                dbgMapVal = 0

                                If colPtr >= 1 And colPtr <= targetTbl.ListColumns.Count Then
                                    dbgTargetColName = targetTbl.ListColumns(colPtr).Name
                                Else
                                    dbgTargetColName = "#colPtr out of target bounds"
                                End If

                                If colPtr >= LBound(mapSrcCol) And colPtr <= UBound(mapSrcCol) Then
                                    dbgMapVal = mapSrcCol(colPtr)
                                Else
                                    dbgMapVal = -999
                                End If

                                If rr = 1 Or rr = srcRows Then
                                    Debug.Print "READ:", _
                                                "rr=" & rr, _
                                                "cc=" & cc, _
                                                "colPtr=" & colPtr, _
                                                "target=" & dbgTargetColName, _
                                                "map=" & dbgMapVal
                                End If

                                If dbgMapVal > 0 Then
                                    outArr(rr, cc) = srcData(rr, dbgMapVal)
                                Else
                                    outArr(rr, cc) = vbNullString
                                End If

                                cc = cc + 1
                            Next colPtr
                        Next rr

                        dbgWriteAddr = tBody.Columns(segStart).Resize(srcRows, segW).Address(0, 0)

                        Debug.Print "WRITE:", _
                                    "segStart=" & segStart, _
                                    "segEnd=" & segEnd, _
                                    "segW=" & segW, _
                                    "rows=" & srcRows
                        Debug.Print "WRITE ADDRESS:", dbgWriteAddr

                        tBody.Columns(segStart).Resize(srcRows, segW).Value2 = outArr

                        On Error GoTo HardFail
                    End If

                    segStart = 0
                End If
            End If
        Next j

        If Not tableFailed Then
            If Not okTables.Exists(tableKey) Then okTables.Add tableKey, True
        Else
            If Not failTables.Exists(tableKey) Then failTables.Add tableKey, "Failed (table-level issue)"
        End If

AfterCopy:
        On Error Resume Next
        If Not targetTbl Is Nothing Then
            If targetTbl.ShowAutoFilter Then
                If targetTbl.AutoFilter.FilterMode Then targetTbl.AutoFilter.ShowAllData
            End If
        End If

        If guardActive Then
            SheetGuard_End tWs, reprotect, passwd
            guardActive = False
        End If
        On Error GoTo HardFail

NextRow:
        If guardActive Then
            On Error Resume Next
            SheetGuard_End tWs, reprotect, passwd
            guardActive = False
            On Error GoTo HardFail
        End If
    Next r

    GoTo Finalize

SegmentFail:
    Debug.Print String(80, "!")
    Debug.Print "SEGMENT FAIL"
    Debug.Print "Err:", Err.Number, Err.description
    Debug.Print "Table:", wsName, tblName
    Debug.Print "segStart:", segStart
    Debug.Print "segEnd:", segEnd
    Debug.Print "segW:", segW
    Debug.Print "rr:", rr
    Debug.Print "cc:", cc
    Debug.Print "colPtr:", colPtr
    Debug.Print "dbgTargetColName:", dbgTargetColName
    Debug.Print "dbgMapVal:", dbgMapVal
    Debug.Print "srcRows:", srcRows
    Debug.Print "srcCols:", srcCols
    Debug.Print "tgtRows:", tgtRows
    Debug.Print "tgtCols:", tgtCols
    Debug.Print "srcData row bounds:", LBound(srcData, 1), "to", UBound(srcData, 1)
    Debug.Print "srcData col bounds:", LBound(srcData, 2), "to", UBound(srcData, 2)
    Debug.Print "tBody address:", IIf(tBody Is Nothing, "Nothing", tBody.Address(0, 0))
    Debug.Print "dbgWriteAddr:", dbgWriteAddr

    ErrString = ErrString & vbNewLine & _
                "Segment fail in " & wsName & " :: " & tblName & _
                " | segStart=" & segStart & _
                " | segEnd=" & segEnd & _
                " | rr=" & rr & _
                " | cc=" & cc & _
                " | colPtr=" & colPtr & _
                " | targetCol=" & dbgTargetColName & _
                " | map=" & dbgMapVal & _
                " | Err " & Err.Number & " - " & Err.description

    tableFailed = True
    If Not failTables.Exists(tableKey) Then failTables.Add tableKey, "Segment failure"

    On Error GoTo HardFail
    Resume AfterCopy

Finalize:
    On Error Resume Next

    If WhatIf Then
        Debug.Print "WhatIf: Would CalculateFull + lock/protect routines."
    Else
        If Not hadHardFail Then
            Application.CalculateFull
            Relock_All_Tables passwd
            ProtectAllSheets passwd
        End If
    End If

    On Error GoTo 0

    okList = vbNullString
    For Each k In okTables.keys
        okList = okList & vbNewLine & Replace$(CStr(k), "|", " :: ")
    Next k
    If Len(okList) = 0 Then okList = vbNewLine & "(none)"

    failList = vbNullString
    For Each k In failTables.keys
        failList = failList & vbNewLine & Replace$(CStr(k), "|", " :: ") & "  -->  " & CStr(failTables(k))
    Next k
    If Len(failList) = 0 Then failList = vbNewLine & "(none)"

    summaryMsg = "Tables updated:" & okList & vbNewLine & vbNewLine & _
                 "Tables NOT updated:" & failList

    errMsg = "Errors (detail):" & vbNewLine & IIf(Len(ErrString) = 0, "(none)", ErrString)

    If WhatIf Then
        Debug.Print "DRY-RUN complete."
        Debug.Print summaryMsg
        If Len(ErrString) > 0 Then Debug.Print errMsg
    Else
        MsgBox summaryMsg, vbInformation, "Tracking Update Result"
        If Len(ErrString) > 0 Then MsgBox errMsg, vbExclamation, "Tracking Update Errors"
    End If

CleanExit:
    On Error Resume Next
    If guardActive Then SheetGuard_End tWs, reprotect, passwd
    AppGuard_End
    Exit Sub

HardFail:
    errNum = Err.Number
    errDesc = Err.description
    hadHardFail = True

    On Error Resume Next
    If guardActive Then
        SheetGuard_End tWs, reprotect, passwd
        guardActive = False
    End If
    LogError procName, errNum, errDesc
    On Error GoTo 0

    ErrString = ErrString & vbNewLine & "Hard fail (" & procName & "): " & errNum & " - " & errDesc
    Resume Finalize
End Sub

'====================================================================================
' Helper: Resize table to exact number of data rows (0 allowed)
'====================================================================================
Public Sub ResizeListObjectExact(ByVal lo As ListObject, ByVal rowCount As Long)

    Dim ws As Worksheet
    Dim newRange As Range
    Dim topLeft As Range
    Dim totalRows As Long
    Dim totalCols As Long

    If lo Is Nothing Then Exit Sub
    If rowCount < 1 Then
        Err.Raise vbObjectError + 7001, "ResizeListObjectExact", _
                  "rowCount must be >= 1. Use ClearListObjectData for zero-row tables."
    End If

    Set ws = lo.Parent
    Set topLeft = lo.HeaderRowRange.Cells(1, 1)

    totalCols = lo.ListColumns.Count
    totalRows = rowCount + 1   'header + data rows

    Set newRange = ws.Range(topLeft, topLeft.Offset(totalRows - 1, totalCols - 1))

    lo.Resize newRange
End Sub

Public Sub ClearListObjectData(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub

    On Error GoTo SafeExit
    Application.EnableEvents = False

    If Not lo.DataBodyRange Is Nothing Then
        lo.DataBodyRange.Delete
    End If

SafeExit:
    Application.EnableEvents = True
End Sub

Public Sub ClearTableToHeaderOnly(ByVal lo As ListObject)
    ResizeListObjectExact lo, 0
End Sub


Public Sub UpdateTablesFromActiveWorkbook_Tracking_old( _
    Optional ByVal WhatIf As Boolean = False, _
    Optional ByVal passwd As String = vbNullString)

    '====================================================================================
    ' Fast table sync driven by tbl_Tables.
    '
    ' Performance strategy:
    '   - Read source DataBodyRange ONCE into an array
    '   - Build target->source column mapping ONCE
    '   - Write back in as few Range.Value2 assignments as possible by batching
    '     contiguous non-calculated TARGET columns into segments.
    '
    ' Safety:
    '   - Fully qualified workbook/sheet/table
    '   - Never overwrite calculated columns
    '   - Guarded sheet protection + app state restoration
    '====================================================================================

    Dim procName As String: procName = "UpdateTablesFromActiveWorkbook_Tracking"

    Dim tWbk As Workbook, sWbk As Workbook
    Dim loDriver As ListObject
    Dim r As ListRow

    Dim wsName As String, tblName As String
    Dim doUpdate As Boolean

    Dim frm As frmSelect
    Dim selectedWbName As String

    Dim sWs As Worksheet, tWs As Worksheet
    Dim sourceTbl As ListObject, targetTbl As ListObject
    Dim wsDriver As Worksheet

    Dim srcRows As Long, tgtRows As Long
    Dim srcCols As Long, tgtCols As Long

    Dim srcDict As Object             ' key=normalized header, item=source col index (1-based)
    Dim hdrArr As Variant
    Dim c As Long
    Dim keyName As String

    Dim tgtNorm As String
    Dim srcIndex As Long

    Dim srcData As Variant            ' 2D [1..srcRows,1..srcCols]
    Dim calcCol() As Boolean          ' [1..tgtCols]
    Dim mapSrcCol() As Long           ' [1..tgtCols] -> 0 if missing, else source col

    Dim reprotect As TSheetGuardState
    Dim guardActive As Boolean

    Dim ErrString As String

    Dim cWS As Long, cTbl As Long, cUpd As Long

    Dim okTables As Object            ' key = "Sheet|Table"
    Dim failTables As Object          ' key = "Sheet|Table", item = reason
    Dim tableKey As String
    Dim tableFailed As Boolean
    Dim didClear As Boolean

    Dim okList As String, failList As String
    Dim summaryMsg As String, errMsg As String
    Dim k As Variant

    ' Missing-source decision policy:
    '   -1 = not chosen yet (live only)
    '    1 = clear ALL missing-source targets
    '    2 = ask EACH time (live only); in WhatIf, we log instead of prompting
    Dim missingPolicy As Long
    Dim resp As VbMsgBoxResult
    Dim perResp As VbMsgBoxResult

    Dim hadHardFail As Boolean
    Dim errNum As Long
    Dim errDesc As String

    Dim j As Long
    Dim segStart As Long, segEnd As Long, segW As Long
    Dim outArr As Variant
    Dim rr As Long, cc As Long
    Dim tBody As Range
    Dim colPtr As Long
    Dim isBreak As Boolean

    On Error GoTo HardFail

    missingPolicy = IIf(WhatIf, 2, -1)
    hadHardFail = False
    guardActive = False

    Set tWbk = ThisWorkbook

    AppGuard_Begin xlCalculationManual, IIf(WhatIf, "DRY-RUN: Updating Tracking tables…", "Updating Tracking tables…")
    Debug.Print "Checkpoint: AppGuard_Begin complete"

    Set okTables = CreateObject("Scripting.Dictionary")
    okTables.CompareMode = vbTextCompare

    Set failTables = CreateObject("Scripting.Dictionary")
    failTables.CompareMode = vbTextCompare

    '------------------------------------------------------------
    ' Select SOURCE workbook via form
    '------------------------------------------------------------
    Set frm = New frmSelect
    frm.FrmType = 1
    frm.LoadCombo
    frm.Show
    selectedWbName = frm.SelectedWorkbookName
    Unload frm
    Set frm = Nothing

    If Len(selectedWbName) = 0 Then
        MsgBox "No workbook selected. Exiting.", vbExclamation
        GoTo CleanExit
    End If

    Set sWbk = Nothing
    On Error Resume Next
    Set sWbk = Application.Workbooks(selectedWbName)
    On Error GoTo HardFail

    If sWbk Is Nothing Then
        MsgBox "Workbook '" & selectedWbName & "' is not open / not found.", vbExclamation
        GoTo CleanExit
    End If

    Debug.Print "Checkpoint: Source workbook resolved -> " & sWbk.Name

    If Not WhatIf Then
        If MsgBox("Transfer data from '" & sWbk.Name & "' to this Tracking workbook?", _
                  vbYesNo + vbQuestion, "Tracking Convert") = vbNo Then
            GoTo CleanExit
        End If
    End If

    '------------------------------------------------------------
    ' Intro copy (best-effort; non-fatal)
    '------------------------------------------------------------
    If WhatIf Then
        Debug.Print "WhatIf: Would copy Intro!B2:B6 from [" & sWbk.Name & "] to [" & tWbk.Name & "]"
    Else
        On Error Resume Next
        tWbk.Worksheets("Intro").Range("B2:B6").Value2 = sWbk.Worksheets("Intro").Range("B2:B6").Value2
        On Error GoTo HardFail
    End If

    Debug.Print "Checkpoint: Intro copy complete"

    '------------------------------------------------------------
    ' Resolve driver table tbl_Tables (search once)
    '------------------------------------------------------------
    Set loDriver = Nothing
    Set wsDriver = Nothing

    On Error Resume Next
    Set wsDriver = GetWorksheetOfTable(tWbk, "tbl_Tables")
    If Not wsDriver Is Nothing Then Set loDriver = wsDriver.ListObjects("tbl_Tables")
    On Error GoTo HardFail

    If loDriver Is Nothing Then
        ErrString = ErrString & vbNewLine & "Driver table 'tbl_Tables' not found in target workbook."
        GoTo Finalize
    End If

    Debug.Print "Checkpoint: Driver table resolved"

    cWS = GetTableColIndex(loDriver, "Worksheet")
    cTbl = GetTableColIndex(loDriver, "Table")
    cUpd = GetTableColIndex(loDriver, "Update")

    If cWS = 0 Or cTbl = 0 Or cUpd = 0 Then
        ErrString = ErrString & vbNewLine & "tbl_Tables must contain headers: Worksheet, Table, Update."
        GoTo Finalize
    End If

    If loDriver.ListRows.Count = 0 Then GoTo Finalize

    Debug.Print "Checkpoint: Entering driver loop"

    '------------------------------------------------------------
    ' Driver loop
    '------------------------------------------------------------
    For Each r In loDriver.ListRows

        tableFailed = False
        didClear = False
        guardActive = False

        wsName = Trim$(CStr(r.Range(1, cWS).Value2))
        tblName = Trim$(CStr(r.Range(1, cTbl).Value2))
        doUpdate = myCoerceBool(r.Range(1, cUpd).Value2)

        If Not doUpdate Then GoTo NextRow
        If Len(wsName) = 0 Or Len(tblName) = 0 Then GoTo NextRow

        tableKey = wsName & "|" & tblName

        Debug.Print "Processing: " & tableKey

        '----------------------------
        ' Resolve TARGET sheet/table first
        '----------------------------
        Set tWs = Nothing
        On Error Resume Next
        Set tWs = tWbk.Worksheets(wsName)
        On Error GoTo HardFail

        If tWs Is Nothing Then
            tableFailed = True
            ErrString = ErrString & vbNewLine & "Sheet missing: [" & wsName & "] (TARGET)."
            If Not failTables.Exists(tableKey) Then failTables.Add tableKey, "Target sheet missing"
            GoTo NextRow
        End If

        Set targetTbl = Nothing
        On Error Resume Next
        Set targetTbl = tWs.ListObjects(tblName)
        On Error GoTo HardFail

        If targetTbl Is Nothing Then
            tableFailed = True
            ErrString = ErrString & vbNewLine & "No table '" & tblName & "' in TARGET sheet '" & wsName & "'."
            If Not failTables.Exists(tableKey) Then failTables.Add tableKey, "Target table missing"
            GoTo NextRow
        End If

        '----------------------------
        ' Resolve SOURCE sheet/table
        '----------------------------
        Set sWs = Nothing
        On Error Resume Next
        Set sWs = sWbk.Worksheets(wsName)
        On Error GoTo HardFail

        If sWs Is Nothing Then
            tableFailed = True
            Debug.Print "MISSING SOURCE SHEET detected | WhatIf=" & WhatIf & " | missingPolicy=" & missingPolicy & " | " & wsName & " :: " & tblName

            HandleMissingSource missingPolicy, WhatIf, tWs, targetTbl, passwd, _
                                "SOURCE sheet not found for:" & vbNewLine & wsName & " :: " & tblName, _
                                didClear, resp, perResp

            If resp = vbCancel Or perResp = vbCancel Then GoTo CleanExit

            If didClear Then
                ErrString = ErrString & vbNewLine & "Not found (missing SOURCE sheet): " & wsName & " :: " & tblName & " -> TARGET cleared."
                If Not failTables.Exists(tableKey) Then failTables.Add tableKey, "Missing source sheet (target cleared)"
            Else
                ErrString = ErrString & vbNewLine & "Not found (missing SOURCE sheet): " & wsName & " :: " & tblName & " -> TARGET left unchanged."
                If Not failTables.Exists(tableKey) Then failTables.Add tableKey, "Missing source sheet (target NOT cleared)"
            End If

            GoTo NextRow
        End If

        Set sourceTbl = Nothing
        On Error Resume Next
        Set sourceTbl = sWs.ListObjects(tblName)
        On Error GoTo HardFail

        If sourceTbl Is Nothing Then
            tableFailed = True
            Debug.Print "MISSING SOURCE TABLE detected | WhatIf=" & WhatIf & " | missingPolicy=" & missingPolicy & " | " & wsName & " :: " & tblName

            HandleMissingSource missingPolicy, WhatIf, tWs, targetTbl, passwd, _
                                "SOURCE table not found for:" & vbNewLine & wsName & " :: " & tblName, _
                                didClear, resp, perResp

            If resp = vbCancel Or perResp = vbCancel Then GoTo CleanExit

            If didClear Then
                ErrString = ErrString & vbNewLine & "Not found (missing SOURCE table): " & wsName & " :: " & tblName & " -> TARGET cleared."
                If Not failTables.Exists(tableKey) Then failTables.Add tableKey, "Missing source table (target cleared)"
            Else
                ErrString = ErrString & vbNewLine & "Not found (missing SOURCE table): " & wsName & " :: " & tblName & " -> TARGET left unchanged."
                If Not failTables.Exists(tableKey) Then failTables.Add tableKey, "Missing source table (target NOT cleared)"
            End If

            GoTo NextRow
        End If

        '----------------------------
        ' Best-effort: clear filters/hidden
        '----------------------------
        On Error Resume Next
        If sourceTbl.ShowAutoFilter Then
            If sourceTbl.AutoFilter.FilterMode Then sourceTbl.AutoFilter.ShowAllData
        End If
        If targetTbl.ShowAutoFilter Then
            If targetTbl.AutoFilter.FilterMode Then targetTbl.AutoFilter.ShowAllData
        End If
        sourceTbl.Range.Rows.Hidden = False
        targetTbl.Range.Rows.Hidden = False
        On Error GoTo HardFail

        srcRows = sourceTbl.ListRows.Count

        '----------------------------
        ' Work on TARGET sheet unprotected
        '----------------------------
        reprotect = SheetGuard_Begin(tWs, passwd)
        guardActive = True

        '=========================================================
        ' Source has 0 rows: clear target automatically (no prompt)
        '=========================================================
        If srcRows = 0 Then
            tableFailed = True

            If WhatIf Then
                Debug.Print "WhatIf: " & wsName & " :: " & tblName & " -> Source has 0 rows. Would clear TARGET to header only."
            Else
                ClearTableToHeaderOnly targetTbl
            End If

            ErrString = ErrString & vbNewLine & "Not found (source empty): " & wsName & " :: " & tblName & " (0 rows) -> TARGET cleared."
            If Not failTables.Exists(tableKey) Then failTables.Add tableKey, "Source empty (target cleared)"
            GoTo AfterCopy
        End If

        '------------------------------------------------------------
        ' Build normalized SOURCE header dictionary once
        '------------------------------------------------------------
        Set srcDict = CreateObject("Scripting.Dictionary")
        srcDict.CompareMode = vbTextCompare

        hdrArr = sourceTbl.HeaderRowRange.Value2
        If IsArray(hdrArr) Then
            For c = LBound(hdrArr, 2) To UBound(hdrArr, 2)
                keyName = NormalizeHeader(CStr(hdrArr(1, c)))
                If Len(keyName) > 0 Then
                    If Not srcDict.Exists(keyName) Then
                        srcDict.Add keyName, c
                    Else
                        tableFailed = True
                        ErrString = ErrString & vbNewLine & _
                                    "Duplicate SOURCE header after normalize: [" & keyName & "] in '" & wsName & " :: " & tblName & "'."
                        If Not failTables.Exists(tableKey) Then failTables.Add tableKey, "Duplicate source headers after normalize"
                    End If
                End If
            Next c
        Else
            keyName = NormalizeHeader(CStr(hdrArr))
            If Len(keyName) > 0 Then
                If Not srcDict.Exists(keyName) Then
                    srcDict.Add keyName, 1
                Else
                    tableFailed = True
                    ErrString = ErrString & vbNewLine & _
                                "Duplicate SOURCE header after normalize: [" & keyName & "] in '" & wsName & " :: " & tblName & "'."
                    If Not failTables.Exists(tableKey) Then failTables.Add tableKey, "Duplicate source headers after normalize"
                End If
            End If
        End If

        '------------------------------------------------------------
        ' Resize TARGET rows to match SOURCE rows (exact)
        '------------------------------------------------------------
        If WhatIf Then
            Debug.Print "WhatIf: " & wsName & " :: " & tblName & " -> Would resize TARGET to " & srcRows & " rows."
        Else
            ResizeListObjectRowsExact targetTbl, srcRows
        End If

        tgtRows = targetTbl.ListRows.Count
        If tgtRows <> srcRows Then
            tableFailed = True
            ErrString = ErrString & vbNewLine & wsName & " :: " & tblName & _
                        ": Target rows (" & tgtRows & ") <> Source rows (" & srcRows & ")."
            If Not failTables.Exists(tableKey) Then failTables.Add tableKey, "Row mismatch after resize"
            GoTo AfterCopy
        End If

        '------------------------------------------------------------
        ' Read SOURCE data once into array
        '------------------------------------------------------------
        srcData = sourceTbl.DataBodyRange.Value2
        srcData = myEnsure2D(srcData)

        srcCols = sourceTbl.ListColumns.Count
        tgtCols = targetTbl.ListColumns.Count

        ReDim calcCol(1 To tgtCols) As Boolean
        ReDim mapSrcCol(1 To tgtCols) As Long

        '------------------------------------------------------------
        ' Precompute calculated flags + target->source column mapping once
        '------------------------------------------------------------
        For j = 1 To tgtCols

            calcCol(j) = IsCalculatedColumnFast(targetTbl, j)

            If calcCol(j) Then
                mapSrcCol(j) = -1 ' sentinel; never write
            Else
                tgtNorm = NormalizeHeader(targetTbl.ListColumns(j).Name)
                If srcDict.Exists(tgtNorm) Then
                    srcIndex = CLng(srcDict(tgtNorm))
                    mapSrcCol(j) = srcIndex
                Else
                    mapSrcCol(j) = 0 ' missing in source => clear this column
                    ErrString = ErrString & vbNewLine & _
                                "Missing in Source: [" & targetTbl.ListColumns(j).Name & "] -> [" & tgtNorm & "] in '" & wsName & " :: " & tblName & "'."
                End If
            End If

        Next j

        '------------------------------------------------------------
        ' Fast write: batch contiguous NON-calculated target columns
        '------------------------------------------------------------
        Set tBody = Nothing
        On Error Resume Next
        Set tBody = targetTbl.DataBodyRange
        On Error GoTo HardFail

        If tBody Is Nothing Then
            tableFailed = True
            ErrString = ErrString & vbNewLine & wsName & " :: " & tblName & ": Target DataBodyRange missing after resize."
            If Not failTables.Exists(tableKey) Then failTables.Add tableKey, "DataBodyRange missing"
            GoTo AfterCopy
        End If

        segStart = 0
        For j = 1 To tgtCols + 1

            If j > tgtCols Then
                isBreak = True
            Else
                isBreak = calcCol(j)
            End If

            If segStart = 0 Then
                If j <= tgtCols Then
                    If Not calcCol(j) Then segStart = j
                End If
            Else
                If isBreak Then
                    segEnd = j - 1
                    segW = segEnd - segStart + 1

                    If WhatIf Then
                        Debug.Print "WhatIf: " & wsName & " :: " & tblName & _
                                    " -> Would write TARGET cols " & segStart & " to " & segEnd & " in one batch."
                    Else
                        ReDim outArr(1 To srcRows, 1 To segW)

                        For rr = 1 To srcRows
                            cc = 1
                            For colPtr = segStart To segEnd
                                If mapSrcCol(colPtr) > 0 Then
                                    outArr(rr, cc) = srcData(rr, mapSrcCol(colPtr))
                                Else
                                    outArr(rr, cc) = vbNullString
                                End If
                                cc = cc + 1
                            Next colPtr
                        Next rr

                        tBody.Columns(segStart).Resize(srcRows, segW).Value2 = outArr
                    End If

                    segStart = 0
                End If
            End If
        Next j

        If Not tableFailed Then
            If Not okTables.Exists(tableKey) Then okTables.Add tableKey, True
        Else
            If Not failTables.Exists(tableKey) Then failTables.Add tableKey, "Failed (table-level issue)"
        End If

AfterCopy:
        On Error Resume Next
        If Not targetTbl Is Nothing Then
            If targetTbl.ShowAutoFilter Then
                If targetTbl.AutoFilter.FilterMode Then targetTbl.AutoFilter.ShowAllData
            End If
        End If

        If guardActive Then
            SheetGuard_End tWs, reprotect, pwd
            guardActive = False
        End If
        On Error GoTo HardFail

NextRow:
        If guardActive Then
            On Error Resume Next
            SheetGuard_End tWs, reprotect, pwd
            guardActive = False
            On Error GoTo HardFail
        End If
    Next r

Finalize:
    On Error Resume Next

    If WhatIf Then
        Debug.Print "WhatIf: Would CalculateFull + lock/protect routines."
    Else
        If Not hadHardFail Then
            Application.CalculateFull
            Relock_All_Tables passwd
            ProtectAllSheets passwd
        End If
    End If

    On Error GoTo 0

    okList = vbNullString
    For Each k In okTables.keys
        okList = okList & vbNewLine & Replace$(CStr(k), "|", " :: ")
    Next k
    If Len(okList) = 0 Then okList = vbNewLine & "(none)"

    failList = vbNullString
    For Each k In failTables.keys
        failList = failList & vbNewLine & Replace$(CStr(k), "|", " :: ") & "  -->  " & CStr(failTables(k))
    Next k
    If Len(failList) = 0 Then failList = vbNewLine & "(none)"

    summaryMsg = "Tables updated:" & okList & vbNewLine & vbNewLine & _
                 "Tables NOT updated:" & failList

    errMsg = "Errors (detail):" & vbNewLine & IIf(Len(ErrString) = 0, "(none)", ErrString)

    If WhatIf Then
        Debug.Print "DRY-RUN complete."
        Debug.Print summaryMsg
        If Len(ErrString) > 0 Then Debug.Print errMsg
    Else
        MsgBox summaryMsg, vbInformation, "Tracking Update Result"
        If Len(ErrString) > 0 Then MsgBox errMsg, vbExclamation, "Tracking Update Errors"
    End If

CleanExit:
    On Error Resume Next
    If guardActive Then SheetGuard_End tWs, reprotect, pwd
    AppGuard_End
    Exit Sub

HardFail:
    errNum = Err.Number
    errDesc = Err.description
    hadHardFail = True

    On Error Resume Next
    If guardActive Then
        SheetGuard_End tWs, reprotect, pwd
        guardActive = False
    End If
    LogError procName, errNum, errDesc
    On Error GoTo 0

    ErrString = ErrString & vbNewLine & "Hard fail (" & procName & "): " & errNum & " - " & errDesc
    Resume Finalize
End Sub



'-----------------------------
' Helpers (module-local)
'-----------------------------
Private Function myCoerceBool(ByVal v As Variant) As Boolean
    ' Accepts TRUE/FALSE, 1/0, "Yes"/"No", "Y"/"N", and numeric text.
    Dim s As String

    If IsError(v) Then
        myCoerceBool = False
        Exit Function
    End If

    Select Case VarType(v)
        Case vbBoolean
            myCoerceBool = CBool(v)
            Exit Function
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            myCoerceBool = (CDbl(v) <> 0)
            Exit Function
    End Select

    s = UCase$(Trim$(CStr(v)))
    If Len(s) = 0 Then
        myCoerceBool = False
    ElseIf s = "TRUE" Or s = "YES" Or s = "Y" Then
        myCoerceBool = True
    ElseIf s = "FALSE" Or s = "NO" Or s = "N" Then
        myCoerceBool = False
    Else
        myCoerceBool = (val(s) <> 0)
    End If
End Function

Private Function myEnsure2D(ByVal v As Variant) As Variant
    ' Normalises Excel range reads so downstream code can treat it as (1..r, 1..c).
    ' - 1 cell => returns (1..1,1..1)
    ' - 1 row or 1 col range from Excel is already 2D, but we keep this for safety.
    Dim tmp(1 To 1, 1 To 1) As Variant

    If IsArray(v) Then
        myEnsure2D = v
    Else
        tmp(1, 1) = v
        myEnsure2D = tmp
    End If
End Function

Private Sub HandleMissingSource( _
    ByRef missingPolicy As Long, _
    ByVal WhatIf As Boolean, _
    ByVal tWs As Worksheet, _
    ByVal targetTbl As ListObject, _
    ByVal passwd As String, _
    ByVal promptBody As String, _
    ByRef didClear As Boolean, _
    ByRef firstResp As VbMsgBoxResult, _
    ByRef perResp As VbMsgBoxResult)

    ' Centralised missing-source handling.
    ' Guarantees:
    '   - Prompts only when WhatIf = False
    '   - Sets firstResp/perResp deterministically so caller can reliably detect Cancel
    '   - Respects missingPolicy:
    '       -1 = not chosen yet
    '        1 = clear all missing-source targets (no per-table prompt)
    '        2 = prompt per missing table

    Dim reprotect As TSheetGuardState

    ' Sentinel defaults so caller can distinguish "not asked" from a real response.
    didClear = False
    firstResp = 0
    perResp = 0

    Debug.Print "ENTER HandleMissingSource | WhatIf=" & WhatIf & " | missingPolicy=" & missingPolicy & _
                " | TargetSheet=" & tWs.Name & " | TargetTable=" & targetTbl.Name

    If WhatIf Then
        ' Dry-run: no prompts, no clears; caller should log intent.
        Debug.Print "WhatIf: " & Replace$(promptBody, vbNewLine, " ")
        Exit Sub
    End If

    If missingPolicy = -1 Then
        firstResp = MsgBox( _
            "One or more SOURCE sheets/tables are missing." & vbNewLine & vbNewLine & _
            "YES    = Clear ALL corresponding TARGET tables automatically" & vbNewLine & _
            "NO     = Decide case-by-case (prompt for each missing table)" & vbNewLine & _
            "CANCEL = Abort update", _
            vbYesNoCancel + vbQuestion, _
            "Missing SOURCE - Action")

        If firstResp = vbCancel Then Exit Sub
        If firstResp = vbYes Then
            missingPolicy = 1
        Else
            missingPolicy = 2
        End If
    End If

    If missingPolicy = 1 Then
        reprotect = SheetGuard_Begin(tWs, passwd)
        ClearTableToHeaderOnly targetTbl
        SheetGuard_End tWs, reprotect, pwd
        didClear = True
        Exit Sub
    End If

    ' missingPolicy = 2: ask per occurrence
    perResp = MsgBox(promptBody & vbNewLine & vbNewLine & "Clear TARGET table?", _
                     vbYesNoCancel + vbQuestion, _
                     "Missing SOURCE")

    If perResp = vbCancel Then Exit Sub

    If perResp = vbYes Then
        reprotect = SheetGuard_Begin(tWs, passwd)
        ClearTableToHeaderOnly targetTbl
        SheetGuard_End tWs, reprotect, pwd
        didClear = True
    Else
        didClear = False
    End If
End Sub
'====================================================================================
' Tracking Workbook Variant
'------------------------------------------------------------------------------------
' Updates tables in ThisWorkbook from a selected open workbook using driver table tbl_Tables.
'
' Assumptions (based on your notes):
'   - Tracking workbook has the same Intro sheet (copies B2:B6).
'   - Driver table is a ListObject named "tbl_Tables"
'   - tbl_Tables contains at least these headers:
'         "Worksheet"   (sheet name)
'         "Table"       (ListObject name)
'         "Update"      (TRUE/FALSE)  -> only rows with TRUE are processed
'   - Copy is header-mapped (normalized header names).
'   - Skip calculated/formula columns in TARGET (same behaviour as your original).
'   - Source table with 0 rows => target is cleared to header-only and message logged.
'
' Backtest support:
'   - Pass WhatIf:=True to do a dry-run (no writes) while printing actions to Immediate window.
'====================================================================================





'====================================================================================
' Helper: Select an open workbook (excluding ThisWorkbook)
'====================================================================================
Private Function SelectOpenWorkbook(ByVal excludeWbk As Workbook) As Workbook
    Dim wb As Workbook
    Dim list As String, i As Long
    Dim names() As String
    Dim pick As Variant

    ReDim names(1 To Application.Workbooks.Count)
    i = 0

    For Each wb In Application.Workbooks
        If Not wb Is excludeWbk Then
            i = i + 1
            names(i) = wb.Name
            list = list & vbNewLine & i & ") " & wb.Name
        End If
    Next wb

    If i = 0 Then
        MsgBox "No other workbooks are open.", vbExclamation
        Exit Function
    End If

    pick = Application.InputBox( _
        Prompt:="Select SOURCE workbook number:" & vbNewLine & list, _
        Title:="Select Source Workbook", _
        Type:=1)

    If pick = False Then Exit Function
    If CLng(pick) < 1 Or CLng(pick) > i Then
        MsgBox "Invalid selection.", vbExclamation
        Exit Function
    End If

    Set SelectOpenWorkbook = Application.Workbooks(names(CLng(pick)))
End Function



'====================================================================================
' Helper: Get ListObject column index by header name (normalized match)
'====================================================================================
Public Function ListObject_ColIndex(ByVal lo As ListObject, ByVal headerName As String) As Long
    Dim i As Long, tgt As String
    tgt = NormalizeHeader(headerName)

    For i = 1 To lo.ListColumns.Count
        If NormalizeHeader(lo.ListColumns(i).Name) = tgt Then
            ListObject_ColIndex = i
            Exit Function
        End If
    Next i
End Function


'====================================================================================
' Helper: Detect calculated/formula column (skip copying into these)
'====================================================================================
Public Function IsCalculatedColumnFast(ByVal lo As ListObject, ByVal colIndex As Long) As Boolean
    On Error GoTo ErrHandler

    Dim rng As Range
    Dim formulaCells As Range
    Dim nonBlankCount As Long
    Dim formulaCount As Long

    IsCalculatedColumnFast = False

    If lo Is Nothing Then Exit Function
    If colIndex < 1 Or colIndex > lo.ListColumns.Count Then Exit Function
    If lo.ListRows.Count = 0 Then Exit Function

    Set rng = lo.ListColumns(colIndex).DataBodyRange
    If rng Is Nothing Then Exit Function

    nonBlankCount = Application.WorksheetFunction.CountA(rng)
    If nonBlankCount = 0 Then Exit Function

    On Error Resume Next
    Set formulaCells = rng.SpecialCells(xlCellTypeFormulas)
    On Error GoTo ErrHandler

    If formulaCells Is Nothing Then
        formulaCount = 0
    Else
        formulaCount = formulaCells.Cells.Count
    End If

    IsCalculatedColumnFast = (formulaCount > 0 And formulaCount = nonBlankCount)

    Exit Function

ErrHandler:
    Err.Raise Err.Number, "IsCalculatedColumnFast(colIndex=" & colIndex & ", table=" & lo.Name & ")", Err.description
End Function

'====================================================================================
' Helper: Normalize header text for mapping
'====================================================================================
Public Function NormalizeHeader(ByVal v As Variant) As String
    Dim s As String
    s = CStr(Nz(v, vbNullString))

    ' Trim + remove NBSP
    s = Replace(s, ChrW$(160), " ")
    s = Trim$(s)

    ' Collapse multiple spaces
    Do While InStr(1, s, "  ", vbBinaryCompare) > 0
        s = Replace(s, "  ", " ")
    Loop

    NormalizeHeader = LCase$(s)
End Function

'====================================================================================
' Helper: Pivot filter clearing (best-effort)
'====================================================================================
Public Sub ClearAllPivotFilters()
    Dim ws As Worksheet
    Dim pt As PivotTable

    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.ClearAllFilters
        Next pt
    Next ws
End Sub


'====================================================================================
' Helper: Null/Empty -> default
'====================================================================================
Public Function Nz(ByVal v As Variant, ByVal defaultValue As Variant) As Variant
    If IsError(v) Then
        Nz = defaultValue
    ElseIf IsEmpty(v) Then
        Nz = defaultValue
    ElseIf VarType(v) = vbString And Len(v) = 0 Then
        Nz = defaultValue
    Else
        Nz = v
    End If
End Function





'====================================================================================
' ProtectAllSheets
'
' Purpose:
'   Re-protect every worksheet in ThisWorkbook using consistent protection rules.
'
' Features:
'   - Safe reset (unprotect then protect)
'   - UserInterfaceOnly:=True (VBA can still write)
'   - Allows filtering / sorting / pivots
'   - Optional password support
'   - Handles chart sheets safely
'
' Recommended:
'   Call AFTER Relock_All_Tables
'====================================================================================
Public Sub ProtectAllSheets(Optional ByVal pwd As String = vbNullString)

    Dim ws As Worksheet
    Dim ch As Chart

    On Error GoTo CleanExit

    Application.ScreenUpdating = False

    '------------------------------------------
    ' Protect all Worksheets
    '------------------------------------------
    For Each ws In ThisWorkbook.Worksheets

        On Error Resume Next
        ws.Unprotect password:=pwd
        On Error GoTo 0

        ws.Protect password:=pwd, _
                   UserInterfaceOnly:=True, _
                   AllowFiltering:=True, _
                   AllowSorting:=True, _
                   AllowUsingPivotTables:=True, _
                   AllowFormattingCells:=False, _
                   AllowFormattingColumns:=False, _
                   AllowFormattingRows:=False, _
                   AllowInsertingRows:=False, _
                   AllowDeletingRows:=False, _
                   AllowInsertingColumns:=False, _
                   AllowDeletingColumns:=False

    Next ws

    '------------------------------------------
    ' Protect all Chart Sheets (if any)
    '------------------------------------------
    For Each ch In ThisWorkbook.Charts
        On Error Resume Next
        ch.Unprotect password:=pwd
        ch.Protect password:=pwd
        On Error GoTo 0
    Next ch

CleanExit:
    Application.ScreenUpdating = True

End Sub


'====================================================================================
' Relock_All_Tables  (uses modGuardsAndTables)
'
' Purpose:
'   Re-applies locking logic across all tables in ThisWorkbook, then PROTECTS ALL sheets.
'
' Rules:
'   - Table header row LOCKED (prevents editing column names)
'   - Calculated columns locked
'   - Any cell containing a formula locked
'   - Plain value-entry cells unlocked
'   - Any cell that was already locked BEFORE this macro runs stays locked (never unlock locked cells)
'   - Sheet protected with UserInterfaceOnly:=True
'   - Outlining enabled after protection (group expand/collapse works)
'
' Dependencies:
'   - modGuardsAndTables: AppGuard_Begin / AppGuard_End
'   - IsCalculatedColumnFast(lo, colIndex) must exist (your existing helper)
'   - Public Const pwd As String = "Walz25!" (or similar) must exist
'====================================================================================
Public Sub Relock_All_Tables(Optional ByVal passwd As String = vbNullString)

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim col As ListColumn
    Dim rng As Range, rngFormulas As Range

    Dim errLog As String
    Dim ctx As String

    On Error GoTo FatalFail

    If Len(passwd) = 0 Then passwd = pwd

    '---------------------------------------------
    ' Use your guard module for app state handling
    '---------------------------------------------
    AppGuard_Begin True, "Relock_All_Tables..."

    For Each ws In ThisWorkbook.Worksheets

        ctx = "Sheet: " & ws.Name

        '-------------------------------------------------
        ' Reset protection state (log failure, continue)
        '-------------------------------------------------
        On Error Resume Next
        ws.Unprotect password:=passwd
        If Err.Number <> 0 Then
            errLog = errLog & vbCrLf & "Unprotect failed | " & ctx & " | " & Err.Number & " - " & Err.description
            Err.Clear
        End If
        On Error GoTo FatalFail

        '-------------------------------------------------
        ' Process each table
        '-------------------------------------------------
        For Each lo In ws.ListObjects

            ctx = "Sheet: " & ws.Name & " | Table: " & lo.Name

            ' LOCK header row (prevents editing column names)
            On Error Resume Next
            If Not lo.HeaderRowRange Is Nothing Then lo.HeaderRowRange.Locked = True
            If Err.Number <> 0 Then
                errLog = errLog & vbCrLf & "Header lock failed | " & ctx & " | " & Err.Number & " - " & Err.description
                Err.Clear
            End If
            On Error GoTo FatalFail

            ' Skip if no data rows
            If lo.ListRows.Count = 0 Then GoTo NextTable

            For Each col In lo.ListColumns

                ctx = "Sheet: " & ws.Name & " | Table: " & lo.Name & " | Col: " & col.Name

                On Error Resume Next
                Set rng = col.DataBodyRange
                If Err.Number <> 0 Or rng Is Nothing Then
                    errLog = errLog & vbCrLf & "DataBodyRange failed | " & ctx & " | " & Err.Number & " - " & Err.description
                    Err.Clear
                    Set rng = Nothing
                    On Error GoTo FatalFail
                    GoTo NextColumn
                End If
                On Error GoTo FatalFail

                '-----------------------------------------
                ' Preserve any pre-existing locked cells
                '-----------------------------------------
                Dim prevLocked As Variant
                prevLocked = rng.Locked ' Boolean for 1-cell, 2D variant for multi-cell

                '-----------------------------------------
                ' Apply lock rules
                '-----------------------------------------
                If IsCalculatedColumnFast(lo, col.Index) Then
                    rng.Locked = True
                Else
                    ' Default unlock (we will restore previous locked cells afterward)
                    rng.Locked = False

                    ' Lock formula cells only (fast)
                    Set rngFormulas = Nothing
                    On Error Resume Next
                    Set rngFormulas = rng.SpecialCells(xlCellTypeFormulas)
                    If Err.Number <> 0 Then
                        ' Normal if there are no formulas
                        Err.Clear
                    ElseIf Not rngFormulas Is Nothing Then
                        rngFormulas.Locked = True
                    End If
                    On Error GoTo FatalFail
                End If

                '-----------------------------------------
                ' Restore any cells that were locked before
                ' (never unlock originally locked cells)
                '-----------------------------------------
                RestorePreviouslyLocked rng, prevLocked

NextColumn:
                Set rng = Nothing
                Set rngFormulas = Nothing

            Next col

NextTable:
        Next lo

        '-------------------------------------------------
        ' Protect ALL sheets (enforced governance)
        '-------------------------------------------------
        ctx = "Sheet: " & ws.Name

        On Error Resume Next
        ws.Protect password:=passwd, _
                   UserInterfaceOnly:=True, _
                   AllowFiltering:=True, _
                   AllowSorting:=True, _
                   AllowUsingPivotTables:=True, _
                   AllowFormattingCells:=False, _
                   AllowFormattingColumns:=True, _
                   AllowFormattingRows:=True, _
                   AllowInsertingRows:=False, _
                   AllowDeletingRows:=False, _
                   AllowInsertingColumns:=False, _
                   AllowDeletingColumns:=False

        If Err.Number <> 0 Then
            errLog = errLog & vbCrLf & "Protect failed | " & ctx & " | " & Err.Number & " - " & Err.description
            Err.Clear
        End If

        ' Enable outlining so users can expand/collapse grouping after protection
        ws.EnableOutlining = True
        If Err.Number <> 0 Then
            errLog = errLog & vbCrLf & "EnableOutlining failed | " & ctx & " | " & Err.Number & " - " & Err.description
            Err.Clear
        End If
        On Error GoTo FatalFail

    Next ws

CleanExit:
    AppGuard_End

    If Len(errLog) > 0 Then
        Err.Raise vbObjectError + 701, "Relock_All_Tables", "Completed with issues:" & vbCrLf & errLog
    End If

    Exit Sub

FatalFail:
    On Error Resume Next
    AppGuard_End
    On Error GoTo 0

    Err.Raise Err.Number, "Relock_All_Tables", "Fatal error at [" & ctx & "]: " & Err.description

End Sub

'====================================================================================
' RestorePreviouslyLocked
' - Ensures any cell that was locked before the macro ran stays locked.
' - Handles both single-cell (Boolean) and multi-cell (2D Variant) Locked captures.
'====================================================================================
Private Sub RestorePreviouslyLocked(ByVal rng As Range, ByVal prevLocked As Variant)

    If rng Is Nothing Then Exit Sub

    If IsVariant2DArray(prevLocked) Then
        Dim r As Long, c As Long
        For r = 1 To UBound(prevLocked, 1)
            For c = 1 To UBound(prevLocked, 2)
                If prevLocked(r, c) = True Then
                    rng.Cells(r, c).Locked = True
                End If
            Next c
        Next r
    Else
        If CBool(prevLocked) = True Then rng.Locked = True
    End If

End Sub

'====================================================================================
' IsVariant2DArray
' - True if v is a 2D array (e.g. range properties over multi-cell ranges)
'====================================================================================
Private Function IsVariant2DArray(ByVal v As Variant) As Boolean
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


Sub Relock()
    Relock_All_Tables
End Sub
