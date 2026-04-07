Attribute VB_Name = "modLoadChecksheet"
Option Explicit
'================================================================================
' SECTION 8 — LOAD CHECKSHEET LINE-AREA HELPERS (NAMED RANGE DRIVEN)
'================================================================================


'=========================================================
' Populate FRM 496 (Load Checksheet) from tbl_DD
'
' - Reads the current docket number from the form
' - Finds the matching row in tbl_DD
' - Writes matching tbl_DD values into the form value column
'   where form label text in column A matches the tbl_DD header
' - Explicitly fills Total Load Weight: using tracking helper
' - Validates Transport Type so downstream tracking load uses
'   the correct docket field
'=========================================================
Public Sub FRM496_FillFrom_tbl_DD_ByDocket()

    Const TBL_NAME As String = "tbl_DD"
    Const FORM_SHEET As String = "Load Checksheet"

    Const LABEL_COL As Long = 1      'A
    Const VALUE_COL As Long = 3      'C

    Const DOCKET_HDR As String = "Delivery Docket Number:"
    Const DOCKET_LABEL As String = "Delivery Docket Number:"
    Const TYPE_HDR As String = "Transport Type"
    Const TOTAL_WT_LABEL As String = "Total Load Weight:"
    Const WALZ_CONTACT_HDR As String = "Walz Contact:"
    Const WALZ_CONTACT_LABEL As String = "Walz Contact:"

    Dim wsForm As Worksheet
    Dim lo As ListObject

    Dim docketNo As String
    Dim dockKey As String
    Dim rowIdx As Long

    Dim formMap As Object            ' label -> row number on form
    Dim hdrMap As Object             ' cleaned header -> column number in tbl_DD

    Dim lastFormRow As Long
    Dim r As Long
    Dim lbl As String
    Dim hdr As String
    Dim colIdx As Long

    Dim ddTypeCol As Long
    Dim transportType As String

    Dim totalLoadWt As Double

    On Error GoTo CleanFail

    Set wsForm = ThisWorkbook.Worksheets(FORM_SHEET)
    Set lo = GetTableByName(TBL_NAME)

    If lo Is Nothing Then
        MsgBox "Table [" & TBL_NAME & "] not found.", vbExclamation
        Exit Sub
    End If

    If Not TableHasData(lo) Then
        MsgBox TBL_NAME & " has no data rows.", vbExclamation
        Exit Sub
    End If

    '---------------------------------------------
    ' Get current docket from the form
    '---------------------------------------------
    docketNo = GetFormDocketNumberFromLabel(wsForm, DOCKET_LABEL, LABEL_COL, VALUE_COL)
    dockKey = CleanKey(docketNo)

    If Len(dockKey) = 0 Then
        MsgBox "Delivery Docket Number is blank on the form.", vbExclamation
        Exit Sub
    End If

    '---------------------------------------------
    ' Find matching row in tbl_DD
    '---------------------------------------------
    rowIdx = FindTableRowIndexByValue_Clean(lo, DOCKET_HDR, dockKey)
    If rowIdx = 0 Then
        MsgBox "Delivery Docket Number [" & docketNo & "] not found in " & TBL_NAME & ".", vbExclamation
        Exit Sub
    End If

    '---------------------------------------------
    ' Validate transport type
    '---------------------------------------------
    ddTypeCol = TableColIndexByHeaderClean(lo, TYPE_HDR)
    If ddTypeCol = 0 Then
        MsgBox TBL_NAME & " is missing required column [" & TYPE_HDR & "].", vbExclamation
        Exit Sub
    End If

    transportType = NzText(lo.DataBodyRange.Cells(rowIdx, ddTypeCol).Value2)

    If Len(GetTrackingDocketHeaderByTransportType(transportType)) = 0 Then
        MsgBox "Invalid Transport Type for docket [" & docketNo & "] in " & TBL_NAME & ": [" & transportType & "]." & vbCrLf & _
               "Valid values are: Subcon, TPP, Site.", vbExclamation
        Exit Sub
    End If

    '---------------------------------------------
    ' Build form label map (column A label text -> row)
    '---------------------------------------------
    Set formMap = CreateObject("Scripting.Dictionary")
    formMap.CompareMode = 1 'TextCompare

    lastFormRow = wsForm.Cells(wsForm.Rows.Count, LABEL_COL).End(xlUp).Row

    For r = 1 To lastFormRow
        lbl = Trim$(Replace(CStr(wsForm.Cells(r, LABEL_COL).Value2), vbLf, " "))
        If Len(lbl) > 0 Then
            If Not formMap.Exists(CleanKey(lbl)) Then
                formMap.Add CleanKey(lbl), r
            End If
        End If
    Next r

    '---------------------------------------------
    ' Build tbl_DD header map (cleaned header -> column index)
    '---------------------------------------------
    Set hdrMap = CreateObject("Scripting.Dictionary")
    hdrMap.CompareMode = 1 'TextCompare

    For colIdx = 1 To lo.ListColumns.Count
        hdr = Trim$(Replace(CStr(lo.ListColumns(colIdx).name), vbLf, " "))
        If Len(hdr) > 0 Then
            If Not hdrMap.Exists(CleanKey(hdr)) Then
                hdrMap.Add CleanKey(hdr), colIdx
            End If
        End If
    Next colIdx

    '---------------------------------------------
    ' Push matching tbl_DD headers into form labels
    ' Rule: if form label text matches a tbl_DD header text,
    ' write that value into column C on the same row.
    '---------------------------------------------
    Dim k As Variant
    For Each k In formMap.keys
        If hdrMap.Exists(k) Then
            wsForm.Cells(formMap(k), VALUE_COL).Value = lo.DataBodyRange.Cells(rowIdx, hdrMap(k)).Value2
        End If
    Next k

    '---------------------------------------------
    ' Explicitly force docket field back in, just to be safe
    '---------------------------------------------
    SetFormValueByLabel wsForm, DOCKET_LABEL, docketNo, LABEL_COL, VALUE_COL

    '---------------------------------------------
    ' Explicit Walz Contact fill
    '---------------------------------------------
    If hdrMap.Exists(CleanKey(WALZ_CONTACT_HDR)) And formMap.Exists(CleanKey(WALZ_CONTACT_LABEL)) Then
        wsForm.Cells(formMap(CleanKey(WALZ_CONTACT_LABEL)), VALUE_COL).Value = _
            lo.DataBodyRange.Cells(rowIdx, hdrMap(CleanKey(WALZ_CONTACT_HDR))).Value2
    End If

    '---------------------------------------------
    ' Fill Total Load Weight from tracking helper
    '---------------------------------------------
    totalLoadWt = GetTotalLoadWeightForDocket(docketNo)
    SetFormValueByLabel wsForm, TOTAL_WT_LABEL, FormatKg(totalLoadWt), LABEL_COL, VALUE_COL

    Exit Sub

CleanFail:
    MsgBox "FRM496_FillFrom_tbl_DD_ByDocket failed:" & vbCrLf & Err.Description, vbExclamation

End Sub

'=========================================================
' Pull matching tbl_Tracking rows for current docket,
' determine Transport Type from tbl_DD first,
' then use the correct tbl_Tracking docket column:
'
'   Subcon -> Load Sheet No. to Subcontractor
'   TPP    -> Load Sheet No. to TPP
'   Site   -> Delivery Docket #
'
' Line outputs:
'   Qty
'   Asset Number
'   Description/Tag Number
'   Line Weight = Assembly Quantity * Load Weight each
'   Transport Dimensions
'=========================================================
Public Sub FRM496_LoadLines_FromTracking()

    Const FORM_SHEET As String = "Load Checksheet"
    Const LABEL_COL As Long = 1
    Const VALUE_COL As Long = 3
    Const DOCKET_LABEL As String = "Delivery Docket Number:"

    Const DD_TBL As String = "tbl_DD"
    Const DD_DOCKET_HDR As String = "Delivery Docket Number:"
    Const DD_TYPE_HDR As String = "Transport Type"

    Const TRACK_TBL As String = "tbl_Tracking"
    Const HDR_QTY As String = "Assembly Quantity"
    Const HDR_ASSET As String = "Asset Number"
    Const HDR_DESC As String = "Description/Tag Number"
    Const HDR_WT_EACH As String = "Load Weight each"
    Const HDR_DIMS As String = "Transport Dimensions"

    Const RNG_HDR As String = "rng_LC_Header"
    Const RNG_BOT As String = "rng_LC_Bottom"

    Dim wsForm As Worksheet
    Dim docketNo As String
    Dim dockKey As String

    Dim loDD As ListObject
    Dim loT As ListObject

    Dim ddRowIdx As Long
    Dim ddTypeCol As Long
    Dim transportType As String
    Dim trackDocketHdr As String

    Dim a As Variant
    Dim r As Long
    Dim n As Long

    Dim colDock As Long
    Dim colQty As Long
    Dim colAsset As Long
    Dim colDesc As Long
    Dim colWtEach As Long
    Dim colDims As Long

    Dim qtyVal As Double
    Dim wtEachVal As Double
    Dim lineWt As Double

    Dim items As Object
    Dim itemRec As Variant
    Dim neededRows As Long

    Dim sgState As TSheetGuardState
    Dim guardStarted As Boolean

    On Error GoTo CleanFail

    Set wsForm = ThisWorkbook.Worksheets(FORM_SHEET)

    ' Correct call style for the existing guard function:
    sgState = SheetGuard_Begin(wsForm)
    guardStarted = True

    '---------------------------------------------
    ' Get docket number from form
    '---------------------------------------------
    docketNo = GetFormDocketNumberFromLabel(wsForm, DOCKET_LABEL, LABEL_COL, VALUE_COL)
    dockKey = CleanKey(docketNo)

    If Len(dockKey) = 0 Then
        MsgBox "Delivery Docket Number is blank on the form.", vbExclamation
        GoTo CleanExit
    End If

    '---------------------------------------------
    ' Step 1: Find docket in tbl_DD and get Transport Type
    '---------------------------------------------
    Set loDD = GetTableByName(DD_TBL)
    If loDD Is Nothing Then
        MsgBox "Table [" & DD_TBL & "] not found.", vbExclamation
        GoTo CleanExit
    End If

    If Not TableHasData(loDD) Then
        MsgBox DD_TBL & " has no data rows.", vbExclamation
        GoTo CleanExit
    End If

    ddRowIdx = FindTableRowIndexByValue_Clean(loDD, DD_DOCKET_HDR, dockKey)
    If ddRowIdx = 0 Then
        MsgBox "Delivery Docket Number [" & docketNo & "] not found in " & DD_TBL & ".", vbExclamation
        GoTo CleanExit
    End If

    ddTypeCol = TableColIndexByHeaderClean(loDD, DD_TYPE_HDR)
    If ddTypeCol = 0 Then
        MsgBox DD_TBL & " is missing required column [" & DD_TYPE_HDR & "].", vbExclamation
        GoTo CleanExit
    End If

    transportType = NzText(loDD.DataBodyRange.Cells(ddRowIdx, ddTypeCol).Value2)
    trackDocketHdr = GetTrackingDocketHeaderByTransportType(transportType)

    If Len(trackDocketHdr) = 0 Then
        MsgBox "Invalid Transport Type for docket [" & docketNo & "] in " & DD_TBL & ": [" & transportType & "]." & vbCrLf & _
               "Valid values are: Subcon, TPP, Site.", vbExclamation
        GoTo CleanExit
    End If

    '---------------------------------------------
    ' Step 2: Load tracking rows using mapped docket column
    '---------------------------------------------
    Set loT = GetTableByName(TRACK_TBL)
    If loT Is Nothing Then
        MsgBox "Table [" & TRACK_TBL & "] not found.", vbExclamation
        GoTo CleanExit
    End If

    If Not TableHasData(loT) Then
        MsgBox TRACK_TBL & " has no data rows.", vbExclamation
        GoTo CleanExit
    End If

    colDock = TableColIndexByHeaderClean(loT, trackDocketHdr)
    colQty = TableColIndexByHeaderClean(loT, HDR_QTY)
    colAsset = TableColIndexByHeaderClean(loT, HDR_ASSET)
    colDesc = TableColIndexByHeaderClean(loT, HDR_DESC)
    colWtEach = TableColIndexByHeaderClean(loT, HDR_WT_EACH)
    colDims = TableColIndexByHeaderClean(loT, HDR_DIMS)

    If colDock = 0 Or colQty = 0 Or colAsset = 0 Or colDesc = 0 Or colWtEach = 0 Or colDims = 0 Then
        MsgBox "Missing one or more required columns in " & TRACK_TBL & "." & vbCrLf & _
               "Need: [" & trackDocketHdr & "], [" & HDR_QTY & "], [" & HDR_ASSET & "], [" & HDR_DESC & "], [" & HDR_WT_EACH & "], [" & HDR_DIMS & "]", vbExclamation
        GoTo CleanExit
    End If

    a = loT.DataBodyRange.Value2
    n = UBound(a, 1)

    Set items = CreateObject("Scripting.Dictionary")
    items.CompareMode = 1 'TextCompare

    '---------------------------------------------
    ' Collect only matching rows
    '---------------------------------------------
    For r = 1 To n
        If StrComp(CleanKey(a(r, colDock)), dockKey, vbTextCompare) = 0 Then

            qtyVal = SafeCDbl(a(r, colQty))
            wtEachVal = SafeCDbl(a(r, colWtEach))
            lineWt = qtyVal * wtEachVal

            itemRec = Array( _
                a(r, colQty), _
                a(r, colAsset), _
                a(r, colDesc), _
                lineWt, _
                a(r, colDims) _
            )

            items.Add CLng(r), itemRec
        End If
    Next r

    neededRows = items.Count

    If neededRows = 0 Then
        ClearLoadChecksheetLines wsForm, RNG_HDR, RNG_BOT, True
        GoTo CleanExit
    End If

    EnsureRowsBetweenNamedRanges wsForm, RNG_HDR, RNG_BOT, neededRows
    WriteLoadChecksheetLines_Ordered wsForm, RNG_HDR, items
    AdjustRowHeights_MergedSafe wsForm, RNG_HDR, RNG_BOT, neededRows

CleanExit:
    On Error Resume Next
    If guardStarted Then
        SheetGuard_End wsForm, sgState
    End If
    On Error GoTo 0
    Exit Sub

CleanFail:
    On Error Resume Next
    If guardStarted Then
        SheetGuard_End wsForm, sgState
    End If
    On Error GoTo 0

    MsgBox "FRM496_LoadLines_FromTracking failed:" & vbCrLf & Err.Description, vbExclamation

End Sub
Public Function SafeCDbl(ByVal v As Variant) As Double
    On Error GoTo EH

    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If Len(Trim$(CStr(v))) = 0 Then Exit Function

    SafeCDbl = CDbl(v)
    Exit Function

EH:
    SafeCDbl = 0
End Function

'=========================================================
' Formats a weight for the form.
'=========================================================
Public Function FormatKg(ByVal v As Double) As String
    FormatKg = Format$(v, "#,##0.00") & " Kg"
End Function

'=========================================================
' Writes a value into the value column for a given form label.
' Looks for the label in LABEL_COL and writes to VALUE_COL.
'=========================================================
Public Sub SetFormValueByLabel(ByVal ws As Worksheet, _
                               ByVal labelText As String, _
                               ByVal outValue As Variant, _
                               Optional ByVal labelCol As Long = 1, _
                               Optional ByVal valueCol As Long = 3)

    Dim lastRow As Long
    Dim r As Long
    Dim cellText As String

    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, labelCol).End(xlUp).Row

    For r = 1 To lastRow
        cellText = Trim$(CStr(ws.Cells(r, labelCol).Value))
        If StrComp(cellText, Trim$(labelText), vbTextCompare) = 0 Then
            ws.Cells(r, valueCol).Value = outValue
            Exit Sub
        End If
    Next r

End Sub

'=========================================================
' Returns total load weight for a docket by summing:
'   Assembly Quantity * Load Weight each
' across matching tbl_Tracking rows for the docket.
'=========================================================
Public Function GetTotalLoadWeightForDocket(ByVal docketNo As String) As Double

    Const DD_TBL As String = "tbl_DD"
    Const DD_DOCKET_HDR As String = "Delivery Docket Number:"
    Const DD_TYPE_HDR As String = "Transport Type"

    Const TRACK_TBL As String = "tbl_Tracking"
    Const HDR_QTY As String = "Assembly Quantity"
    Const HDR_WT_EACH As String = "Load Weight each"

    Dim dockKey As String

    Dim loDD As ListObject
    Dim loT As ListObject

    Dim ddRowIdx As Long
    Dim ddTypeCol As Long
    Dim transportType As String
    Dim trackDocketHdr As String

    Dim colDock As Long
    Dim colQty As Long
    Dim colWtEach As Long

    Dim a As Variant
    Dim r As Long
    Dim n As Long

    Dim qtyVal As Double
    Dim wtEachVal As Double
    Dim totalWt As Double

    On Error GoTo Fail

    dockKey = CleanKey(docketNo)
    If Len(dockKey) = 0 Then Exit Function

    Set loDD = GetTableByName(DD_TBL)
    If loDD Is Nothing Then Exit Function
    If Not TableHasData(loDD) Then Exit Function

    ddRowIdx = FindTableRowIndexByValue_Clean(loDD, DD_DOCKET_HDR, dockKey)
    If ddRowIdx = 0 Then Exit Function

    ddTypeCol = TableColIndexByHeaderClean(loDD, DD_TYPE_HDR)
    If ddTypeCol = 0 Then Exit Function

    transportType = NzText(loDD.DataBodyRange.Cells(ddRowIdx, ddTypeCol).Value2)
    trackDocketHdr = GetTrackingDocketHeaderByTransportType(transportType)
    If Len(trackDocketHdr) = 0 Then Exit Function

    Set loT = GetTableByName(TRACK_TBL)
    If loT Is Nothing Then Exit Function
    If Not TableHasData(loT) Then Exit Function

    colDock = TableColIndexByHeaderClean(loT, trackDocketHdr)
    colQty = TableColIndexByHeaderClean(loT, HDR_QTY)
    colWtEach = TableColIndexByHeaderClean(loT, HDR_WT_EACH)

    If colDock = 0 Or colQty = 0 Or colWtEach = 0 Then Exit Function

    a = loT.DataBodyRange.Value2
    n = UBound(a, 1)

    totalWt = 0

    For r = 1 To n
        If StrComp(CleanKey(a(r, colDock)), dockKey, vbTextCompare) = 0 Then
            qtyVal = SafeCDbl(a(r, colQty))
            wtEachVal = SafeCDbl(a(r, colWtEach))
            totalWt = totalWt + (qtyVal * wtEachVal)
        End If
    Next r

    GetTotalLoadWeightForDocket = totalWt
    Exit Function

Fail:
    GetTotalLoadWeightForDocket = 0
End Function

'=========================================================
' Helper: map Transport Type -> tbl_Tracking docket column
'=========================================================
Private Function GetTrackingDocketHeaderByTransportType(ByVal transportType As String) As String

    Select Case UCase$(Trim$(transportType))
        Case "SUBCON"
            GetTrackingDocketHeaderByTransportType = "Load Sheet No. to Subcontractor"
        Case "TPP"
            GetTrackingDocketHeaderByTransportType = "Load Sheet No. to TPP"
        Case "SITE"
            GetTrackingDocketHeaderByTransportType = "Delivery Docket # "
        Case Else
            GetTrackingDocketHeaderByTransportType = vbNullString
    End Select

End Function


'=========================================================
' Helper: safe text conversion
'=========================================================
Private Function NzText(ByVal v As Variant) As String
    If IsError(v) Then
        NzText = vbNullString
    ElseIf IsNull(v) Or IsEmpty(v) Then
        NzText = vbNullString
    Else
        NzText = Trim$(CStr(v))
    End If
End Function

'------------------------------------------------------------
' EnsureRowsBetweenNamedRanges
' Ensures exactly neededRows exist between header marker and bottom marker.
' Inserts/deletes rows and copies template formatting/formulas (A:F only).
'------------------------------------------------------------
Public Sub EnsureRowsBetweenNamedRanges( _
    ByVal ws As Worksheet, _
    ByVal headerName As String, _
    ByVal bottomName As String, _
    ByVal neededRows As Long _
)
    Dim rHeader As Range, rBottom As Range
    Dim headerRow As Long, bottomRow As Long
    Dim dataStart As Long, dataEnd As Long
    Dim capacity As Long, delta As Long
    Dim templateRow As Long

    Const FIRST_COL As Long = 1   'A
    Const LAST_COL As Long = 6    'F

    Dim srcBand As Range, dstBand As Range
    Dim prevSU As Boolean, prevEE As Boolean
    Dim prevCalc As XlCalculation

    If ws Is Nothing Then Exit Sub

    Set rHeader = ws.Range(headerName)
    Set rBottom = ws.Range(bottomName)

    headerRow = rHeader.Row
    bottomRow = rBottom.Row

    If bottomRow <= headerRow + 1 Then Exit Sub
    If neededRows < 0 Then neededRows = 0

    dataStart = headerRow + 1
    dataEnd = bottomRow - 1
    capacity = dataEnd - dataStart + 1
    If capacity < 0 Then capacity = 0

    templateRow = dataStart

    prevSU = Application.ScreenUpdating
    prevEE = Application.EnableEvents
    prevCalc = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanFail

    If neededRows > capacity Then
        delta = neededRows - capacity

        ws.Rows(bottomRow).Resize(delta).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

        Set srcBand = ws.Range(ws.Cells(templateRow, FIRST_COL), ws.Cells(templateRow, LAST_COL))
        Set dstBand = ws.Range(ws.Cells(bottomRow, FIRST_COL), ws.Cells(bottomRow + delta - 1, LAST_COL))

        srcBand.Copy Destination:=dstBand

    ElseIf neededRows < capacity Then
        delta = capacity - neededRows
        If delta > 0 Then
            ws.Rows(dataStart + neededRows).Resize(delta).Delete
        End If
    End If

CleanExit:
    Application.CutCopyMode = False
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEE
    Application.ScreenUpdating = prevSU
    Exit Sub

CleanFail:
    Resume CleanExit
End Sub

'------------------------------------------------------------
' WriteLoadChecksheetLines
' Writes items into the line-area starting at header row + 1.
' items: Scripting.Dictionary; each item is a 1D array (1..4): qty, asset, desc, weight
'------------------------------------------------------------
Public Sub WriteLoadChecksheetLines(ByVal ws As Worksheet, ByVal headerName As String, ByVal items As Object)
    Dim rHeader As Range
    Dim startRow As Long, startCol As Long
    Dim i As Long, k As Variant, rec As Variant

    If ws Is Nothing Then Exit Sub
    If items Is Nothing Then Exit Sub
    If items.Count = 0 Then Exit Sub

    Set rHeader = ws.Range(headerName)
    startRow = rHeader.Row + 1
    startCol = rHeader.Column

    i = 0
    For Each k In items.keys
        i = i + 1
        rec = items(k)

        ws.Cells(startRow + i - 1, startCol + 0).Value = rec(1)
        ws.Cells(startRow + i - 1, startCol + 1).Value = rec(2)
        ws.Cells(startRow + i - 1, startCol + 2).Value = rec(3)
        ws.Cells(startRow + i - 1, startCol + 4).Value = rec(4)
    Next k
End Sub

'------------------------------------------------------------
' WriteLoadChecksheetLines_Ordered
'
' Writes form lines using actual header labels on the sheet,
' so merged/non-contiguous columns do not break the output.
'
' Record structure:
'   rec(0) = Qty
'   rec(1) = Equipment No / Asset Number
'   rec(2) = Description
'   rec(3) = Weight
'   rec(4) = Dimensions
'------------------------------------------------------------
Public Sub WriteLoadChecksheetLines_Ordered(ByVal ws As Worksheet, ByVal headerName As String, ByVal items As Object)

    Dim rHeader As Range
    Dim startRow As Long
    Dim headerRow As Long

    Dim colQty As Long
    Dim colEquip As Long
    Dim colDesc As Long
    Dim colWt As Long
    Dim colDims As Long

    Dim keys As Variant
    Dim i As Long, j As Long
    Dim tmp As Variant
    Dim rec As Variant

    If ws Is Nothing Then Exit Sub
    If items Is Nothing Then Exit Sub
    If items.Count = 0 Then Exit Sub

    Set rHeader = ws.Range(headerName)

    headerRow = rHeader.Row
    startRow = rHeader.Row + 1

    'Find actual form columns by label text
    colQty = FindHeaderColumnByText(ws, headerRow, "Qty:")
    colEquip = FindHeaderColumnByText(ws, headerRow, "Equipment No:")
    colDesc = FindHeaderColumnByText(ws, headerRow, "Description:")
    colWt = FindHeaderColumnByText(ws, headerRow, "Weight (kg):")
    colDims = FindHeaderColumnByText(ws, headerRow, "Dimensions")

    If colQty = 0 Or colEquip = 0 Or colDesc = 0 Or colWt = 0 Or colDims = 0 Then
        Err.Raise vbObjectError + 513, "WriteLoadChecksheetLines_Ordered", _
                  "Could not locate one or more form header columns." & vbCrLf & _
                  "Found => Qty:" & colQty & ", Equip:" & colEquip & ", Desc:" & colDesc & _
                  ", Weight:" & colWt & ", Dims:" & colDims
    End If

    keys = items.keys

    'Simple numeric sort on dictionary keys
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If CLng(keys(j)) < CLng(keys(i)) Then
                tmp = keys(i)
                keys(i) = keys(j)
                keys(j) = tmp
            End If
        Next j
    Next i

    'Write output rows
    For i = LBound(keys) To UBound(keys)

        rec = items(keys(i))

        ws.Cells(startRow + i, colQty).Value = rec(0)      'Qty
        ws.Cells(startRow + i, colEquip).Value = rec(1)    'Equipment No / Asset
        ws.Cells(startRow + i, colDesc).Value = rec(2)     'Description
        ws.Cells(startRow + i, colWt).Value = rec(3)       'Weight
        ws.Cells(startRow + i, colDims).Value = rec(4)     'Dimensions

    Next i

End Sub

'------------------------------------------------------------
' FindHeaderColumnByText
'
' Looks across a header row and returns the first worksheet
' column whose displayed header text contains the target text.
' Handles merged cells by reading the top-left cell of MergeArea.
'------------------------------------------------------------
Private Function FindHeaderColumnByText(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal findText As String) As Long

    Dim lastCol As Long
    Dim c As Long
    Dim txt As String

    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol

        If ws.Cells(headerRow, c).MergeCells Then
            txt = CStr(ws.Cells(headerRow, c).MergeArea.Cells(1, 1).Value2)
        Else
            txt = CStr(ws.Cells(headerRow, c).Value2)
        End If

        txt = Trim$(Replace(txt, vbLf, " "))
        If Len(txt) > 0 Then
            If InStr(1, txt, findText, vbTextCompare) > 0 Then
                FindHeaderColumnByText = c
                Exit Function
            End If
        End If

    Next c

End Function

'------------------------------------------------------------
' ClearLoadChecksheetLines
' Clears the line-items area between header marker and bottom marker.
' keepTemplateRow=True: keeps 1 formatted template row (header+1) and clears contents.
' keepTemplateRow=False: removes all rows between header and bottom.
'------------------------------------------------------------
Public Sub ClearLoadChecksheetLines( _
    ByVal ws As Worksheet, _
    ByVal headerName As String, _
    ByVal bottomName As String, _
    Optional ByVal keepTemplateRow As Boolean = True _
)
    Dim rHeader As Range, rBottom As Range
    Dim headerRow As Long, bottomRow As Long
    Dim dataStart As Long, dataEnd As Long
    Dim cap As Long

    If ws Is Nothing Then Exit Sub

    Set rHeader = ws.Range(headerName)
    Set rBottom = ws.Range(bottomName)

    headerRow = rHeader.Row
    bottomRow = rBottom.Row

    If bottomRow <= headerRow + 1 Then Exit Sub

    dataStart = headerRow + 1
    dataEnd = bottomRow - 1
    cap = dataEnd - dataStart + 1
    If cap <= 0 Then Exit Sub

    If keepTemplateRow Then
        If cap > 1 Then
            ws.Rows(dataStart + 1).Resize(cap - 1).Delete
        End If
        ws.Rows(dataStart).ClearContents
    Else
        ws.Rows(dataStart).Resize(cap).Delete
    End If
End Sub

Public Sub AdjustRowHeights_MergedSafe( _
    ByVal ws As Worksheet, _
    ByVal headerName As String, _
    ByVal bottomName As String, _
    ByVal neededRows As Long _
)
    Const COL_C_OFFSET As Long = 2   'A=0,B=1,C=2 relative to startCol
    Const MAX_HEIGHT As Double = 140#
    Const MIN_HEIGHT As Double = 15#
    Const EXTRA_PADDING As Double = 8#   'total extra room top/bottom
    Const MIN_LINES_FOR_PADDING As Long = 2

    Dim rHeader As Range, rBottom As Range
    Dim startRow As Long, startCol As Long
    Dim bottomRow As Long
    Dim lastDataRow As Long
    Dim r As Long

    Dim cCell As Range
    Dim targetCell As Range

    Dim tempWS As Worksheet
    Dim tempCell As Range
    Dim mergeCols As Long
    Dim i As Long
    Dim totalColWidth As Double
    Dim fittedHeight As Double
    Dim textVal As String
    Dim lineCountEstimate As Long

    If ws Is Nothing Then Exit Sub
    If neededRows <= 0 Then Exit Sub

    Set rHeader = ws.Range(headerName)
    Set rBottom = ws.Range(bottomName)

    startRow = rHeader.Row + 1
    startCol = rHeader.Column
    bottomRow = rBottom.Row

    lastDataRow = startRow + neededRows - 1
    If lastDataRow > bottomRow - 1 Then lastDataRow = bottomRow - 1
    If lastDataRow < startRow Then Exit Sub

    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set tempWS = GetOrCreateRowHeightTempSheet(ThisWorkbook)
    Set tempCell = tempWS.Range("A1")

    For r = startRow To lastDataRow

        Set cCell = ws.Cells(r, startCol + COL_C_OFFSET)

        If cCell.MergeCells Then
            Set targetCell = cCell.MergeArea
        Else
            Set targetCell = cCell
        End If

        textVal = CStr(targetCell.Cells(1, 1).Value2)

        tempWS.Cells.Clear

        With tempCell
            .Value = textVal
            .WrapText = True
            .NumberFormat = "@"

            .Font.name = targetCell.Cells(1, 1).Font.name
            .Font.Size = targetCell.Cells(1, 1).Font.Size
            .Font.Bold = targetCell.Cells(1, 1).Font.Bold
            .Font.Italic = targetCell.Cells(1, 1).Font.Italic

            .HorizontalAlignment = targetCell.Cells(1, 1).HorizontalAlignment
            .VerticalAlignment = xlTop
        End With

        If targetCell.MergeCells Then
            mergeCols = targetCell.Columns.Count
        Else
            mergeCols = 1
        End If

        totalColWidth = 0
        For i = 1 To mergeCols
            totalColWidth = totalColWidth + ws.Columns(targetCell.Column + i - 1).ColumnWidth
        Next i

        tempWS.Columns(1).ColumnWidth = totalColWidth

        tempWS.Rows(1).EntireRow.AutoFit
        fittedHeight = tempWS.Rows(1).RowHeight

        'Add breathing room for wrapped text
        lineCountEstimate = CountDisplayLines(textVal)
        If lineCountEstimate >= MIN_LINES_FOR_PADDING Then
            fittedHeight = fittedHeight + EXTRA_PADDING
        Else
            fittedHeight = fittedHeight + 2
        End If

        If fittedHeight < MIN_HEIGHT Then fittedHeight = MIN_HEIGHT
        If fittedHeight > MAX_HEIGHT Then fittedHeight = MAX_HEIGHT

        With targetCell
            .WrapText = True
            .VerticalAlignment = xlTop
        End With

        ws.Rows(r).RowHeight = fittedHeight

    Next r

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Err.Raise Err.Number, "AdjustRowHeights_MergedSafe", Err.Description
End Sub

Private Function GetOrCreateRowHeightTempSheet(ByVal wb As Workbook) As Worksheet
    Const TEMP_SHEET_NAME As String = "zz_RowHeightTemp"

    On Error Resume Next
    Set GetOrCreateRowHeightTempSheet = wb.Worksheets(TEMP_SHEET_NAME)
    On Error GoTo 0

    If GetOrCreateRowHeightTempSheet Is Nothing Then
        Set GetOrCreateRowHeightTempSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        
        With GetOrCreateRowHeightTempSheet
            .name = TEMP_SHEET_NAME
            .Visible = xlSheetVeryHidden
        End With
    End If

    '---------------------------------------------
    ' FORCE UNPROTECTED STATE (critical)
    '---------------------------------------------
    With GetOrCreateRowHeightTempSheet
        If .ProtectContents Then
            .Unprotect pwd  'add password if you use one
        End If
    End With

End Function

Private Function CountDisplayLines(ByVal s As String) As Long
    Dim arr() As String

    If Len(s) = 0 Then
        CountDisplayLines = 1
    Else
        arr = Split(s, vbLf)
        CountDisplayLines = UBound(arr) - LBound(arr) + 1
        If CountDisplayLines < 1 Then CountDisplayLines = 1
    End If
End Function



Public Sub AdjustRowHeights_MergedSafe_old( _
    ByVal ws As Worksheet, _
    ByVal headerName As String, _
    ByVal bottomName As String, _
    ByVal neededRows As Long _
)
    Const COL_C_OFFSET As Long = 2   'A=0,B=1,C=2 relative to startCol
    Const MAX_HEIGHT As Double = 60#
    Const MIN_HEIGHT As Double = 15#
    Const POINTS_PER_LINE As Double = 15#
    Const PIX_PER_CHARACTER As Double = 4 'Approx 7 pixels per character at Calibri 11

    Dim rHeader As Range, rBottom As Range
    Dim startRow As Long, startCol As Long
    Dim bottomRow As Long
    Dim lastDataRow As Long
    Dim r As Long
    Dim cCell As Range
    Dim targetCell As Range

    Dim textLen As Long
    Dim approxLines As Long
    Dim approxHeight As Double
    Dim totalWidth As Double

    If neededRows <= 0 Then Exit Sub
    If ws Is Nothing Then Exit Sub

    Set rHeader = ws.Range(headerName)
    Set rBottom = ws.Range(bottomName)

    startRow = rHeader.Row + 1
    startCol = rHeader.Column
    bottomRow = rBottom.Row

    lastDataRow = startRow + neededRows - 1
    If lastDataRow > bottomRow - 1 Then lastDataRow = bottomRow - 1
    If lastDataRow < startRow Then Exit Sub

    Application.ScreenUpdating = False

    For r = startRow To lastDataRow

        Set cCell = ws.Cells(r, startCol + COL_C_OFFSET)

        'If merged, use the entire merged area
        If cCell.MergeCells Then
            Set targetCell = cCell.MergeArea
        Else
            Set targetCell = cCell
        End If

        targetCell.WrapText = True

        textLen = Len(CStr(targetCell.Cells(1, 1).Value))

        'Use total merged width
        totalWidth = targetCell.Width

        If textLen = 0 Then
            approxLines = 1
        Else
            
            approxLines = Application.WorksheetFunction.RoundUp((textLen * PIX_PER_CHARACTER) / totalWidth, 0)
            If approxLines < 1 Then approxLines = 1
        End If

        approxHeight = approxLines * POINTS_PER_LINE

        If approxHeight < MIN_HEIGHT Then approxHeight = MIN_HEIGHT
        If approxHeight > MAX_HEIGHT Then approxHeight = MAX_HEIGHT

        ws.Rows(r).RowHeight = approxHeight + 2

    Next r

    Application.ScreenUpdating = True
End Sub


