Attribute VB_Name = "modTracking"
Sub Print_Tracking_Schedule()
'
' Hide blank rows

   ActiveSheet.ListObjects("tbl_tracking").Range.AutoFilter _
    Field:=3, Criteria1:="<>"
   
' Format_Page Macro
' Format page to print tracking schedule and save as .pdf file
' Unmerge client name

    Range("D1:F1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With

' Unmerge project name

    Range("D2:F2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With

' Unmerge project name

    Range("D3:F3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With

' Unmerge project name

    Range("D4:F4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    
 ' Hide columns E
    
'    Columns("E:E").Select (command not needed removed by S.Hooper)
'    Range("E5").Activate (command needs to be line 6 now due to new line added ref next code line. changed by S.Hooper)
    Range("E6").Activate
    Selection.EntireColumn.Hidden = True
    
' Re-merge column D to F to show client name
    Range("D1:F4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        '.Merge Across:=True
    End With

'' Re-merge column D to F to show project name
'
'    Range("D2:F2").Select
'    With Selection
'        .MergeCells = True
'        .WrapText = False
'        .Orientation = 0
'        .AddIndent = False
'        .ShrinkToFit = False
'        .HorizontalAlignment = xlCenter
'        .VerticalAlignment = xlCenter
'    End With
'
' ' Re-merge column D to F to show fab job number
'
'    Range("D3:F3").Select
'    With Selection
'        .MergeCells = True
'        .WrapText = False
'        .Orientation = 0
'        .AddIndent = False
'        .ShrinkToFit = False
'        .HorizontalAlignment = xlCenter
'        .VerticalAlignment = xlCenter
'    End With
'
'' Re-merge column D to F to show paint job number
'
'    Range("D4:F4").Select
'    With Selection
'        .MergeCells = True
'        .WrapText = False
'        .Orientation = 0
'        .AddIndent = False
'        .ShrinkToFit = False
'        .HorizontalAlignment = xlCenter
'        .VerticalAlignment = xlCenter
'    End With

' Unmerge G1 to BH2 (Tracking Schedule)
    Range("G1:BH4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With

' Hide progress data columns
    Columns("K:K").Select
    Range("K3").Activate
    Selection.EntireColumn.Hidden = True
    Columns("N:AW").Select
    Range("N3").Activate
    Selection.EntireColumn.Hidden = True
    Columns("BA:BD").Select
    Range("BA3").Activate
    Selection.EntireColumn.Hidden = True
    Columns("BF:BH").Select
    Range("BF3").Activate
    Selection.EntireColumn.Hidden = True
    
' Re-merge G1 to BH2 (Tracking Schedule)
    Range("G1:BH4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        '.MergeCells = True
    End With

' Save as pdf code
    Dim wsA As Worksheet
    Dim wbA As Workbook
    Dim strTime As String
    Dim strName As String
    Dim strPath As String
    Dim strFile As String
    Dim strPathFile As String
    Dim myFile As Variant
    On Error GoTo ErrHandler
    
    Set wbA = ActiveWorkbook
    Set wsA = ActiveSheet
    strTime = Format(Now(), "yyyymmdd")
    
    ' get job number as a variable
    
    Dim jobNo As String
    Worksheets("Weekly Report").Activate
    jobNo = Cells(3, "C").Value
    
    'get active workbook folder, if saved
    strPath = wbA.Path
    If strPath = "" Then
      strPath = Application.DefaultFilePath
    End If
    strPath = strPath & "\"
    
    'replace spaces and periods in sheet name
    strName = Replace(wsA.name, " ", "")
    strName = Replace(strName, ".", "_")
    
    'create default name for savng file
    strFile = jobNo & "_" & strName & "_" & strTime & ".pdf"
    strPathFile = strPath & strFile
    
    'use can enter name and
    ' select folder for file
    myFile = Application.GetSaveAsFilename _
        (InitialFileName:=strPathFile, _
            FileFilter:="PDF Files (*.pdf), *.pdf", _
            Title:="Select Folder and FileName to save")
    
    'export to PDF if a folder was selected
    If myFile <> "False" Then
        wsA.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=myFile, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
        'confirmation message with file info
        MsgBox "PDF file has been created: " _
          & vbCrLf _
          & myFile
    End If
    
exitHandler:
        Exit Sub
ErrHandler:
        MsgBox "Could not create PDF file"
        Resume exitHandler

End Sub
Sub Print_Delivery_Tracking_Schedule()

   Dim ws As Worksheet
    Dim lo As ListObject

    '--- adjust sheet name if required
    Set ws = ThisWorkbook.Worksheets("Tracking Schedule")
    Set lo = ws.ListObjects("tbl_Tracking")

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanFail

    '1) Work inside tbl_Tracking: clear any filters
    'With lo
    '    If .AutoFilter.FilterMode Then .AutoFilter.ShowAllData
        'If filters exist but no data is filtered, ShowAllData can error - ignore safely
    'End With

    '2) Clear current filters (belt + braces: also clears worksheet filter state)
    'On Error Resume Next
    'If ws.FilterMode Then ws.ShowAllData
    'On Error GoTo CleanFail

    '3) Unhide all columns
    ws.Columns.Hidden = False

    '4) Hide all columns between O and AY, J and N
    ws.Columns("O:AY").Hidden = True
    ws.Columns("J:N").Hidden = True

' Save as pdf code
    Dim wsA As Worksheet
    Dim wbA As Workbook
    Dim strTime As String
    Dim strName As String
    Dim strPath As String
    Dim strFile As String
    Dim strPathFile As String
    Dim myFile As Variant
    On Error GoTo CleanFail
    
    Set wbA = ActiveWorkbook
    Set wsA = ActiveSheet
    strTime = Format(Now(), "yyyymmdd")
    
    ' get job number as a variable
    
    Dim jobNo As String
    Worksheets("Weekly Report").Activate
    jobNo = Cells(3, "C").Value
    
    'get active workbook folder, if saved
    strPath = wbA.Path
    If strPath = "" Then
      strPath = Application.DefaultFilePath
    End If
    strPath = strPath & "\"
    
    'replace spaces and periods in sheet name
    strName = Replace(wsA.name, " ", "")
    strName = Replace(strName, ".", "_")
    
    'create default name for savng file
    strFile = jobNo & "_" & strName & "_" & strTime & ".pdf"
    strPathFile = strPath & strFile
    
    'use can enter name and
    ' select folder for file
    myFile = Application.GetSaveAsFilename _
        (InitialFileName:=strPathFile, _
            FileFilter:="PDF Files (*.pdf), *.pdf", _
            Title:="Select Folder and FileName to save")
    
    'export to PDF if a folder was selected
    If myFile <> "False" Then
        wsA.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=myFile, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
        'confirmation message with file info
        MsgBox "PDF file has been created: " _
          & vbCrLf _
          & myFile
    End If


CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    'If ShowAllData failed because nothing was filtered, it’s not a real failure.
    Resume CleanExit

End Sub


Sub Unhide_Tracking_Schedule()
'
' Unhide_All Macro
' Unhide all hidden columns
'

'
    Columns("A:BJ").Select
    Range("BJ1").Activate
    Selection.EntireColumn.Hidden = False
    
' Re-merge column D to F to show client name
    Range("D1:F1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = True
    End With
' Re-merge column D to F to show project name

    Range("D2:F2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = True
    End With
' Align Macro
'

    Range("G1:BH2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
End Sub




'-------------------------------------------------------
' SplitSelectedTrackingRow
' Description:
' Splits the currently selected data row in tbl_Tracking by:
'   1) validating a single selected table row,
'   2) prompting for a split quantity,
'   3) inserting a new row directly beneath the selected row,
'   4) copying only nominated worksheet-column bands from the original row
'      to the new row,
'   5) assigning the split quantity to the new row, and
'   6) reducing the original row quantity accordingly.
'
' Inputs:
' - Active selection must intersect exactly one data row in tbl_Tracking
' - Worksheet: "Tracking Schedule"
' - Table: "tbl_Tracking"
' - Quantity header: "Assembly Quantity"
'
' Outputs:
' - Inserts one new ListRow
' - Updates Assembly Quantity in both original and new rows
' - Displays completion / validation / error messages
'
' Dependencies:
' - AppGuard_Begin / AppGuard_End
' - SheetGuard_Begin / SheetGuard_End
' - TSheetGuardState
'
' Assumptions:
' - The nominated copy bands are worksheet columns, not table-relative columns
' - Full split is intentionally blocked so the original row never goes to zero
'-------------------------------------------------------
Public Sub SplitSelectedTrackingRow()

    Const PROC_NAME As String = "SplitSelectedTrackingRow"
    Const WS_NAME   As String = "Tracking Schedule"
    Const TBL_NAME  As String = "tbl_Tracking"
    Const QTY_HDR   As String = "Assembly Quantity"

    Dim ws As Worksheet
    Dim lo As ListObject

    Dim selRowIndex As Long
    Dim qtyColIndex As Long

    Dim originalQty As Double
    Dim splitQty As Double
    Dim remainingQty As Double

    Dim newListRow As ListRow
    Dim newRowIndex As Long

    Dim sgState As TSheetGuardState
    Dim appGuardStarted As Boolean
    Dim sheetGuardStarted As Boolean

    On Error GoTo ErrHandler

    '-------------------------------------------------------
    ' Begin application guard for safe performance mode
    '-------------------------------------------------------
    AppGuard_Begin True, "Splitting tracking row..."
    appGuardStarted = True

    '-------------------------------------------------------
    ' Resolve worksheet and table
    '-------------------------------------------------------
    Set ws = ThisWorkbook.Worksheets(WS_NAME)
    Set lo = ws.ListObjects(TBL_NAME)

    '-------------------------------------------------------
    ' Temporarily unprotect sheet using modGuardsAndTables
    '-------------------------------------------------------
    sgState = SheetGuard_Begin(ws)
    sheetGuardStarted = True

    '-------------------------------------------------------
    ' Validate table has data rows
    '-------------------------------------------------------
    If lo.DataBodyRange Is Nothing Then
        MsgBox "Table '" & TBL_NAME & "' has no data rows.", vbExclamation, PROC_NAME
        GoTo SafeExit
    End If

    '-------------------------------------------------------
    ' Validate current selection resolves to exactly one table data row
    '-------------------------------------------------------
    selRowIndex = GetSingleSelectedTableRowIndex(lo, Selection)
    If selRowIndex = 0 Then GoTo SafeExit

    '-------------------------------------------------------
    ' Resolve and validate quantity column
    '-------------------------------------------------------
    qtyColIndex = GetTableColumnIndexSafe(lo, QTY_HDR)
    If qtyColIndex = 0 Then
        MsgBox "Column '" & QTY_HDR & "' was not found in table '" & TBL_NAME & "'.", _
               vbCritical, PROC_NAME
        GoTo SafeExit
    End If

    If Not IsNumeric(lo.DataBodyRange.Cells(selRowIndex, qtyColIndex).Value2) Then
        MsgBox "'" & QTY_HDR & "' on the selected row is blank or non-numeric.", _
               vbExclamation, PROC_NAME
        GoTo SafeExit
    End If

    originalQty = CDbl(lo.DataBodyRange.Cells(selRowIndex, qtyColIndex).Value2)

    If originalQty <= 0 Then
        MsgBox "Selected row has an invalid '" & QTY_HDR & "' value of " & _
               Format$(originalQty, "0.########") & ".", _
               vbExclamation, PROC_NAME
        GoTo SafeExit
    End If

    '-------------------------------------------------------
    ' Prompt user for split quantity
    '-------------------------------------------------------
    splitQty = PromptForSplitQuantity(originalQty, QTY_HDR)
    If splitQty <= 0 Then GoTo SafeExit

    '-------------------------------------------------------
    ' Block full split so original row does not become zero
    '-------------------------------------------------------
    If splitQty >= originalQty Then
        MsgBox "Split quantity must be less than the current '" & QTY_HDR & "' (" & _
               Format$(originalQty, "0.########") & ").", _
               vbExclamation, PROC_NAME
        GoTo SafeExit
    End If

    remainingQty = originalQty - splitQty

    '-------------------------------------------------------
    ' Insert new row directly below selected row
    '-------------------------------------------------------
    Set newListRow = lo.ListRows.Add(Position:=selRowIndex + 1)
    newRowIndex = newListRow.Index

    '-------------------------------------------------------
    ' Copy only required worksheet column bands from selected row to new row
    ' Worksheet bands:
    '   A:F, H:I, N:Q, BC:BD, BN:BO
    '-------------------------------------------------------
    CopyWorksheetColumnBandsToNewTableRow lo, selRowIndex, newRowIndex

    '-------------------------------------------------------
    ' Set split quantity in new row and reduce original row quantity
    '-------------------------------------------------------
    lo.DataBodyRange.Cells(newRowIndex, qtyColIndex).Value2 = splitQty
    lo.DataBodyRange.Cells(selRowIndex, qtyColIndex).Value2 = remainingQty

    MsgBox "Row split complete." & vbCrLf & _
           "Original " & QTY_HDR & ": " & Format$(originalQty, "0.########") & vbCrLf & _
           "New row " & QTY_HDR & ": " & Format$(splitQty, "0.########") & vbCrLf & _
           "Remaining on original row: " & Format$(remainingQty, "0.########"), _
           vbInformation, PROC_NAME

SafeExit:
    On Error Resume Next

    If sheetGuardStarted Then
        SheetGuard_End ws, sgState
    End If

    If appGuardStarted Then
        AppGuard_End
    End If

    On Error GoTo 0
    Exit Sub

ErrHandler:
    MsgBox "Error in " & PROC_NAME & ":" & vbCrLf & Err.Description, vbCritical, PROC_NAME
    Resume SafeExit

End Sub

'-------------------------------------------------------
' GetSingleSelectedTableRowIndex
' Description:
' Validates that the supplied selection intersects exactly one data row
' within the specified ListObject, then returns that row index relative
' to the table's DataBodyRange.
'
' Inputs:
' - lo: target ListObject
' - sel: current Excel selection
'
' Outputs:
' - Returns 0 if selection is invalid
' - Returns 1-based DataBodyRange row index if valid
'
' Assumptions:
' - A selection spanning multiple rows is rejected
' - Header-only or outside-table selections are rejected
'-------------------------------------------------------
Private Function GetSingleSelectedTableRowIndex(ByVal lo As ListObject, _
                                                ByVal sel As Object) As Long
    Dim rngSel As Range
    Dim rngHit As Range
    Dim firstRowAbs As Long
    Dim lastRowAbs As Long

    On Error GoTo Fail

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If sel Is Nothing Then Exit Function
    If Not TypeOf sel Is Range Then Exit Function

    Set rngSel = sel
    Set rngHit = Intersect(rngSel, lo.DataBodyRange)

    If rngHit Is Nothing Then
        MsgBox "Please select exactly one data row inside table '" & lo.name & "'.", _
               vbExclamation, "Invalid Selection"
        Exit Function
    End If

    firstRowAbs = rngHit.Row
    lastRowAbs = rngHit.Row + rngHit.Rows.Count - 1

    If firstRowAbs <> lastRowAbs Then
        MsgBox "Please select cells from one table row only.", _
               vbExclamation, "Invalid Selection"
        Exit Function
    End If

    GetSingleSelectedTableRowIndex = firstRowAbs - lo.DataBodyRange.Row + 1
    Exit Function

Fail:
    GetSingleSelectedTableRowIndex = 0
End Function

'-------------------------------------------------------
' GetTableColumnIndexSafe
' Description:
' Returns the 1-based ListColumn index for the supplied header caption.
' Safe version that returns 0 instead of raising an error.
'
' Inputs:
' - lo: target ListObject
' - headerText: header name to find
'
' Outputs:
' - 0 if not found
' - 1-based column index if found
'-------------------------------------------------------
Private Function GetTableColumnIndexSafe(ByVal lo As ListObject, _
                                         ByVal headerText As String) As Long
    Dim lc As ListColumn

    If lo Is Nothing Then Exit Function

    For Each lc In lo.ListColumns
        If StrComp(Trim$(lc.name), Trim$(headerText), vbTextCompare) = 0 Then
            GetTableColumnIndexSafe = lc.Index
            Exit Function
        End If
    Next lc
End Function

'-------------------------------------------------------
' PromptForSplitQuantity
' Description:
' Prompts the user to enter a split quantity and validates the result.
'
' Inputs:
' - currentQty: current quantity on selected row
' - qtyHeader: header label used in prompt messaging
'
' Outputs:
' - Returns 0 if cancelled or invalid
' - Returns validated Double if successful
'-------------------------------------------------------
Private Function PromptForSplitQuantity(ByVal currentQty As Double, _
                                        ByVal qtyHeader As String) As Double
    Dim response As Variant
    Dim enteredValue As Double

    response = Application.InputBox( _
                    Prompt:="Enter split quantity for '" & qtyHeader & "'." & vbCrLf & _
                            "Current value: " & Format$(currentQty, "0.########"), _
                    Title:="Split Tracking Row", _
                    Type:=1)

    If VarType(response) = vbBoolean Then
        If response = False Then Exit Function
    End If

    If Not IsNumeric(response) Then
        MsgBox "Split quantity must be numeric.", vbExclamation, "Invalid Quantity"
        Exit Function
    End If

    enteredValue = CDbl(response)

    If enteredValue <= 0 Then
        MsgBox "Split quantity must be greater than zero.", vbExclamation, "Invalid Quantity"
        Exit Function
    End If

    PromptForSplitQuantity = enteredValue
End Function

'-------------------------------------------------------
' CopyWorksheetColumnBandsToNewTableRow
' Description:
' Copies only approved worksheet-column bands from the original table row
' to the new table row.
'
' Inputs:
' - lo: target ListObject
' - sourceRowIndex: original row index within DataBodyRange
' - targetRowIndex: new row index within DataBodyRange
'
' Outputs:
' - Copies formulas and values from source row to target row
'-------------------------------------------------------
Private Sub CopyWorksheetColumnBandsToNewTableRow(ByVal lo As ListObject, _
                                                  ByVal sourceRowIndex As Long, _
                                                  ByVal targetRowIndex As Long)

    CopyWorksheetColsToNewTableRowRange lo, sourceRowIndex, targetRowIndex, 1, 6     ' A:F
    CopyWorksheetColsToNewTableRowRange lo, sourceRowIndex, targetRowIndex, 8, 9     ' H:I
    CopyWorksheetColsToNewTableRowRange lo, sourceRowIndex, targetRowIndex, 14, 17   ' N:Q
    CopyWorksheetColsToNewTableRowRange lo, sourceRowIndex, targetRowIndex, 55, 56   ' BC:BD
    CopyWorksheetColsToNewTableRowRange lo, sourceRowIndex, targetRowIndex, 66, 67   ' BN:BO

End Sub

'-------------------------------------------------------
' CopyWorksheetColsToNewTableRowRange
' Description:
' Converts a worksheet-column band into a table-relative column band and
' copies the source row cells to the target row cells using direct
' formula assignment. This preserves formulas and constants.
'
' Inputs:
' - lo: target ListObject
' - sourceRowIndex: row to copy from
' - targetRowIndex: row to copy to
' - wsColStart: worksheet start column number
' - wsColEnd: worksheet end column number
'
' Outputs:
' - Writes matching cells into the inserted row
'
' Assumptions:
' - Table may start in any worksheet column
' - Any portion of the worksheet band outside the table is ignored
'-------------------------------------------------------
Private Sub CopyWorksheetColsToNewTableRowRange(ByVal lo As ListObject, _
                                                ByVal sourceRowIndex As Long, _
                                                ByVal targetRowIndex As Long, _
                                                ByVal wsColStart As Long, _
                                                ByVal wsColEnd As Long)

    Dim tblWsFirstCol As Long
    Dim tblWsLastCol As Long
    Dim copyWsStart As Long
    Dim copyWsEnd As Long
    Dim tblColStart As Long
    Dim tblColEnd As Long
    Dim srcRng As Range
    Dim tgtRng As Range

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If wsColEnd < wsColStart Then Exit Sub

    tblWsFirstCol = lo.Range.Column
    tblWsLastCol = lo.Range.Column + lo.ListColumns.Count - 1

    copyWsStart = Application.Max(wsColStart, tblWsFirstCol)
    copyWsEnd = Application.Min(wsColEnd, tblWsLastCol)

    If copyWsStart > copyWsEnd Then Exit Sub

    tblColStart = copyWsStart - tblWsFirstCol + 1
    tblColEnd = copyWsEnd - tblWsFirstCol + 1

    If tblColStart < 1 Or tblColEnd > lo.ListColumns.Count Then
        Err.Raise vbObjectError + 513, _
                  "CopyWorksheetColsToNewTableRowRange", _
                  "Calculated table-relative column range is outside table bounds."
    End If

    Set srcRng = lo.DataBodyRange.Rows(sourceRowIndex).Cells(1, tblColStart).Resize(1, tblColEnd - tblColStart + 1)
    Set tgtRng = lo.DataBodyRange.Rows(targetRowIndex).Cells(1, tblColStart).Resize(1, tblColEnd - tblColStart + 1)

    tgtRng.Formula = srcRng.Formula

End Sub

'===============================================================================
' Copies one worksheet column band into the destination table row.
' Converts worksheet columns to table-relative columns first.
'===============================================================================
Private Sub CopyWorksheetBand( _
    ByVal lo As ListObject, _
    ByVal srcRowIndex As Long, _
    ByVal dstRowIndex As Long, _
    ByVal wsColStartLetter As String, _
    ByVal wsColEndLetter As String)

    Dim wsStartCol As Long
    Dim wsEndCol As Long
    Dim tblStartCol As Long
    Dim tblEndCol As Long
    Dim i As Long

    wsStartCol = ColumnLetterToNumber(wsColStartLetter)
    wsEndCol = ColumnLetterToNumber(wsColEndLetter)

    If wsStartCol <= 0 Or wsEndCol <= 0 Then Exit Sub
    If wsEndCol < wsStartCol Then Exit Sub

    'Convert worksheet columns into table-relative columns
    tblStartCol = wsStartCol - lo.Range.Column + 1
    tblEndCol = wsEndCol - lo.Range.Column + 1

    'Clamp to table bounds in case band partially falls outside the table
    If tblStartCol < 1 Then tblStartCol = 1
    If tblEndCol > lo.ListColumns.Count Then tblEndCol = lo.ListColumns.Count

    If tblStartCol > lo.ListColumns.Count Then Exit Sub
    If tblEndCol < 1 Then Exit Sub
    If tblEndCol < tblStartCol Then Exit Sub

    For i = tblStartCol To tblEndCol
        lo.DataBodyRange.Cells(dstRowIndex, i).Value = lo.DataBodyRange.Cells(srcRowIndex, i).Value
    Next i
End Sub

'===============================================================================
' Converts Excel column letter(s) to column number.
' Examples:
'   A  -> 1
'   Z  -> 26
'   AA -> 27
'===============================================================================
Private Function ColumnLetterToNumber(ByVal colLetter As String) As Long

    Dim i As Long
    Dim result As Long
    Dim ch As String

    colLetter = UCase$(Trim$(colLetter))
    If Len(colLetter) = 0 Then Exit Function

    For i = 1 To Len(colLetter)
        ch = Mid$(colLetter, i, 1)
        If ch < "A" Or ch > "Z" Then
            ColumnLetterToNumber = 0
            Exit Function
        End If
        result = result * 26 + (asc(ch) - asc("A") + 1)
    Next i

    ColumnLetterToNumber = result
End Function


Public Sub GotoLinkFromCell(ByVal Target As Range)
    Dim srcTbl As ListObject
    Dim clickedColIndex As Long
    Dim clickedColName As String
    Dim findValue As String

    Dim wsTrack As Worksheet
    Dim trackTbl As ListObject
    Dim trackFindColName As String
    Dim trackFindColIndex As Long

    Dim searchRange As Range
    Dim foundCell As Range

    Dim shGuard As TSheetGuardState

    On Error GoTo ErrHandler

    '------------------------------------------------------------
    ' Validate target selection
    '------------------------------------------------------------
    If Target Is Nothing Then Exit Sub
    If Target.Cells.CountLarge <> 1 Then Exit Sub

    ' Must be inside a table
    If Target.ListObject Is Nothing Then Exit Sub
    Set srcTbl = Target.ListObject

    ' Determine which source table column was clicked
    clickedColIndex = Target.Column - srcTbl.Range.Column + 1
    If clickedColIndex < 1 Or clickedColIndex > srcTbl.ListColumns.Count Then Exit Sub

    clickedColName = srcTbl.ListColumns(clickedColIndex).name
    findValue = Trim$(CStr(Target.Value2))

    ' Nothing to search for
    If Len(findValue) = 0 Then Exit Sub

    '------------------------------------------------------------
    ' Map source column to destination column in tbl_Tracking
    '------------------------------------------------------------
    Select Case clickedColName
        Case "Assembly No."
            trackFindColName = "Asset Number"
        
        Case "Mark Number/ Assembly/ ID"
            trackFindColName = "Asset Number"
        
        Case "Workpack"
            trackFindColName = "Workpack"

        Case Else
            Exit Sub
    End Select

    '------------------------------------------------------------
    ' Get destination worksheet and table
    '------------------------------------------------------------
    Set wsTrack = ThisWorkbook.Worksheets("Tracking Schedule")
    Set trackTbl = wsTrack.ListObjects("tbl_Tracking")

    ' Ensure destination column exists
    On Error GoTo BadColumn
    trackFindColIndex = trackTbl.ListColumns(trackFindColName).Index
    On Error GoTo ErrHandler

    ' If table has no data rows, there is nothing to find
    If trackTbl.DataBodyRange Is Nothing Then
        MsgBox "Table 'tbl_Tracking' has no data rows to search.", _
               vbExclamation, "No Data"
        GoTo SafeExit
    End If

    '------------------------------------------------------------
    ' Build search range from the mapped destination column
    '------------------------------------------------------------
    Set searchRange = trackTbl.ListColumns(trackFindColName).DataBodyRange

    '------------------------------------------------------------
    ' Find exact match in destination column
    ' - LookIn:=xlValues ensures we search displayed values
    ' - LookAt:=xlWhole enforces exact match
    ' - MatchCase:=False keeps it user-friendly
    '------------------------------------------------------------
    Set foundCell = searchRange.Find(What:=findValue, _
                                     After:=searchRange.Cells(searchRange.Cells.Count), _
                                     LookIn:=xlValues, _
                                     LookAt:=xlWhole, _
                                     SearchOrder:=xlByRows, _
                                     SearchDirection:=xlNext, _
                                     matchCase:=False)

    If foundCell Is Nothing Then
        MsgBox "No match found in '" & trackFindColName & "' for:" & vbCrLf & vbCrLf & _
               findValue, vbInformation, "Not Found"
        GoTo SafeExit
    End If

    Application.ScreenUpdating = False

    '------------------------------------------------------------
    ' Activate destination sheet and select the matched cell
    '------------------------------------------------------------
    shGuard = SheetGuard_Begin(wsTrack, pwd)

    wsTrack.Activate
    Application.GoTo foundCell, True

SafeExit:
    On Error Resume Next
    SheetGuard_End wsTrack, shGuard, pwd
    On Error GoTo 0

    Application.ScreenUpdating = True
    Exit Sub

BadColumn:
    MsgBox "Couldn't find column '" & trackFindColName & _
           "' in table 'tbl_Tracking' on 'Tracking Schedule'.", _
           vbExclamation, "Column Missing"
    GoTo SafeExit

ErrHandler:
    MsgBox "GotoLinkFromCell failed: " & Err.Description, _
           vbExclamation, "Macro Error"
    GoTo SafeExit
End Sub

Public Sub UnbindLinkPassShortcut()
    Application.OnKey "^+g"
End Sub

Public Sub BindLinkPassShortcut()
    ' Recommended: Ctrl+Shift+G
    Application.OnKey "^+g", "LinkPass"
End Sub

Public Sub LinkPass()
Attribute LinkPass.VB_ProcData.VB_Invoke_Func = "g\n14"
    ' Shortcut entry point.
    ' Runs against the currently selected cell.

    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Cells.CountLarge <> 1 Then Exit Sub

    GotoLinkFromCell Selection.Cells(1, 1)
    
End Sub

'-------------------------------------------------------
' AddSelectedTrackingRowsToMTO_VisibleOnly
' Description:
'   Processes the current selection from tbl_Tracking,
'   but only for visible selected data rows, and appends
'   one new line item per valid selected row into tbl_MTO.
'
'   Only one field is transferred:
'       tbl_Tracking.[Asset Number] -> tbl_MTO.[Assembly No.]
'
' Inputs:
'   - Current worksheet selection
'   - Source table: tbl_Tracking
'   - Target table: tbl_MTO
'   - Helper procedures/functions:
'       GetWorksheetOfTable
'       TableColIndexRequired
'       AppGuard_Begin / AppGuard_End
'       SheetGuard_Begin / SheetGuard_End
'       IsTableDataRowVisible
'
' Outputs:
'   - Appends new rows into tbl_MTO
'   - Displays a completion summary
'
' Assumptions:
'   - tbl_Tracking contains a column named "Asset Number"
'   - tbl_MTO contains a column named "Assembly No."
'   - Sheet protection helpers are available in modGuardsAndTables
'   - Selection may include multiple areas and partial row selections
'-------------------------------------------------------
Public Sub AddSelectedTrackingRowsToMTO_VisibleOnly()

    Const PROC_NAME As String = "AddSelectedTrackingRowsToMTO_VisibleOnly"
    Const SOURCE_TBL As String = "tbl_Tracking"
    Const TARGET_TBL As String = "tbl_MTO"
    Const SOURCE_HDR_ASSET As String = "Asset Number"
    Const TARGET_HDR_ASSEMBLY As String = "Assembly No."

    '-------------------------------------------------------
    ' Core Excel objects
    '-------------------------------------------------------
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim loSource As ListObject
    Dim loTarget As ListObject

    '-------------------------------------------------------
    ' Protection state
    '-------------------------------------------------------
    Dim sgTarget As TSheetGuardState

    '-------------------------------------------------------
    ' Column indexes
    '-------------------------------------------------------
    Dim idxSourceAsset As Long
    Dim idxTargetAssembly As Long

    '-------------------------------------------------------
    ' Selection handling
    '-------------------------------------------------------
    Dim selRng As Range
    Dim hitRng As Range
    Dim area As Range
    Dim rw As Range
    Dim relRow As Long

    '-------------------------------------------------------
    ' Working collections
    '-------------------------------------------------------
    Dim dictRows As Object   ' Unique visible selected source row numbers
    Dim key As Variant

    '-------------------------------------------------------
    ' Working values
    '-------------------------------------------------------
    Dim assetNo As String

    '-------------------------------------------------------
    ' Bulk write variables
    '-------------------------------------------------------
    Dim oldRows As Long
    Dim newRows As Long
    Dim addCount As Long
    Dim firstNewRow As Long
    Dim writeCols As Long
    Dim outArr() As Variant
    Dim outR As Long

    '-------------------------------------------------------
    ' Counters / reporting
    '-------------------------------------------------------
    Dim selectionRowTouches As Long
    Dim visibleRowsReviewed As Long
    Dim addedCount As Long
    Dim skipHiddenRows As Long
    Dim skipBlankAsset As Long

    '-------------------------------------------------------
    ' Reporting buffers
    '-------------------------------------------------------
    Dim detail As String
    Dim msg As String

    On Error GoTo ErrHandler

    '-------------------------------------------------------
    ' Begin application performance guard
    '-------------------------------------------------------
    AppGuard_Begin

    '-------------------------------------------------------
    ' Resolve source and target worksheets / tables
    '-------------------------------------------------------
    Set wsSource = GetWorksheetOfTable(ThisWorkbook, SOURCE_TBL)
    Set wsTarget = GetWorksheetOfTable(ThisWorkbook, TARGET_TBL)

    If wsSource Is Nothing Then
        Err.Raise 5, PROC_NAME, "Worksheet for table '" & SOURCE_TBL & "' was not found."
    End If

    If wsTarget Is Nothing Then
        Err.Raise 5, PROC_NAME, "Worksheet for table '" & TARGET_TBL & "' was not found."
    End If

    Set loSource = wsSource.ListObjects(SOURCE_TBL)
    Set loTarget = wsTarget.ListObjects(TARGET_TBL)

    If loSource Is Nothing Then
        Err.Raise 5, PROC_NAME, "Table '" & SOURCE_TBL & "' was not found."
    End If

    If loTarget Is Nothing Then
        Err.Raise 5, PROC_NAME, "Table '" & TARGET_TBL & "' was not found."
    End If

    '-------------------------------------------------------
    ' Validate source table has data rows
    '-------------------------------------------------------
    If loSource.DataBodyRange Is Nothing Then
        MsgBox "Source table '" & SOURCE_TBL & "' has no data rows.", vbExclamation, PROC_NAME
        GoTo SafeExit
    End If

    '-------------------------------------------------------
    ' Validate current selection intersects source table body
    '-------------------------------------------------------
    Set selRng = Selection
    If selRng Is Nothing Then
        MsgBox "Select one or more rows within " & SOURCE_TBL & " and run again.", vbExclamation, PROC_NAME
        GoTo SafeExit
    End If

    Set hitRng = Intersect(selRng, loSource.DataBodyRange)
    If hitRng Is Nothing Then
        MsgBox "No valid selection found inside " & SOURCE_TBL & ".", vbExclamation, PROC_NAME
        GoTo SafeExit
    End If

    '-------------------------------------------------------
    ' Resolve required headers
    '-------------------------------------------------------
    idxSourceAsset = TableColIndexRequired(loSource, SOURCE_HDR_ASSET)
    idxTargetAssembly = TableColIndexRequired(loTarget, TARGET_HDR_ASSEMBLY)

    '-------------------------------------------------------
    ' Collect unique visible selected source rows
    '
    ' Notes:
    '   - A user may select multiple cells within the same row
    '   - A user may select multiple areas
    '   - We only want each visible table row once
    '-------------------------------------------------------
    Set dictRows = CreateObject("Scripting.Dictionary")
    dictRows.CompareMode = vbTextCompare

    For Each area In hitRng.Areas
        For Each rw In area.Rows

            selectionRowTouches = selectionRowTouches + 1

            relRow = rw.Row - loSource.DataBodyRange.Row + 1

            If relRow >= 1 And relRow <= loSource.DataBodyRange.Rows.Count Then
                If IsTableDataRowVisible(loSource, relRow) Then
                    If Not dictRows.Exists(CStr(relRow)) Then
                        dictRows.Add CStr(relRow), relRow
                    End If
                Else
                    skipHiddenRows = skipHiddenRows + 1
                End If
            End If

        Next rw
    Next area

    If dictRows.Count = 0 Then
        MsgBox "No visible selected data rows were found inside " & SOURCE_TBL & ".", vbExclamation, PROC_NAME
        GoTo SafeExit
    End If

    visibleRowsReviewed = dictRows.Count

    '-------------------------------------------------------
    ' Count valid rows to add
    '
    ' Only rows with a non-blank Asset Number are appended.
    ' Blank Asset Numbers are skipped and reported.
    '-------------------------------------------------------
    addCount = 0

    For Each key In dictRows.keys
        relRow = CLng(dictRows(key))
        assetNo = Trim$(CStr(loSource.DataBodyRange.Cells(relRow, idxSourceAsset).Value2))

        If Len(assetNo) = 0 Then
            skipBlankAsset = skipBlankAsset + 1
            AppendLine detail, "Row " & relRow & " skipped: blank Asset Number."
        Else
            addCount = addCount + 1
        End If
    Next key

    If addCount = 0 Then
        msg = "MTO transfer complete." & vbCrLf & vbCrLf & _
              "Selection rows intersected: " & selectionRowTouches & vbCrLf & _
              "Visible table rows reviewed: " & visibleRowsReviewed & vbCrLf & _
              "Rows skipped - hidden / filtered out: " & skipHiddenRows & vbCrLf & _
              "Rows skipped - blank Asset Number: " & skipBlankAsset & vbCrLf & _
              "Rows added to " & TARGET_TBL & ": 0"

        If Len(detail) > 0 Then
            msg = msg & vbCrLf & vbCrLf & "Detail:" & vbCrLf & detail
        End If

        MsgBox msg, vbInformation, PROC_NAME
        GoTo SafeExit
    End If

    '-------------------------------------------------------
    ' Unprotect target sheet before resizing and writing
    '-------------------------------------------------------
    sgTarget = SheetGuard_Begin(wsTarget)

    '-------------------------------------------------------
    ' Determine current and new target row counts
    '-------------------------------------------------------
    If loTarget.DataBodyRange Is Nothing Then
        oldRows = 0
    Else
        oldRows = loTarget.DataBodyRange.Rows.Count
    End If

    newRows = oldRows + addCount
    writeCols = loTarget.ListColumns.Count

    '-------------------------------------------------------
    ' Resize target table once only
    '
    ' ListObject.Range includes the header row, so total rows
    ' required = data rows + 1 header row.
    '-------------------------------------------------------
    loTarget.Resize loTarget.Range.Resize(RowSize:=newRows + 1, ColumnSize:=loTarget.Range.Columns.Count)

    '-------------------------------------------------------
    ' Build output array for appended rows only
    '
    ' Only the target field [Assembly No.] is populated.
    ' All other target columns are left blank so table defaults,
    ' formulas, or downstream processes can handle them.
    '-------------------------------------------------------
    ReDim outArr(1 To addCount, 1 To writeCols)

    outR = 0
    For Each key In dictRows.keys
        relRow = CLng(dictRows(key))
        assetNo = Trim$(CStr(loSource.DataBodyRange.Cells(relRow, idxSourceAsset).Value2))

        If Len(assetNo) > 0 Then
            outR = outR + 1
            outArr(outR, idxTargetAssembly) = assetNo
        End If
    Next key

    '-------------------------------------------------------
    ' Write appended rows in a single operation
    '-------------------------------------------------------
    firstNewRow = oldRows + 1
    loTarget.DataBodyRange.Rows(firstNewRow).Resize(addCount, writeCols).Value = outArr

    '-------------------------------------------------------
    ' Force calculated/formula columns in tbl_MTO to fill down
    ' after appending new rows via bulk write
    '-------------------------------------------------------
    FillDownCalculatedColumns loTarget

    addedCount = addCount

    '-------------------------------------------------------
    ' Display completion summary
    '-------------------------------------------------------
    msg = "MTO transfer complete." & vbCrLf & vbCrLf & _
          "Selection rows intersected: " & selectionRowTouches & vbCrLf & _
          "Visible table rows reviewed: " & visibleRowsReviewed & vbCrLf & _
          "Rows skipped - hidden / filtered out: " & skipHiddenRows & vbCrLf & _
          "Rows skipped - blank Asset Number: " & skipBlankAsset & vbCrLf & _
          "Rows added to " & TARGET_TBL & ": " & addedCount

    If Len(detail) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Detail:" & vbCrLf & detail
    End If

    '-------------------------------------------------------
    ' Activate tbl_MTO and select the newly added rows
    '
    ' This brings the user directly to the inserted block
    ' rather than the first row of the table.
    '-------------------------------------------------------
    If Not loTarget.DataBodyRange Is Nothing Then
        wsTarget.Activate
        loTarget.DataBodyRange.Cells(firstNewRow, idxTargetAssembly) _
            .Resize(addCount, 1).Select
        Application.GoTo loTarget.DataBodyRange.Cells(firstNewRow, idxTargetAssembly), True
    End If

    MsgBox msg, vbInformation, PROC_NAME

SafeExit:
    On Error Resume Next
    SheetGuard_End wsTarget, sgTarget
    AppGuard_End
    On Error GoTo 0
    Exit Sub

ErrHandler:
    Dim errMsg As String

    errMsg = "Error in " & PROC_NAME & ":" & vbCrLf & _
             Err.Number & " - " & Err.Description

    MsgBox errMsg, vbExclamation, PROC_NAME
    Resume SafeExit

End Sub

'====================================================================================
' AddSelectedTrackingRowsToInstall_VisibleOnly_Fast
'
' Purpose:
'   Processes the current selection from tbl_Tracking, but only visible rows.
'   Consolidates duplicate Asset Numbers, sums Assembly Quantity, validates data,
'   skips existing assets already in tbl_Install, then bulk-appends new rows into
'   tbl_Install using a single table resize + single array write.
'
' Why this version is faster:
'   - Avoids ListRows.Add in a loop
'   - Resizes tbl_Install once only
'   - Writes all new rows in one array assignment
'
' Requirements:
'   - GetWorksheetOfTable
'   - AppGuard_Begin / AppGuard_End
'   - SheetGuard_Begin / SheetGuard_End
'   - UpdateEarnedValue
'====================================================================================
Public Sub AddSelectedTrackingRowsToInstall_VisibleOnly()

    Const PROC_NAME As String = "AddSelectedTrackingRowsToInstall_VisibleOnly_Fast"
    Const TRACK_TBL As String = "tbl_Tracking"
    Const INSTALL_TBL As String = "tbl_Install"

    '----------------------------------------
    ' Core Excel objects
    '----------------------------------------
    Dim wsTrack As Worksheet
    Dim wsInstall As Worksheet
    Dim loTrack As ListObject
    Dim loInstall As ListObject

    '----------------------------------------
    ' Protection states
    '----------------------------------------
    Dim shInstall As TSheetGuardState

    '----------------------------------------
    ' Source column indexes
    '----------------------------------------
    Dim idxTrack_Asset As Long
    Dim idxTrack_AssyQty As Long
    Dim idxTrack_ProgQty As Long

    '----------------------------------------
    ' Target column indexes
    '----------------------------------------
    Dim idxInst_Asset As Long
    Dim idxInst_ProgQty As Long
    Dim idxInst_AssyQty As Long   'Optional

    '----------------------------------------
    ' Selection objects
    '----------------------------------------
    Dim selRng As Range
    Dim hitRng As Range
    Dim area As Range
    Dim rw As Range

    '----------------------------------------
    ' Dictionaries
    '----------------------------------------
    Dim dictRows As Object         'Unique visible selected source row numbers
    Dim dictGrouped As Object      'asset -> Array(asset, summedAssyQty, progQty)
    Dim dictConflicts As Object    'asset -> True
    Dim dictExisting As Object     'existing tbl_Install assets
    Dim dictToAdd As Object        'final assets to insert

    '----------------------------------------
    ' Working variables
    '----------------------------------------
    Dim key As Variant
    Dim relRow As Long
    Dim assetNo As String
    Dim assyQty As Double
    Dim progQty As Double
    Dim rec As Variant

    '----------------------------------------
    ' Bulk write variables
    '----------------------------------------
    Dim oldRows As Long
    Dim newRows As Long
    Dim addCount As Long
    Dim firstNewRow As Long
    Dim outArr() As Variant
    Dim writeCols As Long
    Dim outR As Long

    '----------------------------------------
    ' Counters / reporting
    '----------------------------------------
    Dim selectionRowTouches As Long
    Dim visibleRowsReviewed As Long
    Dim validUniqueAssets As Long
    Dim addedCount As Long
    Dim skipHiddenRows As Long
    Dim skipBlankAsset As Long
    Dim skipBadProgQty As Long
    Dim skipBadAssyQty As Long
    Dim skipAlreadyExists As Long
    Dim skipConflictProgQty As Long

    '----------------------------------------
    ' Reporting buffers
    '----------------------------------------
    Dim summary As String
    Dim conflictList As String
    Dim existingList As String

    On Error GoTo ErrHandler

    AppGuard_Begin

    '----------------------------------------
    ' Resolve worksheets / tables
    '----------------------------------------
    Set wsTrack = GetWorksheetOfTable(ThisWorkbook, TRACK_TBL)
    Set wsInstall = GetWorksheetOfTable(ThisWorkbook, INSTALL_TBL)

    If wsTrack Is Nothing Then Err.Raise 5, PROC_NAME, "Worksheet for table '" & TRACK_TBL & "' was not found."
    If wsInstall Is Nothing Then Err.Raise 5, PROC_NAME, "Worksheet for table '" & INSTALL_TBL & "' was not found."

    Set loTrack = wsTrack.ListObjects(TRACK_TBL)
    Set loInstall = wsInstall.ListObjects(INSTALL_TBL)

    If loTrack Is Nothing Then Err.Raise 5, PROC_NAME, "Table '" & TRACK_TBL & "' was not found."
    If loInstall Is Nothing Then Err.Raise 5, PROC_NAME, "Table '" & INSTALL_TBL & "' was not found."

    If loTrack.DataBodyRange Is Nothing Then
        MsgBox "Source table '" & TRACK_TBL & "' has no data rows.", vbExclamation, PROC_NAME
        GoTo SafeExit
    End If

    '----------------------------------------
    ' Validate selection
    '----------------------------------------
    Set selRng = Selection
    If selRng Is Nothing Then
        MsgBox "Select one or more rows within " & TRACK_TBL & " and run again.", vbExclamation, PROC_NAME
        GoTo SafeExit
    End If

    Set hitRng = Intersect(selRng, loTrack.DataBodyRange)
    If hitRng Is Nothing Then
        MsgBox "No valid selection found inside " & TRACK_TBL & ".", vbExclamation, PROC_NAME
        GoTo SafeExit
    End If

    '----------------------------------------
    ' Resolve required source headers
    '----------------------------------------
    idxTrack_Asset = TableColIndexRequired(loTrack, "Asset Number")
    idxTrack_AssyQty = TableColIndexRequired(loTrack, "Assembly Quantity")
    idxTrack_ProgQty = TableColIndexRequired(loTrack, "Progress Unit Qty")

    '----------------------------------------
    ' Resolve required target headers
    '----------------------------------------
    idxInst_Asset = TableColIndexRequired(loInstall, "Mark Number/ Assembly/ ID")
    idxInst_ProgQty = TableColIndexRequired(loInstall, "Progress Unit Qty")

    ' Optional quantity field in tbl_Install
    idxInst_AssyQty = TableColIndexAny(loInstall, _
                                       "Assembly Quantity", _
                                       "Qty", _
                                       "Quantity", _
                                       "Install Qty", _
                                       "Unit Qty")

    '----------------------------------------
    ' Create dictionaries
    '----------------------------------------
    Set dictRows = CreateObject("Scripting.Dictionary")
    Set dictGrouped = CreateObject("Scripting.Dictionary")
    Set dictConflicts = CreateObject("Scripting.Dictionary")
    Set dictExisting = CreateObject("Scripting.Dictionary")
    Set dictToAdd = CreateObject("Scripting.Dictionary")

    dictRows.CompareMode = vbTextCompare
    dictGrouped.CompareMode = vbTextCompare
    dictConflicts.CompareMode = vbTextCompare
    dictExisting.CompareMode = vbTextCompare
    dictToAdd.CompareMode = vbTextCompare

    '----------------------------------------
    ' Existing install asset lookup
    '----------------------------------------
    BuildExistingInstallLookup loInstall, idxInst_Asset, dictExisting

    '----------------------------------------
    ' Collect unique visible selected rows only
    '----------------------------------------
    For Each area In hitRng.Areas
        For Each rw In area.Rows

            selectionRowTouches = selectionRowTouches + 1

            relRow = rw.Row - loTrack.DataBodyRange.Row + 1

            If relRow >= 1 And relRow <= loTrack.DataBodyRange.Rows.Count Then
                If IsTableDataRowVisible(loTrack, relRow) Then
                    If Not dictRows.Exists(CStr(relRow)) Then
                        dictRows.Add CStr(relRow), relRow
                    End If
                Else
                    skipHiddenRows = skipHiddenRows + 1
                End If
            End If

        Next rw
    Next area

    If dictRows.Count = 0 Then
        MsgBox "No visible selected data rows were found inside " & TRACK_TBL & ".", vbExclamation, PROC_NAME
        GoTo SafeExit
    End If

    visibleRowsReviewed = dictRows.Count

    '----------------------------------------
    ' Validate and group source rows by Asset Number
    '----------------------------------------
    For Each key In dictRows.keys

        relRow = CLng(dictRows(key))

        assetNo = Trim$(CStr(loTrack.DataBodyRange.Cells(relRow, idxTrack_Asset).Value2))
        If Len(assetNo) = 0 Then
            skipBlankAsset = skipBlankAsset + 1
            AppendLine summary, "Row " & relRow & " skipped: blank Asset Number."
            GoTo NextRow
        End If

        If Not TryGetPositiveNumber(loTrack.DataBodyRange.Cells(relRow, idxTrack_ProgQty).Value2, progQty) Then
            skipBadProgQty = skipBadProgQty + 1
            AppendLine summary, "Asset '" & assetNo & "' skipped: invalid Progress Unit Qty."
            GoTo NextRow
        End If

        If Not TryGetNumber(loTrack.DataBodyRange.Cells(relRow, idxTrack_AssyQty).Value2, assyQty) Then
            skipBadAssyQty = skipBadAssyQty + 1
            AppendLine summary, "Asset '" & assetNo & "' skipped: invalid Assembly Quantity."
            GoTo NextRow
        End If

        If dictConflicts.Exists(assetNo) Then GoTo NextRow

        If Not dictGrouped.Exists(assetNo) Then
            rec = Array(assetNo, assyQty, progQty)
            dictGrouped.Add assetNo, rec
        Else
            rec = dictGrouped(assetNo)

            If CDbl(rec(2)) <> progQty Then
                dictGrouped.Remove assetNo
                dictConflicts(assetNo) = True
                skipConflictProgQty = skipConflictProgQty + 1
                AppendLine summary, "Asset '" & assetNo & "' skipped: conflicting Progress Unit Qty in selected visible rows."
                GoTo NextRow
            End If

            rec(1) = CDbl(rec(1)) + assyQty
            dictGrouped(assetNo) = rec
        End If

NextRow:
    Next key

    validUniqueAssets = dictGrouped.Count

    '----------------------------------------
    ' Filter out assets already in tbl_Install
    '----------------------------------------
    For Each key In dictGrouped.keys
        assetNo = CStr(key)

        If dictExisting.Exists(assetNo) Then
            skipAlreadyExists = skipAlreadyExists + 1
            AppendLine existingList, assetNo
        Else
            dictToAdd.Add assetNo, dictGrouped(assetNo)
        End If
    Next key

    addCount = dictToAdd.Count

    '----------------------------------------
    ' Build conflict list for reporting
    '----------------------------------------
    If dictConflicts.Count > 0 Then
        For Each key In dictConflicts.keys
            AppendLine conflictList, CStr(key)
        Next key
    End If

    '----------------------------------------
    ' If nothing to add, do not resize the table
    '----------------------------------------
    If addCount = 0 Then
        GoTo ShowSummary
    End If

    '----------------------------------------
    ' Unprotect target sheet before resizing / writing
    '----------------------------------------
    shInstall = SheetGuard_Begin(wsInstall)

    '----------------------------------------
    ' Work out existing / new row counts
    '----------------------------------------
    If loInstall.DataBodyRange Is Nothing Then
        oldRows = 0
    Else
        oldRows = loInstall.DataBodyRange.Rows.Count
    End If

    newRows = oldRows + addCount
    writeCols = loInstall.ListColumns.Count

    '----------------------------------------
    ' Resize tbl_Install once only
    '
    ' loInstall.Range includes header row.
    ' So resize row count = data rows + 1 header row.
    '----------------------------------------
    loInstall.Resize loInstall.Range.Resize(newRows + 1, loInstall.Range.Columns.Count)

    '----------------------------------------
    ' Build output array for new appended rows only
    '----------------------------------------
    ReDim outArr(1 To addCount, 1 To writeCols)

    outR = 1
    For Each key In dictToAdd.keys
        rec = dictToAdd(key)

        outArr(outR, idxInst_Asset) = CStr(rec(0))
        outArr(outR, idxInst_ProgQty) = CDbl(rec(2))

        If idxInst_AssyQty > 0 Then
            outArr(outR, idxInst_AssyQty) = CDbl(rec(1))
        End If

        outR = outR + 1
    Next key

    '----------------------------------------
    ' Write appended rows in one hit
    '----------------------------------------
    firstNewRow = oldRows + 1
    loInstall.DataBodyRange.Rows(firstNewRow).Resize(addCount, writeCols).Value = outArr

    addedCount = addCount

    '----------------------------------------
    ' Run downstream update
    '----------------------------------------
    If addedCount > 0 Then
        UpdateEarnedValue
    End If

ShowSummary:
    Dim msg As String

    msg = "Install transfer complete." & vbCrLf & vbCrLf & _
          "Selection rows intersected: " & selectionRowTouches & vbCrLf & _
          "Visible table rows reviewed: " & visibleRowsReviewed & vbCrLf & _
          "Rows skipped - hidden / filtered out: " & skipHiddenRows & vbCrLf & _
          "Valid unique assets found: " & validUniqueAssets & vbCrLf & _
          "Assets added to tbl_Install: " & addedCount & vbCrLf & _
          "Rows skipped - blank Asset Number: " & skipBlankAsset & vbCrLf & _
          "Rows skipped - invalid Progress Unit Qty: " & skipBadProgQty & vbCrLf & _
          "Rows skipped - invalid Assembly Quantity: " & skipBadAssyQty & vbCrLf & _
          "Assets skipped - already in tbl_Install: " & skipAlreadyExists & vbCrLf & _
          "Assets skipped - conflicting Progress Unit Qty: " & skipConflictProgQty

    If Len(existingList) > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "Already in tbl_Install:" & vbCrLf & existingList
    End If

    If Len(conflictList) > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "Conflicting Progress Unit Qty:" & vbCrLf & conflictList
    End If

    If Len(summary) > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "Detail:" & vbCrLf & summary
    End If

    '-------------------------------------------------------
    ' Activate tbl_Install and select the newly added rows
    '
    ' This brings the user directly to the inserted block
    ' rather than the first row of the table.
    '-------------------------------------------------------
    If Not loInstall.DataBodyRange Is Nothing Then
        wsInstall.Activate
        loInstall.DataBodyRange.Cells(firstNewRow, idxTrack_Asset) _
            .Resize(addCount, 1).Select
        Application.GoTo loInstall.DataBodyRange.Cells(firstNewRow, idxTrack_Asset), True
    End If

    MsgBox msg, vbInformation, PROC_NAME

SafeExit:
    On Error Resume Next
    SheetGuard_End wsInstall, shInstall
    AppGuard_End
    On Error GoTo 0
    Exit Sub

ErrHandler:
    Dim errMsg As String
    errMsg = "Error in " & PROC_NAME & ":" & vbCrLf & _
             Err.Number & " - " & Err.Description
    MsgBox errMsg, vbExclamation, PROC_NAME
    Resume SafeExit

End Sub

