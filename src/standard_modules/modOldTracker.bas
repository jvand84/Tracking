Attribute VB_Name = "modOldTracker"
Option Explicit

Public Sub ImportfromOld()

    BuildConfiguredTables_FromSelection

End Sub

'-------------------------------------------------------
' BuildConfiguredTables_FromSelection
'
' Purpose:
'   Prompts the user to select an open workbook via
'   frmSelect, resolves the selected workbook from the
'   workbook name returned by the form, and then builds
'   the configured tables in that workbook.
'
' Requirements:
'   - frmSelect must return the selected workbook name
'   - The selected workbook must already be open
'   - modGuardsAndTables must be available
'
' Notes:
'   - Cancelling or leaving no selection exits safely
'   - No fallback to ActiveWorkbook or ThisWorkbook
'-------------------------------------------------------
Public Sub BuildConfiguredTables_FromSelection()

    '-------------------------------------------------------
    ' Variable Declarations
    '-------------------------------------------------------
    Dim frm As frmSelect
    Dim selectedName As String
    Dim targetWb As Workbook

    '-------------------------------------------------------
    ' Guard Initialisation
    '-------------------------------------------------------
    On Error GoTo ErrHandler
    AppGuard_Begin

    '-------------------------------------------------------
    ' Launch workbook selection form
    '-------------------------------------------------------
    Set frm = New frmSelect
    frm.FrmType = 1
    frm.LoadCombo
    frm.Show

    selectedName = Trim$(frm.SelectedWorkbookName)

    Unload frm
    Set frm = Nothing

    '-------------------------------------------------------
    ' Validate selection
    '-------------------------------------------------------
    If Len(selectedName) = 0 Then
        MsgBox "No workbook selected. Exiting.", vbExclamation
        GoTo CleanExit
    End If

    Set targetWb = GetOpenWorkbookByName(selectedName)

    If targetWb Is Nothing Then
        Err.Raise vbObjectError + 2100, _
                  "BuildConfiguredTables_FromSelection", _
                  "The selected workbook '" & selectedName & "' is not open or could not be resolved."
    End If

    '-------------------------------------------------------
    ' Execute main processing
    '-------------------------------------------------------
    BuildConfiguredTablesFromWorkbook targetWb

CleanExit:
    '-------------------------------------------------------
    ' Cleanup
    '-------------------------------------------------------
    On Error Resume Next
    AppGuard_End
    On Error GoTo 0
    Exit Sub

ErrHandler:
    '-------------------------------------------------------
    ' Error Handling
    '-------------------------------------------------------
    LogError "BuildConfiguredTables_FromSelection", Err.Number, Err.Description
    MsgBox "BuildConfiguredTables_FromSelection failed:" & vbCrLf & _
           Err.Description, vbCritical, "Error"
    Resume CleanExit

End Sub

Public Sub BuildConfiguredTablesFromWorkbook(ByVal wb As Workbook)

    '-------------------------------------------------------
    ' Variable Declarations
    '-------------------------------------------------------
    Dim targetWb As Workbook
    Dim mapSheets As Object
    Dim wsName As Variant
    Dim ws As Worksheet
    Dim sgState As TSheetGuardState
    Dim processedCount As Long
    Dim skippedCount As Long
    Dim errorCount As Long
    Dim summary As String
    Dim errSummary As String
    Dim mapItem As Variant
    Dim tableName As String
    Dim headerRow As Long

    '-------------------------------------------------------
    ' Guard Initialisation
    '-------------------------------------------------------
    On Error GoTo ErrHandler

    '-------------------------------------------------------
    ' Resolve workbook and configuration map
    '-------------------------------------------------------
    If wb Is Nothing Then
        Err.Raise vbObjectError + 2000, _
                  "BuildConfiguredTablesFromWorkbook", _
                  "No workbook supplied."
    End If

    Set targetWb = wb
    Set mapSheets = BuildWorksheetTableMap()

    '-------------------------------------------------------
    ' Main Processing Logic
    '-------------------------------------------------------
    For Each wsName In mapSheets.keys

        Set ws = Nothing
        On Error Resume Next
        Set ws = targetWb.Worksheets(CStr(wsName))
        On Error GoTo ErrHandler

        If ws Is Nothing Then
            skippedCount = skippedCount + 1
            errSummary = errSummary & vbCrLf & "Missing sheet: " & CStr(wsName)
        Else
            On Error GoTo SheetErrHandler

            sgState = SheetGuard_Begin(ws)

            mapItem = mapSheets(CStr(wsName))
            tableName = CStr(mapItem(0))
            headerRow = CLng(mapItem(1))

            '-------------------------------------------------------
            ' Hard reset worksheet state before rebuilding table
            '-------------------------------------------------------
            ClearWorksheetAndTableFilters ws
            RemoveAllTablesFromWorksheet ws

            '-------------------------------------------------------
            ' Build the required configured table
            '-------------------------------------------------------
            CreateOrResizeConfiguredTable _
                ws:=ws, _
                tableName:=tableName, _
                headerRow:=headerRow

            processedCount = processedCount + 1

SheetCleanup:
            On Error Resume Next
            SheetGuard_End ws, sgState
            On Error GoTo ErrHandler
            GoTo NextSheet

SheetErrHandler:
            errorCount = errorCount + 1

            LogError "BuildConfiguredTablesFromWorkbook." & ws.name, Err.Number, Err.Description

            errSummary = errSummary & vbCrLf & _
                         ws.name & ": (" & Err.Number & ") " & Err.Description

            Resume SheetCleanup
        End If

NextSheet:
    Next wsName

    '-------------------------------------------------------
    ' Final Summary
    '-------------------------------------------------------
    summary = "Configured table build complete." & vbCrLf & vbCrLf & _
              "Processed: " & processedCount & vbCrLf & _
              "Skipped: " & skippedCount & vbCrLf & _
              "Errors: " & errorCount

    If Len(errSummary) > 0 Then
        summary = summary & vbCrLf & vbCrLf & "Details:" & errSummary
    End If

    MsgBox summary, vbInformation, "Build Configured Tables"
    Exit Sub

ErrHandler:
    LogError "BuildConfiguredTablesFromWorkbook", Err.Number, Err.Description
    MsgBox "BuildConfiguredTablesFromWorkbook failed:" & vbCrLf & _
           Err.Description, vbCritical, "Error"

End Sub

Private Sub CreateOrResizeConfiguredTable(ByVal ws As Worksheet, _
                                          ByVal tableName As String, _
                                          ByVal headerRow As Long)

    '-------------------------------------------------------
    ' Variable Declarations
    '-------------------------------------------------------
    Dim lastCol As Long
    Dim lastRow As Long
    Dim firstCol As Long
    Dim rngTable As Range
    Dim lo As ListObject
    Dim wbLo As ListObject

    On Error GoTo ErrHandler

    '-------------------------------------------------------
    ' Validation
    '-------------------------------------------------------
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1005, _
                  "CreateOrResizeConfiguredTable", _
                  "Worksheet reference is Nothing."
    End If

    If Len(Trim$(tableName)) = 0 Then
        Err.Raise vbObjectError + 1006, _
                  "CreateOrResizeConfiguredTable", _
                  "Table name is blank for sheet '" & ws.name & "'."
    End If

    If headerRow < 1 Or headerRow > ws.Rows.Count Then
        Err.Raise vbObjectError + 1000, _
                  "CreateOrResizeConfiguredTable", _
                  "Invalid configured header row for sheet '" & ws.name & "'."
    End If

    '-------------------------------------------------------
    ' Resolve target range
    '-------------------------------------------------------
    firstCol = GetFirstUsedColumnInRow(ws, headerRow)
    If firstCol = 0 Then
        Err.Raise vbObjectError + 1100, _
                  "CreateOrResizeConfiguredTable", _
                  "Could not determine first header column on sheet '" & ws.name & "'."
    End If

    lastCol = GetLastUsedColumnInRow(ws, headerRow)
    If lastCol = 0 Then
        Err.Raise vbObjectError + 1001, _
                  "CreateOrResizeConfiguredTable", _
                  "Could not determine last header column on sheet '" & ws.name & "'."
    End If

    lastRow = GetLastUsedRowInWorksheet(ws)
    If lastRow = 0 Then
        Err.Raise vbObjectError + 1002, _
                  "CreateOrResizeConfiguredTable", _
                  "Could not determine last used row on sheet '" & ws.name & "'."
    End If

    If lastRow < headerRow Then
        Err.Raise vbObjectError + 1003, _
                  "CreateOrResizeConfiguredTable", _
                  "Detected last row is above the header row on sheet '" & ws.name & "'."
    End If

    Set rngTable = ws.Range(ws.Cells(headerRow, firstCol), ws.Cells(lastRow, lastCol))

    '-------------------------------------------------------
    ' Validate table-name uniqueness across workbook
    '-------------------------------------------------------
    Set wbLo = FindListObjectByName(ws.Parent, tableName)

    If Not wbLo Is Nothing Then
        If Not wbLo.Parent Is ws Then
            Err.Raise vbObjectError + 1004, _
                      "CreateOrResizeConfiguredTable", _
                      "Table name '" & tableName & "' already exists on worksheet '" & _
                      wbLo.Parent.name & "'."
        Else
            ' Remove existing same-sheet table object but keep data
            wbLo.Unlist
        End If
    End If

    '-------------------------------------------------------
    ' Create fresh configured table
    '-------------------------------------------------------
    Set lo = ws.ListObjects.Add(SourceType:=xlSrcRange, _
                                Source:=rngTable, _
                                XlListObjectHasHeaders:=xlYes)

    lo.name = tableName

    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "CreateOrResizeConfiguredTable", Err.Description

End Sub

'-------------------------------------------------------
' ClearWorksheetAndTableFilters
'
' Purpose:
'   Removes active filters and worksheet AutoFilter state
'   before table resize / create operations.
'
' Notes:
'   - Clears existing ListObject filters
'   - Removes worksheet-level AutoFilter state
'   - Does not use AutoFilter toggle tricks
'-------------------------------------------------------
Private Sub ClearWorksheetAndTableFilters(ByVal ws As Worksheet)

    '-------------------------------------------------------
    ' Variable Declarations
    '-------------------------------------------------------
    Dim lo As ListObject

    On Error GoTo ErrHandler

    '-------------------------------------------------------
    ' Clear filters on existing tables
    '-------------------------------------------------------
    For Each lo In ws.ListObjects
        If Not lo.AutoFilter Is Nothing Then
            If lo.AutoFilter.FilterMode Then
                lo.AutoFilter.ShowAllData
            End If
        End If
    Next lo

    '-------------------------------------------------------
    ' Clear worksheet-level filters
    '-------------------------------------------------------
    If ws.FilterMode Then
        ws.ShowAllData
    End If

    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If

    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "ClearWorksheetAndTableFilters", Err.Description

End Sub


'-------------------------------------------------------
' GetOpenWorkbookByName
'
' Purpose:
'   Returns a reference to an open workbook matching the
'   supplied workbook name.
'
' Inputs:
'   - workbookName: Name of the workbook (e.g. "File.xlsx")
'
' Outputs:
'   - Returns Workbook object if found
'   - Returns Nothing if not found
'
' Requirements:
'   - Workbook must already be open in Excel
'
' Notes:
'   - Match is case-insensitive
'   - Does NOT open workbooks from disk
'   - Uses Application.Workbooks collection only
'-------------------------------------------------------
Public Function GetOpenWorkbookByName(ByVal workbookName As String) As Workbook

    '-------------------------------------------------------
    ' Variable Declarations
    '-------------------------------------------------------
    Dim wb As Workbook
    Dim searchName As String

    On Error GoTo ErrHandler

    '-------------------------------------------------------
    ' Validation
    '-------------------------------------------------------
    searchName = Trim$(workbookName)

    If Len(searchName) = 0 Then
        Exit Function
    End If

    '-------------------------------------------------------
    ' Main Processing Logic
    '-------------------------------------------------------
    For Each wb In Application.Workbooks
        If StrComp(wb.name, searchName, vbTextCompare) = 0 Then
            Set GetOpenWorkbookByName = wb
            Exit Function
        End If
    Next wb

    '-------------------------------------------------------
    ' Not found ? return Nothing
    '-------------------------------------------------------
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "GetOpenWorkbookByName", Err.Description

End Function

'-------------------------------------------------------
' BuildWorksheetTableMap
'
' Purpose:
'   Returns a worksheet-to-table mapping used by the
'   configured table build routine.
'
' Outputs:
'   - Scripting.Dictionary where:
'       Key   = Worksheet name
'       Item  = Variant array:
'               (0) = Target ListObject name
'               (1) = Header row number
'
' Requirements:
'   - Worksheet names must match the workbook exactly
'   - Table names must be unique across the workbook
'   - Header row numbers must be valid worksheet row numbers
'
' Notes:
'   - Dictionary compare mode is text compare so key
'     matching is case-insensitive
'   - This is a controlled configuration map and should
'     be updated here if worksheet names change
'-------------------------------------------------------
Public Function BuildWorksheetTableMap() As Object

    '-------------------------------------------------------
    ' Variable Declarations
    '-------------------------------------------------------
    Dim dict As Object   ' Scripting.Dictionary

    On Error GoTo ErrHandler

    '-------------------------------------------------------
    ' Initialise dictionary
    '-------------------------------------------------------
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    '-------------------------------------------------------
    ' Add worksheet-to-table mappings
    ' Format:
    '   dict.Add SheetName, Array(TableName, HeaderRowNumber)
    '-------------------------------------------------------
    dict.Add "Tracking Schedule", Array("tbl_Tracking", 6)
    dict.Add "Structural Installation", Array("tmp_SInstall", 2)
    dict.Add "Piping Installation", Array("tmp_PInstall", 2)
    dict.Add "Structural Material Take Off", Array("tmp_SMTO", 7)
    dict.Add "Piping Material Takeoff", Array("tmp_PMTO", 4)
    dict.Add "Piping Welding Traceability", Array("tbl_Trace_Pipe", 4)
    dict.Add "Structural Welding Traceability", Array("tbl_Trace_Structural", 4)

    '-------------------------------------------------------
    ' Return configured mapping
    '-------------------------------------------------------
    Set BuildWorksheetTableMap = dict
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "BuildWorksheetTableMap", Err.Description

End Function

Public Function GetHeaderRow(ByVal ws As Worksheet) As Long

    Dim f As Range

    On Error GoTo ErrHandler

    ' Find first used cell anywhere on the sheet
    Set f = ws.Cells.Find(What:="*", _
                          After:=ws.Cells(ws.Rows.Count, ws.Columns.Count), _
                          LookIn:=xlFormulas, _
                          LookAt:=xlPart, _
                          SearchOrder:=xlByRows, _
                          SearchDirection:=xlNext, _
                          matchCase:=False)

    If f Is Nothing Then
        GetHeaderRow = 0
    Else
        GetHeaderRow = f.Row
    End If

    Exit Function

ErrHandler:
    Err.Raise Err.Number, "GetHeaderRow", Err.Description

End Function

'-------------------------------------------------------
' GetLastUsedColumnInRow
'
' Purpose:
'   Returns the last used column number in the specified
'   worksheet row.
'
' Inputs:
'   - ws: Target worksheet
'   - rowNum: Row number to inspect
'
' Outputs:
'   - Returns 0 if the row contains no used cells
'   - Otherwise returns the last used column number
'-------------------------------------------------------
Public Function GetLastUsedColumnInRow(ByVal ws As Worksheet, _
                                       ByVal rowNum As Long) As Long

    '-------------------------------------------------------
    ' Variable Declarations
    '-------------------------------------------------------
    Dim f As Range

    On Error GoTo ErrHandler

    '-------------------------------------------------------
    ' Validation
    '-------------------------------------------------------
    If ws Is Nothing Then
        Err.Raise vbObjectError + 3010, _
                  "GetLastUsedColumnInRow", _
                  "Worksheet reference is Nothing."
    End If

    If rowNum < 1 Or rowNum > ws.Rows.Count Then
        Exit Function
    End If

    '-------------------------------------------------------
    ' Find last used cell in the row
    '-------------------------------------------------------
    Set f = ws.Rows(rowNum).Find(What:="*", _
                                 After:=ws.Cells(rowNum, 1), _
                                 LookIn:=xlFormulas, _
                                 LookAt:=xlPart, _
                                 SearchOrder:=xlByColumns, _
                                 SearchDirection:=xlPrevious, _
                                 matchCase:=False)

    If f Is Nothing Then
        GetLastUsedColumnInRow = 0
    Else
        GetLastUsedColumnInRow = f.Column
    End If

    Exit Function

ErrHandler:
    Err.Raise Err.Number, "GetLastUsedColumnInRow", Err.Description

End Function

'-------------------------------------------------------
' GetLastUsedRowInWorksheet
'
' Purpose:
'   Returns the last used row anywhere on the worksheet.
'
' Inputs:
'   - ws: Target worksheet
'
' Outputs:
'   - Returns 0 if the worksheet is blank
'   - Otherwise returns the last used row number
'-------------------------------------------------------
Public Function GetLastUsedRowInWorksheet(ByVal ws As Worksheet) As Long

    '-------------------------------------------------------
    ' Variable Declarations
    '-------------------------------------------------------
    Dim f As Range

    On Error GoTo ErrHandler

    '-------------------------------------------------------
    ' Validation
    '-------------------------------------------------------
    If ws Is Nothing Then
        Err.Raise vbObjectError + 3020, _
                  "GetLastUsedRowInWorksheet", _
                  "Worksheet reference is Nothing."
    End If

    '-------------------------------------------------------
    ' Find last used cell on worksheet
    '-------------------------------------------------------
    Set f = ws.Cells.Find(What:="*", _
                          After:=ws.Cells(1, 1), _
                          LookIn:=xlFormulas, _
                          LookAt:=xlPart, _
                          SearchOrder:=xlByRows, _
                          SearchDirection:=xlPrevious, _
                          matchCase:=False)

    If f Is Nothing Then
        GetLastUsedRowInWorksheet = 0
    Else
        GetLastUsedRowInWorksheet = f.Row
    End If

    Exit Function

ErrHandler:
    Err.Raise Err.Number, "GetLastUsedRowInWorksheet", Err.Description

End Function



'-------------------------------------------------------
' GetFirstUsedColumnInRow
'
' Purpose:
'   Returns the first used column number in the specified
'   row of a worksheet.
'
' Inputs:
'   - ws: Target worksheet
'   - rowNum: Row number to inspect
'
' Outputs:
'   - Returns 0 if the row contains no used cells
'   - Otherwise returns the first used column number
'
' Notes:
'   - Uses Find for reliability (handles blanks, formats,
'     and non-contiguous data)
'   - Looks in formulas so constants and formulas are
'     both treated as used cells
'-------------------------------------------------------
Public Function GetFirstUsedColumnInRow(ByVal ws As Worksheet, _
                                        ByVal rowNum As Long) As Long

    '-------------------------------------------------------
    ' Variable Declarations
    '-------------------------------------------------------
    Dim f As Range

    On Error GoTo ErrHandler

    '-------------------------------------------------------
    ' Validation
    '-------------------------------------------------------
    If ws Is Nothing Then
        Err.Raise vbObjectError + 3050, _
                  "GetFirstUsedColumnInRow", _
                  "Worksheet reference is Nothing."
    End If

    If rowNum < 1 Or rowNum > ws.Rows.Count Then
        Exit Function
    End If

    '-------------------------------------------------------
    ' Find first used cell in the row
    '-------------------------------------------------------
    Set f = ws.Rows(rowNum).Find(What:="*", _
                                 After:=ws.Cells(rowNum, ws.Columns.Count), _
                                 LookIn:=xlFormulas, _
                                 LookAt:=xlPart, _
                                 SearchOrder:=xlByColumns, _
                                 SearchDirection:=xlNext, _
                                 matchCase:=False)

    If f Is Nothing Then
        GetFirstUsedColumnInRow = 0
    Else
        GetFirstUsedColumnInRow = f.Column
    End If

    Exit Function

ErrHandler:
    Err.Raise Err.Number, "GetFirstUsedColumnInRow", Err.Description

End Function

'-------------------------------------------------------
' RemoveAllTablesFromWorksheet
'
' Purpose:
'   Removes all Excel table objects (ListObjects) from the
'   specified worksheet while preserving the underlying data.
'
' Notes:
'   - Uses Unlist, not Delete
'   - Keeps the worksheet values in place
'   - Removes the table object and its table name
'-------------------------------------------------------
Private Sub RemoveAllTablesFromWorksheet(ByVal ws As Worksheet)

    '-------------------------------------------------------
    ' Variable Declarations
    '-------------------------------------------------------
    Dim i As Long

    On Error GoTo ErrHandler

    If ws Is Nothing Then
        Err.Raise vbObjectError + 1200, _
                  "RemoveAllTablesFromWorksheet", _
                  "Worksheet reference is Nothing."
    End If

    '-------------------------------------------------------
    ' Remove all table objects but keep data
    '-------------------------------------------------------
    For i = ws.ListObjects.Count To 1 Step -1
        ws.ListObjects(i).Unlist
    Next i

    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "RemoveAllTablesFromWorksheet", Err.Description

End Sub


'=========================================================
' Returns a dictionary mapping:
'   Key   = Old Structural MTO header
'   Item  = New Combined MTO header
'
' Any old header not present in the dictionary is treated
' as "not mapped".
'=========================================================
Public Function GetStructuralMTO_To_CombinedMTO_Map() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    dict.CompareMode = vbTextCompare
    
    '-------------------------------
    ' Direct / logical mappings
    '-------------------------------
    dict("Workpack") = "Workpack"
    dict("Dwng No.") = "Ass Dwg No."
    dict("Rev") = "Rev"
    dict("Assembly No.") = "Assembly No."
    dict("Mark No.") = "Mark No."
    dict("Description") = "Description"
    dict("Profile") = "Profile"
    dict("Grade") = "Grade"
    dict("Length (mm)") = "Length (mm)"
    dict("Width (mm)") = "Width (mm)"
    dict("Quantity") = "Unit Quantity"
    'dict("TOTAL") = "Total Qty"
    dict("Area") = "Area (ea)"
    dict("Weight (KG)") = "Weight (KG) ea"
    dict("Supplier") = "Supplier"
    dict("PO No.") = "PO No."
    dict("Comments") = "Comments"
    dict("x") = "x"
    dict("Heat Number") = "Heat Number"
    dict("Promise Delivery Date") = "Promise Delivery Date"
    dict("Actual Delivery Date") = "Actual Delivery Date"
    dict("Delivery Docket") = "Delivery Docket"
    dict("TYPE") = "Type"
    
    Set GetStructuralMTO_To_CombinedMTO_Map = dict
End Function



'=========================================================
' Returns a dictionary mapping:
'   Key   = Old Piping MTO header
'   Item  = New Combined MTO header
'
' Any old header not present in the dictionary is treated
' as "not mapped".
'=========================================================
Public Function GetPipingMTO_To_CombinedMTO_Map() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    dict.CompareMode = vbTextCompare
    
    '-------------------------------
    ' Direct / logical mappings
    '-------------------------------
    dict("Workpack") = "Workpack"
    dict("Drawing Number") = "Part Dwg"
    dict("Drawing Number ") = "Part Dwg"   'covers trailing-space header
    dict("Rev") = "Rev"
    dict("Spool Number") = "Assembly No."
    dict("Item Ref") = "Mark No."
    dict("Size 1") = "Size 1"
    dict("Size 2") = "Size 2"
    dict("Schedule") = "Type"
    dict("Grade") = "Grade"
    dict("Description") = "Description"
    dict("Pipe Cut Length (mm)") = "Length (mm)"
    dict("Fitting Length (mm)") = "Fitting Length (mm)"
    dict("QTY") = "Unit Quantity"
    dict("Supplier") = "Supplier"
    dict("Purchase Order") = "PO No."
    dict("Delivery Date") = "Actual Delivery Date"
    dict("Delivery Docket") = "Delivery Docket"
    dict("Comments") = "Comments"
    dict("Total Length (m)") = "Length (m)"
    dict("SQM") = "Area (ea)"
    dict("KG") = "Weight (KG) ea"
    
    Set GetPipingMTO_To_CombinedMTO_Map = dict
End Function

