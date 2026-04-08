Attribute VB_Name = "modMTO"
Option Explicit

' Cache for stock lengths across runs (session-level)
Private dictStockCache As Object

Sub ClearFilter_MTO_Safe()
    Dim lo As ListObject
    On Error Resume Next
    
    Set lo = ThisWorkbook.Sheets("MTO").ListObjects("tbl_MTO")
    If Not lo Is Nothing Then
        If Not lo.AutoFilter Is Nothing Then lo.AutoFilter.ShowAllData
    End If
    
    On Error GoTo 0
End Sub

'===============================
'  SHEET PROTECTION ROUTINES
'===============================
Sub UnprotectSheet(ws As Worksheet, pwd As String)
    ws.Unprotect password:=pwd
End Sub

Sub ReprotectSheet(ws As Worksheet, pwd As String)
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
End Sub


Sub ExportFiltered_MTO()
    ExportFiltered_MTO_WithSubtotals
End Sub

Sub ExportFiltered_MTO_WithSubtotals()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim rngVisibleData As Range
    Dim rngCopy As Range
    Dim r As Range
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim fName As Variant
    Dim colCount As Long
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("MTO")
    Set lo = ws.ListObjects("tbl_MTO")
    
    ' Make sure table has data
    If lo.DataBodyRange Is Nothing Then
        MsgBox "tbl_MTO has no data rows.", vbExclamation
        Exit Sub
    End If
    
    ' Build visible-data range manually (ignore SpecialCells)
    For Each r In lo.DataBodyRange.Rows
        If Not r.EntireRow.Hidden Then
            If rngVisibleData Is Nothing Then
                Set rngVisibleData = r
            Else
                Set rngVisibleData = Union(rngVisibleData, r)
            End If
        End If
    Next r
    
    If rngVisibleData Is Nothing Then
        MsgBox "No visible data rows to export for tbl_MTO.", vbExclamation
        Exit Sub
    End If
    
    ' Add header row to the export range
    Set rngCopy = Union(lo.HeaderRowRange, rngVisibleData)
    
    colCount = lo.Range.Columns.Count
    
    ' Create new workbook
    Set wbNew = Workbooks.Add(xlWBATWorksheet)
    Set wsNew = wbNew.Sheets(1)
    wsNew.Name = "Export"
    
    ' Copy values + formats
    rngCopy.Copy
    With wsNew.Range("A1")
        .PasteSpecial xlPasteValuesAndNumberFormats
        .PasteSpecial xlPasteFormats
    End With
    Application.CutCopyMode = False
    
    ' Copy column widths
    For i = 1 To colCount
        wsNew.Columns(i).ColumnWidth = lo.Range.Columns(i).ColumnWidth
    Next i
    
    ' Apply filters on the new sheet
    wsNew.Range("A1").CurrentRegion.AutoFilter
    
    ' Freeze top row (adjust column to taste)
    wsNew.Activate
    wsNew.Range("F2").Select
    ActiveWindow.FreezePanes = True
    
    ' -------------------------
    ' ADD SUBTOTAL ROW
    ' -------------------------
    lastRow = wsNew.Cells(wsNew.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Label
    wsNew.Cells(lastRow, "A").Value = "SUBTOTALS:"
    
    ' Count in column E
    wsNew.Cells(lastRow, "E").Formula = "=SUBTOTAL(3,E2:E" & lastRow - 1 & ")"
    
    ' Subtotal sums in T, U, V
    wsNew.Cells(lastRow, "T").Formula = "=SUBTOTAL(9,T2:T" & lastRow - 1 & ")"
    wsNew.Cells(lastRow, "U").Formula = "=SUBTOTAL(9,U2:U" & lastRow - 1 & ")"
    wsNew.Cells(lastRow, "V").Formula = "=SUBTOTAL(9,V2:V" & lastRow - 1 & ")"
    
    ' Format subtotal row
    With wsNew.Rows(lastRow)
        .Font.Bold = True
        .Interior.Color = RGB(220, 220, 220) ' light grey
    End With
    
    ' Optional: tidy up
    ' wsNew.UsedRange.Columns.AutoFit
    
    ' Ask where to save
    fName = Application.GetSaveAsFilename( _
                InitialFileName:="MTO Export.xlsx", _
                FileFilter:="Excel Files (*.xlsx), *.xlsx")
    
    If fName = False Then Exit Sub  ' Cancelled
    
    wbNew.SaveAs fName, FileFormat:=xlOpenXMLWorkbook
    MsgBox "Filtered MTO exported with subtotals.", vbInformation
End Sub




'==== ENTRY POINTS FOR BUTTONS ======================================

Public Sub NestPipingCuts()
    NestCuts_Fast "PIPE"
End Sub

Public Sub NestStructuralCuts()
    NestCuts_Fast "STRUCT"
End Sub

'-------------------------------------------------------
'-------------------------------------------------------
' NestCuts_Fast
' – Builds a nested cutting layout from tbl_MTO
' – Supports PIPE and STRUCT modes
' – Optionally scopes the run to the selected Workpack
' – Uses nestable profiles from tbl_MatSpec
' – Resolves stock length per Grade / Size 1 / Profile group
' – Expands quantities into individual pieces
' – Performs a First-Fit Decreasing pass with secondary backfill
' – Writes detail and summary output to the "Nesting" worksheet
'
' Requirements / Dependencies:
' – AppGuard_Begin / AppGuard_End
' – SheetGuard_Begin / SheetGuard_End
' – GetWorksheetOfTable
' – GetNestableProfiles
' – ResolveStockLength
' – QuickSort2D
' – Global password variable: pwd
' – Global / module-level cache: dictStockCache
'
' Notes:
' – Existing behaviour is intentionally preserved.
' – This procedure always clears existing table filters on tbl_MTO.
' – If a selected cell is within tbl_MTO[Workpack], the run is scoped
'   to that Workpack only.
'-------------------------------------------------------
'-------------------------------------------------------
Private Sub NestCuts_Fast(ByVal mode As String)

    Const PROC_NAME As String = "NestCuts_Fast"

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim outWS As Worksheet

    Dim shGuardMTO As TSheetGuardState
    Dim shGuardOut As TSheetGuardState
    Dim mtoGuardOpen As Boolean
    Dim outGuardOpen As Boolean

    '-------------------------------------------------------
    ' Required tbl_MTO column indexes
    '-------------------------------------------------------
    Dim colType As Long
    Dim colGrade As Long
    Dim colSize1 As Long
    Dim colProfile As Long
    Dim colAssDwg As Long
    Dim colLength As Long
    Dim colTotalQty As Long
    Dim colPipePrep As Long
    Dim colMarkNo As Long
    Dim colDisc As Long
    Dim colWorkpack As Long

    '-------------------------------------------------------
    ' Grouping / lookup structures
    '-------------------------------------------------------
    Dim dictGroups As Object
    Dim dictNestableProfiles As Object
    Dim key As String

    '-------------------------------------------------------
    ' Input / user settings
    '-------------------------------------------------------
    Dim cutAllowance As Double
    Dim cutInput As Variant
    Dim cutText As String

    '-------------------------------------------------------
    ' Source data and row-level working variables
    '-------------------------------------------------------
    Dim dataArr As Variant
    Dim nRows As Long
    Dim totalPieces As Long

    Dim i As Long
    Dim q As Long

    Dim thisType As String
    Dim thisDisc As String
    Dim thisLength As Double
    Dim thisUsed As Double
    Dim thisQty As Long

    Dim grade As Variant
    Dim size1 As Variant
    Dim profile As Variant
    Dim profileKey As String
    Dim assDwg As Variant
    Dim pipePrep As Variant
    Dim markNo As Variant
    Dim thisWorkpack As Variant

    '-------------------------------------------------------
    ' Optional Workpack filter based on current selection
    '-------------------------------------------------------
    Dim filterWorkpack As Variant
    Dim haveFilterWorkpack As Boolean

    '-------------------------------------------------------
    ' Per-group storage object
    ' grp(0)  = Grade
    ' grp(1)  = Size 1
    ' grp(2)  = Profile
    ' grp(3)  = unused / reserved
    ' grp(4)  = unused / reserved
    ' grp(5)  = Collection of actual lengths
    ' grp(6)  = Collection of mark numbers
    ' grp(7)  = total piece count in group
    ' grp(8)  = Collection of used lengths (length + kerf)
    ' grp(9)  = Collection of pipe prep values
    ' grp(10) = stock length
    ' grp(11) = Collection of assembly drawing numbers
    ' grp(12) = Collection of workpacks
    '-------------------------------------------------------
    Dim grp As Variant
    Dim colLen As Collection
    Dim colMark As Collection
    Dim colUsed As Collection
    Dim colPrep As Collection
    Dim colAss As Collection
    Dim colWork As Collection
    Dim stdLen As Double

    '-------------------------------------------------------
    ' Per-group processing arrays
    '-------------------------------------------------------
    Dim k As Variant
    Dim n As Long

    Dim arrPiecesActual() As Double
    Dim arrPiecesUsed() As Double
    Dim arrMarks() As Variant
    Dim arrPrep() As Variant
    Dim arrAss() As Variant
    Dim arrWork() As Variant

    ' arr2D columns:
    ' 1 = used length
    ' 2 = mark no
    ' 3 = actual length
    ' 4 = pipe prep
    ' 5 = ass dwg
    ' 6 = workpack
    Dim arr2D() As Variant

    '-------------------------------------------------------
    ' Nesting / stick allocation
    '-------------------------------------------------------
    Dim stickRem() As Double
    Dim stickCount As Long
    Dim assignStick() As Long
    Dim s As Long
    Dim placed As Boolean

    Dim s1 As Long
    Dim s2 As Long

    Dim stickHasPiece() As Boolean
    Dim newStickNum() As Long
    Dim newStickRem() As Double
    Dim newCount As Long

    '-------------------------------------------------------
    ' Output arrays
    '-------------------------------------------------------
    Dim outArr As Variant
    Dim outIdx As Long

    Dim expandedArr() As Variant
    Dim trimmedArr() As Variant
    Dim newIdx As Long

    '-------------------------------------------------------
    ' Summary output arrays
    '-------------------------------------------------------
    Dim summaryArr As Variant
    Dim summaryIdx As Long
    Dim groupCount As Long
    Dim totalStockLen As Double
    Dim totalWaste As Double
    Dim wastePct As Double
    Dim stickIdx As Long
    Dim usedOnStick As Double
    Dim i2 As Long
    Dim pieceCount As Long

    '-------------------------------------------------------
    ' Sorting / formatting variables
    '-------------------------------------------------------
    Dim r1 As Long
    Dim r2 As Long
    Dim tmp(1 To 12) As Variant

    Dim curStick As Variant
    Dim prevStick As Variant
    Dim lastStick As Variant
    Dim runLen As Double

    Dim r As Long
    Dim c As Long
    Dim lastRow As Long
    Dim msgScope As String

    ' Formatting objects for alternating stick shading
    Dim rngCF As Range
    Dim fcOdd As FormatCondition
    Dim fcEven As FormatCondition
    Dim rr As Long
    Dim currStick2 As Variant
    Dim prevSt As Variant

    On Error GoTo ErrHandler

    '-------------------------------------------------------
    ' Application guard
    '-------------------------------------------------------
    AppGuard_Begin

    '-------------------------------------------------------
    ' Locate source worksheet and table
    '-------------------------------------------------------
    Set ws = GetWorksheetOfTable(ThisWorkbook, "tbl_MTO")
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1000, PROC_NAME, _
                  "Could not locate worksheet for table 'tbl_MTO'."
    End If

    Set lo = ws.ListObjects("tbl_MTO")
    If lo Is Nothing Then
        Err.Raise vbObjectError + 1001, PROC_NAME, _
                  "Could not locate table 'tbl_MTO'."
    End If

    '-------------------------------------------------------
    ' Release protection on source sheet before working with
    ' filters / reading / downstream worksheet operations.
    ' Guard state is always restored in SafeExit.
    '-------------------------------------------------------
    shGuardMTO = SheetGuard_Begin(ws, pwd)
    mtoGuardOpen = True

    '-------------------------------------------------------
    ' Always clear any active filters on tbl_MTO.
    ' This preserves existing behaviour and ensures the array
    ' load reads the full table, not just visible rows.
    '-------------------------------------------------------
    On Error Resume Next
    If Not lo.AutoFilter Is Nothing Then
        lo.AutoFilter.ShowAllData
    End If
    On Error GoTo ErrHandler

    '-------------------------------------------------------
    ' Resolve required column indexes up front.
    ' Any missing header should fail fast and clearly.
    '-------------------------------------------------------
    colAssDwg = lo.ListColumns("Ass Dwg No.").Index
    colType = lo.ListColumns("Type").Index
    colGrade = lo.ListColumns("Grade").Index
    colSize1 = lo.ListColumns("Size 1").Index
    colProfile = lo.ListColumns("Profile").Index
    colLength = lo.ListColumns("Length (mm)").Index
    colTotalQty = lo.ListColumns("Total Qty").Index
    colPipePrep = lo.ListColumns("Pipe Prep").Index
    colMarkNo = lo.ListColumns("Mark No.").Index
    colDisc = lo.ListColumns("Discipline").Index
    colWorkpack = lo.ListColumns("Workpack").Index

    '-------------------------------------------------------
    ' Optional Workpack scoping:
    ' If the current selection intersects tbl_MTO[Workpack],
    ' use the selected value as a filter for this run.
    '
    ' Existing behaviour is preserved:
    ' – only a non-blank selected value activates the filter
    ' – any error while evaluating Selection is ignored
    '-------------------------------------------------------
    On Error Resume Next
    If Not lo.ListColumns("Workpack").DataBodyRange Is Nothing Then
        If Not Intersect(Selection, lo.ListColumns("Workpack").DataBodyRange) Is Nothing Then
            filterWorkpack = Selection.Value
            If Len(Trim$(CStr(filterWorkpack))) > 0 Then
                haveFilterWorkpack = True
            End If
        End If
    End If
    On Error GoTo ErrHandler

    '-------------------------------------------------------
    ' Load profiles flagged as nestable from tbl_MatSpec.
    ' If the dependency fails or returns no allowed profiles,
    ' exit cleanly.
    '-------------------------------------------------------
    Set dictNestableProfiles = GetNestableProfiles()
    If dictNestableProfiles Is Nothing Then GoTo SafeExit

    If dictNestableProfiles.Count = 0 Then
        MsgBox "No profiles marked 'Y' for nesting in tbl_MatSpec.", vbExclamation
        GoTo SafeExit
    End If

    '-------------------------------------------------------
    ' Prompt for global cut allowance (kerf)
    ' Default = 3 mm
    ' Blank   = use default
    ' Zero    = valid (e.g. cropper)
    ' Cancel  = exit
    '-------------------------------------------------------
    cutInput = Application.InputBox( _
        Prompt:="Enter cut allowance per cut (mm)." & vbCrLf & _
                "Leave BLANK to use default 3 mm." & vbCrLf & _
                "Enter 0 if using a cropper.", _
        Title:="Cut Allowance (Kerf)", _
        Type:=2)

    ' Application.InputBox returns Boolean False on cancel
    If TypeName(cutInput) = "Boolean" Then
        If cutInput = False Then GoTo SafeExit
    End If

    cutText = Trim$(CStr(cutInput))

    If cutText = "" Then
        cutAllowance = 3
    ElseIf IsNumeric(cutText) Then
        cutAllowance = CDbl(cutText)

        If cutAllowance < 0 Then
            MsgBox "Cut allowance cannot be negative.", vbExclamation
            GoTo SafeExit
        End If
    Else
        MsgBox "Cut allowance must be a number.", vbExclamation
        GoTo SafeExit
    End If

    '-------------------------------------------------------
    ' Ensure stock-length cache exists.
    ' Existing behaviour preserved.
    '-------------------------------------------------------
    If dictStockCache Is Nothing Then
        Set dictStockCache = CreateObject("Scripting.Dictionary")
    End If

    '-------------------------------------------------------
    ' Load tbl_MTO body into memory.
    ' This is the main performance-saving step.
    '-------------------------------------------------------
    If lo.DataBodyRange Is Nothing Then
        MsgBox "tbl_MTO contains no data rows.", vbInformation
        GoTo SafeExit
    End If

    dataArr = lo.DataBodyRange.Value
    nRows = UBound(dataArr, 1)

    '-------------------------------------------------------
    ' First pass:
    ' Build grouped collections of all qualifying pieces.
    '
    ' Group key = Grade|Size 1|Profile
    '
    ' A source row may contribute multiple pieces based on
    ' Total Qty, so each piece is expanded into the collections.
    '-------------------------------------------------------
    Set dictGroups = CreateObject("Scripting.Dictionary")
    totalPieces = 0

    For i = 1 To nRows

        thisType = UCase$(Trim$(CStr(dataArr(i, colType))))
        thisDisc = UCase$(Trim$(CStr(dataArr(i, colDisc))))
        thisWorkpack = dataArr(i, colWorkpack)

        '---------------------------------------------------
        ' Optional Workpack filter
        '---------------------------------------------------
        If haveFilterWorkpack Then
            If CStr(thisWorkpack) <> CStr(filterWorkpack) Then GoTo NextRow
        End If

        '---------------------------------------------------
        ' Mode filter
        ' PIPE   -> Type must be PIPE
        ' STRUCT -> Discipline must be STRUCTURAL
        '---------------------------------------------------
        If mode = "PIPE" Then
            If thisType <> "PIPE" Then GoTo NextRow
        ElseIf mode = "STRUCT" Then
            If thisDisc <> "STRUCTURAL" Then GoTo NextRow
        End If

        '---------------------------------------------------
        ' Quantity / length validation
        ' Existing behaviour: non-positive rows are skipped
        '---------------------------------------------------
        thisLength = CDbl(dataArr(i, colLength))
        thisQty = CLng(val(dataArr(i, colTotalQty)))

        If thisLength <= 0 Or thisQty <= 0 Then GoTo NextRow

        grade = dataArr(i, colGrade)
        size1 = dataArr(i, colSize1)
        profile = dataArr(i, colProfile)
        assDwg = dataArr(i, colAssDwg)
        pipePrep = dataArr(i, colPipePrep)
        markNo = dataArr(i, colMarkNo)

        '---------------------------------------------------
        ' Only process profiles that are marked nestable
        ' in tbl_MatSpec.
        '---------------------------------------------------
        profileKey = UCase$(Trim$(CStr(profile)))
        If Not dictNestableProfiles.Exists(profileKey) Then GoTo NextRow

        '---------------------------------------------------
        ' Group by Grade|Size1|Profile
        '---------------------------------------------------
        key = CStr(grade) & "|" & CStr(size1) & "|" & CStr(profile)

        If dictGroups.Exists(key) Then
            grp = dictGroups(key)
            stdLen = grp(10)
        Else
            '------------------------------------------------
            ' Resolve stock length once per group.
            ' Dependency may prompt / cache / lookup based on
            ' your existing ResolveStockLength implementation.
            '------------------------------------------------
            stdLen = ResolveStockLength(key, profile, grade, size1, mode)
            If stdLen <= 0 Then GoTo SafeExit

            ReDim grp(0 To 12)
            grp(0) = grade
            grp(1) = size1
            grp(2) = profile
            grp(3) = Empty
            grp(4) = Empty
            Set grp(5) = New Collection   ' actual lengths
            Set grp(6) = New Collection   ' mark numbers
            grp(7) = 0                    ' piece count
            Set grp(8) = New Collection   ' used lengths (len + kerf)
            Set grp(9) = New Collection   ' pipe prep
            grp(10) = stdLen              ' stock length
            Set grp(11) = New Collection  ' ass dwg no.
            Set grp(12) = New Collection  ' workpack

            dictGroups.Add key, grp
        End If

        '---------------------------------------------------
        ' Used length includes kerf allowance per piece.
        ' This preserves existing logic exactly.
        '---------------------------------------------------
        thisUsed = thisLength + cutAllowance

        '---------------------------------------------------
        ' Hard stop if any single piece cannot fit in the
        ' nominated stock length.
        '---------------------------------------------------
        If thisUsed > stdLen Then
            MsgBox "Piece length + cut allowance exceeds stock length:" & vbCrLf & _
                   "Piece length = " & thisLength & " mm, Cut = " & cutAllowance & " mm" & vbCrLf & _
                   "Total = " & thisUsed & " mm, Stock = " & stdLen & " mm" & vbCrLf & _
                   "Row: " & (lo.DataBodyRange.Row + i - 1), vbCritical
            GoTo SafeExit
        End If

        '---------------------------------------------------
        ' Expand quantity into individual pieces
        '---------------------------------------------------
        Set colLen = grp(5)
        Set colMark = grp(6)
        Set colUsed = grp(8)
        Set colPrep = grp(9)
        Set colAss = grp(11)
        Set colWork = grp(12)

        For q = 1 To thisQty
            colLen.Add thisLength
            colMark.Add markNo
            colUsed.Add thisUsed
            colPrep.Add pipePrep
            colAss.Add assDwg
            colWork.Add thisWorkpack
        Next q

        grp(7) = grp(7) + thisQty
        dictGroups(key) = grp

        totalPieces = totalPieces + thisQty

NextRow:
    Next i

    '-------------------------------------------------------
    ' Nothing qualified after all filters / validations
    '-------------------------------------------------------
    If totalPieces = 0 Then
        If haveFilterWorkpack Then
            MsgBox "No qualifying rows found to nest for mode " & mode & _
                   " and Workpack " & CStr(filterWorkpack) & ".", vbInformation
        Else
            MsgBox "No qualifying rows found to nest for mode " & mode & ".", vbInformation
        End If
        GoTo SafeExit
    End If

    '-------------------------------------------------------
    ' Prepare output arrays
    '-------------------------------------------------------
    ReDim outArr(1 To totalPieces, 1 To 12)
    outIdx = 1

    groupCount = dictGroups.Count
    If groupCount > 0 Then
        ReDim summaryArr(1 To groupCount, 1 To 8)
        summaryIdx = 1
    End If

    '-------------------------------------------------------
    ' Process each group independently
    '
    ' Steps:
    ' 1) Move collection data into arrays
    ' 2) Sort descending by used length (QuickSort2D)
    ' 3) First-Fit Decreasing assignment to sticks
    ' 4) Backfill/compress later sticks into earlier sticks
    ' 5) Renumber used sticks compactly
    ' 6) Write detail rows and group summary
    '-------------------------------------------------------
    For Each k In dictGroups.keys

        grp = dictGroups(k)

        Set colLen = grp(5)
        Set colMark = grp(6)
        Set colUsed = grp(8)
        Set colPrep = grp(9)
        Set colAss = grp(11)
        Set colWork = grp(12)

        stdLen = grp(10)
        n = grp(7)

        If n = 0 Then GoTo NextGroup

        ReDim arrPiecesActual(1 To n)
        ReDim arrPiecesUsed(1 To n)
        ReDim arrMarks(1 To n)
        ReDim arrPrep(1 To n)
        ReDim arrAss(1 To n)
        ReDim arrWork(1 To n)

        For i = 1 To n
            arrPiecesActual(i) = CDbl(colLen(i))
            arrPiecesUsed(i) = CDbl(colUsed(i))
            arrMarks(i) = colMark(i)
            arrPrep(i) = colPrep(i)
            arrAss(i) = colAss(i)
            arrWork(i) = colWork(i)
        Next i

        '---------------------------------------------------
        ' Build sortable 2D array
        '---------------------------------------------------
        ReDim arr2D(1 To n, 1 To 6)
        For i = 1 To n
            arr2D(i, 1) = arrPiecesUsed(i)
            arr2D(i, 2) = arrMarks(i)
            arr2D(i, 3) = arrPiecesActual(i)
            arr2D(i, 4) = arrPrep(i)
            arr2D(i, 5) = arrAss(i)
            arr2D(i, 6) = arrWork(i)
        Next i

        '---------------------------------------------------
        ' Sort by used length descending.
        ' Existing dependency / behaviour preserved.
        '---------------------------------------------------
        If n > 1 Then QuickSort2D arr2D, 1, n

        '---------------------------------------------------
        ' First-Fit Decreasing (FFD)
        ' – Each piece is placed into the first stick with
        '   enough remainder
        ' – If none fit, a new stick is created
        '---------------------------------------------------
        ReDim stickRem(1 To 1)
        stickCount = 0
        ReDim assignStick(1 To n)

        For i = 1 To n
            thisUsed = arr2D(i, 1)
            placed = False

            For s = 1 To stickCount
                If stickRem(s) >= thisUsed Then
                    stickRem(s) = stickRem(s) - thisUsed
                    assignStick(i) = s
                    placed = True
                    Exit For
                End If
            Next s

            If Not placed Then
                stickCount = stickCount + 1

                If stickCount = 1 Then
                    stickRem(1) = stdLen - thisUsed
                Else
                    ReDim Preserve stickRem(1 To stickCount)
                    stickRem(stickCount) = stdLen - thisUsed
                End If

                assignStick(i) = stickCount
            End If
        Next i

        '---------------------------------------------------
        ' Backfill / compression pass
        ' Try moving pieces from later sticks into earlier
        ' sticks if they fit.
        '
        ' This preserves your existing behaviour exactly.
        '---------------------------------------------------
        For s1 = 1 To stickCount - 1
            For s2 = s1 + 1 To stickCount
                For i = 1 To n
                    If assignStick(i) = s2 Then
                        thisUsed = arr2D(i, 1)

                        If stickRem(s1) >= thisUsed Then
                            stickRem(s1) = stickRem(s1) - thisUsed
                            stickRem(s2) = stickRem(s2) + thisUsed
                            assignStick(i) = s1
                        End If
                    End If
                Next i
            Next s2
        Next s1

        '---------------------------------------------------
        ' Compress stick numbering so only sticks that still
        ' contain pieces are numbered sequentially.
        '---------------------------------------------------
        ReDim stickHasPiece(1 To stickCount)
        For i = 1 To n
            stickHasPiece(assignStick(i)) = True
        Next i

        ReDim newStickNum(1 To stickCount)
        newCount = 0

        For s = 1 To stickCount
            If stickHasPiece(s) Then
                newCount = newCount + 1
                newStickNum(s) = newCount
            End If
        Next s

        ReDim newStickRem(1 To newCount)
        For s = 1 To newCount
            newStickRem(s) = stdLen
        Next s

        For i = 1 To n
            s = newStickNum(assignStick(i))
            assignStick(i) = s
            newStickRem(s) = newStickRem(s) - arr2D(i, 1)
        Next i

        stickCount = newCount
        stickRem = newStickRem

        '---------------------------------------------------
        ' Write group detail rows into outArr
        '
        ' Output columns:
        ' 1  Workpack
        ' 2  Ass Dwg No.
        ' 3  Profile
        ' 4  Mark No.
        ' 5  Length (mm)
        ' 6  Pipe Prep
        ' 7  Grade
        ' 8  Size 1
        ' 9  Stick No
        ' 10 Running Length (filled later)
        ' 11 Stock Length
        ' 12 Waste on Stick
        '---------------------------------------------------
        For i = 1 To n
            s = assignStick(i)

            outArr(outIdx, 1) = arr2D(i, 6)   ' Workpack
            outArr(outIdx, 2) = arr2D(i, 5)   ' Ass Dwg No.
            outArr(outIdx, 3) = grp(2)        ' Profile
            outArr(outIdx, 4) = arr2D(i, 2)   ' Mark No.
            outArr(outIdx, 5) = arr2D(i, 3)   ' Length (mm)
            outArr(outIdx, 6) = arr2D(i, 4)   ' Pipe Prep
            outArr(outIdx, 7) = grp(0)        ' Grade
            outArr(outIdx, 8) = grp(1)        ' Size 1
            outArr(outIdx, 9) = s             ' Stick No
            outArr(outIdx, 10) = 0            ' Running length placeholder
            outArr(outIdx, 11) = stdLen       ' Stock Length
            outArr(outIdx, 12) = stickRem(s)  ' Waste on Stick

            outIdx = outIdx + 1
        Next i

        '---------------------------------------------------
        ' Build summary row for this group
        '
        ' Note:
        ' Waste is recalculated based on actual piece lengths
        ' plus kerf between pieces on a stick.
        ' This preserves existing summary logic.
        '---------------------------------------------------
        If groupCount > 0 Then
            totalStockLen = stickCount * stdLen
            totalWaste = 0

            For stickIdx = 1 To stickCount
                pieceCount = 0
                usedOnStick = 0

                For i2 = 1 To n
                    If assignStick(i2) = stickIdx Then
                        pieceCount = pieceCount + 1
                        usedOnStick = usedOnStick + arr2D(i2, 3)
                    End If
                Next i2

                If pieceCount > 1 Then
                    usedOnStick = usedOnStick + cutAllowance * (pieceCount - 1)
                End If

                totalWaste = totalWaste + (stdLen - usedOnStick)
            Next stickIdx

            If totalStockLen > 0 Then
                wastePct = totalWaste / totalStockLen
            Else
                wastePct = 0
            End If

            summaryArr(summaryIdx, 1) = grp(0)
            summaryArr(summaryIdx, 2) = grp(1)
            summaryArr(summaryIdx, 3) = grp(2)
            summaryArr(summaryIdx, 4) = stdLen
            summaryArr(summaryIdx, 5) = stickCount
            summaryArr(summaryIdx, 6) = totalStockLen
            summaryArr(summaryIdx, 7) = totalWaste
            summaryArr(summaryIdx, 8) = wastePct

            summaryIdx = summaryIdx + 1
        End If

NextGroup:
    Next k

    '-------------------------------------------------------
    ' Sort detail output by:
    ' 1) Grade
    ' 2) Size 1
    ' 3) Stick No
    '
    ' Existing nested-loop sort retained to preserve
    ' behaviour exactly.
    '-------------------------------------------------------
    For r1 = 1 To totalPieces - 1
        For r2 = r1 + 1 To totalPieces

            If (CStr(outArr(r1, 7)) > CStr(outArr(r2, 7))) _
            Or (CStr(outArr(r1, 7)) = CStr(outArr(r2, 7)) And CStr(outArr(r1, 8)) > CStr(outArr(r2, 8))) _
            Or (CStr(outArr(r1, 7)) = CStr(outArr(r2, 7)) And CStr(outArr(r1, 8)) = CStr(outArr(r2, 8)) And CLng(outArr(r1, 9)) > CLng(outArr(r2, 9))) Then

                For i = 1 To 12
                    tmp(i) = outArr(r1, i)
                    outArr(r1, i) = outArr(r2, i)
                    outArr(r2, i) = tmp(i)
                Next i
            End If

        Next r2
    Next r1

    '-------------------------------------------------------
    ' Insert a blank output row between stick changes.
    ' This is for readability on the Nesting sheet.
    '
    ' Existing behaviour preserved by creating a larger
    ' temporary array and then trimming it back.
    '-------------------------------------------------------
    If totalPieces > 0 Then
        ReDim expandedArr(1 To totalPieces * 2, 1 To 12)
        newIdx = 0

        For r = 1 To totalPieces
            curStick = outArr(r, 9)

            If r > 1 Then
                prevStick = outArr(r - 1, 9)
                If curStick <> prevStick Then
                    newIdx = newIdx + 1   ' blank separator row
                End If
            End If

            newIdx = newIdx + 1
            For c = 1 To 12
                expandedArr(newIdx, c) = outArr(r, c)
            Next c
        Next r

        ReDim trimmedArr(1 To newIdx, 1 To 12)
        For r = 1 To newIdx
            For c = 1 To 12
                trimmedArr(r, c) = expandedArr(r, c)
            Next c
        Next r

        outArr = trimmedArr
        totalPieces = newIdx
    End If

    '-------------------------------------------------------
    ' Calculate running length within each stick.
    '
    ' Logic:
    ' – first item on a stick = its own length
    ' – subsequent items add kerf + length
    ' – blank separator rows stay blank
    '-------------------------------------------------------
    If totalPieces > 0 Then
        lastStick = vbNullString
        runLen = 0
        pieceCount = 0

        For r = 1 To totalPieces
            curStick = outArr(r, 9)

            If IsEmpty(curStick) Or curStick = "" Then
                outArr(r, 10) = Empty
            Else
                If curStick <> lastStick Then
                    pieceCount = 1

                    If IsNumeric(outArr(r, 5)) Then
                        runLen = CDbl(outArr(r, 5))
                    Else
                        runLen = 0
                    End If

                    lastStick = curStick
                Else
                    pieceCount = pieceCount + 1

                    If IsNumeric(outArr(r, 5)) Then
                        runLen = runLen + cutAllowance + CDbl(outArr(r, 5))
                    End If
                End If

                outArr(r, 10) = runLen
            End If
        Next r
    End If

    '-------------------------------------------------------
    ' Locate or create output sheet
    '-------------------------------------------------------
    On Error Resume Next
    Set outWS = ThisWorkbook.Worksheets("Nesting")
    On Error GoTo ErrHandler

    If outWS Is Nothing Then
        Set outWS = ThisWorkbook.Worksheets.Add(After:=ws)
        outWS.Name = "Nesting"
    End If

    '-------------------------------------------------------
    ' Release protection on output sheet before clearing /
    ' writing / formatting.
    '-------------------------------------------------------
    shGuardOut = SheetGuard_Begin(outWS, pwd)
    outGuardOpen = True

    '-------------------------------------------------------
    ' Clear and rebuild output sheet contents
    '-------------------------------------------------------
    outWS.Cells.Clear

    With outWS

        '---------------------------------------------------
        ' Detail section headers
        '---------------------------------------------------
        .Range("A1").Value = "Workpack"
        .Range("B1").Value = "Ass Dwg No."
        .Range("C1").Value = "Profile"
        .Range("D1").Value = "Mark No."
        .Range("E1").Value = "Length (mm)"
        .Range("F1").Value = "Pipe Prep"
        .Range("G1").Value = "Grade"
        .Range("H1").Value = "Size 1"
        .Range("I1").Value = "Stick No"
        .Range("J1").Value = "Running Length (mm)"
        .Range("K1").Value = "Stock Length (mm)"
        .Range("L1").Value = "Waste on Stick (mm)"
        .Range("M1").Value = "Heat No."

        If totalPieces > 0 Then
            .Range("A2").Resize(totalPieces, 12).Value = outArr
        End If

        '---------------------------------------------------
        ' Summary section headers
        '---------------------------------------------------
        .Range("N1").Value = "Grade"
        .Range("O1").Value = "Size 1"
        .Range("P1").Value = "Profile"
        .Range("Q1").Value = "Stock Length (mm)"
        .Range("R1").Value = "No. of Sticks"
        .Range("S1").Value = "Total Stock Length (mm)"
        .Range("T1").Value = "Total Waste (mm)"
        .Range("U1").Value = "Waste (%)"

        If groupCount > 0 And summaryIdx > 1 Then
            .Range("N2").Resize(summaryIdx - 1, 8).Value = summaryArr
            .Range("U2:U" & summaryIdx).NumberFormat = "0.0%"
        End If

        '---------------------------------------------------
        ' Determine last used row of detail output.
        ' Minimum = 1 to keep print setup safe.
        '---------------------------------------------------
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        If lastRow < 1 Then lastRow = 1

        '---------------------------------------------------
        ' Detail formatting
        ' – alternating shading by odd/even stick number
        ' – light grid borders
        ' – heavy top/bottom borders per stick block
        '---------------------------------------------------
        If totalPieces > 0 Then

            Set rngCF = .Range("A2:M" & lastRow)
            rngCF.FormatConditions.Delete

            Set fcOdd = rngCF.FormatConditions.Add( _
                Type:=xlExpression, _
                Formula1:="=AND($I2<>"""",MOD($I2,2)=1)")
            fcOdd.Interior.Color = RGB(235, 235, 235)

            Set fcEven = rngCF.FormatConditions.Add( _
                Type:=xlExpression, _
                Formula1:="=AND($I2<>"""",MOD($I2,2)=0)")
            fcEven.Interior.Pattern = xlNone

            With .Range("A2:M" & lastRow).Borders
                .LineStyle = xlContinuous
                .Weight = xlHairline
                .Color = RGB(200, 200, 200)
            End With

            prevSt = vbNullString

            For rr = 2 To lastRow
                currStick2 = .Cells(rr, "I").Value

                If currStick2 <> "" Then
                    If rr = 2 Or currStick2 <> prevSt Then
                        With .Range("A" & rr & ":M" & rr).Borders(xlTop)
                            .LineStyle = xlContinuous
                            .Weight = xlMedium
                            .Color = RGB(0, 0, 0)
                        End With
                    End If

                    If rr = lastRow Or currStick2 <> .Cells(rr + 1, "I").Value Then
                        With .Range("A" & rr & ":M" & rr).Borders(xlBottom)
                            .LineStyle = xlContinuous
                            .Weight = xlMedium
                            .Color = RGB(0, 0, 0)
                        End With
                    End If

                    prevSt = currStick2
                End If
            Next rr
        End If

        '---------------------------------------------------
        ' Print setup
        '---------------------------------------------------
        .PageSetup.PrintArea = "A1:M" & lastRow

        With .PageSetup
            .Orientation = xlLandscape
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .TopMargin = Application.InchesToPoints(0.5)
            .BottomMargin = Application.InchesToPoints(0.5)
            .LeftMargin = Application.InchesToPoints(0.5)
            .RightMargin = Application.InchesToPoints(0.5)
            .HeaderMargin = Application.InchesToPoints(0.3)
            .FooterMargin = Application.InchesToPoints(0.3)
            .CenterHorizontally = True
            .CenterVertically = False
            .PrintGridlines = False
            .PrintHeadings = False
            .PrintTitleRows = "$1:$1"
        End With

        .Columns.AutoFit
    End With

    '-------------------------------------------------------
    ' Completion message
    '-------------------------------------------------------
    If haveFilterWorkpack Then
        msgScope = "Workpack " & CStr(filterWorkpack)
    Else
        msgScope = "All workpacks"
    End If

    MsgBox "Cut nesting complete for mode: " & mode & vbCrLf & _
           "Scope: " & msgScope & vbCrLf & _
           "Cut allowance: " & cutAllowance & " mm per cut.", vbInformation

SafeExit:
    On Error Resume Next

    '-------------------------------------------------------
    ' Always restore protection states
    '-------------------------------------------------------
    If outGuardOpen Then SheetGuard_End outWS, shGuardOut, pwd
    If mtoGuardOpen Then SheetGuard_End ws, shGuardMTO, pwd

    AppGuard_End
    On Error GoTo 0
    Exit Sub

ErrHandler:
    MsgBox "Error in " & PROC_NAME & ":" & vbCrLf & _
           Err.Number & " - " & Err.description, vbCritical
    Resume SafeExit

End Sub

'==== MAIN NESTING PROCEDURE =======================================

Private Sub NestCuts_Fastold(ByVal mode As String)

    Const PROC_NAME As String = "NestCuts_Fast"
    
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim outWS As Worksheet
    
    Dim shGuardMTO As TSheetGuardState
    Dim shGuardOut As TSheetGuardState
    Dim mtoGuardOpen As Boolean
    Dim outGuardOpen As Boolean
    
    Dim colType As Long, colGrade As Long, colSize1 As Long
    Dim colProfile As Long, colAssDwg As Long, colLength As Long
    Dim colTotalQty As Long, colPipePrep As Long, colMarkNo As Long
    Dim colDisc As Long, colWorkpack As Long
    
    Dim dictGroups As Object
    Dim dictNestableProfiles As Object
    Dim key As String
    
    Dim cutAllowance As Double
    Dim totalPieces As Long
    Dim i As Long, q As Long
    
    Dim dataArr As Variant
    Dim nRows As Long
    
    Dim outArr As Variant
    Dim outIdx As Long
    
    ' summary array (per Grade/Size/Profile group)
    Dim summaryArr As Variant
    Dim summaryIdx As Long
    Dim groupCount As Long
    Dim totalStockLen As Double, totalWaste As Double, wastePct As Double
    
    Dim filterWorkpack As Variant
    Dim haveFilterWorkpack As Boolean
    
    Dim cutInput As Variant
    Dim cutText As String
    
    Dim thisType As String, thisDisc As String
    Dim thisLength As Double
    Dim thisUsed As Double
    Dim thisQty As Long
    Dim grade As Variant, size1 As Variant, profile As Variant
    Dim profileKey As String
    Dim assDwg As Variant, pipePrep As Variant, markNo As Variant
    Dim thisWorkpack As Variant
    
    Dim grp As Variant
    Dim colLen As Collection, colMark As Collection
    Dim colUsed As Collection, colPrep As Collection, colAss As Collection
    Dim colWork As Collection
    Dim stdLen As Double
    
    Dim k As Variant
    Dim n As Long
    Dim arrPiecesActual() As Double
    Dim arrPiecesUsed() As Double
    Dim arrMarks() As Variant
    Dim arrPrep() As Variant
    Dim arrAss() As Variant
    Dim arrWork() As Variant
    Dim arr2D() As Variant
    Dim stickRem() As Double
    Dim stickCount As Long
    Dim assignStick() As Long
    Dim s As Long
    Dim placed As Boolean
    
    Dim r1 As Long, r2 As Long
    Dim tmp(1 To 12) As Variant
    
    Dim expandedArr() As Variant
    Dim trimmedArr() As Variant
    Dim newIdx As Long
    Dim curStick As Variant, prevStick As Variant
    Dim r As Long, c As Long
    
    Dim lastStick As Variant
    Dim runLen As Double
    Dim pieceCount As Long
    
    Dim lastRow As Long
    Dim msgScope As String
    
    On Error GoTo ErrHandler
    
    '-----------------------------
    '  App guard / table lookup
    '-----------------------------
    AppGuard_Begin
    
    Set ws = GetWorksheetOfTable(ThisWorkbook, "tbl_MTO")
    If ws Is Nothing Then Err.Raise vbObjectError + 1000, PROC_NAME, "Could not locate worksheet for table 'tbl_MTO'."
    
    Set lo = ws.ListObjects("tbl_MTO")
    If lo Is Nothing Then Err.Raise vbObjectError + 1001, PROC_NAME, "Could not locate table 'tbl_MTO'."
    
    ' Unlock MTO sheet before touching filters / reading table
    shGuardMTO = SheetGuard_Begin(ws, pwd)
    mtoGuardOpen = True
    
    ' Always start with NO filters on tbl_MTO
    On Error Resume Next
    If Not lo.AutoFilter Is Nothing Then
        lo.AutoFilter.ShowAllData
    End If
    On Error GoTo ErrHandler
    
    '-----------------------------
    '  Column indexes
    '-----------------------------
    colAssDwg = lo.ListColumns("Ass Dwg No.").Index
    colType = lo.ListColumns("Type").Index
    colGrade = lo.ListColumns("Grade").Index
    colSize1 = lo.ListColumns("Size 1").Index
    colProfile = lo.ListColumns("Profile").Index
    colLength = lo.ListColumns("Length (mm)").Index
    colTotalQty = lo.ListColumns("Total Qty").Index
    colPipePrep = lo.ListColumns("Pipe Prep").Index
    colMarkNo = lo.ListColumns("Mark No.").Index
    colDisc = lo.ListColumns("Discipline").Index
    colWorkpack = lo.ListColumns("Workpack").Index
    
    '-----------------------------
    '  Optional Workpack filter based on selection
    '-----------------------------
    On Error Resume Next
    If Not Intersect(Selection, lo.ListColumns("Workpack").DataBodyRange) Is Nothing Then
        filterWorkpack = Selection.Value
        If Len(Trim$(CStr(filterWorkpack))) > 0 Then
            haveFilterWorkpack = True
        End If
    End If
    On Error GoTo ErrHandler
    
    '-----------------------------
    '  Nestable profiles from tbl_MatSpec
    '-----------------------------
    Set dictNestableProfiles = GetNestableProfiles()
    If dictNestableProfiles Is Nothing Then GoTo SafeExit
    
    If dictNestableProfiles.Count = 0 Then
        MsgBox "No profiles marked 'Y' for nesting in tbl_MatSpec.", vbExclamation
        GoTo SafeExit
    End If
    
    '-----------------------------
    '  Global cut allowance (kerf) – default 3 mm
    '-----------------------------
    cutInput = Application.InputBox( _
        Prompt:="Enter cut allowance per cut (mm)." & vbCrLf & _
                "Leave BLANK to use default 3 mm." & vbCrLf & _
                "Enter 0 if using a cropper.", _
        Title:="Cut Allowance (Kerf)", _
        Type:=2)
    
    ' Only treat Boolean False as Cancel
    If TypeName(cutInput) = "Boolean" Then
        If cutInput = False Then GoTo SafeExit
    End If
    
    cutText = Trim$(CStr(cutInput))
    
    If cutText = "" Then
        cutAllowance = 3
    ElseIf IsNumeric(cutText) Then
        cutAllowance = CDbl(cutText)
        If cutAllowance < 0 Then
            MsgBox "Cut allowance cannot be negative.", vbExclamation
            GoTo SafeExit
        End If
    Else
        MsgBox "Cut allowance must be a number.", vbExclamation
        GoTo SafeExit
    End If
    
    ' Ensure stock cache exists
    If dictStockCache Is Nothing Then
        Set dictStockCache = CreateObject("Scripting.Dictionary")
    End If
    
    '-----------------------------
    '  Load MTO into array
    '-----------------------------
    If lo.DataBodyRange Is Nothing Then
        MsgBox "tbl_MTO contains no data rows.", vbInformation
        GoTo SafeExit
    End If
    
    dataArr = lo.DataBodyRange.Value
    nRows = UBound(dataArr, 1)
    
    '-----------------------------
    '  First pass – group pieces
    '-----------------------------
    Set dictGroups = CreateObject("Scripting.Dictionary")
    totalPieces = 0
    
    For i = 1 To nRows
        thisType = UCase$(Trim$(CStr(dataArr(i, colType))))
        thisDisc = UCase$(Trim$(CStr(dataArr(i, colDisc))))
        thisWorkpack = dataArr(i, colWorkpack)
        
        ' Workpack filter
        If haveFilterWorkpack Then
            If CStr(thisWorkpack) <> CStr(filterWorkpack) Then GoTo NextRow
        End If
        
        ' Mode filter
        If mode = "PIPE" Then
            If thisType <> "PIPE" Then GoTo NextRow
        ElseIf mode = "STRUCT" Then
            If thisDisc <> "STRUCTURAL" Then GoTo NextRow
        End If
        
        thisLength = CDbl(dataArr(i, colLength))
        thisQty = CLng(val(dataArr(i, colTotalQty)))
        If thisLength <= 0 Or thisQty <= 0 Then GoTo NextRow
        
        grade = dataArr(i, colGrade)
        size1 = dataArr(i, colSize1)
        profile = dataArr(i, colProfile)
        assDwg = dataArr(i, colAssDwg)
        pipePrep = dataArr(i, colPipePrep)
        markNo = dataArr(i, colMarkNo)
        
        ' Must exist in tbl_MatSpec with Nesting = Y
        profileKey = UCase$(Trim$(CStr(profile)))
        If Not dictNestableProfiles.Exists(profileKey) Then GoTo NextRow
        
        ' Group key: Grade|Size1|Profile
        key = CStr(grade) & "|" & CStr(size1) & "|" & CStr(profile)
        
        If dictGroups.Exists(key) Then
            grp = dictGroups(key)
            stdLen = grp(10)
        Else
            ' Resolve stock length for this group (cache + spec + prompt)
            stdLen = ResolveStockLength(key, profile, grade, size1, mode)
            If stdLen <= 0 Then GoTo SafeExit
            
            ReDim grp(0 To 12)
            grp(0) = grade
            grp(1) = size1
            grp(2) = profile
            grp(3) = Empty
            grp(4) = Empty
            Set grp(5) = New Collection   ' lengths
            Set grp(6) = New Collection   ' marks
            grp(7) = 0                    ' count
            Set grp(8) = New Collection   ' used lengths (len+kerf)
            Set grp(9) = New Collection   ' pipe prep
            grp(10) = stdLen              ' stock length
            Set grp(11) = New Collection  ' ass dwg
            Set grp(12) = New Collection  ' workpack
            dictGroups.Add key, grp
        End If
        
        thisUsed = thisLength + cutAllowance
        If thisUsed > stdLen Then
            MsgBox "Piece length + cut allowance exceeds stock length:" & vbCrLf & _
                   "Piece length = " & thisLength & " mm, Cut = " & cutAllowance & " mm" & vbCrLf & _
                   "Total = " & thisUsed & " mm, Stock = " & stdLen & " mm" & vbCrLf & _
                   "Row: " & (lo.DataBodyRange.Row + i - 1), vbCritical
            GoTo SafeExit
        End If
        
        Set colLen = grp(5)
        Set colMark = grp(6)
        Set colUsed = grp(8)
        Set colPrep = grp(9)
        Set colAss = grp(11)
        Set colWork = grp(12)
        
        For q = 1 To thisQty
            colLen.Add thisLength
            colMark.Add markNo
            colUsed.Add thisUsed
            colPrep.Add pipePrep
            colAss.Add assDwg
            colWork.Add thisWorkpack
        Next q
        
        grp(7) = grp(7) + thisQty
        dictGroups(key) = grp
        
        totalPieces = totalPieces + thisQty
        
NextRow:
    Next i
    
    If totalPieces = 0 Then
        If haveFilterWorkpack Then
            MsgBox "No qualifying rows found to nest for mode " & mode & _
                   " and Workpack " & CStr(filterWorkpack) & ".", vbInformation
        Else
            MsgBox "No qualifying rows found to nest for mode " & mode & ".", vbInformation
        End If
        GoTo SafeExit
    End If
    
    '-----------------------------
    '  Prepare arrays
    '-----------------------------
    ReDim outArr(1 To totalPieces, 1 To 12)
    outIdx = 1
    
    groupCount = dictGroups.Count
    If groupCount > 0 Then
        ReDim summaryArr(1 To groupCount, 1 To 8)
        summaryIdx = 1
    End If
    
    '-----------------------------
    '  Process each group (FFD + backfill)
    '-----------------------------
    For Each k In dictGroups.keys
        grp = dictGroups(k)
        Set colLen = grp(5)
        Set colMark = grp(6)
        Set colUsed = grp(8)
        Set colPrep = grp(9)
        Set colAss = grp(11)
        Set colWork = grp(12)
        stdLen = grp(10)
        n = grp(7)
        If n = 0 Then GoTo NextGroup
        
        ReDim arrPiecesActual(1 To n)
        ReDim arrPiecesUsed(1 To n)
        ReDim arrMarks(1 To n)
        ReDim arrPrep(1 To n)
        ReDim arrAss(1 To n)
        ReDim arrWork(1 To n)
        
        For i = 1 To n
            arrPiecesActual(i) = CDbl(colLen(i))
            arrPiecesUsed(i) = CDbl(colUsed(i))
            arrMarks(i) = colMark(i)
            arrPrep(i) = colPrep(i)
            arrAss(i) = colAss(i)
            arrWork(i) = colWork(i)
        Next i
        
        ' arr2D: [used, mark, length, prep, ass, workpack]
        ReDim arr2D(1 To n, 1 To 6)
        For i = 1 To n
            arr2D(i, 1) = arrPiecesUsed(i)
            arr2D(i, 2) = arrMarks(i)
            arr2D(i, 3) = arrPiecesActual(i)
            arr2D(i, 4) = arrPrep(i)
            arr2D(i, 5) = arrAss(i)
            arr2D(i, 6) = arrWork(i)
        Next i
        
        If n > 1 Then QuickSort2D arr2D, 1, n
        
        ' FFD
        ReDim stickRem(1 To 1)
        stickCount = 0
        ReDim assignStick(1 To n)
        
        For i = 1 To n
            thisUsed = arr2D(i, 1)
            placed = False
            
            For s = 1 To stickCount
                If stickRem(s) >= thisUsed Then
                    stickRem(s) = stickRem(s) - thisUsed
                    assignStick(i) = s
                    placed = True
                    Exit For
                End If
            Next s
            
            If Not placed Then
                stickCount = stickCount + 1
                If stickCount = 1 Then
                    stickRem(1) = stdLen - thisUsed
                Else
                    ReDim Preserve stickRem(1 To stickCount)
                    stickRem(stickCount) = stdLen - thisUsed
                End If
                assignStick(i) = stickCount
            End If
        Next i
        
        ' Backfill/compress
        Dim s1 As Long, s2 As Long
        For s1 = 1 To stickCount - 1
            For s2 = s1 + 1 To stickCount
                For i = 1 To n
                    If assignStick(i) = s2 Then
                        thisUsed = arr2D(i, 1)
                        If stickRem(s1) >= thisUsed Then
                            stickRem(s1) = stickRem(s1) - thisUsed
                            stickRem(s2) = stickRem(s2) + thisUsed
                            assignStick(i) = s1
                        End If
                    End If
                Next i
            Next s2
        Next s1
        
        Dim stickHasPiece() As Boolean
        Dim newStickNum() As Long
        Dim newStickRem() As Double
        Dim newCount As Long
        
        ReDim stickHasPiece(1 To stickCount)
        For i = 1 To n
            stickHasPiece(assignStick(i)) = True
        Next i
        
        ReDim newStickNum(1 To stickCount)
        newCount = 0
        For s = 1 To stickCount
            If stickHasPiece(s) Then
                newCount = newCount + 1
                newStickNum(s) = newCount
            End If
        Next s
        
        ReDim newStickRem(1 To newCount)
        For s = 1 To newCount
            newStickRem(s) = stdLen
        Next s
        
        For i = 1 To n
            s = newStickNum(assignStick(i))
            assignStick(i) = s
            newStickRem(s) = newStickRem(s) - arr2D(i, 1)
        Next i
        
        stickCount = newCount
        stickRem = newStickRem
        
        ' Write group to outArr
        For i = 1 To n
            s = assignStick(i)
            outArr(outIdx, 1) = arr2D(i, 6)   ' Workpack
            outArr(outIdx, 2) = arr2D(i, 5)   ' Ass Dwg No.
            outArr(outIdx, 3) = grp(2)        ' Profile
            outArr(outIdx, 4) = arr2D(i, 2)   ' Mark No.
            outArr(outIdx, 5) = arr2D(i, 3)   ' Length (mm)
            outArr(outIdx, 6) = arr2D(i, 4)   ' Pipe Prep
            outArr(outIdx, 7) = grp(0)        ' Grade
            outArr(outIdx, 8) = grp(1)        ' Size 1
            outArr(outIdx, 9) = s             ' Stick No
            outArr(outIdx, 10) = 0            ' Running length
            outArr(outIdx, 11) = stdLen       ' Stock Length
            outArr(outIdx, 12) = stickRem(s)  ' Waste on stick
            outIdx = outIdx + 1
        Next i
        
        ' Summary row
        If groupCount > 0 Then
            Dim stickIdx As Long, usedOnStick As Double
            Dim i2 As Long
            
            totalStockLen = stickCount * stdLen
            totalWaste = 0
            
            For stickIdx = 1 To stickCount
                pieceCount = 0
                usedOnStick = 0
                
                For i2 = 1 To n
                    If assignStick(i2) = stickIdx Then
                        pieceCount = pieceCount + 1
                        usedOnStick = usedOnStick + arr2D(i2, 3)
                    End If
                Next i2
                
                If pieceCount > 1 Then
                    usedOnStick = usedOnStick + cutAllowance * (pieceCount - 1)
                End If
                
                totalWaste = totalWaste + (stdLen - usedOnStick)
            Next stickIdx
            
            If totalStockLen > 0 Then
                wastePct = totalWaste / totalStockLen
            Else
                wastePct = 0
            End If
            
            summaryArr(summaryIdx, 1) = grp(0)
            summaryArr(summaryIdx, 2) = grp(1)
            summaryArr(summaryIdx, 3) = grp(2)
            summaryArr(summaryIdx, 4) = stdLen
            summaryArr(summaryIdx, 5) = stickCount
            summaryArr(summaryIdx, 6) = totalStockLen
            summaryArr(summaryIdx, 7) = totalWaste
            summaryArr(summaryIdx, 8) = wastePct
            
            summaryIdx = summaryIdx + 1
        End If
        
NextGroup:
    Next k
    
    '-----------------------------
    '  Sort detail by Grade, Size1, StickNo
    '-----------------------------
    For r1 = 1 To totalPieces - 1
        For r2 = r1 + 1 To totalPieces
            If (CStr(outArr(r1, 7)) > CStr(outArr(r2, 7))) _
            Or (CStr(outArr(r1, 7)) = CStr(outArr(r2, 7)) And CStr(outArr(r1, 8)) > CStr(outArr(r2, 8))) _
            Or (CStr(outArr(r1, 7)) = CStr(outArr(r2, 7)) And CStr(outArr(r1, 8)) = CStr(outArr(r2, 8)) And CLng(outArr(r1, 9)) > CLng(outArr(r2, 9))) Then
                
                For i = 1 To 12
                    tmp(i) = outArr(r1, i)
                    outArr(r1, i) = outArr(r2, i)
                    outArr(r2, i) = tmp(i)
                Next i
            End If
        Next r2
    Next r1
    
    '-----------------------------
    ' Insert blank rows between Stick changes
    '-----------------------------
    If totalPieces > 0 Then
        ReDim expandedArr(1 To totalPieces * 2, 1 To 12)
        newIdx = 0
        
        For r = 1 To totalPieces
            curStick = outArr(r, 9)
            
            If r > 1 Then
                prevStick = outArr(r - 1, 9)
                If curStick <> prevStick Then
                    newIdx = newIdx + 1
                End If
            End If
            
            newIdx = newIdx + 1
            For c = 1 To 12
                expandedArr(newIdx, c) = outArr(r, c)
            Next c
        Next r
        
        ReDim trimmedArr(1 To newIdx, 1 To 12)
        For r = 1 To newIdx
            For c = 1 To 12
                trimmedArr(r, c) = expandedArr(r, c)
            Next c
        Next r
        
        outArr = trimmedArr
        totalPieces = newIdx
    End If
    
    '-----------------------------
    ' Calculate running length per stick
    '-----------------------------
    If totalPieces > 0 Then
        lastStick = vbNullString
        runLen = 0
        pieceCount = 0
        
        For r = 1 To totalPieces
            curStick = outArr(r, 9)
            
            If IsEmpty(curStick) Or curStick = "" Then
                outArr(r, 10) = Empty
            Else
                If curStick <> lastStick Then
                    pieceCount = 1
                    If IsNumeric(outArr(r, 5)) Then
                        runLen = CDbl(outArr(r, 5))
                    Else
                        runLen = 0
                    End If
                    lastStick = curStick
                Else
                    pieceCount = pieceCount + 1
                    If IsNumeric(outArr(r, 5)) Then
                        runLen = runLen + cutAllowance + CDbl(outArr(r, 5))
                    End If
                End If
                
                outArr(r, 10) = runLen
            End If
        Next r
    End If
    
    '-----------------------------
    '  Dump to "Nesting" sheet + formatting
    '-----------------------------
    On Error Resume Next
    Set outWS = ThisWorkbook.Worksheets("Nesting")
    On Error GoTo ErrHandler
    
    If outWS Is Nothing Then
        Set outWS = ThisWorkbook.Worksheets.Add(After:=ws)
        outWS.Name = "Nesting"
    End If
    
    ' Unlock output sheet before clearing / formatting
    shGuardOut = SheetGuard_Begin(outWS, pwd)
    outGuardOpen = True
    
    outWS.Cells.Clear
    
    With outWS
        .Range("A1").Value = "Workpack"
        .Range("B1").Value = "Ass Dwg No."
        .Range("C1").Value = "Profile"
        .Range("D1").Value = "Mark No."
        .Range("E1").Value = "Length (mm)"
        .Range("F1").Value = "Pipe Prep"
        .Range("G1").Value = "Grade"
        .Range("H1").Value = "Size 1"
        .Range("I1").Value = "Stick No"
        .Range("J1").Value = "Running Length (mm)"
        .Range("K1").Value = "Stock Length (mm)"
        .Range("L1").Value = "Waste on Stick (mm)"
        .Range("M1").Value = "Heat No."
        
        If totalPieces > 0 Then
            .Range("A2").Resize(totalPieces, 12).Value = outArr
        End If
        
        .Range("N1").Value = "Grade"
        .Range("O1").Value = "Size 1"
        .Range("P1").Value = "Profile"
        .Range("Q1").Value = "Stock Length (mm)"
        .Range("R1").Value = "No. of Sticks"
        .Range("S1").Value = "Total Stock Length (mm)"
        .Range("T1").Value = "Total Waste (mm)"
        .Range("U1").Value = "Waste (%)"
        
        If groupCount > 0 And summaryIdx > 1 Then
            .Range("N2").Resize(summaryIdx - 1, 8).Value = summaryArr
            .Range("U2:U" & summaryIdx).NumberFormat = "0.0%"
        End If
        
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        If lastRow < 1 Then lastRow = 1
        
        If totalPieces > 0 Then
            Dim rngCF As Range
            Dim fcOdd As FormatCondition
            Dim fcEven As FormatCondition
            Dim rr As Long
            Dim currStick2 As Variant, prevSt As Variant
            
            Set rngCF = .Range("A2:M" & lastRow)
            rngCF.FormatConditions.Delete
            
            Set fcOdd = rngCF.FormatConditions.Add( _
                Type:=xlExpression, _
                Formula1:="=AND($I2<>"""",MOD($I2,2)=1)")
            fcOdd.Interior.Color = RGB(235, 235, 235)
            
            Set fcEven = rngCF.FormatConditions.Add( _
                Type:=xlExpression, _
                Formula1:="=AND($I2<>"""",MOD($I2,2)=0)")
            fcEven.Interior.Pattern = xlNone
            
            With .Range("A2:M" & lastRow).Borders
                .LineStyle = xlContinuous
                .Weight = xlHairline
                .Color = RGB(200, 200, 200)
            End With
            
            prevSt = vbNullString
            For rr = 2 To lastRow
                currStick2 = .Cells(rr, "I").Value
                If currStick2 <> "" Then
                    If rr = 2 Or currStick2 <> prevSt Then
                        With .Range("A" & rr & ":M" & rr).Borders(xlTop)
                            .LineStyle = xlContinuous
                            .Weight = xlMedium
                            .Color = RGB(0, 0, 0)
                        End With
                    End If
                    
                    If rr = lastRow Or currStick2 <> .Cells(rr + 1, "I").Value Then
                        With .Range("A" & rr & ":M" & rr).Borders(xlBottom)
                            .LineStyle = xlContinuous
                            .Weight = xlMedium
                            .Color = RGB(0, 0, 0)
                        End With
                    End If
                    
                    prevSt = currStick2
                End If
            Next rr
        End If
        
        .PageSetup.PrintArea = "A1:M" & lastRow
        With .PageSetup
            .Orientation = xlLandscape
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .TopMargin = Application.InchesToPoints(0.5)
            .BottomMargin = Application.InchesToPoints(0.5)
            .LeftMargin = Application.InchesToPoints(0.5)
            .RightMargin = Application.InchesToPoints(0.5)
            .HeaderMargin = Application.InchesToPoints(0.3)
            .FooterMargin = Application.InchesToPoints(0.3)
            .CenterHorizontally = True
            .CenterVertically = False
            .PrintGridlines = False
            .PrintHeadings = False
            .PrintTitleRows = "$1:$1"
        End With
        
        .Columns.AutoFit
    End With
    
    If haveFilterWorkpack Then
        msgScope = "Workpack " & CStr(filterWorkpack)
    Else
        msgScope = "All workpacks"
    End If
    
    MsgBox "Cut nesting complete for mode: " & mode & vbCrLf & _
           "Scope: " & msgScope & vbCrLf & _
           "Cut allowance: " & cutAllowance & " mm per cut.", vbInformation
    
SafeExit:
    On Error Resume Next
    
    If outGuardOpen Then SheetGuard_End outWS, shGuardOut, pwd
    If mtoGuardOpen Then SheetGuard_End ws, shGuardMTO, pwd
    
    AppGuard_End
    On Error GoTo 0
    Exit Sub

ErrHandler:
    MsgBox "Error in " & PROC_NAME & ":" & vbCrLf & _
           Err.Number & " - " & Err.description, vbCritical
    
    Resume SafeExit

End Sub



'==== LOOKUP NESTABLE PROFILES FROM tbl_MatSpec ====================

Private Function GetNestableProfiles() As Object
    Const MAT_SHEET As String = "Material Properties"
    Const MAT_TABLE As String = "tbl_MatSpec"
    Const MAT_DESC_COL_NAME As String = "Description"
    Const MAT_NEST_COL_NAME As String = "Nesting"   ' column with Y
    
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim descCol As Long, nestCol As Long
    Dim arr As Variant
    Dim i As Long
    Dim key As String
    Dim dict As Object
    
    On Error GoTo ErrHandler
    
    Set ws = ThisWorkbook.Worksheets(MAT_SHEET)
    Set lo = ws.ListObjects(MAT_TABLE)
    
    descCol = lo.ListColumns(MAT_DESC_COL_NAME).Index
    nestCol = lo.ListColumns(MAT_NEST_COL_NAME).Index
    
    arr = lo.DataBodyRange.Value
    Set dict = CreateObject("Scripting.Dictionary")
    
    For i = 1 To UBound(arr, 1)
        If UCase$(Trim$(CStr(arr(i, nestCol)))) = "Y" Then
            key = UCase$(Trim$(CStr(arr(i, descCol))))
            If Len(key) > 0 Then
                If Not dict.Exists(key) Then dict.Add key, True
            End If
        End If
    Next i
    
    Set GetNestableProfiles = dict
    Exit Function
    
ErrHandler:
    MsgBox "Error reading tbl_MatSpec. Check sheet/table/column names." & vbCrLf & _
           "Sheet: " & MAT_SHEET & vbCrLf & _
           "Table: " & MAT_TABLE & vbCrLf & _
           "Description col: " & MAT_DESC_COL_NAME & vbCrLf & _
           "Nesting flag col: " & MAT_NEST_COL_NAME, vbCritical
    Set GetNestableProfiles = Nothing
End Function

'==== DEFAULT STOCK LENGTH RESOLUTION (SPEC + CACHE + PROMPT) ======

Private Function ResolveStockLength(ByVal groupKey As String, _
                                    ByVal profile As Variant, _
                                    ByVal grade As Variant, _
                                    ByVal size1 As Variant, _
                                    ByVal mode As String) As Double
    Dim defaultStock As Variant
    Dim msg As String
    Dim useDefault As VbMsgBoxResult
    Dim stockInput As Variant, stockText As String
    Dim stdLen As Double
    
    ' Ensure cache exists
    If dictStockCache Is Nothing Then
        Set dictStockCache = CreateObject("Scripting.Dictionary")
    End If
    
    ' 1) If we have a cached value for this group, confirm using it
    If dictStockCache.Exists(groupKey) Then
        stdLen = CDbl(dictStockCache(groupKey))
        msg = "Last used stock length for this group:" & vbCrLf & _
              "Grade: " & grade & vbCrLf & _
              "Size 1: " & size1 & vbCrLf & _
              "Profile: " & profile & vbCrLf & vbCrLf & _
              "Stock length: " & stdLen & " mm" & vbCrLf & _
              "Use this value?"
        
        useDefault = MsgBox(msg, vbYesNoCancel + vbQuestion, _
                            "Stock Length for Group (" & mode & ")")
        
        If useDefault = vbYes Then
            ResolveStockLength = stdLen
            Exit Function
        ElseIf useDefault = vbCancel Then
            ResolveStockLength = 0
            Exit Function
        End If
        ' If No, fall through to spec/default/prompt logic
    End If
    
    ' 2) Try to get stock length from spec (tbl_MatSpec[Stock (mm)])
    defaultStock = GetDefaultStockLength(profile)
    
    msg = "Enter stock length (mm) for:" & vbCrLf & _
          "Grade: " & grade & vbCrLf & _
          "Size 1: " & size1 & vbCrLf & _
          "Profile: " & profile & vbCrLf & vbCrLf
    
    If Not IsEmpty(defaultStock) Then
        msg = msg & "Default from specification: " & defaultStock & " mm" & vbCrLf & _
                    "Use this value?"
        useDefault = MsgBox(msg, vbYesNoCancel + vbQuestion, _
                            "Stock Length for Group (" & mode & ")")
        If useDefault = vbCancel Then
            ResolveStockLength = 0
            Exit Function
        End If
        If useDefault = vbYes Then
            stdLen = CDbl(defaultStock)
            GoTo StockDone
        End If
    Else
        msg = msg & "No spec stock length found." & vbCrLf
    End If
    
    ' 3) Ask user for stock length
    stockInput = Application.InputBox( _
        Prompt:=msg & vbCrLf & "(Leave blank or enter 0 for default 12000 mm)", _
        Title:="Stock Length for Group (" & mode & ")", _
        Type:=2)   'text input
    
    If stockInput = False Then
        ResolveStockLength = 0
        Exit Function
    End If
    
    stockText = Trim$(CStr(stockInput))
    
    If stockText = "" Or stockText = "0" Then
        stdLen = 12000
    ElseIf IsNumeric(stockText) Then
        stdLen = CDbl(stockText)
        If stdLen <= 0 Then stdLen = 12000
    Else
        MsgBox "Stock length must be a number.", vbExclamation
        ResolveStockLength = 0
        Exit Function
    End If
    
StockDone:
    ResolveStockLength = stdLen
    
    ' Update cache for this session
    If Not dictStockCache.Exists(groupKey) Then
        dictStockCache.Add groupKey, stdLen
    Else
        dictStockCache(groupKey) = stdLen
    End If
End Function

Private Function GetDefaultStockLength(ByVal profileValue As Variant) As Variant
    Const MAT_SHEET As String = "Material Properties"
    Const MAT_TABLE As String = "tbl_MatSpec"
    Const COL_DESC As String = "Description"
    Const COL_STOCK As String = "Stock (mm)"
    
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim descCol As Long, stockCol As Long
    Dim arr As Variant
    Dim i As Long
    Dim prof As String
    
    On Error GoTo ErrHandler
    
    prof = UCase$(Trim$(CStr(profileValue)))
    
    Set ws = ThisWorkbook.Worksheets(MAT_SHEET)
    Set lo = ws.ListObjects(MAT_TABLE)
    
    descCol = lo.ListColumns(COL_DESC).Index
    stockCol = lo.ListColumns(COL_STOCK).Index
    
    arr = lo.DataBodyRange.Value
    
    For i = 1 To UBound(arr, 1)
        If UCase$(Trim$(CStr(arr(i, descCol)))) = prof Then
            If IsNumeric(arr(i, stockCol)) Then
                GetDefaultStockLength = arr(i, stockCol)
                Exit Function
            End If
        End If
    Next i
    
    GetDefaultStockLength = Empty
    Exit Function
    
ErrHandler:
    GetDefaultStockLength = Empty
End Function

'==== QuickSort for 2D array: sort desc on col 1 ===================

Private Sub QuickSort2D(ByRef arr As Variant, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long
    Dim pivot As Double
    Dim tmp1 As Double
    Dim tmp2 As Variant
    Dim tmp3 As Variant
    Dim tmp4 As Variant
    Dim tmp5 As Variant
    Dim tmp6 As Variant
    
    i = first
    j = last
    pivot = arr((first + last) \ 2, 1)
    
    Do While i <= j
        Do While arr(i, 1) > pivot
            i = i + 1
        Loop
        Do While arr(j, 1) < pivot
            j = j - 1
        Loop
        
        If i <= j Then
            ' used length
            tmp1 = arr(i, 1)
            arr(i, 1) = arr(j, 1)
            arr(j, 1) = tmp1
            
            ' mark
            tmp2 = arr(i, 2)
            arr(i, 2) = arr(j, 2)
            arr(j, 2) = tmp2
            
            ' actual length
            tmp3 = arr(i, 3)
            arr(i, 3) = arr(j, 3)
            arr(j, 3) = tmp3
            
            ' prep
            tmp4 = arr(i, 4)
            arr(i, 4) = arr(j, 4)
            arr(j, 4) = tmp4
            
            ' ass dwg
            tmp5 = arr(i, 5)
            arr(i, 5) = arr(j, 5)
            arr(j, 5) = tmp5
            
            ' workpack
            tmp6 = arr(i, 6)
            arr(i, 6) = arr(j, 6)
            arr(j, 6) = tmp6
            
            i = i + 1
            j = j - 1
        End If
    Loop
    
    If first < j Then QuickSort2D arr, first, j
    If i < last Then QuickSort2D arr, i, last
End Sub


