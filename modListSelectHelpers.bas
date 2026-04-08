Attribute VB_Name = "modListSelectHelpers"
Option Explicit

'=======================================================
' modPricebookHelpers
'
' Purpose:
'   Helper routines for frmListSelect table resolution,
'   validation, selected-row mapping, and safe invocation
'   of modGuardsAndTables utilities.
'
' Notes:
'   - Written to be compile-safe even if helper procedure
'     signatures in modGuardsAndTables vary
'   - Uses late-bound Application.Run wrappers for:
'       AppGuard_Begin
'       AppGuard_End
'       SheetGuard_Begin
'       SheetGuard_End
'       LogError
'=======================================================

'-------------------------------------------------------
' FindListObjectInWorkbook
'
' Purpose:
'   Locate a ListObject by name anywhere in the workbook.
'
' Returns:
'   The matching ListObject, or Nothing if not found.
'-------------------------------------------------------
Public Function FindListObjectInWorkbook(ByVal tableName As String, _
                                         ByVal wb As Workbook) As ListObject
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                Set FindListObjectInWorkbook = lo
                Exit Function
            End If
        Next lo
    Next ws

CleanExit:
    Exit Function

ErrHandler:
    SafeLogError "FindListObjectInWorkbook", Err.Number, Err.description
    Resume CleanExit
End Function

'-------------------------------------------------------
' GetTableColumnIndex
'
' Purpose:
'   Return the 1-based index of a header within a table.
'
' Returns:
'   0 if header is not found.
'-------------------------------------------------------
Public Function GetTableColumnIndex(ByVal lo As ListObject, _
                                    ByVal headerText As String) As Long
    On Error GoTo ErrHandler

    Dim lc As ListColumn

    If lo Is Nothing Then Exit Function

    For Each lc In lo.ListColumns
        If StrComp(Trim$(lc.Name), Trim$(headerText), vbTextCompare) = 0 Then
            GetTableColumnIndex = lc.Index
            Exit Function
        End If
    Next lc

CleanExit:
    Exit Function

ErrHandler:
    SafeLogError "GetTableColumnIndex", Err.Number, Err.description
    Resume CleanExit
End Function

'-------------------------------------------------------
' GetSelectedTableRowIndexes
'
' Purpose:
'   Determines which data rows in a table intersect the
'   user's current selection.
'
' Arguments:
'   lo         - target table
'   sel        - current Excel selection
'   rowIdxOut  - output array of unique 1-based table row
'                indexes inside DataBodyRange
'
' Returns:
'   Number of selected table rows found.
'
' Notes:
'   - Handles multi-area selections
'   - Ignores selection outside the table
'   - Returns unique row indexes only
'-------------------------------------------------------
Public Function GetSelectedTableRowIndexes(ByVal lo As ListObject, _
                                           ByVal sel As Object, _
                                           ByRef rowIdxOut() As Long) As Long
    On Error GoTo ErrHandler

    Dim dict As Object
    Dim rngBody As Range
    Dim rngHit As Range
    Dim area As Range
    Dim rw As Range
    Dim tableRowIdx As Long
    Dim i As Long
    Dim k As Variant

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If sel Is Nothing Then Exit Function
    If TypeName(sel) <> "Range" Then Exit Function

    Set rngBody = lo.DataBodyRange
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    For Each area In sel.Areas
        Set rngHit = Intersect(area.EntireRow, rngBody.EntireRow)

        If Not rngHit Is Nothing Then
            For Each rw In rngHit.Rows
                If rw.Row >= rngBody.Row And rw.Row < rngBody.Row + rngBody.Rows.Count Then
                    tableRowIdx = rw.Row - rngBody.Row + 1
                    If tableRowIdx >= 1 And tableRowIdx <= rngBody.Rows.Count Then
                        If Not dict.Exists(CStr(tableRowIdx)) Then
                            dict.Add CStr(tableRowIdx), tableRowIdx
                        End If
                    End If
                End If
            Next rw
        End If
    Next area

    If dict.Count = 0 Then Exit Function

    ReDim rowIdxOut(1 To dict.Count)
    i = 0
    For Each k In dict.keys
        i = i + 1
        rowIdxOut(i) = CLng(dict(k))
    Next k

    GetSelectedTableRowIndexes = dict.Count

CleanExit:
    Exit Function

ErrHandler:
    SafeLogError "GetSelectedTableRowIndexes", Err.Number, Err.description
    GetSelectedTableRowIndexes = 0
    Erase rowIdxOut
    Resume CleanExit
End Function

'-------------------------------------------------------
' NzText
'
' Purpose:
'   Safe text coercion for Null / Empty / Error values.
'-------------------------------------------------------
Public Function NzText(ByVal v As Variant, Optional ByVal Fallback As String = "") As String
    On Error GoTo ErrHandler

    If IsError(v) Then
        NzText = Fallback
    ElseIf IsNull(v) Then
        NzText = Fallback
    ElseIf IsEmpty(v) Then
        NzText = Fallback
    Else
        NzText = Trim$(CStr(v))
    End If

CleanExit:
    Exit Function

ErrHandler:
    NzText = Fallback
    Resume CleanExit
End Function

'-------------------------------------------------------
' IsValueBlank
'
' Purpose:
'   Returns True if a value should be treated as blank.
'-------------------------------------------------------
Public Function IsValueBlank(ByVal v As Variant) As Boolean
    On Error GoTo ErrHandler

    If IsError(v) Then
        IsValueBlank = True
    ElseIf IsNull(v) Then
        IsValueBlank = True
    ElseIf IsEmpty(v) Then
        IsValueBlank = True
    Else
        IsValueBlank = (Len(Trim$(CStr(v))) = 0)
    End If

CleanExit:
    Exit Function

ErrHandler:
    IsValueBlank = True
    Resume CleanExit
End Function

'-------------------------------------------------------
' SafeAppGuardBegin
'
' Purpose:
'   Calls modGuardsAndTables.AppGuard_Begin if available.
'-------------------------------------------------------
Public Sub SafeAppGuardBegin()
    On Error Resume Next
    Application.Run "AppGuard_Begin"
    On Error GoTo 0
End Sub

'-------------------------------------------------------
' SafeAppGuardEnd
'
' Purpose:
'   Calls modGuardsAndTables.AppGuard_End if available.
'-------------------------------------------------------
Public Sub SafeAppGuardEnd()
    On Error Resume Next
    Application.Run "AppGuard_End"
    On Error GoTo 0
End Sub

'-------------------------------------------------------
' SafeSheetGuardBegin
'
' Purpose:
'   Calls modGuardsAndTables.SheetGuard_Begin if available.
'-------------------------------------------------------
Public Sub SafeSheetGuardBegin(ByVal ws As Worksheet)
    On Error Resume Next
    Application.Run "SheetGuard_Begin", ws
    On Error GoTo 0
End Sub

'-------------------------------------------------------
' SafeSheetGuardEnd
'
' Purpose:
'   Calls modGuardsAndTables.SheetGuard_End if available.
'-------------------------------------------------------
Public Sub SafeSheetGuardEnd(ByVal ws As Worksheet)
    On Error Resume Next
    Application.Run "SheetGuard_End", ws
    On Error GoTo 0
End Sub

'-------------------------------------------------------
' SafeLogError
'
' Purpose:
'   Logs errors through modGuardsAndTables.LogError if
'   available. Falls back silently if helper is absent.
'
' Notes:
'   Because LogError signatures vary between workbooks,
'   this wrapper attempts common call patterns.
'-------------------------------------------------------
Public Sub SafeLogError(ByVal procName As String, _
                        ByVal errNumber As Long, _
                        ByVal errDescription As String)
    On Error Resume Next

    Application.Run "LogError", procName, errNumber, errDescription
    If Err.Number <> 0 Then
        Err.Clear
        Application.Run "LogError", procName, errDescription
    End If

    On Error GoTo 0
End Sub


'-------------------------------------------------------
' ShowListSelectForTable
' Description:
'   Creates and shows frmListSelect for the supplied
'   configuration key.
'-------------------------------------------------------
Public Sub ShowListSelectForTable(ByVal tblName As String)

    On Error GoTo ErrHandler

    Dim frm As frmListSelect

    Set frm = New frmListSelect
    frm.InitialiseForTable tblName
    frm.Show vbModeless

    Exit Sub

ErrHandler:
    SafeLogError "ShowListSelectForTable", Err.Number, Err.description
    MsgBox "Unable to open the selection form." & vbCrLf & vbCrLf & _
           "Reason: " & Err.description, vbExclamation, "Selection Form"
End Sub

