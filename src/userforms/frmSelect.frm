VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelect 
   Caption         =   "Select Workbook"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmSelect.frx":0000
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public vSelection As String
Public FrmType As Integer

Public Property Get SelectedWorkbookName() As String
    SelectedWorkbookName = vSelection
End Property


Private Sub UserForm_Initialize()
    Dim xlLeft As Long, xlTop As Long
    Dim xlWidth As Long, xlHeight As Long
    Dim frmWidth As Long, frmHeight As Long

    ' Get Excel window position and size
    With Application
        xlLeft = .Left
        xlTop = .Top
        xlWidth = .Width
        xlHeight = .Height
    End With

    ' Get form size in points (approximate)
    frmWidth = Me.Width
    frmHeight = Me.Height

    ' Center the form over the Excel window
    Me.Left = xlLeft + (xlWidth - frmWidth) / 2
    Me.Top = xlTop + (xlHeight - frmHeight) / 2
    
    If cmbSelection.ListCount > 0 Then
        cmbSelection.ListIndex = 0 ' Select first item by default
    End If
    
End Sub

Sub LoadCombo()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim val1 As String, val2 As String
    
    Select Case FrmType
    
    Case 1 ' Populate ComboBox with open workbooks (excluding ThisWorkbook)
        For Each wb In Application.Workbooks
            If wb.name <> ThisWorkbook.name Then
                cmbSelection.AddItem wb.name
            End If
        Next wb
        lbl1.Caption = "Select Workbook to Get Data from."
        frmSelect.Caption = "Select Workbook"
    Case 2 ' Populate ComboBox with PCC Claims
        
        Set ws = ThisWorkbook.Sheets("PC Register")
        Set tbl = ws.ListObjects("tblPCC")
        ' Clear and fill combo box
        cmbSelection.Clear
        For i = 1 To tbl.ListRows.Count
            cmbSelection.AddItem tbl.DataBodyRange.Cells(i, 1).Value
        Next i
        lbl1.Caption = "Select Progress Claim to populate."
        frmSelect.Caption = "Select Progress Claim"
    Case 3
        Set ws = ThisWorkbook.Sheets("Table Names Summary")
        Set tbl = ws.ListObjects("tblTables")
        With cmbSelection
            .Clear
            
            For i = 1 To tbl.ListRows.Count
                val1 = tbl.DataBodyRange.Cells(i, 1).Value ' Column 1
                If tbl.DataBodyRange.Cells(i, 6) = True Then
                    If Not ItemExistsInCombo(cmbSelection, val1) Then
                        .AddItem
                        .list(.ListCount - 1, 0) = val1
                    End If
                End If
            Next i
        End With
        lbl1.Caption = "Select Sheet."
        frmSelect.Caption = "Select Sheet to Navigate to."
    End Select
    
    If cmbSelection.ListCount > 0 Then
        cmbSelection.ListIndex = 0 ' Select first item by default
    End If
End Sub

Private Sub btnOK_Click()
    If cmbSelection.ListIndex = -1 Then
        MsgBox "Please select.", vbExclamation
        Exit Sub
    End If
    vSelection = cmbSelection.Value
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    vSelection = ""
    Me.Hide
End Sub

'====================================================================================
' ItemExistsInCombo
'
' Purpose:
'   Returns TRUE if a value exists in a ComboBox list.
'
' Parameters:
'   cbo            -> MSForms.ComboBox control
'   searchValue    -> Value to search for
'   Optional exactMatch (default TRUE)
'
' Behaviour:
'   - Case-insensitive
'   - Safe if ComboBox empty
'   - Does NOT alter selection
'
'====================================================================================
Public Function ItemExistsInCombo( _
        ByVal cbo As MSForms.ComboBox, _
        ByVal searchValue As String, _
        Optional ByVal exactMatch As Boolean = True) As Boolean

    Dim i As Long
    Dim itemText As String
    Dim tgt As String

    On Error GoTo SafeExit

    If cbo Is Nothing Then Exit Function
    If cbo.ListCount = 0 Then Exit Function

    tgt = Trim$(LCase$(searchValue))

    For i = 0 To cbo.ListCount - 1

        itemText = Trim$(LCase$(cbo.list(i)))

        If exactMatch Then
            If itemText = tgt Then
                ItemExistsInCombo = True
                Exit Function
            End If
        Else
            If InStr(1, itemText, tgt, vbTextCompare) > 0 Then
                ItemExistsInCombo = True
                Exit Function
            End If
        End If

    Next i

SafeExit:
End Function


