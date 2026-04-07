VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmListSelect 
   Caption         =   "Pricebook"
   ClientHeight    =   6465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7080
   OleObjectBlob   =   "frmListSelect.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmListSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------
' frmListSelect
' Description:
'   Reusable lookup form that loads a configured source
'   table, allows dynamic search, and writes the selected
'   ID value back into selected rows of a configured
'   destination table.
'
' Inputs:
'   - InitialiseForTable(ByVal tblName As String)
'
' Outputs:
'   - Updates destination table values for selected rows
'
' Dependencies:
'   - SafeAppGuardBegin / SafeAppGuardEnd
'   - SafeSheetGuardBegin / SafeSheetGuardEnd
'   - SafeLogError
'   - FindListObjectInWorkbook
'   - GetTableColumnIndex
'   - GetSelectedTableRowIndexes
'   - IsValueBlank
'   - NzText
'   - Sort2DArrayByTwoKeys
'
' Assumptions:
'   - The calling code passes a valid configuration key
'   - The user has selected one or more rows in the
'     configured destination table before clicking Select
'-------------------------------------------------------

Private Const FORM_NAME As String = "frmListSelect"

'-------------------------------------------------------
' Cached lookup data
'   Column 0 = ID / Code
'   Column 1 = Description
'-------------------------------------------------------
Private mLookupData As Variant
Private mHasLookupData As Boolean

'-------------------------------------------------------
' Runtime configuration
'-------------------------------------------------------
Private mConfigKey As String
Private mSourceTableName As String
Private mPasteTableName As String
Private mIdHeader As String
Private mDescHeader As String
Private mPasteColumnHeader As String

Private mCaptionText As String
Private mTaskLabelText As String
Private mCodeLabelText As String
Private mDescLabelText As String

Private mIsInitialised As Boolean

'=======================================================
' Public Initialiser
'=======================================================

'-------------------------------------------------------
' InitialiseForTable
' Description:
'   Configures the form for the supplied table key.
'
' Inputs:
'   tblName - logical configuration key
'
' Outputs:
'   - Resolves source/destination tables
'   - Updates form captions/labels
'-------------------------------------------------------
Public Sub InitialiseForTable(ByVal tblName As String)

    On Error GoTo ErrHandler

    ResetFormState

    mConfigKey = Trim$(tblName)
    If Len(mConfigKey) = 0 Then
        Err.Raise vbObjectError + 7100, FORM_NAME & ".InitialiseForTable", _
                  "A tbl_Name/configuration key is required."
    End If

    ResolveConfiguration mConfigKey

    Me.Caption = mCaptionText
    Me.lblTask.Caption = mTaskLabelText
    Me.lblCode.Caption = mCodeLabelText
    Me.lblDesc.Caption = mDescLabelText

    mIsInitialised = True
    Exit Sub

ErrHandler:
    SafeLogError FORM_NAME & ".InitialiseForTable", Err.Number, Err.Description
    MsgBox "Unable to initialise the form." & vbCrLf & vbCrLf & _
           "Reason: " & Err.Description, vbExclamation, "Selection Form"
End Sub

'=======================================================
' Form Lifecycle
'=======================================================

'-------------------------------------------------------
' UserForm_Initialize
' Description:
'   Sets up controls only. Configuration is supplied
'   separately via InitialiseForTable.
'-------------------------------------------------------
Private Sub UserForm_Initialize()

    On Error GoTo ErrHandler

    With Me.lbxPricebook
        .Clear
        .ColumnCount = 2
        .BoundColumn = 1
        .ColumnWidths = "100 pt;280 pt"
        .ListStyle = fmListStylePlain
        .MultiSelect = fmMultiSelectSingle
    End With

    Me.txtSearch.Value = vbNullString

    Me.Caption = "Selection"
    Me.lblTask.Caption = "Select an item"
    Me.lblCode.Caption = "Code"
    Me.lblDesc.Caption = "Description"

    PositionFormTopRight
    Exit Sub

ErrHandler:
    SafeLogError FORM_NAME & ".UserForm_Initialize", Err.Number, Err.Description
    MsgBox "Unable to initialise the form controls." & vbCrLf & vbCrLf & _
           "Reason: " & Err.Description, vbExclamation, "Selection Form"
End Sub

'-------------------------------------------------------
' UserForm_Activate
' Description:
'   Loads configured source data once the caller has
'   passed tbl_Name into InitialiseForTable.
'-------------------------------------------------------
Private Sub UserForm_Activate()

    On Error GoTo ErrHandler

    If Not mIsInitialised Then Exit Sub
    If mHasLookupData Then Exit Sub

    LoadConfiguredSourceData
    Exit Sub

ErrHandler:
    SafeLogError FORM_NAME & ".UserForm_Activate", Err.Number, Err.Description
    MsgBox "Unable to load the selection list." & vbCrLf & vbCrLf & _
           "Reason: " & Err.Description, vbExclamation, Me.Caption
End Sub

'=======================================================
' Button Events
'=======================================================

Private Sub cmdCancel_Click()

    On Error GoTo ErrHandler
    Unload Me
    Exit Sub

ErrHandler:
    SafeLogError FORM_NAME & ".cmdCancel_Click", Err.Number, Err.Description
    MsgBox "Unable to close the form." & vbCrLf & vbCrLf & _
           "Reason: " & Err.Description, vbExclamation, Me.Caption
End Sub

'-------------------------------------------------------
' cmdPBSelect_Click
' Description:
'   Writes the selected code from the listbox into the
'   selected rows of the configured destination table.
'-------------------------------------------------------
Private Sub cmdPBSelect_Click()

    On Error GoTo ErrHandler

    Dim loPaste As ListObject
    Dim wsPaste As Worksheet
    Dim selectedId As String
    Dim selectedDesc As String
    Dim pasteRowIdx() As Long
    Dim selectedCount As Long
    Dim idxPaste As Long
    Dim arrPasteCol As Variant
    Dim i As Long

    SafeAppGuardBegin

    If Not mIsInitialised Then
        Err.Raise vbObjectError + 7110, FORM_NAME & ".cmdPBSelect_Click", _
                  "The form has not been initialised."
    End If

    If Me.lbxPricebook.ListIndex < 0 Then
        MsgBox "Select an item first.", vbInformation, Me.Caption
        GoTo CleanExit
    End If

    selectedId = NzText(Me.lbxPricebook.list(Me.lbxPricebook.ListIndex, 0))
    selectedDesc = NzText(Me.lbxPricebook.list(Me.lbxPricebook.ListIndex, 1))

    If Len(selectedId) = 0 Then
        MsgBox "The selected item does not contain a valid code.", vbExclamation, Me.Caption
        GoTo CleanExit
    End If

    Set loPaste = FindListObjectInWorkbook(mPasteTableName, ThisWorkbook)
    If loPaste Is Nothing Then
        Err.Raise vbObjectError + 7111, FORM_NAME & ".cmdPBSelect_Click", _
                  "Could not locate paste table '" & mPasteTableName & "'."
    End If

    Set wsPaste = loPaste.Parent

    idxPaste = GetTableColumnIndex(loPaste, mPasteColumnHeader)
    If idxPaste = 0 Then
        Err.Raise vbObjectError + 7112, FORM_NAME & ".cmdPBSelect_Click", _
                  "Header '" & mPasteColumnHeader & "' was not found in paste table '" & mPasteTableName & "'."
    End If

    If loPaste.DataBodyRange Is Nothing Then
        MsgBox "Paste table '" & mPasteTableName & "' contains no data rows.", vbInformation, Me.Caption
        GoTo CleanExit
    End If

    selectedCount = GetSelectedTableRowIndexes(loPaste, Selection, pasteRowIdx)
    If selectedCount = 0 Then
        MsgBox "Select one or more rows in '" & mPasteTableName & "' before clicking Select.", _
               vbInformation, Me.Caption
        GoTo CleanExit
    End If

    If MsgBox("Are you sure you want to change the selected rows to:" & vbCrLf & vbCrLf & _
              selectedId & " - " & selectedDesc & "?", _
              vbYesNo + vbQuestion, Me.Caption) = vbNo Then
        GoTo CleanExit
    End If

    SafeSheetGuardBegin wsPaste

    arrPasteCol = loPaste.ListColumns(idxPaste).DataBodyRange.Value2

    If IsArray(arrPasteCol) Then
        For i = 1 To selectedCount
            arrPasteCol(pasteRowIdx(i), 1) = selectedId
        Next i
        loPaste.ListColumns(idxPaste).DataBodyRange.Value2 = arrPasteCol
    Else
        loPaste.ListColumns(idxPaste).DataBodyRange.Value2 = selectedId
    End If

CleanExit:
    On Error Resume Next
    If Not wsPaste Is Nothing Then SafeSheetGuardEnd wsPaste
    SafeAppGuardEnd
    On Error GoTo 0
    Exit Sub

ErrHandler:
    SafeLogError FORM_NAME & ".cmdPBSelect_Click", Err.Number, Err.Description
    MsgBox "Unable to apply the selected value." & vbCrLf & vbCrLf & _
           "Reason: " & Err.Description, vbExclamation, Me.Caption
    Resume CleanExit
End Sub

'=======================================================
' Search / List Events
'=======================================================

Private Sub txtSearch_Change()

    On Error GoTo ErrHandler

    If Not mHasLookupData Then Exit Sub
    LoadLookupListFromCache Me.txtSearch.Text
    Exit Sub

ErrHandler:
    SafeLogError FORM_NAME & ".txtSearch_Change", Err.Number, Err.Description
End Sub

Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error GoTo ErrHandler

    Select Case KeyCode
        Case vbKeyEscape
            Me.txtSearch.Value = vbNullString
            KeyCode = 0
    End Select

    Exit Sub

ErrHandler:
    SafeLogError FORM_NAME & ".txtSearch_KeyDown", Err.Number, Err.Description
End Sub

Private Sub lbxPricebook_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo ErrHandler

    If Me.lbxPricebook.ListIndex >= 0 Then
        cmdPBSelect_Click
    End If

    Exit Sub

ErrHandler:
    SafeLogError FORM_NAME & ".lbxPricebook_DblClick", Err.Number, Err.Description
End Sub

'=======================================================
' Configuration
'=======================================================

'-------------------------------------------------------
' ResolveConfiguration
' Description:
'   Maps tbl_Name to:
'   - SourceTable
'   - PasteTable
'   - ID header
'   - Description header
'   - Paste column
'   - labels / caption
'
' Important:
'   Add a Case block for each supported lookup workflow.
'-------------------------------------------------------
Private Sub ResolveConfiguration(ByVal configKey As String)

    On Error GoTo ErrHandler

    Select Case LCase$(Trim$(configKey))

        Case "tbl_Install"
            mSourceTableName = "tbl_Pricebook"
            mPasteTableName = "tbl_Install"
            mIdHeader = "Comm Code"
            mDescHeader = "Description"
            mPasteColumnHeader = "Commodity"

            mCaptionText = "Pricebook Selection"
            mTaskLabelText = "Select a Pricebook item"
            mCodeLabelText = "Comm Code"
            mDescLabelText = "Description"

        Case "tbl_rfqdistribution"
            mSourceTableName = "tbl_RFQ"
            mPasteTableName = "tbl_RFQDistribution"
            mIdHeader = "RFQID"
            mDescHeader = "RFQ_Description"
            mPasteColumnHeader = "RFQID"

            mCaptionText = "RFQ Selection"
            mTaskLabelText = "Select an item"
            mCodeLabelText = "Code"
            mDescLabelText = "Description"
            
        Case "tbl_pricebook"
            mSourceTableName = "tbl_ROCType"
            mPasteTableName = "tbl_Pricebook"
            mIdHeader = "RulesOfCredit_Desc"
            mDescHeader = "Comments"
            mPasteColumnHeader = "ROC"

            mCaptionText = "ROC Selection"
            mTaskLabelText = "Select a ROC item"
            mCodeLabelText = "ROC"
            mDescLabelText = "Description"

        Case "tbl_workpackage"
            mSourceTableName = "tbl_workpackage"
            mPasteTableName = "tbl_Tracking"
            mIdHeader = "Workpackage"
            mDescHeader = "Description"
            mPasteColumnHeader = "Workpack"

            mCaptionText = "Workpack Selection"
            mTaskLabelText = "Select a Workpack"
            mCodeLabelText = "Workpack"
            mDescLabelText = "Description"

        Case "lcs_subcon"
            mSourceTableName = "tbl_DD"
            mPasteTableName = "tbl_Tracking"
            mIdHeader = "Delivery Docket Number:"
            mDescHeader = "Dispatch Date"
            mPasteColumnHeader = "Load Sheet No. to Subcontractor"

            mCaptionText = "Delivery Docket Selection"
            mTaskLabelText = "Select a Delivery Docket"
            mCodeLabelText = "Delivery Docket"
            mDescLabelText = "Dispatch Date"
            
        Case "lcs_dd"
            mSourceTableName = "tbl_DD"
            mPasteTableName = "tbl_Tracking"
            mIdHeader = "Delivery Docket Number:"
            mDescHeader = "Dispatch Date"
            mPasteColumnHeader = "Delivery Docket # "

            mCaptionText = "Delivery Docket Selection"
            mTaskLabelText = "Select a Delivery Docket"
            mCodeLabelText = "Delivery Docket"
            mDescLabelText = "Dispatch Date"
            
        Case "lcs_tpp"
            mSourceTableName = "tbl_DD"
            mPasteTableName = "tbl_Tracking"
            mIdHeader = "Delivery Docket Number:"
            mDescHeader = "Dispatch Date"
            mPasteColumnHeader = "Load Sheet No. to TPP"

            mCaptionText = "Delivery Docket Selection"
            mTaskLabelText = "Select a Delivery Docket"
            mCodeLabelText = "Delivery Docket"
            mDescLabelText = "Dispatch Date"
            

        Case Else
            Err.Raise vbObjectError + 7120, FORM_NAME & ".ResolveConfiguration", _
                      "Unsupported tbl_Name/configuration key: '" & configKey & "'."
    End Select

    Exit Sub

ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'=======================================================
' Source Load / Search
'=======================================================

'-------------------------------------------------------
' LoadConfiguredSourceData
' Description:
'   Reads configured source table into an in-memory cache,
'   sorts it, then loads the listbox.
'-------------------------------------------------------
Private Sub LoadConfiguredSourceData()

    On Error GoTo ErrHandler

    Dim loSource As ListObject
    Dim arrSource As Variant
    Dim arrOut() As Variant
    Dim idxId As Long
    Dim idxDesc As Long
    Dim rowCount As Long
    Dim validCount As Long
    Dim r As Long

    SafeAppGuardBegin

    mHasLookupData = False
    EraseLookupData

    Set loSource = FindListObjectInWorkbook(mSourceTableName, ThisWorkbook)
    If loSource Is Nothing Then
        Err.Raise vbObjectError + 7130, FORM_NAME & ".LoadConfiguredSourceData", _
                  "Could not locate source table '" & mSourceTableName & "'."
    End If

    idxId = GetTableColumnIndex(loSource, mIdHeader)
    idxDesc = GetTableColumnIndex(loSource, mDescHeader)

    If idxId = 0 Then
        Err.Raise vbObjectError + 7131, FORM_NAME & ".LoadConfiguredSourceData", _
                  "Header '" & mIdHeader & "' was not found in source table '" & mSourceTableName & "'."
    End If

    If idxDesc = 0 Then
        Err.Raise vbObjectError + 7132, FORM_NAME & ".LoadConfiguredSourceData", _
                  "Header '" & mDescHeader & "' was not found in source table '" & mSourceTableName & "'."
    End If

    Me.lbxPricebook.Clear
    Me.txtSearch.Value = vbNullString

    If loSource.DataBodyRange Is Nothing Then GoTo CleanExit

    arrSource = loSource.DataBodyRange.Value2
    rowCount = UBound(arrSource, 1)

    ReDim arrOut(1 To rowCount, 1 To 2)
    
    For r = 1 To rowCount
        If Not IsValueBlank(arrSource(r, idxId)) Then
            validCount = validCount + 1
            arrOut(validCount, 1) = NzText(arrSource(r, idxId))
            arrOut(validCount, 2) = CoerceDateForDisplay(arrSource(r, idxDesc), "dd-mmm-yyyy")
        End If
    Next r
    
    If validCount = 0 Then GoTo CleanExit
    
    '-------------------------------------------------------
    ' Trim to exact row count using a new array because
    ' VBA cannot ReDim Preserve the first dimension
    '-------------------------------------------------------
    If validCount < rowCount Then
        Dim arrTrimmed() As Variant
        Dim i As Long
    
        ReDim arrTrimmed(1 To validCount, 1 To 2)
    
        For i = 1 To validCount
            arrTrimmed(i, 1) = arrOut(i, 1)
            arrTrimmed(i, 2) = arrOut(i, 2)
        Next i
    
        arrOut = arrTrimmed
    End If

    Sort2DArrayByTwoKeys arrOut, 1, 2, True, True

    mLookupData = ToZeroBased2D(arrOut)
    mHasLookupData = True

    LoadLookupListFromCache vbNullString

CleanExit:
    SafeAppGuardEnd
    Exit Sub

ErrHandler:
    SafeLogError FORM_NAME & ".LoadConfiguredSourceData", Err.Number, Err.Description
    MsgBox "Unable to load the lookup list." & vbCrLf & vbCrLf & _
           "Reason: " & Err.Description, vbExclamation, Me.Caption
    Resume CleanExit
End Sub

'-------------------------------------------------------
' LoadLookupListFromCache
' Description:
'   Filters cached lookup data and reloads the listbox.
'-------------------------------------------------------
Private Sub LoadLookupListFromCache(ByVal searchText As String)

    On Error GoTo ErrHandler

    Dim arrFiltered() As Variant
    Dim searchKey As String
    Dim idText As String
    Dim descText As String
    Dim rowCount As Long
    Dim hitCount As Long
    Dim r As Long
    Dim outRow As Long

    Me.lbxPricebook.Clear

    If Not mHasLookupData Then Exit Sub
    If Not IsArray(mLookupData) Then Exit Sub

    rowCount = UBound(mLookupData, 1) - LBound(mLookupData, 1) + 1
    If rowCount <= 0 Then Exit Sub

    searchKey = LCase$(Trim$(searchText))

    If Len(searchKey) = 0 Then
        Me.lbxPricebook.list = mLookupData
        Exit Sub
    End If

    For r = LBound(mLookupData, 1) To UBound(mLookupData, 1)
        idText = LCase$(NzText(mLookupData(r, 0)))
        descText = LCase$(NzText(mLookupData(r, 1)))

        If InStr(1, idText, searchKey, vbTextCompare) > 0 _
        Or InStr(1, descText, searchKey, vbTextCompare) > 0 Then
            hitCount = hitCount + 1
        End If
    Next r

    If hitCount = 0 Then Exit Sub

    ReDim arrFiltered(0 To hitCount - 1, 0 To 1)

    outRow = -1
    For r = LBound(mLookupData, 1) To UBound(mLookupData, 1)
        idText = LCase$(NzText(mLookupData(r, 0)))
        descText = LCase$(NzText(mLookupData(r, 1)))

        If InStr(1, idText, searchKey, vbTextCompare) > 0 _
        Or InStr(1, descText, searchKey, vbTextCompare) > 0 Then

            outRow = outRow + 1
            arrFiltered(outRow, 0) = mLookupData(r, 0)
            arrFiltered(outRow, 1) = mLookupData(r, 1)
        End If
    Next r

    Me.lbxPricebook.list = arrFiltered
    Exit Sub

ErrHandler:
    SafeLogError FORM_NAME & ".LoadLookupListFromCache", Err.Number, Err.Description
    MsgBox "Unable to refresh the search list." & vbCrLf & vbCrLf & _
           "Reason: " & Err.Description, vbExclamation, Me.Caption
End Sub

'=======================================================
' Helpers
'=======================================================

Private Sub PositionFormTopRight()

    On Error Resume Next

    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height * 0.125)
    Me.Left = Application.Left + Application.Width - Me.Width - (Application.Width * 0.125)
End Sub

Private Sub ResetFormState()

    mConfigKey = vbNullString
    mSourceTableName = vbNullString
    mPasteTableName = vbNullString
    mIdHeader = vbNullString
    mDescHeader = vbNullString
    mPasteColumnHeader = vbNullString

    mCaptionText = vbNullString
    mTaskLabelText = vbNullString
    mCodeLabelText = vbNullString
    mDescLabelText = vbNullString

    mIsInitialised = False
    mHasLookupData = False

    EraseLookupData
End Sub

Private Sub EraseLookupData()
    On Error Resume Next
    Erase mLookupData
    On Error GoTo 0
End Sub

Private Function ToZeroBased2D(ByVal arrIn As Variant) As Variant

    Dim rInL As Long, rInU As Long
    Dim cInL As Long, cInU As Long
    Dim r As Long, c As Long
    Dim arrOut() As Variant

    rInL = LBound(arrIn, 1)
    rInU = UBound(arrIn, 1)
    cInL = LBound(arrIn, 2)
    cInU = UBound(arrIn, 2)

    ReDim arrOut(0 To rInU - rInL, 0 To cInU - cInL)

    For r = rInL To rInU
        For c = cInL To cInU
            arrOut(r - rInL, c - cInL) = arrIn(r, c)
        Next c
    Next r

    ToZeroBased2D = arrOut
End Function

Public Function IsUserFormLoaded(ByVal formName As String) As Boolean
    Dim frm As Object
    For Each frm In VBA.UserForms
        If StrComp(frm.name, formName, vbTextCompare) = 0 Then
            IsUserFormLoaded = True
            Exit Function
        End If
    Next frm
End Function

Private Function IsDateHeader(ByVal headerText As String) As Boolean

    Dim s As String
    s = LCase$(Trim$(headerText))

    Select Case s
        Case "date", "dispatch date", "delivery date", "required date", "issued date"
            IsDateHeader = True
        Case Else
            IsDateHeader = False
    End Select

End Function
