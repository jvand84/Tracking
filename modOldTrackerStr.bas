Attribute VB_Name = "modOldTrackerStr"
Option Explicit

'-------------------------------------------------------
' Module: Structural Mapping Helper
'
' Purpose:
'   Provides helper routines for converting legacy
'   structural MTO profile/grade combinations into the
'   equivalent new structural profile attributes.
'
' Main Outputs:
'   - Dictionary mapping old profile + old grade keys
'     to new structural attributes
'   - Resolver function to return the mapped attributes
'
' Public Procedures:
'   - BuildStructuralProfileAttributeMap
'   - TryResolveStructuralAttributes
'   - Example_UseStructuralProfileMap
'
' Dependencies / Assumptions:
'   - Source profile master table is supplied as a valid
'     ListObject
'   - HeaderMapFromListObject resolves required columns
'   - The new profile master is the source of truth
'   - Legacy plate profiles are resolved using the old
'     profile text + old grade/class combination
'
' Notes:
'   - Dictionary compare mode is text compare
'   - Mapping stores the NEW description/profile in the
'     output value array
'   - Duplicate keys are ignored intentionally so the
'     first valid encountered mapping is retained
'-------------------------------------------------------

'=========================================================
' BuildStructuralProfileAttributeMap
'
' Purpose:
'   Builds a lookup dictionary from the NEW structural
'   profile master table.
'
' Key:
'   old-profile | old-grade
'
' Item:
'   Variant array (1 To 6)
'       1 = Discipline
'       2 = Type
'       3 = Grade
'       4 = Size 1
'       5 = Size 2
'       6 = Profile (new description)
'
' Inputs:
'   - loNewProfiles: ListObject containing the new profile
'     master data
'
' Required headers in loNewProfiles:
'   - Discipline
'   - Type
'   - Description
'   - Size 1
'   - Size 2
'   - Class
'
' Behaviour:
'   - Adds a direct-match key for each profile based on
'     Description + Class
'   - Adds extra legacy plate keys for plate rows so that
'     old profiles such as "3PL" can map to new profile
'     descriptions such as "3PL CS 250"
'
' Notes:
'   - Empty DataBodyRange returns an empty dictionary
'   - Missing required headers will raise an error
'=========================================================
Public Function BuildStructuralProfileAttributeMap(ByVal loNewProfiles As ListObject) As Object

    '-------------------------------------------------------
    ' Variable declarations
    '-------------------------------------------------------
    Dim dict As Object
    Dim arr As Variant
    Dim hdr As Object
    Dim r As Long

    Dim idxDiscipline As Long
    Dim idxType As Long
    Dim idxDesc As Long
    Dim idxSize1 As Long
    Dim idxSize2 As Long
    Dim idxClass As Long

    Dim newDiscipline As String
    Dim newType As String
    Dim newDesc As String
    Dim newSize1 As String
    Dim newSize2 As String
    Dim newClass As String

    Dim oldProfileKey As String
    Dim valueArr(1 To 6) As Variant
    Dim legacyPlateProfile As String

    On Error GoTo ErrHandler

    '-------------------------------------------------------
    ' Initialise dictionary
    '-------------------------------------------------------
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    '-------------------------------------------------------
    ' Validate input table
    '-------------------------------------------------------
    If loNewProfiles Is Nothing Then
        Err.Raise vbObjectError + 5000, _
                  "BuildStructuralProfileAttributeMap", _
                  "loNewProfiles is Nothing."
    End If

    If loNewProfiles.DataBodyRange Is Nothing Then
        Set BuildStructuralProfileAttributeMap = dict
        Exit Function
    End If

    '-------------------------------------------------------
    ' Load table data and resolve header positions
    '-------------------------------------------------------
    arr = loNewProfiles.DataBodyRange.Value2
    Set hdr = HeaderMapFromListObject(loNewProfiles)

    idxDiscipline = GetRequiredHeaderIndex(hdr, "discipline", "BuildStructuralProfileAttributeMap")
    idxType = GetRequiredHeaderIndex(hdr, "type", "BuildStructuralProfileAttributeMap")
    idxDesc = GetRequiredHeaderIndex(hdr, "description", "BuildStructuralProfileAttributeMap")
    idxSize1 = GetRequiredHeaderIndex(hdr, "size1", "BuildStructuralProfileAttributeMap")
    idxSize2 = GetRequiredHeaderIndex(hdr, "size2", "BuildStructuralProfileAttributeMap")
    idxClass = GetRequiredHeaderIndex(hdr, "class", "BuildStructuralProfileAttributeMap")

    '-------------------------------------------------------
    ' Build lookup dictionary row by row
    '-------------------------------------------------------
    For r = 1 To UBound(arr, 1)

        newDiscipline = NzText(arr(r, idxDiscipline))
        newType = NzText(arr(r, idxType))
        newDesc = NzText(arr(r, idxDesc))
        newSize1 = NzText(arr(r, idxSize1))
        newSize2 = NzText(arr(r, idxSize2))
        newClass = NzText(arr(r, idxClass))

        If Len(newDesc) > 0 Then

            '-------------------------------------------------------
            ' Standard direct-match key:
            '   new description + normalized class/grade
            '-------------------------------------------------------
            oldProfileKey = MakeOldStructuralLookupKey(newDesc, newClass)

            valueArr(1) = newDiscipline
            valueArr(2) = newType
            valueArr(3) = newClass
            valueArr(4) = newSize1
            valueArr(5) = newSize2
            valueArr(6) = newDesc

            If Not dict.Exists(oldProfileKey) Then
                dict.Add oldProfileKey, valueArr
            End If

            '-------------------------------------------------------
            ' Plate special handling:
            '
            ' New descriptions may be:
            '   3PL CS 250
            '   3PL CS 350
            '   3PL 316SS
            '
            ' Old profile may only be:
            '   3PL
            '
            ' The legacy plate key is therefore built from:
            '   old profile text + old grade/class
            '-------------------------------------------------------
            If StrComp(UCase$(newType), "PL", vbBinaryCompare) = 0 Then
                legacyPlateProfile = ExtractLegacyPlateProfile(newDesc)

                If Len(legacyPlateProfile) > 0 Then
                    oldProfileKey = MakeOldStructuralLookupKey(legacyPlateProfile, newClass)

                    valueArr(1) = newDiscipline
                    valueArr(2) = newType
                    valueArr(3) = newClass
                    valueArr(4) = newSize1
                    valueArr(5) = newSize2
                    valueArr(6) = newDesc

                    If Not dict.Exists(oldProfileKey) Then
                        dict.Add oldProfileKey, valueArr
                    End If
                End If
            End If

        End If
    Next r

    Set BuildStructuralProfileAttributeMap = dict
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "BuildStructuralProfileAttributeMap", Err.description

End Function

'=========================================================
' TryResolveStructuralAttributes
'
' Purpose:
'   Attempts to resolve an old profile + old grade pair
'   to the mapped new structural attributes.
'
' Inputs:
'   - oldProfile: legacy profile text
'   - oldGrade: legacy grade/class text
'   - dictMap: mapping dictionary previously built by
'     BuildStructuralProfileAttributeMap
'
' Outputs:
'   Returns:
'   - True if resolved
'   - False if not resolved
'
' ByRef outputs:
'   - outDiscipline
'   - outType
'   - outGrade
'   - outSize1
'   - outSize2
'   - outProfile
'
' Notes:
'   - Outputs are reset to blank before resolution
'   - A Nothing dictionary returns False safely
'=========================================================
Public Function TryResolveStructuralAttributes( _
    ByVal oldProfile As String, _
    ByVal oldGrade As String, _
    ByVal dictMap As Object, _
    ByRef outDiscipline As String, _
    ByRef outType As String, _
    ByRef outGrade As String, _
    ByRef outSize1 As String, _
    ByRef outSize2 As String, _
    ByRef outProfile As String) As Boolean

    '-------------------------------------------------------
    ' Variable declarations
    '-------------------------------------------------------
    Dim key As String
    Dim v As Variant

    On Error GoTo ErrHandler

    '-------------------------------------------------------
    ' Default outputs
    '-------------------------------------------------------
    outDiscipline = vbNullString
    outType = vbNullString
    outGrade = vbNullString
    outSize1 = vbNullString
    outSize2 = vbNullString
    outProfile = vbNullString
    TryResolveStructuralAttributes = False

    '-------------------------------------------------------
    ' Validate dictionary
    '-------------------------------------------------------
    If dictMap Is Nothing Then Exit Function

    '-------------------------------------------------------
    ' Resolve lookup key and return mapped values
    '-------------------------------------------------------
    key = MakeOldStructuralLookupKey(oldProfile, oldGrade)

    If dictMap.Exists(key) Then
        v = dictMap(key)

        outDiscipline = NzText(v(1))
        outType = NzText(v(2))
        outGrade = NzText(v(3))
        outSize1 = NzText(v(4))
        outSize2 = NzText(v(5))
        outProfile = NzText(v(6))

        TryResolveStructuralAttributes = True
    End If

    Exit Function

ErrHandler:
    Err.Raise Err.Number, "TryResolveStructuralAttributes", Err.description

End Function

'=========================================================
' Example_UseStructuralProfileMap
'
' Purpose:
'   Demonstration / debug routine showing how to build the
'   mapping dictionary and resolve each row from an old
'   structural MTO table.
'
' Assumptions:
'   - Worksheet "Old Structural MTO" contains table
'     "tblOldStructural"
'   - Worksheet "Profile Master" contains table
'     "tblProfiles"
'   - Old table contains headers:
'       Profile
'       Grade
'
' Notes:
'   - Intended as a test/debug helper
'   - Writes results to the Immediate Window
'=========================================================
Public Sub Example_UseStructuralProfileMap()

    '-------------------------------------------------------
    ' Variable declarations
    '-------------------------------------------------------
    Dim loOld As ListObject
    Dim loNewProfiles As ListObject
    Dim dictMap As Object

    Dim arrOld As Variant
    Dim hdrOld As Object
    Dim r As Long

    Dim idxOldProfile As Long
    Dim idxOldGrade As Long

    Dim discipline As String
    Dim typ As String
    Dim grade As String
    Dim size1 As String
    Dim size2 As String
    Dim profile As String

    On Error GoTo ErrHandler

    '-------------------------------------------------------
    ' Resolve source tables
    '-------------------------------------------------------
    Set loOld = ThisWorkbook.Worksheets("Old Structural MTO").ListObjects("tblOldStructural")
    Set loNewProfiles = ThisWorkbook.Worksheets("Profile Master").ListObjects("tblProfiles")

    '-------------------------------------------------------
    ' Build dictionary from new profile master
    '-------------------------------------------------------
    Set dictMap = BuildStructuralProfileAttributeMap(loNewProfiles)

    If loOld.DataBodyRange Is Nothing Then Exit Sub

    '-------------------------------------------------------
    ' Read old structural table into memory
    '-------------------------------------------------------
    arrOld = loOld.DataBodyRange.Value2
    Set hdrOld = HeaderMapFromListObject(loOld)

    idxOldProfile = GetRequiredHeaderIndex(hdrOld, "profile", "Example_UseStructuralProfileMap")
    idxOldGrade = GetRequiredHeaderIndex(hdrOld, "grade", "Example_UseStructuralProfileMap")

    '-------------------------------------------------------
    ' Test row-by-row resolution
    '-------------------------------------------------------
    For r = 1 To UBound(arrOld, 1)

        If TryResolveStructuralAttributes( _
            NzText(arrOld(r, idxOldProfile)), _
            NzText(arrOld(r, idxOldGrade)), _
            dictMap, _
            discipline, typ, grade, size1, size2, profile) Then

            Debug.Print "Row " & r & _
                        " | Discipline=" & discipline & _
                        " | Type=" & typ & _
                        " | Grade=" & grade & _
                        " | Size1=" & size1 & _
                        " | Size2=" & size2 & _
                        " | Profile=" & profile
        Else
            Debug.Print "Row " & r & " UNMAPPED: " & _
                        NzText(arrOld(r, idxOldProfile)) & " | " & NzText(arrOld(r, idxOldGrade))
        End If

    Next r

    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "Example_UseStructuralProfileMap", Err.description

End Sub

'=========================================================
' MakeOldStructuralLookupKey
'
' Purpose:
'   Builds the normalized dictionary key used for legacy
'   structural lookup.
'
' Key format:
'   normalized-profile | normalized-grade
'=========================================================
Private Function MakeOldStructuralLookupKey(ByVal oldProfile As String, _
                                            ByVal oldGrade As String) As String

    MakeOldStructuralLookupKey = NormalizeStructuralProfileText(oldProfile) & "|" & _
                                 NormalizeStructuralGradeText(oldGrade)

End Function

'=========================================================
' NormalizeStructuralProfileText
'
' Purpose:
'   Normalizes profile text for reliable lookup matching.
'
' Behaviour:
'   - trims text
'   - converts tabs / non-breaking spaces to spaces
'   - removes asterisks
'   - collapses repeated spaces
'   - returns uppercase
'=========================================================
Private Function NormalizeStructuralProfileText(ByVal s As String) As String

    Dim t As String

    t = NzText(s)
    t = Replace(t, vbTab, " ")
    t = Replace(t, Chr$(160), " ")
    t = Replace(t, "*", "")

    Do While InStr(1, t, "  ", vbBinaryCompare) > 0
        t = Replace(t, "  ", " ")
    Loop

    NormalizeStructuralProfileText = UCase$(Trim$(t))

End Function

'=========================================================
' NormalizeStructuralGradeText
'
' Purpose:
'   Normalizes grade/class text for reliable lookup
'   matching.
'
' Behaviour:
'   Common mappings:
'   - anything containing 316 -> 316SS
'   - anything containing 350 / C350 -> 350
'   - anything containing 250 / C250 -> 250
'=========================================================
Private Function NormalizeStructuralGradeText(ByVal s As String) As String

    Dim t As String

    t = UCase$(Trim$(NzText(s)))
    t = Replace(t, " ", "")

    Select Case True
        Case InStr(1, t, "316", vbBinaryCompare) > 0
            NormalizeStructuralGradeText = "316SS"

        Case InStr(1, t, "350", vbBinaryCompare) > 0 _
          Or InStr(1, t, "C350", vbBinaryCompare) > 0
            NormalizeStructuralGradeText = "350"

        Case InStr(1, t, "250", vbBinaryCompare) > 0 _
          Or InStr(1, t, "C250", vbBinaryCompare) > 0
            NormalizeStructuralGradeText = "250"

        Case Else
            NormalizeStructuralGradeText = t
    End Select

End Function

'=========================================================
' ExtractLegacyPlateProfile
'
' Purpose:
'   Extracts the legacy plate profile token from a new
'   plate description.
'
' Example:
'   Input:  "3PL CS 250"
'   Output: "3PL"
'
' Notes:
'   - Returns blank if no token can be extracted
'=========================================================
Private Function ExtractLegacyPlateProfile(ByVal newPlateDesc As String) As String

    Dim t As String
    Dim p As Long

    t = NormalizeStructuralProfileText(newPlateDesc)
    p = InStr(1, t, " ", vbBinaryCompare)

    If p > 1 Then
        ExtractLegacyPlateProfile = Left$(t, p - 1)
    Else
        ExtractLegacyPlateProfile = vbNullString
    End If

End Function

'=========================================================
' HeaderMapFromListObject
'
' Purpose:
'   Builds a header dictionary from a ListObject where:
'     Key   = normalized header name
'     Item  = 1-based column index
'
' Notes:
'   - Duplicate normalized headers will be overwritten by
'     the last matching column, which is standard VBA
'     dictionary behaviour
'=========================================================
Private Function HeaderMapFromListObject(ByVal lo As ListObject) As Object

    Dim d As Object
    Dim i As Long
    Dim hdrName As String

    On Error GoTo ErrHandler

    If lo Is Nothing Then
        Err.Raise vbObjectError + 5010, _
                  "HeaderMapFromListObject", _
                  "ListObject reference is Nothing."
    End If

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    For i = 1 To lo.ListColumns.Count
        hdrName = NormalizeHeader(lo.ListColumns(i).Name)
        d(hdrName) = i
    Next i

    Set HeaderMapFromListObject = d
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "HeaderMapFromListObject", Err.description

End Function

'=========================================================
' GetRequiredHeaderIndex
'
' Purpose:
'   Returns a required header index from the supplied
'   header map dictionary.
'
' Inputs:
'   - hdr: header dictionary
'   - normalizedHeaderName: normalized header key to find
'   - sourceProc: calling procedure name for better errors
'
' Behaviour:
'   - Raises a controlled error if the header is missing
'=========================================================
Private Function GetRequiredHeaderIndex(ByVal hdr As Object, _
                                        ByVal normalizedHeaderName As String, _
                                        ByVal sourceProc As String) As Long

    If hdr Is Nothing Then
        Err.Raise vbObjectError + 5020, _
                  sourceProc, _
                  "Header dictionary is Nothing."
    End If

    If Not hdr.Exists(normalizedHeaderName) Then
        Err.Raise vbObjectError + 5021, _
                  sourceProc, _
                  "Required header '" & normalizedHeaderName & "' was not found."
    End If

    GetRequiredHeaderIndex = CLng(hdr(normalizedHeaderName))

End Function

'=========================================================
' NormalizeHeader
'
' Purpose:
'   Normalizes header text so source headers can be looked
'   up consistently regardless of spaces, underscores, or
'   hyphens.
'=========================================================
Private Function NormalizeHeader(ByVal s As String) As String

    Dim t As String

    t = LCase$(Trim$(s))
    t = Replace(t, " ", "")
    t = Replace(t, "_", "")
    t = Replace(t, "-", "")

    NormalizeHeader = t

End Function

'=========================================================
' NzText
'
' Purpose:
'   Safe text coercion helper.
'
' Behaviour:
'   - Error values return blank
'   - Null values return blank
'   - Other values return trimmed string
'=========================================================
Private Function NzText(ByVal v As Variant) As String

    If IsError(v) Then
        NzText = vbNullString
    ElseIf IsNull(v) Then
        NzText = vbNullString
    Else
        NzText = Trim$(CStr(v))
    End If

End Function

'-------------------------------------------------------
' Test_StructuralMapping_FromSelectedWorkbook
'
' Purpose:
'   - Prompts user to select an OPEN workbook via frmSelect
'   - Reads old structural MTO from the selected workbook
'   - Uses tbl_MatSpec in ThisWorkbook as the mapping source
'   - Resolves structural attributes using the mapping helper
'   - Writes results to a temporary test sheet in ThisWorkbook
'
' Output Sheet:
'   zz_Test_StructuralMap
'
' Output Columns:
'   Old Profile
'   Old Grade
'   Description
'   Discipline
'   Type
'   Grade
'   Size 1
'   Size 2
'   New Profile
'   Status
'
' Requirements:
'   - frmSelect available
'   - GetOpenWorkbookByName available
'   - FindListObjectByName available
'   - BuildStructuralProfileAttributeMap available
'   - TryResolveStructuralAttributes available
'   - HeaderMapFromListObject available
'   - GetRequiredHeaderIndex available
'   - NzText available
'
' Notes:
'   - Selected workbook must already be open
'   - Mapping source is tbl_MatSpec in ThisWorkbook
'   - Temp sheet is deleted/recreated on each run
'-------------------------------------------------------
Public Sub Test_StructuralMapping_FromSelectedWorkbook()

    '-------------------------------------------------------
    ' Variable declarations
    '-------------------------------------------------------
    Dim frm As frmSelect
    Dim selectedName As String
    Dim srcWb As Workbook
    Dim toolWb As Workbook
    
    Dim wsTemp As Worksheet
    Dim loOld As ListObject
    Dim loProfiles As ListObject
    Dim dictMap As Object
    
    Dim arrOld As Variant
    Dim arrOut() As Variant
    Dim hdrOld As Object
    
    Dim r As Long
    Dim idxProfile As Long
    Dim idxGrade As Long
    Dim idxDesc As Long
    
    Dim discipline As String
    Dim typ As String
    Dim grade As String
    Dim size1 As String
    Dim size2 As String
    Dim profile As String
    
    Dim oldProfile As String
    Dim oldGrade As String
    Dim oldDesc As String
    
    Const TEMP_SHEET As String = "zz_Test_StructuralMap"
    
    On Error GoTo ErrHandler
    
    Set toolWb = ThisWorkbook
    
    '-------------------------------------------------------
    ' Prompt user to select the source workbook
    '-------------------------------------------------------
    Set frm = New frmSelect
    frm.FrmType = 1
    frm.LoadCombo
    frm.Show
    
    selectedName = Trim$(frm.SelectedWorkbookName)
    
    Unload frm
    Set frm = Nothing
    
    If Len(selectedName) = 0 Then
        MsgBox "No workbook selected. Exiting.", vbExclamation
        Exit Sub
    End If
    
    '-------------------------------------------------------
    ' Resolve selected source workbook
    '-------------------------------------------------------
    Set srcWb = GetOpenWorkbookByName(selectedName)
    
    If srcWb Is Nothing Then
        Err.Raise vbObjectError + 6000, _
                  "Test_StructuralMapping_FromSelectedWorkbook", _
                  "The selected workbook '" & selectedName & "' is not open or could not be resolved."
    End If
    
    '-------------------------------------------------------
    ' Resolve source old structural MTO table
    '-------------------------------------------------------
    Set loOld = srcWb.Worksheets("Structural Material Take Off").ListObjects("tmp_SMTO")
    
    If loOld Is Nothing Then
        Err.Raise vbObjectError + 6001, _
                  "Test_StructuralMapping_FromSelectedWorkbook", _
                  "Table 'tblOldStructural' was not found in worksheet 'Old Structural MTO' in workbook '" & srcWb.Name & "'."
    End If
    
    If loOld.DataBodyRange Is Nothing Then
        MsgBox "Old Structural MTO table is empty.", vbExclamation
        Exit Sub
    End If
    
    '-------------------------------------------------------
    ' Resolve mapping source table from ThisWorkbook
    '-------------------------------------------------------
    Set loProfiles = FindListObjectByName(toolWb, "tbl_MatSpec")
    
    If loProfiles Is Nothing Then
        Err.Raise vbObjectError + 6002, _
                  "Test_StructuralMapping_FromSelectedWorkbook", _
                  "Table 'tbl_MatSpec' was not found in ThisWorkbook."
    End If
    
    If loProfiles.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 6003, _
                  "Test_StructuralMapping_FromSelectedWorkbook", _
                  "Table 'tbl_MatSpec' is empty in ThisWorkbook."
    End If
    
    '-------------------------------------------------------
    ' Build mapping dictionary from tbl_MatSpec
    '-------------------------------------------------------
    Set dictMap = BuildStructuralProfileAttributeMap(loProfiles)
    
    '-------------------------------------------------------
    ' Read source old structural table into memory
    '-------------------------------------------------------
    arrOld = loOld.DataBodyRange.Value2
    Set hdrOld = HeaderMapFromListObject(loOld)
    
    idxProfile = GetRequiredHeaderIndex(hdrOld, "profile", "Test_StructuralMapping_FromSelectedWorkbook")
    idxGrade = GetRequiredHeaderIndex(hdrOld, "grade", "Test_StructuralMapping_FromSelectedWorkbook")
    idxDesc = GetRequiredHeaderIndex(hdrOld, "description", "Test_StructuralMapping_FromSelectedWorkbook")
    
    '-------------------------------------------------------
    ' Prepare output array
    '-------------------------------------------------------
    ReDim arrOut(1 To UBound(arrOld, 1), 1 To 10)
    
    '-------------------------------------------------------
    ' Resolve each old row against mapping dictionary
    '-------------------------------------------------------
    For r = 1 To UBound(arrOld, 1)
        
        oldProfile = NzText(arrOld(r, idxProfile))
        oldGrade = NzText(arrOld(r, idxGrade))
        oldDesc = NzText(arrOld(r, idxDesc))
        
        If TryResolveStructuralAttributes( _
            oldProfile, _
            oldGrade, _
            dictMap, _
            discipline, _
            typ, _
            grade, _
            size1, _
            size2, _
            profile) Then
            
            arrOut(r, 1) = oldProfile
            arrOut(r, 2) = oldGrade
            arrOut(r, 3) = oldDesc
            arrOut(r, 4) = discipline
            arrOut(r, 5) = typ
            arrOut(r, 6) = grade
            arrOut(r, 7) = size1
            arrOut(r, 8) = size2
            arrOut(r, 9) = profile
            arrOut(r, 10) = "OK"
            
        Else
            
            arrOut(r, 1) = oldProfile
            arrOut(r, 2) = oldGrade
            arrOut(r, 3) = oldDesc
            arrOut(r, 10) = "UNMAPPED"
            
        End If
        
    Next r
    
    '-------------------------------------------------------
    ' Recreate temp sheet in ThisWorkbook
    '-------------------------------------------------------
    On Error Resume Next
    Application.DisplayAlerts = False
    toolWb.Worksheets(TEMP_SHEET).Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrHandler
    
    Set wsTemp = toolWb.Worksheets.Add(After:=toolWb.Worksheets(toolWb.Worksheets.Count))
    wsTemp.Name = TEMP_SHEET
    
    '-------------------------------------------------------
    ' Write output headers
    '-------------------------------------------------------
    wsTemp.Range("A1:J1").Value = Array( _
        "Old Profile", _
        "Old Grade", _
        "Description", _
        "Discipline", _
        "Type", _
        "Grade", _
        "Size 1", _
        "Size 2", _
        "New Profile", _
        "Status")
    
    '-------------------------------------------------------
    ' Write output data in a single hit
    '-------------------------------------------------------
    wsTemp.Range("A2").Resize(UBound(arrOut, 1), UBound(arrOut, 2)).Value = arrOut
    
    '-------------------------------------------------------
    ' Basic formatting
    '-------------------------------------------------------
    With wsTemp
        .Rows(1).Font.Bold = True
        .Columns.AutoFit
    End With
    
    MsgBox "Structural mapping test complete." & vbCrLf & _
           "Source workbook: " & srcWb.Name & vbCrLf & _
           "Rows processed: " & UBound(arrOut, 1), vbInformation
    
    Exit Sub

ErrHandler:
    MsgBox "Test_StructuralMapping_FromSelectedWorkbook failed:" & vbCrLf & _
           Err.description, vbCritical, "Error"
End Sub

