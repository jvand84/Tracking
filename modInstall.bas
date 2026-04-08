Attribute VB_Name = "modInstall"
Option Explicit

Public Sub UpdateEarnedValue()
    Populate_tbl_Install_From_Tracking_Pricebook_And_ROC
    Update_tbl_Install_Forecasts_From_Pricebook
    GreyOutNonMilestoneGates
End Sub

'============================================================
' Populate_tbl_Install_From_Tracking_Pricebook_And_ROC
'
' UPDATE:
'   - Adds "Earned $" = Earned Qty * (Pricebook[Project Sell Unit Rate])
'
' Notes:
'   - Pricebook table name: tblPricebook
'   - Uses CommKey() to normalise Comm Codes (fixes en-dash/spacing issues)
'============================================================
Public Sub Populate_tbl_Install_From_Tracking_Pricebook_And_ROC()

    Const LO_INSTALL As String = "tbl_Install"
    Const LO_TRACKING As String = "tbl_Tracking"
    Const LO_PRICEBOOK As String = "tbl_Pricebook"     ' <-- tbl_pricebook in your wording
    Const LO_ROC As String = "tbl_ROCMilestones"

    ' Install headers
    Const H_INSTALL_KEY As String = "Mark Number/ Assembly/ ID"
    Const H_INSTALL_COMMOD As String = "Commodity"
    Const H_INSTALL_UOM As String = "UOM"
    Const H_INSTALL_DRAWING As String = "Drawing No."
    Const H_INSTALL_DESC As String = "Description"
    Const H_INSTALL_QTY As String = "Qty"
    Const H_INSTALL_WEIGHT As String = "Weight"
    Const H_INSTALL_WORKPACK As String = "Workpack"
    Const H_INSTALL_PROGRESS_UNIT_QTY As String = "Progress Unit Qty"
    Const H_INSTALL_EARNED_QTY As String = "Earned Qty"
    Const H_INSTALL_PCT As String = "%"
    Const H_INSTALL_EARNED_HRS As String = "Earned Hrs"
    Const H_INSTALL_EARNED_DOLLARS As String = "Earned $"   ' <-- NEW

    ' Tracking headers
    Const H_TRACK_KEY As String = "Asset Number"
    Const H_TRACK_DRAWING As String = "Drawing No."
    Const H_TRACK_DESC As String = "Description/Tag Number"
    Const H_TRACK_QTY As String = "Assembly Quantity"
    Const H_TRACK_WEIGHT As String = "MTO Weight (kg)"
    Const H_TRACK_WORKPACK As String = "Workpack"

    ' Pricebook headers
    Const H_PRICE_COMMOD As String = "Comm Code"
    Const H_PRICE_UOM As String = "UOM"
    Const H_PRICE_HRS_PER_UNIT As String = "HRS-Total / unit"
    Const H_PRICE_SELL_RATE As String = "Project Sell Unit Rate"  ' <-- NEW used
    Const H_PRICE_ROC_1 As String = "RulesOfCredit_idx"
    Const H_PRICE_ROC_2 As String = "Rules Of Credit"
    Const H_PRICE_ROC_3 As String = "ROC"
    Const H_PRICE_ROC_4 As String = "RulesOfCredit"

    ' ROC Milestones headers
    Const H_ROC_KEY As String = "RulesOfCredit_idx"
    Const H_ROC_WEIGHT As String = "Weighting"
    Const H_ROC_SEQ As String = "Sequence"
    Const H_ROC_VISIBLE As String = "Visible"

    Dim loI As ListObject, loT As ListObject, loP As ListObject, loR As ListObject
    Dim iArr As Variant, tArr As Variant, pArr As Variant, rArr As Variant

    Dim dictT As Object, dictP As Object, dictROC As Object
    Dim gateColBySeq As Object

    Dim r As Long, iCol As Long

    ' Install col indexes
    Dim cI_Key As Long, cI_Com As Long, cI_UOM As Long
    Dim cI_Draw As Long, cI_Desc As Long, cI_Qty As Long, cI_Wt As Long, cI_WP As Long
    Dim cI_ProgUnit As Long, cI_Earned As Long, cI_Pct As Long, cI_EarnedHrs As Long, cI_EarnedDollars As Long

    ' Tracking col indexes
    Dim cT_Key As Long, cT_Draw As Long, cT_Desc As Long, cT_Qty As Long, cT_Wt As Long, cT_WP As Long

    ' Pricebook col indexes
    Dim cP_Com As Long, cP_UOM As Long, cP_ROC As Long, cP_HrsUnit As Long, cP_SellRate As Long

    ' ROC col indexes
    Dim cR_Key As Long, cR_Weight As Long, cR_Seq As Long, cR_Visible As Long

    ' Working vars
    Dim k As String, commodKey As String, rocKey As String
    Dim vT As Variant, vP As Variant
    Dim milestones As Collection, m As Variant
    Dim seq As Long, w As Double
    Dim gateQty As Double
    Dim earnedSum As Double
    Dim qty As Double, progUnit As Double, denom As Double
    Dim hrsPerUnit As Double, sellRate As Double

    ' Counters
    Dim updatedTracking As Long, missingTracking As Long
    Dim updatedUOM As Long, missingUOM As Long
    Dim updatedEarned As Long, missingROC As Long, missingMilestones As Long

    On Error GoTo Fail
    AppGuard_Begin

    '========================================================
    ' Locate tables (scan workbook)
    '========================================================
    Set loI = FindListObjectByName(ThisWorkbook, LO_INSTALL)
    Set loT = FindListObjectByName(ThisWorkbook, LO_TRACKING)
    Set loP = FindListObjectByName(ThisWorkbook, LO_PRICEBOOK)
    Set loR = FindListObjectByName(ThisWorkbook, LO_ROC)

    If loI Is Nothing Then Err.Raise 5, , "Couldn't find table '" & LO_INSTALL & "'."
    If loT Is Nothing Then Err.Raise 5, , "Couldn't find table '" & LO_TRACKING & "'."
    If loP Is Nothing Then Err.Raise 5, , "Couldn't find table '" & LO_PRICEBOOK & "'."
    If loR Is Nothing Then Err.Raise 5, , "Couldn't find table '" & LO_ROC & "'."

    '========================================================
    ' Column indexes
    '========================================================
    cI_Key = colIndex(loI, H_INSTALL_KEY)
    cI_Com = colIndex(loI, H_INSTALL_COMMOD)
    cI_UOM = colIndex(loI, H_INSTALL_UOM)
    cI_Draw = colIndex(loI, H_INSTALL_DRAWING)
    cI_Desc = colIndex(loI, H_INSTALL_DESC)
    cI_Qty = colIndex(loI, H_INSTALL_QTY)
    cI_Wt = colIndex(loI, H_INSTALL_WEIGHT)
    cI_WP = colIndex(loI, H_INSTALL_WORKPACK)
    cI_ProgUnit = colIndex(loI, H_INSTALL_PROGRESS_UNIT_QTY)
    cI_Earned = colIndex(loI, H_INSTALL_EARNED_QTY)
    cI_Pct = colIndex(loI, H_INSTALL_PCT)
    cI_EarnedHrs = colIndex(loI, H_INSTALL_EARNED_HRS)
    cI_EarnedDollars = colIndex(loI, H_INSTALL_EARNED_DOLLARS) ' NEW

    cT_Key = colIndex(loT, H_TRACK_KEY)
    cT_Draw = colIndex(loT, H_TRACK_DRAWING)
    cT_Desc = colIndex(loT, H_TRACK_DESC)
    cT_Qty = colIndex(loT, H_TRACK_QTY)
    cT_Wt = colIndex(loT, H_TRACK_WEIGHT)
    cT_WP = colIndex(loT, H_TRACK_WORKPACK)

    cP_Com = colIndex(loP, H_PRICE_COMMOD)
    cP_UOM = colIndex(loP, H_PRICE_UOM)
    cP_HrsUnit = colIndex(loP, H_PRICE_HRS_PER_UNIT)
    cP_SellRate = colIndex(loP, H_PRICE_SELL_RATE)
    cP_ROC = ColIndexAny(loP, Array(H_PRICE_ROC_1, H_PRICE_ROC_2, H_PRICE_ROC_3, H_PRICE_ROC_4))

    cR_Key = colIndex(loR, H_ROC_KEY)
    cR_Weight = colIndex(loR, H_ROC_WEIGHT)
    cR_Seq = colIndex(loR, H_ROC_SEQ)
    cR_Visible = ColIndexOptional(loR, H_ROC_VISIBLE)

    '========================================================
    ' Load arrays
    '========================================================
    If loI.DataBodyRange Is Nothing Then GoTo Cleanup
    iArr = loI.DataBodyRange.Value2

    If Not loT.DataBodyRange Is Nothing Then tArr = loT.DataBodyRange.Value2 Else tArr = Empty
    If Not loP.DataBodyRange Is Nothing Then pArr = loP.DataBodyRange.Value2 Else pArr = Empty
    If Not loR.DataBodyRange Is Nothing Then rArr = loR.DataBodyRange.Value2 Else rArr = Empty

    '========================================================
    ' Detect Gate{n}-Qty columns in tbl_Install
    '========================================================
    Set gateColBySeq = CreateObject("Scripting.Dictionary")
    gateColBySeq.CompareMode = vbTextCompare

    For iCol = 1 To loI.ListColumns.Count
        seq = GateSeqFromHeader(loI.ListColumns(iCol).Name)
        If seq > 0 Then gateColBySeq(CStr(seq)) = iCol
    Next iCol

    '========================================================
    ' Build Tracking dictionary
    '========================================================
    Set dictT = CreateObject("Scripting.Dictionary")
    dictT.CompareMode = vbTextCompare

    If Not IsEmpty(tArr) Then
        For r = 1 To UBound(tArr, 1)
            k = KeyOf(tArr(r, cT_Key))
            If Len(k) > 0 Then
                dictT(k) = Array(tArr(r, cT_Draw), tArr(r, cT_Desc), tArr(r, cT_Qty), tArr(r, cT_Wt), tArr(r, cT_WP))
            End If
        Next r
    End If

    '========================================================
    ' Build Pricebook dictionary (CommKey -> Array(UOM, ROC, HrsPerUnit, SellRate))
    '========================================================
    Set dictP = CreateObject("Scripting.Dictionary")
    dictP.CompareMode = vbTextCompare

    If Not IsEmpty(pArr) Then
        For r = 1 To UBound(pArr, 1)
            commodKey = CommKey(pArr(r, cP_Com))
            If Len(commodKey) > 0 Then
                dictP(commodKey) = Array( _
                    pArr(r, cP_UOM), _
                    pArr(r, cP_ROC), _
                    pArr(r, cP_HrsUnit), _
                    pArr(r, cP_SellRate) _
                )
            End If
        Next r
    End If

    '========================================================
    ' Build ROC dictionary
    '========================================================
    Set dictROC = CreateObject("Scripting.Dictionary")
    dictROC.CompareMode = vbTextCompare

    If Not IsEmpty(rArr) Then
        For r = 1 To UBound(rArr, 1)
            rocKey = KeyOf(rArr(r, cR_Key))
            If Len(rocKey) > 0 Then

                If cR_Visible > 0 Then
                    If Not CBoolSafe(rArr(r, cR_Visible), True) Then GoTo NextROCRow
                End If

                If Not dictROC.Exists(rocKey) Then
                    Set milestones = New Collection
                    dictROC.Add rocKey, milestones
                Else
                    Set milestones = dictROC(rocKey)
                End If

                seq = CLngSafe(rArr(r, cR_Seq), 0)
                w = WeightNorm(rArr(r, cR_Weight))
                milestones.Add Array(seq, w)
            End If
NextROCRow:
        Next r
    End If

    '========================================================
    ' Apply to tbl_Install
    '========================================================
    For r = 1 To UBound(iArr, 1)

        ' (1) Tracking populate
        k = KeyOf(iArr(r, cI_Key))
        If Len(k) > 0 And dictT.Exists(k) Then
            vT = dictT(k)
            iArr(r, cI_Draw) = vT(0)
            iArr(r, cI_Desc) = vT(1)
            iArr(r, cI_Qty) = vT(2)
            iArr(r, cI_Wt) = vT(3)
            iArr(r, cI_WP) = vT(4)
            updatedTracking = updatedTracking + 1
        ElseIf Len(k) > 0 Then
            missingTracking = missingTracking + 1
        End If

        ' (2) Pricebook populate (UOM + ROC + Hrs/Unit + SellRate)
        commodKey = CommKey(iArr(r, cI_Com))
        rocKey = vbNullString
        hrsPerUnit = 0#
        sellRate = 0#

        If Len(commodKey) > 0 And dictP.Exists(commodKey) Then
            vP = dictP(commodKey)
            iArr(r, cI_UOM) = vP(0)
            rocKey = KeyOf(vP(1))
            hrsPerUnit = CDblSafe(vP(2), 0#)
            sellRate = CDblSafe(vP(3), 0#)
            updatedUOM = updatedUOM + 1
        ElseIf Len(commodKey) > 0 Then
            missingUOM = missingUOM + 1
        End If

        ' (3) Earned Qty + % + Earned Hrs + Earned $
        earnedSum = 0#
        progUnit = CDblSafe(iArr(r, cI_ProgUnit), 0#)
        qty = CDblSafe(iArr(r, cI_Qty), 0#)

        If progUnit > 0# And qty > 0# And Len(rocKey) > 0 And dictROC.Exists(rocKey) Then

            Set milestones = dictROC(rocKey)

            For Each m In milestones
                seq = CLngSafe(m(0), 0)
                w = CDblSafe(m(1), 0#)
                If seq > 0 And gateColBySeq.Exists(CStr(seq)) Then
                    gateQty = CDblSafe(iArr(r, gateColBySeq(CStr(seq))), 0#)
                    earnedSum = earnedSum + (gateQty * w * progUnit)
                End If
            Next m

            iArr(r, cI_Earned) = earnedSum

            denom = progUnit * qty
            If denom > 0# Then
                iArr(r, cI_Pct) = earnedSum / denom
            Else
                iArr(r, cI_Pct) = 0#
            End If

            iArr(r, cI_EarnedHrs) = earnedSum * hrsPerUnit
            iArr(r, cI_EarnedDollars) = earnedSum * sellRate   ' <-- NEW

            updatedEarned = updatedEarned + 1

        ElseIf Len(rocKey) > 0 Then
            If Not dictROC.Exists(rocKey) Then missingMilestones = missingMilestones + 1
        ElseIf Len(commodKey) > 0 And dictP.Exists(commodKey) Then
            missingROC = missingROC + 1
        End If

    Next r

    '========================================================
    ' Writeback
    '========================================================
    loI.DataBodyRange.Value2 = iArr

Cleanup:
    Debug.Print "Populate_tbl_Install complete:"
    Debug.Print "  Tracking updated: " & updatedTracking & " | missing: " & missingTracking
    Debug.Print "  UOM updated:      " & updatedUOM & " | missing: " & missingUOM
    Debug.Print "  Earned updated:   " & updatedEarned & " | missing ROC: " & missingROC & " | missing milestones: " & missingMilestones
    AppGuard_End
    Exit Sub

Fail:
    Debug.Print "Populate_tbl_Install FAILED: " & Err.Number & " - " & Err.description
    AppGuard_End
    MsgBox "Populate_tbl_Install failed: " & Err.description, vbExclamation, "tbl_Install Populate"
End Sub

'============================================================
' NEW helper: CommKey (normalise commodity codes)
'============================================================
Private Function CommKey(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Or IsEmpty(v) Then Exit Function

    s = CStr(v)
    s = Replace(s, ChrW(160), " ")        ' NBSP -> space
    s = Trim$(s)

    ' Normalise dashes
    s = Replace(s, ChrW(8211), "-")       ' en dash
    s = Replace(s, ChrW(8212), "-")       ' em dash
    s = Replace(s, ChrW(8722), "-")       ' minus sign

    ' Remove internal spaces (handles "410 - 1")
    s = Replace(s, " ", vbNullString)

    CommKey = s
End Function





Private Function colIndex(ByVal lo As ListObject, ByVal headerName As String) As Long
    Dim i As Long, want As String
    want = NormalizeHeader(headerName)
    For i = 1 To lo.ListColumns.Count
        If NormalizeHeader(lo.ListColumns(i).Name) = want Then
            colIndex = i
            Exit Function
        End If
    Next i
    Err.Raise 5, , "Missing column '" & headerName & "' in table '" & lo.Name & "'."
End Function

Private Function ColIndexOptional(ByVal lo As ListObject, ByVal headerName As String) As Long
    On Error GoTo Nope
    ColIndexOptional = colIndex(lo, headerName)
    Exit Function
Nope:
    ColIndexOptional = 0
End Function

Private Function ColIndexAny(ByVal lo As ListObject, ByVal headerCandidates As Variant) As Long
    Dim i As Long
    For i = LBound(headerCandidates) To UBound(headerCandidates)
        On Error Resume Next
        ColIndexAny = colIndex(lo, CStr(headerCandidates(i)))
        If Err.Number = 0 And ColIndexAny > 0 Then
            On Error GoTo 0
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
    Next i
    Err.Raise 5, , "Missing ROC column in table '" & lo.Name & "'. Tried: " & JoinVariant(headerCandidates, ", ")
End Function

Private Function NormalizeHeader(ByVal s As String) As String
    Dim t As String
    t = LCase$(Trim$(CStr(s)))
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    NormalizeHeader = t
End Function

Private Function GateSeqFromHeader(ByVal headerText As String) As Long
    Dim s As String, p As Long, i As Long, ch As String, digits As String
    s = NormalizeHeader(headerText)
    If InStr(1, s, "gate", vbTextCompare) = 0 Then Exit Function
    If InStr(1, s, "qty", vbTextCompare) = 0 Then Exit Function

    p = InStr(1, s, "gate", vbTextCompare)
    digits = vbNullString
    For i = p + 4 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "#" Then
            digits = digits & ch
        ElseIf Len(digits) > 0 Then
            Exit For
        End If
    Next i

    If Len(digits) = 0 Then Exit Function
    GateSeqFromHeader = CLng(digits)
End Function

Private Function WeightNorm(ByVal v As Variant) As Double
    Dim s As String, d As Double
    If IsError(v) Or IsEmpty(v) Then Exit Function
    s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function

    If Right$(s, 1) = "%" Then
        s = Left$(s, Len(s) - 1)
        d = CDblSafe(s, 0#) / 100#
    Else
        d = CDblSafe(s, 0#)
        If d > 1# Then d = d / 100#
    End If
    WeightNorm = d
End Function

Private Function KeyOf(ByVal v As Variant) As String
    If IsError(v) Then KeyOf = vbNullString Else KeyOf = Trim$(CStr(v))
End Function

Private Function CDblSafe(ByVal v As Variant, ByVal defaultValue As Double) As Double
    On Error GoTo Bad
    If IsError(v) Or IsEmpty(v) Or Len(Trim$(CStr(v))) = 0 Then
        CDblSafe = defaultValue
    Else
        CDblSafe = CDbl(v)
    End If
    Exit Function
Bad:
    CDblSafe = defaultValue
End Function

Private Function CLngSafe(ByVal v As Variant, ByVal defaultValue As Long) As Long
    On Error GoTo Bad
    If IsError(v) Or IsEmpty(v) Or Len(Trim$(CStr(v))) = 0 Then
        CLngSafe = defaultValue
    Else
        CLngSafe = CLng(v)
    End If
    Exit Function
Bad:
    CLngSafe = defaultValue
End Function

Private Function CBoolSafe(ByVal v As Variant, ByVal defaultValue As Boolean) As Boolean
    On Error GoTo Bad
    If IsError(v) Or IsEmpty(v) Or Len(Trim$(CStr(v))) = 0 Then
        CBoolSafe = defaultValue
    Else
        CBoolSafe = CBool(v)
    End If
    Exit Function
Bad:
    CBoolSafe = defaultValue
End Function

Private Function JoinVariant(ByVal v As Variant, ByVal delim As String) As String
    Dim i As Long, s As String
    For i = LBound(v) To UBound(v)
        If i > LBound(v) Then s = s & delim
        s = s & CStr(v(i))
    Next i
    JoinVariant = s
End Function


'============================================================
' GreyOutNonMilestoneGates
'
' Purpose:
'   Greys out the Gate n-Date and Gate n-Qty cells in tbl_Install
'   for any Gate numbers that DO NOT exist as milestones for the
'   row's ROC (from tbl_Pricebook -> RulesOfCredit_idx).
'
' Logic:
'   - Build Commodity -> ROC dictionary from tbl_Pricebook
'   - Build ROC -> GateMask(1..10) from tbl_ROCMilestones (Visible only if present)
'   - For each tbl_Install row:
'       For gate 1..10:
'         If gate not in ROC milestones -> grey out Gate n-Date + Gate n-Qty
'         Else -> clear grey (xlNone)
'
' Notes:
'   - Uses AppGuard_Begin / AppGuard_End (modGuardsAndTables)
'   - Uses your existing helpers: FindListObjectByName, ColIndex, ColIndexAny,
'     ColIndexOptional, NormalizeHeader, KeyOf, CLngSafe, CBoolSafe
'============================================================
Public Sub GreyOutNonMilestoneGates()

    Const LO_INSTALL As String = "tbl_Install"
    Const LO_PRICEBOOK As String = "tbl_Pricebook"
    Const LO_ROC As String = "tbl_ROCMilestones"

    ' Install headers
    Const H_INSTALL_COMMOD As String = "Commodity"

    ' Pricebook headers
    Const H_PRICE_COMMOD As String = "Comm Code"
    Const H_PRICE_ROC_1 As String = "RulesOfCredit_idx"
    Const H_PRICE_ROC_2 As String = "Rules Of Credit"
    Const H_PRICE_ROC_3 As String = "ROC"
    Const H_PRICE_ROC_4 As String = "RulesOfCredit"

    ' ROC headers
    Const H_ROC_KEY As String = "RulesOfCredit_idx"
    Const H_ROC_SEQ As String = "Sequence"
    Const H_ROC_VISIBLE As String = "Visible"

    Const GATE_MIN As Long = 1
    Const GATE_MAX As Long = 10

    Dim loI As ListObject, loP As ListObject, loR As ListObject
    Dim iArr As Variant, pArr As Variant, rArr As Variant
    Dim dictComToROC As Object, dictROCMask As Object

    Dim cI_Com As Long
    Dim cP_Com As Long, cP_ROC As Long
    Dim cR_Key As Long, cR_Seq As Long, cR_Visible As Long

    Dim gateQtyCol(1 To GATE_MAX) As Long
    Dim gateDateCol(1 To GATE_MAX) As Long

    Dim r As Long, iCol As Long, g As Long
    Dim commod As String, rocKey As String
    Dim mask As Variant
    Dim seq As Long

    Dim grayColor As Long
    grayColor = RGB(190, 190, 190)

    On Error GoTo Fail
    AppGuard_Begin

    '--------------------------------------------------------
    ' Locate tables
    '--------------------------------------------------------
    Set loI = FindListObjectByName(ThisWorkbook, LO_INSTALL)
    Set loP = FindListObjectByName(ThisWorkbook, LO_PRICEBOOK)
    Set loR = FindListObjectByName(ThisWorkbook, LO_ROC)

    If loI Is Nothing Then Err.Raise 5, , "Couldn't find table '" & LO_INSTALL & "'."
    If loP Is Nothing Then Err.Raise 5, , "Couldn't find table '" & LO_PRICEBOOK & "'."
    If loR Is Nothing Then Err.Raise 5, , "Couldn't find table '" & LO_ROC & "'."

    If loI.DataBodyRange Is Nothing Then GoTo Cleanup

    '--------------------------------------------------------
    ' Column indexes
    '--------------------------------------------------------
    cI_Com = colIndex(loI, H_INSTALL_COMMOD)

    cP_Com = colIndex(loP, H_PRICE_COMMOD)
    cP_ROC = ColIndexAny(loP, Array(H_PRICE_ROC_1, H_PRICE_ROC_2, H_PRICE_ROC_3, H_PRICE_ROC_4))

    cR_Key = colIndex(loR, H_ROC_KEY)
    cR_Seq = colIndex(loR, H_ROC_SEQ)
    cR_Visible = ColIndexOptional(loR, H_ROC_VISIBLE)

    '--------------------------------------------------------
    ' Load arrays (fast reads)
    '--------------------------------------------------------
    iArr = loI.DataBodyRange.Value2
    If Not loP.DataBodyRange Is Nothing Then pArr = loP.DataBodyRange.Value2 Else pArr = Empty
    If Not loR.DataBodyRange Is Nothing Then rArr = loR.DataBodyRange.Value2 Else rArr = Empty

    '--------------------------------------------------------
    ' Detect Gate columns in tbl_Install
    '--------------------------------------------------------
    For iCol = 1 To loI.ListColumns.Count
        seq = GateSeqFromHeader_Qty(loI.ListColumns(iCol).Name)
        If seq >= GATE_MIN And seq <= GATE_MAX Then gateQtyCol(seq) = iCol

        seq = GateSeqFromHeader_Date(loI.ListColumns(iCol).Name)
        If seq >= GATE_MIN And seq <= GATE_MAX Then gateDateCol(seq) = iCol
    Next iCol

    '--------------------------------------------------------
    ' Build Commodity -> ROC dictionary
    '--------------------------------------------------------
    Set dictComToROC = CreateObject("Scripting.Dictionary")
    dictComToROC.CompareMode = vbTextCompare

    If Not IsEmpty(pArr) Then
        For r = 1 To UBound(pArr, 1)
            commod = CommKey(pArr(r, cP_Com))
            If Len(commod) > 0 Then
                dictComToROC(commod) = KeyOf(pArr(r, cP_ROC))
            End If
        Next r
    End If

    '--------------------------------------------------------
    ' Build ROC -> GateMask(1..10) dictionary
    '   mask(g)=True if Sequence g exists (Visible only if present)
    '--------------------------------------------------------
    Set dictROCMask = CreateObject("Scripting.Dictionary")
    dictROCMask.CompareMode = vbTextCompare

    If Not IsEmpty(rArr) Then
        For r = 1 To UBound(rArr, 1)

            rocKey = KeyOf(rArr(r, cR_Key))
            If Len(rocKey) = 0 Then GoTo NextROCRow

            If cR_Visible > 0 Then
                If Not CBoolSafe(rArr(r, cR_Visible), True) Then GoTo NextROCRow
            End If

            seq = CLngSafe(rArr(r, cR_Seq), 0)
            If seq < GATE_MIN Or seq > GATE_MAX Then GoTo NextROCRow

            If Not dictROCMask.Exists(rocKey) Then
                mask = MakeGateMaskFalse(GATE_MAX) ' 1..10 False
                dictROCMask.Add rocKey, mask
            Else
                mask = dictROCMask(rocKey)
            End If

            mask(seq) = True
            dictROCMask(rocKey) = mask

NextROCRow:
        Next r
    End If

    '--------------------------------------------------------
    ' Apply grey formatting per row based on ROC milestone mask
    '--------------------------------------------------------
    Dim cellQty As Range, cellDate As Range
    Dim hasMask As Boolean

    For r = 1 To UBound(iArr, 1)

        commod = CommKey(iArr(r, cI_Com))
        rocKey = vbNullString
        hasMask = False

        If Len(commod) > 0 And dictComToROC.Exists(commod) Then
            rocKey = KeyOf(dictComToROC(commod))
        End If

        If Len(rocKey) > 0 And dictROCMask.Exists(rocKey) Then
            mask = dictROCMask(rocKey)
            hasMask = True
        Else
            ' No ROC / no milestones -> treat as no milestones (grey everything we can)
            mask = MakeGateMaskFalse(GATE_MAX)
        End If

        For g = GATE_MIN To GATE_MAX

            ' Skip if the gate columns don't exist in the table
            If gateQtyCol(g) = 0 And gateDateCol(g) = 0 Then GoTo NextGate

            If gateQtyCol(g) > 0 Then
                Set cellQty = loI.DataBodyRange.Cells(r, gateQtyCol(g))
            Else
                Set cellQty = Nothing
            End If

            If gateDateCol(g) > 0 Then
                Set cellDate = loI.DataBodyRange.Cells(r, gateDateCol(g))
            Else
                Set cellDate = Nothing
            End If

            If mask(g) = False Then
                If Not cellQty Is Nothing Then cellQty.Interior.Color = grayColor
                If Not cellDate Is Nothing Then cellDate.Interior.Color = grayColor
            Else
                ' Gate is a valid milestone for this ROC -> clear grey
                If Not cellQty Is Nothing Then cellQty.Interior.Pattern = xlNone
                If Not cellDate Is Nothing Then cellDate.Interior.Pattern = xlNone
            End If

NextGate:
        Next g
    Next r

Cleanup:
    AppGuard_End
    Exit Sub

Fail:
    AppGuard_End
    MsgBox "GreyOutNonMilestoneGates failed: " & Err.description, vbExclamation, "Gate Formatting"
End Sub

'============================================================
' Helpers (local)
'============================================================

Private Function MakeGateMaskFalse(ByVal gateMax As Long) As Variant
    ' Returns a 1..gateMax Boolean array (all False) in a Variant
    Dim a() As Boolean
    Dim i As Long
    ReDim a(1 To gateMax)
    For i = 1 To gateMax
        a(i) = False
    Next i
    MakeGateMaskFalse = a
End Function

Private Function GateSeqFromHeader_Qty(ByVal headerText As String) As Long
    ' Detects: "Gate 1-Qty", "Gate 1 Qty", "Gate1-Qty" etc (must contain "qty")
    Dim s As String, p As Long, i As Long, ch As String, digits As String
    s = NormalizeHeader(headerText)
    If InStr(1, s, "gate", vbTextCompare) = 0 Then Exit Function
    If InStr(1, s, "qty", vbTextCompare) = 0 Then Exit Function

    p = InStr(1, s, "gate", vbTextCompare)
    digits = vbNullString

    For i = p + 4 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "#" Then
            digits = digits & ch
        ElseIf Len(digits) > 0 Then
            Exit For
        End If
    Next i

    If Len(digits) = 0 Then Exit Function
    GateSeqFromHeader_Qty = CLng(digits)
End Function

Private Function GateSeqFromHeader_Date(ByVal headerText As String) As Long
    ' Detects: "Gate 1-Date", "Gate 1 Date", "Gate1-Date" etc (must contain "date")
    Dim s As String, p As Long, i As Long, ch As String, digits As String
    s = NormalizeHeader(headerText)
    If InStr(1, s, "gate", vbTextCompare) = 0 Then Exit Function
    If InStr(1, s, "date", vbTextCompare) = 0 Then Exit Function

    p = InStr(1, s, "gate", vbTextCompare)
    digits = vbNullString

    For i = p + 4 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "#" Then
            digits = digits & ch
        ElseIf Len(digits) > 0 Then
            Exit For
        End If
    Next i

    If Len(digits) = 0 Then Exit Function
    GateSeqFromHeader_Date = CLng(digits)
End Function


'============================================================
' Update_tbl_Install_Forecasts_From_Pricebook
'
' tblPricebook provides:
'   - "HRS-Total / unit"
'   - "Project Sell Unit Rate"
'
' tbl_Install updates:
'   - "Forecast Hrs" = Progress Unit Qty * Qty * HRS-Total / unit
'   - "Forecast $"   = Progress Unit Qty * Qty * Project Sell Unit Rate
'
' NOTE:
'   Does NOT write "HRS-Total / unit" or "Project Sell Unit Rate" into tbl_Install,
'   because those columns are not present in loI.
'============================================================


Public Sub Update_tbl_Install_Forecasts_From_Pricebook()

    Const LO_INSTALL As String = "tbl_Install"
    Const LO_PRICEBOOK As String = "tbl_Pricebook"

    ' Install headers
    Const H_INSTALL_COMMOD As String = "Commodity"
    Const H_INSTALL_PROGRESS_UNIT_QTY As String = "Progress Unit Qty"
    Const H_INSTALL_QTY As String = "Qty"
    Const H_INSTALL_FORECAST_HRS As String = "Forecast Hrs"
    Const H_INSTALL_FORECAST_DOLLARS As String = "Forecast $"

    ' Pricebook headers
    Const H_PRICE_COMMOD As String = "Comm Code"
    Const H_PRICE_SELL_RATE As String = "Project Sell Unit Rate"
    Const H_PRICE_HRS_UNIT As String = "HRS-Total / unit"

    Dim loI As ListObject, loP As ListObject
    Dim iArr As Variant, pArr As Variant
    Dim dictP As Object

    Dim cI_Com As Long, cI_ProgUnit As Long, cI_Qty As Long
    Dim cI_ForecastHrs As Long, cI_ForecastDollars As Long

    Dim cP_Com As Long, cP_SellRate As Long, cP_HrsUnit As Long

    Dim r As Long
    Dim commodKey As String
    Dim vP As Variant
    Dim progUnit As Double, qty As Double
    Dim sellRate As Double, hrsUnit As Double

    Dim hits As Long, misses As Long

    On Error GoTo Fail
    AppGuard_Begin

    ' Locate tables
    Set loI = FindListObjectByName(ThisWorkbook, LO_INSTALL)
    Set loP = FindListObjectByName(ThisWorkbook, LO_PRICEBOOK)

    If loI Is Nothing Then Err.Raise 5, , "Couldn't find table '" & LO_INSTALL & "'."
    If loP Is Nothing Then Err.Raise 5, , "Couldn't find table '" & LO_PRICEBOOK & "'."
    If loI.DataBodyRange Is Nothing Then GoTo Cleanup

    ' Column indexes
    cI_Com = colIndex(loI, H_INSTALL_COMMOD)
    cI_ProgUnit = colIndex(loI, H_INSTALL_PROGRESS_UNIT_QTY)
    cI_Qty = colIndex(loI, H_INSTALL_QTY)
    cI_ForecastHrs = colIndex(loI, H_INSTALL_FORECAST_HRS)
    cI_ForecastDollars = colIndex(loI, H_INSTALL_FORECAST_DOLLARS)

    cP_Com = colIndex(loP, H_PRICE_COMMOD)
    cP_SellRate = colIndex(loP, H_PRICE_SELL_RATE)
    cP_HrsUnit = colIndex(loP, H_PRICE_HRS_UNIT)

    ' Load arrays
    iArr = loI.DataBodyRange.Value2
    If Not loP.DataBodyRange Is Nothing Then
        pArr = loP.DataBodyRange.Value2
    Else
        pArr = Empty
    End If

    ' Build Pricebook dictionary: CommKey -> Array(SellRate, HrsUnit)
    Set dictP = CreateObject("Scripting.Dictionary")
    dictP.CompareMode = vbTextCompare

    If Not IsEmpty(pArr) Then
        For r = 1 To UBound(pArr, 1)
            commodKey = CommKey(pArr(r, cP_Com))
            If Len(commodKey) > 0 Then
                dictP(commodKey) = Array(pArr(r, cP_SellRate), pArr(r, cP_HrsUnit))
            End If
        Next r
    End If

    ' Apply to Install
    For r = 1 To UBound(iArr, 1)

        commodKey = CommKey(iArr(r, cI_Com))
        progUnit = CDblSafe(iArr(r, cI_ProgUnit), 0#)
        qty = CDblSafe(iArr(r, cI_Qty), 0#)

        If Len(commodKey) > 0 And dictP.Exists(commodKey) Then

            vP = dictP(commodKey)
            sellRate = CDblSafe(vP(0), 0#)
            hrsUnit = CDblSafe(vP(1), 0#)

            iArr(r, cI_ForecastHrs) = progUnit * qty * hrsUnit
            iArr(r, cI_ForecastDollars) = progUnit * qty * sellRate

            hits = hits + 1
        Else
            iArr(r, cI_ForecastHrs) = 0#
            iArr(r, cI_ForecastDollars) = 0#
            misses = misses + 1
        End If

    Next r

    ' Writeback
    loI.DataBodyRange.Value2 = iArr

Cleanup:
    Debug.Print "Forecast update complete. Hits=" & hits & " Misses=" & misses
    AppGuard_End
    Exit Sub

Fail:
    Debug.Print "Update_tbl_Install_Forecasts_From_Pricebook FAILED: " & Err.Number & " - " & Err.description
    AppGuard_End
    MsgBox "Forecast update failed: " & Err.description, vbExclamation, "Forecast Update"
End Sub








'====================================================================================
' BuildExistingInstallLookup
'
' Reads the existing asset values from tbl_Install and loads them into a dictionary.
' Used to prevent duplicate inserts.
'====================================================================================
Public Sub BuildExistingInstallLookup(ByVal lo As ListObject, ByVal assetCol As Long, ByRef dict As Object)

    Dim arr As Variant
    Dim i As Long
    Dim v As String

    If lo Is Nothing Then Exit Sub
    If dict Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    arr = lo.ListColumns(assetCol).DataBodyRange.Value2

    If IsArray(arr) Then
        For i = 1 To UBound(arr, 1)
            v = Trim$(CStr(arr(i, 1)))
            If Len(v) > 0 Then
                If Not dict.Exists(v) Then dict.Add v, True
            End If
        Next i
    Else
        v = Trim$(CStr(arr))
        If Len(v) > 0 Then
            If Not dict.Exists(v) Then dict.Add v, True
        End If
    End If

End Sub









