Attribute VB_Name = "modStructTrace"
Sub Print_Structural_Material_Traceability()
Attribute Print_Structural_Material_Traceability.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Hide blank rows

   ActiveSheet.Range("$A$3:$V$3000").AutoFilter Field:=2, Criteria1:="<>"
    
' Unmerge C1 to U5 (Tracking Schedule title)
    Range("C1:U5").Select
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
    Columns("L:O").Select
    Range("L7").Activate
    Selection.EntireColumn.Hidden = True
    Columns("S:V").Select
    Range("S7").Activate
    Selection.EntireColumn.Hidden = True
        
' Re-merge C1 to U5 (Tracking Schedule title)
    Range("C1:U5").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = True
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
    strName = Replace(wsA.Name, " ", "")
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
