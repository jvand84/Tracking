Attribute VB_Name = "modExport"

Option Explicit

' ===== CONFIG =====
Const EXPORT_PATH As String = "C:\Users\jvand\OneDrive\Walz\998 VBA Projects\Tracking\src\"
' ==================

' Export all VBA modules, classes, and forms to the EXPORT_PATH folder
Public Sub ExportVBAModules()
    Dim vbComp As Object
    Dim filePath As String
    Dim newExport_Path As String
    
    If Dir(EXPORT_PATH, vbDirectory) = "" Then
        MsgBox "Export folder not found: " & EXPORT_PATH, vbCritical
        Exit Sub
    End If
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1 ' Standard module, Class, Form
                newExport_Path = EXPORT_PATH & "standard_modules\"
                filePath = newExport_Path & vbComp.name & FileExtension(vbComp.Type)
                vbComp.Export filePath
            Case 2
                newExport_Path = EXPORT_PATH & "class_modules\"
                filePath = newExport_Path & vbComp.name & FileExtension(vbComp.Type)
                vbComp.Export filePath
            Case 3
                newExport_Path = EXPORT_PATH & "userforms\"
                filePath = newExport_Path & vbComp.name & FileExtension(vbComp.Type)
                vbComp.Export filePath
        End Select
    Next vbComp
    
    MsgBox "Modules exported to: " & EXPORT_PATH, vbInformation
End Sub

' Import all .bas/.cls/.frm files from EXPORT_PATH into the workbook (robust)
Public Sub ImportVBAModules_Safe()
    Dim fso As Object, folder As Object, file As Object
    Dim vbComps As Object, vbComp As Object
    Dim compName As String, ext As String
    Dim okCount As Long, failCount As Long, skipCount As Long
    Dim logText As String, logPath As String
    Dim ts As Object

    Const CT_STD As Long = 1          ' vbext_ct_StdModule
    Const CT_CLASS As Long = 2        ' vbext_ct_ClassModule
    Const CT_MSFORM As Long = 3       ' vbext_ct_MSForm
    Const CT_DOC As Long = 100        ' vbext_ct_Document

    On Error GoTo HardFail

    '--- sanity: export folder present?
    If Len(EXPORT_PATH) = 0 Or Dir(EXPORT_PATH, vbDirectory) = "" Then
        MsgBox "Import folder not found: " & EXPORT_PATH, vbCritical
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(EXPORT_PATH)

    '--- sanity: can we access the VBProject?
    On Error Resume Next
    Set vbComps = ThisWorkbook.VBProject.VBComponents
    If Err.Number <> 0 Then
        Dim errNum As Long, eMsg As String
        errNum = Err.Number: eMsg = Err.Description
        On Error GoTo 0
        MsgBox "Cannot access VBA project (" & errNum & "): " & eMsg & vbCrLf & vbCrLf & _
               "Enable: File > Options > Trust Center > Trust Center Settings... > " & _
               "Macro Settings > check “Trust access to the VBA project object model”.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    logText = "Import started: " & Now & vbCrLf & "Source: " & EXPORT_PATH & vbCrLf & String(60, "-") & vbCrLf

    '--- iterate files
    For Each file In folder.Files
        ext = LCase$(fso.GetExtensionName(file.name))
        If ext = "bas" Or ext = "cls" Or ext = "frm" Then
            compName = fso.GetBaseName(file.name)

            ' Try find existing component with same name
            Set vbComp = Nothing
            On Error Resume Next
            Set vbComp = vbComps(compName)
            On Error GoTo 0

            ' If exists, try to remove (but never remove document modules)
            If Not vbComp Is Nothing Then
                If vbComp.Type = CT_DOC Then
                    skipCount = skipCount + 1
                    logText = logText & "SKIP remove (Document module): " & compName & vbCrLf
                Else
                    On Error Resume Next
                    vbComps.Remove vbComp
                    If Err.Number <> 0 Then
                        failCount = failCount + 1
                        logText = logText & "FAIL remove " & compName & " -> " & Err.Number & ": " & Err.Description & vbCrLf
                        Err.Clear
                    Else
                        logText = logText & "OK   removed existing: " & compName & vbCrLf
                    End If
                    On Error GoTo 0
                End If
            End If

            ' Import new copy
            On Error Resume Next
            vbComps.Import file.Path
            If Err.Number <> 0 Then
                failCount = failCount + 1
                logText = logText & "FAIL import " & file.name & " -> " & Err.Number & ": " & Err.Description & vbCrLf
                Err.Clear
            Else
                okCount = okCount + 1
                logText = logText & "OK   import " & file.name & vbCrLf
            End If
            On Error GoTo 0
        End If
    Next file

    '--- write log file to the export folder
    logPath = fso.BuildPath(EXPORT_PATH, "ImportLog_" & Format(Now, "yyyymmdd_hhnnss") & ".txt")
    On Error Resume Next
    Set ts = fso.OpenTextFile(logPath, 2, True) ' ForWriting, Create:=True
    ts.Write logText
    ts.Close
    On Error GoTo 0

    '--- summary
    MsgBox "VBA import finished." & vbCrLf & vbCrLf & _
           "Imported: " & okCount & vbCrLf & _
           "Failed:   " & failCount & vbCrLf & _
           "Skipped:  " & skipCount & " (document modules not removed)" & vbCrLf & vbCrLf & _
           "Log: " & logPath, _
           IIf(failCount > 0, vbExclamation, vbInformation)

    Exit Sub

HardFail:
    MsgBox "Unexpected error (" & Err.Number & "): " & Err.Description, vbCritical
End Sub



' Helper function to get correct file extension
Private Function FileExtension(compType As Long) As String
    Select Case compType
        Case 1: FileExtension = ".bas" ' Module
        Case 2: FileExtension = ".cls" ' Class
        Case 3: FileExtension = ".frm" ' Form
        Case Else: FileExtension = ".txt"
    End Select
End Function

Public Function CleanNameLocal(ByVal s As String) As String
    Dim t As String, i As Long, ch As String, o As String
    t = Trim$(LCase$(s))
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch Like "[a-z0-9%]" Then o = o & ch
    Next i
    CleanNameLocal = o
End Function
