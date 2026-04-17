Attribute VB_Name = "VBAExporter"
Option Explicit

' Diger modullere bagimlilik yok Ã¢â‚¬â€ self-contained.
' Requirement: Excel Options -> Trust Center -> Macro Settings
'              -> "Trust access to the VBA project object model" ON

Private Function SrcPath() As String
    SrcPath = ThisWorkbook.path & "\src\"
End Function

' ============================================
' IMPORT: src/ -> Excel
' ============================================

Public Sub ImportAllModules()
    On Error GoTo NoVBAccess
    Dim dummy As Object
    Set dummy = ThisWorkbook.VBProject.VBComponents
    On Error GoTo ErrorHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim importPath As String
    importPath = SrcPath()

    If Not fso.FolderExists(importPath) Then
        MsgBox "Import path not found: " & importPath, vbCritical, "VBAImporter"
        Exit Sub
    End If

    Dim proj As Object
    Set proj = ThisWorkbook.VBProject

    Dim file As Object
    Dim ext As String
    Dim moduleName As String
    Dim comp As Object
    Dim tempPath As String
    Dim importCount As Long, docCount As Long, skipCount As Long

    For Each file In fso.GetFolder(importPath).files
        ext = LCase(fso.GetExtensionName(file.name))
        moduleName = fso.GetBaseName(file.name)

        Select Case ext
        Case "bas", "cls"
            If moduleName = "VBAExporter" Then skipCount = skipCount + 1: GoTo NextFile

            Dim modCode As String
            modCode = StripHeader(ReadUTF8File(file.path))

            Set comp = Nothing
            On Error Resume Next
            Set comp = proj.VBComponents(moduleName)
            On Error GoTo ErrorHandler

            If Not comp Is Nothing Then
                ' Var olan modulu yerinde guncelle Remove+Import duplicate isim yaratir
                With comp.CodeModule
                    If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
                    .AddFromString modCode
                End With
            Else
                ' Yeni modul ANSI temp dosyasi uzerinden import
                tempPath = Environ("TEMP") & "\" & file.name
                ReencodeToANSI file.path, tempPath
                proj.VBComponents.Import tempPath
                On Error Resume Next: Kill tempPath: On Error GoTo ErrorHandler
            End If
            importCount = importCount + 1

        Case "frm"
            If moduleName = "VBAExporter" Then skipCount = skipCount + 1: GoTo NextFile

            tempPath = Environ("TEMP") & "\" & file.name
            ReencodeToANSI file.path, tempPath

            Dim frxSrc As String: frxSrc = importPath & moduleName & ".frx"
            If fso.FileExists(frxSrc) Then
                fso.CopyFile frxSrc, Environ("TEMP") & "\" & moduleName & ".frx", True
            End If

            Set comp = Nothing
            On Error Resume Next
            Set comp = proj.VBComponents(moduleName)
            On Error GoTo ErrorHandler
            If Not comp Is Nothing Then proj.VBComponents.Remove comp

            proj.VBComponents.Import tempPath
            importCount = importCount + 1

            On Error Resume Next
            Kill tempPath
            Kill Environ("TEMP") & "\" & moduleName & ".frx"
            On Error GoTo ErrorHandler

        Case "doccls"
            Dim docCode As String
            docCode = StripHeader(ReadUTF8File(file.path))

            Set comp = Nothing
            On Error Resume Next
            Set comp = proj.VBComponents(moduleName)
            On Error GoTo ErrorHandler

            If Not comp Is Nothing Then
                With comp.CodeModule
                    If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
                    If Len(Trim(docCode)) > 0 Then .AddFromString docCode
                End With
                docCount = docCount + 1
            End If
        End Select
NextFile:
    Next file

    MsgBox "Import tamamlandi!" & vbCrLf & vbCrLf & _
           "Modul: " & importCount & vbCrLf & _
           "Document modul: " & docCount & vbCrLf & _
           "Atlandi (self): " & skipCount, _
           vbInformation, "VBAImporter"
    Exit Sub

NoVBAccess:
    MsgBox "VBA Project erisimi reddedildi." & vbCrLf & vbCrLf & _
           "Excel Options -> Trust Center -> Trust Center Settings" & vbCrLf & _
           "-> Macro Settings -> Trust access to the VBA project object model", _
           vbCritical, "VBAImporter"
    Exit Sub

ErrorHandler:
    MsgBox "'" & moduleName & "' import hatasi:" & vbCrLf & Err.Description, vbCritical, "VBAImporter"
End Sub

' ============================================
' EXPORT: Excel -> src/
' ============================================

Public Sub ExportAllModulesUTF8()
    On Error GoTo NoVBAccess
    Dim dummy As Object
    Set dummy = ThisWorkbook.VBProject.VBComponents
    On Error GoTo ErrorHandler

    Dim exportPath As String
    exportPath = SrcPath()

    If Not CreateObject("Scripting.FileSystemObject").FolderExists(exportPath) Then
        If MsgBox("Export folder not found: " & exportPath & vbCrLf & "Create it?", vbYesNo) = vbYes Then
            MkDir exportPath
        Else
            Exit Sub
        End If
    End If

    Dim vbComp As Object
    Dim ext As String
    Dim tempPath As String
    Dim exportCount As Integer

    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        ext = GetExtension(vbComp.Type)
        If ext = "" Then GoTo NextComp

        tempPath = Environ("TEMP") & "\" & vbComp.name & ext

        On Error Resume Next
        vbComp.Export tempPath
        If Err.Number <> 0 Then Err.Clear: GoTo NextComp
        On Error GoTo ErrorHandler

        ReencodeToUTF8BOM tempPath, exportPath & vbComp.name & ext

        ' UserForm icin .frx binary dosyasini da src'ye kopyala
        If ext = ".frm" Then
            Dim frxTemp As String: frxTemp = Environ("TEMP") & "\" & vbComp.name & ".frx"
            If Len(Dir(frxTemp)) > 0 Then
                FileCopy frxTemp, exportPath & vbComp.name & ".frx"
                On Error Resume Next: Kill frxTemp: On Error GoTo ErrorHandler
            End If
        End If

        On Error Resume Next: Kill tempPath: On Error GoTo ErrorHandler
        exportCount = exportCount + 1
NextComp:
    Next vbComp

    MsgBox exportCount & " modules exported (UTF-8 BOM) to:" & vbCrLf & exportPath, vbInformation, "VBAExporter"
    Exit Sub

NoVBAccess:
    MsgBox "VBA Project access denied." & vbCrLf & vbCrLf & _
           "Excel Options -> Trust Center -> Trust Center Settings" & vbCrLf & _
           "-> Macro Settings -> Trust access to the VBA project object model", _
           vbCritical, "VBAExporter"
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "VBAExporter"
End Sub

' ============================================
' Helpers
' ============================================

Private Sub ReencodeToUTF8BOM(sourcePath As String, destPath As String)
    Dim fileNo As Integer, content As String
    fileNo = FreeFile
    Open sourcePath For Input As #fileNo
    content = Input(LOF(fileNo), fileNo)
    Close #fileNo

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .WriteText content
        .SaveToFile destPath, 2
        .Close
    End With
End Sub

Private Sub ReencodeToANSI(sourcePath As String, destPath As String)
    Dim content As String
    content = ReadUTF8File(sourcePath)
    Dim fileNo As Integer
    fileNo = FreeFile
    Open destPath For Output As #fileNo
    Print #fileNo, content;
    Close #fileNo
End Sub

Private Function ReadUTF8File(filePath As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = 2
        .Charset = "utf-8"
        .Open
        .LoadFromFile filePath
        ReadUTF8File = .ReadText
        .Close
    End With
End Function

Private Function GetExtension(compType As Integer) As String
    Select Case compType
        Case 1:   GetExtension = ".bas"
        Case 2:   GetExtension = ".cls"
        Case 3:   GetExtension = ".frm"
        Case 100: GetExtension = ".doccls"
        Case Else: GetExtension = ""
    End Select
End Function

Private Function StripHeader(code As String) As String
    ' VERSION/BEGIN..END/Attribute satirlarini soyar, VBA kodunu doner
    Dim lines() As String
    If InStr(code, vbCrLf) > 0 Then
        lines = Split(code, vbCrLf)
    Else
        lines = Split(code, vbLf)
    End If

    Dim i As Long, cnt As Long
    Dim result() As String
    ReDim result(UBound(lines))
    Dim state As String
    state = "header"

    For i = 0 To UBound(lines)
        Dim ln As String
        ln = lines(i)
        If right(ln, 1) = Chr(13) Then ln = left(ln, Len(ln) - 1)

        Select Case state
        Case "header"
            If ln Like "VERSION *" Or ln Like "Attribute VB_*" Then
            ElseIf ln = "BEGIN" Then
                state = "begin_block"
            Else
                state = "code"
                result(cnt) = ln: cnt = cnt + 1
            End If
        Case "begin_block"
            If ln = "END" Or ln = "End" Then state = "header"
        Case "code"
            result(cnt) = ln: cnt = cnt + 1
        End Select
    Next i

    If cnt = 0 Then Exit Function
    ReDim Preserve result(cnt - 1)
    StripHeader = Join(result, vbCrLf)
End Function
