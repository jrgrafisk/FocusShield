Attribute VB_Name = "modInstaller"
Option Explicit
' ONE-TIME INSTALLER - not part of the add-in itself.
'
' Import ONLY this one file (Filer > Importer fil... - just this file, no
' multi-select needed), then run ImportAllModules once (F5 with the cursor
' inside it). It imports every other .bas/.cls file in the "vba" folder for
' you, including ThisWorkbook's code - no more one-file-at-a-time importing.
'
' Requires "Trust access to the VBA project object model" to be turned on
' first (File > Options > Trust Center > Trust Center Settings > Macro
' Settings). It is a one-time checkbox - see the README for exact steps.
'
' You can delete this module afterwards (right-click it in Project
' Explorer > Remove modInstaller) - it is not part of the add-in itself.

Public Sub ImportAllModules()
    Dim proj As Object
    On Error Resume Next
    Set proj = ThisWorkbook.VBProject
    On Error GoTo 0
    If proj Is Nothing Then
        MsgBox DK("Kunne ikke tilg~aa~ VBA-projektet.") & vbLf & vbLf & _
              DK("Sl~aa~ ""Trust access to the VBA project object model"" til f~oe~rst: " & _
              "Filer > Indstillinger > Sikkerhedscenter > Indstillinger for " & _
              "Sikkerhedscenter > Makroindstillinger, s~ae~t flueben, tryk OK to gange, " & _
              "og pr~oe~v igen."), vbCritical, "Installation"
        Exit Sub
    End If

    Dim folder As String
    folder = PickFolder()
    If folder = "" Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folder) Then
        MsgBox "Mappen findes ikke.", vbCritical, "Installation"
        Exit Sub
    End If

    Dim imported As Long, skipped As Long, failed As String
    Dim file As Object
    For Each file In fso.GetFolder(folder).Files
        Dim ext As String, baseName As String
        ext = LCase$(fso.GetExtensionName(file.Name))
        baseName = LCase$(file.Name)
        If baseName = "_installer.bas" Then
            ' this file itself - already pasted in by hand, skip it
            skipped = skipped + 1
        ElseIf baseName = "thisworkbook.cls" Then
            On Error Resume Next
            proj.VBComponents("ThisWorkbook").CodeModule.AddFromFile file.Path
            If Err.Number <> 0 Then
                failed = failed & vbLf & "ThisWorkbook.cls (" & Err.Description & ")"
            Else
                imported = imported + 1
            End If
            On Error GoTo 0
        ElseIf ext = "bas" Or ext = "cls" Then
            On Error Resume Next
            proj.VBComponents.Import file.Path
            If Err.Number <> 0 Then
                failed = failed & vbLf & file.Name & " (" & Err.Description & ")"
            Else
                imported = imported + 1
            End If
            On Error GoTo 0
        Else
            skipped = skipped + 1
        End If
    Next file

    Dim msg As String
    msg = imported & " fil(er) importeret."
    If failed <> "" Then
        msg = msg & vbLf & vbLf & DK("Fejlede (import~ee~r disse manuelt via Filer > " & _
             "Import~ee~r fil):") & failed
    End If
    msg = msg & vbLf & vbLf & DK("Du kan nu slette dette modul (modInstaller) - h~oe~jreklik " & _
         "det i Project Explorer > Remove modInstaller.")
    msg = msg & vbLf & vbLf & DK("N~ae~ste skridt: gem projektmappen som Excel-tilf~oe~jelsesprogram " & _
         "(*.xlam), luk den, og aktiv~ee~r den under Filer > Indstillinger > " & _
         "Tilf~oe~jelsesprogrammer. Se README.md.")
    MsgBox msg, vbInformation, "Installation"
End Sub

Private Function PickFolder() As String
    Dim fd As Object
    Set fd = Application.FileDialog(4)  ' msoFileDialogFolderPicker
    fd.Title = DK("V~ae~lg mappen 'vba' (den du pakkede ud fra zip-filen)")
    If fd.Show <> -1 Then
        PickFolder = ""
        Exit Function
    End If
    PickFolder = fd.SelectedItems(1)
End Function
