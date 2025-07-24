Attribute VB_Name = "modSync"
Option Explicit

'--- modSync.bas ---
'Exports every VBA component to sub-folders inside a "src" tree and re-imports
'them.  **Requires the workbook to be opened from a local or synced drive path.**
'If the file is opened directly from SharePoint/Teams via an https:// URL the
'user is warned and the operation is cancelled (no silent fallback).
'
'Folder layout – mirrors the VBE tree
'  src\Objects\        sheet / ThisWorkbook modules (.cls)
'  src\Modules\        standard modules (.bas)
'  src\ClassModules\   class modules (.cls)
'  src\Forms\          UserForms (.frm + .frx)
'
'Each export also writes (or refreshes) helper Git files:
'  - **.gitattributes**  (by default at the **repo root** next to the workbook)
'  - **.gitignore**      (same location)
'  - **README.md**       (only if it does *not* already exist)
'
'Set WRITE_GIT_AT_ROOT = False if you prefer those files inside src\ instead.
'
'Workbook/worksheet (Document) modules that are effectively empty (only
'"Option Explicit" and whitespace) are **skipped on export** to avoid clutter.
'On export, files that no longer correspond to any component (deleted or
'emptied) are **removed** from disk to keep the src tree tidy.

Const SRC_ROOT As String = "src"                   'name of the export root
Const GIT_ATTRIB As String = ".gitattributes"
Const GIT_IGNORE As String = ".gitignore"
Const README_FILE As String = "README.md"
Const WRITE_GIT_AT_ROOT As Boolean = True          'place git files at workbook folder

'====================  Ribbon wrappers  ====================
Public Sub ExportProject(control As Object)
    DoExportProject
End Sub

Public Sub ImportProject(control As Object)
    DoImportProject
End Sub

'====================  MAIN ROUTINES  ======================
Public Sub DoExportAddin()
    Dim wb As Workbook: Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Sub

    Dim rootPath As String: rootPath = GetRootPath(wb)   '…\src\
    If rootPath = "" Then Exit Sub
    Dim repoPath As String: repoPath = GetRepoPath(wb)   '…\ (workbook folder)

    Dim exported As Object: Set exported = CreateObject("Scripting.Dictionary")

    Dim comp As Object, subDir As String, fullPath As String
    For Each comp In wb.VBProject.VBComponents
        'Skip empty document modules (only Option Explicit / whitespace)
        If comp.Type = vbext_ct_Document Then
            If IsDocModuleEmpty(comp) Then GoTo NextComponent
        End If
        'Skip empty standard/class modules too
        If (comp.Type = vbext_ct_StdModule Or comp.Type = vbext_ct_ClassModule) _
           And IsCodeEmpty(comp) Then GoTo NextComponent

        subDir = rootPath & CompFolder(comp.Type) & "\"
        EnsureFolder subDir
        fullPath = subDir & comp.Name & Ext(comp.Type)
        comp.Export fullPath
        exported(AddSlash(fullPath)) = True
        If comp.Type = vbext_ct_MSForm Then
            exported(AddSlash(subDir & comp.Name & ".frx")) = True
        End If
NextComponent:
    Next

    PruneStaleFiles rootPath, exported

    If WRITE_GIT_AT_ROOT Then
        WriteGitAttributes repoPath
        WriteGitIgnore repoPath
        WriteReadme repoPath
    Else
        WriteGitAttributes rootPath
        WriteGitIgnore rootPath
        WriteReadme rootPath
    End If
End Sub

Public Sub DoExportProject()
    Dim wb As Workbook: Set wb = TargetWB()
    If wb Is Nothing Then Exit Sub

    Dim rootPath As String: rootPath = GetRootPath(wb)   '…\src\
    If rootPath = "" Then Exit Sub
    Dim repoPath As String: repoPath = GetRepoPath(wb)   '…\ (workbook folder)

    Dim exported As Object: Set exported = CreateObject("Scripting.Dictionary")

    Dim comp As Object, subDir As String, fullPath As String
    For Each comp In wb.VBProject.VBComponents
        'Skip empty document modules (only Option Explicit / whitespace)
        If comp.Type = vbext_ct_Document Then
            If IsDocModuleEmpty(comp) Then GoTo NextComponent
        End If
        'Skip empty standard/class modules too
        If (comp.Type = vbext_ct_StdModule Or comp.Type = vbext_ct_ClassModule) _
           And IsCodeEmpty(comp) Then GoTo NextComponent

        subDir = rootPath & CompFolder(comp.Type) & "\"
        EnsureFolder subDir
        fullPath = subDir & comp.Name & Ext(comp.Type)
        comp.Export fullPath                        'writes .frm+.frx automatically
        exported(AddSlash(fullPath)) = True
        If comp.Type = vbext_ct_MSForm Then
            exported(AddSlash(subDir & comp.Name & ".frx")) = True
        End If
NextComponent:
    Next

    'Remove files on disk that weren't (re)exported this run
    PruneStaleFiles rootPath, exported

    'Git helpers
    If WRITE_GIT_AT_ROOT Then
        WriteGitAttributes repoPath
        WriteGitIgnore repoPath
        WriteReadme repoPath
    Else
        WriteGitAttributes rootPath
        WriteGitIgnore rootPath
        WriteReadme rootPath
    End If
End Sub

Public Sub DoImportProject()
    Dim wb As Workbook: Set wb = TargetWB()
    If wb Is Nothing Then Exit Sub

    Dim rootPath As String: rootPath = GetRootPath(wb)
    If rootPath = "" Then Exit Sub
    If Dir(rootPath, vbDirectory) = "" Then
        MsgBox "Nothing to import – folder '" & rootPath & "' not found.", vbExclamation
        Exit Sub
    End If

    '-- remove all non-document components first
    Dim vc As Object
    For Each vc In wb.VBProject.VBComponents
        If vc.Type <> vbext_ct_Document Then wb.VBProject.VBComponents.Remove vc
    Next

    '-- iterate expected sub-folders
    Dim subFolder As Variant, f As String, vbComp As Object, filePath As String
    For Each subFolder In Array("Modules", "ClassModules", "Forms", "Objects", "Misc")
        filePath = rootPath & subFolder & "\"
        If Dir(filePath, vbDirectory) <> "" Then
            f = Dir(filePath & "*.*")
            Do While Len(f) > 0
                If LCase$(Right$(f, 4)) = ".frx" Then GoSub SkipFile 'ignore binary partner

                Dim tgtName As String: tgtName = Split(f, ".")(0)
                Set vbComp = Nothing
                On Error Resume Next
                Set vbComp = wb.VBProject.VBComponents(tgtName)
                On Error GoTo 0

                If vbComp Is Nothing Then
                    wb.VBProject.VBComponents.Import filePath & f
                Else
                    Dim txt As String
                    txt = CreateObject("Scripting.FileSystemObject") _
                          .OpenTextFile(filePath & f, 1).ReadAll
                    txt = CleanCode(txt)
                    With vbComp.CodeModule
                        .DeleteLines 1, .CountOfLines
                        .InsertLines 1, txt
                    End With
                End If
SkipFile:
                f = Dir
            Loop
        End If
    Next subFolder
End Sub

'====================  PATH / FILE HELPERS  ========================
Private Function GetRootPath(wb As Workbook) As String
    Dim p As String: p = wb.Path
    If p = "" Then
        MsgBox "Please save the workbook first.", vbExclamation
        Exit Function
    End If
    If LCase$(Left$(p, 4)) = "http" Then
        MsgBox "This workbook is open directly from SharePoint/Teams. " & _
               "Please open it from your local OneDrive sync folder or map it " & _
               "to a drive letter before running the export/import.", vbExclamation
        Exit Function
    End If
    If Right$(p, 1) <> "\" Then p = p & "\"
    p = p & SRC_ROOT & "\"
    EnsureFolder p
    GetRootPath = p
End Function

Private Function GetRepoPath(wb As Workbook) As String
    Dim p As String: p = wb.Path
    If Right$(p, 1) <> "\" Then p = p & "\"
    GetRepoPath = p
End Function

Private Sub EnsureFolder(fPath As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(fPath) Then fso.CreateFolder fPath
End Sub

'Delete any .bas/.cls/.frm/.frx file that wasn't exported this run
Private Sub PruneStaleFiles(rootPath As String, exported As Object)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim subFolder As Variant, folderPath As String
    For Each subFolder In Array("Modules", "ClassModules", "Forms", "Objects", "Misc")
        folderPath = rootPath & subFolder & "\"
        If fso.FolderExists(folderPath) Then
            Dim f As Object
            For Each f In fso.GetFolder(folderPath).Files
                Dim Ext As String: Ext = LCase$(fso.GetExtensionName(f.Path))
                If Ext = "bas" Or Ext = "cls" Or Ext = "frm" Or Ext = "frx" Then
                    If Not exported.Exists(AddSlash(f.Path)) Then
                        On Error Resume Next
                        f.Delete True
                        On Error GoTo 0
                    End If
                End If
            Next
        End If
    Next
End Sub

Private Function AddSlash(p As String) As String
    AddSlash = Replace$(p, "/", "\")  'normalise for dictionary keys
End Function

'====================  GIT FILE WRITERS  ======================
Private Sub WriteGitAttributes(basePath As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim aPath As String: aPath = basePath & GIT_ATTRIB

    Dim txt As String
    txt = "# Auto-generated by VBA Sync Add-in on " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
          "# Treat VBA text modules as LF-normalised text files" & vbCrLf & _
          "*.bas text eol=lf" & vbCrLf & _
          "*.cls text eol=lf" & vbCrLf & _
          "*.frm text eol=lf" & vbCrLf & _
          vbCrLf & _
          "# Binary partner of UserForms" & vbCrLf & _
          "*.frx binary" & vbCrLf & _
          vbCrLf & _
          "# Ignore diff for Excel workbooks" & vbCrLf & _
          "*.xls* binary" & vbCrLf

    Dim ts
    If fso.FileExists(aPath) Then
        Set ts = fso.OpenTextFile(aPath, 2)
    Else
        Set ts = fso.CreateTextFile(aPath, True)
    End If
    ts.Write txt
    ts.Close
End Sub

Private Sub WriteGitIgnore(basePath As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim iPath As String: iPath = basePath & GIT_IGNORE

    Dim txt As String
    txt = "# Auto-generated by VBA Sync Add-in on " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
          "# Ignore Excel/Office temp and cache files" & vbCrLf & _
          "~$*" & vbCrLf & _
          "*.tmp" & vbCrLf & _
          "*.bak" & vbCrLf & _
          "*.log" & vbCrLf & _
          "*.ldb" & vbCrLf & _
          "*.laccdb" & vbCrLf & _
          "*.asd" & vbCrLf & _
          "*.wbk" & vbCrLf & _
          vbCrLf & _
          "# Office autosave / lock files" & vbCrLf & _
          "*.owner" & vbCrLf & _
          vbCrLf & _
          "# OS cruft" & vbCrLf & _
          "Thumbs.db" & vbCrLf & _
          ".DS_Store" & vbCrLf & _
          vbCrLf & _
          "# IDE/project folders (optional)" & vbCrLf & _
          ".vs/" & vbCrLf & _
          ".idea/" & vbCrLf

    Dim ts
    If fso.FileExists(iPath) Then
        Set ts = fso.OpenTextFile(iPath, 2)
    Else
        Set ts = fso.CreateTextFile(iPath, True)
    End If
    ts.Write txt
    ts.Close
End Sub

Private Sub WriteReadme(basePath As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim rPath As String: rPath = basePath & README_FILE
    If fso.FileExists(rPath) Then Exit Sub   'do NOT overwrite

    Dim txt As String
    txt = "# Excel VBA Git Sync" & vbCrLf & vbCrLf & _
          "This repo mirrors the VBA project of the workbook in the `src/` folder, " & _
          "so you can version, diff, and review code comfortably in Git/VS Code." & vbCrLf & vbCrLf & _
          "## Workflow" & vbCrLf & _
          "1. Open the workbook locally (not directly over SharePoint/Teams https URLs)." & vbCrLf & _
          "2. Click **VBA Sync ? Export** to write all modules to `src/` (sub-folders mirror the VBE tree)." & vbCrLf & _
          "3. Work in VS Code / your Git client (lint, AI assist, diffs, PRs)." & vbCrLf & _
          "4. Click **VBA Sync ? Import** to push changes back into the workbook." & vbCrLf & vbCrLf & _
          "## Structure" & vbCrLf & _
          "``""" & vbCrLf & _
          "src/" & vbCrLf & _
          "  Modules/        ' .bas" & vbCrLf & _
          "  ClassModules/   ' .cls" & vbCrLf & _
          "  Forms/          ' .frm + .frx" & vbCrLf & _
          "  Objects/        ' ThisWorkbook / Sheets as .cls (code only)" & vbCrLf & _
          "``""" & vbCrLf & vbCrLf & _
          "## Notes" & vbCrLf & _
          "- Empty document modules (only `Option Explicit`) aren’t exported." & vbCrLf & _
          "- Export removes files that no longer exist in the project." & vbCrLf & _
          "- `.gitattributes` and `.gitignore` are auto-generated (by default at the repo root)." & vbCrLf & _
          "- SharePoint/Teams URLs are blocked: open the synced local copy instead." & vbCrLf

    Dim ts
    Set ts = fso.CreateTextFile(rPath, True)
    ts.Write txt
    ts.Close
End Sub

'====================  COMPONENT HELPERS  ===================
Private Function CompFolder(t As Long) As String
    Select Case t
        Case vbext_ct_StdModule:     CompFolder = "Modules"
        Case vbext_ct_ClassModule:   CompFolder = "ClassModules"
        Case vbext_ct_MSForm:        CompFolder = "Forms"
        Case vbext_ct_Document:      CompFolder = "Objects"
        Case Else:                   CompFolder = "Misc"
    End Select
End Function

Private Function Ext(t As Long) As String
    Select Case t
        Case vbext_ct_StdModule:     Ext = ".bas"
        Case vbext_ct_ClassModule:   Ext = ".cls"
        Case vbext_ct_MSForm:        Ext = ".frm"   'paired .frx auto-exported
        Case vbext_ct_Document:      Ext = ".cls"   'sheet / ThisWorkbook code-behind
        Case Else:                   Ext = ".bas"
    End Select
End Function

'Return True if a Document module contains no code other than Option Explicit / whitespace
Private Function IsDocModuleEmpty(vbComp As Object) As Boolean
    Dim cm As Object: Set cm = vbComp.CodeModule
    Dim txt As String
    txt = cm.Lines(1, cm.CountOfLines)
    txt = CleanCode(txt)

    Dim ln As Variant, hasRealCode As Boolean
    For Each ln In Split(txt, vbCrLf)
        Dim t As String: t = Trim$(ln)
        If Len(t) > 0 And LCase$(t) <> "option explicit" Then
            hasRealCode = True
            Exit For
        End If
    Next
    IsDocModuleEmpty = Not hasRealCode
End Function

'Generic emptiness check for Std/Class modules
Private Function IsCodeEmpty(vbComp As Object) As Boolean
    Dim cm As Object: Set cm = vbComp.CodeModule
    Dim txt As String
    If cm.CountOfLines = 0 Then
        IsCodeEmpty = True
        Exit Function
    End If
    txt = cm.Lines(1, cm.CountOfLines)
    txt = CleanCode(txt)

    Dim ln As Variant
    For Each ln In Split(txt, vbCrLf)
        Dim t As String: t = Trim$(ln)
        If Len(t) > 0 And LCase$(t) <> "option explicit" Then
            IsCodeEmpty = False
            Exit Function
        End If
    Next
    IsCodeEmpty = True
End Function

'====================  UTILITY HELPERS  =====================
Private Function TargetWB() As Workbook
    Dim wb As Workbook: Set wb = Application.ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "No workbook is active.", vbExclamation
    ElseIf wb.IsAddin Then
        MsgBox "Switch to the workbook you want to export, not the add-in tab.", vbExclamation
        Set wb = Nothing
    End If
    Set TargetWB = wb
End Function

Private Function CleanCode(src As String) As String
    Dim ln As Variant, out$, inBegin As Boolean, t As String
    For Each ln In Split(src, vbCrLf)
        t = Trim$(ln)
        Select Case True
            Case t Like "VERSION *", t Like "Attribute VB_*":
            Case Left$(t, 5) = "BEGIN":                       inBegin = True
            Case inBegin And t = "END":                        inBegin = False
            Case inBegin:
            Case Else:                                          out = out & ln & vbCrLf
        End Select
    Next
    CleanCode = RTrim$(out)
End Function


