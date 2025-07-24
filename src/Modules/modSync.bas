Attribute VB_Name = "modSync"
Option Explicit

'--- modSync.bas ---
'Exports every VBA component to sub-folders inside a "src" tree and re-imports
'them.  **Requires the workbook to be opened from a local or synced drive path.**
'If the file is opened directly from SharePoint/Teams via an https:// URL the
'user is warned and the operation is cancelled (no silent fallback).
'
'NEW: Also extracts Excel file structure (workbook.xml, table definitions,
'worksheet schemas) to enable version control and AI collaboration on both
'VBA code AND Excel data models.
'
'Folder layout - mirrors the VBE tree + Excel structure
'  src\Objects\        sheet / ThisWorkbook modules (.cls)
'  src\Modules\        standard modules (.bas)
'  src\ClassModules\   class modules (.cls)
'  src\Forms\          UserForms (.frm + .frx)
'  src\Excel\          Excel file structure (NEW)
'    \workbook.xml     workbook structure and sheet definitions
'    \tables\          table definitions (*.xml)
'    \worksheets\      worksheet schemas (sheet*.xml)
'
'Each export also writes (or refreshes) helper Git files:
'  - **.gitattributes**  (by default at the **repo root** next to the workbook)
'  - **.gitignore**      (same location)
'  - **README.md**       (only if it does *not* already exist)
'
'Set WRITE_GIT_AT_ROOT = False if you prefer those files inside src\ instead.
'Set EXTRACT_EXCEL_STRUCTURE = False to disable Excel structure extraction.
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
Const EXTRACT_EXCEL_STRUCTURE As Boolean = True    'NEW: extract Excel file structure

'====================  Ribbon wrappers  ====================
Public Sub ExportProject(control As Object)
    DoExportProject
End Sub

Public Sub ImportProject(control As Object)
    DoImportProject
End Sub

'====================  MAIN ROUTINES  ======================
Private Sub DoExportAddin()
    Dim wb As Workbook: Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Sub

    Dim rootPath As String: rootPath = GetRootPath(wb)   '...\src\
    If rootPath = "" Then Exit Sub
    Dim repoPath As String: repoPath = GetRepoPath(wb)   '...\ (workbook folder)

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
        fullPath = subDir & comp.Name & GetExt(comp.Type)
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

Private Sub DoExportProject()
    Dim wb As Workbook: Set wb = TargetWB()
    If wb Is Nothing Then Exit Sub

    Dim rootPath As String: rootPath = GetRootPath(wb)   '...\src\
    If rootPath = "" Then Exit Sub
    Dim repoPath As String: repoPath = GetRepoPath(wb)   '...\ (workbook folder)

    Dim exported As Object: Set exported = CreateObject("Scripting.Dictionary")

    ' Export VBA components
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
        fullPath = subDir & comp.Name & GetExt(comp.Type)
        comp.Export fullPath                        'writes .frm+.frx automatically
        exported(AddSlash(fullPath)) = True
        If comp.Type = vbext_ct_MSForm Then
            exported(AddSlash(subDir & comp.Name & ".frx")) = True
        End If
NextComponent:
    Next

    'NEW: Export Excel file structure
    If EXTRACT_EXCEL_STRUCTURE Then
        ExtractExcelStructure wb, rootPath, exported
    End If

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

'NEW: Extract Excel file structure for version control and collaboration
Private Sub ExtractExcelStructure(wb As Workbook, rootPath As String, exported As Object)
    On Error GoTo ExcelStructureError
    
    Dim excelDir As String: excelDir = rootPath & "Excel\"
    EnsureFolder excelDir
    
    ' Create temporary copy of workbook as ZIP
    Dim tempZip As String: tempZip = wb.Path & "\" & wb.Name & ".temp.zip"
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile wb.FullName, tempZip
    
    ' Extract Excel structure using Shell
    Dim tempExtract As String: tempExtract = wb.Path & "\temp_excel_extract"
    EnsureFolder tempExtract
    
    ' Use PowerShell to extract ZIP (more reliable than Shell.Application)
    Dim psCmd As String
    psCmd = "powershell -Command ""Expand-Archive '" & tempZip & "' -DestinationPath '" & tempExtract & "' -Force"""
    CreateObject("WScript.Shell").Run psCmd, 0, True
    
    ' Copy key Excel files to src/Excel/
    CopyExcelFile tempExtract & "\xl\workbook.xml", excelDir, "workbook.xml", exported
    
    ' Copy table definitions
    Dim tablesDir As String: tablesDir = excelDir & "tables\"
    If fso.FolderExists(tempExtract & "\xl\tables") Then
        EnsureFolder tablesDir
        Dim tableFile As Object
        For Each tableFile In fso.GetFolder(tempExtract & "\xl\tables").Files
            If LCase(fso.GetExtensionName(tableFile.Name)) = "xml" Then
                CopyExcelFile tableFile.Path, tablesDir, tableFile.Name, exported
            End If
        Next
    End If
    
    ' Copy worksheet structure (first 200 lines only to avoid huge data files)
    Dim worksheetsDir As String: worksheetsDir = excelDir & "worksheets\"
    If fso.FolderExists(tempExtract & "\xl\worksheets") Then
        EnsureFolder worksheetsDir
        Dim wsFile As Object
        For Each wsFile In fso.GetFolder(tempExtract & "\xl\worksheets").Files
            If LCase(fso.GetExtensionName(wsFile.Name)) = "xml" And wsFile.Name <> "_rels" Then
                CopyExcelFileWithLimit wsFile.Path, worksheetsDir, wsFile.Name, exported, 200
            End If
        Next
    End If
    
    ' Create Excel structure summary
    CreateExcelStructureSummary wb, excelDir, exported
    
    ' Cleanup temporary files
    On Error Resume Next
    fso.DeleteFile tempZip, True
    fso.DeleteFolder tempExtract, True
    On Error GoTo 0
    
    Exit Sub
    
ExcelStructureError:
    ' Cleanup on error
    On Error Resume Next
    If fso.FileExists(tempZip) Then fso.DeleteFile tempZip, True
    If fso.FolderExists(tempExtract) Then fso.DeleteFolder tempExtract, True
    On Error GoTo 0
    ' Continue without Excel structure if extraction fails
End Sub

Private Sub CopyExcelFile(sourcePath As String, destDir As String, fileName As String, exported As Object)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(sourcePath) Then
        Dim destPath As String: destPath = destDir & fileName
        fso.CopyFile sourcePath, destPath, True
        exported(AddSlash(destPath)) = True
    End If
End Sub

Private Sub CopyExcelFileWithLimit(sourcePath As String, destDir As String, fileName As String, exported As Object, maxLines As Long)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(sourcePath) Then
        Dim sourceText As String
        sourceText = fso.OpenTextFile(sourcePath, 1).ReadAll
        
        ' Limit to first N lines to avoid huge worksheet data files
        Dim Lines As Variant: Lines = Split(sourceText, vbCrLf)
        If UBound(Lines) > maxLines Then
            ReDim Preserve Lines(0 To maxLines)
            sourceText = Join(Lines, vbCrLf) & vbCrLf & _
                        "<!-- Truncated at " & maxLines & " lines by VBA Sync to avoid large files -->"
        End If
        
        Dim destPath As String: destPath = destDir & fileName
        Dim ts As Object: Set ts = fso.CreateTextFile(destPath, True)
        ts.Write sourceText
        ts.Close
        exported(AddSlash(destPath)) = True
    End If
End Sub

Private Sub CreateExcelStructureSummary(wb As Workbook, excelDir As String, exported As Object)
    Dim summaryPath As String: summaryPath = excelDir & "STRUCTURE_SUMMARY.md"
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim summary As String
    summary = "# Excel File Structure Summary" & vbCrLf & vbCrLf
    summary = summary & "Generated: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    summary = summary & "Workbook: " & wb.Name & vbCrLf & vbCrLf
    
    ' Worksheet summary
    summary = summary & "## Worksheets" & vbCrLf
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        summary = summary & "- **" & ws.Name & "**"
        If ws.UsedRange.Rows.Count > 1 Then
            summary = summary & " (" & ws.UsedRange.Rows.Count & " rows, " & ws.UsedRange.Columns.Count & " cols)"
        End If
        summary = summary & vbCrLf
    Next
    summary = summary & vbCrLf
    
    ' Table summary
    summary = summary & "## Excel Tables" & vbCrLf
    Dim tableCount As Long: tableCount = 0
    For Each ws In wb.Worksheets
        Dim tbl As ListObject
        For Each tbl In ws.ListObjects
            tableCount = tableCount + 1
            summary = summary & "- **" & tbl.Name & "** (" & ws.Name & ")"
            summary = summary & " - " & tbl.ListRows.Count & " rows, " & tbl.ListColumns.Count & " columns" & vbCrLf
        Next
    Next
    If tableCount = 0 Then summary = summary & "- No Excel tables found" & vbCrLf
    summary = summary & vbCrLf
    
    ' Named ranges summary
    summary = summary & "## Named Ranges" & vbCrLf
    If wb.Names.Count > 0 Then
        Dim nm As Name
        For Each nm In wb.Names
            On Error Resume Next
            summary = summary & "- **" & nm.Name & "**: " & nm.RefersTo & vbCrLf
            On Error GoTo 0
        Next
    Else
        summary = summary & "- No named ranges found" & vbCrLf
    End If
    summary = summary & vbCrLf
    
    summary = summary & "## Files Included" & vbCrLf
    summary = summary & "- `workbook.xml` - Overall workbook structure" & vbCrLf
    summary = summary & "- `tables/*.xml` - Excel table definitions" & vbCrLf
    summary = summary & "- `worksheets/*.xml` - Worksheet schemas (first 200 lines)" & vbCrLf
    
    Dim ts As Object: Set ts = fso.CreateTextFile(summaryPath, True)
    ts.Write summary
    ts.Close
    exported(AddSlash(summaryPath)) = True
End Sub

Private Sub DoImportProject()
    Dim wb As Workbook: Set wb = TargetWB()
    If wb Is Nothing Then Exit Sub

    Dim rootPath As String: rootPath = GetRootPath(wb)
    If rootPath = "" Then Exit Sub
    If Dir(rootPath, vbDirectory) = "" Then
        MsgBox "Nothing to import - folder '" & rootPath & "' not found.", vbExclamation
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
    
    ' Note: Excel structure import is not implemented as it would require
    ' complex workbook reconstruction. The extracted XML files are for
    ' version control, collaboration, and AI assistance purposes only.
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

'Delete any .bas/.cls/.frm/.frx/.xml file that wasn't exported this run
Private Sub PruneStaleFiles(rootPath As String, exported As Object)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim subFolder As Variant, folderPath As String
    ' Updated to include Excel subfolder
    For Each subFolder In Array("Modules", "ClassModules", "Forms", "Objects", "Misc", "Excel", "Excel/tables", "Excel/worksheets")
        folderPath = rootPath & subFolder & "\"
        If fso.FolderExists(folderPath) Then
            Dim f As Object
            For Each f In fso.GetFolder(folderPath).Files
                Dim ext As String: ext = LCase$(fso.GetExtensionName(f.Path))
                If ext = "bas" Or ext = "cls" Or ext = "frm" Or ext = "frx" Or ext = "xml" Or ext = "md" Then
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
          "# Excel structure files" & vbCrLf & _
          "*.xml text eol=lf" & vbCrLf & _
          "*.md text eol=lf" & vbCrLf & _
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
          "# VBA Sync temporary files" & vbCrLf & _
          "*.temp.zip" & vbCrLf & _
          "temp_excel_extract/" & vbCrLf & _
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
    txt = "# Excel VBA Sync - Version Control for Excel Applications" & vbCrLf & vbCrLf & _
          "** Bring modern software development practices to Excel VBA! **" & vbCrLf & vbCrLf & _
          "This tool automatically exports your Excel VBA project AND Excel file structure to a clean folder hierarchy, " & _
          "enabling Git version control, AI assistance, team collaboration, and professional development workflows " & _
          "for Excel-based applications." & vbCrLf & vbCrLf & _
          "## KEY BENEFITS" & vbCrLf & _
          "- **AI Collaboration**: AI can understand both your VBA code AND Excel data models" & vbCrLf & _
          "- **Version Control**: Full Git history of code changes, table schemas, and worksheet structure" & vbCrLf & _
          "- **Team Development**: Review changes, manage pull requests, and collaborate like software teams" & vbCrLf & _
          "- **Code Intelligence**: Syntax highlighting, linting, and IDE features in VS Code" & vbCrLf
    
    txt = txt & "- **Data Model Tracking**: Version control Excel table definitions and worksheet schemas" & vbCrLf & _
          "- **Backup & Recovery**: Never lose VBA code changes again" & vbCrLf & vbCrLf & _
          "## QUICK START" & vbCrLf & _
          "1. **Install**: Copy `VBA Sync.xlam` to your Excel add-ins folder and enable it" & vbCrLf & _
          "2. **Open**: Open your Excel workbook locally (avoid SharePoint direct links)" & vbCrLf & _
          "3. **Export**: Click **VBA Sync > Export** to create the `src/` folder structure" & vbCrLf & _
          "4. **Develop**: Edit code in VS Code, use Git for version control, get AI assistance" & vbCrLf & _
          "5. **Import**: Click **VBA Sync > Import** to load changes back into Excel" & vbCrLf & vbCrLf & _
          "## WHAT GETS EXPORTED" & vbCrLf & "```" & vbCrLf
    
    txt = txt & "src/" & vbCrLf & _
          "|-- Modules/              # Standard VBA modules (.bas)" & vbCrLf & _
          "|-- ClassModules/         # VBA class modules (.cls)" & vbCrLf & _
          "|-- Forms/                # UserForms (.frm + .frx)" & vbCrLf & _
          "|-- Objects/              # ThisWorkbook & Sheet modules (.cls)" & vbCrLf & _
          "`-- Excel/                # Excel file structure (NEW!)" & vbCrLf & _
          "    |-- workbook.xml      # Workbook structure & named ranges" & vbCrLf & _
          "    |-- tables/           # Excel table definitions (*.xml)" & vbCrLf & _
          "    |-- worksheets/       # Worksheet schemas (*.xml)" & vbCrLf & _
          "    `-- STRUCTURE_SUMMARY.md  # Human-readable data model summary" & vbCrLf
    
    txt = txt & "```" & vbCrLf & vbCrLf & _
          "Plus auto-generated Git configuration files:" & vbCrLf & _
          "- `.gitattributes` - Proper line endings for VBA files" & vbCrLf & _
          "- `.gitignore` - Excludes Excel temp files and system cruft" & vbCrLf & _
          "- `README.md` - This file (created once, never overwritten)" & vbCrLf & vbCrLf & _
          "## REAL-WORLD USE CASES" & vbCrLf & _
          "- **Financial Models**: Version control formulas, table schemas, and VBA business logic" & vbCrLf & _
          "- **Reporting Tools**: Track changes to data processing pipelines and report generation" & vbCrLf & _
          "- **Dashboard Applications**: Collaborate on interactive Excel apps with professional workflows" & vbCrLf & _
          "- **Data Integration**: Manage API connections, database queries, and ETL processes" & vbCrLf
    
    txt = txt & "- **Automation Scripts**: Version control Excel automation with full change history" & vbCrLf & vbCrLf & _
          "## INSTALLATION" & vbCrLf & _
          "1. Download `VBA Sync.xlam` from the repository" & vbCrLf & _
          "2. Copy to your Excel add-ins folder (usually `%APPDATA%\Microsoft\AddIns\`)" & vbCrLf & _
          "3. Open Excel > File > Options > Add-ins > Excel Add-ins > Browse" & vbCrLf & _
          "4. Select `VBA Sync.xlam` and check the box to enable it" & vbCrLf & _
          "5. Look for the **VBA Sync** ribbon tab" & vbCrLf & vbCrLf & _
          "## PRO TIPS" & vbCrLf & _
          "- **Git Integration**: Initialize a Git repository in your workbook folder for full version control" & vbCrLf & _
          "- **VS Code**: Install VBA language extensions for syntax highlighting and IntelliSense" & vbCrLf
    
    txt = txt & "- **AI Assistance**: Tools like GitHub Copilot can now understand your Excel data models" & vbCrLf & _
          "- **Team Workflow**: Use Git branches and pull requests for collaborative Excel development" & vbCrLf & _
          "- **Documentation**: The auto-generated `STRUCTURE_SUMMARY.md` helps onboard new team members" & vbCrLf & vbCrLf & _
          "## IMPORTANT NOTES" & vbCrLf & _
          "- **Local Files Only**: Must open Excel files from local/synced folders (not SharePoint URLs)" & vbCrLf & _
          "- **VBA Only Import**: Excel structure export is for versioning; import only updates VBA code" & vbCrLf & _
          "- **Smart Filtering**: Empty modules are skipped; removed components are cleaned up automatically" & vbCrLf & _
          "- **File Size**: Worksheet XML files are truncated at 200 lines to prevent huge files" & vbCrLf & _
          "- **Macro Security**: Ensure macro security settings allow the add-in to run" & vbCrLf & vbCrLf
    
    txt = txt & "## LICENSE" & vbCrLf & _
          "This project is released under the MIT License - feel free to use, modify, and distribute!" & vbCrLf & vbCrLf & _
          "## CREDITS" & vbCrLf & _
          "Created by **Arnaud Lavignolle** at **Axiom Project Services Pty Ltd**" & vbCrLf & _
          "Built to bridge the gap between Excel development and modern software engineering practices." & vbCrLf & vbCrLf & _
          "---" & vbCrLf & _
          "*Generated by VBA Sync on " & Format(Now, "yyyy-mm-dd") & " - Happy coding!*"

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

Private Function GetExt(t As Long) As String
    Select Case t
        Case vbext_ct_StdModule:     GetExt = ".bas"
        Case vbext_ct_ClassModule:   GetExt = ".cls"
        Case vbext_ct_MSForm:        GetExt = ".frm"   'paired .frx auto-exported
        Case vbext_ct_Document:      GetExt = ".cls"   'sheet / ThisWorkbook code-behind
        Case Else:                   GetExt = ".bas"
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


