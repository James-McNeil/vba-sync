# Excel VBA Git Sync

This repo mirrors the VBA project of the workbook in the `src/` folder, so you can version, diff, and review code comfortably in Git/VS Code.

## Workflow
1. Open the workbook locally (not directly over SharePoint/Teams https URLs).
2. Click **VBA Sync ? Export** to write all modules to `src/` (sub-folders mirror the VBE tree).
3. Work in VS Code / your Git client (lint, AI assist, diffs, PRs).
4. Click **VBA Sync ? Import** to push changes back into the workbook.

## Structure
``"
src/
  Modules/        ' .bas
  ClassModules/   ' .cls
  Forms/          ' .frm + .frx
  Objects/        ' ThisWorkbook / Sheets as .cls (code only)
``"

## Notes
- Empty document modules (only `Option Explicit`) aren’t exported.
- Export removes files that no longer exist in the project.
- `.gitattributes` and `.gitignore` are auto-generated (by default at the repo root).
- SharePoint/Teams URLs are blocked: open the synced local copy instead.
