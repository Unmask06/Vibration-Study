copilot_prompt.md
Objective

Create a minimal, Git-friendly VBA import/export toolchain so I can:

Edit .bas/.cls/.frm in VS Code with Copilot

Import them into an Excel .xlsm to run/debug in VBE

Export back to plain text for Git commits

Project Name

vba-import-export-starter

Target Structure

Create the repository with this exact tree:

vba-import-export-starter/
  src/
    Modules/        # .bas go here
    Classes/        # .cls go here
    Forms/          # .frm (+ .frx) go here
  tools/
    DevTools.bas         # VBA module with ImportAll / ExportAll / helpers
  workbook/
    App.xlsm             # sample Excel macro-enabled workbook (with empty VBAProject)
  .gitignore
  README.md


If binary generation of App.xlsm is not possible in Agent mode, create a placeholder text file at workbook/README-App.txt explaining how to create an empty .xlsm and import tools/DevTools.bas manually. Keep the rest of the project intact.

Files to Generate (exact content requirements)
1) .gitignore

Use this content:

# Office binaries (don’t version them unless absolutely necessary)
workbook/*.xlsm
workbook/*.xlsb

# Form binary resources (optionally use Git LFS if you need them)
src/Forms/*.frx

# VS Code noise
.vscode/

2) README.md

Write a concise README that covers:

What this is (VBA import/export workflow with VS Code + Git)

Prereqs: Excel desktop, “Trust access to the VBA project object model” enabled for dev; Git; VS Code

Setup:

Clone repo

Create a blank App.xlsm in workbook/ (or use existing)

Import tools/DevTools.bas into App.xlsm (Alt+F11 → File → Import File)

Ensure reference to Microsoft Visual Basic for Applications Extensibility (VBIDE) is enabled (VBE: Tools → References)

Turn on Trust access to the VBA project object model (Excel → Options → Trust Center → Trust Center Settings → Macro Settings)

Workflow:

Edit code in src/ (.bas/.cls/.frm)

In Excel, run ImportAll from DevTools module → test/debug in VBE

When stable, run ExportAll to refresh src/ → commit to Git

Notes:

.frm exports will also produce .frx (binary)

Skip removing ThisWorkbook/Sheets on import

Consider code signing for deployment

Troubleshooting (trust access, missing VBIDE reference, file paths)

3) tools/DevTools.bas

Create one standard module named DevTools with the following production-ready VBA (no placeholders):

Public procedures:

ExportAll(): exports all components in the host workbook’s VBProject to:

src/Modules/<Name>.bas for standard modules

src/Classes/<Name>.cls for class modules

src/Forms/<Name>.frm for forms (and .frx is emitted automatically)

ImportAll(): imports from the src/ tree into the host workbook’s VBProject

Before import, remove all existing non-document components (keep ThisWorkbook and any Sheet*)

Private helpers:

MkDirIfMissing(path As String)

RemoveAllCode(targetWB As Workbook)

ProjectRoot() As String that returns the workbook’s folder (assume DevTools is imported into workbook/App.xlsm)

SrcPath(subFolder As String) As String that concatenates to src/<subFolder>

UX:

MsgBox "Export complete." / "Import complete." on success

Basic error handler with Err.Raise details

Exact Implementation Requirements (Copilot must satisfy these):

Use VBIDE.VBComponent and vbext_ct_* type constants

Skip removing vbext_ct_Document components in RemoveAllCode

Use Dir() loops for importing *.bas / *.cls / *.frm

Ensure directories exist before export/import

No hardcoded absolute paths; everything relative to the host workbook’s folder

4) src/ seed files

Add tiny starter examples so the flow works on day one:

src/Modules/ModMain.bas with:

Option Explicit

Public Sub HelloWorld() that Debug.Print "Hello from ModMain"

src/Classes/CLogger.cls with:

Option Explicit

Private pPrefix As String

Public Property Let Prefix(ByVal v As String) / Get

Public Sub Info(ByVal msg As String) → Debug.Print with prefix

src/Forms/FrmAbout.frm:

Provide a minimal, valid text .frm scaffold so it can import; also include one Label control (Caption: “VBA Import/Export Starter”). (If generating a .frm is not supported, document in README how to create a dummy form and export once.)

If Agent mode cannot synthesize a proper .frm, skip the file and note the limitation in README.

Coding Standards & Constraints

Use Option Explicit everywhere.

Keep procedures small & named descriptively.

No external dependencies beyond VBIDE reference.

Defensive file I/O (check folders/files exist).

Don’t remove document modules on import.

No absolute paths; always derive from the workbook location.

Acceptance Criteria (Copilot must self-check)

Repo structure exactly matches the tree above (or substitutes a placeholder where .xlsm cannot be created).

DevTools.bas compiles in Excel VBE with the VBIDE reference set.

ExportAll:

Creates src/Modules, src/Classes, src/Forms if missing

Exports all non-document components to proper folders and extensions

ImportAll:

Removes all non-document components

Imports all .bas/.cls/.frm files found under src/ subfolders

README clearly explains setup, trust settings, and workflow.

.gitignore excludes .xlsm/.xlsb, .frx, .vscode/.

Stretch Goals (optional, separate commits)

Add a PowerShell script tools/dev.ps1 to open Excel and run ImportAll automatically.

Add a simple Rubberduck-friendly testing pattern example in src/Modules/.

Document Git LFS setup for .frx or .xlsm if versioning binaries is required.

Now generate all files and content.