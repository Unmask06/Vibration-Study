# VBA Import/Export Starter

A minimal, Git-friendly VBA import/export toolchain that enables seamless development workflow between VS Code and Excel VBE.

## What This Is

This project provides a complete workflow for VBA development that combines the best of both worlds:
- **VS Code**: Modern editing experience with GitHub Copilot, Git integration, and powerful extensions
- **Excel VBE**: Native debugging, testing, and runtime environment

The toolchain allows you to:
- Edit `.bas/.cls/.frm` files in VS Code with full Copilot assistance
- Import them into an Excel `.xlsm` workbook for testing and debugging
- Export changes back to plain text files for Git version control
- Maintain a clean, organized project structure

## Prerequisites

- **Excel Desktop**: Required for VBA development and debugging
- **Git**: For version control of your VBA source code
- **VS Code**: Recommended for modern editing experience
- **Trust Settings**: Enable "Trust access to the VBA project object model" in Excel

## Project Structure

```
vba-import-export-starter/
├── src/
│   ├── Modules/        # Standard modules (.bas files)
│   ├── Classes/        # Class modules (.cls files)
│   └── Forms/          # UserForms (.frm + .frx files)
├── tools/
│   └── DevTools.bas    # Import/Export utility module
├── workbook/
│   └── App.xlsm        # Excel macro-enabled workbook
├── logs/               # Debug and log files
├── .gitignore          # Git ignore patterns
└── README.md           # This file
```

## Setup Instructions

### 1. Clone and Initialize

```bash
git clone <your-repo-url>
cd vba-import-export-starter
```

### 2. Create Excel Workbook

1. Open Excel
2. Create a new blank workbook
3. Save as `workbook/App.xlsm` (Excel Macro-Enabled Workbook format)

### 3. Configure Excel Trust Settings

**Critical**: These settings must be enabled for the import/export functionality to work.

1. Open Excel → **File** → **Options**
2. Navigate to **Trust Center** → **Trust Center Settings**
3. Go to **Macro Settings**
4. Check **"Trust access to the VBA project object model"**
5. Click **OK** to save

### 4. Import DevTools Module

1. Open your `App.xlsm` workbook
2. Press **Alt+F11** to open the VBA Editor
3. Go to **File** → **Import File**
4. Select `tools/DevTools.bas`
5. The DevTools module should now appear in your VBA project

### 5. Enable VBIDE Reference

1. In the VBA Editor, go to **Tools** → **References**
2. Check **"Microsoft Visual Basic for Applications Extensibility 5.3"**
3. Click **OK**

## Development Workflow

### Daily Workflow

1. **Edit in VS Code**: Modify your VBA code in the `src/` directory
   - `src/Modules/` for standard modules (`.bas`)
   - `src/Classes/` for class modules (`.cls`)
   - `src/Forms/` for UserForms (`.frm`)

2. **Import to Excel**: 
   - Open your `App.xlsm` workbook
   - Press **Alt+F11** to open VBA Editor
   - Run `DevTools.ImportAll` (or press F5 while in the DevTools module)
   - All your changes are now loaded into Excel

3. **Test and Debug**: Use Excel's native debugging features
   - Set breakpoints, step through code
   - Use the Immediate Window for testing
   - Run your procedures to validate functionality

4. **Export Changes**:
   - After making changes in Excel VBE, run `DevTools.ExportAll`
   - Your changes are written back to the `src/` directory
   - Commit your changes to Git

### DevTools Commands

- **`DevTools.ImportAll`**: Imports all VBA files from `src/` into Excel
- **`DevTools.ExportAll`**: Exports all VBA components from Excel to `src/`
- **`DevTools.ListAllComponents`**: Debug helper to list all VBA components
- **`DevTools.ShowPaths`**: Debug helper to show file paths used by DevTools

### Example Usage

```vba
' In Excel VBA Editor, run these commands:
Sub TestWorkflow()
    ' Import latest code from VS Code
    DevTools.ImportAll
    
    ' Run your code
    ModMain.HelloWorld
    
    ' Test the logger class
    Dim logger As CLogger
    Set logger = New CLogger
    logger.Prefix = "Test"
    logger.Info "Testing the import/export workflow"
    
    ' Export any changes back to files
    DevTools.ExportAll
End Sub
```

## Files and Components

### Core Tools

- **`tools/DevTools.bas`**: The main import/export utility with comprehensive error handling
- **`workbook/App.xlsm`**: Your working Excel file (not version controlled)

### Example Components

- **`src/Modules/ModMain.bas`**: Example standard module with helper functions
- **`src/Classes/CLogger.cls`**: Example class module demonstrating logging functionality
- **`src/Forms/FrmAbout.frm`**: Example UserForm with basic UI elements

### Your Project Files

Your existing VBA modules have been organized in `src/Modules/`:
- `CalculationEngine.bas`
- `DataStructures.bas`
- `ParameterManager.bas`
- `T28_TableDriven.bas`
- `T28_TableDriven_New.bas`
- `T28_UI_Calc.bas`
- `TableEventHandler.bas`
- `TableValidationManager.bas`
- `ValveListWorksheet_EventHandler.bas`

## Git Integration

### What Gets Versioned

✅ **Included in Git**:
- All source code in `src/` directory
- `tools/DevTools.bas`
- Documentation and configuration files

❌ **Excluded from Git** (see `.gitignore`):
- Excel binary files (`*.xlsm`, `*.xlsb`)
- Form binary resources (`*.frx`)
- VS Code settings (`.vscode/`)
- Temporary and backup files

### Recommended Git Workflow

```bash
# After making changes in VS Code
git add src/
git commit -m "feat: add new calculation module"

# After testing and making changes in Excel
# (Run DevTools.ExportAll first)
git add src/
git commit -m "fix: correct validation logic in TableEventHandler"

# Push to remote repository
git push origin main
```

## Troubleshooting

### Common Issues

**Import/Export Not Working**
- ✅ Verify "Trust access to the VBA project object model" is enabled
- ✅ Check that VBIDE reference is enabled in VBA Editor
- ✅ Ensure you're running from the correct workbook location
- ✅ Run `DevTools.ShowPaths` to verify directory paths

**File Not Found Errors**
- ✅ Check that `src/` directories exist
- ✅ Verify workbook is saved in the correct location
- ✅ Run `DevTools.ShowPaths` to debug path issues

**Permission Errors**
- ✅ Ensure Excel is not running as administrator
- ✅ Check file and folder permissions
- ✅ Verify antivirus isn't blocking VBA operations

**Code Changes Not Appearing**
- ✅ Save files in VS Code before importing
- ✅ Run `DevTools.ImportAll` after editing in VS Code
- ✅ Run `DevTools.ExportAll` after editing in Excel VBE

### Debug Commands

```vba
' Check current components
DevTools.ListAllComponents

' Verify paths
DevTools.ShowPaths

' Manual component removal (if needed)
DevTools.RemoveAllCode ThisWorkbook
```

## Advanced Usage

### Code Signing for Deployment

For production deployment, consider signing your VBA code:

1. Obtain a code signing certificate
2. In VBA Editor: **Tools** → **Digital Signature**
3. Select your certificate and sign the project

### Working with Forms

- `.frm` files contain the form definition and code
- `.frx` files contain binary form resources (automatically generated)
- Edit form layouts in Excel VBE, then export to update `.frm` files
- Form resources (`.frx`) are excluded from Git by default

### Collaboration

- Each developer should have their own `App.xlsm` workbook
- Only source code files in `src/` are shared via Git
- Use meaningful commit messages to track changes
- Consider branching strategies for feature development

## Best Practices

- **Always run `ImportAll` before starting work in Excel**
- **Always run `ExportAll` before committing to Git**
- **Use descriptive names for modules and procedures**
- **Include `Option Explicit` in all modules**
- **Test thoroughly in Excel before exporting changes**
- **Commit frequently with meaningful messages**

## License

This project structure and tooling can be freely used and modified for your VBA development needs.

---

*Happy VBA coding with modern tools! 🚀*