# Excel Workbook Setup Instructions

## Creating App.xlsm

Since Excel workbook files cannot be automatically generated in this environment, you need to create the Excel workbook manually. Follow these steps:

### Step 1: Create the Workbook

1. Open Microsoft Excel
2. Create a new blank workbook
3. Save the file as `App.xlsm` in this directory (`workbook/`)
4. Choose "Excel Macro-Enabled Workbook (*.xlsm)" as the file type

### Step 2: Import DevTools Module

1. With `App.xlsm` open, press **Alt+F11** to open the VBA Editor
2. In the VBA Editor, go to **File** → **Import File**
3. Navigate to the `tools/` directory in your project
4. Select `DevTools.bas` and click **Open**
5. The DevTools module should now appear in your VBA Project Explorer

### Step 3: Enable Required References

1. In the VBA Editor, go to **Tools** → **References**
2. Scroll down and check **"Microsoft Visual Basic for Applications Extensibility 5.3"**
3. Click **OK**

### Step 4: Configure Trust Settings

**IMPORTANT**: This step is required for the import/export functionality to work.

1. Go to **File** → **Options** in Excel
2. Navigate to **Trust Center** → **Trust Center Settings**
3. Go to **Macro Settings**
4. Check **"Trust access to the VBA project object model"**
5. Click **OK** to save the changes

### Step 5: Test the Setup

1. In the VBA Editor, open the DevTools module
2. Place your cursor in the `ImportAll` subroutine
3. Press **F5** to run the procedure
4. You should see a message box saying "Import complete" with the number of components imported

### Verification

Your VBA Project should now contain:
- **DevTools** (imported from tools/)
- **ModMain** (imported from src/Modules/)
- **CLogger** (imported from src/Classes/)
- **FrmAbout** (imported from src/Forms/)
- **Plus your existing project modules** (imported from src/Modules/)

### Troubleshooting

If you encounter issues:

1. **"Permission denied" or "Programmatic access to VBA not trusted"**
   - Ensure you've enabled "Trust access to the VBA project object model" in Excel Options

2. **"Reference not found"**
   - Make sure the VBIDE reference is enabled in Tools → References

3. **"Path not found"**
   - Verify that the workbook is saved in the correct location (workbook/ directory)
   - Run `DevTools.ShowPaths` to see the paths being used

4. **No components imported**
   - Check that the src/ directories contain the expected .bas, .cls, and .frm files
   - Verify the directory structure matches the expected layout

### Next Steps

Once your workbook is set up:

1. You can edit VBA code in VS Code (in the `src/` directory)
2. Use `DevTools.ImportAll` to bring changes into Excel for testing
3. Use `DevTools.ExportAll` to save changes back to the src files
4. Commit your changes to Git (the .xlsm file is excluded by .gitignore)

See the main README.md for detailed workflow instructions.

---

**Note**: The App.xlsm file should not be committed to Git as it's excluded in .gitignore. Each developer should create their own workbook following these instructions.