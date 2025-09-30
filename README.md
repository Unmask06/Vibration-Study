# Vibration Study Calculator

A comprehensive Excel-based calculator for valve vibration analysis using VBA, featuring a modern development workflow with VS Code integration and Git version control.

## What This Is

This project provides a complete solution for valve vibration calculations based on EI (Energy Institute) guidelines. It combines modern development practices with Excel's powerful calculation capabilities:

- **VS Code**: Modern editing with GitHub Copilot, Git integration, and powerful extensions
- **Excel VBE**: Native debugging, testing, and runtime environment for VBA calculations
- **Git-Friendly**: Clean source code management with automated import/export tools

The calculator supports multiple valve closure scenarios:
- **Liquid Close**: Standard valve closure calculations
- **Gas Open Rapid**: Gas system rapid opening analysis
- **Liquid Open**: Liquid system opening calculations

## Prerequisites

- **Excel Desktop**: Required for VBA calculations and debugging
- **Git**: For version control of source code
- **VS Code**: Recommended for modern editing experience
- **Trust Settings**: Enable "Trust access to the VBA project object model" in Excel

## Project Structure

```
vibration-study/
├── src/
│   ├── Modules/           # VBA standard modules
│   │   ├── CalculationEngine.bas     # All calculation logic
│   │   ├── DataStructures.bas        # Data types and validation
│   │   ├── ValveListGenerator.bas    # Worksheet operations
│   │   └── DevTools.bas             # Development utilities
│   ├── Classes/           # Class modules (.cls files)
│   └── Forms/             # UserForms (.frm + .frx files)
├── tools/
│   └── DevTools.bas       # Import/Export utility module
├── workbook/
│   └── App.xlsm           # Excel macro-enabled workbook
│   └── README-App.txt     # Workbook-specific documentation
├── Ref Files/             # Reference documentation
├── logs/                  # Debug and log files
├── REORGANIZATION_SUMMARY.md  # Module consolidation details
├── .gitignore             # Git ignore patterns
└── README.md              # This file
```

## Setup Instructions

### 1. Clone and Initialize

```bash
git clone <your-repo-url>
cd vibration-study
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

1. **Edit in VS Code**: Modify VBA code in the `src/Modules/` directory
   - `CalculationEngine.bas` - Add new calculation methods
   - `DataStructures.bas` - Modify data types or validation rules
   - `ValveListGenerator.bas` - Update worksheet operations
   - `DevTools.bas` - Enhance development utilities

2. **Import to Excel**:
   - Open your `App.xlsm` workbook
   - Press **Alt+F11** to open VBA Editor
   - Run `DevTools.ImportAll` (or press F5 while in the DevTools module)
   - All your changes are now loaded into Excel

3. **Test and Debug**: Use Excel's native debugging features
   - Set breakpoints in calculation functions
   - Use the Immediate Window for testing calculations
   - Validate results against EI guidelines

4. **Export Changes**:
   - After making changes in Excel VBE, run `DevTools.ExportAll`
   - Your changes are written back to the `src/` directory
   - Commit your changes to Git

### DevTools Commands

- **`DevTools.ImportAll`**: Imports all VBA files from `src/` into Excel
- **`DevTools.ExportAll`**: Exports all VBA components from Excel to `src/`
- **`DevTools.ListAllComponents`**: Debug helper to list all VBA components
- **`DevTools.ShowPaths`**: Debug helper to show file paths used by DevTools
- **`DevTools.TestCalculation`**: Quick test of calculation functionality
- **`DevTools.ShowProjectInfo`**: Display project information

### Example Usage

```vba
' In Excel VBA Editor, run these commands:
Sub TestVibrationCalculations()
    ' Import latest code from VS Code
    DevTools.ImportAll
    
    ' Initialize the calculation system
    ValveListGenerator.InitializeModules
    
    ' Generate valve data from ValveList table
    ValveListGenerator.Generate_Inputs_From_tbValveList
    
    ' Run calculations for all valves
    ValveListGenerator.RunCalculations
    
    ' Export any changes back to files
    DevTools.ExportAll
End Sub

' Example of direct calculation usage
Sub CalculateSingleValve()
    Dim inputs As ValveInputs
    Dim result As CalculationResult
    
    ' Populate valve inputs (example values)
    inputs.rho = 1000#        ' Fluid density
    inputs.c0 = 1200#         ' Speed of sound
    inputs.v = 2.5#           ' Velocity
    inputs.Dint_mm = 200#     ' Internal diameter
    inputs.Dext_mm = 250#     ' External diameter
    inputs.T_mm = 10#         ' Wall thickness
    inputs.casetype = "liqclose"
    
    ' Perform calculation
    result = CalculationEngine.CalculateByCase(inputs.casetype, inputs)
    
    ' Display results
    Debug.Print "Ppeak: " & result.Ppeak & " Pa"
    Debug.Print "Fmax: " & result.Fmax & " kN"
    Debug.Print "LOF: " & result.LOF
End Sub
```

## Module Architecture

### Core Modules

#### **CalculationEngine.bas**
**Purpose**: All mathematical calculations and engineering formulas
- `CalculateByCase()` - Main calculation dispatcher for different valve types
- `CalculateLiquidClose()` - Liquid valve closure calculations (Joukowsky equation)
- `CalculateGasOpenRapid()` - Gas rapid opening analysis
- `CalculateLiquidOpen()` - Liquid opening calculations
- `CalculateWaveSpeed()` - Fluid wave speed calculations
- `CalculateFlim()` - Force limit calculations (EI methodology)

#### **DataStructures.bas**
**Purpose**: Data types, validation, and parameter management
- `ValveInputs` type - Complete valve parameter structure
- `CalculationResult` type - Calculation output structure
- Parameter indexing and lookup functions
- Data validation for Excel tables
- Unit conversion utilities (Pa ↔ barg)

#### **ValveListGenerator.bas**
**Purpose**: Excel worksheet operations and data management
- `Generate_Inputs_From_tbValveList()` - Create input columns from valve table
- `RunCalculations()` - Execute calculations for all valves
- `ClearValveData()` - Reset worksheet data
- `RefreshValidations()` - Update data validation rules

#### **DevTools.bas**
**Purpose**: Development utilities and import/export functionality
- VBA import/export between VS Code and Excel
- Development testing utilities
- Project information and debugging tools

### Dependencies

```
ValveListGenerator.bas ────┐
                   │       │
                   ▼       ▼
         DataStructures.bas    CalculationEngine.bas
                   ▲
                   │
                   ▼
              DevTools.bas
```

**Initialization Order**:
1. DataStructures (provides types)
2. CalculationEngine (uses ValveInputs type)
3. ValveListGenerator (uses both)
4. DevTools (uses DataStructures)

## Calculation Methodology

### Supported Case Types

1. **Liquid Close** (`liqclose`)
   - Joukowsky pressure surge equation: ΔP = ρ × c × v
   - Force calculation: F = ΔP × Area
   - Load factor: LOF = F / Flim

2. **Gas Open Rapid** (`gasopenrapid`)
   - Placeholder for EI gas opening calculations
   - Framework ready for implementation

3. **Liquid Open** (`liqopen`)
   - Placeholder for EI liquid opening calculations
   - Framework ready for implementation

### EI Guidelines Compliance

- Force limit calculations based on EI Guidelines (T2.8)
- Pipe support type considerations (Anchor, Guide, Sliding, None)
- Material properties and dimensional validation
- Comprehensive input validation and error handling

## Git Integration

### What Gets Versioned

✅ **Included in Git**:
- All source code in `src/` directory
- `tools/DevTools.bas`
- Documentation and configuration files
- `REORGANIZATION_SUMMARY.md`

❌ **Excluded from Git** (see `.gitignore`):
- Excel binary files (`*.xlsm`, `*.xlsb`)
- Form binary resources (`*.frx`)
- VS Code settings (`.vscode/`)
- Temporary and backup files
- `logs/` directory contents

### Recommended Git Workflow

```bash
# After implementing new calculation method
git add src/Modules/CalculationEngine.bas
git commit -m "feat: add gas opening rapid calculation per EI guidelines"

# After updating data validation
git add src/Modules/DataStructures.bas
git commit -m "fix: improve valve input validation for negative values"

# After testing in Excel and exporting changes
git add src/
git commit -m "test: validate calculations against EI reference cases"

# Push to remote repository
git push origin main
```

## Troubleshooting

### Common Issues

**Calculation Errors**
- ✅ Verify all required valve parameters are provided
- ✅ Check units (pressure in barg, dimensions in mm)
- ✅ Validate fluid properties are physically reasonable
- ✅ Ensure case type matches supported values

**Import/Export Not Working**
- ✅ Verify "Trust access to the VBA project object model" is enabled
- ✅ Check that VBIDE reference is enabled in VBA Editor
- ✅ Ensure you're running from the correct workbook location
- ✅ Run `DevTools.ShowPaths` to verify directory paths

**File Not Found Errors**
- ✅ Check that `src/Modules/` directory exists
- ✅ Verify workbook is saved in the correct location
- ✅ Run `DevTools.ShowPaths` to debug path issues

**Validation Errors**
- ✅ Run `ValveListGenerator.RefreshValidations()` to update dropdowns
- ✅ Check that Data sheet contains validation lists
- ✅ Ensure table structure matches expected format

### Debug Commands

```vba
' Check current components
DevTools.ListAllComponents

' Verify paths
DevTools.ShowPaths

' Test calculation engine
DevTools.TestCalculation

' Reset module variables
DevTools.ResetModuleVariables

' Quick validation setup
DevTools.QuickSetupValidations
```

## Advanced Usage

### Adding New Calculation Cases

1. Add new case type to `CalculationEngine.CalculateByCase()`
2. Implement calculation method following EI guidelines
3. Update validation rules in `DataStructures.bas`
4. Test thoroughly against reference cases

### Custom Validation Rules

Modify `DataStructures.bas` to add custom validation:
```vba
Public Function ValidateValveInputs(inputs As ValveInputs) As String
    ' Add your custom validation logic here
End Function
```

### Working with Large Datasets

For performance with many valves:
- Use `ValveListGenerator.ClearValveData()` to reset
- Process valves in batches if needed
- Monitor memory usage in large workbooks

## Best Practices

### Development
- **Always run `ImportAll` before starting work in Excel**
- **Always run `ExportAll` before committing to Git**
- **Test calculations against known reference cases**
- **Include `Option Explicit` in all modules**
- **Use meaningful variable names following EI conventions**

### Calculations
- **Validate all inputs before calculations**
- **Check units and dimensional consistency**
- **Handle edge cases (zero values, negative pressures)**
- **Document calculation assumptions and limitations**

### Version Control
- **Commit frequently with descriptive messages**
- **Use feature branches for new calculation methods**
- **Tag releases when calculation methods are validated**
- **Document changes in commit messages**

## Reference Documentation

- **EI Guidelines T2.8**: Available in `Ref Files/` directory
- **REORGANIZATION_SUMMARY.md**: Details of module consolidation
- **workbook/README-App.txt**: Workbook-specific usage instructions

## License

This vibration study calculator is designed for engineering analysis and can be freely used and modified for valve vibration calculations.

---

*Engineering calculations with modern development practices! 🔧📊*

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

Your consolidated VBA modules have been organized in `src/Modules/`:
- `CalculationEngine.bas` - All calculation logic and engineering formulas
- `DataStructures.bas` - Data types, validation, and parameter management
- `ValveListGenerator.bas` - Worksheet operations and data management
- `DevTools.bas` - Development utilities and import/export tools

### Reference Materials

- **`Ref Files/`**: EI Guidelines and reference documentation
- **`REORGANIZATION_SUMMARY.md`**: Complete details of module consolidation
- **`workbook/README-App.txt`**: Workbook-specific usage instructions

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

**Calculation Errors**
- ✅ Verify all required valve parameters are provided
- ✅ Check units (pressure in barg, dimensions in mm)
- ✅ Validate fluid properties are physically reasonable
- ✅ Ensure case type matches supported values

**Import/Export Not Working**
- ✅ Verify "Trust access to the VBA project object model" is enabled
- ✅ Check that VBIDE reference is enabled in VBA Editor
- ✅ Ensure you're running from the correct workbook location
- ✅ Run `DevTools.ShowPaths` to verify directory paths

**File Not Found Errors**
- ✅ Check that `src/Modules/` directory exists
- ✅ Verify workbook is saved in the correct location
- ✅ Run `DevTools.ShowPaths` to debug path issues

**Validation Errors**
- ✅ Run `ValveListGenerator.RefreshValidations()` to update dropdowns
- ✅ Check that Data sheet contains validation lists
- ✅ Ensure table structure matches expected format

### Debug Commands

```vba
' Check current components
DevTools.ListAllComponents

' Verify paths
DevTools.ShowPaths

' Test calculation engine
DevTools.TestCalculation

' Reset module variables
DevTools.ResetModuleVariables

' Quick validation setup
DevTools.QuickSetupValidations

' Show project information
DevTools.ShowProjectInfo
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

### Development
- **Always run `ImportAll` before starting work in Excel**
- **Always run `ExportAll` before committing to Git**
- **Test calculations against known reference cases**
- **Include `Option Explicit` in all modules**
- **Use meaningful variable names following EI conventions**

### Calculations
- **Validate all inputs before calculations**
- **Check units and dimensional consistency**
- **Handle edge cases (zero values, negative pressures)**
- **Document calculation assumptions and limitations**

### Version Control
- **Commit frequently with descriptive messages**
- **Use feature branches for new calculation methods**
- **Tag releases when calculation methods are validated**
- **Document changes in commit messages**

## Reference Documentation

- **EI Guidelines T2.8**: Available in `Ref Files/` directory
- **REORGANIZATION_SUMMARY.md**: Details of module consolidation
- **workbook/README-App.txt**: Workbook-specific usage instructions

## License

This vibration study calculator is designed for engineering analysis and can be freely used and modified for valve vibration calculations.

---

*Engineering calculations with modern development practices! 🔧�*