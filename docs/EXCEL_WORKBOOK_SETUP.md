# Excel Workbook Setup Instructions

## Creating the `.xlsm` File from VBA Modules

This guide shows you how to manually create the Excel workbook and import all VBA modules.

---

## Method 1: Manual Creation (Recommended for MacOS)

### Step 1: Create New Workbook

1. Open Microsoft Excel for Mac
2. Click **File** → **New Workbook**
3. Save immediately as:
   - **Name**: `MyFinancialLife.xlsm`
   - **Format**: Excel Macro-Enabled Workbook (*.xlsm)
   - **Location**: Your preferred directory

### Step 2: Access VBA Editor

**Option A**: Keyboard Shortcut
- Press `Alt + F11` (or `Option + F11` or `Fn + Option + F11` depending on your Mac)

**Option B**: Menu
- Go to **Tools** → **Macro** → **Visual Basic Editor**

**Option C**: Developer Tab
1. Enable Developer tab: **Excel** → **Preferences** → **Ribbon & Toolbar**
2. Check "Developer" in Main Tabs
3. Click **Developer** tab → **Visual Basic**

### Step 3: Import VBA Modules

In the VBA Editor:

1. Right-click on **VBAProject (MyFinancialLife.xlsm)**
2. Select **Import File...**
3. Navigate to the `vba-modules` folder
4. Import files **in this exact order**:

```
1. modConfig.bas          (Configuration management)
2. modUtils.bas           (Utility functions)
3. modImport.bas          (Import engine)
4. modNormalize.bas       (Data normalization)
5. modTransfers.bas       (Transfer detection)
6. modInvestments.bas     (Investment detection)
7. modClassify.bas        (Classification engine)
8. modErrors.bas          (Error handling)
9. modUI.bas              (User interface)
```

**Important**: Import order matters because modules have dependencies.

### Step 4: Verify Installation

Your Project Explorer should show:

```
VBAProject (MyFinancialLife.xlsm)
│
├── Microsoft Excel Objects
│   ├── Sheet1 (Sheet1)
│   ├── Sheet2 (Sheet2)
│   ├── Sheet3 (Sheet3)
│   └── ThisWorkbook
│
└── Modules
    ├── modConfig
    ├── modUtils
    ├── modImport
    ├── modNormalize
    ├── modTransfers
    ├── modInvestments
    ├── modClassify
    ├── modErrors
    └── modUI
```

### Step 5: Save and Close VBA Editor

1. Press `Cmd + S` to save
2. Close VBA Editor (returns to Excel)

### Step 6: Enable Macros

1. You may see security warning: **"Macros have been disabled"**
2. Click **Enable Content** button
3. If not prompted, set permanently:
   - **Excel** → **Preferences** → **Security & Privacy**
   - Select **"Enable all macros"** (or "Warn before enabling")

### Step 7: Initialize System

1. Press `Alt + F8` (or `Fn + Option + F8`) to open Macro dialog
2. Select **InitializeSystem** from the list
3. Click **Run**

This creates all necessary sheets:
- Dashboard
- Config
- Categories
- Transactions
- Errors
- Audit
- RawData (hidden)

---

## Method 2: Quick Import via VBA Script

If you want to automate the import process:

### Create Import Script

1. Open VBA Editor
2. Insert a new module: **Insert** → **Module**
3. Paste this code:

```vba
Sub ImportAllModules()
    Dim modulePath As String
    Dim modules() As String
    Dim i As Integer

    ' SET THIS PATH TO YOUR VBA-MODULES FOLDER
    modulePath = "/Users/yourusername/MyLife/vba-modules/"

    ' Module files to import
    modules = Split("modConfig.bas,modUtils.bas,modImport.bas," & _
                   "modNormalize.bas,modTransfers.bas,modInvestments.bas," & _
                   "modClassify.bas,modErrors.bas,modUI.bas", ",")

    ' Import each module
    For i = LBound(modules) To UBound(modules)
        On Error Resume Next
        ThisWorkbook.VBProject.VBComponents.Import modulePath & modules(i)

        If Err.Number <> 0 Then
            MsgBox "Error importing: " & modules(i) & vbCrLf & Err.Description
        Else
            Debug.Print "Imported: " & modules(i)
        End If

        Err.Clear
    Next i

    MsgBox "Import complete! Remove this temporary module.", vbInformation
End Sub
```

4. **Update the `modulePath`** to your actual path
5. Run the script: Press `F5`
6. After successful import, delete this temporary module

---

## Method 3: Using Template File (If Provided)

If you received a pre-built `.xlsm` file:

1. Open the file in Excel for Mac
2. Click **Enable Macros** when prompted
3. Go to Dashboard sheet
4. Click **Initialize System**
5. Done!

---

## Troubleshooting Import Issues

### Problem: "Cannot import file"

**Cause**: VBA security settings too restrictive

**Solution**:
1. **Excel** → **Preferences** → **Security & Privacy**
2. Check **"Trust access to the VBA project object model"**
3. Try import again

### Problem: "Import fails silently"

**Cause**: File path incorrect or files not accessible

**Solution**:
1. Verify all `.bas` files exist in `vba-modules` folder
2. Check file permissions (should be readable)
3. Use absolute paths, not relative

### Problem: "Module already exists"

**Cause**: Trying to import module twice

**Solution**:
1. In VBA Editor, right-click module name
2. Select **Remove [ModuleName]**
3. Choose **No** when asked to export
4. Import the module again

### Problem: "Compile error" after import

**Cause**: Modules imported out of order, creating dependency issues

**Solution**:
1. Remove all imported modules
2. Re-import in the correct order (see Step 3 above)
3. In VBA Editor: **Debug** → **Compile VBAProject**
4. Fix any errors shown

---

## Verifying Successful Installation

### Test 1: Run InitializeSystem

1. `Alt + F8` → Select `InitializeSystem` → Run
2. Should create all sheets without errors
3. Check for success message

### Test 2: Check Module Code

1. Open VBA Editor
2. Double-click `modConfig`
3. Should see code starting with `Attribute VB_Name = "modConfig"`
4. No compile errors

### Test 3: Run Quick Test

In VBA Immediate Window (`Cmd + G`):

```vba
? modConfig.GetConfig("FuzzyThreshold", "80")
```

Should return: `80`

### Test 4: Test Button

1. Go to Dashboard sheet
2. Click **Show Statistics** button
3. Should display statistics dialog

---

## Post-Installation Configuration

After successful installation:

### 1. Configure File Paths

Go to **Config** sheet and set:
```
ImportPath1: /Users/yourusername/Documents/bank_statement.csv
ImportPath2: /Users/yourusername/Documents/credit_card.csv
ImportPath3: (optional)
```

### 2. Review Categories

Go to **Categories** sheet:
- Default categories are loaded
- Customize as needed
- Add your specific keywords

### 3. Test Import

1. Place sample files in configured paths
2. Dashboard → **Import Transactions**
3. Review results

---

## Directory Structure After Setup

Your final structure should be:

```
MyFinancialLife.xlsm         (The Excel file)
├── Dashboard               (Sheet - visible)
├── Config                  (Sheet - visible)
├── Categories              (Sheet - visible)
├── Transactions            (Sheet - visible)
├── Errors                  (Sheet - visible)
├── Audit                   (Sheet - visible)
└── RawData                 (Sheet - hidden)

VBA Modules (inside .xlsm):
├── modConfig
├── modUtils
├── modImport
├── modNormalize
├── modTransfers
├── modInvestments
├── modClassify
├── modErrors
└── modUI
```

---

## Exporting VBA Modules (Backup)

To export your VBA code for backup:

1. Open VBA Editor
2. For each module:
   - Right-click module name
   - Select **Export File...**
   - Save to backup location

Or use this macro:

```vba
Sub ExportAllModules()
    Dim comp As VBComponent
    Dim exportPath As String

    exportPath = "/Users/yourusername/MyLife/vba-backup/"

    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = vbext_ct_StdModule Then
            comp.Export exportPath & comp.Name & ".bas"
            Debug.Print "Exported: " & comp.Name
        End If
    Next comp

    MsgBox "Export complete!", vbInformation
End Sub
```

---

## Security Settings

### Recommended Settings

**Excel** → **Preferences** → **Security & Privacy**:

- ✅ **Macro Settings**: "Warn before enabling all macros"
- ✅ **Trust access to the VBA project object model**: Enabled
- ✅ **Trusted Locations**: Add your workbook folder (optional)

### For Maximum Security

- Only enable macros for files you trust
- Keep workbook in secure location
- Use FileVault encryption
- Password-protect the Excel file

---

## Next Steps

1. ✅ VBA modules imported
2. ✅ System initialized
3. ✅ Sheets created
4. → Configure file paths (Config sheet)
5. → Customize categories (Categories sheet)
6. → Import your first data file
7. → Review and classify transactions

**Ready to start!** Go to Dashboard and click "Import Transactions"

---

## Additional Resources

- **Installation Guide**: `INSTALLATION_GUIDE.md`
- **User Guide**: `USER_GUIDE.md`
- **Technical Spec**: `TECHNICAL_SPECIFICATION.md`

---

## Support Checklist

If you encounter issues:

- [ ] All 9 VBA modules imported?
- [ ] Saved as `.xlsm` format?
- [ ] Macros enabled?
- [ ] InitializeSystem run successfully?
- [ ] All sheets created?
- [ ] No compile errors in VBA?

If all checked and still issues, review the troubleshooting section above.

---

**Version**: 1.0
**Last Updated**: January 2026
