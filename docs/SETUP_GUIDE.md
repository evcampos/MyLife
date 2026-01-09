# Setup Guide - [MY LIFE] Financial Management System

## Prerequisites

### System Requirements
- **Excel for MacOS** (tested on versions 2019 and later)
- **MacOS** 10.14 or later
- **Macros enabled** in Excel
- Minimum 2GB free disk space for data files

### Optional Requirements
- Access to bank/card statement exports (CSV or Excel format)
- Investment transaction reports
- Index data files (CDI, SELIC, IPCA, USD/BRL, Fed Funds)

---

## Installation

### Step 1: Download the System

1. Clone or download this repository to your Mac
2. Extract to a location like `/Users/YourName/Documents/MyLife`
3. Ensure all files are present:
   - `vba-modules/` folder with 14 `.bas` files
   - `docs/` folder with documentation
   - `samples/` folder with sample data
   - `README.md` and `PROJECT_STATUS.md`

### Step 2: Create the Excel Workbook

1. Open **Excel for Mac**
2. Create a new blank workbook
3. Save it as `MyLife.xlsm` (Excel Macro-Enabled Workbook)
4. Save it in the same directory as the repository

---

## Initial Configuration

### Step 3: Enable Macros

1. In Excel, go to **Excel > Preferences**
2. Click on **Security & Privacy**
3. Under **Macro Settings**, select:
   - ☑️ **Enable all macros**
   - ☑️ **Trust access to the VBA project object model**
4. Click **OK**

### Step 4: Import VBA Modules

You have two options:

#### Option A: Import Modules Manually (Recommended)

1. In Excel, press `Alt + F11` (or `Option + F11`) to open VBA Editor
2. In the VBA Editor, go to **File > Import File**
3. Navigate to the `vba-modules/` folder
4. Import each `.bas` file one by one:
   - `modConfig.bas`
   - `modUtils.bas`
   - `modMain.bas`
   - `modImportBanks.bas`
   - `modImportCards.bas`
   - `modImportInvestments.bas`
   - `modImportOPUS.bas`
   - `modImportDebts.bas`
   - `modClassification.bas`
   - `modCorrelation.bas`
   - `modIndexes.bas`
   - `modCapitalCost.bas`
   - `modDashboard.bas`
   - `modHealthCheck.bas`

5. After importing all modules, close the VBA Editor

#### Option B: Copy/Paste Code

1. Open VBA Editor (`Alt + F11`)
2. For each `.bas` file:
   - Insert a new module (**Insert > Module**)
   - Open the `.bas` file in a text editor
   - Copy all contents
   - Paste into the new module
3. Rename each module to match the filename (without `.bas`)

### Step 5: Initialize the System

1. In VBA Editor, press `F5` or click **Run**
2. Select `modConfig.InitializeSystem` from the macro list
3. Click **Run**
4. Wait for the system to create all required sheets
5. Click **OK** when the success message appears

**Result:** You should now see 11 worksheets created:
- FILES PATHS
- FILES STRUCTURE
- BANKS
- CARDS
- INVESTMENTS
- OPUS
- DEBTS
- INDEXES
- CATEGORIES
- DASHBOARD
- HEALTH_CHECK

---

## Configuration

### Step 6: Configure File Paths

1. Go to the **FILES PATHS** sheet
2. For each data source you have, enter the full path to your file:

| Source | FilePath | Active |
|--------|----------|--------|
| ITAU_BANK | /Users/YourName/Documents/Data/itau_extrato.csv | TRUE |
| NUBANK_BANK | /Users/YourName/Documents/Data/nubank_extrato.xlsx | TRUE |
| C6_BANK | | FALSE |
| BB_BANK | | FALSE |
| ITAU_CARD | /Users/YourName/Documents/Data/itau_fatura.csv | TRUE |
| NUBANK_CARD | | FALSE |
| C6_CARD | | FALSE |
| BB_CARD | | FALSE |
| ITAU_INV | /Users/YourName/Documents/Data/itau_investimentos.csv | TRUE |
| NUBANK_INV | | FALSE |
| C6_INV | | FALSE |
| BB_INV | | FALSE |
| XP_INV | | FALSE |
| BTG_INV | | FALSE |
| OPUS | /Users/YourName/Documents/Data/opus_positions.xlsx | TRUE |
| DEBTS | /Users/YourName/Documents/Data/debts.csv | TRUE |
| INDEXES | /Users/YourName/Documents/Data/indexes.csv | TRUE |

**Tips:**
- Use **full paths** starting with `/Users/`
- Set **Active = TRUE** only for sources you have data for
- Leave **FilePath empty** or set **Active = FALSE** for unused sources
- MacOS paths are case-sensitive

### Step 7: Configure Categories

1. Go to the **CATEGORIES** sheet
2. Review the pre-loaded sample categories
3. Add your own categories or modify existing ones:

| Category | Subcategory | Keywords | Priority |
|----------|-------------|----------|----------|
| Food & Dining | Restaurants | RESTAURANT,RESTAURANTE,BAR,CAFE | 10 |
| Food & Dining | Groceries | SUPERMARKET,MERCADO,GROCERY | 10 |
| Transportation | Gas | POSTO,GAS,SHELL,IPIRANGA,PETROBRAS | 10 |
| Transportation | Public Transit | METRO,BUS,UBER,TAXI,99POP | 8 |
| Housing | Rent | ALUGUEL,RENT,RENTAL | 10 |
| Housing | Utilities | LUZ,AGUA,GAS,ENERGIA,WATER,ELECTRIC | 9 |

**Priority levels:**
- **10** = Highest priority (exact match wins)
- **8-9** = High priority
- **5-7** = Medium priority
- **1-4** = Low priority

### Step 8: Review File Structure Expectations

1. Go to the **FILES STRUCTURE** sheet
2. This sheet shows expected column names for each source
3. The system auto-detects columns, but you can add alternative names here

**Example:**
```
Source: ITAU_BANK
ExpectedColumns: Date,Description,Value,Balance
AlternativeNames: Data,Descricao,Valor,Saldo
```

---

## Testing the Setup

### Step 9: Test with Sample Data

1. If you don't have real data yet, use the sample files in `samples/` folder
2. Copy sample files to a test location
3. Update **FILES PATHS** sheet with sample file paths
4. Set **Active = TRUE** for sample sources

### Step 10: Run Initial Import

1. Go to the **DASHBOARD** sheet
2. In VBA Editor, run `modImportBanks.ImportAllBanks`
3. Check for any error messages
4. Verify data appears in the **BANKS** sheet

**If successful:**
- You should see transactions imported
- Dates should be properly formatted
- Amounts should be numeric

**If errors occur:**
- Check file paths are correct
- Verify file format (CSV or Excel)
- Check column names match expected structure
- Review error message for specific issue

### Step 11: Run Classification

1. In VBA Editor, run `modClassification.ClassifyAllTransactions`
2. Check the **BANKS** sheet for populated categories
3. Review unclassified transactions

### Step 12: Run Health Check

1. In VBA Editor, run `modHealthCheck.RunHealthCheck`
2. Review the **HEALTH_CHECK** sheet
3. Address any FAIL or WARNING items

---

## Common Issues & Solutions

### Issue: "File not found" error

**Solution:**
- Verify the full path is correct (copy from Finder)
- Check file permissions (must be readable)
- Ensure no typos in the path
- MacOS paths start with `/Users/` not `~/`

### Issue: "Could not detect required columns"

**Solution:**
- Open your CSV/Excel file and check column names
- Add alternative column names to **FILES STRUCTURE** sheet
- Ensure first row contains headers
- Check delimiter (comma, semicolon, or tab)

### Issue: Macros disabled or not running

**Solution:**
- Re-enable macros in **Excel > Preferences > Security & Privacy**
- Save the workbook as `.xlsm` format
- Close and reopen Excel
- Check for any security warnings and click "Enable Macros"

### Issue: "Object not found" or module errors

**Solution:**
- Verify all 14 VBA modules are imported
- Check module names match exactly (case-sensitive)
- Re-import any missing modules
- Restart Excel

### Issue: Dates not importing correctly

**Solution:**
- Check your system's date format (System Preferences > Language & Region)
- Verify source file has dates in recognizable format (YYYY-MM-DD, DD/MM/YYYY, etc.)
- Adjust `modUtils.ParseDate` function if needed

### Issue: Non-English characters not displaying

**Solution:**
- Ensure CSV files are saved with UTF-8 encoding
- Open CSV in TextEdit and re-save with UTF-8 encoding
- Excel should handle accented characters automatically

---

## Advanced Configuration

### Custom Button Setup (Optional)

You can add buttons to the DASHBOARD for easier access:

1. Go to **DASHBOARD** sheet
2. Go to **Insert > Shapes** or **Developer > Insert > Button**
3. Draw a button on the sheet
4. Right-click the button and select **Assign Macro**
5. Choose a macro like `modImportBanks.ImportAllBanks`
6. Repeat for other common operations

**Recommended buttons:**
- **Import Banks** → `modImportBanks.ImportAllBanks`
- **Import Cards** → `modImportCards.ImportAllCards`
- **Import Investments** → `modImportInvestments.ImportAllInvestments`
- **Classify** → `modClassification.ClassifyAllTransactions`
- **Correlate** → `modCorrelation.CorrelateAllInvestments`
- **Update Dashboard** → `modDashboard.UpdateDashboard`
- **Health Check** → `modHealthCheck.RunHealthCheck`

### Keyboard Shortcuts (Optional)

1. In VBA Editor, go to **Tools > Macros**
2. Select a macro
3. Click **Options**
4. Assign a keyboard shortcut (e.g., `Ctrl+Shift+I` for Import)

---

## Data Backup

### Recommended Backup Strategy

1. **Daily:** Save a copy of `MyLife.xlsm` with date suffix
   - Example: `MyLife_2026-01-09.xlsm`

2. **Weekly:** Export key sheets to CSV
   - BANKS, CARDS, INVESTMENTS, OPUS, DEBTS

3. **Monthly:** Full system backup including source files
   - Copy entire MyLife folder to external drive or cloud

### Automated Backup (Optional)

Create a VBA macro to auto-save backups:

```vba
Sub BackupWorkbook()
    Dim backupPath As String
    backupPath = ThisWorkbook.Path & "/backups/MyLife_" & Format(Date, "yyyy-mm-dd") & ".xlsm"
    ThisWorkbook.SaveCopyAs backupPath
End Sub
```

---

## Next Steps

After setup is complete:

1. ✅ Read the **USER_GUIDE.md** for detailed usage instructions
2. ✅ Review **VBA_MODULES.md** to understand the code structure
3. ✅ Import your first real data files
4. ✅ Run classification and correlation
5. ✅ Build your dashboard
6. ✅ Schedule regular imports (weekly or monthly)

---

## Getting Help

If you encounter issues:

1. Check **HEALTH_CHECK** sheet for diagnostic information
2. Review **ARCHITECTURE.md** for system design details
3. Check the VBA Immediate Window for debug logs
4. Verify all prerequisites are met
5. Try with sample data first to isolate issues

---

## System Maintenance

### Weekly Tasks
- Import new transactions
- Run classification
- Run correlation
- Update dashboard
- Run health check

### Monthly Tasks
- Review unclassified transactions
- Update categories as needed
- Import index data
- Calculate capital costs
- Backup workbook

### Quarterly Tasks
- Review and optimize category rules
- Archive old data if needed
- Update documentation
- Review system performance

---

**Setup Version:** 1.0
**Last Updated:** January 2026
**Compatible with:** Excel for MacOS 2019+
