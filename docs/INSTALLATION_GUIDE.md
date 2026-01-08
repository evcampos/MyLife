# Installation & Setup Guide

## Table of Contents
1. [System Requirements](#system-requirements)
2. [Installation Steps](#installation-steps)
3. [Initial Configuration](#initial-configuration)
4. [Sample Data](#sample-data)
5. [Troubleshooting](#troubleshooting)

---

## System Requirements

### Hardware
- Mac computer (Intel or Apple Silicon)
- 8GB RAM minimum (16GB recommended)
- 500MB free disk space

### Software
- **Microsoft Excel for MacOS**
  - Version: 2016 or later (2019, 2021, Office 365 recommended)
  - Macros must be enabled
- **MacOS**: 10.13 (High Sierra) or later

---

## Installation Steps

### Step 1: Create the Excel Workbook

1. Open Microsoft Excel for Mac
2. Create a new blank workbook
3. Save as: `MyFinancialLife.xlsm` (Excel Macro-Enabled Workbook)
   - **Important**: Must be `.xlsm` format to support VBA macros

### Step 2: Import VBA Modules

1. In Excel, press `Alt + F11` (or `Fn + Option + F11` on some Macs) to open the VBA Editor
2. In the VBA Editor, go to `File` → `Import File...`
3. Import each `.bas` file in this order:
   - `modConfig.bas`
   - `modUtils.bas`
   - `modImport.bas`
   - `modNormalize.bas`
   - `modTransfers.bas`
   - `modInvestments.bas`
   - `modClassify.bas`
   - `modErrors.bas`
   - `modUI.bas`

4. After importing all modules, your VBA Project should look like:
   ```
   VBAProject (MyFinancialLife.xlsm)
   ├── Microsoft Excel Objects
   │   └── ThisWorkbook
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

5. Close the VBA Editor and return to Excel

### Step 3: Enable Macros

1. If you see a security warning: "Macros have been disabled"
2. Click **Enable Content** or **Enable Macros**
3. You may need to:
   - Go to `Excel` → `Preferences` → `Security & Privacy`
   - Set Macro Settings to "Enable all macros"

### Step 4: Initialize the System

1. Press `Alt + F8` (or `Fn + Option + F8`) to open the Macro dialog
2. Select `InitializeSystem` from the list
3. Click **Run**

This will create all necessary sheets:
- Dashboard (control panel)
- Config (settings)
- Categories (category mappings)
- Transactions (processed data)
- Errors (unresolved issues)
- Audit (log)
- RawData (hidden)

---

## Initial Configuration

### Configure File Paths

1. Go to the **Config** sheet
2. Set your import file paths:
   - `ImportPath1`: `/Users/yourusername/Documents/BankStatement.csv`
   - `ImportPath2`: `/Users/yourusername/Documents/CreditCard.xlsx`
   - `ImportPath3`: (optional third path)

#### MacOS Path Examples:
```
/Users/yourusername/Documents/Finances/BankStatement.csv
/Users/yourusername/Downloads/CreditCard_2024.xlsx
/Users/yourusername/Library/Mobile Documents/com~apple~CloudDocs/Statements/Investment.csv
```

#### iCloud Drive Paths:
```
/Users/yourusername/Library/Mobile Documents/com~apple~CloudDocs/Documents/Statement.csv
```

### Configure Settings

In the **Config** sheet, adjust these settings as needed:

| Setting | Default | Description |
|---------|---------|-------------|
| DateFormat | dd/mm/yyyy | Date format in your files |
| DefaultCurrency | USD | Your base currency |
| FuzzyThreshold | 80 | Matching sensitivity (0-100) |
| TransferTolerance | 0.001 | Amount tolerance for transfers |
| TransferDateDays | 1 | Date tolerance for transfers (days) |
| SkipDuplicates | TRUE | Skip duplicate transactions |
| AutoBackup | TRUE | Enable automatic backups |

### Customize Categories

1. Go to the **Categories** sheet
2. Review the default category mappings
3. Add your custom categories:
   - **Category**: Primary classification (e.g., "Food & Dining")
   - **Subcategory**: Secondary detail (e.g., "Restaurants")
   - **Keywords**: Comma-separated search terms (e.g., "CHIPOTLE,PANERA,OLIVE GARDEN")
   - **Priority**: 1-200 (lower = higher priority)
   - **IsActive**: TRUE or FALSE

#### Tips for Effective Categories:
- Use specific keywords for high accuracy
- Include variations: "RESTAURANT,RESTAURANTE,RESTAU"
- Include common misspellings
- Use partial matches: "AMAZON" will match "AMAZON.COM", "AMAZON PRIME", etc.
- Higher priority rules are checked first

---

## Sample Data

### Required CSV/Excel Format

Your input files should have these columns (names can vary):

#### Minimum Required:
- **Date** (or Data, Fecha, Datum)
- **Description** (or Descripción, Histórico)
- **Amount** (or Valor, Monto)

#### Optional:
- **Debit/Credit** (or Débito/Crédito)
- **Account** (or Conta, Cuenta)
- **Balance** (or Saldo)

### Example CSV Format:

```csv
Date,Description,Amount,Account
01/15/2024,WHOLE FOODS MARKET,-125.50,****1234
01/15/2024,PAYROLL DEPOSIT,3500.00,****1234
01/16/2024,SHELL GAS STATION,-45.00,****1234
01/17/2024,NETFLIX SUBSCRIPTION,-15.99,****5678
```

### Example with Debit/Credit Columns:

```csv
Date,Description,Debit,Credit,Balance
01/15/2024,WHOLE FOODS MARKET,125.50,,5432.10
01/15/2024,PAYROLL DEPOSIT,,3500.00,8932.10
01/16/2024,SHELL GAS STATION,45.00,,8887.10
```

---

## Usage Workflow

### 1. First Time Import

1. Go to **Dashboard** sheet
2. Click **Import Transactions**
3. Wait for processing (may take 30-60 seconds)
4. Review the summary statistics

### 2. Review Results

1. Check **Transactions** sheet for classified data
2. Check **Errors** sheet for unresolved items
3. Manually categorize any unclassified transactions

### 3. Resolve Errors

For each error in the **Errors** sheet:
1. Review the description and amount
2. Determine the correct category/subcategory
3. Enter values in SuggestedCategory and SuggestedSubcategory columns
4. Go to Dashboard → Click **Clear Resolved Errors**

### 4. Improve Classification

When you manually categorize transactions:
1. The system learns from your corrections
2. New keywords are suggested automatically
3. Click **Update Categories** to reload mappings
4. Click **Reprocess Data** to apply new rules

### 5. Export Results

1. Go to Dashboard
2. Click **Export Transactions**
3. Choose save location
4. Open CSV in Excel, Google Sheets, or analysis tools

---

## Troubleshooting

### Macros Don't Run

**Problem**: "Cannot run the macro. The macro may not be available..."

**Solutions**:
1. Check that file is saved as `.xlsm` format
2. Verify all VBA modules are imported
3. Go to `Excel` → `Preferences` → `Security & Privacy`
4. Enable all macros
5. Close and reopen the file

### Import Finds No Files

**Problem**: "No data imported. Please check your file paths"

**Solutions**:
1. Verify paths in Config sheet are absolute paths
2. Check files exist at specified locations
3. Ensure you have read permissions
4. MacOS paths are case-sensitive
5. Test path in Terminal: `ls "/path/to/file.csv"`

### Import Fails to Parse File

**Problem**: "Could not detect required columns"

**Solutions**:
1. Verify file has Date, Description, and Amount (or Debit/Credit) columns
2. Check that first row contains headers
3. Ensure file is not corrupted
4. Try opening file in Excel manually first
5. Check for special characters in column names

### Classification Accuracy is Low

**Problem**: Many transactions are "Unclassified"

**Solutions**:
1. Add more keywords to Categories sheet
2. Lower FuzzyThreshold in Config (try 70 instead of 80)
3. Include common misspellings and variations
4. Use partial keywords ("REST" instead of "RESTAURANT")
5. Check keyword priority (lower numbers = higher priority)

### Performance is Slow

**Problem**: Import takes more than 2 minutes

**Solutions**:
1. Close other applications
2. Break large files into smaller chunks
3. Disable SkipDuplicates if not needed
4. Increase Excel's memory allocation
5. Process one file at a time

### Errors About Duplicate Names

**Problem**: "A sheet named 'Categories' already exists"

**Solutions**:
1. Don't run InitializeSystem multiple times
2. If sheets exist, skip initialization
3. Or delete existing sheets before re-initializing

### MacOS-Specific Issues

**Problem**: Path with spaces doesn't work

**Solution**:
- Paths with spaces are supported
- Ensure path is in Config value exactly as shown in Finder
- Example: `/Users/John Smith/Documents/File.csv`

**Problem**: iCloud Drive files not accessible

**Solution**:
- Ensure iCloud Drive is synced
- Check file is downloaded (not just in cloud)
- Use full iCloud path: `/Users/username/Library/Mobile Documents/...`

---

## Performance Optimization

### For Large Datasets (>10,000 transactions)

1. **Process in batches**:
   - Import one month at a time
   - Combine results later

2. **Optimize categories**:
   - Use higher priority for common transactions
   - Reduce number of low-priority rules

3. **Disable duplicate checking**:
   - Set `SkipDuplicates` = FALSE if you're sure there are no dupes

4. **Close unnecessary applications**:
   - Give Excel maximum RAM

---

## Backup & Recovery

### Create Manual Backup

1. Go to Dashboard
2. Click **Backup Workbook**
3. Backup saved with timestamp

### Restore from Backup

1. Open backup file
2. Save As with original name
3. All data and settings preserved

### Auto-Backup (Optional)

Set `AutoBackup` = TRUE in Config sheet to enable automatic backups before each import.

---

## Security & Privacy

### Data Storage
- All data stays local on your Mac
- No cloud synchronization (unless you use iCloud Drive)
- No external connections
- No telemetry or analytics

### Sensitive Information
- Account numbers are masked (last 4 digits only)
- Descriptions preserved for classification
- Consider encrypting the `.xlsm` file if needed

---

## Next Steps

1. ✅ Complete installation
2. ✅ Configure file paths
3. ✅ Customize categories
4. ✅ Run first import
5. 📊 Analyze your financial data!

For additional help, see:
- `TECHNICAL_SPECIFICATION.md` - Detailed system architecture
- `README.md` - Overview and features
- `USER_GUIDE.md` - Advanced usage tips

---

## Support

For issues or questions:
1. Check Audit sheet for error logs
2. Review this troubleshooting section
3. Consult technical documentation
4. Verify Excel for Mac version compatibility

**Version**: 1.0
**Last Updated**: January 2026
