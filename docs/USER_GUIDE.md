# User Guide - [MY LIFE] Financial Management System

## Table of Contents

1. [Overview](#overview)
2. [Daily Workflow](#daily-workflow)
3. [Importing Data](#importing-data)
4. [Classification](#classification)
5. [Correlation](#correlation)
6. [Dashboard Usage](#dashboard-usage)
7. [Reports & Analysis](#reports--analysis)
8. [Troubleshooting](#troubleshooting)
9. [Best Practices](#best-practices)

---

## Overview

**[MY LIFE]** is a comprehensive financial management system that consolidates data from multiple sources into a single executive dashboard. This guide will help you use the system effectively for day-to-day financial management.

### System Capabilities

- ✅ Import transactions from multiple banks and credit cards
- ✅ Automatically classify transactions by category
- ✅ Correlate investment movements with bank transactions
- ✅ Track external investments and debts
- ✅ Calculate capital costs using financial indexes
- ✅ Generate executive dashboards with KPIs
- ✅ Validate data integrity with health checks

---

## Daily Workflow

### Typical Usage Pattern

```
1. Export new statements from banks/cards
   ↓
2. Save files to configured import folders
   ↓
3. Run import macros (Banks, Cards, Investments)
   ↓
4. Run classification
   ↓
5. Run correlation (if applicable)
   ↓
6. Review unclassified/uncorrelated items
   ↓
7. Update dashboard
   ↓
8. Run health check
   ↓
9. Review dashboard and KPIs
```

### Quick Start Checklist

Before each session:
- [ ] Have latest bank/card statements downloaded
- [ ] Check file paths are correct in FILES PATHS sheet
- [ ] Open MyLife.xlsm with macros enabled

---

## Importing Data

### Bank Transactions

**Purpose:** Import checking account movements from banks

**Steps:**
1. Go to VBA Editor (`Alt + F11`)
2. Run macro: `modImportBanks.ImportAllBanks`
3. Wait for completion message
4. Check **BANKS** sheet for imported data

**What gets imported:**
- Transaction date
- Description
- Amount (positive = deposit, negative = withdrawal)
- Bank name

**Supported formats:**
- CSV (comma, semicolon, or tab delimited)
- Excel (.xlsx, .xls, .xlsm)

**Auto-detected columns:**
- Date: `DATE`, `DATA`, `FECHA`
- Description: `DESCRIPTION`, `DESCRICAO`, `HISTORICO`, `MEMO`
- Value: `VALUE`, `VALOR`, `AMOUNT`

**Example bank CSV:**
```csv
Date,Description,Value,Balance
2024-01-15,TRANSFERENCIA PIX,-500.00,4500.00
2024-01-16,SALARIO,5000.00,9500.00
2024-01-17,POSTO SHELL,-150.00,9350.00
```

### Credit Card Transactions

**Purpose:** Import credit card statements with installment tracking

**Steps:**
1. Run macro: `modImportCards.ImportAllCards`
2. Check **CARDS** sheet for imported data

**What gets imported:**
- Purchase date
- Description/Merchant
- Amount
- Installment info (e.g., "2/12" for 2nd of 12 payments)
- Card number (last 4 digits)
- Category (if provided by bank)

**Example card CSV:**
```csv
PurchaseDate,Description,Installment,Value,Category
2024-01-10,RESTAURANTE ABC,1/1,-150.00,Food
2024-01-12,LOJA XYZ,3/12,-1200.00,Shopping
```

### Investment Transactions

**Purpose:** Import investment applications and redemptions

**Steps:**
1. Run macro: `modImportInvestments.ImportAllInvestments`
2. Check **INVESTMENTS** sheet for imported data

**What gets imported:**
- Transaction date
- Investment description/product
- Amount
- Transaction type (Application or Redemption)

**Type detection:**
The system automatically detects if a transaction is:
- **Application** (investment purchase): Keywords like APLICACAO, APPLICATION, INVEST, APORTE, BUY
- **Redemption** (investment sale): Keywords like RESGATE, REDEMPTION, WITHDRAW, SAQUE, SELL

**Example investment CSV:**
```csv
Date,Description,Value,Type
2024-01-15,APLICACAO CDB ITAU,5000.00,Application
2024-02-10,RESGATE CDB ITAU,5150.00,Redemption
```

### OPUS (External Investments)

**Purpose:** Import positions from external investment platforms

**Steps:**
1. Run macro: `modImportOPUS.ImportOPUS`
2. Check **OPUS** sheet for imported data

**What gets imported:**
- Type (Asset or Liability)
- Company/Investment name
- Investment cost (initial value)
- Capital cost (interest rate or percentage)
- Updated cost (calculated value)
- Currency (BRL, USD, etc.)

**Example OPUS CSV:**
```csv
Type,Company,InvestmentCost,CapitalCost,UpdatedCost,Currency
Asset,Startup ABC,100000.00,5.5%,105500.00,USD
Asset,Real Estate Fund,50000.00,CDI,52000.00,BRL
```

### Debts

**Purpose:** Track personal loans and debts with interest

**Steps:**
1. Run macro: `modImportDebts.ImportDebts`
2. Check **DEBTS** sheet for imported data

**What gets imported:**
- Creditor name
- Interest rate
- Amount paid
- Updated amount (with interest)
- Currency
- Last update date

**Example debts CSV:**
```csv
Creditor,InterestRate,AmountPaid,Currency,LastUpdate
Bank XYZ,2.5%,10000.00,BRL,2024-01-01
Friend Loan,0%,5000.00,BRL,2023-12-15
```

### Financial Indexes

**Purpose:** Import index data for capital cost calculations

**Steps:**
1. Run macro: `modIndexes.ImportIndexData`
2. System automatically calculates cumulative factors
3. Check **INDEXES** sheet for imported data

**Supported indexes:**
- **CDI** (Brazil interbank rate)
- **SELIC** (Brazil base rate)
- **IPCA** (Brazil inflation)
- **USD/BRL** (Dollar exchange rate)
- **FEDFUNDS** (US Federal Funds rate)

**Example indexes CSV:**
```csv
Index,Date,Value
CDI,2024-01-01,10.85
CDI,2024-01-02,10.85
CDI,2024-01-03,10.90
SELIC,2024-01-01,11.75
```

---

## Classification

### Automatic Classification

**Purpose:** Categorize transactions based on keywords

**How it works:**
1. **Exact Match** - Searches for keywords in transaction description
2. **Proximity Match** - Uses text similarity if no exact match found
3. **Manual Resolution** - Flag as "Unclassified" if no match

**Running classification:**
```vba
modClassification.ClassifyAllTransactions
```

**Result:** Category and Subcategory populated in BANKS sheet

### Managing Categories

**Location:** CATEGORIES sheet

**Structure:**
| Category | Subcategory | Keywords | Priority |
|----------|-------------|----------|----------|
| Food & Dining | Restaurants | RESTAURANT,BAR,CAFE | 10 |
| Food & Dining | Groceries | SUPERMARKET,MERCADO | 10 |

**Adding new categories:**
1. Go to **CATEGORIES** sheet
2. Add new row with:
   - Category name
   - Subcategory (optional)
   - Keywords (comma-separated, uppercase)
   - Priority (1-10, higher = more important)

**Tips:**
- Use multiple keywords for better matching
- Include variations (RESTAURANT, RESTAURANTE)
- Higher priority wins if multiple matches
- Keep keywords specific to avoid false matches

### Reviewing Unclassified

**Check count:**
```vba
modClassification.GetUnclassifiedCount()
```

**Finding unclassified:**
1. Go to **BANKS** sheet
2. Filter **Category** column for blank or "Unclassified"
3. Review descriptions
4. Add missing keywords to **CATEGORIES** sheet
5. Re-run classification

### Manual Classification

**For specific transactions:**
1. Find transaction in **BANKS** sheet
2. Manually enter Category and Subcategory
3. Optionally teach the system:
   ```vba
   modClassification.LearnFromManualClassification("STORE NAME", "Shopping", "Clothing")
   ```

This adds the keyword to CATEGORIES for future automatic matching.

### Reclassify All

**When to use:** After adding new category rules

**Steps:**
```vba
modClassification.ReclassifyAll
```

⚠️ **Warning:** This clears all existing classifications and re-runs the engine.

---

## Correlation

### Investment-Bank Correlation

**Purpose:** Link investment applications/redemptions with corresponding bank movements

**Why it matters:**
- Ensures balanced books (money out = investment in)
- Tracks investment lifecycle
- Validates data integrity
- Enables accurate net worth calculation

### How Correlation Works

**Algorithm:**
1. For each **Investment Application**:
   - Search BANKS for withdrawal (negative amount)
   - Within ±3 days of investment date
   - Matching amount (within R$ 0.01 tolerance)
   - Not already correlated

2. For each **Investment Redemption**:
   - Search BANKS for deposit (positive amount)
   - Within ±3 days of redemption date
   - Matching amount (within R$ 0.01 tolerance)
   - Not already correlated

3. When match found:
   - Generate unique CorrelationID
   - Link both transactions

**Example:**
```
Bank Transaction:
Date: 2024-01-15
Amount: -5000.00
CorrelationID: CORR-20240115143022-1234

Investment Transaction:
Date: 2024-01-15
Amount: 5000.00
Type: Application
CorrelationID: CORR-20240115143022-1234
```

### Running Correlation

**Steps:**
```vba
modCorrelation.CorrelateAllInvestments
```

**What happens:**
1. Clears existing correlations
2. Correlates applications (investment ← bank withdrawal)
3. Correlates redemptions (investment → bank deposit)
4. Shows count of correlations created

### Reviewing Correlation Results

**Check uncorrelated count:**
```vba
modCorrelation.GetUncorrelatedInvestmentCount()
```

**Validate balance:**
```vba
modCorrelation.ValidateCorrelationBalance()
```

**Manual correlation:**

If automatic correlation fails, manually correlate:
```vba
modCorrelation.ManualCorrelate(bankRow, investmentRow)
```

Example:
```vba
' Link bank row 150 with investment row 25
modCorrelation.ManualCorrelate(150, 25)
```

### Troubleshooting Correlation

**Issue: Uncorrelated investments**

Possible causes:
- Dates differ by more than 3 days
- Amounts don't match exactly
- Bank transaction already correlated
- Investment transaction missing from import

**Solution:**
1. Check dates are close (±3 days tolerance)
2. Verify amounts match
3. Use manual correlation if needed
4. Check both sheets have the transactions

---

## Dashboard Usage

### Understanding the Dashboard

**Location:** DASHBOARD sheet

**Components:**
1. **Filters** (Row 1) - Year, Month, Institution, Currency
2. **KPIs** (Row 2-3) - Total Income, Total Expenses, Balance
3. **Consolidated Cash** - Bank balances by institution
4. **Consolidated Cards** - Card totals
5. **Consolidated Transactions** - By category
6. **Consolidated Debts** - Debt summary

### Updating the Dashboard

**Full update:**
```vba
modDashboard.UpdateDashboard
```

**This refreshes:**
- All KPI calculations
- All consolidated tables
- Named ranges

**When to update:**
- After importing new data
- After classification
- After correlation
- Before reviewing metrics

### Using Filters

**Set filters in row 1:**

| Filter | Options | Effect |
|--------|---------|--------|
| Year | 2024, 2025, 2026, All | Filter by year |
| Month | Jan-Dec, All | Filter by month |
| Institution | ITAU, NUBANK, C6, BB, All | Filter by bank |
| Currency | BRL, USD, All | Filter by currency |

⚠️ **Note:** Current version calculates totals without filtering. Enhanced filtering requires formulas (planned for v2.0).

### Understanding KPIs

**Total Income:**
- Sum of all positive bank transactions
- Represents money coming in

**Total Expenses:**
- Sum of all negative bank transactions (absolute value)
- Represents money going out

**Balance:**
- Total Income - Total Expenses
- Net cash flow position

**Named ranges:**
You can reference KPIs in formulas:
```excel
=Total_Income
=Total_Expenses
=Balance
```

### Consolidated Views

**Cash Consolidated:**
- Shows balance by bank
- Includes investment positions
- Useful for seeing where money is located

**Cards Consolidated:**
- Total spending by card
- Helps track card usage

**Transactions Consolidated:**
- Spending by category
- Shows where money goes
- Useful for budgeting

**Debts Consolidated:**
- All debts with updated amounts
- Track interest accrual
- Monitor payment progress

### Creating Charts

**Manual chart creation:**
1. Select data range (e.g., category totals)
2. Go to **Insert > Chart**
3. Choose chart type (Pie, Column, Line)
4. Format as needed

**Recommended charts:**
- **Monthly Trend:** Line chart of income vs expenses over time
- **Category Breakdown:** Pie chart of expenses by category
- **Bank Distribution:** Pie chart of balances by bank

---

## Capital Cost Calculations

### Purpose

Track how investments and debts grow over time using financial indexes.

### Calculating Capital Costs

**For all OPUS and Debts:**
```vba
modCapitalCost.CalculateAllCapitalCosts
```

**What it does:**
- Updates OPUS investments with current values
- Updates debt amounts with accrued interest
- Uses appropriate index based on currency

### Index Selection

**Automatic selection:**
- **BRL investments/debts** → CDI index (default)
- **USD investments/debts** → Fed Funds index (default)

**Explicit index in data:**
If your capital cost column contains:
- `"CDI"` → Uses CDI
- `"SELIC"` → Uses SELIC
- `"IPCA"` → Uses IPCA
- `"5.5%"` → Uses simple interest at 5.5%

### Formula

```
UpdatedValue = InitialValue × (CumulativeFactor[end] / CumulativeFactor[start])
```

**Example:**
```
Initial Investment: R$ 10,000
Start Date: 2024-01-01 (Factor: 1.0000)
End Date: 2024-12-31 (Factor: 1.1085)
Updated Value: R$ 10,000 × (1.1085 / 1.0000) = R$ 11,085
```

### Viewing Results

**Check updated values:**
- **OPUS sheet** - Column "UpdatedCost"
- **DEBTS sheet** - Column "UpdatedAmount"

**Calculate ROI:**
```vba
modCapitalCost.CalculateROI(initialValue, currentValue)
```

Returns percentage return.

### Total Summaries

**Total investment value:**
```vba
modCapitalCost.GetTotalInvestmentValue()
```

**Total debt value:**
```vba
modCapitalCost.GetTotalDebtValue()
```

---

## Reports & Analysis

### Health Check Report

**Purpose:** Validate system integrity and data quality

**Running health check:**
```vba
modHealthCheck.RunHealthCheck
```

**Checks performed:**
1. **Import Status** - Data loaded from all sources
2. **Data Integrity** - Valid dates, amounts, descriptions
3. **Correlations** - Investments linked to bank movements
4. **Classifications** - Transactions categorized
5. **Indexes** - Index data available
6. **Balances** - Total inflows/outflows calculated

**Result codes:**
- 🟢 **PASS** - Check successful
- 🟡 **WARNING** - Minor issue, review recommended
- 🔴 **FAIL** - Critical issue, fix required

**Location:** HEALTH_CHECK sheet

**Interpreting results:**

| Status | Action Required |
|--------|-----------------|
| PASS | None, all good |
| WARNING | Review details, may need attention |
| FAIL | Must fix before relying on data |

### Quick System Status

**One-line summary:**
```vba
modHealthCheck.GetSystemStatus()
```

Returns: `"Banks: 150 | Cards: 45 | Investments: 12 | Unclassified: 5 | Uncorrelated: 2"`

### Export Data

**Export to CSV:**
1. Right-click sheet tab
2. Select **Move or Copy**
3. Choose **(new book)**
4. Save as CSV

**Recommended exports:**
- Monthly BANKS data for archival
- CATEGORIES for backup
- DASHBOARD summary for reports

---

## Troubleshooting

### Common Issues

#### Import Errors

**"File not found"**
- Check FILES PATHS sheet has correct path
- Verify file exists at that location
- Ensure no typos in path

**"Could not detect required columns"**
- Open source file and check column names
- Add alternative names to FILES STRUCTURE
- Ensure first row has headers

**"No data imported"**
- Check Active = TRUE for that source
- Verify file has data rows (not just headers)
- Check file format is supported (CSV or Excel)

#### Classification Issues

**Many unclassified transactions**
- Add more keywords to CATEGORIES
- Check keyword spelling
- Use common words from descriptions
- Increase priority for important categories

**Wrong categories assigned**
- Check for conflicting keywords
- Adjust priority levels
- Make keywords more specific
- Review category order

#### Correlation Issues

**High uncorrelated count**
- Check date tolerance (default ±3 days)
- Verify amounts match exactly
- Look for timing differences between sources
- Use manual correlation if needed

**Correlation imbalance**
- Run `ValidateCorrelationBalance()`
- Check for duplicate correlations
- Verify amount signs (positive/negative)
- Re-run correlation process

#### Dashboard Issues

**KPIs show zero**
- Ensure data imported to BANKS sheet
- Run UpdateDashboard
- Check for filter issues

**Consolidated tables empty**
- Run UpdateDashboard
- Check source sheets have data
- Verify sheet names match constants

### Getting More Help

1. Check **HEALTH_CHECK** sheet for diagnostic info
2. Review VBA Immediate Window for debug logs
3. Consult **ARCHITECTURE.md** for technical details
4. Try with sample data to isolate issues

---

## Best Practices

### Weekly Routine

1. **Monday morning:**
   - Export weekend transactions
   - Import to system
   - Classify and correlate
   - Review dashboard

2. **Friday afternoon:**
   - Check unclassified items
   - Add new categories if needed
   - Run health check
   - Update any manual entries

### Monthly Routine

1. **First of month:**
   - Import previous month's complete data
   - Import credit card statements
   - Reconcile with bank balances

2. **Mid-month:**
   - Update index data
   - Calculate capital costs
   - Review investment performance

3. **End of month:**
   - Export monthly reports
   - Backup workbook
   - Archive old data if needed

### Data Quality Tips

1. **Import early and often**
   - Don't wait for month-end
   - Import weekly or even daily
   - Easier to catch errors early

2. **Review unclassified immediately**
   - Add categories while transaction is fresh
   - Improves future auto-classification
   - Maintains clean data

3. **Validate correlations**
   - Run balance check after correlation
   - Investigate imbalances immediately
   - Keep investment records accurate

4. **Backup regularly**
   - Save workbook with date suffix
   - Keep source files archived
   - Export key sheets to CSV

5. **Document changes**
   - Note when adding new categories
   - Track manual correlations
   - Record any customizations

### Performance Tips

1. **Disable automatic calculation during imports:**
   - Already handled by import macros
   - Significantly speeds up large imports

2. **Close unnecessary workbooks:**
   - Keep only MyLife.xlsm open
   - Reduces memory usage

3. **Archive old data:**
   - After 2-3 years, move old transactions to separate workbook
   - Keeps current workbook fast

4. **Optimize categories:**
   - Remove unused categories
   - Consolidate similar keywords
   - Use higher priority for common transactions

---

## Keyboard Shortcuts

### Recommended Macros

Assign these shortcuts for quick access:

| Shortcut | Macro | Purpose |
|----------|-------|---------|
| Ctrl+Shift+B | ImportAllBanks | Import banks |
| Ctrl+Shift+C | ImportAllCards | Import cards |
| Ctrl+Shift+I | ImportAllInvestments | Import investments |
| Ctrl+Shift+L | ClassifyAllTransactions | Classify |
| Ctrl+Shift+O | CorrelateAllInvestments | Correlate |
| Ctrl+Shift+D | UpdateDashboard | Update dashboard |
| Ctrl+Shift+H | RunHealthCheck | Health check |

---

## Appendix: Sample Workflows

### Workflow 1: Monthly Reconciliation

```
1. Export bank statements (all accounts)
2. Export card statements (all cards)
3. Import banks → modImportBanks.ImportAllBanks
4. Import cards → modImportCards.ImportAllCards
5. Classify → modClassification.ClassifyAllTransactions
6. Review unclassified (should be < 5%)
7. Update dashboard → modDashboard.UpdateDashboard
8. Compare dashboard totals with bank balances
9. Run health check → modHealthCheck.RunHealthCheck
10. Fix any FAIL items
11. Export monthly report
12. Backup workbook
```

### Workflow 2: Investment Tracking

```
1. Export investment transactions
2. Export bank statements (same period)
3. Import both → modImportInvestments, modImportBanks
4. Correlate → modCorrelation.CorrelateAllInvestments
5. Check uncorrelated count
6. Manually correlate any missing links
7. Validate balance → ValidateCorrelationBalance()
8. Import index data → modIndexes.ImportIndexData
9. Calculate capital costs → modCapitalCost.CalculateAllCapitalCosts
10. Review updated investment values
11. Calculate ROI
```

### Workflow 3: Budget Analysis

```
1. Import all data for period
2. Classify all transactions
3. Update dashboard
4. Go to Consolidated Transactions table
5. Export to CSV
6. Analyze spending by category in Excel
7. Compare to budget
8. Identify areas to reduce spending
9. Add more detailed subcategories if needed
10. Re-classify with new categories
```

---

**User Guide Version:** 1.0
**Last Updated:** January 2026
**Compatible with:** [MY LIFE] v1.0
