# VBA Modules Documentation - [MY LIFE]

## Table of Contents

1. [Module Overview](#module-overview)
2. [Core Modules](#core-modules)
3. [Import Modules](#import-modules)
4. [Processing Modules](#processing-modules)
5. [Dashboard & Validation](#dashboard--validation)
6. [Function Reference](#function-reference)
7. [Development Guidelines](#development-guidelines)

---

## Module Overview

The [MY LIFE] system consists of 14 VBA modules organized into 4 layers:

```
┌─────────────────────────────────────────────────┐
│          ORCHESTRATION LAYER                    │
│              modMain                            │
└─────────────────────────────────────────────────┘
                     ↓
┌──────────────┬──────────────┬──────────────────┐
│   Import     │  Processing  │   Dashboard      │
│   Layer      │   Layer      │   & Validation   │
│              │              │                  │
│ Banks        │ Classification│  Dashboard      │
│ Cards        │ Correlation  │  HealthCheck    │
│ Investments  │ Indexes      │                 │
│ OPUS         │ CapitalCost  │                 │
│ Debts        │              │                 │
└──────────────┴──────────────┴──────────────────┘
                     ↓
┌─────────────────────────────────────────────────┐
│          FOUNDATION LAYER                       │
│        modConfig, modUtils                      │
└─────────────────────────────────────────────────┘
```

### Module Dependencies

| Module | Depends On |
|--------|------------|
| modConfig | None |
| modUtils | None |
| modMain | All modules |
| modImportBanks | modConfig, modUtils |
| modImportCards | modConfig, modUtils |
| modImportInvestments | modConfig, modUtils |
| modImportOPUS | modConfig, modUtils |
| modImportDebts | modConfig, modUtils |
| modClassification | modConfig, modUtils |
| modCorrelation | modUtils |
| modIndexes | modUtils |
| modCapitalCost | modIndexes, modUtils |
| modDashboard | All data modules |
| modHealthCheck | All modules |

---

## Core Modules

### modConfig

**Purpose:** Central configuration management

**Responsibilities:**
- Define all sheet name constants
- Define named range constants
- System initialization
- Create all required worksheets
- Manage configuration settings

**Key Constants:**
```vba
' Sheet Names
Public Const SHEET_FILES_PATHS As String = "FILES PATHS"
Public Const SHEET_BANKS As String = "BANKS"
Public Const SHEET_CARDS As String = "CARDS"
' ... etc

' Named Ranges
Public Const NR_TOTAL_INCOME As String = "Total_Income"
Public Const NR_TOTAL_EXPENSES As String = "Total_Expenses"
Public Const NR_BALANCE As String = "Balance"
```

**Key Functions:**
- `InitializeSystem()` - Creates all worksheets and structures
- `GetFilePath(sourceName)` - Retrieves file path from config
- `CreateXXXSheet()` - Individual sheet creation functions
- `ArrangeSheets()` - Orders sheets logically

**Usage:**
```vba
' Initialize system on first run
modConfig.InitializeSystem

' Get configured file path
Dim path As String
path = modConfig.GetFilePath("ITAU_BANK")
```

**Notes:**
- Run `InitializeSystem()` only once during setup
- All sheet names centralized here
- No hardcoded values elsewhere in system

---

### modUtils

**Purpose:** Utility functions library

**Responsibilities:**
- String manipulation and cleaning
- Date parsing and validation
- Numeric operations
- File operations (MacOS compatible)
- ID generation
- Progress indication
- Logging
- Currency handling
- Excel operations

**Key Functions:**

**String Operations:**
```vba
CleanText(text) As String
' Normalizes text: trim, uppercase, remove double spaces

ContainsAny(text, keywords, delimiter) As Boolean
' Checks if text contains any of the keywords

TextSimilarity(str1, str2) As Integer
' Returns similarity score 0-100
```

**Date Operations:**
```vba
ParseDate(dateStr) As Date
' Converts string to date safely

DatesMatch(date1, date2, toleranceDays) As Boolean
' Checks if dates are within tolerance

IsValidDate(value) As Boolean
' Validates if value is a valid date
```

**Numeric Operations:**
```vba
AmountsMatch(amount1, amount2, tolerance) As Boolean
' Checks if amounts are equal within tolerance (default 0.01)

RoundCurrency(value) As Double
' Rounds to 2 decimal places
```

**File Operations:**
```vba
FileExists(filePath) As Boolean
' Checks file existence (MacOS compatible)

GetFileExtension(fileName) As String
' Returns lowercase file extension
```

**ID Generation:**
```vba
GenerateUniqueID(prefix) As String
' Generates unique ID: PREFIX-yyyymmddhhnnss-random
```

**Progress & Logging:**
```vba
ShowProgress(message, current, total)
' Displays progress in status bar

ClearProgress()
' Clears status bar

LogDebug(message)
' Writes to Immediate Window with timestamp
```

**Excel Operations:**
```vba
ClearSheetData(ws)
' Clears sheet contents keeping headers

GetLastRow(ws, col) As Long
' Returns last row with data in column

AutoFitColumns(ws, fromCol, toCol)
' Auto-fits column widths
```

**Transaction Hashing:**
```vba
GenerateTransactionHash(transDate, amount, description) As String
' Creates hash for duplicate detection
```

**Usage Examples:**
```vba
' Clean and normalize text
Dim clean As String
clean = modUtils.CleanText("  Restaurant ABC  ")
' Result: "RESTAURANT ABC"

' Check if description contains keywords
If modUtils.ContainsAny(description, "UBER,TAXI,99POP") Then
    category = "Transportation"
End If

' Check if dates match within 3 days
If modUtils.DatesMatch(invDate, bankDate, 3) Then
    ' Correlate transactions
End If

' Show progress
modUtils.ShowProgress "Processing...", 50, 100
' Status bar: "Processing... (50%)"
```

---

### modMain

**Purpose:** Master orchestration and control

**Responsibilities:**
- Coordinate import operations
- Coordinate processing operations
- Button/UI handlers
- Main workflow orchestration

**Key Functions:**
```vba
RunFullImport()
' Imports all sources: banks, cards, investments, OPUS, debts

RunFullProcessing()
' Classification + Correlation + Dashboard update

RunCompleteWorkflow()
' Full import → Processing → Health Check
```

**Typical orchestration:**
```vba
Public Sub RunCompleteWorkflow()
    ' Import all data
    Call modImportBanks.ImportAllBanks
    Call modImportCards.ImportAllCards
    Call modImportInvestments.ImportAllInvestments
    Call modImportOPUS.ImportOPUS
    Call modImportDebts.ImportDebts

    ' Process data
    Call modClassification.ClassifyAllTransactions
    Call modCorrelation.CorrelateAllInvestments

    ' Update dashboard
    Call modDashboard.UpdateDashboard

    ' Validate
    Call modHealthCheck.RunHealthCheck
End Sub
```

**Usage:**
- Attach to buttons on DASHBOARD
- Create keyboard shortcuts
- Schedule with macros

---

## Import Modules

All import modules follow the same pattern established in `modImportBanks`.

### Common Pattern

```vba
' 1. Main import function
Public Function ImportAllXXX() As Long
    ' Loops through configured sources
    ' Returns total imported count
End Function

' 2. File type dispatcher
Private Function ImportXXXFile(sourceName, filePath) As Long
    ' Detects file extension
    ' Calls CSV or Excel importer
End Function

' 3. CSV importer
Private Function ImportXXXCSV(sourceName, filePath) As Long
    ' Opens file
    ' Detects delimiter
    ' Maps columns
    ' Parses rows
    ' Returns imported count
End Function

' 4. Excel importer
Private Function ImportXXXExcel(sourceName, filePath) As Long
    ' Opens workbook
    ' Reads headers
    ' Maps columns
    ' Parses rows
    ' Returns imported count
End Function

' 5. Row parser
Private Function ParseAndAddXXXRow(...) As Boolean
    ' Validates data
    ' Adds to worksheet
    ' Returns success/fail
End Function

' Helper functions
Private Function DetectDelimiter(headerLine) As String
Private Function FindColumn(headers, possibleNames) As Integer
Private Function GetColumnValue(rowData, colIndex) As String
```

### modImportBanks

**Purpose:** Import bank checking account transactions

**Sheet Target:** BANKS

**Columns Written:**
1. Bank name
2. Transaction date
3. Description
4. Value
5. Category (empty, filled by classification)
6. Subcategory (empty, filled by classification)
7. CorrelationID (empty, filled by correlation)
8. ImportDate (timestamp)

**Supported Sources:**
- ITAU_BANK
- NUBANK_BANK
- C6_BANK
- BB_BANK

**Auto-detected Columns:**
- Date: `DATE`, `DATA`, `FECHA`
- Description: `DESCRIPTION`, `DESCRICAO`, `HISTORICO`, `MEMO`
- Value: `VALUE`, `VALOR`, `AMOUNT`

**Special Logic:**
- Skips balance lines (value = 0 with "SALDO" in description)
- Negative values = withdrawals/payments
- Positive values = deposits/income

---

### modImportCards

**Purpose:** Import credit card transactions with installments

**Sheet Target:** CARDS

**Columns Written:**
1. Bank name
2. Card number
3. Purchase date
4. Category (from file if available)
5. Description
6. Installment (e.g., "3/12")
7. Value
8. ImportDate

**Supported Sources:**
- ITAU_CARD
- NUBANK_CARD
- C6_CARD
- BB_CARD

**Auto-detected Columns:**
- Date: `DATE`, `DATA`, `PURCHASEDATE`, `PURCHASE_DATE`
- Description: `DESCRIPTION`, `DESCRICAO`, `MERCHANT`, `ESTABELECIMENTO`
- Value: `VALUE`, `VALOR`, `AMOUNT`
- Category: `CATEGORY`, `CATEGORIA`, `TYPE`
- Installment: `INSTALLMENT`, `PARCELA`, `INSTALLMENTS`
- Card: `CARD`, `CARTAO`, `CARDNUMBER`, `FINAL`

**Special Logic:**
- Defaults installment to "1/1" if not provided
- Card category pre-filled but can be overridden by classification

---

### modImportInvestments

**Purpose:** Import investment transactions for correlation

**Sheet Target:** INVESTMENTS

**Columns Written:**
1. Bank/Institution
2. Transaction date
3. Description
4. Value
5. Type (Application/Redemption)
6. CorrelationID (empty, filled by correlation)
7. ImportDate

**Supported Sources:**
- ITAU_INV
- NUBANK_INV
- C6_INV
- BB_INV
- XP_INV
- BTG_INV

**Auto-detected Columns:**
- Date: `DATE`, `DATA`, `TRANSACTION_DATE`, `MOVIMENTACAO`
- Description: `DESCRIPTION`, `DESCRICAO`, `PRODUCT`, `PRODUTO`
- Value: `VALUE`, `VALOR`, `AMOUNT`
- Type: `TYPE`, `TIPO`, `OPERATION`, `OPERACAO`

**Type Detection:**
```vba
' Application keywords
APLICACAO, APPLICATION, INVEST, APORTE, BUY, COMPRA

' Redemption keywords
RESGATE, REDEMPTION, REDEEM, WITHDRAW, SAQUE, SELL, VENDA
```

---

### modImportOPUS

**Purpose:** Import external investment positions

**Sheet Target:** OPUS

**Columns Written:**
1. Type (Asset/Liability)
2. Company name
3. Investment cost
4. Capital cost rate
5. Updated cost
6. Currency
7. ImportDate

**Source:** Single OPUS file

**Auto-detected Columns:**
- Type: `TYPE`, `TIPO`, `ASSET_LIABILITY`
- Company: `COMPANY`, `EMPRESA`, `NAME`, `NOME`
- InvestmentCost: `INVESTMENTCOST`, `INVESTMENT_COST`, `INITIAL_VALUE`
- CapitalCost: `CAPITALCOST`, `CAPITAL_COST`, `INTEREST`, `JUROS`
- UpdatedCost: `UPDATEDCOST`, `UPDATED_COST`, `CURRENT_VALUE`
- Currency: `CURRENCY`, `MOEDA`

**Type Normalization:**
```vba
' Asset keywords
ASSET, ATIVO, INVESTMENT, INVESTIMENTO

' Liability keywords
LIABILITY, PASSIVO, DEBT, DIVIDA
```

---

### modImportDebts

**Purpose:** Import debt tracking information

**Sheet Target:** DEBTS

**Columns Written:**
1. Creditor
2. Interest rate
3. Amount paid
4. Updated amount
5. Currency
6. Last update date
7. ImportDate

**Source:** Single DEBTS file

**Auto-detected Columns:**
- Creditor: `CREDITOR`, `CREDOR`, `BANK`, `BANCO`, `LENDER`
- Rate: `INTERESTRATE`, `INTEREST_RATE`, `RATE`, `TAXA`
- Paid: `AMOUNTPAID`, `AMOUNT_PAID`, `VALOR_PAGO`, `INITIAL`
- Updated: `UPDATEDAMOUNT`, `UPDATED_AMOUNT`, `CURRENT`
- Currency: `CURRENCY`, `MOEDA`
- Date: `LASTUPDATE`, `LAST_UPDATE`, `DATE`, `DATA`

**Defaults:**
- Currency: BRL if not specified
- LastUpdate: Today's date if not specified

---

## Processing Modules

### modClassification

**Purpose:** Automatic transaction classification

**Algorithm:**
```
Phase 1: Exact Match
  FOR each category rule:
    IF any keyword IN transaction description:
      RETURN category with highest priority

Phase 2: Proximity Match
  FOR each category rule:
    score = TextSimilarity(description, keywords)
    IF score >= 70 AND score > bestScore:
      SAVE as best match
  RETURN best match if score >= 70

Phase 3: No Match
  RETURN "Unclassified"
```

**Key Functions:**

```vba
ClassifyAllTransactions() As Long
' Classifies all unclassified bank transactions
' Returns count of classified transactions

ClassifyTransaction(description, outCategory, outSubcategory) As Boolean
' Classifies single transaction
' Returns true if classified

ExactMatch(description, outCategory, outSubcategory) As Boolean
' Phase 1: Exact keyword matching

ProximityMatch(description, outCategory, outSubcategory) As Boolean
' Phase 2: Similarity-based matching (threshold: 70)

GetUnclassifiedCount() As Long
' Returns count of unclassified transactions

LearnFromManualClassification(description, category, subcategory)
' Adds keyword to CATEGORIES for future matching

ReclassifyAll()
' Clears all classifications and re-runs
```

**Configuration:**
- `SIMILARITY_THRESHOLD = 70` (0-100 scale)

**Usage:**
```vba
' Classify all
Dim count As Long
count = modClassification.ClassifyAllTransactions()

' Check unclassified
Dim unclassified As Long
unclassified = modClassification.GetUnclassifiedCount()

' Teach system from manual entry
modClassification.LearnFromManualClassification("UBER TRIP", "Transportation", "Rideshare")
```

---

### modCorrelation

**Purpose:** Link investment movements with bank transactions

**Algorithm:**
```
FOR each investment application:
  SEARCH banks WHERE:
    - date within ±3 days
    - amount = -investment.amount (withdrawal)
    - not already correlated
  IF found:
    CREATE unique CorrelationID
    LINK both transactions

FOR each investment redemption:
  SEARCH banks WHERE:
    - date within ±3 days
    - amount = +investment.amount (deposit)
    - not already correlated
  IF found:
    CREATE unique CorrelationID
    LINK both transactions
```

**Key Functions:**

```vba
CorrelateAllInvestments() As Long
' Correlates all investments with bank transactions
' Returns count of correlations created

ClearExistingCorrelations()
' Clears all correlation IDs

CorrelateApplications() As Long
' Links investment applications with bank withdrawals

CorrelateRedemptions() As Long
' Links investment redemptions with bank deposits

GetUncorrelatedInvestmentCount() As Long
' Returns count of uncorrelated investments

ValidateCorrelationBalance() As Boolean
' Checks that all correlations balance to zero
' Returns true if balanced

ManualCorrelate(bankRow, invRow)
' Manually links specific transactions
```

**Configuration:**
- `DATE_TOLERANCE_DAYS = 3` (±3 days)
- `AMOUNT_TOLERANCE = 0.01` (R$ 0.01)

**Usage:**
```vba
' Correlate all
Dim count As Long
count = modCorrelation.CorrelateAllInvestments()

' Check balance
Dim balanced As Boolean
balanced = modCorrelation.ValidateCorrelationBalance()

' Manual link
modCorrelation.ManualCorrelate(bankRow:=150, invRow:=25)
```

---

### modIndexes

**Purpose:** Manage financial indexes and calculate cumulative factors

**Algorithm:**
```
CumulativeFactor[0] = 1.0

FOR each date:
  dailyRate = annualRate / 252  (252 business days/year)
  CumulativeFactor[n] = CumulativeFactor[n-1] × (1 + dailyRate)
```

**Key Functions:**

```vba
CalculateAllCumulativeFactors()
' Calculates cumulative factors for all indexes

CalculateCumulativeFactorsForIndex(indexName)
' Calculates for specific index

GetCumulativeFactor(indexName, targetDate) As Double
' Returns cumulative factor for date

GetIndexValue(indexName, targetDate) As Double
' Returns index value for date

ImportIndexData() As Long
' Imports index data from file
' Auto-calculates cumulative factors
```

**Supported Indexes:**
- CDI (Brazil)
- SELIC (Brazil)
- IPCA (Brazil)
- USD/BRL (Exchange rate)
- FEDFUNDS (US)

**Configuration:**
- `BUSINESS_DAYS_PER_YEAR = 252`

**Usage:**
```vba
' Import and calculate
modIndexes.ImportIndexData()

' Get factor for date
Dim factor As Double
factor = modIndexes.GetCumulativeFactor("CDI", #1/15/2024#)

' Get index value
Dim rate As Double
rate = modIndexes.GetIndexValue("SELIC", Date)
```

---

### modCapitalCost

**Purpose:** Calculate updated values using index adjustments

**Formula:**
```
UpdatedValue = InitialValue × (CumulativeFactor[end] / CumulativeFactor[start])
```

**Key Functions:**

```vba
CalculateAllCapitalCosts()
' Updates all OPUS and DEBTS with calculated values

CalculateOPUSCapitalCosts()
' Updates OPUS investment values

CalculateDebtCapitalCosts()
' Updates debt amounts with interest

DetermineIndexForCurrency(currency, rateText) As String
' Selects appropriate index:
'   BRL → CDI (default)
'   USD → FEDFUNDS (default)
'   Or explicit index in rateText

CalculateUpdatedValue(initialValue, indexName, endDate, startDate) As Double
' Applies index adjustment to value

CalculateSimpleInterest(principal, rateText) As Double
' Fallback for percentage-based rates

CalculateROI(initialValue, currentValue) As Double
' Returns percentage return

GetTotalInvestmentValue() As Double
' Sum of all OPUS assets

GetTotalDebtValue() As Double
' Sum of all debts
```

**Usage:**
```vba
' Calculate all
modCapitalCost.CalculateAllCapitalCosts()

' Calculate specific value
Dim updated As Double
updated = modCapitalCost.CalculateUpdatedValue( _
    initialValue:=10000, _
    indexName:="CDI", _
    endDate:=Date, _
    startDate:=#1/1/2024#)

' Get totals
Dim totalInv As Double
totalInv = modCapitalCost.GetTotalInvestmentValue()
```

---

## Dashboard & Validation

### modDashboard

**Purpose:** Build and update executive dashboard

**Responsibilities:**
- Calculate KPIs (Income, Expenses, Balance)
- Build consolidated tables
- Define named ranges
- Create filter dropdowns
- Format dashboard

**Key Functions:**

```vba
UpdateDashboard()
' Updates all dashboard components

UpdateKPIs()
' Calculates and writes KPI values to dashboard

UpdateConsolidatedCash()
' Aggregates bank balances by institution

UpdateConsolidatedCards()
' Aggregates card totals

UpdateConsolidatedTransactions()
' Aggregates by category

UpdateConsolidatedDebts()
' Lists debts with updated amounts

DefineNamedRanges()
' Creates named ranges for KPIs

CreateFilterDropdowns()
' Creates data validation dropdowns

FormatDashboard()
' Applies formatting and styling

SetupDashboard()
' One-time setup: filters + format + update
```

**Dashboard Layout:**
```
Row 1: Filters (Year, Month, Institution, Currency)
Row 2-3: KPIs (Income, Expenses, Balance)
Row 7+: Consolidated Cash table
Row 7+ (Col E): Consolidated Cards table
Row 34+: Consolidated Transactions table
Row 34+ (Col E): Consolidated Debts table
```

**Usage:**
```vba
' First time setup
modDashboard.SetupDashboard()

' Regular updates
modDashboard.UpdateDashboard()
```

---

### modHealthCheck

**Purpose:** System validation and health reporting

**Checks Performed:**
1. **Import Status** - Data loaded from sources
2. **Data Integrity** - Valid dates, amounts, descriptions
3. **Correlations** - Investments linked, balanced
4. **Classifications** - Transactions categorized
5. **Indexes** - Index data available
6. **Balances** - Calculations accurate

**Key Functions:**

```vba
RunHealthCheck()
' Runs all checks, writes to HEALTH_CHECK sheet

CheckImportStatus(ws, startRow) As Long
' Validates import counts

CheckDataIntegrity(ws, startRow) As Long
' Validates data quality

CheckCorrelations(ws, startRow) As Long
' Validates correlation status

CheckClassifications(ws, startRow) As Long
' Validates classification status

CheckIndexes(ws, startRow) As Long
' Validates index data

CheckBalances(ws, startRow) As Long
' Validates balance calculations

ShowHealthSummary(ws)
' Shows summary with pass/warning/fail counts

GetSystemStatus() As String
' Returns one-line status summary
```

**Status Codes:**
- `PASS` (Green) - Check successful
- `WARNING` (Orange) - Review recommended
- `FAIL` (Red) - Critical issue

**Usage:**
```vba
' Full health check
modHealthCheck.RunHealthCheck()

' Quick status
Dim status As String
status = modHealthCheck.GetSystemStatus()
' Returns: "Banks: 150 | Cards: 45 | Investments: 12 | ..."
```

---

## Function Reference

### Quick Reference Table

| Module | Public Functions | Purpose |
|--------|------------------|---------|
| **modConfig** | `InitializeSystem()` | Setup |
| | `GetFilePath(source)` | Config |
| **modUtils** | `CleanText(text)` | String |
| | `ContainsAny(text, keywords)` | String |
| | `TextSimilarity(str1, str2)` | String |
| | `ParseDate(dateStr)` | Date |
| | `DatesMatch(d1, d2, tol)` | Date |
| | `AmountsMatch(a1, a2, tol)` | Numeric |
| | `FileExists(path)` | File |
| | `GetFileExtension(name)` | File |
| | `GenerateUniqueID(prefix)` | ID |
| | `ShowProgress(msg, cur, tot)` | UI |
| | `LogDebug(message)` | Debug |
| | `GetLastRow(ws, col)` | Excel |
| **modMain** | `RunFullImport()` | Orchestration |
| | `RunFullProcessing()` | Orchestration |
| | `RunCompleteWorkflow()` | Orchestration |
| **modImportBanks** | `ImportAllBanks()` | Import |
| **modImportCards** | `ImportAllCards()` | Import |
| **modImportInvestments** | `ImportAllInvestments()` | Import |
| **modImportOPUS** | `ImportOPUS()` | Import |
| **modImportDebts** | `ImportDebts()` | Import |
| **modClassification** | `ClassifyAllTransactions()` | Processing |
| | `GetUnclassifiedCount()` | Query |
| | `ReclassifyAll()` | Processing |
| **modCorrelation** | `CorrelateAllInvestments()` | Processing |
| | `GetUncorrelatedInvestmentCount()` | Query |
| | `ValidateCorrelationBalance()` | Validation |
| | `ManualCorrelate(bRow, iRow)` | Manual |
| **modIndexes** | `ImportIndexData()` | Import |
| | `CalculateAllCumulativeFactors()` | Processing |
| | `GetCumulativeFactor(idx, date)` | Query |
| **modCapitalCost** | `CalculateAllCapitalCosts()` | Processing |
| | `CalculateUpdatedValue(...)` | Calculation |
| | `GetTotalInvestmentValue()` | Query |
| | `GetTotalDebtValue()` | Query |
| **modDashboard** | `UpdateDashboard()` | Update |
| | `SetupDashboard()` | Setup |
| **modHealthCheck** | `RunHealthCheck()` | Validation |
| | `GetSystemStatus()` | Query |

---

## Development Guidelines

### Code Standards

**1. Error Handling**

All public functions must have error handlers:
```vba
Public Function MyFunction() As Long
    On Error GoTo ErrorHandler

    ' Function code here

    MyFunction = result
    Exit Function

ErrorHandler:
    MsgBox modUtils.FormatErrorMessage("ModuleName", "MyFunction", Err.Description), vbCritical
    MyFunction = 0
End Function
```

**2. Progress Indication**

Long-running operations must show progress:
```vba
modUtils.ShowProgress "Processing...", currentItem, totalItems
DoEvents  ' Allow UI updates
```

Always clear progress when done:
```vba
modUtils.ClearProgress
```

**3. Logging**

Use debug logging for troubleshooting:
```vba
modUtils.LogDebug "Starting import from: " & filePath
```

**4. Screen Updating**

Disable for performance, always re-enable:
```vba
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

' ... processing ...

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
```

**5. Function Naming**

- Public functions: `PascalCase`
- Private functions: `PascalCase`
- Variables: `camelCase`
- Constants: `UPPER_SNAKE_CASE`

**6. Comments**

- Header block for each module
- Function purpose comments
- Complex logic explanation
- TODO markers for future work

**Example:**
```vba
' ========================================================================
' Module: modExample
' Purpose: Example module for documentation
' Description: Shows coding standards and patterns
' Platform: Excel for MacOS
' Author: [MY LIFE] Financial System v1.0
' ========================================================================

Option Explicit

' ========================================================================
' Main Function - Does Something Important
' ========================================================================
Public Function DoSomething(param As String) As Long
    On Error GoTo ErrorHandler

    ' Variable declarations
    Dim result As Long
    Dim ws As Worksheet

    ' Initialize
    result = 0
    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_BANKS)

    ' Main logic
    ' TODO: Add validation

    DoSomething = result
    Exit Function

ErrorHandler:
    MsgBox modUtils.FormatErrorMessage("modExample", "DoSomething", Err.Description), vbCritical
    DoSomething = 0
End Function
```

### Adding New Features

**1. New Import Source**

To add a new bank/card:
1. Copy `modImportBanks.bas` to `modImportNewSource.bas`
2. Update source names array
3. Adjust column mappings
4. Add to FILES PATHS configuration
5. Update modMain orchestration

**2. New Classification Category**

Simply add to CATEGORIES sheet:
- No code changes needed
- System auto-detects new categories

**3. New Index**

1. Add to INDEXES sheet
2. Update `modIndexes.CalculateAllCumulativeFactors()`
3. Update `modCapitalCost.DetermineIndexForCurrency()` if needed

**4. New Dashboard KPI**

1. Add calculation to `modDashboard.UpdateKPIs()`
2. Add named range in `modDashboard.DefineNamedRanges()`
3. Update dashboard layout

### Testing

**Unit Testing:**
```vba
' Test function in Immediate Window
? modUtils.TextSimilarity("RESTAURANT", "RESTAURANTE")
' Result: 90

? modUtils.DatesMatch(#1/15/2024#, #1/17/2024#, 3)
' Result: True
```

**Integration Testing:**
1. Use sample data files
2. Run full workflow
3. Validate results in HEALTH_CHECK

**Debugging:**
```vba
' Add breakpoints (F9)
' Step through code (F8)
' Watch variables (Add Watch)
' Check Immediate Window for logs
```

### Performance Optimization

**1. Batch Operations**
```vba
' Bad: Update cells one by one
For i = 1 To 1000
    ws.Cells(i, 1).Value = data(i)
Next

' Good: Update range at once
ws.Range("A1:A1000").Value = data
```

**2. Array Processing**
```vba
' Load to array, process, write back
Dim data As Variant
data = ws.Range("A1:A1000").Value

For i = 1 To UBound(data)
    data(i, 1) = UCase(data(i, 1))
Next

ws.Range("A1:A1000").Value = data
```

**3. Indexed Lookups**
```vba
' Use Dictionary for fast lookups
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")

' Add items
dict.Add "key", "value"

' Fast lookup
If dict.Exists("key") Then
    result = dict("key")
End If
```

---

## Appendix: Module Metrics

| Module | Lines | Functions | Complexity |
|--------|-------|-----------|------------|
| modConfig | ~600 | 15 | Medium |
| modUtils | ~400 | 25 | Low |
| modMain | ~200 | 5 | Low |
| modImportBanks | ~350 | 10 | Medium |
| modImportCards | ~420 | 10 | Medium |
| modImportInvestments | ~440 | 11 | Medium |
| modImportOPUS | ~380 | 10 | Medium |
| modImportDebts | ~390 | 10 | Medium |
| modClassification | ~320 | 8 | High |
| modCorrelation | ~350 | 9 | High |
| modIndexes | ~420 | 12 | Medium |
| modCapitalCost | ~350 | 10 | Medium |
| modDashboard | ~380 | 11 | Medium |
| modHealthCheck | ~450 | 10 | High |
| **Total** | **~5,650** | **156** | |

---

**Module Documentation Version:** 1.0
**Last Updated:** January 2026
**System Version:** [MY LIFE] v1.0
