# System Architecture - [MY LIFE]

## Overview

[MY LIFE] is designed as a modular, production-grade financial management system built entirely in VBA for Excel MacOS.

## Architecture Principles

1. **Modular Design**: Each functional area is a separate VBA module
2. **Configuration-Driven**: No hardcoded paths or settings
3. **Data Integrity**: Automatic correlation and validation
4. **Extensibility**: Easy to add new data sources or features
5. **User-Friendly**: Dashboard-driven with automated workflows

---

## Module Architecture

```
┌─────────────────────────────────────────────────────────┐
│                     DASHBOARD                            │
│            (Executive View & Controls)                   │
└─────────────────────────────────────────────────────────┘
                          ↓
┌─────────────────────────────────────────────────────────┐
│                  ORCHESTRATION LAYER                     │
│               (modMain - Master Control)                 │
└─────────────────────────────────────────────────────────┘
                          ↓
┌──────────────┬──────────────┬──────────────┬────────────┐
│   Import     │  Correlation │ Calculation  │ Validation │
│   Layer      │    Engine    │   Engine     │   Engine   │
└──────────────┴──────────────┴──────────────┴────────────┘
         ↓              ↓             ↓             ↓
┌─────────────────────────────────────────────────────────┐
│              CONFIGURATION & UTILITIES                   │
│          (modConfig, modUtils, modIndexes)               │
└─────────────────────────────────────────────────────────┘
```

---

## VBA Modules

### Core Modules

| Module | Responsibility | Dependencies |
|--------|----------------|--------------|
| `modMain` | Master orchestration, button handlers | All modules |
| `modConfig` | Configuration management, file paths | None |
| `modUtils` | Utility functions, helpers | None |

### Import Modules

| Module | Responsibility | Dependencies |
|--------|----------------|--------------|
| `modImportBanks` | Import bank transactions | modConfig, modUtils |
| `modImportCards` | Import credit card data | modConfig, modUtils |
| `modImportInvestments` | Import investments | modConfig, modUtils, modCorrelation |
| `modImportOPUS` | Import external investments | modConfig, modUtils |
| `modImportDebts` | Import debt information | modConfig, modUtils |

### Processing Modules

| Module | Responsibility | Dependencies |
|--------|----------------|--------------|
| `modClassification` | Transaction classification | modConfig, modUtils |
| `modCorrelation` | Investment-bank correlation | modUtils |
| `modCapitalCost` | Capital cost calculations | modIndexes, modUtils |
| `modIndexes` | Index management, cumulative factors | modUtils |

### Dashboard & Validation

| Module | Responsibility | Dependencies |
|--------|----------------|--------------|
| `modDashboard` | Dashboard updates, KPIs | All data modules |
| `modHealthCheck` | Data validation, integrity | All modules |

---

## Data Flow

### Import Flow

```
FILES PATHS Config
        ↓
   Import Modules
        ↓
   Raw Data Sheets (BANKS, CARDS, etc.)
        ↓
  Classification Engine
        ↓
 Correlation Engine (for investments)
        ↓
   Classified Data
        ↓
     DASHBOARD
```

### Investment Correlation Flow

```
BANKS Transactions          INVESTMENTS Transactions
        ↓                            ↓
        └────────→ Correlation Engine ←───────┘
                          ↓
                   Match Algorithm:
                   - Same date (±3 days)
                   - Opposite amounts
                   - Bank ↔ Investment
                          ↓
                   Link with CorrelationID
                          ↓
              Validated Balanced Movements
```

### Capital Cost Calculation Flow

```
INDEXES Sheet
    ↓
Cumulative Factor Calculation
    ↓
Historical Index Values
    ↓
Capital Cost Formula:
    Initial Value × Cumulative Factor(from → to)
    ↓
Updated Values
```

---

## Worksheet Definitions

### Configuration Sheets

#### FILES PATHS
```
| Source      | FilePath                              | Active |
|-------------|---------------------------------------|--------|
| ITAU_BANK   | /Users/.../itau_extrato.csv          | TRUE   |
| NUBANK_CARD | /Users/.../nubank_fatura.xlsx        | TRUE   |
```

#### FILES STRUCTURE
```
| Source      | ExpectedColumns                           |
|-------------|-------------------------------------------|
| ITAU_BANK   | Date,Description,Value,Balance           |
| NUBANK_CARD | Date,Description,Installment,Value       |
```

### Data Sheets

#### BANKS
```
| Bank   | Date       | Description         | Value    | Category | Subcategory | CorrelationID |
|--------|------------|---------------------|----------|----------|-------------|---------------|
| ITAU   | 2024-01-15 | TRANSFERENCIA       | -5000.00 | Transfer | Investment  | CORR-001     |
```

#### CARDS
```
| Bank   | Card | PurchaseDate | Category | Description  | Installment | Value  |
|--------|------|--------------|----------|--------------|-------------|--------|
| NUBANK | 1234 | 2024-01-10   | Food     | RESTAURANT   | 1/1         | -150.00|
```

#### INVESTMENTS
```
| Bank   | Date       | Description    | Value   | Category    | CorrelationID |
|--------|------------|----------------|---------|-------------|---------------|
| ITAU   | 2024-01-15 | APLICACAO CDB  | 5000.00 | Application | CORR-001     |
```

#### OPUS
```
| Type   | Company      | InvestmentCost | CapitalCost | UpdatedCost | Currency |
|--------|--------------|----------------|-------------|-------------|----------|
| Asset  | Startup ABC  | 100000.00      | 5.5%        | 105500.00   | USD      |
```

#### DEBTS
```
| Creditor  | InterestRate | AmountPaid | UpdatedAmount | Currency | LastUpdate |
|-----------|--------------|------------|---------------|----------|------------|
| Bank XYZ  | 2.5%         | 10000.00   | 10250.00      | BRL      | 2024-01-31 |
```

#### INDEXES
```
| Index  | Date       | Value  | CumulativeFactor |
|--------|------------|--------|------------------|
| CDI    | 2024-01-01 | 0.1085 | 1.0000          |
| CDI    | 2024-01-02 | 0.1085 | 1.0003          |
| CDI    | 2024-01-03 | 0.1090 | 1.0006          |
```

#### CATEGORIES
```
| Category       | Subcategory  | Keywords                    | Priority |
|----------------|--------------|----------------------------|----------|
| Food & Dining  | Restaurants  | RESTAURANT,RESTAURANTE,BAR | 10       |
| Transportation | Gas          | POSTO,GAS,SHELL,IPIRANGA   | 10       |
```

### Dashboard Sheet

#### DASHBOARD Structure

**Filters Section** (Top)
```
Year: [Dropdown]    Month: [Dropdown]    Institution: [Dropdown]    Currency: [Dropdown]
```

**KPIs Section**
```
┌─────────────────┬─────────────────┬─────────────────┐
│  Total Income   │ Total Expenses  │    Balance      │
│  R$ 150,000.00  │  R$ 95,000.00   │  R$ 55,000.00   │
└─────────────────┴─────────────────┴─────────────────┘
```

**Consolidated Tables**
1. Consolidate Cash (by Bank/Investment)
2. Consolidate Cards (by Card/Bank)
3. Consolidate Transactions (by Category)
4. Consolidate Net Debts (by Creditor)

**Charts**
1. Monthly Trend (Line Chart: Income vs Expenses vs Balance)
2. Category Breakdown (Pie Chart: Expenses by Category)

---

## Key Algorithms

### 1. Investment Correlation Algorithm

```
FOR each investment transaction:
    IF type = "Application":
        Search BANKS for:
            - Date within ±3 days
            - Value = -investment.value
            - No existing correlation
        IF found:
            Link with unique CorrelationID
            Mark both as correlated

    IF type = "Redemption":
        Search BANKS for:
            - Date within ±3 days
            - Value = +investment.value
            - No existing correlation
        IF found:
            Link with unique CorrelationID
            Mark both as correlated
```

### 2. Transaction Classification Algorithm

```
Phase 1: Exact Match
    FOR each keyword in CATEGORIES:
        IF keyword IN UPPER(description):
            RETURN category, subcategory

Phase 2: Proximity Match
    FOR each category:
        score = TextSimilarity(description, keywords)
        IF score > threshold:
            RETURN category with confidence score

Phase 3: Manual Resolution
    IF no match:
        LOG to unclassified list
        PROMPT user for manual mapping
        SAVE mapping for future use
```

### 3. Cumulative Factor Calculation

```
CumulativeFactor[0] = 1.0

FOR each date from start to end:
    dailyRate = annualRate / 252  (252 = business days/year)
    CumulativeFactor[n] = CumulativeFactor[n-1] × (1 + dailyRate)
```

### 4. Capital Cost Calculation

```
UpdatedValue = InitialValue × (
    CumulativeFactor[endDate] / CumulativeFactor[startDate]
)
```

---

## Health Check Validations

### Import Validation
- All configured paths are accessible
- All imports completed successfully
- Row counts match expected ranges

### Data Integrity
- All transactions have dates
- All values are numeric
- No duplicate transactions (by hash)

### Correlation Validation
- All investment applications have bank withdrawals
- All investment redemptions have bank deposits
- Amounts match within tolerance (0.01%)

### Classification Validation
- All transactions classified
- No orphan categories
- Classification confidence tracked

### Index Validation
- All indexes have daily data
- No gaps in cumulative factors
- Indexes up to date (last 30 days)

### Balance Validation
- Total inflows - total outflows = net position
- Cross-check with bank balances (if provided)

---

## Named Ranges

### KPIs
- `Total_Income`: Sum of all income transactions (filtered)
- `Total_Expenses`: Sum of all expense transactions (filtered)
- `Balance`: Total_Income - Total_Expenses

### Dashboard Data
- `Monthly_Trend`: Range for trend chart
- `Category_Breakdown`: Range for category pie chart
- `Cash_Consolidated`: Consolidated cash table
- `Cards_Consolidated`: Consolidated cards table
- `Transactions_Consolidated`: Consolidated transactions table
- `Debts_Consolidated`: Consolidated debts table

---

## Error Handling Strategy

### Levels

1. **Configuration Errors**: Missing paths, invalid settings
   - Action: Halt execution, prompt user to fix

2. **Import Errors**: File not found, format mismatch
   - Action: Log error, skip file, continue with others

3. **Data Errors**: Missing values, invalid formats
   - Action: Mark row as error, log to health check

4. **Correlation Errors**: No matching movement
   - Action: Flag as unmatched, report in health check

5. **Calculation Errors**: Missing index data
   - Action: Use last known value, flag warning

---

## Performance Considerations

### Optimization Strategies

1. **Batch Operations**: Process in chunks, update screen once
2. **Indexed Lookups**: Use dictionaries for fast searches
3. **Minimal Sheet Access**: Load data to arrays, process, write back
4. **Smart Recalculation**: Disable automatic calculation during imports
5. **Efficient Correlation**: Hash-based matching instead of nested loops

### Expected Performance

| Operation | 1K Rows | 10K Rows | 50K Rows |
|-----------|---------|----------|----------|
| Import    | 2 sec   | 10 sec   | 60 sec   |
| Classify  | 3 sec   | 15 sec   | 90 sec   |
| Correlate | 2 sec   | 12 sec   | 75 sec   |
| Dashboard | 1 sec   | 5 sec    | 30 sec   |

---

## Extension Points

### Adding New Data Source

1. Create `modImport[Source].bas`
2. Add source to FILES PATHS
3. Add structure to FILES STRUCTURE
4. Add import button to DASHBOARD
5. Update health check

### Adding New Index

1. Add row to INDEXES sheet
2. Add calculation logic to modIndexes
3. Update capital cost formulas

### Adding New Dashboard View

1. Add consolidated table structure
2. Create aggregation logic
3. Add chart if needed
4. Update named ranges

---

## Security & Privacy

- All data stored locally in Excel file
- No external API calls (except optional index updates)
- File can be password-protected
- VBA code can be locked

---

## Maintenance

### Daily
- Import new transactions
- Review health check

### Weekly
- Update indexes (if not automatic)
- Review unclassified transactions
- Validate correlations

### Monthly
- Review dashboard metrics
- Update capital costs
- Backup workbook

---

**Version**: 1.0
**Last Updated**: January 2026
