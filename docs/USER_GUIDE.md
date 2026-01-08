# User Guide - Advanced Features & Tips

## Table of Contents
1. [Understanding the System](#understanding-the-system)
2. [Advanced Classification](#advanced-classification)
3. [Transfer Detection](#transfer-detection)
4. [Investment Tracking](#investment-tracking)
5. [Error Resolution Strategies](#error-resolution-strategies)
6. [Custom Workflows](#custom-workflows)
7. [Best Practices](#best-practices)

---

## Understanding the System

### Data Flow Overview

```
Input Files → Import → Raw Data → Normalization →
   ↓
Transactions Table
   ↓
   ├─→ Transfer Detection
   ├─→ Investment Detection
   └─→ Category Classification
       ↓
   Final Categorized Data
```

### Sheet Purposes

| Sheet | Purpose | User Editable |
|-------|---------|---------------|
| **Dashboard** | Control center | Buttons only |
| **Config** | Settings & paths | ✅ Yes |
| **Categories** | Category rules | ✅ Yes |
| **Transactions** | Final output | View only |
| **Errors** | Unresolved issues | ✅ Resolve manually |
| **Audit** | System log | View only |
| **RawData** | Import staging | Hidden |

---

## Advanced Classification

### Understanding Match Priority

The classification engine uses a two-phase approach:

#### Phase 1: Exact Matching
- Searches for exact keyword matches in description
- Case-insensitive
- Fastest and most accurate
- Uses Priority field (1 = highest)

#### Phase 2: Fuzzy Matching
- Calculates similarity score (0-100%)
- Uses Levenshtein distance algorithm
- Threshold configurable (default: 80%)
- Slower but catches variations

### Optimizing Category Rules

#### 1. Priority Strategy

```
Priority 1-10:    Critical, high-frequency transactions
Priority 10-50:   Common categories
Priority 50-100:  Specific merchants
Priority 100+:    Catch-all rules
```

**Example**:
```
Category: Food & Dining
Subcategory: Groceries
Keywords: WHOLE FOODS,TRADER JOES,SAFEWAY,GROCERY
Priority: 10
```

#### 2. Keyword Best Practices

**✅ Good Keywords**:
- Specific: `STARBUCKS` not `COFFEE`
- Partial: `AMAZON` catches `AMAZON.COM`, `AMAZON PRIME`
- Multiple variations: `RESTAURANT,RESTAURANTE,RESTAU`

**❌ Avoid**:
- Too generic: `FOOD`, `STORE`, `MARKET`
- Too specific: `STARBUCKS #1234 MAIN ST` (won't match other locations)
- Overlapping: Don't use same keyword in multiple categories

#### 3. Testing Classification

After adding new rules:
1. Add keywords to Categories sheet
2. Click **Update Categories**
3. Click **Reprocess Data**
4. Check results in Transactions sheet
5. Adjust FuzzyThreshold if needed

### Handling Ambiguous Transactions

Some transactions are hard to classify automatically:

**Strategy 1: Manual Override**
1. Find transaction in Errors sheet
2. Enter category in SuggestedCategory column
3. System learns and suggests keywords

**Strategy 2: Create Specific Rules**
1. Extract unique identifier from description
2. Add as high-priority keyword
3. Example: "PAYPAL *SPOTIFY" → Add "PAYPAL *SPOTIFY" keyword

**Strategy 3: Use Subcategories**
1. Same merchant, different purposes
2. Create specific subcategories
3. Example:
   - AMAZON → Shopping:General
   - AMAZON PRIME VIDEO → Entertainment:Streaming

---

## Transfer Detection

### How It Works

Transfers are detected when:
1. **Date Match**: Within N days (default: 1)
2. **Amount Match**: Opposite amounts within tolerance
3. **Account Difference**: Different accounts
4. **Sign Opposite**: One debit, one credit

### Configuration

Adjust in Config sheet:
```
TransferTolerance: 0.001      (0.1% difference allowed)
TransferDateDays: 1           (check ±1 day)
```

### Common Scenarios

#### Scenario 1: Same-Day Transfers
```
Account A: -500.00 on 01/15
Account B: +500.00 on 01/15
→ Automatically detected
```

#### Scenario 2: Multi-Day Transfers
```
Account A: -500.00 on 01/15
Account B: +500.00 on 01/17
→ Detected if TransferDateDays ≥ 2
```

#### Scenario 3: Transfers with Fees
```
Account A: -500.00 on 01/15
Account B: +495.00 on 01/15 (5.00 fee)
→ NOT detected (amounts don't match)
```

**Solution**: Lower TransferTolerance or manually link

### Manual Transfer Linking

For transfers not auto-detected:

1. Note TransactionIDs from Transactions sheet
2. Open VBA Editor (Alt+F11)
3. Run in Immediate Window:
   ```vba
   modTransfers.LinkTransactions "TXN-123", "TXN-456"
   ```

### Unlinking Transfers

If transfer detected incorrectly:

```vba
modTransfers.UnlinkTransfer "TXN-123"
```

---

## Investment Tracking

### Supported Investment Types

| Type | Keywords | Amount Sign |
|------|----------|-------------|
| Application | APLICACAO, INVESTMENT, BUY | Negative |
| Redemption | RESGATE, REDEMPTION, SELL | Positive |
| Liquidation | LIQUIDACAO, CLOSURE | Positive |
| Dividend | DIVIDEND, DISTRIBUTION | Positive |
| Interest | INTEREST, JUROS, YIELD | Positive |

### Investment Detection Logic

The system checks:
1. **Keywords**: Matches investment terms
2. **Amount Sign**: Verifies money flow direction
3. **Institution**: Recognizes investment accounts

### Manual Investment Classification

For undetected investments:

1. Find transaction in Transactions sheet
2. Note TransactionID
3. Run in VBA Immediate Window:
   ```vba
   modInvestments.ClassifyAsInvestment "TXN-123", "Application"
   ```

### Investment Reports

Generate detailed investment analysis:
1. Go to Dashboard
2. Click **Investment Report**
3. View summary by type and total net investment

---

## Error Resolution Strategies

### Error Types

#### CLASSIFICATION_ERROR
**Cause**: No category match found
**Solution**:
1. Review description in Errors sheet
2. Determine correct category
3. Enter in SuggestedCategory column
4. System suggests keywords automatically

#### NORMALIZATION_ERROR
**Cause**: Invalid data format
**Solution**:
1. Check source file for corruption
2. Verify date formats
3. Check for special characters
4. Re-import after fixing

#### IMPORT_ERROR
**Cause**: File read/parse failure
**Solution**:
1. Verify file path exists
2. Check file permissions
3. Ensure file is not open in another app
4. Check for file corruption

### Batch Error Resolution

For multiple similar errors:

1. Resolve first error manually
2. System suggests keyword
3. Add keyword to Categories sheet
4. Click **Reprocess Data**
5. Similar transactions auto-resolved

### Error Prevention

**Before Importing**:
- Validate file formats
- Check for missing columns
- Remove summary rows
- Ensure consistent date formats

---

## Custom Workflows

### Workflow 1: Monthly Reconciliation

```
1. Export bank statements (CSV)
2. Place in configured import folder
3. Dashboard → Import Transactions
4. Review Errors sheet
5. Resolve unclassified items
6. Dashboard → Export Transactions
7. Analyze in Excel/Sheets
```

### Workflow 2: Multi-Account Consolidation

```
1. Configure multiple ImportPaths:
   - ImportPath1: Checking account
   - ImportPath2: Savings account
   - ImportPath3: Credit card
2. Import all at once
3. System detects inter-account transfers
4. Consolidated view in Transactions sheet
```

### Workflow 3: Investment Portfolio Tracking

```
1. Import brokerage statements
2. System detects investment transactions
3. Dashboard → Investment Report
4. Analyze contributions vs redemptions
5. Track dividend income
```

### Workflow 4: Category Analysis

```
1. Import transactions
2. Export Transactions to CSV
3. Open in Excel
4. Create pivot table:
   - Rows: Category, Subcategory
   - Values: Sum of Amount
5. Analyze spending patterns
```

### Workflow 5: Year-End Review

```
1. Import all months (Jan-Dec)
2. Dashboard → Show Statistics
3. Review classification success rate
4. Dashboard → Investment Report
5. Export for tax preparation
```

---

## Best Practices

### Data Management

#### 1. Regular Imports
- Import monthly for best accuracy
- Don't wait to accumulate large files
- Smaller batches process faster

#### 2. Category Maintenance
- Review Categories sheet quarterly
- Add keywords as you discover patterns
- Deactivate unused categories (IsActive = FALSE)

#### 3. Backup Strategy
- Enable AutoBackup in Config
- Keep backups for 12 months
- Store backups securely

#### 4. Error Hygiene
- Resolve errors promptly
- Clear resolved errors monthly
- Document resolution patterns

### Classification Tips

#### For High Accuracy

1. **Start Broad, Then Refine**
   ```
   First pass:  GROCERY → Food & Dining:Groceries
   Refinement:  WHOLE FOODS → Food & Dining:Organic Groceries
   ```

2. **Use Negative Keywords** (Advanced)
   ```
   Category: Transportation:Gas
   Keywords: SHELL,CHEVRON,EXXON
   Exclude: SHELL CONVENIENCE (convenience store, not gas)
   ```

3. **Seasonal Adjustments**
   ```
   Priority changes:
   - Summer: Lower priority for heating utilities
   - Winter: Higher priority for holiday shopping
   ```

### Performance Tips

#### For Large Datasets (>50,000 transactions)

1. **Batch Processing**
   - Process one year at a time
   - Combine results in master file

2. **Category Optimization**
   - Keep active rules under 200
   - Disable low-priority rules
   - Use exact matches when possible

3. **Excel Settings**
   - Increase memory allocation
   - Disable automatic calculation during import
   - Close other applications

### Security & Privacy

#### Sensitive Data Protection

1. **Account Masking**
   - System automatically masks accounts
   - Only last 4 digits shown

2. **Description Privacy**
   - Full descriptions preserved for accuracy
   - Consider manual sanitization for sharing

3. **File Encryption**
   - Use MacOS FileVault
   - Password-protect Excel file
   - Store backups securely

---

## Keyboard Shortcuts

### Excel for Mac

| Shortcut | Action |
|----------|--------|
| Alt + F11 | Open VBA Editor |
| Alt + F8 | Run Macro |
| Cmd + S | Save Workbook |
| Cmd + Shift + N | New Sheet |

### VBA Editor

| Shortcut | Action |
|----------|--------|
| F5 | Run Current Macro |
| Cmd + G | Open Immediate Window |
| Cmd + R | Show Project Explorer |

---

## Formulas & Analysis

### Useful Excel Formulas

#### Monthly Spending by Category
```excel
=SUMIFS(Transactions[Amount], Transactions[Category], "Food & Dining",
        Transactions[Date], ">=1/1/2024", Transactions[Date], "<=1/31/2024")
```

#### Average Transaction Amount
```excel
=AVERAGE(Transactions[Amount])
```

#### Count Transactions by Category
```excel
=COUNTIF(Transactions[Category], "Food & Dining")
```

#### Spending Trend (Month over Month)
```excel
=SUMIFS(Transactions[Amount], Transactions[Date], ">="&DATE(2024,1,1))
```

---

## Troubleshooting Common Scenarios

### Scenario 1: Transfer Detected Incorrectly

**Problem**: Two unrelated transactions linked as transfer

**Solution**:
1. Note TransactionID of either transaction
2. VBA Immediate Window: `modTransfers.UnlinkTransfer "TXN-123"`
3. Transactions now independent

### Scenario 2: Same Merchant, Different Categories

**Problem**: AMAZON used for both shopping and streaming

**Solution**:
1. Use specific keywords:
   - `AMAZON PRIME VIDEO` → Entertainment
   - `AMAZON FRESH` → Groceries
   - `AMAZON` → Shopping (lower priority)

### Scenario 3: Foreign Currency Transactions

**Problem**: Amounts in different currencies not normalized

**Current Limitation**: System uses amounts as-is

**Workaround**:
1. Convert amounts manually before import
2. Or add currency conversion in future enhancement

### Scenario 4: Recurring Subscriptions

**Problem**: Want to track subscription totals

**Solution**:
1. Create specific subcategory: "Subscriptions"
2. Add all subscription keywords
3. Use Excel formula to sum: `=SUMIF(Transactions[Subcategory], "Subscriptions", Transactions[Amount])`

---

## Advanced VBA Customization

### Custom Classification Logic

To add custom classification rules, edit `modClassify.bas`:

```vba
' Add custom logic in ClassifyTransaction function
If InStr(description, "CUSTOM PATTERN") > 0 Then
    result.Category = "Custom Category"
    result.Subcategory = "Custom Sub"
    result.Success = True
End If
```

### Custom Reports

Create custom report button:

```vba
Public Sub MyCustomReport()
    ' Your custom analysis code
    MsgBox "Custom report complete!"
End Sub
```

Add to Dashboard via VBA Editor.

---

## Appendix: Sample Category Structures

### Conservative Budget Categories

```
Income
├── Salary
├── Bonus
└── Other Income

Housing
├── Rent/Mortgage
├── Utilities
├── Maintenance
└── Insurance

Transportation
├── Gas
├── Public Transit
└── Maintenance

Food
├── Groceries
└── Dining Out

Personal
└── Miscellaneous
```

### Detailed Budget Categories

```
Income
├── Salary
├── Bonus
├── Investment Income
└── Side Income

Housing
├── Mortgage
├── Property Tax
├── HOA Fees
├── Home Insurance
├── Utilities - Electric
├── Utilities - Gas
├── Utilities - Water
├── Internet/Cable
└── Home Maintenance

Transportation
├── Auto Payment
├── Auto Insurance
├── Gas & Fuel
├── Parking
├── Tolls
├── Public Transit
├── Ride Share
└── Maintenance & Repairs

Food & Dining
├── Groceries
├── Restaurants
├── Fast Food
├── Coffee Shops
└── Alcohol & Bars

Healthcare
├── Insurance
├── Doctor Visits
├── Prescriptions
├── Dental
└── Vision

Entertainment
├── Streaming Services
├── Movies & Events
├── Hobbies
└── Travel

Personal Care
├── Clothing
├── Hair & Beauty
└── Fitness

Financial
├── Savings
├── Investments
├── Debt Payments
└── Bank Fees
```

---

## Getting Help

### Resources

1. **Documentation**:
   - `README.md` - Overview
   - `INSTALLATION_GUIDE.md` - Setup instructions
   - `TECHNICAL_SPECIFICATION.md` - System details

2. **In-Application**:
   - Dashboard → Show Statistics
   - Dashboard → Error Report
   - Audit sheet (system logs)

3. **Community**:
   - Share best practices
   - Contribute category keywords
   - Report issues

---

**Version**: 1.0
**Last Updated**: January 2026
