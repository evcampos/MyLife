# Technical Specification - Financial Import & Classification System

## 1. System Architecture

### 1.1 Module Structure
The system is organized into 9 primary VBA modules:

| Module | Purpose | Dependencies |
|--------|---------|--------------|
| modConfig | Configuration management, file paths | None |
| modUtils | Helper functions, string manipulation | None |
| modImport | File parsing and import | modConfig, modUtils |
| modNormalize | Data standardization | modUtils |
| modClassify | Category classification | modUtils, modConfig |
| modTransfers | Transfer detection | modUtils |
| modInvestments | Investment detection | modUtils |
| modErrors | Error logging and management | modConfig |
| modUI | User interface handlers | All modules |

### 1.2 Excel Workbook Structure

| Sheet Name | Purpose | User Editable |
|------------|---------|---------------|
| Dashboard | Control panel with buttons | View only |
| Config | File paths and settings | Yes |
| Categories | Category mapping table | Yes |
| Transactions | Normalized transaction data | View only |
| Errors | Unresolved classification issues | Yes (for manual mapping) |
| RawData | Imported raw data (hidden) | No |
| Audit | Processing log | View only |

## 2. Data Flow

```
Input Files → Import Engine → Raw Data → Normalization →
Transfer Detection → Investment Detection → Classification →
Transactions / Errors
```

## 3. Configuration Schema

### 3.1 Config Sheet Structure
| Column | Type | Description |
|--------|------|-------------|
| ConfigKey | Text | Setting identifier |
| ConfigValue | Text | Setting value |
| Description | Text | Human-readable description |
| Category | Text | Config grouping |

### 3.2 Standard Config Keys
- `ImportPath1`, `ImportPath2`, etc. - File paths
- `DateFormat` - Expected date format
- `DefaultCurrency` - Base currency
- `FuzzyThreshold` - Matching sensitivity (0-100)
- `AutoBackup` - Enable automatic backups

## 4. Import Engine Specification

### 4.1 Supported Input Formats
- CSV (comma, semicolon, tab-delimited)
- Excel (.xlsx, .xls)
- Fixed-width text files

### 4.2 Required Columns (flexible mapping)
The import engine auto-detects columns containing:
- Date (Date, Data, Fecha, etc.)
- Description (Description, Descripción, Histórico, etc.)
- Amount/Debit/Credit (Valor, Amount, Débito, Crédito, etc.)
- Account/Card (Conta, Tarjeta, Account, etc.)

### 4.3 Import Process
1. Read file format
2. Detect column mappings
3. Parse all rows
4. Store in RawData sheet
5. Log import statistics

## 5. Normalization Engine

### 5.1 Normalization Rules
- **Dates**: Convert to Excel date format
- **Amounts**: Standardize sign (debit=negative, credit=positive)
- **Descriptions**: Trim, uppercase for matching
- **Institutions**: Extract from file name or config
- **Accounts**: Normalize account identifiers

### 5.2 Duplicate Detection
- Hash: Date + Amount + Description (first 30 chars)
- Skip exact duplicates
- Flag near-duplicates for review

## 6. Classification Engine

### 6.1 Categories Table Structure
| Column | Type | Description |
|--------|------|-------------|
| Category | Text | Primary category |
| Subcategory | Text | Secondary classification |
| Keywords | Text | Comma-separated search terms |
| Priority | Integer | Matching priority (1=highest) |
| IsActive | Boolean | Enable/disable rule |

### 6.2 Classification Algorithm

#### Phase 1: Exact Matching
- Search for exact keyword in description
- Case-insensitive
- Whole-word matching preferred
- Highest priority rules first

#### Phase 2: Fuzzy Matching
- Calculate similarity score for each keyword
- Use Levenshtein-like distance
- Threshold configurable (default 80%)
- Return best match above threshold

#### Phase 3: Learning
- Track manual mappings
- Suggest new keywords
- Auto-update Categories table (optional)

### 6.3 Matching Algorithms

**Exact Match**:
```
IF INSTR(UCASE(Description), UCASE(Keyword)) > 0 THEN Match
```

**Fuzzy Match**:
```
Score = (MatchingChars / MaxLength) * 100
IF Score >= Threshold THEN Match
```

## 7. Transfer Detection

### 7.1 Detection Criteria
Transactions are linked as transfers when:
- Same date (±1 day tolerance)
- Opposite amounts (within 0.1% tolerance)
- Different accounts
- Similar descriptions (optional)

### 7.2 Transfer Classification
- Category: "Transfer"
- Subcategory: "Internal Transfer"
- Generate unique TransferID
- Link both sides with TransferID

## 8. Investment Detection

### 8.1 Investment Keywords
- Application: "APLICACAO", "INVESTMENT", "BUY", "PURCHASE"
- Redemption: "RESGATE", "REDEMPTION", "SELL", "WITHDRAWAL"
- Liquidation: "LIQUIDACAO", "CLOSURE", "LIQUIDATION"

### 8.2 Investment Processing
1. Detect investment keywords
2. Classify as Application/Redemption/Liquidation
3. Generate derived records if needed
4. Assign Category: "Investment"
5. Assign Subcategory based on type

## 9. Error Handling

### 9.1 Error Types
- `IMPORT_ERROR`: File read/parse failure
- `NORMALIZATION_ERROR`: Invalid data format
- `CLASSIFICATION_ERROR`: No category match
- `VALIDATION_ERROR`: Business rule violation

### 9.2 Error Sheet Structure
| Column | Description |
|--------|-------------|
| ErrorID | Unique identifier |
| ErrorType | Type from above |
| TransactionDate | Original transaction date |
| Description | Original description |
| Amount | Transaction amount |
| IssueDescription | What went wrong |
| Resolution | User's fix (manual mapping) |
| Status | OPEN / RESOLVED |

## 10. User Interface

### 10.1 Dashboard Buttons
- **Initialize System**: Setup sheets and validation
- **Import Transactions**: Run full import process
- **Reprocess Data**: Re-classify existing data
- **Update Categories**: Reload category mappings
- **Clear Errors**: Remove resolved errors
- **Export Report**: Generate summary report
- **Backup Data**: Save current state

### 10.2 Form Dialogs (optional)
- Path Configuration Form
- Manual Category Mapping Form
- Import Settings Form

## 11. Performance Requirements

- Import 10,000 transactions in < 30 seconds
- Classification of 10,000 records in < 60 seconds
- Memory usage < 500MB
- Support files up to 100MB

## 12. Data Validation Rules

### 12.1 Transactions Table
- Date: Required, must be valid date
- Amount: Required, must be numeric
- Description: Required, max 500 chars
- Category: Required after classification
- Account: Optional

### 12.2 Categories Table
- Category: Required, max 50 chars
- Keywords: Required, at least one
- Priority: Default = 100

## 13. MacOS Compatibility Notes

### 13.1 File Path Handling
- Use MacOS-style paths: `/Users/username/...`
- Support iCloud Drive paths
- Use `Dir()` function for file existence checks
- Handle spaces in paths with quotes

### 13.2 Known Limitations
- No Windows API calls
- Use native VBA functions only
- Excel for Mac specific features noted in code

## 14. Security & Privacy

- No external connections required
- All data processing local
- No telemetry or logging to external services
- User data never leaves the Excel file

## 15. Extensibility

### 15.1 Adding New File Formats
Extend `modImport` with new parser function

### 15.2 Adding New Classification Rules
Add rows to Categories table, no code changes

### 15.3 Adding New Detection Logic
Extend `modTransfers` or `modInvestments` modules

## 16. Testing Strategy

- Unit tests for each module function
- Integration tests for full pipeline
- Sample data files for regression testing
- MacOS Excel compatibility testing

## Version History

- v1.0 (January 2026): Initial release
