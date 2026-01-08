# Sample Data Files for [MY LIFE]

This folder contains sample CSV files for testing all import modules.

## 📁 Files Overview

### Bank Statements
- **itau_bank_statement.csv** - ITAÚ checking account transactions
- **nubank_bank_statement.csv** - NUBANK checking account transactions

**Format**: Date, Description, Value, Balance

### Credit Card Statements
- **itau_card_statement.csv** - ITAÚ credit card with installments
- **nubank_card_statement.csv** - NUBANK credit card with installments

**Format**: PurchaseDate, Category, Description, Installment, Value

### Investment Transactions
- **investments_transactions.csv** - Investment applications, redemptions, dividends

**Format**: Date, Description, Value, Type

### OPUS External Positions
- **opus_positions.csv** - External investment positions (assets and liabilities)

**Format**: Type, Company, InvestmentCost, CapitalCost, UpdatedCost, Currency

### Debts
- **debts.csv** - Loans and debts with interest rates

**Format**: Creditor, InterestRate, AmountPaid, Currency

### Financial Indexes
- **indexes.csv** - CDI, SELIC, IPCA, USD/BRL, FED_FUNDS rates

**Format**: Index, Date, Value

## 🧪 How to Use These Files

### Step 1: Copy to Your Local Machine
Place these files in a folder on your Mac, for example:
```
/Users/yourname/Documents/MyLifeData/
```

### Step 2: Configure Paths in Excel
Open your MyLife.xlsm workbook and go to the **FILES PATHS** sheet.

Update the file paths to point to where you saved the sample files:

```
Source          | FilePath                                        | Active
----------------|------------------------------------------------|--------
ITAU_BANK       | /Users/yourname/Documents/MyLifeData/itau_bank_statement.csv      | TRUE
NUBANK_BANK     | /Users/yourname/Documents/MyLifeData/nubank_bank_statement.csv    | TRUE
ITAU_CARD       | /Users/yourname/Documents/MyLifeData/itau_card_statement.csv      | TRUE
NUBANK_CARD     | /Users/yourname/Documents/MyLifeData/nubank_card_statement.csv    | TRUE
INVESTMENTS     | /Users/yourname/Documents/MyLifeData/investments_transactions.csv | TRUE
OPUS            | /Users/yourname/Documents/MyLifeData/opus_positions.csv           | TRUE
DEBTS           | /Users/yourname/Documents/MyLifeData/debts.csv                    | TRUE
```

### Step 3: Run Import
Go to the **DASHBOARD** sheet and click the **"Import All Data"** button (or run `modMain.MasterImport`).

### Step 4: Review Results
Check the data sheets:
- **BANKS** - Should have 21 transactions
- **CARDS** - Should have 21 transactions
- **INVESTMENTS** - Should have 7 transactions
- **OPUS** - Should have 6 positions
- **DEBTS** - Should have 4 loans

## 📊 What These Samples Demonstrate

### Bank Transactions Correlation
The sample includes **correlated movements**:

**Investment Application:**
- ITAÚ Bank: 2024-01-05, "APLICACAO CDB ITAU", -10,000.00 (withdrawal)
- Investments: 2024-01-05, "APLICACAO CDB ITAU BANK", +10,000.00 (application)

**Investment Redemption:**
- NUBANK Bank: 2024-01-20, "Resgate investimento LCI", +8,000.00 (deposit)
- Investments: 2024-01-20, "RESGATE LCI NUBANK", -8,000.00 (redemption)

These should be automatically linked by the correlation engine.

### Card Installments
Cards include both single-payment and installment transactions:
- MAGAZINE LUIZA NOTEBOOK: 1/10 (first of 10 installments)
- APPLE STORE IPHONE: 1/12 (first of 12 installments)
- ZARA ROUPAS: 1/3 (first of 3 installments)

### Multi-Currency
OPUS and DEBTS include both BRL and USD entries to test multi-currency handling.

### Classification Testing
Transactions include keywords that should match the default categories:
- "RESTAURANTE" → Food & Dining
- "POSTO" → Transportation:Gas
- "FARMACIA" → Healthcare:Pharmacy
- "UBER" → Transportation:Ride Share

## 🔧 Creating Your Own Sample Files

To create your own test data:

1. **Use the same column structure** as shown in these samples
2. **Save as CSV** (comma-separated, UTF-8 encoding)
3. **Use consistent date format**: YYYY-MM-DD or DD/MM/YYYY
4. **Include both positive and negative values**
5. **Test edge cases**: duplicates, missing data, special characters

## ⚠️ Important Notes

### For MacOS Users:
- Use forward slashes in paths: `/Users/...`
- Avoid spaces in filenames when possible
- If using iCloud Drive: `/Users/yourname/Library/Mobile Documents/com~apple~CloudDocs/...`

### Data Privacy:
- These are **fake sample data** for testing only
- When using real data, ensure proper security measures
- Never commit real financial data to git repositories

## 📈 Expected Import Results

After importing these sample files, you should see:

**Summary Statistics:**
- Total Bank Transactions: 21
- Total Card Transactions: 21
- Total Investments: 7
- Total OPUS Positions: 6 (3 assets, 3 liabilities)
- Total Debts: 4
- Total Index Entries: 21

**Correlations to Detect:**
- 2 investment correlations (application + redemption)

**Classifications:**
- ~80% should auto-classify with default categories
- ~20% may need manual mapping

## 🚀 Next Steps

After successfully importing these samples:

1. Review classification accuracy in BANKS and CARDS sheets
2. Check correlations in INVESTMENTS sheet (CorrelationID column)
3. Verify OPUS positions loaded correctly
4. Check DEBTS with interest calculations
5. Review DASHBOARD for consolidated view

---

**Version**: 1.0
**Last Updated**: January 2026
