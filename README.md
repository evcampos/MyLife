# [MY LIFE] - Family Office Financial Management System

## Overview

**[MY LIFE]** is a comprehensive Excel-based financial management system designed for entrepreneurs and family offices. It consolidates all financial information from multiple sources (banks, credit cards, investments, external assets, and debts) into a single executive dashboard with automated imports, data normalization, metric calculations, and health monitoring.

## Platform

- **Excel for MacOS**
- **100% VBA automation**
- No external dependencies
- Single `.xlsm` workbook

## Key Features

### 1. **Multi-Source Data Import**
- Banks: ITAÚ, NUBANK, C6, BANCO DO BRASIL
- Credit Cards: Multiple cards with installment tracking
- Investments: Linked to bank movements
- OPUS: External investment positions
- Debts: Personal loans with interest tracking

### 2. **Automatic Investment Correlation**
- Links bank withdrawals → investment applications
- Links investment redemptions → bank deposits
- Ensures balanced, traceable movements

### 3. **Capital Cost Calculations**
- Multi-currency support (BRL/USD)
- Index-based adjustments (CDI, SELIC, IPCA, Fed Funds)
- Historical cumulative factor tracking

### 4. **Smart Transaction Classification**
- Exact keyword matching
- Proximity/fuzzy matching
- Manual mapping with learning
- Persistent classification rules

### 5. **Executive Dashboard**
- Consolidated KPIs (Income, Expenses, Balance)
- Multi-dimensional filtering (Year, Month, Institution, Currency)
- Consolidated views: Cash, Cards, Transactions, Debts
- Visual charts and trend analysis

### 6. **Health Check System**
- Validates all imports
- Checks data integrity
- Verifies correlations
- Reports inconsistencies

## Worksheet Structure

| Sheet | Purpose |
|-------|---------|
| **FILES PATHS** | Central configuration of all import paths |
| **FILES STRUCTURE** | Column structure definitions for validation |
| **BANKS** | Checking account transactions |
| **CARDS** | Credit card transactions with installments |
| **INVESTMENTS** | Investment transactions (correlated with banks) |
| **OPUS** | External investment positions |
| **DEBTS** | Loans and debt tracking |
| **INDEXES** | Financial indexes with cumulative factors |
| **CATEGORIES** | Transaction classification rules |
| **DASHBOARD** | Executive KPI and consolidated views |
| **HEALTH_CHECK** | Validation results |

## Quick Start

1. **Open** `MyLife.xlsm` in Excel for Mac
2. **Enable Macros** when prompted
3. **Configure Paths**: Go to `FILES PATHS` sheet and set your import file locations
4. **Review Categories**: Customize classification rules in `CATEGORIES` sheet
5. **Import Data**: Use buttons on `DASHBOARD` to run imports
6. **Review Dashboard**: Analyze consolidated financial view

## Named Ranges (For Advanced Users)

The system exposes key metrics as named ranges for formulas:

- `Total_Income`
- `Total_Expenses`
- `Balance`
- `Monthly_Trend`
- `Category_Breakdown`

## Documentation

- `docs/ARCHITECTURE.md` - System architecture and module design
- `docs/USER_GUIDE.md` - Detailed usage instructions
- `docs/VBA_MODULES.md` - VBA module documentation
- `docs/SETUP_GUIDE.md` - Installation and configuration

## Version

**Version 1.0** - January 2026

## Status

🚧 **Production-Grade Financial System**

Built for real-world family office financial management.
