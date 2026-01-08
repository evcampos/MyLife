# Excel VBA Financial Import & Classification Engine (MacOS)

## Overview
A comprehensive Excel-based financial data import and classification system for MacOS, implemented entirely in VBA.

## Features
- **Multi-source Import**: Bank statements, credit cards, investment accounts
- **Intelligent Classification**: Exact and fuzzy matching with machine-learning-like logic
- **Transfer Detection**: Automatic identification and linking of transfers
- **Investment Tracking**: Applications, redemptions, and liquidations
- **Error Management**: Comprehensive audit trail and error resolution
- **User-Friendly Interface**: Buttons and forms for all operations

## System Requirements
- Microsoft Excel for MacOS (2016 or later)
- Macros enabled
- ~50MB free disk space

## Project Structure
```
MyLife/
├── vba-modules/          # VBA source code modules
│   ├── modConfig.bas     # Configuration management
│   ├── modUtils.bas      # Utility functions
│   ├── modImport.bas     # Import engine
│   ├── modNormalize.bas  # Data normalization
│   ├── modClassify.bas   # Classification engine
│   ├── modTransfers.bas  # Transfer detection
│   ├── modInvestments.bas# Investment detection
│   ├── modErrors.bas     # Error handling
│   └── modUI.bas         # User interface
├── docs/                 # Documentation
├── templates/            # Excel templates
└── samples/             # Sample data files
```

## Installation
1. Open the provided `.xlsm` file
2. Enable macros when prompted
3. Go to the Config sheet and set your file paths
4. Click "Initialize System" button

## Usage
1. **Configure Paths**: Set file paths in Config sheet
2. **Import Data**: Click "Import Transactions" button
3. **Review Results**: Check Transactions sheet for classified data
4. **Resolve Errors**: Review Errors sheet and update categories
5. **Reprocess**: Click "Reprocess" to apply new mappings

## Version
Version 1.0 - January 2026

## Author
Built with Claude Code for comprehensive financial management
