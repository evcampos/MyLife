Attribute VB_Name = "modConfig"
' ========================================================================
' Module: modConfig
' Purpose: Configuration Management for [MY LIFE]
' Description: Handles all system configuration including file paths,
'              sheet names, and settings. No hardcoded paths allowed.
' Platform: Excel for MacOS
' Author: [MY LIFE] Financial System v1.0
' ========================================================================

Option Explicit

' ========================================================================
' SHEET NAME CONSTANTS
' ========================================================================
Public Const SHEET_FILES_PATHS As String = "FILES PATHS"
Public Const SHEET_FILES_STRUCTURE As String = "FILES STRUCTURE"
Public Const SHEET_BANKS As String = "BANKS"
Public Const SHEET_CARDS As String = "CARDS"
Public Const SHEET_INVESTMENTS As String = "INVESTMENTS"
Public Const SHEET_OPUS As String = "OPUS"
Public Const SHEET_DEBTS As String = "DEBTS"
Public Const SHEET_INDEXES As String = "INDEXES"
Public Const SHEET_CATEGORIES As String = "CATEGORIES"
Public Const SHEET_DASHBOARD As String = "DASHBOARD"
Public Const SHEET_HEALTH_CHECK As String = "HEALTH_CHECK"

' ========================================================================
' NAMED RANGE CONSTANTS
' ========================================================================
Public Const NR_TOTAL_INCOME As String = "Total_Income"
Public Const NR_TOTAL_EXPENSES As String = "Total_Expenses"
Public Const NR_BALANCE As String = "Balance"
Public Const NR_MONTHLY_TREND As String = "Monthly_Trend"
Public Const NR_CATEGORY_BREAKDOWN As String = "Category_Breakdown"

' ========================================================================
' Initialize System - Create All Required Sheets
' ========================================================================
Public Sub InitializeSystem()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    MsgBox "Initializing [MY LIFE] system..." & vbCrLf & vbCrLf & _
           "This will create all required sheets and structures.", _
           vbInformation, "[MY LIFE] Setup"

    ' Create all required sheets
    Call CreateFilesPathsSheet
    Call CreateFilesStructureSheet
    Call CreateBanksSheet
    Call CreateCardsSheet
    Call CreateInvestmentsSheet
    Call CreateOPUSSheet
    Call CreateDebtsSheet
    Call CreateIndexesSheet
    Call CreateCategoriesSheet
    Call CreateDashboardSheet
    Call CreateHealthCheckSheet

    ' Arrange sheets in logical order
    Call ArrangeSheets

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "[MY LIFE] system initialized successfully!" & vbCrLf & vbCrLf & _
           "Next steps:" & vbCrLf & _
           "1. Configure file paths in FILES PATHS sheet" & vbCrLf & _
           "2. Review CATEGORIES sheet" & vbCrLf & _
           "3. Go to DASHBOARD to start importing data", _
           vbInformation, "Initialization Complete"

    ' Navigate to Dashboard
    ThisWorkbook.Worksheets(SHEET_DASHBOARD).Activate

    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error initializing system: " & Err.Description, vbCritical, "Initialization Error"
End Sub

' ========================================================================
' Create FILES PATHS Sheet
' ========================================================================
Private Sub CreateFilesPathsSheet()
    Dim ws As Worksheet

    If Not SheetExists(SHEET_FILES_PATHS) Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = SHEET_FILES_PATHS
    Else
        Set ws = ThisWorkbook.Worksheets(SHEET_FILES_PATHS)
        ws.Cells.Clear
    End If

    With ws
        ' Headers
        .Cells(1, 1).Value = "Source"
        .Cells(1, 2).Value = "FilePath"
        .Cells(1, 3).Value = "Active"
        .Cells(1, 4).Value = "Description"

        ' Format headers
        .Range("A1:D1").Font.Bold = True
        .Range("A1:D1").Interior.Color = RGB(68, 114, 196)
        .Range("A1:D1").Font.Color = RGB(255, 255, 255)

        ' Add example configurations
        .Cells(2, 1).Value = "ITAU_BANK"
        .Cells(2, 2).Value = "/Users/yourname/Documents/Finances/itau_extrato.csv"
        .Cells(2, 3).Value = "TRUE"
        .Cells(2, 4).Value = "ITAÚ checking account"

        .Cells(3, 1).Value = "NUBANK_BANK"
        .Cells(3, 2).Value = "/Users/yourname/Documents/Finances/nubank_extrato.csv"
        .Cells(3, 3).Value = "TRUE"
        .Cells(3, 4).Value = "NUBANK checking account"

        .Cells(4, 1).Value = "C6_BANK"
        .Cells(4, 2).Value = "/Users/yourname/Documents/Finances/c6_extrato.csv"
        .Cells(4, 3).Value = "FALSE"
        .Cells(4, 4).Value = "C6 checking account"

        .Cells(5, 1).Value = "BB_BANK"
        .Cells(5, 2).Value = "/Users/yourname/Documents/Finances/bb_extrato.csv"
        .Cells(5, 3).Value = "FALSE"
        .Cells(5, 4).Value = "Banco do Brasil checking account"

        .Cells(7, 1).Value = "ITAU_CARD"
        .Cells(7, 2).Value = "/Users/yourname/Documents/Finances/itau_fatura.csv"
        .Cells(7, 3).Value = "TRUE"
        .Cells(7, 4).Value = "ITAÚ credit card"

        .Cells(8, 1).Value = "NUBANK_CARD"
        .Cells(8, 2).Value = "/Users/yourname/Documents/Finances/nubank_fatura.csv"
        .Cells(8, 3).Value = "TRUE"
        .Cells(8, 4).Value = "NUBANK credit card"

        .Cells(9, 1).Value = "C6_CARD"
        .Cells(9, 2).Value = "/Users/yourname/Documents/Finances/c6_fatura.csv"
        .Cells(9, 3).Value = "FALSE"
        .Cells(9, 4).Value = "C6 credit card"

        .Cells(11, 1).Value = "INVESTMENTS"
        .Cells(11, 2).Value = "/Users/yourname/Documents/Finances/investments.csv"
        .Cells(11, 3).Value = "TRUE"
        .Cells(11, 4).Value = "Investment transactions"

        .Cells(13, 1).Value = "OPUS"
        .Cells(13, 2).Value = "/Users/yourname/Documents/Finances/opus.csv"
        .Cells(13, 3).Value = "TRUE"
        .Cells(13, 4).Value = "OPUS external investments"

        .Cells(15, 1).Value = "DEBTS"
        .Cells(15, 2).Value = "/Users/yourname/Documents/Finances/debts.csv"
        .Cells(15, 3).Value = "TRUE"
        .Cells(15, 4).Value = "Debts and loans"

        ' Auto-fit columns
        .Columns("A:D").AutoFit
        .Columns("B:B").ColumnWidth = 60
    End With
End Sub

' ========================================================================
' Create FILES STRUCTURE Sheet
' ========================================================================
Private Sub CreateFilesStructureSheet()
    Dim ws As Worksheet

    If Not SheetExists(SHEET_FILES_STRUCTURE) Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = SHEET_FILES_STRUCTURE
    Else
        Set ws = ThisWorkbook.Worksheets(SHEET_FILES_STRUCTURE)
        ws.Cells.Clear
    End If

    With ws
        ' Headers
        .Cells(1, 1).Value = "Source"
        .Cells(1, 2).Value = "ExpectedColumns"
        .Cells(1, 3).Value = "Description"

        ' Format headers
        .Range("A1:C1").Font.Bold = True
        .Range("A1:C1").Interior.Color = RGB(68, 114, 196)
        .Range("A1:C1").Font.Color = RGB(255, 255, 255)

        ' Define expected structures
        .Cells(2, 1).Value = "BANK"
        .Cells(2, 2).Value = "Date,Description,Value,Balance"
        .Cells(2, 3).Value = "Standard bank statement format"

        .Cells(3, 1).Value = "CARD"
        .Cells(3, 2).Value = "PurchaseDate,Category,Description,Installment,Value"
        .Cells(3, 3).Value = "Credit card statement with installments"

        .Cells(4, 1).Value = "INVESTMENT"
        .Cells(4, 2).Value = "Date,Description,Value,Type"
        .Cells(4, 3).Value = "Investment transactions"

        .Cells(5, 1).Value = "OPUS"
        .Cells(5, 2).Value = "Type,Company,InvestmentCost,CapitalCost,UpdatedCost,Currency"
        .Cells(5, 3).Value = "OPUS external investments"

        .Cells(6, 1).Value = "DEBTS"
        .Cells(6, 2).Value = "Creditor,InterestRate,AmountPaid,Currency"
        .Cells(6, 3).Value = "Debt tracking"

        ' Auto-fit columns
        .Columns("A:C").AutoFit
        .Columns("B:B").ColumnWidth = 60
    End With
End Sub

' ========================================================================
' Create BANKS Sheet
' ========================================================================
Private Sub CreateBanksSheet()
    Dim ws As Worksheet

    If Not SheetExists(SHEET_BANKS) Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = SHEET_BANKS
    Else
        Set ws = ThisWorkbook.Worksheets(SHEET_BANKS)
    End If

    ' Only create headers if empty
    If ws.Cells(1, 1).Value = "" Then
        With ws
            .Cells(1, 1).Value = "Bank"
            .Cells(1, 2).Value = "Date"
            .Cells(1, 3).Value = "Description"
            .Cells(1, 4).Value = "Value"
            .Cells(1, 5).Value = "Category"
            .Cells(1, 6).Value = "Subcategory"
            .Cells(1, 7).Value = "CorrelationID"
            .Cells(1, 8).Value = "ImportDate"

            ' Format headers
            .Range("A1:H1").Font.Bold = True
            .Range("A1:H1").Interior.Color = RGB(68, 114, 196)
            .Range("A1:H1").Font.Color = RGB(255, 255, 255)

            ' Auto-fit
            .Columns("A:H").AutoFit
        End With
    End If
End Sub

' ========================================================================
' Create CARDS Sheet
' ========================================================================
Private Sub CreateCardsSheet()
    Dim ws As Worksheet

    If Not SheetExists(SHEET_CARDS) Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = SHEET_CARDS
    Else
        Set ws = ThisWorkbook.Worksheets(SHEET_CARDS)
    End If

    If ws.Cells(1, 1).Value = "" Then
        With ws
            .Cells(1, 1).Value = "Bank"
            .Cells(1, 2).Value = "CardNumber"
            .Cells(1, 3).Value = "PurchaseDate"
            .Cells(1, 4).Value = "Category"
            .Cells(1, 5).Value = "Description"
            .Cells(1, 6).Value = "Installment"
            .Cells(1, 7).Value = "Value"
            .Cells(1, 8).Value = "ImportDate"

            ' Format headers
            .Range("A1:H1").Font.Bold = True
            .Range("A1:H1").Interior.Color = RGB(68, 114, 196)
            .Range("A1:H1").Font.Color = RGB(255, 255, 255)

            .Columns("A:H").AutoFit
        End With
    End If
End Sub

' ========================================================================
' Create INVESTMENTS Sheet
' ========================================================================
Private Sub CreateInvestmentsSheet()
    Dim ws As Worksheet

    If Not SheetExists(SHEET_INVESTMENTS) Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = SHEET_INVESTMENTS
    Else
        Set ws = ThisWorkbook.Worksheets(SHEET_INVESTMENTS)
    End If

    If ws.Cells(1, 1).Value = "" Then
        With ws
            .Cells(1, 1).Value = "Bank"
            .Cells(1, 2).Value = "Date"
            .Cells(1, 3).Value = "Description"
            .Cells(1, 4).Value = "Value"
            .Cells(1, 5).Value = "Type"
            .Cells(1, 6).Value = "Category"
            .Cells(1, 7).Value = "CorrelationID"
            .Cells(1, 8).Value = "ImportDate"

            ' Format headers
            .Range("A1:H1").Font.Bold = True
            .Range("A1:H1").Interior.Color = RGB(68, 114, 196)
            .Range("A1:H1").Font.Color = RGB(255, 255, 255)

            .Columns("A:H").AutoFit
        End With
    End If
End Sub

' ========================================================================
' Create OPUS Sheet
' ========================================================================
Private Sub CreateOPUSSheet()
    Dim ws As Worksheet

    If Not SheetExists(SHEET_OPUS) Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = SHEET_OPUS
    Else
        Set ws = ThisWorkbook.Worksheets(SHEET_OPUS)
    End If

    If ws.Cells(1, 1).Value = "" Then
        With ws
            .Cells(1, 1).Value = "Type"
            .Cells(1, 2).Value = "Company"
            .Cells(1, 3).Value = "InvestmentCost"
            .Cells(1, 4).Value = "CapitalCost"
            .Cells(1, 5).Value = "UpdatedCost"
            .Cells(1, 6).Value = "Currency"
            .Cells(1, 7).Value = "LastUpdate"

            ' Format headers
            .Range("A1:G1").Font.Bold = True
            .Range("A1:G1").Interior.Color = RGB(68, 114, 196)
            .Range("A1:G1").Font.Color = RGB(255, 255, 255)

            .Columns("A:G").AutoFit
        End With
    End If
End Sub

' ========================================================================
' Create DEBTS Sheet
' ========================================================================
Private Sub CreateDebtsSheet()
    Dim ws As Worksheet

    If Not SheetExists(SHEET_DEBTS) Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = SHEET_DEBTS
    Else
        Set ws = ThisWorkbook.Worksheets(SHEET_DEBTS)
    End If

    If ws.Cells(1, 1).Value = "" Then
        With ws
            .Cells(1, 1).Value = "Creditor"
            .Cells(1, 2).Value = "InterestRate"
            .Cells(1, 3).Value = "AmountPaid"
            .Cells(1, 4).Value = "UpdatedAmount"
            .Cells(1, 5).Value = "Currency"
            .Cells(1, 6).Value = "LastUpdate"

            ' Format headers
            .Range("A1:F1").Font.Bold = True
            .Range("A1:F1").Interior.Color = RGB(68, 114, 196)
            .Range("A1:F1").Font.Color = RGB(255, 255, 255)

            .Columns("A:F").AutoFit
        End With
    End If
End Sub

' ========================================================================
' Create INDEXES Sheet
' ========================================================================
Private Sub CreateIndexesSheet()
    Dim ws As Worksheet

    If Not SheetExists(SHEET_INDEXES) Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = SHEET_INDEXES
    Else
        Set ws = ThisWorkbook.Worksheets(SHEET_INDEXES)
    End If

    If ws.Cells(1, 1).Value = "" Then
        With ws
            .Cells(1, 1).Value = "Index"
            .Cells(1, 2).Value = "Date"
            .Cells(1, 3).Value = "Value"
            .Cells(1, 4).Value = "CumulativeFactor"

            ' Format headers
            .Range("A1:D1").Font.Bold = True
            .Range("A1:D1").Interior.Color = RGB(68, 114, 196)
            .Range("A1:D1").Font.Color = RGB(255, 255, 255)

            .Columns("A:D").AutoFit
        End With
    End If
End Sub

' ========================================================================
' Create CATEGORIES Sheet
' ========================================================================
Private Sub CreateCategoriesSheet()
    Dim ws As Worksheet

    If Not SheetExists(SHEET_CATEGORIES) Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = SHEET_CATEGORIES
    Else
        Set ws = ThisWorkbook.Worksheets(SHEET_CATEGORIES)
    End If

    If ws.Cells(1, 1).Value = "" Then
        With ws
            .Cells(1, 1).Value = "Category"
            .Cells(1, 2).Value = "Subcategory"
            .Cells(1, 3).Value = "Keywords"
            .Cells(1, 4).Value = "Priority"

            ' Format headers
            .Range("A1:D1").Font.Bold = True
            .Range("A1:D1").Interior.Color = RGB(68, 114, 196)
            .Range("A1:D1").Font.Color = RGB(255, 255, 255)

            ' Add sample categories
            Call AddSampleCategories(ws)

            .Columns("A:D").AutoFit
            .Columns("C:C").ColumnWidth = 50
        End With
    End If
End Sub

' ========================================================================
' Add Sample Categories
' ========================================================================
Private Sub AddSampleCategories(ws As Worksheet)
    Dim row As Long
    row = 2

    ' Income
    ws.Cells(row, 1).Value = "Income": ws.Cells(row, 2).Value = "Salary": ws.Cells(row, 3).Value = "SALARIO,SALARY,VENCIMENTO,PAYROLL": ws.Cells(row, 4).Value = 1: row = row + 1
    ws.Cells(row, 1).Value = "Income": ws.Cells(row, 2).Value = "Investment": ws.Cells(row, 3).Value = "DIVIDEND,RENDIMENTO,YIELD": ws.Cells(row, 4).Value = 1: row = row + 1

    ' Food
    ws.Cells(row, 1).Value = "Food & Dining": ws.Cells(row, 2).Value = "Restaurants": ws.Cells(row, 3).Value = "RESTAURANTE,RESTAURANT,BAR,LANCHONETE": ws.Cells(row, 4).Value = 10: row = row + 1
    ws.Cells(row, 1).Value = "Food & Dining": ws.Cells(row, 2).Value = "Groceries": ws.Cells(row, 3).Value = "MERCADO,SUPERMERCADO,SUPERMARKET": ws.Cells(row, 4).Value = 10: row = row + 1

    ' Transportation
    ws.Cells(row, 1).Value = "Transportation": ws.Cells(row, 2).Value = "Gas": ws.Cells(row, 3).Value = "POSTO,GAS,SHELL,IPIRANGA,PETROBRAS": ws.Cells(row, 4).Value = 10: row = row + 1
    ws.Cells(row, 1).Value = "Transportation": ws.Cells(row, 2).Value = "Uber/Taxi": ws.Cells(row, 3).Value = "UBER,99,TAXI,CABIFY": ws.Cells(row, 4).Value = 10: row = row + 1

    ' Utilities
    ws.Cells(row, 1).Value = "Bills & Utilities": ws.Cells(row, 2).Value = "Internet": ws.Cells(row, 3).Value = "INTERNET,CLARO,VIVO,TIM,OI": ws.Cells(row, 4).Value = 10: row = row + 1
    ws.Cells(row, 1).Value = "Bills & Utilities": ws.Cells(row, 2).Value = "Electric": ws.Cells(row, 3).Value = "ELETRICIDADE,ENERGIA,LIGHT,CPFL": ws.Cells(row, 4).Value = 10: row = row + 1

    ' Healthcare
    ws.Cells(row, 1).Value = "Healthcare": ws.Cells(row, 2).Value = "Pharmacy": ws.Cells(row, 3).Value = "FARMACIA,PHARMACY,DROGARIA": ws.Cells(row, 4).Value = 10: row = row + 1

    ' Investment
    ws.Cells(row, 1).Value = "Investment": ws.Cells(row, 2).Value = "Application": ws.Cells(row, 3).Value = "APLICACAO,APPLICATION,INVESTIMENTO": ws.Cells(row, 4).Value = 5: row = row + 1
    ws.Cells(row, 1).Value = "Investment": ws.Cells(row, 2).Value = "Redemption": ws.Cells(row, 3).Value = "RESGATE,REDEMPTION,LIQUIDACAO": ws.Cells(row, 4).Value = 5: row = row + 1
End Sub

' ========================================================================
' Create DASHBOARD Sheet (Placeholder)
' ========================================================================
Private Sub CreateDashboardSheet()
    Dim ws As Worksheet

    If Not SheetExists(SHEET_DASHBOARD) Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = SHEET_DASHBOARD
    Else
        Set ws = ThisWorkbook.Worksheets(SHEET_DASHBOARD)
    End If

    ' Dashboard will be built by modDashboard module
End Sub

' ========================================================================
' Create HEALTH_CHECK Sheet
' ========================================================================
Private Sub CreateHealthCheckSheet()
    Dim ws As Worksheet

    If Not SheetExists(SHEET_HEALTH_CHECK) Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = SHEET_HEALTH_CHECK
    Else
        Set ws = ThisWorkbook.Worksheets(SHEET_HEALTH_CHECK)
    End If

    ' Health check will be populated by modHealthCheck module
End Sub

' ========================================================================
' Arrange Sheets in Logical Order
' ========================================================================
Private Sub ArrangeSheets()
    On Error Resume Next

    Dim sheetOrder() As String
    Dim i As Integer

    sheetOrder = Split("DASHBOARD,FILES PATHS,FILES STRUCTURE,BANKS,CARDS,INVESTMENTS,OPUS,DEBTS,INDEXES,CATEGORIES,HEALTH_CHECK", ",")

    For i = LBound(sheetOrder) To UBound(sheetOrder)
        If SheetExists(sheetOrder(i)) Then
            ThisWorkbook.Worksheets(sheetOrder(i)).Move Before:=ThisWorkbook.Worksheets(1)
        End If
    Next i
End Sub

' ========================================================================
' Check if Sheet Exists
' ========================================================================
Public Function SheetExists(sheetName As String) As Boolean
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)

    SheetExists = Not ws Is Nothing
End Function

' ========================================================================
' Get File Path for Source
' ========================================================================
Public Function GetFilePath(sourceName As String) As String
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_FILES_PATHS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 1).Value)) = UCase(Trim(sourceName)) Then
            If UCase(Trim(ws.Cells(i, 3).Value)) = "TRUE" Then
                GetFilePath = ws.Cells(i, 2).Value
                Exit Function
            End If
        End If
    Next i

    GetFilePath = ""
End Function

' ========================================================================
' Get All Active File Paths
' ========================================================================
Public Function GetAllActivePaths() As Collection
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim paths As Collection

    Set paths = New Collection
    Set ws = ThisWorkbook.Worksheets(SHEET_FILES_PATHS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 3).Value)) = "TRUE" Then
            paths.Add Array(ws.Cells(i, 1).Value, ws.Cells(i, 2).Value)
        End If
    Next i

    Set GetAllActivePaths = paths
End Function
