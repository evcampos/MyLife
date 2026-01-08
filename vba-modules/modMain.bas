Attribute VB_Name = "modMain"
' ========================================================================
' Module: modMain
' Purpose: Main Orchestration and Button Handlers
' Description: Master control for [MY LIFE] system
' Platform: Excel for MacOS
' Author: [MY LIFE] Financial System v1.0
' ========================================================================

Option Explicit

' ========================================================================
' MASTER IMPORT - Import All Data Sources
' ========================================================================
Public Sub MasterImport()
    On Error GoTo ErrorHandler

    Dim startTime As Double
    Dim banksCount As Long
    Dim cardsCount As Long
    Dim investCount As Long

    startTime = Timer

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    MsgBox "Starting [MY LIFE] data import..." & vbCrLf & vbCrLf & _
           "This may take a few moments.", vbInformation, "[MY LIFE] Import"

    ' Import Banks
    modUtils.ShowProgress "Importing bank transactions..."
    banksCount = modImportBanks.ImportAllBanks()

    ' TODO: Import Cards (modImportCards)
    ' cardsCount = modImportCards.ImportAllCards()

    ' TODO: Import Investments (modImportInvestments)
    ' investCount = modImportInvestments.ImportAllInvestments()

    ' TODO: Import OPUS (modImportOPUS)
    ' Call modImportOPUS.ImportOPUS

    ' TODO: Import Debts (modImportDebts)
    ' Call modImportDebts.ImportDebts

    ' TODO: Classification (modClassification)
    ' Call modClassification.ClassifyAllTransactions

    ' TODO: Correlation (modCorrelation)
    ' Call modCorrelation.CorrelateInvestments

    ' TODO: Update Dashboard (modDashboard)
    ' Call modDashboard.UpdateDashboard

    modUtils.ClearProgress

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    Dim elapsed As Double
    elapsed = Timer - startTime

    MsgBox "Import complete!" & vbCrLf & vbCrLf & _
           "Banks: " & banksCount & " transactions" & vbCrLf & _
           "Cards: " & cardsCount & " transactions" & vbCrLf & _
           "Investments: " & investCount & " transactions" & vbCrLf & vbCrLf & _
           "Time: " & Format(elapsed, "0.0") & " seconds", _
           vbInformation, "[MY LIFE] Import Complete"

    ' Navigate to Dashboard
    ThisWorkbook.Worksheets(modConfig.SHEET_DASHBOARD).Activate

    Exit Sub

ErrorHandler:
    modUtils.ClearProgress
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error during import: " & Err.Description, vbCritical, "[MY LIFE] Import Error"
End Sub

' ========================================================================
' Update Dashboard
' ========================================================================
Public Sub UpdateDashboard()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    ' TODO: modDashboard.UpdateDashboard

    Application.ScreenUpdating = True

    MsgBox "Dashboard updated successfully!", vbInformation, "[MY LIFE]"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error updating dashboard: " & Err.Description, vbCritical, "[MY LIFE] Error"
End Sub

' ========================================================================
' Run Health Check
' ========================================================================
Public Sub RunHealthCheck()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    ' TODO: modHealthCheck.RunHealthCheck

    Application.ScreenUpdating = True

    MsgBox "Health check complete!" & vbCrLf & _
           "Review results in HEALTH_CHECK sheet.", _
           vbInformation, "[MY LIFE] Health Check"

    ' Navigate to Health Check sheet
    ThisWorkbook.Worksheets(modConfig.SHEET_HEALTH_CHECK).Activate

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error during health check: " & Err.Description, vbCritical, "[MY LIFE] Error"
End Sub

' ========================================================================
' Clear All Data
' ========================================================================
Public Sub ClearAllData()
    On Error GoTo ErrorHandler

    Dim response As VbMsgBoxResult

    response = MsgBox("WARNING: This will delete all imported data!" & vbCrLf & vbCrLf & _
                     "Are you sure?", vbYesNo + vbCritical, "[MY LIFE] Clear Data")

    If response = vbNo Then Exit Sub

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    ' Clear data from all data sheets
    modUtils.ClearSheetData ThisWorkbook.Worksheets(modConfig.SHEET_BANKS)
    modUtils.ClearSheetData ThisWorkbook.Worksheets(modConfig.SHEET_CARDS)
    modUtils.ClearSheetData ThisWorkbook.Worksheets(modConfig.SHEET_INVESTMENTS)
    modUtils.ClearSheetData ThisWorkbook.Worksheets(modConfig.SHEET_OPUS)
    modUtils.ClearSheetData ThisWorkbook.Worksheets(modConfig.SHEET_DEBTS)
    modUtils.ClearSheetData ThisWorkbook.Worksheets(modConfig.SHEET_HEALTH_CHECK)

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "All data cleared.", vbInformation, "[MY LIFE]"

    Exit Sub

ErrorHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Error clearing data: " & Err.Description, vbCritical, "[MY LIFE] Error"
End Sub

' ========================================================================
' Show System Information
' ========================================================================
Public Sub ShowSystemInfo()
    On Error Resume Next

    Dim info As String
    Dim banksCount As Long
    Dim cardsCount As Long
    Dim investCount As Long

    banksCount = modUtils.GetLastRow(ThisWorkbook.Worksheets(modConfig.SHEET_BANKS), 1) - 1
    cardsCount = modUtils.GetLastRow(ThisWorkbook.Worksheets(modConfig.SHEET_CARDS), 1) - 1
    investCount = modUtils.GetLastRow(ThisWorkbook.Worksheets(modConfig.SHEET_INVESTMENTS), 1) - 1

    info = "[MY LIFE] Financial Management System" & vbCrLf & _
           "═══════════════════════════════════════" & vbCrLf & vbCrLf & _
           "Current Data:" & vbCrLf & _
           "  Bank Transactions: " & IIf(banksCount < 0, 0, banksCount) & vbCrLf & _
           "  Card Transactions: " & IIf(cardsCount < 0, 0, cardsCount) & vbCrLf & _
           "  Investments: " & IIf(investCount < 0, 0, investCount) & vbCrLf & vbCrLf & _
           "Version: 1.0" & vbCrLf & _
           "Platform: Excel for MacOS"

    MsgBox info, vbInformation, "[MY LIFE] System Info"
End Sub
