Attribute VB_Name = "modDashboard"
' ========================================================================
' Module: modDashboard
' Purpose: Dashboard Generation and Management
' Description: Builds and updates the executive dashboard with KPIs,
'              consolidated views, and charts
' Platform: Excel for MacOS
' Author: [MY LIFE] Financial System v1.0
' ========================================================================

Option Explicit

' ========================================================================
' Update Dashboard - Main Entry Point
' ========================================================================
Public Sub UpdateDashboard()
    On Error GoTo ErrorHandler

    modUtils.ShowProgress "Updating dashboard..."
    modUtils.LogDebug "Starting dashboard update"

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Update all dashboard components
    Call UpdateKPIs
    Call UpdateConsolidatedCash
    Call UpdateConsolidatedCards
    Call UpdateConsolidatedTransactions
    Call UpdateConsolidatedDebts
    Call DefineNamedRanges

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    modUtils.ClearProgress
    modUtils.LogDebug "Dashboard update complete"

    MsgBox "Dashboard updated successfully!", vbInformation

    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    modUtils.ClearProgress
    MsgBox modUtils.FormatErrorMessage("modDashboard", "UpdateDashboard", Err.Description), vbCritical
End Sub

' ========================================================================
' Update KPIs
' ========================================================================
Private Sub UpdateKPIs()
    On Error Resume Next

    Dim ws As Worksheet
    Dim wsBank As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim value As Double
    Dim totalIncome As Double
    Dim totalExpenses As Double
    Dim balance As Double

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_DASHBOARD)
    Set wsBank = ThisWorkbook.Worksheets(modConfig.SHEET_BANKS)

    lastRow = modUtils.GetLastRow(wsBank, 1)

    totalIncome = 0
    totalExpenses = 0

    ' Calculate from bank transactions
    For i = 2 To lastRow
        value = wsBank.Cells(i, 4).Value

        If value > 0 Then
            totalIncome = totalIncome + value
        Else
            totalExpenses = totalExpenses + Abs(value)
        End If
    Next i

    balance = totalIncome - totalExpenses

    ' Write KPIs to dashboard (assuming specific cells)
    ' Row 3 for KPIs
    ws.Cells(3, 2).Value = totalIncome
    ws.Cells(3, 4).Value = totalExpenses
    ws.Cells(3, 6).Value = balance

    ' Format as currency
    ws.Cells(3, 2).NumberFormat = "R$ #,##0.00"
    ws.Cells(3, 4).NumberFormat = "R$ #,##0.00"
    ws.Cells(3, 6).NumberFormat = "R$ #,##0.00"
End Sub

' ========================================================================
' Update Consolidated Cash (Banks + Investments)
' ========================================================================
Private Sub UpdateConsolidatedCash()
    On Error Resume Next

    Dim ws As Worksheet
    Dim wsBank As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim startRow As Integer
    Dim currentRow As Integer
    Dim dict As Object
    Dim bank As String
    Dim value As Double
    Dim key As Variant

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_DASHBOARD)
    Set wsBank = ThisWorkbook.Worksheets(modConfig.SHEET_BANKS)

    ' Use dictionary to aggregate by bank
    Set dict = CreateObject("Scripting.Dictionary")

    lastRow = modUtils.GetLastRow(wsBank, 1)

    ' Aggregate bank balances
    For i = 2 To lastRow
        bank = Trim(CStr(wsBank.Cells(i, 1).Value))
        value = wsBank.Cells(i, 4).Value

        If dict.Exists(bank) Then
            dict(bank) = dict(bank) + value
        Else
            dict.Add bank, value
        End If
    Next i

    ' Write to dashboard (starting at row 8)
    startRow = 8
    currentRow = startRow

    ' Clear existing data
    ws.Range("A" & startRow & ":C" & (startRow + 20)).ClearContents

    ' Headers
    ws.Cells(startRow - 1, 1).Value = "Bank"
    ws.Cells(startRow - 1, 2).Value = "Balance"
    ws.Cells(startRow - 1, 3).Value = "Type"

    ' Write data
    For Each key In dict.Keys
        ws.Cells(currentRow, 1).Value = key
        ws.Cells(currentRow, 2).Value = dict(key)
        ws.Cells(currentRow, 3).Value = "Bank"
        ws.Cells(currentRow, 2).NumberFormat = "R$ #,##0.00"
        currentRow = currentRow + 1
    Next key
End Sub

' ========================================================================
' Update Consolidated Cards
' ========================================================================
Private Sub UpdateConsolidatedCards()
    On Error Resume Next

    Dim ws As Worksheet
    Dim wsCard As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim startRow As Integer
    Dim currentRow As Integer
    Dim dict As Object
    Dim card As String
    Dim value As Double
    Dim key As Variant

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_DASHBOARD)
    Set wsCard = ThisWorkbook.Worksheets(modConfig.SHEET_CARDS)

    ' Use dictionary to aggregate by card
    Set dict = CreateObject("Scripting.Dictionary")

    lastRow = modUtils.GetLastRow(wsCard, 1)

    ' Aggregate card balances
    For i = 2 To lastRow
        card = Trim(CStr(wsCard.Cells(i, 1).Value))
        value = wsCard.Cells(i, 7).Value

        If dict.Exists(card) Then
            dict(card) = dict(card) + value
        Else
            dict.Add card, value
        End If
    Next i

    ' Write to dashboard (starting at row 8, column E)
    startRow = 8
    currentRow = startRow

    ' Clear existing data
    ws.Range("E" & startRow & ":G" & (startRow + 20)).ClearContents

    ' Headers
    ws.Cells(startRow - 1, 5).Value = "Card"
    ws.Cells(startRow - 1, 6).Value = "Total"
    ws.Cells(startRow - 1, 7).Value = "Type"

    ' Write data
    For Each key In dict.Keys
        ws.Cells(currentRow, 5).Value = key
        ws.Cells(currentRow, 6).Value = dict(key)
        ws.Cells(currentRow, 7).Value = "Card"
        ws.Cells(currentRow, 6).NumberFormat = "R$ #,##0.00"
        currentRow = currentRow + 1
    Next key
End Sub

' ========================================================================
' Update Consolidated Transactions by Category
' ========================================================================
Private Sub UpdateConsolidatedTransactions()
    On Error Resume Next

    Dim ws As Worksheet
    Dim wsBank As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim startRow As Integer
    Dim currentRow As Integer
    Dim dict As Object
    Dim category As String
    Dim value As Double
    Dim key As Variant

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_DASHBOARD)
    Set wsBank = ThisWorkbook.Worksheets(modConfig.SHEET_BANKS)

    ' Use dictionary to aggregate by category
    Set dict = CreateObject("Scripting.Dictionary")

    lastRow = modUtils.GetLastRow(wsBank, 1)

    ' Aggregate by category
    For i = 2 To lastRow
        category = Trim(CStr(wsBank.Cells(i, 5).Value))
        value = wsBank.Cells(i, 4).Value

        If category = "" Then category = "Unclassified"

        If dict.Exists(category) Then
            dict(category) = dict(category) + value
        Else
            dict.Add category, value
        End If
    Next i

    ' Write to dashboard (starting at row 35)
    startRow = 35
    currentRow = startRow

    ' Clear existing data
    ws.Range("A" & startRow & ":B" & (startRow + 30)).ClearContents

    ' Headers
    ws.Cells(startRow - 1, 1).Value = "Category"
    ws.Cells(startRow - 1, 2).Value = "Amount"

    ' Write data
    For Each key In dict.Keys
        ws.Cells(currentRow, 1).Value = key
        ws.Cells(currentRow, 2).Value = dict(key)
        ws.Cells(currentRow, 2).NumberFormat = "R$ #,##0.00"
        currentRow = currentRow + 1
    Next key
End Sub

' ========================================================================
' Update Consolidated Debts
' ========================================================================
Private Sub UpdateConsolidatedDebts()
    On Error Resume Next

    Dim ws As Worksheet
    Dim wsDebt As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim startRow As Integer
    Dim currentRow As Integer
    Dim creditor As String
    Dim updatedAmount As Double
    Dim currency As String

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_DASHBOARD)
    Set wsDebt = ThisWorkbook.Worksheets(modConfig.SHEET_DEBTS)

    lastRow = modUtils.GetLastRow(wsDebt, 1)

    ' Write to dashboard (starting at row 35, column E)
    startRow = 35
    currentRow = startRow

    ' Clear existing data
    ws.Range("E" & startRow & ":G" & (startRow + 20)).ClearContents

    ' Headers
    ws.Cells(startRow - 1, 5).Value = "Creditor"
    ws.Cells(startRow - 1, 6).Value = "Amount"
    ws.Cells(startRow - 1, 7).Value = "Currency"

    ' Write data
    For i = 2 To lastRow
        creditor = Trim(CStr(wsDebt.Cells(i, 1).Value))
        updatedAmount = wsDebt.Cells(i, 4).Value
        currency = Trim(CStr(wsDebt.Cells(i, 5).Value))

        ws.Cells(currentRow, 5).Value = creditor
        ws.Cells(currentRow, 6).Value = updatedAmount
        ws.Cells(currentRow, 7).Value = currency
        ws.Cells(currentRow, 6).NumberFormat = "#,##0.00"

        currentRow = currentRow + 1
    Next i
End Sub

' ========================================================================
' Define Named Ranges for Easy Reference
' ========================================================================
Private Sub DefineNamedRanges()
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_DASHBOARD)

    ' Delete existing named ranges if they exist
    ThisWorkbook.Names(modConfig.NR_TOTAL_INCOME).Delete
    ThisWorkbook.Names(modConfig.NR_TOTAL_EXPENSES).Delete
    ThisWorkbook.Names(modConfig.NR_BALANCE).Delete

    ' Create named ranges for KPIs
    ThisWorkbook.Names.Add Name:=modConfig.NR_TOTAL_INCOME, _
                          RefersTo:=ws.Cells(3, 2)

    ThisWorkbook.Names.Add Name:=modConfig.NR_TOTAL_EXPENSES, _
                          RefersTo:=ws.Cells(3, 4)

    ThisWorkbook.Names.Add Name:=modConfig.NR_BALANCE, _
                          RefersTo:=ws.Cells(3, 6)
End Sub

' ========================================================================
' Create Filter Dropdowns
' ========================================================================
Public Sub CreateFilterDropdowns()
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_DASHBOARD)

    ' Create dropdowns at row 1
    ' Year dropdown
    With ws.Cells(1, 2).Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="2024,2025,2026,All"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    ' Month dropdown
    With ws.Cells(1, 4).Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec,All"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    ' Institution dropdown
    With ws.Cells(1, 6).Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="ITAU,NUBANK,C6,BB,All"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    ' Currency dropdown
    With ws.Cells(1, 8).Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="BRL,USD,All"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    ' Set default values
    ws.Cells(1, 2).Value = "All"
    ws.Cells(1, 4).Value = "All"
    ws.Cells(1, 6).Value = "All"
    ws.Cells(1, 8).Value = "BRL"
End Sub

' ========================================================================
' Format Dashboard
' ========================================================================
Public Sub FormatDashboard()
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_DASHBOARD)

    Application.ScreenUpdating = False

    ' Set column widths
    ws.Columns("A:H").ColumnWidth = 15

    ' Format headers (row 1)
    With ws.Range("A1:H1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Format KPI section (row 2-3)
    ws.Cells(2, 1).Value = "Total Income"
    ws.Cells(2, 3).Value = "Total Expenses"
    ws.Cells(2, 5).Value = "Balance"

    With ws.Range("A2:F3")
        .Font.Bold = True
        .Font.Size = 14
    End With

    Application.ScreenUpdating = True
End Sub

' ========================================================================
' Quick Dashboard Setup (Run Once)
' ========================================================================
Public Sub SetupDashboard()
    On Error GoTo ErrorHandler

    Call CreateFilterDropdowns
    Call FormatDashboard
    Call UpdateDashboard

    MsgBox "Dashboard setup complete!", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox modUtils.FormatErrorMessage("modDashboard", "SetupDashboard", Err.Description), vbCritical
End Sub
