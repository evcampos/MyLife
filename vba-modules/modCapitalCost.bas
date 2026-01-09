Attribute VB_Name = "modCapitalCost"
' ========================================================================
' Module: modCapitalCost
' Purpose: Capital Cost Calculations
' Description: Calculates updated values for investments and debts
'              using index-based adjustments
' Platform: Excel for MacOS
' Author: [MY LIFE] Financial System v1.0
' ========================================================================

Option Explicit

' ========================================================================
' Calculate All Capital Costs
' ========================================================================
Public Sub CalculateAllCapitalCosts()
    On Error GoTo ErrorHandler

    modUtils.ShowProgress "Calculating capital costs..."
    modUtils.LogDebug "Starting capital cost calculation"

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Calculate for OPUS investments
    Call CalculateOPUSCapitalCosts

    ' Calculate for Debts
    Call CalculateDebtCapitalCosts

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    modUtils.ClearProgress
    modUtils.LogDebug "Capital cost calculation complete"

    MsgBox "Capital costs calculated successfully!", vbInformation

    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    modUtils.ClearProgress
    MsgBox modUtils.FormatErrorMessage("modCapitalCost", "CalculateAllCapitalCosts", Err.Description), vbCritical
End Sub

' ========================================================================
' Calculate OPUS Capital Costs
' ========================================================================
Private Sub CalculateOPUSCapitalCosts()
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim investCost As Double
    Dim capitalCostRate As String
    Dim currency As String
    Dim updatedCost As Double
    Dim indexName As String

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_OPUS)
    lastRow = modUtils.GetLastRow(ws, 1)

    For i = 2 To lastRow
        investCost = ws.Cells(i, 3).Value
        capitalCostRate = Trim(CStr(ws.Cells(i, 4).Value))
        currency = Trim(CStr(ws.Cells(i, 6).Value))

        ' Determine index based on currency and rate
        indexName = DetermineIndexForCurrency(currency, capitalCostRate)

        ' Calculate updated cost
        If indexName <> "" Then
            updatedCost = CalculateUpdatedValue(investCost, indexName, Date)
        Else
            ' Use simple percentage if specified
            updatedCost = CalculateSimpleInterest(investCost, capitalCostRate)
        End If

        ' Update the sheet
        ws.Cells(i, 5).Value = updatedCost
    Next i
End Sub

' ========================================================================
' Calculate Debt Capital Costs
' ========================================================================
Private Sub CalculateDebtCapitalCosts()
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim amountPaid As Double
    Dim interestRate As String
    Dim currency As String
    Dim lastUpdate As Date
    Dim updatedAmount As Double
    Dim indexName As String

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_DEBTS)
    lastRow = modUtils.GetLastRow(ws, 1)

    For i = 2 To lastRow
        amountPaid = ws.Cells(i, 3).Value
        interestRate = Trim(CStr(ws.Cells(i, 2).Value))
        currency = Trim(CStr(ws.Cells(i, 5).Value))
        lastUpdate = ws.Cells(i, 6).Value

        ' Determine index based on currency and rate
        indexName = DetermineIndexForCurrency(currency, interestRate)

        ' Calculate updated amount
        If indexName <> "" Then
            updatedAmount = CalculateUpdatedValue(amountPaid, indexName, Date, lastUpdate)
        Else
            ' Use simple percentage if specified
            updatedAmount = CalculateSimpleInterest(amountPaid, interestRate)
        End If

        ' Update the sheet
        ws.Cells(i, 4).Value = updatedAmount
    Next i
End Sub

' ========================================================================
' Determine Index Based on Currency
' ========================================================================
Private Function DetermineIndexForCurrency(currency As String, rateText As String) As String
    On Error Resume Next

    Dim cleanRate As String
    cleanRate = modUtils.CleanText(rateText)

    ' Check for explicit index mention
    If modUtils.ContainsAny(cleanRate, "CDI") Then
        DetermineIndexForCurrency = "CDI"
        Exit Function
    ElseIf modUtils.ContainsAny(cleanRate, "SELIC") Then
        DetermineIndexForCurrency = "SELIC"
        Exit Function
    ElseIf modUtils.ContainsAny(cleanRate, "IPCA") Then
        DetermineIndexForCurrency = "IPCA"
        Exit Function
    ElseIf modUtils.ContainsAny(cleanRate, "FED,FEDFUNDS") Then
        DetermineIndexForCurrency = "FEDFUNDS"
        Exit Function
    End If

    ' Default based on currency
    Select Case UCase(Trim(currency))
        Case "BRL"
            DetermineIndexForCurrency = "CDI"  ' Default for BRL
        Case "USD"
            DetermineIndexForCurrency = "FEDFUNDS"  ' Default for USD
        Case Else
            DetermineIndexForCurrency = ""  ' No default
    End Select
End Function

' ========================================================================
' Calculate Updated Value Using Index
' ========================================================================
Public Function CalculateUpdatedValue(initialValue As Double, _
                                      indexName As String, _
                                      Optional endDate As Date, _
                                      Optional startDate As Date) As Double
    On Error GoTo ErrorHandler

    Dim startFactor As Double
    Dim endFactor As Double
    Dim adjustmentFactor As Double
    Dim updatedValue As Double

    ' Use first available date if start date not specified
    If startDate = 0 Then
        startDate = DateSerial(2020, 1, 1)  ' Default start date
    End If

    If endDate = 0 Then
        endDate = Date
    End If

    ' Get cumulative factors
    startFactor = modIndexes.GetCumulativeFactor(indexName, startDate)
    endFactor = modIndexes.GetCumulativeFactor(indexName, endDate)

    ' Calculate adjustment
    If startFactor > 0 Then
        adjustmentFactor = endFactor / startFactor
    Else
        adjustmentFactor = 1#
    End If

    ' Calculate updated value
    updatedValue = initialValue * adjustmentFactor

    CalculateUpdatedValue = updatedValue

    Exit Function

ErrorHandler:
    CalculateUpdatedValue = initialValue  ' Return original on error
End Function

' ========================================================================
' Calculate Simple Interest (fallback when no index available)
' ========================================================================
Private Function CalculateSimpleInterest(principal As Double, rateText As String) As Double
    On Error GoTo ErrorHandler

    Dim rate As Double
    Dim cleanRate As String

    ' Extract percentage from text
    cleanRate = rateText
    cleanRate = Replace(cleanRate, "%", "")
    cleanRate = Replace(cleanRate, " ", "")

    If IsNumeric(cleanRate) Then
        rate = CDbl(cleanRate) / 100
        CalculateSimpleInterest = principal * (1 + rate)
    Else
        ' Can't parse rate, return original
        CalculateSimpleInterest = principal
    End If

    Exit Function

ErrorHandler:
    CalculateSimpleInterest = principal
End Function

' ========================================================================
' Calculate Return on Investment
' ========================================================================
Public Function CalculateROI(initialValue As Double, currentValue As Double) As Double
    On Error Resume Next

    If initialValue > 0 Then
        CalculateROI = ((currentValue - initialValue) / initialValue) * 100
    Else
        CalculateROI = 0
    End If
End Function

' ========================================================================
' Get Total Investment Value (Sum of all OPUS assets)
' ========================================================================
Public Function GetTotalInvestmentValue() As Double
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim itemType As String
    Dim updatedCost As Double
    Dim total As Double

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_OPUS)
    lastRow = modUtils.GetLastRow(ws, 1)

    total = 0

    For i = 2 To lastRow
        itemType = Trim(CStr(ws.Cells(i, 1).Value))
        updatedCost = ws.Cells(i, 5).Value

        If UCase(itemType) = "ASSET" Then
            total = total + updatedCost
        End If
    Next i

    GetTotalInvestmentValue = total

    Exit Function

ErrorHandler:
    GetTotalInvestmentValue = 0
End Function

' ========================================================================
' Get Total Debt Value (Sum of all debts)
' ========================================================================
Public Function GetTotalDebtValue() As Double
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim updatedAmount As Double
    Dim total As Double

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_DEBTS)
    lastRow = modUtils.GetLastRow(ws, 1)

    total = 0

    For i = 2 To lastRow
        updatedAmount = ws.Cells(i, 4).Value
        total = total + updatedAmount
    Next i

    GetTotalDebtValue = total

    Exit Function

ErrorHandler:
    GetTotalDebtValue = 0
End Function
