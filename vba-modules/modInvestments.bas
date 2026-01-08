Attribute VB_Name = "modInvestments"
' ========================================================================
' Module: modInvestments
' Purpose: Investment Detection Logic
' Description: Detects and classifies investment transactions including
'              applications, redemptions, and liquidations
' Platform: Excel for MacOS
' Author: Financial Import System v1.0
' ========================================================================

Option Explicit

' Investment types
Public Const INV_APPLICATION As String = "Application"
Public Const INV_REDEMPTION As String = "Redemption"
Public Const INV_LIQUIDATION As String = "Liquidation"
Public Const INV_DIVIDEND As String = "Dividend"
Public Const INV_INTEREST As String = "Interest"

' ========================================================================
' Main Investment Detection Function
' ========================================================================
Public Function DetectAllInvestments() As Long
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim trans As modNormalize.NormalizedTransaction
    Dim investmentCount As Long
    Dim investmentType As String

    modUtils.LogMessage "Starting investment detection", "INFO"
    modUtils.ShowProgress "Detecting investments..."

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_TRANSACTIONS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    investmentCount = 0

    ' Process each transaction
    For i = 2 To lastRow
        modUtils.ShowProgress "Detecting investments...", i - 1, lastRow - 1

        trans = GetTransactionFromRow(ws, i)

        ' Skip if already categorized as transfer
        If trans.TransferID <> "" Then GoTo NextTransaction

        ' Detect investment type
        investmentType = DetectInvestmentType(trans.Description, trans.NormalizedAmount)

        If investmentType <> "" Then
            ' Update transaction
            ws.Cells(i, 5).Value = "Investment"
            ws.Cells(i, 6).Value = investmentType
            ws.Cells(i, 10).Value = investmentType

            investmentCount = investmentCount + 1

            modUtils.LogMessage "Investment detected: " & investmentType & " - " & trans.Description, "DEBUG"
        End If

NextTransaction:
    Next i

    modUtils.ClearProgress
    modUtils.LogMessage "Investment detection complete: " & investmentCount & " investments found", "INFO"

    DetectAllInvestments = investmentCount

    Exit Function

ErrorHandler:
    modUtils.LogMessage "Error in DetectAllInvestments: " & Err.Description, "ERROR"
    modUtils.ClearProgress
    DetectAllInvestments = 0
End Function

' ========================================================================
' Detect Investment Type from Description
' ========================================================================
Private Function DetectInvestmentType(description As String, amount As Double) As String
    On Error Resume Next

    Dim cleanDesc As String

    cleanDesc = modUtils.CleanText(description)

    ' Application keywords
    If modUtils.ContainsAny(cleanDesc, "APLICACAO,APPLICATION,INVESTMENT,INVESTIMENTO,BUY,COMPRA,PURCHASE,APORTE") Then
        ' Applications are typically negative (money out)
        DetectInvestmentType = INV_APPLICATION
        Exit Function
    End If

    ' Redemption keywords
    If modUtils.ContainsAny(cleanDesc, "RESGATE,REDEMPTION,RESCUE,WITHDRAWAL,SAQUE,SELL,VENDA") Then
        ' Redemptions are typically positive (money in)
        DetectInvestmentType = INV_REDEMPTION
        Exit Function
    End If

    ' Liquidation keywords
    If modUtils.ContainsAny(cleanDesc, "LIQUIDACAO,LIQUIDATION,CLOSURE,ENCERRAMENTO,LIQUIDATE") Then
        ' Liquidations are typically positive (money in)
        DetectInvestmentType = INV_LIQUIDATION
        Exit Function
    End If

    ' Dividend keywords
    If modUtils.ContainsAny(cleanDesc, "DIVIDEND,DIVIDENDO,DIV,DISTRIBUTION") Then
        DetectInvestmentType = INV_DIVIDEND
        Exit Function
    End If

    ' Interest keywords
    If modUtils.ContainsAny(cleanDesc, "INTEREST,JUROS,RENDIMENTO,YIELD,EARNINGS") Then
        DetectInvestmentType = INV_INTEREST
        Exit Function
    End If

    ' Check for common fund/security identifiers
    If modUtils.ContainsAny(cleanDesc, "FUND,FUNDO,STOCK,ACAO,BOND,TITULO,DEBENTURE,CDB,LCI,LCA,TESOURO") Then
        ' Use amount to infer type
        If amount < 0 Then
            DetectInvestmentType = INV_APPLICATION
        Else
            DetectInvestmentType = INV_REDEMPTION
        End If
        Exit Function
    End If

    DetectInvestmentType = ""
End Function

' ========================================================================
' Get Transaction from Row
' ========================================================================
Private Function GetTransactionFromRow(ws As Worksheet, row As Long) As modNormalize.NormalizedTransaction
    On Error Resume Next

    Dim trans As modNormalize.NormalizedTransaction

    trans.TransactionID = ws.Cells(row, 1).Value
    trans.TransactionDate = ws.Cells(row, 2).Value
    trans.Description = ws.Cells(row, 3).Value
    trans.NormalizedAmount = ws.Cells(row, 4).Value
    trans.Category = ws.Cells(row, 5).Value
    trans.Subcategory = ws.Cells(row, 6).Value
    trans.Account = ws.Cells(row, 7).Value
    trans.Institution = ws.Cells(row, 8).Value
    trans.TransferID = ws.Cells(row, 9).Value
    trans.InvestmentType = ws.Cells(row, 10).Value

    GetTransactionFromRow = trans
End Function

' ========================================================================
' Classify Investment by Institution
' ========================================================================
Public Function ClassifyInvestmentByInstitution(trans As modNormalize.NormalizedTransaction) As String
    On Error Resume Next

    Dim institution As String

    institution = modUtils.CleanText(trans.Institution)

    ' Common investment institutions
    Select Case True
        Case modUtils.ContainsAny(institution, "VANGUARD,FIDELITY,SCHWAB,ETRADE")
            ClassifyInvestmentByInstitution = "Brokerage"
        Case modUtils.ContainsAny(institution, "401K,IRA,RETIREMENT,APOSENTADORIA")
            ClassifyInvestmentByInstitution = "Retirement"
        Case modUtils.ContainsAny(institution, "SAVINGS,POUPANCA,CD,CERTIFICATE")
            ClassifyInvestmentByInstitution = "Savings"
        Case Else
            ClassifyInvestmentByInstitution = "Other Investment"
    End Select
End Function

' ========================================================================
' Get Investment Portfolio Summary
' ========================================================================
Public Function GetInvestmentSummary() As String
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim applications As Long
    Dim redemptions As Long
    Dim dividends As Long
    Dim totalApplied As Double
    Dim totalRedeemed As Double
    Dim summary As String

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_TRANSACTIONS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Count investment types
    For i = 2 To lastRow
        If ws.Cells(i, 5).Value = "Investment" Then
            Select Case ws.Cells(i, 10).Value
                Case INV_APPLICATION
                    applications = applications + 1
                    totalApplied = totalApplied + Abs(ws.Cells(i, 4).Value)
                Case INV_REDEMPTION, INV_LIQUIDATION
                    redemptions = redemptions + 1
                    totalRedeemed = totalRedeemed + Abs(ws.Cells(i, 4).Value)
                Case INV_DIVIDEND, INV_INTEREST
                    dividends = dividends + 1
            End Select
        End If
    Next i

    summary = "Investment Summary:" & vbCrLf & _
              "Applications: " & applications & " (" & Format(totalApplied, "$#,##0.00") & ")" & vbCrLf & _
              "Redemptions: " & redemptions & " (" & Format(totalRedeemed, "$#,##0.00") & ")" & vbCrLf & _
              "Dividends/Interest: " & dividends & vbCrLf & _
              "Net Investment: " & Format(totalRedeemed - totalApplied, "$#,##0.00")

    GetInvestmentSummary = summary
End Function

' ========================================================================
' Get Investment Statistics
' ========================================================================
Public Function GetInvestmentStats() As String
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim investmentCount As Long

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_TRANSACTIONS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        If ws.Cells(i, 5).Value = "Investment" Then
            investmentCount = investmentCount + 1
        End If
    Next i

    GetInvestmentStats = "Investments Detected: " & investmentCount
End Function

' ========================================================================
' Manual Investment Classification
' ========================================================================
Public Sub ClassifyAsInvestment(transactionID As String, investmentType As String)
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_TRANSACTIONS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Find and update transaction
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = transactionID Then
            ws.Cells(i, 5).Value = "Investment"
            ws.Cells(i, 6).Value = investmentType
            ws.Cells(i, 10).Value = investmentType

            modUtils.LogMessage "Manual investment classification: " & transactionID & " -> " & investmentType, "INFO"
            Exit Sub
        End If
    Next i
End Sub

' ========================================================================
' Clear Investment Classification
' ========================================================================
Public Sub ClearInvestmentClassification(transactionID As String)
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_TRANSACTIONS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Find and clear classification
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = transactionID Then
            If ws.Cells(i, 5).Value = "Investment" Then
                ws.Cells(i, 5).Value = ""
                ws.Cells(i, 6).Value = ""
                ws.Cells(i, 10).Value = ""

                modUtils.LogMessage "Investment classification cleared: " & transactionID, "INFO"
            End If
            Exit Sub
        End If
    Next i
End Sub

' ========================================================================
' Generate Investment Report
' ========================================================================
Public Sub GenerateInvestmentReport()
    On Error Resume Next

    Dim ws As Worksheet
    Dim reportWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim reportRow As Long

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_TRANSACTIONS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Create or clear report sheet
    On Error Resume Next
    Set reportWs = ThisWorkbook.Worksheets("InvestmentReport")
    If reportWs Is Nothing Then
        Set reportWs = ThisWorkbook.Worksheets.Add
        reportWs.Name = "InvestmentReport"
    Else
        reportWs.Cells.Clear
    End If
    On Error GoTo 0

    ' Setup headers
    With reportWs
        .Cells(1, 1).Value = "Date"
        .Cells(1, 2).Value = "Description"
        .Cells(1, 3).Value = "Type"
        .Cells(1, 4).Value = "Amount"
        .Cells(1, 5).Value = "Institution"
        .Range("A1:E1").Font.Bold = True
        .Range("A1:E1").Interior.Color = RGB(68, 114, 196)
        .Range("A1:E1").Font.Color = RGB(255, 255, 255)
    End With

    reportRow = 2

    ' Copy investment transactions
    For i = 2 To lastRow
        If ws.Cells(i, 5).Value = "Investment" Then
            reportWs.Cells(reportRow, 1).Value = ws.Cells(i, 2).Value ' Date
            reportWs.Cells(reportRow, 2).Value = ws.Cells(i, 3).Value ' Description
            reportWs.Cells(reportRow, 3).Value = ws.Cells(i, 10).Value ' Investment Type
            reportWs.Cells(reportRow, 4).Value = ws.Cells(i, 4).Value ' Amount
            reportWs.Cells(reportRow, 5).Value = ws.Cells(i, 8).Value ' Institution

            reportRow = reportRow + 1
        End If
    Next i

    ' Format
    reportWs.Columns("A:E").AutoFit

    modUtils.LogMessage "Investment report generated", "INFO"
    MsgBox "Investment report generated successfully!", vbInformation
End Sub
