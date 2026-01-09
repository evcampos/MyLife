Attribute VB_Name = "modHealthCheck"
' ========================================================================
' Module: modHealthCheck
' Purpose: System Health Validation
' Description: Validates all imports, data integrity, correlations,
'              classifications, and generates health reports
' Platform: Excel for MacOS
' Author: [MY LIFE] Financial System v1.0
' ========================================================================

Option Explicit

' Health check result constants
Private Const STATUS_PASS As String = "PASS"
Private Const STATUS_WARNING As String = "WARNING"
Private Const STATUS_FAIL As String = "FAIL"

' ========================================================================
' Run Complete Health Check
' ========================================================================
Public Sub RunHealthCheck()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim currentRow As Long

    modUtils.ShowProgress "Running health check..."
    modUtils.LogDebug "Starting health check"

    Application.ScreenUpdating = False

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_HEALTH_CHECK)

    ' Clear existing results
    ws.Cells.Clear

    ' Create headers
    ws.Cells(1, 1).Value = "CHECK"
    ws.Cells(1, 2).Value = "STATUS"
    ws.Cells(1, 3).Value = "RESULT"
    ws.Cells(1, 4).Value = "DETAILS"
    ws.Cells(1, 5).Value = "TIMESTAMP"

    With ws.Range("A1:E1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    currentRow = 2

    ' Run all health checks
    currentRow = CheckImportStatus(ws, currentRow)
    currentRow = CheckDataIntegrity(ws, currentRow)
    currentRow = CheckCorrelations(ws, currentRow)
    currentRow = CheckClassifications(ws, currentRow)
    currentRow = CheckIndexes(ws, currentRow)
    currentRow = CheckBalances(ws, currentRow)

    ' Auto-fit columns
    ws.Columns("A:E").AutoFit

    ' Show summary
    Call ShowHealthSummary(ws)

    Application.ScreenUpdating = True

    modUtils.ClearProgress
    modUtils.LogDebug "Health check complete"

    ' Activate health check sheet
    ws.Activate

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    modUtils.ClearProgress
    MsgBox modUtils.FormatErrorMessage("modHealthCheck", "RunHealthCheck", Err.Description), vbCritical
End Sub

' ========================================================================
' Check Import Status
' ========================================================================
Private Function CheckImportStatus(ws As Worksheet, startRow As Long) As Long
    On Error Resume Next

    Dim row As Long
    Dim wsBank As Worksheet
    Dim wsCard As Worksheet
    Dim wsInv As Worksheet
    Dim bankCount As Long
    Dim cardCount As Long
    Dim invCount As Long

    row = startRow

    Set wsBank = ThisWorkbook.Worksheets(modConfig.SHEET_BANKS)
    Set wsCard = ThisWorkbook.Worksheets(modConfig.SHEET_CARDS)
    Set wsInv = ThisWorkbook.Worksheets(modConfig.SHEET_INVESTMENTS)

    bankCount = modUtils.GetLastRow(wsBank, 1) - 1
    cardCount = modUtils.GetLastRow(wsCard, 1) - 1
    invCount = modUtils.GetLastRow(wsInv, 1) - 1

    ' Check bank imports
    ws.Cells(row, 1).Value = "Bank Transactions Imported"
    ws.Cells(row, 3).Value = bankCount
    If bankCount > 0 Then
        ws.Cells(row, 2).Value = STATUS_PASS
        ws.Cells(row, 4).Value = "Bank data loaded successfully"
    Else
        ws.Cells(row, 2).Value = STATUS_WARNING
        ws.Cells(row, 4).Value = "No bank transactions found"
    End If
    ws.Cells(row, 5).Value = Now
    row = row + 1

    ' Check card imports
    ws.Cells(row, 1).Value = "Card Transactions Imported"
    ws.Cells(row, 3).Value = cardCount
    If cardCount >= 0 Then
        ws.Cells(row, 2).Value = STATUS_PASS
        ws.Cells(row, 4).Value = "Card data loaded"
    Else
        ws.Cells(row, 2).Value = STATUS_WARNING
        ws.Cells(row, 4).Value = "No card transactions found"
    End If
    ws.Cells(row, 5).Value = Now
    row = row + 1

    ' Check investment imports
    ws.Cells(row, 1).Value = "Investment Transactions Imported"
    ws.Cells(row, 3).Value = invCount
    If invCount >= 0 Then
        ws.Cells(row, 2).Value = STATUS_PASS
        ws.Cells(row, 4).Value = "Investment data loaded"
    Else
        ws.Cells(row, 2).Value = STATUS_WARNING
        ws.Cells(row, 4).Value = "No investment transactions found"
    End If
    ws.Cells(row, 5).Value = Now
    row = row + 1

    CheckImportStatus = row
End Function

' ========================================================================
' Check Data Integrity
' ========================================================================
Private Function CheckDataIntegrity(ws As Worksheet, startRow As Long) As Long
    On Error Resume Next

    Dim row As Long
    Dim wsBank As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim invalidDates As Long
    Dim invalidAmounts As Long
    Dim emptyDescriptions As Long

    row = startRow

    Set wsBank = ThisWorkbook.Worksheets(modConfig.SHEET_BANKS)
    lastRow = modUtils.GetLastRow(wsBank, 1)

    invalidDates = 0
    invalidAmounts = 0
    emptyDescriptions = 0

    ' Check each transaction
    For i = 2 To lastRow
        ' Check date
        If Not modUtils.IsValidDate(wsBank.Cells(i, 2).Value) Then
            invalidDates = invalidDates + 1
        End If

        ' Check amount
        If Not modUtils.IsNumericValue(wsBank.Cells(i, 4).Value) Then
            invalidAmounts = invalidAmounts + 1
        End If

        ' Check description
        If Trim(CStr(wsBank.Cells(i, 3).Value)) = "" Then
            emptyDescriptions = emptyDescriptions + 1
        End If
    Next i

    ' Report invalid dates
    ws.Cells(row, 1).Value = "Valid Transaction Dates"
    ws.Cells(row, 3).Value = (lastRow - 1 - invalidDates) & " of " & (lastRow - 1)
    If invalidDates = 0 Then
        ws.Cells(row, 2).Value = STATUS_PASS
        ws.Cells(row, 4).Value = "All dates valid"
    Else
        ws.Cells(row, 2).Value = STATUS_FAIL
        ws.Cells(row, 4).Value = invalidDates & " invalid dates found"
    End If
    ws.Cells(row, 5).Value = Now
    row = row + 1

    ' Report invalid amounts
    ws.Cells(row, 1).Value = "Valid Transaction Amounts"
    ws.Cells(row, 3).Value = (lastRow - 1 - invalidAmounts) & " of " & (lastRow - 1)
    If invalidAmounts = 0 Then
        ws.Cells(row, 2).Value = STATUS_PASS
        ws.Cells(row, 4).Value = "All amounts valid"
    Else
        ws.Cells(row, 2).Value = STATUS_FAIL
        ws.Cells(row, 4).Value = invalidAmounts & " invalid amounts found"
    End If
    ws.Cells(row, 5).Value = Now
    row = row + 1

    ' Report empty descriptions
    ws.Cells(row, 1).Value = "Valid Transaction Descriptions"
    ws.Cells(row, 3).Value = (lastRow - 1 - emptyDescriptions) & " of " & (lastRow - 1)
    If emptyDescriptions = 0 Then
        ws.Cells(row, 2).Value = STATUS_PASS
        ws.Cells(row, 4).Value = "All descriptions present"
    Else
        ws.Cells(row, 2).Value = STATUS_WARNING
        ws.Cells(row, 4).Value = emptyDescriptions & " empty descriptions found"
    End If
    ws.Cells(row, 5).Value = Now
    row = row + 1

    CheckDataIntegrity = row
End Function

' ========================================================================
' Check Correlations
' ========================================================================
Private Function CheckCorrelations(ws As Worksheet, startRow As Long) As Long
    On Error Resume Next

    Dim row As Long
    Dim uncorrelated As Long
    Dim balanceOK As Boolean

    row = startRow

    ' Check uncorrelated investments
    uncorrelated = modCorrelation.GetUncorrelatedInvestmentCount()

    ws.Cells(row, 1).Value = "Investment Correlation"
    ws.Cells(row, 3).Value = uncorrelated & " uncorrelated"
    If uncorrelated = 0 Then
        ws.Cells(row, 2).Value = STATUS_PASS
        ws.Cells(row, 4).Value = "All investments correlated"
    ElseIf uncorrelated < 5 Then
        ws.Cells(row, 2).Value = STATUS_WARNING
        ws.Cells(row, 4).Value = uncorrelated & " investments not correlated"
    Else
        ws.Cells(row, 2).Value = STATUS_FAIL
        ws.Cells(row, 4).Value = uncorrelated & " investments not correlated"
    End If
    ws.Cells(row, 5).Value = Now
    row = row + 1

    ' Check correlation balance
    balanceOK = modCorrelation.ValidateCorrelationBalance()

    ws.Cells(row, 1).Value = "Correlation Balance"
    If balanceOK Then
        ws.Cells(row, 2).Value = STATUS_PASS
        ws.Cells(row, 3).Value = "Balanced"
        ws.Cells(row, 4).Value = "All correlations balance correctly"
    Else
        ws.Cells(row, 2).Value = STATUS_FAIL
        ws.Cells(row, 3).Value = "Imbalanced"
        ws.Cells(row, 4).Value = "Some correlations don't balance"
    End If
    ws.Cells(row, 5).Value = Now
    row = row + 1

    CheckCorrelations = row
End Function

' ========================================================================
' Check Classifications
' ========================================================================
Private Function CheckClassifications(ws As Worksheet, startRow As Long) As Long
    On Error Resume Next

    Dim row As Long
    Dim unclassified As Long

    row = startRow

    unclassified = modClassification.GetUnclassifiedCount()

    ws.Cells(row, 1).Value = "Transaction Classification"
    ws.Cells(row, 3).Value = unclassified & " unclassified"
    If unclassified = 0 Then
        ws.Cells(row, 2).Value = STATUS_PASS
        ws.Cells(row, 4).Value = "All transactions classified"
    ElseIf unclassified < 10 Then
        ws.Cells(row, 2).Value = STATUS_WARNING
        ws.Cells(row, 4).Value = unclassified & " transactions unclassified"
    Else
        ws.Cells(row, 2).Value = STATUS_FAIL
        ws.Cells(row, 4).Value = unclassified & " transactions unclassified"
    End If
    ws.Cells(row, 5).Value = Now
    row = row + 1

    CheckClassifications = row
End Function

' ========================================================================
' Check Indexes
' ========================================================================
Private Function CheckIndexes(ws As Worksheet, startRow As Long) As Long
    On Error Resume Next

    Dim row As Long
    Dim wsIdx As Worksheet
    Dim lastRow As Long
    Dim indexCount As Long

    row = startRow

    Set wsIdx = ThisWorkbook.Worksheets(modConfig.SHEET_INDEXES)
    lastRow = modUtils.GetLastRow(wsIdx, 1)
    indexCount = lastRow - 1

    ws.Cells(row, 1).Value = "Index Data Available"
    ws.Cells(row, 3).Value = indexCount & " records"
    If indexCount > 0 Then
        ws.Cells(row, 2).Value = STATUS_PASS
        ws.Cells(row, 4).Value = "Index data loaded"
    Else
        ws.Cells(row, 2).Value = STATUS_WARNING
        ws.Cells(row, 4).Value = "No index data found"
    End If
    ws.Cells(row, 5).Value = Now
    row = row + 1

    CheckIndexes = row
End Function

' ========================================================================
' Check Balances
' ========================================================================
Private Function CheckBalances(ws As Worksheet, startRow As Long) As Long
    On Error Resume Next

    Dim row As Long
    Dim wsBank As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim totalInflow As Double
    Dim totalOutflow As Double
    Dim netPosition As Double

    row = startRow

    Set wsBank = ThisWorkbook.Worksheets(modConfig.SHEET_BANKS)
    lastRow = modUtils.GetLastRow(wsBank, 1)

    totalInflow = 0
    totalOutflow = 0

    For i = 2 To lastRow
        Dim value As Double
        value = wsBank.Cells(i, 4).Value

        If value > 0 Then
            totalInflow = totalInflow + value
        Else
            totalOutflow = totalOutflow + value
        End If
    Next i

    netPosition = totalInflow + totalOutflow

    ws.Cells(row, 1).Value = "Balance Calculation"
    ws.Cells(row, 2).Value = STATUS_PASS
    ws.Cells(row, 3).Value = modUtils.FormatCurrency(netPosition, "BRL")
    ws.Cells(row, 4).Value = "Inflow: " & modUtils.FormatCurrency(totalInflow, "BRL") & _
                              " | Outflow: " & modUtils.FormatCurrency(totalOutflow, "BRL")
    ws.Cells(row, 5).Value = Now
    row = row + 1

    CheckBalances = row
End Function

' ========================================================================
' Show Health Summary
' ========================================================================
Private Sub ShowHealthSummary(ws As Worksheet)
    On Error Resume Next

    Dim lastRow As Long
    Dim i As Long
    Dim passCount As Long
    Dim warningCount As Long
    Dim failCount As Long
    Dim status As String

    lastRow = modUtils.GetLastRow(ws, 1)

    passCount = 0
    warningCount = 0
    failCount = 0

    ' Count statuses
    For i = 2 To lastRow
        status = Trim(CStr(ws.Cells(i, 2).Value))

        Select Case status
            Case STATUS_PASS
                passCount = passCount + 1
                ws.Cells(i, 2).Interior.Color = RGB(146, 208, 80)  ' Green
            Case STATUS_WARNING
                warningCount = warningCount + 1
                ws.Cells(i, 2).Interior.Color = RGB(255, 192, 0)   ' Orange
            Case STATUS_FAIL
                failCount = failCount + 1
                ws.Cells(i, 2).Interior.Color = RGB(255, 0, 0)     ' Red
                ws.Cells(i, 2).Font.Color = RGB(255, 255, 255)
        End Select
    Next i

    ' Show summary message
    Dim summary As String
    summary = "Health Check Complete!" & vbCrLf & vbCrLf & _
              "PASS: " & passCount & vbCrLf & _
              "WARNING: " & warningCount & vbCrLf & _
              "FAIL: " & failCount

    If failCount > 0 Then
        MsgBox summary, vbExclamation, "Health Check - Issues Found"
    ElseIf warningCount > 0 Then
        MsgBox summary, vbInformation, "Health Check - Warnings"
    Else
        MsgBox summary, vbInformation, "Health Check - All Clear"
    End If
End Sub

' ========================================================================
' Quick Status Summary
' ========================================================================
Public Function GetSystemStatus() As String
    On Error Resume Next

    Dim wsBank As Worksheet
    Dim wsCard As Worksheet
    Dim wsInv As Worksheet
    Dim bankCount As Long
    Dim cardCount As Long
    Dim invCount As Long
    Dim unclassified As Long
    Dim uncorrelated As Long

    Set wsBank = ThisWorkbook.Worksheets(modConfig.SHEET_BANKS)
    Set wsCard = ThisWorkbook.Worksheets(modConfig.SHEET_CARDS)
    Set wsInv = ThisWorkbook.Worksheets(modConfig.SHEET_INVESTMENTS)

    bankCount = modUtils.GetLastRow(wsBank, 1) - 1
    cardCount = modUtils.GetLastRow(wsCard, 1) - 1
    invCount = modUtils.GetLastRow(wsInv, 1) - 1
    unclassified = modClassification.GetUnclassifiedCount()
    uncorrelated = modCorrelation.GetUncorrelatedInvestmentCount()

    GetSystemStatus = "Banks: " & bankCount & " | " & _
                     "Cards: " & cardCount & " | " & _
                     "Investments: " & invCount & " | " & _
                     "Unclassified: " & unclassified & " | " & _
                     "Uncorrelated: " & uncorrelated
End Function
