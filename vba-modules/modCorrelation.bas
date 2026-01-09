Attribute VB_Name = "modCorrelation"
' ========================================================================
' Module: modCorrelation
' Purpose: Investment-Bank Correlation Engine
' Description: Correlates investment applications/redemptions with
'              corresponding bank transactions
' Platform: Excel for MacOS
' Author: [MY LIFE] Financial System v1.0
' ========================================================================

Option Explicit

' Correlation tolerance
Private Const DATE_TOLERANCE_DAYS As Integer = 3
Private Const AMOUNT_TOLERANCE As Double = 0.01

' ========================================================================
' Main Correlation Function - Correlate All Investments
' ========================================================================
Public Function CorrelateAllInvestments() As Long
    On Error GoTo ErrorHandler

    Dim totalCorrelated As Long

    modUtils.ShowProgress "Correlating investments with bank transactions..."
    modUtils.LogDebug "Starting investment correlation"

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Clear existing correlations
    Call ClearExistingCorrelations

    ' Correlate applications (bank → investment)
    totalCorrelated = totalCorrelated + CorrelateApplications()

    ' Correlate redemptions (investment → bank)
    totalCorrelated = totalCorrelated + CorrelateRedemptions()

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    modUtils.ClearProgress
    modUtils.LogDebug "Correlation complete: " & totalCorrelated & " correlations"

    CorrelateAllInvestments = totalCorrelated

    Exit Function

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    modUtils.ClearProgress
    MsgBox modUtils.FormatErrorMessage("modCorrelation", "CorrelateAllInvestments", Err.Description), vbCritical
    CorrelateAllInvestments = 0
End Function

' ========================================================================
' Clear Existing Correlations
' ========================================================================
Private Sub ClearExistingCorrelations()
    On Error Resume Next

    Dim wsBank As Worksheet
    Dim wsInv As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set wsBank = ThisWorkbook.Worksheets(modConfig.SHEET_BANKS)
    Set wsInv = ThisWorkbook.Worksheets(modConfig.SHEET_INVESTMENTS)

    ' Clear bank correlations
    lastRow = modUtils.GetLastRow(wsBank, 1)
    For i = 2 To lastRow
        wsBank.Cells(i, 7).Value = ""  ' CorrelationID column
    Next i

    ' Clear investment correlations
    lastRow = modUtils.GetLastRow(wsInv, 1)
    For i = 2 To lastRow
        wsInv.Cells(i, 6).Value = ""  ' CorrelationID column
    Next i
End Sub

' ========================================================================
' Correlate Applications (Investment Applications → Bank Withdrawals)
' ========================================================================
Private Function CorrelateApplications() As Long
    On Error GoTo ErrorHandler

    Dim wsBank As Worksheet
    Dim wsInv As Worksheet
    Dim invLastRow As Long
    Dim bankLastRow As Long
    Dim i As Long, j As Long
    Dim invDate As Date
    Dim invAmount As Double
    Dim invType As String
    Dim bankDate As Date
    Dim bankAmount As Double
    Dim bankCorrelation As String
    Dim correlated As Long
    Dim correlationID As String

    Set wsBank = ThisWorkbook.Worksheets(modConfig.SHEET_BANKS)
    Set wsInv = ThisWorkbook.Worksheets(modConfig.SHEET_INVESTMENTS)

    invLastRow = modUtils.GetLastRow(wsInv, 1)
    bankLastRow = modUtils.GetLastRow(wsBank, 1)

    correlated = 0

    ' Loop through investment transactions
    For i = 2 To invLastRow
        invType = Trim(CStr(wsInv.Cells(i, 5).Value))

        ' Only process applications
        If UCase(invType) = "APPLICATION" Or modUtils.ContainsAny(invType, "APLICACAO,INVEST,APORTE") Then
            invDate = wsInv.Cells(i, 2).Value
            invAmount = wsInv.Cells(i, 4).Value

            ' Look for matching bank withdrawal
            For j = 2 To bankLastRow
                bankDate = wsBank.Cells(j, 2).Value
                bankAmount = wsBank.Cells(j, 4).Value
                bankCorrelation = Trim(CStr(wsBank.Cells(j, 7).Value))

                ' Check if bank transaction is not already correlated
                If bankCorrelation = "" Then
                    ' Check if dates match within tolerance
                    If modUtils.DatesMatch(invDate, bankDate, DATE_TOLERANCE_DAYS) Then
                        ' Check if amounts match (bank withdrawal should be negative)
                        If modUtils.AmountsMatch(invAmount, -bankAmount, AMOUNT_TOLERANCE) Then
                            ' Found a match!
                            correlationID = modUtils.GenerateUniqueID("CORR")

                            wsInv.Cells(i, 6).Value = correlationID
                            wsBank.Cells(j, 7).Value = correlationID

                            correlated = correlated + 1
                            Exit For  ' Move to next investment
                        End If
                    End If
                End If
            Next j
        End If
    Next i

    CorrelateApplications = correlated

    Exit Function

ErrorHandler:
    MsgBox modUtils.FormatErrorMessage("modCorrelation", "CorrelateApplications", Err.Description), vbCritical
    CorrelateApplications = 0
End Function

' ========================================================================
' Correlate Redemptions (Investment Redemptions → Bank Deposits)
' ========================================================================
Private Function CorrelateRedemptions() As Long
    On Error GoTo ErrorHandler

    Dim wsBank As Worksheet
    Dim wsInv As Worksheet
    Dim invLastRow As Long
    Dim bankLastRow As Long
    Dim i As Long, j As Long
    Dim invDate As Date
    Dim invAmount As Double
    Dim invType As String
    Dim bankDate As Date
    Dim bankAmount As Double
    Dim bankCorrelation As String
    Dim correlated As Long
    Dim correlationID As String

    Set wsBank = ThisWorkbook.Worksheets(modConfig.SHEET_BANKS)
    Set wsInv = ThisWorkbook.Worksheets(modConfig.SHEET_INVESTMENTS)

    invLastRow = modUtils.GetLastRow(wsInv, 1)
    bankLastRow = modUtils.GetLastRow(wsBank, 1)

    correlated = 0

    ' Loop through investment transactions
    For i = 2 To invLastRow
        invType = Trim(CStr(wsInv.Cells(i, 5).Value))

        ' Only process redemptions
        If UCase(invType) = "REDEMPTION" Or modUtils.ContainsAny(invType, "RESGATE,REDEEM,WITHDRAW,SAQUE") Then
            invDate = wsInv.Cells(i, 2).Value
            invAmount = wsInv.Cells(i, 4).Value

            ' Look for matching bank deposit
            For j = 2 To bankLastRow
                bankDate = wsBank.Cells(j, 2).Value
                bankAmount = wsBank.Cells(j, 4).Value
                bankCorrelation = Trim(CStr(wsBank.Cells(j, 7).Value))

                ' Check if bank transaction is not already correlated
                If bankCorrelation = "" Then
                    ' Check if dates match within tolerance
                    If modUtils.DatesMatch(invDate, bankDate, DATE_TOLERANCE_DAYS) Then
                        ' Check if amounts match (bank deposit should be positive)
                        ' Investment redemption amount could be positive or negative
                        If modUtils.AmountsMatch(Abs(invAmount), bankAmount, AMOUNT_TOLERANCE) Then
                            ' Found a match!
                            correlationID = modUtils.GenerateUniqueID("CORR")

                            wsInv.Cells(i, 6).Value = correlationID
                            wsBank.Cells(j, 7).Value = correlationID

                            correlated = correlated + 1
                            Exit For  ' Move to next investment
                        End If
                    End If
                End If
            Next j
        End If
    Next i

    CorrelateRedemptions = correlated

    Exit Function

ErrorHandler:
    MsgBox modUtils.FormatErrorMessage("modCorrelation", "CorrelateRedemptions", Err.Description), vbCritical
    CorrelateRedemptions = 0
End Function

' ========================================================================
' Get Uncorrelated Investment Count
' ========================================================================
Public Function GetUncorrelatedInvestmentCount() As Long
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim count As Long

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_INVESTMENTS)
    lastRow = modUtils.GetLastRow(ws, 1)

    count = 0

    For i = 2 To lastRow
        If Trim(CStr(ws.Cells(i, 6).Value)) = "" Then
            count = count + 1
        End If
    Next i

    GetUncorrelatedInvestmentCount = count

    Exit Function

ErrorHandler:
    GetUncorrelatedInvestmentCount = 0
End Function

' ========================================================================
' Get Correlation Balance Status
' ========================================================================
Public Function ValidateCorrelationBalance() As Boolean
    On Error GoTo ErrorHandler

    Dim wsBank As Worksheet
    Dim wsInv As Worksheet
    Dim bankLastRow As Long
    Dim invLastRow As Long
    Dim i As Long
    Dim correlationID As String
    Dim bankAmount As Double
    Dim invAmount As Double
    Dim balanceOK As Boolean
    Dim dict As Object
    Dim key As Variant

    Set wsBank = ThisWorkbook.Worksheets(modConfig.SHEET_BANKS)
    Set wsInv = ThisWorkbook.Worksheets(modConfig.SHEET_INVESTMENTS)

    ' Use dictionary to track correlations
    Set dict = CreateObject("Scripting.Dictionary")

    ' Collect bank transactions
    bankLastRow = modUtils.GetLastRow(wsBank, 1)
    For i = 2 To bankLastRow
        correlationID = Trim(CStr(wsBank.Cells(i, 7).Value))
        If correlationID <> "" Then
            bankAmount = wsBank.Cells(i, 4).Value

            If dict.Exists(correlationID) Then
                dict(correlationID) = dict(correlationID) + bankAmount
            Else
                dict.Add correlationID, bankAmount
            End If
        End If
    Next i

    ' Check investment transactions
    invLastRow = modUtils.GetLastRow(wsInv, 1)
    For i = 2 To invLastRow
        correlationID = Trim(CStr(wsInv.Cells(i, 6).Value))
        If correlationID <> "" Then
            invAmount = wsInv.Cells(i, 4).Value

            If dict.Exists(correlationID) Then
                dict(correlationID) = dict(correlationID) + invAmount
            Else
                dict.Add correlationID, invAmount
            End If
        End If
    Next i

    ' Validate that all correlations balance to near zero
    balanceOK = True
    For Each key In dict.Keys
        If Abs(dict(key)) > AMOUNT_TOLERANCE Then
            balanceOK = False
            modUtils.LogDebug "Correlation " & key & " imbalance: " & dict(key)
        End If
    Next key

    ValidateCorrelationBalance = balanceOK

    Exit Function

ErrorHandler:
    ValidateCorrelationBalance = False
End Function

' ========================================================================
' Manual Correlation Helper
' ========================================================================
Public Sub ManualCorrelate(bankRow As Long, invRow As Long)
    On Error GoTo ErrorHandler

    Dim wsBank As Worksheet
    Dim wsInv As Worksheet
    Dim correlationID As String

    Set wsBank = ThisWorkbook.Worksheets(modConfig.SHEET_BANKS)
    Set wsInv = ThisWorkbook.Worksheets(modConfig.SHEET_INVESTMENTS)

    correlationID = modUtils.GenerateUniqueID("CORR-MANUAL")

    wsBank.Cells(bankRow, 7).Value = correlationID
    wsInv.Cells(invRow, 6).Value = correlationID

    MsgBox "Manual correlation created: " & correlationID, vbInformation

    Exit Sub

ErrorHandler:
    MsgBox modUtils.FormatErrorMessage("modCorrelation", "ManualCorrelate", Err.Description), vbCritical
End Sub
