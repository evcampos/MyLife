Attribute VB_Name = "modTransfers"
' ========================================================================
' Module: modTransfers
' Purpose: Transfer Detection Logic
' Description: Automatically detects and links transfer transactions
'              between accounts based on date, amount, and patterns
' Platform: Excel for MacOS
' Author: Financial Import System v1.0
' ========================================================================

Option Explicit

' Transfer match structure
Private Type TransferMatch
    Transaction1ID As String
    Transaction2ID As String
    MatchScore As Double
    TransferID As String
End Type

' ========================================================================
' Main Transfer Detection Function
' ========================================================================
Public Function DetectAllTransfers() As Long
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim trans1 As modNormalize.NormalizedTransaction
    Dim trans2 As modNormalize.NormalizedTransaction
    Dim transferCount As Long
    Dim dateTolerance As Integer
    Dim amountTolerance As Double
    Dim transferID As String

    modUtils.LogMessage "Starting transfer detection", "INFO"
    modUtils.ShowProgress "Detecting transfers..."

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_TRANSACTIONS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Get configuration
    dateTolerance = modConfig.GetConfigInt("TransferDateDays", 1)
    amountTolerance = modConfig.GetConfigDouble("TransferTolerance", 0.001)

    transferCount = 0

    ' Compare all transactions
    For i = 2 To lastRow - 1
        ' Skip if already has TransferID
        If ws.Cells(i, 9).Value <> "" Then GoTo NextTransaction1

        ' Get transaction 1
        trans1 = GetTransactionFromRow(ws, i)

        ' Look for matching transfer
        For j = i + 1 To lastRow
            ' Skip if already has TransferID
            If ws.Cells(j, 9).Value <> "" Then GoTo NextTransaction2

            ' Get transaction 2
            trans2 = GetTransactionFromRow(ws, j)

            ' Check if this is a transfer match
            If IsTransferMatch(trans1, trans2, dateTolerance, amountTolerance) Then
                ' Generate unique transfer ID
                transferID = modUtils.GenerateUniqueID("TRF")

                ' Update both transactions
                ws.Cells(i, 9).Value = transferID
                ws.Cells(j, 9).Value = transferID

                ' Set categories
                ws.Cells(i, 5).Value = "Transfer"
                ws.Cells(i, 6).Value = "Internal Transfer"
                ws.Cells(j, 5).Value = "Transfer"
                ws.Cells(j, 6).Value = "Internal Transfer"

                transferCount = transferCount + 1

                modUtils.LogMessage "Transfer detected: " & trans1.Description & " <-> " & trans2.Description, "DEBUG"

                ' Exit inner loop after finding match
                Exit For
            End If

NextTransaction2:
        Next j

NextTransaction1:
        modUtils.ShowProgress "Detecting transfers...", i - 1, lastRow - 2
    Next i

    modUtils.ClearProgress
    modUtils.LogMessage "Transfer detection complete: " & transferCount & " transfers found", "INFO"

    DetectAllTransfers = transferCount

    Exit Function

ErrorHandler:
    modUtils.LogMessage "Error in DetectAllTransfers: " & Err.Description, "ERROR"
    modUtils.ClearProgress
    DetectAllTransfers = 0
End Function

' ========================================================================
' Check if Two Transactions are a Transfer Match
' ========================================================================
Private Function IsTransferMatch(trans1 As modNormalize.NormalizedTransaction, _
                                trans2 As modNormalize.NormalizedTransaction, _
                                dateTolerance As Integer, _
                                amountTolerance As Double) As Boolean
    On Error Resume Next

    ' Check date proximity
    If Not modUtils.DatesMatch(trans1.TransactionDate, trans2.TransactionDate, dateTolerance) Then
        IsTransferMatch = False
        Exit Function
    End If

    ' Check opposite amounts (one debit, one credit)
    If Not modUtils.AmountsMatch(Abs(trans1.NormalizedAmount), Abs(trans2.NormalizedAmount), amountTolerance) Then
        IsTransferMatch = False
        Exit Function
    End If

    ' Must be opposite signs
    If (trans1.NormalizedAmount > 0 And trans2.NormalizedAmount > 0) Or _
       (trans1.NormalizedAmount < 0 And trans2.NormalizedAmount < 0) Then
        IsTransferMatch = False
        Exit Function
    End If

    ' Must be different accounts
    If trans1.Account = trans2.Account And trans1.Account <> "" Then
        IsTransferMatch = False
        Exit Function
    End If

    ' Check for transfer keywords in description (optional, increases confidence)
    Dim hasTransferKeyword As Boolean
    hasTransferKeyword = modUtils.ContainsAny(trans1.Description & " " & trans2.Description, _
                        "TRANSFER,TRANSFERENCIA,TRANSF,TRASFERIMENTO,INTERNAL,ACCOUNT TRANSFER")

    ' If no keywords, require exact amount match
    If Not hasTransferKeyword Then
        If Abs(trans1.NormalizedAmount + trans2.NormalizedAmount) > 0.01 Then
            IsTransferMatch = False
            Exit Function
        End If
    End If

    IsTransferMatch = True
End Function

' ========================================================================
' Get Transaction from Worksheet Row
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

    GetTransactionFromRow = trans
End Function

' ========================================================================
' Find Transfer Partner
' ========================================================================
Public Function FindTransferPartner(transactionID As String) As String
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim transferID As String

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_TRANSACTIONS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Find transaction and get its TransferID
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = transactionID Then
            transferID = ws.Cells(i, 9).Value
            Exit For
        End If
    Next i

    ' Find partner with same TransferID
    If transferID <> "" Then
        For i = 2 To lastRow
            If ws.Cells(i, 9).Value = transferID And ws.Cells(i, 1).Value <> transactionID Then
                FindTransferPartner = ws.Cells(i, 1).Value
                Exit Function
            End If
        Next i
    End If

    FindTransferPartner = ""
End Function

' ========================================================================
' Clear All Transfer Detections
' ========================================================================
Public Sub ClearAllTransfers()
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_TRANSACTIONS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Clear TransferID column
    If lastRow > 1 Then
        ws.Range("I2:I" & lastRow).ClearContents
    End If

    modUtils.LogMessage "All transfer detections cleared", "INFO"
End Sub

' ========================================================================
' Get Transfer Statistics
' ========================================================================
Public Function GetTransferStats() As String
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim transferCount As Long
    Dim uniqueTransfers As Object

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_TRANSACTIONS)
    Set uniqueTransfers = CreateObject("Scripting.Dictionary")

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        If ws.Cells(i, 9).Value <> "" Then
            If Not uniqueTransfers.exists(ws.Cells(i, 9).Value) Then
                uniqueTransfers.Add ws.Cells(i, 9).Value, 1
            End If
        End If
    Next i

    GetTransferStats = "Transfers Detected: " & uniqueTransfers.Count
End Function

' ========================================================================
' Manual Transfer Link
' ========================================================================
Public Sub LinkTransactions(trans1ID As String, trans2ID As String)
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim transferID As String
    Dim row1 As Long, row2 As Long

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_TRANSACTIONS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Find rows
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = trans1ID Then row1 = i
        If ws.Cells(i, 1).Value = trans2ID Then row2 = i
    Next i

    If row1 = 0 Or row2 = 0 Then Exit Sub

    ' Generate transfer ID
    transferID = modUtils.GenerateUniqueID("TRF")

    ' Link transactions
    ws.Cells(row1, 9).Value = transferID
    ws.Cells(row2, 9).Value = transferID

    ' Set categories
    ws.Cells(row1, 5).Value = "Transfer"
    ws.Cells(row1, 6).Value = "Internal Transfer"
    ws.Cells(row2, 5).Value = "Transfer"
    ws.Cells(row2, 6).Value = "Internal Transfer"

    modUtils.LogMessage "Manual transfer link created: " & trans1ID & " <-> " & trans2ID, "INFO"
End Sub

' ========================================================================
' Unlink Transfer
' ========================================================================
Public Sub UnlinkTransfer(transactionID As String)
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim transferID As String

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_TRANSACTIONS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Find transaction and get TransferID
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = transactionID Then
            transferID = ws.Cells(i, 9).Value
            Exit For
        End If
    Next i

    If transferID = "" Then Exit Sub

    ' Clear TransferID for all transactions with this ID
    For i = 2 To lastRow
        If ws.Cells(i, 9).Value = transferID Then
            ws.Cells(i, 9).Value = ""
            ' Clear transfer categories
            If ws.Cells(i, 5).Value = "Transfer" Then
                ws.Cells(i, 5).Value = ""
                ws.Cells(i, 6).Value = ""
            End If
        End If
    Next i

    modUtils.LogMessage "Transfer unlinked: " & transferID, "INFO"
End Sub
