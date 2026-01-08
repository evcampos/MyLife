Attribute VB_Name = "modNormalize"
' ========================================================================
' Module: modNormalize
' Purpose: Data Normalization Engine
' Description: Standardizes raw transaction data including dates, amounts,
'              descriptions, and handles duplicate detection
' Platform: Excel for MacOS
' Author: Financial Import System v1.0
' ========================================================================

Option Explicit

' Normalized transaction structure
Public Type NormalizedTransaction
    TransactionID As String
    TransactionDate As Date
    Description As String
    OriginalDescription As String
    NormalizedAmount As Double
    Account As String
    Institution As String
    TransactionType As String
    Category As String
    Subcategory As String
    TransferID As String
    InvestmentType As String
    TransactionHash As String
    SourceFile As String
    IsDuplicate As Boolean
    ErrorFlag As Boolean
    ErrorMessage As String
End Type

' ========================================================================
' Main Normalization Function
' ========================================================================
Public Function NormalizeAllData() As Long
    On Error GoTo ErrorHandler

    Dim rawWs As Worksheet
    Dim transWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim trans As NormalizedTransaction
    Dim normalizedCount As Long
    Dim duplicateHashes As Object

    modUtils.LogMessage "Starting data normalization", "INFO"
    modUtils.ShowProgress "Normalizing data..."

    ' Get worksheets
    Set rawWs = ThisWorkbook.Worksheets(modConfig.SHEET_RAWDATA)
    Set transWs = GetOrCreateTransactionsSheet()

    ' Clear existing transactions
    ClearTransactionsSheet transWs

    ' Initialize duplicate tracking
    Set duplicateHashes = CreateObject("Scripting.Dictionary")

    ' Get row count
    lastRow = rawWs.Cells(rawWs.Rows.Count, 1).End(xlUp).Row

    normalizedCount = 0

    ' Process each raw transaction
    For i = 2 To lastRow
        modUtils.ShowProgress "Normalizing data...", i - 1, lastRow - 1

        ' Normalize transaction
        trans = NormalizeTransaction(rawWs, i)

        ' Check for duplicates
        If modConfig.GetConfigBool("SkipDuplicates", True) Then
            If duplicateHashes.exists(trans.TransactionHash) Then
                trans.IsDuplicate = True
                modUtils.LogMessage "Duplicate detected: " & trans.Description, "DEBUG"
                ' Skip duplicate
                GoTo NextTransaction
            Else
                duplicateHashes.Add trans.TransactionHash, True
            End If
        End If

        ' Write normalized transaction
        Call WriteNormalizedTransaction(transWs, trans)
        normalizedCount = normalizedCount + 1

NextTransaction:
    Next i

    modUtils.ClearProgress
    modUtils.LogMessage "Normalization complete: " & normalizedCount & " transactions", "INFO"

    NormalizeAllData = normalizedCount

    Exit Function

ErrorHandler:
    modUtils.LogMessage "Error in NormalizeAllData: " & Err.Description, "ERROR"
    modUtils.ClearProgress
    NormalizeAllData = 0
End Function

' ========================================================================
' Normalize Single Transaction
' ========================================================================
Private Function NormalizeTransaction(ws As Worksheet, row As Long) As NormalizedTransaction
    On Error Resume Next

    Dim trans As NormalizedTransaction

    ' Initialize
    trans.TransactionID = modUtils.GenerateUniqueID("TXN")
    trans.ErrorFlag = False
    trans.IsDuplicate = False

    ' Read raw data
    trans.TransactionDate = ws.Cells(row, 1).Value
    trans.OriginalDescription = ws.Cells(row, 2).Value
    trans.NormalizedAmount = ws.Cells(row, 3).Value
    trans.Account = ws.Cells(row, 4).Value
    trans.TransactionType = ws.Cells(row, 5).Value
    trans.SourceFile = ws.Cells(row, 6).Value
    trans.TransactionHash = ws.Cells(row, 8).Value

    ' Normalize date
    If Not modUtils.IsValidDate(trans.TransactionDate) Then
        trans.ErrorFlag = True
        trans.ErrorMessage = "Invalid date"
    End If

    ' Normalize description
    trans.Description = NormalizeDescription(trans.OriginalDescription)

    ' Normalize amount
    trans.NormalizedAmount = modUtils.RoundCurrency(trans.NormalizedAmount)

    ' Extract institution from file name
    trans.Institution = ExtractInstitution(trans.SourceFile)

    ' Normalize account
    trans.Account = NormalizeAccount(trans.Account)

    ' Initialize empty fields (to be filled by classification)
    trans.Category = ""
    trans.Subcategory = ""
    trans.TransferID = ""
    trans.InvestmentType = ""

    NormalizeTransaction = trans
End Function

' ========================================================================
' Normalize Description
' ========================================================================
Private Function NormalizeDescription(description As String) As String
    On Error Resume Next

    Dim result As String

    result = Trim(description)

    ' Remove excessive whitespace
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop

    ' Remove common noise words (optional)
    ' result = Replace(result, "COMPRA ", "")
    ' result = Replace(result, "PAGAMENTO ", "")

    ' Keep original case for readability
    NormalizeDescription = result
End Function

' ========================================================================
' Extract Institution from File Name
' ========================================================================
Private Function ExtractInstitution(fileName As String) As String
    On Error Resume Next

    Dim result As String

    result = modUtils.GetFileName(fileName)

    ' Remove extension
    If InStr(result, ".") > 0 Then
        result = Left(result, InStrRev(result, ".") - 1)
    End If

    ' Extract institution from common patterns
    ' Example: "Bank_Statement_2024.csv" -> "Bank"
    If InStr(result, "_") > 0 Then
        result = Split(result, "_")(0)
    ElseIf InStr(result, "-") > 0 Then
        result = Split(result, "-")(0)
    End If

    ' Common institution name mapping
    result = UCase(Trim(result))

    Select Case True
        Case InStr(result, "CHASE") > 0
            ExtractInstitution = "Chase Bank"
        Case InStr(result, "CITI") > 0
            ExtractInstitution = "Citibank"
        Case InStr(result, "BOFA") > 0 Or InStr(result, "BANKOFAMERICA") > 0
            ExtractInstitution = "Bank of America"
        Case InStr(result, "WELLS") > 0
            ExtractInstitution = "Wells Fargo"
        Case InStr(result, "AMEX") > 0 Or InStr(result, "AMERICAN") > 0
            ExtractInstitution = "American Express"
        Case Else
            ExtractInstitution = result
    End Select
End Function

' ========================================================================
' Normalize Account
' ========================================================================
Private Function NormalizeAccount(account As String) As String
    On Error Resume Next

    Dim result As String

    result = Trim(account)

    ' Mask account numbers for privacy (keep last 4 digits)
    If Len(result) > 4 And IsNumeric(result) Then
        result = "****" & Right(result, 4)
    End If

    NormalizeAccount = result
End Function

' ========================================================================
' Get or Create Transactions Sheet
' ========================================================================
Private Function GetOrCreateTransactionsSheet() As Worksheet
    On Error Resume Next

    Dim ws As Worksheet

    If Not modConfig.SheetExists(modConfig.SHEET_TRANSACTIONS) Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = modConfig.SHEET_TRANSACTIONS
        Call InitializeTransactionsSheet(ws)
    Else
        Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_TRANSACTIONS)
    End If

    Set GetOrCreateTransactionsSheet = ws
End Function

' ========================================================================
' Initialize Transactions Sheet
' ========================================================================
Private Sub InitializeTransactionsSheet(ws As Worksheet)
    On Error Resume Next

    With ws
        ' Setup headers
        .Cells(1, 1).Value = "TransactionID"
        .Cells(1, 2).Value = "Date"
        .Cells(1, 3).Value = "Description"
        .Cells(1, 4).Value = "Amount"
        .Cells(1, 5).Value = "Category"
        .Cells(1, 6).Value = "Subcategory"
        .Cells(1, 7).Value = "Account"
        .Cells(1, 8).Value = "Institution"
        .Cells(1, 9).Value = "TransferID"
        .Cells(1, 10).Value = "InvestmentType"
        .Cells(1, 11).Value = "TransactionType"
        .Cells(1, 12).Value = "OriginalDescription"
        .Cells(1, 13).Value = "SourceFile"
        .Cells(1, 14).Value = "Hash"

        ' Format headers
        .Range("A1:N1").Font.Bold = True
        .Range("A1:N1").Interior.Color = RGB(68, 114, 196)
        .Range("A1:N1").Font.Color = RGB(255, 255, 255)

        ' Format columns
        .Columns("B:B").NumberFormat = "mm/dd/yyyy" ' Date
        .Columns("D:D").NumberFormat = "$#,##0.00"  ' Amount

        ' Auto-fit
        .Columns("A:N").AutoFit

        ' Freeze header row
        .Rows(2).Select
        ActiveWindow.FreezePanes = True
        .Cells(1, 1).Select
    End With
End Sub

' ========================================================================
' Clear Transactions Sheet
' ========================================================================
Private Sub ClearTransactionsSheet(ws As Worksheet)
    On Error Resume Next

    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow > 1 Then
        ws.Range("A2:N" & lastRow).ClearContents
    End If
End Sub

' ========================================================================
' Write Normalized Transaction
' ========================================================================
Private Sub WriteNormalizedTransaction(ws As Worksheet, trans As NormalizedTransaction)
    On Error Resume Next

    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    With ws
        .Cells(lastRow, 1).Value = trans.TransactionID
        .Cells(lastRow, 2).Value = trans.TransactionDate
        .Cells(lastRow, 3).Value = trans.Description
        .Cells(lastRow, 4).Value = trans.NormalizedAmount
        .Cells(lastRow, 5).Value = trans.Category
        .Cells(lastRow, 6).Value = trans.Subcategory
        .Cells(lastRow, 7).Value = trans.Account
        .Cells(lastRow, 8).Value = trans.Institution
        .Cells(lastRow, 9).Value = trans.TransferID
        .Cells(lastRow, 10).Value = trans.InvestmentType
        .Cells(lastRow, 11).Value = trans.TransactionType
        .Cells(lastRow, 12).Value = trans.OriginalDescription
        .Cells(lastRow, 13).Value = trans.SourceFile
        .Cells(lastRow, 14).Value = trans.TransactionHash
    End With
End Sub

' ========================================================================
' Update Transaction Field
' ========================================================================
Public Sub UpdateTransactionField(transactionID As String, fieldName As String, value As Variant)
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim colIndex As Integer

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_TRANSACTIONS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Find column index
    Select Case UCase(fieldName)
        Case "CATEGORY": colIndex = 5
        Case "SUBCATEGORY": colIndex = 6
        Case "TRANSFERID": colIndex = 9
        Case "INVESTMENTTYPE": colIndex = 10
        Case Else: Exit Sub
    End Select

    ' Find transaction and update
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = transactionID Then
            ws.Cells(i, colIndex).Value = value
            Exit Sub
        End If
    Next i
End Sub

' ========================================================================
' Get All Transactions
' ========================================================================
Public Function GetAllTransactions() As Collection
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim trans As NormalizedTransaction
    Dim transactions As Collection

    Set transactions = New Collection
    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_TRANSACTIONS)

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        trans.TransactionID = ws.Cells(i, 1).Value
        trans.TransactionDate = ws.Cells(i, 2).Value
        trans.Description = ws.Cells(i, 3).Value
        trans.NormalizedAmount = ws.Cells(i, 4).Value
        trans.Category = ws.Cells(i, 5).Value
        trans.Subcategory = ws.Cells(i, 6).Value
        trans.Account = ws.Cells(i, 7).Value
        trans.Institution = ws.Cells(i, 8).Value
        trans.TransferID = ws.Cells(i, 9).Value
        trans.InvestmentType = ws.Cells(i, 10).Value
        trans.TransactionType = ws.Cells(i, 11).Value
        trans.OriginalDescription = ws.Cells(i, 12).Value
        trans.SourceFile = ws.Cells(i, 13).Value
        trans.TransactionHash = ws.Cells(i, 14).Value

        transactions.Add trans
    Next i

    Set GetAllTransactions = transactions
End Function

' ========================================================================
' Get Transaction by ID
' ========================================================================
Public Function GetTransactionByID(transactionID As String) As NormalizedTransaction
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim trans As NormalizedTransaction

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_TRANSACTIONS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = transactionID Then
            trans.TransactionID = ws.Cells(i, 1).Value
            trans.TransactionDate = ws.Cells(i, 2).Value
            trans.Description = ws.Cells(i, 3).Value
            trans.NormalizedAmount = ws.Cells(i, 4).Value
            trans.Category = ws.Cells(i, 5).Value
            trans.Subcategory = ws.Cells(i, 6).Value
            trans.Account = ws.Cells(i, 7).Value
            trans.Institution = ws.Cells(i, 8).Value
            trans.TransferID = ws.Cells(i, 9).Value
            trans.InvestmentType = ws.Cells(i, 10).Value
            trans.TransactionType = ws.Cells(i, 11).Value
            trans.OriginalDescription = ws.Cells(i, 12).Value
            trans.SourceFile = ws.Cells(i, 13).Value
            trans.TransactionHash = ws.Cells(i, 14).Value

            GetTransactionByID = trans
            Exit Function
        End If
    Next i
End Function

' ========================================================================
' Get Transaction Statistics
' ========================================================================
Public Function GetNormalizationStats() As String
    On Error Resume Next

    Dim ws As Worksheet
    Dim totalTrans As Long
    Dim totalAmount As Double
    Dim i As Long
    Dim stats As String

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_TRANSACTIONS)
    totalTrans = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row - 1

    For i = 2 To totalTrans + 1
        totalAmount = totalAmount + ws.Cells(i, 4).Value
    Next i

    stats = "Total Transactions: " & totalTrans & vbCrLf & _
            "Total Amount: " & Format(totalAmount, "$#,##0.00")

    GetNormalizationStats = stats
End Function
