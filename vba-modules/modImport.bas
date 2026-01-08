Attribute VB_Name = "modImport"
' ========================================================================
' Module: modImport
' Purpose: File Import Engine
' Description: Handles importing data from multiple file formats including
'              CSV, Excel, and text files with intelligent column detection
' Platform: Excel for MacOS
' Author: Financial Import System v1.0
' ========================================================================

Option Explicit

' Import result structure
Public Type ImportResult
    Success As Boolean
    RowsImported As Long
    ErrorMessage As String
    SourceFile As String
    ImportDate As Date
End Type

' Column mapping structure
Private Type ColumnMap
    DateCol As Integer
    DescriptionCol As Integer
    AmountCol As Integer
    DebitCol As Integer
    CreditCol As Integer
    AccountCol As Integer
    TransactionTypeCol As Integer
    BalanceCol As Integer
End Type

' ========================================================================
' Main Import Function - Import All Configured Files
' ========================================================================
Public Function ImportAllFiles() As Collection
    On Error GoTo ErrorHandler

    Dim results As Collection
    Dim paths As Collection
    Dim path As Variant
    Dim result As ImportResult

    Set results = New Collection
    Set paths = modConfig.GetImportPaths()

    modUtils.LogMessage "Starting import process", "INFO"
    modUtils.ShowProgress "Initializing import..."

    ' Initialize RawData sheet
    Call InitializeRawDataSheet

    ' Import each configured path
    For Each path In paths
        modUtils.ShowProgress "Importing: " & modUtils.GetFileName(CStr(path))

        result = ImportFile(CStr(path))
        results.Add result

        If result.Success Then
            modUtils.LogMessage "Successfully imported: " & result.SourceFile & _
                               " (" & result.RowsImported & " rows)", "INFO"
        Else
            modUtils.LogMessage "Failed to import: " & result.SourceFile & _
                               " - " & result.ErrorMessage, "ERROR"
        End If
    Next path

    modUtils.ClearProgress
    Set ImportAllFiles = results

    Exit Function

ErrorHandler:
    modUtils.LogMessage "Error in ImportAllFiles: " & Err.Description, "ERROR"
    modUtils.ClearProgress
    Set ImportAllFiles = New Collection
End Function

' ========================================================================
' Import Single File
' ========================================================================
Public Function ImportFile(filePath As String) As ImportResult
    On Error GoTo ErrorHandler

    Dim result As ImportResult
    Dim extension As String

    result.SourceFile = filePath
    result.ImportDate = Now
    result.Success = False
    result.RowsImported = 0

    ' Check file exists
    If Dir(filePath) = "" Then
        result.ErrorMessage = "File not found"
        ImportFile = result
        Exit Function
    End If

    ' Determine file type and call appropriate parser
    extension = modUtils.GetFileExtension(filePath)

    Select Case extension
        Case "csv", "txt"
            result = ImportCSVFile(filePath)
        Case "xlsx", "xls", "xlsm"
            result = ImportExcelFile(filePath)
        Case Else
            result.ErrorMessage = "Unsupported file format: " & extension
    End Select

    ImportFile = result
    Exit Function

ErrorHandler:
    result.Success = False
    result.ErrorMessage = "Import error: " & Err.Description
    ImportFile = result
End Function

' ========================================================================
' Import CSV File
' ========================================================================
Private Function ImportCSVFile(filePath As String) As ImportResult
    On Error GoTo ErrorHandler

    Dim result As ImportResult
    Dim fileNum As Integer
    Dim lineText As String
    Dim delimiter As String
    Dim headers() As String
    Dim rowData() As String
    Dim colMap As ColumnMap
    Dim rowCount As Long
    Dim ws As Worksheet

    result.SourceFile = filePath
    result.ImportDate = Now
    result.Success = False
    result.RowsImported = 0

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_RAWDATA)

    ' Open file
    fileNum = FreeFile
    Open filePath For Input As #fileNum

    ' Read header line
    Line Input #fileNum, lineText

    ' Detect delimiter
    delimiter = DetectDelimiter(lineText)
    headers = Split(lineText, delimiter)

    ' Map columns
    colMap = MapColumns(headers)

    ' Validate column mapping
    If Not ValidateColumnMap(colMap) Then
        result.ErrorMessage = "Could not detect required columns (Date, Description, Amount)"
        Close #fileNum
        ImportCSVFile = result
        Exit Function
    End If

    ' Read data rows
    rowCount = 0
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText

        ' Skip empty lines
        If Trim(lineText) <> "" Then
            rowData = Split(lineText, delimiter)

            ' Parse and add row to RawData
            If ParseAndAddRow(ws, rowData, colMap, modUtils.GetFileName(filePath)) Then
                rowCount = rowCount + 1
            End If
        End If
    Loop

    Close #fileNum

    result.Success = True
    result.RowsImported = rowCount
    ImportCSVFile = result

    Exit Function

ErrorHandler:
    On Error Resume Next
    Close #fileNum
    result.Success = False
    result.ErrorMessage = "CSV import error: " & Err.Description
    ImportCSVFile = result
End Function

' ========================================================================
' Import Excel File
' ========================================================================
Private Function ImportExcelFile(filePath As String) As ImportResult
    On Error GoTo ErrorHandler

    Dim result As ImportResult
    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim destWs As Worksheet
    Dim headers() As String
    Dim colMap As ColumnMap
    Dim lastRow As Long
    Dim i As Long
    Dim rowCount As Long
    Dim colCount As Integer
    Dim rowData() As String

    result.SourceFile = filePath
    result.ImportDate = Now
    result.Success = False
    result.RowsImported = 0

    Set destWs = ThisWorkbook.Worksheets(modConfig.SHEET_RAWDATA)

    ' Open source workbook
    Application.ScreenUpdating = False
    Set sourceWb = Workbooks.Open(filePath, ReadOnly:=True)
    Set sourceWs = sourceWb.Worksheets(1) ' Assume first sheet

    ' Read headers
    lastRow = sourceWs.Cells(sourceWs.Rows.Count, 1).End(xlUp).Row
    colCount = sourceWs.Cells(1, sourceWs.Columns.Count).End(xlToLeft).Column

    ReDim headers(1 To colCount)
    For i = 1 To colCount
        headers(i) = CStr(sourceWs.Cells(1, i).Value)
    Next i

    ' Map columns
    colMap = MapColumns(headers)

    ' Validate column mapping
    If Not ValidateColumnMap(colMap) Then
        result.ErrorMessage = "Could not detect required columns"
        sourceWb.Close SaveChanges:=False
        Application.ScreenUpdating = True
        ImportExcelFile = result
        Exit Function
    End If

    ' Import data rows
    rowCount = 0
    For i = 2 To lastRow
        ReDim rowData(1 To colCount)

        ' Read row data
        Dim j As Integer
        For j = 1 To colCount
            rowData(j) = CStr(sourceWs.Cells(i, j).Value)
        Next j

        ' Parse and add row
        If ParseAndAddRow(destWs, rowData, colMap, modUtils.GetFileName(filePath)) Then
            rowCount = rowCount + 1
        End If
    Next i

    sourceWb.Close SaveChanges:=False
    Application.ScreenUpdating = True

    result.Success = True
    result.RowsImported = rowCount
    ImportExcelFile = result

    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not sourceWb Is Nothing Then sourceWb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    result.Success = False
    result.ErrorMessage = "Excel import error: " & Err.Description
    ImportExcelFile = result
End Function

' ========================================================================
' Detect CSV Delimiter
' ========================================================================
Private Function DetectDelimiter(headerLine As String) As String
    On Error Resume Next

    Dim commaCount As Integer
    Dim semicolonCount As Integer
    Dim tabCount As Integer

    commaCount = Len(headerLine) - Len(Replace(headerLine, ",", ""))
    semicolonCount = Len(headerLine) - Len(Replace(headerLine, ";", ""))
    tabCount = Len(headerLine) - Len(Replace(headerLine, vbTab, ""))

    ' Return most common delimiter
    If tabCount > commaCount And tabCount > semicolonCount Then
        DetectDelimiter = vbTab
    ElseIf semicolonCount > commaCount Then
        DetectDelimiter = ";"
    Else
        DetectDelimiter = ","
    End If
End Function

' ========================================================================
' Map Columns - Intelligent Column Detection
' ========================================================================
Private Function MapColumns(headers As Variant) As ColumnMap
    On Error Resume Next

    Dim colMap As ColumnMap
    Dim i As Integer
    Dim header As String

    ' Initialize
    colMap.DateCol = 0
    colMap.DescriptionCol = 0
    colMap.AmountCol = 0
    colMap.DebitCol = 0
    colMap.CreditCol = 0
    colMap.AccountCol = 0
    colMap.TransactionTypeCol = 0
    colMap.BalanceCol = 0

    ' Scan headers
    For i = LBound(headers) To UBound(headers)
        header = modUtils.CleanText(CStr(headers(i)))

        ' Date column
        If colMap.DateCol = 0 Then
            If modUtils.ContainsAny(header, "DATE,DATA,FECHA,DATUM") Then
                colMap.DateCol = i
            End If
        End If

        ' Description column
        If colMap.DescriptionCol = 0 Then
            If modUtils.ContainsAny(header, "DESCRIPTION,DESCRIP,HISTORIC,MEMO,DETAILS,DETALLE") Then
                colMap.DescriptionCol = i
            End If
        End If

        ' Amount column
        If colMap.AmountCol = 0 Then
            If modUtils.ContainsAny(header, "AMOUNT,VALOR,MONTO,IMPORTE") And _
               Not modUtils.ContainsAny(header, "BALANCE,SALDO") Then
                colMap.AmountCol = i
            End If
        End If

        ' Debit column
        If colMap.DebitCol = 0 Then
            If modUtils.ContainsAny(header, "DEBIT,DÉBITO,DEBITO,WITHDRAWAL,SAIDA") Then
                colMap.DebitCol = i
            End If
        End If

        ' Credit column
        If colMap.CreditCol = 0 Then
            If modUtils.ContainsAny(header, "CREDIT,CRÉDITO,CREDITO,DEPOSIT,ENTRADA") Then
                colMap.CreditCol = i
            End If
        End If

        ' Account column
        If colMap.AccountCol = 0 Then
            If modUtils.ContainsAny(header, "ACCOUNT,CONTA,CUENTA,CARD,TARJETA") Then
                colMap.AccountCol = i
            End If
        End If

        ' Transaction Type column
        If colMap.TransactionTypeCol = 0 Then
            If modUtils.ContainsAny(header, "TYPE,TIPO,TRANSACTION TYPE") Then
                colMap.TransactionTypeCol = i
            End If
        End If

        ' Balance column
        If colMap.BalanceCol = 0 Then
            If modUtils.ContainsAny(header, "BALANCE,SALDO") Then
                colMap.BalanceCol = i
            End If
        End If
    Next i

    MapColumns = colMap
End Function

' ========================================================================
' Validate Column Mapping
' ========================================================================
Private Function ValidateColumnMap(colMap As ColumnMap) As Boolean
    On Error Resume Next

    ' Must have at minimum: Date and Description
    ' Amount can be in single column or split into Debit/Credit
    ValidateColumnMap = (colMap.DateCol > 0 And colMap.DescriptionCol > 0 And _
                        (colMap.AmountCol > 0 Or (colMap.DebitCol > 0 Or colMap.CreditCol > 0)))
End Function

' ========================================================================
' Parse and Add Row to RawData
' ========================================================================
Private Function ParseAndAddRow(ws As Worksheet, rowData As Variant, _
                                colMap As ColumnMap, sourceFile As String) As Boolean
    On Error GoTo ErrorHandler

    Dim lastRow As Long
    Dim transDate As Date
    Dim description As String
    Dim amount As Double
    Dim debitAmount As Double
    Dim creditAmount As Double
    Dim account As String
    Dim transType As String

    ' Get values from row
    transDate = modUtils.ParseDate(GetColumnValue(rowData, colMap.DateCol))
    description = GetColumnValue(rowData, colMap.DescriptionCol)

    ' Get amount (handle both single amount and debit/credit columns)
    If colMap.AmountCol > 0 Then
        amount = modUtils.ExtractNumber(GetColumnValue(rowData, colMap.AmountCol))
    ElseIf colMap.DebitCol > 0 Or colMap.CreditCol > 0 Then
        debitAmount = modUtils.ExtractNumber(GetColumnValue(rowData, colMap.DebitCol))
        creditAmount = modUtils.ExtractNumber(GetColumnValue(rowData, colMap.CreditCol))

        If debitAmount <> 0 Then
            amount = -Abs(debitAmount)
            transType = "DEBIT"
        ElseIf creditAmount <> 0 Then
            amount = Abs(creditAmount)
            transType = "CREDIT"
        End If
    End If

    account = GetColumnValue(rowData, colMap.AccountCol)

    If colMap.TransactionTypeCol > 0 Then
        transType = GetColumnValue(rowData, colMap.TransactionTypeCol)
    End If

    ' Validate required fields
    If Not modUtils.IsValidDate(transDate) Or description = "" Then
        ParseAndAddRow = False
        Exit Function
    End If

    ' Skip balance lines (usually have zero amount and "SALDO" in description)
    If amount = 0 And modUtils.ContainsAny(description, "SALDO,BALANCE,TOTAL") Then
        ParseAndAddRow = False
        Exit Function
    End If

    ' Add to RawData sheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(lastRow, 1).Value = transDate
    ws.Cells(lastRow, 2).Value = description
    ws.Cells(lastRow, 3).Value = amount
    ws.Cells(lastRow, 4).Value = account
    ws.Cells(lastRow, 5).Value = transType
    ws.Cells(lastRow, 6).Value = sourceFile
    ws.Cells(lastRow, 7).Value = Now ' Import timestamp
    ws.Cells(lastRow, 8).Value = modUtils.GenerateTransactionHash(transDate, amount, description)

    ParseAndAddRow = True
    Exit Function

ErrorHandler:
    ParseAndAddRow = False
End Function

' ========================================================================
' Get Column Value Helper
' ========================================================================
Private Function GetColumnValue(rowData As Variant, colIndex As Integer) As String
    On Error Resume Next

    If colIndex > 0 And colIndex <= UBound(rowData) Then
        GetColumnValue = Trim(CStr(rowData(colIndex)))
    Else
        GetColumnValue = ""
    End If
End Function

' ========================================================================
' Initialize RawData Sheet
' ========================================================================
Private Sub InitializeRawDataSheet()
    On Error Resume Next

    Dim ws As Worksheet

    ' Create or clear RawData sheet
    If Not modConfig.SheetExists(modConfig.SHEET_RAWDATA) Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = modConfig.SHEET_RAWDATA
    Else
        Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_RAWDATA)
        ws.Cells.Clear
    End If

    ' Setup headers
    With ws
        .Cells(1, 1).Value = "Date"
        .Cells(1, 2).Value = "Description"
        .Cells(1, 3).Value = "Amount"
        .Cells(1, 4).Value = "Account"
        .Cells(1, 5).Value = "TransactionType"
        .Cells(1, 6).Value = "SourceFile"
        .Cells(1, 7).Value = "ImportTimestamp"
        .Cells(1, 8).Value = "TransactionHash"

        ' Format headers
        .Range("A1:H1").Font.Bold = True
        .Range("A1:H1").Interior.Color = RGB(68, 114, 196)
        .Range("A1:H1").Font.Color = RGB(255, 255, 255)
    End With

    ' Hide sheet
    ws.Visible = xlSheetVeryHidden
End Sub

' ========================================================================
' Get Import Statistics
' ========================================================================
Public Function GetImportStats() As String
    On Error Resume Next

    Dim ws As Worksheet
    Dim totalRows As Long
    Dim uniqueFiles As Object
    Dim i As Long
    Dim stats As String

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_RAWDATA)
    totalRows = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row - 1

    Set uniqueFiles = CreateObject("Scripting.Dictionary")

    For i = 2 To totalRows + 1
        If Not uniqueFiles.exists(ws.Cells(i, 6).Value) Then
            uniqueFiles.Add ws.Cells(i, 6).Value, 1
        End If
    Next i

    stats = "Total Rows: " & totalRows & vbCrLf & _
            "Source Files: " & uniqueFiles.Count

    GetImportStats = stats
End Function
