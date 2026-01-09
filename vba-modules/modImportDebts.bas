Attribute VB_Name = "modImportDebts"
' ========================================================================
' Module: modImportDebts
' Purpose: Import Debt Information
' Description: Imports debt tracking data with interest rates
' Platform: Excel for MacOS
' Author: [MY LIFE] Financial System v1.0
' ========================================================================

Option Explicit

' ========================================================================
' Main Import Function - Import Debts
' ========================================================================
Public Function ImportDebts() As Long
    On Error GoTo ErrorHandler

    Dim filePath As String
    Dim totalImported As Long

    filePath = modConfig.GetFilePath("DEBTS")

    If filePath = "" Then
        modUtils.LogDebug "No debts file path configured"
        ImportDebts = 0
        Exit Function
    End If

    modUtils.ShowProgress "Importing debts..."
    modUtils.LogDebug "Starting debts import from: " & filePath

    If modUtils.FileExists(filePath) Then
        totalImported = ImportDebtsFile(filePath)
    Else
        modUtils.LogDebug "Debts file not found: " & filePath
        MsgBox "Debts file not found: " & filePath, vbExclamation
        totalImported = 0
    End If

    modUtils.ClearProgress
    modUtils.LogDebug "Debts import complete: " & totalImported & " records"

    ImportDebts = totalImported

    Exit Function

ErrorHandler:
    modUtils.ClearProgress
    MsgBox modUtils.FormatErrorMessage("modImportDebts", "ImportDebts", Err.Description), vbCritical
    ImportDebts = 0
End Function

' ========================================================================
' Import Debts File
' ========================================================================
Private Function ImportDebtsFile(filePath As String) As Long
    On Error GoTo ErrorHandler

    Dim extension As String
    Dim imported As Long

    extension = modUtils.GetFileExtension(filePath)

    Select Case extension
        Case "csv", "txt"
            imported = ImportDebtsCSV(filePath)
        Case "xlsx", "xls", "xlsm"
            imported = ImportDebtsExcel(filePath)
        Case Else
            MsgBox "Unsupported file format for debts: " & extension, vbExclamation
            imported = 0
    End Select

    ImportDebtsFile = imported

    Exit Function

ErrorHandler:
    MsgBox modUtils.FormatErrorMessage("modImportDebts", "ImportDebtsFile", Err.Description), vbCritical
    ImportDebtsFile = 0
End Function

' ========================================================================
' Import CSV Debts File
' ========================================================================
Private Function ImportDebtsCSV(filePath As String) As Long
    On Error GoTo ErrorHandler

    Dim fileNum As Integer
    Dim lineText As String
    Dim delimiter As String
    Dim headers() As String
    Dim rowData() As String
    Dim ws As Worksheet
    Dim imported As Long

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_DEBTS)

    ' Open file
    fileNum = FreeFile
    Open filePath For Input As #fileNum

    ' Read header line
    Line Input #fileNum, lineText

    ' Detect delimiter
    delimiter = DetectDelimiter(lineText)
    headers = Split(lineText, delimiter)

    ' Map columns
    Dim colCreditor As Integer, colRate As Integer, colPaid As Integer
    Dim colUpdated As Integer, colCurrency As Integer, colDate As Integer

    colCreditor = FindColumn(headers, "CREDITOR,CREDOR,BANK,BANCO,LENDER")
    colRate = FindColumn(headers, "INTERESTRATE,INTEREST_RATE,RATE,TAXA")
    colPaid = FindColumn(headers, "AMOUNTPAID,AMOUNT_PAID,VALOR_PAGO,INITIAL")
    colUpdated = FindColumn(headers, "UPDATEDAMOUNT,UPDATED_AMOUNT,CURRENT,VALOR_ATUALIZADO")
    colCurrency = FindColumn(headers, "CURRENCY,MOEDA")
    colDate = FindColumn(headers, "LASTUPDATE,LAST_UPDATE,DATE,DATA")

    If colCreditor = -1 Or colRate = -1 Or colPaid = -1 Then
        Close #fileNum
        MsgBox "Could not detect required columns in debts file", vbExclamation
        ImportDebtsCSV = 0
        Exit Function
    End If

    ' Read data rows
    imported = 0
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText

        If Trim(lineText) <> "" Then
            rowData = Split(lineText, delimiter)

            If ParseAndAddDebtRow(ws, rowData, colCreditor, colRate, colPaid, _
                                 colUpdated, colCurrency, colDate) Then
                imported = imported + 1
            End If
        End If
    Loop

    Close #fileNum

    ImportDebtsCSV = imported

    Exit Function

ErrorHandler:
    On Error Resume Next
    Close #fileNum
    MsgBox modUtils.FormatErrorMessage("modImportDebts", "ImportDebtsCSV", Err.Description), vbCritical
    ImportDebtsCSV = 0
End Function

' ========================================================================
' Import Excel Debts File
' ========================================================================
Private Function ImportDebtsExcel(filePath As String) As Long
    On Error GoTo ErrorHandler

    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim destWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim imported As Long
    Dim headers() As Variant
    Dim colCount As Integer

    Set destWs = ThisWorkbook.Worksheets(modConfig.SHEET_DEBTS)

    ' Open source workbook
    Application.ScreenUpdating = False
    Set sourceWb = Workbooks.Open(filePath, ReadOnly:=True)
    Set sourceWs = sourceWb.Worksheets(1)

    ' Read headers
    lastRow = sourceWs.Cells(sourceWs.Rows.Count, 1).End(xlUp).row
    colCount = sourceWs.Cells(1, sourceWs.Columns.Count).End(xlToLeft).Column

    ReDim headers(1 To colCount)
    For i = 1 To colCount
        headers(i) = CStr(sourceWs.Cells(1, i).Value)
    Next i

    ' Map columns
    Dim colCreditor As Integer, colRate As Integer, colPaid As Integer
    Dim colUpdated As Integer, colCurrency As Integer, colDate As Integer

    colCreditor = FindColumn(headers, "CREDITOR,CREDOR,BANK,BANCO,LENDER")
    colRate = FindColumn(headers, "INTERESTRATE,INTEREST_RATE,RATE,TAXA")
    colPaid = FindColumn(headers, "AMOUNTPAID,AMOUNT_PAID,VALOR_PAGO,INITIAL")
    colUpdated = FindColumn(headers, "UPDATEDAMOUNT,UPDATED_AMOUNT,CURRENT,VALOR_ATUALIZADO")
    colCurrency = FindColumn(headers, "CURRENCY,MOEDA")
    colDate = FindColumn(headers, "LASTUPDATE,LAST_UPDATE,DATE,DATA")

    If colCreditor = -1 Or colRate = -1 Or colPaid = -1 Then
        sourceWb.Close SaveChanges:=False
        Application.ScreenUpdating = True
        MsgBox "Could not detect required columns in debts file", vbExclamation
        ImportDebtsExcel = 0
        Exit Function
    End If

    ' Import data rows
    imported = 0
    For i = 2 To lastRow
        Dim rowData() As Variant
        ReDim rowData(1 To colCount)

        Dim j As Integer
        For j = 1 To colCount
            rowData(j) = sourceWs.Cells(i, j).Value
        Next j

        If ParseAndAddDebtRow(destWs, rowData, colCreditor, colRate, colPaid, _
                             colUpdated, colCurrency, colDate) Then
            imported = imported + 1
        End If
    Next i

    sourceWb.Close SaveChanges:=False
    Application.ScreenUpdating = True

    ImportDebtsExcel = imported

    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not sourceWb Is Nothing Then sourceWb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    MsgBox modUtils.FormatErrorMessage("modImportDebts", "ImportDebtsExcel", Err.Description), vbCritical
    ImportDebtsExcel = 0
End Function

' ========================================================================
' Parse and Add Debt Row
' ========================================================================
Private Function ParseAndAddDebtRow(ws As Worksheet, rowData As Variant, _
                                    colCreditor As Integer, colRate As Integer, _
                                    colPaid As Integer, colUpdated As Integer, _
                                    colCurrency As Integer, colDate As Integer) As Boolean
    On Error GoTo ErrorHandler

    Dim lastRow As Long
    Dim creditor As String
    Dim interestRate As String
    Dim amountPaid As Double
    Dim updatedAmount As Double
    Dim currency As String
    Dim lastUpdate As Date

    ' Get values
    creditor = GetColumnValue(rowData, colCreditor)
    interestRate = GetColumnValue(rowData, colRate)
    amountPaid = CDbl(GetColumnValue(rowData, colPaid))

    ' Optional fields
    If colUpdated > 0 Then
        Dim updatedVal As String
        updatedVal = GetColumnValue(rowData, colUpdated)
        If updatedVal <> "" Then
            updatedAmount = CDbl(updatedVal)
        Else
            updatedAmount = amountPaid
        End If
    Else
        updatedAmount = amountPaid
    End If

    If colCurrency > 0 Then
        currency = GetColumnValue(rowData, colCurrency)
    Else
        currency = "BRL"  ' Default currency
    End If

    If colDate > 0 Then
        Dim dateVal As String
        dateVal = GetColumnValue(rowData, colDate)
        If dateVal <> "" Then
            lastUpdate = modUtils.ParseDate(dateVal)
        Else
            lastUpdate = Date
        End If
    Else
        lastUpdate = Date
    End If

    ' Validate
    If creditor = "" Then
        ParseAndAddDebtRow = False
        Exit Function
    End If

    ' Add to sheet
    lastRow = modUtils.GetLastRow(ws, 1) + 1

    ws.Cells(lastRow, 1).Value = creditor
    ws.Cells(lastRow, 2).Value = interestRate
    ws.Cells(lastRow, 3).Value = amountPaid
    ws.Cells(lastRow, 4).Value = updatedAmount
    ws.Cells(lastRow, 5).Value = currency
    ws.Cells(lastRow, 6).Value = lastUpdate
    ws.Cells(lastRow, 7).Value = Now ' ImportDate

    ParseAndAddDebtRow = True

    Exit Function

ErrorHandler:
    ParseAndAddDebtRow = False
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

    If tabCount > commaCount And tabCount > semicolonCount Then
        DetectDelimiter = vbTab
    ElseIf semicolonCount > commaCount Then
        DetectDelimiter = ";"
    Else
        DetectDelimiter = ","
    End If
End Function

' ========================================================================
' Find Column Index
' ========================================================================
Private Function FindColumn(headers As Variant, possibleNames As String) As Integer
    On Error Resume Next

    Dim i As Integer
    Dim names() As String
    Dim name As Variant

    names = Split(possibleNames, ",")

    For i = LBound(headers) To UBound(headers)
        For Each name In names
            If modUtils.ContainsAny(CStr(headers(i)), CStr(name)) Then
                FindColumn = i
                Exit Function
            End If
        Next name
    Next i

    FindColumn = -1
End Function

' ========================================================================
' Get Column Value
' ========================================================================
Private Function GetColumnValue(rowData As Variant, colIndex As Integer) As String
    On Error Resume Next

    If colIndex > 0 And colIndex <= UBound(rowData) Then
        GetColumnValue = Trim(CStr(rowData(colIndex)))
    Else
        GetColumnValue = ""
    End If
End Function
