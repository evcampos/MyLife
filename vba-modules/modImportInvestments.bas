Attribute VB_Name = "modImportInvestments"
' ========================================================================
' Module: modImportInvestments
' Purpose: Import Investment Transactions
' Description: Imports investment transactions (applications/redemptions)
'              for correlation with bank movements
' Platform: Excel for MacOS
' Author: [MY LIFE] Financial System v1.0
' ========================================================================

Option Explicit

' ========================================================================
' Main Import Function - Import All Investment Sources
' ========================================================================
Public Function ImportAllInvestments() As Long
    On Error GoTo ErrorHandler

    Dim sources() As String
    Dim source As Variant
    Dim filePath As String
    Dim totalImported As Long

    sources = Array("ITAU_INV", "NUBANK_INV", "C6_INV", "BB_INV", "XP_INV", "BTG_INV")

    modUtils.ShowProgress "Importing investment transactions..."
    modUtils.LogDebug "Starting investment imports"

    For Each source In sources
        filePath = modConfig.GetFilePath(CStr(source))

        If filePath <> "" Then
            modUtils.LogDebug "Importing from: " & source

            If modUtils.FileExists(filePath) Then
                totalImported = totalImported + ImportInvestmentFile(CStr(source), filePath)
            Else
                modUtils.LogDebug "File not found: " & filePath
            End If
        End If
    Next source

    modUtils.ClearProgress
    modUtils.LogDebug "Investment imports complete: " & totalImported & " transactions"

    ImportAllInvestments = totalImported

    Exit Function

ErrorHandler:
    modUtils.ClearProgress
    MsgBox modUtils.FormatErrorMessage("modImportInvestments", "ImportAllInvestments", Err.Description), vbCritical
    ImportAllInvestments = 0
End Function

' ========================================================================
' Import Single Investment File
' ========================================================================
Private Function ImportInvestmentFile(sourceName As String, filePath As String) As Long
    On Error GoTo ErrorHandler

    Dim extension As String
    Dim imported As Long

    extension = modUtils.GetFileExtension(filePath)

    Select Case extension
        Case "csv", "txt"
            imported = ImportInvestmentCSV(sourceName, filePath)
        Case "xlsx", "xls", "xlsm"
            imported = ImportInvestmentExcel(sourceName, filePath)
        Case Else
            MsgBox "Unsupported file format for " & sourceName & ": " & extension, vbExclamation
            imported = 0
    End Select

    ImportInvestmentFile = imported

    Exit Function

ErrorHandler:
    MsgBox modUtils.FormatErrorMessage("modImportInvestments", "ImportInvestmentFile", Err.Description), vbCritical
    ImportInvestmentFile = 0
End Function

' ========================================================================
' Import CSV Investment File
' ========================================================================
Private Function ImportInvestmentCSV(sourceName As String, filePath As String) As Long
    On Error GoTo ErrorHandler

    Dim fileNum As Integer
    Dim lineText As String
    Dim delimiter As String
    Dim headers() As String
    Dim rowData() As String
    Dim ws As Worksheet
    Dim bankName As String
    Dim imported As Long

    ' Extract bank name from source
    bankName = Replace(Replace(sourceName, "_INV", ""), "_INVEST", "")

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_INVESTMENTS)

    ' Open file
    fileNum = FreeFile
    Open filePath For Input As #fileNum

    ' Read header line
    Line Input #fileNum, lineText

    ' Detect delimiter
    delimiter = DetectDelimiter(lineText)
    headers = Split(lineText, delimiter)

    ' Map columns
    Dim colDate As Integer, colDesc As Integer, colValue As Integer, colType As Integer

    colDate = FindColumn(headers, "DATE,DATA,TRANSACTION_DATE,MOVIMENTACAO")
    colDesc = FindColumn(headers, "DESCRIPTION,DESCRICAO,PRODUCT,PRODUTO,INVESTMENT")
    colValue = FindColumn(headers, "VALUE,VALOR,AMOUNT")
    colType = FindColumn(headers, "TYPE,TIPO,OPERATION,OPERACAO")

    If colDate = -1 Or colDesc = -1 Or colValue = -1 Then
        Close #fileNum
        MsgBox "Could not detect required columns in " & sourceName, vbExclamation
        ImportInvestmentCSV = 0
        Exit Function
    End If

    ' Read data rows
    imported = 0
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText

        If Trim(lineText) <> "" Then
            rowData = Split(lineText, delimiter)

            If ParseAndAddInvestmentRow(ws, bankName, rowData, colDate, colDesc, colValue, colType) Then
                imported = imported + 1
            End If
        End If
    Loop

    Close #fileNum

    ImportInvestmentCSV = imported

    Exit Function

ErrorHandler:
    On Error Resume Next
    Close #fileNum
    MsgBox modUtils.FormatErrorMessage("modImportInvestments", "ImportInvestmentCSV", Err.Description), vbCritical
    ImportInvestmentCSV = 0
End Function

' ========================================================================
' Import Excel Investment File
' ========================================================================
Private Function ImportInvestmentExcel(sourceName As String, filePath As String) As Long
    On Error GoTo ErrorHandler

    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim destWs As Worksheet
    Dim bankName As String
    Dim lastRow As Long
    Dim i As Long
    Dim imported As Long
    Dim headers() As Variant
    Dim colCount As Integer

    ' Extract bank name
    bankName = Replace(Replace(sourceName, "_INV", ""), "_INVEST", "")

    Set destWs = ThisWorkbook.Worksheets(modConfig.SHEET_INVESTMENTS)

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
    Dim colDate As Integer, colDesc As Integer, colValue As Integer, colType As Integer

    colDate = FindColumn(headers, "DATE,DATA,TRANSACTION_DATE,MOVIMENTACAO")
    colDesc = FindColumn(headers, "DESCRIPTION,DESCRICAO,PRODUCT,PRODUTO,INVESTMENT")
    colValue = FindColumn(headers, "VALUE,VALOR,AMOUNT")
    colType = FindColumn(headers, "TYPE,TIPO,OPERATION,OPERACAO")

    If colDate = -1 Or colDesc = -1 Or colValue = -1 Then
        sourceWb.Close SaveChanges:=False
        Application.ScreenUpdating = True
        MsgBox "Could not detect required columns in " & sourceName, vbExclamation
        ImportInvestmentExcel = 0
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

        If ParseAndAddInvestmentRow(destWs, bankName, rowData, colDate, colDesc, colValue, colType) Then
            imported = imported + 1
        End If
    Next i

    sourceWb.Close SaveChanges:=False
    Application.ScreenUpdating = True

    ImportInvestmentExcel = imported

    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not sourceWb Is Nothing Then sourceWb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    MsgBox modUtils.FormatErrorMessage("modImportInvestments", "ImportInvestmentExcel", Err.Description), vbCritical
    ImportInvestmentExcel = 0
End Function

' ========================================================================
' Parse and Add Investment Row
' ========================================================================
Private Function ParseAndAddInvestmentRow(ws As Worksheet, bankName As String, _
                                          rowData As Variant, colDate As Integer, _
                                          colDesc As Integer, colValue As Integer, _
                                          colType As Integer) As Boolean
    On Error GoTo ErrorHandler

    Dim lastRow As Long
    Dim transDate As Date
    Dim description As String
    Dim value As Double
    Dim transType As String

    ' Get values
    transDate = modUtils.ParseDate(GetColumnValue(rowData, colDate))
    description = GetColumnValue(rowData, colDesc)
    value = CDbl(GetColumnValue(rowData, colValue))

    ' Determine transaction type
    If colType > 0 Then
        transType = DetectInvestmentType(GetColumnValue(rowData, colType))
    Else
        transType = DetectInvestmentType(description)
    End If

    ' Validate
    If Not modUtils.IsValidDate(transDate) Or description = "" Then
        ParseAndAddInvestmentRow = False
        Exit Function
    End If

    ' Add to sheet
    lastRow = modUtils.GetLastRow(ws, 1) + 1

    ws.Cells(lastRow, 1).Value = bankName
    ws.Cells(lastRow, 2).Value = transDate
    ws.Cells(lastRow, 3).Value = description
    ws.Cells(lastRow, 4).Value = value
    ws.Cells(lastRow, 5).Value = transType
    ws.Cells(lastRow, 6).Value = "" ' CorrelationID (to be filled by correlation)
    ws.Cells(lastRow, 7).Value = Now ' ImportDate

    ParseAndAddInvestmentRow = True

    Exit Function

ErrorHandler:
    ParseAndAddInvestmentRow = False
End Function

' ========================================================================
' Detect Investment Type (Application or Redemption)
' ========================================================================
Private Function DetectInvestmentType(text As String) As String
    On Error Resume Next

    Dim cleanedText As String
    cleanedText = modUtils.CleanText(text)

    ' Check for application keywords
    If modUtils.ContainsAny(cleanedText, "APLICACAO,APPLICATION,INVEST,APORTE,BUY,COMPRA") Then
        DetectInvestmentType = "Application"
        Exit Function
    End If

    ' Check for redemption keywords
    If modUtils.ContainsAny(cleanedText, "RESGATE,REDEMPTION,REDEEM,WITHDRAW,SAQUE,SELL,VENDA") Then
        DetectInvestmentType = "Redemption"
        Exit Function
    End If

    ' Default based on sign (positive = application, negative = redemption)
    DetectInvestmentType = "Unknown"
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
