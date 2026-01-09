Attribute VB_Name = "modImportCards"
' ========================================================================
' Module: modImportCards
' Purpose: Import Credit Card Transactions
' Description: Imports credit card statements with installment tracking
' Platform: Excel for MacOS
' Author: [MY LIFE] Financial System v1.0
' ========================================================================

Option Explicit

' ========================================================================
' Main Import Function - Import All Card Sources
' ========================================================================
Public Function ImportAllCards() As Long
    On Error GoTo ErrorHandler

    Dim sources() As String
    Dim source As Variant
    Dim filePath As String
    Dim totalImported As Long

    sources = Array("ITAU_CARD", "NUBANK_CARD", "C6_CARD", "BB_CARD")

    modUtils.ShowProgress "Importing credit card transactions..."
    modUtils.LogDebug "Starting card imports"

    For Each source In sources
        filePath = modConfig.GetFilePath(CStr(source))

        If filePath <> "" Then
            modUtils.LogDebug "Importing from: " & source

            If modUtils.FileExists(filePath) Then
                totalImported = totalImported + ImportCardFile(CStr(source), filePath)
            Else
                modUtils.LogDebug "File not found: " & filePath
            End If
        End If
    Next source

    modUtils.ClearProgress
    modUtils.LogDebug "Card imports complete: " & totalImported & " transactions"

    ImportAllCards = totalImported

    Exit Function

ErrorHandler:
    modUtils.ClearProgress
    MsgBox modUtils.FormatErrorMessage("modImportCards", "ImportAllCards", Err.Description), vbCritical
    ImportAllCards = 0
End Function

' ========================================================================
' Import Single Card File
' ========================================================================
Private Function ImportCardFile(sourceName As String, filePath As String) As Long
    On Error GoTo ErrorHandler

    Dim extension As String
    Dim imported As Long

    extension = modUtils.GetFileExtension(filePath)

    Select Case extension
        Case "csv", "txt"
            imported = ImportCardCSV(sourceName, filePath)
        Case "xlsx", "xls", "xlsm"
            imported = ImportCardExcel(sourceName, filePath)
        Case Else
            MsgBox "Unsupported file format for " & sourceName & ": " & extension, vbExclamation
            imported = 0
    End Select

    ImportCardFile = imported

    Exit Function

ErrorHandler:
    MsgBox modUtils.FormatErrorMessage("modImportCards", "ImportCardFile", Err.Description), vbCritical
    ImportCardFile = 0
End Function

' ========================================================================
' Import CSV Card File
' ========================================================================
Private Function ImportCardCSV(sourceName As String, filePath As String) As Long
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
    bankName = Replace(sourceName, "_CARD", "")

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_CARDS)

    ' Open file
    fileNum = FreeFile
    Open filePath For Input As #fileNum

    ' Read header line
    Line Input #fileNum, lineText

    ' Detect delimiter
    delimiter = DetectDelimiter(lineText)
    headers = Split(lineText, delimiter)

    ' Map columns
    Dim colDate As Integer, colDesc As Integer, colValue As Integer
    Dim colCategory As Integer, colInstallment As Integer, colCard As Integer

    colDate = FindColumn(headers, "DATE,DATA,PURCHASEDATE,PURCHASE_DATE")
    colDesc = FindColumn(headers, "DESCRIPTION,DESCRICAO,HISTORICO,MERCHANT,ESTABELECIMENTO")
    colValue = FindColumn(headers, "VALUE,VALOR,AMOUNT")
    colCategory = FindColumn(headers, "CATEGORY,CATEGORIA,TYPE")
    colInstallment = FindColumn(headers, "INSTALLMENT,PARCELA,INSTALLMENTS")
    colCard = FindColumn(headers, "CARD,CARTAO,CARDNUMBER,FINAL")

    If colDate = -1 Or colDesc = -1 Or colValue = -1 Then
        Close #fileNum
        MsgBox "Could not detect required columns in " & sourceName, vbExclamation
        ImportCardCSV = 0
        Exit Function
    End If

    ' Read data rows
    imported = 0
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText

        If Trim(lineText) <> "" Then
            rowData = Split(lineText, delimiter)

            If ParseAndAddCardRow(ws, bankName, rowData, colDate, colDesc, colValue, _
                                 colCategory, colInstallment, colCard) Then
                imported = imported + 1
            End If
        End If
    Loop

    Close #fileNum

    ImportCardCSV = imported

    Exit Function

ErrorHandler:
    On Error Resume Next
    Close #fileNum
    MsgBox modUtils.FormatErrorMessage("modImportCards", "ImportCardCSV", Err.Description), vbCritical
    ImportCardCSV = 0
End Function

' ========================================================================
' Import Excel Card File
' ========================================================================
Private Function ImportCardExcel(sourceName As String, filePath As String) As Long
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
    bankName = Replace(sourceName, "_CARD", "")

    Set destWs = ThisWorkbook.Worksheets(modConfig.SHEET_CARDS)

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
    Dim colDate As Integer, colDesc As Integer, colValue As Integer
    Dim colCategory As Integer, colInstallment As Integer, colCard As Integer

    colDate = FindColumn(headers, "DATE,DATA,PURCHASEDATE,PURCHASE_DATE")
    colDesc = FindColumn(headers, "DESCRIPTION,DESCRICAO,HISTORICO,MERCHANT,ESTABELECIMENTO")
    colValue = FindColumn(headers, "VALUE,VALOR,AMOUNT")
    colCategory = FindColumn(headers, "CATEGORY,CATEGORIA,TYPE")
    colInstallment = FindColumn(headers, "INSTALLMENT,PARCELA,INSTALLMENTS")
    colCard = FindColumn(headers, "CARD,CARTAO,CARDNUMBER,FINAL")

    If colDate = -1 Or colDesc = -1 Or colValue = -1 Then
        sourceWb.Close SaveChanges:=False
        Application.ScreenUpdating = True
        MsgBox "Could not detect required columns in " & sourceName, vbExclamation
        ImportCardExcel = 0
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

        If ParseAndAddCardRow(destWs, bankName, rowData, colDate, colDesc, colValue, _
                             colCategory, colInstallment, colCard) Then
            imported = imported + 1
        End If
    Next i

    sourceWb.Close SaveChanges:=False
    Application.ScreenUpdating = True

    ImportCardExcel = imported

    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not sourceWb Is Nothing Then sourceWb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    MsgBox modUtils.FormatErrorMessage("modImportCards", "ImportCardExcel", Err.Description), vbCritical
    ImportCardExcel = 0
End Function

' ========================================================================
' Parse and Add Card Row
' ========================================================================
Private Function ParseAndAddCardRow(ws As Worksheet, bankName As String, _
                                    rowData As Variant, colDate As Integer, _
                                    colDesc As Integer, colValue As Integer, _
                                    colCategory As Integer, colInstallment As Integer, _
                                    colCard As Integer) As Boolean
    On Error GoTo ErrorHandler

    Dim lastRow As Long
    Dim transDate As Date
    Dim description As String
    Dim value As Double
    Dim category As String
    Dim installment As String
    Dim cardNumber As String

    ' Get values
    transDate = modUtils.ParseDate(GetColumnValue(rowData, colDate))
    description = GetColumnValue(rowData, colDesc)
    value = CDbl(GetColumnValue(rowData, colValue))

    ' Optional fields
    If colCategory > 0 Then
        category = GetColumnValue(rowData, colCategory)
    Else
        category = ""
    End If

    If colInstallment > 0 Then
        installment = GetColumnValue(rowData, colInstallment)
    Else
        installment = "1/1"  ' Default: single payment
    End If

    If colCard > 0 Then
        cardNumber = GetColumnValue(rowData, colCard)
    Else
        cardNumber = ""
    End If

    ' Validate
    If Not modUtils.IsValidDate(transDate) Or description = "" Then
        ParseAndAddCardRow = False
        Exit Function
    End If

    ' Add to sheet
    lastRow = modUtils.GetLastRow(ws, 1) + 1

    ws.Cells(lastRow, 1).Value = bankName
    ws.Cells(lastRow, 2).Value = cardNumber
    ws.Cells(lastRow, 3).Value = transDate
    ws.Cells(lastRow, 4).Value = category
    ws.Cells(lastRow, 5).Value = description
    ws.Cells(lastRow, 6).Value = installment
    ws.Cells(lastRow, 7).Value = value
    ws.Cells(lastRow, 8).Value = Now ' ImportDate

    ParseAndAddCardRow = True

    Exit Function

ErrorHandler:
    ParseAndAddCardRow = False
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
