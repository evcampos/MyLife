Attribute VB_Name = "modImportOPUS"
' ========================================================================
' Module: modImportOPUS
' Purpose: Import OPUS External Investment Positions
' Description: Imports external investment positions (assets & liabilities)
'              with multi-currency support
' Platform: Excel for MacOS
' Author: [MY LIFE] Financial System v1.0
' ========================================================================

Option Explicit

' ========================================================================
' Main Import Function - Import OPUS Data
' ========================================================================
Public Function ImportOPUS() As Long
    On Error GoTo ErrorHandler

    Dim filePath As String
    Dim totalImported As Long

    filePath = modConfig.GetFilePath("OPUS")

    If filePath = "" Then
        modUtils.LogDebug "No OPUS file path configured"
        ImportOPUS = 0
        Exit Function
    End If

    modUtils.ShowProgress "Importing OPUS positions..."
    modUtils.LogDebug "Starting OPUS import from: " & filePath

    If modUtils.FileExists(filePath) Then
        totalImported = ImportOPUSFile(filePath)
    Else
        modUtils.LogDebug "OPUS file not found: " & filePath
        MsgBox "OPUS file not found: " & filePath, vbExclamation
        totalImported = 0
    End If

    modUtils.ClearProgress
    modUtils.LogDebug "OPUS import complete: " & totalImported & " positions"

    ImportOPUS = totalImported

    Exit Function

ErrorHandler:
    modUtils.ClearProgress
    MsgBox modUtils.FormatErrorMessage("modImportOPUS", "ImportOPUS", Err.Description), vbCritical
    ImportOPUS = 0
End Function

' ========================================================================
' Import OPUS File
' ========================================================================
Private Function ImportOPUSFile(filePath As String) As Long
    On Error GoTo ErrorHandler

    Dim extension As String
    Dim imported As Long

    extension = modUtils.GetFileExtension(filePath)

    Select Case extension
        Case "csv", "txt"
            imported = ImportOPUSCSV(filePath)
        Case "xlsx", "xls", "xlsm"
            imported = ImportOPUSExcel(filePath)
        Case Else
            MsgBox "Unsupported file format for OPUS: " & extension, vbExclamation
            imported = 0
    End Select

    ImportOPUSFile = imported

    Exit Function

ErrorHandler:
    MsgBox modUtils.FormatErrorMessage("modImportOPUS", "ImportOPUSFile", Err.Description), vbCritical
    ImportOPUSFile = 0
End Function

' ========================================================================
' Import CSV OPUS File
' ========================================================================
Private Function ImportOPUSCSV(filePath As String) As Long
    On Error GoTo ErrorHandler

    Dim fileNum As Integer
    Dim lineText As String
    Dim delimiter As String
    Dim headers() As String
    Dim rowData() As String
    Dim ws As Worksheet
    Dim imported As Long

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_OPUS)

    ' Open file
    fileNum = FreeFile
    Open filePath For Input As #fileNum

    ' Read header line
    Line Input #fileNum, lineText

    ' Detect delimiter
    delimiter = DetectDelimiter(lineText)
    headers = Split(lineText, delimiter)

    ' Map columns
    Dim colType As Integer, colCompany As Integer, colInvestCost As Integer
    Dim colCapitalCost As Integer, colUpdatedCost As Integer, colCurrency As Integer

    colType = FindColumn(headers, "TYPE,TIPO,ASSET_LIABILITY")
    colCompany = FindColumn(headers, "COMPANY,EMPRESA,NAME,NOME")
    colInvestCost = FindColumn(headers, "INVESTMENTCOST,INVESTMENT_COST,INITIAL_VALUE,VALOR_INVESTIDO")
    colCapitalCost = FindColumn(headers, "CAPITALCOST,CAPITAL_COST,INTEREST,JUROS")
    colUpdatedCost = FindColumn(headers, "UPDATEDCOST,UPDATED_COST,CURRENT_VALUE,VALOR_ATUALIZADO")
    colCurrency = FindColumn(headers, "CURRENCY,MOEDA")

    If colType = -1 Or colCompany = -1 Or colInvestCost = -1 Then
        Close #fileNum
        MsgBox "Could not detect required columns in OPUS file", vbExclamation
        ImportOPUSCSV = 0
        Exit Function
    End If

    ' Read data rows
    imported = 0
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText

        If Trim(lineText) <> "" Then
            rowData = Split(lineText, delimiter)

            If ParseAndAddOPUSRow(ws, rowData, colType, colCompany, colInvestCost, _
                                 colCapitalCost, colUpdatedCost, colCurrency) Then
                imported = imported + 1
            End If
        End If
    Loop

    Close #fileNum

    ImportOPUSCSV = imported

    Exit Function

ErrorHandler:
    On Error Resume Next
    Close #fileNum
    MsgBox modUtils.FormatErrorMessage("modImportOPUS", "ImportOPUSCSV", Err.Description), vbCritical
    ImportOPUSCSV = 0
End Function

' ========================================================================
' Import Excel OPUS File
' ========================================================================
Private Function ImportOPUSExcel(filePath As String) As Long
    On Error GoTo ErrorHandler

    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim destWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim imported As Long
    Dim headers() As Variant
    Dim colCount As Integer

    Set destWs = ThisWorkbook.Worksheets(modConfig.SHEET_OPUS)

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
    Dim colType As Integer, colCompany As Integer, colInvestCost As Integer
    Dim colCapitalCost As Integer, colUpdatedCost As Integer, colCurrency As Integer

    colType = FindColumn(headers, "TYPE,TIPO,ASSET_LIABILITY")
    colCompany = FindColumn(headers, "COMPANY,EMPRESA,NAME,NOME")
    colInvestCost = FindColumn(headers, "INVESTMENTCOST,INVESTMENT_COST,INITIAL_VALUE,VALOR_INVESTIDO")
    colCapitalCost = FindColumn(headers, "CAPITALCOST,CAPITAL_COST,INTEREST,JUROS")
    colUpdatedCost = FindColumn(headers, "UPDATEDCOST,UPDATED_COST,CURRENT_VALUE,VALOR_ATUALIZADO")
    colCurrency = FindColumn(headers, "CURRENCY,MOEDA")

    If colType = -1 Or colCompany = -1 Or colInvestCost = -1 Then
        sourceWb.Close SaveChanges:=False
        Application.ScreenUpdating = True
        MsgBox "Could not detect required columns in OPUS file", vbExclamation
        ImportOPUSExcel = 0
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

        If ParseAndAddOPUSRow(destWs, rowData, colType, colCompany, colInvestCost, _
                             colCapitalCost, colUpdatedCost, colCurrency) Then
            imported = imported + 1
        End If
    Next i

    sourceWb.Close SaveChanges:=False
    Application.ScreenUpdating = True

    ImportOPUSExcel = imported

    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not sourceWb Is Nothing Then sourceWb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    MsgBox modUtils.FormatErrorMessage("modImportOPUS", "ImportOPUSExcel", Err.Description), vbCritical
    ImportOPUSExcel = 0
End Function

' ========================================================================
' Parse and Add OPUS Row
' ========================================================================
Private Function ParseAndAddOPUSRow(ws As Worksheet, rowData As Variant, _
                                    colType As Integer, colCompany As Integer, _
                                    colInvestCost As Integer, colCapitalCost As Integer, _
                                    colUpdatedCost As Integer, colCurrency As Integer) As Boolean
    On Error GoTo ErrorHandler

    Dim lastRow As Long
    Dim itemType As String
    Dim company As String
    Dim investCost As Double
    Dim capitalCost As String
    Dim updatedCost As Double
    Dim currency As String

    ' Get values
    itemType = GetColumnValue(rowData, colType)
    company = GetColumnValue(rowData, colCompany)
    investCost = CDbl(GetColumnValue(rowData, colInvestCost))

    ' Optional fields
    If colCapitalCost > 0 Then
        capitalCost = GetColumnValue(rowData, colCapitalCost)
    Else
        capitalCost = "0%"
    End If

    If colUpdatedCost > 0 Then
        Dim updatedVal As String
        updatedVal = GetColumnValue(rowData, colUpdatedCost)
        If updatedVal <> "" Then
            updatedCost = CDbl(updatedVal)
        Else
            updatedCost = investCost
        End If
    Else
        updatedCost = investCost
    End If

    If colCurrency > 0 Then
        currency = GetColumnValue(rowData, colCurrency)
    Else
        currency = "BRL"  ' Default currency
    End If

    ' Validate
    If company = "" Or itemType = "" Then
        ParseAndAddOPUSRow = False
        Exit Function
    End If

    ' Normalize type
    itemType = NormalizeType(itemType)

    ' Add to sheet
    lastRow = modUtils.GetLastRow(ws, 1) + 1

    ws.Cells(lastRow, 1).Value = itemType
    ws.Cells(lastRow, 2).Value = company
    ws.Cells(lastRow, 3).Value = investCost
    ws.Cells(lastRow, 4).Value = capitalCost
    ws.Cells(lastRow, 5).Value = updatedCost
    ws.Cells(lastRow, 6).Value = currency
    ws.Cells(lastRow, 7).Value = Now ' ImportDate

    ParseAndAddOPUSRow = True

    Exit Function

ErrorHandler:
    ParseAndAddOPUSRow = False
End Function

' ========================================================================
' Normalize Type (Asset/Liability)
' ========================================================================
Private Function NormalizeType(rawType As String) As String
    On Error Resume Next

    Dim cleanedType As String
    cleanedType = modUtils.CleanText(rawType)

    If modUtils.ContainsAny(cleanedType, "ASSET,ATIVO,INVESTMENT,INVESTIMENTO") Then
        NormalizeType = "Asset"
    ElseIf modUtils.ContainsAny(cleanedType, "LIABILITY,PASSIVO,DEBT,DIVIDA") Then
        NormalizeType = "Liability"
    Else
        NormalizeType = rawType ' Keep original if can't classify
    End If
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
