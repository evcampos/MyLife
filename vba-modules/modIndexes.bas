Attribute VB_Name = "modIndexes"
' ========================================================================
' Module: modIndexes
' Purpose: Financial Index Management
' Description: Manages financial indexes (CDI, SELIC, IPCA, USD/BRL, Fed Funds)
'              and calculates cumulative factors for capital cost calculations
' Platform: Excel for MacOS
' Author: [MY LIFE] Financial System v1.0
' ========================================================================

Option Explicit

' Business days per year for index calculations
Private Const BUSINESS_DAYS_PER_YEAR As Integer = 252

' ========================================================================
' Calculate All Cumulative Factors
' ========================================================================
Public Sub CalculateAllCumulativeFactors()
    On Error GoTo ErrorHandler

    modUtils.ShowProgress "Calculating cumulative factors for indexes..."
    modUtils.LogDebug "Starting cumulative factor calculation"

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Call CalculateCumulativeFactorsForIndex("CDI")
    Call CalculateCumulativeFactorsForIndex("SELIC")
    Call CalculateCumulativeFactorsForIndex("IPCA")
    Call CalculateCumulativeFactorsForIndex("USD/BRL")
    Call CalculateCumulativeFactorsForIndex("FEDFUNDS")

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    modUtils.ClearProgress
    modUtils.LogDebug "Cumulative factor calculation complete"

    MsgBox "Cumulative factors calculated successfully!", vbInformation

    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    modUtils.ClearProgress
    MsgBox modUtils.FormatErrorMessage("modIndexes", "CalculateAllCumulativeFactors", Err.Description), vbCritical
End Sub

' ========================================================================
' Calculate Cumulative Factors for Specific Index
' ========================================================================
Private Sub CalculateCumulativeFactorsForIndex(indexName As String)
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentIndex As String
    Dim currentDate As Date
    Dim currentValue As Double
    Dim previousFactor As Double
    Dim currentFactor As Double
    Dim dailyRate As Double
    Dim annualRate As Double

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_INDEXES)
    lastRow = modUtils.GetLastRow(ws, 1)

    previousFactor = 1#  ' Start with 1.0

    For i = 2 To lastRow
        currentIndex = Trim(CStr(ws.Cells(i, 1).Value))

        If UCase(currentIndex) = UCase(indexName) Then
            currentDate = ws.Cells(i, 2).Value
            currentValue = ws.Cells(i, 3).Value

            ' Calculate cumulative factor
            ' Formula: CumulativeFactor[n] = CumulativeFactor[n-1] × (1 + dailyRate)
            annualRate = currentValue / 100  ' Convert percentage to decimal
            dailyRate = annualRate / BUSINESS_DAYS_PER_YEAR

            currentFactor = previousFactor * (1 + dailyRate)

            ' Write cumulative factor
            ws.Cells(i, 4).Value = currentFactor

            previousFactor = currentFactor
        End If
    Next i
End Sub

' ========================================================================
' Get Cumulative Factor for Date
' ========================================================================
Public Function GetCumulativeFactor(indexName As String, targetDate As Date) As Double
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentIndex As String
    Dim currentDate As Date
    Dim factor As Double
    Dim found As Boolean

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_INDEXES)
    lastRow = modUtils.GetLastRow(ws, 1)

    found = False
    factor = 1#

    ' Search for exact date or closest prior date
    For i = 2 To lastRow
        currentIndex = Trim(CStr(ws.Cells(i, 1).Value))
        currentDate = ws.Cells(i, 2).Value

        If UCase(currentIndex) = UCase(indexName) Then
            If currentDate <= targetDate Then
                factor = ws.Cells(i, 4).Value
                found = True
            ElseIf currentDate > targetDate Then
                ' Passed the target date, use previous value
                Exit For
            End If
        End If
    Next i

    If found Then
        GetCumulativeFactor = factor
    Else
        GetCumulativeFactor = 1#  ' Default
    End If

    Exit Function

ErrorHandler:
    GetCumulativeFactor = 1#
End Function

' ========================================================================
' Get Index Value for Date
' ========================================================================
Public Function GetIndexValue(indexName As String, targetDate As Date) As Double
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentIndex As String
    Dim currentDate As Date
    Dim value As Double
    Dim found As Boolean

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_INDEXES)
    lastRow = modUtils.GetLastRow(ws, 1)

    found = False
    value = 0

    ' Search for exact date or closest prior date
    For i = 2 To lastRow
        currentIndex = Trim(CStr(ws.Cells(i, 1).Value))
        currentDate = ws.Cells(i, 2).Value

        If UCase(currentIndex) = UCase(indexName) Then
            If currentDate <= targetDate Then
                value = ws.Cells(i, 3).Value
                found = True
            ElseIf currentDate > targetDate Then
                ' Passed the target date, use previous value
                Exit For
            End If
        End If
    Next i

    GetIndexValue = value

    Exit Function

ErrorHandler:
    GetIndexValue = 0
End Function

' ========================================================================
' Import Index Data from File
' ========================================================================
Public Function ImportIndexData() As Long
    On Error GoTo ErrorHandler

    Dim filePath As String
    Dim totalImported As Long

    filePath = modConfig.GetFilePath("INDEXES")

    If filePath = "" Then
        modUtils.LogDebug "No indexes file path configured"
        ImportIndexData = 0
        Exit Function
    End If

    modUtils.ShowProgress "Importing index data..."
    modUtils.LogDebug "Starting index import from: " & filePath

    If modUtils.FileExists(filePath) Then
        totalImported = ImportIndexFile(filePath)
    Else
        modUtils.LogDebug "Index file not found: " & filePath
        MsgBox "Index file not found: " & filePath, vbExclamation
        totalImported = 0
    End If

    modUtils.ClearProgress
    modUtils.LogDebug "Index import complete: " & totalImported & " records"

    ' Calculate cumulative factors after import
    If totalImported > 0 Then
        Call CalculateAllCumulativeFactors
    End If

    ImportIndexData = totalImported

    Exit Function

ErrorHandler:
    modUtils.ClearProgress
    MsgBox modUtils.FormatErrorMessage("modIndexes", "ImportIndexData", Err.Description), vbCritical
    ImportIndexData = 0
End Function

' ========================================================================
' Import Index File
' ========================================================================
Private Function ImportIndexFile(filePath As String) As Long
    On Error GoTo ErrorHandler

    Dim extension As String
    Dim imported As Long

    extension = modUtils.GetFileExtension(filePath)

    Select Case extension
        Case "csv", "txt"
            imported = ImportIndexCSV(filePath)
        Case "xlsx", "xls", "xlsm"
            imported = ImportIndexExcel(filePath)
        Case Else
            MsgBox "Unsupported file format for indexes: " & extension, vbExclamation
            imported = 0
    End Select

    ImportIndexFile = imported

    Exit Function

ErrorHandler:
    MsgBox modUtils.FormatErrorMessage("modIndexes", "ImportIndexFile", Err.Description), vbCritical
    ImportIndexFile = 0
End Function

' ========================================================================
' Import CSV Index File
' ========================================================================
Private Function ImportIndexCSV(filePath As String) As Long
    On Error GoTo ErrorHandler

    Dim fileNum As Integer
    Dim lineText As String
    Dim delimiter As String
    Dim headers() As String
    Dim rowData() As String
    Dim ws As Worksheet
    Dim imported As Long
    Dim lastRow As Long
    Dim indexName As String
    Dim indexDate As Date
    Dim indexValue As Double

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_INDEXES)

    ' Open file
    fileNum = FreeFile
    Open filePath For Input As #fileNum

    ' Read header line
    Line Input #fileNum, lineText
    delimiter = DetectDelimiter(lineText)
    headers = Split(lineText, delimiter)

    ' Map columns
    Dim colIndex As Integer, colDate As Integer, colValue As Integer
    colIndex = FindColumn(headers, "INDEX,INDICE,NAME")
    colDate = FindColumn(headers, "DATE,DATA")
    colValue = FindColumn(headers, "VALUE,VALOR,RATE")

    If colIndex = -1 Or colDate = -1 Or colValue = -1 Then
        Close #fileNum
        MsgBox "Could not detect required columns in index file", vbExclamation
        ImportIndexCSV = 0
        Exit Function
    End If

    ' Read data rows
    imported = 0
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText

        If Trim(lineText) <> "" Then
            rowData = Split(lineText, delimiter)

            indexName = Trim(CStr(rowData(colIndex)))
            indexDate = modUtils.ParseDate(Trim(CStr(rowData(colDate))))
            indexValue = CDbl(Trim(CStr(rowData(colValue))))

            If modUtils.IsValidDate(indexDate) And indexName <> "" Then
                lastRow = modUtils.GetLastRow(ws, 1) + 1
                ws.Cells(lastRow, 1).Value = indexName
                ws.Cells(lastRow, 2).Value = indexDate
                ws.Cells(lastRow, 3).Value = indexValue
                ws.Cells(lastRow, 4).Value = 1#  ' Will be calculated later

                imported = imported + 1
            End If
        End If
    Loop

    Close #fileNum

    ImportIndexCSV = imported

    Exit Function

ErrorHandler:
    On Error Resume Next
    Close #fileNum
    MsgBox modUtils.FormatErrorMessage("modIndexes", "ImportIndexCSV", Err.Description), vbCritical
    ImportIndexCSV = 0
End Function

' ========================================================================
' Import Excel Index File
' ========================================================================
Private Function ImportIndexExcel(filePath As String) As Long
    On Error GoTo ErrorHandler

    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim destWs As Worksheet
    Dim lastRow As Long
    Dim destLastRow As Long
    Dim i As Long
    Dim imported As Long

    Set destWs = ThisWorkbook.Worksheets(modConfig.SHEET_INDEXES)

    ' Open source workbook
    Application.ScreenUpdating = False
    Set sourceWb = Workbooks.Open(filePath, ReadOnly:=True)
    Set sourceWs = sourceWb.Worksheets(1)

    lastRow = sourceWs.Cells(sourceWs.Rows.Count, 1).End(xlUp).row

    imported = 0
    For i = 2 To lastRow
        destLastRow = modUtils.GetLastRow(destWs, 1) + 1
        destWs.Cells(destLastRow, 1).Value = sourceWs.Cells(i, 1).Value  ' Index
        destWs.Cells(destLastRow, 2).Value = sourceWs.Cells(i, 2).Value  ' Date
        destWs.Cells(destLastRow, 3).Value = sourceWs.Cells(i, 3).Value  ' Value
        destWs.Cells(destLastRow, 4).Value = 1#  ' Cumulative factor placeholder

        imported = imported + 1
    Next i

    sourceWb.Close SaveChanges:=False
    Application.ScreenUpdating = True

    ImportIndexExcel = imported

    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not sourceWb Is Nothing Then sourceWb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    MsgBox modUtils.FormatErrorMessage("modIndexes", "ImportIndexExcel", Err.Description), vbCritical
    ImportIndexExcel = 0
End Function

' ========================================================================
' Detect CSV Delimiter
' ========================================================================
Private Function DetectDelimiter(headerLine As String) As String
    On Error Resume Next

    Dim commaCount As Integer
    Dim semicolonCount As Integer

    commaCount = Len(headerLine) - Len(Replace(headerLine, ",", ""))
    semicolonCount = Len(headerLine) - Len(Replace(headerLine, ";", ""))

    If semicolonCount > commaCount Then
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
