Attribute VB_Name = "modUtils"
' ========================================================================
' Module: modUtils
' Purpose: Utility and Helper Functions
' Description: Common functions used throughout the system including
'              string manipulation, date handling, and fuzzy matching
' Platform: Excel for MacOS
' Author: Financial Import System v1.0
' ========================================================================

Option Explicit

' ========================================================================
' String Utility Functions
' ========================================================================

' Clean and normalize text for matching
Public Function CleanText(text As String) As String
    On Error Resume Next

    Dim result As String

    result = Trim(text)
    result = UCase(result)

    ' Remove multiple spaces
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop

    ' Remove special characters that might interfere with matching
    result = Replace(result, vbTab, " ")
    result = Replace(result, vbCr, " ")
    result = Replace(result, vbLf, " ")

    CleanText = result
End Function

' Extract numbers from text
Public Function ExtractNumber(text As String) As Double
    On Error Resume Next

    Dim i As Integer
    Dim result As String
    Dim char As String
    Dim decimalFound As Boolean

    decimalFound = False

    For i = 1 To Len(text)
        char = Mid(text, i, 1)

        If IsNumeric(char) Then
            result = result & char
        ElseIf (char = "." Or char = ",") And Not decimalFound Then
            result = result & "."
            decimalFound = True
        ElseIf char = "-" And i = 1 Then
            result = result & char
        End If
    Next i

    If result <> "" Then
        ExtractNumber = CDbl(result)
    Else
        ExtractNumber = 0
    End If
End Function

' Check if string contains any of the keywords
Public Function ContainsAny(text As String, keywords As String, Optional delimiter As String = ",") As Boolean
    On Error Resume Next

    Dim keywordArray() As String
    Dim keyword As Variant

    text = CleanText(text)
    keywordArray = Split(keywords, delimiter)

    For Each keyword In keywordArray
        If InStr(1, text, CleanText(CStr(keyword)), vbTextCompare) > 0 Then
            ContainsAny = True
            Exit Function
        End If
    Next keyword

    ContainsAny = False
End Function

' Get first keyword match
Public Function GetFirstMatch(text As String, keywords As String, Optional delimiter As String = ",") As String
    On Error Resume Next

    Dim keywordArray() As String
    Dim keyword As Variant

    text = CleanText(text)
    keywordArray = Split(keywords, delimiter)

    For Each keyword In keywordArray
        If InStr(1, text, CleanText(CStr(keyword)), vbTextCompare) > 0 Then
            GetFirstMatch = Trim(CStr(keyword))
            Exit Function
        End If
    Next keyword

    GetFirstMatch = ""
End Function

' ========================================================================
' Fuzzy Matching Functions
' ========================================================================

' Calculate similarity between two strings (0-100)
Public Function FuzzyMatch(str1 As String, str2 As String) As Integer
    On Error Resume Next

    Dim maxLen As Integer
    Dim distance As Integer
    Dim similarity As Double

    str1 = CleanText(str1)
    str2 = CleanText(str2)

    ' Handle empty strings
    If str1 = "" Or str2 = "" Then
        FuzzyMatch = 0
        Exit Function
    End If

    ' Exact match
    If str1 = str2 Then
        FuzzyMatch = 100
        Exit Function
    End If

    ' Calculate Levenshtein distance
    distance = LevenshteinDistance(str1, str2)
    maxLen = Application.WorksheetFunction.Max(Len(str1), Len(str2))

    ' Calculate similarity percentage
    similarity = (1 - (distance / maxLen)) * 100

    FuzzyMatch = CInt(similarity)
End Function

' Levenshtein Distance Algorithm
Private Function LevenshteinDistance(str1 As String, str2 As String) As Integer
    On Error Resume Next

    Dim len1 As Integer, len2 As Integer
    Dim i As Integer, j As Integer
    Dim cost As Integer
    Dim d() As Integer

    len1 = Len(str1)
    len2 = Len(str2)

    ReDim d(0 To len1, 0 To len2)

    For i = 0 To len1
        d(i, 0) = i
    Next i

    For j = 0 To len2
        d(0, j) = j
    Next j

    For i = 1 To len1
        For j = 1 To len2
            If Mid(str1, i, 1) = Mid(str2, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If

            d(i, j) = Application.WorksheetFunction.Min( _
                d(i - 1, j) + 1, _
                d(i, j - 1) + 1, _
                d(i - 1, j - 1) + cost)
        Next j
    Next i

    LevenshteinDistance = d(len1, len2)
End Function

' Find best fuzzy match from a list
Public Function BestFuzzyMatch(text As String, candidateList As String, _
                               Optional delimiter As String = ",", _
                               Optional threshold As Integer = 80) As String
    On Error Resume Next

    Dim candidates() As String
    Dim candidate As Variant
    Dim score As Integer
    Dim bestScore As Integer
    Dim bestMatch As String

    candidates = Split(candidateList, delimiter)
    bestScore = 0
    bestMatch = ""

    For Each candidate In candidates
        score = FuzzyMatch(text, CStr(candidate))

        If score > bestScore Then
            bestScore = score
            bestMatch = Trim(CStr(candidate))
        End If
    Next candidate

    If bestScore >= threshold Then
        BestFuzzyMatch = bestMatch
    Else
        BestFuzzyMatch = ""
    End If
End Function

' ========================================================================
' Date Utility Functions
' ========================================================================

' Parse date from various formats
Public Function ParseDate(dateStr As String, Optional dateFormat As String = "") As Date
    On Error Resume Next

    Dim result As Date

    dateStr = Trim(dateStr)

    ' Try direct conversion
    result = CDate(dateStr)
    If Err.Number = 0 Then
        ParseDate = result
        Exit Function
    End If
    Err.Clear

    ' Try with format hints
    If InStr(dateStr, "/") > 0 Then
        ' Likely dd/mm/yyyy or mm/dd/yyyy
        result = ParseSlashDate(dateStr)
    ElseIf InStr(dateStr, "-") > 0 Then
        ' Likely yyyy-mm-dd
        result = ParseDashDate(dateStr)
    End If

    ParseDate = result
End Function

' Parse date with slashes
Private Function ParseSlashDate(dateStr As String) As Date
    On Error Resume Next

    Dim parts() As String
    Dim day As Integer, month As Integer, year As Integer

    parts = Split(dateStr, "/")

    If UBound(parts) = 2 Then
        ' Assume dd/mm/yyyy format (common in many countries)
        day = CInt(parts(0))
        month = CInt(parts(1))
        year = CInt(parts(2))

        ' Handle 2-digit years
        If year < 100 Then
            If year > 50 Then
                year = 1900 + year
            Else
                year = 2000 + year
            End If
        End If

        ParseSlashDate = DateSerial(year, month, day)
    End If
End Function

' Parse date with dashes
Private Function ParseDashDate(dateStr As String) As Date
    On Error Resume Next

    Dim parts() As String
    Dim day As Integer, month As Integer, year As Integer

    parts = Split(dateStr, "-")

    If UBound(parts) = 2 Then
        ' Assume yyyy-mm-dd format (ISO)
        year = CInt(parts(0))
        month = CInt(parts(1))
        day = CInt(parts(2))

        ParseDashDate = DateSerial(year, month, day)
    End If
End Function

' Check if two dates are within tolerance
Public Function DatesMatch(date1 As Date, date2 As Date, Optional toleranceDays As Integer = 0) As Boolean
    On Error Resume Next

    DatesMatch = (Abs(date1 - date2) <= toleranceDays)
End Function

' ========================================================================
' Numeric Utility Functions
' ========================================================================

' Check if two amounts are equal within tolerance
Public Function AmountsMatch(amount1 As Double, amount2 As Double, Optional tolerance As Double = 0.01) As Boolean
    On Error Resume Next

    AmountsMatch = (Abs(amount1 - amount2) <= tolerance)
End Function

' Round to 2 decimal places
Public Function RoundCurrency(value As Double) As Double
    On Error Resume Next

    RoundCurrency = Round(value, 2)
End Function

' Normalize amount sign (negative for debits, positive for credits)
Public Function NormalizeAmount(amount As Double, transactionType As String) As Double
    On Error Resume Next

    transactionType = UCase(Trim(transactionType))

    Select Case transactionType
        Case "DEBIT", "DÉBITO", "DEBITO", "D", "WITHDRAWAL", "SAIDA"
            NormalizeAmount = -Abs(amount)
        Case "CREDIT", "CRÉDITO", "CREDITO", "C", "DEPOSIT", "ENTRADA"
            NormalizeAmount = Abs(amount)
        Case Else
            NormalizeAmount = amount
    End Select
End Function

' ========================================================================
' Data Validation Functions
' ========================================================================

' Check if value is a valid date
Public Function IsValidDate(value As Variant) As Boolean
    On Error Resume Next

    Dim testDate As Date
    testDate = CDate(value)

    IsValidDate = (Err.Number = 0 And testDate > 0)
End Function

' Check if value is numeric
Public Function IsNumericValue(value As Variant) As Boolean
    On Error Resume Next

    IsNumericValue = IsNumeric(value)
End Function

' Validate required field
Public Function IsNotEmpty(value As Variant) As Boolean
    On Error Resume Next

    If IsNull(value) Then
        IsNotEmpty = False
    ElseIf VarType(value) = vbString Then
        IsNotEmpty = (Trim(CStr(value)) <> "")
    Else
        IsNotEmpty = True
    End If
End Function

' ========================================================================
' Hash and ID Generation Functions
' ========================================================================

' Generate hash for duplicate detection
Public Function GenerateTransactionHash(transDate As Date, amount As Double, description As String) As String
    On Error Resume Next

    Dim hashInput As String

    ' Create hash from key fields
    hashInput = Format(transDate, "yyyymmdd") & "|" & _
                Format(amount, "0.00") & "|" & _
                Left(CleanText(description), 30)

    GenerateTransactionHash = GetStringHash(hashInput)
End Function

' Simple string hash function
Private Function GetStringHash(text As String) As String
    On Error Resume Next

    Dim i As Long
    Dim hashValue As Long
    Dim char As String

    hashValue = 0

    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        hashValue = ((hashValue * 31) + Asc(char)) Mod 2147483647
    Next i

    GetStringHash = CStr(hashValue)
End Function

' Generate unique ID
Public Function GenerateUniqueID(Optional prefix As String = "TXN") As String
    On Error Resume Next

    Dim timestamp As String
    Dim random As String

    timestamp = Format(Now, "yyyymmddhhnnss")
    random = CStr(Int(Rnd * 9999))

    GenerateUniqueID = prefix & "-" & timestamp & "-" & random
End Function

' ========================================================================
' Array and Collection Functions
' ========================================================================

' Convert delimited string to array
Public Function StringToArray(text As String, Optional delimiter As String = ",") As Variant
    On Error Resume Next

    StringToArray = Split(text, delimiter)
End Function

' Check if value exists in array
Public Function InArray(value As Variant, arr As Variant) As Boolean
    On Error Resume Next

    Dim element As Variant

    For Each element In arr
        If element = value Then
            InArray = True
            Exit Function
        End If
    Next element

    InArray = False
End Function

' ========================================================================
' File and Path Functions
' ========================================================================

' Extract file name from path
Public Function GetFileName(fullPath As String) As String
    On Error Resume Next

    Dim parts() As String

    ' Handle both Windows and Mac path separators
    If InStr(fullPath, "/") > 0 Then
        parts = Split(fullPath, "/")
    Else
        parts = Split(fullPath, "\")
    End If

    GetFileName = parts(UBound(parts))
End Function

' Extract file extension
Public Function GetFileExtension(fileName As String) As String
    On Error Resume Next

    Dim parts() As String

    parts = Split(fileName, ".")

    If UBound(parts) > 0 Then
        GetFileExtension = LCase(parts(UBound(parts)))
    Else
        GetFileExtension = ""
    End If
End Function

' Check if file is CSV
Public Function IsCSVFile(fileName As String) As Boolean
    IsCSVFile = (GetFileExtension(fileName) = "csv")
End Function

' Check if file is Excel
Public Function IsExcelFile(fileName As String) As Boolean
    Dim ext As String
    ext = GetFileExtension(fileName)
    IsExcelFile = (ext = "xlsx" Or ext = "xls" Or ext = "xlsm")
End Function

' ========================================================================
' Logging Functions
' ========================================================================

' Log message to Audit sheet
Public Sub LogMessage(message As String, Optional logLevel As String = "INFO")
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long

    ' Get or create Audit sheet
    If Not modConfig.SheetExists(modConfig.SHEET_AUDIT) Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = modConfig.SHEET_AUDIT
        ws.Cells(1, 1).Value = "Timestamp"
        ws.Cells(1, 2).Value = "Level"
        ws.Cells(1, 3).Value = "Message"
        ws.Range("A1:C1").Font.Bold = True
    Else
        Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_AUDIT)
    End If

    ' Add log entry
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(lastRow, 1).Value = Now
    ws.Cells(lastRow, 2).Value = logLevel
    ws.Cells(lastRow, 3).Value = message
End Sub

' ========================================================================
' Progress Indicator Functions
' ========================================================================

' Show progress in status bar
Public Sub ShowProgress(message As String, Optional current As Long = 0, Optional total As Long = 0)
    On Error Resume Next

    Dim percent As String

    If total > 0 And current > 0 Then
        percent = " (" & Format((current / total), "0%") & ")"
    Else
        percent = ""
    End If

    Application.StatusBar = message & percent
    DoEvents
End Sub

' Clear progress indicator
Public Sub ClearProgress()
    On Error Resume Next

    Application.StatusBar = False
End Sub
