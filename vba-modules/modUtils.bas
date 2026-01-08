Attribute VB_Name = "modUtils"
' ========================================================================
' Module: modUtils
' Purpose: Utility Functions for [MY LIFE]
' Description: Common helper functions used throughout the system
' Platform: Excel for MacOS
' Author: [MY LIFE] Financial System v1.0
' ========================================================================

Option Explicit

' ========================================================================
' String Utility Functions
' ========================================================================

' Clean and normalize text
Public Function CleanText(text As String) As String
    On Error Resume Next

    Dim result As String

    result = Trim(text)
    result = UCase(result)

    ' Remove multiple spaces
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop

    CleanText = result
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

' Calculate text similarity (0-100)
Public Function TextSimilarity(str1 As String, str2 As String) As Integer
    On Error Resume Next

    Dim maxLen As Integer
    Dim distance As Integer
    Dim similarity As Double

    str1 = CleanText(str1)
    str2 = CleanText(str2)

    If str1 = "" Or str2 = "" Then
        TextSimilarity = 0
        Exit Function
    End If

    If str1 = str2 Then
        TextSimilarity = 100
        Exit Function
    End If

    ' Simple similarity: count matching characters
    Dim matches As Integer
    Dim i As Integer

    matches = 0
    maxLen = Application.WorksheetFunction.Max(Len(str1), Len(str2))

    For i = 1 To Application.WorksheetFunction.Min(Len(str1), Len(str2))
        If Mid(str1, i, 1) = Mid(str2, i, 1) Then
            matches = matches + 1
        End If
    Next i

    similarity = (matches / maxLen) * 100
    TextSimilarity = CInt(similarity)
End Function

' ========================================================================
' Date Utility Functions
' ========================================================================

' Parse date from string
Public Function ParseDate(dateStr As String) As Date
    On Error Resume Next

    Dim result As Date

    dateStr = Trim(dateStr)

    ' Try direct conversion
    result = CDate(dateStr)

    If Err.Number = 0 Then
        ParseDate = result
    Else
        ParseDate = 0
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

' ========================================================================
' File Operations
' ========================================================================

' Check if file exists (MacOS compatible)
Public Function FileExists(filePath As String) As Boolean
    On Error Resume Next

    Dim result As String
    result = Dir(filePath)

    FileExists = (result <> "")
End Function

' Get file extension
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

' ========================================================================
' ID Generation
' ========================================================================

' Generate unique ID
Public Function GenerateUniqueID(Optional prefix As String = "ID") As String
    On Error Resume Next

    Dim timestamp As String
    Dim random As String

    timestamp = Format(Now, "yyyymmddhhnnss")
    random = CStr(Int(Rnd * 9999))

    GenerateUniqueID = prefix & "-" & timestamp & "-" & random
End Function

' ========================================================================
' Progress Indication
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

' Clear progress
Public Sub ClearProgress()
    On Error Resume Next

    Application.StatusBar = False
End Sub

' ========================================================================
' Logging
' ========================================================================

' Log message to Immediate Window (for debugging)
Public Sub LogDebug(message As String)
    Debug.Print Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & message
End Sub

' ========================================================================
' Data Validation
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
' Currency Handling
' ========================================================================

' Get currency symbol
Public Function GetCurrencySymbol(currency As String) As String
    Select Case UCase(Trim(currency))
        Case "BRL"
            GetCurrencySymbol = "R$"
        Case "USD"
            GetCurrencySymbol = "$"
        Case "EUR"
            GetCurrencySymbol = "€"
        Case Else
            GetCurrencySymbol = ""
    End Select
End Function

' Format currency value
Public Function FormatCurrency(value As Double, currency As String) As String
    Dim symbol As String
    symbol = GetCurrencySymbol(currency)

    FormatCurrency = symbol & " " & Format(value, "#,##0.00")
End Function

' ========================================================================
' Array Operations
' ========================================================================

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
' Excel Operations
' ========================================================================

' Clear sheet contents (keep headers)
Public Sub ClearSheetData(ws As Worksheet)
    On Error Resume Next

    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    If lastRow > 1 Then
        ws.Range("A2:ZZ" & lastRow).ClearContents
    End If
End Sub

' Get last row with data
Public Function GetLastRow(ws As Worksheet, Optional col As Integer = 1) As Long
    On Error Resume Next

    GetLastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).row
End Function

' Auto-fit columns
Public Sub AutoFitColumns(ws As Worksheet, fromCol As String, toCol As String)
    On Error Resume Next

    ws.Range(fromCol & ":" & toCol).Columns.AutoFit
End Sub

' ========================================================================
' Error Message Formatting
' ========================================================================

' Format error message
Public Function FormatErrorMessage(moduleName As String, functionName As String, errorDesc As String) As String
    FormatErrorMessage = "Error in " & moduleName & "." & functionName & ": " & errorDesc
End Function

' ========================================================================
' Transaction Hash (for duplicate detection)
' ========================================================================

' Generate hash for duplicate detection
Public Function GenerateTransactionHash(transDate As Date, amount As Double, description As String) As String
    On Error Resume Next

    Dim hashInput As String

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
