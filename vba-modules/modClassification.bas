Attribute VB_Name = "modClassification"
' ========================================================================
' Module: modClassification
' Purpose: Transaction Classification Engine
' Description: Classifies transactions using exact matching and proximity
'              matching algorithms with learning capability
' Platform: Excel for MacOS
' Author: [MY LIFE] Financial System v1.0
' ========================================================================

Option Explicit

' Classification threshold for proximity matching (0-100)
Private Const SIMILARITY_THRESHOLD As Integer = 70

' ========================================================================
' Main Classification Function - Classify All Transactions
' ========================================================================
Public Function ClassifyAllTransactions() As Long
    On Error GoTo ErrorHandler

    Dim totalClassified As Long

    modUtils.ShowProgress "Classifying transactions..."
    modUtils.LogDebug "Starting transaction classification"

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Classify banks
    totalClassified = totalClassified + ClassifyBankTransactions()

    ' Classify cards (already have categories, but may need refinement)
    ' Cards are usually pre-categorized by the bank, so we skip them
    ' unless the category is empty

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    modUtils.ClearProgress
    modUtils.LogDebug "Classification complete: " & totalClassified & " transactions"

    ClassifyAllTransactions = totalClassified

    Exit Function

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    modUtils.ClearProgress
    MsgBox modUtils.FormatErrorMessage("modClassification", "ClassifyAllTransactions", Err.Description), vbCritical
    ClassifyAllTransactions = 0
End Function

' ========================================================================
' Classify Bank Transactions
' ========================================================================
Private Function ClassifyBankTransactions() As Long
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim description As String
    Dim category As String
    Dim subcategory As String
    Dim classified As Long

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_BANKS)
    lastRow = modUtils.GetLastRow(ws, 1)

    classified = 0

    For i = 2 To lastRow
        ' Check if already classified
        If Trim(CStr(ws.Cells(i, 5).Value)) = "" Then
            description = Trim(CStr(ws.Cells(i, 3).Value))

            If description <> "" Then
                ' Attempt classification
                If ClassifyTransaction(description, category, subcategory) Then
                    ws.Cells(i, 5).Value = category
                    ws.Cells(i, 6).Value = subcategory
                    classified = classified + 1
                End If
            End If
        End If
    Next i

    ClassifyBankTransactions = classified

    Exit Function

ErrorHandler:
    MsgBox modUtils.FormatErrorMessage("modClassification", "ClassifyBankTransactions", Err.Description), vbCritical
    ClassifyBankTransactions = 0
End Function

' ========================================================================
' Classify Single Transaction
' ========================================================================
Public Function ClassifyTransaction(description As String, _
                                    ByRef outCategory As String, _
                                    ByRef outSubcategory As String) As Boolean
    On Error GoTo ErrorHandler

    Dim cleanDesc As String
    cleanDesc = modUtils.CleanText(description)

    ' Phase 1: Exact keyword match
    If ExactMatch(cleanDesc, outCategory, outSubcategory) Then
        ClassifyTransaction = True
        Exit Function
    End If

    ' Phase 2: Proximity match
    If ProximityMatch(cleanDesc, outCategory, outSubcategory) Then
        ClassifyTransaction = True
        Exit Function
    End If

    ' Phase 3: No match found
    outCategory = "Unclassified"
    outSubcategory = ""
    ClassifyTransaction = False

    Exit Function

ErrorHandler:
    ClassifyTransaction = False
End Function

' ========================================================================
' Exact Keyword Match
' ========================================================================
Private Function ExactMatch(description As String, _
                            ByRef outCategory As String, _
                            ByRef outSubcategory As String) As Boolean
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim category As String
    Dim subcategory As String
    Dim keywords As String
    Dim priority As Integer
    Dim bestPriority As Integer
    Dim found As Boolean

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_CATEGORIES)
    lastRow = modUtils.GetLastRow(ws, 1)

    found = False
    bestPriority = 0

    ' Search through categories
    For i = 2 To lastRow
        category = Trim(CStr(ws.Cells(i, 1).Value))
        subcategory = Trim(CStr(ws.Cells(i, 2).Value))
        keywords = Trim(CStr(ws.Cells(i, 3).Value))
        priority = CInt(ws.Cells(i, 4).Value)

        If keywords <> "" Then
            If modUtils.ContainsAny(description, keywords) Then
                ' Found a match - check if it's higher priority
                If priority >= bestPriority Then
                    outCategory = category
                    outSubcategory = subcategory
                    bestPriority = priority
                    found = True
                End If
            End If
        End If
    Next i

    ExactMatch = found

    Exit Function

ErrorHandler:
    ExactMatch = False
End Function

' ========================================================================
' Proximity Match (Similarity-based)
' ========================================================================
Private Function ProximityMatch(description As String, _
                                ByRef outCategory As String, _
                                ByRef outSubcategory As String) As Boolean
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim category As String
    Dim subcategory As String
    Dim keywords As String
    Dim similarity As Integer
    Dim bestSimilarity As Integer
    Dim keywordArray() As String
    Dim keyword As Variant
    Dim maxSim As Integer

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_CATEGORIES)
    lastRow = modUtils.GetLastRow(ws, 1)

    bestSimilarity = 0

    ' Search through categories
    For i = 2 To lastRow
        category = Trim(CStr(ws.Cells(i, 1).Value))
        subcategory = Trim(CStr(ws.Cells(i, 2).Value))
        keywords = Trim(CStr(ws.Cells(i, 3).Value))

        If keywords <> "" Then
            ' Check similarity with each keyword
            keywordArray = Split(keywords, ",")
            maxSim = 0

            For Each keyword In keywordArray
                similarity = modUtils.TextSimilarity(description, CStr(keyword))
                If similarity > maxSim Then
                    maxSim = similarity
                End If
            Next keyword

            ' If best match so far, save it
            If maxSim > bestSimilarity And maxSim >= SIMILARITY_THRESHOLD Then
                outCategory = category
                outSubcategory = subcategory
                bestSimilarity = maxSim
            End If
        End If
    Next i

    ProximityMatch = (bestSimilarity >= SIMILARITY_THRESHOLD)

    Exit Function

ErrorHandler:
    ProximityMatch = False
End Function

' ========================================================================
' Get Unclassified Transactions
' ========================================================================
Public Function GetUnclassifiedCount() As Long
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim count As Long

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_BANKS)
    lastRow = modUtils.GetLastRow(ws, 1)

    count = 0

    For i = 2 To lastRow
        If Trim(CStr(ws.Cells(i, 5).Value)) = "" Or _
           CStr(ws.Cells(i, 5).Value) = "Unclassified" Then
            count = count + 1
        End If
    Next i

    GetUnclassifiedCount = count

    Exit Function

ErrorHandler:
    GetUnclassifiedCount = 0
End Function

' ========================================================================
' Learn From Manual Classification
' ========================================================================
Public Sub LearnFromManualClassification(description As String, _
                                         category As String, _
                                         subcategory As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim existingKeywords As String
    Dim newKeyword As String
    Dim found As Boolean

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_CATEGORIES)
    lastRow = modUtils.GetLastRow(ws, 1)

    ' Extract significant keyword from description (first word usually)
    Dim words() As String
    words = Split(modUtils.CleanText(description), " ")

    If UBound(words) >= 0 Then
        newKeyword = words(0)  ' Use first word as keyword

        ' Find category row
        found = False
        For i = 2 To lastRow
            If Trim(CStr(ws.Cells(i, 1).Value)) = category And _
               Trim(CStr(ws.Cells(i, 2).Value)) = subcategory Then
                existingKeywords = Trim(CStr(ws.Cells(i, 3).Value))

                ' Add keyword if not already present
                If Not modUtils.ContainsAny(existingKeywords, newKeyword) Then
                    If existingKeywords <> "" Then
                        ws.Cells(i, 3).Value = existingKeywords & "," & newKeyword
                    Else
                        ws.Cells(i, 3).Value = newKeyword
                    End If
                End If

                found = True
                Exit For
            End If
        Next i

        ' If category doesn't exist, create it
        If Not found Then
            lastRow = lastRow + 1
            ws.Cells(lastRow, 1).Value = category
            ws.Cells(lastRow, 2).Value = subcategory
            ws.Cells(lastRow, 3).Value = newKeyword
            ws.Cells(lastRow, 4).Value = 5  ' Medium priority
        End If
    End If

    Exit Sub

ErrorHandler:
    modUtils.LogDebug "Error learning from classification: " & Err.Description
End Sub

' ========================================================================
' Reclassify All (useful after adding new rules)
' ========================================================================
Public Sub ReclassifyAll()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Clear existing classifications
    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_BANKS)
    lastRow = modUtils.GetLastRow(ws, 1)

    For i = 2 To lastRow
        ws.Cells(i, 5).Value = ""
        ws.Cells(i, 6).Value = ""
    Next i

    ' Re-run classification
    Call ClassifyAllTransactions

    MsgBox "Reclassification complete!", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox modUtils.FormatErrorMessage("modClassification", "ReclassifyAll", Err.Description), vbCritical
End Sub
