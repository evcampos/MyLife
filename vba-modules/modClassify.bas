Attribute VB_Name = "modClassify"
' ========================================================================
' Module: modClassify
' Purpose: Category Classification Engine
' Description: Classifies transactions using exact and fuzzy matching
'              against a configurable categories table
' Platform: Excel for MacOS
' Author: Financial Import System v1.0
' ========================================================================

Option Explicit

' Classification result structure
Private Type ClassificationResult
    Success As Boolean
    Category As String
    Subcategory As String
    MatchedKeyword As String
    MatchScore As Integer
    MatchMethod As String ' "EXACT" or "FUZZY"
End Type

' Category rule structure
Private Type CategoryRule
    Category As String
    Subcategory As String
    Keywords As String
    Priority As Integer
    IsActive As Boolean
End Type

' ========================================================================
' Main Classification Function - Classify All Transactions
' ========================================================================
Public Function ClassifyAllTransactions() As Long
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim description As String
    Dim category As String
    Dim subcategory As String
    Dim classifiedCount As Long
    Dim result As ClassificationResult

    modUtils.LogMessage "Starting transaction classification", "INFO"
    modUtils.ShowProgress "Classifying transactions..."

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_TRANSACTIONS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    classifiedCount = 0

    ' Load category rules
    Call LoadCategoryRules

    ' Process each transaction
    For i = 2 To lastRow
        modUtils.ShowProgress "Classifying transactions...", i - 1, lastRow - 1

        ' Skip if already categorized (transfers, investments)
        category = ws.Cells(i, 5).Value
        If category <> "" And category <> "Unclassified" Then
            GoTo NextTransaction
        End If

        ' Get description
        description = ws.Cells(i, 3).Value

        ' Classify transaction
        result = ClassifyTransaction(description)

        If result.Success Then
            ' Update transaction
            ws.Cells(i, 5).Value = result.Category
            ws.Cells(i, 6).Value = result.Subcategory

            classifiedCount = classifiedCount + 1

            modUtils.LogMessage "Classified: " & description & " -> " & _
                               result.Category & "/" & result.Subcategory & _
                               " [" & result.MatchMethod & " - " & result.MatchScore & "%]", "DEBUG"
        Else
            ' Mark as unclassified
            ws.Cells(i, 5).Value = "Unclassified"

            ' Log error
            Call modErrors.LogClassificationError( _
                ws.Cells(i, 1).Value, _
                ws.Cells(i, 2).Value, _
                description, _
                ws.Cells(i, 4).Value, _
                "No matching category found")
        End If

NextTransaction:
    Next i

    modUtils.ClearProgress
    modUtils.LogMessage "Classification complete: " & classifiedCount & " transactions classified", "INFO"

    ClassifyAllTransactions = classifiedCount

    Exit Function

ErrorHandler:
    modUtils.LogMessage "Error in ClassifyAllTransactions: " & Err.Description, "ERROR"
    modUtils.ClearProgress
    ClassifyAllTransactions = 0
End Function

' ========================================================================
' Classify Single Transaction
' ========================================================================
Private Function ClassifyTransaction(description As String) As ClassificationResult
    On Error Resume Next

    Dim result As ClassificationResult
    Dim exactResult As ClassificationResult
    Dim fuzzyResult As ClassificationResult
    Dim fuzzyThreshold As Integer

    ' Initialize
    result.Success = False
    result.MatchScore = 0

    ' Phase 1: Try exact matching first
    exactResult = ExactMatch(description)

    If exactResult.Success Then
        ClassifyTransaction = exactResult
        Exit Function
    End If

    ' Phase 2: Try fuzzy matching
    fuzzyThreshold = modConfig.GetConfigInt("FuzzyThreshold", 80)
    fuzzyResult = FuzzyMatch(description, fuzzyThreshold)

    If fuzzyResult.Success Then
        ClassifyTransaction = fuzzyResult
        Exit Function
    End If

    ' No match found
    result.Success = False
    ClassifyTransaction = result
End Function

' ========================================================================
' Phase 1: Exact Keyword Matching
' ========================================================================
Private Function ExactMatch(description As String) As ClassificationResult
    On Error Resume Next

    Dim result As ClassificationResult
    Dim rules As Collection
    Dim rule As CategoryRule
    Dim cleanDesc As String
    Dim keywords() As String
    Dim keyword As Variant

    result.Success = False
    result.MatchMethod = "EXACT"

    cleanDesc = modUtils.CleanText(description)

    ' Get category rules sorted by priority
    Set rules = GetCategoryRules()

    ' Try each rule
    For Each rule In rules
        If rule.IsActive Then
            keywords = Split(rule.Keywords, ",")

            ' Check each keyword
            For Each keyword In keywords
                If Trim(keyword) <> "" Then
                    If InStr(1, cleanDesc, modUtils.CleanText(CStr(keyword)), vbTextCompare) > 0 Then
                        ' Exact match found!
                        result.Success = True
                        result.Category = rule.Category
                        result.Subcategory = rule.Subcategory
                        result.MatchedKeyword = Trim(CStr(keyword))
                        result.MatchScore = 100
                        result.MatchMethod = "EXACT"

                        ExactMatch = result
                        Exit Function
                    End If
                End If
            Next keyword
        End If
    Next rule

    ExactMatch = result
End Function

' ========================================================================
' Phase 2: Fuzzy Matching
' ========================================================================
Private Function FuzzyMatch(description As String, threshold As Integer) As ClassificationResult
    On Error Resume Next

    Dim result As ClassificationResult
    Dim rules As Collection
    Dim rule As CategoryRule
    Dim cleanDesc As String
    Dim keywords() As String
    Dim keyword As Variant
    Dim score As Integer
    Dim bestScore As Integer
    Dim bestRule As CategoryRule
    Dim bestKeyword As String

    result.Success = False
    result.MatchMethod = "FUZZY"

    cleanDesc = modUtils.CleanText(description)
    bestScore = 0

    ' Get category rules
    Set rules = GetCategoryRules()

    ' Try fuzzy matching against each rule
    For Each rule In rules
        If rule.IsActive Then
            keywords = Split(rule.Keywords, ",")

            ' Check each keyword
            For Each keyword In keywords
                If Trim(keyword) <> "" Then
                    ' Calculate similarity score
                    score = modUtils.FuzzyMatch(cleanDesc, CStr(keyword))

                    If score > bestScore Then
                        bestScore = score
                        bestRule = rule
                        bestKeyword = Trim(CStr(keyword))
                    End If
                End If
            Next keyword
        End If
    Next rule

    ' Check if best score meets threshold
    If bestScore >= threshold Then
        result.Success = True
        result.Category = bestRule.Category
        result.Subcategory = bestRule.Subcategory
        result.MatchedKeyword = bestKeyword
        result.MatchScore = bestScore
        result.MatchMethod = "FUZZY"
    End If

    FuzzyMatch = result
End Function

' ========================================================================
' Load Category Rules (Private Cache)
' ========================================================================
Private categoryRulesCache As Collection

Private Sub LoadCategoryRules()
    On Error Resume Next

    ' Clear cache
    Set categoryRulesCache = New Collection

    ' Load from Categories sheet
    Set categoryRulesCache = GetCategoryRules()
End Sub

' ========================================================================
' Get Category Rules from Categories Sheet
' ========================================================================
Private Function GetCategoryRules() As Collection
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim rules As Collection
    Dim rule As CategoryRule
    Dim lastRow As Long
    Dim i As Long

    Set rules = New Collection

    ' Check if Categories sheet exists
    If Not modConfig.SheetExists(modConfig.SHEET_CATEGORIES) Then
        Call InitializeCategoriesSheet
    End If

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_CATEGORIES)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Load each rule
    For i = 2 To lastRow
        rule.Category = Trim(ws.Cells(i, 1).Value)
        rule.Subcategory = Trim(ws.Cells(i, 2).Value)
        rule.Keywords = Trim(ws.Cells(i, 3).Value)

        ' Priority (default 100)
        If IsNumeric(ws.Cells(i, 4).Value) Then
            rule.Priority = CInt(ws.Cells(i, 4).Value)
        Else
            rule.Priority = 100
        End If

        ' IsActive (default True)
        Dim activeValue As String
        activeValue = UCase(Trim(ws.Cells(i, 5).Value))
        rule.IsActive = (activeValue = "" Or activeValue = "TRUE" Or activeValue = "YES" Or activeValue = "1")

        ' Only add valid rules
        If rule.Category <> "" And rule.Keywords <> "" Then
            rules.Add rule
        End If
    Next i

    ' Sort rules by priority (higher priority first)
    Set rules = SortRulesByPriority(rules)

    Set GetCategoryRules = rules
    Exit Function

ErrorHandler:
    modUtils.LogMessage "Error loading category rules: " & Err.Description, "ERROR"
    Set GetCategoryRules = New Collection
End Function

' ========================================================================
' Sort Rules by Priority (Simple Bubble Sort)
' ========================================================================
Private Function SortRulesByPriority(rules As Collection) As Collection
    On Error Resume Next

    Dim i As Long, j As Long
    Dim temp As CategoryRule
    Dim arr() As CategoryRule
    Dim sortedRules As Collection

    ' Convert collection to array
    ReDim arr(1 To rules.Count)
    For i = 1 To rules.Count
        arr(i) = rules(i)
    Next i

    ' Bubble sort (descending priority)
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j).Priority > arr(i).Priority Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i

    ' Convert back to collection
    Set sortedRules = New Collection
    For i = 1 To UBound(arr)
        sortedRules.Add arr(i)
    Next i

    Set SortRulesByPriority = sortedRules
End Function

' ========================================================================
' Initialize Categories Sheet with Default Categories
' ========================================================================
Public Sub InitializeCategoriesSheet()
    On Error Resume Next

    Dim ws As Worksheet

    ' Create sheet
    If Not modConfig.SheetExists(modConfig.SHEET_CATEGORIES) Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = modConfig.SHEET_CATEGORIES
    Else
        Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_CATEGORIES)
        ws.Cells.Clear
    End If

    ' Setup headers
    With ws
        .Cells(1, 1).Value = "Category"
        .Cells(1, 2).Value = "Subcategory"
        .Cells(1, 3).Value = "Keywords"
        .Cells(1, 4).Value = "Priority"
        .Cells(1, 5).Value = "IsActive"

        ' Format headers
        .Range("A1:E1").Font.Bold = True
        .Range("A1:E1").Interior.Color = RGB(68, 114, 196)
        .Range("A1:E1").Font.Color = RGB(255, 255, 255)
    End With

    ' Add default categories
    Call AddDefaultCategories(ws)

    ' Auto-fit columns
    ws.Columns("A:E").AutoFit

    modUtils.LogMessage "Categories sheet initialized with default categories", "INFO"
End Sub

' ========================================================================
' Add Default Category Rules
' ========================================================================
Private Sub AddDefaultCategories(ws As Worksheet)
    On Error Resume Next

    Dim row As Long
    row = 2

    ' Income categories
    Call AddCategory(ws, row, "Income", "Salary", "SALARY,PAYROLL,WAGE,SALARIO,VENCIMENTO", 1)
    Call AddCategory(ws, row, "Income", "Bonus", "BONUS,INCENTIVE,COMMISSION", 2)
    Call AddCategory(ws, row, "Income", "Other Income", "REFUND,REIMBURSEMENT,REEMBOLSO", 3)

    ' Food & Dining
    Call AddCategory(ws, row, "Food & Dining", "Restaurants", "RESTAURANT,CAFE,COFFEE,BAR,DINER,PIZZERIA", 10)
    Call AddCategory(ws, row, "Food & Dining", "Groceries", "GROCERY,SUPERMARKET,MARKET,WHOLE FOODS,TRADER,SAFEWAY", 10)
    Call AddCategory(ws, row, "Food & Dining", "Fast Food", "MCDONALDS,BURGER,TACO,SUBWAY,KFC,WENDYS", 10)

    ' Transportation
    Call AddCategory(ws, row, "Transportation", "Gas & Fuel", "GAS,FUEL,SHELL,CHEVRON,EXXON,PETROL", 10)
    Call AddCategory(ws, row, "Transportation", "Public Transit", "METRO,BUS,TRAIN,SUBWAY,TRANSIT,UBER,LYFT,TAXI", 10)
    Call AddCategory(ws, row, "Transportation", "Parking", "PARKING,PARK,GARAGE", 10)
    Call AddCategory(ws, row, "Transportation", "Auto Maintenance", "AUTO,REPAIR,MECHANIC,OIL CHANGE,TIRE", 10)

    ' Shopping
    Call AddCategory(ws, row, "Shopping", "Clothing", "CLOTHING,APPAREL,FASHION,ZARA,H&M,GAP", 10)
    Call AddCategory(ws, row, "Shopping", "Electronics", "ELECTRONICS,APPLE,BEST BUY,COMPUTER,PHONE", 10)
    Call AddCategory(ws, row, "Shopping", "General Merchandise", "AMAZON,WALMART,TARGET,COSTCO,SHOPPING", 10)

    ' Bills & Utilities
    Call AddCategory(ws, row, "Bills & Utilities", "Electric", "ELECTRIC,ELECTRICITY,POWER,UTILITY", 10)
    Call AddCategory(ws, row, "Bills & Utilities", "Water", "WATER,SEWER", 10)
    Call AddCategory(ws, row, "Bills & Utilities", "Internet & Phone", "INTERNET,PHONE,MOBILE,CELLULAR,VERIZON,ATT,COMCAST", 10)
    Call AddCategory(ws, row, "Bills & Utilities", "Cable & Streaming", "CABLE,NETFLIX,HULU,DISNEY,SPOTIFY,HBO", 10)

    ' Housing
    Call AddCategory(ws, row, "Housing", "Rent", "RENT,RENTAL,LEASE", 5)
    Call AddCategory(ws, row, "Housing", "Mortgage", "MORTGAGE,HOME LOAN", 5)
    Call AddCategory(ws, row, "Housing", "Home Improvement", "HOME DEPOT,LOWES,HARDWARE,IMPROVEMENT", 10)

    ' Healthcare
    Call AddCategory(ws, row, "Healthcare", "Doctor", "DOCTOR,PHYSICIAN,MEDICAL,CLINIC", 10)
    Call AddCategory(ws, row, "Healthcare", "Pharmacy", "PHARMACY,CVS,WALGREENS,PRESCRIPTION,MEDICINE", 10)
    Call AddCategory(ws, row, "Healthcare", "Dental", "DENTAL,DENTIST,ORTHODONTIC", 10)

    ' Entertainment
    Call AddCategory(ws, row, "Entertainment", "Movies & Events", "MOVIE,CINEMA,THEATER,CONCERT,EVENT,TICKET", 10)
    Call AddCategory(ws, row, "Entertainment", "Hobbies", "HOBBY,SPORT,GYM,FITNESS", 10)

    ' Financial
    Call AddCategory(ws, row, "Financial", "Bank Fees", "FEE,CHARGE,SERVICE CHARGE,BANK FEE", 10)
    Call AddCategory(ws, row, "Financial", "Interest Charges", "INTEREST CHARGE,FINANCE CHARGE", 10)
    Call AddCategory(ws, row, "Financial", "ATM Withdrawal", "ATM,CASH WITHDRAWAL,SAQUE", 10)

    ' Personal
    Call AddCategory(ws, row, "Personal", "Hair & Beauty", "SALON,BARBER,SPA,BEAUTY", 10)
    Call AddCategory(ws, row, "Personal", "Personal Care", "PERSONAL CARE,COSMETIC", 10)

    ' Education
    Call AddCategory(ws, row, "Education", "Tuition", "TUITION,SCHOOL,UNIVERSITY,COLLEGE", 5)
    Call AddCategory(ws, row, "Education", "Books & Supplies", "BOOK,TEXTBOOK,SUPPLIES,STATIONERY", 10)

    ' Miscellaneous
    Call AddCategory(ws, row, "Miscellaneous", "Other", "MISC,OTHER,VARIOUS", 200)
End Sub

' ========================================================================
' Add Single Category
' ========================================================================
Private Sub AddCategory(ws As Worksheet, ByRef row As Long, _
                       category As String, subcategory As String, _
                       keywords As String, priority As Integer)
    ws.Cells(row, 1).Value = category
    ws.Cells(row, 2).Value = subcategory
    ws.Cells(row, 3).Value = keywords
    ws.Cells(row, 4).Value = priority
    ws.Cells(row, 5).Value = "TRUE"
    row = row + 1
End Sub

' ========================================================================
' Manual Category Assignment
' ========================================================================
Public Sub AssignCategory(transactionID As String, category As String, subcategory As String)
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_TRANSACTIONS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Find and update transaction
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = transactionID Then
            ws.Cells(i, 5).Value = category
            ws.Cells(i, 6).Value = subcategory

            modUtils.LogMessage "Manual category assignment: " & transactionID & " -> " & category & "/" & subcategory, "INFO"
            Exit Sub
        End If
    Next i
End Sub

' ========================================================================
' Get Classification Statistics
' ========================================================================
Public Function GetClassificationStats() As String
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim classifiedCount As Long
    Dim unclassifiedCount As Long
    Dim category As String

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_TRANSACTIONS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        category = ws.Cells(i, 5).Value

        If category = "" Or category = "Unclassified" Then
            unclassifiedCount = unclassifiedCount + 1
        Else
            classifiedCount = classifiedCount + 1
        End If
    Next i

    GetClassificationStats = "Classified: " & classifiedCount & vbCrLf & _
                            "Unclassified: " & unclassifiedCount & vbCrLf & _
                            "Success Rate: " & Format(classifiedCount / (classifiedCount + unclassifiedCount), "0%")
End Function

' ========================================================================
' Suggest New Category Mapping
' ========================================================================
Public Sub SuggestMapping(description As String, category As String, subcategory As String)
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim keyword As String

    ' Extract potential keyword from description
    keyword = ExtractKeyword(description)

    If keyword = "" Then Exit Sub

    ' Add to Categories sheet
    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_CATEGORIES)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(lastRow, 1).Value = category
    ws.Cells(lastRow, 2).Value = subcategory
    ws.Cells(lastRow, 3).Value = keyword
    ws.Cells(lastRow, 4).Value = 50 ' Medium priority
    ws.Cells(lastRow, 5).Value = "TRUE"

    modUtils.LogMessage "New mapping suggested: " & keyword & " -> " & category & "/" & subcategory, "INFO"
End Sub

' ========================================================================
' Extract Keyword from Description
' ========================================================================
Private Function ExtractKeyword(description As String) As String
    On Error Resume Next

    Dim words() As String
    Dim word As Variant
    Dim cleanDesc As String

    cleanDesc = modUtils.CleanText(description)

    ' Remove common noise words
    cleanDesc = Replace(cleanDesc, "COMPRA ", "")
    cleanDesc = Replace(cleanDesc, "PAGAMENTO ", "")
    cleanDesc = Replace(cleanDesc, "PURCHASE ", "")

    ' Split into words
    words = Split(cleanDesc, " ")

    ' Return first significant word
    For Each word In words
        If Len(word) > 3 Then
            ExtractKeyword = CStr(word)
            Exit Function
        End If
    Next word

    ExtractKeyword = ""
End Function
