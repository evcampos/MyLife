Attribute VB_Name = "modErrors"
' ========================================================================
' Module: modErrors
' Purpose: Error Handling and Audit System
' Description: Manages error logging, unresolved transactions, and
'              provides audit trail functionality
' Platform: Excel for MacOS
' Author: Financial Import System v1.0
' ========================================================================

Option Explicit

' Error types
Public Const ERR_IMPORT As String = "IMPORT_ERROR"
Public Const ERR_NORMALIZATION As String = "NORMALIZATION_ERROR"
Public Const ERR_CLASSIFICATION As String = "CLASSIFICATION_ERROR"
Public Const ERR_VALIDATION As String = "VALIDATION_ERROR"

' Error status
Public Const STATUS_OPEN As String = "OPEN"
Public Const STATUS_RESOLVED As String = "RESOLVED"

' ========================================================================
' Initialize Errors Sheet
' ========================================================================
Public Sub InitializeErrorsSheet()
    On Error Resume Next

    Dim ws As Worksheet

    If Not modConfig.SheetExists(modConfig.SHEET_ERRORS) Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = modConfig.SHEET_ERRORS
    Else
        Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_ERRORS)
        ' Don't clear - preserve existing errors
    End If

    ' Setup headers if empty
    If ws.Cells(1, 1).Value = "" Then
        With ws
            .Cells(1, 1).Value = "ErrorID"
            .Cells(1, 2).Value = "ErrorType"
            .Cells(1, 3).Value = "TransactionID"
            .Cells(1, 4).Value = "Date"
            .Cells(1, 5).Value = "Description"
            .Cells(1, 6).Value = "Amount"
            .Cells(1, 7).Value = "IssueDescription"
            .Cells(1, 8).Value = "SuggestedCategory"
            .Cells(1, 9).Value = "SuggestedSubcategory"
            .Cells(1, 10).Value = "Status"
            .Cells(1, 11).Value = "CreatedDate"

            ' Format headers
            .Range("A1:K1").Font.Bold = True
            .Range("A1:K1").Interior.Color = RGB(192, 0, 0) ' Red for errors
            .Range("A1:K1").Font.Color = RGB(255, 255, 255)
        End With
    End If

    ws.Columns("A:K").AutoFit
End Sub

' ========================================================================
' Log Classification Error
' ========================================================================
Public Sub LogClassificationError(transactionID As String, transDate As Date, _
                                  description As String, amount As Double, _
                                  issueDesc As String)
    On Error Resume Next

    Call LogError(ERR_CLASSIFICATION, transactionID, transDate, _
                 description, amount, issueDesc, "", "")
End Sub

' ========================================================================
' Log Import Error
' ========================================================================
Public Sub LogImportError(sourceFile As String, issueDesc As String)
    On Error Resume Next

    Call LogError(ERR_IMPORT, "", Date, sourceFile, 0, issueDesc, "", "")
End Sub

' ========================================================================
' Log Normalization Error
' ========================================================================
Public Sub LogNormalizationError(transactionID As String, transDate As Date, _
                                 description As String, amount As Double, _
                                 issueDesc As String)
    On Error Resume Next

    Call LogError(ERR_NORMALIZATION, transactionID, transDate, _
                 description, amount, issueDesc, "", "")
End Sub

' ========================================================================
' Generic Log Error Function
' ========================================================================
Private Sub LogError(errorType As String, transactionID As String, transDate As Date, _
                    description As String, amount As Double, issueDesc As String, _
                    Optional suggestedCategory As String = "", _
                    Optional suggestedSubcategory As String = "")
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim errorID As String

    ' Initialize sheet if needed
    Call InitializeErrorsSheet

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_ERRORS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ' Generate error ID
    errorID = modUtils.GenerateUniqueID("ERR")

    ' Add error entry
    With ws
        .Cells(lastRow, 1).Value = errorID
        .Cells(lastRow, 2).Value = errorType
        .Cells(lastRow, 3).Value = transactionID
        .Cells(lastRow, 4).Value = transDate
        .Cells(lastRow, 5).Value = description
        .Cells(lastRow, 6).Value = amount
        .Cells(lastRow, 7).Value = issueDesc
        .Cells(lastRow, 8).Value = suggestedCategory
        .Cells(lastRow, 9).Value = suggestedSubcategory
        .Cells(lastRow, 10).Value = STATUS_OPEN
        .Cells(lastRow, 11).Value = Now

        ' Highlight error row
        .Range(.Cells(lastRow, 1), .Cells(lastRow, 11)).Interior.Color = RGB(255, 235, 235)
    End With
End Sub

' ========================================================================
' Get All Open Errors
' ========================================================================
Public Function GetOpenErrors() As Collection
    On Error Resume Next

    Dim ws As Worksheet
    Dim errors As Collection
    Dim lastRow As Long
    Dim i As Long

    Set errors = New Collection

    If Not modConfig.SheetExists(modConfig.SHEET_ERRORS) Then
        Set GetOpenErrors = errors
        Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_ERRORS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        If ws.Cells(i, 10).Value = STATUS_OPEN Then
            errors.Add Array(ws.Cells(i, 1).Value, ws.Cells(i, 2).Value, _
                           ws.Cells(i, 5).Value, ws.Cells(i, 7).Value)
        End If
    Next i

    Set GetOpenErrors = errors
End Function

' ========================================================================
' Resolve Error
' ========================================================================
Public Sub ResolveError(errorID As String, Optional category As String = "", _
                       Optional subcategory As String = "")
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim transactionID As String

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_ERRORS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Find and resolve error
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = errorID Then
            ' Update status
            ws.Cells(i, 10).Value = STATUS_RESOLVED

            ' Update suggested resolution
            If category <> "" Then
                ws.Cells(i, 8).Value = category
                ws.Cells(i, 9).Value = subcategory
            End If

            ' Clear highlighting
            ws.Range(ws.Cells(i, 1), ws.Cells(i, 11)).Interior.Color = RGB(235, 255, 235) ' Light green

            ' Update transaction if provided
            transactionID = ws.Cells(i, 3).Value
            If transactionID <> "" And category <> "" Then
                Call modClassify.AssignCategory(transactionID, category, subcategory)

                ' Suggest adding to category mappings
                Call modClassify.SuggestMapping(ws.Cells(i, 5).Value, category, subcategory)
            End If

            modUtils.LogMessage "Error resolved: " & errorID, "INFO"
            Exit Sub
        End If
    Next i
End Sub

' ========================================================================
' Clear All Resolved Errors
' ========================================================================
Public Sub ClearResolvedErrors()
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim deleteCount As Long

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_ERRORS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Delete resolved errors (from bottom to top)
    For i = lastRow To 2 Step -1
        If ws.Cells(i, 10).Value = STATUS_RESOLVED Then
            ws.Rows(i).Delete
            deleteCount = deleteCount + 1
        End If
    Next i

    modUtils.LogMessage "Cleared " & deleteCount & " resolved errors", "INFO"

    ' If no errors remain, hide or delete sheet
    If ws.Cells(ws.Rows.Count, 1).End(xlUp).Row = 1 Then
        ws.Visible = xlSheetVeryHidden
        MsgBox "All errors have been resolved!", vbInformation
    Else
        MsgBox deleteCount & " resolved errors have been cleared.", vbInformation
    End If
End Sub

' ========================================================================
' Clear All Errors
' ========================================================================
Public Sub ClearAllErrors()
    On Error Resume Next

    Dim ws As Worksheet
    Dim response As VbMsgBoxResult

    If Not modConfig.SheetExists(modConfig.SHEET_ERRORS) Then Exit Sub

    response = MsgBox("This will delete ALL error records. Are you sure?", _
                     vbYesNo + vbExclamation, "Clear All Errors")

    If response = vbYes Then
        Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_ERRORS)
        ws.Cells.Clear

        ' Recreate headers
        Call InitializeErrorsSheet

        modUtils.LogMessage "All errors cleared", "INFO"
        MsgBox "All errors have been cleared.", vbInformation
    End If
End Sub

' ========================================================================
' Get Error Statistics
' ========================================================================
Public Function GetErrorStats() As String
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim openCount As Long
    Dim resolvedCount As Long
    Dim byType As Object

    Set byType = CreateObject("Scripting.Dictionary")

    If Not modConfig.SheetExists(modConfig.SHEET_ERRORS) Then
        GetErrorStats = "No errors found"
        Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_ERRORS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        If ws.Cells(i, 10).Value = STATUS_OPEN Then
            openCount = openCount + 1
        Else
            resolvedCount = resolvedCount + 1
        End If

        ' Count by type
        Dim errType As String
        errType = ws.Cells(i, 2).Value
        If Not byType.exists(errType) Then
            byType.Add errType, 0
        End If
        byType(errType) = byType(errType) + 1
    Next i

    Dim stats As String
    stats = "Error Statistics:" & vbCrLf & _
            "Open: " & openCount & vbCrLf & _
            "Resolved: " & resolvedCount & vbCrLf & vbCrLf & _
            "By Type:" & vbCrLf

    Dim key As Variant
    For Each key In byType.Keys
        stats = stats & "  " & key & ": " & byType(key) & vbCrLf
    Next key

    GetErrorStats = stats
End Function

' ========================================================================
' Check for Unresolved Errors
' ========================================================================
Public Function HasUnresolvedErrors() As Boolean
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    If Not modConfig.SheetExists(modConfig.SHEET_ERRORS) Then
        HasUnresolvedErrors = False
        Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_ERRORS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        If ws.Cells(i, 10).Value = STATUS_OPEN Then
            HasUnresolvedErrors = True
            Exit Function
        End If
    Next i

    HasUnresolvedErrors = False
End Function

' ========================================================================
' Export Errors to CSV
' ========================================================================
Public Sub ExportErrorsToCSV()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim filePath As String
    Dim fileNum As Integer
    Dim i As Long, j As Integer
    Dim line As String

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_ERRORS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "No errors to export.", vbInformation
        Exit Sub
    End If

    ' Get save location
    filePath = Application.GetSaveAsFilename( _
        InitialFileName:="Errors_" & Format(Now, "yyyymmdd_hhnnss") & ".csv", _
        FileFilter:="CSV Files (*.csv), *.csv")

    If filePath = "False" Then Exit Sub

    ' Write to CSV
    fileNum = FreeFile
    Open filePath For Output As #fileNum

    ' Write headers
    line = ""
    For j = 1 To 11
        line = line & ws.Cells(1, j).Value
        If j < 11 Then line = line & ","
    Next j
    Print #fileNum, line

    ' Write data
    For i = 2 To lastRow
        line = ""
        For j = 1 To 11
            line = line & """" & ws.Cells(i, j).Value & """"
            If j < 11 Then line = line & ","
        Next j
        Print #fileNum, line
    Next i

    Close #fileNum

    modUtils.LogMessage "Errors exported to: " & filePath, "INFO"
    MsgBox "Errors exported successfully to:" & vbCrLf & filePath, vbInformation

    Exit Sub

ErrorHandler:
    On Error Resume Next
    Close #fileNum
    MsgBox "Error exporting errors: " & Err.Description, vbCritical
End Sub

' ========================================================================
' Validate All Transactions
' ========================================================================
Public Function ValidateAllTransactions() As Long
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim errorCount As Long

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_TRANSACTIONS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    modUtils.LogMessage "Starting transaction validation", "INFO"

    For i = 2 To lastRow
        ' Check for missing required fields
        If Not modUtils.IsValidDate(ws.Cells(i, 2).Value) Then
            Call LogNormalizationError(ws.Cells(i, 1).Value, ws.Cells(i, 2).Value, _
                                      ws.Cells(i, 3).Value, ws.Cells(i, 4).Value, _
                                      "Invalid date")
            errorCount = errorCount + 1
        End If

        If ws.Cells(i, 3).Value = "" Then
            Call LogNormalizationError(ws.Cells(i, 1).Value, ws.Cells(i, 2).Value, _
                                      ws.Cells(i, 3).Value, ws.Cells(i, 4).Value, _
                                      "Missing description")
            errorCount = errorCount + 1
        End If

        If Not IsNumeric(ws.Cells(i, 4).Value) Then
            Call LogNormalizationError(ws.Cells(i, 1).Value, ws.Cells(i, 2).Value, _
                                      ws.Cells(i, 3).Value, ws.Cells(i, 4).Value, _
                                      "Invalid amount")
            errorCount = errorCount + 1
        End If
    Next i

    ValidateAllTransactions = errorCount

    modUtils.LogMessage "Validation complete: " & errorCount & " errors found", "INFO"
End Function

' ========================================================================
' Generate Error Report
' ========================================================================
Public Sub GenerateErrorReport()
    On Error Resume Next

    Dim stats As String

    stats = GetErrorStats()

    MsgBox stats, vbInformation, "Error Report"
End Sub
