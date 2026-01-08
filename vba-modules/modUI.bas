Attribute VB_Name = "modUI"
' ========================================================================
' Module: modUI
' Purpose: User Interface Components
' Description: Main orchestration functions and button handlers that
'              coordinate all system operations
' Platform: Excel for MacOS
' Author: Financial Import System v1.0
' ========================================================================

Option Explicit

' ========================================================================
' Initialize Complete System
' ========================================================================
Public Sub InitializeSystem()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    modUtils.LogMessage "=== System Initialization Started ===", "INFO"

    ' Initialize all components
    Call modConfig.InitializeConfig
    Call modClassify.InitializeCategoriesSheet
    Call modErrors.InitializeErrorsSheet
    Call CreateDashboard

    ' Show sheets in correct order
    Call ArrangeSheets

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    modUtils.LogMessage "=== System Initialization Complete ===", "INFO"

    MsgBox "System initialized successfully!" & vbCrLf & vbCrLf & _
           "Next steps:" & vbCrLf & _
           "1. Configure file paths in the Config sheet" & vbCrLf & _
           "2. Review and customize the Categories sheet" & vbCrLf & _
           "3. Click 'Import Transactions' to begin", _
           vbInformation, "Initialization Complete"

    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error during initialization: " & Err.Description, vbCritical
    modUtils.LogMessage "Error during initialization: " & Err.Description, "ERROR"
End Sub

' ========================================================================
' Main Import Process - Full Pipeline
' ========================================================================
Public Sub ImportTransactions()
    On Error GoTo ErrorHandler

    Dim startTime As Double
    Dim results As Collection
    Dim importCount As Long
    Dim normalizeCount As Long
    Dim transferCount As Long
    Dim investmentCount As Long
    Dim classifyCount As Long
    Dim errorCount As Long

    startTime = Timer

    ' Disable updates for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    modUtils.LogMessage "=== Full Import Process Started ===", "INFO"

    ' Step 1: Import files
    modUtils.ShowProgress "Step 1/6: Importing files..."
    Set results = modImport.ImportAllFiles()

    ' Count imported rows
    Dim result As modImport.ImportResult
    For Each result In results
        If result.Success Then
            importCount = importCount + result.RowsImported
        End If
    Next result

    If importCount = 0 Then
        MsgBox "No data imported. Please check your file paths in the Config sheet.", _
               vbExclamation, "Import Failed"
        GoTo Cleanup
    End If

    ' Step 2: Normalize data
    modUtils.ShowProgress "Step 2/6: Normalizing data..."
    normalizeCount = modNormalize.NormalizeAllData()

    ' Step 3: Detect transfers
    modUtils.ShowProgress "Step 3/6: Detecting transfers..."
    transferCount = modTransfers.DetectAllTransfers()

    ' Step 4: Detect investments
    modUtils.ShowProgress "Step 4/6: Detecting investments..."
    investmentCount = modInvestments.DetectAllInvestments()

    ' Step 5: Classify transactions
    modUtils.ShowProgress "Step 5/6: Classifying transactions..."
    classifyCount = modClassify.ClassifyAllTransactions()

    ' Step 6: Validate
    modUtils.ShowProgress "Step 6/6: Validating results..."
    errorCount = modErrors.ValidateAllTransactions()

    ' Show results
    Dim elapsed As Double
    elapsed = Timer - startTime

    Dim summary As String
    summary = "Import Complete!" & vbCrLf & vbCrLf & _
              "Statistics:" & vbCrLf & _
              "━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf & _
              "Rows Imported: " & importCount & vbCrLf & _
              "Transactions Normalized: " & normalizeCount & vbCrLf & _
              "Transfers Detected: " & transferCount & vbCrLf & _
              "Investments Detected: " & investmentCount & vbCrLf & _
              "Transactions Classified: " & classifyCount & vbCrLf & _
              "Validation Errors: " & errorCount & vbCrLf & vbCrLf & _
              "Processing Time: " & Format(elapsed, "0.0") & " seconds"

    modUtils.LogMessage "=== Full Import Process Complete ===", "INFO"
    modUtils.LogMessage summary, "INFO"

Cleanup:
    modUtils.ClearProgress
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    ' Show summary
    MsgBox summary, vbInformation, "Import Complete"

    ' Navigate to Transactions sheet
    ThisWorkbook.Worksheets(modConfig.SHEET_TRANSACTIONS).Activate

    Exit Sub

ErrorHandler:
    modUtils.ClearProgress
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error during import: " & Err.Description, vbCritical
    modUtils.LogMessage "Error during import: " & Err.Description, "ERROR"
End Sub

' ========================================================================
' Reprocess Existing Data
' ========================================================================
Public Sub ReprocessData()
    On Error GoTo ErrorHandler

    Dim response As VbMsgBoxResult

    response = MsgBox("This will re-classify and re-analyze all existing transactions." & vbCrLf & _
                     "Continue?", vbYesNo + vbQuestion, "Reprocess Data")

    If response = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    modUtils.LogMessage "=== Reprocessing Data ===", "INFO"

    ' Clear existing classifications
    Call modTransfers.ClearAllTransfers

    ' Re-run detection and classification
    Call modTransfers.DetectAllTransfers
    Call modInvestments.DetectAllInvestments
    Call modClassify.ClassifyAllTransactions

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Data reprocessed successfully!", vbInformation

    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error during reprocessing: " & Err.Description, vbCritical
End Sub

' ========================================================================
' Update Categories from Sheet
' ========================================================================
Public Sub UpdateCategories()
    On Error Resume Next

    ' Reload category rules
    Call modClassify.InitializeCategoriesSheet

    MsgBox "Category mappings updated." & vbCrLf & vbCrLf & _
           "Click 'Reprocess Data' to apply new mappings to existing transactions.", _
           vbInformation, "Categories Updated"

    modUtils.LogMessage "Category mappings updated", "INFO"
End Sub

' ========================================================================
' Clear All Data
' ========================================================================
Public Sub ClearAllData()
    On Error Resume Next

    Dim response As VbMsgBoxResult

    response = MsgBox("WARNING: This will delete ALL imported data and classifications!" & vbCrLf & vbCrLf & _
                     "Are you sure?", vbYesNo + vbCritical, "Clear All Data")

    If response = vbNo Then Exit Sub

    Application.DisplayAlerts = False

    ' Clear data sheets
    Dim ws As Worksheet

    If modConfig.SheetExists(modConfig.SHEET_TRANSACTIONS) Then
        ThisWorkbook.Worksheets(modConfig.SHEET_TRANSACTIONS).Cells.Clear
    End If

    If modConfig.SheetExists(modConfig.SHEET_RAWDATA) Then
        ThisWorkbook.Worksheets(modConfig.SHEET_RAWDATA).Cells.Clear
    End If

    If modConfig.SheetExists(modConfig.SHEET_AUDIT) Then
        ThisWorkbook.Worksheets(modConfig.SHEET_AUDIT).Cells.Clear
    End If

    Application.DisplayAlerts = True

    MsgBox "All data cleared.", vbInformation

    modUtils.LogMessage "All data cleared by user", "INFO"
End Sub

' ========================================================================
' Show Statistics Dashboard
' ========================================================================
Public Sub ShowStatistics()
    On Error Resume Next

    Dim stats As String

    stats = "System Statistics" & vbCrLf & _
            "═══════════════════════════════" & vbCrLf & vbCrLf

    stats = stats & modImport.GetImportStats() & vbCrLf & vbCrLf

    stats = stats & modNormalize.GetNormalizationStats() & vbCrLf & vbCrLf

    stats = stats & modTransfers.GetTransferStats() & vbCrLf & vbCrLf

    stats = stats & modInvestments.GetInvestmentStats() & vbCrLf & vbCrLf

    stats = stats & modClassify.GetClassificationStats() & vbCrLf & vbCrLf

    stats = stats & modErrors.GetErrorStats()

    MsgBox stats, vbInformation, "System Statistics"
End Sub

' ========================================================================
' Export Data to CSV
' ========================================================================
Public Sub ExportTransactions()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim filePath As String
    Dim fileNum As Integer
    Dim i As Long, j As Integer
    Dim line As String

    Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_TRANSACTIONS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "No transactions to export.", vbInformation
        Exit Sub
    End If

    ' Get save location
    filePath = Application.GetSaveAsFilename( _
        InitialFileName:="Transactions_" & Format(Now, "yyyymmdd_hhnnss") & ".csv", _
        FileFilter:="CSV Files (*.csv), *.csv")

    If filePath = "False" Then Exit Sub

    ' Write to CSV
    fileNum = FreeFile
    Open filePath For Output As #fileNum

    ' Write headers
    line = ""
    For j = 1 To 14
        line = line & ws.Cells(1, j).Value
        If j < 14 Then line = line & ","
    Next j
    Print #fileNum, line

    ' Write data
    For i = 2 To lastRow
        line = ""
        For j = 1 To 14
            Dim cellValue As String
            cellValue = ws.Cells(i, j).Value
            cellValue = Replace(cellValue, """", """""") ' Escape quotes
            line = line & """" & cellValue & """"
            If j < 14 Then line = line & ","
        Next j
        Print #fileNum, line
    Next i

    Close #fileNum

    MsgBox "Transactions exported successfully to:" & vbCrLf & filePath, vbInformation

    modUtils.LogMessage "Transactions exported to: " & filePath, "INFO"

    Exit Sub

ErrorHandler:
    On Error Resume Next
    Close #fileNum
    MsgBox "Error exporting transactions: " & Err.Description, vbCritical
End Sub

' ========================================================================
' Create Dashboard Sheet
' ========================================================================
Private Sub CreateDashboard()
    On Error Resume Next

    Dim ws As Worksheet
    Dim btn As Button
    Dim btnTop As Integer
    Dim btnLeft As Integer
    Dim btnWidth As Integer
    Dim btnHeight As Integer

    ' Create or get Dashboard sheet
    If Not modConfig.SheetExists(modConfig.SHEET_DASHBOARD) Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = modConfig.SHEET_DASHBOARD
        ws.Move Before:=ThisWorkbook.Worksheets(1)
    Else
        Set ws = ThisWorkbook.Worksheets(modConfig.SHEET_DASHBOARD)
        ws.Buttons.Delete ' Clear existing buttons
    End If

    ' Setup dashboard
    With ws
        .Cells.Clear
        .Cells(1, 1).Value = "Financial Import & Classification System"
        .Cells(1, 1).Font.Size = 18
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Color = RGB(68, 114, 196)

        .Cells(3, 1).Value = "Main Operations"
        .Cells(3, 1).Font.Size = 12
        .Cells(3, 1).Font.Bold = True

        .Cells(12, 1).Value = "Maintenance"
        .Cells(12, 1).Font.Size = 12
        .Cells(12, 1).Font.Bold = True

        .Cells(18, 1).Value = "Reports & Analysis"
        .Cells(18, 1).Font.Size = 12
        .Cells(18, 1).Font.Bold = True
    End With

    ' Button dimensions
    btnWidth = 200
    btnHeight = 30
    btnLeft = 20

    ' Main operation buttons
    btnTop = 80
    Set btn = ws.Buttons.Add(btnLeft, btnTop, btnWidth, btnHeight)
    btn.Caption = "🔧 Initialize System"
    btn.OnAction = "InitializeSystem"

    btnTop = btnTop + btnHeight + 10
    Set btn = ws.Buttons.Add(btnLeft, btnTop, btnWidth, btnHeight)
    btn.Caption = "📥 Import Transactions"
    btn.OnAction = "ImportTransactions"

    btnTop = btnTop + btnHeight + 10
    Set btn = ws.Buttons.Add(btnLeft, btnTop, btnWidth, btnHeight)
    btn.Caption = "🔄 Reprocess Data"
    btn.OnAction = "ReprocessData"

    btnTop = btnTop + btnHeight + 10
    Set btn = ws.Buttons.Add(btnLeft, btnTop, btnWidth, btnHeight)
    btn.Caption = "📋 Update Categories"
    btn.OnAction = "UpdateCategories"

    ' Maintenance buttons
    btnTop = 240
    Set btn = ws.Buttons.Add(btnLeft, btnTop, btnWidth, btnHeight)
    btn.Caption = "✅ Clear Resolved Errors"
    btn.OnAction = "modErrors.ClearResolvedErrors"

    btnTop = btnTop + btnHeight + 10
    Set btn = ws.Buttons.Add(btnLeft, btnTop, btnWidth, btnHeight)
    btn.Caption = "🗑️ Clear All Data"
    btn.OnAction = "ClearAllData"

    ' Reports buttons
    btnTop = 360
    Set btn = ws.Buttons.Add(btnLeft, btnTop, btnWidth, btnHeight)
    btn.Caption = "📊 Show Statistics"
    btn.OnAction = "ShowStatistics"

    btnTop = btnTop + btnHeight + 10
    Set btn = ws.Buttons.Add(btnLeft, btnTop, btnWidth, btnHeight)
    btn.Caption = "💾 Export Transactions"
    btn.OnAction = "ExportTransactions"

    btnTop = btnTop + btnHeight + 10
    Set btn = ws.Buttons.Add(btnLeft, btnTop, btnWidth, btnHeight)
    btn.Caption = "📈 Investment Report"
    btn.OnAction = "modInvestments.GenerateInvestmentReport"

    btnTop = btnTop + btnHeight + 10
    Set btn = ws.Buttons.Add(btnLeft, btnTop, btnWidth, btnHeight)
    btn.Caption = "⚠️ Error Report"
    btn.OnAction = "modErrors.GenerateErrorReport"

    ' Add instructions
    ws.Cells(24, 1).Value = "Quick Start Guide:"
    ws.Cells(24, 1).Font.Bold = True
    ws.Cells(25, 1).Value = "1. Click 'Initialize System' to set up all sheets"
    ws.Cells(26, 1).Value = "2. Go to Config sheet and set your file paths"
    ws.Cells(27, 1).Value = "3. Review Categories sheet and customize as needed"
    ws.Cells(28, 1).Value = "4. Click 'Import Transactions' to process your files"
    ws.Cells(29, 1).Value = "5. Review Errors sheet and resolve any issues"

    ws.Columns("A:A").ColumnWidth = 50
End Sub

' ========================================================================
' Arrange Sheets in Logical Order
' ========================================================================
Private Sub ArrangeSheets()
    On Error Resume Next

    Dim sheetOrder() As String
    Dim i As Integer

    sheetOrder = Split("Dashboard,Config,Categories,Transactions,Errors,Audit,RawData", ",")

    For i = LBound(sheetOrder) To UBound(sheetOrder)
        If modConfig.SheetExists(sheetOrder(i)) Then
            ThisWorkbook.Worksheets(sheetOrder(i)).Move Before:=ThisWorkbook.Worksheets(1)
        End If
    Next i

    ' Hide technical sheets
    If modConfig.SheetExists(modConfig.SHEET_RAWDATA) Then
        ThisWorkbook.Worksheets(modConfig.SHEET_RAWDATA).Visible = xlSheetVeryHidden
    End If
End Sub

' ========================================================================
' Backup Workbook
' ========================================================================
Public Sub BackupWorkbook()
    On Error GoTo ErrorHandler

    Dim backupPath As String
    Dim backupFolder As String

    backupFolder = modConfig.GetConfig("BackupFolder", "")

    If backupFolder = "" Then
        backupFolder = ThisWorkbook.Path
    End If

    backupPath = backupFolder & "/Backup_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsm"

    ThisWorkbook.SaveCopyAs backupPath

    MsgBox "Backup created: " & backupPath, vbInformation

    modUtils.LogMessage "Backup created: " & backupPath, "INFO"

    Exit Sub

ErrorHandler:
    MsgBox "Error creating backup: " & Err.Description, vbCritical
End Sub

' ========================================================================
' Quick Help
' ========================================================================
Public Sub ShowHelp()
    On Error Resume Next

    Dim helpText As String

    helpText = "Financial Import & Classification System" & vbCrLf & _
               "═══════════════════════════════════════════" & vbCrLf & vbCrLf & _
               "MAIN FEATURES:" & vbCrLf & _
               "• Import from multiple file formats (CSV, Excel)" & vbCrLf & _
               "• Automatic transfer detection" & vbCrLf & _
               "• Investment transaction classification" & vbCrLf & _
               "• Smart category matching (exact + fuzzy)" & vbCrLf & _
               "• Error tracking and resolution" & vbCrLf & vbCrLf & _
               "WORKFLOW:" & vbCrLf & _
               "1. Configure → Set file paths in Config sheet" & vbCrLf & _
               "2. Import → Process all configured files" & vbCrLf & _
               "3. Review → Check Transactions and Errors sheets" & vbCrLf & _
               "4. Resolve → Fix any unclassified transactions" & vbCrLf & _
               "5. Export → Generate reports or export data" & vbCrLf & vbCrLf & _
               "For detailed documentation, see README.md"

    MsgBox helpText, vbInformation, "System Help"
End Sub
