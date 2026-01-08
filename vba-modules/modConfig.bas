Attribute VB_Name = "modConfig"
' ========================================================================
' Module: modConfig
' Purpose: Configuration Management
' Description: Handles all system configuration including file paths,
'              settings, and preferences
' Platform: Excel for MacOS
' Author: Financial Import System v1.0
' ========================================================================

Option Explicit

' Configuration Sheet Constants
Private Const CONFIG_SHEET As String = "Config"
Private Const CONFIG_COL_KEY As Integer = 1
Private Const CONFIG_COL_VALUE As Integer = 2
Private Const CONFIG_COL_DESC As Integer = 3
Private Const CONFIG_COL_CATEGORY As Integer = 4

' Default Configuration Values
Private Const DEFAULT_DATE_FORMAT As String = "dd/mm/yyyy"
Private Const DEFAULT_CURRENCY As String = "USD"
Private Const DEFAULT_FUZZY_THRESHOLD As Integer = 80
Private Const DEFAULT_AUTO_BACKUP As String = "TRUE"

' Sheet Name Constants
Public Const SHEET_DASHBOARD As String = "Dashboard"
Public Const SHEET_CONFIG As String = "Config"
Public Const SHEET_CATEGORIES As String = "Categories"
Public Const SHEET_TRANSACTIONS As String = "Transactions"
Public Const SHEET_ERRORS As String = "Errors"
Public Const SHEET_RAWDATA As String = "RawData"
Public Const SHEET_AUDIT As String = "Audit"

' ========================================================================
' Initialize Configuration System
' ========================================================================
Public Sub InitializeConfig()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet

    ' Check if Config sheet exists
    If Not SheetExists(CONFIG_SHEET) Then
        ' Create Config sheet
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = CONFIG_SHEET

        ' Setup headers
        With ws
            .Cells(1, CONFIG_COL_KEY).Value = "ConfigKey"
            .Cells(1, CONFIG_COL_VALUE).Value = "ConfigValue"
            .Cells(1, CONFIG_COL_DESC).Value = "Description"
            .Cells(1, CONFIG_COL_CATEGORY).Value = "Category"

            ' Format headers
            .Range(.Cells(1, 1), .Cells(1, 4)).Font.Bold = True
            .Range(.Cells(1, 1), .Cells(1, 4)).Interior.Color = RGB(68, 114, 196)
            .Range(.Cells(1, 1), .Cells(1, 4)).Font.Color = RGB(255, 255, 255)
        End With

        ' Add default configurations
        Call AddDefaultConfigs
    End If

    ' Validate configuration
    Call ValidateConfig

    Exit Sub

ErrorHandler:
    MsgBox "Error initializing configuration: " & Err.Description, vbCritical, "Config Error"
End Sub

' ========================================================================
' Add Default Configuration Values
' ========================================================================
Private Sub AddDefaultConfigs()
    On Error Resume Next

    ' File Path Configurations
    Call SetConfig("ImportPath1", "", "Primary import file path", "Paths")
    Call SetConfig("ImportPath2", "", "Secondary import file path", "Paths")
    Call SetConfig("ImportPath3", "", "Tertiary import file path", "Paths")
    Call SetConfig("BackupFolder", "", "Backup file location", "Paths")

    ' Format Configurations
    Call SetConfig("DateFormat", DEFAULT_DATE_FORMAT, "Date format for parsing", "Format")
    Call SetConfig("DefaultCurrency", DEFAULT_CURRENCY, "Default currency", "Format")
    Call SetConfig("DecimalSeparator", ".", "Decimal separator", "Format")
    Call SetConfig("ThousandSeparator", ",", "Thousand separator", "Format")

    ' Processing Configurations
    Call SetConfig("FuzzyThreshold", CStr(DEFAULT_FUZZY_THRESHOLD), "Fuzzy match threshold (0-100)", "Processing")
    Call SetConfig("TransferTolerance", "0.001", "Transfer amount tolerance (decimal)", "Processing")
    Call SetConfig("TransferDateDays", "1", "Transfer date tolerance (days)", "Processing")
    Call SetConfig("SkipDuplicates", "TRUE", "Skip duplicate transactions", "Processing")

    ' System Configurations
    Call SetConfig("AutoBackup", DEFAULT_AUTO_BACKUP, "Enable automatic backups", "System")
    Call SetConfig("LogLevel", "INFO", "Logging level (DEBUG/INFO/ERROR)", "System")
    Call SetConfig("MaxImportRows", "100000", "Maximum rows to import", "System")

    ' Auto-fit columns
    ThisWorkbook.Worksheets(CONFIG_SHEET).Columns("A:D").AutoFit
End Sub

' ========================================================================
' Get Configuration Value
' ========================================================================
Public Function GetConfig(configKey As String, Optional defaultValue As String = "") As String
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, CONFIG_COL_KEY).End(xlUp).Row

    ' Search for config key
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, CONFIG_COL_KEY).Value)) = UCase(Trim(configKey)) Then
            GetConfig = ws.Cells(i, CONFIG_COL_VALUE).Value
            Exit Function
        End If
    Next i

    ' Not found, return default
    GetConfig = defaultValue
    Exit Function

ErrorHandler:
    GetConfig = defaultValue
End Function

' ========================================================================
' Set Configuration Value
' ========================================================================
Public Sub SetConfig(configKey As String, configValue As String, _
                     Optional description As String = "", _
                     Optional category As String = "General")
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim found As Boolean

    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, CONFIG_COL_KEY).End(xlUp).Row

    ' Search for existing key
    found = False
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, CONFIG_COL_KEY).Value)) = UCase(Trim(configKey)) Then
            ' Update existing
            ws.Cells(i, CONFIG_COL_VALUE).Value = configValue
            If description <> "" Then
                ws.Cells(i, CONFIG_COL_DESC).Value = description
            End If
            If category <> "" Then
                ws.Cells(i, CONFIG_COL_CATEGORY).Value = category
            End If
            found = True
            Exit For
        End If
    Next i

    ' Add new if not found
    If Not found Then
        lastRow = lastRow + 1
        ws.Cells(lastRow, CONFIG_COL_KEY).Value = configKey
        ws.Cells(lastRow, CONFIG_COL_VALUE).Value = configValue
        ws.Cells(lastRow, CONFIG_COL_DESC).Value = description
        ws.Cells(lastRow, CONFIG_COL_CATEGORY).Value = category
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error setting configuration: " & Err.Description, vbCritical, "Config Error"
End Sub

' ========================================================================
' Get All Import Paths
' ========================================================================
Public Function GetImportPaths() As Collection
    On Error GoTo ErrorHandler

    Dim paths As Collection
    Dim i As Integer
    Dim path As String

    Set paths = New Collection

    ' Get all ImportPath configurations
    For i = 1 To 10 ' Support up to 10 paths
        path = GetConfig("ImportPath" & i, "")
        If path <> "" Then
            ' Validate path exists (MacOS compatible)
            If PathExists(path) Then
                paths.Add path
            End If
        End If
    Next i

    Set GetImportPaths = paths
    Exit Function

ErrorHandler:
    Set GetImportPaths = New Collection
End Function

' ========================================================================
' Check if Path Exists (MacOS Compatible)
' ========================================================================
Private Function PathExists(path As String) As Boolean
    On Error Resume Next

    Dim result As String

    ' Check if path is a file or directory
    result = Dir(path)

    If result <> "" Then
        PathExists = True
    Else
        ' Try as directory
        result = Dir(path & "/*")
        PathExists = (result <> "")
    End If
End Function

' ========================================================================
' Validate Configuration
' ========================================================================
Public Function ValidateConfig() As Boolean
    On Error GoTo ErrorHandler

    Dim issues As String
    Dim paths As Collection

    issues = ""

    ' Check critical configurations
    If GetConfig("FuzzyThreshold") = "" Then
        issues = issues & "- FuzzyThreshold not configured" & vbCrLf
    End If

    If GetConfig("DateFormat") = "" Then
        issues = issues & "- DateFormat not configured" & vbCrLf
    End If

    ' Check if at least one import path is configured
    Set paths = GetImportPaths()
    If paths.Count = 0 Then
        issues = issues & "- No valid import paths configured" & vbCrLf
    End If

    ' Display issues if any
    If issues <> "" Then
        MsgBox "Configuration issues found:" & vbCrLf & vbCrLf & issues, _
               vbExclamation, "Configuration Validation"
        ValidateConfig = False
    Else
        ValidateConfig = True
    End If

    Exit Function

ErrorHandler:
    ValidateConfig = False
End Function

' ========================================================================
' Check if Sheet Exists
' ========================================================================
Public Function SheetExists(sheetName As String) As Boolean
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)

    SheetExists = Not ws Is Nothing
End Function

' ========================================================================
' Get Configuration as Integer
' ========================================================================
Public Function GetConfigInt(configKey As String, Optional defaultValue As Integer = 0) As Integer
    On Error Resume Next

    Dim value As String
    value = GetConfig(configKey, CStr(defaultValue))

    GetConfigInt = CInt(value)

    If Err.Number <> 0 Then
        GetConfigInt = defaultValue
    End If
End Function

' ========================================================================
' Get Configuration as Boolean
' ========================================================================
Public Function GetConfigBool(configKey As String, Optional defaultValue As Boolean = False) As Boolean
    On Error Resume Next

    Dim value As String
    value = UCase(GetConfig(configKey, IIf(defaultValue, "TRUE", "FALSE")))

    GetConfigBool = (value = "TRUE" Or value = "YES" Or value = "1")
End Function

' ========================================================================
' Get Configuration as Double
' ========================================================================
Public Function GetConfigDouble(configKey As String, Optional defaultValue As Double = 0#) As Double
    On Error Resume Next

    Dim value As String
    value = GetConfig(configKey, CStr(defaultValue))

    GetConfigDouble = CDbl(value)

    If Err.Number <> 0 Then
        GetConfigDouble = defaultValue
    End If
End Function

' ========================================================================
' Export Configuration to Dictionary
' ========================================================================
Public Function GetAllConfigs() As Object
    On Error GoTo ErrorHandler

    Dim dict As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set dict = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)

    lastRow = ws.Cells(ws.Rows.Count, CONFIG_COL_KEY).End(xlUp).Row

    For i = 2 To lastRow
        If ws.Cells(i, CONFIG_COL_KEY).Value <> "" Then
            dict(ws.Cells(i, CONFIG_COL_KEY).Value) = ws.Cells(i, CONFIG_COL_VALUE).Value
        End If
    Next i

    Set GetAllConfigs = dict
    Exit Function

ErrorHandler:
    Set GetAllConfigs = CreateObject("Scripting.Dictionary")
End Function
