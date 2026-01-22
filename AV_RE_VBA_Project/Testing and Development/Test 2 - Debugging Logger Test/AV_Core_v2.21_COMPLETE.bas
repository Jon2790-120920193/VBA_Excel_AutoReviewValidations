Attribute VB_Name = "AV_Core"
Option Explicit

' ======================================================
' AV_Core v2.1 - COMPLETE & CORRECTED
' Core services: configuration, mapping, debug, table caching
' ALL FUNCTIONS INCLUDED - Phase 1 + Phase 2 compatible
' ======================================================

Private Const MODULE_NAME As String = "AV_Core"

' ======================================================
' GLOBAL STATE
' ======================================================

Public ValidationStartTime As Single
Public ValidationCancelTimeout As Single
Public ValidationCancelFlag As Boolean

' Debug flags
Public DebugFlags As Object
Public GlobalDebugOn As Boolean
Private DebugInitialized As Boolean

' Auto-validation mapping cache
Private gAutoValidationMap As Object

' Validation table cache
Private Type ValidationTableCache
    GIWValidation As ListObject
    ElectricityPairs As ListObject
    PlumbingPairs As ListObject
    HeatSourcePairs As ListObject
    HeatAnyRef As ListObject
    IsLoaded As Boolean
End Type
Private gTableCache As ValidationTableCache

' ======================================================
' VALIDATION TARGET STRUCTURE
' ======================================================

Public Type ValidationTarget
    TableName As String
    Enabled As Boolean
    Mode As String
    KeyColumnHeader As String
End Type

Public Type ValidationConfig
    Targets() As ValidationTarget
    TargetCount As Long
    Language As String
    IsEnglish As Boolean
    TimeoutSeconds As Long
End Type

' ======================================================
' LOAD VALIDATION CONFIGURATION
' ======================================================

Public Function LoadValidationConfig() As ValidationConfig
    Dim config As ValidationConfig
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(AV_Constants.CONFIG_SHEET_NAME)
    
    ' Load timeout setting
    config.TimeoutSeconds = AV_Constants.VALIDATION_TIMEOUT_SECONDS
    
    ' TODO: Load language from ValidationSettings table when created
    ' For now, default to English
    config.Language = AV_Constants.LANGUAGE_ENGLISH
    config.IsEnglish = True
    
    ' Load validation targets
    Dim tblTargets As ListObject
    Set tblTargets = AV_DataAccess.GetTable(wsConfig, AV_Constants.TBL_VALIDATION_TARGETS)
    
    If tblTargets Is Nothing Then
        config.TargetCount = 0
        LoadValidationConfig = config
        Exit Function
    End If
    
    ' Count enabled targets
    Dim r As ListRow
    Dim enabledCount As Long
    enabledCount = 0
    
    For Each r In tblTargets.ListRows
        If LCase(Trim(CStr(r.Range.Cells(1, tblTargets.ListColumns(AV_Constants.COL_VT_ENABLED).Index).Value))) = "true" Then
            enabledCount = enabledCount + 1
        End If
    Next r
    
    If enabledCount = 0 Then
        config.TargetCount = 0
        LoadValidationConfig = config
        Exit Function
    End If
    
    ' Load targets
    ReDim config.Targets(1 To enabledCount)
    Dim idx As Long
    idx = 0
    
    For Each r In tblTargets.ListRows
        If LCase(Trim(CStr(r.Range.Cells(1, tblTargets.ListColumns(AV_Constants.COL_VT_ENABLED).Index).Value))) = "true" Then
            idx = idx + 1
            
            config.Targets(idx).TableName = Trim(CStr(r.Range.Cells(1, tblTargets.ListColumns(AV_Constants.COL_VT_TABLE_NAME).Index).Value))
            config.Targets(idx).Enabled = True
            config.Targets(idx).Mode = Trim(CStr(r.Range.Cells(1, tblTargets.ListColumns(AV_Constants.COL_VT_MODE).Index).Value))
            config.Targets(idx).KeyColumnHeader = Trim(CStr(r.Range.Cells(1, tblTargets.ListColumns(AV_Constants.COL_VT_KEY_COLUMN).Index).Value))
        End If
    Next r
    
    config.TargetCount = enabledCount
    LoadValidationConfig = config
End Function

' ======================================================
' CONFIGURATION VALIDATION
' ======================================================

Public Function ValidateConfiguration(ByRef errorMessage As String) As Boolean
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(AV_Constants.CONFIG_SHEET_NAME)
    
    ' Check critical tables exist
    Dim criticalTables As Variant
    criticalTables = Array( _
        AV_Constants.TBL_VALIDATION_TARGETS, _
        AV_Constants.TBL_AUTO_VAL_MAPPING, _
        AV_Constants.TBL_AUTO_FORMAT _
    )
    
    Dim tblName As Variant
    For Each tblName In criticalTables
        If Not AV_DataAccess.TableExists(wsConfig, CStr(tblName)) Then
            errorMessage = AV_Constants.FormatString(AV_Constants.ERR_CONFIG_TABLE_MISSING, tblName)
            ValidateConfiguration = False
            Exit Function
        End If
    Next tblName
    
    ' Load and validate targets
    Dim config As ValidationConfig
    config = LoadValidationConfig()
    
    If config.TargetCount = 0 Then
        errorMessage = AV_Constants.ERR_NO_VALIDATION_TARGETS
        ValidateConfiguration = False
        Exit Function
    End If
    
    ' Check each target table exists (by ListObject name)
    Dim i As Long
    For i = 1 To config.TargetCount
        ' Find the ListObject by name
        Dim tblFound As ListObject
        Set tblFound = FindListObjectByName(config.Targets(i).TableName)
        
        If tblFound Is Nothing Then
            errorMessage = "Table not found: " & config.Targets(i).TableName & vbCrLf & _
                          "Check ValidationTargets table - TableName should be the Excel Table (ListObject) name."
            ValidateConfiguration = False
            Exit Function
        End If
    Next i
    
    ValidateConfiguration = True
End Function

' ======================================================
' AUTO-VALIDATION MAP (TABLE-BASED)
' ======================================================

Public Function GetAutoValidationMap(Optional wsConfig As Worksheet) As Object
    On Error GoTo ErrorHandler
    
    ' Check if using cached version
    If Not gAutoValidationMap Is Nothing Then
        AV_UI.AppendUserLog "[GetAutoValidationMap] Using cached version (" & gAutoValidationMap.Count & " items)"
        Set GetAutoValidationMap = gAutoValidationMap
        Exit Function
    End If
    
    AV_UI.AppendUserLog "[GetAutoValidationMap] Loading validation mappings..."
    
    If wsConfig Is Nothing Then
        Set wsConfig = ThisWorkbook.Sheets(AV_Constants.CONFIG_SHEET_NAME)
    End If
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim tbl As ListObject
    Set tbl = AV_DataAccess.GetTable(wsConfig, AV_Constants.TBL_AUTO_VAL_MAPPING)
    
    If tbl Is Nothing Then
        AV_UI.AppendUserLog "[GetAutoValidationMap] ERROR: Table not found"
        Set GetAutoValidationMap = dict
        Exit Function
    End If
    
    DebugMessage "Table found: " & tbl.Name & " (" & tbl.ListRows.Count & " rows)", MODULE_NAME
    
    Dim r As ListRow
    Dim devFunc As String
    Dim item As Object
    Dim rowIndex As Long
    Dim successCount As Long
    Dim skipCount As Long
    
    ' Get column indices once
    Dim colDevFunc As Long, colDrop As Long, colPrefixEN As Long, colPrefixFR As Long
    Dim colReview As Long, colAutoVal As Long, colRuleTable As Long
    
    On Error Resume Next
    colDevFunc = tbl.ListColumns(AV_Constants.COL_AVCPM_DEV_FUNC_NAMES).Index
    colDrop = tbl.ListColumns(AV_Constants.COL_AVCPM_DROP_COLUMN).Index
    colPrefixEN = tbl.ListColumns(AV_Constants.COL_AVCPM_PREFIX_EN).Index
    colPrefixFR = tbl.ListColumns(AV_Constants.COL_AVCPM_PREFIX_FR).Index
    colReview = tbl.ListColumns(AV_Constants.COL_AVCPM_REVIEW_COLUMN_HEADER).Index
    colAutoVal = tbl.ListColumns(AV_Constants.COL_AVCPM_AUTO_VALIDATE).Index
    colRuleTable = tbl.ListColumns(AV_Constants.COL_AVCPM_RULE_TABLE).Index
    
    If Err.Number <> 0 Then
        AV_UI.AppendUserLog "[GetAutoValidationMap] ERROR: Column not found - " & Err.Description
        Err.Clear
        Set GetAutoValidationMap = dict
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    rowIndex = 0
    successCount = 0
    skipCount = 0
    
    For Each r In tbl.ListRows
        rowIndex = rowIndex + 1
        
        On Error Resume Next
        
        ' Read Dev Function Names
        devFunc = "Validate_Column_" & Trim(CStr(tbl.DataBodyRange(rowIndex, colDevFunc).Value))
        
        If Err.Number <> 0 Then
            DebugMessage "Row " & rowIndex & " ERROR reading Dev Function: " & Err.Description, MODULE_NAME
            Err.Clear
            skipCount = skipCount + 1
            GoTo NextRow
        End If
        
        ' Check if empty
        If devFunc = "Validate_Column_" Then
            DebugMessage "Row " & rowIndex & " SKIPPED - empty Dev Function Name", MODULE_NAME
            skipCount = skipCount + 1
            GoTo NextRow
        End If
        
        DebugMessage "Row " & rowIndex & " Processing: " & devFunc, MODULE_NAME
        
        Set item = CreateObject("Scripting.Dictionary")
        
        ' Read all columns
        item("DropColHeader") = Trim(CStr(tbl.DataBodyRange(rowIndex, colDrop).Value))
        item("PrefixEN") = Trim(CStr(tbl.DataBodyRange(rowIndex, colPrefixEN).Value))
        item("PrefixFR") = Trim(CStr(tbl.DataBodyRange(rowIndex, colPrefixFR).Value))
        item("ColumnRef") = Trim(CStr(tbl.DataBodyRange(rowIndex, colReview).Value))
        item("AutoValidate") = (LCase(Trim(CStr(tbl.DataBodyRange(rowIndex, colAutoVal).Value))) = "true")
        item("RuleTable") = Trim(CStr(tbl.DataBodyRange(rowIndex, colRuleTable).Value))
        
        If Err.Number <> 0 Then
            DebugMessage "Row " & rowIndex & " ERROR reading data: " & Err.Description, MODULE_NAME
            Err.Clear
            skipCount = skipCount + 1
            GoTo NextRow
        End If
        
        On Error GoTo ErrorHandler
        
        ' Add to dictionary
        Set dict(devFunc) = item
        successCount = successCount + 1
        
NextRow:
    Next r
    
    AV_UI.AppendUserLog "Loaded " & successCount & " validation mappings"
    DebugMessage "Success: " & successCount & " | Skipped: " & skipCount, MODULE_NAME
    
    Set gAutoValidationMap = dict
    Set GetAutoValidationMap = dict
    Exit Function
    
ErrorHandler:
    AV_UI.AppendUserLog "CRITICAL ERROR in GetAutoValidationMap"
    AV_UI.AppendUserLog "Error #" & Err.Number & ": " & Err.Description
    DebugMessage "CRITICAL ERROR at row " & rowIndex & ": " & Err.Number & " - " & Err.Description, MODULE_NAME
    Set GetAutoValidationMap = dict
End Function

Public Function GetRuleTableNameFromAutoValMap(AutoValMap As Object, ByVal DevFuncName As String, ByVal DefaultRuleTable As String) As String
    If AutoValMap Is Nothing Then
        GetRuleTableNameFromAutoValMap = DefaultRuleTable
        Exit Function
    End If
    
    Dim fullFuncName As String
    fullFuncName = "Validate_Column_" & DevFuncName
    
    If AutoValMap.Exists(fullFuncName) Then
        Dim ruleTable As String
        ruleTable = AutoValMap(fullFuncName)("RuleTable")
        If Len(Trim(ruleTable)) > 0 Then
            GetRuleTableNameFromAutoValMap = ruleTable
        Else
            GetRuleTableNameFromAutoValMap = DefaultRuleTable
        End If
    Else
        GetRuleTableNameFromAutoValMap = DefaultRuleTable
    End If
End Function

' ======================================================
' VALIDATION TABLE CACHE
' ======================================================

Public Sub ClearTableCache()
    gTableCache.IsLoaded = False
    Set gTableCache.GIWValidation = Nothing
    Set gTableCache.ElectricityPairs = Nothing
    Set gTableCache.PlumbingPairs = Nothing
    Set gTableCache.HeatSourcePairs = Nothing
    Set gTableCache.HeatAnyRef = Nothing
End Sub

Public Sub ClearAutoValidationMapCache()
    ' Clear the cached AutoValidationMap
    Set gAutoValidationMap = Nothing
End Sub

Public Function GetValidationTable(tableName As String) As ListObject
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(AV_Constants.CONFIG_SHEET_NAME)
    
    If Not gTableCache.IsLoaded Then
        On Error Resume Next
        Set gTableCache.GIWValidation = wsConfig.ListObjects(AV_Constants.TBL_GIW_VALIDATION)
        Set gTableCache.ElectricityPairs = wsConfig.ListObjects(AV_Constants.TBL_ELECTRICITY_PAIRS)
        Set gTableCache.PlumbingPairs = wsConfig.ListObjects(AV_Constants.TBL_PLUMBING_PAIRS)
        Set gTableCache.HeatSourcePairs = wsConfig.ListObjects(AV_Constants.TBL_HEAT_SOURCE_PAIRS)
        Set gTableCache.HeatAnyRef = wsConfig.ListObjects(AV_Constants.TBL_HEAT_ANY_REF)
        On Error GoTo 0
        gTableCache.IsLoaded = True
    End If
    
    Select Case tableName
        Case AV_Constants.TBL_GIW_VALIDATION
            Set GetValidationTable = gTableCache.GIWValidation
        Case AV_Constants.TBL_ELECTRICITY_PAIRS
            Set GetValidationTable = gTableCache.ElectricityPairs
        Case AV_Constants.TBL_PLUMBING_PAIRS
            Set GetValidationTable = gTableCache.PlumbingPairs
        Case AV_Constants.TBL_HEAT_SOURCE_PAIRS
            Set GetValidationTable = gTableCache.HeatSourcePairs
        Case AV_Constants.TBL_HEAT_ANY_REF
            Set GetValidationTable = gTableCache.HeatAnyRef
        Case Else
            Set GetValidationTable = AV_DataAccess.GetTable(wsConfig, tableName)
    End Select
End Function

' ======================================================
' ROW VALIDATION DECISIONS (LEGACY - NEEDS UPDATE)
' ======================================================

Public Function ShouldValidateRow(ByVal rowNum As Long, wsTarget As Worksheet, Optional ByVal ForceValidation As Boolean = False) As Boolean
    If ForceValidation Then
        ShouldValidateRow = True
        Exit Function
    End If
    
    ' TODO: Update to use table-based ForceValidationTable
    ' For now, default to validate all rows
    ShouldValidateRow = True
End Function

Public Function ValidationTimeoutReached() As Boolean
    If ValidationCancelTimeout <= 0 Then Exit Function
    ValidationTimeoutReached = (Timer - ValidationStartTime) >= ValidationCancelTimeout
End Function

' ======================================================
' COLUMN METADATA (LEGACY - Still needed by AV_Engine)
' ======================================================

Public Function GetValidationColumns(wsConfig As Worksheet) As Object
    ' LEGACY FUNCTION - Reads from cells B6, B7, etc.
    ' TODO: Phase 3 - Replace with table-based approach
    Dim dict As Object
    Dim r As Long

    Set dict = CreateObject("Scripting.Dictionary")
    r = 6

    Do While Trim(wsConfig.Range("B" & r).Value) <> ""
        dict(wsConfig.Range("B" & r).Value) = wsConfig.Range("C" & r).Value
        r = r + 1
    Loop

    Set GetValidationColumns = dict
End Function

Public Function GetDDMValidationColumns(wsConfig As Worksheet) As Object
    ' Gets DDM (dropdown menu) validation columns configuration
    Dim DDMRefTable As ListObject
    Dim r As ListRow
    Dim dict As Object
    Dim ReferenceTable As Object
    Dim ReferenceTableName As String
    Dim StartRowIndex As Long
    Dim EndRowMaxIndex As Long
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Get DDM sheet info
    On Error Resume Next
    Set ReferenceTable = GetDDMSheetInfo(wsConfig)
    On Error GoTo 0
    
    If ReferenceTable Is Nothing Then
        Set GetDDMValidationColumns = dict
        Exit Function
    End If
    
    ReferenceTableName = ReferenceTable("ValidationTableName")
    StartRowIndex = ReferenceTable("StartRowIndex")
    EndRowMaxIndex = ReferenceTable("EndRowIndex")
    
    On Error Resume Next
    Set DDMRefTable = wsConfig.ListObjects(AV_Constants.TBL_AUTO_CHECK_VALIDATION)
    On Error GoTo 0
    
    If DDMRefTable Is Nothing Then
        Set GetDDMValidationColumns = dict
        Exit Function
    End If
    
    Dim i As Long
    i = 0
    For Each r In DDMRefTable.ListRows
        i = i + 1
        Dim autoCheckVal As String
        autoCheckVal = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns(AV_Constants.COL_ACDV_AUTO_CHECK).Index).Value))
        
        If StrComp(autoCheckVal, "TRUE", vbTextCompare) = 0 Then
            Dim key As String
            key = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns(AV_Constants.COL_ACDV_REVIEW_COLUMN_NAME).Index).Value))
            
            Dim item As Object
            Set item = CreateObject("Scripting.Dictionary")
            
            item("ReviewLetter") = key
            item("ColumnNameFR") = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns(AV_Constants.COL_ACDV_COLUMN_NAME_FR).Index).Value))
            item("ColumnNameEN") = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns(AV_Constants.COL_ACDV_COLUMN_NAME).Index).Value))
            item("MenuFieldEN") = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns(AV_Constants.COL_ACDV_MENU_FIELD_EN).Index).Value))
            item("MenuFieldFR") = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns(AV_Constants.COL_ACDV_MENU_FIELD_FR).Index).Value))
            item("CommentDropCol") = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns(AV_Constants.COL_ACDV_AUTO_COMMENT).Index).Value))
            
            Dim NonEmptyRangeEN As Range
            Dim NonEmptyRangeFR As Range
            
            Set NonEmptyRangeEN = GetNonEmptyRangeInColumn(ReferenceTableName, item("MenuFieldEN"), StartRowIndex, EndRowMaxIndex)
            Set NonEmptyRangeFR = GetNonEmptyRangeInColumn(ReferenceTableName, item("MenuFieldFR"), StartRowIndex, EndRowMaxIndex)
            
            Dim listEN As Variant, listFR As Variant
            
            If Not NonEmptyRangeEN Is Nothing Then
                listEN = GetValuesAsList(NonEmptyRangeEN)
                If IsArray(listEN) Then item("ValidColumnListEN") = listEN
            Else
                item("ValidColumnListEN") = Array()
            End If
            
            If Not NonEmptyRangeFR Is Nothing Then
                listFR = GetValuesAsList(NonEmptyRangeFR)
                If IsArray(listFR) Then item("ValidColumnListFR") = listFR
            Else
                item("ValidColumnListFR") = Array()
            End If
            
            dict.Add key, item
        End If
    Next r
    
    Set GetDDMValidationColumns = dict
End Function

Private Function GetDDMSheetInfo(wsConfig As Worksheet) As Object
    Dim tbl As ListObject
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    Set tbl = wsConfig.ListObjects(AV_Constants.TBL_DDM_FIELDS_INFO)
    On Error GoTo 0
    
    If tbl Is Nothing Then Exit Function
    
    dict("ValidationTableName") = CStr(tbl.DataBodyRange.Cells(1, 2).Value)
    dict("StartRowIndex") = CLng(tbl.DataBodyRange.Cells(2, 2).Value)
    dict("EndRowIndex") = CLng(tbl.DataBodyRange.Cells(3, 2).Value)
    
    Set GetDDMSheetInfo = dict
End Function

Private Function GetNonEmptyRangeInColumn(sheetName As String, colLetter As String, startRow As Long, endRow As Long) As Range
    Dim ws As Worksheet
    Dim checkRange As Range
    Dim lastNonEmptyRow As Long
    Dim cell As Range
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Function
    If startRow <= 0 Or endRow < startRow Then Exit Function
    
    Set checkRange = ws.Range(colLetter & startRow & ":" & colLetter & endRow)
    lastNonEmptyRow = 0
    
    For Each cell In checkRange.Cells
        If Trim(CStr(cell.Value)) <> "" Then lastNonEmptyRow = cell.Row
    Next cell
    
    If lastNonEmptyRow = 0 Then Exit Function
    
    Set GetNonEmptyRangeInColumn = ws.Range(colLetter & startRow & ":" & colLetter & lastNonEmptyRow)
End Function

Private Function GetValuesAsList(rng As Range) As Variant
    Dim cell As Range
    Dim valuesList() As String
    Dim Count As Long
    
    If rng Is Nothing Then Exit Function
    
    For Each cell In rng.Cells
        If Trim(CStr(cell.Value)) <> "" Then
            Count = Count + 1
            ReDim Preserve valuesList(1 To Count)
            valuesList(Count) = Trim(CStr(cell.Value))
        End If
    Next cell
    
    If Count > 0 Then
        GetValuesAsList = valuesList
    Else
        GetValuesAsList = Array()
    End If
End Function

' ======================================================
' DEBUG SYSTEM
' ======================================================

Public Sub InitDebugFlags(Optional ByVal ForceReload As Boolean = False)
    Dim wsConfig As Worksheet
    Dim tbl As ListObject
    Dim r As ListRow

    If DebugInitialized And Not ForceReload Then Exit Sub

    Set DebugFlags = CreateObject("Scripting.Dictionary")
    GlobalDebugOn = False

    Set wsConfig = ThisWorkbook.Sheets(AV_Constants.CONFIG_SHEET_NAME)
    
    On Error Resume Next
    Set tbl = wsConfig.ListObjects(AV_Constants.TBL_GLOBAL_DEBUG)
    On Error GoTo 0

    ' GlobalDebugOptions has single column with value ON/OFF
    If Not tbl Is Nothing Then
        If tbl.ListRows.Count > 0 Then
            GlobalDebugOn = (UCase(Trim(tbl.DataBodyRange(1, 1).Value)) = "ON")
        End If
    End If

    On Error Resume Next
    Set tbl = wsConfig.ListObjects(AV_Constants.TBL_DEBUG_CONTROLS)
    On Error GoTo 0

    If Not tbl Is Nothing Then
        For Each r In tbl.ListRows
            DebugFlags(r.Range(1, 1).Value) = (UCase(Trim(r.Range(1, 2).Value)) = "ON")
        Next r
    End If

    DebugInitialized = True
End Sub

Public Sub DebugMessage(ByVal Msg As String, Optional ByVal ModuleName As String = "")
    If Not DebugInitialized Then InitDebugFlags

    If GlobalDebugOn Then
        Debug.Print "[DEBUG] " & ModuleName & " :: " & Msg
    ElseIf ModuleName <> "" Then
        If DebugFlags.Exists(ModuleName) Then
            If DebugFlags(ModuleName) Then
                Debug.Print "[DEBUG] " & ModuleName & " :: " & Msg
            End If
        End If
    End If
End Sub

' ======================================================
' SAFE HELPERS
' ======================================================

Public Function FindListObjectByName(tableName As String) As ListObject
    ' Search all worksheets for a ListObject with the given name
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set tbl = ws.ListObjects(tableName)
        On Error GoTo 0
        
        If Not tbl Is Nothing Then
            Set FindListObjectByName = tbl
            Exit Function
        End If
    Next ws
    
    Set FindListObjectByName = Nothing
End Function

Public Function SafeTrim(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Then
        SafeTrim = ""
    Else
        SafeTrim = Trim(CStr(v))
    End If
End Function
