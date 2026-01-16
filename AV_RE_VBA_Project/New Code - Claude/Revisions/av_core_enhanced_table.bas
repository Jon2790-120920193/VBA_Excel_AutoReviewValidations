Attribute VB_Name = "AV_Core"
Option Explicit

' ======================================================
' AV_Core
' Core services: configuration, mapping, debug
' NOW FULLY TABLE-BASED - NO CELL REFERENCES
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
    
    ' Check each target sheet and table exists
    Dim i As Long
    For i = 1 To config.TargetCount
        ' Check sheet exists
        If Not AV_DataAccess.WorksheetExists(config.Targets(i).TableName) Then
            errorMessage = AV_Constants.FormatString(AV_Constants.ERR_TARGET_SHEET_MISSING, config.Targets(i).TableName)
            ValidateConfiguration = False
            Exit Function
        End If
        
        ' Check table exists in sheet
        Dim wsTarget As Worksheet
        Set wsTarget = ThisWorkbook.Sheets(config.Targets(i).TableName)
        
        If wsTarget.ListObjects.Count = 0 Then
            errorMessage = AV_Constants.FormatString(AV_Constants.ERR_TARGET_TABLE_MISSING, config.Targets(i).TableName, config.Targets(i).TableName)
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
    If Not gAutoValidationMap Is Nothing Then
        Set GetAutoValidationMap = gAutoValidationMap
        Exit Function
    End If
    
    If wsConfig Is Nothing Then
        Set wsConfig = ThisWorkbook.Sheets(AV_Constants.CONFIG_SHEET_NAME)
    End If
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim tbl As ListObject
    Set tbl = AV_DataAccess.GetTable(wsConfig, AV_Constants.TBL_AUTO_VAL_MAPPING)
    
    If tbl Is Nothing Then
        Set GetAutoValidationMap = dict
        Exit Function
    End If
    
    Dim r As ListRow
    Dim devFunc As String
    Dim item As Object
    
    For Each r In tbl.ListRows
        devFunc = "Validate_Column_" & Trim(CStr(r.Range.Cells(1, tbl.ListColumns(AV_Constants.COL_AVCPM_DEV_FUNC_NAMES).Index).Value))
        
        Set item = CreateObject("Scripting.Dictionary")
        item("DropColHeader") = Trim(CStr(r.Range.Cells(1, tbl.ListColumns(AV_Constants.COL_AVCPM_DROP_COLUMN).Index).Value))
        item("PrefixEN") = Trim(CStr(r.Range.Cells(1, tbl.ListColumns(AV_Constants.COL_AVCPM_PREFIX_EN).Index).Value))
        item("PrefixFR") = Trim(CStr(r.Range.Cells(1, tbl.ListColumns(AV_Constants.COL_AVCPM_PREFIX_FR).Index).Value))
        item("ColumnRef") = Trim(CStr(r.Range.Cells(1, tbl.ListColumns(AV_Constants.COL_AVCPM_REVIEW_COLUMN_HEADER).Index).Value))
        item("AutoValidate") = (LCase(Trim(CStr(r.Range.Cells(1, tbl.ListColumns(AV_Constants.COL_AVCPM_AUTO_VALIDATE).Index).Value))) = "true")
        item("RuleTable") = Trim(CStr(r.Range.Cells(1, tbl.ListColumns(AV_Constants.COL_AVCPM_RULE_TABLE).Index).Value))
        
        If devFunc <> "Validate_Column_" Then
            dict(devFunc) = item
        End If
    Next r
    
    Set gAutoValidationMap = dict
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

    If Not tbl Is Nothing Then
        For Each r In tbl.ListRows
            If LCase(Trim(r.Range(1, 1).Value)) = "global" Then
                GlobalDebugOn = (LCase(Trim(r.Range(1, 2).Value)) = "true")
            End If
        Next r
    End If

    On Error Resume Next
    Set tbl = wsConfig.ListObjects(AV_Constants.TBL_DEBUG_CONTROLS)
    On Error GoTo 0

    If Not tbl Is Nothing Then
        For Each r In tbl.ListRows
            DebugFlags(r.Range(1, 1).Value) = (LCase(Trim(r.Range(1, 2).Value)) = "true")
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

Public Function SafeTrim(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Then
        SafeTrim = ""
    Else
        SafeTrim = Trim(CStr(v))
    End If
End Function
