Attribute VB_Name = "AV_Core"
Option Explicit

' ======================================================
' AV_Core v2.6
' Core services: configuration, mapping, caching, debug
' FIXED: Header-based lookups instead of column letters
' KEY CHANGES:
'   - GetAutoValidationMap reads "ReviewSheet Column Header" not "ReviewSheet Column Letter"
'   - GetDDMValidationColumns reads "ReviewSheet Column Name" not "ReviewSheet Column Letter"
'   - GetValidMenuValues uses header-based table lookup
'   - Added GetCellByHeader helper function
' ======================================================

Private Const MODULE_NAME As String = "AV_Core"

' ======================================================
' GLOBAL STATE
' ======================================================

Public ValidationStartTime As Single
Public ValidationCancelTimeout As Single
Public ValidationCancelFlag As Boolean
Public BulkValidationInProgress As Boolean

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
    Set wsConfig = ThisWorkbook.Sheets("Config")
    
    config.TimeoutSeconds = 10000
    config.Language = "English"
    config.IsEnglish = True
    
    Dim tblTargets As ListObject
    On Error Resume Next
    Set tblTargets = wsConfig.ListObjects("ValidationTargets")
    On Error GoTo 0
    
    If tblTargets Is Nothing Then
        config.TargetCount = 0
        LoadValidationConfig = config
        Exit Function
    End If
    
    Dim r As ListRow
    Dim enabledCount As Long
    enabledCount = 0
    
    For Each r In tblTargets.ListRows
        If LCase(Trim(CStr(r.Range.Cells(1, tblTargets.ListColumns("Enabled").Index).Value))) = "true" Then
            enabledCount = enabledCount + 1
        End If
    Next r
    
    If enabledCount = 0 Then
        config.TargetCount = 0
        LoadValidationConfig = config
        Exit Function
    End If
    
    ReDim config.Targets(1 To enabledCount)
    Dim idx As Long
    idx = 0
    
    For Each r In tblTargets.ListRows
        If LCase(Trim(CStr(r.Range.Cells(1, tblTargets.ListColumns("Enabled").Index).Value))) = "true" Then
            idx = idx + 1
            config.Targets(idx).TableName = Trim(CStr(r.Range.Cells(1, tblTargets.ListColumns("TableName").Index).Value))
            config.Targets(idx).Enabled = True
            config.Targets(idx).Mode = Trim(CStr(r.Range.Cells(1, tblTargets.ListColumns("Mode").Index).Value))
            config.Targets(idx).KeyColumnHeader = Trim(CStr(r.Range.Cells(1, tblTargets.ListColumns("Key Column (Header Name)").Index).Value))
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
    Set wsConfig = ThisWorkbook.Sheets("Config")
    
    Dim criticalTables As Variant
    criticalTables = Array("ValidationTargets", "AutoValidationCommentPrefixMappingTable", "AutoFormatOnFullValidation")
    
    Dim tblName As Variant
    Dim tbl As ListObject
    For Each tblName In criticalTables
        Set tbl = Nothing
        On Error Resume Next
        Set tbl = wsConfig.ListObjects(CStr(tblName))
        On Error GoTo 0
        
        If tbl Is Nothing Then
            errorMessage = "Critical configuration table missing: " & tblName
            ValidateConfiguration = False
            Exit Function
        End If
    Next tblName
    
    Dim config As ValidationConfig
    config = LoadValidationConfig()
    
    If config.TargetCount = 0 Then
        errorMessage = "No validation targets enabled."
        ValidateConfiguration = False
        Exit Function
    End If
    
    ValidateConfiguration = True
End Function

' ======================================================
' AUTO-VALIDATION MAP
' KEY FIX: Reads "ReviewSheet Column Header" not "ReviewSheet Column Letter"
' ======================================================

Public Function GetAutoValidationMap(Optional wsConfig As Worksheet) As Object
    If Not gAutoValidationMap Is Nothing Then
        Set GetAutoValidationMap = gAutoValidationMap
        Exit Function
    End If
    
    If wsConfig Is Nothing Then
        Set wsConfig = ThisWorkbook.Sheets("Config")
    End If
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = wsConfig.ListObjects("AutoValidationCommentPrefixMappingTable")
    On Error GoTo 0
    
    If tbl Is Nothing Then
        Set GetAutoValidationMap = dict
        Exit Function
    End If
    
    Dim r As ListRow
    Dim devFunc As String
    Dim item As Object
    
    For Each r In tbl.ListRows
        devFunc = "Validate_Column_" & SafeTrim(r.Range.Cells(1, tbl.ListColumns("Dev Function Names").Index).Value)
        
        Set item = CreateObject("Scripting.Dictionary")
        item("DropColHeader") = SafeTrim(r.Range.Cells(1, tbl.ListColumns("Drop in Column").Index).Value)
        item("PrefixEN") = SafeTrim(r.Range.Cells(1, tbl.ListColumns("Prefix to message").Index).Value)
        item("PrefixFR") = SafeTrim(r.Range.Cells(1, tbl.ListColumns("(FR) Prefix to message").Index).Value)
        
        ' *** KEY FIX: Now reads "ReviewSheet Column Header" (header name, not column letter) ***
        item("ColumnRef") = SafeTrim(r.Range.Cells(1, tbl.ListColumns("ReviewSheet Column Header").Index).Value)
        
        item("AutoValidate") = (LCase(SafeTrim(r.Range.Cells(1, tbl.ListColumns("AutoValidate").Index).Value)) = "true")
        item("RuleTable") = SafeTrim(r.Range.Cells(1, tbl.ListColumns("RuleTableName").Index).Value)
        
        If devFunc <> "Validate_Column_" Then
            Set dict(devFunc) = item
        End If
    Next r
    
    Set gAutoValidationMap = dict
    Set GetAutoValidationMap = dict
End Function

Public Sub ClearAutoValidationMapCache()
    Set gAutoValidationMap = Nothing
End Sub

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
    Set wsConfig = ThisWorkbook.Sheets("Config")
    
    If Not gTableCache.IsLoaded Then
        On Error Resume Next
        Set gTableCache.GIWValidation = wsConfig.ListObjects("GIWValidationTable")
        Set gTableCache.ElectricityPairs = wsConfig.ListObjects("ElectricityPairValidation")
        Set gTableCache.PlumbingPairs = wsConfig.ListObjects("PlumbingPairValidation")
        Set gTableCache.HeatSourcePairs = wsConfig.ListObjects("HeatSourcePairValidation")
        Set gTableCache.HeatAnyRef = wsConfig.ListObjects("HeatSourceANYRefTable")
        On Error GoTo 0
        gTableCache.IsLoaded = True
    End If
    
    Select Case tableName
        Case "GIWValidationTable"
            Set GetValidationTable = gTableCache.GIWValidation
        Case "ElectricityPairValidation"
            Set GetValidationTable = gTableCache.ElectricityPairs
        Case "PlumbingPairValidation"
            Set GetValidationTable = gTableCache.PlumbingPairs
        Case "HeatSourcePairValidation"
            Set GetValidationTable = gTableCache.HeatSourcePairs
        Case "HeatSourceANYRefTable"
            Set GetValidationTable = gTableCache.HeatAnyRef
        Case Else
            On Error Resume Next
            Set GetValidationTable = wsConfig.ListObjects(tableName)
            On Error GoTo 0
    End Select
End Function

' ======================================================
' ROW VALIDATION DECISIONS
' ======================================================

Public Function ShouldValidateRow(ByVal rowNum As Long, wsTarget As Worksheet, _
    Optional TargetTable As ListObject = Nothing, _
    Optional ByVal ForceValidation As Boolean = False) As Boolean
    
    If ForceValidation Then
        ShouldValidateRow = True
        Exit Function
    End If
    
    Dim tbl As ListObject
    Dim wsConfig As Worksheet
    Dim colHeaderToCheck As String
    Dim buildingValue As String
    Dim checkValue As String
    Dim i As Long

    Set wsConfig = ThisWorkbook.Sheets("Config")
    
    On Error Resume Next
    Set tbl = wsConfig.ListObjects("ForceValidationTable")
    On Error GoTo 0

    If tbl Is Nothing Then
        ShouldValidateRow = False
        Exit Function
    End If
    
    If TargetTable Is Nothing Then
        ShouldValidateRow = False
        Exit Function
    End If

    For i = 1 To tbl.ListRows.Count
        colHeaderToCheck = Trim(tbl.DataBodyRange(i, tbl.ListColumns("Column").Index).Value)
        buildingValue = Trim(tbl.DataBodyRange(i, tbl.ListColumns("IsBuildingColumnValue").Index).Value)
        
        If colHeaderToCheck <> "" Then
            Dim targetColIndex As Long
            targetColIndex = 0
            On Error Resume Next
            targetColIndex = TargetTable.ListColumns(colHeaderToCheck).Index
            On Error GoTo 0

            If targetColIndex > 0 Then
                Dim tableDataRow As Long
                tableDataRow = rowNum - TargetTable.DataBodyRange.Row + 1
                
                If tableDataRow >= 1 And tableDataRow <= TargetTable.ListRows.Count Then
                    checkValue = Trim(CStr(TargetTable.DataBodyRange(tableDataRow, targetColIndex).Value))
                    
                    If buildingValue = "" And checkValue = "" Then
                        ShouldValidateRow = True
                        Exit Function
                    End If

                    If buildingValue <> "" And StrComp(buildingValue, checkValue, vbTextCompare) = 0 Then
                        ShouldValidateRow = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i

    ShouldValidateRow = False
End Function

Public Function ValidationTimeoutReached() As Boolean
    If ValidationCancelTimeout <= 0 Then Exit Function
    ValidationTimeoutReached = (Timer - ValidationStartTime) >= ValidationCancelTimeout
End Function

' ======================================================
' DDM VALIDATION COLUMNS (SIMPLE VALIDATIONS)
' KEY FIX: Now uses "ReviewSheet Column Name" (header name, not column letter)
' ======================================================

Public Function GetDDMValidationColumns(wsConfig As Worksheet) As Object
    Dim DDMRefTable As ListObject
    Dim r As ListRow
    Dim dict As Object
    Dim ReferenceTable As Object
    Dim ReferenceTableName As String
    Dim StartRowIndex As Long
    Dim EndRowMaxIndex As Long
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    Set ReferenceTable = GetDDMSheetInfo(wsConfig)
    On Error GoTo 0
    
    If ReferenceTable Is Nothing Then
        DebugMessage "GetDDMValidationColumns: DDMFieldsInfo not found", MODULE_NAME
        Set GetDDMValidationColumns = dict
        Exit Function
    End If
    
    ReferenceTableName = ReferenceTable("ValidationTableName")
    StartRowIndex = ReferenceTable("StartRowIndex")
    EndRowMaxIndex = ReferenceTable("EndRowIndex")
    
    DebugMessage "GetDDMValidationColumns: MenuFields sheet = " & ReferenceTableName, MODULE_NAME
    
    On Error Resume Next
    Set DDMRefTable = wsConfig.ListObjects("AutoCheckDataValidationTable")
    On Error GoTo 0
    
    If DDMRefTable Is Nothing Then
        DebugMessage "GetDDMValidationColumns: AutoCheckDataValidationTable not found", MODULE_NAME
        Set GetDDMValidationColumns = dict
        Exit Function
    End If
    
    Dim i As Long
    i = 0
    For Each r In DDMRefTable.ListRows
        i = i + 1
        Dim autoCheckVal As String
        autoCheckVal = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns("AutoCheck").Index).Value))
        
        If StrComp(autoCheckVal, "TRUE", vbTextCompare) = 0 Then
            ' *** KEY FIX: Use "ReviewSheet Column Name" (header name, not column letter) ***
            Dim targetHeaderName As String
            targetHeaderName = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns("ReviewSheet Column Name").Index).Value))
            
            If Len(targetHeaderName) = 0 Then
                DebugMessage "GetDDMValidationColumns: Empty ReviewSheet Column Name at row " & i, MODULE_NAME
                GoTo SkipRow
            End If
            
            Dim item As Object
            Set item = CreateObject("Scripting.Dictionary")
            
            ' Store header name as the key identifier
            item("TargetHeaderName") = targetHeaderName
            item("ColumnNameFR") = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns("Column Name (FR)").Index).Value))
            item("ColumnNameEN") = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns("Column Name").Index).Value))
            item("MenuFieldEN") = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns("MenuField Column (EN)").Index).Value))
            item("MenuFieldFR") = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns("MenuField Column (FR)").Index).Value))
            item("CommentDropCol") = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns("AutoComment Column").Index).Value))
            
            ' Get valid values from menu fields table using header-based lookup
            Dim listEN As Variant, listFR As Variant
            
            listEN = GetValidMenuValues(ReferenceTableName, item("MenuFieldEN"))
            listFR = GetValidMenuValues(ReferenceTableName, item("MenuFieldFR"))
            
            If IsArray(listEN) Then
                item("ValidColumnListEN") = listEN
                DebugMessage "GetDDMValidationColumns: " & targetHeaderName & " EN has " & (UBound(listEN) - LBound(listEN) + 1) & " values", MODULE_NAME
            Else
                item("ValidColumnListEN") = Array()
                DebugMessage "GetDDMValidationColumns: " & targetHeaderName & " EN list EMPTY", MODULE_NAME
            End If
            
            If IsArray(listFR) Then
                item("ValidColumnListFR") = listFR
            Else
                item("ValidColumnListFR") = Array()
            End If
            
            dict.Add targetHeaderName, item
            DebugMessage "GetDDMValidationColumns: Mapped " & targetHeaderName, MODULE_NAME
        End If
SkipRow:
    Next r
    
    DebugMessage "GetDDMValidationColumns: Total mapped = " & dict.Count, MODULE_NAME
    Set GetDDMValidationColumns = dict
End Function
' ======================================================
' GET VALID MENU VALUES (Helper)
' KEY FIX: Now reads by header name from the menu table
' ======================================================

Private Function GetValidMenuValues(sheetName As String, headerName As String) As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim colIndex As Long
    Dim values() As String
    Dim Count As Long
    Dim i As Long
    Dim cellVal As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        DebugMessage "GetValidMenuValues: Sheet not found: " & sheetName, MODULE_NAME
        GetValidMenuValues = Array()
        Exit Function
    End If
    
    If ws.ListObjects.Count = 0 Then
        DebugMessage "GetValidMenuValues: No table on sheet: " & sheetName, MODULE_NAME
        GetValidMenuValues = Array()
        Exit Function
    End If
    
    Set tbl = ws.ListObjects(1)
    
    On Error Resume Next
    colIndex = tbl.ListColumns(headerName).Index
    On Error GoTo 0
    
    If colIndex = 0 Then
        DebugMessage "GetValidMenuValues: Column not found: " & headerName & " in " & sheetName, MODULE_NAME
        GetValidMenuValues = Array()
        Exit Function
    End If
    
    Count = 0
    For i = 1 To tbl.ListRows.Count
        cellVal = Trim(CStr(tbl.DataBodyRange(i, colIndex).Value))
        If Len(cellVal) > 0 Then
            Count = Count + 1
            ReDim Preserve values(1 To Count)
            values(Count) = cellVal
        End If
    Next i
    
    If Count > 0 Then
        GetValidMenuValues = values
    Else
        GetValidMenuValues = Array()
    End If
End Function

Private Function GetDDMSheetInfo(wsConfig As Worksheet) As Object
    Dim tbl As ListObject
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    Set tbl = wsConfig.ListObjects("DDMFieldsInfo")
    On Error GoTo 0
    
    If tbl Is Nothing Then Exit Function
    
    dict("ValidationTableName") = CStr(tbl.DataBodyRange.Cells(1, 2).Value)
    dict("StartRowIndex") = CLng(tbl.DataBodyRange.Cells(2, 2).Value)
    dict("EndRowIndex") = CLng(tbl.DataBodyRange.Cells(3, 2).Value)
    
    Set GetDDMSheetInfo = dict
End Function

' ======================================================
' CELL LOOKUP BY HEADER (NEW HELPER)
' ======================================================

Public Function GetCellByHeader(TargetTable As ListObject, rowNum As Long, headerName As String) As Range
    If TargetTable Is Nothing Then Exit Function
    If Len(headerName) = 0 Then Exit Function
    
    Dim colIndex As Long
    On Error Resume Next
    colIndex = TargetTable.ListColumns(headerName).Index
    On Error GoTo 0
    
    If colIndex = 0 Then
        DebugMessage "GetCellByHeader: Column not found: " & headerName, MODULE_NAME
        Exit Function
    End If
    
    Dim tableRow As Long
    tableRow = rowNum - TargetTable.DataBodyRange.Row + 1
    
    If tableRow < 1 Or tableRow > TargetTable.ListRows.Count Then
        Exit Function
    End If
    
    Set GetCellByHeader = TargetTable.DataBodyRange(tableRow, colIndex)
End Function

' ======================================================
' DEBUG SYSTEM
' ======================================================

Public Sub InitDebugFlags(Optional ByVal ForceReload As Boolean = False)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim r As ListRow

    If DebugInitialized And Not ForceReload Then Exit Sub

    Set DebugFlags = CreateObject("Scripting.Dictionary")
    GlobalDebugOn = False

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Config")
    Set tbl = ws.ListObjects("GlobalDebugOptions")
    On Error GoTo 0

    If Not tbl Is Nothing Then
        For Each r In tbl.ListRows
            GlobalDebugOn = (LCase(Trim(CStr(r.Range(1, 1).Value))) = "true")
        Next r
    End If

    On Error Resume Next
    Set tbl = ws.ListObjects("DebugControls")
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
