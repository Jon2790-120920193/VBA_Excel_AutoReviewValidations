Attribute VB_Name = "AV_Core"
Option Explicit

' ======================================================
' AV_Core
' Core services: mapping, routing, row logic, debug flags
' VERSION: 2.5 - Added BulkValidationInProgress flag
' ======================================================

Private Const MODULE_NAME As String = "AV_Core"

' ======================================================
' GLOBAL STATE (DECLARED AT TOP — DO NOT MOVE)
' ======================================================

Public ValidationStartTime As Single
Public ValidationCancelTimeout As Single
Public ValidationCancelFlag As Boolean

' NEW in v2.5: Flag to prevent Worksheet_Change cascade during bulk validation
Public BulkValidationInProgress As Boolean

' Debug flags
Public DebugFlags As Object          ' Scripting.Dictionary
Public GlobalDebugOn As Boolean
Private DebugInitialized As Boolean

' Auto-validation mapping cache
Private gAutoValidationMap As Object


' ======================================================
' CONSTANTS
' ======================================================

' Mapping table
Private Const MAPPING_TABLE_NAME As String = "AutoValidationCommentPrefixMappingTable"

' System tags
Public Const SYSTEM_TAG_START As String = "[[SYS_TAG"
Public Const SYSTEM_TAG_END As String = "]]"
Public Const SYSTEM_COMMENT_TAG As String = "[[SYS_COMMENT]]"

' Formatting
Public Const FALLBACKFORMAT As String = "Default"


' ======================================================
' DEBUG INITIALIZATION
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
            If LCase(Trim(r.Range(1, 1).Value)) = "global" Then
                GlobalDebugOn = (LCase(Trim(r.Range(1, 2).Value)) = "true")
            End If
        Next r
    End If

    On Error Resume Next
    Set tbl = ws.ListObjects("DebugControls")
    On Error GoTo 0

    If Not tbl Is Nothing Then
        For Each r In tbl.ListRows
            DebugFlags(r.Range(1, 1).Value) = _
                (LCase(Trim(r.Range(1, 2).Value)) = "true")
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
' AUTO-VALIDATION MAP
' ======================================================

Public Function GetAutoValidationMap(Optional wsConfig As Worksheet) As Object
    Dim tbl As ListObject
    Dim r As ListRow
    Dim dict As Object

    ' Use cached version if available
    If Not gAutoValidationMap Is Nothing Then
        Set GetAutoValidationMap = gAutoValidationMap
        Exit Function
    End If

    ' Default sheet
    If wsConfig Is Nothing Then
        Set wsConfig = ThisWorkbook.Sheets("Config")
    End If

    Set dict = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Set tbl = wsConfig.ListObjects(MAPPING_TABLE_NAME)
    On Error GoTo 0

    If tbl Is Nothing Then
        DebugMessage "AutoValidationCommentPrefixMappingTable not found", MODULE_NAME
        Set GetAutoValidationMap = dict
        Exit Function
    End If
    
    DebugMessage "Table found: " & MAPPING_TABLE_NAME & " (" & tbl.ListRows.Count & " rows)", MODULE_NAME

    ' Build dictionary: Key = "Validate_Column_FuncName", Value = Dictionary with metadata
    Dim devFunc As String, item As Object
    Dim rowNum As Long
    rowNum = 0
    
    For Each r In tbl.ListRows
        rowNum = rowNum + 1
        devFunc = "Validate_Column_" & SafeTrim(r.Range(1, tbl.ListColumns("Dev Function Names").Index).Value)
        
        Set item = CreateObject("Scripting.Dictionary")
        
        ' Try new column name first, fall back to legacy
        On Error Resume Next
        item("DropColHeader") = SafeTrim(r.Range(1, tbl.ListColumns("Drop in Column").Index).Value)
        item("PrefixEN") = SafeTrim(r.Range(1, tbl.ListColumns("Prefix to message").Index).Value)
        item("PrefixFR") = SafeTrim(r.Range(1, tbl.ListColumns("(FR) Prefix to message").Index).Value)
        
        ' Try "ReviewSheet Column Header" first (new), then "ReviewSheet Column Letter" (legacy)
        Dim colRef As String
        colRef = ""
        colRef = SafeTrim(r.Range(1, tbl.ListColumns("ReviewSheet Column Header").Index).Value)
        If Len(colRef) = 0 Then
            colRef = SafeTrim(r.Range(1, tbl.ListColumns("ReviewSheet Column Letter").Index).Value)
        End If
        item("ColumnRef") = colRef
        
        item("AutoValidate") = (LCase(SafeTrim(r.Range(1, tbl.ListColumns("AutoValidate").Index).Value)) = "true")
        
        ' Optional: RuleTableName
        Dim ruleTableCol As Long
        ruleTableCol = 0
        ruleTableCol = tbl.ListColumns("RuleTableName").Index
        If ruleTableCol > 0 Then
            item("RuleTable") = SafeTrim(r.Range(1, ruleTableCol).Value)
        Else
            item("RuleTable") = ""
        End If
        On Error GoTo 0
        
        If devFunc <> "Validate_Column_" Then
            dict(devFunc) = item
            DebugMessage "Row " & rowNum & " Processing: " & devFunc, MODULE_NAME
        End If
    Next r
    
    DebugMessage "Success: " & dict.Count & " | Skipped: " & (tbl.ListRows.Count - dict.Count), MODULE_NAME

    Set gAutoValidationMap = dict
    Set GetAutoValidationMap = dict
End Function


Public Function GetRuleTableNameFromAutoValMap( _
        AutoValMap As Object, _
        ByVal DevFuncName As String, _
        ByVal DefaultRuleTable As String) As String

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
' CLEAR CACHE (call when config changes)
' ======================================================

Public Sub ClearAutoValidationMapCache()
    Set gAutoValidationMap = Nothing
    DebugMessage "AutoValidationMap cache cleared", MODULE_NAME
End Sub


' ======================================================
' ROW-LEVEL DECISIONS
' ======================================================

Public Function ShouldValidateRow( _
        ByVal rowNum As Long, _
        wsTarget As Worksheet, _
        Optional ByVal ForceValidation As Boolean = False) As Boolean

    If ForceValidation Then
        ShouldValidateRow = True
        Exit Function
    End If

    ' Check ForceValidationTable
    Dim tbl As ListObject
    Dim wsConfig As Worksheet
    Dim colToCheck As String
    Dim buildingValue As String
    Dim checkValue As String
    Dim i As Long
    Dim TargetCol As Range

    Set wsConfig = ThisWorkbook.Sheets("Config")
    
    On Error Resume Next
    Set tbl = wsConfig.ListObjects("ForceValidationTable")
    On Error GoTo 0

    If tbl Is Nothing Then
        ShouldValidateRow = False
        Exit Function
    End If

    For i = 1 To tbl.ListRows.Count
        colToCheck = Trim(tbl.DataBodyRange(i, tbl.ListColumns("Column").Index).Value)
        buildingValue = Trim(tbl.DataBodyRange(i, tbl.ListColumns("IsBuildingColumnValue").Index).Value)
        
        If colToCheck <> "" Then
            On Error Resume Next
            Set TargetCol = wsTarget.Range(colToCheck & "1")
            On Error GoTo 0

            If Not TargetCol Is Nothing Then
                checkValue = Trim(wsTarget.Cells(rowNum, TargetCol.Column).Value)
                
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
    Next i

    ShouldValidateRow = False
End Function


Public Function ValidationTimeoutReached() As Boolean
    If ValidationCancelTimeout <= 0 Then Exit Function
    ValidationTimeoutReached = (Timer - ValidationStartTime) >= ValidationCancelTimeout
End Function


' ======================================================
' COLUMN METADATA
' ======================================================

Public Function GetValidationColumns(wsConfig As Worksheet) As Object
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
        DebugMessage "DDMFieldsInfo table not found or invalid", MODULE_NAME
        Set GetDDMValidationColumns = dict
        Exit Function
    End If
    
    ReferenceTableName = ReferenceTable("ValidationTableName")
    StartRowIndex = ReferenceTable("StartRowIndex")
    EndRowMaxIndex = ReferenceTable("EndRowIndex")
    
    DebugMessage "DDM Reference: " & ReferenceTableName & " (rows " & StartRowIndex & "-" & EndRowMaxIndex & ")", MODULE_NAME
    
    On Error Resume Next
    Set DDMRefTable = wsConfig.ListObjects("AutoCheckDataValidationTable")
    On Error GoTo 0
    
    If DDMRefTable Is Nothing Then
        DebugMessage "AutoCheckDataValidationTable not found", MODULE_NAME
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
            ' Get the review sheet column name (header in target table)
            Dim reviewColName As String
            On Error Resume Next
            reviewColName = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns("ReviewSheet Column Name").Index).Value))
            On Error GoTo 0
            
            If Len(reviewColName) = 0 Then
                DebugMessage "Row " & i & ": Missing ReviewSheet Column Name, skipping", MODULE_NAME
                GoTo NextRow
            End If
            
            Dim item As Object
            Set item = CreateObject("Scripting.Dictionary")
            
            ' Use the column NAME as the key (for header-based lookup)
            item("ReviewLetter") = reviewColName  ' This is now a header name, not a letter
            item("ColumnNameFR") = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns("Column Name (FR)").Index).Value))
            item("ColumnNameEN") = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns("Column Name").Index).Value))
            item("MenuFieldEN") = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns("MenuField Column (EN)").Index).Value))
            item("MenuFieldFR") = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns("MenuField Column (FR)").Index).Value))
            item("CommentDropCol") = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns("AutoComment Column").Index).Value))
            
            DebugMessage "DDM Config: " & reviewColName & " -> EN:" & item("MenuFieldEN") & " FR:" & item("MenuFieldFR"), MODULE_NAME
            
            Dim NonEmptyRangeEN As Range
            Dim NonEmptyRangeFR As Range
            
            Set NonEmptyRangeEN = GetNonEmptyRangeInColumn(ReferenceTableName, item("MenuFieldEN"), StartRowIndex, EndRowMaxIndex)
            Set NonEmptyRangeFR = GetNonEmptyRangeInColumn(ReferenceTableName, item("MenuFieldFR"), StartRowIndex, EndRowMaxIndex)
            
            Dim listEN As Variant, listFR As Variant
            
            If Not NonEmptyRangeEN Is Nothing Then
                listEN = GetValuesAsList(NonEmptyRangeEN)
                If IsArray(listEN) Then
                    item("ValidColumnListEN") = listEN
                    DebugMessage "  EN values loaded: " & UBound(listEN) & " items", MODULE_NAME
                End If
            Else
                item("ValidColumnListEN") = Array()
                DebugMessage "  EN values: NONE FOUND for column '" & item("MenuFieldEN") & "'", MODULE_NAME
            End If
            
            If Not NonEmptyRangeFR Is Nothing Then
                listFR = GetValuesAsList(NonEmptyRangeFR)
                If IsArray(listFR) Then
                    item("ValidColumnListFR") = listFR
                    DebugMessage "  FR values loaded: " & UBound(listFR) & " items", MODULE_NAME
                End If
            Else
                item("ValidColumnListFR") = Array()
                DebugMessage "  FR values: NONE FOUND for column '" & item("MenuFieldFR") & "'", MODULE_NAME
            End If
            
            dict.Add reviewColName, item
        End If
NextRow:
    Next r
    
    DebugMessage "GetDDMValidationColumns: " & dict.Count & " columns configured", MODULE_NAME
    Set GetDDMValidationColumns = dict
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


Private Function GetNonEmptyRangeInColumn(sheetName As String, colHeader As String, startRow As Long, endRow As Long) As Range
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim col As ListColumn
    Dim checkRange As Range
    Dim lastNonEmptyRow As Long
    Dim cell As Range
    Dim colNum As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        DebugMessage "GetNonEmptyRangeInColumn: Sheet '" & sheetName & "' not found", MODULE_NAME
        Exit Function
    End If
    
    If startRow <= 0 Or endRow < startRow Then Exit Function
    
    ' Try to find column by header name in first row or table
    ' First, check if there's a table on this sheet
    If ws.ListObjects.Count > 0 Then
        Set tbl = ws.ListObjects(1)
        On Error Resume Next
        Set col = tbl.ListColumns(colHeader)
        On Error GoTo 0
        
        If Not col Is Nothing Then
            ' Return the column's data
            Set GetNonEmptyRangeInColumn = col.DataBodyRange
            Exit Function
        End If
    End If
    
    ' Fallback: search header row for column name
    Dim headerRow As Range
    Set headerRow = ws.Rows(1)
    
    Dim foundCell As Range
    Set foundCell = headerRow.Find(What:=colHeader, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    
    If foundCell Is Nothing Then
        DebugMessage "GetNonEmptyRangeInColumn: Column '" & colHeader & "' not found in sheet '" & sheetName & "'", MODULE_NAME
        Exit Function
    End If
    
    colNum = foundCell.Column
    
    ' Get range from startRow to endRow
    Set checkRange = ws.Range(ws.Cells(startRow, colNum), ws.Cells(endRow, colNum))
    lastNonEmptyRow = 0
    
    For Each cell In checkRange.Cells
        If Trim(CStr(cell.Value)) <> "" Then lastNonEmptyRow = cell.Row
    Next cell
    
    If lastNonEmptyRow = 0 Then Exit Function
    
    Set GetNonEmptyRangeInColumn = ws.Range(ws.Cells(startRow, colNum), ws.Cells(lastNonEmptyRow, colNum))
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
' SAFE HELPERS
' ======================================================

Public Function SafeTrim(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Then
        SafeTrim = ""
    Else
        SafeTrim = Trim(CStr(v))
    End If
End Function
