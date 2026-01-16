Attribute VB_Name = "AV_Core"
Option Explicit

' ======================================================
' AV_Core
' Core services: mapping, routing, row logic, debug flags
' ======================================================


' ======================================================
' GLOBAL STATE (DECLARED AT TOP — DO NOT MOVE)
' ======================================================

Public ValidationStartTime As Single
Public ValidationCancelTimeout As Single
Public ValidationCancelFlag As Boolean

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
        Set GetAutoValidationMap = dict
        Exit Function
    End If

    ' Build dictionary: Key = "Validate_Column_FuncName", Value = Dictionary with metadata
    Dim devFunc As String, item As Object
    For Each r In tbl.ListRows
        devFunc = "Validate_Column_" & SafeTrim(r.Range(1, tbl.ListColumns("Dev Function Names").Index).Value)
        
        Set item = CreateObject("Scripting.Dictionary")
        item("DropColHeader") = SafeTrim(r.Range(1, tbl.ListColumns("Drop in Column").Index).Value)
        item("PrefixEN") = SafeTrim(r.Range(1, tbl.ListColumns("Prefix to message").Index).Value)
        item("PrefixFR") = SafeTrim(r.Range(1, tbl.ListColumns("(FR) Prefix to message").Index).Value)
        item("ColumnRef") = SafeTrim(r.Range(1, tbl.ListColumns("ReviewSheet Column Letter").Index).Value)
        item("AutoValidate") = (LCase(SafeTrim(r.Range(1, tbl.ListColumns("AutoValidate").Index).Value)) = "true")
        
        If devFunc <> "Validate_Column_" Then
            dict(devFunc) = item
        End If
    Next r

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

    If AutoValMap.Exists(DevFuncName) Then
        GetRuleTableNameFromAutoValMap = AutoValMap(DevFuncName)(1)
    Else
        GetRuleTableNameFromAutoValMap = DefaultRuleTable
    End If
End Function


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
        Set GetDDMValidationColumns = dict
        Exit Function
    End If
    
    ReferenceTableName = ReferenceTable("ValidationTableName")
    StartRowIndex = ReferenceTable("StartRowIndex")
    EndRowMaxIndex = ReferenceTable("EndRowIndex")
    
    On Error Resume Next
    Set DDMRefTable = wsConfig.ListObjects("AutoCheckDataValidationTable")
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
        autoCheckVal = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns("AutoCheck").Index).Value))
        
        If StrComp(autoCheckVal, "TRUE", vbTextCompare) = 0 Then
            Dim key As String
            key = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns("ReviewSheet Column Letter").Index).Value))
            
            Dim item As Object
            Set item = CreateObject("Scripting.Dictionary")
            
            item("ReviewLetter") = key
            item("ColumnNameFR") = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns("Column Name (FR)").Index).Value))
            item("ColumnNameEN") = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns("Column Name").Index).Value))
            item("MenuFieldEN") = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns("MenuField Column (EN)").Index).Value))
            item("MenuFieldFR") = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns("MenuField Column (FR)").Index).Value))
            item("CommentDropCol") = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns("AutoComment Column").Index).Value))
            
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
    Set tbl = wsConfig.ListObjects("DDMFieldsInfo")
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
' SAFE HELPERS
' ======================================================

Public Function SafeTrim(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Then
        SafeTrim = ""
    Else
        SafeTrim = Trim(CStr(v))
    End If
End Function
