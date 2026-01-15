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
Public Const AV2_FALLBACKFORMAT As String = "Default"


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

Public Function GetAutoValidationMap(wsConfig As Worksheet) As Object
    Dim tbl As ListObject
    Dim r As ListRow
    Dim dict As Object

    If Not gAutoValidationMap Is Nothing Then
        Set GetAutoValidationMap = gAutoValidationMap
        Exit Function
    End If

    Set dict = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Set tbl = wsConfig.ListObjects(MAPPING_TABLE_NAME)
    On Error GoTo 0

    If tbl Is Nothing Then
        Set GetAutoValidationMap = dict
        Exit Function
    End If

    For Each r In tbl.ListRows
        dict(r.Range(1, 1).Value) = Array( _
            r.Range(1, 2).Value, _
            r.Range(1, 3).Value _
        )
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

    ' Default behavior: always validate
    ShouldValidateRow = True
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
        dict(wsConfig.Range("C" & r).Value) = wsConfig.Range("B" & r).Value
        r = r + 1
    Loop

    Set GetValidationColumns = dict
End Function


Public Function GetDDMValidationColumns(wsConfig As Worksheet) As Object
    Dim dict As Object
    Dim r As Long

    Set dict = CreateObject("Scripting.Dictionary")
    r = 6

    Do While Trim(wsConfig.Range("B" & r).Value) <> ""
        dict(wsConfig.Range("B" & r).Value) = wsConfig.Range("C" & r).Value
        r = r + 1
    Loop

    Set GetDDMValidationColumns = dict
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


