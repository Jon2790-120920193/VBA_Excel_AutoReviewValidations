Attribute VB_Name = "AV_Validators"
Option Explicit

' ======================================================
' AV_Validators.bas v2.1
' Public entry points for validation (routing layer)
' All complex logic delegated to AV_ValidationRules
' UPDATED: Uses AV_DataAccess and AV_Constants
' ======================================================

Private Const MODULE_NAME As String = "AV_Validators"

' ======================================================
' PUBLIC ENTRY POINTS (DO NOT RENAME)
' These are called dynamically via Application.Run
' ======================================================

' -------------------- Electricity --------------------
Public Sub Validate_Column_Electricity(cell As Range, sheetName As String, _
    Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)
    
    AV_ValidationRules.ValidatePairedFields cell, sheetName, "Electricity", "Electricity_Metered", _
        AV_Constants.TBL_ELECTRICITY_PAIRS, english, FormatMap, AutoValMap
End Sub

Public Sub Validate_Column_Electricity_Metered(cell As Range, sheetName As String, _
    Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)
    
    AV_ValidationRules.ValidatePairedFields cell, sheetName, "Electricity_Metered", "Electricity", _
        AV_Constants.TBL_ELECTRICITY_PAIRS, english, FormatMap, AutoValMap
End Sub

' -------------------- Plumbing --------------------
Public Sub Validate_Column_Plumbing(cell As Range, sheetName As String, _
    Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)
    
    AV_ValidationRules.ValidatePairedFields cell, sheetName, "Plumbing", "Water_Metered", _
        AV_Constants.TBL_PLUMBING_PAIRS, english, FormatMap, AutoValMap
End Sub

Public Sub Validate_Column_Water_Metered(cell As Range, sheetName As String, _
    Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)
    
    AV_ValidationRules.ValidatePairedFields cell, sheetName, "Water_Metered", "Plumbing", _
        AV_Constants.TBL_PLUMBING_PAIRS, english, FormatMap, AutoValMap
End Sub

' -------------------- GIW --------------------
Public Sub Validate_Column_GIWQuantity(cell As Range, sheetName As String, _
    Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)
    
    If AV_ValidationRules.Validate_GIWQuantity(cell, sheetName, "GIWQuantity", english, FormatMap, AutoValMap) Then
        Dim siblingCell As Range
        Set siblingCell = GetSiblingCell(cell, sheetName, "GIWIncluded", AutoValMap)
        If Not siblingCell Is Nothing Then
            AV_ValidationRules.Validate_GIWIncluded siblingCell, sheetName, "GIWIncluded", english, FormatMap, AutoValMap
        End If
    End If
End Sub

Public Sub Validate_Column_GIWIncluded(cell As Range, sheetName As String, _
    Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)
    
    If AV_ValidationRules.Validate_GIWIncluded(cell, sheetName, "GIWIncluded", english, FormatMap, AutoValMap) Then
        Dim siblingCell As Range
        Set siblingCell = GetSiblingCell(cell, sheetName, "GIWQuantity", AutoValMap)
        If Not siblingCell Is Nothing Then
            AV_ValidationRules.Validate_GIWQuantity siblingCell, sheetName, "GIWQuantity", english, FormatMap, AutoValMap
        End If
    End If
End Sub

' -------------------- Heat --------------------
Public Sub Validate_Column_Heat_Source(cell As Range, sheetName As String, _
    Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)
    
    AV_ValidationRules.Validate_HeatPairs cell, sheetName, "Heat_Source", english, 0, FormatMap, AutoValMap
End Sub

Public Sub Validate_Column_Heat_Metered(cell As Range, sheetName As String, _
    Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)
    
    AV_ValidationRules.Validate_HeatPairs cell, sheetName, "Heat_Metered", english, 0, FormatMap, AutoValMap
End Sub

' -------------------- Construction Date --------------------
Public Sub Validate_Column_Construction_Date(cell As Range, sheetName As String, _
    Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)
    
    AV_ValidationRules.Validate_ConstructionDate cell, sheetName, english, FormatMap, AutoValMap
End Sub

' ======================================================
' SHARED HELPER - GET SIBLING CELL (Enhanced v2.1)
' ======================================================
Public Function GetSiblingCell(cell As Range, sheetName As String, TargetFuncName As String, _
                              Optional AutoValMap As Object) As Range
    Dim ws As Worksheet
    Dim wsConfig As Worksheet
    Dim tbl As ListObject
    Dim colLetter As String
    Dim funcName As String
    
    Set ws = ThisWorkbook.Sheets(sheetName)
    Set wsConfig = ThisWorkbook.Sheets(AV_Constants.CONFIG_SHEET_NAME)

    ' Load AutoValMap if not provided
    If AutoValMap Is Nothing Then
        Set AutoValMap = AV_Core.GetAutoValidationMap(wsConfig)
    End If
    
    If AutoValMap Is Nothing Then
        AV_Core.DebugMessage "AutoValidation map not available", MODULE_NAME
        Exit Function
    End If
    
    ' Build full function name
    funcName = "Validate_Column_" & TargetFuncName
    
    ' Check if function exists in map
    If Not AutoValMap.Exists(funcName) Then
        AV_Core.DebugMessage "Function '" & funcName & "' not found in AutoValidation map", MODULE_NAME
        Exit Function
    End If
    
    ' Get column reference from map
    colLetter = AV_Core.SafeTrim(AutoValMap(funcName)("ColumnRef"))
    
    If Len(colLetter) = 0 Then
        AV_Core.DebugMessage "No column reference found for '" & funcName & "'", MODULE_NAME
        Exit Function
    End If
    
    ' Build cell reference
    On Error Resume Next
    Set GetSiblingCell = ws.Range(colLetter & cell.Row)
    On Error GoTo 0
    
    If GetSiblingCell Is Nothing Then
        AV_Core.DebugMessage "Unable to build cell reference: " & colLetter & cell.Row, MODULE_NAME
    End If
End Function
