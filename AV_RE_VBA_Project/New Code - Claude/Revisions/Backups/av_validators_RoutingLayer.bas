Attribute VB_Name = "AV_Validators"
Option Explicit

' ======================================================
' AV_Validators.bas
' Public entry points for validation (routing layer)
' All complex logic delegated to AV_ValidationRules
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
        "ElectricityPairValidation", english, FormatMap, AutoValMap
End Sub

Public Sub Validate_Column_Electricity_Metered(cell As Range, sheetName As String, _
    Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)
    
    AV_ValidationRules.ValidatePairedFields cell, sheetName, "Electricity_Metered", "Electricity", _
        "ElectricityPairValidation", english, FormatMap, AutoValMap
End Sub

' -------------------- Plumbing --------------------
Public Sub Validate_Column_Plumbing(cell As Range, sheetName As String, _
    Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)
    
    AV_ValidationRules.ValidatePairedFields cell, sheetName, "Plumbing", "Water_Metered", _
        "PlumbingPairValidation", english, FormatMap, AutoValMap
End Sub

Public Sub Validate_Column_Water_Metered(cell As Range, sheetName As String, _
    Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)
    
    AV_ValidationRules.ValidatePairedFields cell, sheetName, "Water_Metered", "Plumbing", _
        "PlumbingPairValidation", english, FormatMap, AutoValMap
End Sub

' -------------------- GIW --------------------
Public Sub Validate_Column_GIWQuantity(cell As Range, sheetName As String, _
    Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)
    
    If AV_ValidationRules.Validate_GIWQuantity(cell, sheetName, "GIWQuantity", english, FormatMap, AutoValMap) Then
        Dim siblingCell As Range
        Set siblingCell = GetSiblingCell(cell, sheetName, "GIWIncluded")
        If Not siblingCell Is Nothing Then
            AV_ValidationRules.Validate_GIWIncluded siblingCell, sheetName, "GIWIncluded", english, FormatMap, AutoValMap
        End If
    End If
End Sub

Public Sub Validate_Column_GIWIncluded(cell As Range, sheetName As String, _
    Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)
    
    If AV_ValidationRules.Validate_GIWIncluded(cell, sheetName, "GIWIncluded", english, FormatMap, AutoValMap) Then
        Dim siblingCell As Range
        Set siblingCell = GetSiblingCell(cell, sheetName, "GIWQuantity")
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
' SHARED HELPER - GET SIBLING CELL
' ======================================================
Public Function GetSiblingCell(cell As Range, sheetName As String, TargetFuncName As String) As Range
    Dim ws As Worksheet
    Dim wsConfig As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    Set wsConfig = ThisWorkbook.Sheets("Config")

    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsConfig.ListObjects("AutoValidationCommentPrefixMappingTable")
    On Error GoTo 0
    
    If lo Is Nothing Then Exit Function

    Dim r As ListRow
    For Each r In lo.ListRows
        If Trim(r.Range.Columns("Dev Function Names").Value) = TargetFuncName Then
            On Error Resume Next
            Set GetSiblingCell = ws.Range(r.Range.Columns("ReviewSheet Column Letter").Value & cell.Row)
            On Error GoTo 0
            Exit Function
        End If
    Next r
End Function
