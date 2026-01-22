Attribute VB_Name = "AV_Validators"
Option Explicit

' ======================================================
' AV_Validators.bas v2.6
' Public entry points for validation (routing layer)
' All complex logic delegated to AV_ValidationRules
' KEY FIX: GetSiblingCell now uses "ReviewSheet Column Header" not "ReviewSheet Column Letter"
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
' SHARED HELPER - GET SIBLING CELL
' KEY FIX: Now uses "ReviewSheet Column Header" (header name, not column letter)
' Also uses AutoValMap cache if provided
' ======================================================
Public Function GetSiblingCell(cell As Range, sheetName As String, TargetFuncName As String, _
    Optional AutoValMap As Object = Nothing) As Range
    
    Dim ws As Worksheet
    Dim wsConfig As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    Set wsConfig = ThisWorkbook.Sheets("Config")
    
    ' First try using cached AutoValMap
    If Not AutoValMap Is Nothing Then
        Dim fullFuncName As String
        fullFuncName = "Validate_Column_" & TargetFuncName
        
        If AutoValMap.Exists(fullFuncName) Then
            Dim targetHeaderName As String
            targetHeaderName = AutoValMap(fullFuncName)("ColumnRef")
            
            If Len(targetHeaderName) > 0 Then
                ' Use the current target table from AV_Engine
                If Not AV_Engine.CurrentTargetTable Is Nothing Then
                    Set GetSiblingCell = AV_Core.GetCellByHeader(AV_Engine.CurrentTargetTable, cell.Row, targetHeaderName)
                    Exit Function
                End If
            End If
        End If
    End If
    
    ' Fallback: Read directly from config table
    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsConfig.ListObjects("AutoValidationCommentPrefixMappingTable")
    On Error GoTo 0
    
    If lo Is Nothing Then Exit Function

    Dim r As ListRow
    For Each r In lo.ListRows
        If Trim(r.Range.Cells(1, lo.ListColumns("Dev Function Names").Index).Value) = TargetFuncName Then
            ' *** KEY FIX: Use "ReviewSheet Column Header" not "ReviewSheet Column Letter" ***
            Dim headerName As String
            headerName = Trim(CStr(r.Range.Cells(1, lo.ListColumns("ReviewSheet Column Header").Index).Value))
            
            If Len(headerName) > 0 Then
                ' Use header-based lookup with current target table
                If Not AV_Engine.CurrentTargetTable Is Nothing Then
                    Set GetSiblingCell = AV_Core.GetCellByHeader(AV_Engine.CurrentTargetTable, cell.Row, headerName)
                End If
            End If
            Exit Function
        End If
    Next r
End Function
