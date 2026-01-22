Attribute VB_Name = "AV_Validators"
Option Explicit

' ======================================================
' AV_Validators.bas
' Public entry points for validation (routing layer)
' All complex logic delegated to AV_ValidationRules
' VERSION: 2.3 - Uses table header lookup
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
' Uses table header lookup via AutoValMap
' ======================================================
Public Function GetSiblingCell(cell As Range, sheetName As String, TargetFuncName As String, _
                               Optional AutoValMap As Object = Nothing) As Range
    
    ' Get the target table from engine
    Dim targetTable As ListObject
    Set targetTable = AV_Engine.CurrentTargetTable
    
    If targetTable Is Nothing Then
        AV_Core.DebugMessage "GetSiblingCell: No target table available", MODULE_NAME
        Exit Function
    End If
    
    ' Try to use AutoValMap if provided
    If Not AutoValMap Is Nothing Then
        Dim funcKey As String
        funcKey = "Validate_Column_" & TargetFuncName
        
        If AutoValMap.Exists(funcKey) Then
            Dim mapItem As Object
            Set mapItem = AutoValMap(funcKey)
            
            If mapItem.Exists("ColumnRef") Then
                Dim headerName As String
                headerName = CStr(mapItem("ColumnRef"))
                
                If Len(headerName) > 0 Then
                    ' Use table header lookup
                    Set GetSiblingCell = AV_DataAccess.GetCellByTableHeader(targetTable, headerName, cell.Row)
                    
                    If GetSiblingCell Is Nothing Then
                        AV_Core.DebugMessage "GetSiblingCell: Column '" & headerName & "' not found", MODULE_NAME
                    End If
                    Exit Function
                End If
            End If
        End If
    End If

    ' Fallback: try the config table directly
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets("Config")
    
    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsConfig.ListObjects("AutoValidationCommentPrefixMappingTable")
    On Error GoTo 0
    
    If lo Is Nothing Then
        AV_Core.DebugMessage "GetSiblingCell: AutoValidationCommentPrefixMappingTable not found", MODULE_NAME
        Exit Function
    End If

    Dim r As ListRow
    Dim headerFromTable As String
    
    For Each r In lo.ListRows
        If Trim(CStr(r.Range.Cells(1, lo.ListColumns("Dev Function Names").Index).Value)) = TargetFuncName Then
            ' Get column header from "ReviewSheet Column Header" column
            On Error Resume Next
            headerFromTable = Trim(CStr(r.Range.Cells(1, lo.ListColumns("ReviewSheet Column Header").Index).Value))
            On Error GoTo 0
            
            If Len(headerFromTable) > 0 Then
                Set GetSiblingCell = AV_DataAccess.GetCellByTableHeader(targetTable, headerFromTable, cell.Row)
                
                If GetSiblingCell Is Nothing Then
                    AV_Core.DebugMessage "GetSiblingCell: Column '" & headerFromTable & "' not found in table", MODULE_NAME
                End If
            End If
            Exit Function
        End If
    Next r
    
    AV_Core.DebugMessage "GetSiblingCell: Function '" & TargetFuncName & "' not found in mapping table", MODULE_NAME
End Function
