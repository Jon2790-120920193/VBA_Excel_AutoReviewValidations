Attribute VB_Name = "AV_Validators"
Option Explicit

' ======================================================
' AV_Validators.bas
' Public entry points for validation (routing layer)
' All complex logic delegated to AV_ValidationRules
' VERSION: 2.2 - Uses header-based cell lookup
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
' Updated to work with both column letters and header names
' ======================================================
Public Function GetSiblingCell(cell As Range, sheetName As String, TargetFuncName As String, _
                               Optional AutoValMap As Object = Nothing) As Range
    Dim ws As Worksheet
    Dim wsConfig As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    Set wsConfig = ThisWorkbook.Sheets("Config")
    
    ' Try to use AutoValMap if provided
    If Not AutoValMap Is Nothing Then
        Dim funcKey As String
        funcKey = "Validate_Column_" & TargetFuncName
        
        If AutoValMap.Exists(funcKey) Then
            Dim mapItem As Object
            Set mapItem = AutoValMap(funcKey)
            
            If mapItem.Exists("ColumnRef") Then
                Dim colRef As String
                colRef = CStr(mapItem("ColumnRef"))
                
                If Len(colRef) > 0 Then
                    ' Use smart cell lookup (handles both letters and header names)
                    Dim targetTable As ListObject
                    Set targetTable = AV_Engine.CurrentTargetTable
                    
                    Set GetSiblingCell = AV_DataAccess.GetCellSmart(ws, colRef, cell.Row, targetTable)
                    
                    If GetSiblingCell Is Nothing Then
                        AV_Core.DebugMessage "Unable to build cell reference: " & colRef & cell.Row, MODULE_NAME
                    End If
                    Exit Function
                End If
            End If
        End If
    End If

    ' Fallback: try the config table directly
    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsConfig.ListObjects("AutoValidationCommentPrefixMappingTable")
    On Error GoTo 0
    
    If lo Is Nothing Then Exit Function

    Dim r As ListRow
    Dim colRefFromTable As String
    
    For Each r In lo.ListRows
        If Trim(CStr(r.Range.Cells(1, lo.ListColumns("Dev Function Names").Index).Value)) = TargetFuncName Then
            ' Try to get column reference - check both old and new column names
            On Error Resume Next
            colRefFromTable = Trim(CStr(r.Range.Cells(1, lo.ListColumns("ReviewSheet Column Header").Index).Value))
            If Len(colRefFromTable) = 0 Then
                colRefFromTable = Trim(CStr(r.Range.Cells(1, lo.ListColumns("ReviewSheet Column Letter").Index).Value))
            End If
            On Error GoTo 0
            
            If Len(colRefFromTable) > 0 Then
                ' Use smart cell lookup
                Dim tbl As ListObject
                Set tbl = AV_Engine.CurrentTargetTable
                
                Set GetSiblingCell = AV_DataAccess.GetCellSmart(ws, colRefFromTable, cell.Row, tbl)
                
                If GetSiblingCell Is Nothing Then
                    AV_Core.DebugMessage "Unable to build cell reference: " & colRefFromTable & cell.Row, MODULE_NAME
                End If
            End If
            Exit Function
        End If
    Next r
End Function
