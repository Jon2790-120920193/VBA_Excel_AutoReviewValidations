Attribute VB_Name = "AV_Validators"
Option Explicit

' ======================================================
' AV_Validators.bas
' Public entry points for validation (routing layer)
' All complex logic delegated to AV_ValidationRules
' VERSION: 2.5 - Table-based header lookup
' ======================================================

Private Const MODULE_NAME As String = "AV_Validators"
Public Const MODULE_VERSION As String = "2.5"

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
' SHARED HELPER - GET SIBLING CELL (v2.5 - Header-based)
' Now uses column headers instead of column letters
' ======================================================
Public Function GetSiblingCell(cell As Range, sheetName As String, TargetFuncName As String, _
                               Optional AutoValMap As Object = Nothing) As Range
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' Try to use AV_Engine's current target table
    Dim targetTable As ListObject
    Set targetTable = AV_Engine.CurrentTargetTable
    
    ' If no current table, try to find the table containing the cell
    If targetTable Is Nothing Then
        Set targetTable = GetTableContainingCell(cell)
    End If
    
    If targetTable Is Nothing Then
        AV_Core.DebugMessage "GetSiblingCell: No target table found for " & TargetFuncName, MODULE_NAME
        Exit Function
    End If
    
    ' Get the column header for the target function
    Dim targetHeader As String
    targetHeader = GetColumnHeaderForFunction(TargetFuncName, AutoValMap)
    
    If Len(targetHeader) = 0 Then
        AV_Core.DebugMessage "GetSiblingCell: No column header found for " & TargetFuncName, MODULE_NAME
        Exit Function
    End If
    
    ' Find the column in the target table
    Dim col As ListColumn
    On Error Resume Next
    Set col = targetTable.ListColumns(targetHeader)
    On Error GoTo ErrHandler
    
    If col Is Nothing Then
        AV_Core.DebugMessage "GetSiblingCell: Column '" & targetHeader & "' not found in table", MODULE_NAME
        Exit Function
    End If
    
    ' Return the cell at the same row but in the sibling column
    Set GetSiblingCell = ws.Cells(cell.Row, col.DataBodyRange.Column)
    Exit Function
    
ErrHandler:
    AV_Core.DebugMessage "GetSiblingCell ERROR: " & Err.Description, MODULE_NAME
End Function


' ======================================================
' HELPER: Get column header for a function name
' Looks up in AutoValidationCommentPrefixMappingTable
' ======================================================
Private Function GetColumnHeaderForFunction(funcName As String, Optional AutoValMap As Object = Nothing) As String
    ' First try the cached AutoValMap if provided
    If Not AutoValMap Is Nothing Then
        Dim fullFuncName As String
        fullFuncName = "Validate_Column_" & funcName
        
        If AutoValMap.Exists(fullFuncName) Then
            Dim mapItem As Object
            Set mapItem = AutoValMap(fullFuncName)
            If mapItem.Exists("ColumnRef") Then
                GetColumnHeaderForFunction = CStr(mapItem("ColumnRef"))
                Exit Function
            End If
        End If
    End If
    
    ' Fall back to table lookup
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets("Config")
    
    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsConfig.ListObjects("AutoValidationCommentPrefixMappingTable")
    On Error GoTo 0
    
    If lo Is Nothing Then Exit Function
    
    ' Find the column index for "ReviewSheet Column Header" (new) or "ReviewSheet Column Letter" (legacy)
    Dim colIdx As Long
    colIdx = 0
    
    On Error Resume Next
    colIdx = lo.ListColumns("ReviewSheet Column Header").Index
    On Error GoTo 0
    
    ' If new column not found, try legacy column name
    If colIdx = 0 Then
        On Error Resume Next
        colIdx = lo.ListColumns("ReviewSheet Column Letter").Index
        On Error GoTo 0
    End If
    
    If colIdx = 0 Then Exit Function
    
    ' Find the function row
    Dim funcColIdx As Long
    On Error Resume Next
    funcColIdx = lo.ListColumns("Dev Function Names").Index
    On Error GoTo 0
    
    If funcColIdx = 0 Then Exit Function
    
    Dim r As ListRow
    For Each r In lo.ListRows
        If StrComp(Trim(CStr(r.Range.Cells(1, funcColIdx).Value)), funcName, vbTextCompare) = 0 Then
            GetColumnHeaderForFunction = Trim(CStr(r.Range.Cells(1, colIdx).Value))
            Exit Function
        End If
    Next r
End Function


' ======================================================
' HELPER: Find table containing a cell
' ======================================================
Private Function GetTableContainingCell(cell As Range) As ListObject
    If cell Is Nothing Then Exit Function
    
    Dim ws As Worksheet
    Set ws = cell.Worksheet
    
    Dim tbl As ListObject
    For Each tbl In ws.ListObjects
        If Not Intersect(cell, tbl.Range) Is Nothing Then
            Set GetTableContainingCell = tbl
            Exit Function
        End If
    Next tbl
End Function
