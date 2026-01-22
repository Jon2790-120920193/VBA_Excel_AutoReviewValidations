Attribute VB_Name = "AV_Engine"
Option Explicit

' ======================================================
' AV_Engine.bas v2.1 - FIXED ValidateSingleRow
' NOW USES TABLE COLUMN HEADERS instead of column letters
' ======================================================

Private Const MODULE_NAME As String = "AV_Engine"

' ======================================================
' VALIDATE SINGLE ROW - FIXED FOR TABLE-BASED COLUMNS
' ======================================================
Public Sub ValidateSingleRow(wsData As Worksheet, rowNum As Long, AdvFunctionMap As Object, english As Boolean, FormatMap As Object)
    Dim colLetter As Variant
    Dim funcName As String
    Dim TargetCell As Range
    Dim mapItem As Object
    Dim AutoValidate As Boolean
    Dim TargetColumnHeader As String
    Dim tbl As ListObject
    Dim colIndex As Long
    Dim relativeRow As Long

    ' Find the table in this worksheet
    On Error Resume Next
    Set tbl = wsData.ListObjects(1)  ' Assumes first table on sheet
    On Error GoTo 0
    
    If tbl Is Nothing Then
        AV_Core.DebugMessage "ERROR: No ListObject found in sheet " & wsData.Name, MODULE_NAME
        Exit Sub
    End If

    For Each colLetter In AdvFunctionMap.Keys
        Set mapItem = AdvFunctionMap(colLetter)
        funcName = CStr(colLetter)

        ' Retrieve AutoValidate flag
        AutoValidate = False
        If mapItem.Exists("AutoValidate") Then
            AutoValidate = mapItem("AutoValidate")
        End If
        
        ' Retrieve ColumnRef (header name)
        TargetColumnHeader = ""
        If mapItem.Exists("ColumnRef") Then
            TargetColumnHeader = CStr(mapItem("ColumnRef"))
        End If
        
        If Len(TargetColumnHeader) = 0 Then
            AV_Core.DebugMessage "WARNING: Missing ColumnRef for " & funcName, MODULE_NAME
            GoTo SkipToNext
        End If
        
        ' Skip if AutoValidate = False
        If AutoValidate = False Then
            AV_Core.DebugMessage "Skipping " & funcName & " (AutoValidate=False)", MODULE_NAME
            GoTo SkipToNext
        End If

        ' Find column by header name in table
        colIndex = 0
        On Error Resume Next
        colIndex = tbl.ListColumns(TargetColumnHeader).Index
        On Error GoTo 0
        
        If colIndex = 0 Then
            AV_Core.DebugMessage "WARNING: Column '" & TargetColumnHeader & "' not found in table for " & funcName, MODULE_NAME
            GoTo SkipToNext
        End If
        
        ' Calculate relative row within table (rowNum is absolute, need relative to table start)
        relativeRow = rowNum - tbl.HeaderRowRange.Row
        
        ' Get the cell at this row and column within the table
        On Error Resume Next
        Set TargetCell = tbl.ListColumns(TargetColumnHeader).DataBodyRange.Cells(relativeRow, 1)
        On Error GoTo 0
        
        If Not TargetCell Is Nothing Then
            On Error GoTo ValidationError
            AV_Core.DebugMessage "Validating row " & rowNum & ", column '" & TargetColumnHeader & "' with " & funcName, MODULE_NAME
            Application.Run funcName, TargetCell, wsData.Name, english, FormatMap, AdvFunctionMap
            On Error GoTo 0
        Else
            AV_Core.DebugMessage "WARNING: Could not get cell for row " & rowNum & ", column '" & TargetColumnHeader & "'", MODULE_NAME
        End If
        
SkipToNext:
    Next colLetter

    ' Only log every N rows to reduce clutter
    If rowNum Mod AV_Constants.VALIDATION_DETAILED_LOG_INTERVAL = 0 Then
        AV_UI.AppendUserLog "Row " & rowNum & " validation complete"
    End If
    
    Exit Sub

ValidationError:
    AV_Core.DebugMessage "Error validating row " & rowNum & ", column " & TargetColumnHeader & ", function: " & funcName & " - " & Err.Description, MODULE_NAME
    AV_UI.AppendUserLog "Warning: Error in row " & rowNum & ", column " & TargetColumnHeader
    Resume Next
End Sub
