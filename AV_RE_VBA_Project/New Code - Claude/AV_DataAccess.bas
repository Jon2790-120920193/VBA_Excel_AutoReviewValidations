Attribute VB_Name = "AV_DataAccess"
Option Explicit

' ======================================================
' AV_DataAccess
' Centralized table and data access layer
' VERSION: 2.3 - Fully table-based, header lookup
' ALL table operations should go through this module
' ======================================================

Private Const MODULE_NAME As String = "AV_DataAccess"

' ======================================================
' TABLE EXISTENCE & VALIDATION
' ======================================================

Public Function TableExists(ws As Worksheet, tableName As String) As Boolean
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    TableExists = Not (tbl Is Nothing)
    On Error GoTo 0
End Function

Public Function WorksheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    WorksheetExists = Not (ws Is Nothing)
    On Error GoTo 0
End Function

Public Function ColumnExistsInTable(tbl As ListObject, columnName As String) As Boolean
    If tbl Is Nothing Then Exit Function
    
    Dim col As ListColumn
    On Error Resume Next
    Set col = tbl.ListColumns(columnName)
    ColumnExistsInTable = Not (col Is Nothing)
    On Error GoTo 0
End Function

' ======================================================
' GET TABLE REFERENCES
' ======================================================

Public Function GetTable(ws As Worksheet, tableName As String) As ListObject
    On Error Resume Next
    Set GetTable = ws.ListObjects(tableName)
    On Error GoTo 0
    
    If GetTable Is Nothing Then
        AV_Core.DebugMessage "Table not found: " & tableName & " in sheet: " & ws.Name, MODULE_NAME
    End If
End Function

Public Function GetFirstTable(ws As Worksheet) As ListObject
    If ws.ListObjects.Count > 0 Then
        Set GetFirstTable = ws.ListObjects(1)
    Else
        AV_Core.DebugMessage "No tables found on sheet: " & ws.Name, MODULE_NAME
    End If
End Function

Public Function FindTableByName(tableName As String) As ListObject
    ' Search all worksheets for a table by name
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set tbl = ws.ListObjects(tableName)
        On Error GoTo 0
        
        If Not tbl Is Nothing Then
            Set FindTableByName = tbl
            Exit Function
        End If
    Next ws
    
    AV_Core.DebugMessage "Table '" & tableName & "' not found in any worksheet", MODULE_NAME
End Function

' ======================================================
' CELL LOOKUP BY TABLE HEADER
' Core function for header-based cell access
' ======================================================

Public Function GetCellByTableHeader(tbl As ListObject, headerName As String, rowNum As Long) As Range
    ' Returns the cell at the given row for the column with headerName
    ' rowNum is the ABSOLUTE worksheet row number
    
    If tbl Is Nothing Then
        AV_Core.DebugMessage "GetCellByTableHeader: Table is Nothing", MODULE_NAME
        Exit Function
    End If
    
    If tbl.DataBodyRange Is Nothing Then
        AV_Core.DebugMessage "GetCellByTableHeader: Table has no data", MODULE_NAME
        Exit Function
    End If
    
    ' Find the column
    Dim col As ListColumn
    On Error Resume Next
    Set col = tbl.ListColumns(headerName)
    On Error GoTo 0
    
    If col Is Nothing Then
        AV_Core.DebugMessage "GetCellByTableHeader: Column '" & headerName & "' not found in table " & tbl.Name, MODULE_NAME
        Exit Function
    End If
    
    ' Validate row is within table bounds
    Dim tableStartRow As Long
    Dim tableEndRow As Long
    tableStartRow = tbl.DataBodyRange.Row
    tableEndRow = tableStartRow + tbl.DataBodyRange.Rows.Count - 1
    
    If rowNum < tableStartRow Or rowNum > tableEndRow Then
        AV_Core.DebugMessage "GetCellByTableHeader: Row " & rowNum & " is outside table range (" & tableStartRow & "-" & tableEndRow & ")", MODULE_NAME
        Exit Function
    End If
    
    ' Return the cell
    Set GetCellByTableHeader = tbl.Parent.Cells(rowNum, col.DataBodyRange.Column)
End Function

Public Function GetCellByTableHeaderSafe(tbl As ListObject, headerName As String, rowNum As Long, _
                                         Optional suppressWarnings As Boolean = False) As Range
    ' Same as GetCellByTableHeader but with option to suppress debug warnings
    ' Useful when checking if column exists without generating log spam
    
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    
    Dim col As ListColumn
    On Error Resume Next
    Set col = tbl.ListColumns(headerName)
    On Error GoTo 0
    
    If col Is Nothing Then Exit Function
    
    Dim tableStartRow As Long
    Dim tableEndRow As Long
    tableStartRow = tbl.DataBodyRange.Row
    tableEndRow = tableStartRow + tbl.DataBodyRange.Rows.Count - 1
    
    If rowNum < tableStartRow Or rowNum > tableEndRow Then Exit Function
    
    Set GetCellByTableHeaderSafe = tbl.Parent.Cells(rowNum, col.DataBodyRange.Column)
End Function

' ======================================================
' SMART CELL LOOKUP
' Works with current engine's target table
' ======================================================

Public Function GetCellSmart(ws As Worksheet, colRef As String, rowNum As Long, _
                             Optional tbl As ListObject = Nothing) As Range
    ' Smart function that works with header names
    ' Uses the current target table from AV_Engine if not provided
    
    If Len(colRef) = 0 Then Exit Function
    
    ' Get table reference
    If tbl Is Nothing Then
        Set tbl = AV_Engine.CurrentTargetTable
    End If
    
    If tbl Is Nothing Then
        AV_Core.DebugMessage "GetCellSmart: No table available", MODULE_NAME
        Exit Function
    End If
    
    ' Use header-based lookup
    Set GetCellSmart = GetCellByTableHeader(tbl, colRef, rowNum)
End Function

' ======================================================
' GET TABLE COLUMN RANGE
' ======================================================

Public Function GetTableColumn(tbl As ListObject, columnName As String) As Range
    If tbl Is Nothing Then Exit Function
    
    On Error Resume Next
    Set GetTableColumn = tbl.ListColumns(columnName).DataBodyRange
    On Error GoTo 0
    
    If GetTableColumn Is Nothing Then
        AV_Core.DebugMessage "Column not found: " & columnName & " in table: " & tbl.Name, MODULE_NAME
    End If
End Function

' ======================================================
' GET VALUES FROM TABLES
' ======================================================

Public Function GetTableValue(tbl As ListObject, rowIndex As Long, columnName As String) As Variant
    ' rowIndex is 1-based within the table data (not worksheet row)
    If tbl Is Nothing Then
        GetTableValue = Empty
        Exit Function
    End If
    
    On Error Resume Next
    GetTableValue = tbl.ListColumns(columnName).DataBodyRange(rowIndex).value
    On Error GoTo 0
End Function

Public Function GetTableValueByKey(tbl As ListObject, _
                                   keyColumn As String, _
                                   keyValue As Variant, _
                                   valueColumn As String) As Variant
    If tbl Is Nothing Then Exit Function
    
    Dim r As ListRow
    Dim keyColIndex As Long, valColIndex As Long
    
    On Error Resume Next
    keyColIndex = tbl.ListColumns(keyColumn).Index
    valColIndex = tbl.ListColumns(valueColumn).Index
    On Error GoTo 0
    
    If keyColIndex = 0 Or valColIndex = 0 Then Exit Function
    
    For Each r In tbl.ListRows
        If r.Range.Cells(1, keyColIndex).value = keyValue Then
            GetTableValueByKey = r.Range.Cells(1, valColIndex).value
            Exit Function
        End If
    Next r
End Function

' ======================================================
' GET ENTIRE ROW AS DICTIONARY
' ======================================================

Public Function GetTableRow(tbl As ListObject, keyColumn As String, keyValue As Variant) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    If tbl Is Nothing Then
        Set GetTableRow = dict
        Exit Function
    End If
    
    Dim r As ListRow
    Dim col As ListColumn
    Dim keyColIndex As Long
    
    On Error Resume Next
    keyColIndex = tbl.ListColumns(keyColumn).Index
    On Error GoTo 0
    
    If keyColIndex = 0 Then
        Set GetTableRow = dict
        Exit Function
    End If
    
    For Each r In tbl.ListRows
        If r.Range.Cells(1, keyColIndex).value = keyValue Then
            For Each col In tbl.ListColumns
                dict(col.Name) = r.Range.Cells(1, col.Index).value
            Next col
            Set GetTableRow = dict
            Exit Function
        End If
    Next r
    
    Set GetTableRow = dict
End Function

' ======================================================
' CHECK VALUE EXISTENCE
' ======================================================

Public Function TableContainsValue(tbl As ListObject, _
                                   columnName As String, _
                                   searchValue As Variant, _
                                   Optional caseSensitive As Boolean = False) As Boolean
    Dim colRange As Range
    Set colRange = GetTableColumn(tbl, columnName)
    
    If colRange Is Nothing Then Exit Function
    
    Dim cell As Range
    For Each cell In colRange.Cells
        If caseSensitive Then
            If cell.value = searchValue Then
                TableContainsValue = True
                Exit Function
            End If
        Else
            If StrComp(CStr(cell.value), CStr(searchValue), vbTextCompare) = 0 Then
                TableContainsValue = True
                Exit Function
            End If
        End If
    Next cell
End Function

' ======================================================
' GET COLUMN VALUES AS ARRAY
' ======================================================

Public Function GetColumnValues(tbl As ListObject, columnName As String) As Variant
    Dim colRange As Range
    Set colRange = GetTableColumn(tbl, columnName)
    
    If colRange Is Nothing Then
        GetColumnValues = Array()
        Exit Function
    End If
    
    Dim values() As Variant
    ReDim values(1 To colRange.Rows.Count)
    
    Dim i As Long
    For i = 1 To colRange.Rows.Count
        values(i) = colRange.Cells(i, 1).value
    Next i
    
    GetColumnValues = values
End Function

' ======================================================
' TABLE INFO HELPERS
' ======================================================

Public Function GetTableRowCount(tbl As ListObject) As Long
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    GetTableRowCount = tbl.DataBodyRange.Rows.Count
End Function

Public Function GetTableStartRow(tbl As ListObject) As Long
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    GetTableStartRow = tbl.DataBodyRange.Row
End Function

Public Function GetTableEndRow(tbl As ListObject) As Long
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    GetTableEndRow = tbl.DataBodyRange.Row + tbl.DataBodyRange.Rows.Count - 1
End Function

Public Function GetTableColumnNames(tbl As ListObject, Optional maxColumns As Long = 0) As String
    ' Returns comma-separated list of column names
    If tbl Is Nothing Then
        GetTableColumnNames = "(no table)"
        Exit Function
    End If
    
    Dim result As String
    Dim col As ListColumn
    Dim Count As Long
    Count = 0
    
    For Each col In tbl.ListColumns
        Count = Count + 1
        If maxColumns > 0 And Count > maxColumns Then
            result = result & " ... (" & tbl.ListColumns.Count & " total)"
            Exit For
        End If
        If Len(result) > 0 Then result = result & ", "
        result = result & col.Name
    Next col
    
    GetTableColumnNames = result
End Function

' ======================================================
' EN/FR HEADER SUPPORT
' ======================================================

Public Function GetFRHeaderEquivalent(enHeader As String) As String
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets("Config")
    
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = wsConfig.ListObjects("ENFRHeaderNamesTable")
    On Error GoTo 0
    
    If tbl Is Nothing Then Exit Function
    
    GetFRHeaderEquivalent = GetTableValueByKey(tbl, _
        "EN - ENMenuSelectionMenuFields Table Header", _
        enHeader, _
        "FR - ENMenuSelectionMenuFields Table Header")
End Function

Public Function GetENHeaderEquivalent(frHeader As String) As String
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets("Config")
    
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = wsConfig.ListObjects("ENFRHeaderNamesTable")
    On Error GoTo 0
    
    If tbl Is Nothing Then Exit Function
    
    GetENHeaderEquivalent = GetTableValueByKey(tbl, _
        "FR - ENMenuSelectionMenuFields Table Header", _
        frHeader, _
        "EN - ENMenuSelectionMenuFields Table Header")
End Function
