Attribute VB_Name = "AV_DataAccess"
Option Explicit

' ======================================================
' AV_DataAccess
' Centralized table and data access layer
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

Public Function ColumnExistsInTable(ws As Worksheet, tableName As String, columnName As String) As Boolean
    Dim tbl As ListObject
    Dim col As ListColumn
    
    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    If Not tbl Is Nothing Then
        Set col = tbl.ListColumns(columnName)
        ColumnExistsInTable = Not (col Is Nothing)
    End If
    On Error GoTo 0
End Function

' ======================================================
' GET TABLE / COLUMN REFERENCES
' ======================================================

Public Function GetTable(ws As Worksheet, tableName As String) As ListObject
    On Error Resume Next
    Set GetTable = ws.ListObjects(tableName)
    On Error GoTo 0
    
    If GetTable Is Nothing Then
        AV_Core.DebugMessage "Table not found: " & tableName & " in sheet: " & ws.Name, MODULE_NAME
    End If
End Function

Public Function GetTableColumn(ws As Worksheet, tableName As String, columnName As String) As Range
    Dim tbl As ListObject
    Set tbl = GetTable(ws, tableName)
    
    If tbl Is Nothing Then Exit Function
    
    On Error Resume Next
    Set GetTableColumn = tbl.ListColumns(columnName).DataBodyRange
    On Error GoTo 0
    
    If GetTableColumn Is Nothing Then
        AV_Core.DebugMessage "Column not found: " & columnName & " in table: " & tableName, MODULE_NAME
    End If
End Function

' ======================================================
' GET SINGLE VALUES FROM TABLES
' ======================================================

Public Function GetTableValue(ws As Worksheet, _
                              tableName As String, _
                              rowIndex As Long, _
                              columnName As String) As Variant
    Dim tbl As ListObject
    Set tbl = GetTable(ws, tableName)
    
    If tbl Is Nothing Then
        GetTableValue = Empty
        Exit Function
    End If
    
    On Error Resume Next
    GetTableValue = tbl.ListColumns(columnName).DataBodyRange(rowIndex).Value
    On Error GoTo 0
End Function

Public Function GetTableValueByKey(ws As Worksheet, _
                                   tableName As String, _
                                   keyColumn As String, _
                                   keyValue As Variant, _
                                   valueColumn As String) As Variant
    Dim tbl As ListObject
    Set tbl = GetTable(ws, tableName)
    
    If tbl Is Nothing Then Exit Function
    
    Dim r As ListRow
    Dim keyColIndex As Long, valColIndex As Long
    
    On Error Resume Next
    keyColIndex = tbl.ListColumns(keyColumn).Index
    valColIndex = tbl.ListColumns(valueColumn).Index
    On Error GoTo 0
    
    If keyColIndex = 0 Or valColIndex = 0 Then Exit Function
    
    For Each r In tbl.ListRows
        If r.Range.Cells(1, keyColIndex).Value = keyValue Then
            GetTableValueByKey = r.Range.Cells(1, valColIndex).Value
            Exit Function
        End If
    Next r
End Function

' ======================================================
' GET ENTIRE ROWS AS DICTIONARIES
' ======================================================

Public Function GetTableRow(ws As Worksheet, _
                            tableName As String, _
                            keyColumn As String, _
                            keyValue As Variant) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim tbl As ListObject
    Set tbl = GetTable(ws, tableName)
    
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
        If r.Range.Cells(1, keyColIndex).Value = keyValue Then
            For Each col In tbl.ListColumns
                dict(col.Name) = r.Range.Cells(1, col.Index).Value
            Next col
            Set GetTableRow = dict
            Exit Function
        End If
    Next r
    
    Set GetTableRow = dict
End Function

Public Function GetTableRowByIndex(ws As Worksheet, _
                                   tableName As String, _
                                   rowIndex As Long) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim tbl As ListObject
    Set tbl = GetTable(ws, tableName)
    
    If tbl Is Nothing Then
        Set GetTableRowByIndex = dict
        Exit Function
    End If
    
    If rowIndex < 1 Or rowIndex > tbl.ListRows.Count Then
        Set GetTableRowByIndex = dict
        Exit Function
    End If
    
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        dict(col.Name) = tbl.DataBodyRange(rowIndex, col.Index).Value
    Next col
    
    Set GetTableRowByIndex = dict
End Function

' ======================================================
' CHECK VALUE EXISTENCE
' ======================================================

Public Function TableContainsValue(ws As Worksheet, _
                                   tableName As String, _
                                   columnName As String, _
                                   searchValue As Variant, _
                                   Optional caseSensitive As Boolean = False) As Boolean
    Dim colRange As Range
    Set colRange = GetTableColumn(ws, tableName, columnName)
    
    If colRange Is Nothing Then Exit Function
    
    Dim cell As Range
    For Each cell In colRange.Cells
        If caseSensitive Then
            If cell.Value = searchValue Then
                TableContainsValue = True
                Exit Function
            End If
        Else
            If StrComp(CStr(cell.Value), CStr(searchValue), vbTextCompare) = 0 Then
                TableContainsValue = True
                Exit Function
            End If
        End If
    Next cell
End Function

' ======================================================
' GET ALL VALUES FROM COLUMN (AS ARRAY)
' ======================================================

Public Function GetColumnValues(ws As Worksheet, _
                                tableName As String, _
                                columnName As String) As Variant
    Dim colRange As Range
    Set colRange = GetTableColumn(ws, tableName, columnName)
    
    If colRange Is Nothing Then
        GetColumnValues = Array()
        Exit Function
    End If
    
    ' Return as 1D array
    Dim values() As Variant
    ReDim values(1 To colRange.Rows.Count)
    
    Dim i As Long
    For i = 1 To colRange.Rows.Count
        values(i) = colRange.Cells(i, 1).Value
    Next i
    
    GetColumnValues = values
End Function

Public Function GetColumnValuesFiltered(ws As Worksheet, _
                                       tableName As String, _
                                       columnName As String, _
                                       filterColumn As String, _
                                       filterValue As Variant) As Variant
    Dim tbl As ListObject
    Set tbl = GetTable(ws, tableName)
    
    If tbl Is Nothing Then
        GetColumnValuesFiltered = Array()
        Exit Function
    End If
    
    Dim valueColIndex As Long, filterColIndex As Long
    On Error Resume Next
    valueColIndex = tbl.ListColumns(columnName).Index
    filterColIndex = tbl.ListColumns(filterColumn).Index
    On Error GoTo 0
    
    If valueColIndex = 0 Or filterColIndex = 0 Then
        GetColumnValuesFiltered = Array()
        Exit Function
    End If
    
    Dim results As Collection
    Set results = New Collection
    
    Dim r As ListRow
    For Each r In tbl.ListRows
        If r.Range.Cells(1, filterColIndex).Value = filterValue Then
            On Error Resume Next
            results.Add r.Range.Cells(1, valueColIndex).Value
            On Error GoTo 0
        End If
    Next r
    
    If results.Count = 0 Then
        GetColumnValuesFiltered = Array()
        Exit Function
    End If
    
    Dim arr() As Variant
    ReDim arr(1 To results.Count)
    Dim i As Long
    For i = 1 To results.Count
        arr(i) = results(i)
    Next i
    
    GetColumnValuesFiltered = arr
End Function

' ======================================================
' FIND HEADER IN TARGET SHEET (EN/FR SUPPORT)
' ======================================================

Public Function FindColumnByHeader(ws As Worksheet, _
                                   tableName As String, _
                                   headerName As String, _
                                   Optional tryFRHeader As Boolean = False) As Range
    Dim tbl As ListObject
    Set tbl = GetTable(ws, tableName)
    
    If tbl Is Nothing Then Exit Function
    
    ' Try exact match first
    On Error Resume Next
    Set FindColumnByHeader = tbl.ListColumns(headerName).DataBodyRange
    On Error GoTo 0
    
    ' If found, return
    If Not FindColumnByHeader Is Nothing Then Exit Function
    
    ' If tryFRHeader enabled, try to find FR equivalent
    If tryFRHeader Then
        Dim frHeader As String
        frHeader = GetFRHeaderEquivalent(headerName)
        
        If frHeader <> "" Then
            On Error Resume Next
            Set FindColumnByHeader = tbl.ListColumns(frHeader).DataBodyRange
            On Error GoTo 0
        End If
    End If
End Function

Public Function GetFRHeaderEquivalent(enHeader As String) As String
    ' Look up FR header from ENFRHeaderNamesTable
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(AV_Constants.CONFIG_SHEET_NAME)
    
    GetFRHeaderEquivalent = GetTableValueByKey( _
        wsConfig, _
        AV_Constants.TBL_ENFR_HEADER_MAPPING, _
        AV_Constants.COL_ENFR_EN_HEADER, _
        enHeader, _
        AV_Constants.COL_ENFR_FR_HEADER _
    )
End Function

Public Function GetENHeaderEquivalent(frHeader As String) As String
    ' Look up EN header from ENFRHeaderNamesTable
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(AV_Constants.CONFIG_SHEET_NAME)
    
    GetENHeaderEquivalent = GetTableValueByKey( _
        wsConfig, _
        AV_Constants.TBL_ENFR_HEADER_MAPPING, _
        AV_Constants.COL_ENFR_FR_HEADER, _
        frHeader, _
        AV_Constants.COL_ENFR_EN_HEADER _
    )
End Function

' ======================================================
' GET CELL FROM TABLE BY HEADER (SPECIFIC ROW)
' ======================================================

Public Function GetCellByHeader(ws As Worksheet, _
                               tableName As String, _
                               rowIndex As Long, _
                               headerName As String) As Range
    Dim tbl As ListObject
    Set tbl = GetTable(ws, tableName)
    
    If tbl Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > tbl.ListRows.Count Then Exit Function
    
    On Error Resume Next
    Set GetCellByHeader = tbl.ListColumns(headerName).DataBodyRange(rowIndex)
    On Error GoTo 0
End Function
