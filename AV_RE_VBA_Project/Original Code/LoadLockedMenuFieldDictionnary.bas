Attribute VB_Name = "LoadLockedMenuFieldDictionnary"
Option Explicit

'==========================
' Returns Dictionary: Key = ColumnLetter, Value = ValidationFunctionName
'==========================
Public Function GetDDMValidationColumns(wsConfig As Worksheet) As Object
    Debug.Print "=== Running GetDDMValidationColumns on sheet: " & wsConfig.Name & " ==="
    
    Dim DDMRefTable As ListObject, r As ListRow
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim ReferenceTable As Object
    Dim ReferenceTableName As String
    Dim StartRowIndex As Long
    Dim EndRowMaxIndex As Long
    
    On Error Resume Next
    Set ReferenceTable = GetDDMSheetInfo(wsConfig)
    On Error GoTo 0
    
    If ReferenceTable Is Nothing Then
        Debug.Print "Error: GetDDMSheetInfo returned Nothing."
        Exit Function
    End If
    
    ReferenceTableName = ReferenceTable("ValidationTableName")
    StartRowIndex = ReferenceTable("StartRowIndex")
    EndRowMaxIndex = ReferenceTable("EndRowIndex")
    
    Debug.Print "Validation Table Name: " & ReferenceTableName
    Debug.Print "StartRowIndex: " & StartRowIndex & ", EndRowIndex: " & EndRowMaxIndex
    
    ' Try to get the validation config table
    On Error Resume Next
    Set DDMRefTable = wsConfig.ListObjects("AutoCheckDataValidationTable")
    On Error GoTo 0
    
    If DDMRefTable Is Nothing Then
        Debug.Print "Error: Could not find 'AutoCheckDataValidationTable' in " & wsConfig.Name
        Exit Function
    End If
    
    Debug.Print "Found AutoCheckDataValidationTable with " & DDMRefTable.ListRows.count & " rows."
    
    ' Loop each row of the validation table
    Dim i As Long
    i = 0
    For Each r In DDMRefTable.ListRows
        i = i + 1
        Dim autoCheckVal As String
        autoCheckVal = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns("AutoCheck").Index).value))
        
        If StrComp(autoCheckVal, "TRUE", vbTextCompare) = 0 Then
            Debug.Print "Row " & i & ": AutoCheck = TRUE"
            
            Dim key As String
            key = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns("ReviewSheet Column Letter").Index).value))
            Debug.Print "  > Key (Column Letter): " & key
            
            Dim item As Object
            Set item = CreateObject("Scripting.Dictionary")
            
            item("ReviewLetter") = key
            item("ColumnNameFR") = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns("Column Name (FR)").Index).value))
            item("ColumnNameEN") = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns("Column Name").Index).value))
            item("MenuFieldEN") = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns("MenuField Column (EN)").Index).value))
            item("MenuFieldFR") = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns("MenuField Column (FR)").Index).value))
            item("CommentDropCol") = Trim(CStr(r.Range.Cells(1, DDMRefTable.ListColumns("AutoComment Column").Index).value))
            
            ' Performance Efficiency Step
            Dim NonEmptyRangeEN As Range
            Dim NonEmptyRangeFR As Range
            
            Set NonEmptyRangeEN = GetNonEmptyRangeInColumn(ReferenceTableName, item("MenuFieldEN"), StartRowIndex, EndRowMaxIndex)
            Set NonEmptyRangeFR = GetNonEmptyRangeInColumn(ReferenceTableName, item("MenuFieldFR"), StartRowIndex, EndRowMaxIndex)
            
            ' Store values as Variant arrays (not Set)
            Dim listEN As Variant, listFR As Variant
            
            If Not NonEmptyRangeEN Is Nothing Then
                listEN = GetValuesAsList(NonEmptyRangeEN)
                If IsArray(listEN) Then
                    item("ValidColumnListEN") = listEN
                    Debug.Print "    EN Range found: " & NonEmptyRangeEN.Address & _
                                " (" & UBound(listEN) + 1 & " values)"
                End If
            Else
                Debug.Print "    EN Range is empty or missing."
                item("ValidColumnListEN") = Array()
            End If
            
            If Not NonEmptyRangeFR Is Nothing Then
                listFR = GetValuesAsList(NonEmptyRangeFR)
                If IsArray(listFR) Then
                    item("ValidColumnListFR") = listFR
                    Debug.Print "    FR Range found: " & NonEmptyRangeFR.Address & _
                                " (" & UBound(listFR) + 1 & " values)"
                End If
            Else
                Debug.Print "    FR Range is empty or missing."
                item("ValidColumnListFR") = Array()
            End If
            
            dict.Add key, item
            AppendUserLog "Column: " & CStr(key) & ":" & item("ColumnNameEN") & " ON"

        Else
            Debug.Print "Row " & i & ": AutoCheck = FALSE, skipping."
        End If
    Next r
    
    Debug.Print "Total validated columns found: " & dict.count
    Set GetDDMValidationColumns = dict
    Debug.Print "=== GetDDMValidationColumns completed successfully ==="
End Function


'==========================
' Returns the Menu Locked Field Sheet Name, Start and End Row index.
'==========================
Private Function GetDDMSheetInfo(wsConfig As Worksheet) As Object
    Debug.Print "-- Running GetDDMSheetInfo for sheet: " & wsConfig.Name
    Dim Table As ListObject
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    Set Table = wsConfig.ListObjects("DDMFieldsInfo")
    On Error GoTo 0
    
    If Table Is Nothing Then
        Debug.Print "Error: 'DDMFieldsInfo' table not found in " & wsConfig.Name
        Exit Function
    End If
    
    dict("ValidationTableName") = CStr(Table.DataBodyRange.Cells(1, 2).value)
    dict("StartRowIndex") = CLng(Table.DataBodyRange.Cells(2, 2).value)
    dict("EndRowIndex") = CLng(Table.DataBodyRange.Cells(3, 2).value)
    
    Debug.Print "  TableName: " & dict("ValidationTableName")
    Debug.Print "  StartRow: " & dict("StartRowIndex") & " | EndRow: " & dict("EndRowIndex")
    
    Set GetDDMSheetInfo = dict
End Function


'==========================
' Returns the non-empty range in a specific column
'==========================
Private Function GetNonEmptyRangeInColumn(sheetName As String, _
                                          colLetter As String, _
                                          startRow As Long, _
                                          endRow As Long) As Range
    Debug.Print "  >> GetNonEmptyRangeInColumn: Sheet=" & sheetName & _
                ", Col=" & colLetter & ", Rows " & startRow & "-" & endRow
    
    Dim ws As Worksheet
    Dim checkRange As Range
    Dim lastNonEmptyRow As Long
    Dim cell As Range
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Debug.Print "  !! Error: Worksheet '" & sheetName & "' not found."
        Exit Function
    End If
    If startRow <= 0 Or endRow < startRow Then
        Debug.Print "  !! Invalid row range provided."
        Exit Function
    End If
    
    Set checkRange = ws.Range(colLetter & startRow & ":" & colLetter & endRow)
    lastNonEmptyRow = 0
    
    For Each cell In checkRange.Cells
        If Trim(CStr(cell.value)) <> "" Then lastNonEmptyRow = cell.row
    Next cell
    
    If lastNonEmptyRow = 0 Then
        Debug.Print "  >> No non-empty cells found in column " & colLetter
        Exit Function
    End If
    
    Set GetNonEmptyRangeInColumn = ws.Range(colLetter & startRow & ":" & colLetter & lastNonEmptyRow)
    Debug.Print "  >> NonEmptyRange: " & GetNonEmptyRangeInColumn.Address
End Function


'==========================
' Collects values from a Range into a Variant array of strings
'==========================
Public Function GetValuesAsList(rng As Range) As Variant
    Dim cell As Range
    Dim valuesList() As String
    Dim count As Long
    
    If rng Is Nothing Then
        Debug.Print "GetValuesAsList: Input range is Nothing."
        Exit Function
    End If
    
    For Each cell In rng.Cells
        If Trim(CStr(cell.value)) <> "" Then
            count = count + 1
            ReDim Preserve valuesList(1 To count)
            valuesList(count) = Trim(CStr(cell.value))
        End If
    Next cell
    
    Debug.Print "GetValuesAsList: Collected " & count & " values."
    
    If count > 0 Then
        GetValuesAsList = valuesList
    Else
        GetValuesAsList = Array()
    End If
End Function


