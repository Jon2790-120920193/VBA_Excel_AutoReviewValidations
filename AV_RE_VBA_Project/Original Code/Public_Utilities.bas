Attribute VB_Name = "Public_Utilities"

' Retrieves the Key Column Letter from the Configuration sheet
Public Function GetKeyColumnStr(Optional RowNum As Long = 1, Optional wsConfig As Worksheet) As String
    On Error GoTo 0
    If wsConfig Is Nothing Then
        Set wsConfig = ThisWorkbook.Worksheets("Config")
        DebugMessage "Config Sheet Loaded from default '" & wsConfig.Name & "'", "Public_Utilities"
    End If
    Dim colLetter As String
    colLetter = UCase(wsConfig.Range("B5").value)
    
    GetKeyColumnStr = colLetter & CStr(RowNum)

End Function

' Returns Dictionary: Key = ColumnLetter, Value = ValidationFunctionName
Public Function GetValidationColumns(wsConfig As Worksheet) As Object
    ' Public_Utilities Module
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim colLetter As String, funcName As String, Column_Name As String
    i = 6 ' Starting at row 6

    Do While wsConfig.Range("B" & i).value <> ""
        colLetter = Trim(wsConfig.Range("B" & i).value)
        funcName = Trim(wsConfig.Range("C" & i).value)
        Column_Name = Trim(wsConfig.Range("A" & i).value)
        
        AppendUserLog Column_Name & " mapped to column " & colLetter
        If Len(colLetter) > 0 And Len(funcName) > 0 Then
            dict(colLetter) = funcName
        End If
        i = i + 1
    Loop
    
    AppendUserLog "-----------------------------------------------", False
    AppendUserLog "Advanced Autovalidation Configurations Loaded"
    AppendUserLog "-----------------------------------------------", False

    Set GetValidationColumns = dict
End Function

' Returns true if module exists in this workbook
Public Function ModuleExists(funcName As String) As Boolean
    ' Public_Utilities Module
    Dim comp As Object
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If InStr(1, comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines), funcName, vbTextCompare) > 0 Then
            ModuleExists = True
            Exit Function
        End If
    Next comp
    ModuleExists = False
End Function

' --- Checks if a column exists in a table by header name ---
Public Function ColumnExists(tbl As ListObject, headerName As String) As Boolean
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        If StrComp(Trim(col.Name), Trim(headerName), vbTextCompare) = 0 Then
            ColumnExists = True
            Exit Function
        End If
    Next col
End Function

' Calls function based on the AdvFunctionMap Object (Dictionary)
Public Sub ValidateSingleRow(wsData As Worksheet, RowNum As Long, AdvFunctionMap As Object, english As Boolean, FormatMap As Object)
    ' Public_Utilities Module
    Dim colLetter As Variant
    Dim funcName As String
    Dim TargetCell As Range
    Dim mapItem As Object
    Dim AutoValidate As Boolean
    Dim TargetColumnLet As String

    For Each colLetter In AdvFunctionMap.Keys
        Set mapItem = AdvFunctionMap(colLetter)
        funcName = CStr(colLetter)

        ' --- Retrieve AutoValidate flag ---
        AutoValidate = False
        If mapItem.Exists("AutoValidate") Then
            AutoValidate = mapItem("AutoValidate")
        End If
        
        ' --- Retrieve ColumnRef safely ---
        TargetColumnLet = ""
        If mapItem.Exists("ColumnRef") Then
            TargetColumnLet = CStr(mapItem("ColumnRef"))
        End If
        
        If Len(TargetColumnLet) = 0 Then
            Debug.Print "[ValidateSingleRow] WARNING: Missing ColumnRef for " & funcName
            GoTo SkipToNext
        End If
        
        ' --- Skip if AutoValidate = False ---
        If AutoValidate = False Then
            Debug.Print "[ValidateSingleRow] Skipping " & funcName & " (AutoValidate=False)"
            GoTo SkipToNext
        End If

        ' --- Proceed with validation ---
        On Error Resume Next
        Set TargetCell = wsData.Range(TargetColumnLet & RowNum)
        On Error GoTo 0
        
        If Not TargetCell Is Nothing Then
            On Error GoTo ValidationError
            Application.Run funcName, TargetCell, wsData.Name, english, FormatMap, AdvFunctionMap
            On Error GoTo 0
        End If
        
SkipToNext:
    Next colLetter

    AppendUserLog "---Row " & RowNum & " Validation Complete---"
    Exit Sub

ValidationError:
    DebugMessage "[ValidateSingleRow] Public_Utilities Error: Row " & RowNum & " ? Validating column '" & colLetter & "' using function: " & funcName
    AppendUserLog "Error during validation: Row " & RowNum & " ? Validating column '" & colLetter & "' using function: " & funcName
    Resume Next
End Sub


' Returns true if the row is flagged for validation based on the ForceValidationTable found in the Config file (Column IsBuildingColumnValue found in the Column Column specified)
Public Function ShouldValidateRow(RowNum As Long, wsData As Worksheet, Optional validateOnBlankMatch As Boolean = True) As Boolean
    ' Public_Utilities Module
    Dim tbl As ListObject
    Dim wsConfig As Worksheet
    Dim colToCheck As String
    Dim checkValue As String
    Dim i As Long
    Dim buildingValue As String
    Dim TargetCol As Range

    Set wsConfig = ThisWorkbook.Sheets("Config")
    
    On Error Resume Next
    Set tbl = wsConfig.ListObjects("ForceValidationTable")
    On Error GoTo 0

    If tbl Is Nothing Then
        Debug.Print "Public_Utilities Module [Validation] ForceValidationTable not found in Config sheet."
        ShouldValidateRow = False
        Exit Function
    End If

    ' Loop through each row in ForceValidationTable
    For i = 1 To tbl.ListRows.count
        colToCheck = Trim(tbl.DataBodyRange(i, tbl.ListColumns("Column").Index).value)
        buildingValue = Trim(tbl.DataBodyRange(i, tbl.ListColumns("IsBuildingColumnValue").Index).value)
        
        If colToCheck <> "" Then
            On Error Resume Next
            Set TargetCol = wsData.Range(colToCheck & "1")
            On Error GoTo 0

            If Not TargetCol Is Nothing Then
                checkValue = Trim(wsData.Cells(RowNum, TargetCol.Column).value)
                
                ' === 1. If validateOnBlankMatch is enabled, and both are blank, allow ===
                If validateOnBlankMatch Then
                    If buildingValue = "" And checkValue = "" Then
                        ShouldValidateRow = True
                        Exit Function
                    End If
                End If

                ' === 2. Standard case-insensitive match ===
                If buildingValue <> "" And StrComp(buildingValue, checkValue, vbTextCompare) = 0 Then
                    ShouldValidateRow = True
                    Exit Function
                End If
            End If
        End If
    Next i

    ' No match found
    ShouldValidateRow = False
End Function

' Helper: interpret a truth-like cell (TRUE, "TRUE", "Yes", "YES", 1)
Public Function CBoolString(v As Variant) As Boolean
    ' Public_Utilities Module
    On Error Resume Next
    If IsEmpty(v) Then
        CBoolString = False: Exit Function
    End If
    Dim s As String
    s = UCase(Trim(CStr(v)))
    If s = "TRUE" Or s = "YES" Or s = "1" Then
        CBoolString = True
    Else
        CBoolString = False
    End If
End Function

' Helper: get header index from dict by matching list of possible header names
Public Function GetHeaderIndex(headerMap As Object, names As Variant) As Long
    ' Public_Utilities Module
    Dim i As Long
    For i = LBound(names) To UBound(names)
        If headerMap.Exists(names(i)) Then
            GetHeaderIndex = headerMap(names(i))
            Exit Function
        End If
        ' try trimmed and case-insensitive variants:
        Dim k As Variant
        For Each k In headerMap.Keys
            If StrComp(Trim(CStr(k)), Trim(CStr(names(i))), vbTextCompare) = 0 Then
                GetHeaderIndex = headerMap(k)
                Exit Function
            End If
        Next k
    Next i
    GetHeaderIndex = 0
End Function


' Convert column letter number
Public Function ColumnLetterToNumber(colLetter As String) As Long
    ' Public_Utilities Module
    Dim i As Long, result As Long
    For i = 1 To Len(colLetter)
        result = result * 26 + (Asc(UCase(Mid(colLetter, i, 1))) - 64)
    Next i
    ColumnLetterToNumber = result
End Function

Public Function ColumnNumberToLetter(colNum As Long) As String
    Dim div As Long, modNum As Long
    Dim colLetter As String
    div = colNum
    colLetter = ""
    
    Do While div > 0
        modNum = (div - 1) Mod 26
        colLetter = Chr(65 + modNum) & colLetter
        div = (div - modNum - 1) \ 26
    Loop
    
    ColumnNumberToLetter = colLetter
End Function


' SafeTrim used earlier
Public Function SafeTrim(v As Variant) As String
    ' Public_Utilities Module
    If IsError(v) Or IsNull(v) Then SafeTrim = "" Else SafeTrim = Trim(CStr(v))
End Function

' Forced Column Number by Header Name retrieval as Long
Public Function GetColNumberByName(ws As Worksheet, headerName As String) As Long
    ' Public_Utilities Module
    Dim headerRange As Range
    Dim headerCell As Range
    Dim firstRow As Range
    Dim tbl As ListObject

    On Error Resume Next
    ' Try to find the header in any table on the sheet
    For Each tbl In ws.ListObjects
        Set headerRange = tbl.HeaderRowRange.Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole)
        If Not headerRange Is Nothing Then
            GetColNumberByName = headerRange.Column
            Exit Function
        End If
    Next tbl

    ' If not found in tables, try in row 1
    Set firstRow = ws.Rows(1)
    Set headerCell = firstRow.Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole)
    If Not headerCell Is Nothing Then
        GetColNumberByName = headerCell.Column
        Exit Function
    End If

    ' If not found at all
    GetColNumberByName = 0
End Function

'-----------------------------------------------
' Returns a Range on the same sheet as rowRange,
' based on the column header from a table
'-----------------------------------------------
Public Function GetCellFromTableColumnHeader(Table As ListObject, rowRange As Range, ColumnHeader As String) As Range
    Dim colLetter As String
    Dim colIndex As Long
    Dim headerValue As String

    ' Find the column index in the table
    On Error Resume Next
    colIndex = Table.ListColumns(ColumnHeader).Index
    On Error GoTo 0
    
    If colIndex = 0 Then
        Err.Raise vbObjectError + 515, , _
            "Column '" & ColumnHeader & "' not found in table '" & Table.Name & "'"
        Exit Function
    End If

    ' === Get the column letter from the first data row (row 1 of DataBodyRange) ===
    headerValue = Trim(CStr(Table.DataBodyRange.Cells(1, colIndex).value))
    
    If Len(headerValue) = 0 Then
        Err.Raise vbObjectError + 516, , _
            "No column letter found under header '" & ColumnHeader & "' in table '" & Table.Name & "'"
        Exit Function
    End If
    
    ' Return the Range on the same sheet as rowRange
    Set GetCellFromTableColumnHeader = rowRange.Worksheet.Range(headerValue & rowRange.row)
End Function



' Returns UserForm state
Public Function IsUserFormLoaded(formName As String) As Boolean
    ' Public_Utilities Module
    Dim frm As Object
    For Each frm In VBA.UserForms
        If StrComp(frm.Name, formName, vbTextCompare) = 0 Then
            IsUserFormLoaded = True
            Exit Function
        End If
    Next frm
    IsUserFormLoaded = False
End Function

'======================================================
' Function: GetCellByLetter
' Purpose : Safely return a Range object from column letter and row number
' Example : Set c = GetCellByLetter(ws, "AC", 24)
'======================================================
Public Function GetCellByLetter(ws As Worksheet, _
                                colLetter As String, _
                                RowNum As Long) As Range
    Dim colNum As Long
    On Error GoTo ErrHandler
    
    ' Convert column letter to number
    colNum = Range(colLetter & "1").Column
    
    ' Return the range
    Set GetCellByLetter = ws.Cells(RowNum, colNum)
    Exit Function

ErrHandler:
    Debug.Print "?? Invalid column letter: " & colLetter
    Set GetCellByLetter = Nothing
End Function
