Attribute VB_Name = "AV_Engine"
Option Explicit

' ======================================================
' AV_Engine.bas
' Validation orchestration & execution engine
' VERSION: 2.4 - With version tracking & enhanced diagnostics
' FULLY TABLE-BASED - NO CELL REFERENCES
' ======================================================

Private Const MODULE_NAME As String = "AV_Engine"

' PUBLIC VERSION CONSTANT - Check this to verify correct module is loaded
Public Const MODULE_VERSION As String = "2.4"

' Module-level references for current validation context
Private mCurrentTargetTable As ListObject
Private mCurrentTargetSheet As Worksheet
Private mCurrentTableName As String

' ======================================================
' PUBLIC PROPERTIES - CURRENT TABLE ACCESS
' ======================================================

Public Property Get CurrentTargetTable() As ListObject
    Set CurrentTargetTable = mCurrentTargetTable
End Property

Public Property Set CurrentTargetTable(tbl As ListObject)
    Set mCurrentTargetTable = tbl
End Property

Public Property Get CurrentTargetSheet() As Worksheet
    Set CurrentTargetSheet = mCurrentTargetSheet
End Property

Public Property Get CurrentTableName() As String
    CurrentTableName = mCurrentTableName
End Property

' ======================================================
' PUBLIC ENTRY POINTS
' ======================================================

Public Sub RunFullValidation(Optional ByVal sheetName As String = "", Optional ByVal english As Boolean = True)
    RunFullValidationMaster sheetName, english
End Sub

' ======================================================
' MAIN VALIDATION EXECUTION
' Fully table-based - derives ALL settings from tables
' Version 2.4: Enhanced diagnostics and version tracking
' ======================================================
Public Sub RunFullValidationMaster(Optional ByVal sheetName As String = "", Optional ByVal english As Boolean = True)

    Dim wsConfig As Worksheet
    Dim wsTarget As Worksheet
    Dim targetTableName As String
    Dim rowNum As Long, i As Long
    Dim keyRows() As Long
    Dim keyCount As Long
    
    ' Table-derived ranges (NO CELL REFERENCES)
    Dim tableStartRow As Long
    Dim tableEndRow As Long
    Dim keyColNum As Long
    Dim keyColumnHeader As String

    Dim AdvFunctionMap As Object
    Dim FormatMap As Object
    Dim colMetaDict As Object
    Dim validateSmartFuncColMap As Object
    Dim colReviewedColumnList As Collection

    On Error GoTo ErrHandler

    ' Initialize UI / logging
    AV_UI.ShowValidationTrackerForm
    AV_UI.AppendUserLog "=========================================="
    AV_UI.AppendUserLog "AV_Engine v" & MODULE_VERSION & " - Table-Based Validation"
    AV_UI.AppendUserLog "=========================================="
    
    AV_Core.InitDebugFlags
    AV_Core.DebugMessage "AV_Engine.MODULE_VERSION = " & MODULE_VERSION, MODULE_NAME

    ' Cancel / timeout flags
    AV_Core.ValidationStartTime = Timer
    AV_Core.ValidationCancelTimeout = 10000
    AV_Core.ValidationCancelFlag = False

    AV_UI.AppendUserLog "Timeout: " & AV_Core.ValidationCancelTimeout & " seconds"

    ' Load configuration sheet
    Set wsConfig = ThisWorkbook.Sheets("Config")

    ' ======================================================
    ' GET TARGET FROM ValidationTargets TABLE (not cells!)
    ' ======================================================
    AV_UI.AppendUserLog ""
    AV_UI.AppendUserLog "--- CONFIGURATION ---"
    
    Dim validationTargets As ListObject
    On Error Resume Next
    Set validationTargets = wsConfig.ListObjects("ValidationTargets")
    On Error GoTo 0
    
    If validationTargets Is Nothing Then
        AV_UI.AppendUserLog "ERROR: ValidationTargets table not found"
        AV_UI.AppendUserLog "Create this table in Config sheet with columns:"
        AV_UI.AppendUserLog "  TableName, Enabled, Mode, Key Column (Header Name)"
        GoTo Cleanup
    End If
    
    AV_Core.DebugMessage "ValidationTargets table found: " & validationTargets.ListRows.Count & " rows", MODULE_NAME
    
    ' Find first enabled target (or use passed sheetName to find specific one)
    Dim targetRow As ListRow
    Dim foundTarget As Boolean
    foundTarget = False
    
    For Each targetRow In validationTargets.ListRows
        Dim isEnabled As String
        isEnabled = UCase(Trim(CStr(targetRow.Range.Cells(1, validationTargets.ListColumns("Enabled").Index).Value)))
        
        If isEnabled = "TRUE" Then
            targetTableName = Trim(CStr(targetRow.Range.Cells(1, validationTargets.ListColumns("TableName").Index).Value))
            
            ' If sheetName passed, only use that specific one
            If Len(sheetName) > 0 Then
                If StrComp(targetTableName, sheetName, vbTextCompare) = 0 Then
                    foundTarget = True
                    Exit For
                End If
            Else
                ' Use first enabled target
                foundTarget = True
                Exit For
            End If
        End If
    Next targetRow
    
    If Not foundTarget Then
        AV_UI.AppendUserLog "ERROR: No enabled target found in ValidationTargets"
        GoTo Cleanup
    End If
    
    AV_UI.AppendUserLog "Target table: " & targetTableName
    mCurrentTableName = targetTableName
    
    ' ======================================================
    ' FIND THE TARGET TABLE (ListObject) BY NAME
    ' Search all worksheets
    ' ======================================================
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set tbl = ws.ListObjects(targetTableName)
        On Error GoTo 0
        
        If Not tbl Is Nothing Then
            Set mCurrentTargetTable = tbl
            Set mCurrentTargetSheet = ws
            Set wsTarget = ws
            Exit For
        End If
    Next ws
    
    If mCurrentTargetTable Is Nothing Then
        AV_UI.AppendUserLog "ERROR: Table '" & targetTableName & "' not found"
        AV_UI.AppendUserLog "Check table exists and name matches exactly"
        GoTo Cleanup
    End If
    
    AV_UI.AppendUserLog "Found on sheet: " & wsTarget.Name
    
    ' ======================================================
    ' AUTO-DETECT TABLE RANGE FROM ListObject
    ' This is the KEY fix - use actual table bounds
    ' ======================================================
    If mCurrentTargetTable.DataBodyRange Is Nothing Then
        AV_UI.AppendUserLog "ERROR: Table '" & targetTableName & "' has no data rows"
        GoTo Cleanup
    End If
    
    tableStartRow = mCurrentTargetTable.DataBodyRange.Row
    tableEndRow = tableStartRow + mCurrentTargetTable.DataBodyRange.Rows.Count - 1
    
    AV_UI.AppendUserLog "Table range: Row " & tableStartRow & " to " & tableEndRow
    AV_UI.AppendUserLog "Total rows: " & mCurrentTargetTable.DataBodyRange.Rows.Count
    
    AV_Core.DebugMessage "TABLE RANGE DETECTION:", MODULE_NAME
    AV_Core.DebugMessage "  tableStartRow = " & tableStartRow, MODULE_NAME
    AV_Core.DebugMessage "  tableEndRow = " & tableEndRow, MODULE_NAME
    AV_Core.DebugMessage "  (derived from ListObject.DataBodyRange)", MODULE_NAME
    
    ' ======================================================
    ' GET KEY COLUMN FROM ValidationTargets
    ' ======================================================
    On Error Resume Next
    Dim keyColIdx As Long
    keyColIdx = validationTargets.ListColumns("Key Column (Header Name)").Index
    If keyColIdx > 0 Then
        keyColumnHeader = Trim(CStr(targetRow.Range.Cells(1, keyColIdx).Value))
    End If
    On Error GoTo 0
    
    ' If not specified, use first column of target table
    If Len(keyColumnHeader) = 0 Then
        keyColumnHeader = mCurrentTargetTable.ListColumns(1).Name
        AV_UI.AppendUserLog "Key column (default): " & keyColumnHeader
    Else
        AV_UI.AppendUserLog "Key column: " & keyColumnHeader
    End If
    
    ' Find key column in the table
    Dim keyCol As ListColumn
    On Error Resume Next
    Set keyCol = mCurrentTargetTable.ListColumns(keyColumnHeader)
    On Error GoTo 0
    
    If keyCol Is Nothing Then
        AV_UI.AppendUserLog "ERROR: Key column '" & keyColumnHeader & "' not found"
        AV_UI.AppendUserLog "Available: " & GetTableColumnList(mCurrentTargetTable)
        GoTo Cleanup
    End If
    
    keyColNum = keyCol.DataBodyRange.Column

    ' ======================================================
    ' LOAD MAPPINGS (cached)
    ' ======================================================
    AV_UI.AppendUserLog ""
    AV_UI.AppendUserLog "--- LOADING MAPPINGS ---"
    
    Set AdvFunctionMap = AV_Core.GetAutoValidationMap(wsConfig)
    Set FormatMap = AV_Format.LoadFormatMap(wsConfig)
    Set colMetaDict = AV_Core.GetDDMValidationColumns(wsConfig)
    Set validateSmartFuncColMap = AV_Core.GetValidationColumns(wsConfig)

    If AdvFunctionMap Is Nothing Or AdvFunctionMap.Count = 0 Then
        AV_UI.AppendUserLog "ERROR: No validation functions mapped"
        GoTo Cleanup
    End If
    
    AV_UI.AppendUserLog "Validation functions: " & AdvFunctionMap.Count
    AV_UI.AppendUserLog "Format definitions: " & FormatMap.Count
    
    ' ======================================================
    ' DIAGNOSTIC: Compare mapped headers vs actual headers
    ' ======================================================
    AV_UI.AppendUserLog ""
    AV_UI.AppendUserLog "--- HEADER MAPPING CHECK ---"
    
    Dim mapKey As Variant
    Dim mappedHeader As String
    Dim headerFound As Boolean
    Dim missingHeaders As String
    Dim foundHeaders As Long
    Dim missingCount As Long
    
    missingHeaders = ""
    foundHeaders = 0
    missingCount = 0
    
    For Each mapKey In AdvFunctionMap.Keys
        Dim mapItem As Object
        Set mapItem = AdvFunctionMap(mapKey)
        mappedHeader = mapItem("ColumnRef")
        
        ' Check if this header exists in the target table
        headerFound = False
        Dim checkCol As ListColumn
        On Error Resume Next
        Set checkCol = mCurrentTargetTable.ListColumns(mappedHeader)
        headerFound = Not (checkCol Is Nothing)
        On Error GoTo 0
        
        If headerFound Then
            AV_Core.DebugMessage "OK: " & mapKey & " -> '" & mappedHeader & "'", MODULE_NAME
            foundHeaders = foundHeaders + 1
        Else
            AV_Core.DebugMessage "MISSING: " & mapKey & " -> '" & mappedHeader & "'", MODULE_NAME
            If Len(missingHeaders) > 0 Then missingHeaders = missingHeaders & ", "
            missingHeaders = missingHeaders & mappedHeader
            missingCount = missingCount + 1
        End If
    Next mapKey
    
    AV_UI.AppendUserLog "Headers found: " & foundHeaders & "/" & AdvFunctionMap.Count
    
    If Len(missingHeaders) > 0 Then
        AV_UI.AppendUserLog "MISSING: " & missingHeaders
        AV_UI.AppendUserLog "(Update AutoValidationCommentPrefixMappingTable)"
    End If

    AV_UI.AppendUserLog ""
    AV_UI.AppendUserLog "--- BUILDING ROW LIST ---"
    AV_UI.SetAutoValidationInitialized True

    ' ======================================================
    ' BUILD ROW LIST FROM TABLE (auto-detected range)
    ' CRITICAL: Only scan within table bounds
    ' ======================================================
    ReDim keyRows(1 To mCurrentTargetTable.DataBodyRange.Rows.Count)
    keyCount = 0

    AV_Core.DebugMessage "Scanning rows " & tableStartRow & " to " & tableEndRow & " for key values", MODULE_NAME
    
    For rowNum = tableStartRow To tableEndRow
        If Trim(CStr(wsTarget.Cells(rowNum, keyColNum).Value)) <> "" Then
            keyCount = keyCount + 1
            keyRows(keyCount) = rowNum
        End If
    Next rowNum

    If keyCount = 0 Then
        AV_UI.AppendUserLog "ERROR: No rows with key values found"
        AV_UI.AppendUserLog "Check column '" & keyColumnHeader & "' has data"
        GoTo Cleanup
    End If

    ReDim Preserve keyRows(1 To keyCount)

    AV_Core.DebugMessage "Rows with key values: " & keyCount, MODULE_NAME
    AV_UI.AppendUserLog "Rows to validate: " & keyCount
    AV_UI.AppendUserLog "First row: " & keyRows(1) & " | Last row: " & keyRows(keyCount)
    
    AV_UI.AppendUserLog ""
    AV_UI.AppendUserLog "--- VALIDATING ROWS ---"

    ' ======================================================
    ' MAIN ROW LOOP
    ' ======================================================
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim validatedCount As Long
    Dim skippedCount As Long
    validatedCount = 0
    skippedCount = 0
    
    For i = LBound(keyRows) To UBound(keyRows)
        rowNum = keyRows(i)

        If i Mod 10 = 0 Then DoEvents

        If AV_Core.ValidationCancelFlag Then
            AV_UI.AppendUserLog "Cancelled by user"
            GoTo Cleanup
        End If

        If AV_Core.ValidationTimeoutReached() Then
            AV_UI.AppendUserLog "Timeout reached"
            GoTo Cleanup
        End If

        If AV_Core.ShouldValidateRow(rowNum, wsTarget, True) Then
            ValidateSingleRow wsTarget, rowNum, AdvFunctionMap, english, FormatMap, mCurrentTargetTable
            validatedCount = validatedCount + 1
        Else
            skippedCount = skippedCount + 1
        End If
        
        ' Progress update
        If i Mod 50 = 0 Then
            AV_UI.AppendUserLog "Progress: " & i & " / " & keyCount
        End If
    Next i

    AV_UI.AppendUserLog ""
    AV_UI.AppendUserLog "--- ADVANCED VALIDATION COMPLETE ---"
    AV_UI.AppendUserLog "Validated: " & validatedCount & " | Skipped: " & skippedCount
    AV_UI.SetAdvancedValidationCompleted True

    ' ======================================================
    ' POST-PASS: SIMPLE DATA VALIDATION
    ' ======================================================
    AV_UI.AppendUserLog ""
    AV_UI.AppendUserLog "--- SIMPLE VALIDATION ---"
    AV_Core.DebugMessage "Starting RunAutoCheckDataValidation()", MODULE_NAME

    Set colReviewedColumnList = BuildCollectionOfColumnHeaders(colMetaDict, validateSmartFuncColMap, mCurrentTargetTable)

    RunAutoCheckDataValidation wsConfig, wsTarget, keyRows, keyColNum, english, FormatMap, colMetaDict, colReviewedColumnList

    AV_Core.DebugMessage "RunAutoCheckDataValidation() completed", MODULE_NAME

Cleanup:
    Set mCurrentTargetTable = Nothing
    Set mCurrentTargetSheet = Nothing
    mCurrentTableName = ""
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    AV_UI.AppendUserLog ""
    AV_UI.AppendUserLog "=========================================="
    AV_UI.AppendUserLog "VALIDATION COMPLETE"
    AV_UI.AppendUserLog "=========================================="
    
    AV_Core.DebugMessage "RunFullValidationMaster completed at " & Now, MODULE_NAME
    Exit Sub

ErrHandler:
    AV_UI.AppendUserLog ""
    AV_UI.AppendUserLog "ERROR in RunFullValidationMaster"
    AV_UI.AppendUserLog "Error #" & Err.Number & ": " & Err.Description
    AV_UI.BringFormToFront ValidationTrackerForm
    Resume Cleanup
End Sub


' ======================================================
' VALIDATE SINGLE ROW
' Uses table header lookup, not column letters
' ======================================================
Public Sub ValidateSingleRow(wsData As Worksheet, rowNum As Long, AdvFunctionMap As Object, _
                             english As Boolean, FormatMap As Object, _
                             Optional targetTable As ListObject = Nothing)
    Dim funcKey As Variant
    Dim funcName As String
    Dim TargetCell As Range
    Dim mapItem As Object
    Dim AutoValidate As Boolean
    Dim ColumnHeader As String

    ' Use passed table or module-level reference
    If targetTable Is Nothing Then
        Set targetTable = mCurrentTargetTable
    End If
    
    If targetTable Is Nothing Then
        AV_Core.DebugMessage "ValidateSingleRow: No target table available", MODULE_NAME
        Exit Sub
    End If

    For Each funcKey In AdvFunctionMap.Keys
        Set mapItem = AdvFunctionMap(funcKey)
        funcName = CStr(funcKey)

        ' Check AutoValidate flag
        AutoValidate = False
        If mapItem.Exists("AutoValidate") Then
            AutoValidate = mapItem("AutoValidate")
        End If
        
        If Not AutoValidate Then
            GoTo SkipToNext
        End If
        
        ' Get column header from mapping
        ColumnHeader = ""
        If mapItem.Exists("ColumnRef") Then
            ColumnHeader = CStr(mapItem("ColumnRef"))
        End If
        
        If Len(ColumnHeader) = 0 Then
            AV_Core.DebugMessage "Row " & rowNum & ": " & funcName & " - missing ColumnRef", MODULE_NAME
            GoTo SkipToNext
        End If

        ' Get target cell using table header lookup
        Set TargetCell = GetCellByTableHeader(targetTable, ColumnHeader, rowNum)
        
        If TargetCell Is Nothing Then
            ' Only log once per function, not per row
            GoTo SkipToNext
        End If
        
        On Error GoTo ValidationError
        Application.Run funcName, TargetCell, wsData.Name, english, FormatMap, AdvFunctionMap
        On Error GoTo 0
        
SkipToNext:
    Next funcKey

    Exit Sub

ValidationError:
    AV_Core.DebugMessage "Validation Error: Row " & rowNum & " - " & funcName & " - " & Err.Description, MODULE_NAME
    Resume SkipToNext
End Sub


' ======================================================
' GET CELL BY TABLE HEADER
' Core function: finds cell at row for given column header
' ======================================================
Private Function GetCellByTableHeader(tbl As ListObject, headerName As String, rowNum As Long) As Range
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    
    ' Find the column by header name
    Dim col As ListColumn
    On Error Resume Next
    Set col = tbl.ListColumns(headerName)
    On Error GoTo 0
    
    If col Is Nothing Then
        ' Don't log here - creates too much noise for missing columns
        Exit Function
    End If
    
    ' Check if rowNum is within table range
    Dim tableStartRow As Long
    Dim tableEndRow As Long
    tableStartRow = tbl.DataBodyRange.Row
    tableEndRow = tableStartRow + tbl.DataBodyRange.Rows.Count - 1
    
    If rowNum < tableStartRow Or rowNum > tableEndRow Then
        AV_Core.DebugMessage "GetCellByTableHeader: Row " & rowNum & " outside range (" & tableStartRow & "-" & tableEndRow & ")", MODULE_NAME
        Exit Function
    End If
    
    ' Return the cell at intersection of column and row
    Set GetCellByTableHeader = tbl.Parent.Cells(rowNum, col.DataBodyRange.Column)
End Function


' ======================================================
' GET TABLE COLUMN LIST (for diagnostics)
' ======================================================
Private Function GetTableColumnList(tbl As ListObject) As String
    If tbl Is Nothing Then
        GetTableColumnList = "(no table)"
        Exit Function
    End If
    
    Dim result As String
    Dim col As ListColumn
    Dim Count As Long
    Count = 0
    
    For Each col In tbl.ListColumns
        Count = Count + 1
        If Count <= 10 Then
            If Len(result) > 0 Then result = result & ", "
            result = result & col.Name
        End If
    Next col
    
    If tbl.ListColumns.Count > 10 Then
        result = result & " ... (" & tbl.ListColumns.Count & " total)"
    End If
    
    GetTableColumnList = result
End Function


' ======================================================
' BUILD COLLECTION OF COLUMN HEADERS
' Updated to work with headers instead of letters
' ======================================================
Public Function BuildCollectionOfColumnHeaders(colMetaDict As Object, validateSmartFuncColMap As Object, tbl As ListObject) As Collection
    Dim colHeaders As New Collection
    Dim key As Variant
    Dim existsDict As Object
    Set existsDict = CreateObject("Scripting.Dictionary")
    
    If tbl Is Nothing Then
        Set BuildCollectionOfColumnHeaders = colHeaders
        Exit Function
    End If
    
    ' Add all keys from colMetaDict that exist in the target table
    If Not colMetaDict Is Nothing Then
        For Each key In colMetaDict.Keys
            Dim headerName As String
            headerName = CStr(key)
            
            ' Check if this header exists in the table
            Dim col As ListColumn
            On Error Resume Next
            Set col = tbl.ListColumns(headerName)
            On Error GoTo 0
            
            If Not col Is Nothing Then
                If Not existsDict.Exists(headerName) Then
                    existsDict(headerName) = True
                    colHeaders.Add headerName
                End If
            End If
        Next key
    End If
    
    Set BuildCollectionOfColumnHeaders = colHeaders
End Function


' ======================================================
' RUN AUTO CHECK DATA VALIDATION
' ======================================================
Public Sub RunAutoCheckDataValidation(wsConfig As Worksheet, _
                                     wsTarget As Worksheet, _
                                     keyRows() As Long, _
                                     keyColNum As Long, _
                                     english As Boolean, _
                                     Optional FormatMap As Object, _
                                     Optional colMetaDict As Object, _
                                     Optional RevColHeaderList As Collection)

    On Error GoTo ErrHandler

    Dim meta As Object
    Dim colKey As Variant
    Dim i As Long, rowNum As Long
    Dim cellValue As String
    Dim found As Boolean
    Dim progressCount As Long
    Dim totalValid As Long
    
    Dim dropColMsgs As Object
    Dim dropColKey As Variant
    Dim cMsgErrorType As String
    Dim cMsg As Variant
    Dim msgArr(1 To 3) As Variant
    Dim DCMsgTxt As String

    If wsConfig Is Nothing Or wsTarget Is Nothing Then Exit Sub
    If LBound(keyRows) > UBound(keyRows) Then Exit Sub
    If colMetaDict Is Nothing Then Exit Sub

    totalValid = UBound(keyRows) - LBound(keyRows) + 1
    AV_Core.DebugMessage "Simple validation on " & totalValid & " rows", MODULE_NAME

    For i = LBound(keyRows) To UBound(keyRows)
        rowNum = keyRows(i)
        
        If rowNum <= 0 Or rowNum > wsTarget.Rows.Count Then GoTo SkipRow
        
        Set dropColMsgs = CreateObject("Scripting.Dictionary")
        
        For Each colKey In colMetaDict.Keys
            Set meta = colMetaDict(colKey)
            
            If Not meta.Exists("ReviewLetter") Then GoTo SkipCol
            If Not meta.Exists("ValidColumnListEN") Then meta("ValidColumnListEN") = Array()
            If Not meta.Exists("ValidColumnListFR") Then meta("ValidColumnListFR") = Array()
            If Not meta.Exists("ColumnNameEN") Then meta("ColumnNameEN") = ""
            If Not meta.Exists("ColumnNameFR") Then meta("ColumnNameFR") = ""
            If Not meta.Exists("CommentDropCol") Then meta("CommentDropCol") = ""
            
            ' Get cell value using table lookup
            Dim dataCell As Range
            Set dataCell = GetCellByTableHeader(mCurrentTargetTable, CStr(meta("ReviewLetter")), rowNum)
            
            cellValue = ""
            If Not dataCell Is Nothing Then
                cellValue = Trim(CStr(dataCell.Value))
            End If
            
            If Len(cellValue) = 0 Then GoTo SkipCol

            found = False
            If IsArray(meta("ValidColumnListEN")) Then found = ExistsInArray(meta("ValidColumnListEN"), cellValue)
            If Not found And IsArray(meta("ValidColumnListFR")) Then found = ExistsInArray(meta("ValidColumnListFR"), cellValue)

            If Not found Then
                Dim errorMsg As String
                If english Then
                    errorMsg = meta("ColumnNameEN") & " - Invalid value '" & cellValue & "'"
                Else
                    errorMsg = meta("ColumnNameFR") & " - Valeur invalide '" & cellValue & "'"
                End If

                If Not dropColMsgs.Exists(meta("CommentDropCol")) Then
                    Set dropColMsgs(meta("CommentDropCol")) = CreateObject("Scripting.Dictionary")
                End If
                
                msgArr(1) = meta("ReviewLetter")
                msgArr(2) = errorMsg
                msgArr(3) = "Error"
                dropColMsgs(meta("CommentDropCol")).Add dropColMsgs(meta("CommentDropCol")).Count + 1, msgArr
            Else
                If Not dataCell Is Nothing Then
                    If AV_Format.getFormatType(dataCell, FormatMap) = "Error" Then
                        If Not dropColMsgs.Exists(meta("CommentDropCol")) Then
                            Set dropColMsgs(meta("CommentDropCol")) = CreateObject("Scripting.Dictionary")
                        End If
                        
                        msgArr(1) = meta("ReviewLetter")
                        msgArr(2) = ""
                        msgArr(3) = "Default"
                        dropColMsgs(meta("CommentDropCol")).Add dropColMsgs(meta("CommentDropCol")).Count + 1, msgArr
                    End If
                End If
            End If

SkipCol:
        Next colKey
                
        For Each dropColKey In dropColMsgs.Keys
            For Each cMsg In dropColMsgs(dropColKey).Items
                DCMsgTxt = cMsg(2)
                cMsgErrorType = CStr(cMsg(3))
                AV_Format.WriteSystemTagToDropColumn wsTarget, CStr(dropColKey), rowNum, CStr(cMsg(1)), DCMsgTxt, cMsgErrorType, FormatMap
            Next cMsg
        Next dropColKey
        
        ' Format key cell based on row validation results
        Dim rowRange As Range
        Set rowRange = BuildRowRangeFromHeaders(wsTarget, RevColHeaderList, rowNum)
        
        If Not rowRange Is Nothing Then
            AV_Format.FormatKeyCell rowRange, FormatMap
        End If
        
        progressCount = progressCount + 1
        If progressCount Mod 10 = 0 Then DoEvents
        If progressCount Mod 100 = 0 Then
            AV_UI.AppendUserLog "Simple validation: " & progressCount & " / " & totalValid
        End If

SkipRow:
    Next i

    AV_Core.DebugMessage "Simple validation complete: " & progressCount & " rows", MODULE_NAME
    AV_UI.AppendUserLog "Simple validation complete: " & progressCount & " rows"
    AV_UI.SetLMenuValCompletedCB True

    Exit Sub

ErrHandler:
    AV_Core.DebugMessage "RunAutoCheckDataValidation ERROR: " & Err.Number & " - " & Err.Description, MODULE_NAME
    AV_UI.AppendUserLog "Simple validation ERROR: " & Err.Description
End Sub


' ======================================================
' BUILD ROW RANGE FROM HEADERS
' ======================================================
Public Function BuildRowRangeFromHeaders(ws As Worksheet, colHeaders As Collection, rowNum As Long) As Range
    Dim headerName As Variant
    Dim cellRange As Range
    Dim combinedRange As Range
    
    If colHeaders Is Nothing Or colHeaders.Count = 0 Then Exit Function
    If ws Is Nothing Then Exit Function
    If rowNum < 1 Then Exit Function
    If mCurrentTargetTable Is Nothing Then Exit Function
    
    For Each headerName In colHeaders
        Set cellRange = GetCellByTableHeader(mCurrentTargetTable, CStr(headerName), rowNum)
        
        If Not cellRange Is Nothing Then
            If combinedRange Is Nothing Then
                Set combinedRange = cellRange
            Else
                Set combinedRange = Union(combinedRange, cellRange)
            End If
        End If
        Set cellRange = Nothing
    Next headerName
    
    Set BuildRowRangeFromHeaders = combinedRange
End Function


' ======================================================
' EXISTS IN ARRAY HELPER
' ======================================================
Private Function ExistsInArray(arr As Variant, val As String) As Boolean
    Dim v As Variant
    If Not IsArray(arr) Then Exit Function
    If IsEmpty(arr) Then Exit Function
    On Error Resume Next
    For Each v In arr
        If StrComp(Trim(CStr(v)), val, vbTextCompare) = 0 Then
            ExistsInArray = True
            Exit Function
        End If
    Next v
End Function
