Attribute VB_Name = "AV_Engine"
Option Explicit

' ======================================================
' AV_Engine.bas
' Validation orchestration & execution engine
' VERSION: 2.2 - Uses header-based cell lookup
' ======================================================

Private Const MODULE_NAME As String = "AV_Engine"

' Module-level reference to current target table
' This allows validators to find cells by header name
Private mCurrentTargetTable As ListObject

' ======================================================
' PUBLIC PROPERTY - CURRENT TABLE ACCESS
' ======================================================

Public Property Get CurrentTargetTable() As ListObject
    Set CurrentTargetTable = mCurrentTargetTable
End Property

Public Property Set CurrentTargetTable(tbl As ListObject)
    Set mCurrentTargetTable = tbl
End Property

' ======================================================
' PUBLIC ENTRY POINTS
' ======================================================

Public Sub RunFullValidation(Optional ByVal sheetName As String = "", Optional ByVal english As Boolean = True)
    RunFullValidationMaster sheetName, english
End Sub

' ======================================================
' MAIN VALIDATION EXECUTION
' ======================================================
Public Sub RunFullValidationMaster(Optional ByVal sheetName As String = "", Optional ByVal english As Boolean = True)

    Dim wsConfig As Worksheet
    Dim wsTarget As Worksheet
    Dim dataSheetName As String
    Dim startRow As Long, endRow As Long
    Dim keyColLetter As String, keyColNum As Long
    Dim rowNum As Long, i As Long
    Dim keyRows() As Long
    Dim keyCount As Long, maxRows As Long
    Dim rowValues As Variant

    Dim AdvFunctionMap As Object
    Dim FormatMap As Object
    Dim colMetaDict As Object
    Dim validateSmartFuncColMap As Object
    Dim colReviewedColumnList As Collection

    On Error GoTo ErrHandler

    ' Initialize UI / logging
    AV_UI.ShowValidationTrackerForm
    AV_UI.AppendUserLog "Initializing Full Validation Master"
    
    AV_Core.InitDebugFlags

    ' Cancel / timeout flags
    AV_Core.ValidationStartTime = Timer
    AV_Core.ValidationCancelTimeout = 10000
    AV_Core.ValidationCancelFlag = False

    AV_UI.AppendUserLog "Validation timeout set to " & AV_Core.ValidationCancelTimeout & " seconds"

    ' Load configuration
    Set wsConfig = ThisWorkbook.Sheets("Config")

    If sheetName = "" Then
        dataSheetName = Trim(wsConfig.Range("B3").Value)
    Else
        dataSheetName = sheetName
    End If

    Set wsTarget = ThisWorkbook.Sheets(dataSheetName)
    
    ' Get the target table (first table on the sheet)
    Set mCurrentTargetTable = AV_DataAccess.GetFirstTable(wsTarget)
    
    If mCurrentTargetTable Is Nothing Then
        AV_UI.AppendUserLog "ERROR: No table found on sheet " & dataSheetName
        AV_UI.AppendUserLog "The target sheet must contain an Excel Table (ListObject)"
        GoTo Cleanup
    End If
    
    AV_UI.AppendUserLog "Target table: " & mCurrentTargetTable.Name

    startRow = CLng(wsConfig.Range("B4").Value)
    endRow = startRow + CLng(wsConfig.Range("D4").Value)

    keyColLetter = Trim(wsConfig.Range("B5").Value)
    
    ' Handle key column - could be letter or header name
    If AV_DataAccess.IsColumnLetter(keyColLetter) Then
        keyColNum = wsTarget.Range(keyColLetter & "1").Column
    Else
        ' It's a header name - find the column
        Dim keyCol As ListColumn
        On Error Resume Next
        Set keyCol = mCurrentTargetTable.ListColumns(keyColLetter)
        On Error GoTo 0
        If Not keyCol Is Nothing Then
            keyColNum = keyCol.DataBodyRange.Column
        Else
            AV_UI.AppendUserLog "ERROR: Key column '" & keyColLetter & "' not found in table"
            GoTo Cleanup
        End If
    End If

    AV_UI.AppendUserLog "Target sheet: " & dataSheetName
    AV_UI.AppendUserLog "Row range: " & startRow & " to " & endRow

    ' Load mappings
    Set AdvFunctionMap = AV_Core.GetAutoValidationMap(wsConfig)
    Set FormatMap = AV_Format.LoadFormatMap(wsConfig)
    Set colMetaDict = AV_Core.GetDDMValidationColumns(wsConfig)
    Set validateSmartFuncColMap = AV_Core.GetValidationColumns(wsConfig)

    If AdvFunctionMap Is Nothing Or AdvFunctionMap.Count = 0 Then
        AV_UI.AppendUserLog "No validation functions mapped. Aborting."
        GoTo Cleanup
    End If

    AV_UI.AppendUserLog "-----------------------------------------------"
    AV_UI.AppendUserLog "Advanced Autovalidation Configurations Loaded"
    AV_UI.AppendUserLog "-----------------------------------------------"
    AV_UI.SetAutoValidationInitialized True

    ' Pre-compute rows with keys
    maxRows = endRow - startRow + 1
    ReDim keyRows(1 To maxRows)
    keyCount = 0

    For rowNum = startRow To endRow
        If Trim(CStr(wsTarget.Cells(rowNum, keyColNum).Value)) <> "" Then
            keyCount = keyCount + 1
            keyRows(keyCount) = rowNum
        End If
    Next rowNum

    If keyCount = 0 Then
        AV_UI.AppendUserLog "No valid rows found. Exiting."
        GoTo Cleanup
    End If

    ReDim Preserve keyRows(1 To keyCount)

    AV_Core.DebugMessage "Number of rows with keys: " & keyCount, MODULE_NAME
    AV_UI.AppendUserLog "Number of rows identified for validation: " & CStr(keyCount)
    AV_UI.AppendUserLog "-----------------------------------------------"
    AV_UI.AppendUserLog "Cycling through each row identified for validation"

    ' MAIN ROW LOOP
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For i = LBound(keyRows) To UBound(keyRows)
        rowNum = keyRows(i)

        If i Mod 10 = 0 Then DoEvents

        If AV_Core.ValidationCancelFlag Then
            AV_UI.AppendUserLog "Validation cancelled by user."
            GoTo Cleanup
        End If

        If AV_Core.ValidationTimeoutReached() Then
            AV_UI.AppendUserLog "Validation stopped due to timeout."
            GoTo Cleanup
        End If

        If AV_Core.ShouldValidateRow(rowNum, wsTarget, True) Then
            rowValues = wsTarget.Rows(rowNum).Value
            ValidateSingleRow wsTarget, rowNum, AdvFunctionMap, english, FormatMap, mCurrentTargetTable
        End If
    Next i

    AV_UI.AppendUserLog "-----------------------------------------------"
    AV_UI.AppendUserLog "ADVANCED AUTO VALIDATIONS COMPLETE"
    AV_UI.AppendUserLog "-----------------------------------------------"
    AV_UI.SetAdvancedValidationCompleted True

    ' Post-pass: simple data validation
    AV_UI.AppendUserLog "Initiating standard data validation check..."
    AV_Core.DebugMessage "Starting RunAutoCheckDataValidation() pass.", MODULE_NAME

    Set colReviewedColumnList = BuildCollectionOfColumnLetters(colMetaDict, validateSmartFuncColMap)

    RunAutoCheckDataValidation wsConfig, wsTarget, keyRows, keyColNum, english, FormatMap, colMetaDict, colReviewedColumnList

    AV_Core.DebugMessage "RunAutoCheckDataValidation() completed.", MODULE_NAME

Cleanup:
    Set mCurrentTargetTable = Nothing  ' Clear reference
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    AV_Core.DebugMessage "RunFullValidationMaster completed at " & Now & ".", MODULE_NAME
    Exit Sub

ErrHandler:
    AV_UI.AppendUserLog "ERROR in RunFullValidationMaster"
    AV_UI.AppendUserLog "Error #" & Err.Number & ": " & Err.Description
    AV_UI.BringFormToFront ValidationTrackerForm
    Resume Cleanup
End Sub


' ======================================================
' VALIDATE SINGLE ROW
' Updated to use table-aware cell lookup
' ======================================================
Public Sub ValidateSingleRow(wsData As Worksheet, rowNum As Long, AdvFunctionMap As Object, _
                             english As Boolean, FormatMap As Object, _
                             Optional targetTable As ListObject = Nothing)
    Dim colKey As Variant
    Dim funcName As String
    Dim TargetCell As Range
    Dim mapItem As Object
    Dim AutoValidate As Boolean
    Dim ColumnRef As String
    Dim DisplayName As String

    ' Use passed table or module-level reference
    If targetTable Is Nothing Then
        Set targetTable = mCurrentTargetTable
    End If

    For Each colKey In AdvFunctionMap.Keys
        Set mapItem = AdvFunctionMap(colKey)
        funcName = CStr(colKey)

        ' Retrieve AutoValidate flag
        AutoValidate = False
        If mapItem.Exists("AutoValidate") Then
            AutoValidate = mapItem("AutoValidate")
        End If
        
        ' Retrieve ColumnRef safely (could be letter or header name)
        ColumnRef = ""
        If mapItem.Exists("ColumnRef") Then
            ColumnRef = CStr(mapItem("ColumnRef"))
        End If
        
        ' Get display name for logging
        DisplayName = ColumnRef
        
        If Len(ColumnRef) = 0 Then
            Debug.Print "[ValidateSingleRow] WARNING: Missing ColumnRef for " & funcName
            GoTo SkipToNext
        End If
        
        ' Skip if AutoValidate = False
        If AutoValidate = False Then
            Debug.Print "[ValidateSingleRow] Skipping " & funcName & " (AutoValidate=False)"
            GoTo SkipToNext
        End If

        ' Get target cell using smart lookup (handles both letters and header names)
        Set TargetCell = AV_DataAccess.GetCellSmart(wsData, ColumnRef, rowNum, targetTable)
        
        If TargetCell Is Nothing Then
            AV_Core.DebugMessage "WARNING: Column '" & ColumnRef & "' not found in table for " & funcName, MODULE_NAME
            GoTo SkipToNext
        End If
        
        AV_Core.DebugMessage "Validating row " & rowNum & ", column '" & DisplayName & "' with " & funcName, MODULE_NAME
        
        On Error GoTo ValidationError
        Application.Run funcName, TargetCell, wsData.Name, english, FormatMap, AdvFunctionMap
        On Error GoTo 0
        
SkipToNext:
    Next colKey

    AV_UI.AppendUserLog "---Row " & rowNum & " Validation Complete---"
    Exit Sub

ValidationError:
    AV_Core.DebugMessage "[ValidateSingleRow] Error: Row " & rowNum & " - Column '" & ColumnRef & "' - Function: " & funcName & " - " & Err.Description, MODULE_NAME
    AV_UI.AppendUserLog "Error during validation: Row " & rowNum & " - " & funcName
    Resume SkipToNext
End Sub


' ======================================================
' BUILD COLLECTION OF COLUMN LETTERS
' ======================================================
Public Function BuildCollectionOfColumnLetters(colMetaDict As Object, validateSmartFuncColMap As Object) As Collection
    Dim colLetters As New Collection
    Dim key As Variant
    Dim existsDict As Object
    Set existsDict = CreateObject("Scripting.Dictionary")
    
    ' Add all keys from colMetaDict
    For Each key In colMetaDict.Keys
        If Not existsDict.Exists(UCase(key)) Then
            existsDict(UCase(key)) = True
            colLetters.Add UCase(key)
        End If
    Next key
    
    ' Add all keys from validateSmartFuncColMap
    For Each key In validateSmartFuncColMap.Keys
        If Not existsDict.Exists(UCase(key)) Then
            existsDict(UCase(key)) = True
            colLetters.Add UCase(key)
        End If
    Next key
    
    Set BuildCollectionOfColumnLetters = colLetters
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
                                     Optional RevColLetList As Collection)

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

    AV_UI.AppendUserLog "Standard Validation Configuration Map completed"

    If wsConfig Is Nothing Or wsTarget Is Nothing Then Exit Sub
    If LBound(keyRows) > UBound(keyRows) Then Exit Sub

    totalValid = UBound(keyRows) - LBound(keyRows) + 1
    AV_Core.DebugMessage "Starting validation on " & totalValid & " rows.", MODULE_NAME

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
            
            cellValue = ""
            On Error Resume Next
            cellValue = Trim(CStr(wsTarget.Range(meta("ReviewLetter") & rowNum).Value))
            On Error GoTo 0
            If Len(cellValue) = 0 Then GoTo SkipCol

            found = False
            If IsArray(meta("ValidColumnListEN")) Then found = ExistsInArray(meta("ValidColumnListEN"), cellValue)
            If Not found And IsArray(meta("ValidColumnListFR")) Then found = ExistsInArray(meta("ValidColumnListFR"), cellValue)

            If Not found Then
                Dim errorMsg As String
                If english Then
                    errorMsg = meta("ColumnNameEN") & " - Invalid value '" & cellValue & "' : Select a valid value from the list."
                Else
                    errorMsg = meta("ColumnNameFR") & " - Valeur invalide '" & cellValue & "' . Sélectionner une valeur valide."
                End If

                If Not dropColMsgs.Exists(meta("CommentDropCol")) Then
                    Set dropColMsgs(meta("CommentDropCol")) = CreateObject("Scripting.Dictionary")
                End If
                
                msgArr(1) = meta("ReviewLetter")
                msgArr(2) = errorMsg
                msgArr(3) = "Error"
                dropColMsgs(meta("CommentDropCol")).Add dropColMsgs(meta("CommentDropCol")).Count + 1, msgArr
            Else
                Dim CellFormCheck As Range
                Dim CellRangeString As String
                CellRangeString = CStr(meta("ReviewLetter")) & rowNum
                Set CellFormCheck = wsTarget.Range(CellRangeString)
                If AV_Format.getFormatType(CellFormCheck, FormatMap) = "Error" Then
                    If Not dropColMsgs.Exists(meta("CommentDropCol")) Then
                        Set dropColMsgs(meta("CommentDropCol")) = CreateObject("Scripting.Dictionary")
                    End If
                    
                    msgArr(1) = meta("ReviewLetter")
                    msgArr(2) = ""
                    msgArr(3) = "Default"
                    dropColMsgs(meta("CommentDropCol")).Add dropColMsgs(meta("CommentDropCol")).Count + 1, msgArr
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
        
        Dim rowRange As Range
        Set rowRange = BuildRowRangeFromColumns(wsTarget, RevColLetList, rowNum)
        
        AV_Format.FormatKeyCell rowRange, FormatMap
        
        progressCount = progressCount + 1
        If progressCount Mod 10 = 0 Then DoEvents
        If progressCount Mod 25 = 0 Then AV_UI.AppendUserLog "[RunAutoCheckDataValidation] Progress: " & progressCount & " / " & totalValid

SkipRow:
    Next i

    AV_Core.DebugMessage "Progress: " & progressCount & " / " & totalValid, MODULE_NAME
    AV_Core.DebugMessage "RunAutoCheckDataValidation completed.", MODULE_NAME
    AV_UI.AppendUserLog "[RunAutoCheckDataValidation] Progress: " & progressCount & " / " & totalValid
    AV_UI.AppendUserLog "Standard menu accessible (locked-menu) field validation Completed."
    AV_UI.SetLegacyMenuValidationCompleted True

    Exit Sub

ErrHandler:
    AV_Core.DebugMessage "RunAutoCheckDataValidation ERROR: " & Err.Number & " - " & Err.Description, MODULE_NAME
    AV_UI.AppendUserLog "RunAutoCheckDataValidation ERROR: " & Err.Number & " - " & Err.Description
End Sub


' ======================================================
' BUILD ROW RANGE FROM COLUMNS
' ======================================================
Public Function BuildRowRangeFromColumns(ws As Worksheet, colLetters As Collection, rowNum As Long) As Range
    Dim colLetter As Variant
    Dim cellRange As Range
    Dim combinedRange As Range
    
    If colLetters Is Nothing Or colLetters.Count = 0 Then Exit Function
    If ws Is Nothing Then Exit Function
    If rowNum < 1 Then Exit Function
    
    For Each colLetter In colLetters
        On Error Resume Next
        Set cellRange = ws.Range(UCase(colLetter) & rowNum)
        On Error GoTo 0
        
        If Not cellRange Is Nothing Then
            If combinedRange Is Nothing Then
                Set combinedRange = cellRange
            Else
                Set combinedRange = Union(combinedRange, cellRange)
            End If
        End If
        Set cellRange = Nothing
    Next colLetter
    
    Set BuildRowRangeFromColumns = combinedRange
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
