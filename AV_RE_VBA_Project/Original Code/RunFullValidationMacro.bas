Attribute VB_Name = "RunFullValidationMacro"
Option Explicit
Public ValidationStartTime As Single
Public ValidationCancelTimeout As Single
Public ValidationCancelFlag As Boolean

' === Master full-validation runner with cancel/timeout support ===

' Master full-validation runner that mimics the SheetValidationTrigger behavior,
' then runs AutoCheckDataValidationTable checks against a DDM reference sheet.

Public Sub RunFullValidationMaster()
    Const MODULE_NAME As String = "ValidationMaster"

    Dim wsConfig As Worksheet, wsTarget As Worksheet
    Dim dataSheetName As String
    Dim startRow As Long, endRow As Long
    Dim keyColLetter As String, keyColNum As Long
    Dim validateSmartFuncColMap As Object
    Dim RowNum As Long, i As Long
    Dim langControl As String, english As Boolean
    Dim FormatMap As Object
    
    Dim keyRows() As Long
    Dim keyCount As Long, maxRows As Long
    Dim rowValues As Variant
    
    On Error GoTo ErrHandler
    
    '--- Initialize Validation Tracker Form ---
    ShowValidationTrackerForm
    AppendUserLog "Initializing Full Validation Master"
    
    InitDebugFlags
    
    ' --- Timer & cancel flags ---
    ValidationStartTime = Timer
    ValidationCancelTimeout = 10000
    ValidationCancelFlag = False
    AppendUserLog "Validation will run no longer than " & ValidationCancelTimeout & "s"
    
    ' --- Load Configurations ---
    AppendUserLog "Loading Configurations..."
    Set wsConfig = ThisWorkbook.Sheets("Config")
    dataSheetName = Trim(wsConfig.Range("B3").value)
    langControl = Trim(wsConfig.Range("M1").value)
    english = (langControl = "English")
    
    AppendUserLog "Current Sheet Language set to: " & langControl
    AppendUserLog "-----------------------------------------------"
    AppendUserLog "Configuration Sheet Loaded: '" & wsConfig.Name & "'"
    AppendUserLog "-----------------------------------------------"
    AppendUserLog "Setting Validation sequence targets..."
    AppendUserLog "Sheet name: '" & dataSheetName & "'"
    
    Set wsTarget = ThisWorkbook.Sheets(dataSheetName)
    
    startRow = CLng(wsConfig.Range("B4").value)
    endRow = startRow + CLng(wsConfig.Range("D4").value)
    keyColLetter = Trim(wsConfig.Range("B5").value)
    keyColNum = wsTarget.Range(keyColLetter & "1").Column
    
    AppendUserLog "Row Key set to: " & Trim(wsConfig.Range("A5").value) & " (" & keyColLetter & ")"
    AppendUserLog "Row Range: " & CStr(startRow) & ":" & CStr(endRow)
    ValidationTrackerForm.setAutoValInitCB True
    AppendUserLog "-----------------------------------------------"
    AppendUserLog "Initializing Advanced Autovalidation Configurations"
    AppendUserLog "-----------------------------------------------"
    
    
    ' INITIALIZING MAPPING DICTIONNARIES
    ' --- Initialize validation column mapping ---
    Set validateSmartFuncColMap = GetValidationColumns(wsConfig)
    
    ' --- Initialize Function Name Mapping ---
    Dim AdvFunctionMap As Object
    Set AdvFunctionMap = GetAutoValidationMap(wsConfig)
    
    ' --- Initialize Format Assignment Mapping ---
    Set FormatMap = LoadFormatMap(wsConfig)
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    DebugMessage "=== RunFullValidationMaster scanning rows " & startRow & "–" & endRow & " ===", MODULE_NAME
    DebugMessage "[RunFullValidationMaster] Number of Smart Functions Mapped successfully: " & AdvFunctionMap.count & " loaded", MODULE_NAME
    
    ' --- Precompute rows with keys ---
    maxRows = endRow - startRow + 1
    ReDim keyRows(1 To maxRows)
    keyCount = 0
    For RowNum = startRow To endRow
        If Trim(CStr(wsTarget.Cells(RowNum, keyColNum).value)) <> "" Then
            keyCount = keyCount + 1
            keyRows(keyCount) = RowNum
        End If
    Next RowNum
    If keyCount = 0 Then
        AppendUserLog "No rows found for validation. Exiting."
        GoTo Cleanup
    End If
    
    ' Shrink keyRows array to actual size
    ReDim Preserve keyRows(1 To keyCount)
    
    DebugMessage "[RunFullValidationMaster] " & keyCount & " rows have valid keys. Progress bar target: " & keyCount, MODULE_NAME
    AppendUserLog "Number of rows identified for validation: " & CStr(keyCount)
    AppendUserLog "-----------------------------------------------"
    AppendUserLog "Cycling through each row identified for validation"
    
    '-------------------------------------
    ' Initialize Validation Column Map
    '-------------------------------------
    Dim colMetaDict As Object
    Set colMetaDict = GetDDMValidationColumns(wsConfig)
    If colMetaDict Is Nothing Or colMetaDict.count = 0 Then
        DebugMessage "[RunAutoCheckDataValidation] No validation columns found in Config sheet.", "RunAutoCheckDataValidation"
        Exit Sub
    End If
    
    ' --- Loop over key rows ---
    For i = LBound(keyRows) To UBound(keyRows)
        RowNum = keyRows(i)
        
        ' UI responsiveness
        If i Mod 10 = 0 Then DoEvents
        
        ' Cancel / Timeout checks
        If ValidationCancelFlag Then
            AppendUserLog "Validation Cancelled by the user on row " & RowNum
            DebugMessage "[RunFullValidationMaster] Cancelled by user at row " & RowNum, MODULE_NAME
            GoTo Cleanup
        End If
        If ValidationTimeoutReached() Then
            AppendUserLog "Validation timer elapsed and ended by system on row " & RowNum
            DebugMessage "[RunFullValidationMaster] Cancelled due to timeout at row " & RowNum, MODULE_NAME
            GoTo Cleanup
        End If
        
        ' Check ForceValidation
        If ShouldValidateRow(RowNum, wsTarget, True) Then
            DebugMessage "? Validating Row " & RowNum, MODULE_NAME
            
            ' Optional: read row values once (performance optimization)
            rowValues = wsTarget.Rows(RowNum).value
            
            ValidateSingleRow wsTarget, RowNum, AdvFunctionMap, english, FormatMap
        Else
            DebugMessage "Skipping row " & RowNum & " (ForceValidation=False)", MODULE_NAME
        End If
    Next i
    
    AppendUserLog "-----------------------------------------------"
    AppendUserLog "-----------------------------------------------"
    AppendUserLog "ADVANCED AUTO VALIDATIONS COMPLETE"
    AppendUserLog "-----------------------------------------------"
    ValidationTrackerForm.setAdvValCompletedCB True
    
    AppendUserLog "Initiating standard data validation check..."
    DebugMessage "Starting RunAutoCheckDataValidation() pass.", MODULE_NAME

    
    ' Compile all columns that will be reviewed from both mapping dictionnaries
    Dim colReviewedColumnList As Collection
    Set colReviewedColumnList = BuildCollectionOfColumnLetters(colMetaDict, validateSmartFuncColMap)
    
    Call RunAutoCheckDataValidation(wsConfig, wsTarget, keyRows, keyColNum, english, FormatMap, colMetaDict, colReviewedColumnList)
    
    
    ' Use colMetaDict and validateSmartFuncColMap to set row/scan range
    'Call FormatKeyCell(validateColMap, FormatMap)
    DebugMessage "RunAutoCheckDataValidation() completed.", MODULE_NAME
    
Cleanup:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    DebugMessage "RunFullValidationMaster completed at " & Now & ".", MODULE_NAME
    Exit Sub

ErrHandler:
    DebugMessage "RunFullValidationMaster ERROR: " & Err.Number & " - " & Err.Description, MODULE_NAME
    AppendUserLog "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    AppendUserLog "Validation was interrupted due to an unexpected error"
    AppendUserLog "ERROR in RunFullValidationMaster"
    AppendUserLog "ERROR #" & Err.Number
    AppendUserLog "Description:"
    AppendUserLog Err.Description
    AppendUserLog "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    BringFormToFront ValidationTrackerForm
    Resume Cleanup
End Sub

Public Function BuildCollectionOfColumnLetters(colMetaDict As Object, validateSmartFuncColMap As Object) As Collection
    Dim colLetters As New Collection
    Dim key As Variant
    Dim existsDict As Object
    Set existsDict = CreateObject("Scripting.Dictionary") ' for uniqueness
    
    ' --- Add all keys from colMetaDict ---
    For Each key In colMetaDict.Keys
        If Not existsDict.Exists(UCase(key)) Then
            existsDict(UCase(key)) = True
            colLetters.Add UCase(key)
        End If
    Next key
    
    ' --- Add all keys from validateSmartFuncColMap ---
    For Each key In validateSmartFuncColMap.Keys
        If Not existsDict.Exists(UCase(key)) Then
            existsDict(UCase(key)) = True
            colLetters.Add UCase(key)
        End If
    Next key
    
    Set BuildCollectionOfColumnLetters = colLetters
End Function


' ==========================
' Timeout helper function
' ==========================
Public Function ValidationTimeoutReached() As Boolean
    ' Returns True if the elapsed time exceeds the timeout
    If ValidationCancelTimeout <= 0 Then
        ValidationTimeoutReached = False
    Else
        ValidationTimeoutReached = (Timer - ValidationStartTime > ValidationCancelTimeout)
    End If
End Function

' === Optional: manual cancel macro attached to a button ===
Public Sub CancelValidation()
    ValidationCancelFlag = True
End Sub


' ========================================================
' RunAutoCheckDataValidation
' Optimized version — preloads validation metadata
' ========================================================
' ========================================================
' RunAutoCheckDataValidation (Finalized for current architecture)
' ========================================================
'=====================================================
' Fixed RunAutoCheckDataValidation with defensive checks
Public Sub RunAutoCheckDataValidation(wsConfig As Worksheet, _
                                     wsTarget As Worksheet, _
                                     keyRows() As Long, _
                                     keyColNum As Long, _
                                     english As Boolean, Optional FormatMap As Object, Optional colMetaDict As Object, Optional RevColLetList As Collection)

    On Error GoTo ErrHandler

    Dim meta As Object
    Dim colKey As Variant
    Dim i As Long, RowNum As Long
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
    


    AppendUserLog "Standard Validation Configuration Map completed"

    '-------------------------------------
    ' Validate Inputs
    '-------------------------------------
    If wsConfig Is Nothing Or wsTarget Is Nothing Then Exit Sub
    If LBound(keyRows) > UBound(keyRows) Then Exit Sub

    totalValid = UBound(keyRows) - LBound(keyRows) + 1
    DebugMessage "[RunAutoCheckDataValidation] Starting validation on " & totalValid & " rows.", "RunAutoCheckDataValidation"

    '-------------------------------------
    ' Loop through rows
    '-------------------------------------
    For i = LBound(keyRows) To UBound(keyRows)
        RowNum = keyRows(i)
        
        ' Defensive: ensure row exists
        If RowNum <= 0 Or RowNum > wsTarget.Rows.count Then GoTo SkipRow
        
        ' Initialize dictionary to accumulate messages per drop column
        Set dropColMsgs = CreateObject("Scripting.Dictionary")
        
        ' Loop through validation columns
        For Each colKey In colMetaDict.Keys
            Set meta = colMetaDict(colKey)
            
            ' Ensure required keys exist
            If Not meta.Exists("ReviewLetter") Then GoTo SkipCol
            If Not meta.Exists("ValidColumnListEN") Then meta("ValidColumnListEN") = Array()
            If Not meta.Exists("ValidColumnListFR") Then meta("ValidColumnListFR") = Array()
            If Not meta.Exists("ColumnNameEN") Then meta("ColumnNameEN") = ""
            If Not meta.Exists("ColumnNameFR") Then meta("ColumnNameFR") = ""
            If Not meta.Exists("CommentDropCol") Then meta("CommentDropCol") = ""
            
            ' Safely get cell value
            cellValue = ""
            On Error Resume Next
            cellValue = Trim(CStr(wsTarget.Range(meta("ReviewLetter") & RowNum).value))
            On Error GoTo 0
            If Len(cellValue) = 0 Then GoTo SkipCol

            ' Check existence in EN/FR lists
            found = False
            If IsArray(meta("ValidColumnListEN")) Then found = ExistsInArray(meta("ValidColumnListEN"), cellValue)
            If Not found And IsArray(meta("ValidColumnListFR")) Then found = ExistsInArray(meta("ValidColumnListFR"), cellValue)

            ' Flag invalids
            If Not found Then
                Dim errorMsg As String
                If english Then
                    errorMsg = meta("ColumnNameEN") & " - Invalid value '" & cellValue & "' : Select a valid value from the list."
                Else
                    errorMsg = meta("ColumnNameFR") & " - Valeur invalide '" & cellValue & "' . Sélectionner une valeur valide."
                End If

                ' Append to dropColMsgs dictionary
                If Not dropColMsgs.Exists(meta("CommentDropCol")) Then
                    Set dropColMsgs(meta("CommentDropCol")) = CreateObject("Scripting.Dictionary")
                End If
                
                ' Prepare array: 1=sourceColLetter, 2=msgText
                msgArr(1) = meta("ReviewLetter") ' Column letter
                msgArr(2) = errorMsg ' Error Message
                msgArr(3) = "Error" ' Format Type String
                dropColMsgs(meta("CommentDropCol")).Add dropColMsgs(meta("CommentDropCol")).count + 1, msgArr
            Else
                Dim CellFormCheck As Range
                Dim CellRangeString As String
                CellRangeString = CStr(meta("ReviewLetter")) & RowNum
                Set CellFormCheck = wsTarget.Range(CellRangeString)
                If getFormatType(CellFormCheck, FormatMap) = "Error" Then
                    ' Append to dropColMsgs dictionary
                    If Not dropColMsgs.Exists(meta("CommentDropCol")) Then
                        Set dropColMsgs(meta("CommentDropCol")) = CreateObject("Scripting.Dictionary")
                    End If
                    
                    ' Prepare array: 1=sourceColLetter, 2=msgText
                    msgArr(1) = meta("ReviewLetter") ' Column letter
                    msgArr(2) = "" ' Error Message
                    msgArr(3) = "Default" ' Format Type String
                    dropColMsgs(meta("CommentDropCol")).Add dropColMsgs(meta("CommentDropCol")).count + 1, msgArr
                End If
            End If

SkipCol:
        Next colKey
                
        ' Write all accumulated messages to drop columns
        For Each dropColKey In dropColMsgs.Keys
            For Each cMsg In dropColMsgs(dropColKey).Items
                ' cMsg is a 2-element array: 1=sourceColLetter, 2=msgText, 3=Error type string - must match FormatMap
                DCMsgTxt = cMsg(2)
                cMsgErrorType = CStr(cMsg(3))
                WriteSystemTagToDropColumn wsTarget, CStr(dropColKey), RowNum, CStr(cMsg(1)), DCMsgTxt, cMsgErrorType, FormatMap
            Next cMsg
        Next dropColKey
        
        ' Build row range from RevColLetList
        Dim rowRange As Range
        Set rowRange = BuildRowRangeFromColumns(wsTarget, RevColLetList, RowNum)
        
        ' Set Row Key Format to highest priority correction format found for the user.
        Call FormatKeyCell(rowRange, FormatMap)
        
        
        ' Progress / UI responsiveness
        progressCount = progressCount + 1
        If progressCount Mod 10 = 0 Then DoEvents
        If progressCount Mod 25 = 0 Then AppendUserLog "[RunAutoCheckDataValidation] Progress: " & progressCount & " / " & totalValid

SkipRow:
    Next i

    ' Final logging
    DebugMessage "[RunAutoCheckDataValidation] Progress: " & progressCount & " / " & totalValid, "RunAutoCheckDataValidation"
    DebugMessage "[RunAutoCheckDataValidation] Completed.", "RunAutoCheckDataValidation"
    AppendUserLog "[RunAutoCheckDataValidation] Progress: " & progressCount & " / " & totalValid
    AppendUserLog "Standard menu accessible (locked-menu) field validation Completed."
    ValidationTrackerForm.setLMenuValCompletedCB True

    Exit Sub

ErrHandler:
    DebugMessage "RunAutoCheckDataValidation ERROR: " & Err.Number & " - " & Err.Description, "RunAutoCheckDataValidation"
    AppendUserLog "RunAutoCheckDataValidation ERROR: " & Err.Number & " - " & Err.Description
End Sub


Public Function BuildRowRangeFromColumns(ws As Worksheet, colLetters As Collection, RowNum As Long) As Range
    Dim colLetter As Variant
    Dim cellRange As Range
    Dim combinedRange As Range
    
    If colLetters Is Nothing Or colLetters.count = 0 Then Exit Function
    If ws Is Nothing Then Exit Function
    If RowNum < 1 Then Exit Function
    
    ' --- Build combined range dynamically ---
    For Each colLetter In colLetters
        On Error Resume Next
        Set cellRange = ws.Range(UCase(colLetter) & RowNum)
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

' ========================================================
' Helper: Check if value exists in array
' ========================================================
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


Private Function GetDropColLetterForValidation(wsConfig As Worksheet, reviewColLetter As String) As String
    Dim tbl As ListObject, r As ListRow
    Dim reviewCol As String, dropCol As String
    
    On Error Resume Next
    Set tbl = wsConfig.ListObjects("AutoCheckDataValidationTable")
    On Error GoTo 0
    If tbl Is Nothing Then Exit Function
    
    For Each r In tbl.ListRows
        reviewCol = Trim(CStr(r.Range.Cells(1, tbl.ListColumns("ReviewSheet Column Letter").Index).value))
        If reviewCol = reviewColLetter Then
            dropCol = Trim(CStr(r.Range.Cells(1, tbl.ListColumns("AutoComment Column").Index).value))
            GetDropColLetterForValidation = dropCol
            Exit Function
        End If
    Next r
End Function






