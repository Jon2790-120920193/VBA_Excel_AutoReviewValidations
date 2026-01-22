Attribute VB_Name = "AV_Engine"
Option Explicit

' ======================================================
' AV_Engine.bas
' Validation orchestration & execution engine
' UPDATED FOR v2.1: Table-based configuration
' ======================================================

Private Const MODULE_NAME As String = "AV_Engine"

' ======================================================
' PUBLIC ENTRY POINTS
' ======================================================

Public Sub RunFullValidation(Optional ByVal sheetName As String = "", Optional ByVal english As Boolean = True)
    RunFullValidationMaster sheetName, english
End Sub

' ======================================================
' MAIN VALIDATION EXECUTION (v2.1 - TABLE-BASED)
' ======================================================
Public Sub RunFullValidationMaster(Optional ByVal sheetName As String = "", Optional ByVal english As Boolean = True)

    Dim wsConfig As Worksheet
    Dim config As AV_Core.ValidationConfig
    Dim errorMsg As String
    
    Dim AdvFunctionMap As Object
    Dim FormatMap As Object
    Dim colMetaDict As Object
    Dim validateSmartFuncColMap As Object
    
    Dim targetIndex As Long
    
    On Error GoTo ErrHandler

    ' Initialize UI / logging
    AV_UI.ShowValidationTrackerForm
    DoEvents  ' Give form time to initialize
    AV_UI.AppendUserLog "=== Initializing Full Validation Master v2.1 ==="
    
    AV_Core.InitDebugFlags

    ' Set configuration sheet reference
    Set wsConfig = ThisWorkbook.Sheets(AV_Constants.CONFIG_SHEET_NAME)

    ' ======================================================
    ' STEP 1: VALIDATE CONFIGURATION
    ' ======================================================
    AV_UI.AppendUserLog "Step 1: Validating configuration..."
    
    If Not AV_Core.ValidateConfiguration(errorMsg) Then
        AV_UI.AppendUserLog "CONFIGURATION ERROR:"
        AV_UI.AppendUserLog errorMsg
        AV_UI.AppendUserLog "Validation aborted. Please fix configuration and try again."
        MsgBox "Configuration Error:" & vbCrLf & vbCrLf & errorMsg, vbCritical, "Validation Aborted"
        GoTo Cleanup
    End If
    
    AV_UI.AppendUserLog "Configuration validated successfully."
    
    ' ======================================================
    ' STEP 2: LOAD CONFIGURATION
    ' ======================================================
    AV_UI.AppendUserLog "Step 2: Loading validation configuration..."
    
    config = AV_Core.LoadValidationConfig()
    
    If config.TargetCount = 0 Then
        AV_UI.AppendUserLog "WARNING: No validation targets enabled."
        AV_UI.AppendUserLog "Check ValidationTargets table - at least one target must have Enabled=TRUE"
        MsgBox "No validation targets enabled." & vbCrLf & _
               "Please enable at least one target in ValidationTargets table.", _
               vbExclamation, "No Targets"
        GoTo Cleanup
    End If
    
    AV_UI.AppendUserLog "Language: " & config.Language
    AV_UI.AppendUserLog "Enabled targets: " & config.TargetCount
    
    For targetIndex = 1 To config.TargetCount
        AV_UI.AppendUserLog "  - " & config.Targets(targetIndex).TableName & _
                           " (Mode: " & config.Targets(targetIndex).Mode & ")"
    Next targetIndex

    ' Cancel / timeout flags
    AV_Core.ValidationStartTime = Timer
    AV_Core.ValidationCancelTimeout = AV_Constants.VALIDATION_TIMEOUT_SECONDS
    AV_Core.ValidationCancelFlag = False

    AV_UI.AppendUserLog "Timeout: " & AV_Core.ValidationCancelTimeout & " seconds"

    ' ======================================================
    ' STEP 3: LOAD MAPPING & FORMAT DATA
    ' ======================================================
    AV_UI.AppendUserLog "Step 3: Loading validation mappings..."
    
    Set AdvFunctionMap = AV_Core.GetAutoValidationMap(wsConfig)
    Set FormatMap = AV_Format.LoadFormatMap(wsConfig)
    Set colMetaDict = AV_Core.GetDDMValidationColumns(wsConfig)
    Set validateSmartFuncColMap = AV_Core.GetValidationColumns(wsConfig)

    If AdvFunctionMap Is Nothing Or AdvFunctionMap.Count = 0 Then
        AV_UI.AppendUserLog "ERROR: No validation functions mapped."
        AV_UI.AppendUserLog "Check AutoValidationCommentPrefixMappingTable"
        GoTo Cleanup
    End If

    AV_UI.AppendUserLog "Validation functions loaded: " & AdvFunctionMap.Count
    AV_UI.AppendUserLog "Format mappings loaded: " & FormatMap.Count
    AV_UI.SetAutoValidationInitialized True

    ' ======================================================
    ' STEP 4: VALIDATE EACH TARGET
    ' ======================================================
    AV_UI.AppendUserLog "==============================================="
    AV_UI.AppendUserLog "Step 4: Processing validation targets..."
    AV_UI.AppendUserLog "==============================================="
    
    For targetIndex = 1 To config.TargetCount
        ' Check for cancellation
        If AV_Core.ValidationCancelFlag Then
            AV_UI.AppendUserLog "Validation cancelled by user."
            GoTo Cleanup
        End If
        
        If AV_Core.ValidationTimeoutReached() Then
            AV_UI.AppendUserLog "Validation stopped due to timeout."
            GoTo Cleanup
        End If
        
        ' Process this target
        Call ProcessValidationTarget( _
            config.Targets(targetIndex), _
            config.IsEnglish, _
            AdvFunctionMap, _
            FormatMap, _
            colMetaDict, _
            validateSmartFuncColMap _
        )
    Next targetIndex

    ' ======================================================
    ' STEP 5: CLEANUP & COMPLETION
    ' ======================================================
    AV_UI.AppendUserLog "==============================================="
    AV_UI.AppendUserLog "VALIDATION COMPLETE"
    AV_UI.AppendUserLog "==============================================="
    AV_Core.DebugMessage "All validation targets processed successfully.", MODULE_NAME

Cleanup:
    ' Clear cached tables to free memory
    AV_Core.ClearTableCache
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    AV_Core.DebugMessage "RunFullValidationMaster completed at " & Now & ".", MODULE_NAME
    Exit Sub

ErrHandler:
    AV_UI.AppendUserLog "CRITICAL ERROR in RunFullValidationMaster"
    AV_UI.AppendUserLog "Error #" & Err.Number & ": " & Err.Description
    AV_UI.AppendUserLog "Source: " & Err.Source
    AV_UI.BringFormToFront ValidationTrackerForm
    MsgBox "Critical Error:" & vbCrLf & vbCrLf & _
           "Error #" & Err.Number & ": " & Err.Description, _
           vbCritical, "Validation Error"
    Resume Cleanup
End Sub


' ======================================================
' PROCESS VALIDATION TARGET (NEW IN v2.1)
' ======================================================
Private Sub ProcessValidationTarget( _
    ByRef target As AV_Core.ValidationTarget, _
    ByVal english As Boolean, _
    ByRef AdvFunctionMap As Object, _
    ByRef FormatMap As Object, _
    ByRef colMetaDict As Object, _
    ByRef validateSmartFuncColMap As Object _
)
    Dim wsTarget As Worksheet
    Dim tblTarget As ListObject
    Dim keyColIndex As Long
    Dim rowNum As Long, i As Long
    Dim keyRows() As Long
    Dim keyCount As Long
    Dim dataRow As ListRow
    Dim colReviewedColumnList As Collection
    
    On Error GoTo ErrHandler
    
    AV_UI.AppendUserLog "-----------------------------------------------"
    AV_UI.AppendUserLog "Processing target: " & target.TableName
    AV_UI.AppendUserLog "-----------------------------------------------"
    
    ' Find target ListObject by name
    Set tblTarget = AV_Core.FindListObjectByName(target.TableName)
    
    If tblTarget Is Nothing Then
        AV_UI.AppendUserLog "ERROR: Table '" & target.TableName & "' not found. Skipping."
        Exit Sub
    End If
    
    Set wsTarget = tblTarget.Parent
    AV_UI.AppendUserLog "Found in sheet: " & wsTarget.Name
    AV_UI.AppendUserLog "Rows: " & tblTarget.ListRows.Count
    
    ' Find key column
    On Error Resume Next
    keyColIndex = tblTarget.ListColumns(target.KeyColumnHeader).Index
    On Error GoTo ErrHandler
    
    If keyColIndex = 0 Then
        AV_UI.AppendUserLog "ERROR: Key column '" & target.KeyColumnHeader & "' not found. Skipping."
        Exit Sub
    End If
    
    AV_UI.AppendUserLog "Key column: " & target.KeyColumnHeader & " (Index: " & keyColIndex & ")"
    
    ' ======================================================
    ' BUILD ROW LIST
    ' ======================================================
    AV_UI.AppendUserLog "Identifying rows to validate..."
    
    ReDim keyRows(1 To tblTarget.ListRows.Count)
    keyCount = 0
    
    For Each dataRow In tblTarget.ListRows
        ' Get actual row number in worksheet
        rowNum = dataRow.Range.Row
        
        ' Check if key column has value
        If Trim(CStr(dataRow.Range.Cells(1, keyColIndex).Value)) <> "" Then
            ' Check if this row should be validated
            If AV_Core.ShouldValidateRow(rowNum, wsTarget, True) Then
                keyCount = keyCount + 1
                keyRows(keyCount) = rowNum
            End If
        End If
    Next dataRow
    
    If keyCount = 0 Then
        AV_UI.AppendUserLog "No rows to validate in this target. Skipping."
        Exit Sub
    End If
    
    ReDim Preserve keyRows(1 To keyCount)
    
    AV_Core.DebugMessage "Rows to validate: " & keyCount, MODULE_NAME
    AV_UI.AppendUserLog "Rows identified: " & keyCount
    
    ' ======================================================
    ' MAIN ROW VALIDATION LOOP
    ' ======================================================
    AV_UI.AppendUserLog "Beginning row validation..."
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For i = LBound(keyRows) To UBound(keyRows)
        rowNum = keyRows(i)

        ' Progress updates
        If i Mod AV_Constants.VALIDATION_PROGRESS_UPDATE_INTERVAL = 0 Then
            DoEvents
            AV_UI.AppendUserLog "Progress: " & i & " / " & keyCount & " rows processed"
        End If

        ' Check cancellation
        If AV_Core.ValidationCancelFlag Then
            AV_UI.AppendUserLog "Validation cancelled by user."
            Exit Sub
        End If

        If AV_Core.ValidationTimeoutReached() Then
            AV_UI.AppendUserLog "Validation stopped due to timeout."
            Exit Sub
        End If

        ' Validate this row
        ValidateSingleRow wsTarget, rowNum, AdvFunctionMap, english, FormatMap
    Next i

    AV_UI.AppendUserLog "Row validation complete for " & target.TableName

    ' ======================================================
    ' POST-PASS: SIMPLE DATA VALIDATION
    ' ======================================================
    AV_UI.AppendUserLog "Running simple dropdown validation..."
    AV_Core.DebugMessage "Starting RunAutoCheckDataValidation() pass.", MODULE_NAME

    Set colReviewedColumnList = BuildCollectionOfColumnLetters(colMetaDict, validateSmartFuncColMap)
    
    ' Get key column number for formatting
    Dim keyColNum As Long
    keyColNum = tblTarget.ListColumns(target.KeyColumnHeader).DataBodyRange.Column

    RunAutoCheckDataValidation wsTarget, keyRows, keyColNum, english, FormatMap, colMetaDict, colReviewedColumnList

    AV_Core.DebugMessage "RunAutoCheckDataValidation() completed.", MODULE_NAME
    AV_UI.AppendUserLog "Target validation complete: " & target.TableName
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    Exit Sub

ErrHandler:
    AV_UI.AppendUserLog "ERROR in ProcessValidationTarget"
    AV_UI.AppendUserLog "Target: " & target.TableName
    AV_UI.AppendUserLog "Error #" & Err.Number & ": " & Err.Description
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


' ======================================================
' VALIDATE SINGLE ROW
' ======================================================
Public Sub ValidateSingleRow(wsData As Worksheet, rowNum As Long, AdvFunctionMap As Object, english As Boolean, FormatMap As Object)
    Dim colLetter As Variant
    Dim funcName As String
    Dim TargetCell As Range
    Dim mapItem As Object
    Dim AutoValidate As Boolean
    Dim TargetColumnLet As String

    For Each colLetter In AdvFunctionMap.Keys
        Set mapItem = AdvFunctionMap(colLetter)
        funcName = CStr(colLetter)

        ' Retrieve AutoValidate flag
        AutoValidate = False
        If mapItem.Exists("AutoValidate") Then
            AutoValidate = mapItem("AutoValidate")
        End If
        
        ' Retrieve ColumnRef safely
        TargetColumnLet = ""
        If mapItem.Exists("ColumnRef") Then
            TargetColumnLet = CStr(mapItem("ColumnRef"))
        End If
        
        If Len(TargetColumnLet) = 0 Then
            AV_Core.DebugMessage "WARNING: Missing ColumnRef for " & funcName, MODULE_NAME
            GoTo SkipToNext
        End If
        
        ' Skip if AutoValidate = False
        If AutoValidate = False Then
            AV_Core.DebugMessage "Skipping " & funcName & " (AutoValidate=False)", MODULE_NAME
            GoTo SkipToNext
        End If

        ' Proceed with validation
        On Error Resume Next
        Set TargetCell = wsData.Range(TargetColumnLet & rowNum)
        On Error GoTo 0
        
        If Not TargetCell Is Nothing Then
            On Error GoTo ValidationError
            Application.Run funcName, TargetCell, wsData.Name, english, FormatMap, AdvFunctionMap
            On Error GoTo 0
        End If
        
SkipToNext:
    Next colLetter

    ' Only log every N rows to reduce clutter
    If rowNum Mod AV_Constants.VALIDATION_DETAILED_LOG_INTERVAL = 0 Then
        AV_UI.AppendUserLog "Row " & rowNum & " validation complete"
    End If
    
    Exit Sub

ValidationError:
    AV_Core.DebugMessage "Error validating row " & rowNum & ", column " & colLetter & ", function: " & funcName, MODULE_NAME
    AV_UI.AppendUserLog "Warning: Error in row " & rowNum & ", column " & colLetter
    Resume Next
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
' RUN AUTO CHECK DATA VALIDATION (UPDATED v2.1)
' ======================================================
Public Sub RunAutoCheckDataValidation(wsTarget As Worksheet, _
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

    If wsTarget Is Nothing Then Exit Sub
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
                msgArr(3) = AV_Constants.FORMAT_ERROR
                dropColMsgs(meta("CommentDropCol")).Add dropColMsgs(meta("CommentDropCol")).Count + 1, msgArr
            Else
                ' Check if cell was previously marked as error
                Dim CellFormCheck As Range
                Dim CellRangeString As String
                CellRangeString = CStr(meta("ReviewLetter")) & rowNum
                Set CellFormCheck = wsTarget.Range(CellRangeString)
                If AV_Format.getFormatType(CellFormCheck, FormatMap) = AV_Constants.FORMAT_ERROR Then
                    If Not dropColMsgs.Exists(meta("CommentDropCol")) Then
                        Set dropColMsgs(meta("CommentDropCol")) = CreateObject("Scripting.Dictionary")
                    End If
                    
                    msgArr(1) = meta("ReviewLetter")
                    msgArr(2) = ""
                    msgArr(3) = AV_Constants.FORMAT_DEFAULT
                    dropColMsgs(meta("CommentDropCol")).Add dropColMsgs(meta("CommentDropCol")).Count + 1, msgArr
                End If
            End If

SkipCol:
        Next colKey
                
        ' Write all messages for this row
        For Each dropColKey In dropColMsgs.Keys
            For Each cMsg In dropColMsgs(dropColKey).Items
                DCMsgTxt = cMsg(2)
                cMsgErrorType = CStr(cMsg(3))
                AV_Format.WriteSystemTagToDropColumn wsTarget, CStr(dropColKey), rowNum, CStr(cMsg(1)), DCMsgTxt, cMsgErrorType, FormatMap
            Next cMsg
        Next dropColKey
        
        ' Format key cell
        Dim rowRange As Range
        Set rowRange = BuildRowRangeFromColumns(wsTarget, RevColLetList, rowNum)
        AV_Format.FormatKeyCell rowRange, FormatMap
        
        progressCount = progressCount + 1
        If progressCount Mod AV_Constants.VALIDATION_PROGRESS_UPDATE_INTERVAL = 0 Then DoEvents

SkipRow:
    Next i

    AV_Core.DebugMessage "Progress: " & progressCount & " / " & totalValid, MODULE_NAME
    AV_Core.DebugMessage "RunAutoCheckDataValidation completed.", MODULE_NAME
    AV_UI.AppendUserLog "Simple validation complete: " & progressCount & " rows processed"
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
