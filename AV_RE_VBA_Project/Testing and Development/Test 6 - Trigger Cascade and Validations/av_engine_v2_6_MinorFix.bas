Attribute VB_Name = "AV_Engine"
Option Explicit

' ======================================================
' AV_Engine.bas v2.6
' Validation orchestration & execution engine
' FIXED: Header-based cell lookups instead of column letters
' ======================================================

Private Const MODULE_NAME As String = "AV_Engine"
Public Const MODULE_VERSION As String = "2.6"

' Current target table for use by other modules (like AV_Validators.GetSiblingCell)
Public CurrentTargetTable As ListObject

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
    Dim config As AV_Core.ValidationConfig
    Dim errorMsg As String
    Dim i As Long

    On Error GoTo ErrHandler

    AV_UI.ShowValidationTrackerForm
    AV_UI.AppendUserLog "=== Auto-Validation Engine v" & MODULE_VERSION & " ==="
    AV_UI.AppendUserLog "Initializing at " & Format(Now, "yyyy-mm-dd hh:mm:ss")
    
    AV_Core.InitDebugFlags
    AV_Core.DebugMessage "RunFullValidationMaster started", MODULE_NAME

    AV_Core.BulkValidationInProgress = True
    AV_Core.ValidationStartTime = Timer
    AV_Core.ValidationCancelTimeout = 10000
    AV_Core.ValidationCancelFlag = False

    Set wsConfig = ThisWorkbook.Sheets("Config")
    
    If Not AV_Core.ValidateConfiguration(errorMsg) Then
        AV_UI.AppendUserLog "ERROR: " & errorMsg
        GoTo Cleanup
    End If
    
    AV_UI.AppendUserLog "Configuration validated successfully."
    
    config = AV_Core.LoadValidationConfig()
    
    AV_UI.AppendUserLog "Language: " & config.Language
    AV_UI.AppendUserLog "Enabled targets: " & config.TargetCount
    
    For i = 1 To config.TargetCount
        AV_UI.AppendUserLog "  - " & config.Targets(i).tableName & " (Mode: " & config.Targets(i).Mode & ")"
    Next i
    
    AV_UI.AppendUserLog "-----------------------------------------------"
    AV_UI.SetAutoValidationInitialized True

    For i = 1 To config.TargetCount
        If AV_Core.ValidationCancelFlag Then
            AV_UI.AppendUserLog "Validation cancelled by user."
            GoTo Cleanup
        End If
        
        If AV_Core.ValidationTimeoutReached() Then
            AV_UI.AppendUserLog "Validation stopped due to timeout."
            GoTo Cleanup
        End If
        
        ProcessValidationTarget config.Targets(i), wsConfig, english
    Next i
    
    AV_UI.AppendUserLog "==============================================="
    AV_UI.AppendUserLog "VALIDATION COMPLETE"
    AV_UI.AppendUserLog "==============================================="

Cleanup:
    AV_Core.BulkValidationInProgress = False
    AV_Core.ClearTableCache
    AV_Core.ClearAutoValidationMapCache
    Set CurrentTargetTable = Nothing
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    AV_Core.DebugMessage "RunFullValidationMaster completed at " & Now, MODULE_NAME
    Exit Sub

ErrHandler:
    AV_UI.AppendUserLog "ERROR in RunFullValidationMaster"
    AV_UI.AppendUserLog "Error #" & Err.Number & ": " & Err.description
    AV_Core.DebugMessage "ERROR #" & Err.Number & ": " & Err.description, MODULE_NAME
    Resume Cleanup
End Sub
' ======================================================
' PROCESS SINGLE VALIDATION TARGET
' ======================================================
Private Sub ProcessValidationTarget(target As AV_Core.ValidationTarget, wsConfig As Worksheet, english As Boolean)
    
    Dim wsTarget As Worksheet
    Dim tblTarget As ListObject
    Dim AdvFunctionMap As Object
    Dim FormatMap As Object
    Dim colMetaDict As Object
    Dim rowNum As Long, i As Long
    Dim keyRows() As Long
    Dim keyCount As Long
    Dim keyColIndex As Long
    
    On Error GoTo ErrHandler
    
    AV_UI.AppendUserLog "Processing target: " & target.tableName
    
    Set tblTarget = FindTableByName(target.tableName)
    
    If tblTarget Is Nothing Then
        AV_UI.AppendUserLog "  ERROR: Table not found: " & target.tableName
        Exit Sub
    End If
    
    Set CurrentTargetTable = tblTarget
    Set wsTarget = tblTarget.Parent
    
    AV_UI.AppendUserLog "  Table: " & target.tableName & " (Rows: " & tblTarget.ListRows.Count & ")"
    
    On Error Resume Next
    keyColIndex = tblTarget.ListColumns(target.KeyColumnHeader).Index
    On Error GoTo 0
    
    If keyColIndex = 0 Then
        AV_UI.AppendUserLog "  ERROR: Key column not found: " & target.KeyColumnHeader
        Exit Sub
    End If
    
    AV_UI.AppendUserLog "  Key column: " & target.KeyColumnHeader & " (Index: " & keyColIndex & ")"
    
    Set AdvFunctionMap = AV_Core.GetAutoValidationMap(wsConfig)
    Set FormatMap = AV_Format.LoadFormatMap(wsConfig)
    Set colMetaDict = AV_Core.GetDDMValidationColumns(wsConfig)
    
    If AdvFunctionMap Is Nothing Or AdvFunctionMap.Count = 0 Then
        AV_UI.AppendUserLog "  WARNING: No validation functions mapped."
    Else
        AV_UI.AppendUserLog "  Advanced validations loaded: " & AdvFunctionMap.Count
    End If
    
    If colMetaDict Is Nothing Or colMetaDict.Count = 0 Then
        AV_UI.AppendUserLog "  WARNING: No simple validations mapped."
    Else
        AV_UI.AppendUserLog "  Simple validations loaded: " & colMetaDict.Count
    End If
    
    PrintHeaderDiagnostics tblTarget, AdvFunctionMap
    
    AV_UI.AppendUserLog "  Identifying rows to validate..."
    
    Dim startRow As Long, endRow As Long
    startRow = tblTarget.DataBodyRange.Row
    endRow = startRow + tblTarget.ListRows.Count - 1
    
    ReDim keyRows(1 To tblTarget.ListRows.Count)
    keyCount = 0
    
    For rowNum = startRow To endRow
        Dim keyVal As String
        keyVal = Trim(CStr(tblTarget.DataBodyRange(rowNum - startRow + 1, keyColIndex).value))
        
        If Len(keyVal) > 0 Then
            If AV_Core.ShouldValidateRow(rowNum, wsTarget, tblTarget, True) Then
                keyCount = keyCount + 1
                keyRows(keyCount) = rowNum
            End If
        End If
    Next rowNum
    
    If keyCount = 0 Then
        AV_UI.AppendUserLog "  No rows identified for validation."
        Exit Sub
    End If
    
    ReDim Preserve keyRows(1 To keyCount)
    AV_UI.AppendUserLog "  Rows identified: " & keyCount
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    AV_UI.AppendUserLog "  Beginning row validation..."
    
    For i = 1 To keyCount
        rowNum = keyRows(i)
        
        If i Mod 10 = 0 Then
            DoEvents
            AV_UI.AppendUserLog "  Progress: " & i & " / " & keyCount & " rows processed"
        End If
        
        If AV_Core.ValidationCancelFlag Then
            AV_UI.AppendUserLog "  Validation cancelled."
            Exit For
        End If
        
        If AV_Core.ValidationTimeoutReached() Then
            AV_UI.AppendUserLog "  Validation timeout."
            Exit For
        End If
        
        ValidateSingleRow tblTarget, rowNum, AdvFunctionMap, english, FormatMap
    Next i
    
    AV_UI.AppendUserLog "  Row validation complete for " & target.tableName
    AV_UI.SetAdvancedValidationCompleted True
    
    AV_UI.AppendUserLog "  Running simple dropdown validation..."
    RunAutoCheckDataValidation tblTarget, keyRows, keyColIndex, english, FormatMap, colMetaDict
    AV_UI.AppendUserLog "  Simple validation complete: " & keyCount & " rows processed"
    AV_UI.SetLegacyMenuValidationCompleted True
    
    AV_UI.AppendUserLog "  Target validation complete: " & target.tableName
    Exit Sub

ErrHandler:
    AV_UI.AppendUserLog "  ERROR processing target " & target.tableName
    AV_UI.AppendUserLog "  Error #" & Err.Number & ": " & Err.description
    AV_Core.DebugMessage "ProcessValidationTarget ERROR: " & Err.description, MODULE_NAME
End Sub

' ======================================================
' PRINT HEADER DIAGNOSTICS
' ======================================================
Private Sub PrintHeaderDiagnostics(tblTarget As ListObject, AdvFunctionMap As Object)
    Dim funcKey As Variant
    Dim mapItem As Object
    Dim headerName As String
    Dim colIndex As Long
    
    AV_Core.DebugMessage "-----------------------------------------------", MODULE_NAME
    AV_Core.DebugMessage "DIAGNOSTIC: Header Mapping Check", MODULE_NAME
    AV_Core.DebugMessage "-----------------------------------------------", MODULE_NAME
    
    For Each funcKey In AdvFunctionMap.Keys
        Set mapItem = AdvFunctionMap(funcKey)
        headerName = mapItem("ColumnRef")
        
        colIndex = 0
        On Error Resume Next
        colIndex = tblTarget.ListColumns(headerName).Index
        On Error GoTo 0
        
        If colIndex > 0 Then
            AV_Core.DebugMessage "OK: " & funcKey & " -> '" & headerName & "' found at index " & colIndex, MODULE_NAME
        Else
            AV_Core.DebugMessage "MISSING: " & funcKey & " -> '" & headerName & "' NOT in table", MODULE_NAME
        End If
    Next funcKey
End Sub
' ======================================================
' VALIDATE SINGLE ROW (Header-Based)
' ======================================================
Public Sub ValidateSingleRow(tblTarget As ListObject, rowNum As Long, AdvFunctionMap As Object, english As Boolean, FormatMap As Object)
    Dim funcKey As Variant
    Dim funcName As String
    Dim TargetCell As Range
    Dim mapItem As Object
    Dim AutoValidate As Boolean
    Dim targetHeaderName As String
    Dim colIndex As Long

    For Each funcKey In AdvFunctionMap.Keys
        Set mapItem = AdvFunctionMap(funcKey)
        funcName = CStr(funcKey)

        AutoValidate = False
        If mapItem.Exists("AutoValidate") Then
            AutoValidate = mapItem("AutoValidate")
        End If
        
        targetHeaderName = ""
        If mapItem.Exists("ColumnRef") Then
            targetHeaderName = CStr(mapItem("ColumnRef"))
        End If
        
        If Len(targetHeaderName) = 0 Then
            AV_Core.DebugMessage "WARNING: Missing ColumnRef for " & funcName, MODULE_NAME
            GoTo SkipToNext
        End If
        
        If AutoValidate = False Then
            GoTo SkipToNext
        End If

        colIndex = 0
        On Error Resume Next
        colIndex = tblTarget.ListColumns(targetHeaderName).Index
        On Error GoTo 0
        
        If colIndex = 0 Then
            AV_Core.DebugMessage "Column not found: " & targetHeaderName & " for " & funcName, MODULE_NAME
            GoTo SkipToNext
        End If
        
        Set TargetCell = AV_Core.GetCellByHeader(tblTarget, rowNum, targetHeaderName)
        
        If Not TargetCell Is Nothing Then
            On Error GoTo ValidationError
            Application.Run funcName, TargetCell, tblTarget.Parent.Name, english, FormatMap, AdvFunctionMap
            On Error GoTo 0
        End If
        
SkipToNext:
    Next funcKey

    Exit Sub

ValidationError:
    AV_Core.DebugMessage "[ValidateSingleRow] Error: Row " & rowNum & " - " & funcName & " - " & Err.description, MODULE_NAME
    Resume SkipToNext
End Sub

' ======================================================
' RUN AUTO CHECK DATA VALIDATION (Header-Based)
' ======================================================
Public Sub RunAutoCheckDataValidation(tblTarget As ListObject, _
                                     keyRows() As Long, _
                                     keyColNum As Long, _
                                     english As Boolean, _
                                     Optional FormatMap As Object, _
                                     Optional colMetaDict As Object)

    On Error GoTo ErrHandler

    Dim meta As Object
    Dim colKey As Variant
    Dim i As Long, rowNum As Long
    Dim cellValue As String
    Dim found As Boolean
    Dim progressCount As Long
    Dim totalValid As Long
    Dim wsTarget As Worksheet
    
    Dim dropColMsgs As Object
    Dim dropColKey As Variant
    Dim cMsgErrorType As String
    Dim cMsg As Variant
    Dim msgArr(1 To 3) As Variant
    Dim DCMsgTxt As String

    Set wsTarget = tblTarget.Parent

    If tblTarget Is Nothing Then Exit Sub
    If LBound(keyRows) > UBound(keyRows) Then Exit Sub
    If colMetaDict Is Nothing Or colMetaDict.Count = 0 Then Exit Sub

    totalValid = UBound(keyRows) - LBound(keyRows) + 1
    AV_Core.DebugMessage "[RunAutoCheckDataValidation] Starting on " & totalValid & " rows.", MODULE_NAME

    For i = LBound(keyRows) To UBound(keyRows)
        rowNum = keyRows(i)
        
        Set dropColMsgs = CreateObject("Scripting.Dictionary")
        
        For Each colKey In colMetaDict.Keys
            Set meta = colMetaDict(colKey)
            
            If Not meta.Exists("TargetHeaderName") Then GoTo SkipCol
            If Not meta.Exists("ValidColumnListEN") Then meta("ValidColumnListEN") = Array()
            If Not meta.Exists("ValidColumnListFR") Then meta("ValidColumnListFR") = Array()
            If Not meta.Exists("ColumnNameEN") Then meta("ColumnNameEN") = ""
            If Not meta.Exists("ColumnNameFR") Then meta("ColumnNameFR") = ""
            If Not meta.Exists("CommentDropCol") Then meta("CommentDropCol") = ""
            
            Dim targetHeaderName As String
            targetHeaderName = meta("TargetHeaderName")
            
            Dim TargetCell As Range
            Set TargetCell = AV_Core.GetCellByHeader(tblTarget, rowNum, targetHeaderName)
            
            If TargetCell Is Nothing Then GoTo SkipCol
            
            cellValue = Trim(CStr(TargetCell.value))
            If Len(cellValue) = 0 Then GoTo SkipCol

            found = False
            If IsArray(meta("ValidColumnListEN")) Then found = ExistsInArray(meta("ValidColumnListEN"), cellValue)
            If Not found And IsArray(meta("ValidColumnListFR")) Then found = ExistsInArray(meta("ValidColumnListFR"), cellValue)

            If Not found Then
                Dim errorMsg As String
                If english Then
                    errorMsg = meta("ColumnNameEN") & " - Invalid value '" & cellValue & "' : Select a valid value from the list."
                Else
                    errorMsg = meta("ColumnNameFR") & " - Valeur invalide '" & cellValue & "' . Selectionner une valeur valide."
                End If

                If Not dropColMsgs.Exists(meta("CommentDropCol")) Then
                    Set dropColMsgs(meta("CommentDropCol")) = CreateObject("Scripting.Dictionary")
                End If
                
                msgArr(1) = targetHeaderName
                msgArr(2) = errorMsg
                msgArr(3) = "Error"
                dropColMsgs(meta("CommentDropCol")).Add dropColMsgs(meta("CommentDropCol")).Count + 1, msgArr
            Else
                If AV_Format.getFormatType(TargetCell, FormatMap) = "Error" Then
                    If Not dropColMsgs.Exists(meta("CommentDropCol")) Then
                        Set dropColMsgs(meta("CommentDropCol")) = CreateObject("Scripting.Dictionary")
                    End If
                    
                    msgArr(1) = targetHeaderName
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
                AV_Format.WriteSystemTagToDropColumn tblTarget, CStr(dropColKey), rowNum, CStr(cMsg(1)), DCMsgTxt, cMsgErrorType, FormatMap
            Next cMsg
        Next dropColKey
        
        FormatKeyCell tblTarget, rowNum, keyColNum, FormatMap
        
        progressCount = progressCount + 1
        If progressCount Mod 10 = 0 Then DoEvents
        If progressCount Mod 25 = 0 Then AV_UI.AppendUserLog "  [Simple] Progress: " & progressCount & " / " & totalValid

SkipRow:
    Next i

    AV_Core.DebugMessage "[RunAutoCheckDataValidation] Completed.", MODULE_NAME
    Exit Sub

ErrHandler:
    AV_Core.DebugMessage "RunAutoCheckDataValidation ERROR: " & Err.Number & " - " & Err.description, MODULE_NAME
    AV_UI.AppendUserLog "RunAutoCheckDataValidation ERROR: " & Err.Number & " - " & Err.description
End Sub

' ======================================================
' FORMAT KEY CELL
' ======================================================
Private Sub FormatKeyCell(tblTarget As ListObject, rowNum As Long, keyColIndex As Long, FormatMap As Object)
    Dim KeyCell As Range
    Dim tableRow As Long
    
    tableRow = rowNum - tblTarget.DataBodyRange.Row + 1
    If tableRow < 1 Or tableRow > tblTarget.ListRows.Count Then Exit Sub
    
    Set KeyCell = tblTarget.DataBodyRange(tableRow, keyColIndex)
    
    Dim rowRange As Range
    Set rowRange = tblTarget.ListRows(tableRow).Range
    
    AV_Format.FormatKeyCell rowRange, FormatMap
End Sub

' ======================================================
' FIND TABLE BY NAME
' ======================================================
Private Function FindTableByName(tableName As String) As ListObject
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            If tbl.Name = tableName Then
                Set FindTableByName = tbl
                Exit Function
            End If
        Next tbl
    Next ws
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
    On Error GoTo 0
End Function
