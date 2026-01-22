Attribute VB_Name = "Test5_Debug450"
Option Explicit

' ======================================================
' Test5_Debug450
' Diagnostic to find the exact location of Error #450
' ======================================================

Public Sub Test5_FindError450()
    Debug.Print "=========================================="
    Debug.Print "TEST 5: DEBUGGING ERROR #450"
    Debug.Print "=========================================="
    Debug.Print ""
    
    ' Test 1: Check AV_Core module exists and has BulkValidationInProgress
    Debug.Print "Step 1: Testing AV_Core.BulkValidationInProgress..."
    On Error Resume Next
    AV_Core.BulkValidationInProgress = True
    If Err.Number <> 0 Then
        Debug.Print "  ERROR #" & Err.Number & ": " & Err.Description
        Debug.Print "  -> BulkValidationInProgress may not be declared in AV_Core"
        Err.Clear
    Else
        Debug.Print "  OK - Set to True"
        AV_Core.BulkValidationInProgress = False
        Debug.Print "  OK - Set to False"
    End If
    On Error GoTo 0
    Debug.Print ""
    
    ' Test 2: Check AV_Core.InitDebugFlags
    Debug.Print "Step 2: Testing AV_Core.InitDebugFlags..."
    On Error Resume Next
    AV_Core.InitDebugFlags
    If Err.Number <> 0 Then
        Debug.Print "  ERROR #" & Err.Number & ": " & Err.Description
        Err.Clear
    Else
        Debug.Print "  OK"
    End If
    On Error GoTo 0
    Debug.Print ""
    
    ' Test 3: Check AV_Core.DebugMessage
    Debug.Print "Step 3: Testing AV_Core.DebugMessage..."
    On Error Resume Next
    AV_Core.DebugMessage "Test message", "Test5_Debug450"
    If Err.Number <> 0 Then
        Debug.Print "  ERROR #" & Err.Number & ": " & Err.Description
        Err.Clear
    Else
        Debug.Print "  OK"
    End If
    On Error GoTo 0
    Debug.Print ""
    
    ' Test 4: Check AV_Core.ValidationStartTime
    Debug.Print "Step 4: Testing AV_Core.ValidationStartTime..."
    On Error Resume Next
    AV_Core.ValidationStartTime = Timer
    If Err.Number <> 0 Then
        Debug.Print "  ERROR #" & Err.Number & ": " & Err.Description
        Err.Clear
    Else
        Debug.Print "  OK - Value: " & AV_Core.ValidationStartTime
    End If
    On Error GoTo 0
    Debug.Print ""
    
    ' Test 5: Check AV_Core.ValidationCancelTimeout
    Debug.Print "Step 5: Testing AV_Core.ValidationCancelTimeout..."
    On Error Resume Next
    AV_Core.ValidationCancelTimeout = 10000
    If Err.Number <> 0 Then
        Debug.Print "  ERROR #" & Err.Number & ": " & Err.Description
        Err.Clear
    Else
        Debug.Print "  OK - Value: " & AV_Core.ValidationCancelTimeout
    End If
    On Error GoTo 0
    Debug.Print ""
    
    ' Test 6: Check AV_Core.ValidationCancelFlag
    Debug.Print "Step 6: Testing AV_Core.ValidationCancelFlag..."
    On Error Resume Next
    AV_Core.ValidationCancelFlag = False
    If Err.Number <> 0 Then
        Debug.Print "  ERROR #" & Err.Number & ": " & Err.Description
        Err.Clear
    Else
        Debug.Print "  OK - Value: " & AV_Core.ValidationCancelFlag
    End If
    On Error GoTo 0
    Debug.Print ""
    
    ' Test 7: Check AV_UI.ShowValidationTrackerForm
    Debug.Print "Step 7: Testing AV_UI.ShowValidationTrackerForm..."
    On Error Resume Next
    AV_UI.ShowValidationTrackerForm
    If Err.Number <> 0 Then
        Debug.Print "  ERROR #" & Err.Number & ": " & Err.Description
        Err.Clear
    Else
        Debug.Print "  OK - Form shown"
    End If
    On Error GoTo 0
    Debug.Print ""
    
    ' Test 8: Check AV_UI.AppendUserLog
    Debug.Print "Step 8: Testing AV_UI.AppendUserLog..."
    On Error Resume Next
    AV_UI.AppendUserLog "Test log message"
    If Err.Number <> 0 Then
        Debug.Print "  ERROR #" & Err.Number & ": " & Err.Description
        Err.Clear
    Else
        Debug.Print "  OK - Message logged"
    End If
    On Error GoTo 0
    Debug.Print ""
    
    ' Test 9: Check ValidationTargets table access
    Debug.Print "Step 9: Testing ValidationTargets table access..."
    Dim wsConfig As Worksheet
    Dim validationTargets As ListObject
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("Config")
    Set validationTargets = wsConfig.ListObjects("ValidationTargets")
    If Err.Number <> 0 Then
        Debug.Print "  ERROR #" & Err.Number & ": " & Err.Description
        Err.Clear
    ElseIf validationTargets Is Nothing Then
        Debug.Print "  WARNING: ValidationTargets table not found"
    Else
        Debug.Print "  OK - Found with " & validationTargets.ListRows.Count & " rows"
    End If
    On Error GoTo 0
    Debug.Print ""
    
    ' Test 10: Check GetAutoValidationMap
    Debug.Print "Step 10: Testing AV_Core.GetAutoValidationMap..."
    On Error Resume Next
    Dim autoValMap As Object
    Set autoValMap = AV_Core.GetAutoValidationMap()
    If Err.Number <> 0 Then
        Debug.Print "  ERROR #" & Err.Number & ": " & Err.Description
        Err.Clear
    ElseIf autoValMap Is Nothing Then
        Debug.Print "  WARNING: GetAutoValidationMap returned Nothing"
    Else
        Debug.Print "  OK - Found " & autoValMap.Count & " mappings"
    End If
    On Error GoTo 0
    Debug.Print ""
    
    ' Test 11: Check GetValidationTable
    Debug.Print "Step 11: Testing AV_Core.GetValidationTable..."
    On Error Resume Next
    Dim tbl As ListObject
    Set tbl = AV_Core.GetValidationTable("GIWValidationTable")
    If Err.Number <> 0 Then
        Debug.Print "  ERROR #" & Err.Number & ": " & Err.Description
        Err.Clear
    ElseIf tbl Is Nothing Then
        Debug.Print "  WARNING: GetValidationTable returned Nothing"
    Else
        Debug.Print "  OK - Found table with " & tbl.ListRows.Count & " rows"
    End If
    On Error GoTo 0
    Debug.Print ""
    
    ' Test 12: Check AV_Format.LoadFormatMap
    Debug.Print "Step 12: Testing AV_Format.LoadFormatMap..."
    On Error Resume Next
    Dim formatMap As Object
    Set formatMap = AV_Format.LoadFormatMap(wsConfig)
    If Err.Number <> 0 Then
        Debug.Print "  ERROR #" & Err.Number & ": " & Err.Description
        Err.Clear
    ElseIf formatMap Is Nothing Then
        Debug.Print "  WARNING: LoadFormatMap returned Nothing"
    Else
        Debug.Print "  OK - Found " & formatMap.Count & " formats"
    End If
    On Error GoTo 0
    Debug.Print ""
    
    Debug.Print "=========================================="
    Debug.Print "TEST 5 COMPLETE"
    Debug.Print "Review errors above to find the source of #450"
    Debug.Print "=========================================="
End Sub
