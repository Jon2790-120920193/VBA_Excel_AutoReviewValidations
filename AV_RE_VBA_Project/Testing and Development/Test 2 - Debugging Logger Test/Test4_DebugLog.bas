Attribute VB_Name = "Test2_DebugLogger"
Option Explicit

' ======================================================
' TEST 2: DEBUG LOGGER TEST
' Tests GlobalDebugOn functionality
' Version: Phase 2 - Test 2
' Date: 2026-01-19
' ======================================================

Public Sub Test2_RunAll()
    Debug.Print "=========================================="
    Debug.Print "TEST 2: DEBUG LOGGER TEST"
    Debug.Print "Phase 2 Testing - Version 2.1"
    Debug.Print "=========================================="
    Debug.Print ""
    
    Debug.Print "This test verifies that GlobalDebugOn controls debug logging."
    Debug.Print ""
    Debug.Print "STEPS:"
    Debug.Print "1. Run Test2_Part1_GlobalDebugON"
    Debug.Print "2. Set GlobalDebugOn to 'OFF' in table"
    Debug.Print "3. Run Test2_Part2_GlobalDebugOFF"
    Debug.Print ""
    Debug.Print "=========================================="
End Sub

' ======================================================
' PART 1: TEST WITH GLOBALDEBUMON = "ON"
' ======================================================
Public Sub Test2_Part1_GlobalDebugON()
    Debug.Print ""
    Debug.Print "=========================================="
    Debug.Print "TEST 2 - PART 1: GlobalDebugOn = ON"
    Debug.Print "=========================================="
    Debug.Print ""
    
    ' Check current setting
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim currentSetting As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Config")
    Set tbl = ws.ListObjects("GlobalDebugOptions")
    On Error GoTo 0
    
    If tbl Is Nothing Then
        Debug.Print "ERROR: GlobalDebugOptions table not found"
        Exit Sub
    End If
    
    currentSetting = UCase(Trim(tbl.DataBodyRange(1, 1).Value))
    Debug.Print "Current setting in table: " & currentSetting
    
    If currentSetting <> "ON" Then
        Debug.Print ""
        Debug.Print "WARNING: GlobalDebugOn is set to '" & currentSetting & "'"
        Debug.Print "Please set it to 'ON' in the GlobalDebugOptions table first."
        Debug.Print ""
        Exit Sub
    End If
    
    Debug.Print ""
    Debug.Print "Step 1: Initializing debug flags..."
    AV_Core.InitDebugFlags True  ' Force reload
    
    Debug.Print "Step 2: Checking AV_Core.GlobalDebugOn variable..."
    Debug.Print "  AV_Core.GlobalDebugOn = " & AV_Core.GlobalDebugOn
    
    If Not AV_Core.GlobalDebugOn Then
        Debug.Print "  ERROR: GlobalDebugOn is False but table says 'ON'"
        Debug.Print "  InitDebugFlags may not be reading table correctly"
        Exit Sub
    End If
    
    Debug.Print ""
    Debug.Print "Step 3: Testing DebugMessage function..."
    Debug.Print "EXPECTED: You should see [DEBUG] messages below"
    Debug.Print "---"
    
    AV_Core.DebugMessage "Test message 1 - This should appear", "Test2_Part1"
    AV_Core.DebugMessage "Test message 2 - This should also appear", "AV_Core"
    AV_Core.DebugMessage "Test message 3 - This one too", "AV_Engine"
    
    Debug.Print "---"
    Debug.Print ""
    Debug.Print "Step 4: Testing GetAutoValidationMap logging..."
    Debug.Print "(This should show detailed row-by-row processing)"
    Debug.Print ""
    
    ' Clear cache to force reload
    AV_Core.ClearAutoValidationMapCache
    
    ' Load map - should show debug messages
    Dim map As Object
    Set map = AV_Core.GetAutoValidationMap()
    
    Debug.Print ""
    Debug.Print "Map loaded: " & map.Count & " items"
    Debug.Print ""
    Debug.Print "=========================================="
    Debug.Print "TEST 2 - PART 1 COMPLETE"
    Debug.Print ""
    Debug.Print "EXPECTED RESULTS:"
    Debug.Print "- [DEBUG] messages appeared"
    Debug.Print "- Row processing messages appeared"
    Debug.Print ""
    Debug.Print "If you saw debug messages, PART 1 PASSED"
    Debug.Print "Now set GlobalDebugOn to 'OFF' and run Part 2"
    Debug.Print "=========================================="
End Sub

' ======================================================
' PART 2: TEST WITH GLOBALDEBUGUN = "OFF"
' ======================================================
Public Sub Test2_Part2_GlobalDebugOFF()
    Debug.Print ""
    Debug.Print "=========================================="
    Debug.Print "TEST 2 - PART 2: GlobalDebugOn = OFF"
    Debug.Print "=========================================="
    Debug.Print ""
    
    ' Check current setting
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim currentSetting As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Config")
    Set tbl = ws.ListObjects("GlobalDebugOptions")
    On Error GoTo 0
    
    If tbl Is Nothing Then
        Debug.Print "ERROR: GlobalDebugOptions table not found"
        Exit Sub
    End If
    
    currentSetting = UCase(Trim(tbl.DataBodyRange(1, 1).Value))
    Debug.Print "Current setting in table: " & currentSetting
    
    If currentSetting <> "OFF" Then
        Debug.Print ""
        Debug.Print "WARNING: GlobalDebugOn is set to '" & currentSetting & "'"
        Debug.Print "Please set it to 'OFF' in the GlobalDebugOptions table first."
        Debug.Print ""
        Exit Sub
    End If
    
    Debug.Print ""
    Debug.Print "Step 1: Initializing debug flags..."
    AV_Core.InitDebugFlags True  ' Force reload
    
    Debug.Print "Step 2: Checking AV_Core.GlobalDebugOn variable..."
    Debug.Print "  AV_Core.GlobalDebugOn = " & AV_Core.GlobalDebugOn
    
    If AV_Core.GlobalDebugOn Then
        Debug.Print "  ERROR: GlobalDebugOn is True but table says 'OFF'"
        Debug.Print "  InitDebugFlags may not be reading table correctly"
        Exit Sub
    End If
    
    Debug.Print ""
    Debug.Print "Step 3: Testing DebugMessage function..."
    Debug.Print "EXPECTED: NO [DEBUG] messages should appear below"
    Debug.Print "---"
    
    AV_Core.DebugMessage "Test message 1 - This should NOT appear", "Test2_Part2"
    AV_Core.DebugMessage "Test message 2 - This should NOT appear", "AV_Core"
    AV_Core.DebugMessage "Test message 3 - This should NOT appear", "AV_Engine"
    
    Debug.Print "---"
    Debug.Print "(If you see NO messages between the dashes above, that's correct)"
    Debug.Print ""
    
    Debug.Print "Step 4: Testing GetAutoValidationMap logging..."
    Debug.Print "(Should show minimal logging - only progress in UserForm)"
    Debug.Print ""
    
    ' Clear cache to force reload
    AV_Core.ClearAutoValidationMapCache
    
    ' Load map - should NOT show debug messages
    Dim map As Object
    Set map = AV_Core.GetAutoValidationMap()
    
    Debug.Print ""
    Debug.Print "Map loaded: " & map.Count & " items"
    Debug.Print "(No row-by-row debug messages should have appeared)"
    Debug.Print ""
    Debug.Print "=========================================="
    Debug.Print "TEST 2 - PART 2 COMPLETE"
    Debug.Print ""
    Debug.Print "EXPECTED RESULTS:"
    Debug.Print "- NO [DEBUG] messages appeared"
    Debug.Print "- NO row processing messages appeared"
    Debug.Print "- Only this test output visible"
    Debug.Print ""
    Debug.Print "If you saw NO debug messages, PART 2 PASSED"
    Debug.Print "=========================================="
End Sub

' ======================================================
' UTILITY: DISPLAY CURRENT DEBUG STATUS
' ======================================================
Public Sub Test2_ShowStatus()
    Debug.Print ""
    Debug.Print "=========================================="
    Debug.Print "CURRENT DEBUG STATUS"
    Debug.Print "=========================================="
    Debug.Print ""
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim tableSetting As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Config")
    Set tbl = ws.ListObjects("GlobalDebugOptions")
    On Error GoTo 0
    
    If tbl Is Nothing Then
        Debug.Print "ERROR: GlobalDebugOptions table not found"
        Exit Sub
    End If
    
    tableSetting = UCase(Trim(tbl.DataBodyRange(1, 1).Value))
    
    Debug.Print "Table Setting: " & tableSetting
    Debug.Print "Variable Value: " & AV_Core.GlobalDebugOn
    Debug.Print "Initialized: " & (AV_Core.DebugFlags Is Nothing = False)
    Debug.Print ""
    
    If tableSetting = "ON" And AV_Core.GlobalDebugOn Then
        Debug.Print "Status: SYNCHRONIZED - Debug logging ENABLED"
    ElseIf tableSetting = "OFF" And Not AV_Core.GlobalDebugOn Then
        Debug.Print "Status: SYNCHRONIZED - Debug logging DISABLED"
    Else
        Debug.Print "WARNING: MISMATCH between table and variable"
        Debug.Print "Run AV_Core.InitDebugFlags(True) to synchronize"
    End If
    
    Debug.Print "=========================================="
End Sub
