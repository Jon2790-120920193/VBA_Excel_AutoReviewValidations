Attribute VB_Name = "Test3_RunFullValidation"
Option Explicit

' ======================================================
' TEST 3: RUNFULLVALIDATIONMASTER TEST
' Tests end-to-end validation workflow
' Version: Phase 2 - Test 3
' Date: 2026-01-19
' ======================================================

Public Sub Test3_RunAll()
    Debug.Print "=========================================="
    Debug.Print "TEST 3: RUNFULLVALIDATIONMASTER TEST"
    Debug.Print "Phase 2 Testing - Version 2.1"
    Debug.Print "=========================================="
    Debug.Print ""
    
    Debug.Print "This test runs complete end-to-end validation."
    Debug.Print ""
    Debug.Print "STEPS:"
    Debug.Print "1. Run Test3_PreFlightCheck - Verify setup"
    Debug.Print "2. Run Test3_RunValidation - Execute full validation"
    Debug.Print "3. Review results in ValidationTrackerForm and Immediate Window"
    Debug.Print ""
    Debug.Print "=========================================="
End Sub

' ======================================================
' PRE-FLIGHT CHECK
' ======================================================
Public Sub Test3_PreFlightCheck()
    Debug.Print ""
    Debug.Print "=========================================="
    Debug.Print "TEST 3 - PRE-FLIGHT CHECK"
    Debug.Print "=========================================="
    Debug.Print ""
    
    Dim allGood As Boolean
    allGood = True
    
    ' Check 1: GlobalDebugOn setting
    Debug.Print "Check 1: GlobalDebugOn setting..."
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim setting As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Config")
    Set tbl = ws.ListObjects("GlobalDebugOptions")
    On Error GoTo 0
    
    If tbl Is Nothing Then
        Debug.Print "  ❌ FAIL: GlobalDebugOptions table not found"
        allGood = False
    Else
        setting = UCase(Trim(tbl.DataBodyRange(1, 1).value))
        If setting = "ON" Then
            Debug.Print "  ✅ PASS: GlobalDebugOn = ON"
        Else
            Debug.Print "  ⚠️  WARNING: GlobalDebugOn = " & setting & " (expected ON)"
            Debug.Print "     Set to ON for detailed test logging"
        End If
    End If
    
    ' Check 2: AutoValidationMap table
    Debug.Print ""
    Debug.Print "Check 2: AutoValidationCommentPrefixMappingTable..."
    On Error Resume Next
    Set tbl = ws.ListObjects("AutoValidationCommentPrefixMappingTable")
    On Error GoTo 0
    
    If tbl Is Nothing Then
        Debug.Print "  ❌ FAIL: AutoValidationCommentPrefixMappingTable not found"
        allGood = False
    Else
        Debug.Print "  ✅ PASS: Table found with " & tbl.ListRows.Count & " validation functions"
    End If
    
    ' Check 3: AutoFormatOnFullValidation table
    Debug.Print ""
    Debug.Print "Check 3: AutoFormatOnFullValidation table..."
    On Error Resume Next
    Set tbl = ws.ListObjects("AutoFormatOnFullValidation")
    On Error GoTo 0
    
    If tbl Is Nothing Then
        Debug.Print "  ❌ FAIL: AutoFormatOnFullValidation table not found"
        allGood = False
    Else
        Debug.Print "  ✅ PASS: Table found with " & tbl.ListRows.Count & " format types"
    End If
    
    ' Check 4: Validation rule tables
    Debug.Print ""
    Debug.Print "Check 4: Validation rule tables..."
    Dim ruleTables As Variant
    Dim ruleTable As Variant
    Dim missingTables As String
    
    ruleTables = Array("GIWValidationTable", "ElectricityPairValidation", _
                       "PlumbingPairValidation", "HeatSourcePairValidation", _
                       "HeatSourceANYRefTable")
    
    missingTables = ""
    For Each ruleTable In ruleTables
        On Error Resume Next
        Set tbl = ws.ListObjects(CStr(ruleTable))
        On Error GoTo 0
        
        If tbl Is Nothing Then
            missingTables = missingTables & "  - " & ruleTable & vbCrLf
        Else
            Debug.Print "  ✅ " & ruleTable & " (" & tbl.ListRows.Count & " rules)"
        End If
    Next ruleTable
    
    If missingTables <> "" Then
        Debug.Print "  ❌ FAIL: Missing tables:" & vbCrLf & missingTables
        allGood = False
    End If
    
    ' Check 5: Target data sheet
    Debug.Print ""
    Debug.Print "Check 5: Target data sheet..."
    ' Note: This assumes legacy B3 cell reference - will update in Phase 3
    Dim targetSheet As String
    On Error Resume Next
    targetSheet = Trim(ws.Range("B3").value)
    On Error GoTo 0
    
    If targetSheet = "" Then
        Debug.Print "  ⚠️  WARNING: No target sheet specified in B3"
    Else
        Dim wsTarget As Worksheet
        On Error Resume Next
        Set wsTarget = ThisWorkbook.Sheets(targetSheet)
        On Error GoTo 0
        
        If wsTarget Is Nothing Then
            Debug.Print "  ❌ FAIL: Target sheet '" & targetSheet & "' not found"
            allGood = False
        Else
            Debug.Print "  ✅ PASS: Target sheet '" & targetSheet & "' exists"
            
            ' Check for data
            Dim startRow As Long, rowCount As Long
            On Error Resume Next
            startRow = CLng(ws.Range("B4").value)
            rowCount = CLng(ws.Range("D4").value)
            On Error GoTo 0
            
            If startRow > 0 And rowCount > 0 Then
                Debug.Print "     Data range: Row " & startRow & " to " & (startRow + rowCount)
                Debug.Print "     Total rows: " & rowCount
            End If
        End If
    End If
    
    ' Summary
    Debug.Print ""
    Debug.Print "=========================================="
    If allGood Then
        Debug.Print "PRE-FLIGHT CHECK: ✅ ALL SYSTEMS GO"
        Debug.Print ""
        Debug.Print "Ready to run Test3_RunValidation"
    Else
        Debug.Print "PRE-FLIGHT CHECK: ❌ ISSUES FOUND"
        Debug.Print ""
        Debug.Print "Fix issues above before running validation"
    End If
    Debug.Print "=========================================="
End Sub

' ======================================================
' RUN VALIDATION TEST
' ======================================================
Public Sub Test3_RunValidation()
    Debug.Print ""
    Debug.Print "=========================================="
    Debug.Print "TEST 3 - RUNNING FULL VALIDATION"
    Debug.Print "=========================================="
    Debug.Print ""
    
    Debug.Print "Starting validation at " & Format(Now, "hh:mm:ss")
    Debug.Print "Watch ValidationTrackerForm for progress..."
    Debug.Print ""
    Debug.Print "---[VALIDATION OUTPUT BELOW]---"
    Debug.Print ""
    
    On Error GoTo ErrHandler
    
    ' Run the validation
    Dim startTime As Single
    startTime = Timer
    
    AV_Engine.RunFullValidationMaster
    
    Dim elapsed As Single
    elapsed = Timer - startTime
    
    Debug.Print ""
    Debug.Print "---[VALIDATION OUTPUT ABOVE]---"
    Debug.Print ""
    Debug.Print "Validation completed successfully"
    Debug.Print "Elapsed time: " & Format(elapsed, "0.00") & " seconds"
    Debug.Print ""
    Debug.Print "=========================================="
    Debug.Print "TEST 3 COMPLETE"
    Debug.Print ""
    Debug.Print "NEXT STEPS:"
    Debug.Print "1. Check ValidationTrackerForm for completion status"
    Debug.Print "2. Review target sheet for validation results"
    Debug.Print "3. Verify error/autocorrect formatting applied"
    Debug.Print "4. Check drop columns for validation messages"
    Debug.Print "=========================================="
    
    Exit Sub
    
ErrHandler:
    Debug.Print ""
    Debug.Print "=========================================="
    Debug.Print "❌ ERROR DURING VALIDATION"
    Debug.Print "=========================================="
    Debug.Print "Error #" & Err.Number & ": " & Err.Description
    Debug.Print "Source: " & Err.Source
    Debug.Print ""
    Debug.Print "Check ValidationTrackerForm for additional details"
    Debug.Print "=========================================="
End Sub

' ======================================================
' QUICK VALIDATION (SMALL SAMPLE)
' ======================================================
Public Sub Test3_QuickValidation()
    Debug.Print ""
    Debug.Print "=========================================="
    Debug.Print "TEST 3 - QUICK VALIDATION (SMALL SAMPLE)"
    Debug.Print "=========================================="
    Debug.Print ""
    
    ' Save current row count
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Config")
    Dim originalRowCount As Long
    On Error Resume Next
    originalRowCount = CLng(ws.Range("D4").value)
    On Error GoTo 0
    
    If originalRowCount = 0 Then
        Debug.Print "ERROR: No row count in D4"
        Exit Sub
    End If
    
    ' Temporarily set to validate only 10 rows
    ws.Range("D4").value = 10
    Debug.Print "Temporarily validating only 10 rows..."
    Debug.Print ""
    
    ' Run validation
    Test3_RunValidation
    
    ' Restore original row count
    ws.Range("D4").value = originalRowCount
    Debug.Print ""
    Debug.Print "Row count restored to: " & originalRowCount
End Sub

' ======================================================
' UTILITY: CHECK MODULE VERSIONS
' ======================================================
Public Sub Test3_CheckModuleVersions()
    Debug.Print ""
    Debug.Print "=========================================="
    Debug.Print "MODULE VERSION CHECK"
    Debug.Print "=========================================="
    Debug.Print ""
    
    Debug.Print "This checks if Test 2 versions are imported:"
    Debug.Print ""
    
    ' Check if GlobalDebugOn works (indicates v2.1 COMPLETE)
    AV_Core.InitDebugFlags True
    Debug.Print "AV_Core: GlobalDebugOn = " & AV_Core.GlobalDebugOn
    If AV_Core.GlobalDebugOn Then
        Debug.Print "  ✅ AV_Core v2.1 COMPLETE confirmed"
    Else
        Debug.Print "  ⚠️  Check AV_Core version (GlobalDebugOn should be True)"
    End If
    
    Debug.Print ""
    Debug.Print "Expected modules for Test 3:"
    Debug.Print "- AV_Core v2.1 COMPLETE (Test 2 version)"
    Debug.Print "- AV_UI v2.1 Test2 (Test 2 version)"
    Debug.Print "- AV_Engine v2.1 (Phase 2 version)"
    Debug.Print "- AV_Format v2.1 (Phase 2 version)"
    Debug.Print "- AV_Validators v2.1 (Phase 2 version)"
    Debug.Print "- AV_ValidationRules v2.1 (Phase 2 version)"
    Debug.Print ""
    Debug.Print "If using older versions, import latest from project files"
    Debug.Print "=========================================="
End Sub
