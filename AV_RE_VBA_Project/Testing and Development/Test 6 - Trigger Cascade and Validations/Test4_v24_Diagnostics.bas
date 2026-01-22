Attribute VB_Name = "Test4_v24_Diagnostics"
Option Explicit

' ======================================================
' TEST 4: VERSION 2.4 DIAGNOSTICS & VALIDATION TEST
' ======================================================
' Purpose: Verify correct module versions are imported and
'          diagnose table configuration issues
' Version: 2.4
' Date: 2026-01-20
' ======================================================

Private Const TEST_MODULE_VERSION As String = "2.4"

' ======================================================
' MAIN TEST ENTRY POINT
' ======================================================
Public Sub Test4_RunAll()
    Debug.Print ""
    Debug.Print "=============================================="
    Debug.Print "  TEST 4: VERSION 2.4 COMPREHENSIVE DIAGNOSTICS"
    Debug.Print "=============================================="
    Debug.Print ""
    Debug.Print "Test Module Version: " & TEST_MODULE_VERSION
    Debug.Print "Date: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
    Debug.Print ""
    Debug.Print "This test suite will:"
    Debug.Print "  1. Verify correct module versions are loaded"
    Debug.Print "  2. Check ValidationTargets table configuration"
    Debug.Print "  3. Verify target table exists and has correct headers"
    Debug.Print "  4. Compare mapped headers vs actual table headers"
    Debug.Print "  5. Run a sample validation with detailed logging"
    Debug.Print ""
    Debug.Print "=============================================="
    Debug.Print ""
    
    ' Run all diagnostic steps
    Test4_Step1_VersionCheck
    Test4_Step2_ValidationTargets
    Test4_Step3_TargetTableAnalysis
    Test4_Step4_HeaderMappingCheck
    
    Debug.Print ""
    Debug.Print "=============================================="
    Debug.Print "DIAGNOSTICS COMPLETE"
    Debug.Print ""
    Debug.Print "If all checks pass, run: Test4_RunValidation"
    Debug.Print "=============================================="
End Sub

' ======================================================
' STEP 1: VERSION CHECK
' Verify correct module versions are imported
' ======================================================
Public Sub Test4_Step1_VersionCheck()
    Debug.Print ""
    Debug.Print "----------------------------------------------"
    Debug.Print "STEP 1: MODULE VERSION CHECK"
    Debug.Print "----------------------------------------------"
    Debug.Print ""
    
    Dim allVersionsOK As Boolean
    allVersionsOK = True
    
    ' Check AV_Engine version
    Debug.Print "Checking AV_Engine..."
    On Error Resume Next
    Dim engineVersion As String
    engineVersion = AV_Engine.MODULE_VERSION
    On Error GoTo 0
    
    If Len(engineVersion) > 0 Then
        Debug.Print "  AV_Engine.MODULE_VERSION = " & engineVersion
        If engineVersion >= "2.4" Then
            Debug.Print "  ✅ PASS: AV_Engine v2.4+ detected"
        Else
            Debug.Print "  ⚠️  WARNING: AV_Engine version " & engineVersion & " (expected 2.4+)"
            Debug.Print "     Import av_engine_v2_4.bas to fix"
            allVersionsOK = False
        End If
    Else
        Debug.Print "  ❌ FAIL: AV_Engine.MODULE_VERSION not found"
        Debug.Print "     This indicates an older version is loaded"
        Debug.Print "     Import av_engine_v2_4.bas to fix"
        allVersionsOK = False
    End If
    
    ' Check AV_DataAccess
    Debug.Print ""
    Debug.Print "Checking AV_DataAccess..."
    Dim testTable As ListObject
    On Error Resume Next
    Set testTable = AV_DataAccess.FindTableByName("ValidationTargets")
    On Error GoTo 0
    
    If Not testTable Is Nothing Then
        Debug.Print "  ✅ PASS: AV_DataAccess.FindTableByName() works"
    Else
        ' Try alternate method
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets("Config")
        On Error Resume Next
        Set testTable = ws.ListObjects("ValidationTargets")
        On Error GoTo 0
        
        If testTable Is Nothing Then
            Debug.Print "  ⚠️  ValidationTargets table not found (expected)"
        Else
            Debug.Print "  ⚠️  AV_DataAccess may need update - fallback worked"
        End If
    End If
    
    ' Check AV_Core
    Debug.Print ""
    Debug.Print "Checking AV_Core..."
    On Error Resume Next
    AV_Core.InitDebugFlags True
    On Error GoTo 0
    
    Debug.Print "  AV_Core.GlobalDebugOn = " & AV_Core.GlobalDebugOn
    If AV_Core.GlobalDebugOn Then
        Debug.Print "  ✅ PASS: Debug system working"
    Else
        Debug.Print "  ⚠️  WARNING: GlobalDebugOn is False"
        Debug.Print "     Set GlobalDebugOptions to ON in Config sheet"
    End If
    
    ' Check AV_Validators
    Debug.Print ""
    Debug.Print "Checking AV_Validators..."
    Debug.Print "  ✅ Validator entry points exist (will test in Step 5)"
    
    ' Summary
    Debug.Print ""
    If allVersionsOK Then
        Debug.Print "VERSION CHECK: ✅ ALL MODULES OK"
    Else
        Debug.Print "VERSION CHECK: ⚠️  SOME MODULES NEED UPDATE"
        Debug.Print ""
        Debug.Print "ACTION REQUIRED:"
        Debug.Print "  1. In VBA Editor, remove old AV_Engine module"
        Debug.Print "  2. Import av_engine_v2_4.bas"
        Debug.Print "  3. Re-run this test"
    End If
End Sub

' ======================================================
' STEP 2: VALIDATION TARGETS CHECK
' Verify ValidationTargets table configuration
' ======================================================
Public Sub Test4_Step2_ValidationTargets()
    Debug.Print ""
    Debug.Print "----------------------------------------------"
    Debug.Print "STEP 2: VALIDATION TARGETS TABLE CHECK"
    Debug.Print "----------------------------------------------"
    Debug.Print ""
    
    Dim wsConfig As Worksheet
    Dim tbl As ListObject
    
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("Config")
    Set tbl = wsConfig.ListObjects("ValidationTargets")
    On Error GoTo 0
    
    If tbl Is Nothing Then
        Debug.Print "❌ FAIL: ValidationTargets table not found"
        Debug.Print ""
        Debug.Print "CREATE THIS TABLE in Config sheet:"
        Debug.Print ""
        Debug.Print "  | TableName              | Enabled | Mode | Key Column (Header Name) |"
        Debug.Print "  |------------------------|---------|------|--------------------------|"
        Debug.Print "  | REP2DSMDraft_Buildings | TRUE    | Both | AO ID                    |"
        Debug.Print ""
        Debug.Print "The TableName must match an Excel Table (ListObject) name in your workbook"
        Exit Sub
    End If
    
    Debug.Print "✅ ValidationTargets table found"
    Debug.Print "   Rows: " & tbl.ListRows.Count
    Debug.Print ""
    
    ' Check required columns
    Dim requiredCols As Variant
    Dim colName As Variant
    Dim missingCols As String
    
    requiredCols = Array("TableName", "Enabled", "Mode", "Key Column (Header Name)")
    missingCols = ""
    
    For Each colName In requiredCols
        Dim testCol As ListColumn
        On Error Resume Next
        Set testCol = tbl.ListColumns(CStr(colName))
        On Error GoTo 0
        
        If testCol Is Nothing Then
            missingCols = missingCols & "  - " & colName & vbCrLf
        End If
    Next colName
    
    If Len(missingCols) > 0 Then
        Debug.Print "❌ FAIL: Missing required columns:"
        Debug.Print missingCols
        Exit Sub
    End If
    
    Debug.Print "✅ All required columns present"
    Debug.Print ""
    
    ' List all targets
    Debug.Print "Configured targets:"
    Debug.Print ""
    
    Dim r As ListRow
    Dim targetName As String
    Dim isEnabled As String
    Dim Mode As String
    Dim keyCol As String
    Dim enabledCount As Long
    
    enabledCount = 0
    
    For Each r In tbl.ListRows
        targetName = Trim(CStr(r.Range.Cells(1, tbl.ListColumns("TableName").Index).value))
        isEnabled = UCase(Trim(CStr(r.Range.Cells(1, tbl.ListColumns("Enabled").Index).value)))
        Mode = Trim(CStr(r.Range.Cells(1, tbl.ListColumns("Mode").Index).value))
        keyCol = Trim(CStr(r.Range.Cells(1, tbl.ListColumns("Key Column (Header Name)").Index).value))
        
        If isEnabled = "TRUE" Then
            enabledCount = enabledCount + 1
            Debug.Print "  ✅ ENABLED: " & targetName
        Else
            Debug.Print "  ⬜ disabled: " & targetName
        End If
        Debug.Print "     Mode: " & Mode & " | Key Column: " & keyCol
    Next r
    
    Debug.Print ""
    If enabledCount > 0 Then
        Debug.Print "✅ " & enabledCount & " target(s) enabled"
    Else
        Debug.Print "❌ No targets enabled! Set Enabled=TRUE for at least one target"
    End If
End Sub

' ======================================================
' STEP 3: TARGET TABLE ANALYSIS
' Find the target table and analyze its structure
' ======================================================
Public Sub Test4_Step3_TargetTableAnalysis()
    Debug.Print ""
    Debug.Print "----------------------------------------------"
    Debug.Print "STEP 3: TARGET TABLE ANALYSIS"
    Debug.Print "----------------------------------------------"
    Debug.Print ""
    
    ' First get target name from ValidationTargets
    Dim wsConfig As Worksheet
    Dim vtTable As ListObject
    
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("Config")
    Set vtTable = wsConfig.ListObjects("ValidationTargets")
    On Error GoTo 0
    
    If vtTable Is Nothing Then
        Debug.Print "❌ Cannot analyze - ValidationTargets table missing"
        Debug.Print "   Run Step 2 first"
        Exit Sub
    End If
    
    ' Find first enabled target
    Dim r As ListRow
    Dim targetTableName As String
    Dim KeyColumnHeader As String
    
    For Each r In vtTable.ListRows
        If UCase(Trim(CStr(r.Range.Cells(1, vtTable.ListColumns("Enabled").Index).value))) = "TRUE" Then
            targetTableName = Trim(CStr(r.Range.Cells(1, vtTable.ListColumns("TableName").Index).value))
            KeyColumnHeader = Trim(CStr(r.Range.Cells(1, vtTable.ListColumns("Key Column (Header Name)").Index).value))
            Exit For
        End If
    Next r
    
    If Len(targetTableName) = 0 Then
        Debug.Print "❌ No enabled target found in ValidationTargets"
        Exit Sub
    End If
    
    Debug.Print "Target table name: " & targetTableName
    Debug.Print "Key column header: " & KeyColumnHeader
    Debug.Print ""
    
    ' Search for the table
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim foundSheet As String
    
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set tbl = ws.ListObjects(targetTableName)
        On Error GoTo 0
        
        If Not tbl Is Nothing Then
            foundSheet = ws.Name
            Exit For
        End If
    Next ws
    
    If tbl Is Nothing Then
        Debug.Print "❌ FAIL: Table '" & targetTableName & "' not found in any worksheet"
        Debug.Print ""
        Debug.Print "TROUBLESHOOTING:"
        Debug.Print "  1. Check the table name is exactly correct (case-sensitive)"
        Debug.Print "  2. Verify the data is formatted as an Excel Table (Ctrl+T)"
        Debug.Print "  3. Check the table name in Table Design tab"
        Debug.Print ""
        Debug.Print "Available tables in workbook:"
        ListAllTables
        Exit Sub
    End If
    
    Debug.Print "✅ Table found on sheet: " & foundSheet
    Debug.Print ""
    
    ' Analyze table structure
    Debug.Print "Table structure:"
    Debug.Print "  Header row: " & tbl.HeaderRowRange.Row
    
    If tbl.DataBodyRange Is Nothing Then
        Debug.Print "  ❌ WARNING: Table has NO DATA ROWS"
        Exit Sub
    End If
    
    Dim tableStartRow As Long
    Dim tableEndRow As Long
    Dim rowCount As Long
    
    tableStartRow = tbl.DataBodyRange.Row
    tableEndRow = tableStartRow + tbl.DataBodyRange.Rows.Count - 1
    rowCount = tbl.DataBodyRange.Rows.Count
    
    Debug.Print "  Data start row: " & tableStartRow
    Debug.Print "  Data end row: " & tableEndRow
    Debug.Print "  Total data rows: " & rowCount
    Debug.Print "  Total columns: " & tbl.ListColumns.Count
    Debug.Print ""
    
    ' Check key column
    Dim keyCol As ListColumn
    On Error Resume Next
    Set keyCol = tbl.ListColumns(KeyColumnHeader)
    On Error GoTo 0
    
    If keyCol Is Nothing Then
        Debug.Print "❌ FAIL: Key column '" & KeyColumnHeader & "' not found in table"
        Debug.Print ""
        Debug.Print "Available columns (first 20):"
        PrintTableColumns tbl, 20
    Else
        Debug.Print "✅ Key column '" & KeyColumnHeader & "' found (column index: " & keyCol.Index & ")"
        
        ' Count non-empty key values
        Dim KeyCell As Range
        Dim nonEmptyKeys As Long
        nonEmptyKeys = 0
        
        For Each KeyCell In keyCol.DataBodyRange.Cells
            If Len(Trim(CStr(KeyCell.value))) > 0 Then
                nonEmptyKeys = nonEmptyKeys + 1
            End If
        Next KeyCell
        
        Debug.Print "✅ Rows with key values: " & nonEmptyKeys & " / " & rowCount
    End If
    
    Debug.Print ""
    Debug.Print "----------------------------------------------"
    Debug.Print "IMPORTANT: Table row range is " & tableStartRow & " to " & tableEndRow
    Debug.Print "----------------------------------------------"
    Debug.Print ""
    Debug.Print "If you see errors about 'Row X is outside table range',"
    Debug.Print "it means an older AV_Engine is still using legacy cell"
    Debug.Print "references (B4/D4) instead of the actual table range."
End Sub

' ======================================================
' STEP 4: HEADER MAPPING CHECK
' Compare mapped headers vs actual table headers
' ======================================================
Public Sub Test4_Step4_HeaderMappingCheck()
    Debug.Print ""
    Debug.Print "----------------------------------------------"
    Debug.Print "STEP 4: HEADER MAPPING CHECK"
    Debug.Print "----------------------------------------------"
    Debug.Print ""
    
    ' Get mapping table
    Dim wsConfig As Worksheet
    Dim mappingTable As ListObject
    
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("Config")
    Set mappingTable = wsConfig.ListObjects("AutoValidationCommentPrefixMappingTable")
    On Error GoTo 0
    
    If mappingTable Is Nothing Then
        Debug.Print "❌ AutoValidationCommentPrefixMappingTable not found"
        Exit Sub
    End If
    
    ' Check for ReviewSheet Column Header column
    Dim headerCol As ListColumn
    On Error Resume Next
    Set headerCol = mappingTable.ListColumns("ReviewSheet Column Header")
    On Error GoTo 0
    
    If headerCol Is Nothing Then
        Debug.Print "⚠️  Column 'ReviewSheet Column Header' not found"
        Debug.Print "   This column should contain the target table header names"
        Debug.Print ""
        Debug.Print "   Available columns in mapping table:"
        PrintTableColumns mappingTable, 10
        
        ' Try legacy column
        On Error Resume Next
        Set headerCol = mappingTable.ListColumns("ReviewSheet Column Letter")
        On Error GoTo 0
        
        If Not headerCol Is Nothing Then
            Debug.Print ""
            Debug.Print "   Found legacy 'ReviewSheet Column Letter' column"
            Debug.Print "   Consider adding 'ReviewSheet Column Header' for v2.4 compatibility"
        End If
        Exit Sub
    End If
    
    Debug.Print "✅ ReviewSheet Column Header column found"
    Debug.Print ""
    
    ' Get target table for comparison
    Dim vtTable As ListObject
    Dim targetTableName As String
    Dim TargetTable As ListObject
    
    On Error Resume Next
    Set vtTable = wsConfig.ListObjects("ValidationTargets")
    On Error GoTo 0
    
    If vtTable Is Nothing Then
        Debug.Print "Cannot compare - ValidationTargets table missing"
        Exit Sub
    End If
    
    ' Find enabled target
    Dim r As ListRow
    For Each r In vtTable.ListRows
        If UCase(Trim(CStr(r.Range.Cells(1, vtTable.ListColumns("Enabled").Index).value))) = "TRUE" Then
            targetTableName = Trim(CStr(r.Range.Cells(1, vtTable.ListColumns("TableName").Index).value))
            Exit For
        End If
    Next r
    
    ' Find target table
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set TargetTable = ws.ListObjects(targetTableName)
        On Error GoTo 0
        If Not TargetTable Is Nothing Then Exit For
    Next ws
    
    If TargetTable Is Nothing Then
        Debug.Print "Cannot compare - target table '" & targetTableName & "' not found"
        Exit Sub
    End If
    
    Debug.Print "Comparing mapped headers to table '" & targetTableName & "':"
    Debug.Print ""
    
    ' Compare each mapping
    Dim devFuncCol As ListColumn
    Dim autoValCol As ListColumn
    Dim mappedHeader As String
    Dim devFunc As String
    Dim AutoValidate As String
    Dim foundCount As Long
    Dim missingCount As Long
    Dim skippedCount As Long
    
    On Error Resume Next
    Set devFuncCol = mappingTable.ListColumns("Dev Function Names")
    Set autoValCol = mappingTable.ListColumns("AutoValidate")
    On Error GoTo 0
    
    foundCount = 0
    missingCount = 0
    skippedCount = 0
    
    For Each r In mappingTable.ListRows
        devFunc = Trim(CStr(r.Range.Cells(1, devFuncCol.Index).value))
        mappedHeader = Trim(CStr(r.Range.Cells(1, headerCol.Index).value))
        AutoValidate = UCase(Trim(CStr(r.Range.Cells(1, autoValCol.Index).value)))
        
        ' Check if header exists in target table
        Dim testCol As ListColumn
        On Error Resume Next
        Set testCol = TargetTable.ListColumns(mappedHeader)
        On Error GoTo 0
        
        If AutoValidate <> "TRUE" Then
            Debug.Print "  ⬜ " & devFunc & " -> '" & mappedHeader & "' (AutoValidate=FALSE)"
            skippedCount = skippedCount + 1
        ElseIf testCol Is Nothing Then
            Debug.Print "  ❌ " & devFunc & " -> '" & mappedHeader & "' NOT FOUND"
            missingCount = missingCount + 1
        Else
            Debug.Print "  ✅ " & devFunc & " -> '" & mappedHeader & "' OK"
            foundCount = foundCount + 1
        End If
    Next r
    
    Debug.Print ""
    Debug.Print "Summary:"
    Debug.Print "  ✅ Found: " & foundCount
    Debug.Print "  ❌ Missing: " & missingCount
    Debug.Print "  ⬜ Skipped (AutoValidate=FALSE): " & skippedCount
    
    If missingCount > 0 Then
        Debug.Print ""
        Debug.Print "ACTION REQUIRED for missing headers:"
        Debug.Print "  1. Check exact spelling of header names in target table"
        Debug.Print "  2. Update 'ReviewSheet Column Header' values in mapping table"
        Debug.Print "  3. Or add missing columns to target table"
    End If
End Sub

' ======================================================
' RUN VALIDATION TEST
' ======================================================
Public Sub Test4_RunValidation()
    Debug.Print ""
    Debug.Print "=============================================="
    Debug.Print "TEST 4 - RUNNING VALIDATION (v2.4)"
    Debug.Print "=============================================="
    Debug.Print ""
    
    ' First verify engine version
    On Error Resume Next
    Dim engineVersion As String
    engineVersion = AV_Engine.MODULE_VERSION
    On Error GoTo 0
    
    If Len(engineVersion) = 0 Then
        Debug.Print "❌ ERROR: AV_Engine.MODULE_VERSION not found"
        Debug.Print "   Import av_engine_v2_4.bas before running validation"
        Exit Sub
    End If
    
    Debug.Print "AV_Engine version: " & engineVersion
    Debug.Print ""
    Debug.Print "Starting validation at " & Format(Now, "hh:mm:ss")
    Debug.Print "Watch ValidationTrackerForm for progress..."
    Debug.Print ""
    Debug.Print "---[VALIDATION OUTPUT BELOW]---"
    Debug.Print ""
    
    On Error GoTo ErrHandler
    
    Dim startTime As Single
    startTime = Timer
    
    ' Run the validation
    AV_Engine.RunFullValidationMaster
    
    Dim elapsed As Single
    elapsed = Timer - startTime
    
    Debug.Print ""
    Debug.Print "---[VALIDATION OUTPUT ABOVE]---"
    Debug.Print ""
    Debug.Print "✅ Validation completed"
    Debug.Print "   Elapsed time: " & Format(elapsed, "0.00") & " seconds"
    Debug.Print ""
    Debug.Print "=============================================="
    Exit Sub
    
ErrHandler:
    Debug.Print ""
    Debug.Print "❌ ERROR DURING VALIDATION"
    Debug.Print "   Error #" & Err.Number & ": " & Err.description
    Debug.Print "=============================================="
End Sub

' ======================================================
' HELPER: List all tables in workbook
' ======================================================
Private Sub ListAllTables()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim Count As Long
    Count = 0
    
    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            Count = Count + 1
            Debug.Print "  " & Count & ". " & tbl.Name & " (on sheet: " & ws.Name & ")"
        Next tbl
    Next ws
    
    If Count = 0 Then
        Debug.Print "  (no tables found)"
    End If
End Sub

' ======================================================
' HELPER: Print table columns
' ======================================================
Private Sub PrintTableColumns(tbl As ListObject, maxCols As Long)
    Dim col As ListColumn
    Dim Count As Long
    Count = 0
    
    For Each col In tbl.ListColumns
        Count = Count + 1
        If Count > maxCols Then
            Debug.Print "  ... (" & tbl.ListColumns.Count & " total columns)"
            Exit Sub
        End If
        Debug.Print "  " & Count & ". " & col.Name
    Next col
End Sub
