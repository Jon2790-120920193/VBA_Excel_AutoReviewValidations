Attribute VB_Name = "Test1"
Option Explicit

' ======================================================
' TEST MODULE - Phase 2 Validation Tests
' ======================================================

' ======================================================
' DIAGNOSTIC TESTS
' ======================================================

Public Sub Test1_BasicValidation()
    Dim errMsg As String
    Debug.Print AV_Core.ValidateConfiguration(errMsg)
    Debug.Print "Error (if any): " & errMsg
    
    'Result: False Error (if any): Target sheet not found: REP2DSMDraft_Buildings Check ValidationTargets table
    'Note: REP2DSMDRAFT_Buildings is not a sheet, but one of the validation target tables. Would like to attempt to avoid sheet references and go to Excel ListObject References only if possible
    'Second Test (1.1), post AV_Engine and AV_Core change to point to REP2DSMDraft_Buildings as ListObject instead of Sheet Object.
    'Result
    'True
    'Error (if any):

End Sub

Public Sub Test1_DetailedDiagnostic()
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets("Config")
    
    Debug.Print "=== VALIDATION CONFIGURATION DIAGNOSTIC ==="
    Debug.Print ""
    
    ' Step 1: Check critical tables
    Debug.Print "Step 1: Checking critical tables..."
    Dim tables As Variant
    tables = Array("ValidationTargets", "AutoValidationCommentPrefixMappingTable", "AutoFormatOnFullValidation")
    
    Dim tbl As Variant
    For Each tbl In tables
        If AV_DataAccess.TableExists(wsConfig, CStr(tbl)) Then
            Debug.Print "  OK: " & tbl & " exists"
        Else
            Debug.Print "  MISSING: " & tbl
        End If
    Next tbl
    
    ' Step 2: Check targets
    Debug.Print ""
    Debug.Print "Step 2: Loading validation targets..."
    Dim config As AV_Core.ValidationConfig
    config = AV_Core.LoadValidationConfig()
    Debug.Print "  Target count: " & config.TargetCount
    
    If config.TargetCount = 0 Then
        Debug.Print "  ERROR: No enabled targets found"
        Exit Sub
    End If
    
    ' Step 3: Check each target
    Debug.Print ""
    Debug.Print "Step 3: Checking each target..."
    Dim i As Long
    For i = 1 To config.TargetCount
        Debug.Print "  Target " & i & ": " & config.Targets(i).tableName
        
        If AV_DataAccess.WorksheetExists(config.Targets(i).tableName) Then
            Debug.Print "    Sheet exists: YES"
            
            Dim ws As Worksheet
            Set ws = ThisWorkbook.Sheets(config.Targets(i).tableName)
            Debug.Print "    ListObject count: " & ws.ListObjects.Count
            
            If ws.ListObjects.Count = 0 Then
                Debug.Print "    ERROR: No ListObjects found in sheet"
            End If
        Else
            Debug.Print "    ERROR: Sheet does not exist"
        End If
    Next i
    
    ' Step 4: Run actual validation
    Debug.Print ""
    Debug.Print "Step 4: Running ValidateConfiguration..."
    Dim errMsg As String
    Dim result As Boolean
    result = AV_Core.ValidateConfiguration(errMsg)
    
    Debug.Print "  Result: " & result
    If Not result Then
        Debug.Print "  Error: " & errMsg
    End If
    
    Debug.Print ""
    Debug.Print "=== END DIAGNOSTIC ==="
    
'Result: Test1.1
'Step 1: Checking critical tables...
'OK:   ValidationTargets exists
'OK:   AutoValidationCommentPrefixMappingTable exists
'OK:   AutoFormatOnFullValidation exists

'Step 2: Loading validation targets...
  'Target count: 3

'Step 3: Checking each target...
  'target 1: REP2DSMDraft_Buildings
'ERROR:     Sheet Not does
  'target 2: REP2DSMDraft_Works
'ERROR:     Sheet Not does
  'target 3: REP2DSMDraft_Other
'ERROR:     Sheet Not does

'Step 4: Running ValidateConfiguration...
  'Result: True

'=== END DIAGNOSTIC ===
End Sub

Public Sub Test2_LoadConfig()
    Dim config As AV_Core.ValidationConfig
    config = AV_Core.LoadValidationConfig()
    Debug.Print "Target Count: " & config.TargetCount
    'Result Test 1.1: Target Count: 3
End Sub

Public Sub Test3_CachedTable()
    Dim tbl As ListObject
    Set tbl = AV_Core.GetValidationTable(AV_Constants.TBL_GIW_VALIDATION)
    Debug.Print "Table Name: " & tbl.Name

    'Result Test 1.1: Table Name: GIWValidationTable
End Sub

' ======================================================
' TRIGGER FUNCTION - Run Full Validation
' ======================================================

Public Sub TriggerValidation()
    ' Simple trigger - uses defaults (English, no specific sheet)
    AV_Engine.RunFullValidation
    
    'Will wait for testing to be completed first to run.
End Sub

Public Sub TriggerValidation_WithOptions()
    ' Trigger with options
    ' Parameters: sheetName (optional), english (optional, default True)
    
    ' Example: Validate specific sheet
    ' AV_Engine.RunFullValidation "MySheet", True
    
    ' For default behavior (all enabled targets):
    AV_Engine.RunFullValidation
    
    'Will wait for testing to be completed first to run.
End Sub
