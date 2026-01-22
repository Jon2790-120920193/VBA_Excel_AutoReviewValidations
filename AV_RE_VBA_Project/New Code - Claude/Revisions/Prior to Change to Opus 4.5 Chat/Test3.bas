Attribute VB_Name = "Test3"
Public Sub Test1_3_DirectMapTest()
    ' Direct test of GetAutoValidationMap
    ' This bypasses the engine to isolate the problem
    
    On Error GoTo ErrorHandler
    
    ' Show form first
    AV_UI.ShowValidationTrackerForm
    DoEvents
    
    Debug.Print "=== Direct GetAutoValidationMap Test ==="
    
    ' Try to load the map
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets("Config")
    
    Debug.Print "About to call GetAutoValidationMap..."
    DoEvents
    
    Dim result As Object
    Set result = AV_Core.GetAutoValidationMap(wsConfig)
    
    Debug.Print "GetAutoValidationMap returned"
    Debug.Print "Result count: " & result.Count
    
    If result.Count > 0 Then
        Debug.Print "Validation functions loaded:"
        Dim key As Variant
        For Each key In result.Keys
            Debug.Print "  - " & key
        Next key
    End If
    
    Debug.Print "=== Test Complete ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR in Test1_3_DirectMapTest"
    Debug.Print "Error #" & Err.Number
    Debug.Print "Description: " & Err.Description
    Debug.Print "Source: " & Err.Source
    
    MsgBox "Error #" & Err.Number & vbCrLf & Err.Description & vbCrLf & "Source: " & Err.Source
End Sub
