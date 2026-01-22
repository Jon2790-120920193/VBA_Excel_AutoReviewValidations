Public Sub Test1_3_DirectMapTest()
    ' Direct test of GetAutoValidationMap
    ' This bypasses the engine to isolate the problem
    
    On Error GoTo ErrorHandler
    
    ' Show form first
    AV_UI.ShowValidationTrackerForm
    DoEvents
    
    AV_UI.AppendUserLog "=== Direct GetAutoValidationMap Test ==="
    
    ' Try to load the map
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets("Config")
    
    AV_UI.AppendUserLog "About to call GetAutoValidationMap..."
    DoEvents
    
    Dim result As Object
    Set result = AV_Core.GetAutoValidationMap(wsConfig)
    
    AV_UI.AppendUserLog "GetAutoValidationMap returned"
    AV_UI.AppendUserLog "Result count: " & result.Count
    
    If result.Count > 0 Then
        AV_UI.AppendUserLog "Validation functions loaded:"
        Dim key As Variant
        For Each key In result.Keys
            AV_UI.AppendUserLog "  - " & key
        Next key
    End If
    
    AV_UI.AppendUserLog "=== Test Complete ==="
    Exit Sub
    
ErrorHandler:
    AV_UI.AppendUserLog "ERROR in Test1_3_DirectMapTest"
    AV_UI.AppendUserLog "Error #" & Err.Number
    AV_UI.AppendUserLog "Description: " & Err.Description
    AV_UI.AppendUserLog "Source: " & Err.Source
    
    MsgBox "Error #" & Err.Number & vbCrLf & Err.Description & vbCrLf & "Source: " & Err.Source
End Sub
