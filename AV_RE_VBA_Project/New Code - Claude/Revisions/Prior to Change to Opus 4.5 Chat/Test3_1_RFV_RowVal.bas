Attribute VB_Name = "Test3_1_RFV_RowVal"
Sub DiagnoseEngine()
    Debug.Print "=== ENGINE DIAGNOSTIC ==="
    
    ' Check if ProcessValidationTarget exists (Phase 2)
    On Error Resume Next
    Dim hasNewMethod As Boolean
    hasNewMethod = False
    
    ' Try to find ProcessValidationTarget in code
    Dim vbComp As Object
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If vbComp.Name = "AV_Engine" Then
            Dim code As String
            code = vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines)
            If InStr(code, "ProcessValidationTarget") > 0 Then
                hasNewMethod = True
                Debug.Print "AV_Engine: Phase 2 version (has ProcessValidationTarget)"
            Else
                Debug.Print "AV_Engine: Phase 1 version (missing ProcessValidationTarget)"
            End If
            Exit For
        End If
    Next
    
    If Not hasNewMethod Then
        Debug.Print "ERROR: Using old AV_Engine - validations won't run"
        Debug.Print "Solution: Import Phase 2 AV_Engine from project files"
    End If
End Sub

Sub DiagnoseValidationFlow()
    Debug.Print "=== VALIDATION FLOW DIAGNOSTIC ==="
    Debug.Print ""
    
    ' Get first row with key
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Tab - Onglet 1")
    
    ' Find first populated row (assuming row 5 is first)
    Dim testRow As Long
    testRow = 5
    
    Debug.Print "Testing row " & testRow
    Debug.Print ""
    
    ' Load mappings
    Dim autoValMap As Object
    Set autoValMap = AV_Core.GetAutoValidationMap()
    
    Debug.Print "AutoValMap loaded: " & autoValMap.Count & " functions"
    Debug.Print ""
    
    ' Check if ValidateSingleRow exists
    Dim hasValidateSingleRow As Boolean
    hasValidateSingleRow = False
    
    Dim vbComp As Object
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If vbComp.Name = "AV_Engine" Then
            Dim code As String
            code = vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines)
            If InStr(code, "Sub ValidateSingleRow") > 0 Or InStr(code, "Function ValidateSingleRow") > 0 Then
                hasValidateSingleRow = True
                Debug.Print "? ValidateSingleRow found in AV_Engine"
            Else
                Debug.Print "? ValidateSingleRow NOT found in AV_Engine"
            End If
            Exit For
        End If
    Next
    
    If Not hasValidateSingleRow Then
        Debug.Print ""
        Debug.Print "PROBLEM: ValidateSingleRow is missing!"
        Debug.Print "This is why validations don't run."
    End If
    
    Debug.Print ""
    Debug.Print "Checking if ProcessValidationTarget calls ValidateSingleRow..."
    
    ' Check ProcessValidationTarget code
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If vbComp.Name = "AV_Engine" Then
            code = vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines)
            
            ' Find ProcessValidationTarget
            Dim startLine As Long, endLine As Long
            Dim i As Long
            For i = 1 To vbComp.CodeModule.CountOfLines
                If InStr(vbComp.CodeModule.Lines(i, 1), "Sub ProcessValidationTarget") > 0 Or _
                   InStr(vbComp.CodeModule.Lines(i, 1), "Function ProcessValidationTarget") > 0 Then
                    startLine = i
                End If
                If startLine > 0 And InStr(vbComp.CodeModule.Lines(i, 1), "End Sub") > 0 Then
                    endLine = i
                    Exit For
                End If
            Next
            
            If startLine > 0 And endLine > 0 Then
                Dim procCode As String
                procCode = vbComp.CodeModule.Lines(startLine, endLine - startLine + 1)
                
                If InStr(procCode, "ValidateSingleRow") > 0 Then
                    Debug.Print "? ProcessValidationTarget DOES call ValidateSingleRow"
                Else
                    Debug.Print "? ProcessValidationTarget does NOT call ValidateSingleRow"
                    Debug.Print ""
                    Debug.Print "PROBLEM: ProcessValidationTarget only runs RunAutoCheckDataValidation"
                    Debug.Print "Missing the row-by-row validation loop!"
                End If
            End If
            Exit For
        End If
    Next
    
    Debug.Print ""
    Debug.Print "=== DIAGNOSTIC COMPLETE ==="
End Sub

Sub DiagnoseAutoValMap()
    Debug.Print "=== AUTOVALIDATION MAPPING DIAGNOSTIC ==="
    Debug.Print ""
    
    Dim autoValMap As Object
    Set autoValMap = AV_Core.GetAutoValidationMap()
    
    Debug.Print "Total functions mapped: " & autoValMap.Count
    Debug.Print ""
    
    ' Check each mapping
    Dim key As Variant
    For Each key In autoValMap.Keys
        Debug.Print "Function: " & key
        
        Dim item As Object
        Set item = autoValMap(key)
        
        Debug.Print "  ColumnRef: " & item("ColumnRef")
        Debug.Print "  DropColHeader: " & item("DropColHeader")
        Debug.Print "  AutoValidate: " & item("AutoValidate")
        Debug.Print "  PrefixEN: " & item("PrefixEN")
        
        If item("AutoValidate") = False Then
            Debug.Print "  ? DISABLED - This validation will NOT run"
        Else
            Debug.Print "  ? ENABLED"
        End If
        
        Debug.Print ""
    Next key
    
    Debug.Print "=== CHECK YOUR TABLE ==="
    Debug.Print "If all show AutoValidate = False, change them to TRUE in the table"
    Debug.Print "========================="
End Sub

Sub CheckValidateSingleRow()
    Debug.Print "=== VALIDATESINGLEROW CODE CHECK ==="
    
    Dim vbComp As Object
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If vbComp.Name = "AV_Engine" Then
            Dim code As String
            Dim startLine As Long, endLine As Long
            Dim i As Long
            
            ' Find ValidateSingleRow
            For i = 1 To vbComp.CodeModule.CountOfLines
                Dim lineText As String
                lineText = vbComp.CodeModule.Lines(i, 1)
                If InStr(lineText, "Sub ValidateSingleRow") > 0 Or InStr(lineText, "Function ValidateSingleRow") > 0 Then
                    startLine = i
                End If
                If startLine > 0 And InStr(lineText, "End Sub") > 0 Then
                    endLine = i
                    Exit For
                End If
            Next
            
            If startLine > 0 And endLine > 0 Then
                code = vbComp.CodeModule.Lines(startLine, endLine - startLine + 1)
                Debug.Print "ValidateSingleRow found at line " & startLine
                Debug.Print ""
                
                ' Check how it accesses columns
                If InStr(code, "wsData.Range(") > 0 And InStr(code, "& rowNum") > 0 Then
                    Debug.Print "? PROBLEM: Uses Range(letter & rowNum) - expects column LETTERS"
                    Debug.Print "   But ColumnRef contains HEADER NAMES"
                ElseIf InStr(code, ".ListColumns(") > 0 Or InStr(code, "FindColumnByHeader") > 0 Then
                    Debug.Print "? GOOD: Uses table column lookup"
                Else
                    Debug.Print "??  UNKNOWN: Cannot determine column access method"
                End If
                
                Debug.Print ""
                Debug.Print "First 20 lines of ValidateSingleRow:"
                Debug.Print "---"
                Debug.Print vbComp.CodeModule.Lines(startLine, 20)
                Debug.Print "---"
            End If
            Exit For
        End If
    Next
    
    Debug.Print "=== CHECK COMPLETE ==="
End Sub
