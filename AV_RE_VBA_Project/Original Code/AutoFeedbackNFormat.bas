Attribute VB_Name = "AutoFeedbackNFormat"
Option Explicit
' ========================================================
'  Module: AutoFeedbackNFormat
'  Purpose: Central handler for writing validation feedback
'           and applying formatting to review sheet columns.
'
'  Dependencies:
'     - Public_Utilities.SafeTrim
'     - LoadSmartFunctionReferences.GetAutoValidationMap
'     - CellAutoMessageUtilities.WriteSystemTagToDropColumn
'
'  Description:
'     This function is typically called by individual
'     validation functions to write context-specific feedback
'     messages in the correct drop column (determined from the
'     AutoValidationMappingTable).
' ========================================================


' ===============================================
' Adds feedback to the correct column for a validation
' ===============================================
Public Sub AddValidationFeedback(ByVal devFunctionName As String, _
                                 ByVal wsTarget As Worksheet, _
                                 ByVal targetRow As Long, _
                                 ByVal messageText As String, Optional ByVal FormatType As String = "Default", _
                                 Optional ByVal english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)

    Dim map As Object
    Dim dropColHeader As String
    Dim prefixText As String
    Dim fullMessage As String
    Dim TargetColLetter As String
    devFunctionName = "Validate_Column_" & devFunctionName
    
    ' --- Load format mapping
    If FormatMap Is Nothing Then
        Set FormatMap = DefaultFormatMap()
        
        If FormatMap Is Nothing Then
            Debug.Print "Error loading the formatting map from the CellFormatUtilities Module"
        Exit Sub
        End If
    End If
    
    ' --- Load Smart Autovalidation mapping
    If AutoValMap Is Nothing Then
        Debug.Print "[AddValidationFeedback] ? No AutoValidation map loaded."
        Set AutoValMap = GetAutoValidationMap()
        
        If AutoValMap Is Nothing Then
            Debug.Print "Error loading the smart autovalidation mapping from LoadSmartFunctionReferences"
        Exit Sub
        End If
    End If
    
    ' --- Look up called function in the map
    If Not AutoValMap.Exists(devFunctionName) Then
        Debug.Print "[AddValidationFeedback] ? Dev function '" & devFunctionName & "' not found in mapping table."
        Exit Sub
    End If


    Set map = AutoValMap(devFunctionName)
    dropColHeader = SafeTrim(map("DropColHeader"))
    TargetColLetter = SafeTrim(map("ColumnRef"))
    If english Then
        prefixText = SafeTrim(map("PrefixEN"))
    Else
        prefixText = SafeTrim(map("PrefixFR"))
    End If

    ' --- Compose final message
    If Len(prefixText) > 0 Then
        fullMessage = prefixText & " " & messageText
    Else
        fullMessage = messageText
    End If

    ' --- Delegate actual writing to your shared utility
    If FormatType <> "Default" Then
        WriteSystemTagToDropColumn wsTarget, dropColHeader, targetRow, TargetColLetter, fullMessage, FormatType, FormatMap
        Debug.Print "[AddValidationFeedback] ? " & devFunctionName & _
            " | DropCol=" & dropColHeader & _
            " | Row=" & targetRow & _
            " | Msg='" & fullMessage & "'"
    Else
        WriteSystemTagToDropColumn wsTarget, dropColHeader, targetRow, TargetColLetter, fullMessage, FormatType, FormatMap
    End If
    
    


End Sub






