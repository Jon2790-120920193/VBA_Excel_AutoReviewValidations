Attribute VB_Name = "AV_Format"
Option Explicit


' ===== BEGIN MERGED: AV2_CellFormatUtilities.bas =====

Private Const FormatKeyColumn As String = "Formatting Key"
Private Const AutoFormatColumn As String = "Autoformatting"
Private Const GlobalFormatTableName As String = "AutoFormatOnFullValidation"
Private Const PriorityColumn As String = "KeyFlagPriority"

Private Const ConfigSheet As String = "Config"

Private Const ReviewFlagsTableName As String = "ReviewRefColumnTable"
Private Const RevStatusColName As String = "ReviewStatusColumn"
Private Const AutoRevStatusColName As String = "AutoReviewColumnLetter"
Private Const HumanSetStatus As String = "HumanSetRevStatus"

Public Const DebugON As Boolean = "False"

Private Const ReviewStatusTable As String = "ReviewStatusTable"

'====================================================
' Load all format objects (AV2_clsCellFormat) into memory
'====================================================
Public Function AV2_LoadFormatMap(wsConfig As Worksheet) As Object
    Dim Table As ListObject
    Dim dict As Object
    Dim r As ListRow
    Dim key As String
    Dim fmt As AV2_clsCellFormat
    Dim srcCell As Range
    Dim colKey As Long, colFormat As Long, colPriority As Long
    Dim priorityVal As Long
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    Set Table = wsConfig.ListObjects(GlobalFormatTableName)
    On Error GoTo 0
    
    If Table Is Nothing Then
        Debug.Print "?? Table '" & GlobalFormatTableName & "' not found on " & wsConfig.Name
        Set AV2_LoadFormatMap = dict
        Exit Function
    End If
    
    colKey = Table.ListColumns(FormatKeyColumn).Index
    colFormat = Table.ListColumns(AutoFormatColumn).Index
    colPriority = Table.ListColumns(PriorityColumn).Index
    
    For Each r In Table.ListRows
        key = Trim(r.Range.Cells(1, colKey).value)
        Set srcCell = r.Range.Cells(1, colFormat)
        ' Handle empty or invalid values gracefully
        If IsNumeric(r.Range.Cells(1, colPriority).value) Then
            priorityVal = CLng(r.Range.Cells(1, colPriority).value)
        Else
            priorityVal = 0
        End If

        If Len(key) > 0 Then
            Set fmt = getCellFormat(srcCell)
            fmt.Priority = priorityVal   ' < new property assignment
            Set dict(key) = fmt
        End If
    Next r
    
    Set AV2_LoadFormatMap = dict
End Function

Public Function AV2_DefaultFormatMap() As Object
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Worksheets("Config")
    Dim FormatMap As Object
    
    Set FormatMap = AV2_LoadFormatMap(wsConfig)
    
    If FormatMap Is Nothing Then
        Debug.Print "Error loading the formatting map from the CellFormatUtilities Module"
        Exit Function
    End If
    
    
    AV2_DebugMessage "Format Map was loaded from default settings. Loaded from: '" & "Config" & "' Workbook Sheet", "AV2_DefaultFormatMap", DebugON
    Set AV2_DefaultFormatMap = FormatMap
    
End Function


' Receives Row Formatting and detects the Highest Level of Correction
' Sets the Key Column with the respective Priority Level, e.g. Error, AutoCorrect...etc

Public Sub AV2_FormatKeyCell(rowRange As Range, FormatMap As Object)
    ' --- If multiple rows, call recursively for each row ---
    Dim r As Range
    If rowRange.Rows.count > 1 Then
        For Each r In rowRange.Rows
            AV2_FormatKeyCell r, FormatMap
        Next r
        Exit Sub
    End If
    
    Dim cell As Range
    Dim key As String
    Dim fmtInfo As AV2_clsCellFormat
    Dim currentPriority As Long
    Dim highestPriority As Long
    Dim highestKey As String
    Dim KeyCell As Range
    Dim wsConfig As Worksheet
    
    Set wsConfig = ThisWorkbook.Worksheets(ConfigSheet)
    
    ' Load the KeyColumn of the Row.
    Dim colLetter As String
    colLetter = UCase(wsConfig.Range("B5").value) ' ensure uppercase for column letter
    Set KeyCell = rowRange.Worksheet.Range(colLetter & rowRange.row)
    
    highestPriority = -1  ' start low (assuming priority is positive)
    highestKey = ""
    
    For Each cell In rowRange.Cells
        ' Get the format key for this cell
        key = AV2_getFormatType(cell, FormatMap)
        
        If Len(key) > 0 Then
            ' Access the AV2_clsCellFormat directly
            Set fmtInfo = FormatMap(key)
            currentPriority = fmtInfo.Priority
            
            ' Compare to find highest
            If currentPriority > highestPriority Then
                highestPriority = currentPriority
                highestKey = key
            End If
        End If
    Next cell
    
    ' Apply the format with the highest priority
    If Len(highestKey) > 0 Then
        If highestPriority = 2 Then
            Call SetReviewStatus(True, rowRange, True)
        ElseIf highestPriority = 3 Then
            Call SetReviewStatus(True, rowRange, False)
        Else
            Call SetReviewStatus(False, rowRange, False)
        End If
            
        Call AV2_setFormat(KeyCell, highestKey, FormatMap)
    End If
End Sub

Private Sub SetReviewStatus(ReviewRequired As Boolean, rowRange As Range, Optional AutoCorrected As Boolean = False)

    Dim reviewStatus As AV2_revStatusRef
    Dim Table As ListObject
    Dim wsConfig As Worksheet
    Dim RowNum As Long

    ' Instantiate the class
    Set reviewStatus = New AV2_revStatusRef
    
    Set wsConfig = ThisWorkbook.Worksheets(ConfigSheet)

    On Error Resume Next
    Set Table = wsConfig.ListObjects(ReviewFlagsTableName)
    On Error GoTo 0

    If Table Is Nothing Then
        AV2_DebugMessage "?? Table '" & GlobalFormatTableName & "' not found on " & wsConfig.Name, "SetReviewStatus", DebugON
        Exit Sub
    End If

    RowNum = rowRange.row

    ' === Get the proper cells ===
    Set reviewStatus.RevStatusCol = AV2_GetCellFromTableColumnHeader(Table, rowRange, RevStatusColName)
    Set reviewStatus.AutoReviewDropCol = AV2_GetCellFromTableColumnHeader(Table, rowRange, AutoRevStatusColName)
    Set reviewStatus.HumanSetStatusCol = AV2_GetCellFromTableColumnHeader(Table, rowRange, HumanSetStatus)

    ' === Assign values ===
    If ReviewRequired Then
        If AutoCorrected Then
            reviewStatus.AutoReviewDropCol.value = "Auto Corrected"
        Else
            reviewStatus.AutoReviewDropCol.value = "Error"
        End If
    Else
        reviewStatus.AutoReviewDropCol.value = "No Errors Found"
    End If

End Sub




'====================================================
' Capture a cell's complete format into a AV2_clsCellFormat
'====================================================
Private Function getCellFormat(cell As Range) As AV2_clsCellFormat
    Dim f As New AV2_clsCellFormat
    Dim b As Border
    
    With cell
        '--- Base appearance
        f.InteriorColor = .Interior.Color
        f.FontColor = .Font.Color
        f.Bold = .Font.Bold
        f.FontName = .Font.Name
        f.FontSize = .Font.Size
        f.NumberFormat = .NumberFormat
        
        '--- Borders
        With .Borders(xlEdgeTop)
            f.BorderTopColor = .Color
            f.BorderTopLineStyle = .LineStyle
        End With
        With .Borders(xlEdgeBottom)
            f.BorderBottomColor = .Color
            f.BorderBottomLineStyle = .LineStyle
        End With
        With .Borders(xlEdgeLeft)
            f.BorderLeftColor = .Color
            f.BorderLeftLineStyle = .LineStyle
        End With
        With .Borders(xlEdgeRight)
            f.BorderRightColor = .Color
            f.BorderRightLineStyle = .LineStyle
        End With
    End With
    
    Set getCellFormat = f
End Function

Public Sub AV2_setFormat(TargetCell As Range, FormatType As String, FormatMap As Object)
    Dim fmt As AV2_clsCellFormat
    
    If FormatMap.Exists(FormatType) Then
        Set fmt = FormatMap(FormatType)
        ApplyCellFormatting TargetCell, fmt
    Else
        AV2_DebugMessage "FormatMap type: '" & FormatType & "' could not be found in the '" & GlobalFormatTableName & "' Config table", "AV2_setFormat", DebugON
    End If
    
End Sub

Public Function AV2_getFormatType(TargetCell As Range, FormatMap As Object) As String
    Dim cellFormat As AV2_clsCellFormat
    Dim key As Variant
    Dim fmt As AV2_clsCellFormat
    
    ' Get the format of the target cell
    Set cellFormat = getCellFormat(TargetCell)
    
    ' Loop through all known formats in the map
    For Each key In FormatMap.Keys
        Set fmt = FormatMap(key)
        
        If FormatsAreEqual(fmt, cellFormat) Then
            AV2_getFormatType = key
            Exit Function
        End If
    Next key
    
    ' If no match found
    AV2_getFormatType = vbNullString
End Function


'====================================================
' Apply stored format to a target cell
'====================================================
Private Sub ApplyCellFormatting(TargetCell As Range, fmt As AV2_clsCellFormat)
    If fmt Is Nothing Then Exit Sub
    With TargetCell
        ' Base appearance
        .Interior.Color = fmt.InteriorColor
        .Font.Color = fmt.FontColor
        .Font.Bold = fmt.Bold
        .Font.Name = fmt.FontName
        .Font.Size = fmt.FontSize
        .NumberFormat = fmt.NumberFormat
        
        ' Borders
        With .Borders(xlEdgeTop)
            .Color = fmt.BorderTopColor
            .LineStyle = fmt.BorderTopLineStyle
        End With
        With .Borders(xlEdgeBottom)
            .Color = fmt.BorderBottomColor
            .LineStyle = fmt.BorderBottomLineStyle
        End With
        With .Borders(xlEdgeLeft)
            .Color = fmt.BorderLeftColor
            .LineStyle = fmt.BorderLeftLineStyle
        End With
        With .Borders(xlEdgeRight)
            .Color = fmt.BorderRightColor
            .LineStyle = fmt.BorderRightLineStyle
        End With
    End With
End Sub

Private Function FormatsAreEqual(fmt1 As AV2_clsCellFormat, fmt2 As AV2_clsCellFormat) As Boolean
    If fmt1 Is Nothing Or fmt2 Is Nothing Then Exit Function
    
    FormatsAreEqual = _
        (fmt1.InteriorColor = fmt2.InteriorColor) And _
        (fmt1.FontColor = fmt2.FontColor) And _
        (fmt1.Bold = fmt2.Bold) And _
        (fmt1.FontName = fmt2.FontName) And _
        (fmt1.FontSize = fmt2.FontSize) And _
        (fmt1.NumberFormat = fmt2.NumberFormat) And _
        (fmt1.BorderTopColor = fmt2.BorderTopColor) And _
        (fmt1.BorderTopLineStyle = fmt2.BorderTopLineStyle) And _
        (fmt1.BorderBottomColor = fmt2.BorderBottomColor) And _
        (fmt1.BorderBottomLineStyle = fmt2.BorderBottomLineStyle) And _
        (fmt1.BorderLeftColor = fmt2.BorderLeftColor) And _
        (fmt1.BorderLeftLineStyle = fmt2.BorderLeftLineStyle) And _
        (fmt1.BorderRightColor = fmt2.BorderRightColor) And _
        (fmt1.BorderRightLineStyle = fmt2.BorderRightLineStyle)
End Function

' ===== END MERGED: AV2_CellFormatUtilities.bas =====


' ===== BEGIN MERGED: AV2_AutoFeedbackNFormat.bas =====

' ========================================================
'  Module: AutoFeedbackNFormat
'  Purpose: Central handler for writing validation feedback
'           and applying formatting to review sheet columns.
'
'  Dependencies:
'     - Public_Utilities.AV2_SafeTrim
'     - LoadSmartFunctionReferences.AV2_GetAutoValidationMap
'     - CellAutoMessageUtilities.AV2_WriteSystemTagToDropColumn
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
Public Sub AV2_AddValidationFeedback(ByVal devFunctionName As String, _
                                 ByVal wsTarget As Worksheet, _
                                 ByVal targetRow As Long, _
                                 ByVal messageText As String, Optional ByVal FormatType As String = "Default", _
                                 Optional ByVal english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)

    Dim map As Object
    Dim dropColHeader As String
    Dim prefixText As String
    Dim fullMessage As String
    Dim TargetColLetter As String
    devFunctionName = "AV2_Validate_Column_" & devFunctionName
    
    ' --- Load format mapping
    If FormatMap Is Nothing Then
        Set FormatMap = AV2_DefaultFormatMap()
        
        If FormatMap Is Nothing Then
            Debug.Print "Error loading the formatting map from the CellFormatUtilities Module"
        Exit Sub
        End If
    End If
    
    ' --- Load Smart Autovalidation mapping
    If AutoValMap Is Nothing Then
        Debug.Print "[AV2_AddValidationFeedback] ? No AutoValidation map loaded."
        Set AutoValMap = AV2_GetAutoValidationMap()
        
        If AutoValMap Is Nothing Then
            Debug.Print "Error loading the smart autovalidation mapping from LoadSmartFunctionReferences"
        Exit Sub
        End If
    End If
    
    ' --- Look up called function in the map
    If Not AutoValMap.Exists(devFunctionName) Then
        Debug.Print "[AV2_AddValidationFeedback] ? Dev function '" & devFunctionName & "' not found in mapping table."
        Exit Sub
    End If


    Set map = AutoValMap(devFunctionName)
    dropColHeader = AV2_SafeTrim(map("DropColHeader"))
    TargetColLetter = AV2_SafeTrim(map("ColumnRef"))
    If english Then
        prefixText = AV2_SafeTrim(map("PrefixEN"))
    Else
        prefixText = AV2_SafeTrim(map("PrefixFR"))
    End If

    ' --- Compose final message
    If Len(prefixText) > 0 Then
        fullMessage = prefixText & " " & messageText
    Else
        fullMessage = messageText
    End If

    ' --- Delegate actual writing to your shared utility
    If FormatType <> "Default" Then
        AV2_WriteSystemTagToDropColumn wsTarget, dropColHeader, targetRow, TargetColLetter, fullMessage, FormatType, FormatMap
        Debug.Print "[AV2_AddValidationFeedback] ? " & devFunctionName & _
            " | DropCol=" & dropColHeader & _
            " | Row=" & targetRow & _
            " | Msg='" & fullMessage & "'"
    Else
        AV2_WriteSystemTagToDropColumn wsTarget, dropColHeader, targetRow, TargetColLetter, fullMessage, FormatType, FormatMap
    End If
    
    


End Sub

' ===== END MERGED: AV2_AutoFeedbackNFormat.bas =====
