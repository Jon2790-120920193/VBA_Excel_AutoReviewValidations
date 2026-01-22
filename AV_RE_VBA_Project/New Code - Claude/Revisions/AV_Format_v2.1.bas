Attribute VB_Name = "AV_Format"
Option Explicit

' ======================================================
' AV_Format v2.1
' Formatting, feedback routing, and utilities
' UPDATED: Uses AV_Constants and AV_DataAccess
' ======================================================

Private Const MODULE_NAME As String = "AV_Format"

' ======================================================
' FORMAT MAP LOADING (Enhanced v2.1)
' ======================================================

Public Function LoadFormatMap(wsConfig As Worksheet) As Object
    Dim tbl As ListObject
    Dim dict As Object
    Dim key As String
    Dim fmt As clsCellFormat
    Dim srcCell As Range
    Dim priorityVal As Long
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Use AV_DataAccess to get table
    Set tbl = AV_DataAccess.GetTable(wsConfig, AV_Constants.TBL_AUTO_FORMAT)
    
    If tbl Is Nothing Then
        AV_Core.DebugMessage "Table '" & AV_Constants.TBL_AUTO_FORMAT & "' not found", MODULE_NAME
        Set LoadFormatMap = dict
        Exit Function
    End If
    
    Dim r As ListRow
    For Each r In tbl.ListRows
        ' Get key
        key = AV_Core.SafeTrim(AV_DataAccess.GetTableValue(wsConfig, AV_Constants.TBL_AUTO_FORMAT, _
                                                           r.Index, AV_Constants.COL_AF_FORMAT_KEY))
        
        If Len(key) = 0 Then GoTo NextRow
        
        ' Get format source cell
        On Error Resume Next
        Set srcCell = r.Range.Cells(1, tbl.ListColumns(AV_Constants.COL_AF_AUTO_FORMATTING).Index)
        On Error GoTo 0
        
        If srcCell Is Nothing Then GoTo NextRow
        
        ' Get priority
        priorityVal = 0
        On Error Resume Next
        priorityVal = CLng(AV_DataAccess.GetTableValue(wsConfig, AV_Constants.TBL_AUTO_FORMAT, _
                                                       r.Index, AV_Constants.COL_AF_PRIORITY))
        On Error GoTo 0
        
        ' Create format object
        Set fmt = getCellFormat(srcCell)
        fmt.Priority = priorityVal
        Set dict(key) = fmt
        
NextRow:
    Next r
    
    AV_Core.DebugMessage "Loaded " & dict.Count & " format mappings", MODULE_NAME
    Set LoadFormatMap = dict
End Function


Public Function DefaultFormatMap() As Object
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Worksheets(AV_Constants.CONFIG_SHEET_NAME)
    
    Set DefaultFormatMap = LoadFormatMap(wsConfig)
    
    If DefaultFormatMap Is Nothing Then
        AV_Core.DebugMessage "Error loading the formatting map", MODULE_NAME
        Exit Function
    End If
    
    AV_Core.DebugMessage "Format Map loaded from default settings", MODULE_NAME
End Function


' ======================================================
' FORMAT KEY CELL (ROW PRIORITY FORMATTING) - Updated v2.1
' ======================================================

Public Sub FormatKeyCell(rowRange As Range, FormatMap As Object)
    ' If multiple rows, call recursively for each row
    Dim r As Range
    If rowRange.Rows.Count > 1 Then
        For Each r In rowRange.Rows
            FormatKeyCell r, FormatMap
        Next r
        Exit Sub
    End If
    
    Dim cell As Range
    Dim key As String
    Dim fmtInfo As clsCellFormat
    Dim currentPriority As Long
    Dim highestPriority As Long
    Dim highestKey As String
    Dim KeyCell As Range
    
    ' Find key cell for this row
    ' NOTE: This still uses column letter from legacy table
    ' Will be updated in Phase 3 to use header-based lookup
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Worksheets(AV_Constants.CONFIG_SHEET_NAME)
    
    Dim colLetter As String
    On Error Resume Next
    colLetter = UCase(wsConfig.Range("B5").Value)
    On Error GoTo 0
    
    If Len(colLetter) = 0 Then
        AV_Core.DebugMessage "Unable to find key column letter", MODULE_NAME
        Exit Sub
    End If
    
    Set KeyCell = rowRange.Worksheet.Range(colLetter & rowRange.Row)
    
    ' Find highest priority format in row
    highestPriority = -1
    highestKey = ""
    
    For Each cell In rowRange.Cells
        key = getFormatType(cell, FormatMap)
        
        If Len(key) > 0 And FormatMap.Exists(key) Then
            Set fmtInfo = FormatMap(key)
            currentPriority = fmtInfo.Priority
            
            If currentPriority > highestPriority Then
                highestPriority = currentPriority
                highestKey = key
            End If
        End If
    Next cell
    
    ' Apply the format with the highest priority
    If Len(highestKey) > 0 Then
        ' Set review status based on priority
        If highestPriority = 2 Then
            Call SetReviewStatus(True, rowRange, True)  ' Auto-corrected
        ElseIf highestPriority = 3 Then
            Call SetReviewStatus(True, rowRange, False) ' Error
        Else
            Call SetReviewStatus(False, rowRange, False) ' No errors
        End If
            
        Call setFormat(KeyCell, highestKey, FormatMap)
    End If
End Sub


Private Sub SetReviewStatus(ReviewRequired As Boolean, rowRange As Range, Optional AutoCorrected As Boolean = False)
    Dim reviewStatus As revStatusRef
    Dim tbl As ListObject
    Dim wsConfig As Worksheet
    Dim rowNum As Long

    Set reviewStatus = New revStatusRef
    Set wsConfig = ThisWorkbook.Worksheets(AV_Constants.CONFIG_SHEET_NAME)

    ' Use AV_DataAccess to get table
    Set tbl = AV_DataAccess.GetTable(wsConfig, AV_Constants.TBL_REVIEW_REF_COLUMNS)

    If tbl Is Nothing Then
        AV_Core.DebugMessage "Table '" & AV_Constants.TBL_REVIEW_REF_COLUMNS & "' not found", MODULE_NAME
        Exit Sub
    End If

    rowNum = rowRange.Row

    ' Get review status column references
    Set reviewStatus.RevStatusCol = GetCellFromTableColumnHeader(tbl, rowRange, "ReviewStatusColumn")
    Set reviewStatus.AutoReviewDropCol = GetCellFromTableColumnHeader(tbl, rowRange, "AutoReviewColumnLetter")
    Set reviewStatus.HumanSetStatusCol = GetCellFromTableColumnHeader(tbl, rowRange, "HumanSetRevStatus")

    ' Set appropriate status using constants
    If ReviewRequired Then
        If AutoCorrected Then
            reviewStatus.AutoReviewDropCol.Value = AV_Constants.STATUS_AUTO_CORRECTED
        Else
            reviewStatus.AutoReviewDropCol.Value = AV_Constants.STATUS_ERROR
        End If
    Else
        reviewStatus.AutoReviewDropCol.Value = AV_Constants.STATUS_NO_ERRORS
    End If
End Sub


' ======================================================
' CELL FORMAT CAPTURE & APPLICATION
' ======================================================

Private Function getCellFormat(cell As Range) As clsCellFormat
    Dim f As New clsCellFormat
    
    With cell
        f.InteriorColor = .Interior.Color
        f.FontColor = .Font.Color
        f.Bold = .Font.Bold
        f.FontName = .Font.Name
        f.FontSize = .Font.Size
        f.NumberFormat = .NumberFormat
        
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


Public Sub setFormat(TargetCell As Range, FormatType As String, FormatMap As Object)
    Dim fmt As clsCellFormat
    
    If FormatMap.Exists(FormatType) Then
        Set fmt = FormatMap(FormatType)
        ApplyCellFormatting TargetCell, fmt
    Else
        AV_Core.DebugMessage "FormatMap type: '" & FormatType & "' not found", MODULE_NAME
    End If
End Sub


Public Function getFormatType(TargetCell As Range, FormatMap As Object) As String
    Dim cellFormat As clsCellFormat
    Dim key As Variant
    Dim fmt As clsCellFormat
    
    Set cellFormat = getCellFormat(TargetCell)
    
    For Each key In FormatMap.Keys
        Set fmt = FormatMap(key)
        
        If FormatsAreEqual(fmt, cellFormat) Then
            getFormatType = key
            Exit Function
        End If
    Next key
    
    getFormatType = vbNullString
End Function


Private Sub ApplyCellFormatting(TargetCell As Range, fmt As clsCellFormat)
    If fmt Is Nothing Then Exit Sub
    
    With TargetCell
        .Interior.Color = fmt.InteriorColor
        .Font.Color = fmt.FontColor
        .Font.Bold = fmt.Bold
        .Font.Name = fmt.FontName
        .Font.Size = fmt.FontSize
        .NumberFormat = fmt.NumberFormat
        
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


Private Function FormatsAreEqual(fmt1 As clsCellFormat, fmt2 As clsCellFormat) As Boolean
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

' ======================================================
' VALIDATION FEEDBACK (CENTRAL HANDLER) - Updated v2.1
' ======================================================

Public Sub AddValidationFeedback(ByVal devFunctionName As String, _
                                 ByVal wsTarget As Worksheet, _
                                 ByVal targetRow As Long, _
                                 ByVal messageText As String, _
                                 Optional ByVal FormatType As String = "", _
                                 Optional ByVal english As Boolean = True, _
                                 Optional FormatMap As Object, _
                                 Optional AutoValMap As Object)

    Dim map As Object
    Dim dropColHeader As String
    Dim prefixText As String
    Dim fullMessage As String
    Dim TargetColLetter As String
    
    ' Ensure function name has correct prefix
    If Left(devFunctionName, 16) <> "Validate_Column_" Then
        devFunctionName = "Validate_Column_" & devFunctionName
    End If
    
    ' Default format type if not specified
    If Len(FormatType) = 0 Then FormatType = AV_Constants.FORMAT_DEFAULT
    
    ' Load format mapping
    If FormatMap Is Nothing Then
        Set FormatMap = DefaultFormatMap()
        If FormatMap Is Nothing Then
            AV_Core.DebugMessage "Error loading the formatting map", MODULE_NAME
            Exit Sub
        End If
    End If
    
    ' Load AutoValidation mapping
    If AutoValMap Is Nothing Then
        AV_Core.DebugMessage "No AutoValidation map loaded, loading now", MODULE_NAME
        Set AutoValMap = AV_Core.GetAutoValidationMap()
        If AutoValMap Is Nothing Then
            AV_Core.DebugMessage "Error loading the autovalidation mapping", MODULE_NAME
            Exit Sub
        End If
    End If
    
    ' Look up function in the map
    If Not AutoValMap.Exists(devFunctionName) Then
        AV_Core.DebugMessage "Dev function '" & devFunctionName & "' not found in mapping table", MODULE_NAME
        Exit Sub
    End If

    Set map = AutoValMap(devFunctionName)
    dropColHeader = AV_Core.SafeTrim(map("DropColHeader"))
    TargetColLetter = AV_Core.SafeTrim(map("ColumnRef"))
    
    ' Get language-appropriate prefix
    If english Then
        prefixText = AV_Core.SafeTrim(map("PrefixEN"))
    Else
        prefixText = AV_Core.SafeTrim(map("PrefixFR"))
    End If

    ' Compose final message
    If Len(prefixText) > 0 And Len(messageText) > 0 Then
        fullMessage = prefixText & " " & messageText
    ElseIf Len(prefixText) > 0 Then
        fullMessage = prefixText
    Else
        fullMessage = messageText
    End If

    ' Delegate actual writing
    WriteSystemTagToDropColumn wsTarget, dropColHeader, targetRow, TargetColLetter, fullMessage, FormatType, FormatMap
End Sub


' ======================================================
' SYSTEM TAG MESSAGE UTILITIES - Updated v2.1
' ======================================================

Public Sub WriteSystemTagToDropColumn(wsTarget As Worksheet, _
                                      dropColLetter As String, _
                                      rowNum As Long, _
                                      sourceColLetter As String, _
                                      tagText As String, _
                                      Optional FormatType As String = "", _
                                      Optional FormatMap As Object)

    On Error GoTo ErrHandler
    Dim cell As Range
    Dim tagId As String
    Dim fullMsg As String
    Dim existingText As String
    Dim cleanedText As String
    Dim TrgCell As Range

    If wsTarget Is Nothing Then Exit Sub
    If Len(dropColLetter) = 0 Or rowNum <= 0 Then Exit Sub

    ' Default format type
    If Len(FormatType) = 0 Then FormatType = AV_Constants.FORMAT_DEFAULT

    ' Get source and drop cells
    Set TrgCell = GetCellByLetter(wsTarget, sourceColLetter, rowNum)
    Set cell = wsTarget.Range(dropColLetter & rowNum)
    If cell Is Nothing Then Exit Sub

    tagId = "Col " & sourceColLetter

    Application.EnableEvents = False

    ' Always clear any existing tag for this column
    Call ClearSystemTagFromString_KeepOthers(cell, tagId)

    ' Handle "Default" case - cleanup only
    If FormatType = AV_Constants.FORMAT_DEFAULT Then
        If Not FormatMap Is Nothing Then
            setFormat TrgCell, AV_Constants.FORMAT_DEFAULT, FormatMap
        Else
            setFormat TrgCell, AV_Constants.FORMAT_DEFAULT, DefaultFormatMap
        End If
        cell.Value = Trim(cell.Value)
        Application.EnableEvents = True
        Exit Sub
    End If

    ' Build message and append
    fullMsg = AV_Constants.SYSTEM_TAG_START & " " & tagId & ": " & tagText & " " & AV_Constants.SYSTEM_TAG_END

    If FormatMap Is Nothing Then
        Set FormatMap = DefaultFormatMap
    End If

    ' Apply formatting to the triggering cell
    setFormat TrgCell, FormatType, FormatMap

    ' Append the message cleanly
    existingText = Trim(Replace(cell.Value, vbCr, ""))
    Do While Right(existingText, 1) = vbLf
        existingText = Left(existingText, Len(existingText) - 1)
    Loop

    If existingText <> "" Then
        cleanedText = existingText & vbLf & fullMsg
    Else
        cleanedText = fullMsg
    End If

    cell.Value = cleanedText

CleanExit:
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    AV_Core.DebugMessage "WriteSystemTagToDropColumn ERROR: " & Err.Number & " - " & Err.Description, MODULE_NAME
    Resume CleanExit
End Sub


Public Sub ClearSystemTagFromString_KeepOthers(TargetCell As Range, tagId As String)
    Dim txt As String
    Dim sPos As Long, ePos As Long, subLen As Long
    Dim tagStart As String, tagEnd As String
    Dim chunk As String
    Dim attempts As Integer

    If TargetCell Is Nothing Then Exit Sub
    If Len(tagId) = 0 Then Exit Sub

    tagStart = AV_Constants.SYSTEM_TAG_START & " " & tagId & ":"
    tagEnd = AV_Constants.SYSTEM_TAG_END
    txt = CStr(TargetCell.Value)

    If Len(txt) = 0 Then Exit Sub

    Application.EnableEvents = False

    attempts = 0

    Do
        sPos = InStr(1, txt, tagStart, vbTextCompare)
        If sPos = 0 Then Exit Do

        ePos = InStr(sPos, txt, tagEnd, vbTextCompare)
        If ePos = 0 Then
            txt = Left(txt, sPos - 1) & Mid(txt, sPos + Len(tagStart))
        Else
            subLen = ePos - sPos + Len(tagEnd)
            chunk = Mid(txt, sPos, subLen)
            txt = Replace(txt, chunk, "", , 1, vbTextCompare)
        End If

        attempts = attempts + 1
        If attempts > 20 Then Exit Do
    Loop

    ' Normalize whitespace
    txt = Trim(Replace(txt, vbCr, ""))
    Do While InStr(txt, vbLf & vbLf) > 0
        txt = Replace(txt, vbLf & vbLf, vbLf)
    Loop
    If Left(txt, 1) = vbLf Then txt = Mid(txt, 2)
    If Right(txt, 1) = vbLf Then txt = Left(txt, Len(txt) - 1)

    TargetCell.Value = txt

    Application.EnableEvents = True
End Sub


' ======================================================
' UTILITY HELPERS - Updated v2.1
' ======================================================

Public Function GetCellFromTableColumnHeader(tbl As ListObject, rowRange As Range, ColumnHeader As String) As Range
    Dim colLetter As String
    Dim colIndex As Long
    Dim headerValue As String

    On Error Resume Next
    colIndex = tbl.ListColumns(ColumnHeader).Index
    On Error GoTo 0
    
    If colIndex = 0 Then
        Err.Raise vbObjectError + 515, MODULE_NAME, _
            "Column '" & ColumnHeader & "' not found in table '" & tbl.Name & "'"
        Exit Function
    End If

    ' Get the column letter stored in the first data row
    headerValue = Trim(CStr(tbl.DataBodyRange.Cells(1, colIndex).Value))
    
    If Len(headerValue) = 0 Then
        Err.Raise vbObjectError + 516, MODULE_NAME, _
            "No column letter found under header '" & ColumnHeader & "' in table '" & tbl.Name & "'"
        Exit Function
    End If
    
    Set GetCellFromTableColumnHeader = rowRange.Worksheet.Range(headerValue & rowRange.Row)
End Function


Public Function GetCellByLetter(ws As Worksheet, colLetter As String, rowNum As Long) As Range
    Dim colNum As Long
    On Error GoTo ErrHandler
    
    colNum = Range(colLetter & "1").Column
    Set GetCellByLetter = ws.Cells(rowNum, colNum)
    Exit Function

ErrHandler:
    AV_Core.DebugMessage "Invalid column letter: " & colLetter, MODULE_NAME
    Set GetCellByLetter = Nothing
End Function
