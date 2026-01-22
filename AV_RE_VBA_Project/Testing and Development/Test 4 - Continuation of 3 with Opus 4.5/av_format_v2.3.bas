Attribute VB_Name = "AV_Format"
Option Explicit

' ======================================================
' AV_Format
' Formatting, feedback routing, and utilities
' VERSION: 2.3 - Uses table header lookup
' ======================================================

Private Const MODULE_NAME As String = "AV_Format"

' -----------------------------
' CONSTANTS
' -----------------------------
Private Const FormatKeyColumn As String = "Formatting Key"
Private Const AutoFormatColumn As String = "Autoformatting"
Private Const GlobalFormatTableName As String = "AutoFormatOnFullValidation"
Private Const PriorityColumn As String = "KeyFlagPriority"
Private Const ConfigSheet As String = "Config"
Private Const ReviewFlagsTableName As String = "ReviewRefColumnTable"
Private Const RevStatusColName As String = "ReviewStatusColumn"
Private Const AutoRevStatusColName As String = "AutoReviewColumnLetter"
Private Const HumanSetStatus As String = "HumanSetRevStatus"

' System tags (for message tagging)
Public Const SYSTEM_TAG_START As String = "[[SYS_TAG"
Public Const SYSTEM_TAG_END As String = "]]"
Public Const SYSTEM_COMMENT_TAG As String = "[[SYS_COMMENT]]"
Private Const FALLBACKFORMAT As String = "Default"


' ======================================================
' FORMAT MAP LOADING
' ======================================================

Public Function LoadFormatMap(wsConfig As Worksheet) As Object
    Dim tbl As ListObject
    Dim dict As Object
    Dim r As ListRow
    Dim key As String
    Dim fmt As clsCellFormat
    Dim srcCell As Range
    Dim colKey As Long, colFormat As Long, colPriority As Long
    Dim priorityVal As Long
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    Set tbl = wsConfig.ListObjects(GlobalFormatTableName)
    On Error GoTo 0
    
    If tbl Is Nothing Then
        AV_Core.DebugMessage "Table '" & GlobalFormatTableName & "' not found on " & wsConfig.Name, MODULE_NAME
        Set LoadFormatMap = dict
        Exit Function
    End If
    
    colKey = tbl.ListColumns(FormatKeyColumn).Index
    colFormat = tbl.ListColumns(AutoFormatColumn).Index
    colPriority = tbl.ListColumns(PriorityColumn).Index
    
    For Each r In tbl.ListRows
        key = Trim(r.Range.Cells(1, colKey).Value)
        Set srcCell = r.Range.Cells(1, colFormat)
        
        If IsNumeric(r.Range.Cells(1, colPriority).Value) Then
            priorityVal = CLng(r.Range.Cells(1, colPriority).Value)
        Else
            priorityVal = 0
        End If

        If Len(key) > 0 Then
            Set fmt = getCellFormat(srcCell)
            fmt.Priority = priorityVal
            Set dict(key) = fmt
        End If
    Next r
    
    AV_Core.DebugMessage "Loaded " & dict.Count & " format definitions", MODULE_NAME
    Set LoadFormatMap = dict
End Function


Public Function DefaultFormatMap() As Object
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Worksheets("Config")
    
    Set DefaultFormatMap = LoadFormatMap(wsConfig)
End Function


' ======================================================
' FORMAT KEY CELL (ROW PRIORITY FORMATTING)
' ======================================================

Public Sub FormatKeyCell(rowRange As Range, FormatMap As Object)
    If rowRange Is Nothing Then Exit Sub
    
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
    
    ' Get the key column cell for this row using the target table
    Dim targetTable As ListObject
    Set targetTable = AV_Engine.CurrentTargetTable
    
    If targetTable Is Nothing Then
        AV_Core.DebugMessage "FormatKeyCell: No target table available", MODULE_NAME
        Exit Sub
    End If
    
    ' Use first column as key column
    Dim keyColNum As Long
    keyColNum = targetTable.ListColumns(1).DataBodyRange.Column
    Set KeyCell = rowRange.Worksheet.Cells(rowRange.Row, keyColNum)
    
    highestPriority = -1
    highestKey = ""
    
    For Each cell In rowRange.Cells
        key = getFormatType(cell, FormatMap)
        
        If Len(key) > 0 Then
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
        If highestPriority = 2 Then
            Call SetReviewStatus(True, rowRange, True)
        ElseIf highestPriority = 3 Then
            Call SetReviewStatus(True, rowRange, False)
        Else
            Call SetReviewStatus(False, rowRange, False)
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
    Set wsConfig = ThisWorkbook.Worksheets(ConfigSheet)

    On Error Resume Next
    Set tbl = wsConfig.ListObjects(ReviewFlagsTableName)
    On Error GoTo 0

    If tbl Is Nothing Then
        AV_Core.DebugMessage "Table '" & ReviewFlagsTableName & "' not found", MODULE_NAME
        Exit Sub
    End If

    rowNum = rowRange.Row

    Set reviewStatus.RevStatusCol = GetCellFromTableColumnHeader(tbl, rowRange, RevStatusColName)
    Set reviewStatus.AutoReviewDropCol = GetCellFromTableColumnHeader(tbl, rowRange, AutoRevStatusColName)
    Set reviewStatus.HumanSetStatusCol = GetCellFromTableColumnHeader(tbl, rowRange, HumanSetStatus)

    If Not reviewStatus.AutoReviewDropCol Is Nothing Then
        If ReviewRequired Then
            If AutoCorrected Then
                reviewStatus.AutoReviewDropCol.Value = "Auto Corrected"
            Else
                reviewStatus.AutoReviewDropCol.Value = "Error"
            End If
        Else
            reviewStatus.AutoReviewDropCol.Value = "No Errors Found"
        End If
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
    
    If TargetCell Is Nothing Then
        AV_Core.DebugMessage "setFormat: TargetCell is Nothing", MODULE_NAME
        Exit Sub
    End If
    
    If FormatMap Is Nothing Then
        AV_Core.DebugMessage "setFormat: FormatMap is Nothing", MODULE_NAME
        Exit Sub
    End If
    
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
    
    If TargetCell Is Nothing Or FormatMap Is Nothing Then Exit Function
    
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
' VALIDATION FEEDBACK (CENTRAL HANDLER)
' ======================================================

Public Sub AddValidationFeedback(ByVal devFunctionName As String, _
                                 ByVal wsTarget As Worksheet, _
                                 ByVal targetRow As Long, _
                                 ByVal messageText As String, _
                                 Optional ByVal FormatType As String = "Default", _
                                 Optional ByVal english As Boolean = True, _
                                 Optional FormatMap As Object, _
                                 Optional AutoValMap As Object)

    Dim map As Object
    Dim dropColHeader As String
    Dim prefixText As String
    Dim fullMessage As String
    Dim TargetColHeader As String
    
    devFunctionName = "Validate_Column_" & devFunctionName
    
    ' Load format mapping
    If FormatMap Is Nothing Then
        Set FormatMap = DefaultFormatMap()
        If FormatMap Is Nothing Then
            AV_Core.DebugMessage "AddValidationFeedback: Error loading format map", MODULE_NAME
            Exit Sub
        End If
    End If
    
    ' Load Smart Autovalidation mapping
    If AutoValMap Is Nothing Then
        Set AutoValMap = AV_Core.GetAutoValidationMap()
        If AutoValMap Is Nothing Then
            AV_Core.DebugMessage "AddValidationFeedback: Error loading autovalidation mapping", MODULE_NAME
            Exit Sub
        End If
    End If
    
    ' Look up called function in the map
    If Not AutoValMap.Exists(devFunctionName) Then
        AV_Core.DebugMessage "AddValidationFeedback: Function '" & devFunctionName & "' not found in mapping", MODULE_NAME
        Exit Sub
    End If

    Set map = AutoValMap(devFunctionName)
    dropColHeader = AV_Core.SafeTrim(map("DropColHeader"))
    TargetColHeader = AV_Core.SafeTrim(map("ColumnRef"))
    
    If english Then
        prefixText = AV_Core.SafeTrim(map("PrefixEN"))
    Else
        prefixText = AV_Core.SafeTrim(map("PrefixFR"))
    End If

    ' Compose final message
    If Len(prefixText) > 0 Then
        fullMessage = prefixText & " " & messageText
    Else
        fullMessage = messageText
    End If

    ' Delegate actual writing - pass header names
    WriteSystemTagToDropColumn wsTarget, dropColHeader, targetRow, TargetColHeader, fullMessage, FormatType, FormatMap
End Sub


' ======================================================
' SYSTEM TAG MESSAGE UTILITIES
' Updated to use table header lookup
' ======================================================

Public Sub WriteSystemTagToDropColumn(wsTarget As Worksheet, _
                                      dropColHeader As String, _
                                      rowNum As Long, _
                                      sourceColHeader As String, _
                                      tagText As String, _
                                      Optional FormatType As String = FALLBACKFORMAT, _
                                      Optional FormatMap As Object)

    On Error GoTo ErrHandler
    Dim cell As Range
    Dim tagId As String
    Dim fullMsg As String
    Dim existingText As String
    Dim cleanedText As String
    Dim TrgCell As Range

    If wsTarget Is Nothing Then Exit Sub
    If Len(dropColHeader) = 0 Or rowNum <= 0 Then Exit Sub

    ' Get the target table from engine
    Dim targetTable As ListObject
    Set targetTable = AV_Engine.CurrentTargetTable
    
    If targetTable Is Nothing Then
        AV_Core.DebugMessage "WriteSystemTagToDropColumn: No target table available", MODULE_NAME
        Exit Sub
    End If
    
    ' Use table header lookup for both cells
    Set TrgCell = AV_DataAccess.GetCellByTableHeader(targetTable, sourceColHeader, rowNum)
    Set cell = AV_DataAccess.GetCellByTableHeader(targetTable, dropColHeader, rowNum)
    
    If TrgCell Is Nothing Then
        AV_Core.DebugMessage "WriteSystemTagToDropColumn: Source column '" & sourceColHeader & "' not found", MODULE_NAME
        Exit Sub
    End If
    
    If cell Is Nothing Then
        AV_Core.DebugMessage "WriteSystemTagToDropColumn: Drop column '" & dropColHeader & "' not found", MODULE_NAME
        Exit Sub
    End If

    ' Build tag ID using abbreviated header
    If Len(sourceColHeader) > 15 Then
        tagId = "Col " & Left(sourceColHeader, 12) & "..."
    Else
        tagId = "Col " & sourceColHeader
    End If

    Application.EnableEvents = False

    ' Always clear any existing tag for this column
    Call ClearSystemTagFromString_KeepOthers(cell, tagId)

    ' Handle "Default" case - cleanup only
    If FormatType = "Default" Then
        If Not FormatMap Is Nothing Then
            setFormat TrgCell, "Default", FormatMap
        Else
            setFormat TrgCell, "Default", DefaultFormatMap
        End If
        cell.Value = Trim(cell.Value)
        Application.EnableEvents = True
        Exit Sub
    End If

    ' Build message and append
    fullMsg = SYSTEM_TAG_START & " " & tagId & ": " & tagText & " " & SYSTEM_TAG_END

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

    tagStart = SYSTEM_TAG_START & " " & tagId & ":"
    tagEnd = SYSTEM_TAG_END
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
' UTILITY HELPERS
' ======================================================

Public Function GetCellFromTableColumnHeader(tbl As ListObject, rowRange As Range, ColumnHeader As String) As Range
    ' This function reads a column LETTER from a config table and returns the cell
    ' Used for ReviewRefColumnTable which stores column letters
    Dim colIndex As Long
    Dim headerValue As String

    On Error Resume Next
    colIndex = tbl.ListColumns(ColumnHeader).Index
    On Error GoTo 0
    
    If colIndex = 0 Then
        Exit Function
    End If

    ' Get the column letter stored in the table
    headerValue = Trim(CStr(tbl.DataBodyRange.Cells(1, colIndex).Value))
    
    If Len(headerValue) = 0 Then
        Exit Function
    End If
    
    ' If it looks like a column letter (1-3 uppercase letters), use Range
    If Len(headerValue) <= 3 Then
        On Error Resume Next
        Set GetCellFromTableColumnHeader = rowRange.Worksheet.Range(headerValue & rowRange.Row)
        On Error GoTo 0
    Else
        ' It might be a header name - try table lookup
        Dim targetTable As ListObject
        Set targetTable = AV_Engine.CurrentTargetTable
        If Not targetTable Is Nothing Then
            Set GetCellFromTableColumnHeader = AV_DataAccess.GetCellByTableHeader(targetTable, headerValue, rowRange.Row)
        End If
    End If
End Function
