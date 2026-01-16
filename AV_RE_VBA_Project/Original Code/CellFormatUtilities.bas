Attribute VB_Name = "CellFormatUtilities"
Option Explicit

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
' Load all format objects (clsCellFormat) into memory
'====================================================
Public Function LoadFormatMap(wsConfig As Worksheet) As Object
    Dim Table As ListObject
    Dim dict As Object
    Dim r As ListRow
    Dim key As String
    Dim fmt As clsCellFormat
    Dim srcCell As Range
    Dim colKey As Long, colFormat As Long, colPriority As Long
    Dim priorityVal As Long
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    Set Table = wsConfig.ListObjects(GlobalFormatTableName)
    On Error GoTo 0
    
    If Table Is Nothing Then
        Debug.Print "?? Table '" & GlobalFormatTableName & "' not found on " & wsConfig.Name
        Set LoadFormatMap = dict
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
            fmt.Priority = priorityVal   ' <— new property assignment
            Set dict(key) = fmt
        End If
    Next r
    
    Set LoadFormatMap = dict
End Function

Public Function DefaultFormatMap() As Object
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Worksheets("Config")
    Dim FormatMap As Object
    
    Set FormatMap = LoadFormatMap(wsConfig)
    
    If FormatMap Is Nothing Then
        Debug.Print "Error loading the formatting map from the CellFormatUtilities Module"
        Exit Function
    End If
    
    
    DebugMessage "Format Map was loaded from default settings. Loaded from: '" & "Config" & "' Workbook Sheet", "DefaultFormatMap", DebugON
    Set DefaultFormatMap = FormatMap
    
End Function


' Receives Row Formatting and detects the Highest Level of Correction
' Sets the Key Column with the respective Priority Level, e.g. Error, AutoCorrect...etc

Public Sub FormatKeyCell(rowRange As Range, FormatMap As Object)
    ' --- If multiple rows, call recursively for each row ---
    Dim r As Range
    If rowRange.Rows.count > 1 Then
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
        key = getFormatType(cell, FormatMap)
        
        If Len(key) > 0 Then
            ' Access the clsCellFormat directly
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
            
        Call setFormat(KeyCell, highestKey, FormatMap)
    End If
End Sub

Private Sub SetReviewStatus(ReviewRequired As Boolean, rowRange As Range, Optional AutoCorrected As Boolean = False)

    Dim reviewStatus As revStatusRef
    Dim Table As ListObject
    Dim wsConfig As Worksheet
    Dim RowNum As Long

    ' Instantiate the class
    Set reviewStatus = New revStatusRef
    
    Set wsConfig = ThisWorkbook.Worksheets(ConfigSheet)

    On Error Resume Next
    Set Table = wsConfig.ListObjects(ReviewFlagsTableName)
    On Error GoTo 0

    If Table Is Nothing Then
        DebugMessage "?? Table '" & GlobalFormatTableName & "' not found on " & wsConfig.Name, "SetReviewStatus", DebugON
        Exit Sub
    End If

    RowNum = rowRange.row

    ' === Get the proper cells ===
    Set reviewStatus.RevStatusCol = GetCellFromTableColumnHeader(Table, rowRange, RevStatusColName)
    Set reviewStatus.AutoReviewDropCol = GetCellFromTableColumnHeader(Table, rowRange, AutoRevStatusColName)
    Set reviewStatus.HumanSetStatusCol = GetCellFromTableColumnHeader(Table, rowRange, HumanSetStatus)

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
' Capture a cell's complete format into a clsCellFormat
'====================================================
Private Function getCellFormat(cell As Range) As clsCellFormat
    Dim f As New clsCellFormat
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

Public Sub setFormat(TargetCell As Range, FormatType As String, FormatMap As Object)
    Dim fmt As clsCellFormat
    
    If FormatMap.Exists(FormatType) Then
        Set fmt = FormatMap(FormatType)
        ApplyCellFormatting TargetCell, fmt
    Else
        DebugMessage "FormatMap type: '" & FormatType & "' could not be found in the '" & GlobalFormatTableName & "' Config table", "setFormat", DebugON
    End If
    
End Sub

Public Function getFormatType(TargetCell As Range, FormatMap As Object) As String
    Dim cellFormat As clsCellFormat
    Dim key As Variant
    Dim fmt As clsCellFormat
    
    ' Get the format of the target cell
    Set cellFormat = getCellFormat(TargetCell)
    
    ' Loop through all known formats in the map
    For Each key In FormatMap.Keys
        Set fmt = FormatMap(key)
        
        If FormatsAreEqual(fmt, cellFormat) Then
            getFormatType = key
            Exit Function
        End If
    Next key
    
    ' If no match found
    getFormatType = vbNullString
End Function


'====================================================
' Apply stored format to a target cell
'====================================================
Private Sub ApplyCellFormatting(TargetCell As Range, fmt As clsCellFormat)
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




