Attribute VB_Name = "AV_Engine"
Option Explicit

' ======================================================
' AV_Engine.bas
' Validation orchestration & execution engine
' ======================================================

' -----------------------------
' Module-level state (MUST be at top)
' -----------------------------
Public ValidationStartTime As Single
Public ValidationCancelTimeout As Single
Public ValidationCancelFlag As Boolean

Private Const MODULE_NAME As String = "AV_Engine"

' ======================================================
' PUBLIC ENTRY POINTS
' ======================================================

' Legacy-safe wrapper (used by buttons / macros)
Public Sub RunFullValidation(Optional ByVal sheetName As String = "", Optional ByVal english As Boolean = True)
    RunFullValidationMaster sheetName, english
End Sub

' ======================================================
' MAIN VALIDATION EXECUTION
' ======================================================
Public Sub RunFullValidationMaster(Optional ByVal sheetName As String = "", Optional ByVal english As Boolean = True)

    Dim wsConfig As Worksheet
    Dim wsTarget As Worksheet
    Dim dataSheetName As String
    Dim startRow As Long, endRow As Long
    Dim keyColLetter As String, keyColNum As Long
    Dim rowNum As Long, i As Long
    Dim keyRows() As Long
    Dim keyCount As Long, maxRows As Long
    Dim rowValues As Variant

    Dim AdvFunctionMap As Object
    Dim FormatMap As Object
    Dim colMetaDict As Object
    Dim colReviewedColumnList As Collection

    On Error GoTo ErrHandler

    ' -----------------------------
    ' Initialize UI / logging
    ' -----------------------------
    ShowValidationTrackerForm
    AppendUserLog "Initializing Full Validation Master"

    ' -----------------------------
    ' Cancel / timeout flags
    ' -----------------------------
    ValidationStartTime = Timer
    ValidationCancelTimeout = 10000
    ValidationCancelFlag = False

    AppendUserLog "Validation timeout set to " & ValidationCancelTimeout & " seconds"

    ' -----------------------------
    ' Load configuration
    ' -----------------------------
    Set wsConfig = ThisWorkbook.Sheets("Config")

    If sheetName = "" Then
        dataSheetName = Trim(wsConfig.Range("B3").Value)
    Else
        dataSheetName = sheetName
    End If

    Set wsTarget = ThisWorkbook.Sheets(dataSheetName)

    startRow = CLng(wsConfig.Range("B4").Value)
    endRow = startRow + CLng(wsConfig.Range("D4").Value)

    keyColLetter = Trim(wsConfig.Range("B5").Value)
    keyColNum = wsTarget.Range(keyColLetter & "1").Column

    AppendUserLog "Target sheet: " & dataSheetName
    AppendUserLog "Row range: " & startRow & " to " & endRow

    ' -----------------------------
    ' Load mappings
    ' -----------------------------
    Set AdvFunctionMap = GetAutoValidationMap(wsConfig)
    Set FormatMap = LoadFormatMap(wsConfig)
    Set colMetaDict = GetDDMValidationColumns(wsConfig)

    If AdvFunctionMap Is Nothing Or AdvFunctionMap.count = 0 Then
        AppendUserLog "No validation functions mapped. Aborting."
        GoTo Cleanup
    End If

    ' -----------------------------
    ' Pre-compute rows with keys
    ' -----------------------------
    maxRows = endRow - startRow + 1
    ReDim keyRows(1 To maxRows)
    keyCount = 0

    For rowNum = startRow To endRow
        If Trim(CStr(wsTarget.Cells(rowNum, keyColNum).Value)) <> "" Then
            keyCount = keyCount + 1
            keyRows(keyCount) = rowNum
        End If
    Next rowNum

    If keyCount = 0 Then
        AppendUserLog "No valid rows found. Exiting."
        GoTo Cleanup
    End If

    ReDim Preserve keyRows(1 To keyCount)

    ' -----------------------------
    ' MAIN ROW LOOP
    ' -----------------------------
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For i = LBound(keyRows) To UBound(keyRows)
        rowNum = keyRows(i)

        If i Mod 10 = 0 Then DoEvents

        If ValidationCancelFlag Then
            AppendUserLog "Validation cancelled by user."
            GoTo Cleanup
        End If

        If ValidationTimeoutReached() Then
            AppendUserLog "Validation stopped due to timeout."
            GoTo Cleanup
        End If

        If ShouldValidateRow(rowNum, wsTarget, True) Then
            rowValues = wsTarget.Rows(rowNum).Value
            ValidateSingleRow wsTarget, rowNum, AdvFunctionMap, english, FormatMap
        End If
    Next i

    ' -----------------------------
    ' Post-pass: simple data validation
    ' -----------------------------
    Set colReviewedColumnList = BuildCollectionOfColumnHeaders(colMetaDict, AdvFunctionMap)

    RunAutoCheckDataValidation wsConfig, wsTarget, keyRows, keyColNum, english, _
                               FormatMap, colMetaDict, colReviewedColumnList

    AppendUserLog "Advanced Auto Validation completed."

Cleanup:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    AppendUserLog "ERROR in RunFullValidationMaster"
    AppendUserLog "Error #" & Err.Number & ": " & Err.Description
    BringFormToFront ValidationTrackerForm
    Resume Cleanup

End Sub


