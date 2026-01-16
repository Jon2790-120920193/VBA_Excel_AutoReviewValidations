Attribute VB_Name = "LoadSmartFunctionReferences"
Option Explicit

' Module-level cache
Private gAutoValidationMap As Object
Private Const MAPPING_TABLE_NAME As String = "AutoValidationCommentPrefixMappingTable"

' ===============================================
' Returns global AutoValidation map (cached)
' Key = "Validate_Column_<FunctionName>"
' Value = Dictionary with keys:
'   - DropColHeader
'   - PrefixEN
'   - PrefixFR
'   - ColumnRef
' ===============================================
Public Function GetAutoValidationMap(Optional wsConfig As Worksheet) As Object
    On Error GoTo ErrHandler

    ' --- Use cached version if already loaded ---
    If Not gAutoValidationMap Is Nothing Then
        Set GetAutoValidationMap = gAutoValidationMap
        Exit Function
    End If

    ' --- Default sheet ---
    If wsConfig Is Nothing Then
        Set wsConfig = ThisWorkbook.Sheets("Config")
    End If

    ' --- Try to get table ---
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = wsConfig.ListObjects(MAPPING_TABLE_NAME)
    On Error GoTo 0

    If tbl Is Nothing Then
        Debug.Print "[AutoValidationMapping] ERROR: Table '" & MAPPING_TABLE_NAME & "' not found in sheet '" & wsConfig.Name & "'."
        Set gAutoValidationMap = CreateObject("Scripting.Dictionary")
        Set GetAutoValidationMap = gAutoValidationMap
        Exit Function
    End If

    ' --- Validate required headers ---
    Dim requiredHeaders As Variant
    requiredHeaders = Array("Dev Function Names", "Drop in Column", "Prefix to message", "(FR) Prefix to message", "ReviewSheet Column Letter", "AutoValidate")

    Dim hdr As Variant, missingHeaders As String
    For Each hdr In requiredHeaders
        If Not ColumnExists(tbl, CStr(hdr)) Then
            missingHeaders = missingHeaders & vbNewLine & " - " & hdr
        End If
    Next hdr

    If Len(missingHeaders) > 0 Then
        Debug.Print "[AutoValidationMapping] WARNING: Missing columns in table '" & MAPPING_TABLE_NAME & "':" & missingHeaders
        Set gAutoValidationMap = CreateObject("Scripting.Dictionary")
        Set GetAutoValidationMap = gAutoValidationMap
        Exit Function
    End If

    ' --- Build master dictionary ---
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim lr As ListRow
    Dim devFunc As String, dropCol As String, prefixEN As String, prefixFR As String, colRef As String, AutoValidate As Boolean
    
    Dim item As Object
    Dim successCount As Long

    For Each lr In tbl.ListRows
        devFunc = "Validate_Column_" & SafeTrim(lr.Range.Cells(1, tbl.ListColumns("Dev Function Names").Index).value)
        dropCol = SafeTrim(lr.Range.Cells(1, tbl.ListColumns("Drop in Column").Index).value)
        prefixEN = SafeTrim(lr.Range.Cells(1, tbl.ListColumns("Prefix to message").Index).value)
        prefixFR = SafeTrim(lr.Range.Cells(1, tbl.ListColumns("(FR) Prefix to message").Index).value)
        colRef = SafeTrim(lr.Range.Cells(1, tbl.ListColumns("ReviewSheet Column Letter").Index).value)
        AutoValidate = CBoolString(SafeTrim(lr.Range.Cells(1, tbl.ListColumns("AutoValidate").Index).value))

        If devFunc <> "" Then
            Set item = CreateObject("Scripting.Dictionary")
            item("DropColHeader") = dropCol
            item("PrefixEN") = prefixEN
            item("PrefixFR") = prefixFR
            item("ColumnRef") = colRef
            item("AutoValidate") = AutoValidate

            If Not dict.Exists(devFunc) Then
                dict.Add devFunc, item
                successCount = successCount + 1
            Else
                Debug.Print "[AutoValidationMapping] WARNING: Duplicate function name skipped: " & devFunc
            End If
        End If
    Next lr

    Debug.Print "[AutoValidationMapping] Loaded " & successCount & " valid mappings from table '" & MAPPING_TABLE_NAME & "'."

    Set gAutoValidationMap = dict
    Set GetAutoValidationMap = gAutoValidationMap
    Exit Function

ErrHandler:
    Debug.Print "[AutoValidationMapping] ERROR: " & Err.Number & " - " & Err.Description
    Set gAutoValidationMap = CreateObject("Scripting.Dictionary")
    Set GetAutoValidationMap = gAutoValidationMap
End Function


' ===============================================
' Helper: Refresh or Reset cache manually
' ===============================================
Public Sub ResetAutoValidationMap(Optional wsConfig As Worksheet)
    Set gAutoValidationMap = Nothing
    Debug.Print "[AutoValidationMapping] Cache cleared."
    If Not wsConfig Is Nothing Then
        ' Force reload
        Dim temp As Object
        Set temp = GetAutoValidationMap(wsConfig)
    End If
End Sub



