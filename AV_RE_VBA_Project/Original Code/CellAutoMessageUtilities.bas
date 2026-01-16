Attribute VB_Name = "CellAutoMessageUtilities"
Private Const MAPPING_TABLE_NAME As String = "AutoValidationCommentPrefixMappingTable"

' ===================================================================
' ClearSystemTagFromString_KeepOthers
' -------------------------------------------------------------------
' Removes only the tag with the matching tagId (e.g., "Col B")
' from the cell value, preserving other system tags.
' ===================================================================

Option Explicit

' -----------------------------
' SYSTEM TAG MARKERS
' -----------------------------
Public Const SYSTEM_TAG_START As String = "[[SYS_TAG"
Public Const SYSTEM_TAG_END As String = "]]"

' -----------------------------
' SYSTEM COMMENT IDENTIFIER
' Used for legacy comment cleanup
' -----------------------------
Public Const SYSTEM_COMMENT_TAG As String = "[[SYS_COMMENT]]"

' Must be set to the a value (case sensitive) in the AutoFormatOnFullValidation table on the Config table.
' Found under the Formatting Key column

Private Const FALLBACKFORMAT As String = "Default"

' -----------------------------
' TAGGING STYLE GUIDE
' Example tag format in cells:
' [[SYS_TAG Col B: Invalid value 'XYZ' : Select a valid option.]]
'
' Each tag begins with SYSTEM_TAG_START
'     followed by " Col <Letter>:" and the message
' and ends with SYSTEM_TAG_END.
'
' These are written by WriteSystemTagToDropColumn
' and cleared by ClearSystemTagFromString_KeepOthers.
' -----------------------------


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
    txt = CStr(TargetCell.value)

    If Len(txt) = 0 Then Exit Sub

    Application.EnableEvents = False

    ' Defensive: limit max cleanup attempts per cell
    attempts = 0

    ' Loop until no more matches
    Do
        sPos = InStr(1, txt, tagStart, vbTextCompare)
        If sPos = 0 Then Exit Do

        ePos = InStr(sPos, txt, tagEnd, vbTextCompare)
        If ePos = 0 Then
            ' Broken tag (no end marker)
            txt = Left(txt, sPos - 1) & Mid(txt, sPos + Len(tagStart))
        Else
            subLen = ePos - sPos + Len(tagEnd)
            chunk = Mid(txt, sPos, subLen)
            txt = Replace(txt, chunk, "", , 1, vbTextCompare)
        End If

        attempts = attempts + 1
        If attempts > 20 Then Exit Do ' Prevent infinite loop
    Loop

    ' Normalize whitespace and remove excess blank lines
    txt = Trim(Replace(txt, vbCr, ""))
    Do While InStr(txt, vbLf & vbLf) > 0
        txt = Replace(txt, vbLf & vbLf, vbLf)
    Loop
    If Left(txt, 1) = vbLf Then txt = Mid(txt, 2)
    If Right(txt, 1) = vbLf Then txt = Left(txt, Len(txt) - 1)

    ' Commit final text
    TargetCell.value = txt

    Application.EnableEvents = True
End Sub

' Write a single system-tagged message block into the requested drop column cell.
' The tag includes the source column letter so multiple source-columns can write to the same drop column safely.
Public Sub WriteSystemTagToDropColumn(wsTarget As Worksheet, _
                                      dropColLetter As String, _
                                      RowNum As Long, _
                                      sourceColLetter As String, _
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


    ' Validate inputs
    If wsTarget Is Nothing Then Exit Sub
    If Len(dropColLetter) = 0 Or RowNum <= 0 Then Exit Sub

    ' Define target and trigger cells
    Set TrgCell = GetCellByLetter(wsTarget, sourceColLetter, RowNum)
    Set cell = wsTarget.Range(dropColLetter & RowNum)
    If cell Is Nothing Then Exit Sub

    ' Create unique tag ID
    tagId = "Col " & sourceColLetter

    Application.EnableEvents = False

    ' Always clear any existing tag for this column before proceeding
    Call ClearSystemTagFromString_KeepOthers(cell, tagId)

    ' Handle "Default" case – cleanup only, no new message
    If FormatType = "Default" Then
        If Not FormatMap Is Nothing Then
            setFormat TrgCell, "Default", FormatMap
        Else
            setFormat TrgCell, "Default", DefaultFormatMap
        End If
        cell.value = Trim(cell.value) ' remove blank lines if left
        Application.EnableEvents = True
        Exit Sub
    End If

    ' Build message and append
    fullMsg = SYSTEM_TAG_START & " " & tagId & ": " & tagText & " " & SYSTEM_TAG_END

    ' Ensure we have a format map
    If FormatMap Is Nothing Then
        Set FormatMap = DefaultFormatMap
        Debug.Print "Format Map was loaded from default settings"
    End If

    ' Apply formatting to the triggering cell
    setFormat TrgCell, FormatType, FormatMap

    ' Append the message cleanly
    existingText = Trim(Replace(cell.value, vbCr, ""))
    Do While Right(existingText, 1) = vbLf
        existingText = Left(existingText, Len(existingText) - 1)
    Loop

    If existingText <> "" Then
        cleanedText = existingText & vbLf & fullMsg
    Else
        cleanedText = fullMsg
    End If

    cell.value = cleanedText

CleanExit:
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    Debug.Print "WriteSystemTagToDropColumn ERROR: " & Err.Number & " - " & Err.Description
    Resume CleanExit
End Sub




