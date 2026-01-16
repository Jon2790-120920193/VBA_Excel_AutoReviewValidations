Attribute VB_Name = "PlumbingValidations"
Option Explicit

Private Const MODULE_NAME As String = "PlumbingValidations"
Private Const DevFuncName1 As String = "Plumbing"
Private Const DevFuncName2 As String = "Water_Metered"

' === PUBLIC ENTRY POINTS ===

Public Sub Validate_Column_Plumbing(cell As Range, sheetName As String, Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)
    Dim dependentCell As Range
    Set dependentCell = GetDependentCell(cell, sheetName)
    
    If FormatMap Is Nothing Then
        Set FormatMap = DefaultFormatMap()
    End If

    If Not dependentCell Is Nothing Then
        ValidatePairs cell, dependentCell, sheetName, DevFuncName1, english, FormatMap, AutoValMap
    End If
End Sub

Public Sub Validate_Column_Water_Metered(cell As Range, sheetName As String, Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)
    Dim dependentCell As Range
    Set dependentCell = GetDependentCell(cell, sheetName)
    
    If FormatMap Is Nothing Then
        Set FormatMap = DefaultFormatMap()
    End If

    If Not dependentCell Is Nothing Then
        ValidatePairs dependentCell, cell, sheetName, DevFuncName2, english, FormatMap, AutoValMap
    End If
End Sub

Private Function ValidatePairs(cellA As Range, cellB As Range, sheetName As String, CalledFuncName As String, _
                               Optional english As Boolean = True, _
                            Optional FormatMap As Object, Optional AutoValMap As Object) As Boolean
    Static isRunning As Boolean
    If isRunning Then Exit Function
    isRunning = True

    Dim wsConfig As Worksheet
    Dim valA As String, valB As String
    Dim validationTable As ListObject
    Dim tblRow As ListRow
    Dim ConfigValidationPairsTable As String
    Dim isMatchFound As Boolean
    Dim autoCorrectFlag As Boolean
    Dim Msg As String
    Dim correctedA As String, correctedB As String
    Dim OtherFuncName As String
    
    If CalledFuncName = DevFuncName1 Then OtherFuncName = DevFuncName2 Else OtherFuncName = DevFuncName1
    
    Dim wsTargetSheet As Worksheet
    Set wsTargetSheet = ThisWorkbook.Worksheets(sheetName)

    ' === CONFIGURATION SECTION ===
    ConfigValidationPairsTable = "PlumbingPairValidation"
    ' Table Columns:
    '   1 = Input A
    '   2 = Input B
    '   3 = AutoCorrect (TRUE/FALSE)
    '   4 = Corrected A (Optional)
    '   5 = Corrected B (Optional)
    ' ==============================

    Set wsConfig = ThisWorkbook.Sheets("Config")
    valA = Trim(CStr(cellA.value))
    valB = Trim(CStr(cellB.value))

    On Error Resume Next
    Set validationTable = wsConfig.ListObjects(ConfigValidationPairsTable)
    On Error GoTo 0

    If validationTable Is Nothing Then
        DebugMessage cellA.value, "Validation table '" & ConfigValidationPairsTable & "' not found in Config sheet."
        ValidatePairs = False
        GoTo CleanExit
    End If

    isMatchFound = False

    For Each tblRow In validationTable.ListRows
        Dim tableValA As String, tableValB As String
        Dim autoCorrectVal As Variant

        tableValA = Trim(CStr(tblRow.Range(1, 1).value))
        tableValB = Trim(CStr(tblRow.Range(1, 2).value))
        autoCorrectVal = tblRow.Range(1, 3).value

        If StrComp(tableValA, valA, vbTextCompare) = 0 And _
           StrComp(tableValB, valB, vbTextCompare) = 0 Then

            isMatchFound = True
            autoCorrectFlag = (LCase(Trim(CStr(autoCorrectVal))) = "true")

            If autoCorrectFlag Then
                correctedA = Trim(CStr(tblRow.Range(1, 4).value))
                correctedB = Trim(CStr(tblRow.Range(1, 5).value))

                Application.EnableEvents = False
                If correctedA <> "" Then cellA.value = correctedA
                If correctedB <> "" Then cellB.value = correctedB
                Application.EnableEvents = True

                Msg = IIf(english, _
                    "AutoCorrected", _
                    "Valeurs corrigées automatiquement pour correspondre à la configuration valide.")
                
                If correctedA <> valA Then
                    Msg = Msg & " " & valA & " -> " & correctedA
                    AddValidationFeedback OtherFuncName, wsTargetSheet, cellA.row, Msg, "Autocorrect", english, FormatMap, AutoValMap
                    AddValidationFeedback CalledFuncName, wsTargetSheet, cellA.row, "", "Default", english, FormatMap, AutoValMap
                    ValidatePairs = True
                    GoTo CleanExit
                Else
                    Msg = Msg & " " & valB & " -> " & correctedB
                    AddValidationFeedback CalledFuncName, wsTargetSheet, cellA.row, Msg, "Autocorrect", english, FormatMap, AutoValMap
                    AddValidationFeedback OtherFuncName, wsTargetSheet, cellA.row, "", "Default", english, FormatMap, AutoValMap
                    ValidatePairs = True
                    GoTo CleanExit
                End If
            Else
                AddValidationFeedback CalledFuncName, wsTargetSheet, cellA.row, "", "Default", english, FormatMap, AutoValMap
                AddValidationFeedback OtherFuncName, wsTargetSheet, cellA.row, "", "Default", english, FormatMap, AutoValMap
                ValidatePairs = True
                GoTo CleanExit
            End If
        End If
    Next tblRow

    If Not isMatchFound Then
        Msg = IIf(english, _
            "Invalid combination of Plumbing and Water Metered.", _
            "Combinaison invalide de plomberie et de mesure d'eau.")

        AddValidationFeedback CalledFuncName, wsTargetSheet, cellA.row, Msg, "Error", english, FormatMap, AutoValMap
        AddValidationFeedback OtherFuncName, wsTargetSheet, cellA.row, Msg, "Error", english, FormatMap, AutoValMap
        ValidatePairs = False
    End If

CleanExit:
    isRunning = False
End Function

' === DEPENDENCY RESOLVER ===
Private Function GetDependentCell(cell As Range, sheetName As String) As Range
    Dim ws As Worksheet, wsConfig As Worksheet
    Dim RowNum As Long
    Dim FirstColumn As String, SecondColumn As String
    Dim i As Long
    Dim ValueAName As String
    Dim ValueBName As String
    Dim ConfigColLet_Values As String
    Dim ConfigColLet_FunctionNames As String
    Dim ConfigFirstRow As Integer

    ' === CONFIGURATION SECTION ===
    ConfigColLet_Values = "B"
    ConfigColLet_FunctionNames = "C"
    ConfigFirstRow = 10
    ValueAName = "Plumbing"
    ValueBName = "Water_Metered"
    ' === END CONFIGURATION ===

    Set ws = ThisWorkbook.Sheets(sheetName)
    Set wsConfig = ThisWorkbook.Sheets("Config")
    RowNum = cell.row

    FirstColumn = "": SecondColumn = ""
    i = ConfigFirstRow

    Do While wsConfig.Range(ConfigColLet_Values & i).value <> ""
        Select Case Trim(wsConfig.Range(ConfigColLet_FunctionNames & i).value)
            Case ValueAName: FirstColumn = Trim(wsConfig.Range(ConfigColLet_Values & i).value)
            Case ValueBName: SecondColumn = Trim(wsConfig.Range(ConfigColLet_Values & i).value)
        End Select
        i = i + 1
    Loop

    ' === Debug check for missing mappings ===
    If FirstColumn = "" Or SecondColumn = "" Then
        Debug.Print "?? [GetDependentCell] Configuration error:"
        If FirstColumn = "" Then Debug.Print "  - Missing or incorrect config entry for '" & ValueAName & "'. Expected in column '" & ConfigColLet_Values & ConfigFirstRow + 1 & "'"
        If SecondColumn = "" Then Debug.Print "  - Missing or incorrect config entry for '" & ValueBName & "'. Expected in column '" & ConfigColLet_FunctionNames & (ConfigFirstRow + 1) & "'"
        Debug.Print "Please check the 'Config' sheet, starting at row " & ConfigFirstRow & "."
        ' Optional: Show a pop-up
        ' MsgBox "Configuration error in 'Config' sheet. Check Debug window for details.", vbExclamation
        Exit Function
    End If

    ' === Return dependent cell ===
    If cell.Column = ws.Range(FirstColumn & "1").Column Then
        Set GetDependentCell = ws.Range(SecondColumn & RowNum)
    Else
        Set GetDependentCell = ws.Range(FirstColumn & RowNum)
    End If
End Function




