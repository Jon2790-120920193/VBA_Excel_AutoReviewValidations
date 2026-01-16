Attribute VB_Name = "ElecticityValidations"
' === PUBLIC ENTRY POINTS ===
Option Explicit


Public Sub Validate_Column_Electricity(cell As Range, sheetName As String, Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)
    Dim dependentCell As Range
    Set dependentCell = GetDependentCell(cell, sheetName)
    
    If FormatMap Is Nothing Then
        Set FormatMap = DefaultFormatMap()
    End If

    If Not dependentCell Is Nothing Then
        ValidatePairs cell, dependentCell, sheetName, english, FormatMap, AutoValMap
    End If
End Sub

Public Sub Validate_Column_Electricity_Metered(cell As Range, sheetName As String, Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)
    Dim dependentCell As Range
    Set dependentCell = GetDependentCell(cell, sheetName)
    
    If FormatMap Is Nothing Then
        Set FormatMap = DefaultFormatMap()
    End If

    If Not dependentCell Is Nothing Then
        ValidatePairs dependentCell, cell, sheetName, english, FormatMap, AutoValMap
    End If
End Sub

Private Function ValidatePairs(cellA As Range, cellB As Range, sheetName As String, Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object) As Boolean
    Static isRunning As Boolean
    If isRunning Then Exit Function
    isRunning = True

    Dim wsConfig As Worksheet
    Dim wsTargetSheet
    Dim valA As String, valB As String
    Dim validationTable As ListObject
    Dim tblRow As ListRow
    Dim ConfigValidationPairsTable As String
    Dim isMatchFound As Boolean
    Dim autoCorrectFlag As Boolean
    Dim Msg As String
    Dim correctedA As String, correctedB As String
    
    Dim DevFuncName1 As String
    DevFuncName1 = "Electricity"
    Dim DevFuncName2 As String
    DevFuncName2 = "Electricity_Metered"

    ' === CONFIGURATION SECTION ===
    ConfigValidationPairsTable = "ElectricityPairValidation"
    ' Table Columns:
    '   1 = Input A
    '   2 = Input B
    '   3 = AutoCorrect (TRUE/FALSE)
    '   4 = Corrected A (Optional)
    '   5 = Corrected B (Optional)
    ' ==============================

    Set wsConfig = ThisWorkbook.Sheets("Config")
    Set wsTargetSheet = ThisWorkbook.Sheets(sheetName)
    valA = Trim(CStr(cellA.value))
    valB = Trim(CStr(cellB.value))

    On Error Resume Next
    Set validationTable = wsConfig.ListObjects(ConfigValidationPairsTable)
    On Error GoTo 0

    If validationTable Is Nothing Then
        Debug.Print "Validation Table was not found"
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
                    "Values auto-corrected to match valid data pairs.", _
                    "Valeurs corrigées automatiquement pour correspondre à la configuration valide.")

                AddValidationFeedback DevFuncName1, wsTargetSheet, cellA.row, "", "Default", english, FormatMap, AutoValMap
                AddValidationFeedback DevFuncName2, wsTargetSheet, cellA.row, Msg, "Autocorrect", english, FormatMap, AutoValMap
                ValidatePairs = True
                GoTo CleanExit
            Else
                ValidatePairs = True
                    AddValidationFeedback DevFuncName1, wsTargetSheet, cellA.row, "", "Default", english, FormatMap, AutoValMap
                    AddValidationFeedback DevFuncName2, wsTargetSheet, cellA.row, "", "Default", english, FormatMap, AutoValMap
                GoTo CleanExit
            End If
        End If
    Next tblRow

    If Not isMatchFound Then
        Msg = IIf(english, _
            "Invalid combination of Electricity and Electricity Metered.", _
            "Combinaison invalide d'électricité et de mesure électrique.")

        AddValidationFeedback DevFuncName1, wsTargetSheet, cellA.row, Msg, "Error", english, FormatMap, AutoValMap
        AddValidationFeedback DevFuncName2, wsTargetSheet, cellA.row, "", "Error", english, FormatMap, AutoValMap
        
        ValidatePairs = False
        GoTo CleanExit
    End If
    
    AddValidationFeedback DevFuncName1, wsTargetSheet, cellA.row, "", "Default", english, FormatMap, AutoValMap
    AddValidationFeedback DevFuncName2, wsTargetSheet, cellA.row, "", "Default", english, FormatMap, AutoValMap

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
    ConfigFirstRow = 8
    ValueAName = "Electricity"
    ValueBName = "Electricity_Metered"
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






