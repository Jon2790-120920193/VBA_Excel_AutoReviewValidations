Attribute VB_Name = "HeatValidations"
Option Explicit

' =============================================
' HeatValidations Module (Revised with Recursion)
' =============================================

' === PUBLIC ENTRY POINTS ===
Public Sub Validate_Column_Heat_Source(cell As Range, sheetName As String, Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)
    Dim dependentCell As Range
    Set dependentCell = GetDependentCell(cell, sheetName)
    
    If FormatMap Is Nothing Then
        Set FormatMap = DefaultFormatMap()
    End If
    
    ' DebugMessage "[Heat Source Validation] Starting for " & cell.Address
    
    
    If dependentCell Is Nothing Then
        ' DebugMessage "[Heat Source Validation] Could not resolve dependent cell for " & cell.Address
        Exit Sub
    End If
    
    If Not dependentCell Is Nothing Then
        ValidatePairs cell, dependentCell, sheetName, english, 0, FormatMap, AutoValMap
        ' DebugMessage "[Heat Source Validation] Completed for " & cell.Address
    Else
        ' DebugMessage "[Heat Metered Validation] Could not resolve dependent cell for " & cell.Address
    End If
End Sub

Public Sub Validate_Column_Heat_Metered(cell As Range, sheetName As String, Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)
    Dim dependentCell As Range
    Set dependentCell = GetDependentCell(cell, sheetName)
    
    If FormatMap Is Nothing Then
        Set FormatMap = DefaultFormatMap()
    End If
    
    DebugMessage "[Heat Metered Validation] Starting for " & cell.Address
    

    If Not dependentCell Is Nothing Then
        Call ValidatePairs(dependentCell, cell, sheetName, english, 0, FormatMap, AutoValMap)
        DebugMessage "[Heat Metered Validation] Completed for " & cell.Address
    Else
        DebugMessage "[Heat Metered Validation] Could not resolve dependent cell for " & cell.Address
    End If
End Sub


' === INTERNAL VALIDATION LOGIC ===
Private Function ValidatePairs(cellA As Range, cellB As Range, sheetName As String, _
                               Optional english As Boolean = True, _
                               Optional recursionLevel As Integer = 0, Optional FormatMap As Object, Optional AutoValMap As Object) As Boolean
    Static isRunning As Boolean
    If isRunning Then Exit Function
    isRunning = True
    
    Dim wsConfig As Worksheet
    Dim valA As String, valB As String
    Dim validationTable As ListObject, anyTable As ListObject
    Dim tblRow As ListRow, tblRowAny As ListRow
    Dim isMatchFound As Boolean, autoCorrectFlag As Boolean
    Dim Msg As String, correctedA As String, correctedB As String
    Dim tableValA As String, tableValB As String, autoCorrectVal As Variant
    Dim anyInput As String, tempValA As String
    Dim wildcards As Variant, prefix As Variant
    Dim wsTargetSheet As Worksheet
    
    Set wsTargetSheet = ThisWorkbook.Worksheets(sheetName)
    
    Set wsConfig = ThisWorkbook.Sheets("Config")
    valA = Trim(CStr(cellA.value))
    valB = Trim(CStr(cellB.value))
    isMatchFound = False

    DebugMessage "[ValidatePairs] Starting: A=" & valA & " | B=" & valB & ", Sheet=" & sheetName

    Dim DevFuncName1 As String
    DevFuncName1 = "Heat_Source"
    Dim DevFuncName2 As String
    DevFuncName2 = "Heat_Metered"

    ' === STEP 1: Exact match in validation table ===
    On Error Resume Next
    Set validationTable = wsConfig.ListObjects("HeatSourcePairValidation")
    On Error GoTo 0

    If Not validationTable Is Nothing Then
        For Each tblRow In validationTable.ListRows
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
                    
                    AddValidationFeedback DevFuncName1, wsTargetSheet, cellA.row, "", "Default", english, FormatMap, AutoValMap
                    AddValidationFeedback DevFuncName2, wsTargetSheet, cellA.row, "", "Default", english, FormatMap, AutoValMap
                    
                    Msg = IIf(english, _
                        "Minor change performed, auto-corrected to match valid heat source/metered combinations.", _
                        "Correction mineure appliquée automatiquement pour correspondre aux combinaisons valides de source et compteur de chaleur.")
                        
                            If correctedA <> valA Then
                                AddValidationFeedback DevFuncName1, wsTargetSheet, cellA.row, Msg, "Autocorrect", english, FormatMap, AutoValMap
                                AddValidationFeedback DevFuncName2, wsTargetSheet, cellA.row, "", "Default", english, FormatMap, AutoValMap
                                ValidatePairs = True
                                GoTo CleanExit
                            End If
                            
                            If correctedB <> valB Then
                                AddValidationFeedback DevFuncName2, wsTargetSheet, cellA.row, Msg, "Autocorrect", english, FormatMap, AutoValMap
                                AddValidationFeedback DevFuncName1, wsTargetSheet, cellA.row, "", "Default", english, FormatMap, AutoValMap
                                ValidatePairs = True
                                GoTo CleanExit
                            End If
                End If
                AddValidationFeedback DevFuncName2, wsTargetSheet, cellA.row, "", "Default", english, FormatMap, AutoValMap
                AddValidationFeedback DevFuncName2, wsTargetSheet, cellA.row, "", "Default", english, FormatMap, AutoValMap
                ValidatePairs = True
                GoTo CleanExit
            End If
        Next tblRow
    End If


    ' === STEP 2: ANY / ANY(FR) mapping ===
    On Error Resume Next
    Set anyTable = wsConfig.ListObjects("HeatSourceANYRefTable")
    On Error GoTo 0

    If Not anyTable Is Nothing Then
        For Each tblRowAny In anyTable.ListRows
            anyInput = Trim(CStr(tblRowAny.Range(1, 1).value))
            If StrComp(anyInput, valA, vbTextCompare) = 0 Then
                tempValA = IIf(InStr(anyInput, "(FR)") > 0, "ANY(FR)", "ANY")
                
                For Each tblRow In validationTable.ListRows
                    tableValA = Trim(CStr(tblRow.Range(1, 1).value))
                    tableValB = Trim(CStr(tblRow.Range(1, 2).value))
                    autoCorrectVal = tblRow.Range(1, 3).value
                    
                    If StrComp(tableValA, tempValA, vbTextCompare) = 0 And _
                       StrComp(tableValB, valB, vbTextCompare) = 0 Then
                        
                        isMatchFound = True
                        autoCorrectFlag = (LCase(Trim(CStr(autoCorrectVal))) = "true")
                        
                        If autoCorrectFlag Then
                            correctedA = valA
                            correctedB = Trim(CStr(tblRow.Range(1, 5).value))
                            
                            Application.EnableEvents = False
                            If correctedA <> "" Then cellA.value = correctedA
                            If correctedB <> "" Then cellB.value = correctedB
                            Application.EnableEvents = True
                            
                            AddValidationFeedback DevFuncName1, wsTargetSheet, cellA.row, "", "Default", english, FormatMap, AutoValMap
                            AddValidationFeedback DevFuncName2, wsTargetSheet, cellA.row, "", "Default", english, FormatMap, AutoValMap
                            
                            Msg = IIf(english, _
                                "Auto-corrected to match valid heat source/metered combinations.", _
                                "Correction automatiquement pour correspondre aux combinaisons valides de source et compteur de chaleur.")
                                
                            Msg = Msg & " " & valA & " -> " & correctedA
                            If correctedA <> valA Then
                                AddValidationFeedback DevFuncName1, wsTargetSheet, cellA.row, Msg, "Autocorrect", english, FormatMap, AutoValMap
                                AddValidationFeedback DevFuncName2, wsTargetSheet, cellA.row, "", "Default", english, FormatMap, AutoValMap
                                ValidatePairs = True
                                GoTo CleanExit
                            End If
                            
                            
                            Msg = Msg & " " & valB & " -> " & correctedB
                            If correctedB <> valB Then
                                AddValidationFeedback DevFuncName2, wsTargetSheet, cellA.row, Msg, "Autocorrect", english, FormatMap, AutoValMap
                                AddValidationFeedback DevFuncName1, wsTargetSheet, cellA.row, "", "Default", english, FormatMap, AutoValMap
                                ValidatePairs = True
                                GoTo CleanExit
                            End If
                        End If
                        
                        AddValidationFeedback DevFuncName1, wsTargetSheet, cellA.row, "", "Default", english, FormatMap, AutoValMap
                        AddValidationFeedback DevFuncName2, wsTargetSheet, cellA.row, "", "Default", english, FormatMap, AutoValMap
                        ValidatePairs = True
                        GoTo CleanExit
                    End If
                Next tblRow
            End If
        Next tblRowAny
    End If


    ' === STEP 3: Wildcards / Recursive Normalization ===
    wildcards = Array("Installation de chauffage centrale", "Central Heating Plant")

    For Each prefix In wildcards
        Dim basePrefix As String
        basePrefix = prefix
        
        If LCase(Left(valA, Len(basePrefix))) = LCase(basePrefix) Then
            Dim remainder As String
            remainder = Trim(Mid(valA, Len(basePrefix) + 1))
            
            ' Strip redundant punctuation
            Do While Len(remainder) > 0
                Dim firstChar As String
                firstChar = Left(remainder, 1)
                If firstChar = " " Or firstChar = "-" Or firstChar = "–" Or firstChar = "," Then
                    remainder = Mid(remainder, 2)
                Else
                    Exit Do
                End If
            Loop
            
            ' === CASE 1: Missing subtype ===
            If remainder = "" Then
                Msg = IIf(english, _
                    "Central Heating Plant entries must specify a Heat Source after the dash (e.g., 'Central Heating Plant - Natural Gas').", _
                    "Les entrées d'Installation de chauffage centrale doivent préciser une source de chaleur après le tiret (ex. 'Installation de chauffage centrale - Gaz naturel').")
                AddValidationFeedback DevFuncName1, wsTargetSheet, cellA.row, Msg, "Error", english, FormatMap, AutoValMap
                ValidatePairs = True
                GoTo CleanExit
            End If
            
            ' === CASE 2: Normalize format ===
            Dim normalizedA As String
            normalizedA = basePrefix & " - " & remainder
            
            If StrComp(normalizedA, valA, vbTextCompare) <> 0 Then
                Application.EnableEvents = False
                cellA.value = normalizedA
                Application.EnableEvents = True
                
                Msg = IIf(english, _
                    "Central Heating Plant entry. Corrected the delimiter format/spacing.", _
                    "Correction: Installation de chauffage centrale. Nettoyage des tirets et espaces.")
                AddValidationFeedback DevFuncName1, wsTargetSheet, cellA.row, Msg, "Autocorrect", english, FormatMap, AutoValMap
                AddValidationFeedback DevFuncName2, wsTargetSheet, cellA.row, "", "Default", english, FormatMap, AutoValMap
                
                ' === RECURSION: Re-run validation once after normalization ===
                If recursionLevel < 1 Then
                    isRunning = False
                    Call ValidatePairs(cellA, cellB, sheetName, english, recursionLevel + 1, FormatMap, AutoValMap)
                    Exit Function
                End If
            End If
            
            ' === Optional Auto-Correction for Metered ===
            If Trim(valB) = "#" Then
                correctedB = "No"
                Application.EnableEvents = False
                cellB.value = correctedB
                Application.EnableEvents = True
                Msg = IIf(english, _
                    "Auto-corrected Heat Metered to 'No' for Central Heating Plant entry.", _
                    "Compteur de chaleur corrigé automatiquement en 'Non' pour l'entrée Installation de chauffage centrale.")
                AddValidationFeedback DevFuncName2, wsTargetSheet, cellA.row, Msg, "Autocorrect", english, FormatMap, AutoValMap
                AddValidationFeedback DevFuncName1, wsTargetSheet, cellA.row, "", "Default", english, FormatMap, AutoValMap
                ValidatePairs = True
                GoTo CleanExit
            End If
            
            AddValidationFeedback DevFuncName1, wsTargetSheet, cellA.row, "", "Default", english, FormatMap, AutoValMap
            AddValidationFeedback DevFuncName2, wsTargetSheet, cellA.row, "", "Default", english, FormatMap, AutoValMap
             ValidatePairs = True
            GoTo CleanExit
        End If
    Next prefix


    ' === STEP 4: No valid match ===
    If Not isMatchFound Then
        Msg = IIf(english, _
            "Invalid combination of Heat Source and Heat Metered.", _
            "Combinaison invalide de la source de chaleur et du mesurage de la chaleur.")
        AddValidationFeedback DevFuncName1, wsTargetSheet, cellA.row, Msg, "Error", english, FormatMap, AutoValMap
        AddValidationFeedback DevFuncName2, wsTargetSheet, cellA.row, "", "Error", english, FormatMap, AutoValMap
        ValidatePairs = False
        GoTo CleanExit
    End If
    
    AddValidationFeedback DevFuncName1, wsTargetSheet, cellA.row, "", "Default", english, FormatMap, AutoValMap
    AddValidationFeedback DevFuncName2, wsTargetSheet, cellA.row, "", "Default", english, FormatMap, AutoValMap
    ValidatePairs = True

CleanExit:
    DebugMessage "[ValidatePairs] Exiting: A=" & cellA.value & " | B=" & cellB.value & " | MatchFound=" & isMatchFound
    isRunning = False
End Function


' =============================================
' DEPENDENCY RESOLVER
' =============================================
Private Function GetDependentCell(cell As Range, sheetName As String) As Range
    Dim ws As Worksheet, wsConfig As Worksheet
    Dim RowNum As Long
    Dim FirstColumn As String, SecondColumn As String
    Dim i As Long, maxRow As Long
    Dim ValueAName As String, ValueBName As String
    Dim ConfigColLet_Values As String, ConfigColLet_FunctionNames As String
    
    ConfigColLet_Values = "B"
    ConfigColLet_FunctionNames = "C"
    ValueAName = "Heat_Source"
    ValueBName = "Heat_Metered"
    
    Set ws = ThisWorkbook.Sheets(sheetName)
    Set wsConfig = ThisWorkbook.Sheets("Config")
    RowNum = cell.row
    maxRow = wsConfig.Cells(wsConfig.Rows.count, ConfigColLet_FunctionNames).End(xlUp).row
    
    For i = 12 To maxRow
        Dim currentFunc As String, currentCol As String
        currentFunc = UCase(Trim(wsConfig.Range(ConfigColLet_FunctionNames & i).value))
        currentCol = Trim(wsConfig.Range(ConfigColLet_Values & i).value)
        
        If currentFunc = UCase(ValueAName) Then FirstColumn = currentCol
        If currentFunc = UCase(ValueBName) Then SecondColumn = currentCol
    Next i
    
    DebugMessage "[GetDependentCell] Row: " & RowNum & " | FirstColumn=" & FirstColumn & " | SecondColumn=" & SecondColumn
    
    If FirstColumn = "" Or SecondColumn = "" Then
        DebugMessage "!! Could not resolve Heat columns."
        Exit Function
    End If
    
    On Error GoTo RangeError
    If cell.Column = ws.Range(FirstColumn & "1").Column Then
        Set GetDependentCell = ws.Range(SecondColumn & RowNum)
    Else
        Set GetDependentCell = ws.Range(FirstColumn & RowNum)
    End If
    On Error GoTo 0
    Exit Function

RangeError:
    DebugMessage "!! Invalid range for cell: " & cell.Address
    Set GetDependentCell = Nothing
End Function


