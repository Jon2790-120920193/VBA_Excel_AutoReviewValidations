Attribute VB_Name = "AV_ValidationRules"
Option Explicit

' ======================================================
' AV_ValidationRules.bas v2.1
' All validation business logic and rules
' Called by AV_Validators (thin routing layer)
' UPDATED: Uses AV_Core.GetValidationTable() for cached access
'          Uses AV_Constants for all magic numbers
' ======================================================

Private Const MODULE_NAME As String = "AV_ValidationRules"

' ======================================================
' GENERIC PAIR VALIDATION (Electricity / Plumbing)
' ======================================================
Public Sub ValidatePairedFields( _
    cell As Range, sheetName As String, _
    ThisFunc As String, OtherFunc As String, _
    DefaultRuleTable As String, _
    english As Boolean, FormatMap As Object, autoValMap As Object)

    Static isRunning As Boolean
    If isRunning Then Exit Sub
    isRunning = True

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    Dim otherCell As Range
    Set otherCell = AV_Validators.GetSiblingCell(cell, sheetName, OtherFunc, autoValMap)

    If otherCell Is Nothing Then GoTo ExitSafe

    Dim ruleTable As String
    ruleTable = AV_Core.GetRuleTableNameFromAutoValMap(autoValMap, ThisFunc, DefaultRuleTable)

    Dim result As Boolean
    result = RunPairRuleValidation( _
        ws, cell, otherCell, ThisFunc, OtherFunc, _
        ruleTable, english, FormatMap, autoValMap)

ExitSafe:
    isRunning = False
End Sub

' ======================================================
' PAIR RULE ENGINE (TABLE-DRIVEN VALIDATION) - Updated v2.1
' ======================================================
Public Function RunPairRuleValidation( _
    ws As Worksheet, cellA As Range, cellB As Range, _
    FuncA As String, FuncB As String, _
    RuleTableName As String, _
    english As Boolean, FormatMap As Object, autoValMap As Object _
) As Boolean

    ' Use cached table access
    Dim lo As ListObject
    Set lo = AV_Core.GetValidationTable(RuleTableName)

    If lo Is Nothing Then
        AV_Core.DebugMessage "Rule table missing: " & RuleTableName, MODULE_NAME
        RunPairRuleValidation = False
        Exit Function
    End If

    Dim valA As String, valB As String
    valA = AV_Core.SafeTrim(cellA.Value)
    valB = AV_Core.SafeTrim(cellB.Value)

    Dim r As ListRow
    For Each r In lo.ListRows
        If StrComp(valA, AV_Core.SafeTrim(r.Range(1, 1).Value), vbTextCompare) = 0 _
           And StrComp(valB, AV_Core.SafeTrim(r.Range(1, 2).Value), vbTextCompare) = 0 Then

            AV_Format.AddValidationFeedback FuncA, ws, cellA.Row, "", AV_Constants.FORMAT_DEFAULT, english, FormatMap, autoValMap
            AV_Format.AddValidationFeedback FuncB, ws, cellA.Row, "", AV_Constants.FORMAT_DEFAULT, english, FormatMap, autoValMap
            RunPairRuleValidation = True
            Exit Function
        End If
    Next r

    ' No match found - error
    AV_Format.AddValidationFeedback FuncA, ws, cellA.Row, _
        IIf(english, "Invalid value pairing.", "Combinaison invalide."), _
        AV_Constants.FORMAT_ERROR, english, FormatMap, autoValMap

    AV_Format.AddValidationFeedback FuncB, ws, cellA.Row, "", AV_Constants.FORMAT_ERROR, english, FormatMap, autoValMap
    RunPairRuleValidation = False
End Function

' ======================================================
' GIW QUANTITY VALIDATION - Updated v2.1
' Complex logic: handles #,# placeholders, auto-corrections, range validation
' ======================================================
Public Function Validate_GIWQuantity( _
    cell As Range, sheetName As String, _
    funcName As String, _
    english As Boolean, _
    FormatMap As Object, autoValMap As Object _
) As Boolean

    Dim wsTargetSheet As Worksheet
    Set wsTargetSheet = ThisWorkbook.Sheets(sheetName)

    Dim val As String, parts() As String
    Dim num1 As Long, num2 As Long
    Dim Msg As String
    Dim OtherFuncName As String

    If StrComp(funcName, "GIWQuantity", vbTextCompare) = 0 Then
        OtherFuncName = "GIWIncluded"
    Else
        OtherFuncName = "GIWQuantity"
    End If

    val = AV_Core.SafeTrim(cell.Value)
    val = Replace(val, ".", ",")
    val = Replace(val, " ", "")

    ' Normalize value in-cell if changed
    If val <> AV_Core.SafeTrim(cell.Value) Then
        Application.EnableEvents = False
        cell.Value = val
        Application.EnableEvents = True
    End If

    ' Handle (A,B) form - remove parentheses
    If Left$(val, 1) = "(" And Right$(val, 1) = ")" Then
        val = Mid$(val, 2, Len(val) - 2)
        Application.EnableEvents = False
        cell.Value = val
        Application.EnableEvents = True
    End If

    ' Empty check
    If val = "" Then
        Msg = IIf(english, "Cannot be empty", "Ne peut pas être vide.")
        AV_Format.AddValidationFeedback funcName, wsTargetSheet, cell.Row, Msg, AV_Constants.FORMAT_ERROR, english, FormatMap, autoValMap
        Validate_GIWQuantity = False
        Exit Function
    End If

    ' Special placeholder #
    If val = "#" Then
        Application.EnableEvents = False
        cell.Value = "#,#"
        Application.EnableEvents = True

        Msg = IIf(english, "Auto-corrected placeholder", "Correction automatique")
        AV_Format.AddValidationFeedback funcName, wsTargetSheet, cell.Row, Msg, AV_Constants.FORMAT_ERROR, english, FormatMap, autoValMap
        Validate_GIWQuantity = True
        Exit Function
    End If

    ' Single numeric entry -> normalize to n,n
    If IsNumeric(val) And InStr(1, val, ",", vbTextCompare) = 0 Then

        If CLng(val) > AV_Constants.MAX_GIW_VALUE Then
            Msg = IIf(english, _
                      "Max value: " & AV_Constants.MAX_GIW_VALUE & " surpassed", _
                      "Valeur maximale : " & AV_Constants.MAX_GIW_VALUE & " dépassée")
            AV_Format.AddValidationFeedback funcName, wsTargetSheet, cell.Row, Msg, AV_Constants.FORMAT_ERROR, english, FormatMap, autoValMap
            Validate_GIWQuantity = False
            Exit Function
        End If

        num1 = CLng(val)

        Application.EnableEvents = False
        cell.Value = CStr(num1) & "," & CStr(num1)
        Application.EnableEvents = True

        Msg = IIf(english, _
                  "Format has been automatically corrected by the system", _
                  "Le format a été automatiquement corrigé par le système.")
        AV_Format.AddValidationFeedback funcName, wsTargetSheet, cell.Row, Msg, AV_Constants.FORMAT_AUTOCORRECT, english, FormatMap, autoValMap
        AV_Format.AddValidationFeedback OtherFuncName, wsTargetSheet, cell.Row, vbNullString, AV_Constants.FORMAT_DEFAULT, english, FormatMap, autoValMap

        Validate_GIWQuantity = True
        Exit Function
    End If

    ' Must be "A,B" format
    parts = Split(val, ",")
    If UBound(parts) <> 1 Then
        Msg = IIf(english, _
                  "Entry not valid, must be 'Number,Number'", _
                  "Entrée non valide, le format doit être 'Nombre,Nombre'")
        AV_Format.AddValidationFeedback funcName, wsTargetSheet, cell.Row, Msg, AV_Constants.FORMAT_ERROR, english, FormatMap, autoValMap
        Validate_GIWQuantity = False
        Exit Function
    End If

    ' Parse values (allow #)
    On Error GoTo ParseFail
    If parts(0) <> "#" Then num1 = CLng(parts(0))
    If parts(1) <> "#" Then num2 = CLng(parts(1))
    On Error GoTo 0

    AV_Format.AddValidationFeedback funcName, wsTargetSheet, cell.Row, vbNullString, AV_Constants.FORMAT_DEFAULT, english, FormatMap, autoValMap
    Validate_GIWQuantity = True
    Exit Function

ParseFail:
    Msg = IIf(english, _
              "Entry not valid, must be numeric values or '#'", _
              "Entrée non valide : valeurs numériques requises ou '#'.")
    AV_Format.AddValidationFeedback funcName, wsTargetSheet, cell.Row, Msg, AV_Constants.FORMAT_ERROR, english, FormatMap, autoValMap
    Validate_GIWQuantity = False
End Function

' ======================================================
' GIW INCLUDED VALIDATION - Updated v2.1
' Complex logic: validates against quantity based on inclusion rules
' Includes auto-correction for specific cases (#,# -> 0,0)
' ======================================================
Public Function Validate_GIWIncluded( _
    cell As Range, sheetName As String, _
    funcName As String, _
    english As Boolean, _
    FormatMap As Object, autoValMap As Object _
) As Boolean

    Static isRunning As Boolean
    If isRunning Then
        Validate_GIWIncluded = True
        Exit Function
    End If
    isRunning = True

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    Dim GIWIncludedVal As String
    GIWIncludedVal = AV_Core.SafeTrim(cell.Value)

    Dim quantityCell As Range
    Set quantityCell = AV_Validators.GetSiblingCell(cell, sheetName, "GIWQuantity", autoValMap)

    Dim GIWQuantityVal As String
    If Not quantityCell Is Nothing Then GIWQuantityVal = AV_Core.SafeTrim(quantityCell.Value)

    ' Get validation rule table using cached access
    Dim ruleTable As String
    ruleTable = AV_Core.GetRuleTableNameFromAutoValMap(autoValMap, funcName, AV_Constants.TBL_GIW_VALIDATION)

    Dim validationTable As ListObject
    Set validationTable = AV_Core.GetValidationTable(ruleTable)

    If validationTable Is Nothing Then
        AV_Core.DebugMessage "Validation table not found for GIW: " & ruleTable, MODULE_NAME
        AV_Format.AddValidationFeedback funcName, ws, cell.Row, _
            IIf(english, "Configuration error: GIW rule table missing.", "Erreur de configuration : table de règles GIW manquante."), _
            AV_Constants.FORMAT_ERROR, english, FormatMap, autoValMap
        Validate_GIWIncluded = False
        GoTo CleanExit
    End If

    ' Find expected rule for the Included value
    Dim expectedRule As String
    expectedRule = vbNullString

    Dim tblRow As ListRow
    For Each tblRow In validationTable.ListRows
        If StrComp(AV_Core.SafeTrim(tblRow.Range(1, 1).Value), GIWIncludedVal, vbTextCompare) = 0 Then
            expectedRule = AV_Core.SafeTrim(tblRow.Range(1, 2).Value)
            Exit For
        End If
    Next tblRow

    If expectedRule = vbNullString Then
        AV_Format.AddValidationFeedback funcName, ws, cell.Row, _
            IIf(english, "Error: Invalid entry", "Erreur : entrée non valide."), _
            AV_Constants.FORMAT_ERROR, english, FormatMap, autoValMap
        Validate_GIWIncluded = False
        GoTo CleanExit
    End If

    ' Validate quantity value
    If GIWQuantityVal = vbNullString Then
        AV_Format.AddValidationFeedback funcName, ws, cell.Row, _
            IIf(english, "Error: Cannot be empty", "Erreur : ne peut pas être vide."), _
            AV_Constants.FORMAT_ERROR, english, FormatMap, autoValMap
        Validate_GIWIncluded = False
        GoTo CleanExit
    End If

    Dim quantityParts() As String
    quantityParts = Split(GIWQuantityVal, ",")

    If UBound(quantityParts) <> 1 Then
        AV_Format.AddValidationFeedback funcName, ws, cell.Row, _
            IIf(english, "Entry not valid, must be 'Number,Number'", "Entrée non valide, le format doit être 'Nombre,Nombre'"), _
            AV_Constants.FORMAT_ERROR, english, FormatMap, autoValMap
        Validate_GIWIncluded = False
        GoTo CleanExit
    End If

    Dim num1 As Long, num2 As Long
    num1 = -1: num2 = -1

    On Error Resume Next
    If quantityParts(0) <> "#" Then num1 = CLng(quantityParts(0))
    If quantityParts(1) <> "#" Then num2 = CLng(quantityParts(1))
    On Error GoTo 0

    Dim isValid As Boolean
    isValid = False

    ' Validate against expected rule
    Select Case expectedRule
        Case "0"
            If num1 = 0 And num2 = 0 Then isValid = True
        Case "1"
            If num1 > 0 And num2 > 0 And num1 <= num2 Then isValid = True
        Case "#"
            If quantityParts(0) = "#" And quantityParts(1) = "#" Then isValid = True
    End Select

    Dim Msg As String
    If Not isValid Then
        Select Case expectedRule
            Case "0"
                ' Special autocorrect: #,# -> 0,0
                If num1 = -1 And num2 = -1 And quantityParts(0) = "#" And quantityParts(1) = "#" Then
                    Msg = IIf(english, _
                        "Automatic Correction: Changed entry #,# to 0,0", _
                        "Correction automatique : #,# à 0,0")

                    Application.EnableEvents = False
                    quantityCell.Value = "0,0"
                    Application.EnableEvents = True

                    AV_Format.AddValidationFeedback "GIWQuantity", ws, cell.Row, Msg, AV_Constants.FORMAT_AUTOCORRECT, english, FormatMap, autoValMap
                    AV_Format.AddValidationFeedback "GIWIncluded", ws, cell.Row, vbNullString, AV_Constants.FORMAT_DEFAULT, english, FormatMap, autoValMap

                    Validate_GIWIncluded = True
                    GoTo CleanExit
                Else
                    Msg = IIf(english, _
                        "Invalid Entry, value must be 0,0 when GIW Included = 'No'", _
                        "Combinaison invalide : la valeur doit être 0,0 lorsque GIW inclus = 'Non'")
                    AV_Format.AddValidationFeedback "GIWQuantity", ws, cell.Row, Msg, AV_Constants.FORMAT_ERROR, english, FormatMap, autoValMap
                    AV_Format.AddValidationFeedback "GIWIncluded", ws, cell.Row, vbNullString, AV_Constants.FORMAT_ERROR, english, FormatMap, autoValMap
                    Validate_GIWIncluded = False
                    GoTo CleanExit
                End If

            Case "1"
                If num1 > num2 And num1 <> -1 And num2 <> -1 Then
                    Msg = IIf(english, _
                        "'" & num1 & "," & num2 & "' is an invalid entry; GIW (" & num1 & _
                        ") cannot exceed Water Closets (" & num2 & ").", _
                        "Entrée invalide : le nombre de GIW (" & num1 & _
                        ") ne peut excéder le nombre de cabinets de toilette (" & num2 & ").")
                    AV_Format.AddValidationFeedback "GIWQuantity", ws, cell.Row, Msg, AV_Constants.FORMAT_ERROR, english, FormatMap, autoValMap
                    AV_Format.AddValidationFeedback "GIWIncluded", ws, cell.Row, vbNullString, AV_Constants.FORMAT_ERROR, english, FormatMap, autoValMap
                    Validate_GIWIncluded = False
                    GoTo CleanExit
                End If

                Msg = IIf(english, _
                    "GIW Quantity must be positive when GIW Included is 'Yes' or 'Partially'.", _
                    "La quantité GIW doit être positive lorsque GIW inclus est 'Oui' ou 'Partiellement'.")
                AV_Format.AddValidationFeedback "GIWQuantity", ws, cell.Row, Msg, AV_Constants.FORMAT_ERROR, english, FormatMap, autoValMap
                AV_Format.AddValidationFeedback "GIWIncluded", ws, cell.Row, vbNullString, AV_Constants.FORMAT_ERROR, english, FormatMap, autoValMap
                Validate_GIWIncluded = False
                GoTo CleanExit

            Case "#"
                Msg = IIf(english, _
                    "GIW Quantity must be '#,#' when GIW Included is 'Not Applicable'.", _
                    "La quantité GIW doit être '#,#' lorsque GIW inclus est 'Non applicable'.")
                AV_Format.AddValidationFeedback "GIWQuantity", ws, cell.Row, Msg, AV_Constants.FORMAT_ERROR, english, FormatMap, autoValMap
                AV_Format.AddValidationFeedback "GIWIncluded", ws, cell.Row, vbNullString, AV_Constants.FORMAT_ERROR, english, FormatMap, autoValMap
                Validate_GIWIncluded = False
                GoTo CleanExit

            Case Else
                Msg = IIf(english, "Invalid combination.", "Combinaison invalide.")
                AV_Format.AddValidationFeedback "GIWQuantity", ws, cell.Row, Msg, AV_Constants.FORMAT_ERROR, english, FormatMap, autoValMap
                AV_Format.AddValidationFeedback "GIWIncluded", ws, cell.Row, vbNullString, AV_Constants.FORMAT_ERROR, english, FormatMap, autoValMap
                Validate_GIWIncluded = False
                GoTo CleanExit
        End Select
    Else
        ' Valid combination: reset both to default
        AV_Format.AddValidationFeedback "GIWIncluded", ws, cell.Row, vbNullString, AV_Constants.FORMAT_DEFAULT, english, FormatMap, autoValMap
        AV_Format.AddValidationFeedback "GIWQuantity", ws, cell.Row, vbNullString, AV_Constants.FORMAT_DEFAULT, english, FormatMap, autoValMap
        Validate_GIWIncluded = True
    End If

CleanExit:
    isRunning = False
End Function

' ======================================================
' HEAT VALIDATION (MULTI-STAGE) - Updated v2.1
' Complex logic: Exact match -> ANY mapping -> Wildcard normalization -> Invalid
' Includes recursion for normalized values
' ======================================================
Public Sub Validate_HeatPairs( _
    cell As Range, sheetName As String, _
    funcName As String, _
    english As Boolean, _
    recursionLevel As Long, _
    FormatMap As Object, autoValMap As Object _
)

    Static isRunning As Boolean
    If isRunning Then Exit Sub
    isRunning = True

    Dim wsTarget As Worksheet
    Set wsTarget = ThisWorkbook.Sheets(sheetName)

    Dim cellA As Range, cellB As Range
    If funcName = "Heat_Source" Then
        Set cellA = cell
        Set cellB = AV_Validators.GetSiblingCell(cell, sheetName, "Heat_Metered", autoValMap)
    Else
        Set cellB = cell
        Set cellA = AV_Validators.GetSiblingCell(cell, sheetName, "Heat_Source", autoValMap)
    End If

    If cellA Is Nothing Or cellB Is Nothing Then GoTo CleanExit

    Dim valA As String, valB As String
    valA = AV_Core.SafeTrim(cellA.Value)
    valB = AV_Core.SafeTrim(cellB.Value)

    ' Get validation tables using cached access
    Dim ruleTable As String
    ruleTable = AV_Core.GetRuleTableNameFromAutoValMap(autoValMap, "Heat_Source", AV_Constants.TBL_HEAT_SOURCE_PAIRS)

    Dim validationTable As ListObject
    Set validationTable = AV_Core.GetValidationTable(ruleTable)

    Dim Msg As String
    Dim correctedA As String, correctedB As String
    Dim autoCorrect As Boolean
    Dim isMatchFound As Boolean
    isMatchFound = False

    ' ======================================================
    ' STAGE 1 – EXACT MATCH
    ' ======================================================
    If Not validationTable Is Nothing Then
        Dim r As ListRow
        For Each r In validationTable.ListRows
            If StrComp(valA, AV_Core.SafeTrim(r.Range(1, 1).Value), vbTextCompare) = 0 _
               And StrComp(valB, AV_Core.SafeTrim(r.Range(1, 2).Value), vbTextCompare) = 0 Then

                isMatchFound = True
                autoCorrect = (LCase(AV_Core.SafeTrim(r.Range(1, 3).Value)) = "true")

                If autoCorrect Then
                    correctedA = AV_Core.SafeTrim(r.Range(1, 4).Value)
                    correctedB = AV_Core.SafeTrim(r.Range(1, 5).Value)

                    Application.EnableEvents = False
                    If correctedA <> "" Then cellA.Value = correctedA
                    If correctedB <> "" Then cellB.Value = correctedB
                    Application.EnableEvents = True

                    Msg = IIf(english, _
                        "Minor auto-correction applied for Heat Source / Metered.", _
                        "Correction mineure appliquée automatiquement pour la source et le compteur de chaleur.")

                    AV_Format.AddValidationFeedback "Heat_Source", wsTarget, cellA.Row, Msg, AV_Constants.FORMAT_AUTOCORRECT, english, FormatMap, autoValMap
                    AV_Format.AddValidationFeedback "Heat_Metered", wsTarget, cellA.Row, vbNullString, AV_Constants.FORMAT_DEFAULT, english, FormatMap, autoValMap
                Else
                    AV_Format.AddValidationFeedback "Heat_Source", wsTarget, cellA.Row, vbNullString, AV_Constants.FORMAT_DEFAULT, english, FormatMap, autoValMap
                    AV_Format.AddValidationFeedback "Heat_Metered", wsTarget, cellA.Row, vbNullString, AV_Constants.FORMAT_DEFAULT, english, FormatMap, autoValMap
                End If

                GoTo CleanExit
            End If
        Next r
    End If

    ' ======================================================
    ' STAGE 2 – ANY / ANY(FR) RESOLUTION
    ' ======================================================
    Dim anyTable As ListObject
    Set anyTable = AV_Core.GetValidationTable(AV_Constants.TBL_HEAT_ANY_REF)

    If Not anyTable Is Nothing And Not validationTable Is Nothing Then
        Dim rAny As ListRow
        For Each rAny In anyTable.ListRows
            If StrComp(valA, AV_Core.SafeTrim(rAny.Range(1, 1).Value), vbTextCompare) = 0 Then
                Dim normalizedA As String
                normalizedA = IIf(InStr(1, rAny.Range(1, 1).Value, "(FR)", vbTextCompare) > 0, "ANY(FR)", "ANY")

                For Each r In validationTable.ListRows
                    If StrComp(normalizedA, AV_Core.SafeTrim(r.Range(1, 1).Value), vbTextCompare) = 0 _
                       And StrComp(valB, AV_Core.SafeTrim(r.Range(1, 2).Value), vbTextCompare) = 0 Then

                        correctedB = AV_Core.SafeTrim(r.Range(1, 5).Value)

                        Application.EnableEvents = False
                        If correctedB <> "" Then cellB.Value = correctedB
                        Application.EnableEvents = True

                        Msg = IIf(english, _
                            "Auto-corrected Heat Metered based on Heat Source.", _
                            "Correction automatique du compteur de chaleur selon la source.")

                        AV_Format.AddValidationFeedback "Heat_Metered", wsTarget, cellA.Row, Msg, AV_Constants.FORMAT_AUTOCORRECT, english, FormatMap, autoValMap
                        AV_Format.AddValidationFeedback "Heat_Source", wsTarget, cellA.Row, vbNullString, AV_Constants.FORMAT_DEFAULT, english, FormatMap, autoValMap
                        GoTo CleanExit
                    End If
                Next r
            End If
        Next rAny
    End If

    ' ======================================================
    ' STAGE 3 – WILDCARD NORMALIZATION (Central Heating Plant)
    ' ======================================================
    Dim prefixes As Variant
    prefixes = Array("Central Heating Plant", "Installation de chauffage centrale")

    Dim p As Variant
    For Each p In prefixes
        If LCase(Left(valA, Len(p))) = LCase(p) Then

            Dim remainder As String
            remainder = Trim(Mid(valA, Len(p) + 1))

            Do While Left(remainder, 1) = "-" Or Left(remainder, 1) = " "
                remainder = Mid(remainder, 2)
            Loop

            If remainder = "" Then
                Msg = IIf(english, _
                    "Central Heating Plant entries must specify a heat source after the dash.", _
                    "Les entrées d'installation de chauffage centrale doivent préciser une source après le tiret.")
                AV_Format.AddValidationFeedback "Heat_Source", wsTarget, cellA.Row, Msg, AV_Constants.FORMAT_ERROR, english, FormatMap, autoValMap
                GoTo CleanExit
            End If

            Dim normalizedVal As String
            normalizedVal = p & " - " & remainder

            If normalizedVal <> valA Then
                Application.EnableEvents = False
                cellA.Value = normalizedVal
                Application.EnableEvents = True

                Msg = IIf(english, _
                    "Heat Source normalized to standard format.", _
                    "Source de chaleur normalisée au format standard.")

                AV_Format.AddValidationFeedback "Heat_Source", wsTarget, cellA.Row, Msg, AV_Constants.FORMAT_AUTOCORRECT, english, FormatMap, autoValMap

                If recursionLevel < 1 Then
                    isRunning = False
                    Validate_HeatPairs cellA, sheetName, "Heat_Source", english, recursionLevel + 1, FormatMap, autoValMap
                    Exit Sub
                End If
            End If
        End If
    Next p

    ' ======================================================
    ' STAGE 4 – INVALID COMBINATION
    ' ======================================================
    Msg = IIf(english, _
        "Invalid combination of Heat Source and Heat Metered.", _
        "Combinaison invalide de la source de chaleur et du compteur de chaleur.")

    AV_Format.AddValidationFeedback "Heat_Source", wsTarget, cellA.Row, Msg, AV_Constants.FORMAT_ERROR, english, FormatMap, autoValMap
    AV_Format.AddValidationFeedback "Heat_Metered", wsTarget, cellA.Row, vbNullString, AV_Constants.FORMAT_ERROR, english, FormatMap, autoValMap

CleanExit:
    isRunning = False
End Sub

' ======================================================
' CONSTRUCTION DATE VALIDATION - Updated v2.1
' Validates and auto-corrects date format to YYYY.MM.DD
' ======================================================
Public Sub Validate_ConstructionDate( _
    cell As Range, _
    sheetName As String, _
    english As Boolean, _
    FormatMap As Object, _
    autoValMap As Object _
)

    Static isRunning As Boolean
    If isRunning Then Exit Sub
    isRunning = True

    Dim wsTarget As Worksheet
    Set wsTarget = ThisWorkbook.Sheets(sheetName)

    If FormatMap Is Nothing Then
        Set FormatMap = AV_Format.DefaultFormatMap()
    End If

    Dim val As String
    Dim correctedVal As String
    Dim Msg As String

    val = AV_Core.SafeTrim(cell.Value)
    correctedVal = val

    ' Blank check
    If val = "" Then
        Msg = IIf(english, _
            "Construction Date cannot be blank.", _
            "La date de construction ne peut pas être vide.")
        AV_Format.AddValidationFeedback "Construction_Date", wsTarget, cell.Row, Msg, AV_Constants.FORMAT_ERROR, english, FormatMap, autoValMap
        GoTo CleanExit
    End If

    ' Special case: # is valid
    If val = "#" Then
        AV_Format.AddValidationFeedback "Construction_Date", wsTarget, cell.Row, vbNullString, AV_Constants.FORMAT_DEFAULT, english, FormatMap, autoValMap
        GoTo CleanExit
    End If

    ' Auto-correct common separators to "."
    If InStr(correctedVal, "-") > 0 Or InStr(correctedVal, "/") > 0 Then
        correctedVal = Replace(correctedVal, "-", ".")
        correctedVal = Replace(correctedVal, "/", ".")
    End If

    ' Remove accidental double dots
    Do While InStr(correctedVal, "..") > 0
        correctedVal = Replace(correctedVal, "..", ".")
    Loop

    ' Regex validation (YYYY.MM.DD)
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^\d{4}\.\d{2}\.\d{2}$"
    regex.IgnoreCase = True

    If Not regex.Test(correctedVal) Then
        Msg = IIf(english, _
            "Invalid Construction Date format. Expected: YYYY.MM.DD", _
            "Format de date de construction invalide. Format attendu : AAAA.MM.JJ.")
        AV_Format.AddValidationFeedback "Construction_Date", wsTarget, cell.Row, Msg, AV_Constants.FORMAT_ERROR, english, FormatMap, autoValMap
        GoTo CleanExit
    End If

    ' Calendar validation
    If Not IsDate(Replace(correctedVal, ".", "/")) Then
        Msg = IIf(english, _
            "Invalid calendar date. Please verify year, month, and day.", _
            "Date invalide. Veuillez vérifier l'année, le mois et le jour.")
        AV_Format.AddValidationFeedback "Construction_Date", wsTarget, cell.Row, Msg, AV_Constants.FORMAT_ERROR, english, FormatMap, autoValMap
        GoTo CleanExit
    End If

    ' Year range validation
    Dim yearVal As Long
    yearVal = CLng(Left(correctedVal, 4))
    
    If yearVal < AV_Constants.MIN_CONSTRUCTION_YEAR Or yearVal > AV_Constants.MAX_CONSTRUCTION_YEAR Then
        Msg = IIf(english, _
            "Construction year out of valid range (" & AV_Constants.MIN_CONSTRUCTION_YEAR & "-" & AV_Constants.MAX_CONSTRUCTION_YEAR & ").", _
            "Année de construction hors de la plage valide (" & AV_Constants.MIN_CONSTRUCTION_YEAR & "-" & AV_Constants.MAX_CONSTRUCTION_YEAR & ").")
        AV_Format.AddValidationFeedback "Construction_Date", wsTarget, cell.Row, Msg, AV_Constants.FORMAT_ERROR, english, FormatMap, autoValMap
        GoTo CleanExit
    End If

    ' Apply correction if needed
    If correctedVal <> val Then
        Application.EnableEvents = False
        cell.Value = correctedVal
        Application.EnableEvents = True

        Msg = IIf(english, _
            "Date format auto-corrected to YYYY.MM.DD.", _
            "Format de date corrigé automatiquement à AAAA.MM.JJ.")
        AV_Format.AddValidationFeedback "Construction_Date", wsTarget, cell.Row, Msg, AV_Constants.FORMAT_AUTOCORRECT, english, FormatMap, autoValMap
    Else
        AV_Format.AddValidationFeedback "Construction_Date", wsTarget, cell.Row, vbNullString, AV_Constants.FORMAT_DEFAULT, english, FormatMap, autoValMap
    End If

CleanExit:
    isRunning = False
End Sub
