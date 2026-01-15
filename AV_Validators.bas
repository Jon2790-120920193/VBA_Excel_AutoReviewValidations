Attribute VB_Name = "AV_Validators"
Option Explicit

' ======================================================
' AV_Validators.bas
' All validation logic (NO UI / NO formatting helpers)
' ======================================================

Private Const MODULE_NAME As String = "AV_Validators"

' ======================================================
' PUBLIC ENTRY POINTS (DO NOT RENAME)
' These are called dynamically via mapping
' ======================================================

' -------------------- Electricity --------------------
Public Sub Validate_Column_Electricity(cell As Range, sheetName As String, _
    Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)

    Validate_PairedFields cell, sheetName, "Electricity", "Electricity_Metered", _
        "ElectricityPairValidation", english, FormatMap, AutoValMap
End Sub

Public Sub Validate_Column_Electricity_Metered(cell As Range, sheetName As String, _
    Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)

    Validate_PairedFields cell, sheetName, "Electricity_Metered", "Electricity", _
        "ElectricityPairValidation", english, FormatMap, AutoValMap
End Sub

' -------------------- Plumbing --------------------
Public Sub Validate_Column_Plumbing(cell As Range, sheetName As String, _
    Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)

    Validate_PairedFields cell, sheetName, "Plumbing", "Water_Metered", _
        "PlumbingPairValidation", english, FormatMap, AutoValMap
End Sub

Public Sub Validate_Column_Water_Metered(cell As Range, sheetName As String, _
    Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)

    Validate_PairedFields cell, sheetName, "Water_Metered", "Plumbing", _
        "PlumbingPairValidation", english, FormatMap, AutoValMap
End Sub

' -------------------- GIW --------------------
Public Sub Validate_Column_GIWQuantity(cell As Range, sheetName As String, _
    Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)

    If Validate_GIWQuantity(cell, sheetName, "GIWQuantity", english, FormatMap, AutoValMap) Then
        Validate_GIWIncluded GetSiblingCell(cell, sheetName, "GIWIncluded"), _
            sheetName, "GIWIncluded", english, FormatMap, AutoValMap
    End If
End Sub

Public Sub Validate_Column_GIWIncluded(cell As Range, sheetName As String, _
    Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)

    If Validate_GIWIncluded(cell, sheetName, "GIWIncluded", english, FormatMap, AutoValMap) Then
        Validate_GIWQuantity GetSiblingCell(cell, sheetName, "GIWQuantity"), _
            sheetName, "GIWQuantity", english, FormatMap, AutoValMap
    End If
End Sub

' -------------------- Heat --------------------
Public Sub Validate_Column_Heat_Source(cell As Range, sheetName As String, _
    Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)

    Validate_HeatPairs cell, sheetName, "Heat_Source", english, 0, FormatMap, AutoValMap
End Sub

Public Sub Validate_Column_Heat_Metered(cell As Range, sheetName As String, _
    Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)

    Validate_HeatPairs cell, sheetName, "Heat_Metered", english, 0, FormatMap, AutoValMap
End Sub

' -------------------- Construction Date --------------------
Public Sub Validate_Column_Construction_Date(cell As Range, sheetName As String, _
    Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)

    Validate_ConstructionDate cell, sheetName, english, FormatMap, AutoValMap
End Sub

' ======================================================
' GENERIC PAIR VALIDATION (Electricity / Plumbing)
' ======================================================
Private Sub Validate_PairedFields( _
    cell As Range, sheetName As String, _
    ThisFunc As String, OtherFunc As String, _
    DefaultRuleTable As String, _
    english As Boolean, FormatMap As Object, AutoValMap As Object)

    Static isRunning As Boolean
    If isRunning Then Exit Sub
    isRunning = True

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    Dim otherCell As Range
    Set otherCell = GetSiblingCell(cell, sheetName, OtherFunc)

    If otherCell Is Nothing Then GoTo ExitSafe

    Dim ruleTable As String
    ruleTable = GetRuleTableNameFromAutoValMap(AutoValMap, ThisFunc, DefaultRuleTable)

    Dim result As Boolean
    result = RunPairRuleValidation( _
        ws, cell, otherCell, ThisFunc, OtherFunc, _
        ruleTable, english, FormatMap, AutoValMap)

ExitSafe:
    isRunning = False
End Sub

' ======================================================
' PAIR RULE ENGINE (shared)
' ======================================================
Private Function RunPairRuleValidation( _
    ws As Worksheet, cellA As Range, cellB As Range, _
    FuncA As String, FuncB As String, _
    RuleTableName As String, _
    english As Boolean, FormatMap As Object, AutoValMap As Object _
) As Boolean

    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets("Config")

    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsConfig.ListObjects(RuleTableName)
    On Error GoTo 0

    If lo Is Nothing Then
        DebugMessage "Rule table missing: " & RuleTableName, MODULE_NAME
        RunPairRuleValidation = False
        Exit Function
    End If

    Dim valA As String, valB As String
    valA = Trim(CStr(cellA.Value))
    valB = Trim(CStr(cellB.Value))

    Dim r As ListRow
    For Each r In lo.ListRows
        If StrComp(valA, Trim(r.Range(1, 1).Value), vbTextCompare) = 0 _
           And StrComp(valB, Trim(r.Range(1, 2).Value), vbTextCompare) = 0 Then

            AddValidationFeedback FuncA, ws, cellA.Row, "", "Default", english, FormatMap, AutoValMap
            AddValidationFeedback FuncB, ws, cellA.Row, "", "Default", english, FormatMap, AutoValMap
            RunPairRuleValidation = True
            Exit Function
        End If
    Next r

    AddValidationFeedback FuncA, ws, cellA.Row, _
        IIf(english, "Invalid value pairing.", "Combinaison invalide."), _
        "Error", english, FormatMap, AutoValMap

    AddValidationFeedback FuncB, ws, cellA.Row, "", "Error", english, FormatMap, AutoValMap
    RunPairRuleValidation = False
End Function

' ======================================================
' GIW VALIDATION (SPECIAL LOGIC) - INSERTED VERBATIM LOGIC
' ======================================================

Private Function Validate_GIWQuantity( _
    cell As Range, sheetName As String, _
    FuncName As String, _
    english As Boolean, _
    FormatMap As Object, AutoValMap As Object _
) As Boolean

    Dim wsTargetSheet As Worksheet
    Set wsTargetSheet = ThisWorkbook.Sheets(sheetName)

    Dim val As String, parts() As String
    Dim num1 As Long, num2 As Long
    Dim MAX_ALLOWED As Long
    Dim Msg As String
    Dim OtherFuncName As String

    ' Pair function name for defaulting
    If StrComp(FuncName, "GIWQuantity", vbTextCompare) = 0 Then
        OtherFuncName = "GIWIncluded"
    Else
        OtherFuncName = "GIWQuantity"
    End If

    MAX_ALLOWED = 1000

    val = Trim(CStr(cell.Value))
    val = Replace(val, ".", ",")
    val = Replace(val, " ", "")

    ' Normalize value in-cell if we changed punctuation/spacing
    If val <> Trim(CStr(cell.Value)) Then
        Application.EnableEvents = False
        cell.Value = val
        Application.EnableEvents = True
    End If

    ' ---------------------------
    ' Handle (A,B) form
    ' ---------------------------
    If Left$(val, 1) = "(" And Right$(val, 1) = ")" Then
        val = Mid$(val, 2, Len(val) - 2)   ' remove parentheses
        Application.EnableEvents = False
        cell.Value = val                  ' silently correct
        Application.EnableEvents = True
    End If
    ' ---------------------------

    ' --- Basic validation ---
    If val = "" Then
        Msg = IIf(english, "Cannot be empty", "Ne peut pas être vide.")
        AddValidationFeedback FuncName, wsTargetSheet, cell.Row, Msg, "Error", english, FormatMap, AutoValMap
        Validate_GIWQuantity = False
        Exit Function
    End If

    ' --- Special placeholder ---
    If val = "#" Then
        Application.EnableEvents = False
        cell.Value = "#,#"
        Application.EnableEvents = True

        Msg = IIf(english, "Auto-corrected placeholder", "Correction automatique")
        AddValidationFeedback FuncName, wsTargetSheet, cell.Row, Msg, "Error", english, FormatMap, AutoValMap
        Validate_GIWQuantity = True
        Exit Function
    End If

    ' --- Single numeric entry -> normalize to n,n ---
    If IsNumeric(val) And InStr(1, val, ",", vbTextCompare) = 0 Then

        If CLng(val) > MAX_ALLOWED Then
            Msg = IIf(english, _
                      "Max value: " & MAX_ALLOWED & " surpassed", _
                      "Valeur maximale : " & MAX_ALLOWED & " dépassée")
            AddValidationFeedback FuncName, wsTargetSheet, cell.Row, Msg, "Error", english, FormatMap, AutoValMap
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
        AddValidationFeedback FuncName, wsTargetSheet, cell.Row, Msg, "Autocorrect", english, FormatMap, AutoValMap

        ' Ensure paired field returns to default styling/state
        AddValidationFeedback OtherFuncName, wsTargetSheet, cell.Row, vbNullString, "Default", english, FormatMap, AutoValMap

        Validate_GIWQuantity = True
        Exit Function
    End If

    ' --- Must be "A,B" ---
    parts = Split(val, ",")
    If UBound(parts) <> 1 Then
        Msg = IIf(english, _
                  "Entry not valid, must be 'Number,Number'", _
                  "Entrée non valide, le format doit être 'Nombre,Nombre'")
        AddValidationFeedback FuncName, wsTargetSheet, cell.Row, Msg, "Error", english, FormatMap, AutoValMap
        Validate_GIWQuantity = False
        Exit Function
    End If

    ' --- Parse values (allow #) ---
    On Error GoTo ParseFail
    If parts(0) <> "#" Then num1 = CLng(parts(0))
    If parts(1) <> "#" Then num2 = CLng(parts(1))
    On Error GoTo 0

    ' If we got here, treat as valid basic structure
    AddValidationFeedback FuncName, wsTargetSheet, cell.Row, vbNullString, "Default", english, FormatMap, AutoValMap
    Validate_GIWQuantity = True
    Exit Function

ParseFail:
    Msg = IIf(english, _
              "Entry not valid, must be numeric values or '#'", _
              "Entrée non valide : valeurs numériques requises ou '#'.")
    AddValidationFeedback FuncName, wsTargetSheet, cell.Row, Msg, "Error", english, FormatMap, AutoValMap
    Validate_GIWQuantity = False
End Function


Private Function Validate_GIWIncluded( _
    cell As Range, sheetName As String, _
    FuncName As String, _
    english As Boolean, _
    FormatMap As Object, AutoValMap As Object _
) As Boolean

    Static isRunning As Boolean
    If isRunning Then
        Validate_GIWIncluded = True
        Exit Function
    End If
    isRunning = True

    Dim ws As Worksheet, wsConfig As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    Set wsConfig = ThisWorkbook.Sheets("Config")

    Dim GIWIncludedVal As String
    GIWIncludedVal = Trim(CStr(cell.Value))

    Dim quantityCell As Range
    Set quantityCell = GetSiblingCell(cell, sheetName, "GIWQuantity")

    Dim GIWQuantityVal As String
    If Not quantityCell Is Nothing Then GIWQuantityVal = Trim(CStr(quantityCell.Value))

    ' --- Validation table (prefer config from AutoValidation map) ---
    Dim ruleTable As String
    ruleTable = GetRuleTableNameFromAutoValMap(AutoValMap, FuncName, "GIWValidationTable")

    Dim validationTable As ListObject
    On Error Resume Next
    Set validationTable = wsConfig.ListObjects(ruleTable)
    On Error GoTo 0

    If validationTable Is Nothing Then
        DebugMessage "Validation table not found for GIW: " & ruleTable, MODULE_NAME
        AddValidationFeedback FuncName, ws, cell.Row, _
            IIf(english, "Configuration error: GIW rule table missing.", "Erreur de configuration : table de règles GIW manquante."), _
            "Error", english, FormatMap, AutoValMap
        Validate_GIWIncluded = False
        GoTo CleanExit
    End If

    ' --- Find expected rule for the Included value ---
    Dim expectedRule As String
    expectedRule = vbNullString

    Dim tblRow As ListRow
    For Each tblRow In validationTable.ListRows
        If StrComp(Trim(CStr(tblRow.Range(1, 1).Value)), GIWIncludedVal, vbTextCompare) = 0 Then
            expectedRule = Trim(CStr(tblRow.Range(1, 2).Value))
            Exit For
        End If
    Next tblRow

    If expectedRule = vbNullString Then
        AddValidationFeedback FuncName, ws, cell.Row, _
            IIf(english, "Error: Invalid entry", "Erreur : entrée non valide."), _
            "Error", english, FormatMap, AutoValMap
        Validate_GIWIncluded = False
        GoTo CleanExit
    End If

    ' --- Validate quantity value ---
    If GIWQuantityVal = vbNullString Then
        AddValidationFeedback FuncName, ws, cell.Row, _
            IIf(english, "Error: Cannot be empty", "Erreur : ne peut pas être vide."), _
            "Error", english, FormatMap, AutoValMap
        Validate_GIWIncluded = False
        GoTo CleanExit
    End If

    Dim quantityParts() As String
    quantityParts = Split(GIWQuantityVal, ",")

    If UBound(quantityParts) <> 1 Then
        AddValidationFeedback FuncName, ws, cell.Row, _
            IIf(english, "Entry not valid, must be 'Number,Number'", "Entrée non valide, le format doit être 'Nombre,Nombre'"), _
            "Error", english, FormatMap, AutoValMap
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

                    AddValidationFeedback "GIWQuantity", ws, cell.Row, Msg, "Autocorrect", english, FormatMap, AutoValMap
                    AddValidationFeedback "GIWIncluded", ws, cell.Row, vbNullString, "Default", english, FormatMap, AutoValMap

                    Validate_GIWIncluded = True
                    GoTo CleanExit
                Else
                    Msg = IIf(english, _
                        "Invalid Entry, value must be 0,0 when GIW Included = 'No'", _
                        "Combinaison invalide : la valeur doit être 0,0 lorsque GIW inclus = 'Non'")
                    AddValidationFeedback "GIWQuantity", ws, cell.Row, Msg, "Error", english, FormatMap, AutoValMap
                    AddValidationFeedback "GIWIncluded", ws, cell.Row, vbNullString, "Error", english, FormatMap, AutoValMap
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
                    AddValidationFeedback "GIWQuantity", ws, cell.Row, Msg, "Error", english, FormatMap, AutoValMap
                    AddValidationFeedback "GIWIncluded", ws, cell.Row, vbNullString, "Error", english, FormatMap, AutoValMap
                    Validate_GIWIncluded = False
                    GoTo CleanExit
                End If

                Msg = IIf(english, _
                    "GIW Quantity must be positive when GIW Included is 'Yes' or 'Partially'.", _
                    "La quantité GIW doit être positive lorsque GIW inclus est 'Oui' ou 'Partiellement'.")
                AddValidationFeedback "GIWQuantity", ws, cell.Row, Msg, "Error", english, FormatMap, AutoValMap
                AddValidationFeedback "GIWIncluded", ws, cell.Row, vbNullString, "Error", english, FormatMap, AutoValMap
                Validate_GIWIncluded = False
                GoTo CleanExit

            Case "#"
                Msg = IIf(english, _
                    "GIW Quantity must be '#,#' when GIW Included is 'Not Applicable'.", _
                    "La quantité GIW doit être '#,#' lorsque GIW inclus est 'Non applicable'.")
                AddValidationFeedback "GIWQuantity", ws, cell.Row, Msg, "Error", english, FormatMap, AutoValMap
                AddValidationFeedback "GIWIncluded", ws, cell.Row, vbNullString, "Error", english, FormatMap, AutoValMap
                Validate_GIWIncluded = False
                GoTo CleanExit

            Case Else
                Msg = IIf(english, "Invalid combination.", "Combinaison invalide.")
                AddValidationFeedback "GIWQuantity", ws, cell.Row, Msg, "Error", english, FormatMap, AutoValMap
                AddValidationFeedback "GIWIncluded", ws, cell.Row, vbNullString, "Error", english, FormatMap, AutoValMap
                Validate_GIWIncluded = False
                GoTo CleanExit
        End Select

    Else
        ' Valid combination: reset both to default
        AddValidationFeedback "GIWIncluded", ws, cell.Row, vbNullString, "Default", english, FormatMap, AutoValMap
        AddValidationFeedback "GIWQuantity", ws, cell.Row, vbNullString, "Default", english, FormatMap, AutoValMap
        Validate_GIWIncluded = True
    End If

CleanExit:
    isRunning = False
End Function


' ======================================================
' HEAT VALIDATION (MULTI-STAGE, PRESERVED LOGIC)
' ======================================================

Private Sub Validate_HeatPairs( _
    cell As Range, sheetName As String, _
    FuncName As String, _
    english As Boolean, _
    recursionLevel As Long, _
    FormatMap As Object, AutoValMap As Object _
)

    Static isRunning As Boolean
    If isRunning Then Exit Sub
    isRunning = True

    Dim wsTarget As Worksheet
    Dim wsConfig As Worksheet
    Set wsTarget = ThisWorkbook.Sheets(sheetName)
    Set wsConfig = ThisWorkbook.Sheets("Config")

    Dim cellA As Range, cellB As Range
    If FuncName = "Heat_Source" Then
        Set cellA = cell
        Set cellB = GetSiblingCell(cell, sheetName, "Heat_Metered")
    Else
        Set cellB = cell
        Set cellA = GetSiblingCell(cell, sheetName, "Heat_Source")
    End If

    If cellA Is Nothing Or cellB Is Nothing Then GoTo CleanExit

    Dim valA As String, valB As String
    valA = Trim(CStr(cellA.Value))
    valB = Trim(CStr(cellB.Value))

    Dim ruleTable As String
    ruleTable = GetRuleTableNameFromAutoValMap(AutoValMap, "Heat_Source", "HeatSourcePairValidation")

    Dim validationTable As ListObject
    On Error Resume Next
    Set validationTable = wsConfig.ListObjects(ruleTable)
    On Error GoTo 0

    Dim Msg As String
    Dim correctedA As String, correctedB As String
    Dim autoCorrect As Boolean
    Dim isMatchFound As Boolean
    isMatchFound = False

    ' ======================================================
    ' STAGE 1 — EXACT MATCH
    ' ======================================================
    If Not validationTable Is Nothing Then
        Dim r As ListRow
        For Each r In validationTable.ListRows
            If StrComp(valA, Trim(r.Range(1, 1).Value), vbTextCompare) = 0 _
               And StrComp(valB, Trim(r.Range(1, 2).Value), vbTextCompare) = 0 Then

                isMatchFound = True
                autoCorrect = (LCase(Trim(CStr(r.Range(1, 3).Value))) = "true")

                If autoCorrect Then
                    correctedA = Trim(CStr(r.Range(1, 4).Value))
                    correctedB = Trim(CStr(r.Range(1, 5).Value))

                    Application.EnableEvents = False
                    If correctedA <> "" Then cellA.Value = correctedA
                    If correctedB <> "" Then cellB.Value = correctedB
                    Application.EnableEvents = True

                    Msg = IIf(english, _
                        "Minor auto-correction applied for Heat Source / Metered.", _
                        "Correction mineure appliquée automatiquement pour la source et le compteur de chaleur.")

                    AddValidationFeedback "Heat_Source", wsTarget, cellA.Row, Msg, "Autocorrect", english, FormatMap, AutoValMap
                    AddValidationFeedback "Heat_Metered", wsTarget, cellA.Row, vbNullString, "Default", english, FormatMap, AutoValMap
                Else
                    AddValidationFeedback "Heat_Source", wsTarget, cellA.Row, vbNullString, "Default", english, FormatMap, AutoValMap
                    AddValidationFeedback "Heat_Metered", wsTarget, cellA.Row, vbNullString, "Default", english, FormatMap, AutoValMap
                End If

                GoTo CleanExit
            End If
        Next r
    End If

    ' ======================================================
    ' STAGE 2 — ANY / ANY(FR) RESOLUTION
    ' ======================================================
    Dim anyTable As ListObject
    On Error Resume Next
    Set anyTable = wsConfig.ListObjects("HeatSourceANYRefTable")
    On Error GoTo 0

    If Not anyTable Is Nothing And Not validationTable Is Nothing Then
        Dim rAny As ListRow
        For Each rAny In anyTable.ListRows
            If StrComp(valA, Trim(rAny.Range(1, 1).Value), vbTextCompare) = 0 Then
                Dim normalizedA As String
                normalizedA = IIf(InStr(1, rAny.Range(1, 1).Value, "(FR)", vbTextCompare) > 0, "ANY(FR)", "ANY")

                For Each r In validationTable.ListRows
                    If StrComp(normalizedA, Trim(r.Range(1, 1).Value), vbTextCompare) = 0 _
                       And StrComp(valB, Trim(r.Range(1, 2).Value), vbTextCompare) = 0 Then

                        correctedB = Trim(CStr(r.Range(1, 5).Value))

                        Application.EnableEvents = False
                        If correctedB <> "" Then cellB.Value = correctedB
                        Application.EnableEvents = True

                        Msg = IIf(english, _
                            "Auto-corrected Heat Metered based on Heat Source.", _
                            "Correction automatique du compteur de chaleur selon la source.")

                        AddValidationFeedback "Heat_Metered", wsTarget, cellA.Row, Msg, "Autocorrect", english, FormatMap, AutoValMap
                        AddValidationFeedback "Heat_Source", wsTarget, cellA.Row, vbNullString, "Default", english, FormatMap, AutoValMap
                        GoTo CleanExit
                    End If
                Next r
            End If
        Next rAny
    End If

    ' ======================================================
    ' STAGE 3 — WILDCARD NORMALIZATION (Central Heating Plant)
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
                AddValidationFeedback "Heat_Source", wsTarget, cellA.Row, Msg, "Error", english, FormatMap, AutoValMap
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

                AddValidationFeedback "Heat_Source", wsTarget, cellA.Row, Msg, "Autocorrect", english, FormatMap, AutoValMap

                If recursionLevel < 1 Then
                    isRunning = False
                    Validate_HeatPairs cellA, sheetName, "Heat_Source", english, recursionLevel + 1, FormatMap, AutoValMap
                    Exit Sub
                End If
            End If
        End If
    Next p

    ' ======================================================
    ' STAGE 4 — INVALID COMBINATION
    ' ======================================================
    Msg = IIf(english, _
        "Invalid combination of Heat Source and Heat Metered.", _
        "Combinaison invalide de la source de chaleur et du compteur de chaleur.")

    AddValidationFeedback "Heat_Source", wsTarget, cellA.Row, Msg, "Error", english, FormatMap, AutoValMap
    AddValidationFeedback "Heat_Metered", wsTarget, cellA.Row, vbNullString, "Error", english, FormatMap, AutoValMap

CleanExit:
    isRunning = False
End Sub


' ======================================================
' CONSTRUCTION DATE VALIDATION
' ======================================================
Public Sub Validate_Column_Construction_Date( _
    cell As Range, _
    sheetName As String, _
    Optional english As Boolean = True, _
    Optional FormatMap As Object, _
    Optional AutoValMap As Object _
)

    Static isRunning As Boolean
    If isRunning Then Exit Sub
    isRunning = True

    Dim wsTarget As Worksheet
    Set wsTarget = ThisWorkbook.Sheets(sheetName)

    If FormatMap Is Nothing Then
        Set FormatMap = DefaultFormatMap()
    End If

    Dim val As String
    Dim correctedVal As String
    Dim Msg As String

    val = Trim(CStr(cell.Value))
    correctedVal = val

    ' -----------------------------
    ' BLANK CHECK
    ' -----------------------------
    If val = "" Then
        Msg = IIf(english, _
            "Construction Date cannot be blank.", _
            "La date de construction ne peut pas être vide.")
        AddValidationFeedback "Construction_Date", wsTarget, cell.Row, Msg, "Error", english, FormatMap, AutoValMap
        GoTo CleanExit
    End If

    ' -----------------------------
    ' SPECIAL CASE: # IS VALID
    ' -----------------------------
    If val = "#" Then
        AddValidationFeedback "Construction_Date", wsTarget, cell.Row, vbNullString, "Default", english, FormatMap, AutoValMap
        GoTo CleanExit
    End If

    ' -----------------------------
    ' AUTO-CORRECT COMMON SEPARATORS
    ' -----------------------------
    If InStr(correctedVal, "-") > 0 Or InStr(correctedVal, "/") > 0 Then
        correctedVal = Replace(correctedVal, "-", ".")
        correctedVal = Replace(correctedVal, "/", ".")
    End If

    Do While InStr(correctedVal, "..") > 0
        correctedVal = Replace(correctedVal, "..", ".")
    Loop

    ' -----------------------------
    ' REGEX VALIDATION (YYYY.MM.DD)
    ' -----------------------------
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = "^\d{4}\.\d{2}\.\d{2}$"
    regex.IgnoreCase = True

    If Not regex.Test(correctedVal) Then
        Msg = IIf(english, _
            "Invalid Construction Date format. Expected: YYYY.MM.DD", _
            "Format de date de construction invalide. Format attendu : AAAA.MM.JJ.")
        AddValidationFeedback "Construction_Date", wsTarget, cell.Row, Msg, "Error", english, FormatMap, AutoValMap
        GoTo CleanExit
    End If

    ' -----------------------------
    ' CALENDAR VALIDATION
    ' -----------------------------
    If Not IsDate(Replace(correctedVal, ".", "/")) Then
        Msg = IIf(english, _
            "Invalid calendar date. Please verify year, month, and day.", _
            "Date invalide. Veuillez vérifier l'année, le mois et le jour.")
        AddValidationFeedback "Construction_Date", wsTarget, cell.Row, Msg, "Error", english, FormatMap, AutoValMap
        GoTo CleanExit
    End If

    ' -----------------------------
    ' APPLY CORRECTION IF NEEDED
    ' -----------------------------
    If correctedVal <> val Then
        Application.EnableEvents = False
        cell.Value = correctedVal
        Application.EnableEvents = True

        Msg = IIf(english, _
            "Date format auto-corrected to YYYY.MM.DD.", _
            "Format de date corrigé automatiquement à AAAA.MM.JJ.")
        AddValidationFeedback "Construction_Date", wsTarget, cell.Row, Msg, "Autocorrect", english, FormatMap, AutoValMap
    Else
        AddValidationFeedback "Construction_Date", wsTarget, cell.Row, vbNullString, "Default", english, FormatMap, AutoValMap
    End If

CleanExit:
    isRunning = False
End Sub


' ======================================================
' SHARED DEPENDENCY RESOLVER
' ======================================================
Private Function GetSiblingCell(cell As Range, sheetName As String, _
    TargetFuncName As String) As Range

    Dim ws As Worksheet
    Dim wsConfig As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    Set wsConfig = ThisWorkbook.Sheets("Config")

    Dim lo As ListObject
    Set lo = wsConfig.ListObjects("AutoValidationCommentPrefixMappingTable")

    Dim r As ListRow
    For Each r In lo.ListRows
        If Trim(r.Range.Columns("Dev Function Names").Value) = TargetFuncName Then
            Set GetSiblingCell = ws.Range( _
                r.Range.Columns("ReviewSheet Column Header").Value & cell.Row)
            Exit Function
        End If
    Next r
End Function


