Attribute VB_Name = "ValidateConstructionDates"
Option Explicit

Public Sub Validate_Column_Construction_Date(cell As Range, sheetName As String, Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)
    Static isRunning As Boolean
    If isRunning Then Exit Sub
    isRunning = True
    
    Dim wsTargetSheet As Worksheet
    Dim val As String, correctedVal As String
    Dim Msg As String
    Dim pattern As String
    Dim regex As Object
    Dim isValid As Boolean, wasCorrected As Boolean
    
    Dim DevFuncName As String
    DevFuncName = "Construction_Date"
    
    ' Default FormatMap if not provided
    If FormatMap Is Nothing Then
        Set FormatMap = DefaultFormatMap()
    End If
    
    ' Target sheet
    Set wsTargetSheet = ThisWorkbook.Sheets(sheetName)
    
    val = Trim(CStr(cell.value))
    correctedVal = val
    
    ' === BLANK VALUE CHECK ===
    If val = "" Then
        Msg = IIf(english, _
            "Construction Date cannot be blank.", _
            "La date de construction ne peut pas être vide.")
        AddValidationFeedback DevFuncName, wsTargetSheet, cell.row, Msg, "Error", english, FormatMap, AutoValMap
        GoTo CleanExit
    End If
    
    ' === SPECIAL CASE: # is valid ===
    If val = "#" Then
        AddValidationFeedback DevFuncName, wsTargetSheet, cell.row, "", "Default", english, FormatMap, AutoValMap
        GoTo CleanExit
    End If
    
    ' === AUTO-CORRECT COMMON FORMATS ===
    ' Replace "-" or "/" with "."
    If InStr(correctedVal, "-") > 0 Or InStr(correctedVal, "/") > 0 Then
        correctedVal = Replace(correctedVal, "-", ".")
        correctedVal = Replace(correctedVal, "/", ".")
        wasCorrected = True
    End If
    
    ' Remove accidental double dots
    Do While InStr(correctedVal, "..") > 0
        correctedVal = Replace(correctedVal, "..", ".")
    Loop
    
    ' === VALIDATION PATTERN ===
    pattern = "^\d{4}\.\d{2}\.\d{2}$"
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = pattern
    regex.IgnoreCase = True
    
    isValid = regex.Test(correctedVal)
    
    ' === VALIDATION OUTCOME ===
    If isValid Then
        ' Optional: verify if it’s an actual valid calendar date
        If Not IsDate(Replace(correctedVal, ".", "/")) Then
            Msg = IIf(english, _
                "Invalid calendar date. Please verify year, month, and day.", _
                "Date de calendrier invalide. Vérifiez l'année, le mois et le jour.")
            AddValidationFeedback DevFuncName, wsTargetSheet, cell.row, Msg, "Error", english, FormatMap, AutoValMap
            GoTo CleanExit
        End If
        
        ' Apply correction if needed
        If wasCorrected And correctedVal <> val Then
            Application.EnableEvents = False
            cell.value = correctedVal
            Application.EnableEvents = True
            
            Msg = IIf(english, _
                "Date format auto-corrected to YYYY.MM.DD.", _
                "Format de date corrigé automatiquement à AAAA.MM.JJ.")
            AddValidationFeedback DevFuncName, wsTargetSheet, cell.row, Msg, "Autocorrect", english, FormatMap, AutoValMap
        Else
            AddValidationFeedback DevFuncName, wsTargetSheet, cell.row, "", "Default", english, FormatMap, AutoValMap
        End If
    Else
        Msg = IIf(english, _
            "Invalid Construction Date format. Expected: YYYY.MM.DD", _
            "Format de date de construction invalide. Format attendu : AAAA.MM.JJ.")
        AddValidationFeedback DevFuncName, wsTargetSheet, cell.row, Msg, "Error", english, FormatMap, AutoValMap
    End If

CleanExit:
    isRunning = False
End Sub


