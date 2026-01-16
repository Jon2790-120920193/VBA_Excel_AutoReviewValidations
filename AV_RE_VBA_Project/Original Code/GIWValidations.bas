Attribute VB_Name = "GIWValidations"
Option Explicit

Private Const MODULE_NAME As String = "GIWValidations"
Private Const DevFuncName1 As String = "GIWQuantity"
Private Const DevFuncName2 As String = "GIWIncluded"

' === PUBLIC ENTRY POINTS ===

Public Sub Validate_Column_GIWQuantity(cell As Range, sheetName As String, Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)
    If FormatMap Is Nothing Then
        Set FormatMap = DefaultFormatMap()
    End If
    
    If RunGIWQuantityValidation(cell, sheetName, DevFuncName1, english, FormatMap, AutoValMap) Then
        Dim dependentCell As Range
        Set dependentCell = GetDependentCell(cell, sheetName)
        RunGIWIncludedValidation dependentCell, sheetName, DevFuncName2, english, FormatMap, AutoValMap
    End If
End Sub

Public Sub Validate_Column_GIWIncluded(cell As Range, sheetName As String, Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object)
    If FormatMap Is Nothing Then
        Set FormatMap = DefaultFormatMap()
    End If
    
    If RunGIWIncludedValidation(cell, sheetName, DevFuncName1, english, FormatMap, AutoValMap) Then
        Dim dependentCell As Range
        Set dependentCell = GetDependentCell(cell, sheetName)
        RunGIWQuantityValidation dependentCell, sheetName, DevFuncName2, english, FormatMap, AutoValMap
    End If
End Sub

Private Function RunGIWQuantityValidation(cell As Range, sheetName As String, Optional CalledFuncName As String, Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object) As Boolean
'(cellA As Range, cellB As Range, sheetName As String, Optional CalledFuncName As String, Optional english As Boolean = True, Optional formatMap As Object, Optional AutoValMap As Object) As Boolean
    Dim val As String, parts() As String, num1 As Long, num2 As Long, MAX_ALLOWED As Long
    Dim originalValue As String
    Dim Msg As String
    Dim wsTargetSheet As Worksheet
    Dim OtherFuncName As String
    
    If CalledFuncName = DevFuncName1 Then OtherFuncName = DevFuncName2 Else OtherFuncName = DevFuncName1
    
    Set wsTargetSheet = ThisWorkbook.Sheets(sheetName)
    
    MAX_ALLOWED = 1000
    val = Trim(cell.value)
    val = Replace(val, ".", ",")
    val = Replace(val, " ", "")
    
    Application.EnableEvents = False
    If val <> Trim(cell.value) Then cell.value = val
    Application.EnableEvents = True
    
    ' --- Basic validation ---
    If val = "" Then
        If english = True Then
            Msg = "Cannot be empty"
        Else
            Msg = "Ne peut pas être vide."
        End If
        AddValidationFeedback CalledFuncName, wsTargetSheet, cell.row, Msg, "Error", english, FormatMap, AutoValMap
        RunGIWQuantityValidation = False: Exit Function
    End If
    
    If val = "#" Then
        Application.EnableEvents = False
        cell.value = "#,#"
        Application.EnableEvents = True
        
        If english = True Then
            Msg = "Auto-corrected placeholder"
        Else
            Msg = "Correction automatique"
        End If
        
        AddValidationFeedback CalledFuncName, wsTargetSheet, cell.row, Msg, "Error", english, FormatMap, AutoValMap
        RunGIWQuantityValidation = True: Exit Function
    End If
    
    If IsNumeric(val) And InStr(val, ",") = 0 Then
        If CLng(val) > MAX_ALLOWED Then
        
            If english = True Then
                Msg = "Valeur maximale: " & MAX_ALLOWED & " dépassé"
            Else
                Msg = "Max value: " & MAX_ALLOWED & " surpassed"
            End If
            
            AddValidationFeedback CalledFuncName, wsTargetSheet, cell.row, Msg, "Error", english, FormatMap, AutoValMap
            RunGIWQuantityValidation = False: Exit Function
            
        End If
        
        num1 = CLng(val)
        Application.EnableEvents = False
        cell.value = num1 & "," & num1
        Application.EnableEvents = True
        
        If english = True Then
            Msg = "Format has been automatically corrected by the system"
        Else
            Msg = "Le format a été automatiquement corrigé par le système."
        End If
            
        AddValidationFeedback CalledFuncName, wsTargetSheet, cell.row, Msg, "Autocorrect", english, FormatMap, AutoValMap
        AddValidationFeedback OtherFuncName, wsTargetSheet, cell.row, "", "Default", english, FormatMap, AutoValMap
        RunGIWQuantityValidation = True: Exit Function
    End If
    
    parts = Split(val, ",")
    If UBound(parts) <> 1 Then
    
        If english = True Then
            Msg = "Entry not valid, must be 'Number:Number'"
        Else
            Msg = "Entrée non valide, le format doit être 'Nombre:Nombre"
        End If
            
        AddValidationFeedback CalledFuncName, wsTargetSheet, cell.row, Msg, "Error", english, FormatMap, AutoValMap
        RunGIWQuantityValidation = False: Exit Function
    End If
    
    If parts(0) <> "#" Then num1 = CLng(parts(0))
    If parts(1) <> "#" Then num2 = CLng(parts(1))
    
    AddValidationFeedback CalledFuncName, wsTargetSheet, cell.row, Msg, "Default", english, FormatMap, AutoValMap
    
    RunGIWQuantityValidation = True
    
End Function


Private Function RunGIWIncludedValidation(cell As Range, sheetName As String, Optional CalledFuncName As String, Optional english As Boolean = True, Optional FormatMap As Object, Optional AutoValMap As Object) As Boolean
    Dim ws As Worksheet, wsConfig As Worksheet
    Dim RowNum As Long
    Dim GIWIncludedVal As String, GIWQuantityVal As String
    Dim validationTable As ListObject
    Dim tblRow As ListRow
    Dim expectedRule As String
    Dim quantityCell As Range
    Dim quantityParts() As String
    Dim num1 As Long, num2 As Long
    Dim isValid As Boolean
    Dim Msg As String
    Dim OtherFuncName As String
    
    If CalledFuncName = DevFuncName1 Then OtherFuncName = DevFuncName2 Else OtherFuncName = DevFuncName1

    Set ws = ThisWorkbook.Sheets(sheetName)
    Set wsConfig = ThisWorkbook.Sheets("Config")
    RowNum = cell.row
    GIWIncludedVal = Trim(cell.value)
    
    ' --- Get dependent cell dynamically ---
    Set quantityCell = GetDependentCell(cell, sheetName)
    GIWQuantityVal = Trim(quantityCell.value)
    
    ' --- Validation table ---
    On Error Resume Next
    Set validationTable = wsConfig.ListObjects("GIWValidationTable")
    On Error GoTo 0
    If validationTable Is Nothing Then
        DebugMessage ("Validation table in config not found 'GIWValidationTable'")
        RunGIWIncludedValidation = False: Exit Function
    End If
    
    expectedRule = ""
    For Each tblRow In validationTable.ListRows
        If StrComp(Trim(tblRow.Range(1, 1).value), GIWIncludedVal, vbTextCompare) = 0 Then
            expectedRule = Trim(tblRow.Range(1, 2).value)
            Exit For
        End If
    Next tblRow
    
    If expectedRule = "" Then
    
        If english = True Then
            Msg = "Error: Invalid entry"
        Else
            Msg = "Erreur: Entrée non valide."
        End If
        
        AddValidationFeedback CalledFuncName, ws, cell.row, Msg, "Error", english, FormatMap, AutoValMap
        
        RunGIWIncludedValidation = False: Exit Function
    End If
    
    ' --- Validate quantity value ---
    If GIWQuantityVal = "" Then
        If english = True Then
            Msg = "Error: Invalid entry, Cannot be empty"
        Else
            Msg = "Erreur: Entrée non valide. La ne peut pas être vide."
        End If
        
        AddValidationFeedback CalledFuncName, ws, cell.row, Msg, "Error", english, FormatMap, AutoValMap
        
        RunGIWIncludedValidation = False: Exit Function
    End If
    
    quantityParts = Split(GIWQuantityVal, ",")
    If UBound(quantityParts) <> 1 Then
    
        If english = True Then
            Msg = "Entry not valid, must be 'Number:Number'"
        Else
            Msg = "Entrée non valide, le format doit être 'Nombre:Nombre"
        End If
        
        AddValidationFeedback CalledFuncName, ws, cell.row, Msg, "Error", english, FormatMap, AutoValMap
        
        RunGIWIncludedValidation = False: Exit Function
    End If
    
    On Error Resume Next
    If quantityParts(0) <> "#" Then num1 = CLng(quantityParts(0)) Else num1 = "-1"
    If quantityParts(1) <> "#" Then num2 = CLng(quantityParts(1)) Else num2 = "-1"
    
    On Error GoTo 0
    
    Dim GIWQValueStr As String
    GIWQValueStr = "'" & num1 & "," & num2 & "'"
    
    ' --- Validate against expected rule ---
    isValid = False
    Select Case expectedRule
        Case "0": If num1 = 0 And num2 = 0 Then isValid = True
        Case "1": If num1 > 0 And num2 > 0 And num1 <= num2 Then isValid = True
        Case "#": If quantityParts(0) = "#" And quantityParts(1) = "#" Then isValid = True
    End Select
    
    If Not isValid Then
        Select Case expectedRule
            Case "0":
                Msg = IIf(english, "GIW Quantity must be 0,0 when GIW Included is 'No'.", "La quantité GIW doit être 0,0 lorsque GIW Inclus est 'Non'.")
                If num1 = -1 And num2 = -1 Then
                    Msg = IIf(english, "Automatic Correction: Changed entry #,# to 0,0", "Correction Automatique: #,# à 0,0")
                    AddValidationFeedback "GIWQuantity", ws, cell.row, Msg, "Autocorrect", english, FormatMap, AutoValMap
                    AddValidationFeedback "GIWIncluded", ws, cell.row, "", "Default", english, FormatMap, AutoValMap
                    quantityCell.value = 0 & "," & 0
                    RunGIWIncludedValidation = True
                    Exit Function
                ElseIf num1 > 0 Or num2 > 0 Then
                    Msg = IIf(english, "Invalid Entry, value must be 0,0 when GIW Included = 'No'", "Combinaison invalide: valeur doit être 0,0 quand Toilette Inclusive-Quantité = 'Non'")
                    AddValidationFeedback "GIWQuantity", ws, cell.row, Msg, "Error", english, FormatMap, AutoValMap
                    RunGIWIncludedValidation = False
                    Exit Function
                End If
            Case "1":
                Msg = IIf(english, "GIW Quantity must be positive when GIW Included is 'Yes' or 'Partially'.", "La quantité GIW doit être positive lorsque GIW Inclus est 'Oui' ou 'Partiellement'.")
                    If num1 > num2 Then
                        Msg = IIf(english, GIWQValueStr & " is an invlid entry, Number of Gender Inclusive Washrooms (" & num1 & ") cannot be greater than Number of Water Closets (" & num2 & ").", "Entrée invalide : le nombre de toilettes inclusives (" & num1 & ") ne peut excéder le nombre de cabinets de toilette (" & num2 & ").")
                        CalledFuncName = "GIWQuantity"
                End If
                AddValidationFeedback CalledFuncName, ws, cell.row, Msg, "Error", english, FormatMap, AutoValMap
                RunGIWIncludedValidation = False
                Exit Function
            
            Case "#": Msg = IIf(english, "GIW Quantity must be '#,#' when GIW Included is 'Not Applicable'.", "La quantité GIW doit être '#,#' lorsque GIW Inclus est 'Non applicable'.")
            Case Else: Msg = IIf(english, "Invalid combination of GIW Included and Quantity.", "Combinaison invalide de GIW Inclus et Quantité.")
        End Select
        
        AddValidationFeedback CalledFuncName, ws, cell.row, Msg, "Error", english, FormatMap, AutoValMap
        RunGIWIncludedValidation = False
    Else
        AddValidationFeedback CalledFuncName, ws, cell.row, Msg, "Default", english, FormatMap, AutoValMap
        AddValidationFeedback OtherFuncName, ws, cell.row, Msg, "Default", english, FormatMap, AutoValMap
        RunGIWIncludedValidation = True
    End If
End Function

Private Function GetDependentCell(cell As Range, sheetName As String) As Range
    Dim ws As Worksheet, wsConfig As Worksheet
    Dim RowNum As Long
    Dim GIWQuantityCol As String, GIWIncludedCol As String
    
    Set ws = ThisWorkbook.Sheets(sheetName)
    Set wsConfig = ThisWorkbook.Sheets("Config")
    RowNum = cell.row
    
    ' --- Resolve column letters from Config ---
    Dim i As Long
    i = 6
    GIWQuantityCol = "": GIWIncludedCol = ""
    Do While wsConfig.Range("B" & i).value <> ""
        If Trim(wsConfig.Range("C" & i).value) = "GIWQuantity" Then
            GIWQuantityCol = Trim(wsConfig.Range("B" & i).value)
        ElseIf Trim(wsConfig.Range("C" & i).value) = "GIWIncluded" Then
            GIWIncludedCol = Trim(wsConfig.Range("B" & i).value)
        End If
        i = i + 1
    Loop
    
    If cell.Column = ws.Range(GIWQuantityCol & "1").Column Then
        Set GetDependentCell = ws.Range(GIWIncludedCol & RowNum)
    Else
        Set GetDependentCell = ws.Range(GIWQuantityCol & RowNum)
    End If
End Function




