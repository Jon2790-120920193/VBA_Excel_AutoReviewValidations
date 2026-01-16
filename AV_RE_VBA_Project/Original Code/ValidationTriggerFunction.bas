Attribute VB_Name = "ValidationTriggerFunction"
Option Explicit


' === MAIN TRIGGER CALLED FROM SHEET2 (B - Buildings - Bâtiments) ===
Sub SheetValidationTrigger(Target As Range, Optional sheetName As String = "", Optional english As Boolean = True)
    
    If sheetName = "" Then sheetName = Target.Worksheet.Name
    
    Dim wsConfig As Worksheet
    Dim wsTarget As Worksheet
    Dim dataSheetName As String
    Dim keyColLetter As String
    Dim keyColNum As Long
    Dim startRow As Long, endRow As Long
    Dim validateColMap As Object
    Dim TargetColLetter As String
    Dim cellRow As Long, cellCol As Long
    Dim langControl As String
    
    Set wsConfig = ThisWorkbook.Sheets("Config")
    dataSheetName = Trim(wsConfig.Range("B3").value)
    
    ' Sets Language based on User Language Control set
    langControl = Trim(wsConfig.Range("M1").value)
    If langControl = "English" Then
        english = True
    ElseIf langControl = "Français" Then
        english = False
    Else
        Debug.Print "[ValidationTrigger] Invalid language selection in Config!M1: '" & langControl & "'. Defaulting to English."
        english = True ' fallback to English
    End If

    ' Validate data sheet exists
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets(dataSheetName)
    On Error GoTo 0

    If wsTarget Is Nothing Then
        MsgBox "Error: Sheet '" & dataSheetName & "' not found. Check Config!B3.", vbCritical
        Exit Sub
    End If

    ' ? Check that the change is on the correct sheet
    If Target.Worksheet.Name <> wsTarget.Name Then Exit Sub

    ' Get row/col info
    cellRow = Target.row
    cellCol = Target.Column

    ' Get column letter of changed cell
    TargetColLetter = Split(wsTarget.Cells(1, cellCol).Address(0, 0), "1")(0)

    ' Load Config values
    startRow = wsConfig.Range("B4").value
    endRow = startRow + wsConfig.Range("D4").value
    keyColLetter = wsConfig.Range("B5").value
    keyColNum = wsTarget.Range(keyColLetter & "1").Column

    ' Check row within range
    If cellRow < startRow Or cellRow > endRow Then Exit Sub

    ' Check if the key column has a value in this row
    If Trim(wsTarget.Cells(cellRow, keyColNum).value) = "" Then Exit Sub
    
    ' === EARLY VALIDATION CONDITION BASED ON ForceValidationTable ===
    If Not ShouldValidateRow(cellRow, wsTarget) Then
        Debug.Print "[Validation Skipped] Row " & cellRow & " does not meet ForceValidationTable criteria."
        Exit Sub
    End If

    ' Load column-function map
    Set validateColMap = GetValidationColumns(wsConfig)

    ' If column is mapped, run the corresponding validation function
    If validateColMap.Exists(TargetColLetter) Then
        Dim funcName As String
        funcName = "Validate_Column_" & validateColMap(TargetColLetter)
    
        Debug.Print "Attempting to run: " & funcName
        Debug.Print "Running with parameters: " & Target.Address & ", " & wsTarget.Name & ", English=" & english
    
        On Error GoTo ValidationError
        
        Dim FormatMapping As Object
        Set FormatMapping = LoadFormatMap(wsConfig)
        
        Dim AdvFunctionMap As Object
        Set AdvFunctionMap = GetAutoValidationMap(wsConfig)

        Application.Run funcName, Target, wsTarget.Name, english, FormatMapping, AdvFunctionMap
        
        Exit Sub

ValidationError:
    Debug.Print "[ValidationError] Function '" & funcName & "' failed. Error: " & Err.Number & " - " & Err.Description
    Debug.Print "Error running validation function: " & funcName & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation
    Resume Next
    End If
    
End Sub











