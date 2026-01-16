Attribute VB_Name = "modUserFormUtils"
Public Sub NormalizeAllForms()
    Dim frm As Object
    For Each frm In VBA.UserForms
        frm.Zoom = 100
    Next frm
End Sub

Public Sub LogFormUpdate(formName As String, Optional action As String = "Opened")
    Dim ws As Worksheet
    Dim NextRow As Long
    Dim logMsg As String
    
    ' Write to a hidden sheet for long-term tracking (optional)
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("FormUpdateLog")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "FormUpdateLog"
        ws.Visible = xlSheetVeryHidden
        ws.Range("A1:B1").value = Array("Timestamp", "Event")
    End If
    
    NextRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
    logMsg = "[" & formName & "] " & action
    ws.Cells(NextRow, 1).value = Now
    ws.Cells(NextRow, 2).value = logMsg
    
    ' Push to FormUpdateLogListBox if open
    Dim frm As Object
    For Each frm In VBA.UserForms
        On Error Resume Next
        frm.FormUpdateLogListBox.AddItem Format(Now, "hh:mm:ss") & " - " & logMsg
        On Error GoTo 0
    Next frm
End Sub

