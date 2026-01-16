Attribute VB_Name = "AV_UI"
Option Explicit

' ======================================================
' AV_UI v2.1
' UI helpers for ValidationTrackerForm
' No validation logic here
' UPDATED: Added MODULE_NAME constant for consistency
' ======================================================

Private Const MODULE_NAME As String = "AV_UI"

' -----------------------------
' DEBUG MODE (optional logging)
' -----------------------------
Public Const DEBUG_MODE As Boolean = True

' -----------------------------
' FORM HELPERS
' -----------------------------

Public Sub ShowValidationTrackerForm()
    On Error Resume Next

    If ValidationTrackerForm Is Nothing Then
        Load ValidationTrackerForm
    End If

    ValidationTrackerForm.Show vbModeless

    If DEBUG_MODE Then
        AV_Core.DebugMessage "ValidationTrackerForm shown", MODULE_NAME
    End If

    On Error GoTo 0
End Sub


Public Sub BringFormToFront(frm As Object)
    On Error Resume Next

    If frm Is Nothing Then Exit Sub
    If Not frm.Visible Then frm.Show vbModeless

    frm.ZOrder 0

    If DEBUG_MODE Then
        AV_Core.DebugMessage "BringFormToFront: " & TypeName(frm), MODULE_NAME
    End If

    On Error GoTo 0
End Sub


' -----------------------------
' USER LOG (FORM LISTBOX UPDATE)
' -----------------------------

Public Sub AppendUserLog(ByVal Msg As String, Optional ByVal includeTimestamp As Boolean = True)
    On Error GoTo SafeExit

    ' Check if form is loaded and initialized before writing
    If Not IsUserFormLoaded("ValidationTrackerForm") Then Exit Sub

    ' Double-check with form's own status (in case it was loaded but not fully initialized)
    If Not ValidationTrackerForm.getFormStatus() Then Exit Sub

    ' Now safe to write
    Dim entry As String
    If includeTimestamp Then
        entry = "[" & Format(Now, "hh:mm:ss") & "] " & Msg
    Else
        entry = Msg
    End If

    With ValidationTrackerForm.FormUpdateLogListBox
        .AddItem entry
        .ListIndex = .ListCount - 1   ' Scroll to last item
    End With

SafeExit:
    Exit Sub
End Sub


Public Function IsUserFormLoaded(ByVal formName As String) As Boolean
    Dim frm As Object
    On Error Resume Next
    For Each frm In VBA.UserForms
        If StrComp(frm.Name, formName, vbTextCompare) = 0 Then
            IsUserFormLoaded = True
            Exit Function
        End If
    Next frm
    On Error GoTo 0
End Function


' -----------------------------
' VALIDATION STATE UPDATES
' -----------------------------

Public Sub SetAutoValidationInitialized(ByVal isComplete As Boolean)
    On Error Resume Next

    If ValidationTrackerForm Is Nothing Then Exit Sub
    ValidationTrackerForm.setAutoValInitCB isComplete

    If DEBUG_MODE Then
        AV_Core.DebugMessage "AutoValidation Init = " & isComplete, MODULE_NAME
    End If

    On Error GoTo 0
End Sub


Public Sub SetAdvancedValidationCompleted(ByVal isComplete As Boolean)
    On Error Resume Next

    If ValidationTrackerForm Is Nothing Then Exit Sub
    ValidationTrackerForm.setAdvValCompletedCB isComplete

    If DEBUG_MODE Then
        AV_Core.DebugMessage "Advanced Validation Complete = " & isComplete, MODULE_NAME
    End If

    On Error GoTo 0
End Sub


Public Sub SetLegacyMenuValidationCompleted(ByVal isComplete As Boolean)
    On Error Resume Next

    If ValidationTrackerForm Is Nothing Then Exit Sub
    ValidationTrackerForm.setLMenuValCompletedCB isComplete

    If DEBUG_MODE Then
        AV_Core.DebugMessage "Legacy Menu Validation Complete = " & isComplete, MODULE_NAME
    End If

    On Error GoTo 0
End Sub


' -----------------------------
' CANCEL HANDLING
' -----------------------------

Public Sub CancelValidation()
    On Error Resume Next

    AV_Core.ValidationCancelFlag = True

    If DEBUG_MODE Then
        AV_Core.DebugMessage "Validation cancelled by user", MODULE_NAME
    End If

    On Error GoTo 0
End Sub
