Attribute VB_Name = "AV_UI"
Option Explicit

' ======================================================
' AV_UI
' UI helpers for ValidationTrackerForm
' No validation logic here
' ======================================================

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

    If DEBUG_MODE Then Debug.Print "[AV_UI] ValidationTrackerForm shown"

    On Error GoTo 0
End Sub


Public Sub BringFormToFront(frm As Object)
    On Error Resume Next

    If frm Is Nothing Then Exit Sub
    If Not frm.Visible Then frm.Show vbModeless

    frm.ZOrder 0

    If DEBUG_MODE Then Debug.Print "[AV_UI] BringFormToFront: " & TypeName(frm)

    On Error GoTo 0
End Sub


' -----------------------------
' VALIDATION STATE UPDATES
' -----------------------------

Public Sub SetAutoValidationInitialized(ByVal isComplete As Boolean)
    On Error Resume Next

    If ValidationTrackerForm Is Nothing Then Exit Sub
    ValidationTrackerForm.setAutoValInitCB isComplete

    If DEBUG_MODE Then Debug.Print "[AV_UI] AutoValidation Init = " & isComplete

    On Error GoTo 0
End Sub


Public Sub SetAdvancedValidationCompleted(ByVal isComplete As Boolean)
    On Error Resume Next

    If ValidationTrackerForm Is Nothing Then Exit Sub
    ValidationTrackerForm.setAdvValCompletedCB isComplete

    If DEBUG_MODE Then Debug.Print "[AV_UI] Advanced Validation Complete = " & isComplete

    On Error GoTo 0
End Sub


Public Sub SetLegacyMenuValidationCompleted(ByVal isComplete As Boolean)
    On Error Resume Next

    If ValidationTrackerForm Is Nothing Then Exit Sub
    ValidationTrackerForm.setLegacyMenuCompletedCB isComplete

    If DEBUG_MODE Then Debug.Print "[AV_UI] Legacy Menu Validation Complete = " & isComplete

    On Error GoTo 0
End Sub


' -----------------------------
' CANCEL HANDLING
' -----------------------------

Public Sub CancelValidation()
    On Error Resume Next

    ValidationCancelFlag = True

    If DEBUG_MODE Then Debug.Print "[AV_UI] Validation cancelled by user"

    On Error GoTo 0
End Sub

