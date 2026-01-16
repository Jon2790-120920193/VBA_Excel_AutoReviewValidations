Attribute VB_Name = "ValidationTrackerFormControls"
Option Explicit

' Toggle debug mode for optional printouts to Immediate window (Ctrl+G)
Public Const DEBUG_MODE As Boolean = True

' --- Checks if the form is currently loaded ---
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

' === Safely append a message to the ValidationTrackerForm ListBox ===
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

ErrHandler:
    Debug.Print "[AppendUserLog ERROR] " & Err.Number & " - " & Err.Description
    Resume SafeExit
End Sub

' Show the ValidationTrackerForm modelessly (safe)
Public Sub ShowValidationTrackerForm()
    On Error Resume Next
    If ValidationTrackerForm Is Nothing Then
        Load ValidationTrackerForm
    End If
    ValidationTrackerForm.Show vbModeless
    ' Access modUIHelpers
    BringFormToFront ValidationTrackerForm
End Sub


' Close/unload the form safely
Public Sub CloseValidationTrackerForm()
    On Error Resume Next
    If IsUserFormLoaded("ValidationTrackerForm") Then
        Unload ValidationTrackerForm
        Debug.Print "[ValidationTrackerForm] closed."
    End If
    On Error GoTo 0
End Sub





