VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ValidationTrackerForm 
   Caption         =   "Full Validation Tracker"
   ClientHeight    =   4460
   ClientLeft      =   -140
   ClientTop       =   -750
   ClientWidth     =   7070
   OleObjectBlob   =   "ValidationTrackerForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ValidationTrackerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public IsInitialized As Boolean
Public UserLog As String  ' If you haven't already

Private Sub UserForm_Initialize()
    On Error Resume Next
    CheckBox1_AutoValInit.Enabled = False
    CheckBox1_AutoValInit.Locked = False
    CheckBox2_AdvValCompleted.Enabled = False
    CheckBox2_AdvValCompleted.Locked = False
    Checkbox3_LMenuValCompleted.Enabled = False
    Checkbox3_LMenuValCompleted.Enabled = False
    
    ' Log the form open event (writes to hidden sheet + updates FormUpdateLogListBox)
    Call LogFormUpdate(Me.Name)
    
    ' Normalize display for DPI consistency
    Me.Zoom = 100
    
    ' Force consistent pixel size (optional, depends on form design)
    ' You can adjust these if each form uses its own dimensions
    Me.Width = 350
    Me.Height = 360
    
    ' Optional: mark form initialization complete
    IsInitialized = True
End Sub

Public Sub setAutoValInitCB(ByVal Checked As Boolean)
    CheckBox1_AutoValInit.Enabled = True
    CheckBox1_AutoValInit.Locked = True

If Checked Then
    CheckBox1_AutoValInit.Value = True
Else
    CheckBox1_AutoValInit.Value = False
End If
    CheckBox1_AutoValInit.Enabled = False
    CheckBox1_AutoValInit.Locked = False
End Sub

Public Sub setAdvValCompletedCB(ByVal Checked As Boolean)
    CheckBox2_AdvValCompleted.Enabled = True
    CheckBox2_AdvValCompleted.Locked = True

If Checked Then
    CheckBox2_AdvValCompleted.Value = True
Else
    CheckBox2_AdvValCompleted.Value = False
End If
    CheckBox2_AdvValCompleted.Enabled = False
    CheckBox2_AdvValCompleted.Locked = False
End Sub

Public Sub setLMenuValCompletedCB(ByVal Checked As Boolean)
    Checkbox3_LMenuValCompleted.Enabled = True
    Checkbox3_LMenuValCompleted.Locked = True

If Checked Then
    Checkbox3_LMenuValCompleted.Value = True
Else
    Checkbox3_LMenuValCompleted.Value = False
End If
    Checkbox3_LMenuValCompleted.Enabled = False
    Checkbox3_LMenuValCompleted.Locked = False
End Sub

Private Sub UserForm_Terminate()
    IsInitialized = False
End Sub

Public Function getFormStatus() As Boolean
    getFormStatus = IsInitialized
End Function

Private Sub CancelValidationButton_Click()
    Call CancelValidation
End Sub

Private Sub CloseButton_Click()
    Unload Me
End Sub

' ======================================================
'  FORM-LOCAL STUB – LOG FORM UPDATE
' ======================================================
Private Sub LogFormUpdate(ByVal message As String)
    ' TEMPORARY STUB
    ' Centralized logging will be reintroduced later.
    Debug.Print "[ValidationTrackerForm] " & message
End Sub



