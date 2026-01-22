VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ValidationTrackerForm 
   Caption         =   "Full Validation Tracker"
   ClientHeight    =   6585
   ClientLeft      =   -135
   ClientTop       =   -750
   ClientWidth     =   7515
   OleObjectBlob   =   "ValidationTrackerForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ValidationTrackerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ======================================================
' ValidationTrackerForm v2.1
' Progress tracking and user feedback for validation
' UPDATED: Wired up Cancel button, standardized naming
' ======================================================

Option Explicit
Public IsInitialized As Boolean
Public UserLog As String

Private Sub UserForm_Initialize()
    On Error Resume Next
    
    ' Initialize checkboxes (disabled, unlocked by default)
    CheckBox1_AutoValInit.Enabled = False
    CheckBox1_AutoValInit.Locked = False
    CheckBox2_AdvValCompleted.Enabled = False
    CheckBox2_AdvValCompleted.Locked = False
    Checkbox3_LMenuValCompleted.Enabled = False
    Checkbox3_LMenuValCompleted.Locked = False
    
    ' Log the form open event
    Call LogFormUpdate("ValidationTrackerForm initialized")
    
    ' Normalize display for DPI consistency
    Me.Zoom = 100
    
    ' Set consistent form dimensions (optional - adjust as needed)
    Me.Width = 350
    Me.Height = 360
    
    ' Mark form initialization complete
    IsInitialized = True
End Sub

' ======================================================
' PUBLIC METHODS (Called by AV_UI)
' ======================================================

Public Function getFormStatus() As Boolean
    getFormStatus = IsInitialized
End Function

Public Sub setAutoValInitCB(ByVal Checked As Boolean)
    On Error Resume Next
    CheckBox1_AutoValInit.Enabled = True
    CheckBox1_AutoValInit.Locked = True
    
    If Checked Then
        CheckBox1_AutoValInit.value = True
    Else
        CheckBox1_AutoValInit.value = False
    End If
    
    CheckBox1_AutoValInit.Enabled = False
    CheckBox1_AutoValInit.Locked = False
End Sub

Public Sub setAdvValCompletedCB(ByVal Checked As Boolean)
    On Error Resume Next
    CheckBox2_AdvValCompleted.Enabled = True
    CheckBox2_AdvValCompleted.Locked = True
    
    If Checked Then
        CheckBox2_AdvValCompleted.value = True
    Else
        CheckBox2_AdvValCompleted.value = False
    End If
    
    CheckBox2_AdvValCompleted.Enabled = False
    CheckBox2_AdvValCompleted.Locked = False
End Sub

Public Sub setLMenuValCompletedCB(ByVal Checked As Boolean)
    On Error Resume Next
    Checkbox3_LMenuValCompleted.Enabled = True
    Checkbox3_LMenuValCompleted.Locked = True
    
    If Checked Then
        Checkbox3_LMenuValCompleted.value = True
    Else
        Checkbox3_LMenuValCompleted.value = False
    End If
    
    Checkbox3_LMenuValCompleted.Enabled = False
    Checkbox3_LMenuValCompleted.Locked = False
End Sub

' ======================================================
' EVENT HANDLERS
' ======================================================

Private Sub CancelValidationButton_Click()
    ' Call the AV_UI cancel function
    AV_UI.CancelValidation
    
    ' Provide visual feedback
    Me.Caption = "Full Validation Tracker - CANCELLED"
    Call LogFormUpdate("Validation cancelled by user")
End Sub

Private Sub CloseButton_Click()
    Unload Me
End Sub

Private Sub UserForm_Terminate()
    IsInitialized = False
End Sub

' ======================================================
' INTERNAL LOGGING
' ======================================================

Private Sub LogFormUpdate(ByVal message As String)
    ' TEMPORARY STUB - Centralized logging will be reintroduced later
    ' For now, just write to debug
    Debug.Print "[ValidationTrackerForm] " & message
End Sub

' ======================================================
' NOTES FOR FORM DESIGNER:
' ======================================================
' This form requires the following controls:
'
' 1. CheckBox1_AutoValInit (CheckBox)
'    - Caption: "Auto Validation Initialized"
'    - Used to show when validation config is loaded
'
' 2. CheckBox2_AdvValCompleted (CheckBox)
'    - Caption: "Advanced Validation Completed"
'    - Used to show when complex validations finish
'
' 3. Checkbox3_LMenuValCompleted (CheckBox)
'    - Caption: "Menu Validation Completed"
'    - Used to show when dropdown validations finish
'
' 4. FormUpdateLogListBox (ListBox)
'    - Used by AV_UI.AppendUserLog to display messages
'    - Should be multiline, scrollable
'
' 5. CancelValidationButton (CommandButton)
'    - Caption: "Cancel"
'    - Allows user to stop validation mid-process
'
' 6. CloseButton (CommandButton)
'    - Caption: "Close"
'    - Closes the form
' ======================================================
