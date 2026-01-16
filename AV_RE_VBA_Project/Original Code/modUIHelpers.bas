Attribute VB_Name = "modUIHelpers"
Option Explicit

#If VBA7 Then
    ' --- For 64-bit or newer VBA environments ---
    Private Declare PtrSafe Function FindWindowA Lib "user32" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

    Private Declare PtrSafe Function SetWindowPos Lib "user32" _
        (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, _
         ByVal x As Long, ByVal y As Long, _
         ByVal cx As Long, ByVal cy As Long, _
         ByVal wFlags As Long) As Long
#Else
    ' --- For legacy 32-bit VBA environments ---
    Private Declare Function FindWindowA Lib "user32" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

    Private Declare Function SetWindowPos Lib "user32" _
        (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
         ByVal x As Long, ByVal y As Long, _
         ByVal cx As Long, ByVal cy As Long, _
         ByVal wFlags As Long) As Long
#End If

Private Const HWND_TOPMOST As Long = -1
Private Const HWND_NOTOPMOST As Long = -2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_SHOWWINDOW As Long = &H40

' === Brings a UserForm to the front and makes it topmost ===
Public Sub BringFormToFront(frm As Object)
    On Error Resume Next
    Dim hWndForm As LongPtr

    hWndForm = FindWindowA(vbNullString, frm.Caption)

    If hWndForm <> 0 Then
        ' Bring form to top without affecting size or position
        SetWindowPos hWndForm, HWND_TOPMOST, 0, 0, 0, 0, _
                     SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
    Else
        Debug.Print "[BringFormToFront] Could not find window handle for: " & frm.Caption
    End If
End Sub


