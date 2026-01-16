Attribute VB_Name = "GlobalDebugFunction"
' --- GlobalDebugOptions.bas ---
Option Explicit

Public DebugFlags As Object       ' Dictionary: ModuleName -> Boolean
Public GlobalDebugOn As Boolean   ' True if global debug = YES
Private DebugInitialized As Boolean

' Initialize debug flags (once per session or per run)
Public Sub InitDebugFlags(Optional forceReload As Boolean = False)
    Dim wsConfig As Worksheet
    Dim tbl As ListObject
    Dim lr As ListRow
    Dim moduleName As String, enabledVal As String
    
    ' Only initialize once unless forced
    If DebugInitialized And Not forceReload Then Exit Sub
    
    Set DebugFlags = CreateObject("Scripting.Dictionary")
    GlobalDebugOn = False
    DebugInitialized = False
    
    On Error GoTo ErrHandler
    
    Set wsConfig = ThisWorkbook.Sheets("Config")
    
    ' Read global debug option first
    On Error Resume Next
    GlobalDebugOn = (UCase(Trim(wsConfig.ListObjects("GlobalDebugOptions").DataBodyRange(1, 1).value)) = "YES")
    On Error GoTo 0
    
    ' If global debug is off, read per-module debug flags
    If Not GlobalDebugOn Then
        On Error Resume Next
        Set tbl = wsConfig.ListObjects("DebugControls")
        On Error GoTo 0
        
        If Not tbl Is Nothing Then
            For Each lr In tbl.ListRows
                moduleName = Trim(lr.Range.Cells(1, 1).value)
                enabledVal = Trim(lr.Range.Cells(1, 2).value)
                
                If moduleName <> "" Then DebugFlags(moduleName) = (UCase(enabledVal) = "YES")
            Next lr
        End If
    End If
    
    DebugInitialized = True
    Exit Sub
    
ErrHandler:
    Debug.Print "[InitDebugFlags ERROR] " & Err.Number & " - " & Err.Description
End Sub

' --- Unified Debug Print Function ---
Public Sub DebugMessage(Msg As String, Optional moduleName As String = "", Optional DebugON As Boolean = False)
    ' Ensure flags are initialized
If Not DebugInitialized Then InitDebugFlags
    
    ' Global debug overrides all
    If GlobalDebugOn Then
        Debug.Print Msg
        Exit Sub
    End If
    
    ' Module-specific check
    If moduleName <> "" Then
        If Not DebugFlags Is Nothing Then
            If DebugFlags.Exists(moduleName) Then
                If DebugFlags(moduleName) Then Debug.Print Msg
            End If
        End If
    Else
        ' Default: use manual debugON argument
        If DebugON Then Debug.Print Msg
    End If
End Sub

Sub ClearImmediateWindow()
    Application.SendKeys "^g ^a {DEL}"
End Sub


