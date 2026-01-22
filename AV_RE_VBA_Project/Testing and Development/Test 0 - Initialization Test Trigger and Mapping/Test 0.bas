Attribute VB_Name = "Test"
Public Sub Test1()
Dim errMsg As String
Debug.Print AV_Core.ValidateConfiguration(errMsg)
'Result returned FALSE
End Sub

Public Sub Test2()
Dim config As AV_Core.ValidationConfig
config = AV_Core.LoadValidationConfig()
Debug.Print config.TargetCount
'Returned 3
End Sub

Public Sub Test3()
Dim tbl As ListObject
Set tbl = AV_Core.GetValidationTable(AV_Constants.TBL_GIW_VALIDATION)
Debug.Print tbl.Name
'Returned GIWValidationTable
End Sub

' No Trigger Function defined, will need to get this resolved, ref original code.



