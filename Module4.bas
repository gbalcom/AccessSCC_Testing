Attribute VB_Name = "Module4"
Option Compare Database
Option Explicit

'@Folder("Access_SCCTesting")
Private Function MyNewFunction() As String
    ' Comments:
    ' Params  :
    ' Returns : String
    ' Created : 11/05/17 18:03 GB
    ' Modified:
    
    'TVCodeTools ErrorEnablerStart
    On Error GoTo PROC_ERR
    'TVCodeTools ErrorEnablerEnd

    MyNewFunction = "Testing"

    'TVCodeTools ErrorHandlerStart
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox Err.Description, vbCritical, "Module4.MyNewFunction"
    Resume PROC_EXIT
    'TVCodeTools ErrorHandlerEnd

End Function
