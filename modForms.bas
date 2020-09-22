Attribute VB_Name = "modForms"
'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : IsLoaded
' Params : 
'          strFormName As String
' Returns: Boolean
' Desc   : The Function uses parameters strFormName As String for IsLoaded and returns Boolean.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'		Auto-Code Commenter
'__________________________________________________
Public Function IsLoaded(strFormName As String) As Boolean
    On Error GoTo Proc_Err
    Const csProcName As String = "IsLoaded"
    Dim lCount As Long
    For lCount = 0 To Forms.Count - 1
        If (Forms(lCount).Name = strFormName) Then
            IsLoaded = True
            Exit For
        End If
    Next

Proc_Exit:
    GoSub Proc_Cleanup
    Exit Function

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here	
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbcrlf & "modForms->"  & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    Err.Raise Number:=lErrNum, Source:=sErrSource, Description:=sErrDesc
    Exit Function
    
End Function

