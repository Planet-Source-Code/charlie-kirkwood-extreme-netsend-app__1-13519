Attribute VB_Name = "modCurrentSystemInfo"
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function getUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'__________________________________________________
' Scope  : Public
' Type   : Property Get
' Name   : GetUserNamePv
' Params : 
' Returns: Variant
' Desc   : The Property Get uses parameters  for GetUserNamePv and returns Variant.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'		Auto-Code Commenter
'__________________________________________________
Public Property Get GetUserNamePv() As Variant
    On Error GoTo Proc_Err
    Const csProcName As String = "GetUserNamePv"
     Dim sBuffer As String
     Dim lSize As Long
     Dim stemp As String
     
     sBuffer = Space$(255)
     lSize = Len(sBuffer)
     Call getUserName(sBuffer, lSize)
     
     stemp = Left$(sBuffer, lSize)
     GetUserNamePv = Replace(stemp, vbNullChar, "")
     

Proc_Exit:
    GoSub Proc_Cleanup
    Exit Property

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here	
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbcrlf & "modCurrentSystemInfo->"  & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    Err.Raise Number:=lErrNum, Source:=sErrSource, Description:=sErrDesc
    Exit Property
    
End Property

'__________________________________________________
' Scope  : Public
' Type   : Property Get
' Name   : ThreadIdPv
' Params : 
' Returns: Variant
' Desc   : The Property Get uses parameters  for ThreadIdPv and returns Variant.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'		Auto-Code Commenter
'__________________________________________________
Public Property Get ThreadIdPv() As Variant
    On Error GoTo Proc_Err
    Const csProcName As String = "ThreadIdPv"
    ThreadIdPv = GetCurrentThreadId

Proc_Exit:
    GoSub Proc_Cleanup
    Exit Property

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here	
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbcrlf & "modCurrentSystemInfo->"  & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    Err.Raise Number:=lErrNum, Source:=sErrSource, Description:=sErrDesc
    Exit Property
    
End Property

'__________________________________________________
' Scope  : Public
' Type   : Property Get
' Name   : ProcessIdPv
' Params : 
' Returns: Variant
' Desc   : The Property Get uses parameters  for ProcessIdPv and returns Variant.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'		Auto-Code Commenter
'__________________________________________________
Public Property Get ProcessIdPv() As Variant
    On Error GoTo Proc_Err
    Const csProcName As String = "ProcessIdPv"
    ProcessIdPv = GetCurrentProcessId

Proc_Exit:
    GoSub Proc_Cleanup
    Exit Property

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here	
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbcrlf & "modCurrentSystemInfo->"  & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    Err.Raise Number:=lErrNum, Source:=sErrSource, Description:=sErrDesc
    Exit Property
    
End Property



