Attribute VB_Name = "modNetServer"
Option Explicit

Public Declare Function NetServerEnum Lib "Netapi32.dll" (vServername As Any, ByVal lLevel As Long, vBufptr As Any, lPrefmaxlen As Long, lEntriesRead As Long, lTotalEntries As Long, vServerType As Any, ByVal sDomain As String, vResumeHandle As Any) As Long
Public Declare Sub RtlMoveMemory Lib "kernel32" (dest As Any, vSrc As Any, ByVal lSize&)
Public Declare Sub lstrcpyW Lib "kernel32" (vDest As Any, ByVal sSrc As Any)
Declare Sub lstrcpy Lib "kernel32" (vDest As Any, ByVal vSrc As Any)
Declare Sub lstrcpynW Lib "kernel32" (ByVal vDest As Any, ByVal vSrc As Any, lLength As Long)
Public Declare Function NetApiBufferFree Lib "Netapi32.dll" (ByVal lpBuffer As Long) As Long

Declare Function NetWkstaGetInfo Lib "Netapi32.dll" (ByVal sServerName$, ByVal lLevel&, vBuffer As Any) As Long
'Public Declare Function NetWkstaGetInfo Lib "Netapi32.dll" (lpServer As Any, ByVal Level As Long, lpBuffer As Any) As Long

Declare Function NetMessageBufferSend Lib "Netapi32.dll" (ByVal sServerName$, ByVal sMsgName$, ByVal sFromName$, ByVal sMessageText$, ByVal lBufferLength&) As Long


'try these
Public Declare Function NetLocalGroupGetMembers Lib "Netapi32.dll" (ByVal psServer As Long, ByVal psLocalGroup As Long, ByVal lLevel As Long, pBuffer As Long, ByVal lMaxLength As Long, plEntriesRead As Long, plTotalEntries As Long, phResume As Long) As Long
Public Declare Function NetUserGetGroups Lib "Netapi32.dll" (ByVal sServerName$, UserName As Byte, ByVal Level As Long, lpBuffer As Long, ByVal PrefMaxLen As Long, lpEntriesRead As Long, lpTotalEntries As Long) As Long
Public Declare Function NetUserGetInfo Lib "Netapi32.dll" (ByVal sServerName$, UserName As Byte, ByVal Level As Long, lpBuffer As Long) As Long
Public Declare Function NetWkstaUserGetInfo Lib "Netapi32.dll" (ByVal reserved As Any, ByVal Level As Long, lpBuffer As Any) As Long




Type SERVER_INFO_100
    sv100_platform_id As Long
    sv100_servername As Long
End Type

Public Type SERVER_INFO_101
    dw_platform_id As Long
    ptr_name As Long
    dw_ver_major As Long
    dw_ver_minor As Long
    dw_type As Long
    ptr_comment As Long
End Type

Type WKSTA_INFO_100
    wki100_platform_id As Long
    wki100_computername As Long
    wki100_langroup As Long
    wki100_ver_major As Long
    wki100_ver_minor As Long
End Type

Public Enum eServerTypes
    SV_TYPE_WORKSTATION = &H1
    SV_TYPE_SERVER = &H2
    SV_TYPE_SQLSERVER = &H4
    SV_TYPE_DOMAIN_CTRL = &H8
    SV_TYPE_DOMAIN_BAKCTRL = &H10
    SV_TYPE_TIMESOURCE = &H20
    SV_TYPE_AFP = &H40
    SV_TYPE_NOVELL = &H80
    SV_TYPE_DOMAIN_MEMBER = &H100
    SV_TYPE_LOCAL_LIST_ONLY = &H40000000
    SV_TYPE_PRINT = &H200
    SV_TYPE_DIALIN = &H400
    SV_TYPE_XENIX_SERVER = &H800
    SV_TYPE_MFPN = &H4000
    SV_TYPE_NT = &H1000
    SV_TYPE_WFW = &H2000
    SV_TYPE_SERVER_NT = &H8000
    SV_TYPE_POTENTIAL_BROWSER = &H10000
    SV_TYPE_BACKUP_BROWSER = &H20000
    SV_TYPE_MASTER_BROWSER = &H40000
    SV_TYPE_DOMAIN_MASTER = &H80000
    SV_TYPE_DOMAIN_ENUM = &H80000000
    SV_TYPE_WINDOWS = &H400000
    SV_TYPE_ALL = &HFFFFFFFF

End Enum

Private Const mcsTempFile As String = "~tempUserList"

'Public Const SV_TYPE_WORKSTATION = &H1
'Public Const SV_TYPE_SERVER = &H2
'Public Const SV_TYPE_SQLSERVER = &H4
'Public Const SV_TYPE_DOMAIN_CTRL = &H8
'Public Const SV_TYPE_DOMAIN_BAKCTRL = &H10
'Public Const SV_TYPE_TIMESOURCE = &H20
'Public Const SV_TYPE_AFP = &H40
'Public Const SV_TYPE_NOVELL = &H80
'Public Const SV_TYPE_DOMAIN_MEMBER = &H100
'Public Const SV_TYPE_LOCAL_LIST_ONLY = &H40000000
'Public Const SV_TYPE_PRINT = &H200
'Public Const SV_TYPE_DIALIN = &H400
'Public Const SV_TYPE_XENIX_SERVER = &H800
'Public Const SV_TYPE_MFPN = &H4000
'Public Const SV_TYPE_NT = &H1000
'Public Const SV_TYPE_WFW = &H2000
'Public Const SV_TYPE_SERVER_NT = &H8000
'Public Const SV_TYPE_POTENTIAL_BROWSER = &H10000
'Public Const SV_TYPE_BACKUP_BROWSER = &H20000
'Public Const SV_TYPE_MASTER_BROWSER = &H40000
'Public Const SV_TYPE_DOMAIN_MASTER = &H80000
'Public Const SV_TYPE_DOMAIN_ENUM = &H80000000
'Public Const SV_TYPE_WINDOWS = &H400000
'Public Const SV_TYPE_ALL = &HFFFFFFFF
'

'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : GetLocalSystemName
' Params : 
' Returns: Nothing
' Desc   : The Function uses parameters  for GetLocalSystemName and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'		Auto-Code Commenter
'__________________________________________________
Public Function GetLocalSystemName()
    On Error GoTo Proc_Err
    Const csProcName As String = "GetLocalSystemName"
    Dim lReturnCode As Long
    Dim bBuffer(512) As Byte
    Dim I As Integer
    Dim twkstaInfo100 As WKSTA_INFO_100, lwkstaInfo100 As Long
    Dim lwkstaInfo100StructPtr As Long
    Dim sLocalName As String
    
    lReturnCode = NetWkstaGetInfo("", 100, lwkstaInfo100)
 
    lwkstaInfo100StructPtr = lwkstaInfo100
                 
    If lReturnCode = 0 Then
                 
        RtlMoveMemory twkstaInfo100, ByVal lwkstaInfo100StructPtr, Len(twkstaInfo100)
         
        lstrcpyW bBuffer(0), twkstaInfo100.wki100_computername

        I = 0
        Do While bBuffer(I) <> 0
            sLocalName = sLocalName & Chr(bBuffer(I))
            I = I + 2
        Loop
            
        GetLocalSystemName = sLocalName
         
    End If


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
    sErrSource = VBA.Err.Source & vbcrlf & "modNetServer->"  & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    Err.Raise Number:=lErrNum, Source:=sErrSource, Description:=sErrDesc
    Exit Function
    
End Function

'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : GetDomainName
' Params : 
' Returns: String
' Desc   : The Function uses parameters  for GetDomainName and returns String.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'		Auto-Code Commenter
'__________________________________________________
Public Function GetDomainName() As String
    On Error GoTo Proc_Err
    Const csProcName As String = "GetDomainName"
    
    Dim lReturnCode As Long
    Dim bBuffer(512) As Byte
    Dim I As Integer
    Dim twkstaInfo100 As WKSTA_INFO_100, lwkstaInfo100 As Long
    Dim lwkstaInfo100StructPtr As Long
    Dim sDomainName As String
    
    lReturnCode = NetWkstaGetInfo("", 100, lwkstaInfo100)
 
    lwkstaInfo100StructPtr = lwkstaInfo100
                 
    If lReturnCode = 0 Then
                 
        RtlMoveMemory twkstaInfo100, ByVal lwkstaInfo100StructPtr, Len(twkstaInfo100)
         
        lstrcpyW bBuffer(0), twkstaInfo100.wki100_langroup
        
        
        I = 0
        Do While bBuffer(I) <> 0
            sDomainName = sDomainName & Chr(bBuffer(I))
            I = I + 2
        Loop
            
        GetDomainName = sDomainName
         
    End If
        

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
    sErrSource = VBA.Err.Source & vbcrlf & "modNetServer->"  & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    Err.Raise Number:=lErrNum, Source:=sErrSource, Description:=sErrDesc
    Exit Function
    
End Function



