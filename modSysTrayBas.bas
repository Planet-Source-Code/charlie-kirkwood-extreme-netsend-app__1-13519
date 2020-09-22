Attribute VB_Name = "modSysTrayBas"

Public nid As NOTIFYICONDATA
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_USER = &H400
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
    End Type
 
    Public Const NIM_ADD = &H0
    Public Const NIM_MODIFY = &H1
    Public Const NIM_DELETE = &H2
    Public Const NIF_MESSAGE = &H1
    Public Const NIF_ICON = &H2
    Public Const NIF_TIP = &H4
    Public Const WM_MOUSEMOVE = &H200

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

'__________________________________________________
' Scope  : 
' Type   : Sub
' Name   : InitializeTrayIcon
' Params : 
' Returns: Nothing
' Desc   : The Sub uses parameters  for InitializeTrayIcon and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'		Auto-Code Commenter
'__________________________________________________
Sub InitializeTrayIcon()
    On Error GoTo Proc_Err
    Const csProcName As String = "InitializeTrayIcon"

      With nid
        .cbSize = Len(nid)
        .hwnd = frmNetSender.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .szTip = App.Title & vbNullChar
        .hIcon = frmNetSender.Icon
    End With
    
    Shell_NotifyIcon NIM_ADD, nid

Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here	
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbcrlf & "modSysTrayBas->"  & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    Err.Raise Number:=lErrNum, Source:=sErrSource, Description:=sErrDesc
    Exit Sub
    
End Sub

