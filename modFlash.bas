Attribute VB_Name = "modFlash"
Option Explicit

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000

'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : isTransparent
' Params : 
'          ByVal hwnd As Long
' Returns: Boolean
' Desc   : The Function uses parameters ByVal hwnd As Long for isTransparent and returns Boolean.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'		Auto-Code Commenter
'__________________________________________________
Public Function isTransparent(ByVal hwnd As Long) As Boolean


On Error Resume Next
Dim msg As Long

msg = GetWindowLong(hwnd, GWL_EXSTYLE)

If (msg And WS_EX_LAYERED) = WS_EX_LAYERED Then
    isTransparent = True
Else
    isTransparent = False
End If

If Err Then
    isTransparent = False
End If

End Function

'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : EnableTransparanty
' Params : 
'          ByVal hwnd As Long
'          Perc As Integer
' Returns: Long
' Desc   : The Function uses parameters ByVal hwnd As Long and Perc As Integer for EnableTransparanty and returns Long.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'		Auto-Code Commenter
'__________________________________________________
Public Function EnableTransparanty(ByVal hwnd As Long, Perc As Integer) As Long

    Dim msg As Long
    
    On Error Resume Next
    
    If Perc < 0 Or Perc > 255 Then
        EnableTransparanty = 1
    Else
    
    
        msg = GetWindowLong(hwnd, GWL_EXSTYLE)
        
        msg = msg Or WS_EX_LAYERED
        
        SetWindowLong hwnd, GWL_EXSTYLE, msg
        
        SetLayeredWindowAttributes hwnd, 0, Perc, LWA_ALPHA
        
        EnableTransparanty = 0
        End If
    
    If Err Then
        EnableTransparanty = 2
    End If
        
End Function

'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : DisableTransparanty
' Params : 
'          ByVal hwnd As Long
' Returns: Long
' Desc   : The Function uses parameters ByVal hwnd As Long for DisableTransparanty and returns Long.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'		Auto-Code Commenter
'__________________________________________________
Public Function DisableTransparanty(ByVal hwnd As Long) As Long



Dim msg As Long

On Error Resume Next


 msg = GetWindowLong(hwnd, GWL_EXSTYLE)

 msg = msg And Not WS_EX_LAYERED

 SetWindowLong hwnd, GWL_EXSTYLE, msg

 SetLayeredWindowAttributes hwnd, 0, 0, LWA_ALPHA
 DisableTransparanty = 0

If Err Then
    DisableTransparanty = 2
End If
   
End Function

