Attribute VB_Name = "modEffects"
'Declarations for ExplodeForm
Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long  'note error in declare
Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const conHwndTopmost = -1
Public Const conHwndNoTopmost = -2
Public Const conSwpNoActivate = &H10
Public Const conSwpShowWindow = &H40


'__________________________________________________
' Scope  : 
' Type   : Sub
' Name   : ExplodeForm
' Params : 
'          frm As Form
'          Steps As Long
'          Color As VBRUN.ColorConstants
' Returns: Nothing
' Desc   : The Sub uses parameters frm As Form, Steps As Long and Color As VBRUN.ColorConstants for ExplodeForm and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'		Auto-Code Commenter
'__________________________________________________
Sub ExplodeForm(frm As Form, Steps As Long, Color As VBRUN.ColorConstants)
    On Error GoTo Proc_Err
    Const csProcName As String = "ExplodeForm"
   Dim ThisRect As RECT, RectWidth As Integer, RectHeight As Integer, ScreenDevice As Long, NewBrush As Long, OldBrush As Long, I As Long, x As Integer, y As Integer, XRect As Integer, YRect As Integer
   If Steps < 20 Then Steps = 20
   'Zooming speed will be different based on machine speed!
   If Color = 0 Then
      Color = frm.BackColor
   End If
   
   
   Steps = Steps * 10
   'Get current form window dimensions
   GetWindowRect frm.hwnd, ThisRect
   RectWidth = (ThisRect.Right - ThisRect.Left)
   RectHeight = ThisRect.Bottom - ThisRect.Top
   'Get a device handle for the screen
   ScreenDevice = GetDC(0)
   'Create a brush for drawing to the screen
   'and save the old brush
   NewBrush = CreateSolidBrush(Color)
   OldBrush = SelectObject(ScreenDevice, NewBrush)
   For I = 1 To Steps
      XRect = RectWidth * (I / Steps)
      YRect = RectHeight * (I / Steps)
      x = ThisRect.Left + (RectWidth - XRect) / 2
      y = ThisRect.Top + (RectHeight - YRect) / 2
      'Incrementally draw rectangle
      Rectangle ScreenDevice, x, y, x + XRect, y + YRect
   Next I
   'Return old brush and delete screen device context handle
   'Then destroy brush that drew rectangles
   Call SelectObject(ScreenDevice, OldBrush)
   Call ReleaseDC(0, ScreenDevice)
   DeleteObject (NewBrush)

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
    sErrSource = VBA.Err.Source & vbcrlf & "modEffects->"  & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    Err.Raise Number:=lErrNum, Source:=sErrSource, Description:=sErrDesc
    Exit Sub
    
End Sub

'__________________________________________________
' Scope  : Public
' Type   : Sub
' Name   : FormAlwaysOnTop
' Params : 
'          oForm As Form
'          fOnTop As Boolean
' Returns: Nothing
' Desc   : The Sub uses parameters oForm As Form and fOnTop As Boolean for FormAlwaysOnTop and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'		Auto-Code Commenter
'__________________________________________________
Public Sub FormAlwaysOnTop(oForm As Form, fOnTop As Boolean)
    On Error GoTo Proc_Err
    Const csProcName As String = "FormAlwaysOnTop"
    
    Dim vTopSwitch As Variant

    If fOnTop Then
        vTopSwitch = conHwndTopmost
    Else
        vTopSwitch = conHwndNoTopmost
    End If
    
    SetWindowPos oForm.hwnd, _
                vTopSwitch, _
                oForm.Left / Screen.TwipsPerPixelX, _
                oForm.Top / Screen.TwipsPerPixelY, _
                oForm.Width / Screen.TwipsPerPixelX, _
                oForm.Height / Screen.TwipsPerPixelY, _
                conSwdpNoActivate Or conSwpShowWindow
    

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
    sErrSource = VBA.Err.Source & vbcrlf & "modEffects->"  & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    Err.Raise Number:=lErrNum, Source:=sErrSource, Description:=sErrDesc
    Exit Sub
    
End Sub

