Attribute VB_Name = "modScrollingSplashScreen"
Option Explicit


Private Declare Function BitBlt Lib "gdi32" (ByVal hdcDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Const SRCCOPY = &HCC0020

'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : g_funScrollText
' Params : 
'          ByRef v_strTextArray() As String
'          ByRef r_ctlBackgroundBuffer As Control
'          ByRef r_ctlTempBuffer As Control
'          ByRef r_ctlDestinationBuffer As Control
'          ByVal v_lngRGBStartColor As Long
'          ByVal v_lngRGBEndColor As Long
'          ByVal v_lngCurrentY As Long
'          ByVal v_lngLeftMargine As Long
'          ByVal v_enuAlignment As VBRUN.AlignmentConstants
' Returns: Boolean
' Desc   : The Function uses parameters ByRef v_strTextArray() As String, ByRef r_ctlBackgroundBuffer As Control, ByRef r_ctlTempBuffer As Control, ByRef r_ctlDestinationBuffer As Control, ByVal v_lngRGBStartColor As Long, ByVal v_lngRGBEndColor As Long, ByVal v_lngCurrentY As Long, ByVal v_lngLeftMargine As Long and ByVal v_enuAlignment As VBRUN.AlignmentConstants for g_funScrollText and returns Boolean.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'		Auto-Code Commenter
'__________________________________________________
Public Function g_funScrollText(ByRef v_strTextArray() As String, ByRef r_ctlBackgroundBuffer As Control, ByRef r_ctlTempBuffer As Control, ByRef r_ctlDestinationBuffer As Control, ByVal v_lngRGBStartColor As Long, ByVal v_lngRGBEndColor As Long, ByVal v_lngCurrentY As Long, ByVal v_lngLeftMargine As Long, ByVal v_enuAlignment As VBRUN.AlignmentConstants) As Boolean
    On Error GoTo Proc_Err
    Const csProcName As String = "g_funScrollText"
   
   Dim l_lngStartRed   As Long
   Dim l_lngStartGreen As Long
   Dim l_lngStartBlue  As Long
   
   Dim l_lngEndRed     As Long
   Dim l_lngEndGreen   As Long
   Dim l_lngEndBlue    As Long

   Dim l_lngCurrentRed   As Long
   Dim l_lngCurrentGreen As Long
   Dim l_lngCurrentBlue  As Long

   Dim l_sngRedOffset    As Single
   Dim l_sngGreenOffset  As Single
   Dim l_sngBlueOffset   As Single
   
   Dim l_sngTextHeight  As Single
   Dim l_lngScaleHeight As Single
   Dim l_lngScaleWidth  As Single

   Dim l_lngLineNumber     As Long
   Dim l_lngNumberOfLines  As Long


   g_funScrollText = True
   
   l_lngNumberOfLines = UBound(v_strTextArray)
                           
   l_sngTextHeight = r_ctlTempBuffer.TextHeight("ABCDEFGHIJKLMNOPQRSTUVWXYZ")
   l_lngScaleHeight = r_ctlTempBuffer.ScaleHeight
   l_lngScaleWidth = r_ctlTempBuffer.ScaleWidth
                           
   If (v_lngRGBStartColor <> v_lngRGBEndColor) Then
      Call g_subGetRGBColors(v_lngRGBStartColor, l_lngStartRed, l_lngStartGreen, l_lngStartBlue)
      Call g_subGetRGBColors(v_lngRGBEndColor, l_lngEndRed, l_lngEndGreen, l_lngEndBlue)
      
      l_sngRedOffset = (CSng(l_lngEndRed - l_lngStartRed) / (l_lngScaleHeight - l_sngTextHeight))
      l_sngGreenOffset = (CSng(l_lngEndGreen - l_lngStartGreen) / (l_lngScaleHeight - l_sngTextHeight))
      l_sngBlueOffset = (CSng(l_lngEndBlue - l_lngStartBlue) / (l_lngScaleHeight - l_sngTextHeight))
   Else
      Call g_subGetRGBColors(v_lngRGBStartColor, l_lngCurrentRed, l_lngCurrentGreen, l_lngCurrentBlue)
   End If
   
   BitBlt r_ctlTempBuffer.hdc, 0, r_ctlTempBuffer.ScaleTop, l_lngScaleWidth, l_lngScaleHeight, _
          r_ctlBackgroundBuffer.hdc, 0, 0, SRCCOPY
          
   With r_ctlTempBuffer
      For l_lngLineNumber = 0 To l_lngNumberOfLines
         .CurrentY = v_lngCurrentY + (l_lngLineNumber * .FontSize + (6 * l_lngLineNumber))
         If (v_enuAlignment = vbCenter) Then
            .CurrentX = (l_lngScaleWidth - .TextWidth(v_strTextArray(l_lngLineNumber))) / 2
         ElseIf (v_enuAlignment = vbLeftJustify) Then
            .CurrentX = 0
         ElseIf (v_enuAlignment = vbRightJustify) Then
            .CurrentX = l_lngScaleWidth - .TextWidth(v_strTextArray(l_lngLineNumber))
         End If

         .CurrentX = .CurrentX + v_lngLeftMargine
         
         If Not (.CurrentY > l_lngScaleHeight) And _
            Not (.CurrentY < -l_sngTextHeight) Then
            If (v_lngRGBStartColor <> v_lngRGBEndColor) Then
               l_lngCurrentRed = Abs(l_lngEndRed - (l_sngRedOffset * .CurrentY))
               l_lngCurrentGreen = Abs(l_lngEndGreen - (l_sngGreenOffset * .CurrentY))
               l_lngCurrentBlue = Abs(l_lngEndBlue - (l_sngBlueOffset * .CurrentY))
            End If
            
            .ForeColor = RGB(l_lngCurrentRed, l_lngCurrentGreen, l_lngCurrentBlue)

            r_ctlTempBuffer.Print v_strTextArray(l_lngLineNumber)
         End If

         If (l_lngLineNumber = l_lngNumberOfLines) And (.CurrentY <= -l_sngTextHeight) Then
            g_funScrollText = False
         End If
      Next
   End With

   BitBlt r_ctlDestinationBuffer.hdc, 0, r_ctlDestinationBuffer.ScaleTop, r_ctlDestinationBuffer.ScaleWidth, r_ctlDestinationBuffer.ScaleHeight, _
          r_ctlTempBuffer.hdc, 0, 0, SRCCOPY

   r_ctlDestinationBuffer.Refresh


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
    sErrSource = VBA.Err.Source & vbcrlf & "modScrollingSplashScreen->"  & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    Err.Raise Number:=lErrNum, Source:=sErrSource, Description:=sErrDesc
    Exit Function
    
End Function
'__________________________________________________
' Scope  : Public
' Type   : Sub
' Name   : g_subGetRGBColors
' Params : 
' Returns: _
' Desc   : The Sub uses parameters  for g_subGetRGBColors and returns _.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'		Auto-Code Commenter
'__________________________________________________
Public Sub g_subGetRGBColors(ByVal v_lngRGBColor As Long, _
                             ByRef r_lngRedColor As Long, _
                             ByRef r_lngGreenColor As Long, _
                             ByRef r_lngBlueColor As Long)
    On Error GoTo Proc_Err
    Const csProcName As String = "g_subGetRGBColors"
        
    r_lngRedColor = v_lngRGBColor Mod 256
    r_lngGreenColor = (v_lngRGBColor \ &H100) Mod 256
    r_lngBlueColor = (v_lngRGBColor \ &H10000) Mod 256
    

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
    sErrSource = VBA.Err.Source & vbcrlf & "modScrollingSplashScreen->"  & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    Err.Raise Number:=lErrNum, Source:=sErrSource, Description:=sErrDesc
    Exit Sub
    
End Sub


