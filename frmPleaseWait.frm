VERSION 5.00
Begin VB.Form frmPleaseWait 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2550
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4035
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4035
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblDefaultMessage 
      BackColor       =   &H00000000&
      Caption         =   "Please Wait ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   885
      Left            =   345
      TabIndex        =   1
      Top             =   990
      Width           =   3540
   End
   Begin VB.Label lblCustomMessage 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Custom Message Here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1995
      Left            =   210
      TabIndex        =   0
      Top             =   270
      Width           =   3585
   End
End
Attribute VB_Name = "frmPleaseWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Public CustomMessagePs As String
    Dim moHourglass As clsHourglass

'__________________________________________________
' Scope  : Public
' Type   : Sub
' Name   : CustomMessageP
' Params : 
'          sMessage As String
' Returns: Nothing
' Desc   : The Sub uses parameters sMessage As String for CustomMessageP and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'		Auto-Code Commenter
'__________________________________________________
Public Sub CustomMessageP(sMessage As String)
    On Error GoTo Proc_Err
    Const csProcName As String = "CustomMessageP"

    Me.lblCustomMessage.Caption = sMessage
    Me.lblCustomMessage.Visible = True
    Me.lblDefaultMessage.Visible = False
    DoEvents
    Me.Refresh


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
    sErrSource = VBA.Err.Source & vbcrlf & "frmPleaseWait->"  & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly
End Sub

'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : Form_Load
' Params : 
' Returns: Nothing
' Desc   : The Sub uses parameters  for Form_Load and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'		Auto-Code Commenter
'__________________________________________________
Private Sub Form_Load()
    On Error GoTo Proc_Err
    Const csProcName As String = "Form_Load"

    
    'turn on the hourglass while waiting
    Set moHourglass = New clsHourglass
    
    Me.lblCustomMessage.Visible = False
    Me.lblDefaultMessage.Visible = True

    'First, set the position of the form
    Me.Move (Screen.Width / 2) - (Me.Width / 2), (Screen.Height / 2) - (Me.Height / 2)
    
    'Explode form is a little buggy... it leaves the 'explode' part on the screen if user sends a message to himself
    'Next, call the ExplodeForm
    'Call ExplodeForm(frmPleaseWait, 200, vbYellow)
    
    Call FormAlwaysOnTop(Me, True)
    
    DoEvents
    

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
    sErrSource = VBA.Err.Source & vbcrlf & "frmPleaseWait->"  & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly
End Sub

'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : Form_Unload
' Params : 
'          Cancel As Integer
' Returns: Nothing
' Desc   : The Sub uses parameters Cancel As Integer for Form_Unload and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'		Auto-Code Commenter
'__________________________________________________
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Proc_Err
    Const csProcName As String = "Form_Unload"

    Set moHourglass = Nothing
    


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
    sErrSource = VBA.Err.Source & vbcrlf & "frmPleaseWait->"  & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly
End Sub

