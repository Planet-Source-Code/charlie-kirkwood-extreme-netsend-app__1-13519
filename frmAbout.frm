VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About..."
   ClientHeight    =   5700
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6390
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3934.24
   ScaleMode       =   0  'User
   ScaleWidth      =   6000.541
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Disclamer:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   105
      TabIndex        =   6
      Top             =   2895
      Width           =   6180
      Begin VB.Label lblLegalCopyright 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":0000
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1920
         Left            =   180
         TabIndex        =   7
         Top             =   405
         Width           =   5715
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   8
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":0237
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   810
      TabIndex        =   5
      Top             =   1950
      Width           =   5025
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ckirkwood9@yahoo.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   810
      TabIndex        =   4
      Top             =   1635
      Width           =   2070
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Charlie Kirkwood"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   810
      TabIndex        =   3
      Top             =   1410
      Width           =   1920
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Extreme NetSend"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   450
      TabIndex        =   2
      Top             =   510
      Width           =   3585
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Out Of Bounds"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   105
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : cmdOK_Click
' Params : 
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdOK_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'		Auto-Code Commenter
'__________________________________________________
Private Sub cmdOK_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "cmdOK_Click"
    Me.Hide

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
    sErrSource = VBA.Err.Source & vbcrlf & "frmAbout->"  & csProcName
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

    Dim sLegalCopyright As String

    'turn off the icon to the form so it will take on a standard windows icon
    Set Me.Icon = Nothing
    
    sLegalCopyright = Replace(App.LegalCopyright, gcsUserPlaceholder, Trim(GetUserNamePv))
    sLegalCopyright = Replace(sLegalCopyright, gcsCompanyPlaceholder, Trim(App.CompanyName))
    
    Me.lblTitle = App.Title
    Me.lblVersion = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    Me.lblLegalCopyright = sLegalCopyright
    Me.Caption = "About " & App.Title & " ..."
    
    

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
    sErrSource = VBA.Err.Source & vbcrlf & "frmAbout->"  & csProcName
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



