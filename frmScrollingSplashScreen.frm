VERSION 5.00
Begin VB.Form frmScrollingSplashScreen 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timScrollText 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1320
      Top             =   3480
   End
   Begin VB.Frame fraPictureBoxes 
      BackColor       =   &H00FF8080&
      Height          =   4155
      Left            =   60
      TabIndex        =   0
      Top             =   -20
      Width           =   7185
      Begin VB.PictureBox picBackgroundBuffer 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3945
         Left            =   60
         ScaleHeight     =   3945
         ScaleWidth      =   7065
         TabIndex        =   1
         Top             =   150
         Visible         =   0   'False
         Width           =   7065
      End
      Begin VB.PictureBox picTempBuffer 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3945
         Left            =   60
         ScaleHeight     =   3945
         ScaleWidth      =   7065
         TabIndex        =   3
         Top             =   150
         Visible         =   0   'False
         Width           =   7065
      End
      Begin VB.PictureBox picDestinationBuffer 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3945
         Left            =   90
         ScaleHeight     =   3945
         ScaleWidth      =   7035
         TabIndex        =   2
         Top             =   150
         Width           =   7035
      End
   End
End
Attribute VB_Name = "frmScrollingSplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit


Dim m_strTextArray() As String
Dim m_lngCurrentY    As Long

Public UnloadAfterScrollingPf As Boolean

'NOTE:  the project must reference the microsoft scripting library for the filesystemobject

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
       
    Call FormAlwaysOnTop(Me, True)
       
    Me.Caption = App.ProductName
    Call loadScrollingTextM(App.Path & "\ScrollingSplashScreen.txt", m_strTextArray)
    Call m_subInitializePictureBoxes
    Call m_subInitializeTimer(15)
   

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
    sErrSource = VBA.Err.Source & vbcrlf & "frmScrollingSplashScreen->"  & csProcName
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

   Unload Me
   Set frmScrollingSplashScreen = Nothing
   

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
    sErrSource = VBA.Err.Source & vbcrlf & "frmScrollingSplashScreen->"  & csProcName
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
' Name   : timScrollText_Timer
' Params : 
' Returns: Nothing
' Desc   : The Sub uses parameters  for timScrollText_Timer and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'		Auto-Code Commenter
'__________________________________________________
Private Sub timScrollText_Timer()
    On Error GoTo Proc_Err
    Const csProcName As String = "timScrollText_Timer"
    
    Dim l_blnContinueToScrollText As Boolean
    
    
    l_blnContinueToScrollText = g_funScrollText(m_strTextArray, _
                                           picBackgroundBuffer, _
                                           picTempBuffer, _
                                           picDestinationBuffer, _
                                           &HC0FFFF, _
                                           &H80FF&, _
                                           m_lngCurrentY, _
                                           0, vbCenter)
                    
    m_lngCurrentY = m_lngCurrentY - 1
    
    If Not (l_blnContinueToScrollText) Then
        If UnloadAfterScrollingPf Then
            Unload Me
        Else
            timScrollText.Enabled = False
        End If
    End If
    

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
    sErrSource = VBA.Err.Source & vbcrlf & "frmScrollingSplashScreen->"  & csProcName
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
' Name   : m_subInitializePictureBoxes
' Params : 
' Returns: Nothing
' Desc   : The Sub uses parameters  for m_subInitializePictureBoxes and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'		Auto-Code Commenter
'__________________________________________________
Private Sub m_subInitializePictureBoxes()
    On Error GoTo Proc_Err
    Const csProcName As String = "m_subInitializePictureBoxes"
    
   picBackgroundBuffer.ScaleMode = vbPixels
   picBackgroundBuffer.AutoRedraw = True
   picBackgroundBuffer.Visible = False
    
   picTempBuffer.ScaleMode = vbPixels
   picTempBuffer.AutoRedraw = True
   picTempBuffer.Visible = False
    
   picDestinationBuffer.ScaleMode = vbPixels
   picDestinationBuffer.AutoRedraw = True
   picDestinationBuffer.Visible = True
    
   m_lngCurrentY = picDestinationBuffer.ScaleHeight
    

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
    sErrSource = VBA.Err.Source & vbcrlf & "frmScrollingSplashScreen->"  & csProcName
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
' Name   : m_subInitializeTimer
' Params : 
'          ByVal v_intInterval As Integer
' Returns: Nothing
' Desc   : The Sub uses parameters ByVal v_intInterval As Integer for m_subInitializeTimer and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'		Auto-Code Commenter
'__________________________________________________
Private Sub m_subInitializeTimer(ByVal v_intInterval As Integer)
    On Error GoTo Proc_Err
    Const csProcName As String = "m_subInitializeTimer"

   timScrollText.Interval = v_intInterval
   timScrollText.Enabled = True


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
    sErrSource = VBA.Err.Source & vbcrlf & "frmScrollingSplashScreen->"  & csProcName
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

'Private Sub m_subLoadTextArrayFromFile(ByVal v_strFile As String, _
'                                       ByRef v_strTextArray() As String)
'
'   On Error GoTo ERROR_HANDLER
'
'   Dim l_lngIndex As Long
'
'   Open (v_strFile) For Input Access Read Shared As #1
'
'   Do Until EOF(1)
'      ReDim Preserve v_strTextArray(l_lngIndex)
'      Line Input #1, v_strTextArray(l_lngIndex)
'      l_lngIndex = l_lngIndex + 1
'   Loop
'
'
'
'   Close #1
'
'EXIT_HANLDER:
'   Exit Sub
'
'ERROR_HANDLER:
'   ReDim Preserve v_strTextArray(3)
'
'   v_strTextArray(0) = "Error, Unable To Load File"
'   v_strTextArray(1) = ""
'   v_strTextArray(2) = "Contact You Supervisor"
'   v_strTextArray(3) = "For Assistance"
'
'End Sub
'
'
'---------------
'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : loadScrollingTextM
' Params : 
' Returns: _
' Desc   : The Sub uses parameters  for loadScrollingTextM and returns _.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'		Auto-Code Commenter
'__________________________________________________
Private Sub loadScrollingTextM(ByVal v_strFile As String, _
                                       ByRef v_strTextArray() As String)

   On Error GoTo ERROR_HANDLER
    
    Dim ofs As clsFs
    Dim vReplacementListArr(1, 1) As Variant
     
    vReplacementListArr(0, 0) = gcsCompanyPlaceholder
    vReplacementListArr(0, 1) = App.CompanyName
    vReplacementListArr(1, 0) = gcsUserPlaceholder
    vReplacementListArr(1, 1) = GetUserNamePv()
    
    Set ofs = New clsFs
    v_strTextArray = ofs.TextFileToArrayPv(v_strFile, vbCrLf, vReplacementListArr)
    Set ofs = Nothing

    
    
EXIT_HANLDER:
   Exit Sub
   
ERROR_HANDLER:
   ReDim Preserve v_strTextArray(3)
   
   v_strTextArray(0) = "Error, Unable To Load File"
   v_strTextArray(1) = ""
   v_strTextArray(2) = "Contact You Supervisor"
   v_strTextArray(3) = "For Assistance"
   
End Sub


