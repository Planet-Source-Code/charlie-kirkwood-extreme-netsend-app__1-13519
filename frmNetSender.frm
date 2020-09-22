VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmNetsender 
   Caption         =   "Extreme Netsend"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8730
   Icon            =   "frmNetSender.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   8730
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdGetDomains 
      Caption         =   "Get Domains"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   225
      TabIndex        =   0
      Top             =   360
      Width           =   1185
   End
   Begin VB.Frame fraFrameForTabStrip 
      Height          =   5790
      Index           =   1
      Left            =   765
      TabIndex        =   41
      Tag             =   "Netsend Using Group"
      Top             =   1500
      Width           =   7965
      Begin VB.Frame fraFrameForUsersAvailable 
         BorderStyle     =   0  'None
         Height          =   1845
         Index           =   2
         Left            =   3975
         TabIndex        =   48
         Tag             =   "Manual Entry of User"
         Top             =   1425
         Width           =   3990
         Begin VB.CommandButton cmdAddEntry 
            Caption         =   "&Add"
            Height          =   375
            Left            =   975
            TabIndex        =   17
            ToolTipText     =   "Add all available computers to netsend list"
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox txtEnterUser 
            Height          =   330
            Left            =   0
            TabIndex        =   16
            Top             =   210
            Width           =   3945
         End
         Begin VB.CommandButton cmdDummyRemoveAll 
            Caption         =   "Re&move All"
            Height          =   375
            Left            =   2970
            TabIndex        =   19
            ToolTipText     =   "Remove all computers from netsend list"
            Top             =   1440
            Width           =   960
         End
         Begin VB.CommandButton cmdDummyRemove 
            Caption         =   "&Remove"
            Height          =   375
            Left            =   1980
            TabIndex        =   18
            ToolTipText     =   "Remove selected computers from netsend list"
            Top             =   1440
            Width           =   960
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Enter a &user to include in the Netsend List - Click Add"
            Height          =   405
            Left            =   30
            TabIndex        =   15
            Top             =   0
            Width           =   3885
         End
      End
      Begin VB.Frame fraFrameForUsersAvailable 
         BorderStyle     =   0  'None
         Height          =   1845
         Index           =   1
         Left            =   3390
         TabIndex        =   46
         Tag             =   "Users in Selected Group"
         Top             =   570
         Width           =   3990
         Begin VB.ListBox lstUsersAvailable 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1185
            Left            =   0
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   10
            Top             =   210
            Width           =   3930
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   375
            Left            =   990
            TabIndex        =   12
            ToolTipText     =   "Add selected computers to netsend list"
            Top             =   1440
            Width           =   960
         End
         Begin VB.CommandButton cmdAddAll 
            Caption         =   "Add A&ll"
            Height          =   375
            Left            =   0
            TabIndex        =   11
            ToolTipText     =   "Add all available computers to netsend list"
            Top             =   1440
            Width           =   960
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "&Remove"
            Height          =   375
            Left            =   1980
            TabIndex        =   13
            ToolTipText     =   "Remove selected computers from netsend list"
            Top             =   1440
            Width           =   960
         End
         Begin VB.CommandButton cmdRemoveAll 
            Caption         =   "Re&move All"
            Height          =   375
            Left            =   2970
            TabIndex        =   14
            ToolTipText     =   "Remove all computers from netsend list"
            Top             =   1440
            Width           =   960
         End
         Begin VB.Label lblUsersAvailable 
            BackStyle       =   0  'Transparent
            Caption         =   "&Users"
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   2595
         End
         Begin VB.Label lblTotalUsersAvailable 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Members:"
            Height          =   255
            Left            =   2430
            TabIndex        =   47
            Top             =   0
            Width           =   1500
         End
      End
      Begin VB.CheckBox chkClearMessageAfterNetsend 
         Caption         =   "Clear A&fter Netsend"
         Height          =   195
         Left            =   3030
         TabIndex        =   25
         Top             =   4410
         Width           =   1740
      End
      Begin VB.CheckBox chkIncludeDisclaimer 
         Caption         =   "&Include disclaimer"
         Height          =   195
         Left            =   1335
         TabIndex        =   24
         Top             =   4410
         Width           =   1650
      End
      Begin VB.CheckBox chkClearListAfterNetsend 
         Caption         =   "Clear list after &Netsend"
         Height          =   390
         Left            =   3585
         TabIndex        =   21
         Top             =   3825
         Width           =   1470
      End
      Begin VB.CommandButton cmdSaveAsDefault 
         Caption         =   "Sa&ve As"
         Height          =   375
         Left            =   6465
         TabIndex        =   23
         Top             =   3840
         Width           =   1305
      End
      Begin VB.CommandButton cmdLoadDefault 
         Caption         =   "L&oad template"
         Height          =   375
         Left            =   5100
         TabIndex        =   22
         Top             =   3840
         Width           =   1305
      End
      Begin VB.ListBox lstUsersSelected 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   3585
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   20
         Top             =   2625
         Width           =   4170
      End
      Begin VB.TextBox txtMessage 
         Appearance      =   0  'Flat
         Height          =   975
         Left            =   150
         MultiLine       =   -1  'True
         TabIndex        =   27
         ToolTipText     =   "Type your Netsend message here  (NT systems only)"
         Top             =   4650
         Width           =   6555
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   6840
         TabIndex        =   29
         ToolTipText     =   "bye bye birdie"
         Top             =   5265
         Width           =   975
      End
      Begin VB.ListBox lstGroups 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3660
         Left            =   150
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   3315
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "&Send"
         Height          =   375
         Left            =   6825
         TabIndex        =   28
         ToolTipText     =   "Send Netsend message"
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   6825
         TabIndex        =   30
         Top             =   4800
         Width           =   975
      End
      Begin ComctlLib.TabStrip tsTabStripForUsersAvailable 
         Height          =   2370
         Left            =   3570
         TabIndex        =   8
         Top             =   195
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   4180
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   1
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   ""
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label lblNumUsersSelected 
         Alignment       =   1  'Right Justify
         Caption         =   "##"
         Height          =   255
         Left            =   6360
         TabIndex        =   45
         Top             =   4410
         Width           =   345
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Users to Netsend:"
         Height          =   255
         Left            =   4125
         TabIndex        =   44
         Top             =   4410
         Width           =   2205
      End
      Begin VB.Label Label7 
         Caption         =   "Message &text:"
         Height          =   255
         Left            =   150
         TabIndex        =   26
         Top             =   4410
         Width           =   1725
      End
      Begin VB.Label lblTotalGroups 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Groups:"
         Height          =   255
         Left            =   2325
         TabIndex        =   43
         Top             =   270
         Width           =   1140
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Sele&ct Group"
         Height          =   255
         Left            =   150
         TabIndex        =   6
         Top             =   270
         Width           =   3420
      End
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "A&bout..."
      Default         =   -1  'True
      Height          =   375
      Left            =   7425
      TabIndex        =   4
      Top             =   360
      Width           =   1065
   End
   Begin VB.Timer tmrGetTimeOfDay 
      Interval        =   1000
      Left            =   4185
      Top             =   6285
   End
   Begin VB.Timer tmrDoNetSend 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3690
      Top             =   6285
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6810
      Top             =   225
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   6195
      Top             =   210
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7275
      Top             =   270
   End
   Begin VB.CommandButton cmdGetGroups 
      Caption         =   "Get &Groups"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5085
      TabIndex        =   3
      Top             =   360
      Width           =   1065
   End
   Begin VB.ComboBox cboDomain 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1455
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   3600
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   42
      Top             =   7305
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   9763
            Text            =   "System messages here."
            TextSave        =   "System messages here."
            Key             =   "msg"
            Object.Tag             =   ""
            Object.ToolTipText     =   "System things here"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "System Time"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "12/11/00"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "System Date"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraFrameForTabStrip 
      Height          =   5985
      Index           =   2
      Left            =   435
      TabIndex        =   31
      Tag             =   "Print Group/Users"
      Top             =   1320
      Width           =   8220
      Begin VB.CommandButton cmdSaveToFile 
         Caption         =   "Save To File"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   38
         Top             =   4740
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   7380
         TabIndex        =   39
         Text            =   "10"
         Top             =   4305
         Width           =   405
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3570
         Left            =   165
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   37
         Top             =   1545
         Width           =   6300
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Print All Groups"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   165
         TabIndex        =   36
         Top             =   270
         Width           =   2775
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Print All Groups and their Members"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   165
         TabIndex        =   35
         Top             =   555
         Width           =   3345
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Print Selected Group with it's Members"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   165
         TabIndex        =   34
         Top             =   840
         Width           =   3510
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Preview Only - All Groups and thier Members"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   165
         TabIndex        =   33
         Top             =   1155
         Value           =   -1  'True
         Width           =   4125
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   32
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Font Size"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   7395
         TabIndex        =   40
         Top             =   3840
         Width           =   420
      End
   End
   Begin ComctlLib.TabStrip tsMainTabstrip 
      Height          =   6315
      Left            =   225
      TabIndex        =   5
      Top             =   825
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   11139
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "&Enter or Select Domain,Computer Name or IP"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1455
      TabIndex        =   1
      Top             =   90
      Width           =   4185
   End
End
Attribute VB_Name = "frmNetsender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Enum eCommonDialogConst
    ecdlShowOpen = 1
    ecdlShowSave = 2
    ecdlShowColor = 3
    ecdlShowFont = 4
    ecdlShowPrinter = 5
    ecdlShowWinHelp32 = 6
End Enum


Private Const mcsDefaultFile As String = "DefaultNetsendList.txt"

'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : cmdAddEntry_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdAddEntry_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub cmdAddEntry_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "cmdAddEntry_Click"

    If Me.txtEnterUser & "" <> "" Then
        Me.lstUsersSelected.AddItem Me.txtEnterUser.Text
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
    sErrSource = VBA.Err.Source & vbCrLf & "frmNetsender->" & csProcName
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
' Name   : cmdDummyRemove_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdDummyRemove_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub cmdDummyRemove_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "cmdDummyRemove_Click"

    Call cmdRemove_Click


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
    sErrSource = VBA.Err.Source & vbCrLf & "frmNetsender->" & csProcName
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
' Name   : cmdDummyRemoveAll_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdDummyRemoveAll_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub cmdDummyRemoveAll_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "cmdDummyRemoveAll_Click"

    Call cmdRemoveAll_Click


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
    sErrSource = VBA.Err.Source & vbCrLf & "frmNetsender->" & csProcName
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
' Name   : cmdGetDomains_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdGetDomains_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub cmdGetDomains_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "cmdGetDomains_Click"

    frmPleaseWait.Show
    initializeDomainsP
    Unload frmPleaseWait
    

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
    sErrSource = VBA.Err.Source & vbCrLf & "frmNetsender->" & csProcName
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
' Name   : cmdGetGroups_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdGetGroups_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub cmdGetGroups_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "cmdGetGroups_Click"
    
    Call getGroupsP
    

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
    sErrSource = VBA.Err.Source & vbCrLf & "frmNetsender->" & csProcName
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
' Name   : cmdLoadDefault_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdLoadDefault_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub cmdLoadDefault_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "cmdLoadDefault_Click"


    Dim sFile As String
    
    sFile = App.Path & "\" & mcsDefaultFile
    
    sFile = GetTemplateFileName(Me.CD1, ecdlShowOpen, sFile, "*.txt")
    If sFile & "" = "" Then
        If MsgBox("you have not selected a file load, do you want to load your default netsend list?", vbYesNoCancel) = vbYes Then
            sFile = App.Path & "\" & mcsDefaultFile
        End If
    End If
    
    If sFile & "" <> "" Then
        Call GetListFromTextFileP(Me.lstUsersSelected, sFile)
    End If
    
    lblNumUsersSelected.Caption = lstUsersSelected.ListCount


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
    sErrSource = VBA.Err.Source & vbCrLf & "frmNetsender->" & csProcName
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
' Name   : cmdPrint_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdPrint_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub cmdPrint_Click()
    
    On Error Resume Next
    
    CD1.CancelError = False
    
    If Option1.Value = True Then
        Text1.Text = ""
        Text1.Text = "(Domain, Computer Name or IP) - " & cboDomain.Text & vbCrLf & vbCrLf
        Text1.Text = Text1.Text & "(All Groups:)" & vbCrLf
        
        Do Until lstGroups.ListCount = 0
            lstGroups.ListIndex = 0
            Text1.Text = Text1.Text & vbTab & lstGroups.Text & vbCrLf
            lstGroups.RemoveItem lstGroups.ListIndex
        Loop
        
        DoEvents
        DoEvents
        
        CD1.ShowPrinter
        
        DoEvents
        
        Printer.FontSize = Text2.Text
        Printer.Print Text1.Text
        DoEvents
        Printer.EndDoc
        Call getGroupsP
    End If
    
    If Option2.Value = True Then
        
        Text1.Text = ""
        Text1.Text = "(Domain, Computer Name or IP) - " & cboDomain.Text & vbCrLf & vbCrLf
        
        Do Until lstGroups.ListCount = 0
            
            lstGroups.ListIndex = 0
            Call lstGroups_DblClick
            DoEvents
            DoEvents
            Text1.Text = Text1.Text & "(Group) - " & lstGroups.Text & vbCrLf
            Text1.Text = Text1.Text & vbTab & "(Members:) - " & lstUsersAvailable.ListCount & vbCrLf
            DoEvents
            
            Do Until lstUsersAvailable.ListCount = 0
                lstUsersAvailable.ListIndex = 0
                Text1.Text = Text1.Text & vbTab & vbTab & lstUsersAvailable.Text & vbCrLf
                lstUsersAvailable.RemoveItem lstUsersAvailable.ListIndex
            Loop
            
            Text1.Text = Text1.Text & vbCrLf
            DoEvents
            lstGroups.RemoveItem lstGroups.ListIndex
        
        Loop
        
        DoEvents
        DoEvents
        CD1.ShowPrinter
        DoEvents
        Printer.FontSize = Text2.Text
        Printer.Print Text1.Text
        DoEvents
        Printer.EndDoc
        Call getGroupsP
    
    End If
    
    If Option3.Value = True Then
        Text1.Text = ""
        Text1.Text = "(Domain, Computer Name or IP) - " & cboDomain.Text & vbCrLf & vbCrLf
        Text1.Text = Text1.Text & "(Group) - " & lstGroups.Text & vbCrLf
        Text1.Text = Text1.Text & vbTab & "(Members:)" & vbCrLf
        
        Do Until lstUsersAvailable.ListCount = 0
            lstUsersAvailable.ListIndex = 0
            Text1.Text = Text1.Text & vbTab & vbTab & lstUsersAvailable.Text & vbCrLf
            lstUsersAvailable.RemoveItem lstUsersAvailable.ListIndex
        Loop
        
        DoEvents
        DoEvents
        CD1.ShowPrinter
        DoEvents
        Printer.FontSize = Text2.Text
        Printer.Print Text1.Text
        DoEvents
        Printer.EndDoc
        Call lstGroups_DblClick
    End If
    
    If Option4.Value = True Then
        Text1.Text = ""
        Text1.Text = "(Domain, Computer Name or IP) - " & cboDomain.Text & vbCrLf & vbCrLf
        
        Do Until lstGroups.ListCount = 0
            lstGroups.ListIndex = 0
            Call lstGroups_DblClick
            DoEvents
            DoEvents
            Text1.Text = Text1.Text & "(Group) - " & lstGroups.Text & vbCrLf
            Text1.Text = Text1.Text & vbTab & "(Members:) - " & lstUsersAvailable.ListCount & vbCrLf
            DoEvents
            Do Until lstUsersAvailable.ListCount = 0
                lstUsersAvailable.ListIndex = 0
                Text1.Text = Text1.Text & vbTab & vbTab & lstUsersAvailable.Text & vbCrLf
                lstUsersAvailable.RemoveItem lstUsersAvailable.ListIndex
            Loop
            Text1.Text = Text1.Text & vbCrLf
            DoEvents
            lstGroups.RemoveItem lstGroups.ListIndex
        Loop
        
        DoEvents
        DoEvents
        Call getGroupsP
    End If

End Sub

'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : cmdAbout_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdAbout_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub cmdAbout_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "cmdAbout_Click"
    
    frmAbout.Show
    

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
    sErrSource = VBA.Err.Source & vbCrLf & "frmNetsender->" & csProcName
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
' Name   : cmdSaveAsDefault_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdSaveAsDefault_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub cmdSaveAsDefault_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "cmdSaveAsDefault_Click"

    Dim sFile As String
    
    sFile = App.Path & "\" & mcsDefaultFile
    
    sFile = GetTemplateFileName(Me.CD1, ecdlShowSave, sFile, "*.txt")
    If sFile & "" = "" Then
        If MsgBox("you have not selected a file to save to, do you want to save this as your default netsend list?", vbYesNoCancel) = vbYes Then
            sFile = App.Path & "\" & mcsDefaultFile
        End If
    End If
    
    If sFile & "" <> "" Then
        Call SaveListToTextFileP(Me.lstUsersSelected, sFile)
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
    sErrSource = VBA.Err.Source & vbCrLf & "frmNetsender->" & csProcName
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
' Name   : cmdSaveToFile_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdSaveToFile_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub cmdSaveToFile_Click()
    On Error Resume Next
    CD1.CancelError = False
    CD1.Filter = "Text Document (*.txt)|*.txt"
    CD1.ShowSave
    
    If CD1.FileName = "" Then
        Exit Sub
    End If
    
    Open CD1.FileName For Output As #1
    Print #1, Text1.Text
    Close #1
    
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
'               Auto-Code Commenter
'__________________________________________________
Private Sub Form_Load()
    On Error GoTo Proc_Err
    Const csProcName As String = "Form_Load"
    
    Me.Visible = True
    
    frmPleaseWait.Show
    
    initializeTabStrips
    'initializeDomainsP
    
    
    Call GetListFromTextFileP(Me.lstUsersSelected, App.Path & "\" & mcsDefaultFile)
    lblNumUsersSelected.Caption = lstUsersSelected.ListCount
    
    Unload frmPleaseWait
    
    DoEvents
    
    Dim oFrm As frmScrollingSplashScreen
    Set oFrm = New frmScrollingSplashScreen
    oFrm.UnloadAfterScrollingPf = True
    oFrm.Show
    
    
Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    Set oFrm = Nothing
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmNetsender->" & csProcName
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
' Name   : initializeDomainsP
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for initializeDomainsP and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub initializeDomainsP()
    On Error GoTo Proc_Err
    Const csProcName As String = "initializeDomainsP"
    
    Dim namespace As IADsContainer
    Dim domain As IADs
    Dim sOldStatusMessage As String
    
    sOldStatusMessage = StatusBar1.Panels("msg").Text
    StatusBar1.Panels("msg").Text = "Retrieving Domains"
     
    cboDomain.AddItem Me.Winsock1.LocalHostName
     
     'Loads Combobox1 with all the current domains
    Set namespace = GetObject("WinNT:")
    
    For Each domain In namespace
        cboDomain.AddItem domain.Name
    Next
    
    StatusBar1.Panels("msg").Text = sOldStatusMessage
    Set namespace = Nothing
    Set domain = Nothing
    
    
    

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
    sErrSource = VBA.Err.Source & vbCrLf & "frmNetsender->" & csProcName
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
' Name   : GetListFromTextFileP
' Params :
'          oList As Control
'          sFileName As String
'          Optional fSilent As Boolean = False
' Returns: Nothing
' Desc   : The Sub uses parameters oList As Control, sFileName As String and Optional fSilent As Boolean = False for GetListFromTextFileP and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub GetListFromTextFileP(oList As Control, sFileName As String, Optional fSilent As Boolean = False)
    On Error GoTo Proc_Err
    Const csProcName As String = "GetListFromTextFileP"
    'get the default list from the saved file
        
    Dim sOldStatusMessage As String
    Dim ofs As clsFs
    Dim oTextStream As TextStream
    Dim sText As String
    Dim aArr As Variant
    Dim lCount As Long
    Dim lMin As Long
    Dim lMax As Long
        
    
    sOldStatusMessage = StatusBar1.Panels("msg").Text
    
    If Not fSilent Then
        StatusBar1.Panels("msg").Text = "Getting Default Netsend List"
    End If
            
        
    
    Set ofs = New clsFs
    If ofs.FileExists(sFileName) Then
        aArr = ofs.TextFileToArrayPv(sFileName, vbCrLf)
        lMin = LBound(aArr)
        lMax = UBound(aArr)
        
        oList.Clear
        For lCount = lMin To lMax
            
            oList.AddItem Trim(aArr(lCount))
        
        Next
        
    End If

    StatusBar1.Panels("msg").Text = sOldStatusMessage
    Set ofs = Nothing
    Set oTextStream = Nothing
    


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
    sErrSource = VBA.Err.Source & vbCrLf & "frmNetsender->" & csProcName
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
' Name   : SaveListToTextFileP
' Params :
'          oList As Control
'          sFileName As String
'          Optional fSilent As Boolean = False
' Returns: Nothing
' Desc   : The Sub uses parameters oList As Control, sFileName As String and Optional fSilent As Boolean = False for SaveListToTextFileP and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub SaveListToTextFileP(oList As Control, sFileName As String, Optional fSilent As Boolean = False)
    On Error GoTo Proc_Err
    Const csProcName As String = "SaveListToTextFileP"
    
    Dim sOldStatusMessage As String
    Dim ofs As clsFs
    Dim sSaveText As String
    Dim sText As String
    Dim aArr As Variant
    Dim lCount As Long
    Dim lMin As Long
    Dim lMax As Long
    Dim oTextStream As TextStream
    
    
    sOldStatusMessage = StatusBar1.Panels("msg").Text
    
    If Not fSilent Then
        StatusBar1.Panels("msg").Text = "Getting Default Netsend List"
    End If

    'save the list of selected users to the default list
    
    Set ofs = New clsFs
    
    If ofs.FileExists(sFileName) Then
        ofs.DeleteFile (sFileName)
    End If
    
    lMin = 0
    lMax = oList.ListCount - 1
    
    For lCount = lMin To lMax
        sText = sText & oList.List(lCount) & vbCrLf
        
    Next
    
    'remove the last vbcrlf
    sText = Left(sText, Len(sText) - 2)
    
    
    Set oTextStream = ofs.OpenTextFile(sFileName, ForWriting, True)
    oTextStream.Write sText
    
    StatusBar1.Panels("msg").Text = sOldStatusMessage
    
    Set ofs = Nothing
    Set oTextStream = Nothing
    

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
    sErrSource = VBA.Err.Source & vbCrLf & "frmNetsender->" & csProcName
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
' Name   : getGroupsP
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for getGroupsP and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub getGroupsP()

    On Error Resume Next
    
    lstUsersAvailable.Clear
    lstGroups.Clear
    lblUsersAvailable.Caption = "&Users Of " & cboDomain.Text
    
    frmPleaseWait.Show
    
    
    DoEvents
    
    If Me.cboDomain.Text & "" <> "" Then
    
        Dim container As IADsContainer
        Dim containername As String
        containername = cboDomain.Text
        Set container = GetObject("WinNT://" & containername)
        
        container.Filter = Array(gcsUserPlaceholder)
        Dim user As IADsUser
        For Each user In container
        lstUsersAvailable.AddItem user.Name
        Next
        
        container.Filter = Array("Group")
        Dim group As IADsGroup
        For Each group In container
        lstGroups.AddItem group.Name
        Next
        
        Err = 0
        DoEvents
        
    End If
    
    Unload frmPleaseWait


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
'               Auto-Code Commenter
'__________________________________________________
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Proc_Err
    Const csProcName As String = "Form_Unload"

    End


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
    sErrSource = VBA.Err.Source & vbCrLf & "frmNetsender->" & csProcName
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
' Name   : lstGroups_DblClick
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for lstGroups_DblClick and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub lstGroups_DblClick()

    On Error Resume Next
    
    Dim vMember As Variant
    Dim oGroup As IADsGroup
    Dim sGroupName As String
    Dim sGroupDomain As String
    
    lstUsersAvailable.Clear
    lblUsersAvailable.Caption = "&Users Of " & lstGroups.Text
    
    frmPleaseWait.Show
    
    DoEvents
    
    sGroupName = lstGroups.Text
    sGroupDomain = cboDomain.Text
    Set oGroup = GetObject("WinNT://" & sGroupDomain & "/" & sGroupName & ",Group")
    
    For Each vMember In oGroup.Members
        lstUsersAvailable.AddItem vMember.Name
    Next
    
    Err = 0
    
    DoEvents
    
    Unload frmPleaseWait
    
    'set the tab to be on the list of users
    Me.tsTabStripForUsersAvailable.Tabs(1).Selected = True
    
    Set oGroup = Nothing
    

End Sub






'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : lstUsersAvailable_DBLClick
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for lstUsersAvailable_DBLClick and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub lstUsersAvailable_DBLClick()
    On Error GoTo Proc_Err
    Const csProcName As String = "lstUsersAvailable_DBLClick"

    Call cmdAdd_Click


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
    sErrSource = VBA.Err.Source & vbCrLf & "frmNetsender->" & csProcName
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
' Name   : lstUsersSelected_DblClick
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for lstUsersSelected_DblClick and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub lstUsersSelected_DblClick()
    On Error GoTo Proc_Err
    Const csProcName As String = "lstUsersSelected_DblClick"

    Call cmdRemove_Click


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
    sErrSource = VBA.Err.Source & vbCrLf & "frmNetsender->" & csProcName
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
' Name   : Timer1_Timer
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for Timer1_Timer and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub Timer1_Timer()
    On Error GoTo Proc_Err
    Const csProcName As String = "Timer1_Timer"
    lblTotalGroups.Caption = "Total Groups: " & lstGroups.ListCount
    lblTotalUsersAvailable.Caption = "Total Members: " & lstUsersAvailable.ListCount
    
    If cboDomain.Text = "" Then
        cmdGetGroups.Enabled = False
    Else
        cmdGetGroups.Enabled = True
    End If
    
    If lstGroups.ListCount = 0 Then
        cmdPrint.Enabled = False
    Else
        cmdPrint.Enabled = True
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
    sErrSource = VBA.Err.Source & vbCrLf & "frmNetsender->" & csProcName
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
' Name   : tsMainTabstrip_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for tsMainTabstrip_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub tsMainTabstrip_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "tsMainTabstrip_Click"

    HandleTabStripP Me.tsMainTabstrip, Me.fraFrameForTabStrip, True
    


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
    sErrSource = VBA.Err.Source & vbCrLf & "frmNetsender->" & csProcName
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
' Name   : initializeTabStrips
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for initializeTabStrips and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub initializeTabStrips()
    On Error GoTo Proc_Err
    Const csProcName As String = "initializeTabStrips"


    'for some reason, something must run first to allow the getobject below.  so i'll run it AFTER user gets domains


    Dim vMember As Variant
    Dim oGroup As IADsGroup
    Dim sGroupName As String
    Dim sGroupDomain As String
    Dim sUserName As String
    Dim fAdminRights As Boolean
    
    
    CreateTabsFromContainerP Me.tsMainTabstrip, Me.fraFrameForTabStrip
    tsMainTabstrip_Click
    
    CreateTabsFromContainerP Me.tsTabStripForUsersAvailable, Me.fraFrameForUsersAvailable
    tsTabStripForUsersAvailable_Click
    
    DoEvents
    

    'see if the current user is an admin, if so, set the print tab to enabled otherwise disable it
    sUserName = GetUserNamePv
    
    sGroupName = "Administrators"
    sGroupDomain = GetDomainName
    Set oGroup = GetObject("WinNT://" & sGroupDomain & "/" & sGroupName & ",Group")
    
    For Each vMember In oGroup.Members
        If vMember.Name = sUserName Then
            fAdminRights = True
            Exit For
        End If
    Next
    
    DoEvents
    
    If Not fAdminRights Then
        'remove the print tab
        Me.tsMainTabstrip.Tabs.Remove (2)
        Me.fraFrameForTabStrip(2).Enabled = False
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
    sErrSource = VBA.Err.Source & vbCrLf & "frmNetsender->" & csProcName
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
' Name   : tmrDoNetSend_Timer
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for tmrDoNetSend_Timer and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub tmrDoNetSend_Timer()
    On Error GoTo Proc_Err
    Const csProcName As String = "tmrDoNetSend_Timer"
    
    Dim lReturnCode As Long
    Dim sUnicodeToName As String
    Dim sUnicodeFromName As String
    Dim sUnicodeMessage As String
    Dim lMessageLength As Long
    Dim sUserName As String
    Static fSavedList As Boolean
    Static sTempFile As String
    Dim lCount As Long
    Dim lMax As Long
    Dim sTemp As String
    
        
    If Not chkClearListAfterNetsend = vbChecked Then
        If Not fSavedList Then
            sTempFile = App.Path & Format(Date, "YYYYMMDD") & Format(Time, "hhmmss") & "TempList.txt"
            Call SaveListToTextFileP(Me.lstUsersSelected, sTempFile, True)
            fSavedList = True
        End If
    End If
        
    'Debug.Print "timer fired at " & Time
    
    lblNumUsersSelected.Caption = lstUsersSelected.ListCount - 1
    
    If lblNumUsersSelected.Caption = "-1" Then
        
        If chkClearMessageAfterNetsend Then
            txtMessage.Text = vbNullString
        End If
        
        lblNumUsersSelected.Caption = ""
        cmdCancel.Visible = False
        cmdSend.Visible = True
        tmrDoNetSend.Enabled = False
        'Timer1.Enabled = True
        StatusBar1.Panels("msg").Text = "Finished"
        
        Unload frmPleaseWait
        Call getGroupsP
        
        If fSavedList Then
            fSavedList = False
            Call GetListFromTextFileP(Me.lstUsersSelected, sTempFile, True)
            Kill sTempFile
        End If
        

    Else
    
        sUnicodeFromName = StrConv(GetLocalSystemName, vbUnicode)
        sUnicodeToName = StrConv((lstUsersSelected.List(0)), vbUnicode)
        
        sUserName = GetUserNamePv
        sUnicodeMessage = txtMessage.Text & vbCrLf & vbCrLf & vbCrLf & _
                        "___________________________________________________________" & _
                        vbCrLf & vbCrLf & _
                        App.ProductName & " used by: " & sUserName & _
                        vbCrLf & _
                        "___________________________________________________________"
                        
        'include disclaimer or not?
        If chkIncludeDisclaimer = vbChecked Then
                  sTemp = Replace(App.LegalCopyright, gcsCompanyPlaceholder, App.CompanyName)
                  sTemp = Replace(sTemp, gcsUserPlaceholder, sUserName)
                  sUnicodeMessage = sUnicodeMessage & vbCrLf & vbCrLf & sTemp
        End If
        
        sUnicodeMessage = StrConv(sUnicodeMessage, vbUnicode)
                        
        lMessageLength = Len(sUnicodeMessage)
        
        lstUsersSelected.RemoveItem (0)
         
        If Not IsLoaded("frmPleaseWait") Then
            frmPleaseWait.Show
        End If

        StatusBar1.Panels("msg").Text = vbNullString
        
    
        lReturnCode = NetMessageBufferSend("", _
                        sUnicodeToName, _
                        sUnicodeFromName, _
                        sUnicodeMessage, _
                        lMessageLength)
    
    
        If lReturnCode = 0 Then
            StatusBar1.Panels("msg").Text = _
            "Your message was sent correctly..."
        Else
            StatusBar1.Panels("msg").Text = "Error..."
           
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
    sErrSource = VBA.Err.Source & vbCrLf & "frmNetsender->" & csProcName
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
' Name   : tmrGetTimeOfDay_Timer
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for tmrGetTimeOfDay_Timer and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub tmrGetTimeOfDay_Timer()
    On Error GoTo Proc_Err
    Const csProcName As String = "tmrGetTimeOfDay_Timer"
    StatusBar1.Panels(2).Text = Time

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
    sErrSource = VBA.Err.Source & vbCrLf & "frmNetsender->" & csProcName
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
' Name   : cmdSend_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdSend_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub cmdSend_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "cmdSend_Click"
    
    Dim lReturnCode As Long
    Dim sUnicodeToName As String
    Dim sUnicodeFromName As String
    Dim sUnicodeMessage As String
    Dim lMessageLength As Long
    Dim lResponse As Long
    Dim sMsg As String
    Dim fBeginNetsend As Boolean
    
    'if noone was selected (listcount = 0)then user eitherwants to
    '   send a broadcast message to teh domain, or forgot to choose a user
    
    If lstUsersSelected.ListCount = 0 Then
    
        sMsg = "You have not chosen any specific users, would you like to send a broadcast message to " _
            & Me.cboDomain & " domain?"
        lResponse = MsgBox(sMsg, vbQuestion + vbYesNoCancel, App.Title)
        
        If lResponse <> vbYes Then
            StatusBar1.Panels("msg").Text = "Please chose a specific user or users."
        Else
            
            sMsg = "_______" & vbCrLf & vbCrLf & "WARNING" & vbCrLf & _
                    "_______" & vbCrLf & vbCrLf & _
                    "You are about to send a broadcast message to the " & Me.cboDomain.Text & _
                    " domain.  This will send your message to all users in this domain." & _
                    vbCrLf & vbCrLf & "Do you want to send this message?"
            lResponse = MsgBox(sMsg, vbYesNoCancel)
            
            If lResponse = vbYes Then
                'add the domain to the selected users listbox, then send the message
                Me.lstUsersSelected.AddItem Me.cboDomain.Text
                fBeginNetsend = True
            End If
                
        End If
        
    Else
        fBeginNetsend = True
    End If
    
    
    If fBeginNetsend Then
        'turn on the cancel button and start the netsend timer
        cmdCancel.Visible = True
        cmdSend.Visible = False
        tmrDoNetSend.Enabled = True
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
    sErrSource = VBA.Err.Source & vbCrLf & "frmNetsender->" & csProcName
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
' Name   : cmdCancel_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdCancel_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub cmdCancel_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "cmdCancel_Click"
    
    lstUsersSelected.Clear
    Call getGroupsP
    
    tmrDoNetSend.Enabled = False
    If chkClearMessageAfterNetsend Then
        txtMessage.Text = vbNullString
    End If
    lblNumUsersSelected.Caption = ""
    cmdCancel.Visible = False
    cmdSend.Visible = True
    StatusBar1.Panels("msg").Text = "Cancelled by " + GetUserNamePv
    
   

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
    sErrSource = VBA.Err.Source & vbCrLf & "frmNetsender->" & csProcName
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
' Name   : cmdAddAll_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdAddAll_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub cmdAddAll_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "cmdAddAll_Click"

    Dim lCount As Long
    Dim lMax As Long
    
    lMax = Me.lstUsersAvailable.ListCount - 1
    
    For lCount = 0 To lMax
        Me.lstUsersSelected.AddItem Me.lstUsersAvailable.List(lCount)
    Next
    
    'Me.lstUsersAvailable.Clear
    
    'Me.lblNumUsersAvailable = Me.lstUsersAvailable.ListCount
    Me.lblNumUsersSelected = Me.lstUsersSelected.ListCount
        

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
    sErrSource = VBA.Err.Source & vbCrLf & "frmNetsender->" & csProcName
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
' Name   : cmdAdd_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdAdd_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub cmdAdd_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "cmdAdd_Click"
    
    Dim lCount As Long
    Dim lMax As Long
    
    If lstUsersAvailable.ListCount <> 0 And Me.lstUsersAvailable.SelCount <> 0 Then
            
        lMax = Me.lstUsersAvailable.ListCount - 1
        For lCount = lMax To 0 Step -1
            If Me.lstUsersAvailable.Selected(lCount) Then
                Me.lstUsersSelected.AddItem Me.lstUsersAvailable.List(lCount)
                'Me.lstUsersAvailable.RemoveItem (lcount)
            End If
        Next
        
        'Me.lblNumUsersAvailable = Me.lstUsersAvailable.ListCount
        Me.lblNumUsersSelected = Me.lstUsersSelected.ListCount
        
    End If
    
    'set focus back to the list after clicking add
    Me.lstUsersAvailable.SetFocus
    

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
    sErrSource = VBA.Err.Source & vbCrLf & "frmNetsender->" & csProcName
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
' Name   : cmdRemoveAll_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdRemoveAll_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub cmdRemoveAll_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "cmdRemoveAll_Click"
    Dim lCount As Long
    Dim lMax As Long
    
    Me.lstUsersSelected.Clear
    
    Call getGroupsP  'cmdRefreshList_Click
    
    

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
    sErrSource = VBA.Err.Source & vbCrLf & "frmNetsender->" & csProcName
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
' Name   : cmdRemove_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdRemove_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub cmdRemove_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "cmdRemove_Click"
    
    Dim lCount As Long
    Dim lMax As Long
    
    If lstUsersSelected.ListCount <> 0 And Me.lstUsersSelected.SelCount <> 0 Then
            
        lMax = Me.lstUsersSelected.ListCount - 1
        For lCount = lMax To 0 Step -1
            If Me.lstUsersSelected.Selected(lCount) Then
                'Me.lstUsersAvailable.AddItem Me.lstUsersSelected.List(lcount)
                Me.lstUsersSelected.RemoveItem (lCount)
            End If
        Next
        
        'Me.lblNumUsersAvailable = Me.lstUsersAvailable.ListCount
        Me.lblNumUsersSelected = Me.lstUsersSelected.ListCount
        
    End If
    
    'set focus back to users selected so they can remove more
    Me.lstUsersSelected.SetFocus
    
    

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
    sErrSource = VBA.Err.Source & vbCrLf & "frmNetsender->" & csProcName
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
' Name   : cmdExit_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdExit_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub cmdExit_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "cmdExit_Click"

    Unload Me
    End


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
    sErrSource = VBA.Err.Source & vbCrLf & "frmNetsender->" & csProcName
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
' Type   : Function
' Name   : GetTemplateFileName
' Params :
'          oCdlg As CommonDialog
'          Optional eMode As eCommonDialogConst = ecdlShowOpen
'          Optional sDefaultTemplate As String = ""
'          Optional sFilter As String = ""
' Returns: String
' Desc   : The Function uses parameters oCdlg As CommonDialog, Optional eMode As eCommonDialogConst = ecdlShowOpen, Optional sDefaultTemplate As String = "" and Optional sFilter As String = "" for GetTemplateFileName and returns String.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Function GetTemplateFileName(oCdlg As CommonDialog, Optional eMode As eCommonDialogConst = ecdlShowOpen, Optional sDefaultTemplate As String = "", Optional sFilter As String = "") As String

    On Error GoTo Proc_Exit
    Dim sFile As String
    Dim fFileExists As Boolean
    Dim lResponse As Long
    
    oCdlg.DialogTitle = "Template"
    oCdlg.InitDir = App.Path
    oCdlg.DefaultExt = ".txt"
    oCdlg.FileName = sDefaultTemplate
    oCdlg.Filter = sFilter
    oCdlg.CancelError = True
    oCdlg.Flags = cdlOFNHideReadOnly
    oCdlg.Action = eMode
    
    
    
    Select Case eMode
        Case eCommonDialogConst.ecdlShowSave
            If Len(oCdlg.FileName) = 0 Then
                'no file typed in
                Err.Raise Number:=4001, Description:="no file selected"
            
            Else
                sFile = oCdlg.FileName
                fFileExists = CBool(Len(Dir(sFile)))
                If fFileExists Then
                    lResponse = MsgBox("File already exists, overwrite existing file?", vbQuestion + vbYesNoCancel, App.Title)
                    If lResponse <> vbYes Then
                        sFile = ""
                    End If
                End If
            End If
        Case eCommonDialogConst.ecdlShowOpen
            If Len(oCdlg.FileName) > 0 And Len(Dir(oCdlg.FileName)) > 0 Then
                sFile = oCdlg.FileName
            Else
                Err.Raise Number:=4001, Description:="Cannot find file"
            End If
            
    End Select
    
Proc_Exit:

    GetTemplateFileName = sFile

End Function

    
'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : tsTabStripForUsersAvailable_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for tsTabStripForUsersAvailable_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'               Auto-Code Commenter
'__________________________________________________
Private Sub tsTabStripForUsersAvailable_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "tsTabStripForUsersAvailable_Click"

    HandleTabStripP Me.tsTabStripForUsersAvailable, Me.fraFrameForUsersAvailable, True
    

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
    sErrSource = VBA.Err.Source & vbCrLf & "frmNetsender->" & csProcName
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

