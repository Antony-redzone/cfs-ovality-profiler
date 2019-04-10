VERSION 5.00
Begin VB.Form V6Splash 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6825
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   9420
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "V6Splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   60
      TabIndex        =   0
      Top             =   50
      Width           =   9315
      Begin VB.Timer BringToFrontTimer 
         Interval        =   50
         Left            =   6060
         Top             =   5760
      End
      Begin VB.CommandButton cmdRegister 
         Caption         =   "Register Now"
         Height          =   450
         Left            =   3720
         TabIndex        =   8
         Top             =   5400
         Width           =   1935
      End
      Begin VB.Timer SplashDelayTimer 
         Enabled         =   0   'False
         Interval        =   4000
         Left            =   6060
         Top             =   5160
      End
      Begin VB.CommandButton CloseSplash 
         Height          =   615
         Left            =   4320
         Picture         =   "V6Splash.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5280
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblWebsite 
         BackStyle       =   0  'Transparent
         Caption         =   "cleanflowsystems.com"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   375
         Left            =   5520
         TabIndex        =   9
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Image Image2 
         Height          =   1455
         Left            =   4920
         Picture         =   "V6Splash.frx":1CD6
         Stretch         =   -1  'True
         Top             =   1680
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Image Image1 
         Height          =   1575
         Left            =   4920
         Picture         =   "V6Splash.frx":7C1F
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4335
      End
      Begin VB.Image Image3 
         Height          =   6075
         Left            =   120
         Picture         =   "V6Splash.frx":A9A7
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label lblRegistered 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Unregistered Copy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   0
         TabIndex        =   7
         Top             =   4800
         Width           =   9375
      End
      Begin VB.Label lblRevision 
         BackStyle       =   0  'Transparent
         Caption         =   "Revision"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5220
         TabIndex        =   6
         Top             =   6060
         Width           =   255
      End
      Begin VB.Label lblMinor 
         BackStyle       =   0  'Transparent
         Caption         =   "Minor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4980
         TabIndex        =   5
         Top             =   6060
         Width           =   255
      End
      Begin VB.Label lblMajor 
         BackStyle       =   0  'Transparent
         Caption         =   "Major"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4740
         TabIndex        =   4
         Top             =   6060
         Width           =   255
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2700
         TabIndex        =   3
         Top             =   6060
         Width           =   1935
      End
      Begin VB.Label lblCompany 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright Cleanflow Systems 2003 - 2006"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   1
         Top             =   6360
         Width           =   4215
      End
      Begin VB.Image Background 
         Height          =   7020
         Left            =   0
         Picture         =   "V6Splash.frx":B0BB
         Stretch         =   -1  'True
         Top             =   -240
         Width           =   9405
      End
   End
   Begin VB.Image BackgroundImageArray 
      Height          =   855
      Index           =   6
      Left            =   6480
      Picture         =   "V6Splash.frx":1554A
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Image BackgroundImageArray 
      Height          =   855
      Index           =   8
      Left            =   8760
      Picture         =   "V6Splash.frx":1EA3A
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Image BackgroundImageArray 
      Height          =   855
      Index           =   7
      Left            =   7680
      Picture         =   "V6Splash.frx":28EC9
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1080
   End
   Begin VB.Image BackgroundImageArray 
      Height          =   855
      Index           =   5
      Left            =   5160
      Picture         =   "V6Splash.frx":2E940
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Image BackgroundImageArray 
      Height          =   855
      Index           =   4
      Left            =   3885
      Picture         =   "V6Splash.frx":4043B
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Image BackgroundImageArray 
      Height          =   855
      Index           =   3
      Left            =   2640
      Picture         =   "V6Splash.frx":4992B
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Image BackgroundImageArray 
      Height          =   855
      Index           =   2
      Left            =   1395
      Picture         =   "V6Splash.frx":59811
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Image BackgroundImageArray 
      Height          =   855
      Index           =   1
      Left            =   120
      Picture         =   "V6Splash.frx":6374F
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1215
   End
End
Attribute VB_Name = "V6Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Const SW_SHOWNORMAL = 1
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
  "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As _
  String, ByVal lpFile As String, ByVal lpParameters As String, _
  ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub BringToFrontTimer_Timer() 'PCNGL300103
On Error GoTo Err_Handler

V6Splash.ZOrder 0

Exit Sub
Err_Handler:
    MsgBox Err & "-SPL1:" & Error$
End Sub

Private Sub CloseSplash_Click()
    Unload Me
End Sub

Private Sub cmdRegister_Click()
On Error GoTo Err_Handler

Registration.Enabled = True
Unload Me 'PCNGL030203
    
Exit Sub
Err_Handler:
    MsgBox Err & "-SPL2:" & Error$
End Sub



Private Sub Form_Load()
On Error GoTo Err_Handler
ConvertLanguage Me, Language 'PCN2111

'PCN4187 - Adding random images for splash background^^^^^^^^^^^^^^
Randomize
Dim RandomArray(8) As Integer
Dim cnt As Integer

cnt = 0

cnt = ((8 - 1 + 1) * Rnd) + 1

If cnt > 8 Or cnt < 1 Then cnt = 8

Set Me.Background = Me.BackgroundImageArray(cnt)

If cnt = 6 Then
    Image2.Visible = True
    Image1.Visible = False
    Image2.Top = 0
    lblWebsite.ForeColor = RGB(170, 172, 171)
End If

'PCN4187VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV


Me.Left = ClearLineProfilerV6.width / 2 - Me.width / 2
Me.Top = ClearLineProfilerV6.height / 2 - Me.height / 2
lblMajor.Caption = App.Major & " ."
lblMinor.Caption = App.Minor & " ."
lblRevision.Caption = App.Revision
If Registered = True Then
    lblRegistered.Visible = False
    cmdRegister.Visible = False
    SplashDelayTimer.Interval = 1750 'PCNGL230104
    CloseSplash.Visible = True
ElseIf SoftwareConfiguration = "Reader" Then
    lblRegistered.Caption = DisplayMessage("Viewer Copy") 'PCN4297
    lblRegistered.FontSize = 18
    lblRegistered.Top = 5500
    cmdRegister.Visible = False
Else
    SplashDelayTimer.Interval = 5000 'PCNGL230103
End If
SplashDelayTimer.Enabled = True 'Start delay once the form is load, better delay control 'PCNGL230103


Exit Sub
Err_Handler:
    MsgBox Err & "-SPL3:" & Error$
End Sub



Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo Err_Handler

    Screen.MousePointer = vbDefault

Exit Sub
Err_Handler:
MsgBox Err & "-SPL4:" & Error$

End Sub

Private Sub Label3_Click()
   
On Error GoTo Err_Handler
   
   Dim sURL As String
   sURL = "www.cuesinc.com"
   Call ShellExecute(Me.hwnd, "open", sURL, "", LocToSave, SW_SHOWNORMAL) 'PCN2155

Exit Sub
Err_Handler:
MsgBox Err & "-SPL5:" & Error$

End Sub



Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo Err_Handler

    Screen.MousePointer = 5 'PCNGL291102
    'Screen.MouseIcon = LoadPicture(App.Path & MainScreenMouseIcon) 'PCNGL291102

Exit Sub
Err_Handler:
Select Case Err
    Case 53 'Can't find mouse icon
    Case Else
        MsgBox Err & "-SPL6:" & Error$
End Select
End Sub


Private Sub SplashDelayTimer_Timer() 'PCNGL100103
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'SplashDelayTimer_Timer Sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    10/01/03     Building initial framework
'
'Description:
'   If this is a registered copy of the application then close this form
'   after a suitable delay.
'
'Purpose:
'   Close form after the expiry of the timer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
   
SplashDelayTimer.Enabled = False 'PCNGL300103
DoEvents 'PCNGL300103
'Message to remind unregistered user to register the product
'If Registered = False Then
'    MsgBox "Please remember to register the ClearLine Profiler software", vbInformation, "Registration reminder"
'End If
Unload Me

Exit Sub
Err_Handler:
    MsgBox Err & "-SPL7:" & Error$
End Sub

