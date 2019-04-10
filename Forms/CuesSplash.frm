VERSION 5.00
Begin VB.Form CuesSplash 
   BackColor       =   &H00C00000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6465
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4530
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "CuesSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   60
      TabIndex        =   0
      Top             =   50
      Width           =   4410
      Begin VB.Timer BringToFrontTimer 
         Interval        =   50
         Left            =   3540
         Top             =   5400
      End
      Begin VB.CommandButton cmdRegister 
         Caption         =   "Register Now"
         Height          =   450
         Left            =   1200
         TabIndex        =   9
         Top             =   5040
         Width           =   1935
      End
      Begin VB.Timer SplashDelayTimer 
         Enabled         =   0   'False
         Interval        =   4000
         Left            =   3540
         Top             =   4800
      End
      Begin VB.CommandButton CloseSplash 
         Height          =   615
         Left            =   1800
         Picture         =   "CuesSplash.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4920
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label TradeMarkLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "TM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   3960
         TabIndex        =   10
         Top             =   1500
         Width           =   255
      End
      Begin VB.Label lblRegistered 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Unregistered Copy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   4440
         Width           =   4215
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
         Left            =   2700
         TabIndex        =   7
         Top             =   5700
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
         Left            =   2460
         TabIndex        =   6
         Top             =   5700
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
         Left            =   2220
         TabIndex        =   5
         Top             =   5700
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
         Left            =   180
         TabIndex        =   4
         Top             =   5700
         Width           =   1935
      End
      Begin VB.Image SplashImage 
         Height          =   1605
         Left            =   210
         Picture         =   "CuesSplash.frx":1CD6
         Top             =   2520
         Width           =   4065
      End
      Begin VB.Image Image1 
         Height          =   1575
         Left            =   120
         Picture         =   "CuesSplash.frx":17228
         Stretch         =   -1  'True
         Top             =   120
         Width           =   4140
      End
      Begin VB.Label lblCompany 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright Cleanflow Systems 2003"
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
         Left            =   120
         TabIndex        =   1
         Top             =   6000
         Width           =   4215
      End
      Begin VB.Label lblProductName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Standard Edition"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   4200
      End
      Begin VB.Shape Shape1 
         Height          =   6370
         Left            =   0
         Top             =   0
         Width           =   4410
      End
   End
End
Attribute VB_Name = "CuesSplash"
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

CuesSplash.ZOrder 0

Exit Sub
Err_Handler:
    MsgBox Err & error$
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
    MsgBox Err & error$
End Sub



Private Sub Form_Load()
On Error GoTo Err_Handler
ConvertLanguage Me, Language 'PCN2111

Me.left = ClearLineProfilerV5.width / 2 - Me.width / 2
Me.Top = ClearLineProfilerV5.height / 2 - Me.height / 2
lblMajor.Caption = App.Major & " ."
lblMinor.Caption = App.Minor & " ."
lblRevision.Caption = App.Revision
If Registered = True Then
    lblRegistered.Visible = False
    cmdRegister.Visible = False
    SplashDelayTimer.Interval = 1750 'PCNGL230104
    CloseSplash.Visible = True
Else
    SplashDelayTimer.Interval = 5000 'PCNGL230103
End If
SplashDelayTimer.Enabled = True 'Start delay once the form is load, better delay control 'PCNGL230103


Exit Sub
Err_Handler:
    MsgBox Err & error$
End Sub



Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo Err_Handler

    Screen.MousePointer = vbDefault

Exit Sub
Err_Handler:
MsgBox Err & " - " & error$

End Sub

Private Sub Label3_Click()
   
On Error GoTo Err_Handler
   
   Dim sURL As String
   sURL = "www.cuesinc.com"
   Call ShellExecute(Me.hwnd, "open", sURL, "", LocToSave, SW_SHOWNORMAL) 'PCN2155

Exit Sub
Err_Handler:
MsgBox Err & " - " & error$

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
        MsgBox Err & " - " & error$
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
    MsgBox Err & " - " & error$
End Sub

