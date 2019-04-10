VERSION 5.00
Begin VB.Form CLPProgressBar 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Processing"
   ClientHeight    =   855
   ClientLeft      =   195
   ClientTop       =   195
   ClientWidth     =   5850
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   Begin VB.Timer ProgressBarTimer 
      Interval        =   5000
      Left            =   5400
      Top             =   0
   End
   Begin VB.PictureBox ProgressStatus 
      Height          =   400
      Left            =   120
      Picture         =   "CLPProgressBar.frx":0000
      ScaleHeight     =   345
      ScaleWidth      =   5595
      TabIndex        =   0
      Top             =   360
      Width           =   5655
      Begin VB.PictureBox Blanking 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   840
         ScaleHeight     =   375
         ScaleWidth      =   4995
         TabIndex        =   1
         Top             =   0
         Width           =   5000
         Begin VB.Label PercentageLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "0%"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   50
            TabIndex        =   3
            Top             =   35
            Width           =   855
         End
      End
      Begin VB.Line Line5 
         BorderColor     =   &H000040C0&
         X1              =   510
         X2              =   720
         Y1              =   120
         Y2              =   210
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   0
         X2              =   5655
         Y1              =   172
         Y2              =   172
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         X1              =   600
         X2              =   600
         Y1              =   0
         Y2              =   360
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000000FF&
         X1              =   480
         X2              =   720
         Y1              =   50
         Y2              =   290
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000000FF&
         X1              =   480
         X2              =   720
         Y1              =   290
         Y2              =   50
      End
      Begin VB.Line Line6 
         BorderColor     =   &H000000FF&
         X1              =   510
         X2              =   720
         Y1              =   200
         Y2              =   120
      End
      Begin VB.Line Line7 
         BorderColor     =   &H000000FF&
         X1              =   555
         X2              =   645
         Y1              =   70
         Y2              =   260
      End
      Begin VB.Line Line8 
         BorderColor     =   &H000000FF&
         X1              =   565
         X2              =   640
         Y1              =   260
         Y2              =   65
      End
   End
   Begin VB.Label ProgressTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   150
      TabIndex        =   2
      Top             =   45
      Width           =   5535
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00400000&
      X1              =   5830
      X2              =   5830
      Y1              =   0
      Y2              =   1920
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00400000&
      X1              =   5880
      X2              =   0
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00400000&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1920
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00400000&
      X1              =   0
      X2              =   5880
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "CLPProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Public TimerWatchdog As Double

Private Sub Form_Load()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : Form_Load
'Created : 17 November 2003, PCN2401
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Me.ProgressTitle.Font.Charset = LanguageCharset

TimerWatchdog = 0
Blanking.BackColor = &H400000

CLPProgressBar.Top = 8050
CLPProgressBar.Left = 2750

ProgressBarTimer.Enabled = True

Exit Sub
Err_Handler:
    MsgBox Err & "-PB1:" & Error$
End Sub

Private Sub ProgressBarTimer_Timer()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : ProgressBarTimer_Timer
'Created : 17 November 2003, PCN2401
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    : Ensures the progress bar remains in view
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
'Dim PercentFinished As Double

CLPProgressBar.ZOrder 0

Exit Sub
Err_Handler:
    MsgBox Err & "-PB2:" & Error$
End Sub

Sub ProgressBarInitialise(Title As String)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : ProgressBarInitialise
'Created : 17 November 2003, PCN2401
'Updated :
'Prg By  : Geoff Logan
'Param   : Title - The title caption of the progress bar
'Desc    : Initialises the progress bar to 0%
'Usage   : Call ProgressBarInitialise("Loading data") is a standard example.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

'Setup progress title
ProgressTitle.Caption = Title
'Setup laser at start.
Line2.x1 = 0
Line2.x2 = 0

Line3.x1 = -120
Line3.x2 = 120
Line3.y1 = 50
Line3.y2 = 290

Line4.x1 = -120
Line4.x2 = 120
Line4.y1 = 290
Line4.y2 = 50

Line5.x1 = -90
Line5.x2 = 120
Line5.y1 = 120
Line5.y2 = 210

Line6.x1 = -90
Line6.x2 = 120
Line6.y1 = 200
Line6.y2 = 120

Line7.x1 = -45
Line7.x2 = 45
Line7.y1 = 70
Line7.y2 = 260

Line8.x1 = -35
Line8.x2 = 40
Line8.y1 = 260
Line8.y2 = 65

'Setup Blanking bar
Blanking.width = ProgressStatus.width
Blanking.Left = 0

Exit Sub
Err_Handler:
    MsgBox Err & "-PB3:" & Error$
End Sub


Sub ProgressBarPosition(ByVal PercentComplete As Double)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : ProgressBarPosition
'Created : 17 November 2003, PCN2401
'Updated :
'Prg By  : Geoff Logan
'Param   : PercentComplete - The percent complete to set the progress bar
'Desc    : Sets the position of the components of the progress bar to represent
'           the percent complete. Also sets the PercentageLabel caption.
'Usage   : Call ProgressBarPosition(0.55) is a standard example. This call will
'           setup the progress bar for 55% completed.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim ProgressPos As Integer

ProgressPos = (ProgressStatus.width - 170) * PercentComplete

'Setup laser
Line2.x1 = ProgressPos
Line2.x2 = ProgressPos

Line3.x1 = ProgressPos - 120
Line3.x2 = ProgressPos + 120

Line4.x1 = ProgressPos - 120
Line4.x2 = ProgressPos + 120

Line5.x1 = ProgressPos - 90
Line5.x2 = ProgressPos + 120

Line6.x1 = ProgressPos - 90
Line6.x2 = ProgressPos + 120

Line7.x1 = ProgressPos - 45
Line7.x2 = ProgressPos + 45

Line8.x1 = ProgressPos - 35
Line8.x2 = ProgressPos + 40

'Setup Blanking bar
Blanking.Left = ProgressPos + 170
PercentageLabel.Caption = Int(PercentComplete * 100) & "%"

If PercentComplete >= 1 Then
    ProgressBarTimer.Enabled = False
    Unload CLPProgressBar
    DoEvents
    Exit Sub
End If


Exit Sub
Err_Handler:
    MsgBox Err & "-PB4:" & Error$
End Sub


