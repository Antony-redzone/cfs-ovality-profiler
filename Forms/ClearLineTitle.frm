VERSION 5.00
Begin VB.Form ClearLineTitle 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11235
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   11235
   ShowInTaskbar   =   0   'False
   Begin VB.Shape PanelEdgeLine 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   3495
      Left            =   10440
      Top             =   0
      Width           =   75
   End
   Begin VB.Label TitleBarCaption 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "PV Screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   11055
   End
End
Attribute VB_Name = "ClearLineTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo Err_Handler

Me.TitleBarCaption.Font.Charset = LanguageCharset

Me.Top = 0
Me.Left = 0
Me.height = ClearLineScreen.Top
Me.width = ClearLineScreen.width
Me.BackColor = RGB(172, 196, 231) 'PCN4171

Me.PanelEdgeLine.Left = ClearLineScreen.width - 10  'PCN4171
Me.PanelEdgeLine.Top = 0
Me.PanelEdgeLine.BackColor = RGB(54, 106, 179) 'PCN4171

Exit Sub
Err_Handler:
    MsgBox Err & "-CLT1:" & Error$
End Sub
