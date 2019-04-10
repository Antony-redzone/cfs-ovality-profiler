VERSION 5.00
Begin VB.Form DistanceChangeForm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3150
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   990
   ScaleWidth      =   3150
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00EFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox CurrentDistanceTextBox 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0F0F0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label NewDistanceLabel 
      BackColor       =   &H00C0FFFF&
      Caption         =   "New Distance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   300
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label CurrentDistanceLabel 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Current Distance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "DistanceChangeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
