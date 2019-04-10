VERSION 5.00
Begin VB.Form ControlsReports 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1080
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3495
   ScaleWidth      =   1080
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton BtnReportProfilex3 
      BackColor       =   &H00C6C7C6&
      Caption         =   "Multi Report x3"
      Height          =   855
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton BtnReportProfile 
      BackColor       =   &H00C6C7C6&
      Caption         =   "Profile"
      Height          =   855
      Left            =   0
      Picture         =   "ControlsReports.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton BtnReport4in1 
      BackColor       =   &H00C6C7C6&
      Caption         =   "4 in 1"
      Height          =   855
      Left            =   0
      Picture         =   "ControlsReports.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton BtnReportPVGraph 
      BackColor       =   &H00C6C7C6&
      Caption         =   "PV Graph"
      Height          =   975
      Left            =   0
      Picture         =   "ControlsReports.frx":3994
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "ControlsReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
