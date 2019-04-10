VERSION 5.00
Begin VB.Form Confirm1KDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Project Report"
   ClientHeight    =   2340
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Begin VB.Frame ReportTypeFrame 
      Caption         =   "Report type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6975
      Begin VB.OptionButton Flat1K 
         Caption         =   "Flat"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton FlatOvality1K 
         Caption         =   "Flat Ovality"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   5
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame SpanSize 
      Caption         =   "Span"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   6975
      Begin VB.OptionButton SpanOption 
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1296
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton SpanOption 
         Caption         =   "75"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2472
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton SpanOption 
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   3648
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton SpanOption 
         Caption         =   "125"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4824
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton SpanOption 
         Caption         =   "150"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   6000
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton SpanOption 
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   1455
   End
End
Attribute VB_Name = "Confirm1KDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub CancelButton_Click()
    Me.Visible = False
End Sub

Private Sub Confirm1kLabel_Click()

End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim Span As Integer
    Dim FlatLabel As String 'PCN4974
    
    If MeasurementUnits = "mm" Then
        Me.SpanSize.Caption = DisplayMessage("Span in meters")
       Span = 25
    Else
        Me.SpanSize.Caption = DisplayMessage("Span in feet")
        Span = 125
    End If
    
    For i = 0 To 5
        Me.SpanOption(i).Caption = (i + 1) * Span
    Next i
    Call ConvertLanguage(Me, Language) 'PCN4171
    
    If MedianFlat And PVDFileName <> "" Then FlatLabel = DisplayMessage("Deflection Flat") Else FlatLabel = DisplayMessage("Flat")  'PCN4974
    
    Me.Caption = DisplayMessage("Project Report")
    Me.Flat1K.Caption = FlatLabel ' DisplayMessage("Flat") PCN4974
    
    Me.ReportTypeFrame.Caption = DisplayMessage("Report type")
    Me.FlatOvality1K.Caption = FlatLabel & " " & DisplayMessage("Ovality") 'was DisplayMessage("Flat") & " " & DisplayMessage("Ovality") 'PCN4974
    Me.Caption = DisplayMessage("Project Report")
    
    
    

    Me.ZOrder 0
    Me.Visible = True
End Sub

Private Sub OKButton_Click()
 Me.Visible = False
 If Me.Flat1K.value = True Then Load PVReport1K
 If Me.FlatOvality1K.value = True Then Load PVReport2in1

End Sub
