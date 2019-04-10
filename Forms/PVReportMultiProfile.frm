VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form PVReportMultiProfilex8 
   Caption         =   "Multi Profile Report"
   ClientHeight    =   11910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13005
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11910
   ScaleWidth      =   13005
   Begin VB.CommandButton CmdText 
      Height          =   615
      Left            =   1440
      Picture         =   "PVReportMultiProfile.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton CmdMove 
      Height          =   615
      Left            =   2280
      Picture         =   "PVReportMultiProfile.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   0
      Width           =   735
   End
   Begin VB.PictureBox PageFramePictureBox 
      Height          =   12975
      Left            =   720
      ScaleHeight     =   12915
      ScaleWidth      =   17835
      TabIndex        =   5
      Top             =   600
      Width           =   17895
      Begin VB.PictureBox picReportPage 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   11904
         Left            =   240
         ScaleHeight     =   11850
         ScaleWidth      =   16770
         TabIndex        =   6
         Tag             =   "Paper"
         Top             =   120
         Width           =   16834
         Begin VB.TextBox FloatingText 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   2640
            TabIndex        =   27
            Top             =   0
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox FloatingTextDefault 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2640
            TabIndex        =   26
            Text            =   "Default Text Setting"
            Top             =   360
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Shape ShapeArray 
            BackColor       =   &H00000000&
            BackStyle       =   1  'Opaque
            BorderWidth     =   2
            FillStyle       =   0  'Solid
            Height          =   15
            Index           =   0
            Left            =   3000
            Shape           =   1  'Square
            Top             =   11280
            Visible         =   0   'False
            Width           =   15
         End
         Begin VB.Image GraphContainer 
            Appearance      =   0  'Flat
            Height          =   1215
            Left            =   960
            Stretch         =   -1  'True
            Top             =   5280
            Width           =   14865
         End
         Begin VB.Line GraphYDevisionLine 
            Index           =   16
            X1              =   840
            X2              =   15840
            Y1              =   6480
            Y2              =   6480
         End
         Begin VB.Image GraphXScaleContainer 
            Height          =   375
            Left            =   960
            Top             =   6480
            Width           =   14865
         End
         Begin VB.Line Line2 
            X1              =   960
            X2              =   960
            Y1              =   5280
            Y2              =   6600
         End
         Begin VB.Image CLPLogoImage 
            Height          =   705
            Left            =   360
            Picture         =   "PVReportMultiProfile.frx":3994
            Stretch         =   -1  'True
            Top             =   10800
            Width           =   2010
         End
         Begin VB.Image LogoImage 
            Height          =   855
            Left            =   240
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label UnregisteredLabel 
            Alignment       =   2  'Center
            BackColor       =   &H00FF80FF&
            Caption         =   "Unregistered Software"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2880
            TabIndex        =   25
            Top             =   5520
            Width           =   10455
         End
         Begin VB.Label lblTitle 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "Pipe Capacity Analysis Report"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5543
            TabIndex        =   24
            Top             =   0
            Width           =   5775
         End
         Begin VB.Label PrintedLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "Printed:"
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
            Left            =   13560
            TabIndex        =   23
            Top             =   960
            Width           =   855
         End
         Begin VB.Label PrintedData 
            Caption         =   "xxxxxx"
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
            Left            =   14520
            TabIndex        =   22
            Top             =   960
            Width           =   1695
         End
         Begin VB.Line FooterBreakLine 
            X1              =   120
            X2              =   15120
            Y1              =   10680
            Y2              =   10680
         End
         Begin VB.Line HeaderBreakLine 
            X1              =   0
            X2              =   16200
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Label CleanFlowSystemEmailLabel 
            Caption         =   "www.cleanflowsystems.com"
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
            Left            =   7200
            TabIndex        =   21
            Top             =   10800
            Width           =   2295
         End
         Begin VB.Label PhLabel 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ph:"
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
            TabIndex        =   20
            Top             =   840
            Width           =   855
         End
         Begin VB.Label PhData 
            Caption         =   "xxxxxx"
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
            Left            =   3600
            TabIndex        =   19
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label PageLabel 
            Caption         =   "Page: 1/1"
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
            Left            =   15360
            TabIndex        =   18
            Top             =   10800
            Width           =   855
         End
         Begin VB.Label GraphUnitLabel 
            Caption         =   "Meters"
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
            Left            =   15000
            TabIndex        =   17
            Top             =   6600
            Width           =   495
         End
         Begin VB.Label CopyrightLabel 
            Caption         =   "Copyright 2005"
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
            Left            =   13440
            TabIndex        =   16
            Top             =   10800
            Width           =   1455
         End
         Begin VB.Label CleanFlowSystemsWebAddressLabel 
            Caption         =   "www.cleanflowsystems.com"
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
            TabIndex        =   15
            Top             =   10800
            Width           =   2295
         End
         Begin VB.Line GraphYDevisionLine 
            BorderColor     =   &H00808080&
            Index           =   5
            X1              =   960
            X2              =   15840
            Y1              =   6330
            Y2              =   6330
         End
         Begin VB.Line GraphYDevisionLine 
            BorderColor     =   &H00808080&
            Index           =   0
            X1              =   840
            X2              =   15840
            Y1              =   5280
            Y2              =   5280
         End
         Begin VB.Image PVProfileImage 
            Height          =   2625
            Index           =   0
            Left            =   520
            Stretch         =   -1  'True
            Top             =   1440
            Width           =   3500
         End
         Begin VB.Image PVYScaleImage 
            Height          =   2625
            Index           =   0
            Left            =   120
            Top             =   1440
            Width           =   400
         End
         Begin VB.Label ObservationsLabel 
            Caption         =   "Observation Text here"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   4440
            Width           =   3855
         End
         Begin VB.Image PVXScaleImage 
            Height          =   300
            Index           =   0
            Left            =   520
            Top             =   4065
            Width           =   3495
         End
         Begin VB.Image PVProfileImage 
            Height          =   2625
            Index           =   1
            Left            =   4480
            Stretch         =   -1  'True
            Top             =   1440
            Width           =   3495
         End
         Begin VB.Image PVYScaleImage 
            Height          =   2625
            Index           =   1
            Left            =   4080
            Top             =   1440
            Width           =   400
         End
         Begin VB.Label ObservationsLabel 
            Caption         =   "Observation Text here"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   4080
            TabIndex        =   13
            Top             =   4440
            Width           =   3855
         End
         Begin VB.Image PVXScaleImage 
            Height          =   300
            Index           =   1
            Left            =   4480
            Top             =   4065
            Width           =   3495
         End
         Begin VB.Image PVProfileImage 
            Height          =   2625
            Index           =   2
            Left            =   8440
            Stretch         =   -1  'True
            Top             =   1440
            Width           =   3495
         End
         Begin VB.Image PVYScaleImage 
            Height          =   2625
            Index           =   2
            Left            =   8040
            Top             =   1440
            Width           =   400
         End
         Begin VB.Label ObservationsLabel 
            Caption         =   "Observation Text here"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   8040
            TabIndex        =   12
            Top             =   4440
            Width           =   3855
         End
         Begin VB.Image PVXScaleImage 
            Height          =   300
            Index           =   2
            Left            =   8440
            Top             =   4065
            Width           =   3495
         End
         Begin VB.Image PVProfileImage 
            Height          =   2625
            Index           =   3
            Left            =   12400
            Stretch         =   -1  'True
            Top             =   1440
            Width           =   3500
         End
         Begin VB.Image PVYScaleImage 
            Height          =   2625
            Index           =   3
            Left            =   12000
            Top             =   1440
            Width           =   400
         End
         Begin VB.Label ObservationsLabel 
            Caption         =   "Observation Text here"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   12000
            TabIndex        =   11
            Top             =   4440
            Width           =   3855
         End
         Begin VB.Image PVXScaleImage 
            Height          =   300
            Index           =   3
            Left            =   12400
            Top             =   4065
            Width           =   3495
         End
         Begin VB.Image PVProfileImage 
            Height          =   2625
            Index           =   4
            Left            =   520
            Stretch         =   -1  'True
            Top             =   6960
            Width           =   3495
         End
         Begin VB.Image PVYScaleImage 
            Height          =   2625
            Index           =   4
            Left            =   120
            Top             =   6960
            Width           =   400
         End
         Begin VB.Label ObservationsLabel 
            Caption         =   "Observation Text here"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   10
            Top             =   9960
            Width           =   3855
         End
         Begin VB.Image PVXScaleImage 
            Height          =   300
            Index           =   4
            Left            =   520
            Top             =   9585
            Width           =   3495
         End
         Begin VB.Image PVProfileImage 
            Height          =   2625
            Index           =   5
            Left            =   4480
            Stretch         =   -1  'True
            Top             =   6960
            Width           =   3495
         End
         Begin VB.Image PVYScaleImage 
            Height          =   2625
            Index           =   5
            Left            =   4080
            Top             =   6960
            Width           =   400
         End
         Begin VB.Label ObservationsLabel 
            Caption         =   "Observation Text here"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   4080
            TabIndex        =   9
            Top             =   9960
            Width           =   3855
         End
         Begin VB.Image PVXScaleImage 
            Height          =   300
            Index           =   5
            Left            =   4480
            Top             =   9585
            Width           =   3495
         End
         Begin VB.Image PVProfileImage 
            Height          =   2625
            Index           =   6
            Left            =   8440
            Stretch         =   -1  'True
            Top             =   6960
            Width           =   3495
         End
         Begin VB.Image PVYScaleImage 
            Height          =   2625
            Index           =   6
            Left            =   8040
            Top             =   6960
            Width           =   400
         End
         Begin VB.Label ObservationsLabel 
            Caption         =   "Observation Text here"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   6
            Left            =   8040
            TabIndex        =   8
            Top             =   9960
            Width           =   3855
         End
         Begin VB.Image PVXScaleImage 
            Height          =   300
            Index           =   6
            Left            =   8440
            Top             =   9585
            Width           =   3495
         End
         Begin VB.Image PVProfileImage 
            Height          =   2625
            Index           =   7
            Left            =   12400
            Stretch         =   -1  'True
            Top             =   6960
            Width           =   3495
         End
         Begin VB.Image PVYScaleImage 
            Height          =   2625
            Index           =   7
            Left            =   12000
            Top             =   6960
            Width           =   400
         End
         Begin VB.Label ObservationsLabel 
            Caption         =   "Observation Text here"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   7
            Left            =   12000
            TabIndex        =   7
            Top             =   9960
            Width           =   3855
         End
         Begin VB.Image PVXScaleImage 
            Height          =   300
            Index           =   7
            Left            =   12400
            Top             =   9585
            Width           =   3495
         End
         Begin VB.Shape GraphBackgroundShape 
            BackColor       =   &H00FFFFA2&
            FillColor       =   &H00FFFFA2&
            FillStyle       =   0  'Solid
            Height          =   1695
            Left            =   240
            Tag             =   "Back"
            Top             =   5160
            Width           =   15615
         End
         Begin VB.Shape PVProfileGraphBoarder 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   2925
            Index           =   4
            Left            =   120
            Tag             =   "Back"
            Top             =   6960
            Width           =   3900
         End
         Begin VB.Shape PVProfileGraphBoarder 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   2925
            Index           =   5
            Left            =   4080
            Tag             =   "Back"
            Top             =   6960
            Width           =   3900
         End
         Begin VB.Shape PVProfileGraphBoarder 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   2925
            Index           =   6
            Left            =   8040
            Tag             =   "Back"
            Top             =   6960
            Width           =   3900
         End
         Begin VB.Shape PVProfileGraphBoarder 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   2925
            Index           =   7
            Left            =   12000
            Tag             =   "Back"
            Top             =   6960
            Width           =   3900
         End
         Begin VB.Shape PVProfileGraphBoarder 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   2925
            Index           =   2
            Left            =   8040
            Tag             =   "Back"
            Top             =   1440
            Width           =   3900
         End
         Begin VB.Shape PVProfileGraphBoarder 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   2925
            Index           =   3
            Left            =   12000
            Tag             =   "Back"
            Top             =   1440
            Width           =   3900
         End
         Begin VB.Shape PVProfileGraphBoarder 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   2925
            Index           =   1
            Left            =   4080
            Tag             =   "Back"
            Top             =   1440
            Width           =   3900
         End
         Begin VB.Shape PVProfileGraphBoarder 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   2925
            Index           =   0
            Left            =   120
            Tag             =   "Back"
            Top             =   1440
            Width           =   3900
         End
      End
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   615
      Left            =   120
      Picture         =   "PVReportMultiProfile.frx":C69A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton ScaleButton05 
      Caption         =   "0.5"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton ScaleButton10 
      Caption         =   "1"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton ScaleButton15 
      Caption         =   "1.5"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton ScaleButton20 
      Caption         =   "2"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   495
   End
   Begin MSComDlg.CommonDialog FloatingTextDialog 
      Left            =   120
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu FloatingTextMenu 
      Caption         =   "FloatingText"
      Visible         =   0   'False
      Begin VB.Menu FloatingTextFontMenu 
         Caption         =   "Font"
      End
      Begin VB.Menu FloatingTextBackgroundColourMenu 
         Caption         =   "Background Colour"
      End
      Begin VB.Menu FloatingTextDefaultMenu 
         Caption         =   "Rest to default"
      End
      Begin VB.Menu Blank 
         Caption         =   ""
      End
      Begin VB.Menu FloatingTextDeleteMenu 
         Caption         =   "Delete"
      End
      Begin VB.Menu FloatingTextDeleteAllMenu 
         Caption         =   "Delete All"
      End
   End
End
Attribute VB_Name = "PVReportMultiProfilex8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ReportMouseX As Single
Dim ReportMouseY As Single
Dim ReportMouseDown As Boolean
Public PreviewStartFrame As Long
Public PreviewEndFrame As Long
Public RenderScale As Single
Public PrintPreviewAction As String

Private Sub AllFramesButtons_Click()
    GraphStartFrame = 1
    GraphEndFrame = PVDataNoOfLines
    Call PositionReportControls
    Call FillOutPrintForm
    Call GraphSpecificSettings
      
    
    Call RenderForm
End Sub

Private Sub CmdMove_Click()
    PrintPreviewAction = "MoveAll"
End Sub

Private Sub cmdPrint_Click()
    RenderScale = 1
    PVGraphOvalityXScale = 8
    PVGraphOvalityXOffset = -25

    ScreenDrawingType = 1
    ScreenDrawingOrientation = 1

    Call DrawPVGraphsReport
    Call RenderToPrinter.RenderReport(Me, Printer, 1)
    Call DrawPVGraphsReport

    Dim i
    For i = 0 To 7
       Set ScreenDrawing.ReportDummyGraphImage = PVProfileImage(i)
       'If VideoSnapShotMode = SnapShot And CLPScreenMode = Video Then 'PCNGL210103
    '   If CLPScreenMode = SnapShot Then 'PCNGL210103 'PCN4043
    '       DrawSF = ReportDummyGraphImage.width / ClearLineScreen.MainScreen.width  'PCN1835
    '       Call DrawProfilesStartToFinish(Printer) 'PCN3691
    '       Call DrawMainScale(Printer)
    '       PVProfileGraphBoarder(1).Tag = ""
    '   Else
           Call DrawProfilesStartToFinish(Printer, True) 'PCN3691
           Set ScreenDrawing.ReportGraphImageX = PVReportMultiProfilex8.PVXScaleImage(i)
           Set ScreenDrawing.ReportGraphImageY = PVReportMultiProfilex8.PVYScaleImage(i)
           Call DrawMainScale(Printer)
    '   End If
    Next i
    Call Printer.EndDoc
    ScreenDrawingType = 0
    ScreenDrawingOrientation = 0
End Sub

Private Sub Command1_Click()

End Sub

Private Sub CmdText_Click()
    PrintPreviewAction = "DrawText"
End Sub

Private Sub FloatingText_Change(Index As Integer)
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingText_Change(Me, Index)
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-" & error$
    End Select
End Sub

Private Sub FloatingText_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingText_KeyPress(Me, Index, KeyAscii)
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-" & error$
    End Select
End Sub

Private Sub FloatingText_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingText_MouseDown(Me, Index, Button, Shift, X, Y)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-" & error$
    End Select
End Sub

Private Sub FloatingText_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingText_MouseMove(Me, Index, Button, Shift, X, Y)
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-" & error$
    End Select
End Sub

Private Sub FloatingTextBackgroundColourMenu_Click()
On Error GoTo Err_Handler
    
    Call RenderToPrinter.FloatingTextBackgroundColourMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-" & error$
    End Select
End Sub

Private Sub FloatingTextDefaultMenu_Click()
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingTextDefaultMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-" & error$
    End Select
End Sub

Private Sub FloatingTextDeleteAllMenu_Click()
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingTextDeleteAllMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-" & error$
    End Select
End Sub

Private Sub FloatingTextDeleteMenu_Click()
On Error GoTo Err_Handler
    
    Call RenderToPrinter.FloatingTextDeleteMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-" & error$
    End Select
End Sub

Private Sub FloatingTextFontMenu_Click()
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingTextFontMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-" & error$
    End Select
End Sub

Private Sub Form_Load()
    Me.Show
    
    PageFramePictureBox.width = Me.width - PageFramePictureBox.Left
    PageFramePictureBox.height = Me.height
    PageFramePictureBox.Top = 720
    
    RenderScale = 1
    Set PrintPreviewForm = Me
    
    PreviewStartFrame = GraphStartFrame
    PreviewEndFrame = GraphEndFrame
    
    Call PositionReportControls
    Call FillOutPrintForm
    Call MarkForPrinting
    
    Call GraphSpecificSettings
      
    
    Call RenderForm

End Sub

Private Sub RenderForm()
    picReportPage.Cls
    picReportPage.width = Printer.width * RenderScale
    picReportPage.height = Printer.height * RenderScale
    
    PreviewStartFrame = GraphStartFrame
    PreviewEndFrame = GraphEndFrame

    PVGraphOvalityXScale = 8
    PVGraphOvalityXOffset = -25
    

    ScreenDrawingType = 2
    ScreenDrawingOrientation = 1
    
    Call DrawPVGraphsReport
    Call RenderToPrinter.RenderReport(Me, picReportPage, RenderScale)
    Call DrawPVGraphsReport

    Dim i
    For i = 0 To 7
       Set ScreenDrawing.ReportDummyGraphImage = PVProfileImage(i)
       'If VideoSnapShotMode = SnapShot And CLPScreenMode = Video Then 'PCNGL210103
    '   If CLPScreenMode = SnapShot Then 'PCNGL210103
    '       DrawSF = ReportDummyGraphImage.width / ClearLineScreen.MainScreen.width  'PCN1835
    '       Call DrawProfilesStartToFinish(picReportPage) 'PCN3691
     '      Call DrawMainScale(picReportPage)
    '       PVProfileGraphBoarder(0).Tag = ""
    '   Else
           Call DrawProfilesStartToFinish(picReportPage, True) 'PCN3691
           Set ScreenDrawing.ReportGraphImageX = PVReportMultiProfilex8.PVXScaleImage(i)
           Set ScreenDrawing.ReportGraphImageY = PVReportMultiProfilex8.PVYScaleImage(i)
           
           Call DrawMainScale(picReportPage)
 '   End If
    Next i
    
    ScreenDrawingType = 0
    ScreenDrawingOrientation = 0
End Sub



Private Function PositionReportControls()
'****************************************************************************************
'Name    : PositionReportControls
'Created : August 9 05, PCN3593
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    : positions all the report images for display AND for printing!
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
Dim i As Integer

On Error GoTo ManualOrientation
Printer.Orientation = vbPRORLandscape
Printer.PrintQuality = vbPRPQHigh

'Detect the CURRENT page setup of the deault printer
picReportPage.width = Printer.width
picReportPage.height = Printer.height
picReportPage.Left = 1000
picReportPage.Top = 500



ManualOrientationSet:


lblTitle.Left = (picReportPage.width / 2) - (lblTitle.width / 2)

'setup the scroll bar
'scrPageScroll.Left = Screen.width - scrPageScroll.width
'scrPageScroll.Top = RehabImagesReport.Top
'scrPageScroll.height = Screen.height - 500


Exit Function
ManualOrientation:
On Error GoTo Err_Handler

Dim originalheight
Dim originalwidth

originalheight = Printer.height
originalwidth = Printer.width

Printer.height = originalwidth
Printer.width = originalheight

picReportPage.width = Printer.width
picReportPage.height = Printer.height
picReportPage.Left = 1000
picReportPage.Top = 500
GoTo ManualOrientationSet:

Err_Handler:
MsgBox Err & " - " & error$

End Function


Private Sub Form_Resize()
    PageFramePictureBox.width = Me.width - PageFramePictureBox.Left
    PageFramePictureBox.height = Me.height
    PageFramePictureBox.Top = 720
End Sub

Private Sub picReportPage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PrintPreviewAction = "DrawText" Then
        Call RenderToPrinter.FloatingTextAdd(Me, Button, Shift, X, Y)
    Else
        ReportMouseDown = True
    End If
    ReportMouseX = X
    ReportMouseY = Y
End Sub

Private Sub picReportPage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ReportMouseDown Then
        picReportPage.Left = picReportPage.Left + X - ReportMouseX
        picReportPage.Top = picReportPage.Top + Y - ReportMouseY
    End If
End Sub

Private Sub picReportPage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReportMouseDown = False
End Sub

Private Sub FillOutPrintForm()
    Dim i As Long
    Dim ControlType As String
    Dim DisplayUnits As String
    Dim ProfileSlicePosition As Single
    
    Dim LeftLimit As Double
    Dim RightLimit As Double
    
    
    
    PVGraphOvalityXScale = 8
    PVGraphOvalityXOffset = -25
    Call PrecisionVisionGraph.GetGeneralPVGraphData(ScreenDrawing.GraphContainer(0).GraphType)
    LeftLimit = Format(ConvertUnitByGraph(PVXScaleLimitPerL, 0, DisplayUnits), "###0.0")
    RightLimit = Format(ConvertUnitByGraph(PVXScaleLimitPerR, 0, DisplayUnits), "###0.0")
    
    With PVReportMultiProfilex8
  
        .PhData.Caption = PhoneNo
        .LogoImage.Picture = LoadPicture(CompanyLogoPath)
        
        
        .lblTitle.Caption = PrecisionVisionGraph.Label_GraphName(0) & " Profile Report"
        

        .ObservationsLabel(0).Caption = ""
        

    
        'By default make all lable no background
        For i = 0 To .Controls.Count - 1
            ControlType = TypeName(.Controls(i))
            Select Case ControlType
                Case "Label": .Controls(i).BackStyle = 0
            End Select
        Next i

''    If PVDataNoOfLines > 1 Then
''        If MeasurementUnits = "mm" Then
''            DisplayUnits = "mm"
''        Else
''            DisplayUnits = "in"
''        End If
''
''
''
''
''        '^^^^ ***********************************************
''        'Distance
''        'If ConfigInfo.DistanceStart >= 0 Then
''            If MeasurementUnits = "mm" Then
''                .PVKey_Distance_Value = Format(PVDistances(PVFrameNo), "#0.0") & "m"
''            Else
''                .PVKey_Distance_Value = Format(PVDistances(PVFrameNo), "#0") & "ft"
''            End If
''
''        .UnitLabel = PrecisionVisionGraph.PVXScaleLabelUnits(0)
''        'End If
''    End If
        
    'If VideoSnapShotMode = SnapShot And CLPScreenMode = Video Then 'PCNGL210103
    For i = 0 To 7
        If CLPScreenMode = SnapShot Then 'PCNGL210103 'PCN4043
            .PVProfileImage(i).Picture = LoadPicture(LocToSave & "Snapshot.bmp")
        End If
    Next i
        
    End With
    
    ProfileSlicePosition = GraphContainer.width / (CSng(PreviewEndFrame - PreviewStartFrame))
    ProfileSlicePosition = (PVFrameNo - PreviewStartFrame) * ProfileSlicePosition
    If ProfileSlicePosition < 0 Then ProfileSlicePosition = -600
    If ProfileSlicePosition > GraphContainer.width Then ProfileSlicePosition = GraphContainer.width + 300
    
    ProfileSlicePosition = ProfileSlicePosition + GraphContainer.Left
    
    'ProfileSliceRubberBand.X1 = ProfileSlicePosition
    'ProfileSlice.X1 = ProfileSlicePosition
    'ProfileSlice.X2 = ProfileSlicePosition
    
    
    
    Call ScreenDrawing.FormTopMost(PVGraphsKeyForm.hwnd) 'PCN2990

End Sub

Sub GraphSpecificSettings()


End Sub

Sub MarkForPrinting()
Dim i As Integer
Dim ControlType As String

'Draw renderings first that are marked back
For i = 0 To PVReportMultiProfilex8.Controls.Count - 1
    
    With PVReportMultiProfilex8.Controls(i)
        If TypeName(.Container) = "PictureBox" Then
            If .Tag = "Paper" Then
                .Visible = True
            ElseIf .Tag = "Back" Then
                .Visible = False
            ElseIf .Visible Then
                .Tag = "Visible"
                .Visible = False
            Else
                .Visible = False
                .Tag = ""
            End If
        End If
    End With
    
Next i
End Sub

Private Sub ScaleButton05_Click()
    RenderScale = 0.5
    Call RenderForm
End Sub

Private Sub ScaleButton10_Click()
    RenderScale = 1
    Call RenderForm
End Sub

Private Sub ScaleButton15_Click()
    RenderScale = 1.5
    Call RenderForm
End Sub

Private Sub ScaleButton20_Click()
    RenderScale = 2
    Call RenderForm
End Sub


