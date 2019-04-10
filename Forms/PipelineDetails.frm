VERSION 5.00
Object = "{7ACB8247-EB73-11D3-881B-B628C6CB1474}#1.0#0"; "smakejpg.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form PipelineDetails 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   10155
   ClientLeft      =   19995
   ClientTop       =   0
   ClientWidth     =   8940
   ControlBox      =   0   'False
   Icon            =   "PipelineDetails.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10155
   ScaleMode       =   0  'User
   ScaleWidth      =   6123.286
   ShowInTaskbar   =   0   'False
   Begin ClearLineProfiler.CBS_DropDownBox RefShape_DropDownBox 
      Height          =   375
      Left            =   1200
      TabIndex        =   38
      Top             =   6480
      Width           =   2880
      _extentx        =   6773
      _extenty        =   661
   End
   Begin VB.Frame FinishNodeFrame 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Finish Node"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1215
      Left            =   120
      TabIndex        =   19
      Top             =   3640
      Width           =   4065
      Begin VB.TextBox SeaLevelEndHeightTextBox 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2760
         TabIndex        =   44
         ToolTipText     =   "end mean sea level"
         Top             =   840
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox FinishNodeNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   8
         Top             =   180
         Width           =   1905
      End
      Begin VB.TextBox FinishNodeLocation 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   9
         Top             =   510
         Width           =   1905
      End
      Begin VB.Label seaLevelUnitLabel 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "m"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Index           =   0
         Left            =   3840
         TabIndex        =   46
         Top             =   960
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image EndSeaHeightImage 
         Height          =   225
         Left            =   2520
         Picture         =   "PipelineDetails.frx":038A
         Stretch         =   -1  'True
         Top             =   885
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label FinishNodeFrameLbl 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   120
         TabIndex        =   42
         Top             =   0
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label FinishNodeNo_lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Finish No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   120
         MousePointer    =   3  'I-Beam
         TabIndex        =   21
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label FhNodeLoc_lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   120
         MousePointer    =   3  'I-Beam
         TabIndex        =   20
         Top             =   480
         Width           =   1905
      End
   End
   Begin VB.Frame StartNodeFrame 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Start Node"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1575
      Left            =   120
      TabIndex        =   16
      Top             =   2000
      Width           =   4065
      Begin VB.TextBox OffsetTextBox 
         Height          =   285
         Left            =   240
         TabIndex        =   50
         Text            =   "100"
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox DesignGradientTextBox 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2760
         TabIndex        =   49
         ToolTipText     =   "design gradient"
         Top             =   860
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox SeaLevelStartHeightTextBox 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2760
         TabIndex        =   45
         ToolTipText     =   "start mean sea level"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox StartNodeNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   6
         Top             =   180
         Width           =   1905
      End
      Begin VB.TextBox StartNodeLocation 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   7
         Top             =   520
         Width           =   1905
      End
      Begin VB.Label seaLevelUnitLabel 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "o"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Index           =   2
         Left            =   3840
         TabIndex        =   48
         Top             =   840
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   2520
         Picture         =   "PipelineDetails.frx":0AC8
         Stretch         =   -1  'True
         Top             =   910
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label seaLevelUnitLabel 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "m"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Index           =   1
         Left            =   3840
         TabIndex        =   47
         Top             =   1320
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image StartSeaHeightImage 
         Height          =   240
         Left            =   2520
         Picture         =   "PipelineDetails.frx":1442
         Stretch         =   -1  'True
         Top             =   1245
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label StartNodeFrameLbl 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   120
         TabIndex        =   41
         Top             =   0
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label StartNodeNo_lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Start No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   120
         MousePointer    =   3  'I-Beam
         TabIndex        =   18
         Top             =   240
         Width           =   1785
      End
      Begin VB.Label StNodeLoc_lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   120
         MousePointer    =   3  'I-Beam
         TabIndex        =   17
         Top             =   480
         Width           =   1905
      End
   End
   Begin VB.Frame AssetInfo 
      BackColor       =   &H00C0C0C0&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4065
      Begin VB.TextBox sTime 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         TabIndex        =   5
         Top             =   1440
         Width           =   1545
      End
      Begin VB.TextBox sDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         TabIndex        =   4
         Top             =   1110
         Width           =   1545
      End
      Begin VB.TextBox City 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   780
         Width           =   2145
      End
      Begin VB.TextBox SiteID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         TabIndex        =   2
         Top             =   450
         Width           =   1545
      End
      Begin VB.TextBox AssetNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         TabIndex        =   1
         Top             =   120
         Width           =   1545
      End
      Begin VB.Label Time_lbl 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   120
         MousePointer    =   3  'I-Beam
         TabIndex        =   32
         Top             =   1440
         Width           =   2205
      End
      Begin VB.Label Date_lbl 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         MousePointer    =   3  'I-Beam
         TabIndex        =   31
         Top             =   1164
         Width           =   2220
      End
      Begin VB.Label City_lbl 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   120
         MousePointer    =   3  'I-Beam
         TabIndex        =   30
         Top             =   856
         Width           =   1605
      End
      Begin VB.Label SiteID_lbl 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Site ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   120
         MousePointer    =   3  'I-Beam
         TabIndex        =   15
         Top             =   548
         Width           =   2235
      End
      Begin VB.Label AssetNo_lbl 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Asset No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   465
         Left            =   120
         MousePointer    =   3  'I-Beam
         TabIndex        =   14
         Top             =   120
         Width           =   2325
         WordWrap        =   -1  'True
      End
   End
   Begin sMakeJPG.sMJPG JPGMake1 
      Left            =   7440
      Top             =   1200
      _ExtentX        =   767
      _ExtentY        =   556
   End
   Begin VB.Frame PipeDataFrame 
      BackColor       =   &H00C0C0C0&
      Height          =   2055
      Left            =   120
      TabIndex        =   22
      Top             =   4920
      Width           =   4065
      Begin VB.TextBox PipeLength 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         TabIndex        =   12
         Top             =   870
         Width           =   840
      End
      Begin VB.TextBox OutsideDiameter 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         TabIndex        =   11
         Top             =   525
         Width           =   840
      End
      Begin VB.TextBox InternalDiameterExpected 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         TabIndex        =   10
         Top             =   195
         Width           =   840
      End
      Begin VB.TextBox Material 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   13
         Top             =   1200
         Width           =   1905
      End
      Begin VB.Label Shape_lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Reference Shape"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Left            =   120
         MousePointer    =   3  'I-Beam
         TabIndex        =   39
         Top             =   1560
         Width           =   945
      End
      Begin VB.Label PipeLen_lbl 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Pipeline Length"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   120
         MousePointer    =   3  'I-Beam
         TabIndex        =   29
         Top             =   900
         Width           =   2520
      End
      Begin VB.Label lblPipeLengthUnit 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "m"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   3600
         TabIndex        =   28
         Top             =   960
         Width           =   360
      End
      Begin VB.Label OutsideDiameter_lbl 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Outside Diameter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   120
         MousePointer    =   3  'I-Beam
         TabIndex        =   27
         Top             =   615
         Width           =   2520
      End
      Begin VB.Label unit2 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   3555
         TabIndex        =   26
         Top             =   600
         Width           =   360
      End
      Begin VB.Label InternalDiameterExpected_lbl 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Internal Diameter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Left            =   120
         MousePointer    =   3  'I-Beam
         TabIndex        =   25
         Top             =   195
         Width           =   2520
      End
      Begin VB.Label unit1 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   3555
         TabIndex        =   24
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Material_lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Material"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   120
         MousePointer    =   3  'I-Beam
         TabIndex        =   23
         Top             =   1200
         Width           =   1905
      End
   End
   Begin VB.Frame CommentsFrame 
      BackColor       =   &H00C0C0C0&
      Caption         =   "General Comments"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1455
      Left            =   120
      TabIndex        =   36
      Top             =   7200
      Width           =   4065
      Begin RichTextLib.RichTextBox RichGeneralComments 
         Height          =   1575
         Left            =   360
         TabIndex        =   40
         Top             =   720
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   2778
         _Version        =   393217
         BackColor       =   -2147483624
         Enabled         =   -1  'True
         Appearance      =   0
         TextRTF         =   $"PipelineDetails.frx":1B80
      End
      Begin VB.TextBox GeneralComments 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   1155
         Left            =   30
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   240
         Width           =   4000
      End
      Begin VB.Label CommentsFrameLbl 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   120
         TabIndex        =   43
         Top             =   0
         Visible         =   0   'False
         Width           =   45
      End
   End
   Begin VB.Line Line1 
      X1              =   2712.328
      X2              =   2630.136
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Image FormBackground 
      Height          =   8895
      Left            =   0
      Picture         =   "PipelineDetails.frx":1C02
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5025
   End
   Begin VB.Label unit3 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "mm"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   7440
      TabIndex        =   35
      Top             =   1920
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "m"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   1080
      TabIndex        =   34
      Top             =   570
      Width           =   225
   End
   Begin VB.Label Unit6 
      BackColor       =   &H8000000A&
      Caption         =   "m"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   7440
      TabIndex        =   33
      Top             =   1560
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "PipelineDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DataEntryChange As Boolean 'PCNGL230103
 Public CommonDialog1 As Object
'PCN3650
'Public Len_Real As Double
'Public LenRealPercent As Double

'vvvv PCN4344 *************************************
Dim TheRefShapeFiles() As String  'PCN4171
Dim TheRefShapeFileLabels() As String  'PCN4171
'^^^^ *********************************************

Public Area As Double

Dim UnicodeGeneralComments As String



'Private Sub AssetNo_Click()
'
'If LanguageCharset <> 0 Then
'    Call DataEntryForm.SetDataEntryType("AssetNoTextChange", Me.AssetNo.text)
'End If
'
'End Sub

Private Sub AssetNo_KeyDown(KeyCode As Integer, Shift As Integer)

    If LanguageCharset <> 0 Then
        AssetNo.Font.Charset = LanguageCharset
        'Call DataEntryForm.SetDataEntryType("AssetNoTextChange", Me.AssetNo.text)
    End If
    
End Sub

Private Sub AssetNo_KeyUp(KeyCode As Integer, Shift As Integer) 'PCNGL230103
On Error GoTo Err_Handler
Dim SaveFailed As Boolean

DataEntryChange = True

Exit Sub
Err_Handler:
    MsgBox Err & "-PD1:" & Error$
End Sub



Private Sub AssetNo_lbl_Click() 'PCN4171
On Error GoTo Err_Handler

Call DataEntryForm.SetDataEntryType("AssetNoLabelChange", Me.AssetNo_lbl.Caption)

Exit Sub
Err_Handler:
    MsgBox Err & "-PD2:" & Error$
End Sub

Private Sub AssetNo_Validate(Cancel As Boolean) 'PCNGL230103
On Error GoTo Err_Handler
Dim SaveFailed As Boolean

'PCN2189
'If DataEntryChange = True And PVDFileName <> "" Then
'    PipelineInfo.AssetNo = Me.AssetNo
'    Observations.AssetNo = Me.AssetNo
'    Call SaveToFilePipeAndConfigInfo("PipelineInfo", SaveFailed)
'End If

If AssetNo = "" Then 'PCN1832
    PipelineInfo.AssetNo = Empty 'PCN1832
    'Observations.AssetNo = "" 'PCN1832 'PCN4131
Else
    PipelineInfo.AssetNo = AssetNo
    'Observations.AssetNo = AssetNo 'PCN4131
    Call SaveToFilePipeAndConfigInfo("PipelineInfo", SaveFailed) 'PCN1768
End If
Change = False

DataEntryChange = False

Exit Sub
Err_Handler:
    MsgBox Err & "-PD3:" & Error$
End Sub


'Private Sub City_Click()
'If LanguageCharset <> 0 Then
'    Call DataEntryForm.SetDataEntryType("CityTextChange", Me.City.text)
'End If
'End Sub

Private Sub City_KeyDown(KeyCode As Integer, Shift As Integer)
If LanguageCharset <> 0 Then
'    City.Font.Charset = LanguageCharset
'    'Call DataEntryForm.SetDataEntryType("CityTextChange", Me.City.text)
End If
End Sub

Private Sub City_KeyUp(KeyCode As Integer, Shift As Integer) 'PCN1768
On Error GoTo Err_Handler
Dim SaveFailed As Boolean

DataEntryChange = True

Exit Sub
Err_Handler:
    MsgBox Err & "-PD4:" & Error$
End Sub

Private Sub City_lbl_Click() 'PCN4171
On Error GoTo Err_Handler

Call DataEntryForm.SetDataEntryType("CityLabelChange", Me.City_lbl.Caption)

Exit Sub
Err_Handler:
    MsgBox Err & "-PD5:" & Error$
End Sub

Private Sub City_Validate(Cancel As Boolean) 'PCN1768
On Error GoTo Err_Handler
Dim SaveFailed As Boolean

If DataEntryChange = True And PVDFileName <> "" Then
    PipelineInfo.City = Me.City
    Call SaveToFilePipeAndConfigInfo("PipelineInfo", SaveFailed)
End If

DataEntryChange = False

Exit Sub
Err_Handler:
    MsgBox Err & "-PD6:" & Error$
End Sub



Private Sub Date_lbl_Click() 'PCN4171
On Error GoTo Err_Handler

Call DataEntryForm.SetDataEntryType("DateLabelChange", Me.Date_lbl.Caption)

Exit Sub
Err_Handler:
    MsgBox Err & "-PD7:" & Error$
End Sub


'PCN6458
'Private Sub DesignGradientTextBox_KeyUp(KeyCode As Integer, Shift As Integer)
'On Error GoTo Err_Handler
'
'    Dim FailSave As Boolean
'    Dim angle As Double
'    Dim drop As Double
'    Dim Distance As Double
'
'    If IsNumeric(DesignGradientTextBox.text) Then
'        'DesignGradient = CDbl(DesignGradientTextBox.text)
'        Call ScreenDrawing.UpDateInclinationGraph
'        'FontInfo.FontSize = DesignGradient * 1000
'        'FailSave = SaveToFileFontInfo(FailSave)
'    End If
'
'    'PCN6311
'    If Me.DesignGradientTextBox.text = "" And KeyCode = vbKeyReturn Then
'        Call ScreenDrawing.UpDateInclinationGraph
'    End If
'
'
'
'Exit Sub
'Err_Handler:
'    MsgBox Err & "-PD35.5"
'End Sub

Private Sub FhNodeLoc_lbl_Click() 'PCN4171
On Error GoTo Err_Handler

Call DataEntryForm.SetDataEntryType("FhNodeLocLabelChange", Me.FhNodeLoc_lbl.Caption)

Exit Sub
Err_Handler:
    MsgBox Err & "-PD8:" & Error$
End Sub

'Private Sub FinishNodeLocation_Click()
'If LanguageCharset <> 0 Then
'    Call DataEntryForm.SetDataEntryType("FhNodeLocTextChange", Me.FinishNodeLocation.text)
'End If
'End Sub

Private Sub FinishNodeLocation_KeyDown(KeyCode As Integer, Shift As Integer)
If LanguageCharset <> 0 Then
'    Call DataEntryForm.SetDataEntryType("FhNodeLocTextChange", Me.FinishNodeLocation.text)
End If
End Sub

Private Sub FinishNodeLocation_KeyUp(KeyCode As Integer, Shift As Integer) 'PCN1768
On Error GoTo Err_Handler
Dim SaveFailed As Boolean

DataEntryChange = True

Exit Sub
Err_Handler:
    MsgBox Err & "-PD9:" & Error$
End Sub

Private Sub FinishNodeLocation_Validate(Cancel As Boolean) 'PCN1768
On Error GoTo Err_Handler
Dim SaveFailed As Boolean

If DataEntryChange = True And PVDFileName <> "" Then
    PipelineInfo.FinishLocation = Me.FinishNodeLocation
    'Observations.FinishNodeNo = Me.FinishNodeNo  'PCN4131
    Call SaveToFilePipeAndConfigInfo("PipelineInfo", SaveFailed)
End If

DataEntryChange = False

Exit Sub
Err_Handler:
    MsgBox Err & "-PD10:" & Error$
End Sub

'Private Sub FinishNodeNo_Click()
'If LanguageCharset <> 0 Then
'    Call DataEntryForm.SetDataEntryType("FhNodeTextChange", Me.FinishNodeNo.text)
'End If
'End Sub

Private Sub FinishNodeNo_KeyDown(KeyCode As Integer, Shift As Integer)
If LanguageCharset <> 0 Then
'    Call DataEntryForm.SetDataEntryType("FhNodeTextChange", Me.FinishNodeNo.text)
End If
End Sub

Private Sub FinishNodeNo_KeyUp(KeyCode As Integer, Shift As Integer) 'PCN1768
On Error GoTo Err_Handler
Dim SaveFailed As Boolean

DataEntryChange = True

Exit Sub
Err_Handler:
    MsgBox Err & "-PD11:" & Error$
End Sub

Private Sub FinishNodeNo_lbl_Click() 'PCN4171
On Error GoTo Err_Handler

Call DataEntryForm.SetDataEntryType("FhNodeLabelChange", Me.FinishNodeNo_lbl.Caption)

Exit Sub
Err_Handler:
    MsgBox Err & "-PD12:" & Error$
End Sub

Private Sub FinishNodeNo_Validate(Cancel As Boolean) 'PCN1768 'PCN2189
On Error GoTo Err_Handler
Dim SaveFailed As Boolean
 
'PCN2189
'If DataEntryChange = True And PVDFileName <> "" Then
'    PipelineInfo.FinishName = Me.FinishNodeNo
'    Call SaveToFilePipeAndConfigInfo("PipelineInfo", SaveFailed)
'End If

If FinishNodeNo = "" Then 'PCN1832
    PipelineInfo.FinishName = Empty 'PCN1832
    'Observations.FinishNodeNo = "" 'PCN1832 'PCN4131
Else
    PipelineInfo.FinishName = FinishNodeNo
    'Observations.FinishNodeNo = FinishNodeNo 'PCN4131
    Call SaveToFilePipeAndConfigInfo("PipelineInfo", SaveFailed) 'PCN1768
End If
Change = False


DataEntryChange = False

Exit Sub
Err_Handler:
    MsgBox Err & "-PD13:" & Error$
End Sub


Private Sub Form_Activate()
On Error GoTo Err_Handler

PipelineDetails.GeneralComments.text = Trim(PipelineInfo.Comments)  'PCN4171


Exit Sub
Err_Handler:
    MsgBox Err & "-PD14:" & Error$
End Sub

Private Sub Form_Load()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Form_Load Sub  Michelle Lindsay michellelindsay@cbsys.co.nz
'
'Revision history
'   V0.0    Michelle Lindsay,   19/12/02    Adding functionality
'
'Description:
'
'Purpose:
'Show the correct labels for unit measurement as entered in the options page or
'contained in the .ini file.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler


ConvertLanguage Me, Language 'PCN2111





Me.Left = PVPageLeft 'PCNGL060103
Me.Top = PVPageTop
Me.width = PVPageWidth
Me.height = PVPageHeight
'vvvv PCN4171 **************************************
FormBackground.height = PVPageHeight
AssetInfo.BackColor = RGB(172, 196, 231)
StartNodeFrame.BackColor = RGB(172, 196, 231) 'PCN4171
FinishNodeFrame.BackColor = RGB(172, 196, 231) 'PCN4171
PipeDataFrame.BackColor = RGB(172, 196, 231) 'PCN4171
CommentsFrame.BackColor = RGB(172, 196, 231) 'PCN4171
'^^^^ **********************************************

Dim BackupMyFile As String 'PCN4431



'If (IdentifyOperatingSystem = "Windows XP") Then 'ML120203
'    StartNodeFrame.Top = 2650
'    FinishNodeFrame.Top = 4100
'    PipeDataFrame.Top = 5550
'    PipeDetailsToolbar.Top = 8450
''    AssetInfo.Left = 0
''    StartNodeFrame.Left = 0
''    FinishNodeFrame.Left = 0
''    PipeDataFrame.Left = 0
''    PipeDetailsToolbar.left = 0 'PCN2402
''    PipelineDetails.width = PVPageWidth - 150 'PCN2402
'End If

If MeasurementUnits = "mm" Then
    unit1.Caption = "mm"
    unit2.Caption = "mm"
    lblPipeLengthUnit.Caption = "m"
    unit3.Caption = "mm"
    Unit6.Caption = "m"
Else
    unit1.Caption = "in"
    unit2.Caption = "in"
    lblPipeLengthUnit.Caption = "ft"
    unit3.Caption = "in"
    Unit6.Caption = "ft"
    
End If

Me.seaLevelUnitLabel(0).Caption = Me.lblPipeLengthUnit.Caption 'PCN6128
Me.seaLevelUnitLabel(1).Caption = Me.lblPipeLengthUnit.Caption 'PCN6128

'InternalDiameterExpected = 0 'PCN1832
'OutsideDiameter = 0 'PCN1832
'PipeLength = 0 'PCN1832

'Load the form data from the PipelineInfo array 'PCNGL130103
Call CopyPipeDetailsToPipelineForm

'vvvv PCN3809 ********************************
If SoftwareConfiguration = "Reader" Then
    'Me.Enabled = False
    Me.CommentsFrame.Enabled = False
    Me.PipeDataFrame.Enabled = False
    Me.FinishNodeFrame.Enabled = False
    Me.StartNodeFrame.Enabled = False
    Me.AssetInfo.Enabled = False
    Me.RefShape_DropDownBox.Enabled = False
    
End If

'RefShape_DropDownBox.Font.Charset = LanguageCharset


Call PopulateRefShapeDropDown(RefShape_DropDownBox) 'PCN4171 'PCN4344
RefShape_DropDownBox.SetCharset (LanguageCharset)

BackupMyFile = MyFile 'PCN4431

'PCN4431 ''''''''''''''''''''''''''''''''''''
If SoftwareConfiguration = "Reader" Then    '
    BackupMyFile = MyFile                   '
    MyFile = App.Path & "\deploy.ini"       '
End If                                      '
'''''''''''''''''''''''''''''''''''''''''''''

Call LoadPipeDetailsFieldLabels
MyFile = BackupMyFile 'PCN4431

If LanguageCharset <> 0 Then
    Me.AssetNo.Font.Charset = LanguageCharset
    Me.SiteID.Font.Charset = LanguageCharset
    Me.City.Font.Charset = LanguageCharset
    Me.StartNodeNo.Font.Charset = LanguageCharset
    Me.StartNodeLocation.Font.Charset = LanguageCharset
    Me.FinishNodeNo.Font.Charset = LanguageCharset
    Me.FinishNodeLocation.Font.Charset = LanguageCharset
    Me.Material.Font.Charset = LanguageCharset
    Me.GeneralComments.Font.Charset = LanguageCharset
    Me.StartNodeFrameLbl.Font.Charset = LanguageCharset: Me.StartNodeFrameLbl.BackColor = RGB(172, 196, 231): Me.StartNodeFrameLbl.Caption = DisplayMessage(StartNodeFrame.Caption): Me.StartNodeFrameLbl.Visible = True
    Me.FinishNodeFrameLbl.Font.Charset = LanguageCharset: Me.FinishNodeFrameLbl.BackColor = RGB(172, 196, 231): Me.FinishNodeFrameLbl.Caption = DisplayMessage(FinishNodeFrame.Caption): Me.FinishNodeFrameLbl.Visible = True
    Me.CommentsFrameLbl.Font.Charset = LanguageCharset: Me.CommentsFrameLbl.BackColor = RGB(172, 196, 231): Me.CommentsFrameLbl.Caption = DisplayMessage(CommentsFrame.Caption): Me.CommentsFrameLbl.Visible = True
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-PD15:" & Error$
End Sub

Sub LoadPipeDetailsFieldLabels()
On Error GoTo Err_Handler
Dim INIStrValue As String

Call GetINI_ParameterInfoOnly(MyFile, "PipeDetailsAssetNo=", INIStrValue)
If INIStrValue <> "" Then Me.AssetNo_lbl = Trim(INIStrValue)

Call GetINI_ParameterInfoOnly(MyFile, "PipeDetailsSiteID=", INIStrValue)
If INIStrValue <> "" Then Me.SiteID_lbl = Trim(INIStrValue)

Call GetINI_ParameterInfoOnly(MyFile, "PipeDetailsCity=", INIStrValue)
If INIStrValue <> "" Then Me.City_lbl = Trim(INIStrValue)

Call GetINI_ParameterInfoOnly(MyFile, "PipeDetailsDate=", INIStrValue)
If INIStrValue <> "" Then Me.Date_lbl = Trim(INIStrValue)

Call GetINI_ParameterInfoOnly(MyFile, "PipeDetailsTime=", INIStrValue)
If INIStrValue <> "" Then Me.Time_lbl = Trim(INIStrValue)

Call GetINI_ParameterInfoOnly(MyFile, "PipeDetailsStNode=", INIStrValue)
If INIStrValue <> "" Then Me.StartNodeNo_lbl = Trim(INIStrValue)

Call GetINI_ParameterInfoOnly(MyFile, "PipeDetailsStLoc=", INIStrValue)
If INIStrValue <> "" Then Me.StNodeLoc_lbl = Trim(INIStrValue)

Call GetINI_ParameterInfoOnly(MyFile, "PipeDetailsFhNode=", INIStrValue)
If INIStrValue <> "" Then Me.FinishNodeNo_lbl = Trim(INIStrValue)

Call GetINI_ParameterInfoOnly(MyFile, "PipeDetailsFhLoc=", INIStrValue)
If INIStrValue <> "" Then Me.FhNodeLoc_lbl = Trim(INIStrValue)

Call GetINI_ParameterInfoOnly(MyFile, "PipeDetailsIntDiaExp=", INIStrValue)
If INIStrValue <> "" Then Me.InternalDiameterExpected_lbl = INIStrValue

Call GetINI_ParameterInfoOnly(MyFile, "PipeDetailsOutDiaExp=", INIStrValue)
If INIStrValue <> "" Then Me.OutsideDiameter_lbl = Trim(INIStrValue)

Call GetINI_ParameterInfoOnly(MyFile, "PipeDetailsLength=", INIStrValue)
If INIStrValue <> "" Then Me.PipeLen_lbl = Trim(INIStrValue)

Call GetINI_ParameterInfoOnly(MyFile, "PipeDetailsMaterial=", INIStrValue)
If INIStrValue <> "" Then Me.Material_lbl = Trim(INIStrValue)


Exit Sub
Err_Handler:
    MsgBox Err & "-PD16:" & Error$
End Sub

Private Sub GeneralComments_Change()
On Error GoTo Err_Handler
Dim SaveFailed As Boolean

If PVDFileName <> "" Then
    PipelineInfo.Comments = Me.GeneralComments.text
    Call SaveToFilePipeAndConfigInfo("PipelineInfo", SaveFailed)
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-PD17:" & Error$
End Sub

Private Sub GeneralComments_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Handler

    If LanguageCharset <> 0 Then
        RichGeneralComments.Top = GeneralComments.Top
        RichGeneralComments.width = GeneralComments.width
        RichGeneralComments.Left = GeneralComments.Left
        RichGeneralComments.height = GeneralComments.height
        RichGeneralComments.Visible = True
        RichGeneralComments.SetFocus
    End If

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG55:" & Error$
End Sub

Private Sub InternalDiameterExpected_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Handler
    DataEntryChange = True
    Call InternalDiameterEntered(KeyAscii)
Exit Sub
Err_Handler:
    MsgBox Err & "-PD18:" & Error$
End Sub

Private Sub InternalDiameterExpected_lbl_Click() 'PCN4171
On Error GoTo Err_Handler

Call DataEntryForm.SetDataEntryType("IntDiaLabelChange", Me.InternalDiameterExpected_lbl.Caption)

Exit Sub
Err_Handler:
    MsgBox Err & "-PD19:" & Error$
End Sub

Private Sub InternalDiameterExpected_LostFocus()
On Error GoTo Err_Handler
    If DataEntryChange = True Then Call InternalDiameterEntered(vbKeyReturn)
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "20:" & Error$
    End Select
End Sub

Public Sub InternalDiameterExpected_Validate(Cancel As Boolean)
On Error GoTo Err_Handler
    Call InternalDiameterEntered(vbKeyReturn)
Exit Sub
Err_Handler:
Select Case Err
    Case Else: MsgBox Err & "-PD20:" & Error$
End Select
End Sub

Public Sub InternalDiameterEntered(ByVal Key As Integer)
On Error GoTo Err_Handler
    Dim SaveFailed As Boolean
    Dim EnteredDiameter As Single
    Dim i As Long
    Dim InternalDia As Single
    
    Select Case Key
        Case vbKeyEscape: Me.InternalDiameterExpected.text = ExpectedDiameter: Exit Sub
        Case vbKeyReturn
            DataEntryChange = False
            If Me.InternalDiameterExpected.text = "" And PVDFileName <> "" Then
              
                    InternalDia = CalculatePercentile(50, "MedianDiameter")
                    Me.InternalDiameterExpected.text = Round(InternalDia, 2)
                End If

            If Not IsNumeric(Me.InternalDiameterExpected.text) Then Me.InternalDiameterExpected.text = ExpectedDiameter: Exit Sub
            EnteredDiameter = Me.InternalDiameterExpected.text
            If EnteredDiameter <= 0 Then Me.InternalDiameterExpected.text = ExpectedDiameter: Exit Sub
            If EnteredDiameter = ExpectedDiameter Then Exit Sub
            
            PipelineInfo.IntDiameter = EnteredDiameter
            ExpectedDiameter = EnteredDiameter 'PCN3647
            
            If LockedDonut = True Then
                Call DebugForm.LockDonutBtn_Click
            End If
              
                
    
            Call ScreenDrawing.DrawProfilesStartToFinish(ClearLineScreen.SnapShotScreen, True)
            Call ScreenDrawing.DrawProfilesStartToFinish(ClearLineScreen.PVScreen, True)
            
            If PVDFileName = "" Or PVRecording = True Then Exit Sub
            
            Call SaveToFilePipeAndConfigInfo("PipelineInfo", SaveFailed) 'PCN1768
   
            
            
            Call CLPProgressBar.ProgressBarInitialise(DisplayMessage("Processing Data"))
            DoEvents
    
            For i = 0 To (PVDataNoOfLines / 100)
                Call CLPProgressBar.ProgressBarPosition((i / PVDataNoOfLines * 100))
                DoEvents
            
               Call ScreenDrawing.PVDiameterMedianCalcCPP((i * 100), 99 + (i * 100)) 'PCN3540'PCN4974 need to calculate median diameter incase its used with diameter flat.
                Call ScreenDrawing.PVFlat3DCalcCPP((i * 100), 99 + (i * 100)) 'PCN3513
                Call ScreenDrawing.PVCapacityCalcCPP((i * 100), 99 + (i * 100)) 'PCN3540
                Call ScreenDrawing.PVXYDiameterCalcCPP((i * 100), 99 + (i * 100)) 'PCN3540
                Call ScreenDrawing.PVDeltaMaxMinCalcCPP((i * 100), 99 + (i * 100)) 'pcn3540
                Call ScreenDrawing.PVDiameterMaxMinCalcCPP((i * 100), 99 + (i * 100)) 'PCN3540
                Call ScreenDrawing.FixMinMax((i * 100), 99 + (i * 100)) 'PCN6524
'PCN6458                 Call ScreenDrawing.PVInclinationCalc((I * 100), 99 + (I * 100)) 'PCN6128
                
            '    Call ScreenDrawing.PVFractileCalcCPP((i * 100), 99 + (i * 100)) 'PCN3540
                Call ScreenDrawing.PVOvalityCalcCPP((i * 100), 99 + (i * 100)) 'PCN3540
                'Call ScreenDrawing.PVDebrisCalcCPP((i * 100), 99 + (i * 100)) 'PCN4461
                ClearLineScreen.UpdateThreeDimColour
            Next i
'Call CPPSmoothInclination

Call ScreenDrawing.CPPFilterGraphs 'PCN4355
            Call CLPProgressBar.ProgressBarPosition(1)
   
            
            Call ScreenDrawing.DrawPVGraphs

    End Select
    
    If OptionsPage.SmoothCheck.value = vbChecked Then
        ReDim GraphInfoContainer(PVOvalitySmooth).DataSingle(PVDataNoOfLines) 'PCN9999
        ReDim GraphInfoContainer(PVXDiameterSmooth).DataSingle(PVDataNoOfLines)
        ReDim GraphInfoContainer(PVYDiameterSmooth).DataSingle(PVDataNoOfLines)
        ReDim GraphInfoContainer(PVMedianDiameterSmooth).DataSingle(PVDataNoOfLines) 'PCN9999
        ReDim GraphInfoContainer(PVMaxDiameterSmooth).DataSingle(PVDataNoOfLines)
        ReDim GraphInfoContainer(PVMinDiameterSmooth).DataSingle(PVDataNoOfLines)
        ReDim GraphInfoContainer(PVCapacitySmooth).DataSingle(PVDataNoOfLines)
'PCN6458         ReDim GraphInfoContainer(PVInclinationSmooth).DataSingle(PVDataNoOfLines) 'PCN6128

        
        Call CPPSmoothGraphs
'        Call ClearAllGraphsAndRuler
'        Call DrawPVGraphs 'PCN3373
'        Call DrawPVXScale
        
    End If
    
    Call ScreenDrawing.PVDeflectionCalcCPP   'PCN5186
    
    'PCN5186'''''''''''''''''''''
    Call ClearAllGraphsAndRuler '
    Call DrawPVGraphs 'PCN3373  '
    Call DrawPVXScale           '
    '''''''''''''''''''''''''''''
    
    Call PVGraphsKeyForm.DisplayPVGraphsKey 'PCN4920

Exit Sub
Err_Handler:
Select Case Err
    Case 52: Resume Next: 'Bad filename or number 'PCN1863 I haven't worked out why I get this error
    Case 63: Resume Next: 'Bad record number 'PCN1863 I haven't worked out why I get this error
    Case Else: MsgBox Err & "-PD21:" & Error$
    
End Select

End Sub








'Private Sub Material_Click()
'If LanguageCharset <> 0 Then
'    Call DataEntryForm.SetDataEntryType("MaterialTextChange", Me.Material.text)
'End If
'End Sub

Private Sub Material_KeyDown(KeyCode As Integer, Shift As Integer)
If LanguageCharset <> 0 Then
'    Call DataEntryForm.SetDataEntryType("MaterialTextChange", Me.Material.text)
End If
End Sub

Private Sub Material_KeyUp(KeyCode As Integer, Shift As Integer) 'PCN1768
On Error GoTo Err_Handler
Dim SaveFailed As Boolean

DataEntryChange = True

Exit Sub
Err_Handler:
    MsgBox Err & "-PD22:" & Error$
End Sub

Private Sub Material_lbl_Click() 'PCN4171
On Error GoTo Err_Handler

Call DataEntryForm.SetDataEntryType("MaterialLabelChange", Me.Material_lbl.Caption)

Exit Sub
Err_Handler:
   MsgBox Err & "-PD23:" & Error$
End Sub

Private Sub Material_Validate(Cancel As Boolean) 'PCN1768
On Error GoTo Err_Handler
Dim SaveFailed As Boolean

If DataEntryChange = True And PVDFileName <> "" Then
    PipelineInfo.Material = Me.Material
    Call SaveToFilePipeAndConfigInfo("PipelineInfo", SaveFailed)
End If

DataEntryChange = False

Exit Sub
Err_Handler:
    MsgBox Err & "-PD24:" & Error$
End Sub




Private Sub OffsetTextBox_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim FailSave As Boolean
    If Not IsNumeric(OffsetTextBox.text) Then Exit Sub
'PCN6458         Call ScreenDrawing.UpDateInclinationGraph

End Sub

Private Sub OutsideDiameter_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Handler

Dim SaveFailed As Boolean


If OutsideDiameter = "" Then 'PCN1832
    PipelineInfo.ExtDiameter = Empty 'PCN1832
    Call SaveToFilePipeAndConfigInfo("PipelineInfo", SaveFailed)
ElseIf IsNumeric(OutsideDiameter) Then
    PipelineInfo.ExtDiameter = OutsideDiameter
    Call SaveToFilePipeAndConfigInfo("PipelineInfo", SaveFailed)
    Call ScreenDrawing.DrawAll
End If

Exit Sub

Err_Handler:
    MsgBox Err & "-PD25:" & Error$
End Sub

Private Sub OutsideDiameter_lbl_Click() 'PCN4171
On Error GoTo Err_Handler

Call DataEntryForm.SetDataEntryType("OutDiaLabelChange", Me.OutsideDiameter_lbl.Caption)

Exit Sub
Err_Handler:
    MsgBox Err & "-PD26:" & Error$
End Sub

Private Sub OutsideDiameter_Validate(Cancel As Boolean)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'OutsideDiameter_Validate Sub  Michelle Lindsay     michellelindsay@cbsys.co.nz
'
'Revision history"
'   V0.0    Michelle Lindsay,    23/01/2003     Building initial framework
'
'Description:
'
'Purpose: For validating the data entered into the OutsideDiameter field.
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim SaveFailed As Boolean  'PCN1768

If OutsideDiameter = "" Then 'PCN1832
    PipelineInfo.ExtDiameter = Empty 'PCN1832
    'Observations.OutsideDiameter = "" 'PCN1832 'PCN4131
ElseIf Not IsNumeric(OutsideDiameter) And OutsideDiameter <> "" Then 'PCN1832
    'MsgBox DisplayMessage("Outside diameter must be numeric"), vbExclamation 'PCN2111
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Outside diameter must be numeric"): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    'Cancel = True
Else
    PipelineInfo.ExtDiameter = OutsideDiameter
    'Observations.OutsideDiameter = OutsideDiameter 'PCN4131
    Call SaveToFilePipeAndConfigInfo("PipelineInfo", SaveFailed) 'PCN1768
End If
Change = False

Exit Sub
Err_Handler:
    MsgBox Err & "-PD27:" & Error$
End Sub

Private Sub PipeDetailsToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PipeDetailsToolbar_ButtonClick Sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    11/11/02     Building initial framework
'
'Description:
'
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

'Call FisheyeFunctions.LiveFishEyeOFF

'Action on button press
'PCN3217
'PopupReportsToolbar.Visible = False
'PopupViewToolbar.Visible = False

Select Case Button.Key
    Case "OpenFile"
        Call OpenAnyFile("") 'PCN2133
    Case "SaveToFile"
        If PVRecording = True Then 'PCN2379
            'MsgBox DisplayMessage("Stop PVD recording before saving") 'PCN2762
            ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Stop PVD recording before saving"): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
            Exit Sub
        End If
        If Registered = False Then 'PCN1956
            'MsgBox DisplayMessage("Cannot save a .PVD file, please register the software to access this."), vbExclamation 'PCN1956 'PCN2111
            ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Cannot save a .PVD file, please register the software to access this."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
        Else
            Call SaveImageAndOrData 'PCNGL110103
        End If
    Case "Information"
        Dim hwndHelp As Long
        If HelpFilename <> "" And Dir(ReadOnlyAppPath & "Language\" & HelpFilename) = "" Then 'Check whether that file exists actually.'PCN2167 7/8/03 by Abe
            MsgBox HelpFilename & " " & DisplayMessage("file for the language") & _
                "(" & Language & ") " & _
                DisplayMessage("does not exist. Create this file first. The default language(English) is loaded for Help file."), , "Clear Line Profiler" 'PCN2167 7/8/03 by Abe, PCN2171
        End If
        If HelpFilename <> "" And Dir(ReadOnlyAppPath & "Language\" & HelpFilename) <> "" Then  'PCN2167 7/8/03 by Abe ---------v
            hwndHelp = HtmlHelp(hwnd, ReadOnlyAppPath & "Language\" & HelpFilename, HH_DISPLAY_TOPIC, 0)
        Else '-------------------------------------------------------------------------------------------^
            'PCN 1972 LS 8/7/03
            'hwndHelp = HtmlHelp(hwnd, App.Path & "\HelpFile.chm", HH_DISPLAY_TOPIC, 0)
            hwndHelp = HtmlHelp(hwnd, LocToSave & "ProfilerHelp.chm", HH_DISPLAY_TOPIC, 0) 'PCN4400
        End If
    Case "ViewSelection"
        'PCN3217
        Call ScreenDrawing.ToggleViewSelectionPopUp(PipelineDetails)
        'PopupViewToolbar.Visible = True
        'PopupReportsToolbar.Visible = False
    Case "Reports"
        'PCN3217
        Call ScreenDrawing.ToggleReportsPopUp(PipelineDetails)
        'PopupReportsToolbar.Visible = True
        'PopupViewToolbar.Visible = False

' Removed again as GL decided to only have these options on the Options page.
'    Case "Distance"
'        Distance.Show 'PCN2463
'        Distance.ZOrder 0 'PCN2463
'    Case "FishEye" 'FISH-EYE PCN2290
'        Fisheye.Show
'        Fisheye.ZOrder 0
'    Case "ImageProcessing"
'        AutoTune.Show
'        AutoTune.ZOrder 0
    Case Else
End Select


Exit Sub
Err_Handler:
    MsgBox Err & "-PD28:" & Error$

End Sub

Private Sub PipeLen_lbl_Click() 'PCN4171
On Error GoTo Err_Handler

Call DataEntryForm.SetDataEntryType("PipeLenLabelChange", Me.PipeLen_lbl.Caption)

Exit Sub
Err_Handler:
    MsgBox Err & "-PD29:" & Error$
End Sub

Private Sub PipeLength_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Handler
Dim SaveFailed As Boolean

DataEntryChange = True

Exit Sub
Err_Handler:
    MsgBox Err & "-PD30:" & Error$
End Sub

Private Sub PipeLength_Validate(Cancel As Boolean)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PipeLength_Validate Sub     Michelle Lindsay    michellelindsay@cbsys.co.nz
'
'Revision history
'   V0.0    Michelle Lindsay,   23/01/2003    Building initial framework
'
'Description:
'       Validates the PipeLength data entered into the field.
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim SaveFailed As Boolean  'PCN1768

If PipeLength = "" Then 'PCN1832
    PipelineInfo.PipeLength = Empty 'PCN1832
ElseIf Not IsNumeric(PipeLength) And PipeLength <> "" Then 'PCN1832
    'MsgBox DisplayMessage("Pipe length must be numeric"), vbExclamation 'PCN2111
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Pipe length must be numeric"): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Cancel = True
Else
    PipelineInfo.PipeLength = PipeLength
    Call SaveToFilePipeAndConfigInfo("PipelineInfo", SaveFailed) 'PCN1768
End If
Change = False

Exit Sub
Err_Handler:
    MsgBox Err & "-PD31:" & Error$
End Sub

Public Sub PopupReportsToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PopupupReportsToolbar_ButtonClick Sub Michelle Lindsay michellelindsay@cbsys.co.nz
'
'Revision history
'   V0.0    Michelle Lindsay,   12/12/02    Building initial framework
'
'Description:
'       Event for showing correct reports when the user click a button on the
'       pop-up toolbar.
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

On Error GoTo Err_Handler

''Select Case Button.Key
''    'Case SnapShot - Not necessary for this release, ML 07.01.03
''    '    PopupReportsToolbar.Visible = False
''    '    If isopen("LaserImageReport") Then Unload LaserImageReport
''    '    LaserImageReport.Show
''    Case "PVPortrait"
''        'Call PVPortraitReport(Observations) 'PCN2777
''        Load PVReportProfile
''    Case "PVLandscape" 'PCNGL040103
''        Call PVLandscapeReport(PipelineDetails)
''    Case "PVGraph" 'PCN2777
''        'Call PVLandscapeSingleReport(PipelineDetails) 'PCN2777
''        Load PVReportSingle
''    'vvvv PCN3401 ********************************************************
''    Case "4_in_1"
''        'This report shows four different PVGraphs (eg Capacity, Ovality,
''        'Delta or XY and the Flat) on the same page.
''        Call PV4_In_1Report(PipelineDetails)
''    '^^^^ ****************************************************************
''    'vvvv PCN3479 **************************************************
''    Case "MultiProfile"
''        'This report is a multi profile report and is based on the
''        'Profiles at all of the Observations points.
''        Call PVMultiProfileReport(PipelineDetails)
''    '^^^^ **********************************************************
''    Case Else
''
''End Select

Exit Sub
Err_Handler:
    MsgBox Err & "-PD32:" & Error$
End Sub
  
Private Sub RefShape_DropDownBox_OnSelect()
On Error GoTo Error_handler
Dim MsgResp As Variant

If RefShape_DropDownBox.ItemSelected >= 0 Then
    'vvvv PCN4344 ************************************
    If Trim(ReferenceShape(RefShape_DropDownBox.ItemSelected).Use) = "Beta" Then
        'MsgResp = MsgBox(DisplayMessage("The Non-Circular shape functionality is BETA ONLY. Do you wish to proceed?"), vbQuestion + vbYesNo)
        ProfilerMessageBox.MsgBoxYesNo (DisplayMessage("The Non-Circular shape functionality is BETA ONLY. Do you wish to proceed?"))
        MsgResp = PMBAnswer
        If MsgResp <> vbYes Then
            Me.RefShape_DropDownBox.SelectItem = 0
            Exit Sub
        End If
    End If
    '^^^^ ********************************************
    Call SetupForRefShape(ReferenceShape(RefShape_DropDownBox.ItemSelected).Name)
    'Call SetupForRefShape(TheRefShapeFiles(RefShape_DropDownBox.ItemSelected))
    Call ScreenDrawing.DrawProfilesStartToFinish(ClearLineScreen.PVScreen, True)
    Call ScreenDrawing.PVFlat3DCalcCPP(0, PVDataNoOfLines)
    If DrawShapeType = "Circle" Then OptionsPage.MedianDiameterOpt(0).Enabled = True
    If DrawShapeType <> "Circle" Then OptionsPage.MedianDiameterOpt(0).Enabled = False
End If

Exit Sub
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-PD33:" & Error$, vbExclamation
    End Select
End Sub


Private Sub RichGeneralComments_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 27:
            If LanguageCharset <> 0 Then
                UnicodeGeneralComments = RichGeneralComments.text
                GeneralComments.Font.Charset = LanguageCharset
                GeneralComments.text = UnicodeGeneralComments
            End If
            RichGeneralComments.Visible = False
        Case vbKeyReturn
            If LanguageCharset <> 0 Then
                UnicodeGeneralComments = LanguageUtil.ConvertRichToAnsi(RichGeneralComments.TextRTF)
                GeneralComments.Font.Charset = LanguageCharset
                GeneralComments.text = UnicodeGeneralComments
            End If
            RichGeneralComments.Visible = False
    End Select
End Sub

Private Sub sDate_Click()
If LanguageCharset <> 0 Then
    Call DataEntryForm.SetDataEntryType("DateTextChange", Me.sDate.text)
End If
End Sub

Private Sub sDate_KeyDown(KeyCode As Integer, Shift As Integer)
If LanguageCharset <> 0 Then
    Call DataEntryForm.SetDataEntryType("DateTextChange", Me.sDate.text)
End If
End Sub

Private Sub sDate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'sDate_MouseDown Sub     Michelle Lindsay    michellelindsay@cbsys.co.nz
'
'Revision history
'   V0.0    Michelle Lindsay,   23/01/2003    Adding functionality
'
'Description:
'
'Purpose:
'Detect if any details are entered into the sDate field.
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Change = True

Exit Sub
Err_Handler:
    MsgBox Err & "-PD34:" & Error$
    

End Sub

Private Sub sDate_Validate(Cancel As Boolean)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'sDate_Validate Sub      Michelle Lindsay       michellelindsay@cbsys.co.nz
'
'Revision history"
'   V0.0    Michelle Lindsay,    23/01/2003     Building initial framework
'
'Description:
'
'Purpose: For validating the data entered into the sDate field.
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~On Error GoTo Err_handler
Dim SaveFailed As Boolean  'PCN1768

If sDate = "" Then 'PCN1832
    PipelineInfo.Date = Empty 'PCN1832
ElseIf Not IsDate(sDate) And sDate <> "" Then 'PCN1832
    'MsgBox DisplayMessage("Must enter a valid date"), vbExclamation 'PCN2111
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Must enter a valid date"): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Cancel = True
Else
    PipelineInfo.Date = sDate
    Call SaveToFilePipeAndConfigInfo("PipelineInfo", SaveFailed) 'PCN1768
End If
Change = False

Exit Sub
Err_Handler:
    MsgBox Err & "-PD35:" & Error$
End Sub



Private Sub SeaLevelStartHeight_Change()

End Sub



'PCN6458
'Private Sub SeaLevelEndHeightTextBox_KeyUp(KeyCode As Integer, Shift As Integer)
'On Error GoTo Err_Handler
'
'   Dim FailSave As Boolean
'
'    If IsNumeric(SeaLevelEndHeightTextBox.text) Then
'        'SeaLevelEndHeight = CDbl(SeaLevelEndHeightTextBox.text)
'        Call ScreenDrawing.UpDateInclinationGraph
'        'FontInfo.FontColour = SeaLevelEndHeightTextBox.text
'        'FailSave = SaveToFileFontInfo(FailSave)
'    End If
'
'    If Me.SeaLevelEndHeightTextBox.text = "" And KeyCode = vbKeyReturn Then
'        Call ScreenDrawing.UpDateInclinationGraph
'    End If
'
'Exit Sub
'Err_Handler:
'    MsgBox Err & "-PD35.5"
'End Sub

'PCN6458
'Private Sub SeaLevelStartHeightTextBox_KeyUp(KeyCode As Integer, Shift As Integer)
'On Error GoTo Err_Handler
'
'    Dim FailSave As Boolean
'
'    If IsNumeric(SeaLevelStartHeightTextBox.text) Then
'        'SeaLevelStartHeight = CDbl(SeaLevelStartHeightTextBox.text)
'        Call ScreenDrawing.UpDateInclinationGraph
'        'FontInfo.FontType = SeaLevelStartHeightTextBox.text
'        'FailSave = SaveToFileFontInfo(FailSave)
'    End If
'
'    If Me.SeaLevelStartHeightTextBox.text = "" And KeyCode = vbKeyReturn Then
'        Call ScreenDrawing.UpDateInclinationGraph
'    End If
'
'Exit Sub
'Err_Handler:
'    MsgBox Err & "-PD35.5"
'End Sub

'Private Sub SiteID_Click()
'If LanguageCharset <> 0 Then
'    Call DataEntryForm.SetDataEntryType("SiteIDTextChange", Me.SiteID.text)
'End If
'End Sub

Private Sub SiteID_KeyDown(KeyCode As Integer, Shift As Integer)
If LanguageCharset <> 0 Then
    'Call DataEntryForm.SetDataEntryType("SiteIDTextChange", Me.SiteID.text)
End If
End Sub

Private Sub SiteID_KeyUp(KeyCode As Integer, Shift As Integer) 'PCN1768
On Error GoTo Err_Handler
Dim SaveFailed As Boolean

DataEntryChange = True

Exit Sub
Err_Handler:
    MsgBox Err & "-PD36:" & Error$
End Sub

Private Sub SiteID_lbl_Click() 'PCN4171
On Error GoTo Err_Handler

Call DataEntryForm.SetDataEntryType("SiteIDLabelChange", Me.SiteID_lbl.Caption)

Exit Sub
Err_Handler:
    MsgBox Err & "-PD37:" & Error$
End Sub

Private Sub SiteID_Validate(Cancel As Boolean) 'PCN1768
On Error GoTo Err_Handler
Dim SaveFailed As Boolean

'PCN2189
'If DataEntryChange = True And PVDFileName <> "" Then
'    PipelineInfo.SiteID = Me.SiteID
'    Call SaveToFilePipeAndConfigInfo("PipelineInfo", SaveFailed)
'End If

If SiteID = "" Then 'PCN1832
    PipelineInfo.SiteID = Empty 'PCN1832
    'Observations.SiteID = "" 'PCN1832 'PCN4131
Else
    PipelineInfo.SiteID = SiteID
    'Observations.SiteID = SiteID 'PCN4131
    Call SaveToFilePipeAndConfigInfo("PipelineInfo", SaveFailed) 'PCN1768
End If
Change = False

DataEntryChange = False

Exit Sub
Err_Handler:
    MsgBox Err & "-PD38:" & Error$
End Sub



'Private Sub StartNodeLocation_Click()
'If LanguageCharset <> 0 Then
'    Call DataEntryForm.SetDataEntryType("StNodeLocTextChange", Me.StartNodeLocation.text)
'End If
'End Sub

Private Sub StartNodeLocation_KeyDown(KeyCode As Integer, Shift As Integer)
If LanguageCharset <> 0 Then
'    Call DataEntryForm.SetDataEntryType("StNodeLocTextChange", Me.StartNodeLocation.text)
End If
End Sub

Private Sub StartNodeLocation_KeyUp(KeyCode As Integer, Shift As Integer) 'PCN1768
On Error GoTo Err_Handler
Dim SaveFailed As Boolean

DataEntryChange = True

Exit Sub
Err_Handler:
    MsgBox Err & "40:" & Error$
End Sub

Private Sub StartNodeLocation_Validate(Cancel As Boolean) 'PCN1768
On Error GoTo Err_Handler
Dim SaveFailed As Boolean

If DataEntryChange = True And PVDFileName <> "" Then
    PipelineInfo.StartLocation = Me.StartNodeLocation
    Call SaveToFilePipeAndConfigInfo("PipelineInfo", SaveFailed)
End If

DataEntryChange = False

Exit Sub
Err_Handler:
    MsgBox Err & "41:" & Error$
End Sub

'Private Sub StartNodeNo_Click()
'If LanguageCharset <> 0 Then
'    Call DataEntryForm.SetDataEntryType("StNodeTextChange", Me.StartNodeNo.text)
'End If
'End Sub

Private Sub StartNodeNo_KeyDown(KeyCode As Integer, Shift As Integer)
If LanguageCharset <> 0 Then
'    Call DataEntryForm.SetDataEntryType("StNodeTextChange", Me.StartNodeNo.text)
End If
End Sub

Private Sub StartNodeNo_KeyUp(KeyCode As Integer, Shift As Integer) 'PCN1768
On Error GoTo Err_Handler
Dim SaveFailed As Boolean

DataEntryChange = True

Exit Sub
Err_Handler:
    MsgBox Err & "-PD39:" & Error$
End Sub

Private Sub StartNodeNo_lbl_Click() 'PCN4171
On Error GoTo Err_Handler

Call DataEntryForm.SetDataEntryType("StNodeLabelChange", Me.StartNodeNo_lbl.Caption)

Exit Sub
Err_Handler:
    MsgBox Err & "-PD40:" & Error$
End Sub

Private Sub StartNodeNo_Validate(Cancel As Boolean)
On Error GoTo Err_Handler
Dim SaveFailed As Boolean

'PCN2189
'If DataEntryChange = True And PVDFileName <> "" Then
'    PipelineInfo.StartName = Me.StartNodeNo
'    Call SaveToFilePipeAndConfigInfo("PipelineInfo", SaveFailed)
'End If

If StartNodeNo = "" Then 'PCN1832
    PipelineInfo.StartName = Empty 'PCN1832
    'Observations.StartNodeNo = "" 'PCN1832 'PCN4131
Else
    PipelineInfo.StartName = StartNodeNo
    'Observations.StartNodeNo = StartNodeNo 'PCN4131
    Call SaveToFilePipeAndConfigInfo("PipelineInfo", SaveFailed) 'PCN1768
End If
Change = False
DataEntryChange = False

Exit Sub
Err_Handler:
    MsgBox Err & "-PD41:" & Error$
End Sub


Private Sub sTime_Click()
If LanguageCharset <> 0 Then
    Call DataEntryForm.SetDataEntryType("TimeTextChange", Me.sTime.text)
End If
End Sub

Private Sub sTime_KeyDown(KeyCode As Integer, Shift As Integer)
If LanguageCharset <> 0 Then
    Call DataEntryForm.SetDataEntryType("TimeTextChange", Me.sTime.text)
End If
End Sub

Private Sub sTime_Validate(Cancel As Boolean)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'sTime_Validate Sub     Michelle Lindsay        michellelindsay@cbsys.co.nz
'
'Revision history"
'   V0.0    Michelle Lindsay,    23/01/2003     Building initial framework
'
'Description:
'
'Purpose: For validating the data entered into the sTime field.
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim SaveFailed As Boolean 'PCN1768
 
If sTime = "" Then 'PCN1832
    PipelineInfo.Time = Empty 'PCN1832
ElseIf Not IsDate(sTime) And sTime <> "" Then 'PCN1832
    'MsgBox DisplayMessage("Must enter a valid time"), vbExclamation 'PCN2111
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Must enter a valid time"): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Cancel = True
Else
    PipelineInfo.Time = sTime
    Call SaveToFilePipeAndConfigInfo("PipelineInfo", SaveFailed) 'PCN1768
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-PD42:" & Error$
End Sub


Private Sub StNodeLoc_lbl_Click() 'PCN4171
On Error GoTo Err_Handler

Call DataEntryForm.SetDataEntryType("StNodeLocLabelChange", Me.StNodeLoc_lbl.Caption)

Exit Sub
Err_Handler:
    MsgBox Err & "-PD43:" & Error$
End Sub

Private Sub Time_lbl_Click() 'PCN4171
On Error GoTo Err_Handler

Call DataEntryForm.SetDataEntryType("TimeLabelChange", Me.Time_lbl.Caption)

Exit Sub
Err_Handler:
    MsgBox Err & "-PD44:" & Error$
End Sub

Sub PopulateRefShapeDropDown(RefShapeControl As Control)
On Error GoTo Error_handler
    
Dim RefShapeType As String


ReDim TheRefShapeFiles(0)
Call LoadRefShapes(RefShapeControl)


RefShapeControl.Font.Size = 10
RefShapeControl.Font = "MS Sans Serif"
RefShapeControl.Font.Charset = LanguageCharset
RefShapeControl.Font.Bold = True

RefShapeControl.text = DisplayMessage("Circle")

    
    
Exit Sub
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-PD45:" & Error$, vbExclamation
    End Select
End Sub

Sub LoadRefShapes(RefShapeControl As Control)
On Error GoTo Error_handler
    
Dim RefShapeTextFiles As String
    Dim RefShapeFileName As String
    Dim RefShapeType As String
    Dim CountRefShapeFiles As Integer
    Dim i As Integer


CountRefShapeFiles = UBound(ReferenceShape)

ReDim TheRefShapeFiles(CountRefShapeFiles)
ReDim TheRefShapeFileLabels(CountRefShapeFiles)


    

'!!!!!!!!!!!!To add directory and files
'ReDim TheRefShapeFiles(14) 'PCNGL170107
'ReDim TheRefShapeFileLabels(14) 'PCNGL170107
'TheRefShapeFiles(0) = "Circle"
'TheRefShapeFiles(1) = "Egg"
'TheRefShapeFiles(2) = "RinkerEllipse"
'TheRefShapeFiles(3) = "Elliptical ASTM C507"
'TheRefShapeFiles(4) = "HorseShoe"
'TheRefShapeFiles(5) = "Huia Loaf"
'TheRefShapeFiles(6) = "Redoubt"
'TheRefShapeFiles(7) = "Loaf"
'TheRefShapeFiles(8) = "EggA"
'TheRefShapeFiles(9) = "EggB"
'TheRefShapeFiles(10) = "CupCake"
'TheRefShapeFiles(11) = "Bullet"
'TheRefShapeFiles(12) = "Square"
'TheRefShapeFiles(13) = "Mushroom"
'TheRefShapeFiles(14) = "COSRehab"

For i = 0 To CountRefShapeFiles
    TheRefShapeFileLabels(i) = DisplayMessage(ReferenceShape(i).Name)
    Call RefShapeControl.AddItem(TheRefShapeFileLabels(i))
    
Next i


'TheRefShapeFileLabels(0) = DisplayMessage("Circle")
'TheRefShapeFileLabels(1) = DisplayMessage("Egg")
'TheRefShapeFileLabels(2) = DisplayMessage("RinkerEllipse")
'TheRefShapeFileLabels(3) = DisplayMessage("Elliptical ASTM C507")
'TheRefShapeFileLabels(4) = DisplayMessage("HorseShoe")
'TheRefShapeFileLabels(5) = DisplayMessage("Huia Loaf")
'TheRefShapeFileLabels(6) = DisplayMessage("Redoubt")
'TheRefShapeFileLabels(7) = DisplayMessage("Loaf")
'TheRefShapeFileLabels(8) = DisplayMessage("EggA")
'TheRefShapeFileLabels(9) = DisplayMessage("EggB")
'TheRefShapeFileLabels(10) = DisplayMessage("CupCake")
'TheRefShapeFileLabels(11) = DisplayMessage("Bullet")
'TheRefShapeFileLabels(12) = DisplayMessage("Square")
'TheRefShapeFileLabels(13) = DisplayMessage("Mushroom")
'TheRefShapeFileLabels(14) = DisplayMessage("COSRehab")
'
'RefShapeControl.AddItem (TheRefShapeFileLabels(0))
'RefShapeControl.AddItem (TheRefShapeFileLabels(1))
'RefShapeControl.AddItem (TheRefShapeFileLabels(2))
'RefShapeControl.AddItem (TheRefShapeFileLabels(3))
'RefShapeControl.AddItem (TheRefShapeFileLabels(4))
'RefShapeControl.AddItem (TheRefShapeFileLabels(5))
'RefShapeControl.AddItem (TheRefShapeFileLabels(6))
'RefShapeControl.AddItem (TheRefShapeFileLabels(7))
'RefShapeControl.AddItem (TheRefShapeFileLabels(8))
'RefShapeControl.AddItem (TheRefShapeFileLabels(9))
'RefShapeControl.AddItem (TheRefShapeFileLabels(10))
'RefShapeControl.AddItem (TheRefShapeFileLabels(11))
'RefShapeControl.AddItem (TheRefShapeFileLabels(12))
'RefShapeControl.AddItem (TheRefShapeFileLabels(13))
'RefShapeControl.AddItem (TheRefShapeFileLabels(14))

Exit Sub
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-PD46:" & Error$, vbExclamation
    End Select
End Sub

Sub SetupForRefShape(RefShapeType As String)
On Error GoTo Err_Handler

DrawShapeType = Trim(RefShapeType)

'Select Case RefShapeType
'    Case "Circle":    DrawShapeType = "Circle"
'    Case "Egg":       DrawShapeType = "Egg"
'    Case "HorseShoe": DrawShapeType = "SemiElliptical"
'    Case "Bullet":    DrawShapeType = "Bullet"
'    Case "Elliptical ASTM C507":     DrawShapeType = "Elliptical ASTM C507"
'    Case "Huia Loaf":     DrawShapeType = "Huia Loaf"
'    Case "Redoubt":     DrawShapeType = "Redoubt"
'    Case "Loaf":     DrawShapeType = "Loaf"
'    Case "EggA":     DrawShapeType = "EggA"
'    Case "EggB":     DrawShapeType = "EggB"
'    Case "CupCake": DrawShapeType = "CupCake"
'    Case "RinkerEllipse": DrawShapeType = "RinkerEllipse"
'    Case "Square": DrawShapeType = "Square"
'    Case "Mushroom": DrawShapeType = "Mushroom"
'
'    Case "COSRehab": DrawShapeType = "COSRehab"
'
'    Case Else:        DrawShapeType = "Circle"
'
'End Select
If mediatype <> "" Then
    Call ScreenDrawing.DrawProfilesStartToFinish(ClearLineScreen.SnapShotScreen, True)
    Call ScreenDrawing.DrawProfilesStartToFinish(ClearLineScreen.PVScreen, True)
    Call ControlsScreen.ControlsViewSetup
End If


    
Exit Sub
Err_Handler:
    MsgBox Err & "-PD47:" & Error$
End Sub


