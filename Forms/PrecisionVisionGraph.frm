VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form PrecisionVisionGraph 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   13155
   ClientLeft      =   19995
   ClientTop       =   0
   ClientWidth     =   16050
   ControlBox      =   0   'False
   Icon            =   "PrecisionVisionGraph.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   13155
   ScaleWidth      =   16050
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ButtonsSmallImages 
      Left            =   2040
      Top             =   9120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrecisionVisionGraph.frx":038A
            Key             =   "More"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrecisionVisionGraph.frx":0B04
            Key             =   "MoreHighlighted"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ButtonImages 
      Left            =   840
      Top             =   9120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrecisionVisionGraph.frx":127E
            Key             =   "YScale"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrecisionVisionGraph.frx":1B58
            Key             =   "YScaleHighlight"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox GraphTypeSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   350
      Left            =   2880
      Picture         =   "PrecisionVisionGraph.frx":2832
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   112
      Top             =   0
      Width           =   350
   End
   Begin VB.PictureBox RulerGreen 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   3480
      Picture         =   "PrecisionVisionGraph.frx":2F9C
      ScaleHeight     =   60
      ScaleWidth      =   240
      TabIndex        =   111
      Top             =   360
      Width           =   240
   End
   Begin VB.PictureBox RulerRed 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   3600
      Picture         =   "PrecisionVisionGraph.frx":309E
      ScaleHeight     =   60
      ScaleWidth      =   240
      TabIndex        =   110
      Top             =   8520
      Width           =   240
   End
   Begin VB.PictureBox LargeRedRuler 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   5040
      Picture         =   "PrecisionVisionGraph.frx":31A0
      ScaleHeight     =   135
      ScaleWidth      =   255
      TabIndex        =   109
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox LargeGreenRuler 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   4080
      Picture         =   "PrecisionVisionGraph.frx":33B6
      ScaleHeight     =   135
      ScaleWidth      =   255
      TabIndex        =   108
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox ObsPreviewPictureBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   4200
      ScaleHeight     =   2055
      ScaleWidth      =   3015
      TabIndex        =   102
      Top             =   5160
      Visible         =   0   'False
      Width           =   3015
      Begin VB.PictureBox ReferenceShiftControls 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   2655
         ScaleHeight     =   1215
         ScaleWidth      =   270
         TabIndex        =   121
         Top             =   240
         Visible         =   0   'False
         Width           =   275
         Begin VB.CheckBox AdjustRefShapeCentre 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Check1"
            CausesValidation=   0   'False
            Height          =   255
            Left            =   30
            TabIndex        =   124
            Top             =   0
            Width           =   195
         End
         Begin VB.OptionButton ReferenceShiftType 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Option1"
            Height          =   195
            Index           =   0
            Left            =   30
            TabIndex        =   123
            Top             =   480
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.OptionButton ReferenceShiftType 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Option2"
            CausesValidation=   0   'False
            Height          =   255
            Index           =   1
            Left            =   30
            TabIndex        =   122
            Top             =   240
            Width           =   195
         End
      End
      Begin VB.PictureBox ObsPreview 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1721
         Left            =   360
         ScaleHeight     =   1695
         ScaleWidth      =   2265
         TabIndex        =   103
         Top             =   120
         Width           =   2295
      End
      Begin VB.PictureBox WaterShiftControls 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   15
         ScaleHeight     =   975
         ScaleWidth      =   255
         TabIndex        =   117
         Top             =   240
         Visible         =   0   'False
         Width           =   255
         Begin VB.CheckBox AdjustWaterCentre 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Check1"
            Height          =   195
            Left            =   70
            TabIndex        =   120
            Top             =   0
            Width           =   195
         End
         Begin VB.OptionButton WaterShiftType 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Option1"
            Height          =   195
            Index           =   0
            Left            =   70
            TabIndex        =   119
            Top             =   480
            Width           =   255
         End
         Begin VB.OptionButton WaterShiftType 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Option2"
            Height          =   195
            Index           =   1
            Left            =   70
            TabIndex        =   118
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Shape GrayImage 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         Height          =   2175
         Left            =   2920
         Top             =   0
         Width           =   255
      End
      Begin VB.Image PaperClipAttachment 
         Height          =   480
         Left            =   30
         Picture         =   "PrecisionVisionGraph.frx":35CC
         Top             =   1320
         Width           =   300
      End
      Begin VB.Image PaperClipHighLight 
         Height          =   510
         Left            =   30
         Picture         =   "PrecisionVisionGraph.frx":3677
         Stretch         =   -1  'True
         Top             =   1290
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  'Opaque
         Height          =   2055
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.PictureBox LargeRuler 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   4680
      ScaleHeight     =   105
      ScaleWidth      =   225
      TabIndex        =   107
      Top             =   120
      Width           =   255
      Begin VB.Image LargeRulerImage 
         Height          =   225
         Left            =   0
         Picture         =   "PrecisionVisionGraph.frx":36D8
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox RulerMark 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Index           =   0
      Left            =   6960
      Picture         =   "PrecisionVisionGraph.frx":38AA
      ScaleHeight     =   60
      ScaleWidth      =   240
      TabIndex        =   106
      Top             =   120
      Width           =   240
   End
   Begin VB.PictureBox ObsTempPicture 
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6120
      ScaleHeight     =   465
      ScaleWidth      =   585
      TabIndex        =   105
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox ObsPictureStorage 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   5280
      ScaleHeight     =   465
      ScaleWidth      =   585
      TabIndex        =   104
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox CalculateNoLines 
      Height          =   285
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   101
      Text            =   "PrecisionVisionGraph.frx":3A7C
      Top             =   10560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PVGraphSelectBox 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00FFFFFF&
      Height          =   2430
      Left            =   840
      ScaleHeight     =   2400
      ScaleWidth      =   2145
      TabIndex        =   92
      Top             =   360
      Visible         =   0   'False
      Width           =   2175
      Begin VB.Label PVGraphSelect 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inclination"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   125
         Tag             =   "Inclination"
         Top             =   1800
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label FlatTM1 
         BackStyle       =   0  'Transparent
         Caption         =   "TM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   113
         Top             =   120
         Width           =   495
      End
      Begin VB.Label PVGraphSelect 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Min Diameter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   99
         Tag             =   "MinDiameter"
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label PVGraphSelect 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Capacity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   98
         Tag             =   "Capacity"
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label PVGraphSelect 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "XY Diameter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   97
         Tag             =   "XYDiameter"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label PVGraphSelect 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Max Diameter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   96
         Tag             =   "MaxDiameter"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label PVGraphSelect 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ovality"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   95
         Tag             =   "Ovality"
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label PVGraphSelect 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "True Diameter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   94
         Tag             =   "MedianDiameter"
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label PVGraphSelect 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Flat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   93
         Tag             =   "Flat"
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.PictureBox PVYScaleSelectorBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2265
      ScaleWidth      =   480
      TabIndex        =   86
      Top             =   5760
      Visible         =   0   'False
      Width           =   515
      Begin VB.Label PVYScaleSetting 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "x1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   15
         TabIndex        =   100
         Top             =   1920
         Width           =   450
      End
      Begin VB.Label PVYScaleSetting 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "x50"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   15
         TabIndex        =   91
         Top             =   120
         Width           =   450
      End
      Begin VB.Label PVYScaleSetting 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "x25"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   15
         TabIndex        =   90
         Top             =   480
         Width           =   450
      End
      Begin VB.Label PVYScaleSetting 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "x10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   15
         TabIndex        =   89
         Top             =   840
         Width           =   450
      End
      Begin VB.Label PVYScaleSetting 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "x5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   15
         TabIndex        =   88
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label PVYScaleSetting 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "x2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   15
         TabIndex        =   87
         Top             =   1560
         Width           =   450
      End
   End
   Begin VB.PictureBox HistogramOfDiameter 
      BackColor       =   &H00400000&
      Height          =   10000
      Left            =   14640
      ScaleHeight     =   663
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   181
      TabIndex        =   62
      Top             =   1200
      Visible         =   0   'False
      Width           =   2775
      Begin VB.PictureBox HistogramOfDiameterGraph 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   9600
         Left            =   120
         ScaleHeight     =   638
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   911
         TabIndex        =   63
         Top             =   120
         Width           =   13695
         Begin VB.Label DiameterHistogramTitle 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Diameter Histogram"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   375
            Left            =   4560
            TabIndex        =   68
            Top             =   120
            Width           =   3855
         End
         Begin VB.Label HistogramKeyMedian 
            BackStyle       =   0  'Transparent
            Caption         =   "Median Diameter"
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
            Left            =   600
            TabIndex        =   67
            Top             =   840
            Width           =   3015
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            X1              =   16
            X2              =   32
            Y1              =   64
            Y2              =   64
         End
         Begin VB.Label HistogramKeyMax 
            BackStyle       =   0  'Transparent
            Caption         =   "Maximum"
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
            Left            =   600
            TabIndex        =   66
            Top             =   1200
            Width           =   3015
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00800000&
            BorderWidth     =   2
            X1              =   16
            X2              =   32
            Y1              =   88
            Y2              =   88
         End
         Begin VB.Label HistogramKeyMin 
            BackStyle       =   0  'Transparent
            Caption         =   "Minimum"
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
            Left            =   600
            TabIndex        =   65
            Top             =   1560
            Width           =   3015
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00004000&
            BorderWidth     =   2
            X1              =   16
            X2              =   32
            Y1              =   112
            Y2              =   112
         End
         Begin VB.Label DiameterHistogramSubTitle 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Diameter Histogram"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   9720
            TabIndex        =   64
            Top             =   7920
            Width           =   2415
         End
      End
      Begin VB.Image HistogramOfDiameterClose 
         Height          =   255
         Left            =   14640
         Picture         =   "PrecisionVisionGraph.frx":3A82
         Top             =   120
         Width           =   240
      End
   End
   Begin VB.PictureBox PVGraph3DBoarder 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Index           =   3
      Left            =   8640
      Picture         =   "PrecisionVisionGraph.frx":3DF4
      ScaleHeight     =   465
      ScaleWidth      =   60
      TabIndex        =   81
      Top             =   0
      Width           =   60
   End
   Begin VB.PictureBox PVGraph3DBoarder 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Index           =   2
      Left            =   8400
      Picture         =   "PrecisionVisionGraph.frx":3FAA
      ScaleHeight     =   465
      ScaleWidth      =   60
      TabIndex        =   80
      Top             =   0
      Width           =   60
   End
   Begin VB.PictureBox PVGraph3DBoarder 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Index           =   1
      Left            =   8160
      Picture         =   "PrecisionVisionGraph.frx":4160
      ScaleHeight     =   465
      ScaleWidth      =   60
      TabIndex        =   79
      Top             =   0
      Width           =   60
   End
   Begin VB.PictureBox PVGraph3DBoarder 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Index           =   0
      Left            =   7920
      Picture         =   "PrecisionVisionGraph.frx":4316
      ScaleHeight     =   465
      ScaleWidth      =   60
      TabIndex        =   78
      Top             =   0
      Width           =   60
   End
   Begin VB.PictureBox Observation 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   5880
      ScaleHeight     =   3255
      ScaleWidth      =   3015
      TabIndex        =   72
      Top             =   1920
      Visible         =   0   'False
      Width           =   3015
      Begin RichTextLib.RichTextBox RichPVObsEntry 
         Height          =   855
         Left            =   360
         TabIndex        =   126
         Top             =   1920
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1508
         _Version        =   393217
         Enabled         =   -1  'True
         Appearance      =   0
         TextRTF         =   $"PrecisionVisionGraph.frx":44CC
      End
      Begin VB.TextBox PVObsEntry 
         Appearance      =   0  'Flat
         Height          =   675
         Left            =   360
         MaxLength       =   88
         MultiLine       =   -1  'True
         TabIndex        =   74
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox PVObsIndexNo 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   285
         Index           =   1
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox PVObsDistance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0F0F0&
         Height          =   270
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   75
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Image ProfileReportObsButtonImage 
         Height          =   240
         Left            =   2040
         Picture         =   "PrecisionVisionGraph.frx":454E
         Top             =   120
         Width           =   195
      End
      Begin VB.Image DeleteObs 
         Height          =   225
         Left            =   1680
         Picture         =   "PrecisionVisionGraph.frx":4810
         Top             =   120
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image ObservationClose 
         Height          =   240
         Left            =   2640
         Picture         =   "PrecisionVisionGraph.frx":4B22
         Top             =   120
         Width           =   240
      End
      Begin VB.Shape SpeakObsBlank 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   300
         Left            =   2760
         Top             =   320
         Width           =   270
      End
      Begin VB.Image ObservationEye 
         Height          =   240
         Index           =   1
         Left            =   80
         Picture         =   "PrecisionVisionGraph.frx":4E64
         Top             =   120
         Width           =   240
      End
      Begin VB.Label ObservationText 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Obs 1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   360
         MousePointer    =   3  'I-Beam
         TabIndex        =   77
         Top             =   120
         Width           =   2295
         WordWrap        =   -1  'True
      End
      Begin VB.Label PVObsDistanceUnit 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1455
         TabIndex        =   76
         Top             =   1440
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape ObservationBubble 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  'Opaque
         Height          =   3255
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   3015
      End
      Begin VB.Shape ObsGraphBackground 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   7995
         Left            =   0
         Top             =   -840
         Width           =   2895
      End
   End
   Begin VB.PictureBox SlideRuleXScale 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   400
      Left            =   4440
      Picture         =   "PrecisionVisionGraph.frx":51EE
      ScaleHeight     =   405
      ScaleWidth      =   2970
      TabIndex        =   40
      Tag             =   "0"
      Top             =   8040
      Visible         =   0   'False
      Width           =   2970
      Begin VB.PictureBox SlideRuleSlider 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   270
         ScaleHeight     =   330
         ScaleWidth      =   1995
         TabIndex        =   41
         Top             =   30
         Width           =   2000
         Begin VB.PictureBox SlideRuleEdgeLeft 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   0
            MousePointer    =   9  'Size W E
            ScaleHeight     =   495
            ScaleWidth      =   45
            TabIndex        =   43
            Top             =   0
            Width           =   50
         End
         Begin VB.PictureBox SlideRuleEdgeRight 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   1950
            MousePointer    =   9  'Size W E
            ScaleHeight     =   495
            ScaleWidth      =   45
            TabIndex        =   42
            Top             =   0
            Width           =   50
         End
         Begin VB.Image SliderCentreMarker 
            Height          =   105
            Left            =   960
            Picture         =   "PrecisionVisionGraph.frx":910C
            Top             =   10
            Width           =   105
         End
         Begin VB.Line SliderCentreMarkerLine 
            BorderColor     =   &H000000FF&
            BorderWidth     =   2
            X1              =   1000
            X2              =   1000
            Y1              =   480
            Y2              =   20
         End
         Begin VB.Image SlideRuleCoverImage 
            Height          =   330
            Left            =   -240
            Picture         =   "PrecisionVisionGraph.frx":91F6
            Top             =   0
            Width           =   2895
         End
      End
   End
   Begin VB.PictureBox GraphCaptionBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   350
      Index           =   5
      Left            =   7560
      MouseIcon       =   "PrecisionVisionGraph.frx":C410
      MousePointer    =   99  'Custom
      ScaleHeight     =   345
      ScaleWidth      =   2895
      TabIndex        =   59
      Top             =   9840
      Width           =   2895
      Begin VB.Image ReportFlagImage 
         Height          =   240
         Index           =   5
         Left            =   2520
         Picture         =   "PrecisionVisionGraph.frx":C99A
         Tag             =   "True"
         Top             =   15
         Width           =   240
      End
      Begin VB.Label Label_GraphName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Flat3D"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   5
         Left            =   0
         MouseIcon       =   "PrecisionVisionGraph.frx":CD24
         TabIndex        =   60
         Top             =   0
         Width           =   2505
      End
      Begin VB.Label Label_GraphNameShadow 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Flat3D"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   5
         Left            =   10
         TabIndex        =   61
         Top             =   -10
         Width           =   2580
      End
   End
   Begin VB.PictureBox GraphCaptionBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   350
      Index           =   4
      Left            =   7560
      MouseIcon       =   "PrecisionVisionGraph.frx":D2AE
      MousePointer    =   99  'Custom
      ScaleHeight     =   345
      ScaleWidth      =   2895
      TabIndex        =   56
      Top             =   9480
      Width           =   2895
      Begin VB.Image ReportFlagImage 
         Height          =   240
         Index           =   4
         Left            =   2520
         Picture         =   "PrecisionVisionGraph.frx":D838
         Tag             =   "True"
         Top             =   15
         Width           =   240
      End
      Begin VB.Label Label_GraphName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "X Y Diameter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   4
         Left            =   0
         MouseIcon       =   "PrecisionVisionGraph.frx":DBC2
         TabIndex        =   57
         Top             =   0
         Width           =   2505
      End
      Begin VB.Label Label_GraphNameShadow 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "X Y Diameter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   4
         Left            =   10
         TabIndex        =   58
         Top             =   -10
         Width           =   2580
      End
   End
   Begin VB.PictureBox GraphCaptionBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   350
      Index           =   3
      Left            =   7560
      MouseIcon       =   "PrecisionVisionGraph.frx":E14C
      MousePointer    =   99  'Custom
      ScaleHeight     =   345
      ScaleWidth      =   2895
      TabIndex        =   53
      Top             =   9120
      Width           =   2895
      Begin VB.Image ReportFlagImage 
         Height          =   240
         Index           =   3
         Left            =   2520
         Picture         =   "PrecisionVisionGraph.frx":E6D6
         Tag             =   "True"
         Top             =   15
         Width           =   240
      End
      Begin VB.Label Label_GraphName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Delta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   0
         MouseIcon       =   "PrecisionVisionGraph.frx":EA60
         TabIndex        =   54
         Top             =   0
         Width           =   2505
      End
      Begin VB.Label Label_GraphNameShadow 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Delta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   3
         Left            =   10
         TabIndex        =   55
         Top             =   -10
         Width           =   2580
      End
   End
   Begin VB.PictureBox GraphCaptionBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   350
      Index           =   2
      Left            =   7560
      MouseIcon       =   "PrecisionVisionGraph.frx":EFEA
      MousePointer    =   99  'Custom
      ScaleHeight     =   345
      ScaleWidth      =   2895
      TabIndex        =   50
      Top             =   8760
      Width           =   2895
      Begin VB.Image ReportFlagImage 
         Height          =   240
         Index           =   2
         Left            =   2520
         Picture         =   "PrecisionVisionGraph.frx":F574
         Tag             =   "True"
         Top             =   15
         Width           =   240
      End
      Begin VB.Label Label_GraphName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ovality"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   0
         MouseIcon       =   "PrecisionVisionGraph.frx":F8FE
         TabIndex        =   51
         Top             =   0
         Width           =   2505
      End
      Begin VB.Label Label_GraphNameShadow 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ovality"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   2
         Left            =   20
         TabIndex        =   52
         Top             =   0
         Width           =   2505
      End
   End
   Begin VB.PictureBox GraphCaptionBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   7560
      MouseIcon       =   "PrecisionVisionGraph.frx":FE88
      MousePointer    =   99  'Custom
      ScaleHeight     =   375
      ScaleWidth      =   2895
      TabIndex        =   47
      Top             =   8400
      Width           =   2895
      Begin VB.Image ReportFlagImage 
         Height          =   240
         Index           =   1
         Left            =   2520
         Picture         =   "PrecisionVisionGraph.frx":10412
         Tag             =   "True"
         Top             =   15
         Width           =   240
      End
      Begin VB.Label Label_GraphName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Capacity"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   0
         MouseIcon       =   "PrecisionVisionGraph.frx":1079C
         MousePointer    =   99  'Custom
         TabIndex        =   48
         Top             =   0
         Width           =   2505
      End
      Begin VB.Label Label_GraphNameShadow 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Capacity"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   1
         Left            =   5
         TabIndex        =   49
         Top             =   0
         Width           =   2580
      End
   End
   Begin VB.Timer DisplayTimer 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   4200
      Tag             =   "Expand"
      Top             =   9480
   End
   Begin VB.PictureBox PVYScaleWindow 
      BackColor       =   &H00FF8080&
      Height          =   8055
      Left            =   0
      ScaleHeight     =   7995
      ScaleWidth      =   360
      TabIndex        =   0
      Top             =   0
      Width           =   425
      Begin VB.PictureBox PVYScaleImageBox 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C000C0&
         BorderStyle     =   0  'None
         Height          =   7935
         Left            =   -120
         ScaleHeight     =   7935
         ScaleWidth      =   975
         TabIndex        =   82
         Top             =   360
         Width           =   975
         Begin VB.PictureBox PVScaleMarkerFhAdjuster 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   105
            Index           =   0
            Left            =   95
            MouseIcon       =   "PrecisionVisionGraph.frx":10D26
            MousePointer    =   99  'Custom
            Picture         =   "PrecisionVisionGraph.frx":119F0
            ScaleHeight     =   105
            ScaleWidth      =   120
            TabIndex        =   84
            Top             =   1185
            Width           =   120
         End
         Begin VB.PictureBox PVScaleMarkerStAdjuster 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   105
            Index           =   0
            Left            =   95
            MouseIcon       =   "PrecisionVisionGraph.frx":11D7A
            MousePointer    =   99  'Custom
            Picture         =   "PrecisionVisionGraph.frx":12A44
            ScaleHeight     =   105
            ScaleWidth      =   120
            TabIndex        =   83
            Top             =   600
            Width           =   120
         End
         Begin VB.Line YScaleMarkerFh 
            BorderColor     =   &H00FF8080&
            Index           =   0
            X1              =   360
            X2              =   0
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Line YScaleMarkerSt 
            BorderColor     =   &H00FF8080&
            Index           =   0
            X1              =   360
            X2              =   0
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Image PVYScaleImage 
            Appearance      =   0  'Flat
            Height          =   7800
            Left            =   0
            Picture         =   "PrecisionVisionGraph.frx":12DCE
            Top             =   0
            Width           =   390
         End
      End
      Begin VB.Label Y_Units 
         BackColor       =   &H00FF8080&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   60
         TabIndex        =   116
         Top             =   60
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image Y_Units_Frame 
         Height          =   240
         Left            =   60
         Picture         =   "PrecisionVisionGraph.frx":12F71
         Top             =   60
         Width           =   240
      End
      Begin VB.Line PVYScaleZeroMarker 
         BorderColor     =   &H00C0C0C0&
         Index           =   0
         X1              =   360
         X2              =   0
         Y1              =   360
         Y2              =   360
      End
   End
   Begin VB.PictureBox PVGraphWindow 
      BackColor       =   &H00C0C0FF&
      Height          =   8340
      Left            =   420
      ScaleHeight     =   8280
      ScaleWidth      =   3045
      TabIndex        =   3
      Top             =   360
      Width           =   3100
      Begin VB.PictureBox PVXScaleBox 
         Height          =   315
         Index           =   0
         Left            =   -120
         ScaleHeight     =   255
         ScaleWidth      =   3045
         TabIndex        =   34
         Top             =   360
         Width           =   3100
         Begin VB.PictureBox PVXScale 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   0
            Left            =   0
            MouseIcon       =   "PrecisionVisionGraph.frx":132FB
            MousePointer    =   99  'Custom
            ScaleHeight     =   315
            ScaleWidth      =   3075
            TabIndex        =   35
            Top             =   0
            Width           =   3070
            Begin VB.Image XScaleLabel 
               Height          =   360
               Index           =   0
               Left            =   2760
               Picture         =   "PrecisionVisionGraph.frx":13FC5
               Top             =   -40
               Width           =   360
            End
            Begin VB.Image PVScaleClock 
               Height          =   240
               Index           =   0
               Left            =   0
               Picture         =   "PrecisionVisionGraph.frx":1472F
               Top             =   0
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.Label PVXScaleLabelC 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "50"
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   0
               Left            =   1170
               TabIndex        =   39
               Top             =   0
               Visible         =   0   'False
               Width           =   700
            End
            Begin VB.Line PVXScaleZeroMarker 
               Index           =   0
               X1              =   1500
               X2              =   1500
               Y1              =   150
               Y2              =   450
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00C0C0C0&
               Index           =   0
               X1              =   1500
               X2              =   1500
               Y1              =   0
               Y2              =   100
            End
            Begin VB.Line PVXScale2 
               Index           =   0
               X1              =   750
               X2              =   750
               Y1              =   150
               Y2              =   300
            End
            Begin VB.Line PVXScale6 
               Index           =   0
               X1              =   2250
               X2              =   2250
               Y1              =   150
               Y2              =   300
            End
            Begin VB.Line PVXScale1 
               Index           =   0
               X1              =   375
               X2              =   375
               Y1              =   225
               Y2              =   300
            End
            Begin VB.Line PVXScale3 
               Index           =   0
               X1              =   1125
               X2              =   1125
               Y1              =   225
               Y2              =   300
            End
            Begin VB.Line PVXScale5 
               Index           =   0
               X1              =   1875
               X2              =   1875
               Y1              =   225
               Y2              =   300
            End
            Begin VB.Line PVXScale7 
               Index           =   0
               X1              =   2625
               X2              =   2625
               Y1              =   225
               Y2              =   300
            End
            Begin VB.Line PVXScale0 
               Index           =   0
               X1              =   0
               X2              =   0
               Y1              =   150
               Y2              =   300
            End
            Begin VB.Line PVXScale8 
               Index           =   0
               X1              =   2985
               X2              =   2985
               Y1              =   150
               Y2              =   300
            End
            Begin VB.Line XMarkerR 
               BorderColor     =   &H008080FF&
               Index           =   0
               X1              =   1950
               X2              =   1950
               Y1              =   0
               Y2              =   240
            End
            Begin VB.Line XMarkerL 
               BorderColor     =   &H008080FF&
               Index           =   0
               X1              =   1050
               X2              =   1050
               Y1              =   0
               Y2              =   240
            End
            Begin VB.Label PVXScaleLabelUnits 
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   0
               Left            =   0
               TabIndex        =   38
               Top             =   0
               Width           =   105
            End
            Begin VB.Label PVXScaleLabelL 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "-50"
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   0
               Left            =   360
               TabIndex        =   37
               Top             =   0
               Width           =   705
            End
            Begin VB.Label PVXScaleLabelR 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "50"
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   0
               Left            =   1920
               TabIndex        =   36
               Top             =   0
               Width           =   700
            End
         End
      End
      Begin VB.PictureBox PVXScaleBox 
         Height          =   315
         Index           =   2
         Left            =   -120
         ScaleHeight     =   255
         ScaleWidth      =   3045
         TabIndex        =   23
         Top             =   1320
         Width           =   3100
         Begin VB.PictureBox PVXScale 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   2
            Left            =   0
            MouseIcon       =   "PrecisionVisionGraph.frx":14AB9
            MousePointer    =   99  'Custom
            ScaleHeight     =   315
            ScaleWidth      =   3060
            TabIndex        =   24
            Top             =   0
            Width           =   3060
            Begin VB.Image XScaleLabel 
               Height          =   360
               Index           =   2
               Left            =   2760
               Picture         =   "PrecisionVisionGraph.frx":15783
               Top             =   -40
               Width           =   360
            End
            Begin VB.Label PVXScaleLabelR 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "50"
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   2
               Left            =   1920
               TabIndex        =   28
               Top             =   0
               Width           =   705
            End
            Begin VB.Label PVXScaleLabelL 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "-50"
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   2
               Left            =   315
               TabIndex        =   27
               Top             =   0
               Width           =   705
            End
            Begin VB.Label PVXScaleLabelUnits 
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   2
               Left            =   0
               TabIndex        =   26
               Top             =   0
               Width           =   105
            End
            Begin VB.Line XMarkerL 
               BorderColor     =   &H008080FF&
               Index           =   2
               X1              =   1050
               X2              =   1050
               Y1              =   0
               Y2              =   240
            End
            Begin VB.Line XMarkerR 
               BorderColor     =   &H008080FF&
               Index           =   2
               X1              =   1950
               X2              =   1950
               Y1              =   0
               Y2              =   240
            End
            Begin VB.Line PVXScale8 
               Index           =   2
               X1              =   2985
               X2              =   2985
               Y1              =   150
               Y2              =   300
            End
            Begin VB.Line PVXScale0 
               Index           =   2
               X1              =   0
               X2              =   0
               Y1              =   150
               Y2              =   300
            End
            Begin VB.Line PVXScale7 
               Index           =   2
               X1              =   2625
               X2              =   2625
               Y1              =   225
               Y2              =   300
            End
            Begin VB.Line PVXScale5 
               Index           =   2
               X1              =   1875
               X2              =   1875
               Y1              =   225
               Y2              =   300
            End
            Begin VB.Line PVXScale3 
               Index           =   2
               X1              =   1125
               X2              =   1125
               Y1              =   225
               Y2              =   300
            End
            Begin VB.Line PVXScale1 
               Index           =   2
               X1              =   375
               X2              =   375
               Y1              =   225
               Y2              =   300
            End
            Begin VB.Line PVXScale6 
               Index           =   2
               X1              =   2250
               X2              =   2250
               Y1              =   150
               Y2              =   300
            End
            Begin VB.Line PVXScale2 
               Index           =   2
               X1              =   750
               X2              =   750
               Y1              =   150
               Y2              =   300
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00C0C0C0&
               Index           =   2
               X1              =   1500
               X2              =   1500
               Y1              =   0
               Y2              =   100
            End
            Begin VB.Line PVXScaleZeroMarker 
               Index           =   2
               X1              =   1500
               X2              =   1500
               Y1              =   150
               Y2              =   450
            End
            Begin VB.Label PVXScaleLabelC 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "50"
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   2
               Left            =   1170
               TabIndex        =   25
               Top             =   0
               Visible         =   0   'False
               Width           =   705
            End
            Begin VB.Image PVScaleClock 
               Height          =   240
               Index           =   2
               Left            =   0
               Picture         =   "PrecisionVisionGraph.frx":15EED
               Top             =   0
               Visible         =   0   'False
               Width           =   240
            End
         End
      End
      Begin VB.PictureBox PVXScaleBox 
         Height          =   315
         Index           =   3
         Left            =   -120
         ScaleHeight     =   255
         ScaleWidth      =   3045
         TabIndex        =   17
         Top             =   1800
         Width           =   3100
         Begin VB.PictureBox PVXScale 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            CausesValidation=   0   'False
            Height          =   315
            Index           =   3
            Left            =   0
            MouseIcon       =   "PrecisionVisionGraph.frx":16277
            MousePointer    =   99  'Custom
            ScaleHeight     =   315
            ScaleWidth      =   3060
            TabIndex        =   18
            Top             =   0
            Width           =   3060
            Begin VB.Image XScaleLabel 
               Height          =   360
               Index           =   3
               Left            =   2760
               Picture         =   "PrecisionVisionGraph.frx":16F41
               Top             =   -40
               Width           =   360
            End
            Begin VB.Label PVXScaleLabelR 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "50"
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   3
               Left            =   1920
               TabIndex        =   22
               Top             =   0
               Width           =   705
            End
            Begin VB.Label PVXScaleLabelL 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "-50"
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   3
               Left            =   315
               TabIndex        =   21
               Top             =   0
               Width           =   705
            End
            Begin VB.Label PVXScaleLabelUnits 
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   3
               Left            =   0
               TabIndex        =   20
               Top             =   0
               Width           =   225
            End
            Begin VB.Line XMarkerL 
               BorderColor     =   &H008080FF&
               Index           =   3
               X1              =   1050
               X2              =   1050
               Y1              =   0
               Y2              =   240
            End
            Begin VB.Line XMarkerR 
               BorderColor     =   &H008080FF&
               Index           =   3
               X1              =   1950
               X2              =   1950
               Y1              =   0
               Y2              =   240
            End
            Begin VB.Line PVXScale8 
               Index           =   3
               X1              =   2985
               X2              =   2985
               Y1              =   150
               Y2              =   300
            End
            Begin VB.Line PVXScale0 
               Index           =   3
               X1              =   0
               X2              =   0
               Y1              =   150
               Y2              =   300
            End
            Begin VB.Line PVXScale7 
               Index           =   3
               X1              =   2625
               X2              =   2625
               Y1              =   225
               Y2              =   300
            End
            Begin VB.Line PVXScale5 
               Index           =   3
               X1              =   1875
               X2              =   1875
               Y1              =   225
               Y2              =   300
            End
            Begin VB.Line PVXScale3 
               Index           =   3
               X1              =   1125
               X2              =   1125
               Y1              =   225
               Y2              =   300
            End
            Begin VB.Line PVXScale1 
               Index           =   3
               X1              =   375
               X2              =   375
               Y1              =   225
               Y2              =   300
            End
            Begin VB.Line PVXScale6 
               Index           =   3
               X1              =   2250
               X2              =   2250
               Y1              =   150
               Y2              =   300
            End
            Begin VB.Line PVXScale2 
               Index           =   3
               X1              =   750
               X2              =   750
               Y1              =   150
               Y2              =   300
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00C0C0C0&
               Index           =   3
               X1              =   1500
               X2              =   1500
               Y1              =   0
               Y2              =   100
            End
            Begin VB.Line PVXScaleZeroMarker 
               Index           =   3
               X1              =   1500
               X2              =   1500
               Y1              =   150
               Y2              =   450
            End
            Begin VB.Label PVXScaleLabelC 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "50"
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   3
               Left            =   1170
               TabIndex        =   19
               Top             =   0
               Visible         =   0   'False
               Width           =   705
            End
            Begin VB.Image PVScaleClock 
               Height          =   240
               Index           =   3
               Left            =   0
               Picture         =   "PrecisionVisionGraph.frx":176AB
               Top             =   0
               Visible         =   0   'False
               Width           =   240
            End
         End
      End
      Begin VB.PictureBox PVXScaleBox 
         Height          =   315
         Index           =   4
         Left            =   -120
         ScaleHeight     =   255
         ScaleWidth      =   3045
         TabIndex        =   11
         Top             =   2280
         Width           =   3100
         Begin VB.PictureBox PVXScale 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            CausesValidation=   0   'False
            Height          =   315
            Index           =   4
            Left            =   0
            MouseIcon       =   "PrecisionVisionGraph.frx":17A35
            MousePointer    =   99  'Custom
            ScaleHeight     =   315
            ScaleWidth      =   3060
            TabIndex        =   12
            Top             =   0
            Width           =   3060
            Begin VB.Image XScaleLabel 
               Height          =   360
               Index           =   4
               Left            =   2760
               Picture         =   "PrecisionVisionGraph.frx":186FF
               Top             =   -40
               Width           =   360
            End
            Begin VB.Image PVScaleClock 
               Height          =   240
               Index           =   4
               Left            =   0
               Picture         =   "PrecisionVisionGraph.frx":18E69
               Top             =   0
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.Label PVXScaleLabelC 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "50"
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   4
               Left            =   1170
               TabIndex        =   16
               Top             =   0
               Visible         =   0   'False
               Width           =   705
            End
            Begin VB.Line PVXScaleZeroMarker 
               Index           =   4
               X1              =   1500
               X2              =   1500
               Y1              =   150
               Y2              =   450
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00C0C0C0&
               Index           =   4
               X1              =   1500
               X2              =   1500
               Y1              =   0
               Y2              =   100
            End
            Begin VB.Line PVXScale2 
               Index           =   4
               X1              =   750
               X2              =   750
               Y1              =   150
               Y2              =   300
            End
            Begin VB.Line PVXScale6 
               Index           =   4
               X1              =   2250
               X2              =   2250
               Y1              =   150
               Y2              =   300
            End
            Begin VB.Line PVXScale1 
               Index           =   4
               X1              =   375
               X2              =   375
               Y1              =   225
               Y2              =   300
            End
            Begin VB.Line PVXScale3 
               Index           =   4
               X1              =   1125
               X2              =   1125
               Y1              =   225
               Y2              =   300
            End
            Begin VB.Line PVXScale5 
               Index           =   4
               X1              =   1875
               X2              =   1875
               Y1              =   225
               Y2              =   300
            End
            Begin VB.Line PVXScale7 
               Index           =   4
               X1              =   2625
               X2              =   2625
               Y1              =   225
               Y2              =   300
            End
            Begin VB.Line PVXScale0 
               Index           =   4
               X1              =   0
               X2              =   0
               Y1              =   150
               Y2              =   300
            End
            Begin VB.Line PVXScale8 
               Index           =   4
               X1              =   2985
               X2              =   2985
               Y1              =   150
               Y2              =   300
            End
            Begin VB.Line XMarkerR 
               BorderColor     =   &H008080FF&
               Index           =   4
               X1              =   1950
               X2              =   1950
               Y1              =   0
               Y2              =   240
            End
            Begin VB.Line XMarkerL 
               BorderColor     =   &H008080FF&
               Index           =   4
               X1              =   1050
               X2              =   1050
               Y1              =   0
               Y2              =   240
            End
            Begin VB.Label PVXScaleLabelUnits 
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   4
               Left            =   0
               TabIndex        =   15
               Top             =   0
               Width           =   105
            End
            Begin VB.Label PVXScaleLabelL 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "-50"
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   4
               Left            =   315
               TabIndex        =   14
               Top             =   0
               Width           =   705
            End
            Begin VB.Label PVXScaleLabelR 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "50"
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   4
               Left            =   1920
               TabIndex        =   13
               Top             =   0
               Width           =   705
            End
         End
      End
      Begin VB.PictureBox PVXScaleBox 
         Height          =   315
         Index           =   1
         Left            =   -120
         ScaleHeight     =   255
         ScaleWidth      =   3045
         TabIndex        =   10
         Top             =   840
         Width           =   3100
         Begin VB.PictureBox PVXScale 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   1
            Left            =   0
            MouseIcon       =   "PrecisionVisionGraph.frx":191F3
            MousePointer    =   99  'Custom
            ScaleHeight     =   315
            ScaleWidth      =   3060
            TabIndex        =   29
            Top             =   0
            Width           =   3060
            Begin VB.Image XScaleLabel 
               Height          =   360
               Index           =   1
               Left            =   2760
               Picture         =   "PrecisionVisionGraph.frx":19EBD
               Top             =   -40
               Width           =   360
            End
            Begin VB.Image PVScaleClock 
               Height          =   240
               Index           =   1
               Left            =   0
               Picture         =   "PrecisionVisionGraph.frx":1A627
               Top             =   0
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.Label PVXScaleLabelC 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "50"
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   1
               Left            =   1200
               TabIndex        =   33
               Top             =   0
               Visible         =   0   'False
               Width           =   705
            End
            Begin VB.Line PVXScaleZeroMarker 
               Index           =   1
               X1              =   1500
               X2              =   1500
               Y1              =   150
               Y2              =   450
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00C0C0C0&
               Index           =   1
               X1              =   1500
               X2              =   1500
               Y1              =   0
               Y2              =   100
            End
            Begin VB.Line PVXScale2 
               Index           =   1
               X1              =   750
               X2              =   750
               Y1              =   150
               Y2              =   300
            End
            Begin VB.Line PVXScale6 
               Index           =   1
               X1              =   2250
               X2              =   2250
               Y1              =   150
               Y2              =   300
            End
            Begin VB.Line PVXScale1 
               Index           =   1
               X1              =   375
               X2              =   375
               Y1              =   225
               Y2              =   300
            End
            Begin VB.Line PVXScale3 
               Index           =   1
               X1              =   1125
               X2              =   1125
               Y1              =   225
               Y2              =   300
            End
            Begin VB.Line PVXScale5 
               Index           =   1
               X1              =   1875
               X2              =   1875
               Y1              =   225
               Y2              =   300
            End
            Begin VB.Line PVXScale7 
               Index           =   1
               X1              =   2625
               X2              =   2625
               Y1              =   225
               Y2              =   300
            End
            Begin VB.Line PVXScale0 
               Index           =   1
               X1              =   0
               X2              =   0
               Y1              =   150
               Y2              =   300
            End
            Begin VB.Line PVXScale8 
               Index           =   1
               X1              =   2985
               X2              =   2985
               Y1              =   150
               Y2              =   300
            End
            Begin VB.Line XMarkerR 
               BorderColor     =   &H008080FF&
               Index           =   1
               X1              =   1950
               X2              =   1950
               Y1              =   0
               Y2              =   240
            End
            Begin VB.Line XMarkerL 
               BorderColor     =   &H008080FF&
               Index           =   1
               X1              =   1050
               X2              =   1050
               Y1              =   0
               Y2              =   240
            End
            Begin VB.Label PVXScaleLabelUnits 
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   1
               Left            =   0
               TabIndex        =   32
               Top             =   0
               Width           =   105
            End
            Begin VB.Label PVXScaleLabelL 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "-50"
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   1
               Left            =   360
               TabIndex        =   31
               Top             =   0
               Width           =   705
            End
            Begin VB.Label PVXScaleLabelR 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "50"
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   1
               Left            =   1920
               TabIndex        =   30
               Top             =   0
               Width           =   705
            End
         End
      End
      Begin VB.PictureBox PVXScaleBox 
         Height          =   315
         Index           =   5
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   3045
         TabIndex        =   4
         Top             =   2760
         Width           =   3100
         Begin VB.PictureBox PVXScale 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            CausesValidation=   0   'False
            Height          =   315
            Index           =   5
            Left            =   0
            MouseIcon       =   "PrecisionVisionGraph.frx":1A9B1
            MousePointer    =   99  'Custom
            ScaleHeight     =   315
            ScaleWidth      =   3060
            TabIndex        =   5
            Top             =   0
            Width           =   3060
            Begin VB.Label PVXScaleLabelR 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "50"
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   5
               Left            =   1920
               TabIndex        =   9
               Top             =   0
               Width           =   705
            End
            Begin VB.Label PVXScaleLabelL 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "-50"
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   5
               Left            =   315
               TabIndex        =   8
               Top             =   0
               Width           =   705
            End
            Begin VB.Label PVXScaleLabelUnits 
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   5
               Left            =   0
               TabIndex        =   7
               Top             =   0
               Width           =   105
            End
            Begin VB.Line XMarkerL 
               BorderColor     =   &H008080FF&
               Index           =   5
               X1              =   1050
               X2              =   1050
               Y1              =   0
               Y2              =   240
            End
            Begin VB.Line XMarkerR 
               BorderColor     =   &H008080FF&
               Index           =   5
               X1              =   1950
               X2              =   1950
               Y1              =   0
               Y2              =   240
            End
            Begin VB.Line PVXScale8 
               Index           =   5
               X1              =   2985
               X2              =   2985
               Y1              =   120
               Y2              =   270
            End
            Begin VB.Line PVXScale0 
               Index           =   5
               X1              =   0
               X2              =   0
               Y1              =   150
               Y2              =   300
            End
            Begin VB.Line PVXScale7 
               Index           =   5
               X1              =   2625
               X2              =   2625
               Y1              =   225
               Y2              =   300
            End
            Begin VB.Line PVXScale5 
               Index           =   5
               X1              =   1875
               X2              =   1875
               Y1              =   225
               Y2              =   300
            End
            Begin VB.Line PVXScale3 
               Index           =   5
               X1              =   1125
               X2              =   1125
               Y1              =   225
               Y2              =   300
            End
            Begin VB.Line PVXScale1 
               Index           =   5
               X1              =   375
               X2              =   375
               Y1              =   225
               Y2              =   300
            End
            Begin VB.Line PVXScale6 
               Index           =   5
               X1              =   2250
               X2              =   2250
               Y1              =   150
               Y2              =   300
            End
            Begin VB.Line PVXScale2 
               Index           =   5
               X1              =   750
               X2              =   750
               Y1              =   150
               Y2              =   300
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00C0C0C0&
               Index           =   5
               X1              =   1500
               X2              =   1500
               Y1              =   0
               Y2              =   100
            End
            Begin VB.Line PVXScaleZeroMarker 
               Index           =   5
               X1              =   1500
               X2              =   1500
               Y1              =   150
               Y2              =   450
            End
            Begin VB.Label PVXScaleLabelC 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "50"
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   5
               Left            =   1170
               TabIndex        =   6
               Top             =   0
               Visible         =   0   'False
               Width           =   705
            End
            Begin VB.Image PVScaleClock 
               Height          =   240
               Index           =   5
               Left            =   0
               Picture         =   "PrecisionVisionGraph.frx":1B67B
               Top             =   0
               Visible         =   0   'False
               Width           =   240
            End
         End
      End
      Begin VB.Image StartupBackgroundImage 
         Height          =   7455
         Left            =   120
         Picture         =   "PrecisionVisionGraph.frx":1BA05
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2775
      End
      Begin VB.Image DistanceImage 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   0
         Left            =   2640
         Picture         =   "PrecisionVisionGraph.frx":1DB9F
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line YMarkerFh 
         BorderColor     =   &H00800000&
         Index           =   0
         X1              =   0
         X2              =   15000
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line YMarkerSt 
         BorderColor     =   &H00800000&
         Index           =   0
         X1              =   0
         X2              =   15000
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Image PVGraphImage 
         BorderStyle     =   1  'Fixed Single
         Height          =   855
         Index           =   5
         Left            =   600
         MouseIcon       =   "PrecisionVisionGraph.frx":1DE44
         MousePointer    =   99  'Custom
         Picture         =   "PrecisionVisionGraph.frx":1EB0E
         Stretch         =   -1  'True
         Top             =   5160
         Width           =   2655
      End
      Begin VB.Image PVGraphImage 
         BorderStyle     =   1  'Fixed Single
         Height          =   855
         Index           =   4
         Left            =   480
         MouseIcon       =   "PrecisionVisionGraph.frx":1EEB9
         MousePointer    =   99  'Custom
         Picture         =   "PrecisionVisionGraph.frx":1FB83
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   2655
      End
      Begin VB.Image PVGraphImage 
         BorderStyle     =   1  'Fixed Single
         Height          =   855
         Index           =   3
         Left            =   360
         MouseIcon       =   "PrecisionVisionGraph.frx":1FF2E
         MousePointer    =   99  'Custom
         Picture         =   "PrecisionVisionGraph.frx":20BF8
         Stretch         =   -1  'True
         Top             =   4920
         Width           =   2655
      End
      Begin VB.Image PVGraphImage 
         BorderStyle     =   1  'Fixed Single
         Height          =   855
         Index           =   2
         Left            =   240
         MouseIcon       =   "PrecisionVisionGraph.frx":20FA3
         MousePointer    =   99  'Custom
         Picture         =   "PrecisionVisionGraph.frx":21C6D
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   2655
      End
      Begin VB.Image PVGraphImage 
         BorderStyle     =   1  'Fixed Single
         Height          =   855
         Index           =   1
         Left            =   120
         MouseIcon       =   "PrecisionVisionGraph.frx":22018
         MousePointer    =   99  'Custom
         Picture         =   "PrecisionVisionGraph.frx":22CE2
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   2655
      End
      Begin VB.Image PVGraphImage 
         Appearance      =   0  'Flat
         Height          =   960
         Index           =   0
         Left            =   0
         MouseIcon       =   "PrecisionVisionGraph.frx":2308D
         MousePointer    =   99  'Custom
         Picture         =   "PrecisionVisionGraph.frx":23D57
         Stretch         =   -1  'True
         Top             =   4440
         Width           =   2700
      End
   End
   Begin VB.PictureBox PipeDisplay 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8115
      Left            =   3600
      ScaleHeight     =   8115
      ScaleWidth      =   4740
      TabIndex        =   69
      Top             =   480
      Width           =   4740
      Begin VB.PictureBox ReportsPictureStorage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   3240
         ScaleHeight     =   495
         ScaleWidth      =   615
         TabIndex        =   115
         Top             =   480
         Width           =   615
      End
      Begin VB.PictureBox PVObsEye 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   400
         ScaleHeight     =   345
         ScaleWidth      =   405
         TabIndex        =   85
         Top             =   600
         Visible         =   0   'False
         Width           =   400
         Begin VB.Image PVObsEyeImage 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   0
            Picture         =   "PrecisionVisionGraph.frx":24133
            Top             =   0
            Width           =   435
         End
      End
      Begin VB.PictureBox MH_Shadow1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   55
         Picture         =   "PrecisionVisionGraph.frx":2464B
         ScaleHeight     =   300
         ScaleWidth      =   315
         TabIndex        =   71
         Top             =   40
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox MH_Shadow2 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   55
         Picture         =   "PrecisionVisionGraph.frx":24B20
         ScaleHeight     =   405
         ScaleWidth      =   315
         TabIndex        =   70
         Top             =   7290
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Shape ViewIndicator 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         Height          =   1515
         Index           =   0
         Left            =   60
         Top             =   350
         Width           =   75
      End
      Begin VB.Shape PipeMain 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         Height          =   7030
         Left            =   120
         Top             =   300
         Width           =   195
      End
      Begin VB.Image SpeakObs 
         Height          =   480
         Left            =   3600
         Picture         =   "PrecisionVisionGraph.frx":25008
         Top             =   1800
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Shape ViewIndicator 
         BackColor       =   &H0080C0FF&
         BackStyle       =   1  'Opaque
         Height          =   1515
         Index           =   1
         Left            =   120
         Top             =   350
         Width           =   50
      End
      Begin VB.Shape ViewIndicator 
         BackColor       =   &H0080C0FF&
         BackStyle       =   1  'Opaque
         Height          =   1515
         Index           =   2
         Left            =   150
         Top             =   345
         Width           =   135
      End
      Begin VB.Shape ViewIndicator 
         BackColor       =   &H0080C0FF&
         BackStyle       =   1  'Opaque
         Height          =   1515
         Index           =   4
         Left            =   270
         Top             =   350
         Width           =   45
      End
      Begin VB.Shape ViewIndicator 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         Height          =   1515
         Index           =   5
         Left            =   300
         Top             =   350
         Width           =   75
      End
      Begin VB.Shape PipeStartMH 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   1200
         Shape           =   3  'Circle
         Top             =   2040
         Width           =   315
      End
      Begin VB.Shape PipeFinishMH 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   60
         Shape           =   3  'Circle
         Top             =   7260
         Width           =   315
      End
   End
   Begin VB.PictureBox PVGraphForReports 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Index           =   0
      Left            =   6480
      ScaleHeight     =   3495
      ScaleWidth      =   15000
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   15000
   End
   Begin VB.PictureBox PVGraphForReports 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Index           =   1
      Left            =   6480
      ScaleHeight     =   3495
      ScaleWidth      =   15000
      TabIndex        =   2
      Top             =   4500
      Visible         =   0   'False
      Width           =   15000
   End
   Begin VB.PictureBox GraphCaptionBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   350
      Index           =   0
      Left            =   480
      MouseIcon       =   "PrecisionVisionGraph.frx":25CD2
      MousePointer    =   99  'Custom
      ScaleHeight     =   345
      ScaleWidth      =   2895
      TabIndex        =   44
      Top             =   0
      Width           =   2895
      Begin VB.Label FlatTM2 
         BackStyle       =   0  'Transparent
         Caption         =   "TM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   114
         Top             =   45
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image ReportFlagImage 
         Height          =   240
         Index           =   0
         Left            =   2520
         Picture         =   "PrecisionVisionGraph.frx":2625C
         Tag             =   "True"
         Top             =   15
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label_GraphName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Capacity"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   0
         MouseIcon       =   "PrecisionVisionGraph.frx":265E6
         MousePointer    =   99  'Custom
         TabIndex        =   45
         Top             =   0
         Width           =   2505
      End
      Begin VB.Label Label_GraphNameShadow 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Capacity"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   0
         Left            =   5
         TabIndex        =   46
         Top             =   0
         Width           =   2580
      End
   End
   Begin VB.Image ScaleSetting 
      Height          =   480
      Left            =   15
      Picture         =   "PrecisionVisionGraph.frx":26B70
      Top             =   8160
      Width           =   480
   End
   Begin VB.Image PrinterReportImage 
      Height          =   2700
      Left            =   1440
      Picture         =   "PrecisionVisionGraph.frx":2743A
      Top             =   10680
      Width           =   46080
   End
   Begin VB.Image PVGraphExpandBtn 
      Height          =   300
      Left            =   5760
      Picture         =   "PrecisionVisionGraph.frx":29AB9
      Tag             =   "0"
      Top             =   0
      Width           =   465
   End
   Begin VB.Shape ViewIndicator 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   1515
      Index           =   3
      Left            =   3840
      Top             =   840
      Width           =   75
   End
End
Attribute VB_Name = "PrecisionVisionGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim PVYScaleStLastY As Long 'Last mouse Y position - Start marker
Dim PVYScaleFhLastY As Long 'Last mouse Y position - Finish Marker
Dim PVYScaleMoveDir As String 'Indication of mouse move direction
Public ReportsTitleStorage As String
Dim UnicodeObsEntry As String

''vvvv PCN3490 *********************************************************
'Dim ObsIndexArray() As Integer 'Used to store the Observation index at or near the current mouse position.
'Const ObsMaxNumberInBubble As Integer = 5 'Maximum number of Observations for the ObsBubble.
'Dim AddNewObsFlag As Boolean
''^^^^ *****************************************************************

'Dim PVXScaleLimitPerL As Single 'PCNGL2901032 'PCN2337 Made Global for 3D Pipe Colour Limits
'Dim PVXScaleLimitPerR As Single 'PCNGL2901032 'PCN2337 Made Global for 3D Pipe Colour Limits





Function CreatePipeDisplay(Start_X, Start_Offset, End_X)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'CreatePipeDisplay Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    4/11/02     Building function
'
'Description:
'   The Precision Vision Graphs form contains a Pipe Display that shows where
'   along the pipeline the current graph represents. CreatePipeDisplay provides
'   the display with a 3D effect.
'
'Purpose:
'   Provide 3D effect on the Pipe Display.
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim Y_Pos As Integer  'PCN1916
Dim Y_Co As Integer
Dim Y2_Co As Integer
Dim ColourSteps As Integer
Dim PipeMainDrawStep As Integer
Dim X_Co As Integer
Dim X_Step As Integer

Me.MH_Shadow1.Visible = True
Me.MH_Shadow2.Visible = True

PipeDisplay.AutoRedraw = True
PipeDisplay.DrawWidth = 1
ColourSteps = (Me.PipeMain.width / 2) 'Removed for Version 6 (ML040806)

PipeMainDrawStep = 2

X_Co = Start_X + Start_Offset
X_Step = End_X - (X_Co + Start_Offset)

For Y_Pos = 0 To ColourSteps
    Me.PipeDisplay.Line (Me.PipeMain.Left + Y_Pos, X_Co)-(Me.PipeMain.Left + Y_Pos, X_Co + X_Step), RGB(64 + Y_Pos * PipeMainDrawStep, 64 + Y_Pos * PipeMainDrawStep, 64)
    Me.PipeDisplay.Line (Me.PipeMain.Left + Me.PipeMain.width - (10 + Y_Pos), X_Co)-(Me.PipeMain.Left + Me.PipeMain.width - (10 + Y_Pos), X_Co + X_Step), RGB(64 + Y_Pos * PipeMainDrawStep, 64 + Y_Pos * PipeMainDrawStep, 64)
    '************************
    ' Create the pipe shadow
    ' Me.PipeDisplay.BackColor = 12632256 (RGB(198,198,198))
    ' ShadowColour = RGB(118,118,118)
'        If (Y_Pos / 2 < 198) And (Y_Pos < 2 * ColourSteps / 3) Then
'            Y_Co = Me.PipeMain.Left + Me.PipeMain.width + Y_Pos
'            Me.PipeDisplay.Line (Y_Co, X_Co)-(Y_Co, X_Co + X_Step), RGB(118 + Y_Pos / 2, 118 + Y_Pos / 2, 118 + Y_Pos / 2)
'        End If
'        If (118 + (Y_Pos + ColourSteps) / 2) < 198 Then
'            Y2_Co = Me.PipeMain.Left + Me.PipeMain.width + Y_Pos + 2 * ColourSteps / 3
'            Me.PipeDisplay.Line (Y2_Co, X_Co)-(Y2_Co, X_Co + X_Step), RGB(118 + (Y_Pos + ColourSteps) / 2, 118 + (Y_Pos + ColourSteps) / 2, 118 + (Y_Pos + ColourSteps) / 2)
'        End If
    '************************
Next Y_Pos
PipeDisplay.AutoRedraw = False
    
'Call DrawPVGraphLimitLines 'PCNGL050203 PCN3402 removed

Exit Function
Err_Handler:
    MsgBox Err & "-PVG1:" & Error$, vbExclamation
End Function



Private Sub AdjustRefShapeCentre_Click()
On Error GoTo Err_Handler
    If AdjustRefShapeCentre.value = 1 Then
        Call Observations.AddRefShapeCentreChange
    Else
        Call Observations.DelRefShapeCentreChange

    End If
    
    Exit Sub
Err_Handler:
    MsgBox Err & "-PVG2:" & Error$
End Sub

Private Sub AdjustRefShapeCentre_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler



Exit Sub
Err_Handler:
    MsgBox Err & "-PVG3:" & Error$
End Sub

Private Sub AdjustWaterCentre_Click()
On Error GoTo Err_Handler

    If AdjustWaterCentre.value = 1 Then
        Call Observations.AddWaterLevelChange
    Else
        Call Observations.DelWaterLevelChange
    End If

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG4:" & Error$
End Sub

Private Sub DisplayTimer_Timer()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PVGraphResetForm
'Created : 20 Feb 2005, PCN3402
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Expands the PVGraph form to show the PVGraphs
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim FormWidthIncrement As Long

FormWidthIncrement = 1090

If DisplayTimer.Tag = "Expand" Then
    If PrecisionVisionGraph.Left >= FormWidthIncrement Then
        PVGraphExpandBtn.Visible = False
        PipeDisplay.Visible = False
        DoEvents
        PrecisionVisionGraph.width = PrecisionVisionGraph.width + FormWidthIncrement
        PrecisionVisionGraph.Left = PrecisionVisionGraph.Left - FormWidthIncrement
        PVGraphExpandBtn.Left = PVGraphExpandBtn.Left + FormWidthIncrement
        PipeDisplay.Left = PipeDisplay.Left + FormWidthIncrement
        Observation.Left = Observation.Left + FormWidthIncrement
        PVGraphWindow.width = PVGraphWindow.width + FormWidthIncrement
        PVGraphExpandBtn.Visible = True
        PipeDisplay.Visible = True
    Else
        PVGraphExpandBtn.Picture = LoadResPicture(106, vbResBitmap)
        DisplayTimer.Enabled = False
        DisplayTimer.Tag = "Contract"
    End If
Else
    If PrecisionVisionGraph.Left <= PipelineDetails.Left - FormWidthIncrement Then
        PVGraphExpandBtn.Visible = False
        PipeDisplay.Visible = False
        DoEvents
        PrecisionVisionGraph.width = PrecisionVisionGraph.width - FormWidthIncrement
        PrecisionVisionGraph.Left = PrecisionVisionGraph.Left + FormWidthIncrement
        PVGraphExpandBtn.Left = PVGraphExpandBtn.Left - FormWidthIncrement
        PipeDisplay.Left = PipeDisplay.Left - FormWidthIncrement
        Observation.Left = Observation.Left - FormWidthIncrement
        PVGraphWindow.width = PVGraphWindow.width - FormWidthIncrement
        PVGraphExpandBtn.Visible = True
        PipeDisplay.Visible = True
    Else
        PVGraphExpandBtn.Picture = LoadResPicture(108, vbResBitmap)
        DisplayTimer.Enabled = False
        PrecisionVisionGraph.Left = PipelineDetails.Left
        PrecisionVisionGraph.width = PipelineDetails.width
        PVGraphExpandBtn.Left = 3600
        PipeDisplay.Left = 3555
        DisplayTimer.Tag = "Expand"
        
        Call ShowOrHideMultipleGraphs("Hide")
        Call ScreenDrawing.DrawPVGraphs
        
        
    End If
End If


Exit Sub
Err_Handler:
    MsgBox Err & "-PVG5:" & Error$
    PVGraphExpandBtn.Visible = True
    PipeDisplay.Visible = True

End Sub

Private Sub Form_Activate()
On Error GoTo Err_Handler
Dim Start_X As Integer
Dim Start_Offset As Integer
Dim End_X As Integer



'Create pipe display
Start_Offset = 50
Start_X = PipeMain.Top
End_X = PipeMain.Top + PipeMain.height
Call CreatePipeDisplay(Start_X, Start_Offset, End_X)



'Set background of manhole graphics
MH_Shadow1.BackColor = RGB(184, 205, 236)
MH_Shadow2.BackColor = RGB(184, 205, 236)

'vvvv PCN4171 ************************************
'Set background of PVYScaleWindow and Y_Units
PVYScaleWindow.BackColor = RGB(145, 177, 219)
Y_Units.BackColor = RGB(145, 177, 219)
'^^^^ ********************************************

WaterShiftControls.BackColor = &HC0FFFF
ReferenceShiftControls.BackColor = &HC0FFFF


Call Observations.ObsDisplayALL 'Display all of the observations 'PCN3490 'PCN4131


'Bring the View Indicator to the front
Me.ViewIndicator.Item(0).ZOrder 0
Me.ViewIndicator.Item(1).ZOrder 0
Me.ViewIndicator.Item(2).ZOrder 0
'Me.ViewIndicator.Item(3).ZOrder 0
Me.ViewIndicator.Item(4).ZOrder 0
Me.ViewIndicator.Item(5).ZOrder 0

Me.SpeakObs.ZOrder 0
Me.PVObsEye.ZOrder 0

RulerGreen.Top = PVGraphWindow.Top
RulerRed.Top = PVGraphWindow.Top + PVGraphWindow.height - RulerRed.height - PVXScaleBox(0).height
RulerGreen.Left = PVGraphWindow.width + PVGraphWindow.Left
RulerRed.Left = PVGraphWindow.width + PVGraphWindow.Left

'Temporary solution - should store in the ini file 'PCNGL181202
If ImageGraphState(0).GraphType = "" Then 'PCN1942
    CapacityLimitL = -25
    CapacityLimitR = 25
    ImageGraphState(0).GraphType = "Capacity"

    Call RepositionPVXLimitMarkers   'PCNGL181202

    'Call DrawPVGraphLimitLines 'PCNGL280103 PCN3402 Removed
    OvalityLimitL = 0 'Not set
    OvalityLimitR = 6 'PCN4250
    DeltaLimitL = -25
    DeltaLimitR = 25
    Flat3dLimitL = -10 'PCNANTONY 'PCN4250
    Flat3dLimitR = 10 'PCNANTONY 'PCN4250
    
End If

Call SetupPVGraphScreen(ImageGraphState(0).GraphType) 'PCN1943

Call DrawPVGraphs

Call DrawPVXScale


'Call ClearLineScreen.GotoPVProfile(PVFrameNo, True)
If PVDFileName = "" And mediatype <> Video Then
    'StartupBackgroundImage.width = PVGraphImage(0).width
    'StartupBackgroundImage.height = PVGraphImage(0).height
    'StartupBackgroundImage.Top = PVGraphImage(0).Top
    'StartupBackgroundImage.Left = PVGraphImage(0).Left
Else
    StartupBackgroundImage.Visible = 0
End If
'    OptionsPage.SmoothCheck.Top = PVGraphWindow.Top + PVGraphWindow.height
'    OptionsPage.SmoothCheck.BackColor = RGB(183, 205, 236)
'    OptionsPage.SmoothCheck.Top = PipelineDetails.PipeDataFrame.Top + PipelineDetails.PipeDataFrame.height + 105
    
    PVObsEntry.Font.Charset = LanguageCharset
'Don't put this code here, causes problem with drawing PVYScale PCN1850 'Call DrawPVYScaleNumbersAndMarkers(PVYScale, 1, 0) 'PCNGL140103 'PCN1850

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG6:" & Error$

End Sub

Private Sub Form_Click()
On Error GoTo Err_Handler
Dim PVi As Long

'PrecisionVisionGraph.PVGraphScreen(4).AutoRedraw = True
'PrecisionVisionGraph.PVGraphScreen(4).Cls
'PrecisionVisionGraph.PVGraphScreen(4).AutoRedraw = False
'For PVi = 1 To PVDataNoOfLines
'    Call DrawPVGraphXYDiameter(PrecisionVisionGraph.PVGraphScreen(4), PVi, 0)
'Next PVi

'PCN3373 ''''''''''''''''''''''''''
Call DrawPVGraphs
'Call DrawPVGraphAllAtCurrentWindow
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG7:" & Error$

End Sub



Private Sub Form_GotFocus()
On Error GoTo Err_Handler



Exit Sub
Err_Handler:
    MsgBox Err & "-PVG8:" & Error$
End Sub


Private Sub Form_Load()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : Form_Load
'Created : 4 November 2002,
'Updated : 18 November 2003, PCN2402
'Prg By  : Geoff Logan
'Param   :
'Desc    : The Precision Vision Graphs form is for representing the collected profiler
'           data in a useful manner, such as Pipeline Capacity, Out-of-Circle, Flat-3D,
'           Delta and 3D.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim i As Integer
Dim GraphWidth As Integer
Dim GraphSpacing As Integer


ConvertLanguage Me, Language 'PCN2111



PrecisionVisionGraph.Left = PVPageLeft 'PCNGL030103
PrecisionVisionGraph.Top = PVPageTop
PrecisionVisionGraph.width = PVPageWidth
PrecisionVisionGraph.height = PVPageHeight
'vvvv PCN4171 **********************************
Me.BackColor = RGB(184, 205, 236)
For i = 0 To 5
    GraphCaptionBox(i).BackColor = RGB(184, 205, 236)
Next i
PipeDisplay.BackColor = RGB(184, 205, 236)
Me.GraphTypeSelect.BackColor = RGB(184, 205, 236)
'^^^^ ******************************************

 
If (IdentifyOperatingSystem = "Windows XP") Then 'ML120203
    ScaleSetting.Top = 8065 'PCN2402
    PVGraphWindow.height = 8095 'PCN2402 'Was 7800 PCN3373
    'PVYScaleWindow.height = 8140 'PCN2402
    PVYScaleWindow.height = 8095
    
    'PVXScale(0).Top = 8140 'PCN2402 'PCN3373
    
'    PrecisionVisionGraph.width = PVPageWidth - 150 'PCN2402
End If

PVGraphWindow.BackColor = &HFFFFFF
PVYScaleImageBox.BackColor = &HFFFFFF
'PCN3402 Set the graphs up in the PVGraph Screen'''''''''''''''''''''''''''''''''
'                                                                               '
GraphWidth = 2800                                                               '
GraphSpacing = 2800                                                             '
                                                                                '
For i = 0 To 5                                                                  '
    'PCN3373 Setting up the new image graph positions and load bitmaps          '
    PVGraphImage(i).Stretch = False
    PVGraphImage(i).Left = (i - 1) * GraphSpacing                               '
    PVGraphImage(i).Top = 0                                                     '

    GraphCaptionBox(i).Left = (i - 1) * GraphSpacing + PVGraphWindow.Left + 25
    GraphCaptionBox(i).width = GraphWidth
    GraphCaptionBox(i).Top = 0
    GraphCaptionBox(i).height = 300
    Label_GraphName(i).Top = 0
    Label_GraphName(i).Left = 0
    Label_GraphName(i).width = 2505
    Label_GraphNameShadow(i).Top = 0
    Label_GraphNameShadow(i).Left = 20
    Label_GraphNameShadow(i).width = 2505

    If i < 4 Then PVGraph3DBoarder(i).Left = (i * GraphSpacing) + GraphSpacing + PVGraphWindow.Left
Next i                                                                          '
                                                                                
PVScaleClock(5).Left = 100
PVYScaleImageBox.height = PVGraphWindow.height                                                    '
PVYScaleImageBox.Top = PVGraphWindow.Top
PVYScaleImageBox.width = PVYScaleWindow.width
PVYScaleImageBox.Left = 0


PVYScaleImage.Top = 0
PVYScaleImage.Left = 0
PVYScaleImage.width = PVYScaleImageBox.width
PVYScaleImage.height = PVGraphImage(0).height


'
'PVGraphImage(0).width = 3045
PVGraphImage(0).Left = 0  '

GraphCaptionBox(0).Left = PVGraphWindow.Left
GraphCaptionBox(0).width = PVGraphImage(0).width

'PVXScaleBox(0).width = 3045
PVXScaleBox(0).Left = 0                                                                                 '

Call ShowOrHideMultipleGraphs("Hide")                                           '
Call ToggleReportFlag(5)
                                                                                '

For i = 0 To 5
    PVXScaleBox(i).width = PVGraphImage(i).width + 50 'PCN4171
    PVXScaleBox(i).Left = PVGraphImage(i).Left
    PVXScaleBox(i).Top = ScaleSetting.Top - PVGraphWindow.Top - 30
    PVXScale(i).width = PVGraphImage(i).width - 72
    PVXScale(i).Left = 0
    If i > 0 And i < 5 Then
        PVXScaleLabelUnits(i).Left = 0
        XScaleLabel(i).Left = PVGraphImage(i).width - XScaleLabel(i).width
    End If
    PVXScaleLabelUnits(i).width = 210
Next i

PVXScaleBox(0).Left = PVGraphImage(0).Left - 25
PVXScale(0).width = PVGraphImage(0).width - 10
                                                                                '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'****************************************
' Initialise parameters
PipeDisplayMoveLastY = 0
YScaleZoomFactor = 1 'The default value 'PCNGL080103
'These PV ratios are in mm or inches per pixel
'vvvv PCN2829 ******************************
'PVGraphCapacityXRatio = 10 'PCN2829

'PVGraphDeltaXRatio = 10
If PVGraphGeneralXScale = 0 Then PVGraphGeneralXScale = 10
If PVGraphCapacityXScale = 0 Then PVGraphCapacityXScale = 10
If PVGraphOvalityXScale = 0 Then PVGraphOvalityXScale = 10
If PVGraphDeltaXScale = 0 Then PVGraphDeltaXScale = 10
If PVGraphXYDiaXScale = 0 Then PVGraphXYDiaXScale = 10 'PCN2829
If PVGraphDiaMaxMinXScale = 0 Then PVGraphDiaMaxMinXScale = 10
If PVGraphDiaMedianXScale = 0 Then PVGraphDiaMedianXScale = 10

If PVGraphDiaMaxXScale = 0 Then PVGraphDiaMaxXScale = 10


'^^^^ **************************************
'PVGraphYRatio = 50
If PVGraphYRatio < 2 Then 'PCN2121
    PVGraphYRatio = 10 'PCNGL170303 'PCN4258
End If
PVGraphYRatio = 1
Call InitialiseNumberPicArray
'Setup View Indicator
ViewIndicator(0).height = PipeMain.height - 80
ViewIndicator(1).height = PipeMain.height - 80
ViewIndicator(2).height = PipeMain.height - 80
ViewIndicator(3).height = PipeMain.height - 80
ViewIndicator(4).height = PipeMain.height - 80
ViewIndicator(5).height = PipeMain.height - 80
'vvvv PCN2970 **************************************************
' Position YScale makers
PrecisionVisionGraph.YScaleMarkerSt(0).y1 = PrecisionVisionGraph.YMarkerSt(0).y1
PrecisionVisionGraph.YScaleMarkerSt(0).y2 = PrecisionVisionGraph.YMarkerSt(0).y1
PrecisionVisionGraph.YScaleMarkerFh(0).y1 = PrecisionVisionGraph.YMarkerFh(0).y1
PrecisionVisionGraph.YScaleMarkerFh(0).y2 = PrecisionVisionGraph.YMarkerFh(0).y1
PrecisionVisionGraph.PVScaleMarkerStAdjuster(0).Top = PrecisionVisionGraph.YScaleMarkerSt(0).y1 - PrecisionVisionGraph.PVScaleMarkerStAdjuster(0).height
PrecisionVisionGraph.PVScaleMarkerFhAdjuster(0).Top = PrecisionVisionGraph.YScaleMarkerFh(0).y1 + 1
'^^^^ **********************************************************

'''' PCN3441 (6 April 2005, Antony van Iersel)
''
'''Background colour
''PVGraphScreen(0).BackColor = RGB(239, 247, 255)
''PVGraphScreen(1).BackColor = RGB(239, 247, 255)
''PVGraphScreen(2).BackColor = RGB(239, 247, 255)
''PVGraphScreen(3).BackColor = RGB(239, 247, 255)
''PVGraphScreen(4).BackColor = RGB(239, 247, 255)
'''vvvv PCN2970 *********************************************
''PVGraphScreen(5).BackColor = RGB(239, 247, 255)
''PVGraphScreen(6).BackColor = RGB(239, 247, 255)
''PVGraphScreen(7).BackColor = RGB(239, 247, 255)
''PVGraphScreen(8).BackColor = RGB(239, 247, 255)
''PVGraphScreen(9).BackColor = RGB(239, 247, 255)
'^^^^ *****************************************************
'Call DrawPVYScaleNumbersAndMarkers(PVYScale, 1, 0) 'PCN1850
'DoEvents 'PCN1850

'vvvv PCN3490 ***************************************************
'Setup the observation markers
Call Observations.ObsDisplayALL 'PCN4131
'Setup the position of the observation bubble
Observation.Left = 600
SpeakObs.Left = 0
'^^^^ ***********************************************************

Me.Y_Units.Visible = False 'PCN4205

Call SetLimitLines 'PCN2769

'Language setup for Graph Titles PCN3402 (Antony van Iersel, 14 April 2005)
'Capacity'
Call UpdateGraphTitles

Call PopulatePVGraphSelectCaptions 'PCN

'PCN5185 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Replace the populating the the graph captions with a function call
Call PopulatePVGraphSelectCaptions
'vvvv PCN4171 ***************************************************
'PVGraphSelect(0).Caption = GetContainerGraphLabel(0)
'PVGraphSelect(1).Caption = GetContainerGraphLabel(1)
'PVGraphSelect(2).Caption = GetContainerGraphLabel(2)
'PVGraphSelect(3).Caption = GetContainerGraphLabel(3)
'PVGraphSelect(4).Caption = GetContainerGraphLabel(4)
'PVGraphSelect(5).Caption = GetContainerGraphLabel(5)
'PVGraphSelect(6).Caption = GetContainerGraphLabel(8)
''PVGraphSelect(7).Caption = GetContainerGraphLabel(7) 'PCN4461


'^^^^ ***********************************************************



Exit Sub
Err_Handler:
    MsgBox Err & "-PVG9:" & Error$
End Sub

Sub PopulatePVGraphSelectCaptions()
On Error GoTo Err_Handler

PVGraphSelect(0).Caption = GetContainerGraphLabel(0)
PVGraphSelect(1).Caption = GetContainerGraphLabel(1)
PVGraphSelect(2).Caption = GetContainerGraphLabel(2)
PVGraphSelect(3).Caption = GetContainerGraphLabel(3)
PVGraphSelect(4).Caption = GetContainerGraphLabel(4)
PVGraphSelect(5).Caption = GetContainerGraphLabel(5)
PVGraphSelect(6).Caption = GetContainerGraphLabel(8)
'PCN6458 PVGraphSelect(7).Caption = GetContainerGraphLabel(PVInclination)

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG9:" & Error$
End Sub


Private Sub FormBackground_Click()

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ScaleSetting.Picture = Me.ButtonImages.ListImages("YScale").Picture
Me.GraphTypeSelect.Picture = Me.ButtonsSmallImages.ListImages("More").Picture
Me.XScaleLabel(0).Picture = Me.ButtonsSmallImages.ListImages("More").Picture
Me.XScaleLabel(1).Picture = Me.ButtonsSmallImages.ListImages("More").Picture
Me.XScaleLabel(2).Picture = Me.ButtonsSmallImages.ListImages("More").Picture
Me.XScaleLabel(3).Picture = Me.ButtonsSmallImages.ListImages("More").Picture
Me.XScaleLabel(4).Picture = Me.ButtonsSmallImages.ListImages("More").Picture

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG10:" & Error$
End Sub

Private Sub GraphCaptionBox_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    If Index <> PVGraphSelectBox.Tag Then Call CentralMoveRoutine
Exit Sub
Err_Handler:
    Select Case Err
        Case 13: Resume Next
        Case Else
            MsgBox Err & "-PVG11:" & Error$
    End Select
End Sub

Private Sub GraphTypeSelect_Click()
On Error GoTo Err_Handler

If PVGraphSelectBox.Visible = False Then
    Call ExpandPVGraphSelectBox(0)
Else
    Call CollapsePVGraphSelectBox
End If
    
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG12:" & Error$
End Sub

Private Sub GraphTypeSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Call MoreButtonHighlight(Me.GraphTypeSelect, X, Y)

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG13:" & Error$
End Sub

Private Sub HistogramOfDiameter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : MedianDiameterForInspection
'Created : 16 April 2005, PCN3489
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Calculates the median diameter for the entire inspection.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
'Dim HistDiamOriginalHeight As Long
'Dim HistDiamOriginalTop As Long
'Dim HistDiamExpandedHeight As Long
'Dim HistDiamExpandedTop As Long


If Y < HistogramOfDiameterGraph.height Then

'    HistDiamOriginalHeight = 1570
'    HistDiamOriginalTop = ScaleSetting.Top + ScaleSetting.height + 10
'    HistDiamExpandedHeight = 10000
'    HistDiamExpandedTop = 0
'
'    DiameterHistogramSubTitle.Visible = True
'
'    HistogramOfDiameter.Top = HistDiamOriginalTop
'    HistogramOfDiameter.height = HistDiamOriginalHeight
'    HistogramOfDiameter.width = 15000 - PVToolbar.width
'    HistogramOfDiameter.left = 10
'    HistogramOfDiameterGraph.Top = -525
'    HistogramOfDiameterGraph.left = -100

    Call HistogramDiameterGraphSetup("Normal")



End If


Exit Sub
Err_Handler:
   MsgBox Err & "-PVG14:" & Error$
End Sub


Private Sub HistogramOfDiameterGraph_DblClick()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : HistogramOfDiameterGraph_MouseMove
'Created : 16 April 2005, PCN3489
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
'Dim HistDiamOriginalHeight As Long
'Dim HistDiamOriginalTop As Long
'Dim HistDiamExpandedHeight As Long
'Dim HistDiamExpandedTop As Long
'
''HistDiamOriginalHeight = 1600
''HistDiamOriginalTop = 8580
'HistDiamExpandedHeight = 10000
'HistDiamExpandedTop = 0
'
'
'If HistogramOfDiameter.Top <> HistDiamExpandedTop Then
'    HistogramOfDiameter.Top = HistDiamExpandedTop
'    HistogramOfDiameter.height = HistDiamExpandedHeight
'    HistogramOfDiameter.left = 0
'    HistogramOfDiameter.width = 15050
'    HistogramOfDiameterGraph.Top = 10
'    HistogramOfDiameterGraph.left = 10
'    DiameterHistogramSubTitle.Visible = False
'End If


Call HistogramDiameterGraphSetup("Expanded")


Exit Sub
Err_Handler:
    MsgBox Err & "-PVG15:" & Error$
End Sub

Private Sub DeleteObs_Click()
On Error GoTo Err_Handler

Dim Start_X As Integer
Dim Start_Offset As Integer
Dim End_X As Integer

    If UBound(ObsIndexArray) = 1 Then
        Call Observations.ObsDelete(ObsIndexArray(1))
        Call ObservationClose_Click
    End If
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG16:" & Error$
End Sub



Private Sub LargeGreenRuler_Click()
On Error GoTo Err_Handler
    If ConfigInfo.DistanceStart <> InvalidData Then Call ClearLineScreen.GotoStartMarker: PVFrameNo = 1
    If SoftwareConfiguration = "Reader" Then Exit Sub 'PCN4322
        
    Load DataEntryForm
    Call DataEntryForm.SetDataEntryType("DistanceStartChange")
    DataEntryForm.Left = PrecisionVisionGraph.Left + 300
    DataEntryForm.Top = PrecisionVisionGraph.Top + RulerGreen.Top + 600
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG17:" & Error$
End Sub

Private Sub LargeRedRuler_Click()
On Error GoTo Err_Handler
    If ConfigInfo.DistanceFinish <> InvalidData Then Call ClearLineScreen.GotoStopMarker: PVFrameNo = PVDataNoOfLines
    If SoftwareConfiguration = "Reader" Then Exit Sub 'PCN4322
    
    Load DataEntryForm
    Call DataEntryForm.SetDataEntryType("DistanceEndChange")
    DataEntryForm.Left = PrecisionVisionGraph.Left + 300
    DataEntryForm.Top = PrecisionVisionGraph.Top + RulerGreen.Top + PVGraphWindow.height - 1300
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG18:" & Error$
End Sub

Private Sub LargeRulerImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    If SoftwareConfiguration = "Reader" Then Exit Sub 'PCN4322
    ReDim Observations.ObsIndexArray(0) As Integer
    
    Y = Y + LargeRuler.Top - PipeDisplay.Top
    Observations.LastYObsPositionClicked = Y
    Observations.LastObservationType = "Distance"
    If Observations.ObsMouseOverMe(Y) <> 0 Then 'PCN4131
        Call Observations.ObsBubbleWithEyeSetup  'PCN4131
        Me.ProfileReportObsButtonImage.Visible = False 'PCN4366
    End If
Exit Sub
Err_Handler:
MsgBox Err & "-PVG19:" & Error$
End Sub





Private Sub ObservationText_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'****************************************************************************************
'Name    : ObservationText_Click
'Created : 22 March 2005, PCN3490
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Edit observation.
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
Dim ObsFrameNo As Long
Dim ObsFrameStr As String
Dim ObsDistStr As String
Dim FramePos As Integer
Dim PipeObsIndexNo As Integer

If SoftwareConfiguration = "Reader" Then Exit Sub 'PCN4325

PipeObsIndexNo = Observations.ObsIndexArray(Index)

ObsFrameNo = PipeObservations(PipeObsIndexNo).PipeObsFrameNo
If ObsFrameNo <= 0 Then Exit Sub

If Trim(PipeObservations(ObsIndexArray(Index)).PipeObs) = "<<<--I-->>>" Then
    Call ClearLineScreen.GotoPVProfile(PipeObservations(ObsIndexArray(Index)).PipeObsFrameNo, True)
    Load DataEntryForm
    Call DataEntryForm.SetDataEntryType("DistanceChange")
    DataEntryForm.Left = Observation.Left - DataEntryForm.width + PrecisionVisionGraph.Left
    DataEntryForm.Top = Observation.Top - 300
    Exit Sub
End If

If UBound(ObsIndexArray) > 1 Then
    Call Observations.ObsEye_Click(Index)
Else
    If Button = 1 Then Call Observations.ObsAddNewOrUpdateSetup(ObsFrameNo, Observations.ObsIndexArray(1)) 'PCN4131
End If



InvalidData:
    'Do nothing
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG20:" & Error$
End Sub

Private Sub ObsPreview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    If Me.PaperClipHighLight.Visible Then Me.PaperClipHighLight.Visible = False
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG21:" & Error$
End Sub

Private Sub ObsPreviewPictureBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    If Me.PaperClipHighLight.Visible Then Me.PaperClipHighLight.Visible = False
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG22:" & Error$
End Sub

Private Sub PaperClipAttachment_Click()
On Error GoTo Err_Handler
    If SoftwareConfiguration = "Reader" Then Exit Sub 'PCN4322
    Call Observations.AttachObsPicture
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG23:" & Error$
End Sub

Private Sub Label_GraphName_Click(Index As Integer)
On Error GoTo Err_Handler
    If PVGraphSelectBox.Visible = False Then
        Call ExpandPVGraphSelectBox(Index)
    Else
        Call CollapsePVGraphSelectBox
    End If
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG24:" & Error$
End Sub

Private Sub Label_GraphName_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    Dim i As Integer
    
    If Index <> PVGraphSelectBox.Tag Then Call CentralMoveRoutine
    If PVGraphSelectBox.Visible = True Then
        For i = 0 To 7 'PCN6128 from 6 to 7
            If PVGraphSelect(i).Caption = Label_GraphName(Index).Caption Then
                PVGraphSelect(i).BackColor = &H80FFFF
            Else
                PVGraphSelect(i).BackColor = &H80C0FF
            End If
        Next i
    End If
Exit Sub
Err_Handler:
    Select Case Err
        Case 13: Resume Next
        Case Else
            MsgBox Err & "-PVG25:" & Error$
    End Select
End Sub

Private Sub MH_Shadow1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'PCNGL181202
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'MH_Shadow1_MouseDown Sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    17/11/02     Building initial framework
'
'Description:
'
'Purpose:
'   Initialises position of mouse
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

PipeDisplayMoveLastY = ViewIndicator.Item(0).Top
Select Case Button
    Case 1 ' Left Button Press
        Call PipeDisplay_MouseMove(1, 0, ViewIndicator.Item(0).Top - 10, ViewIndicator.Item(0).Top - 25) 'Move viewindicator up
End Select

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG26:" & Error$
End Sub

Private Sub MH_Shadow1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'PCNGL181202
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'MH_Shadow1_MouseMove Sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    17/11/02     Building initial framework
'
'Description:
'
'Purpose:
'   Initialises position of mouse
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Select Case Button
    Case 1 ' Left Button Press
        Call PipeDisplay_MouseMove(1, 0, ViewIndicator.Item(0).Top - 25, ViewIndicator.Item(0).Top - 25) 'Move viewindicator up
End Select

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG27:" & Error$
End Sub

Private Sub MH_Shadow2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'PCNGL181202
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'MH_Shadow1_MouseDown Sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    17/11/02     Building initial framework
'
'Description:
'
'Purpose:
'   Initialises position of mouse
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

PipeDisplayMoveLastY = ViewIndicator.Item(0).Top
Select Case Button
    Case 1 ' Left Button Press
        Call PipeDisplay_MouseMove(1, 0, ViewIndicator.Item(0).Top + 10, ViewIndicator.Item(0).Top + 25) 'Move viewindicator up
End Select

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG28:" & Error$
End Sub

Private Sub MH_Shadow2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'PCNGL181202
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'MH_Shadow1_MouseMove Sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    17/11/02     Building initial framework
'
'Description:
'
'Purpose:
'   Initialises position of mouse
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Select Case Button
    Case 1 ' Left Button Press
        Call PipeDisplay_MouseMove(1, 0, ViewIndicator.Item(0).Top + 25, ViewIndicator.Item(0).Top + 25) 'Move viewindicator up
End Select

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG29:" & Error$
End Sub

Public Sub ObservationClose_Click()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : ObservationClose_Click
'Created : 20 Feb 2005, PCN3490
'Updated : 16 august 2005, Antony van Iersel
'Prg By  : Geoff Logan
'Param   :
'Desc    : Closes the Observation panel
'Usage   :
'Update  : If there was a obs closeby that relist the observations instead of closing
'          of completely
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Observation.Visible = False
SpeakObs.Visible = False
PVObsEye.Visible = False
ObsPreviewPictureBox.Visible = False

Me.PVObsIndexNo(1) = ""
Call Observations.ClearEmptyObservations 'PCN4191 once you right click, an observation is made regaurdless
                                         'if there was an obs writen up or not, now it checks to see if anything
                                         'was entered, if not it will be deleted.

If UBound(Observations.ObsIndexArray) > 1 Then Exit Sub 'If there was only one obs in the area, then close

'If there was other obs close by then relist those observation at the last click position''''
If Observations.ObsMouseOverMe(Observations.LastYObsPositionClicked) <> 0 Then 'PCN4131     '
    If UBound(Observations.ObsIndexArray) > 1 Then                                          '
       Call Observations.ObsBubbleWithEyeSetup  'PCN4131                                    '
    End If                                                                                  '
End If                                                                                      '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Exit Sub
Err_Handler:
    Select Case Err
        Case 9: Exit Sub
        Case Else: MsgBox Err & "-PVG30:" & Error$
    End Select
End Sub

Private Sub ObservationClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : ObservationClose_MouseDown
'Created : 20 Feb 2005, PCN3490
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Changes the close image when the mouse is over the close
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler


ObservationClose.Picture = LoadResPicture(110, vbResBitmap)
DoEvents

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG31:" & Error$
End Sub

Private Sub ObservationClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : ObservationClose_MouseMove
'Created : 20 Feb 2005, PCN3490
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Changes the close image when the mouse is over the close
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler


If X > 50 And X < (ObservationClose.width - 50) And Y > 50 And Y < (ObservationClose.height - 50) Then
    ObservationClose.Picture = LoadResPicture(111, vbResBitmap)
Else
    ObservationClose.Picture = LoadResPicture(109, vbResBitmap)
End If


Exit Sub
Err_Handler:
    MsgBox Err & "-PVG32:" & Error$
End Sub

Private Sub ObservationClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : ObservationClose_MouseUp
'Created : 20 Feb 2005, PCN3490
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Changes the close image when the mouse is over the close
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler


ObservationClose.Picture = LoadResPicture(109, vbResBitmap)


Exit Sub
Err_Handler:
    MsgBox Err & "-PVG33:" & Error$
End Sub

Private Sub ObservationEye_Click(Index As Integer)
On Error GoTo Err_Handler
    If UBound(Observations.ObsIndexArray) > 1 Then
        Call Observations.ObsEye_Click(Index)
    Else
        Call ObservationClose_Click
    End If
    Call AdjustIndicatorAndGraph 'Richard 120210
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG34:" & Error$
End Sub



Private Sub PaperClipAttachment_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    If Not Me.PaperClipHighLight.Visible Then Me.PaperClipHighLight.Visible = True
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG35:" & Error$
End Sub

Private Sub PaperClipHighLight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    If Me.PaperClipHighLight.Visible Then Me.PaperClipHighLight.Visible = False
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG36:" & Error$
End Sub

Private Sub PipeDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PipeDisplay_MouseDown Sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    4/11/02     Building initial framework
'
'Description:
'   By clicking on the Pipe Display View Indicator, the user is able to move
'   the indicator up or down and change what section of the pipeline is
'   represented in the PipeDisplay.
'
'Purpose:
'   Initialises position of mouse
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Dim PVPipeLength As Integer
Dim ViewIndicatorHeight As Integer
Dim IndicatorPosition As Integer
Dim CurrentIndicatorTop As Integer

PVPipeLength = PipeMain.height - (80)
ViewIndicatorHeight = PVPipeLength / PVGraphYRatio

Select Case Button
    Case 1 ' Left Button Press
            PipeDisplayMoveLastY = Y
            IndicatorPosition = Y
            CurrentIndicatorTop = Me.ViewIndicator.Item(0).Top
            If (CurrentIndicatorTop < IndicatorPosition) And (CurrentIndicatorTop + ViewIndicatorHeight) > IndicatorPosition Then
                IndicatorOffset = IndicatorPosition - CurrentIndicatorTop
            End If
End Select
        
'PCN3219
If X > 450 Then
    'Determine if over a observation.
    ReDim Observations.ObsIndexArray(0) As Integer
    Observations.LastYObsPositionClicked = Y
    Observations.LastObservationType = "Observation"
    If Observations.ObsMouseOverMe(Y) <> 0 Then 'PCN4131
        Call Observations.ObsBubbleWithEyeSetup  'PCN4131
    Else
        PVObsEye.Visible = False
    End If
Else
    PVObsEye.Visible = False
End If


Exit Sub
Err_Handler:
    MsgBox Err & "-PVG37:" & Error$
End Sub

Public Sub PipeDisplay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PipeDisplay_MouseMove Sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    4/11/02     Building initial framework
'   V1.0    Richard Ashcroft, 12/02/10  Rebuild completely and simplify
'
'Description:
'   By clicking on the Pipe Display View Indicator, the user is able to move
'   the indicator up or down and change what section of the pipeline is
'   represented in the PipeDisplay.
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Dim PipeMain_Offset As Long
Dim PVPipeLength As Integer
Dim ViewIndicatorHeight As Integer
Dim IndicatorPosition As Integer
'Dim IndicatorOffset As Integer
Dim CurrentIndicatorTop As Integer

PipeMain_Offset = 40
PVPipeLength = PipeMain.height - (2 * PipeMain_Offset)
ViewIndicatorHeight = PVPipeLength / PVGraphYRatio

Call CentralMoveRoutine



'Debug.Print X & " - " & Y
Select Case Button
    Case 1 ' Left Button Press
        If X <= 450 Then
            CurrentIndicatorTop = Me.ViewIndicator.Item(0).Top
            
            IndicatorPosition = Y
            If IndicatorPosition < (PipeMain.Top + PipeMain_Offset + IndicatorOffset) Then IndicatorPosition = (PipeMain.Top + PipeMain_Offset + IndicatorOffset)
            If IndicatorPosition > (PipeMain.Top + PipeMain.height - PipeMain_Offset - ViewIndicatorHeight + IndicatorOffset) Then IndicatorPosition = (PipeMain.Top + PipeMain.height - PipeMain_Offset - ViewIndicatorHeight + IndicatorOffset)
'            If (CurrentIndicatorTop < IndicatorPosition) And (CurrentIndicatorTop + ViewIndicatorHeight) > IndicatorPosition Then
'                IndicatorOffset = IndicatorPosition - CurrentIndicatorTop
'            End If
            'IndicatorPosition = (IndicatorPosition - (IndicatorPosition - CurrentIndicatorTop))
            
            Me.ViewIndicator.Item(0).Top = IndicatorPosition - IndicatorOffset ' - (IndicatorPosition - CurrentIndicatorTop)
            Me.ViewIndicator.Item(1).Top = IndicatorPosition - IndicatorOffset ' - (IndicatorPosition - CurrentIndicatorTop)
            Me.ViewIndicator.Item(2).Top = IndicatorPosition - IndicatorOffset '- (IndicatorPosition - CurrentIndicatorTop)
            'Me.ViewIndicator.Item(3).Top = IndicatorPosition
            Me.ViewIndicator.Item(4).Top = IndicatorPosition - IndicatorOffset ' - (IndicatorPosition - CurrentIndicatorTop)
            Me.ViewIndicator.Item(5).Top = IndicatorPosition - IndicatorOffset ' - (IndicatorPosition - CurrentIndicatorTop)
            
            GraphStartFrame = ((IndicatorPosition - IndicatorOffset - PipeMain.Top - PipeMain_Offset) / PVPipeLength) * PVDataNoOfLines
            GraphEndFrame = GraphStartFrame + (PVDataNoOfLines / PVGraphYRatio)
            'Richard
            Call DrawPVGraphs
            Call RepositionYScaleMarkers    'PCN3373
                
            '^^^^ ***********************************************************
            '^^^^ *******************************
            'Setup Base cover
            'Call SetPositionOfPVGraphBaseCover 'PCN2514 'PCN3373
            'Set last Y
            'PipeDisplayMoveLastY = Y
            
            'Redraw All Graph in the next window if in Current Page
'            Call DrawPVGraphAllAtCurrentWindow 'PCN3373
                
            'End If
            Call AdjustRulerWidths
        End If
    Case Else
    
    Dim ClosestObs As Integer
        
    If X > 450 Then
        'Determine if over a observation.
        Observations.LastObservationType = "Observation"
        ReDim Observations.ObsIndexArray(1) As Integer
        ClosestObs = Observations.ObsMouseOverMe(Y)
        If ClosestObs <> 0 Then  'PCN4131
            PVObsEye.ZOrder 0
            PrecisionVisionGraph.PVObsEye.Top = ObsCalcPos(ClosestObs) - (PrecisionVisionGraph.PVObsEye.height / 2)
            PrecisionVisionGraph.PVObsEye.Visible = True
        Else
            PVObsEye.Visible = False
        End If
    Else
        PVObsEye.Visible = False
    End If
End Select


''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''PipeDisplay_MouseMove Sub  Geoff Logan geofflogan@cbsys.co.nz
''
''Revision history"
''   V0.0    Geoff Logan,    4/11/02     Building initial framework
''
''Description:
''   By clicking on the Pipe Display View Indicator, the user is able to move
''   the indicator up or down and change what section of the pipeline is
''   represented in the PipeDisplay.
''
''Purpose:
''
''
''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'On Error GoTo Err_Handler
'Dim Start_X As Long  'PCN1916
'Dim PipeMain_Offset As Long
'Dim End_X As Long
'Dim X_Step As Long 'PCNGL161202
'Dim X_Top_Limit As Long
'Dim X_Bottom_Limit As Long  'PCNGL161202
'Dim NextTop As Long
'Dim NextBotton As Long 'PCNGL161202
'Dim PI_X_Scale As Double 'PCNGL010503-1
'Dim PV_Y As Long 'PCNGL161202
'Dim GraphSetIndex As Integer 'PCN2970
'Dim PV_Y_Adjusted As Long 'PCN2970
'Dim ObsIndex As Integer 'PCN3490
'
'PipeMain_Offset = 50
'Call CentralMoveRoutine
'
''Debug.Print X & " - " & Y
'Select Case Button
'    Case 1 ' Left Button Press
'        If X <= 450 Then
'            X_Top_Limit = Me.PipeMain.Top + PipeMain_Offset
'            X_Bottom_Limit = Me.PipeMain.Top + Me.PipeMain.height - PipeMain_Offset
'            X_Step = Y - PipeDisplayMoveLastY
'            NextTop = Me.ViewIndicator.Item(0).Top + X_Step
'            NextBotton = NextTop + Me.ViewIndicator.Item(0).height
'            If NextTop < X_Top_Limit Then
'                NextTop = X_Top_Limit
'            End If
'            If NextBotton > X_Bottom_Limit Then
'                NextTop = X_Bottom_Limit - Me.ViewIndicator.Item(0).height
'                NextBotton = X_Bottom_Limit
'                PipeDisplayMoveLastY = Y
'            End If
'            'If PipeDisplayMoveLastY <> 0 Then
'                Me.ViewIndicator.Item(0).Top = NextTop
'                Me.ViewIndicator.Item(1).Top = NextTop
'                Me.ViewIndicator.Item(2).Top = NextTop
'                Me.ViewIndicator.Item(4).Top = NextTop
'                Me.ViewIndicator.Item(5).Top = NextTop
'                'Re-draw pipe 3D shading and shadow
'                If X_Step > 0 Then
'                    Start_X = NextTop - X_Step
'                    End_X = Me.ViewIndicator.Item(0).Top
'                Else
'                    Start_X = Me.ViewIndicator.Item(0).Top + Me.ViewIndicator.Item(0).height - (X_Step + 10)
'                    End_X = Start_X + X_Step
'                End If
'                'Call CreatePipeDisplay(Start_X, 0, End_X)
'                'Determine what is portion of the PV graph is displayed in PVGWindow
'                PI_X_Scale = ViewIndicator.Item(0).height / (PipeMain.height - 2 * PipeMain_Offset)
'                If PI_X_Scale = 0 Then PI_X_Scale = 0.000001 'PCNGL010503-1
'                PV_Y = (Me.ViewIndicator.Item(0).Top - MH_Shadow1.height - 35) / PI_X_Scale
'                GraphStartFrame = (PV_Y / PVGraphYRatio)
'                GraphEndFrame = (PV_Y / PVGraphYRatio) + (7500 / PVGraphYRatio)
'
'                Call DrawPVGraphs
'                Call RepositionYScaleMarkers    'PCN3373
'
'                For GraphSetIndex = 0 To (NoOfPVGraphSets - 1)
'
'
'                'Adjust PV_Y
'                PV_Y_Adjusted = ((GraphSetIndex) * PVGraphHeightLimit) - PV_Y
'
'
'
'                Next GraphSetIndex
'                '^^^^ ***********************************************************
'                '^^^^ *******************************
'                'Setup Base cover
'                'Call SetPositionOfPVGraphBaseCover 'PCN2514 'PCN3373
'                'Set last Y
'                PipeDisplayMoveLastY = Y
'
'                'Redraw All Graph in the next window if in Current Page
'    '            Call DrawPVGraphAllAtCurrentWindow 'PCN3373
'
'            'End If
'            Call AdjustRulerWidths
'        End If
'    Case Else
'    Dim ClosestObs As Integer
'
'    If X > 450 Then
'        'Determine if over a observation.
'        Observations.LastObservationType = "Observation"
'        ReDim Observations.ObsIndexArray(1) As Integer
'        ClosestObs = Observations.ObsMouseOverMe(Y)
'        If ClosestObs <> 0 Then  'PCN4131
'            PVObsEye.ZOrder 0
'            PrecisionVisionGraph.PVObsEye.Top = ObsCalcPos(ClosestObs) - (PrecisionVisionGraph.PVObsEye.height / 2)
'            PrecisionVisionGraph.PVObsEye.Visible = True
'        Else
'            PVObsEye.Visible = False
'        End If
'    Else
'        PVObsEye.Visible = False
'    End If
'End Select
'
''Call DrawPVGraphLimitLines  'PCN1718 'PCN3402 removed

Exit Sub
Err_Handler:
Select Case Err
    Case Else
       MsgBox Err & "-PVG38:" & Error$
End Select
End Sub

Sub AdjustRulerWidths()
On Error GoTo Err_Handler

Dim RulerBottom As Single
Dim RulerTop As Single
Dim ViewBottom As Single
Dim ViewTop As Single

Dim i As Integer



ViewBottom = ViewIndicator(0).Top + ViewIndicator(0).height + PipeDisplay.Top
ViewTop = ViewIndicator(0).Top + PipeDisplay.Top

For i = 1 To RulerMark.Count - 1
    RulerBottom = RulerMark(i).Top + RulerMark(i).height
    RulerTop = RulerMark(i).Top
    If RulerBottom > ViewTop And RulerTop < ViewBottom Then
        RulerMark(i).Left = 3415
    Else
        RulerMark(i).Left = 3530
    End If
Next i

  

    
    
Exit Sub
Err_Handler:
Select Case Err
    Case Else
       MsgBox Err & "-PVG39:" & Error$
End Select
End Sub

Private Sub PipeDisplay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PipeDisplay_MouseDown Sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    4/11/02     Building initial framework
'
'Description:
'   By clicking on the Pipe Display View Indicator, the user is able to move
'   the indicator up or down and change what section of the pipeline is
'   represented in the PipeDisplay.
'
'Purpose:
'   Resets position of mouse
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Select Case Button
    Case 1 ' Left Button Press
            PipeDisplayMoveLastY = 0
End Select


Exit Sub
Err_Handler:
    MsgBox Err & "40:" & Error$
End Sub

Private Sub PopupReportsToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PopupReportstoolbar_ButtonClick Sub Michelle Lindsay michellelindsay@cbsys.co.nz
'
'Revision history
'   V0.0    Michelle Lindsay,   20/12/02    Building initial framework
'   V0.1    Michelle Lindsay,   10/01/03    Alter statement to call functions
'
'Description:
'
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Select Case Button.Key
    'Case SnapShot Not required at this time, ML 07.01.03
    '    PopupReportsToolbar.Visible = False
    '    If isopen("LaserImageReport") Then Unload LaserImageReport
    '    LaserImageReport.Show
    Case "PVPortrait"
        'Call PVPortraitReport(Observations) 'PCN2777
        Load PVReportProfile
    Case "PVLandscape" 'PCNGL040103
        Call PVLandscapeReport(PrecisionVisionGraph)
    Case "PVGraph" 'PCN2777
        'Call PVLandscapeSingleReport(PrecisionVisionGraph)
        Load PVReportSingle
    'vvvv PCN3401 ********************************************************
    Case "4_in_1"
        'This report shows four different PVGraphs (eg Capacity, Ovality,
        'Delta or XY and the Flat) on the same page.
        Load PVReport4in1
    '^^^^ ****************************************************************
    'vvvv PCN3479 **************************************************
    Case "MultiProfile"
        'This report is a multi profile report and is based on the
        'Profiles at all of the Observations points.
        Load PVReportMultiProfilex3
        'Call PVMultiProfileReport(PrecisionVisionGraph)
    '^^^^ **********************************************************
    Case Else
    
End Select

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG41:" & Error$

End Sub

Private Sub PopupViewToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PopupViewToolbar_ButtonClick
'Created : 12 November 2002,
'Updated : 18 November 2003, PCN2402
'Prg By  : Geoff Logan
'Param   :
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

'Action on button press
Select Case Button.Key
    Case "PipeDetailsPage"
        Load PipelineDetails
        PipelineDetails.Show
        PipelineDetails.ZOrder 0

'PCV4131
'    Case "ObsPage"
'        Load Observations
'        Observations.Show
'        Observations.ZOrder 0

    Case "PVGPage"
    
    Case "OptionsPage"
        Load OptionsPage
        OptionsPage.Show
        OptionsPage.ZOrder 0

    Case Else
End Select



Exit Sub
Err_Handler:
    MsgBox Err & "-PVG42:" & Error$
End Sub



Private Sub ProfileReportObsButtonImage_Click()
'PCN     : 4142
'Name    : ProfileReportImage_Click
'Created : 29 August 2006
'Prg By  : Antony van Iersel
'Param   :
'Desc    : This button is used create a profile report from the observation
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim PVDHeaderEmbedded As PVDHeaderEmbeddedType

    If PipeObservations(ObsOpenIndex).PipeObsSnapshotOffset <> 0 Then
        Call PageFunctions.EmbeddedFileExtract(WindowsTempDirectory & EmbedJMPFileNameAndPath, _
                                                PipeObservations(ObsOpenIndex).PipeObsSnapshotOffset, _
                                                PipeObservations(ObsOpenIndex).PipeObsSnapshotLength, _
                                                PVDHeaderEmbedded)
    End If
    RenderScale = 1
    Load PVReportProfile
    PVReportProfile.ObservationsLabel.Caption = PipeObservations(ObsOpenIndex).PipeObs
    PVReportProfile.ObservationsLabel.Tag = "Visible"
    PVReportProfile.ObservationsLabel.Visible = True
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG43:" & Error$
End Sub

Private Sub PVGraphExpandBtn_Click()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PVGraphExpandBtn_Click
'Created : 20 Feb 2005, PCN3402
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : This button is used to expand the PVGraphs form to show 4 PV graphs
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

If PVGraphExpandBtn.Tag = "0" Then
    PVGraphExpandBtn.Tag = "-1"
    Call PVGraphExpand
    Call HistogramDiameterGraphSetup("Open") 'PCN3489
Else
    PVGraphExpandBtn.Tag = "0"
    Call HistogramDiameterGraphSetup("Close") 'PCN3489
    Call PVGraphResetForm
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG45:" & Error$
End Sub



Private Sub PVGraphImage_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PVGraphScreen_Click Sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    5/12/02     Building initial framework
'
'Description:
'       Jump to the corresponding frame of the current cursor location
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
'Dim RepositionToNewY As Long 'PCN2970
Dim GraphIndex As Integer
Dim PVFileLoadError As Boolean
Dim PVYFrame As Single

'This function is disabled while recording 'PCNGL240103
If PVRecording = True Then
    'MsgBox DisplayMessage("This function is disabled while recording."), vbInformation: Exit Sub
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("This function is disabled while recording."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0: Exit Sub
End If

If mediatype = "Live" Then
    'MsgBox DisplayMessage("Must disconnect from video before analysing Precision Vision data."), vbInformation 'PCN2111
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Must disconnect from video before analysing Precision Vision data."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Exit Sub
End If

PVYFrame = PVGraphYFromFrame(PVFrameNo)

Call ClearLineScreen.ProfilerPause

Select Case Button
    Case 1
        PVFrameNo = PVGraphFrameFromY(Y)
        Call ClearLineScreen.GotoPVProfile(PVFrameNo, True)
    Case 2
        'vvvv PCN3809 ********************************
        If SoftwareConfiguration <> "Reader" Then
            If Abs(PVYFrame - Y) > 90 Then
                PVFrameNo = PVGraphFrameFromY(Y)
                Call ClearLineScreen.GotoPVProfile(PVFrameNo, True)
            End If
            Call Observations.ObsAddNewOrUpdateSetup(PVFrameNo)  'PCN4131
            Call Observations.ObsAddNewOrUpdateSetup(PVFrameNo, Observations.ObsIndexArray(1)) 'PCN4131
        End If
        '^^^^ ****************************************
End Select
If mediatype = Video Then Call ClearLineScreen.MoveSlider(PVTimes(PVFrameNo))


'If CLPScreenMode = SnapShot Then Call ClearLineScreen.TakeASnapShot

'Call PVDebugOvalityCalcCPP(PVFrameNo)
'Call PVAutoRotateCPP(PVFrameNo - 1, PVFrameNo - 1, True)
If DebrisOn Then GraphInfoContainer(PVDebris).DataSingle(PVFrameNo) = PVDebrisCalc(PVFrameNo, True) 'PCN4461
Call PVGraphsKeyForm.PVGraphsKeyUpdate 'PCN2990
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG46:" & Error$
End Sub



Public Sub PVGraphScreen_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PVGraphScreen_Click Sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    5/12/02     Building initial framework
'
'Description:
'       Jump to the corresponding frame of the current cursor location
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
'Dim RepositionToNewY As Long 'PCN2970
Dim GraphIndex As Integer
Dim PVFileLoadError As Boolean

'This function is disabled while recording 'PCNGL240103
If PVRecording = True Then
    'MsgBox DisplayMessage("This function is disabled while recording."), vbInformation 'PCN2111
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("This function is disabled while recording."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Exit Sub
End If

If mediatype = "Live" Then
    'MsgBox DisplayMessage("Must disconnect from video before analysing Precision Vision data."), vbInformation 'PCN2111
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Must disconnect from video before analysing Precision Vision data."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Exit Sub
End If

'vvvv PCN2208 *********************************************
'No longer required since all file reading is not locking
'other processes to lock the file read.
'If DrawingFlatGraph = True Then
'    MsgBox "This is disabled while drawing the graph"
'    Exit Sub
'End If
'^^^^ *****************************************************
    
ClearLineScreen.PVScreen.Cls
'Move start and finish Y Markers
'vvvv PCN2970 *****************************************
'Setup markers on the current set of graphs

Call SetupYScaleMarkers(Y)
'^^^^ *************************************************

If PVFrameNo < 1 Then PVFrameNo = 1  'PCNLS2132


If ThreeDRunning = True And CLPScreenMode = ThreeD Then 'PCN2446
    Call d3d_laser_focus(PVFrameNo)
    'Call DrawPVGraphLimitLines 'PCN2240 'PCN3402 removed
    'Exit Sub PCN2875, because it exit here, the Video etc was not updated.
End If

'vvvv Load required profile into the PVData buffer array 'PCNGL140103 *********
'Call LoadPVDataFromFile(PVDFileName, PVFrameNo, 1, PVFileLoadError) 'PCN2164
Call DrawPVProfile_Setup(ClearLineScreen.PVScreen) 'PCN1871 PCN3526
Call DrawProfilesStartToFinish(ClearLineScreen.PVScreen, True) 'PCN1835

'PCNLS Make Picture in Picture draw frame
If PicInPicMode = PV Then
    Call DrawPVProfile_Setup(ClearLineScreen.PVScreenPicInPic) 'PCN3526
    Call DrawProfilesStartToFinish(ClearLineScreen.PVScreenPicInPic, True) 'PCN1835
End If

'Update PV frame status bar
Call ClearLineScreen.SetPVFrameStatus 'PCN4171

'If an video file (eg AVI) is loaded then goto the require frame no 'PCNGL150103
If CheckAVIInitialised = True Then 'Check that the AVI is correctly initialised before running the C code
    Call ClearLineScreen.GotoPVProfile(PVFrameNo, True)
End If
'^^^^ ***************************************************************************

'Call DrawPVGraphLimitLines 'PCNGL280103 'PCN3402 removed

'vvvv PCN1938 *********************************
If ImageGraphState(0).GraphType = "Delta" Then
    Call PVDeltaCalcShowMaxMin(1)
End If
'^^^^ *****************************************

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG47:" & Error$
End Sub














Private Sub PVGraphScreen_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PVGraphScreen_MouseMove
'Created : 14 September 2004, PCN2970
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : For PVGraphScreen MouseMove redraw the limitlines.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

'Call DrawPVGraphLimitLines 'PCN3402 removed

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG48:" & Error$
End Sub


Private Sub PVGraphImage_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    Call CentralMoveRoutine
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG49:" & Error$
End Sub

Private Sub PVGraphSelect_Click(Index As Integer)
On Error GoTo Err_Handler
    Call SetPVGraphSelectSetting(Index, PVGraphSelectBox.Tag)
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG50:" & Error$
End Sub

Private Sub PVGraphSelect_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    Call HighlightPVGraphSelect(Index)
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG51:" & Error$
End Sub

Private Sub PVGraphSelectBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    Call BoundryCheckForCollapse(X, Y) 'Check boundry check to see if the PvGraphSelectBox needs to be collapsed
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG52:" & Error$
End Sub

Private Sub PVObsDistance_Click()
On Error GoTo Err_Handler
    Call ClearLineScreen.GotoPVProfile(PipeObservations(ObsOpenIndex).PipeObsFrameNo, True)
    Load DataEntryForm
    Call DataEntryForm.SetDataEntryType("DistanceChange")
    DataEntryForm.Left = Observation.Left - DataEntryForm.width + PrecisionVisionGraph.Left
    DataEntryForm.Top = Observation.Top - 300
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG53:" & Error$
End Sub

Private Sub PVObsEntry_Change()
'****************************************************************************************
'Name    : PVObsEntry_Change
'Created : 21 March 2005, PCN3490
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Add a new observation to the Pipe Observations on change if new other
'          wise update.
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler

'If Observations.AddNewObsFlag Then
'    Call Observations.ObservationAddNew 'PCN4131
'    Call Observations.ObsDisplayALL  'PCN4131
'Else
'    'Update current PipeObservation record
    Call Observations.ObservationUpdate  'PCN4131
'End If

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG54:" & Error$
End Sub



Private Sub PVObsEntry_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Handler

    If LanguageCharset <> 0 Then
        PVObsEntry.Font.Charset = LanguageCharset
'        RichPVObsEntry.Top = PVObsEntry.Top
'        RichPVObsEntry.width = PVObsEntry.width
'        RichPVObsEntry.Left = PVObsEntry.Left
'        RichPVObsEntry.height = PVObsEntry.height
'        RichPVObsEntry.Visible = True
'        RichPVObsEntry.SetFocus
    End If

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG55:" & Error$
End Sub

Private Sub PVObsEntry_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Handler
    Me.PVObsEntry = Observations.RemoveCrLf(Me.PVObsEntry.text) 'PCN4389
    If KeyCode <> vbKeyReturn Then 'PCN4389
        Call Observations.ObservationUpdate
    Else
        Call Observations.ObsBubbleWithEyeSetup
    End If
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG55:" & Error$
End Sub





Private Sub PVObsEyeImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    Call PipeDisplay_MouseDown(Button, Shift, X + PVObsEye.Left + PVObsEyeImage.Top, Y + PVObsEye.Top + PVObsEyeImage.Top + 2)
    Call AdjustIndicatorAndGraph
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG56:" & Error$
End Sub

Private Sub PVObsEyeImage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    Call PipeDisplay_MouseMove(Button, Shift, X + PVObsEye.Left + PVObsEyeImage.Top, Y + PVObsEye.Top + PVObsEyeImage.Left + 2)
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG57:" & Error$
End Sub




Private Sub PVScaleMarkerFhAdjuster_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PVScaleMarkerStAdjuster_MouseDown Sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    5/12/02     Building initial framework
'
'Description:
'       The finish scale marker is moved up or down the PVYScale by the finish Adjuster
'Purpose:
'       To reset the move direction flag
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

PVYScaleMoveDir = ""
PVYScaleFhLastY = Y
ClearLineScreen.PVScreen.Cls

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG58:" & Error$
End Sub


Private Sub PVScaleMarkerFhAdjuster_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single) 'PCN2970
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PVScaleMarkerFhAdjuster_MouseMove Sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    5/12/02     Building initial framework
'
'Description:
'       The Finish scale marker is moved up or down the PVYScale by the Finish Adjuster
'Purpose:
'       To move of the PVScaleMarker allow the PVYScale
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim GraphIndex As Integer

If Button = 1 Then 'Left mouse press
    'Adjust finish marker
    If Y - PVYScaleFhLastY > 0 And PVYScaleMoveDir = "D" Then
        PVYScaleMoveDir = "D"
    ElseIf Y - PVYScaleFhLastY < 0 And PVYScaleMoveDir = "U" Then
        PVYScaleMoveDir = "U"
    ElseIf Y - PVYScaleFhLastY > 0 And PVYScaleMoveDir = "" Then
        PVYScaleMoveDir = "D"
        PVYScaleFhLastY = Y
        Exit Sub
    ElseIf Y - PVYScaleFhLastY < 0 And PVYScaleMoveDir = "" Then
        PVYScaleMoveDir = "U"
        PVYScaleFhLastY = Y
        Exit Sub
    Else
        PVYScaleFhLastY = Y
        Exit Sub
    End If
    'Adjust markers
    'vvvv PCN2970 ************************************************
    'Determine the Graph set and setup markers
    If YScaleMarkerFh(0).y1 + (Y - PVYScaleFhLastY) >= YScaleMarkerSt(0).y1 Then
        PVScaleMarkerFhAdjuster(0).Top = PVScaleMarkerFhAdjuster(0).Top + (Y - PVYScaleFhLastY)
        YScaleMarkerFh(0).y1 = YScaleMarkerFh(0).y1 + (Y - PVYScaleFhLastY)
    Else
        PVScaleMarkerFhAdjuster(0).Top = YScaleMarkerSt(0).y1 + 1
        YScaleMarkerFh(0).y1 = YScaleMarkerSt(0).y1
    End If
    YScaleMarkerFh(Index).y2 = YScaleMarkerFh(0).y1
    'Dim IndexStart As Integer
    'IndexStart = NoOfPVGraphs * Index
    'For GraphIndex = IndexStart To IndexStart + (NoOfPVGraphs - 1)
        YMarkerFh(0).y1 = YScaleMarkerFh(0).y1
        YMarkerFh(0).y2 = YScaleMarkerFh(0).y1
        PVScaleMarkerFnFrame = PVGraphFrameFromY(YScaleMarkerFh(0).y1)
    'Next GraphIndex
    '^^^^ ********************************************************
    'Call DrawPVGraphLimitLines 'PCNGL050203 'PCN3402 removed
End If
PVYScaleFhLastY = Y

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG59:" & Error$
End Sub

Private Sub PVScaleMarkerFhAdjuster_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single) 'PCN2970
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PVScaleMarkerFhAdjuster_MouseUp Sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    5/12/02     Building initial framework
'
'Description:
'       Once the finish adjuster has been moved, the PV Profile needs to be
'       redrawn from the start to finish markers
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

'vvvv PCN2970 *****************************************
'Setup markers on the current set of graphs
Dim SelectedGraphSet As Integer
SelectedGraphSet = Int(Index / NoOfPVGraphs) + 1
'^^^^ *************************************************

'check what the display mode is 'PCNGL030103
If PicInPicMode = PV Then
    Call DrawProfilesStartToFinish(ClearLineScreen.PVScreenPicInPic, True)
Else
    Call DrawProfilesStartToFinish(ClearLineScreen.PVScreen, True)
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG60:" & Error$
End Sub


Private Sub PVScaleMarkerStAdjuster_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single) 'PCN2970
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PVScaleMarkerStAdjuster_MouseDown Sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    5/12/02     Building initial framework
'
'Description:
'       The Start scale marker is moved up or down the PVYScale by the Start Adjuster
'Purpose:
'       To reset the move direction flag
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

PVYScaleMoveDir = ""
PVYScaleStLastY = Y
'ClearLineScreen.PVScreen.Cls 'PCNGL240103

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG61:" & Error$
End Sub

Private Sub PVScaleMarkerStAdjuster_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single) 'PCN2970
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PVScaleMarkerStAdjuster_MouseDown Sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    5/12/02     Building initial framework
'
'Description:
'       The Start scale marker is moved up or down the PVYScale by the Start Adjuster
'Purpose:
'       To move of the PVScaleMarker allow the PVYScale
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim GraphIndex As Integer
    Dim Frame As Long
    
If Button = 1 And PVRecording <> True Then  'Left mouse press 'Disable when recording PCNGL240103
    'Adjust starter marker
    If Y - PVYScaleStLastY > 0 And PVYScaleMoveDir = "D" Then
        PVYScaleMoveDir = "D"
    ElseIf Y - PVYScaleStLastY < 0 And PVYScaleMoveDir = "U" Then
        PVYScaleMoveDir = "U"
    ElseIf Y - PVYScaleStLastY > 0 And PVYScaleMoveDir = "" Then
        PVYScaleMoveDir = "D"
        PVYScaleStLastY = Y
        Exit Sub
    ElseIf Y - PVYScaleStLastY < 0 And PVYScaleMoveDir = "" Then
        PVYScaleMoveDir = "U"
        PVYScaleStLastY = Y
        Exit Sub
    Else
        PVYScaleStLastY = Y
        Exit Sub
    End If
    'Adjust markers
    'vvvv PCN2970 ************************************************
    'Determine the Graph set and setup markers
    If YScaleMarkerSt(0).y1 + (Y - PVYScaleStLastY) <= YScaleMarkerFh(0).y1 Then
        PVScaleMarkerStAdjuster(0).Top = PVScaleMarkerStAdjuster(0).Top + (Y - PVYScaleStLastY)
        YScaleMarkerSt(0).y1 = YScaleMarkerSt(0).y1 + (Y - PVYScaleStLastY)
    Else
        PVScaleMarkerStAdjuster(0).Top = YScaleMarkerFh(0).y1 - PVScaleMarkerStAdjuster(0).height
        YScaleMarkerSt(0).y1 = YScaleMarkerFh(0).y1
    End If
    YScaleMarkerSt(0).y2 = YScaleMarkerSt(0).y1
    'Dim IndexStart As Integer
    'IndexStart = NoOfPVGraphs * Index
    'For GraphIndex = IndexStart To IndexStart + (NoOfPVGraphs - 1)
        YMarkerSt(0).y1 = YScaleMarkerSt(0).y1
        YMarkerSt(0).y2 = YScaleMarkerSt(0).y1
        
    PVScaleMarkerStFrame = PVGraphFrameFromY(YScaleMarkerSt(0).y1)
    'Next GraphIndex
    '^^^^ ********************************************************
    'Call DrawPVGraphLimitLines 'PCNGL050203 'PCN3402 removed
End If
PVYScaleStLastY = Y



Exit Sub
Err_Handler:
    MsgBox Err & "-PVG62:" & Error$
End Sub

Private Sub PVScaleMarkerStAdjuster_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)  'PCN2970
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PVScaleMarkerStAdjuster_MouseUp Sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    5/12/02     Building initial framework
'
'Description:
'       Once the Start adjuster has been moved, the PV Profile needs to be
'       redrawn from the start to finish markers
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

'vvvv PCN2970 *****************************************
'Setup markers on the current set of graphs
Dim SelectedGraphSet As Integer
SelectedGraphSet = Int(Index / NoOfPVGraphs) + 1
'^^^^ *************************************************
'check what the display mode is 'PCNGL030103
If PicInPicMode = PV Then
    Call DrawProfilesStartToFinish(ClearLineScreen.PVScreenPicInPic, True)
Else
    Call DrawProfilesStartToFinish(ClearLineScreen.PVScreen, True)
End If


Exit Sub
Err_Handler:
    MsgBox Err & "-PVG63:" & Error$
End Sub




Private Sub PVXScale_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PVXScale_MouseDown sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    11/12/02     Building initial framework
'
'Description:
'   By pressing the left or right mouse button, when the mouse is over the PVXScale,
'   the left or right XMarker is reset to the current mouse position.
'   The limit marker must then be redrawn.
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Call ScreenDrawing.SetAndStoreLimitLines(Index, Button, X)

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG64:" & Error$
End Sub

Private Sub PVXScale_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PVXScale_MouseMove sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    11/12/02     Building initial framework
'
'Description:
'   When the mouse moves over the PVXScale window, the tool tip displays the % at the
'   current position.
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
'Dim PVXScaleLimitPerL As Single 'PCNGL2901032 'PCN2337 Made Global for 3D Pipe Colour Limits
'Dim PVXScaleLimitPerR As Single 'PCNGL2901032 'PCN2337 Made Global for 3D Pipe Colour Limits

Call DrawPVXToolTips(Index, X) 'PCNGL2801032'PCNGL290103

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG65:" & Error$
End Sub



Private Sub PVXScaleLabelC_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PVXScaleLabelC_Click sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    11/12/02     Building initial framework
'
'Description:
'   By pressing the left or right mouse button, when the mouse is over the PVXScale,
'   the left or right XMarker is reset to the current mouse position.
'   The limit marker must then be redrawn.
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

'Call RepositionPVXLimitMarkers(Button, X + PVXScaleLabelC(0).left, True) 'PCN2680
Call ScreenDrawing.SetAndStoreLimitLines(Index, Button, X + PVXScaleLabelC(Index).Left)
'Call DrawPVGraphLimitLines 'PCNGL280103

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG67:" & Error$
End Sub

Private Sub PVXScaleLabelC_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PVXScaleLabelC_MouseMove sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    28/01/03     Building initial framework
'
'Description:
'
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
'Dim PVXScaleLimitPerL As Single 'PCNGL2901032 'PCN2337
'Dim PVXScaleLimitPerR As Single 'PCNGL2901032 'PCN2337

Call DrawPVXToolTips(Index, PVXScaleLabelC(Index).Left + X) 'PCNGL2801032

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG68:" & Error$
End Sub

Private Sub PVXScaleLabelL_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PVXScaleLabelL_MouseDown sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    11/12/02     Building initial framework
'
'Description:
'   By pressing the left or right mouse button, when the mouse is over the PVXScale,
'   the left or right XMarker is reset to the current mouse position.
'   The limit marker must then be redrawn.
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

'Call RepositionPVXLimitMarkers(Button, X + PVXScaleLabelL(0).left, True) 'PCN2680
'Call DrawPVGraphLimitLines 'PCNGL280103

Call ScreenDrawing.SetAndStoreLimitLines(Index, Button, X + PVXScaleLabelL(Index).Left)

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG69:" & Error$
End Sub

Private Sub PVXScaleLabelL_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PVXScaleLabelL_MouseMove sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    11/12/02     Building initial framework
'
'Description:
'   When the mouse moves over the PVXScale window, the tool tip displays the % at the
'   current position.
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
'Dim PVXScaleLimitPerL As Single 'PCNGL2901032 'PCN2337
'Dim PVXScaleLimitPerR As Single 'PCNGL2901032 'PCN2337

Call DrawPVXToolTips(Index, PVXScaleLabelL(Index).Left + X) 'PCNGL2801032

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG70:" & Error$
End Sub


Private Sub PVXScaleLabelR_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PVXScaleLabelR_MouseDown sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    11/12/02     Building initial framework
'
'Description:
'   By pressing the left or right mouse button, when the mouse is over the PVXScale,
'   the left or right XMarker is reset to the current mouse position.
'   The limit marker must then be redrawn.
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

'Call RepositionPVXLimitMarkers(Button, X + PVXScaleLabelR(0).left, True) 'PCN2680
'Call DrawPVGraphLimitLines 'PCNGL280103
Call PVXScale_MouseDown(Index, Button, Shift, X + PVXScaleLabelR(Index).Left, Y) 'PCN3373

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG71:" & Error$
End Sub

Private Sub PVXScaleLabelR_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PVXScaleLabelR_MouseMove sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    11/12/02     Building initial framework
'
'Description:
'   When the mouse moves over the PVXScale window, the tool tip displays the % at the
'   current position.
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
'Dim PVXScaleLimitPerL As Single 'PCNGL2901032 'PCN2337
'Dim PVXScaleLimitPerR As Single 'PCNGL2901032 'PCN2337

Call DrawPVXToolTips(Index, PVXScaleLabelR(Index).Left + X) 'PCNGL2801032

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG72:" & Error$
End Sub

Function RepositionPVXLimitMarkers()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PVXScaleLabelL_MouseDown sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    11/12/02     Building initial framework
'
'Description:
'   By pressing the left or right mouse button, when the mouse is over the PVXScale,
'   the left or right XMarker is reset to the current mouse position.
'   The limit marker must then be redrawn.
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    Dim i As Integer
    
    For i = 0 To 5
    
    Call GetGeneralPVGraphData(ImageGraphState(i).GraphType)

        XMarkerL(i).x1 = PVGraphXScaleTwipFromUnit(i, PVXScaleLimitPerL): XMarkerL(i).x2 = XMarkerL(i).x1
        XMarkerR(i).x1 = PVGraphXScaleTwipFromUnit(i, PVXScaleLimitPerR): XMarkerR(i).x2 = XMarkerR(i).x1

    Next i
Exit Function
Err_Handler:
    MsgBox Err & "-PVG73:" & Error$
End Function
Sub StoreLimitLinesInINI()
On Error GoTo Err_Handler
Dim FileLoadError As Boolean

Call INI_WriteBack(MyFile, "LimitCapMin=", CapacityLimitL): ConfigInfo.LimitCapacityL = CapacityLimitL
Call INI_WriteBack(MyFile, "LimitCapMax=", CapacityLimitR): ConfigInfo.LimitCapacityR = CapacityLimitR
    
Call INI_WriteBack(MyFile, "LimitOvalR=", OvalityLimitR): ConfigInfo.LimitOvality = OvalityLimitR

''''Call INI_WriteBack(MyFile, "LimitDeltaMin=", DeltaLimitL): ConfigInfo.LimitDeltaL = DeltaLimitL 'PCN4349
''''Call INI_WriteBack(MyFile, "LimitDeltaMax=", DeltaLimitR): ConfigInfo.LimitDeltaR = DeltaLimitR 'PCN4349

Call INI_WriteBack(MyFile, "LimitXYDiameterL=", XYDiameterLimitL): ConfigInfo.LimitXYDiameterL = XYDiameterLimitL
Call INI_WriteBack(MyFile, "LimitXYDiameterR=", XYDiameterLimitR): ConfigInfo.LimitXYDiameterR = XYDiameterLimitR

Call INI_WriteBack(MyFile, "LimitOvalL=", OvalityLimitL): 'PCN3540 Added 16 August 2005, Antony

Call INI_WriteBack(MyFile, "LimitDiameterMedianL=", DiameterMedianLimitL): ConfigInfo.LimitDeltaL = DiameterMedianLimitL 'PCN3540 Added 16 August 2005, Antony 'PCN4349
Call INI_WriteBack(MyFile, "LimitDiameterMedianR=", DiameterMedianLimitR): ConfigInfo.LimitDeltaR = DiameterMedianLimitR 'PCN3540 Added 16 August 2005, Antony 'PCN4349

Call INI_WriteBack(MyFile, "LimitDiameterMaxR=", DiameterMaxLimitR) 'PCN3540 Added 16 August 2005, Antony 'PCN4349 'PCN4799
Call INI_WriteBack(MyFile, "LimitDiameterMaxL=", DiameterMaxLimitL) 'PCN3540 Added 16 August 2005, Antony 'PCN4349 'PCN4799
    
Call INI_WriteBack(MyFile, "LimitDiameterMinR=", GraphInfoContainer(PVMinDiameter).LimitR) 'PCN3540 Added 16 August 2005, Antony 'PCN4349 'PCN4799
Call INI_WriteBack(MyFile, "LimitDiameterMinL=", GraphInfoContainer(PVMinDiameter).LimitL) 'PCN3540 Added 16 August 2005, Antony 'PCN4349 'PCN4799
    
    
'Save to the ConfigInfo Limit Lines to the PVD file
Call PageFunctions.SaveToFilePipeAndConfigInfo("ConfigInfo", FileLoadError)

   
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG74:" & Error$

End Sub

Private Sub PVXScaleLabelUnits_DblClick(Index As Integer)
On Error GoTo Err_Handler
Dim UnitsIndex As Integer
Dim NoOfGraphs As Integer
Dim UnitType As String


'vvvv PCN4207 ************************************
''NoOfGraphs = UBound(PVGraphOrder)
''For UnitsIndex = 0 To NoOfGraphs
''    If PVGraphOrder(UnitsIndex) = ImageGraphState(Index).PreviousGraphType Then
''        UnitType = PVXScaleUnits(UnitsIndex)
''        Exit For
''    End If
''Next UnitsIndex

If ImageGraphState(0).GraphType = "XYDiameter" And MedianFlat Then Exit Sub 'PCN5186 if its deflection then there is only percentage, like ovality

UnitsIndex = GetGraphInfoIndex(Index)
If GraphInfoContainer(UnitsIndex).PVXScaleUnits = "Per" Then
    If GraphInfoContainer(UnitsIndex).GraphType = "Capacity" Then
        GraphInfoContainer(UnitsIndex).PVXScaleUnits = "Area"
    ElseIf GraphInfoContainer(UnitsIndex).GraphType = "Ovality" Then
        Exit Sub
    Else
        GraphInfoContainer(UnitsIndex).PVXScaleUnits = "Real"
    End If
Else
    GraphInfoContainer(UnitsIndex).PVXScaleUnits = "Per"
End If
'^^^^ ********************************************
Call ScreenDrawing.DrawPVXScale
    
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG75:" & Error$
End Sub

Private Sub PVXScaleLabelUnits_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PVXScaleLabelUnits_MouseDown sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    11/12/02     Building initial framework
'
'Description:
'   By pressing the left or right mouse button, when the mouse is over the PVXScale,
'   the left or right XMarker is reset to the current mouse position.
'   The limit marker must then be redrawn.
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Call RepositionPVXLimitMarkers    'PCN2680
'Call DrawPVGraphLimitLines 'PCNGL280103 'PCN3402 removed

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG76:" & Error$
End Sub

Function DrawPVXToolTips(ByVal Index As Single, X As Single) 'PCNGL290103
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'DrawPVXLimitLines function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    28/01/03     Building initial framework
'
'Description:
'   When the mouse moves over the PVXScale window, the tool tip displays the % at the
'   current position.
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim GraphXaxisCentre As Integer 'PCNGL280103
Dim XScaleScaleFactor As Double 'PCNGL2801032

Dim LeftTool As Double
Dim RightTool As Double
Dim CurrentTool As Double
Dim DisplayUnits As String
Dim RealType As String
Dim GraphType As String
Dim UnitsIndex As Integer


CurrentTool = Format(PVGraphXScaleUnitFromTwip(Index, X), "###0.0")


    Call GetGeneralPVGraphData(ImageGraphState(Index).GraphType)
'Select Case GraphType
'    Case 0: LeftTool = PVXScaleLimitPerL: RightTool = PVXScaleLimitPerR: RealType = "Rad"
'    Case 1: LeftTool = CapacityLimitL: RightTool = CapacityLimitR: RealType = "Area"
'    Case "Ovality": LeftTool = OvalityLimitL: RightTool = OvalityLimitR
'    Case "Delta": LeftTool = DeltaLimitL: RightTool = DeltaLimitR: RealType = "Rad"
'    Case "XYDiameter": LeftTool = XYDiameterLimitL: RightTool = XYDiameterLimitR: RealType = "Dia"
'End Select

UnitsIndex = GetGraphInfoIndex(Index) 'PCN4348
'PCN5186 added the next four lines
If ImageGraphState(0).GraphType = "XYDiameter" And MedianFlat Then
    CurrentTool = CurrentTool
    LeftTool = PVXScaleLimitPerL
    RightTool = PVXScaleLimitPerR
ElseIf GraphInfoContainer(UnitsIndex).PVXScaleUnits <> "Real" Then
    CurrentTool = ConvertRealToPerByGraph(CurrentTool, Index, DisplayUnits)
    LeftTool = ConvertRealToPerByGraph(PVXScaleLimitPerL, Index, DisplayUnits)
    RightTool = ConvertRealToPerByGraph(PVXScaleLimitPerR, Index, DisplayUnits)
Else
    CurrentTool = CurrentTool
    LeftTool = PVXScaleLimitPerL
    RightTool = PVXScaleLimitPerR


End If
 
'PCN4348 CurrentTool which is the moving X scale tooltip, was not rounding on display
PVXScale(Index).ToolTipText = Format(CurrentTool, "###0.0") _
                            & DisplayUnits & " Limits: " _
                            & Format(LeftTool, "###0.0") _
                            & ", " _
                            & Format(RightTool, "###0.0")

PVXScaleLabelL(Index).ToolTipText = PVXScale(Index).ToolTipText
PVXScaleLabelC(Index).ToolTipText = PVXScale(Index).ToolTipText
PVXScaleLabelR(Index).ToolTipText = PVXScale(Index).ToolTipText


Exit Function
Err_Handler:
    MsgBox Err & "-PVG78:" & Error$
End Function

Sub GetGeneralPVGraphData(GraphType As String)
'GetPVXLimits() 'PCN2337 Antony 4 November 2003
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'GetPVXLimits function  Antony van Iersel
'
'Revision history"
'   V0.0    Antony van Iersel,    04/11/03     Building initial framework
'
'Description:
'   PVXScaleLimitPerR and L need to be calculated for 3D pipe, but only
'   those values, instead of writing the calculation twice, one for PVGraph and for
'   the 3D Pipe, enclose it with one Function
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Select Case GraphType 'PCNGL280103
   Case "Capacity"
        PVXScaleLimitPerL = CapacityLimitL
        PVXScaleLimitPerR = CapacityLimitR
        PVGraphGeneralXScale = PVGraphCapacityXScale
        PVGraphGeneralXOffset = PVGraphCapacityXOffset
    Case "Ovality" 'PCNGL280103
        PVXScaleLimitPerL = OvalityLimitL
        PVXScaleLimitPerR = OvalityLimitR
        PVGraphGeneralXScale = PVGraphOvalityXScale
        PVGraphGeneralXOffset = PVGraphOvalityXOffset
    Case "Delta"
        PVXScaleLimitPerL = DeltaLimitL
        PVXScaleLimitPerR = DeltaLimitR
        PVGraphGeneralXScale = PVGraphDeltaXScale
        PVGraphGeneralXOffset = PVGraphDeltaXOffset
    Case "XYDiameter"
        PVXScaleLimitPerL = XYDiameterLimitL
        PVXScaleLimitPerR = XYDiameterLimitR
        PVGraphGeneralXScale = PVGraphXYDiaXScale
        PVGraphGeneralXOffset = PVGraphXYDiaXOffset
    Case "MaxMinDiameter"
        PVXScaleLimitPerL = DiameterMaxMinLimitL
        PVXScaleLimitPerR = DiameterMaxMinLimitR
        PVGraphGeneralXScale = PVGraphDiaMaxMinXScale
        PVGraphGeneralXOffset = PVGraphDiaMaxMinXOffset
    Case "MaxDiameter"
        PVXScaleLimitPerL = DiameterMaxLimitL
        PVXScaleLimitPerR = DiameterMaxLimitR
        PVGraphGeneralXScale = PVGraphDiaMaxXScale
        PVGraphGeneralXOffset = PVGraphDiaMaxXOffset
    Case "MedianDiameter"
        PVXScaleLimitPerL = DiameterMedianLimitL
        PVXScaleLimitPerR = DiameterMedianLimitR
        PVGraphGeneralXScale = PVGraphDiaMedianXScale
        PVGraphGeneralXOffset = PVGraphDiaMedianXOffset
    Case "MinDiameter"
        PVXScaleLimitPerL = GraphInfoContainer(PVMinDiameter).LimitL
        PVXScaleLimitPerR = GraphInfoContainer(PVMinDiameter).LimitR
        PVGraphGeneralXScale = GraphInfoContainer(PVMinDiameter).XScale
        PVGraphGeneralXOffset = GraphInfoContainer(PVMinDiameter).XOffset
'PCN6458     Case "Inclination" 'PCN6128
'PCN6458         PVXScaleLimitPerL = GraphInfoContainer(PVInclination).LimitL
'PCN6458         PVXScaleLimitPerR = GraphInfoContainer(PVInclination).LimitR
'PCN6458         PVGraphGeneralXScale = GraphInfoContainer(PVInclination).XScale
'PCN6458         PVGraphGeneralXOffset = GraphInfoContainer(PVInclination).XOffset
End Select


        
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG79:" & Error$
End Sub

Sub SetFromPVXLimits(Index As Integer) 'PCN3402 Antony 4 May 2005
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'GetPVXLimits function  Antony van Iersel
'
'
'Description:
'   Sets the appropriate graph limit lines from the pvxlimit lines
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

    Dim Graph As String
    
    Graph = ImageGraphState(Index).GraphType
    
    Select Case Graph
        Case "Capacity"
             CapacityLimitL = PVXScaleLimitPerL
             CapacityLimitR = PVXScaleLimitPerR
         Case "Ovality" 'PCNGL280103
             OvalityLimitL = PVXScaleLimitPerL
             OvalityLimitR = PVXScaleLimitPerR
         Case "Delta"
             DeltaLimitL = PVXScaleLimitPerL
             DeltaLimitR = PVXScaleLimitPerR
         Case "XYDiameter"
             XYDiameterLimitL = PVXScaleLimitPerL
             XYDiameterLimitR = PVXScaleLimitPerR
         Case "MaxMinDiameter"
             DiameterMaxMinLimitL = PVXScaleLimitPerL
             DiameterMaxMinLimitR = PVXScaleLimitPerR
         Case "MedianDiameter"
              DiameterMedianLimitL = PVXScaleLimitPerL
              DiameterMedianLimitR = PVXScaleLimitPerR
         Case "MaxDiameter"
              DiameterMaxLimitL = PVXScaleLimitPerL
              DiameterMaxLimitR = PVXScaleLimitPerR
        Case "MinDiameter"
            GraphInfoContainer(PVMinDiameter).LimitL = PVXScaleLimitPerL
            GraphInfoContainer(PVMinDiameter).LimitR = PVXScaleLimitPerR
'PCN6458         Case "Inclination" 'PCN6128
'PCN6458             GraphInfoContainer(PVInclination).LimitL = PVXScaleLimitPerL
'PCN6458             GraphInfoContainer(PVInclination).LimitR = PVXScaleLimitPerR
        

              
    End Select

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG80:" & Error$
End Sub

Function MoveMarkerForRecord(Frame As Long) 'PCNGL290103
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'MoveMarkerForRecord function  Louise Shrimpton louises@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    30/01/03     Building initial framework
'
'Description:
'   This function is called when record is pressed, with the current screen frame
'   this will move the marker to that position.
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim RepositionToNewY As Single 'PCN2970

Dim PVFileLoadError As Boolean

RepositionToNewY = (Frame * PVGraphYRatio) + PVYScaleZeroMarker(0).y1
Call SetupYScaleMarkers(RepositionToNewY)
Call RapidReadPVData(PVFrameNo)

ClearLineScreen.PVScreen.AutoRedraw = True 'PCNGL200103
Call ClearLineScreen.DrawPVGrid(ClearLineScreen.PVScreen, MainScaleGrid / 8, CentreLineX, CentreLineY) 'PCNGL200103
ClearLineScreen.PVScreen.AutoRedraw = True 'PCNGL200103
Call DrawPVProfile(ClearLineScreen.PVScreen, 1) 'Goto frame PVFrameNo PCN3526 stop timer removed
ClearLineScreen.PVScreen.AutoRedraw = False 'PCNGL200103
'Update PV frame status bar
Call ClearLineScreen.SetPVFrameStatus 'PCN4171
'If an video file (eg AVI) is loaded then goto the require frame no 'PCNGL150103
If CheckAVIInitialised = True And PVRecording = True Then 'Check that the AVI is correctly initialised before running the C code
    Call ClearLineScreen.GotoPVProfile(PVFrameNo, True)
End If
'^^^^ ***************************************************************************
    
Exit Function
Err_Handler:
    MsgBox Err & "-PVG81:" & Error$
End Function


Function MoveMarker(Frame As Long) 'PCNLS020203
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'MoveMarker function  Louise Shrimpton louises@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    02/02/03     Building initial framework
'
'Description:
'
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim RepositionToNewY As Single 'PCN2970
'Dim StopTimer As Boolean 'Not used by this sub
Dim GraphIndex As Integer
Dim PVFileLoadError As Boolean


If Frame > PVDataNoOfLines Then
    Frame = PVDataNoOfLines
End If

RepositionToNewY = (Frame * PVGraphYRatio) + PVYScaleZeroMarker(0).y1

'Call GraphSetGetTotal(GraphSet)
Call SetupYScaleMarkers(RepositionToNewY)

'PVFrameNo = Int((YScaleMarkerSt.Y1 - PVYScaleZeroMarker.Y1) / PVGraphYRatio) 'PCNGL121202
PVFrameNo = Frame
'^^^^ ********************************************************

If PVFrameNo = 0 Then PVFrameNo = 1

'vvvv Load required profile into the PVData buffer array 'PCNGL140103 *********

''PCN3526
''Call LoadPVDataFromFile(PVDFileName, PVFrameNo, 1, PVFileLoadError)
Call RapidReadPVData(PVFrameNo)


ClearLineScreen.PVScreen.AutoRedraw = True 'PCNGL200103
'ClearLineScreen.PVScreen.Cls 'PCNGL200103
Call ClearLineScreen.DrawPVGrid(ClearLineScreen.PVScreen, MainScaleGrid / 8, CentreLineX, CentreLineY) 'PCNGL200103
ClearLineScreen.PVScreen.AutoRedraw = True 'PCNGL200103
Call DrawPVProfile(ClearLineScreen.PVScreen, 1) 'Goto frame PVFrameNo PCN3526 stop timer removed
ClearLineScreen.PVScreen.AutoRedraw = False 'PCNGL200103
'Update PV frame status bar
Call ClearLineScreen.SetPVFrameStatus 'PCN4171

'If an video file (eg AVI) is loaded then goto the require frame no 'PCNGL150103

'^^^^ ***************************************************************************
If PicInPicMode = PV Then
    Call DrawProfilesStartToFinish(ClearLineScreen.PVScreenPicInPic, True) 'PCN1835
End If
'Call DrawPVGraphLimitLines 'PCNGL280103 'PCN3402 removed
Call PVGraphsKeyForm.PVGraphsKeyUpdate 'PCN2990
    
Exit Function
Err_Handler:
    MsgBox Err & "-PVG82:" & Error$
End Function


Function ResetPVData() 'PCNLS200203
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'MoveMarker function  Louise Shrimpton louises@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    20/02/03     Building initial framework
'
'Description:
'
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    
''''    CurrentGraph = 1  'PCN4328 -removed

    CurrentPVYScalePageNo = 0 'PCNLS200203
    PVDataNoOfLines = 0 'PCN3274 This initialises it to 0, don't know all consequenses.
    PVFrameNo = 0 'PCN2639 'PCN3526 this should start at 0 for no frames recorded, when the first
                           'frame is recorded then the first should be 1
    NumTimesRecorded = 1 'PCN2639
    RequestFrameNo = 1
    MaxDisplayedFrameNo = 0 'Initialise
    Call InitilisePVProfile(MaxFrameBufferNo)

Exit Function
Err_Handler:
    MsgBox Err & "-PVG83:" & Error$
End Function


Function SetupPVGraphScreen(PVGraphScreenType As String)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'MoveMarker function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    30/05/03    PCN1943   Building initial framework
'
'Description:
'
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Call SetLimitLines 'PCN2769
Call RepositionPVXLimitMarkers 'PCNGL181202
Call SetGeneralReportFlag
Exit Function
Err_Handler:
    MsgBox Err & "-PVG84:" & Error$
End Function






Private Sub PVYScaleImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
Dim FrameNo As Single
Dim PVYFrame As Single

    If PVDFileName = "" Or PVRecording Then Exit Sub
    If SoftwareConfiguration = "Reader" Then Exit Sub 'PCN4322

    PVYFrame = PVGraphYFromFrame(PVFrameNo)
    If Button = 2 Then
        If Abs(PVYFrame - Y) > 90 Then
            PVFrameNo = PVGraphFrameFromY(Y)
            Call ClearLineScreen.GotoPVProfile(PVFrameNo, True)
        End If

    
        Load DataEntryForm
        Call DataEntryForm.SetDataEntryType("DistanceChange")
        DataEntryForm.Left = PrecisionVisionGraph.Left + 200
        DataEntryForm.Top = Y
    End If
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG85:" & Error$
End Sub

Private Sub PVYScaleSelectorBox_LostFocus()
On Error GoTo Err_Handler
    Call ShowPVYScaleSelectorBox
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG86:" & Error$
End Sub

Private Sub PVYScaleSetting_Click(Index As Integer)
On Error GoTo Err_Handler
    Call SetPVYScaleSetting(Index)
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG87:" & Error$
End Sub

Private Sub PVYScaleSetting_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    Call HighlightPVYScaleSelector(Index)

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG88:" & Error$
End Sub

Private Sub ReferenceShiftType_Click(Index As Integer)
On Error GoTo Err_Handler
    If Index = 0 Then
        Call Observations.SetShiftRefereneShape(GRADUAL, ObsOpenIndex)
    Else
        Call Observations.SetShiftRefereneShape(EDGE, ObsOpenIndex)
    End If
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG89:" & Error$
End Sub

Private Sub ReportFlagImage_Click(Index As Integer)
On Error GoTo Err_Handler
    ToggleReportFlag (Index)
    
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG90:" & Error$
End Sub

Private Sub RichPVObsEntry_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Handler
    Select Case KeyAscii
        Case 27:
            If LanguageCharset <> 0 Then
                UnicodeObsEntry = RichPVObsEntry.text
                PVObsEntry.text = UnicodeObsEntry
            End If
            RichPVObsEntry.Visible = False
        Case vbKeyReturn
            If LanguageCharset <> 0 Then
                UnicodeObsEntry = LanguageUtil.ConvertRichToAnsi(RichPVObsEntry.TextRTF)
                PVObsEntry.text = UnicodeObsEntry
            End If
            RichPVObsEntry.Visible = False
    End Select
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG91:" & Error$
End Sub


Private Sub RulerGreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    If (PVDFileName = "" And mediatype <> Video) Or PVRecording Then Exit Sub

    LargeGreenRuler.Visible = True
    LargeGreenRuler.Left = RulerGreen.Left
    LargeGreenRuler.Top = RulerGreen.Top - (LargeGreenRuler.height / 2) + (RulerGreen.height / 2)
    LargeGreenRuler.ZOrder 0
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG91:" & Error$
End Sub

Private Sub RulerRed_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    If (PVDFileName = "" And mediatype <> Video) Or PVRecording Then Exit Sub
    
    LargeRedRuler.Visible = True
    LargeRedRuler.Left = RulerRed.Left
    LargeRedRuler.Top = RulerRed.Top - (LargeRedRuler.height / 2) + (RulerRed.height / 2)
    LargeRedRuler.ZOrder 0
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG92:" & Error$
End Sub

Private Sub RulerMark_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    LargeRuler.Visible = True
    LargeRuler.Left = RulerMark(Index).Left
    LargeRuler.Top = RulerMark(Index).Top - (LargeRuler.height / 2) + (RulerMark(Index).height / 2)
    LargeRuler.ZOrder 0
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG93:" & Error$
End Sub

Private Sub ScaleSetting_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***********************************************************************************
'Name    : ScaleSetting_MouseDown
'Created :  21 August 2003, PCN2121
'Updated :
'
'Prg By  : Geoff Logan
'Param   : Stardard, none are used
'
'Desc    : When this textbox is pressed the focus is sent to the Scale setting on
'           the OptionsPage form.
'Usage   :
'***********************************************************************************
On Error GoTo Err_Handler

If PVYScaleSelectorBox.Visible Then
    PVYScaleSelectorBox.Visible = False
Else
    Call ShowPVYScaleSelectorBox
End If
'OptionsPage.DrawingScale.SetFocus

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG94:" & Error$
End Sub

Function GetPVXLimits_Delta(DeltaLimitPerL As Single, DeltaLimitPerR As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : GetPVXLimits_Delta
'Created : 13 March 2004, PCN2680
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim GraphXaxisCentre As Integer
Dim XScaleScaleFactor As Double

'vvvv PCN2829 **************************************
'Const DeltaXScaleMaxMarker  As Integer = 25 'The maximum marker value 'PCNGL2801032
'XScaleScaleFactor = 100 / DeltaXScaleMaxMarker
''The center of the drawing screen
'GraphXaxisCentre = PVGraphScreen(0).width / 2
'
''DeltaLimitPerL = Round((100 * (DeltaLimitL - GraphXaxisCentre) / ((PVXScale6.X1 - GraphXaxisCentre) * XScaleScaleFactor)), 1) 'PCNGL2801032
''DeltaLimitPerR = Round((100 * (DeltaLimitR - GraphXaxisCentre) / ((PVXScale6.X1 - GraphXaxisCentre) * XScaleScaleFactor)), 1) 'PCNGL2801032

DeltaLimitPerL = Round(PVGraphDeltaXScale * 2 * (DeltaLimitL - PVXScaleZeroMarker(0).x1) / (PVXScaleZeroMarker(0).x1), 1)  'PCN2829
DeltaLimitPerR = Round(PVGraphDeltaXScale * 2 * (DeltaLimitR - PVXScaleZeroMarker(0).x1) / (PVXScaleZeroMarker(0).x1), 1)  'PCN2829
'^^^^ **********************************************

Exit Function
Err_Handler:
    MsgBox Err & "-PVG95:" & Error$
End Function

Function SetLimitLines()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : SetLimitLines
'Created : 15 April 2004, PCN2769
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
'vvvv PCN2854 *****************************************
'Const DeltaXScaleMaxMarker  As Double = 25 'The maximum marker value 'Will be changed to a variable for variable XScale
'Const OvalityXScaleMaxMarker  As Integer = 10 'The maximum marker value 'PCNGL2801032 'PCN2184
'Const CapacityXScaleMaxMarker  As Double = 50 'The maximum marker value 'Will be changed to a variable for variable XScale
'Const XYDiameterXScaleMaxMarker As Double = 10
'^^^^ *************************************************
Dim GraphXaxisCentre As Integer
Dim GraphXaxisOvalityZero As Integer
Dim XScaleScaleFactor As Double

'The center of the drawing screen
'GraphXaxisCentre = PVGraphImage(0).width / 2 'PCN3402

'GraphXaxisOvalityZero = PVXScale2(0).X1 'PCN3402

CapacityLimitL = ConfigInfo.LimitCapacityL
CapacityLimitR = ConfigInfo.LimitCapacityR

OvalityLimitL = ConfigInfo.LimitOvality
OvalityLimitR = ConfigInfo.LimitOvality

DeltaLimitL = ConfigInfo.LimitDeltaL
DeltaLimitR = ConfigInfo.LimitDeltaR

XYDiameterLimitL = ConfigInfo.LimitXYDiameterL
XYDiameterLimitR = ConfigInfo.LimitXYDiameterR

If CapacityLimitL < -10000 Or CapacityLimitL > 10000 Then CapacityLimitL = 0 'ID4796 when given just bad limit values, reset to 0
If CapacityLimitR < -10000 Or CapacityLimitR > 10000 Then CapacityLimitR = 0 'ID4796 when given just bad limit values, reset to 0

If OvalityLimitL < -10000 Or OvalityLimitL > 10000 Then OvalityLimitL = 0 'ID4796 when given just bad limit values, reset to 0
If OvalityLimitR < -10000 Or OvalityLimitR > 10000 Then OvalityLimitR = 0 'ID4796 when given just bad limit values, reset to 0

If DeltaLimitL < -10000 Or DeltaLimitL > 10000 Then DeltaLimitL = 0 'ID4796 when given just bad limit values, reset to 0
If DeltaLimitR < -10000 Or DeltaLimitR > 10000 Then DeltaLimitR = 0 'ID4796 when given just bad limit values, reset to 0

If XYDiameterLimitL < -10000 Or XYDiameterLimitL > 10000 Then XYDiameterLimitL = 0 'ID4796 when given just bad limit values, reset to 0
If XYDiameterLimitR < -10000 Or XYDiameterLimitR > 10000 Then XYDiameterLimitR = 0 'ID4796 when given just bad limit values, reset to 0

Exit Function
Err_Handler:
    MsgBox Err & "-PVG96:" & Error$
End Function

Sub SetupYScaleMarkers(Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : SetupYScaleMarkers
'Created : 14 June 2004, PCN2970
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim RepositionToNewY As Long
Dim IndexStart As Integer
Dim CalcPVFrameNo As Double

'PVFrameNo = PVGraphFrameFromY(Y)
PVScaleMarkerStFrame = PVFrameNo
PVScaleMarkerFnFrame = PVFrameNo

Call RepositionYScaleMarkers




Exit Sub
Err_Handler:
    MsgBox Err & "-PVG97:" & Error$
End Sub



Function SetupPVGraphToDisplay(GraphIndex As Integer)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : SetupPVGraphToDisplay
'Created : 24 August 2004, PCN2970
'Updated :
'Prg By  : Geoff Logan
'Param   : GraphIndex - Specifies the PVGraph index for the first set.
'Desc    : Make visible the require set of PVGraphs (eg Capacity, Ovality etc)
'          and makes all other set invisible.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim PVGraphIndex As Integer
Dim CurrentGraphIndex As Integer
Dim PVGraphSet As Integer


'CurrentGraphIndex = 0
'For PVGraphSet = 1 To NoOfPVGraphSets
'    For PVGraphIndex = 0 To (NoOfPVGraphs - 1)
'        If PVGraphIndex = GraphIndex Then
'            PVGraphScreen(CurrentGraphIndex).Visible = True
'        Else
'            If Not PVGraphScreen(CurrentGraphIndex).AutoRedraw Then   'PCN2970
'                PVGraphScreen(CurrentGraphIndex).AutoRedraw = True 'PCN2970
'                PVGraphScreen(CurrentGraphIndex).Cls 'PCN2970
'                PVGraphScreen(CurrentGraphIndex).AutoRedraw = False 'PCN2970
'            End If
'            PVGraphScreen(CurrentGraphIndex).Visible = False
'        End If
'        CurrentGraphIndex = CurrentGraphIndex + 1
'    Next PVGraphIndex
'Next PVGraphSet
'
''vvvv PCN2970 *******************************
'CurrentPVGraphPageNoLOWER = 0
'If imagegraphstate(0).GraphType = "Flat" Then
'
'    CurrentPVGraphPageNoUPPER = MaxDisplayedFrameNo
'Else
'    CurrentPVGraphPageNoUPPER = 0
'End If
'Call DrawPVGraphAllAtCurrentWindow
''^^^^ ***************************************
Call DrawPVGraphs 'PC3373
Exit Function
Err_Handler:
    MsgBox Err & "-PVG98:" & Error$
End Function

Sub RepositionYScaleMarkers()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : RepositionYScaleMarkers
'Created : 31 August 2004, PCN2970
'Updated :
'Prg By  : Geoff Logan
'Param   : GraphSet -
'          RepositionToNewY - The new Y position
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Dim RepositionToNewYStart
Dim RepositionToNewYFinish


RepositionToNewYStart = PVGraphYFromFrame(PVScaleMarkerStFrame)
RepositionToNewYFinish = PVGraphYFromFrame(PVScaleMarkerFnFrame)

YMarkerSt(0).y1 = RepositionToNewYStart
YMarkerSt(0).y2 = RepositionToNewYStart
YMarkerFh(0).y1 = RepositionToNewYFinish
YMarkerFh(0).y2 = RepositionToNewYFinish
YMarkerSt(0).Visible = True
YMarkerFh(0).Visible = True

PrecisionVisionGraph.YScaleMarkerSt(0).y1 = RepositionToNewYStart
PrecisionVisionGraph.YScaleMarkerSt(0).y2 = RepositionToNewYStart
PrecisionVisionGraph.YScaleMarkerFh(0).y1 = RepositionToNewYFinish
PrecisionVisionGraph.YScaleMarkerFh(0).y2 = RepositionToNewYFinish
PrecisionVisionGraph.PVScaleMarkerStAdjuster(0).Top = RepositionToNewYStart - PrecisionVisionGraph.PVScaleMarkerStAdjuster(0).height 'PCN2***
PrecisionVisionGraph.PVScaleMarkerFhAdjuster(0).Top = RepositionToNewYFinish + 1 'PCN2***
PrecisionVisionGraph.YScaleMarkerSt(0).Visible = True
PrecisionVisionGraph.YScaleMarkerFh(0).Visible = True
PrecisionVisionGraph.PVScaleMarkerStAdjuster(0).Visible = True
PrecisionVisionGraph.PVScaleMarkerFhAdjuster(0).Visible = True



Exit Sub
Err_Handler:
    MsgBox Err & "-PVG99:" & Error$
End Sub


Function PVXLimitLineClear(StartGraphIndex As Integer)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PVXLimitLineClear
'Created : 3 September 2004, PCN2970
'Updated :
'Prg By  : Geoff Logan
'Param   : GraphIndex - The first page of the current graph.
'Desc    : Clears all PVGraph pages for the current imagegraphstate(0).GraphType
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim GraphSetIndex As Integer
Dim GraphIndex As Integer

'''' PCN3441 (6 April 2005, Antony van Iersel)
''
''For GraphSetIndex = 0 To NoOfPVGraphSets - 1
''    GraphIndex = StartGraphIndex + (GraphSetIndex * NoOfPVGraphs)
''    If PrecisionVisionGraph.PVGraphScreen(GraphIndex).AutoRedraw Then
''        PrecisionVisionGraph.PVGraphScreen(GraphIndex).AutoRedraw = False
''        PrecisionVisionGraph.PVGraphScreen(GraphIndex).Cls
''        PrecisionVisionGraph.PVGraphScreen(GraphIndex).AutoRedraw = True
''    Else
''        PrecisionVisionGraph.PVGraphScreen(GraphIndex).Cls
''    End If
''Next GraphSetIndex


Exit Function
Err_Handler:
    MsgBox Err & "-PVG100:" & Error$
End Function

Function PVGraphReset()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PVGraphReset
'Created : 6 September 2004, PCN2970
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Resets and clears all the PVGraphs except the Flat3D (since it does
'          not have a x scale and takes too long to redraw.)
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim GraphSetIndex As Integer
Dim GraphIndex As Integer
Dim StartGraphIndex As Integer

'PCN3373
'For GraphSetIndex = 0 To NoOfPVGraphSets - 1
'    For StartGraphIndex = 0 To NoOfPVGraphs - 1
'        If StartGraphIndex <> 3 Then 'Don't reset the Flat3D graph
'            GraphIndex = StartGraphIndex + (GraphSetIndex * NoOfPVGraphs)
'            If PrecisionVisionGraph.PVGraphScreen(GraphIndex).AutoRedraw Then
'                PrecisionVisionGraph.PVGraphScreen(GraphIndex).Cls
'            Else
'                PrecisionVisionGraph.PVGraphScreen(GraphIndex).AutoRedraw = True
'                PrecisionVisionGraph.PVGraphScreen(GraphIndex).Cls
'                PrecisionVisionGraph.PVGraphScreen(GraphIndex).AutoRedraw = False
'            End If
'        End If
'    Next StartGraphIndex
'Next GraphSetIndex

'CurrentPVGraphPageNoUPPER = 0
'CurrentPVGraphPageNoLOWER = 0

Exit Function
Err_Handler:
    MsgBox Err & "-PVG101:" & Error$
End Function

Function PVGraphResetFlat3D()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PVGraphResetFlat3D
'Created : 6 September 2004, PCN2970
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Resets and clears the Flat3D PVGraphs
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim GraphSetIndex As Integer
Dim GraphIndex As Integer
Dim StartGraphIndex As Integer


'''' PCN3441 (6 April 2005, Antony van Iersel)
''
''StartGraphIndex = 3 'Flat3D
''For GraphSetIndex = 0 To NoOfPVGraphSets - 1
''    GraphIndex = StartGraphIndex + (GraphSetIndex * NoOfPVGraphs)
''    If PrecisionVisionGraph.PVGraphScreen(GraphIndex).AutoRedraw Then
''        PrecisionVisionGraph.PVGraphScreen(GraphIndex).Cls
''    Else
''        PrecisionVisionGraph.PVGraphScreen(GraphIndex).AutoRedraw = True
''        PrecisionVisionGraph.PVGraphScreen(GraphIndex).Cls
''        PrecisionVisionGraph.PVGraphScreen(GraphIndex).AutoRedraw = False
''    End If
''Next GraphSetIndex
''
''CurrentPVGraphPageNoUPPER = 0
''CurrentPVGraphPageNoLOWER = 0
''MaxDisplayedFrameNo = 0

Exit Function
Err_Handler:
    MsgBox Err & "-PVG102:" & Error$
End Function

Function LimitLinesForFlat3D()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : LimitLinesForFlat3D
'Created : 9 September 2004, PCN2970
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Calcs the limit lines for Flat3D
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

'DeltaLimitL = ConfigInfo.LimitDeltaL  'PCN2820
'DeltaLimitR = ConfigInfo.LimitDeltaR  'PCN2820
'Flat3dLimitL = DeltaLimitL 'PCNANTONY
'Flat3dLimitR = DeltaLimitR 'PCNANTONY



Exit Function
Err_Handler:
    MsgBox Err & "-PVG103:" & Error$

End Function

Function PVGraphExpand()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PVGraphExpand
'Created : 20 Feb 2005, PCN3402
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Expands the PVGraph form to show the PVGraphs
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim PVGraphOldLeft As Long
Dim OldLeft As Long
Dim OldTop As Long
Dim PVGraphScreenWidth As Long
Dim PVGraphScreenSpacing As Long
Dim PVGraphIndex As Integer


PVGraphExpandBtn.Picture = LoadResPicture(105, vbResBitmap)
Call ShowOrHideMultipleGraphs("Show")

Call DrawPVGraphs
SlideRuleXScale.Visible = False

DisplayTimer.Enabled = True

Exit Function


PrecisionVisionGraph.Left = 0
PrecisionVisionGraph.width = ClearLineScreen.width + PipelineDetails.width

'Reposition the controls on this form so they look as if the have not moved
'relative to the screen.
PVGraphOldLeft = ClearLineScreen.width
OldLeft = PipeDisplay.Left
PipeDisplay.Left = PVGraphOldLeft + OldLeft
Observation.Left = PipeDisplay.Left + (Observation.Left - OldLeft)
OldLeft = PVGraphExpandBtn.Left
PVGraphExpandBtn.Left = PVGraphOldLeft + OldLeft
''OldLeft = PVToolbar.Left
''PVToolbar.Left = PipeDisplay.Left + OldLeft
'Setup PVGraphScreens for the new layout
OldLeft = PVGraphImage(0).Left
OldTop = PVGraphImage(0).Top
PVGraphScreenWidth = PVGraphImage(0).width
PVGraphWindow.width = 4 * (PVGraphScreenWidth + PVGraphScreenSpacing)
'For PVGraphIndex = 2 To 6 Step 2
'    OldLeft = PVGraphScreen(PVGraphIndex).left
'    PVGraphScreen(PVGraphIndex).left = OldLeft + PVGraphScreenWidth + PVGraphScreenSpacing
'    PVGraphScreen(PVGraphIndex).Top = OldTop
'    PVGraphScreen(PVGraphIndex).Visible = True
'    PVGraphScreen(PVGraphIndex + 1).left = PVGraphScreen(PVGraphIndex).left
'    PVGraphScreen(PVGraphIndex + 1).Top = OldTop
'    PVGraphScreen(PVGraphIndex + 1).Visible = True
'Next PVGraphIndex


Exit Function
Err_Handler:
MsgBox Err & "-PVG104:" & Error$
End Function

Function PVGraphResetForm()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PVGraphResetForm
'Created : 20 Feb 2005, PCN3402
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Expands the PVGraph form to show the PVGraphs
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim PVGraphOldLeft As Long
Dim OldLeft As Long

SlideRuleXScale.Visible = False
PVGraphExpandBtn.Picture = LoadResPicture(107, vbResBitmap)


DisplayTimer.Enabled = True

Exit Function

PrecisionVisionGraph.Left = PipelineDetails.Left
PrecisionVisionGraph.width = PipelineDetails.width

'Reposition the controls on this form so they look as if the have not moved
'relative to the screen.
PVGraphOldLeft = ClearLineScreen.width
OldLeft = PipeDisplay.Left
PipeDisplay.Left = OldLeft - PVGraphOldLeft
Observation.Left = PipeDisplay.Left + (Observation.Left - OldLeft)
OldLeft = PVGraphExpandBtn.Left
PVGraphExpandBtn.Left = OldLeft - PVGraphOldLeft
''OldLeft = PVToolbar.Left
''PVToolbar.Left = OldLeft - PVGraphOldLeft




Exit Function
Err_Handler:
    MsgBox Err & "-PVG105:" & Error$
End Function

Private Sub ScaleSetting_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

If PVYScaleSelectorBox.Visible = True Then Call ShowPVYScaleSelectorBox

If X > 50 And X < 300 And Y > 50 And Y < 350 Then
    Me.ScaleSetting.Picture = Me.ButtonImages.ListImages("YScaleHighlight").Picture
Else
    Me.ScaleSetting.Picture = Me.ButtonImages.ListImages("YScale").Picture
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG106:" & Error$
End Sub

Private Sub SlideRuleEdgeLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : SlideRuleEdgeLeft_MouseDown
'Created : 6 April 2005, PCN3402
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Control of movement of SlideRuleCoverImage and Slider
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

PVYScaleMoveDir = ""
PVYScaleStLastY = X

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG107:" & Error$
End Sub

Private Sub SlideRuleEdgeLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : SlideRuleEdgeLeft_MouseMove
'Created : 6 April 2005, PCN3402
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Control of movement of SlideRuleCoverImage and Slider
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler


If Button = 1 Then
    If SlideRuleSlider.Left + (X - PVYScaleStLastY) > 30 And _
    SlideRuleSlider.Left + SlideRuleSlider.width - 2 * (X - PVYScaleStLastY) < (SlideRuleCoverImage.width + 50) And _
       SlideRuleSlider.width - 2 * (X - PVYScaleStLastY) > SliderCentreMarker.width Then
       
        SlideRuleSlider.Left = SlideRuleSlider.Left + (X - PVYScaleStLastY)
        SlideRuleCoverImage.Left = SlideRuleCoverImage.Left - (X - PVYScaleStLastY)
        SlideRuleSlider.width = SlideRuleSlider.width - 2 * (X - PVYScaleStLastY)
        SlideRuleEdgeRight.Left = SlideRuleSlider.width - SlideRuleEdgeRight.width
        SliderCentreMarkerLine.x1 = SliderCentreMarkerLine.x1 - (X - PVYScaleStLastY)
        SliderCentreMarkerLine.x2 = SliderCentreMarkerLine.x1
        SliderCentreMarker.Left = SliderCentreMarker.Left - (X - PVYScaleStLastY)
        Call TransferSliderValuesToGraph
    End If
End If

'Highlight bar
SlideRuleEdgeLeft.BackColor = &H80FF&


Exit Sub
Err_Handler:
    MsgBox Err & "-PVG108:" & Error$
End Sub

Private Sub SlideRuleEdgeRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : SlideRuleCoverImage_MouseDown
'Created : 6 April 2005, PCN3402
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Control of movement of SlideRuleCoverImage and Slider
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

PVYScaleMoveDir = ""
PVYScaleStLastY = X

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG109:" & Error$
End Sub

Private Sub SlideRuleEdgeRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : SlideRuleCoverImage_MouseDown
'Created : 6 April 2005, PCN3402
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Control of movement of SlideRuleCoverImage and Slider
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

If Button = 1 Then
    If SlideRuleSlider.Left - (X - PVYScaleStLastY) > 30 And _
        SlideRuleSlider.Left + SlideRuleSlider.width + 2 * (X - PVYScaleStLastY) < (SlideRuleCoverImage.width + 50) And _
        SlideRuleSlider.width + 2 * (X - PVYScaleStLastY) > SliderCentreMarker.width Then
        
        SlideRuleSlider.Left = SlideRuleSlider.Left - (X - PVYScaleStLastY)
        SlideRuleCoverImage.Left = SlideRuleCoverImage.Left + (X - PVYScaleStLastY)
        SlideRuleSlider.width = SlideRuleSlider.width + 2 * (X - PVYScaleStLastY)
        SlideRuleEdgeRight.Left = SlideRuleSlider.width - SlideRuleEdgeRight.width
        SliderCentreMarkerLine.x1 = SliderCentreMarkerLine.x1 + (X - PVYScaleStLastY)
        SliderCentreMarkerLine.x2 = SliderCentreMarkerLine.x1
        SliderCentreMarker.Left = SliderCentreMarker.Left + (X - PVYScaleStLastY)
        Call TransferSliderValuesToGraph
    End If
End If

'Highlight bar
SlideRuleEdgeRight.BackColor = &H80FF&

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG110:" & Error$
End Sub

Private Sub SlideRuleCoverImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : SlideRuleCoverImage_MouseDown
'Created : 6 April 2005, PCN3402
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Control of movement of SlideRuleCoverImage and Slider
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim curSelect As StdPicture

PVYScaleMoveDir = ""
PVYScaleStLastY = X
'Setup cursor
Set curSelect = LoadResPicture(109, vbResIcon) 'Hand grip - Move icon
SlideRuleCoverImage.MousePointer = 99
SlideRuleCoverImage.MouseIcon = curSelect

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG111:" & Error$
End Sub

Private Sub SlideRuleCoverImage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : SlideRuleCoverImage_MouseMove
'Created : 6 April 2005, PCN3402
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Control of movement of SlideRuleCoverImage and Slider
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim curSelect As StdPicture


If Button = 1 Then

    If SlideRuleSlider.Left + (X - PVYScaleStLastY) > 30 And SlideRuleSlider.Left + (X - PVYScaleStLastY) < (SlideRuleCoverImage.width - SlideRuleSlider.width + 50) Then
        SlideRuleSlider.Left = SlideRuleSlider.Left + (X - PVYScaleStLastY)
        SlideRuleCoverImage.Left = SlideRuleCoverImage.Left - (X - PVYScaleStLastY)
    Call TransferSliderValuesToGraph
    End If
    PVYScaleStLastY = X
    
Else
    'Check mouse pointer
    If SlideRuleCoverImage.MousePointer <> 99 Then
        'Setup cursor
        Set curSelect = LoadResPicture(108, vbResIcon) 'Hand open - Move icon
        SlideRuleCoverImage.MousePointer = 99
        SlideRuleCoverImage.MouseIcon = curSelect
    End If
End If


'Unhighlight Edge bars
SlideRuleEdgeRight.BackColor = &HFF8080
SlideRuleEdgeLeft.BackColor = &HFF8080


Exit Sub
Err_Handler:
    MsgBox Err & "-PVG112:" & Error$
End Sub

Private Sub SlideRuleCoverImage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : SlideRuleCoverImage_MouseDown
'Created : 6 April 2005, PCN3402
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Control of movement of SlideRuleCoverImage and Slider
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim curSelect As StdPicture

'Setup cursor
Set curSelect = LoadResPicture(108, vbResIcon) 'Hand open - Move icon
SlideRuleCoverImage.MousePointer = 99
SlideRuleCoverImage.MouseIcon = curSelect
Call RefreshSlideRule

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG113:" & Error$
End Sub

Public Sub TransferSliderValuesToGraph()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCN3402
'Name    : TransferSliderValuesToGraph
'Created : 7 April 2005,
'Updated :
'Prg By  : Antony van Iersel
'Param   : None
'Desc    : Transfer the pv graph x scale ruler to the graph that is taged by the
'          slide ruler
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    Dim XCentrePer As Double
    Dim XScalePer As Double
    
    XCentrePer = (SlideRuleSlider.width / 2) + SlideRuleSlider.Left ' Get centre
    
    
    ' Convert to 200% of slider offset by 100%, so the midle is 0 then its + - 100%
    XCentrePer = (XCentrePer / CDbl(SlideRuleXScale.width) * 200) - 100
    
    
    ' Convert to 200% of slider offset by 100%, so the midle is 0 then its + - 100%
    ' Also add the smaller the scale the greater scale effect it has on the graph
    XScalePer = SlideRuleSlider.width 'Get current ruler width
    
    
    
    XScalePer = XScalePer / CDbl(SlideRuleXScale.width) * 200 'Convert to 200%
    'XScalePer = XScalePer / ((200 - XScalePer) / 20)
    
    
    'XScalePer = SlideRuleSlider.width / CDbl(SlideRuleXScale.width) * 200
    'XScalePer = (SlideRuleSlider.width / CDbl(SlideRuleXScale.width) * 200) _
    '            / ((200 - (SlideRuleSlider.width / CDbl(SlideRuleXScale.width) * 200)) / 20)
    
    'XCentrePer = Round((XCentrePer / 10), 1) * 10
    'XScalePer = Round((XScalePer / 10), 1) * 10
    
    XCentrePer = Round(XCentrePer / 5, 0) * 5
    XScalePer = Round(XScalePer / 5, 0) * 5
   
    Select Case SlideRuleXScale.Tag
        Case 0: PVGraphGeneralXScale = XScalePer: PVGraphGeneralXOffset = XCentrePer
        Case 1: PVGraphCapacityXScale = XScalePer: PVGraphCapacityXOffset = XCentrePer
        Case 2: PVGraphOvalityXScale = XScalePer: PVGraphOvalityXOffset = XCentrePer
        Case 3: PVGraphDeltaXScale = XScalePer: PVGraphDeltaXOffset = XCentrePer
        Case 4: PVGraphXYDiaXScale = XScalePer: PVGraphXYDiaXOffset = XCentrePer
     End Select
     
    If SlideRuleXScale.Tag = 0 Then
        Select Case ImageGraphState(0).GraphType
            Case "Capacity": PVGraphCapacityXScale = XScalePer: PVGraphCapacityXOffset = XCentrePer
            Case "Ovality": PVGraphOvalityXScale = XScalePer: PVGraphOvalityXOffset = XCentrePer
            Case "Delta": PVGraphDeltaXScale = XScalePer: PVGraphDeltaXOffset = XCentrePer
            Case "XYDiameter": PVGraphXYDiaXScale = XScalePer: PVGraphXYDiaXOffset = XCentrePer
            Case "MaxMinDiameter": PVGraphDiaMaxMinXScale = XScalePer: PVGraphDiaMaxMinXOffset = XCentrePer
            Case "MedianDiameter": PVGraphDiaMedianXScale = XScalePer: PVGraphDiaMedianXOffset = XCentrePer
            Case "MaxDiameter": PVGraphDiaMaxXScale = XScalePer: PVGraphDiaMaxXOffset = XCentrePer
            Case "MinDiameter": GraphInfoContainer(PVMinDiameter).XScale = XScalePer: GraphInfoContainer(PVMinDiameter).XOffset = XCentrePer
'PCN6458             Case "Inclination": GraphInfoContainer(PVInclination).XScale = XScalePer: GraphInfoContainer(PVInclination).XOffset = XCentrePer 'PCN6128
            

        End Select
    Else
       Call GetGeneralPVGraphData(ImageGraphState(0).GraphType) 'PCN????
    End If
    
    Call INI_WriteBack(MyFile, "PVGraphCapacityXScale=", PVGraphCapacityXScale)
    Call INI_WriteBack(MyFile, "PVGraphOvalityXScale=", PVGraphOvalityXScale)
    Call INI_WriteBack(MyFile, "PVGraphDeltaXScale=", PVGraphDeltaXScale)
    Call INI_WriteBack(MyFile, "PVGraphXYDiaXScale=", PVGraphXYDiaXScale)
    Call INI_WriteBack(MyFile, "PVGraphDiaMaxMinXScale=", PVGraphDiaMaxMinXScale)
    Call INI_WriteBack(MyFile, "PVGraphDiaMedianXScale=", PVGraphDiaMedianXScale)
    Call INI_WriteBack(MyFile, "PVGraphDiaMaxXScale=", PVGraphDiaMaxXScale) 'PCN4799
    Call INI_WriteBack(MyFile, "PVGraphDiaMinXScale=", GraphInfoContainer(PVMinDiameter).XScale) 'PCN4333
'PCN6458     Call INI_WriteBack(MyFile, "PVGraphInclinationXScale=", GraphInfoContainer(PVInclination).XScale) 'PCN6128
    
    
'    Call INI_WriteBack(MyFile, "PVGraphFractileXScale=", PVGraphFractileXScale)
    
    Call INI_WriteBack(MyFile, "PVGraphCapacityXOffset=", PVGraphCapacityXOffset)
    Call INI_WriteBack(MyFile, "PVGraphOvalityXOffset=", PVGraphOvalityXOffset)
    Call INI_WriteBack(MyFile, "PVGraphDeltaXOffset=", PVGraphDeltaXOffset)
    Call INI_WriteBack(MyFile, "PVGraphXYDiaXOffset=", PVGraphXYDiaXOffset)
    Call INI_WriteBack(MyFile, "PVGraphDiaMaxMinXOffset=", PVGraphDiaMaxMinXOffset)
    Call INI_WriteBack(MyFile, "PVGraphDiaMedianXOffset=", PVGraphDiaMedianXOffset) 'PCN4799
    Call INI_WriteBack(MyFile, "PVGraphDiaMaxXOffset=", PVGraphDiaMaxXOffset) 'PCN4799
    Call INI_WriteBack(MyFile, "PVGraphDiaMinXOffset=", GraphInfoContainer(PVMinDiameter).XOffset)
'PCN6458     Call INI_WriteBack(MyFile, "PVGraphInclinationXOffset=", GraphInfoContainer(PVInclination).XOffset) 'PCN6128
    
'    Call INI_WriteBack(MyFile, "PVGraphFractileXOffset=", PVGraphFractileXOffset)
    
     
    Call ScreenDrawing.RefreshGraphXScale(ImageGraphState(SlideRuleXScale.Tag), XScalePer, XCentrePer)
    Call DrawPVXScale
     
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG114:" & Error$

End Sub





Private Sub SlideRuleXScale_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

'Unhighlight Edge bars
SlideRuleEdgeRight.BackColor = &HFF8080
SlideRuleEdgeLeft.BackColor = &HFF8080

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG115:" & Error$
End Sub

Private Sub SmoothCheck_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub WaterShiftType_Click(Index As Integer)
On Error GoTo Err_Handler

    If Index = 0 Then
        Call Observations.SetShiftWaterLevel(GRADUAL, ObsOpenIndex)
    Else
        Call Observations.SetShiftWaterLevel(EDGE, ObsOpenIndex)
    End If
  Exit Sub
Err_Handler:
    MsgBox Err & "-PVG116:" & Error$
End Sub

Public Sub XScaleLabel_Click(Index As Integer)
On Error GoTo Err_Handler
    Dim XScale As Double
    Dim XOffset As Double
    
    If SlideRuleXScale.Visible And (SlideRuleXScale.Tag = Index) Then SlideRuleXScale.Visible = False: Exit Sub
    
    SlideRuleXScale.Visible = True
    SlideRuleXScale.Tag = Index
    SlideRuleXScale.Left = PVXScaleBox(Index).Left + _
                            PVGraphWindow.Left - _
                            (SlideRuleXScale.width - PVXScaleBox(Index).width)
                            
    SlideRuleXScale.Top = PVXScaleBox(Index).Top - _
                           PVXScaleBox(Index).height + _
                           PVGraphWindow.Top - 100
                        
    Call GetGeneralPVGraphData(ImageGraphState(Index).GraphType)
                           
'    Select Case Index
'        case 0: PVGraphGeneralXScale:         XOffset = PVGraphGeneralXOffset
'        Case 1: XScale = PVGraphCapacityXScale: XOffset = PVGraphCapacityXOffset
'        Case 2: XScale = PVGraphOvalityXScale: XOffset = PVGraphOvalityXOffset
'        Case 3: XScale = PVGraphDeltaXScale: XOffset = PVGraphDeltaXOffset
'        Case 4: XScale = PVGraphXYDiaXScale: XOffset = PVGraphXYDiaXOffset
'    End Select
                               
    Call AdjustSlideRule(PVGraphGeneralXScale, PVGraphGeneralXOffset)
                           
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG117:" & Error$

    
End Sub

Sub RefreshSlideRule()
On Error GoTo Err_Handler
    'Call AdjustSlideRule(ImageGraphState(SlideRuleXScale.Tag).XScale * 180, _
    '                ImageGraphState(SlideRuleXScale.Tag).CentreOffset * 100 / 180)
    
    Call AdjustSlideRule(ImageGraphState(SlideRuleXScale.Tag).XScale * 100, _
                    ImageGraphState(SlideRuleXScale.Tag).CentreOffset)
    
'XScale = XScale / 180
'CentreOffset = 180 / 100 * CentreOffset
    
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG118:" & Error$

    
End Sub

Sub AdjustSlideRule(ByVal XScale As Double, ByVal XCentre As Double)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCN3402
'Name    : AdjustSlideRule
'Created : 7 April 2005,
'Updated :
'Prg By  : Antony van Iersel
'Param   : XScale, the scale the slide rule has to represent, in a percentage
'          %100 is the intire slide rule, %10 is 10 per of slide rule etc
'          XCentre, the centre positioning of the slide rule
'Desc    : Set the Slide Rule to the given scale and centre
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

On Error GoTo Err_Handler
    Dim AdjustSlideRuleWidth As Double 'In twips
    Dim AdjustSlideRuleLeft As Double 'In twips
    Dim AdjustSlideRuleCentre As Double 'In twips
    Dim AdjustSlideRuleRight As Double 'In twips
    
    If XScale > 200 Then XScale = 200 ' make sure the scale percentage is no greater than 200%, twice the whole graph
    
    'Slide Rule width is the scale percentage, except right across the
    'scale is 200% not 100%
    AdjustSlideRuleWidth = CDbl(SlideRuleXScale.width) * XScale / 200
    
'    AdjustSlideRuleWidth = (40000 * XScale) / ((4000 * XScale) + 200) * CDbl(SlideRuleXScale.width)
'    XScalePer = XScalePer / CDbl(SlideRuleXScale.width) * 200 'Convert to 200%
'    XScalePer = XScalePer / ((200 - XScalePer) / 20)
    
    
    
    AdjustSlideRuleCentre = CDbl(SlideRuleXScale.width) * (XCentre + 100) / 200
    AdjustSlideRuleLeft = AdjustSlideRuleCentre - (AdjustSlideRuleWidth / 2)
    AdjustSlideRuleRight = AdjustSlideRuleCentre + (AdjustSlideRuleWidth / 2)
    
    If AdjustSlideRuleLeft < 0 Then
        AdjustSlideRuleLeft = 0
        AdjustSlideRuleCentre = AdjustSlideRuleLeft + (AdjustSlideRuleWidth / 2)
        AdjustSlideRuleRight = AdjustSlideRuleWidth
    End If
    
    If AdjustSlideRuleRight > SlideRuleXScale.width Then
        AdjustSlideRuleRight = SlideRuleXScale.width
        AdjustSlideRuleCentre = AdjustSlideRuleRight - (AdjustSlideRuleWidth / 2)
        AdjustSlideRuleLeft = AdjustSlideRuleRight - AdjustSlideRuleWidth
    End If
    
    SlideRuleSlider.width = AdjustSlideRuleWidth
    SlideRuleSlider.Left = AdjustSlideRuleLeft
    SlideRuleCoverImage.Left = -SlideRuleSlider.Left + 30
    
    SliderCentreMarkerLine.x1 = SlideRuleSlider.width / 2
    SliderCentreMarkerLine.x2 = SlideRuleSlider.width / 2
    SliderCentreMarker.Left = (SlideRuleSlider.width / 2) - (SliderCentreMarker.width / 2)
    
    SlideRuleEdgeLeft.Left = 0
    SlideRuleEdgeRight.Left = SlideRuleSlider.width - SlideRuleEdgeRight.width
    
    
    
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG119:" & Error$

    
End Sub

Sub ShowOrHideMultipleGraphs(ByVal ShowHide As String)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCN3402
'Name    : ShowOrHideMultipleGraphs
'Created : 14 April 2005,
'Updated :
'Prg By  : Antony van Iersel
'Parm    : ShowHide, if "Show" then show what is needed for multiple graphs
'                    if "Hide" then hide what is need for single graph
'Desc    : When the PVGraphs are expanded or retracted certain
'          graphs, lables, etc have to be showed or hidden
'          XCentre, the centre positioning of the slide rule
'Usage   : When ever the PVGraph is expanded, or retracted, this is called
'          to match accordingly
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

On Error GoTo Err_Handler
    
    If ShowHide = "Hide" Then
                                                              
        PVGraphImage(0).Visible = True
        PVGraphImage(1).Visible = False
        PVGraphImage(2).Visible = False
        PVXScaleBox(0).Visible = True
        PVXScaleBox(1).Visible = False
        PVXScaleBox(2).Visible = False
        GraphCaptionBox(0).Visible = True
        GraphCaptionBox(1).Visible = False
        GraphCaptionBox(2).Visible = False
        PVGraph3DBoarder(0).Visible = False
    
    Else
        PVGraphImage(0).Visible = False
        PVGraphImage(1).Visible = True
        PVGraphImage(2).Visible = True
        PVXScaleBox(0).Visible = False
        PVXScaleBox(1).Visible = True
        PVXScaleBox(2).Visible = True
        GraphCaptionBox(0).Visible = False
        GraphCaptionBox(1).Visible = True
        GraphCaptionBox(2).Visible = True
        PVGraph3DBoarder(0).Visible = True
    End If
                 
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG120:" & Error$
End Sub

Sub ToggleReportFlag(ByVal Index As Integer)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCN3402
'Name    : ToggleReportFlag
'Created : 14 April 2005,
'Updated :
'Prg By  : Antony van Iersel
'Parm    : Index, which Report Flag to turn of or on.
'Desc    : If the report flag is false, Circle with line thru it
'          it will make it true, no circle with line thru it,
'          and visa versa
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

On Error GoTo Err_Handler
    Dim i As Integer
    Dim GraphIndex As Integer
    
    
    If ReportFlagImage(Index).Tag = False Then Exit Sub
    
    For i = 0 To 5
        ReportFlagImage(i).Picture = LoadResPicture(112, 0)
        ReportFlagImage(i).Tag = True
    Next i
    
    If Index = 0 Then
        ReportFlagImage(0).Picture = LoadResPicture(113, 0)
        ReportFlagImage(0).Tag = False
        Select Case ImageGraphState(0).GraphType
            Case "": GraphIndex = 1
            Case "Capacity": GraphIndex = 1
            Case "Ovality": GraphIndex = 2
            Case "Delta": GraphIndex = 3
            Case "XYDiameter": GraphIndex = 4
            Case "Flat": GraphIndex = 5
        End Select
        ReportFlagImage(GraphIndex).Picture = LoadResPicture(113, 0)
        ReportFlagImage(GraphIndex).Tag = False
    Else
        ReportFlagImage(Index).Picture = LoadResPicture(113, 0)
        ReportFlagImage(Index).Tag = False
        Call SetGeneralReportFlag
    End If
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG121:" & Error$
End Sub

Sub SetGeneralReportFlag()
On Error GoTo Err_Handler
    Dim SetGeneralGraph As Boolean
    
    SetGeneralGraph = True

    If ImageGraphState(0).GraphType = "" And ReportFlagImage(1).Tag = False Then SetGeneralGraph = False
    If ImageGraphState(0).GraphType = "Capacity" And ReportFlagImage(1).Tag = False Then SetGeneralGraph = False
    If ImageGraphState(0).GraphType = "Ovality" And ReportFlagImage(2).Tag = False Then SetGeneralGraph = False
    If ImageGraphState(0).GraphType = "Delta" And ReportFlagImage(3).Tag = False Then SetGeneralGraph = False
    If ImageGraphState(0).GraphType = "XYDiameter" And ReportFlagImage(4).Tag = False Then SetGeneralGraph = False
    If ImageGraphState(0).GraphType = "Flat" And ReportFlagImage(5).Tag = False Then SetGeneralGraph = False
    If SetGeneralGraph = False Then
        ReportFlagImage(0).Picture = LoadResPicture(113, 0)
        ReportFlagImage(0).Tag = False
    Else
        ReportFlagImage(0).Picture = LoadResPicture(112, 0)
        ReportFlagImage(0).Tag = True
    End If
    
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG122:" & Error$
End Sub


Function HistogramDiameterGraphSetup(Configuration As String)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : HistogramOfDiameterGraph_MouseMove
'Created : 16 April 2005, PCN3489
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim HistDiamOriginalHeight As Long
Dim HistDiamOriginalTop As Long
Dim HistDiamExpandedHeight As Long
Dim HistDiamExpandedTop As Long

HistDiamOriginalHeight = 1570
HistDiamOriginalTop = ScaleSetting.Top + ScaleSetting.height + 10
HistDiamExpandedHeight = 10000
HistDiamExpandedTop = 0
HistogramOfDiameter.ZOrder 0

Select Case Configuration
    Case "Normal"
        If HistogramOfDiameter.Top <> HistDiamOriginalTop Then
            HistogramOfDiameter.Visible = True
            DiameterHistogramSubTitle.Visible = True
            HistogramOfDiameter.Top = HistDiamOriginalTop
            HistogramOfDiameter.height = HistDiamOriginalHeight
'            HistogramOfDiameter.width = 15000 - PVToolbar.width
            HistogramOfDiameter.Left = 10
            HistogramOfDiameterGraph.Top = -525
            HistogramOfDiameterGraph.Left = -100
        End If
    Case "Expanded"
        If HistogramOfDiameter.Top <> HistDiamExpandedTop Then
            HistogramOfDiameter.Visible = True
            HistogramOfDiameter.Top = HistDiamExpandedTop
            HistogramOfDiameter.height = HistDiamExpandedHeight
            HistogramOfDiameter.Left = 0
            HistogramOfDiameter.width = 15050
            HistogramOfDiameterGraph.Top = 10
            HistogramOfDiameterGraph.Left = 10
            DiameterHistogramSubTitle.Visible = False
        End If
    Case "Open"
        HistogramOfDiameter.Visible = True
        DiameterHistogramSubTitle.Visible = True
        HistogramOfDiameter.Top = HistDiamOriginalTop
        HistogramOfDiameter.height = HistDiamOriginalHeight
'        HistogramOfDiameter.width = 15000 - PVToolbar.width
        HistogramOfDiameter.Left = 10
        HistogramOfDiameterGraph.Top = -525
        HistogramOfDiameterGraph.Left = -100
    Case Else
        HistogramOfDiameter.Visible = False
    
End Select
    





Exit Function
Err_Handler:
    MsgBox Err & "-PVG123:" & Error$
End Function


Sub ShowPVYScaleSelectorBox()
On Error GoTo Err_Handler


PVYScaleSelectorBox.Visible = True
Select Case PVGraphYRatio
    Case 2: Call HighlightPVYScaleSelector(0)
    Case 5: Call HighlightPVYScaleSelector(1)
    Case 10: Call HighlightPVYScaleSelector(2)
    Case 25: Call HighlightPVYScaleSelector(3)
    Case 50: Call HighlightPVYScaleSelector(4)
End Select

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG124:" & Error$
End Sub
Sub SetPVYScaleSetting(Index)
On Error GoTo Err_Handler
Select Case Index
    Case 0: PVGraphYRatio = 2
    Case 1: PVGraphYRatio = 5
    Case 2: PVGraphYRatio = 10
    Case 3: PVGraphYRatio = 25
    Case 4: PVGraphYRatio = 50
    Case 5: PVGraphYRatio = 1
End Select
    
    DrawPVGraphs
    PipeDisplayMoveLastY = ViewIndicator.Item(0).Top
    Call PipeDisplay_MouseMove(1, 0, 25, 25) 'Move viewindicator to top
    Call ReDimensionIndicators
    
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG125:" & Error$


End Sub

Sub SetPVGraphSelectSetting(Index As Integer, Container As Integer)
On Error GoTo Err_Handler
'Select Case Index
    'vvvv PCN4171 ****************************************************************
'    Case 0: Call ScreenDrawing.GraphSelect("Capacity", Container)
'    Case 1: Call ScreenDrawing.GraphSelect("Ovality", Container)
'    Case 2: Call ScreenDrawing.GraphSelect("MaxDiameter", Container)
'    Case 3: Call ScreenDrawing.GraphSelect("XYDiameter", Container)
'    Case 4: Call ScreenDrawing.GraphSelect("MaxMinDiameter", Container)
'    Case 5: Call ScreenDrawing.GraphSelect("MedianDiameter", Container)
'    Case 6: Call ScreenDrawing.GraphSelect("Flat", Container)

    Call ScreenDrawing.GraphSelect(PVGraphSelect(Index).Tag, Container) 'PCN4333

'    Case 0: Call ScreenDrawing.GraphSelect("Flat", Container)
'    Case 1: Call ScreenDrawing.GraphSelect("MedianDiameter", Container)
'    Case 2: Call ScreenDrawing.GraphSelect("Ovality", Container)
'    Case 3: Call ScreenDrawing.GraphSelect("MaxDiameter", Container)
'    Case 4: Call ScreenDrawing.GraphSelect("XYDiameter", Container)
'    Case 5: Call ScreenDrawing.GraphSelect("Capacity", Container)
    

    '^^^^ ************************************************************************
'End Select
    
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG126:" & Error$


End Sub

Sub HighlightPVYScaleSelector(ByVal Index As Integer)
On Error GoTo Err_Handler

Dim ColourNonHighlight As Long
Dim ColourHighlight As Long

Dim i As Integer

ColourNonHighlight = &H80C0FF
ColourHighlight = &H80FFFF

For i = 0 To 5
    PVYScaleSetting(i).BackColor = ColourNonHighlight
Next i

PVYScaleSetting(Index).BackColor = ColourHighlight
    
    
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG127:" & Error$

End Sub

Sub HighlightPVGraphSelect(ByVal Index As Integer)
On Error GoTo Err_Handler

Dim ColourNonHighlight As Long
Dim ColourHighlight As Long

Dim i As Integer

ColourNonHighlight = &H80C0FF
ColourHighlight = &H80FFFF

For i = 0 To 7 'PCN6128 from 6 to 7
    PVGraphSelect(i).BackColor = ColourNonHighlight
Next i

PVGraphSelect(Index).BackColor = ColourHighlight
    
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG128:" & Error$
End Sub

Function GetContainerGraphLabel(ByVal GraphInfoIndex As Integer) As String
On Error GoTo Err_Handler

Dim GraphType As String
Dim ReturnType As String

GraphType = GraphInfoContainer(GraphInfoIndex).GraphType
'vvvv PCN4171 ***************************************************
Select Case GraphType
    Case "Ovality": ReturnType = DisplayMessage("Ovality")
    Case "Capacity": ReturnType = DisplayMessage("Capacity")
    Case "MaxDiameter": ReturnType = DisplayMessage("Max Diameter")
    Case "MinDiameter": ReturnType = DisplayMessage("Min Diameter")
    Case "XYDiameter": If MedianFlat And PVDFileName <> "" Then ReturnType = DisplayMessage("XY Deflection") Else ReturnType = DisplayMessage("XY Diameter") 'PCN5186 added Deflection label
    Case "Flat": If MedianFlat And PVDFileName <> "" Then ReturnType = DisplayMessage("Deflection Flat") Else ReturnType = DisplayMessage("Flat")  'PCN4974
    Case "MaxMinDiameter": ReturnType = DisplayMessage("Diameter (Max/Min)")
    Case "MedianDiameter": ReturnType = DisplayMessage("Median Diameter")
    Case "Debris": ReturnType = DisplayMessage("Debris") 'PCN4461
'PCN6458     Case "Inclination": ReturnType = DisplayMessage("Inclination")

    Case Else: ReturnType = ""
End Select
'^^^^ ***********************************************************
 
GetContainerGraphLabel = ReturnType

Exit Function
Err_Handler:
    MsgBox Err & "-PVG129:" & Error$
End Function


Function GetCurrentGraphLabel(ByVal Index As Integer) As String
On Error GoTo Err_Handler

Dim GraphType As String
Dim ReturnType As String

GraphType = ImageGraphState(Index).GraphType
'vvvv PCN4171 ***************************************************
Select Case GraphType
    Case "Ovality": ReturnType = DisplayMessage("Ovality")
    Case "Capacity": ReturnType = DisplayMessage("Capacity")
    Case "MaxDiameter": ReturnType = DisplayMessage("Max Diameter")
    Case "XYDiameter": If MedianFlat And PVDFileName <> "" Then ReturnType = DisplayMessage("XY Deflection") Else ReturnType = DisplayMessage("XY Diameter") 'PCN5186 added Deflection label
    Case "Flat": If MedianFlat And PVDFileName <> "" Then ReturnType = DisplayMessage("Deflection Flat") Else ReturnType = DisplayMessage("Flat")  'PCN4974
    Case "MaxMinDiameter": ReturnType = DisplayMessage("Diameter (Max/Min)")
    Case "MedianDiameter": ReturnType = DisplayMessage("Median Diameter")
    Case "MinDiameter": ReturnType = DisplayMessage("Min Diameter") 'PCN4333
'PCN6458     Case "Inclination": ReturnType = DisplayMessage("Inclination") 'PCN6128

    Case Else: ReturnType = ""
End Select
'^^^^ ***********************************************************
GetCurrentGraphLabel = ReturnType

Exit Function
Err_Handler:
    MsgBox Err & "-PVG130:" & Error$
End Function

Sub UpdateGraphTitles()
On Error GoTo Err_Handler

Dim i As Integer

For i = 0 To 5
    Label_GraphName(i).Caption = GetCurrentGraphLabel(i)
    Label_GraphNameShadow(i).Caption = Label_GraphName(i).Caption
Next i

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG131:" & Error$
End Sub

Sub ExpandPVGraphSelectBox(Index As Integer)
On Error GoTo Err_Handler

Dim i As Integer
Dim Inc As Integer

For i = 0 To 7 'PCN6128 from 6 to 7
    If PVGraphSelect(i).Caption = Label_GraphName(Index).Caption Then
        PVGraphSelect(i).BackColor = &H80FFFF
    Else
        PVGraphSelect(i).BackColor = &H80C0FF
    End If
Next i


PVGraphSelectBox.width = 2175
PVGraphSelectBox.Left = GraphCaptionBox(Index).Left + ((GraphCaptionBox(Index).width - PVGraphSelectBox.width) / 2)
PVGraphSelectBox.Top = GraphCaptionBox(Index).Top + GraphCaptionBox(Index).height

PVGraphSelectBox.ZOrder 0
PVGraphSelectBox.Visible = True
PVGraphSelectBox.Tag = Index
i = 0
Inc = 100
Do While i < 1935 'PCN6458 + 255 'PCN6128 add inclination option in drop down box
    PVGraphSelectBox.height = i
    DoEvents
    Sleep (10)
    i = i + Inc
Loop

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG132:" & Error$
End Sub

Sub CollapsePVGraphSelectBox()
On Error GoTo Err_Handler

Dim i As Integer
Dim Inc As Integer
If PVGraphSelectBox.Tag = -1 Then Exit Sub

PVGraphSelectBox.Tag = -1
i = 1935 'PCN6458 + 255 'PCN6128 add inclination option in drop down box
Inc = 100
Do While i > 0
    PVGraphSelectBox.height = i
        DoEvents
    Sleep (10)
    i = i - Inc
Loop

PVGraphSelectBox.Visible = False

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG133:" & Error$
End Sub

Sub BoundryCheckForCollapse(X As Single, Y As Single)
'****************************************************************************************
'Name    : BoundryCheckForCollapse
'Created : 10 August 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   :
'Desc    :
'          wise update.
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
    Static Count As Long
    
    
    
    If (X < 100) Or (X > (PVGraphSelectBox.width - 100) And Y > 100) Or (Y > (PVGraphSelectBox.height - 100)) Then
       Count = Count + 1
       
       Call CollapsePVGraphSelectBox
       
    End If

       
       
Exit Sub
Err_Handler:
    MsgBox Err & "-PVG134:" & Error$
End Sub

Sub CentralMoveRoutine()
'****************************************************************************************
'Name    : CentralMoveRoutine
'Created : 10 August 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   :
'Desc    : Any movement where appropriate is sent here for processing
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
    
    If PVGraphSelectBox.Visible = True Then CollapsePVGraphSelectBox
    LargeRuler.Visible = False
    LargeGreenRuler.Visible = False
    LargeRedRuler.Visible = False

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG135:" & Error$
End Sub

Function GetFrameFromDistance(ByVal Distance As Double)
On Error GoTo Err_Handler

    Dim Diff As Double
    Dim SmallestDiff As Double
    Dim SmallestDiffIndex As Long
    Dim Index As Long
    
    SmallestDiffIndex = 1
    If Trim(ConfigInfo.DistanceProcessMethod) <> "None" Then
        SmallestDiff = Abs(Distance - PVDistances(1))
        
        For Index = 2 To PVDataNoOfLines
            Diff = Abs(Distance - PVDistances(Index))
            If Diff < SmallestDiff Then
                SmallestDiff = Diff
                SmallestDiffIndex = Index
            End If
        Next Index
        GetFrameFromDistance = SmallestDiffIndex
    Else
        GetFrameFromDistance = -1
    End If
Exit Function
Err_Handler:
    MsgBox Err & "-PVG136:" & Error$
End Function

Sub MoveGraph(ByVal Frame As Long)
On Error GoTo Err_Handler
    
'Dim Start_X As Long  'PCN1916
Dim PipeMain_Offset As Long
'Dim End_X As Long
'Dim X_Top_Limit As Long
'Dim X_Bottom_Limit As Long  'PCNGL161202
'Dim NextTop As Long
'Dim NextBotton As Long 'PCNGL161202
'Dim PI_X_Scale As Double 'PCNGL010503-1
'Dim PV_Y As Double 'PCNGL161202
'Dim GraphSetIndex As Integer 'PCN2970

'Dim PV_Y_Adjusted As Long 'PCN2970
'Dim ObsIndex As Integer 'PCN3490
'Dim Position As Single
    
Dim PVPipeLength As Integer
Dim ViewIndicatorHeight As Integer
Dim IndicatorPosition As Integer

PipeMain_Offset = 40
PVPipeLength = PipeMain.height - (2 * PipeMain_Offset)
ViewIndicatorHeight = PVPipeLength / PVGraphYRatio
    

    If PVDFileName = "" Then Exit Sub
    If PVDataNoOfLines = 0 Then Exit Sub

'    Frame = Frame - ((GraphEndFrame - GraphStartFrame) / 2)
'
'    Position = CSng(Frame) / CSng(PVDataNoOfLines) * (PrecisionVisionGraph.PipeMain.height)
'    Position = Position
    Call CentralMoveRoutine

'    X_Top_Limit = Me.PipeMain.Top + PipeMain_Offset
'    X_Bottom_Limit = Me.PipeMain.Top + Me.PipeMain.height - PipeMain_Offset
'    'X_Step = Y - PipeDisplayMoveLastY
'    NextTop = Position
'    NextBotton = NextTop + Me.ViewIndicator.Item(0).height
'    If NextTop < X_Top_Limit Then
'        NextTop = X_Top_Limit
'    End If
'    If NextBotton > X_Bottom_Limit Then
'        NextTop = X_Bottom_Limit - Me.ViewIndicator.Item(0).height
'        NextBotton = X_Bottom_Limit
'        PipeDisplayMoveLastY = Position
'    End If

'        Me.ViewIndicator.Item(0).Top = NextTop
'        Me.ViewIndicator.Item(1).Top = NextTop
'        Me.ViewIndicator.Item(2).Top = NextTop
'        'Me.ViewIndicator.Item(3).Top = NextTop
'        Me.ViewIndicator.Item(4).Top = NextTop
'        Me.ViewIndicator.Item(5).Top = NextTop
'
'        'Re-draw pipe 3D shading and shadow
'        'If X_Step > 0 Then
'        '    Start_X = NextTop - X_Step
'        '    End_X = Me.ViewIndicator.Item(0).Top
'        'Else
'        '    Start_X = Me.ViewIndicator.Item(0).Top + Me.ViewIndicator.Item(0).height - (X_Step + 10)
'        '    End_X = Start_X + X_Step
'        'End If
'
'        'Determine what is portion of the PV graph is displayed in PVGWindow
'        PI_X_Scale = ViewIndicator.Item(0).height / 7460 '(PipeMain.height - 2 * PipeMain_Offset)
'        If PI_X_Scale = 0 Then PI_X_Scale = 0.000001 'PCNGL010503-1
'
'        PV_Y = (Me.ViewIndicator.Item(0).Top - MH_Shadow1.height - 35) / PI_X_Scale
'        GraphStartFrame = (PV_Y / PVGraphYRatio)
'        GraphEndFrame = (PV_Y / PVGraphYRatio) + (7500 / PVGraphYRatio)
        
        IndicatorPosition = PipeMain.Top + PipeMain_Offset
        If IndicatorPosition < (PipeMain.Top + PipeMain_Offset) Then IndicatorPosition = (PipeMain.Top + PipeMain_Offset)
        If IndicatorPosition > (PipeMain.Top + PipeMain.height - PipeMain_Offset - ViewIndicatorHeight) Then IndicatorPosition = (PipeMain.Top + PipeMain.height - PipeMain_Offset - ViewIndicatorHeight)
            
        Me.ViewIndicator.Item(0).Top = IndicatorPosition
        Me.ViewIndicator.Item(1).Top = IndicatorPosition
        Me.ViewIndicator.Item(2).Top = IndicatorPosition
        'Me.ViewIndicator.Item(3).Top = IndicatorPosition
        Me.ViewIndicator.Item(4).Top = IndicatorPosition
        Me.ViewIndicator.Item(5).Top = IndicatorPosition
            
        GraphStartFrame = ((IndicatorPosition - PipeMain.Top - PipeMain_Offset) / PVPipeLength) * PVDataNoOfLines
        GraphEndFrame = GraphStartFrame + (PVDataNoOfLines / PVGraphYRatio)
        
        Call DrawPVGraphs
        Call RepositionYScaleMarkers    'PCN3373
        
            
            
'        PV_Y_Adjusted = ((GraphSetIndex) * PVGraphHeightLimit) - PV_Y


        
        'PipeDisplayMoveLastY = Position
        
    Call AdjustRulerWidths

        


Exit Sub
Err_Handler:
    Select Case Err
        Case 6: Exit Sub
        Case Else: MsgBox Err & "-PVG137:" & Error$
    End Select
End Sub


Sub SetPVGraphY_Units()
On Error GoTo Err_Handler

If Trim(ConfigInfo.DistanceProcessMethod) = "None" Or ConfigInfo.DistanceStart = InvalidData Then
    Me.Y_Units.Visible = False
ElseIf ConfigInfo.Units = "mm" Then
    Me.Y_Units.Caption = "m"
    Me.Y_Units.Visible = True
Else
    Me.Y_Units.Caption = "ft"
    Me.Y_Units.Visible = True
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG138:" & Error$
End Sub

Sub MoreButtonHighlight(ButtonControl As Control, X As Single, Y As Single)
On Error GoTo Err_Handler

If X > 100 And X < 250 And Y > 100 And Y < 250 Then
    ButtonControl.Picture = Me.ButtonsSmallImages.ListImages("MoreHighlighted").Picture
Else
    ButtonControl.Picture = Me.ButtonsSmallImages.ListImages("More").Picture
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG139:" & Error$
End Sub

Private Sub XScaleLabel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Call MoreButtonHighlight(Me.XScaleLabel(Index), X, Y)

Exit Sub
Err_Handler:
    MsgBox Err & "-PVG140:" & Error$
End Sub

Private Sub AdjustIndicatorAndGraph()
'Richard
On Error GoTo Err_Handler

Dim PipeMain_Offset As Integer
Dim PVPipeLength As Integer
Dim IndicatorPosition As Integer

PipeMain_Offset = 40
PVPipeLength = PipeMain.height - (2 * PipeMain_Offset)

'If PVFrameNo < GraphStartFrame Then

    If PVFrameNo > ((PVDataNoOfLines / PVGraphYRatio) / 2) Then
        If PVFrameNo < PVDataNoOfLines - ((PVDataNoOfLines / PVGraphYRatio) / 2) Then
            Me.ViewIndicator.Item(0).Top = ((PVFrameNo / PVDataNoOfLines) * PVPipeLength) + PipeMain.Top + PipeMain_Offset - (Me.ViewIndicator.Item(0).height / 2)
            Me.ViewIndicator.Item(1).Top = Me.ViewIndicator.Item(0).Top
            Me.ViewIndicator.Item(2).Top = Me.ViewIndicator.Item(0).Top
            'Me.ViewIndicator.Item(3).Top = Me.ViewIndicator.Item(0).Top
            Me.ViewIndicator.Item(4).Top = Me.ViewIndicator.Item(0).Top
            Me.ViewIndicator.Item(5).Top = Me.ViewIndicator.Item(0).Top
        Else
            Me.ViewIndicator.Item(0).Top = PipeMain.Top + PipeMain.height - PipeMain_Offset - Me.ViewIndicator.Item(0).height
            Me.ViewIndicator.Item(1).Top = Me.ViewIndicator.Item(0).Top
            Me.ViewIndicator.Item(2).Top = Me.ViewIndicator.Item(0).Top
            'Me.ViewIndicator.Item(3).Top = Me.ViewIndicator.Item(0).Top
            Me.ViewIndicator.Item(4).Top = Me.ViewIndicator.Item(0).Top
            Me.ViewIndicator.Item(5).Top = Me.ViewIndicator.Item(0).Top
        End If
    Else
        Me.ViewIndicator.Item(0).Top = PipeMain.Top + PipeMain_Offset
        Me.ViewIndicator.Item(1).Top = Me.ViewIndicator.Item(0).Top
        Me.ViewIndicator.Item(2).Top = Me.ViewIndicator.Item(0).Top
        'Me.ViewIndicator.Item(3).Top = Me.ViewIndicator.Item(0).Top
        Me.ViewIndicator.Item(4).Top = Me.ViewIndicator.Item(0).Top
        Me.ViewIndicator.Item(5).Top = Me.ViewIndicator.Item(0).Top
    End If
'ElseIf PVFrameNo > GraphStartFrame Then
    
IndicatorPosition = Me.ViewIndicator.Item(0).Top

GraphStartFrame = ((IndicatorPosition - PipeMain.Top - PipeMain_Offset) / PVPipeLength) * PVDataNoOfLines
GraphEndFrame = GraphStartFrame + (PVDataNoOfLines / PVGraphYRatio)
'Richard
Call DrawPVGraphs
Call RepositionYScaleMarkers    'PCN3373
Call AdjustRulerWidths


Exit Sub
Err_Handler:
    MsgBox Err & "-PVG141:" & Error$
End Sub


