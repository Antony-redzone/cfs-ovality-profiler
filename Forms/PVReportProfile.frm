VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form PVReportProfile 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Profile Report"
   ClientHeight    =   16425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15240
   Icon            =   "PVReportProfile.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   16425
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox CmboPrinterList 
      Height          =   315
      Left            =   5160
      MousePointer    =   1  'Arrow
      TabIndex        =   52
      Text            =   "Select a Printer"
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton PrinterSettingsButton 
      Caption         =   "PrinterSettings"
      Height          =   615
      Left            =   120
      TabIndex        =   49
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComDlg.CommonDialog FloatingTextDialog 
      Left            =   120
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox PageFramePictureBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   15480
      Left            =   0
      ScaleHeight     =   15450
      ScaleWidth      =   14985
      TabIndex        =   4
      Top             =   600
      Width           =   15015
      Begin VB.PictureBox picReportPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   16000
         Left            =   120
         ScaleHeight     =   15975
         ScaleWidth      =   14385
         TabIndex        =   5
         Tag             =   "Paper"
         Top             =   120
         Width           =   14415
         Begin VB.TextBox UserTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2400
            MousePointer    =   3  'I-Beam
            TabIndex        =   53
            Top             =   945
            Width           =   6495
         End
         Begin VB.TextBox FloatingTextDefault 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8760
            TabIndex        =   48
            Text            =   "Default Text Setting"
            Top             =   120
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.TextBox FloatingText 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   8520
            TabIndex        =   47
            Top             =   120
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox CommentsTextBox 
            Appearance      =   0  'Flat
            Height          =   765
            Left            =   480
            MousePointer    =   3  'I-Beam
            MultiLine       =   -1  'True
            TabIndex        =   6
            Top             =   13890
            Width           =   10575
         End
         Begin VB.Line Line5 
            X1              =   11040
            X2              =   11040
            Y1              =   5880
            Y2              =   7320
         End
         Begin VB.Line KeyLine2 
            X1              =   13080
            X2              =   13080
            Y1              =   3120
            Y2              =   5280
         End
         Begin VB.Label DiameterLabel 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Ø"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   13200
            TabIndex        =   108
            Top             =   3120
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label PVKey_Flat3D_Value0_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   13200
            TabIndex        =   89
            Top             =   4920
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label PVKey_Flat3D_Value1_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   13200
            TabIndex        =   88
            Top             =   4680
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label PVKey_Flat3D_Value2_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   13200
            TabIndex        =   87
            Top             =   4440
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label PVKey_Flat3D_Value7_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   13200
            TabIndex        =   86
            Top             =   3480
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label PVKey_Flat3D_Value5_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   13200
            TabIndex        =   85
            Top             =   3960
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label PVKey_Flat3D_Value6_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   13200
            TabIndex        =   84
            Top             =   3720
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label PVKey_Flat3D_Value4_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   13200
            TabIndex        =   83
            Top             =   4200
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label PVKey_Flat3D_Value0_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   12360
            TabIndex        =   107
            Top             =   4920
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label PVKey_Flat3D_Value1_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   12360
            TabIndex        =   106
            Top             =   4680
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label PVKey_Flat3D_Value2_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   12360
            TabIndex        =   105
            Top             =   4440
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label PVKey_Flat3D_Value7_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   12360
            TabIndex        =   104
            Top             =   3480
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label PVKey_Flat3D_Value5_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   12360
            TabIndex        =   103
            Top             =   3960
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label PVKey_Flat3D_Value6_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   12360
            TabIndex        =   102
            Top             =   3720
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label PVKey_Flat3D_Value4_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   12360
            TabIndex        =   101
            Top             =   4200
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label PVKey_Flat3D_Value0 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   11640
            TabIndex        =   100
            Top             =   4920
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label PVKey_Flat3D_Value1 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   11640
            TabIndex        =   99
            Top             =   4680
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label PVKey_Flat3D_Value2 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   11640
            TabIndex        =   98
            Top             =   4440
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label PVKey_Flat3D_Value7 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   11640
            TabIndex        =   97
            Top             =   3480
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label PVKey_Flat3D_Value5 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   11640
            TabIndex        =   96
            Top             =   3960
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label PVKey_Flat3D_Value6 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   11640
            TabIndex        =   95
            Top             =   3720
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label PVKey_Flat3D_Value4 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   11640
            TabIndex        =   94
            Top             =   4200
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label RadiusLabel 
            BackColor       =   &H00FFFFFF&
            Caption         =   "r"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   11760
            TabIndex        =   93
            Top             =   3120
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Line KeyLine1 
            X1              =   12285
            X2              =   12285
            Y1              =   3120
            Y2              =   5280
         End
         Begin VB.Label PVKey_MinDiameter_Value_Unit 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0mm"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   8640
            TabIndex        =   92
            Top             =   8940
            Width           =   855
         End
         Begin VB.Label PVKey_MinDiameter_Value_Unit 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   8520
            TabIndex        =   91
            Top             =   9165
            Width           =   975
         End
         Begin VB.Image MinDiameterLabel 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   8280
            Picture         =   "PVReportProfile.frx":038A
            ToolTipText     =   "X"
            Top             =   9000
            Width           =   270
         End
         Begin VB.Label DiameterMinLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Min"
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
            Left            =   8280
            TabIndex        =   90
            Top             =   8670
            Width           =   1335
         End
         Begin VB.Line Line9 
            X1              =   8160
            X2              =   9600
            Y1              =   8670
            Y2              =   8670
         End
         Begin VB.Shape PVKey_Flat3D_Color7 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   270
            Left            =   11280
            Top             =   3480
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Shape PVKey_Flat3D_Color6 
            FillColor       =   &H000096FF&
            FillStyle       =   0  'Solid
            Height          =   270
            Left            =   11280
            Top             =   3720
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Shape PVKey_Flat3D_Color5 
            FillColor       =   &H0014FFFF&
            FillStyle       =   0  'Solid
            Height          =   270
            Left            =   11280
            Top             =   3960
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Shape PVKey_Flat3D_Color0 
            FillColor       =   &H006F4928&
            FillStyle       =   0  'Solid
            Height          =   270
            Left            =   11280
            Top             =   4920
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Shape PVKey_Flat3D_Color1 
            FillColor       =   &H00CC9B5A&
            FillStyle       =   0  'Solid
            Height          =   270
            Left            =   11280
            Top             =   4680
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Shape PVKey_Flat3D_Color2 
            FillColor       =   &H00EEE0B5&
            FillStyle       =   0  'Solid
            Height          =   270
            Left            =   11280
            Top             =   4440
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Shape KeyDataBoarder 
            Height          =   2175
            Left            =   11160
            Top             =   3120
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.Shape PVKey_Flat3D_Color4 
            Height          =   270
            Left            =   11280
            Top             =   4200
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Label FractileLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "90% - Fractile: "
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
            Left            =   480
            TabIndex        =   82
            Top             =   12990
            Width           =   1095
         End
         Begin VB.Label FractileLabelData 
            Caption         =   "xxxxxxxx"
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
            Left            =   1680
            TabIndex        =   81
            Top             =   12990
            Width           =   495
         End
         Begin VB.Label LimitValue1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "xxxx.xx"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   10440
            TabIndex        =   80
            Top             =   10680
            Width           =   615
         End
         Begin VB.Label LimitValue2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "xxxx.xx"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   10440
            TabIndex        =   79
            Top             =   11640
            Width           =   615
         End
         Begin VB.Image ControlsPositionShunt 
            Height          =   16815
            Index           =   2
            Left            =   -120
            Tag             =   "Container"
            Top             =   -240
            Width           =   11295
         End
         Begin VB.Image ControlsPositionShunt 
            Height          =   5175
            Index           =   1
            Left            =   0
            Tag             =   "Container"
            Top             =   9840
            Width           =   11175
         End
         Begin VB.Image ControlsPositionShunt 
            Height          =   6855
            Index           =   0
            Left            =   480
            Tag             =   "Container"
            Top             =   3000
            Width           =   12735
         End
         Begin VB.Image PVYScaleImage 
            Height          =   5325
            Left            =   600
            Top             =   3900
            Width           =   375
         End
         Begin VB.Image PVXScaleImage 
            Height          =   375
            Left            =   960
            Top             =   9210
            Width           =   7095
         End
         Begin VB.Image ImageDimension 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   9720
            Picture         =   "PVReportProfile.frx":06CC
            Top             =   6240
            Width           =   270
         End
         Begin VB.Label PVKey_Capacity_Value_Unit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "0%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   8490
            TabIndex        =   78
            Top             =   7785
            Width           =   900
         End
         Begin VB.Label PVKey_YDiameter_Value_Unit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "0%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   8520
            TabIndex        =   77
            Top             =   7095
            Width           =   900
         End
         Begin VB.Label PVKey_XDiameter_Value_Unit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "0%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   8520
            TabIndex        =   76
            Top             =   6420
            Width           =   900
         End
         Begin VB.Label PVKey_MaxDiameter_Value_Unit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "0%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   8520
            TabIndex        =   75
            Top             =   8475
            Width           =   900
         End
         Begin VB.Image PVKey_MaxDiameter_Icon 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   8280
            Picture         =   "PVReportProfile.frx":0C56
            ToolTipText     =   "X"
            Top             =   8295
            Width           =   270
         End
         Begin VB.Image PVKey_YDiameter_Icon 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   8280
            Picture         =   "PVReportProfile.frx":0F98
            ToolTipText     =   "Y"
            Top             =   6945
            Width           =   270
         End
         Begin VB.Label DimensionValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "0mm"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   9840
            TabIndex        =   74
            Top             =   6240
            Width           =   1020
         End
         Begin VB.Shape UperLimitData 
            Height          =   915
            Left            =   8160
            Top             =   3300
            Width           =   1575
         End
         Begin VB.Label LimitLinesLabel 
            Caption         =   "Limit Lines"
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
            Left            =   8280
            TabIndex        =   73
            Top             =   3360
            Width           =   855
         End
         Begin VB.Label UpperLimitLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "Upper ="
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
            Left            =   8520
            TabIndex        =   72
            Top             =   3720
            Width           =   615
         End
         Begin VB.Label LowerLimitLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "Lower ="
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
            Left            =   8520
            TabIndex        =   71
            Top             =   3960
            Width           =   615
         End
         Begin VB.Label UpperLimitData 
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
            Left            =   9240
            TabIndex        =   70
            Top             =   3720
            Width           =   375
         End
         Begin VB.Label LowerLimitData 
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
            Left            =   9240
            TabIndex        =   69
            Top             =   3960
            Width           =   375
         End
         Begin VB.Line LowerLimitLine 
            X1              =   8280
            X2              =   8520
            Y1              =   4080
            Y2              =   4080
         End
         Begin VB.Label PVKey_MaxDiameter_Value_Unit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "0mm"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   8520
            TabIndex        =   68
            Top             =   8250
            Width           =   900
         End
         Begin VB.Label PVKey_XDiameter_Value_Unit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "0mm"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   8520
            TabIndex        =   67
            Top             =   6195
            Width           =   900
         End
         Begin VB.Label PVKey_Capacity_Value_Unit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "0cm"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   8520
            TabIndex        =   66
            Top             =   7560
            Width           =   900
         End
         Begin VB.Image PVKey_XDiameter_Icon 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   8280
            Picture         =   "PVReportProfile.frx":12DA
            ToolTipText     =   "X"
            Top             =   6270
            Width           =   270
         End
         Begin VB.Image PVKey_Ovality_Icon 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   9720
            Picture         =   "PVReportProfile.frx":161C
            Top             =   6960
            Width           =   270
         End
         Begin VB.Image PVKey_Capacity_Icon 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   8280
            Picture         =   "PVReportProfile.frx":195E
            Top             =   7620
            Width           =   270
         End
         Begin VB.Line UpperLimitLine 
            X1              =   8280
            X2              =   8520
            Y1              =   3840
            Y2              =   3840
         End
         Begin VB.Shape Shape3 
            Height          =   3495
            Left            =   8160
            Top             =   5880
            Width           =   1455
         End
         Begin VB.Label UnitSquare 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   9420
            TabIndex        =   65
            Top             =   7440
            Width           =   255
         End
         Begin VB.Label PVKey_Distance_Icon 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            ForeColor       =   &H00004000&
            Height          =   270
            Left            =   9960
            TabIndex        =   64
            Top             =   9300
            Width           =   270
         End
         Begin VB.Label PVKey_Distance_Value 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "0m"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   10080
            TabIndex        =   63
            Top             =   9300
            Width           =   930
         End
         Begin VB.Label PVKey_Ovality_Value 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "0%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   9960
            TabIndex        =   62
            Top             =   6960
            Width           =   930
         End
         Begin VB.Line Line1 
            X1              =   11040
            X2              =   9600
            Y1              =   5880
            Y2              =   5880
         End
         Begin VB.Line Line3 
            X1              =   11040
            X2              =   8160
            Y1              =   6615
            Y2              =   6615
         End
         Begin VB.Line Line4 
            X1              =   11040
            X2              =   8160
            Y1              =   7305
            Y2              =   7305
         End
         Begin VB.Line Line6 
            X1              =   9600
            X2              =   8160
            Y1              =   7980
            Y2              =   7980
         End
         Begin VB.Label DimensionLabel 
            Caption         =   "True"
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
            Left            =   9720
            TabIndex        =   61
            Top             =   5910
            Width           =   1335
         End
         Begin VB.Label YDiameterLabel 
            Caption         =   "Vertical"
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
            Left            =   8280
            TabIndex        =   60
            Top             =   6615
            Width           =   1335
         End
         Begin VB.Label PVKey_YDiameter_Value_Unit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "0mm"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   8520
            TabIndex        =   59
            Top             =   6870
            Width           =   900
         End
         Begin VB.Label OvalityLabel 
            Caption         =   "Ovality"
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
            Left            =   9720
            TabIndex        =   58
            Top             =   6600
            Width           =   1335
         End
         Begin VB.Label DiameterMaxLabel 
            Caption         =   "Max"
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
            Left            =   8280
            TabIndex        =   57
            Top             =   7980
            Width           =   1335
         End
         Begin VB.Label XDiameterLabel 
            Caption         =   "Horizontal"
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
            Left            =   8280
            TabIndex        =   56
            Top             =   5925
            Width           =   1335
         End
         Begin VB.Label CapacityLabel 
            Caption         =   "Capacity"
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
            Left            =   8280
            TabIndex        =   55
            Top             =   7305
            Width           =   1335
         End
         Begin VB.Label ObservationsLabel 
            BackStyle       =   0  'Transparent
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
            Left            =   600
            TabIndex        =   54
            Top             =   3960
            Width           =   7455
         End
         Begin VB.Image PVProfileImage 
            Height          =   5325
            Left            =   960
            Stretch         =   -1  'True
            Top             =   3900
            Width           =   7095
         End
         Begin VB.Line Line7 
            X1              =   10800
            X2              =   10800
            Y1              =   9480
            Y2              =   9720
         End
         Begin VB.Label CompanyNameLabel 
            Caption         =   "Co"
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
            Left            =   2880
            TabIndex        =   51
            Top             =   14880
            Width           =   5055
         End
         Begin VB.Label GraphUnitLbl 
            BackStyle       =   0  'Transparent
            Caption         =   "m"
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
            Left            =   10800
            TabIndex        =   50
            Top             =   12510
            Width           =   255
         End
         Begin VB.Label UnitLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "unit"
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
            Left            =   360
            TabIndex        =   46
            Top             =   10320
            Width           =   495
         End
         Begin VB.Label YScaleLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "xx.xx"
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
            Index           =   1
            Left            =   360
            TabIndex        =   45
            Top             =   10620
            Width           =   600
         End
         Begin VB.Line ProfileSliceRubberBand 
            X1              =   1680
            X2              =   10800
            Y1              =   9960
            Y2              =   9720
         End
         Begin VB.Line ProfileSlice 
            X1              =   1680
            X2              =   1680
            Y1              =   12720
            Y2              =   9960
         End
         Begin VB.Label YScaleLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "xx.xx"
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
            Index           =   3
            Left            =   360
            TabIndex        =   44
            Top             =   11700
            Width           =   600
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
         Begin VB.Image CLPLogoImage 
            Height          =   705
            Left            =   600
            Picture         =   "PVReportProfile.frx":1CA0
            Stretch         =   -1  'True
            Top             =   14880
            Width           =   2010
         End
         Begin VB.Image LogoImage 
            Height          =   690
            Left            =   360
            Picture         =   "PVReportProfile.frx":2505
            Stretch         =   -1  'True
            Top             =   480
            Width           =   2010
         End
         Begin VB.Label lblTitle 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "Capacity Profile Report"
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
            Left            =   240
            TabIndex        =   43
            Top             =   480
            Width           =   10815
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
            Left            =   9120
            TabIndex        =   42
            Top             =   15120
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
            Left            =   10080
            TabIndex        =   41
            Top             =   15120
            Width           =   855
         End
         Begin VB.Line FooterBreakLine 
            X1              =   480
            X2              =   11040
            Y1              =   14760
            Y2              =   14760
         End
         Begin VB.Line HeaderBreakLine 
            X1              =   480
            X2              =   11040
            Y1              =   1500
            Y2              =   1500
         End
         Begin VB.Label StartNodeLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "Start Node:"
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
            Left            =   480
            TabIndex        =   40
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label StartLocationLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "Start Location:"
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
            Left            =   480
            TabIndex        =   39
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label AssetNoLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "Asset No:"
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
            Left            =   4320
            TabIndex        =   38
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label FinishNodeLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "Finish Node:"
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
            Left            =   4320
            TabIndex        =   37
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label FinishLocationLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "Finish Location:"
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
            Left            =   4320
            TabIndex        =   36
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label PipeMaterialLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "Pipe Material:"
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
            Left            =   8280
            TabIndex        =   35
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label CityLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "City:"
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
            Left            =   480
            TabIndex        =   34
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label PipeLengthLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "Pipe Length:"
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
            Left            =   8400
            TabIndex        =   33
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label PipeDiameterLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "Pipe Diameter:"
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
            Left            =   8400
            TabIndex        =   32
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label DateLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "Date:"
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
            Left            =   8880
            TabIndex        =   31
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label SiteIDLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "Site ID:"
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
            Left            =   480
            TabIndex        =   30
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label SiteIDData 
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
            Left            =   1800
            TabIndex        =   29
            Top             =   1560
            Width           =   2535
         End
         Begin VB.Label StartNodeData 
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
            Left            =   1800
            TabIndex        =   28
            Top             =   2040
            Width           =   2535
         End
         Begin VB.Label StartLocationData 
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
            Left            =   1800
            TabIndex        =   27
            Top             =   2280
            Width           =   2535
         End
         Begin VB.Label AssetNoData 
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
            Left            =   5880
            TabIndex        =   26
            Top             =   1560
            Width           =   3015
         End
         Begin VB.Label FinishNodeData 
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
            Left            =   5880
            TabIndex        =   25
            Top             =   2040
            Width           =   2535
         End
         Begin VB.Label FinishLocationData 
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
            Left            =   5880
            TabIndex        =   24
            Top             =   2280
            Width           =   2535
         End
         Begin VB.Label PipeMaterialData 
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
            Left            =   9840
            TabIndex        =   23
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label CityData 
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
            Left            =   1800
            TabIndex        =   22
            Top             =   1800
            Width           =   4935
         End
         Begin VB.Label PipeLengthData 
            Alignment       =   1  'Right Justify
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
            Left            =   9840
            TabIndex        =   21
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label PipeDiameterData 
            Alignment       =   1  'Right Justify
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
            Left            =   9840
            TabIndex        =   20
            Top             =   2280
            Width           =   615
         End
         Begin VB.Label DateLabelData 
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
            Left            =   9840
            TabIndex        =   19
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label CleanFlowSystemWebLabel 
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
            Left            =   4800
            TabIndex        =   18
            Top             =   15120
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
            Left            =   2760
            TabIndex        =   17
            Top             =   15120
            Width           =   375
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
            Left            =   3120
            TabIndex        =   16
            Top             =   15120
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
            Left            =   10080
            TabIndex        =   15
            Top             =   14880
            Width           =   855
         End
         Begin VB.Label CommentsLabel 
            Caption         =   "Comments"
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
            Left            =   480
            TabIndex        =   14
            Top             =   13200
            Width           =   1215
         End
         Begin VB.Label PipeLengthUnitLabel 
            Caption         =   "m"
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
            Left            =   10560
            TabIndex        =   13
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label PipeDiameterUnitLabel 
            Caption         =   "mm"
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
            Left            =   10560
            TabIndex        =   12
            Top             =   2280
            Width           =   495
         End
         Begin VB.Label CopyrightLabel 
            Caption         =   "Copyright 2006"
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
            Left            =   8040
            TabIndex        =   11
            Top             =   14880
            Width           =   1455
         End
         Begin VB.Label Explination 
            Caption         =   "Ovality 'q' (as per ASTM F 1216 Standard Practice) as a percentage of original pipe versus distance"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   4200
            TabIndex        =   10
            Top             =   12960
            Width           =   6975
            WordWrap        =   -1  'True
         End
         Begin VB.Image GraphContainer 
            Appearance      =   0  'Flat
            Height          =   2175
            Left            =   960
            Stretch         =   -1  'True
            Top             =   10200
            Width           =   9495
         End
         Begin VB.Image GraphXScaleContainer 
            Height          =   375
            Left            =   960
            Top             =   12360
            Width           =   9495
         End
         Begin VB.Label YScaleLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "xx.xx"
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
            Index           =   2
            Left            =   360
            TabIndex        =   9
            Top             =   11160
            Width           =   600
         End
         Begin VB.Label YScaleLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "xx.xx"
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
            Index           =   4
            Left            =   360
            TabIndex        =   8
            Top             =   12240
            Width           =   600
         End
         Begin VB.Label YScaleLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "xx.xx"
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
            Index           =   0
            Left            =   360
            TabIndex        =   7
            Top             =   10080
            Width           =   600
         End
         Begin VB.Shape Shape2 
            Height          =   2175
            Left            =   960
            Top             =   10200
            Width           =   9495
         End
         Begin VB.Shape GraphBackgroundShape 
            BackColor       =   &H00DDDDA2&
            FillColor       =   &H00DDDDA2&
            FillStyle       =   0  'Solid
            Height          =   2775
            Left            =   480
            Tag             =   "Back"
            Top             =   10080
            Width           =   10575
         End
         Begin VB.Shape Shape4 
            FillColor       =   &H00FFFFC0&
            FillStyle       =   0  'Solid
            Height          =   255
            Left            =   9720
            Tag             =   "Back"
            Top             =   9300
            Width           =   255
         End
         Begin VB.Shape PVProfileGraphBoarder 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   6345
            Left            =   600
            Tag             =   "Back"
            Top             =   3240
            Width           =   7455
         End
      End
   End
   Begin VB.CommandButton ScaleButton20 
      Caption         =   "2"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton ScaleButton15 
      Caption         =   "1.5"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton ScaleButton10 
      Caption         =   "1"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton ScaleButton05 
      Caption         =   "0.5"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComDlg.CommonDialog PrinterDialogBox 
      Left            =   120
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Printer Settings"
   End
   Begin VB.Image CloseReport 
      Height          =   480
      Left            =   12840
      MousePointer    =   1  'Arrow
      Picture         =   "PVReportProfile.frx":2D6A
      ToolTipText     =   "Close Report"
      Top             =   90
      Width           =   480
   End
   Begin VB.Image ControlsReport 
      Height          =   720
      Index           =   5
      Left            =   3480
      MousePointer    =   1  'Arrow
      Picture         =   "PVReportProfile.frx":3A34
      Tag             =   "SaveReportToPVD"
      ToolTipText     =   "Save Report To PVD"
      Top             =   -60
      Width           =   720
   End
   Begin VB.Image ControlsReport 
      Height          =   720
      Index           =   4
      Left            =   2640
      MousePointer    =   1  'Arrow
      Picture         =   "PVReportProfile.frx":56FE
      Tag             =   "ZoomOut"
      ToolTipText     =   "Zoom Out"
      Top             =   -60
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsReport 
      Height          =   720
      Index           =   3
      Left            =   4320
      MousePointer    =   1  'Arrow
      Picture         =   "PVReportProfile.frx":73C8
      Tag             =   "Print"
      ToolTipText     =   "Print"
      Top             =   -60
      Width           =   720
   End
   Begin VB.Image ControlsReport 
      Height          =   720
      Index           =   2
      Left            =   1800
      MousePointer    =   1  'Arrow
      Picture         =   "PVReportProfile.frx":9092
      Tag             =   "ZoomIn"
      ToolTipText     =   "Zoom In"
      Top             =   -60
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsReport 
      Height          =   720
      Index           =   1
      Left            =   960
      MousePointer    =   1  'Arrow
      Picture         =   "PVReportProfile.frx":AD5C
      Tag             =   "Move"
      ToolTipText     =   "Move Report or Text"
      Top             =   -60
      Width           =   720
   End
   Begin VB.Image ControlsReport 
      Height          =   720
      Index           =   0
      Left            =   120
      MousePointer    =   1  'Arrow
      Picture         =   "PVReportProfile.frx":CA26
      Tag             =   "DrawText"
      ToolTipText     =   "Text"
      Top             =   -60
      Width           =   720
   End
   Begin VB.Image ControlHighlight 
      Height          =   570
      Left            =   10680
      Picture         =   "PVReportProfile.frx":E6F0
      Top             =   30
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Image ControlHighlightDepressed 
      Height          =   585
      Left            =   11760
      Picture         =   "PVReportProfile.frx":10282
      Top             =   30
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Image ControlsBackPanel 
      Height          =   615
      Left            =   0
      MousePointer    =   1  'Arrow
      Picture         =   "PVReportProfile.frx":11ECC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13335
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
         Caption         =   "Reset to default"
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
Attribute VB_Name = "PVReportProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim ReportMouseDown As Boolean
Public PreviewStartFrame As Long
Public PreviewEndFrame As Long
Dim ShuntY(2) As Single
Dim ShuntX(2) As Single
'Public RenderScale As Single
'Public PrintPreviewAction As String

Private Sub AllFramesButtons_Click()
On Error GoTo Err_Handler
    GraphStartFrame = 1
    GraphEndFrame = PVDataNoOfLines
    Call PositionReportControls
    Call FillOutPrintForm
    Call GraphSpecificSettings
      
    
    Call RenderForm
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RP1:" & Error$
    End Select
End Sub




Sub PrintPVProfileReport()
On Error GoTo Err_Handler
    Dim I As Integer
   

    RenderScale = 1

    Printer.Orientation = vbPRORPortrait
    Printer.PrintQuality = vbPRPQHigh
    PrinterDialogBox.Orientation = cdlPortrait

    ScreenDrawingType = 1
    ScreenDrawingOrientation = 1


    If ScreenDrawing.ImageGraphState(0).GraphType = "Ovality" Then
        Me.Explination.Tag = "Visible"
    End If

    Call DrawPVGraphsReport

    Call RenderToPrinter.RenderReport(Me, Printer, 1)
    Call RenderToPrinter.RenderSingleTextBox(Me.UserTitle, Printer, False) 'PCN4277
    Call RenderToPrinter.RenderSingleTextBox(Me.CommentsTextBox, Printer, True) 'PCN4412
         

    Set ScreenDrawing.ReportDummyGraphImage = PVProfileImage
    
    Set ScreenDrawing.ReportGraphImageX = PVReportProfile.PVXScaleImage
    Set ScreenDrawing.ReportGraphImageY = PVReportProfile.PVYScaleImage
    Call DrawMainScale(Printer)
    Call DrawPVGraphsReport
        For I = 1 To Me.FloatingText.Count - 1 'PCN4412
        Call RenderToPrinter.RenderSingleTextBox(Me.FloatingText(I), Printer, True)
    Next I



    Call Printer.EndDoc
    ScreenDrawingType = 0
    ScreenDrawingOrientation = 0
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RP2:" & Error$
    End Select
End Sub





Private Sub CloseReport_Click()
On Error GoTo Err_Handler

Unload Me
    
Exit Sub
Err_Handler:
    MsgBox Err & "-RP3:" & Error$
End Sub

Private Sub CommentsTextBox_Change()
On Error GoTo Err_Handler
Dim SaveFailed As Boolean

If PVDFileName <> "" Then
    PipelineInfo.Comments = PVReportProfile.CommentsTextBox.text
    Call SaveToFilePipeAndConfigInfo("PipelineInfo", SaveFailed)
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-RP4:" & Error$
End Sub


Private Sub ControlsBackPanel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


If Me.ControlHighlight.Visible Then Me.ControlHighlight.Visible = False
If Me.ControlHighlightDepressed.Visible Then Me.ControlHighlightDepressed.Visible = False
    
Exit Sub
Err_Handler:
    MsgBox Err & "-RP5:" & Error$
End Sub

Private Sub ControlsReport_Click(Index As Integer)
On Error GoTo Err_Handler

Call ExecuteReportButton(Me, Index)
    
Exit Sub
Err_Handler:
    MsgBox Err & "-RP6:" & Error$
End Sub

Private Sub ControlsReport_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


Me.ControlHighlightDepressed.Left = Me.ControlHighlight.Left
Me.ControlHighlightDepressed.Visible = True
Me.ControlHighlight.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-RP7:" & Error$
End Sub

Private Sub ControlsReport_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ControlHighlight.Left = Me.ControlsReport(Index).Left - 100

If Me.ControlHighlight.Visible = False Then Me.ControlHighlight.Visible = True

    
Exit Sub
Err_Handler:
    MsgBox Err & "-RP8:" & Error$
End Sub

Private Sub ControlsReport_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ControlHighlight.Visible = True
Me.ControlHighlightDepressed.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-RP9:" & Error$
End Sub


Private Sub FloatingText_Change(Index As Integer)
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingText_Change(Me, Index)
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RP10:" & Error$
    End Select
End Sub

Private Sub FloatingText_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingText_KeyPress(Me, Index, KeyAscii)
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RP11:" & Error$
    End Select
End Sub

Private Sub FloatingText_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingText_MouseDown(Me, Index, Button, Shift, X, Y)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RP12:" & Error$
    End Select
End Sub

Private Sub FloatingText_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingText_MouseMove(Me, Index, Button, Shift, X, Y)
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RP13:" & Error$
    End Select
End Sub

Private Sub FloatingTextBackgroundColourMenu_Click()
On Error GoTo Err_Handler
    
    Call RenderToPrinter.FloatingTextBackgroundColourMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RP14:" & Error$
    End Select
End Sub

Private Sub FloatingTextDefaultMenu_Click()
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingTextDefaultMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RP15:" & Error$
    End Select
End Sub

Private Sub FloatingTextDeleteAllMenu_Click()
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingTextDeleteAllMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RP16:" & Error$
    End Select
End Sub

Private Sub FloatingTextDeleteMenu_Click()
On Error GoTo Err_Handler
    
    Call RenderToPrinter.FloatingTextDeleteMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RP17:" & Error$
    End Select
End Sub

Private Sub FloatingTextFontMenu_Click()
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingTextFontMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RP18:" & Error$
    End Select
End Sub
Sub InitialiseForm()
On Error GoTo Err_Handler
    Dim I As Integer

    
   Set ImageGraphState(6).PictureImage = PrecisionVisionGraph.PrinterReportImage 'PCN4271
    
    Me.Left = 0
    Me.width = ClearLineProfilerV6.width - 200
    Me.Top = 0
    Me.height = ClearLineProfilerV6.height - 500
'    Me.Show
    
    I = Me.Controls.Count
    ReDim OriginalStateVisible(I)
    ReDim OriginalStateTag(I)
    ReDim OriginalStateLeft(I)
    ReDim OriginalStateTop(I)
    ReDim OriginalStateX1(I)
    ReDim OriginalStateY1(I)
    ReDim OriginalStateX2(I)
    ReDim OriginalStateY2(I)
    
    
    For I = 1 To Me.Controls.Count - 1
        OriginalStateVisible(I) = Me.Controls(I).Visible
        OriginalStateTag(I) = Me.Controls(I).Tag
        OriginalStateLeft(I) = Me.Controls(I).Left
        OriginalStateTop(I) = Me.Controls(I).Top
        OriginalStateX1(I) = Me.Controls(I).x1
        OriginalStateY1(I) = Me.Controls(I).y1
        OriginalStateX2(I) = Me.Controls(I).x2
        OriginalStateY2(I) = Me.Controls(I).y2
    Next I
    
    PageFramePictureBox.Left = 0
    PageFramePictureBox.width = Me.width - PageFramePictureBox.Left
    PageFramePictureBox.height = Me.height
    PageFramePictureBox.Top = 650
    Me.ControlsBackPanel.width = Me.width
    Me.CloseReport.Left = Me.width - 750
    
    Call ConvertLanguage(Me, Language) 'PCN4171
    Me.PageLabel.Caption = DisplayMessage("Page") & " 1/1"


'    RenderScale = 1
    Set PrintPreviewForm = Me
    
    PreviewStartFrame = GraphStartFrame
    PreviewEndFrame = GraphEndFrame
    
    Call PositionReportControls
    Call FillOutPrintForm
    Call MarkForPrinting
    
    Call GraphSpecificSettings
      
    
    Call RenderForm
    
    'Set mouse icon for move
    PrintPreviewAction = "MoveAll"
    Call SetupReportMouseIcon(Me, 108)
    Me.PageFramePictureBox.Visible = True


Exit Sub
Err_Handler:
    Select Case Err
        Case 438, 387, 393
: Resume Next
        Case Else: MsgBox Err & "-RP19:" & Error$
    End Select
    'Resume 'ANT
End Sub



Private Sub Form_Load()
On Error GoTo Err_Handler
'    If CLPScreenMode = ThreeD Then
'            Call ClearLineScreen.D3D_CaptureImage(LocToSave & "Snapshot.bmp")
'            PVReportProfile.PVProfileImage.Picture = LoadPicture(LocToSave & "Snapshot.bmp")
'    End If

    Me.UserTitle.Font.Charset = LanguageCharset
    Me.FloatingText(0).Font.Charset = LanguageCharset
    Me.CommentsLabel.Font.Charset = LanguageCharset
    Me.CommentsTextBox.Font.Charset = LanguageCharset
    
    
    Me.CmboPrinterList.Enabled = False
    Call GetPipeDetailsLabels(Me)
'    Call ClearLineScreen.SnapShotForEmbededFile 'Moved to report button, because we want to be able to
                                                 'load image from observations as well
                                                 
    'PVReportProfile.PVProfileImage.Picture = LoadPicture(WindowsTempDirectory & "CBS\EmbedFile.jpg")
    PVReportProfile.PVProfileImage.Picture = LoadPicture(WindowsTempDirectory & "EmbedFile.jpg") 'ID4601

    Call GetPrinterList(Me)
    Call Me.InitialiseForm
    Me.CmboPrinterList.Enabled = True
    PipelineDetails.GeneralComments.text = Trim(PipelineInfo.Comments)
    
    'Me.UserTitle.text = UserTitleProfile 'PCN4433
    
Exit Sub
Err_Handler:
    Select Case Err
        Case 53: Resume Next 'File Not Found
        Case Else: MsgBox Err & "-RP20:" & Error$
        
    End Select
    
End Sub

Sub RenderForm()
On Error GoTo Err_Handler
    picReportPage.Cls
    picReportPage.width = Printer.width * RenderScale
    picReportPage.height = Printer.height * RenderScale
    
    PreviewStartFrame = GraphStartFrame
    PreviewEndFrame = GraphEndFrame

'    PVGraphOvalityXScale = 8
'    PVGraphOvalityXOffset = -25
    
    ScreenDrawingType = 2
    ScreenDrawingOrientation = 1
    
    Call DrawPVGraphsReport
    Call RenderToPrinter.RenderReport(Me, picReportPage, RenderScale)
    Call DrawPVGraphsReport

    

    Set ScreenDrawing.ReportDummyGraphImage = PVProfileImage

    DrawSF = ReportDummyGraphImage.width / ClearLineScreen.MainScreen.width  'PCN1835

    Set ScreenDrawing.ReportGraphImageX = PVReportProfile.PVXScaleImage
    Set ScreenDrawing.ReportGraphImageY = PVReportProfile.PVYScaleImage
    Call DrawMainScale(picReportPage)

    ScreenDrawingType = 0
    ScreenDrawingOrientation = 0
    
    PVReportProfile.CommentsTextBox.Visible = True
    PVReportProfile.UserTitle.Visible = True

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RP21:" & Error$
    End Select
End Sub



 Sub PositionReportControls()
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
Dim I As Integer




    


If Printer.PaperSize = 1 Then
    ShuntY(0) = -480
    ShuntY(1) = -600
    ShuntX(2) = 600
Else
    ShuntY(0) = 0
    ShuntY(1) = 0
End If

Call ShuntControlsPerContainer


On Error GoTo ManualOrientation

Printer.Orientation = vbPRORPortrait


ManualOrientationSet:

Printer.PrintQuality = vbPRPQHigh
PrinterDialogBox.Orientation = cdlPortrait

'Detect the CURRENT page setup of the deault printer
picReportPage.width = Printer.width
picReportPage.height = Printer.height
picReportPage.Left = 100
picReportPage.Top = 100

Exit Sub

ManualOrientation:

On Error GoTo Err_Handler

Dim originalheight
Dim originalwidth

originalheight = Printer.height
originalwidth = Printer.width

'PCN4367
If originalheight > originalwidth Then GoTo ManualOrientationSet


Printer.height = originalwidth
Printer.width = originalheight

picReportPage.width = Printer.width
picReportPage.height = Printer.height
GoTo ManualOrientationSet
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RP22:" & Error$
    End Select
End Sub


Private Sub Form_Resize()
On Error GoTo Err_Handler
    PageFramePictureBox.width = Me.width - PageFramePictureBox.Left
    PageFramePictureBox.height = Me.height
    PageFramePictureBox.Top = 650
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RP23:" & Error$
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler
Dim FileSaveFail As Boolean
Dim I As Integer


Call SaveToFilePipeObs(FileSaveFail)

'PCN4531
For I = 1 To Me.FloatingText.Count - 1
     Unload Me.FloatingText(I)
 Next I

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RP24:" & Error$
    End Select
End Sub

Private Sub PageFramePictureBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


If Me.ControlHighlight.Visible Then Me.ControlHighlight.Visible = False
If Me.ControlHighlightDepressed.Visible Then Me.ControlHighlightDepressed.Visible = False
    
Exit Sub
Err_Handler:
    MsgBox Err & "-RP25:" & Error$
End Sub

Private Sub picReportPage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
'    If PrintPreviewAction = "DrawText" Then
'        Call RenderToPrinter.FloatingTextAdd(Me, picReportPage, Button, Shift, X, Y)
'    Else
'        ReportMouseDown = True
'    End If
'    ReportMouseX = X
'    ReportMouseY = Y

Call ReportPageMouseDown(Me, picReportPage, Button, Shift, X, Y) 'PCN4193

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RP26:" & Error$
    End Select

End Sub

Private Sub picReportPage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    If ReportMouseDown Then
        picReportPage.Left = picReportPage.Left + X - ReportMouseX
        picReportPage.Top = picReportPage.Top + Y - ReportMouseY
    End If
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RP27:" & Error$
    End Select
End Sub

Private Sub picReportPage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

ReportMouseDown = False
If PrintPreviewAction = "MoveAll" Then
    Call SetupReportMouseIcon(PVReportProfile, 108)
End If

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RP28:" & Error$
    End Select
End Sub


Private Sub FillOutPrintForm()
On Error GoTo Err_Handler
    Dim I As Long
    Dim ControlType As String
    Dim YScaleLabel0 As Double
    Dim YScaleLabel1 As Double
    Dim YScaleLabel2 As Double
    Dim YScaleLabel3 As Double
    Dim YScaleLabel4 As Double
    Dim DisplayUnits As String
    Dim CapacityVal As Double
    Dim OvalityVal As Double
    Dim DeltaMinVal As Double
    Dim DeltaMaxVal As Double 'PCN4333
    Dim XDiameterVal As Double
    Dim YDiameterVal As Double
    Dim DiameterMedianVal As Double
    Dim DiameterMaxVal As Double
    Dim DiameterMinVal As Double 'PCN4333
    Dim ProfileSlicePosition As Single
    
    Dim LeftLimit As Double
    Dim RightLimit As Double
    Dim Limit1Top As Single
    Dim Limit2Top As Single
    
    Dim GraphInfoIndex As Integer
    Dim FractileOne As Single
    Dim FractileTwo As Single
    
    If PVFrameNo < 1 Then PVFrameNo = 1 'PCN4411 this only happens if you have not set a profile position first before saving PVD after a fresh record
    If PVFrameNo > PVDataNoOfLines Then PVFrameNo = PVDataNoOfLines  'PCN4411
    
    
    
    
    YScaleLabel0 = PVGraphXScaleUnitFromTwip(0, PrecisionVisionGraph.PVXScale(0).width / 4 * 4)
    YScaleLabel1 = PVGraphXScaleUnitFromTwip(0, PrecisionVisionGraph.PVXScale(0).width / 4 * 3)
    YScaleLabel2 = PVGraphXScaleUnitFromTwip(0, PrecisionVisionGraph.PVXScale(0).width / 4 * 2)
    YScaleLabel3 = PVGraphXScaleUnitFromTwip(0, PrecisionVisionGraph.PVXScale(0).width / 4 * 1)
    YScaleLabel4 = PVGraphXScaleUnitFromTwip(0, PrecisionVisionGraph.PVXScale(0).width / 4 * 0)
    
    YScaleLabel0 = Format(ConvertUnitByGraph(YScaleLabel0, 0, DisplayUnits), "###0.0")
    YScaleLabel1 = Format(ConvertUnitByGraph(YScaleLabel1, 0, DisplayUnits), "###0.0")
    YScaleLabel2 = Format(ConvertUnitByGraph(YScaleLabel2, 0, DisplayUnits), "###0.0")
    YScaleLabel3 = Format(ConvertUnitByGraph(YScaleLabel3, 0, DisplayUnits), "###0.0")
    YScaleLabel4 = Format(ConvertUnitByGraph(YScaleLabel4, 0, DisplayUnits), "###0.0")
    
 '   PVGraphOvalityXScale = 8
'    PVGraphOvalityXOffset = -25
    Call PrecisionVisionGraph.GetGeneralPVGraphData(ScreenDrawing.ImageGraphState(0).GraphType)
    

'PCN4253 ' The me.limitvalue1 and me.limitvalue2 added. '''''''''''''''''''''''''''''''''
    Me.LimitValue1 = Format(ConvertUnitByGraph(PVXScaleLimitPerL, 0, DisplayUnits), "###0.0")
    Me.LimitValue2 = Format(ConvertUnitByGraph(PVXScaleLimitPerR, 0, DisplayUnits), "###0.0")
    LeftLimit = SafeCDbl(Me.LimitValue1) 'PCN4161                                                   '
    RightLimit = SafeCDbl(Me.LimitValue2) 'PCN4161                                                  '
    Limit1Top = PVGraphXScaleTwipFromUnit(0, PVXScaleLimitPerL)                         '
    Limit2Top = PVGraphXScaleTwipFromUnit(0, PVXScaleLimitPerR)                         '
    Limit1Top = Limit1Top * (Me.GraphContainer.height / PrecisionVisionGraph.PVGraphImage(0).width)
    Limit2Top = Limit2Top * (Me.GraphContainer.height / PrecisionVisionGraph.PVGraphImage(0).width)
    Limit1Top = Me.GraphContainer.height - Limit1Top                                    '
    Limit2Top = Me.GraphContainer.height - Limit2Top                                    '
    Me.LimitValue1.Top = Limit1Top + Me.GraphContainer.Top                              '
    Me.LimitValue2.Top = Limit2Top + Me.GraphContainer.Top - Me.LimitValue1.height      '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    With PVReportProfile
        .SiteIDData.Caption = Trim(PipelineInfo.SiteID)
        .StartNodeData.Caption = Trim(PipelineInfo.StartName)
        .FinishNodeData.Caption = Trim(PipelineInfo.FinishName)
        .StartLocationData.Caption = Trim(PipelineInfo.StartLocation)
        .FinishLocationData.Caption = Trim(PipelineInfo.FinishLocation)
        .AssetNoData.Caption = Trim(PipelineInfo.AssetNo)
        .CityData.Caption = Trim(PipelineInfo.City)
        .PipeLengthData.Caption = Trim(PipelineInfo.PipeLength)
        .PipeMaterialData.Caption = Trim(PipelineInfo.Material)
        .PipeDiameterData.Caption = Trim(PipelineInfo.IntDiameter)
        .DateLabelData.Caption = CStr(PipelineInfo.Date)
        .PrintedData.Caption = CStr(Date)
        .PhData.Caption = PhoneNo
        .LogoImage.Picture = LoadPicture(CompanyLogoPath)
        .CompanyNameLabel = CompanyName
        
        .YScaleLabel(0) = YScaleLabel0
        .YScaleLabel(1) = YScaleLabel1
        .YScaleLabel(2) = YScaleLabel2
        .YScaleLabel(3) = YScaleLabel3
        .YScaleLabel(4) = YScaleLabel4
        
        .UpperLimitData = RightLimit
        .LowerLimitData = LeftLimit
        If PVDFileName <> "" Then
            .lblTitle.Caption = PrecisionVisionGraph.Label_GraphName(0) & " " & DisplayMessage("Profile Report")
        Else
            .lblTitle.Caption = DisplayMessage("Profile Report")
        End If
        
        .CommentsTextBox.text = PipelineInfo.Comments
        If Trim(.ObservationsLabel.Caption) = "Observation Text Here" Then
            .ObservationsLabel.Caption = ""
        End If
        

    
        'By default make all lable no background
        For I = 0 To .Controls.Count - 1
            ControlType = TypeName(.Controls(I))
            Select Case ControlType
                Case "Label": .Controls(I).BackStyle = 0
            End Select
        Next I

    If PVDataNoOfLines > 1 Then
        If MeasurementUnits = "mm" Then
            DisplayUnits = "mm"
        Else
            DisplayUnits = "in"
        End If
    
        CapacityVal = PVCapacityFullData(PVFrameNo) + CapacityDataOffset ' / PVCalculationsMultiplier
        
        OvalityVal = Abs(GraphInfoContainer(PVOvality).DataSingle(PVFrameNo)) 'PCN3540 / PVCalculationsMultiplier

        
'        DeltaMinVal = PVDeltaFullMin(PVFrameNo)
'        DeltaMaxVal = PVDeltaFullMax(PVFrameNo)
        XDiameterVal = PVXDiameterFullData(PVFrameNo)
        YDiameterVal = PVYDiameterFullData(PVFrameNo)
        DiameterMedianVal = PVDiameterMedian(PVFrameNo) + TrueDiameterOffset
        DiameterMaxVal = GraphInfoContainer(PVMaxDiameter).DataDouble(PVFrameNo)
        DiameterMinVal = GraphInfoContainer(PVMinDiameter).DataDouble(PVFrameNo) 'PCN4333
        
        
        '.PVKey_FrameNo_Value = PVFrameNo

        .PVKey_Ovality_Value.Caption = Format(OvalityVal, "#0.0") & "%"
        
        'Capacity 'PCN4974 replacing with what the pvkey is
''        If MeasurementUnits = "mm" Then
''            .PVKey_Capacity_Value_Unit(0).Caption = Format(ConvertPerToReal(CapacityVal, "Area"), "#0.0") & "cm"
''        Else
''            .PVKey_Capacity_Value_Unit(0).Caption = Format(ConvertPerToReal(CapacityVal, "Area"), "#0.00") & "in"
''        End If
''        .PVKey_Capacity_Value_Unit(1).Caption = Format(CapacityVal, "#0.0") & "%"
          .PVKey_Capacity_Value_Unit(0).Caption = PVGraphsKeyForm.PVKey_Capacity_Value_Unit.Caption
          If MedianFlat And PVDFileName <> "" Then .PVKey_Capacity_Value_Unit(1).Caption = ""
          If Not MedianFlat Then .PVKey_Capacity_Value_Unit(1).Caption = PVGraphsKeyForm.PVKey_Capacity_Value.Caption
        

        
        'X Diameter 'PCN4974 replaced with just displaying what the PVKey is'
''        If PVXDiameterFullData(PVFrameNo) > InvalidData Then                                                      '
''            .PVKey_XDiameter_Value_Unit(0).Caption = Format(XDiameterVal, "#0.0") & DisplayUnits
''            .PVKey_XDiameter_Value_Unit(1).Caption = Format(ConvertRealToPer(XDiameterVal, "Dia"), "#0.0") & "%"
''        Else
''            .PVKey_XDiameter_Value_Unit(0).Caption = DisplayUnits
''            .PVKey_XDiameter_Value_Unit(1).Caption = "%"
''        End If
            .PVKey_XDiameter_Value_Unit(0).Caption = PVGraphsKeyForm.PVKey_XDiameter_Value_Unit.Caption
            If MedianFlat And PVDFileName <> "" Then .PVKey_XDiameter_Value_Unit(1).Caption = ""
            If Not MedianFlat Then .PVKey_XDiameter_Value_Unit(1).Caption = PVGraphsKeyForm.PVKey_XDiameter_Value.Caption
        
        'Y Diameter  'PCN4974 replaced with just displaying what the PVKey is '''''''''''''''''''''''''''''''''''''''
''        If PVYDiameterFullData(PVFrameNo) > InvalidData Then                                                      '
''            .PVKey_YDiameter_Value_Unit(0).Caption = Format(YDiameterVal, "#0.0") & DisplayUnits                  '
''            .PVKey_YDiameter_Value_Unit(1).Caption = Format(ConvertRealToPer(YDiameterVal, "Dia"), "#0.0") & "%"  '
''        Else                                                                                                      '
''            .PVKey_YDiameter_Value_Unit(0).Caption = DisplayUnits                                                 '
''            .PVKey_YDiameter_Value_Unit(1).Caption = "%"                                                          '
''        End If'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            .PVKey_YDiameter_Value_Unit(0).Caption = PVGraphsKeyForm.PVKey_YDiameter_Value_Unit.Caption                     '
            If MedianFlat And PVDFileName <> "" Then .PVKey_YDiameter_Value_Unit(1).Caption = ""                                                  '
            If Not MedianFlat Then .PVKey_YDiameter_Value_Unit(1).Caption = PVGraphsKeyForm.PVKey_YDiameter_Value.Caption   '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        'True Diameter
''        If PVDiameterMedian(PVFrameNo) + TrueDiameterOffset > InvalidData Then
''            .DimensionValue.Caption = Format(DiameterMedianVal, "#0.0") & DisplayUnits
''        Else
''            .DimensionValue.Caption = DisplayUnits
''        End If
            .DimensionValue.Caption = PVGraphsKeyForm.DimensionValue.Caption
        
        'Max Diameter
        If MedianFlat And PVDFileName <> "" Then  'PCN4974
            If GraphInfoContainer(PVMaxDiameter).DataDouble(PVFrameNo) > InvalidData Then .PVKey_MaxDiameter_Value_Unit(0).Caption = Format(DiameterMaxVal, "#0.0") & DisplayUnits
            .PVKey_MaxDiameter_Value_Unit(1).Caption = ""
        ElseIf GraphInfoContainer(PVMaxDiameter).DataDouble(PVFrameNo) > InvalidData Then
            .PVKey_MaxDiameter_Value_Unit(0).Caption = Format(DiameterMaxVal, "#0.0") & DisplayUnits
            .PVKey_MaxDiameter_Value_Unit(1).Caption = Format(ConvertRealToPer(DiameterMaxVal, "Dia"), "#0.0") & "%"
        Else
            .PVKey_MaxDiameter_Value_Unit(0).Caption = DisplayUnits
            .PVKey_MaxDiameter_Value_Unit(1).Caption = "%"
        End If

            
        'Min diameter 'PCN4333
        If MedianFlat And PVDFileName <> "" Then 'PCN4974
            If GraphInfoContainer(PVMinDiameter).DataDouble(PVFrameNo) > InvalidData Then .PVKey_MinDiameter_Value_Unit(0).Caption = Format(DiameterMinVal, "#0.0") & DisplayUnits
            .PVKey_MinDiameter_Value_Unit(1).Caption = ""
        ElseIf GraphInfoContainer(PVMinDiameter).DataDouble(PVFrameNo) > InvalidData Then
            .PVKey_MinDiameter_Value_Unit(0).Caption = Format(DiameterMinVal, "#0.0") & DisplayUnits
            .PVKey_MinDiameter_Value_Unit(1).Caption = Format(ConvertRealToPer(DiameterMinVal, "Dia"), "#0.0") & "%"
        Else
            .PVKey_MinDiameter_Value_Unit(0).Caption = DisplayUnits
            .PVKey_MinDiameter_Value_Unit(1).Caption = "%"
        End If

        
        '^^^^ ***********************************************
        'Distance
        'If ConfigInfo.DistanceStart >= 0 Then
            If MeasurementUnits = "mm" Then
                .PVKey_Distance_Value = Format(PVDistances(PVFrameNo), "#0.0") & "m"
                .PVKey_Distance_Icon.Caption = "m" 'PCN4415
            Else
                .PVKey_Distance_Value = Format(PVDistances(PVFrameNo), "#0") & "ft"
                .PVKey_Distance_Icon.Caption = "ft" 'PCN4415
            End If
            
        .UnitLabel = PrecisionVisionGraph.PVXScaleLabelUnits(0)
        'End If
        
        GraphInfoIndex = GetGraphInfoIndex(0)
        If GraphInfoIndex <> 0 Then ' "flat"
            If GraphInfoIndex <> 4 Then '"XYDiameter needs two values
                FractileOne = ScreenDrawing.GraphInfoContainer(GraphInfoIndex).Fractile90
                Me.FractileLabelData = Format(ConvertUnitByGraph(FractileOne, 0, DisplayUnits), "###0.0") + DisplayUnits
            Else
                FractileOne = ScreenDrawing.GraphInfoContainer(GraphInfoIndex).Fractile90
                FractileTwo = ScreenDrawing.GraphInfoContainer(PVYDiameter).Fractile90 'YDiameter
                Me.FractileLabelData = "(X) " + _
                                        Format(ConvertUnitByGraph(FractileOne, 0, DisplayUnits), "###0.0") + _
                                        DisplayUnits + " :" + " (Y) " + _
                                        Format(ConvertUnitByGraph(FractileTwo, 0, DisplayUnits), "###0.0") + _
                                        DisplayUnits
            End If
            
            Me.FractileLabelData = Me.FractileLabelData + ",   " + DisplayMessage("Exceeded limits") + ": " + Format(GraphInfoContainer(GraphInfoIndex).PercentExceeded, "###0.0") + "%"
            
            
        End If
    End If
        
    'If VideoSnapShotMode = SnapShot And CLPScreenMode = Video Then 'PCNGL210103
'    If CLPScreenMode = SnapShot Then 'PCNGL210103 'PCN4043
'        .PVProfileImage.Picture = LoadPicture(LocToSave & "Snapshot.bmp")
'    End If
    
    If MeasurementUnits = "mm" Then
        .PipeLengthUnitLabel = "m"
        .GraphUnitLbl = "m" 'PCN4382
        .PipeDiameterUnitLabel = "mm"
    Else
        .PipeLengthUnitLabel = "ft"
        .GraphUnitLbl = "ft" 'PCN4382
        .PipeDiameterUnitLabel = "in"
    End If
    
    Me.PVKey_Flat3D_Value7_Unit(0) = PVGraphsKeyForm.PVKey_Flat3D_Value7_Unit(0) 'PCN4910 (1) is diameter
    Me.PVKey_Flat3D_Value6_Unit(0) = PVGraphsKeyForm.PVKey_Flat3D_Value6_Unit(0) 'PCN4910 (1) is diameter
    Me.PVKey_Flat3D_Value5_Unit(0) = PVGraphsKeyForm.PVKey_Flat3D_Value5_Unit(0) 'PCN4910 (1) is diameter
    Me.PVKey_Flat3D_Value4_Unit(0) = PVGraphsKeyForm.PVKey_Flat3D_Value4_Unit(0) 'PCN4910 (1) is diameter
    Me.PVKey_Flat3D_Value2_Unit(0) = PVGraphsKeyForm.PVKey_Flat3D_Value2_Unit(0) 'PCN4910 (1) is diameter
    Me.PVKey_Flat3D_Value1_Unit(0) = PVGraphsKeyForm.PVKey_Flat3D_Value1_Unit(0) 'PCN4910 (1) is diameter
    Me.PVKey_Flat3D_Value0_Unit(0) = PVGraphsKeyForm.PVKey_Flat3D_Value0_Unit(0) 'PCN4910 (1) is diameter
    
    Me.PVKey_Flat3D_Value7_Unit(1) = PVGraphsKeyForm.PVKey_Flat3D_Value7_Unit(1) 'PCN4910 (1) is diameter
    Me.PVKey_Flat3D_Value6_Unit(1) = PVGraphsKeyForm.PVKey_Flat3D_Value6_Unit(1) 'PCN4910 (1) is diameter
    Me.PVKey_Flat3D_Value5_Unit(1) = PVGraphsKeyForm.PVKey_Flat3D_Value5_Unit(1) 'PCN4910 (1) is diameter
    Me.PVKey_Flat3D_Value4_Unit(1) = PVGraphsKeyForm.PVKey_Flat3D_Value4_Unit(1) 'PCN4910 (1) is diameter
    Me.PVKey_Flat3D_Value2_Unit(1) = PVGraphsKeyForm.PVKey_Flat3D_Value2_Unit(1) 'PCN4910 (1) is diameter
    Me.PVKey_Flat3D_Value1_Unit(1) = PVGraphsKeyForm.PVKey_Flat3D_Value1_Unit(1) 'PCN4910 (1) is diameter
    Me.PVKey_Flat3D_Value0_Unit(1) = PVGraphsKeyForm.PVKey_Flat3D_Value0_Unit(1) 'PCN4910 (1) is diameter
    
    Me.PVKey_Flat3D_Value7 = PVGraphsKeyForm.PVKey_Flat3D_Value7.text
    Me.PVKey_Flat3D_Value6 = PVGraphsKeyForm.PVKey_Flat3D_Value6.Caption
    Me.PVKey_Flat3D_Value5 = PVGraphsKeyForm.PVKey_Flat3D_Value5.Caption
    Me.PVKey_Flat3D_Value4 = PVGraphsKeyForm.PVKey_Flat3D_Value4.Caption
    Me.PVKey_Flat3D_Value2 = PVGraphsKeyForm.PVKey_Flat3D_Value2.Caption
    Me.PVKey_Flat3D_Value1 = PVGraphsKeyForm.PVKey_Flat3D_Value1.Caption
    Me.PVKey_Flat3D_Value0 = PVGraphsKeyForm.PVKey_Flat3D_Value0.text
    
    End With
    
    ProfileSlicePosition = GraphContainer.width / (CSng(PreviewEndFrame - PreviewStartFrame))
    ProfileSlicePosition = (PVFrameNo - PreviewStartFrame) * ProfileSlicePosition
    If ProfileSlicePosition < 0 Then ProfileSlicePosition = -600
    If ProfileSlicePosition > GraphContainer.width Then ProfileSlicePosition = GraphContainer.width + 300
    
    ProfileSlicePosition = ProfileSlicePosition + GraphContainer.Left
    
    ProfileSliceRubberBand.x1 = ProfileSlicePosition
    ProfileSliceRubberBand.y2 = Line7.y2
    
    ProfileSlice.x1 = ProfileSlicePosition
    ProfileSlice.x2 = ProfileSlicePosition
    
    
    
    Call ScreenDrawing.FormTopMost(PVGraphsKeyForm.hwnd) 'PCN2990
Exit Sub
Err_Handler:
    Select Case Err
        Case 11: Resume Next 'divide by zero , PCN4552
        Case 6: Resume Next 'overflow, PCN4552, these are caused by not PVD
        Case Else: MsgBox Err & "-RP29:" & Error$
    End Select
    
End Sub

Sub GraphSpecificSettings()
On Error GoTo Err_Handler
    Dim I As Integer
    Dim ColourKeysLeftShift As Single
    Dim ColourKeysTopShift As Single

    If ScreenDrawing.ImageGraphState(0).GraphType = "Ovality" Then
        PVReportProfile.UnitLabel = PrecisionVisionGraph.PVXScaleLabelUnits(0)
        

        PVReportProfile.Explination.Visible = True
        PVReportProfile.Explination.Tag = ""
        
        Me.LimitValue2.Visible = False: Me.LimitValue2.Tag = "" 'PCN4253
        
        'PCN4324 if ovality display is less than zero, dont display
        For I = 0 To 4
            If Me.YScaleLabel(I) < 0 Then Me.YScaleLabel(I).Visible = False: Me.YScaleLabel(I).Tag = ""
        Next I
        
    ElseIf ScreenDrawing.ImageGraphState(0).GraphType = "Flat" Then
        With PVReportProfile
        
            .UnitLabel.Visible = False
            .UnitLabel.Tag = ""
            .YScaleLabel(0) = 6
            .YScaleLabel(1) = 3
            .YScaleLabel(2) = 12
            .YScaleLabel(3) = 9
            .YScaleLabel(4) = 6
            .UnitLabel = PrecisionVisionGraph.PVXScaleLabelUnits(0)
        End With
        Me.LimitValue1.Visible = False: Me.LimitValue1.Tag = "" 'PCN4253
        Me.LimitValue2.Visible = False: Me.LimitValue2.Tag = "" 'PCN4253
        Me.FractileLabel.Visible = False: Me.FractileLabelData.Visible = False 'PCN4296
        Me.FractileLabel.Tag = "": Me.FractileLabelData.Tag = ""
        
        Me.LimitLinesLabel.Visible = False: Me.LimitLinesLabel.Tag = ""
        Me.UpperLimitLabel.Visible = False: Me.UpperLimitLabel.Tag = ""
        Me.UperLimitData.Visible = False: Me.UperLimitData.Tag = ""
        Me.LowerLimitLabel.Visible = False: Me.LowerLimitLabel.Tag = ""
        Me.UpperLimitData.Visible = False: UpperLimitData.Tag = ""
        Me.LowerLimitData.Visible = False: LowerLimitData.Tag = ""
        Me.UpperLimitLine.Visible = False: UpperLimitLine.Tag = ""
        Me.LowerLimitLine.Visible = False: LowerLimitLine.Tag = ""

        
        ColourKeysLeftShift = Me.KeyDataBoarder.Left - Me.UperLimitData.Left + 1415 - 1425 'PCN4920
        ColourKeysTopShift = 0
        
        'PCN9420.................................................
        Me.KeyLine1.x1 = Me.KeyLine1.x1 - ColourKeysLeftShift   '
        Me.KeyLine1.x2 = Me.KeyLine1.x2 - ColourKeysLeftShift   '
        Me.KeyLine2.x1 = Me.KeyLine2.x1 - ColourKeysLeftShift   '
        Me.KeyLine2.x2 = Me.KeyLine2.x2 - ColourKeysLeftShift   '
        '........................................................
        Me.DiameterLabel.Left = Me.DiameterLabel.Left - ColourKeysLeftShift
        Me.RadiusLabel.Left = Me.RadiusLabel.Left - ColourKeysLeftShift
        
        
        Me.PVKey_Flat3D_Color7.Left = Me.PVKey_Flat3D_Color7.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Color6.Left = Me.PVKey_Flat3D_Color6.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Color5.Left = Me.PVKey_Flat3D_Color5.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Color4.Left = Me.PVKey_Flat3D_Color4.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Color2.Left = Me.PVKey_Flat3D_Color2.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Color1.Left = Me.PVKey_Flat3D_Color1.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Color0.Left = Me.PVKey_Flat3D_Color0.Left - ColourKeysLeftShift
        
        Me.PVKey_Flat3D_Value7.Left = Me.PVKey_Flat3D_Value7.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value6.Left = Me.PVKey_Flat3D_Value6.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value5.Left = Me.PVKey_Flat3D_Value5.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value4.Left = Me.PVKey_Flat3D_Value4.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value2.Left = Me.PVKey_Flat3D_Value2.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value1.Left = Me.PVKey_Flat3D_Value1.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value0.Left = Me.PVKey_Flat3D_Value0.Left - ColourKeysLeftShift
        
        Me.PVKey_Flat3D_Value7_Unit(0).Left = Me.PVKey_Flat3D_Value7_Unit(0).Left - ColourKeysLeftShift 'PCN4920 (1) is diameter
        Me.PVKey_Flat3D_Value6_Unit(0).Left = Me.PVKey_Flat3D_Value6_Unit(0).Left - ColourKeysLeftShift 'PCN4920 (1) is diameter
        Me.PVKey_Flat3D_Value5_Unit(0).Left = Me.PVKey_Flat3D_Value5_Unit(0).Left - ColourKeysLeftShift 'PCN4920 (1) is diameter
        Me.PVKey_Flat3D_Value4_Unit(0).Left = Me.PVKey_Flat3D_Value4_Unit(0).Left - ColourKeysLeftShift 'PCN4920 (1) is diameter
        Me.PVKey_Flat3D_Value2_Unit(0).Left = Me.PVKey_Flat3D_Value2_Unit(0).Left - ColourKeysLeftShift 'PCN4920 (1) is diameter
        Me.PVKey_Flat3D_Value1_Unit(0).Left = Me.PVKey_Flat3D_Value1_Unit(0).Left - ColourKeysLeftShift 'PCN4920 (1) is diameter
        Me.PVKey_Flat3D_Value0_Unit(0).Left = Me.PVKey_Flat3D_Value0_Unit(0).Left - ColourKeysLeftShift 'PCN4920 (1) is diameter
        
        Me.PVKey_Flat3D_Value7_Unit(1).Left = Me.PVKey_Flat3D_Value7_Unit(1).Left - ColourKeysLeftShift 'PCN4920 (1) is diameter
        Me.PVKey_Flat3D_Value6_Unit(1).Left = Me.PVKey_Flat3D_Value6_Unit(1).Left - ColourKeysLeftShift 'PCN4920 (1) is diameter
        Me.PVKey_Flat3D_Value5_Unit(1).Left = Me.PVKey_Flat3D_Value5_Unit(1).Left - ColourKeysLeftShift 'PCN4920 (1) is diameter
        Me.PVKey_Flat3D_Value4_Unit(1).Left = Me.PVKey_Flat3D_Value4_Unit(1).Left - ColourKeysLeftShift 'PCN4920 (1) is diameter
        Me.PVKey_Flat3D_Value2_Unit(1).Left = Me.PVKey_Flat3D_Value2_Unit(1).Left - ColourKeysLeftShift 'PCN4920 (1) is diameter
        Me.PVKey_Flat3D_Value1_Unit(1).Left = Me.PVKey_Flat3D_Value1_Unit(1).Left - ColourKeysLeftShift 'PCN4920 (1) is diameter
        Me.PVKey_Flat3D_Value0_Unit(1).Left = Me.PVKey_Flat3D_Value0_Unit(1).Left - ColourKeysLeftShift 'PCN4920 (1) is diameter
        Me.KeyDataBoarder.Left = Me.KeyDataBoarder.Left - ColourKeysLeftShift
        
                Me.KeyLine1.Visible = True: Me.KeyLine1.Tag = "Visible"
        Me.KeyLine2.Visible = True: Me.KeyLine2.Tag = "Visible"
        Me.DiameterLabel.Visible = True: Me.DiameterLabel.Tag = "Visible"
        Me.RadiusLabel.Visible = True: Me.RadiusLabel.Tag = "Visible"
        
        Me.PVKey_Flat3D_Color7.Visible = True: Me.PVKey_Flat3D_Color7.Tag = "Visible"
        Me.PVKey_Flat3D_Color6.Visible = True: Me.PVKey_Flat3D_Color6.Tag = "Visible"
        Me.PVKey_Flat3D_Color5.Visible = True: Me.PVKey_Flat3D_Color5.Tag = "Visible"
        Me.PVKey_Flat3D_Color4.Visible = True: Me.PVKey_Flat3D_Color4.Tag = "Visible"
        Me.PVKey_Flat3D_Color2.Visible = True: Me.PVKey_Flat3D_Color2.Tag = "Visible"
        Me.PVKey_Flat3D_Color1.Visible = True: Me.PVKey_Flat3D_Color1.Tag = "Visible"
        Me.PVKey_Flat3D_Color0.Visible = True: Me.PVKey_Flat3D_Color0.Tag = "Visible"
        
        Me.PVKey_Flat3D_Value7.Visible = True: Me.PVKey_Flat3D_Value7.Tag = "Visible"
        Me.PVKey_Flat3D_Value6.Visible = True: Me.PVKey_Flat3D_Value6.Tag = "Visible"
        Me.PVKey_Flat3D_Value5.Visible = True: Me.PVKey_Flat3D_Value5.Tag = "Visible"
        Me.PVKey_Flat3D_Value4.Visible = True: Me.PVKey_Flat3D_Value4.Tag = "Visible"
        Me.PVKey_Flat3D_Value2.Visible = True: Me.PVKey_Flat3D_Value2.Tag = "Visible"
        Me.PVKey_Flat3D_Value1.Visible = True: Me.PVKey_Flat3D_Value1.Tag = "Visible"
        Me.PVKey_Flat3D_Value0.Visible = True: Me.PVKey_Flat3D_Value0.Tag = "Visible"
        
        Me.PVKey_Flat3D_Value7_Unit(0).Visible = True: Me.PVKey_Flat3D_Value7_Unit(0).Tag = "Visible" 'PCN4920 (1) is diameter, (0) is radius
        Me.PVKey_Flat3D_Value6_Unit(0).Visible = True: Me.PVKey_Flat3D_Value6_Unit(0).Tag = "Visible" 'PCN4920 (1) is diameter, (0) is radius
        Me.PVKey_Flat3D_Value5_Unit(0).Visible = True: Me.PVKey_Flat3D_Value5_Unit(0).Tag = "Visible" 'PCN4920 (1) is diameter, (0) is radius
        Me.PVKey_Flat3D_Value4_Unit(0).Visible = True: Me.PVKey_Flat3D_Value4_Unit(0).Tag = "Visible" 'PCN4920 (1) is diameter, (0) is radius
        Me.PVKey_Flat3D_Value2_Unit(0).Visible = True: Me.PVKey_Flat3D_Value2_Unit(0).Tag = "Visible" 'PCN4920 (1) is diameter, (0) is radius
        Me.PVKey_Flat3D_Value1_Unit(0).Visible = True: Me.PVKey_Flat3D_Value1_Unit(0).Tag = "Visible" 'PCN4920 (1) is diameter, (0) is radius
        Me.PVKey_Flat3D_Value0_Unit(0).Visible = True: Me.PVKey_Flat3D_Value0_Unit(0).Tag = "Visible" 'PCN4920 (1) is diameter, (0) is radius
        
        Me.PVKey_Flat3D_Value7_Unit(1).Visible = True: Me.PVKey_Flat3D_Value7_Unit(1).Tag = "Visible" 'PCN4920 (1) is diameter, (0) is radius
        Me.PVKey_Flat3D_Value6_Unit(1).Visible = True: Me.PVKey_Flat3D_Value6_Unit(1).Tag = "Visible" 'PCN4920 (1) is diameter, (0) is radius
        Me.PVKey_Flat3D_Value5_Unit(1).Visible = True: Me.PVKey_Flat3D_Value5_Unit(1).Tag = "Visible" 'PCN4920 (1) is diameter, (0) is radius
        Me.PVKey_Flat3D_Value4_Unit(1).Visible = True: Me.PVKey_Flat3D_Value4_Unit(1).Tag = "Visible" 'PCN4920 (1) is diameter, (0) is radius
        Me.PVKey_Flat3D_Value2_Unit(1).Visible = True: Me.PVKey_Flat3D_Value2_Unit(1).Tag = "Visible" 'PCN4920 (1) is diameter, (0) is radius
        Me.PVKey_Flat3D_Value1_Unit(1).Visible = True: Me.PVKey_Flat3D_Value1_Unit(1).Tag = "Visible" 'PCN4920 (1) is diameter, (0) is radius
        Me.PVKey_Flat3D_Value0_Unit(1).Visible = True: Me.PVKey_Flat3D_Value0_Unit(1).Tag = "Visible" 'PCN4920 (1) is diameter, (0) is radius
        
        Me.KeyDataBoarder.Visible = True: Me.KeyDataBoarder.Tag = "Visible"
        
        
        If MedianFlat And PVDFileName <> "" Then
            Me.PVKey_Flat3D_Value7_Unit(0).Visible = False: Me.PVKey_Flat3D_Value7_Unit(0).Tag = "" 'PCN4920 (0) radial (1) diameter
            Me.PVKey_Flat3D_Value6_Unit(0).Visible = False: Me.PVKey_Flat3D_Value6_Unit(0).Tag = "" 'PCN4920 (0) radial (1) diameter
            Me.PVKey_Flat3D_Value5_Unit(0).Visible = False: Me.PVKey_Flat3D_Value5_Unit(0).Tag = "" 'PCN4920 (0) radial (1) diameter
            Me.PVKey_Flat3D_Value4_Unit(0).Visible = False: Me.PVKey_Flat3D_Value4_Unit(0).Tag = "" 'PCN4920 (0) radial (1) diameter
            Me.PVKey_Flat3D_Value2_Unit(0).Visible = False: Me.PVKey_Flat3D_Value2_Unit(0).Tag = "" 'PCN4920 (0) radial (1) diameter
            Me.PVKey_Flat3D_Value1_Unit(0).Visible = False: Me.PVKey_Flat3D_Value1_Unit(0).Tag = "" 'PCN4920 (0) radial (1) diameter
            Me.PVKey_Flat3D_Value0_Unit(0).Visible = False: Me.PVKey_Flat3D_Value0_Unit(0).Tag = "" 'PCN4920 (0) radial (1) diameter

            Me.KeyLine1.Visible = False: Me.KeyLine1.Tag = ""
            PVKey_Flat3D_Value0.Left = PVKey_Flat3D_Value0.Left + 700
            PVKey_Flat3D_Value1.Left = PVKey_Flat3D_Value1.Left + 700
            PVKey_Flat3D_Value2.Left = PVKey_Flat3D_Value2.Left + 700

'
'            PVKey_Flat3D_Value4_Unit(1).Left = PVKey_Flat3D_Value4_Unit(1).Left + 700
'            PVKey_Flat3D_Value5_Unit(1).Left = PVKey_Flat3D_Value5_Unit(1).Left + 700
'            PVKey_Flat3D_Value6_Unit(1).Left = PVKey_Flat3D_Value6_Unit(1).Left + 700
'            PVKey_Flat3D_Value7_Unit(1).Left = PVKey_Flat3D_Value7_Unit(1).Left + 700

            PVKey_Flat3D_Value4.Left = PVKey_Flat3D_Value4.Left + 700
            PVKey_Flat3D_Value5.Left = PVKey_Flat3D_Value5.Left + 700
            PVKey_Flat3D_Value6.Left = PVKey_Flat3D_Value6.Left + 700
            PVKey_Flat3D_Value7.Left = PVKey_Flat3D_Value7.Left + 700

            PVKey_Flat3D_Color0.Left = PVKey_Flat3D_Color0.Left + 700
            PVKey_Flat3D_Color1.Left = PVKey_Flat3D_Color1.Left + 700
            PVKey_Flat3D_Color2.Left = PVKey_Flat3D_Color2.Left + 700
            PVKey_Flat3D_Color4.Left = PVKey_Flat3D_Color4.Left + 700
            PVKey_Flat3D_Color5.Left = PVKey_Flat3D_Color5.Left + 700
            PVKey_Flat3D_Color6.Left = PVKey_Flat3D_Color6.Left + 700
            PVKey_Flat3D_Color7.Left = PVKey_Flat3D_Color7.Left + 700
            RadiusLabel.Left = RadiusLabel.Left + 700
            KeyDataBoarder.width = KeyDataBoarder.width - 700
            KeyDataBoarder.Left = KeyDataBoarder.Left + 700
            Me.GraphContainer.width = Me.GraphContainer.width + 700
            Me.Shape2.width = Me.Shape2.width + 700
            Me.GraphUnitLbl.Left = Me.GraphUnitLbl.Left + 700
            Me.GraphXScaleContainer.width = Me.GraphXScaleContainer.width + 700
            Me.GraphBackgroundShape.width = Me.GraphBackgroundShape.width + 700
            
            Me.PVKey_XDiameter_Value_Unit(1).Visible = False: Me.PVKey_XDiameter_Value_Unit(1).Tag = ""
            Me.PVKey_YDiameter_Value_Unit(1).Visible = False: Me.PVKey_YDiameter_Value_Unit(1).Tag = ""
            
        
        End If
        
    End If
    
    If MedianFlat And PVDFileName <> "" Then 'PCN5186
        Me.FractileLabel.Visible = False: Me.FractileLabelData.Visible = False 'PCN4296
        Me.FractileLabel.Tag = "": Me.FractileLabelData.Tag = ""
    End If
    
    If ScreenDrawing.ImageGraphState(0).GraphType <> "Ovality" Then
        PVReportProfile.Explination.Visible = False
        PVReportProfile.Explination.Tag = ""
    End If
    
    If ScreenDrawing.ImageGraphState(0).GraphType = "Inclination" Then
        Me.LimitLinesLabel.Visible = False: Me.LimitLinesLabel.Tag = ""
'        Me.UpperLimitLabel.Visible = False: Me.UpperLimitLabel.Tag = ""
'        Me.LowerLimitLabel.Visible = False: Me.LowerLimitLabel.Tag = ""
'        Me.UpperLimitData.Visible = False: UpperLimitData.Tag = ""
'        Me.LowerLimitData = False: LowerLimitData.Tag = ""
        Me.UperLimitData.Visible = False: Me.UperLimitData.Tag = ""
        LimitValue1.Visible = False: Me.LimitValue1.Tag = ""
        Me.LimitValue2.Visible = False: Me.LimitValue2.Tag = ""
         PVReportSingle.Explination.Visible = True
        PVReportSingle.Explination.Tag = "Visible"
        
        PVReportSingle.Explination.Caption = DisplayMessage("The accuracy is dependant on the precision of the provided sea level heights and is an indicator only")
    End If
    
    Dim UserTittleCaption As String
    Dim TittleParse As String
        
    TittleParse = "Profile_" & ScreenDrawing.ImageGraphState(0).GraphType & "="
    Call GetINI_ParameterInfoOnly(MyFile, TittleParse, UserTittleCaption)
    Me.UserTitle.text = UserTittleCaption
    
    If Me.LimitValue2.Top + Me.LimitValue2.height < Me.GraphContainer.Top Then Me.LimitValue2.Visible = False: Me.LimitValue2.Tag = ""
    If Me.LimitValue1.Top > Me.GraphContainer.Top + Me.GraphContainer.height Then Me.LimitValue1.Visible = False: Me.LimitValue1.Tag = ""

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RP30:" & Error$
    End Select
End Sub

Sub MarkForPrinting()
On Error GoTo Err_Handler

Dim I As Integer
Dim ControlType As String

'Draw renderings first that are marked back
For I = 1 To PVReportProfile.Controls.Count - 1
    
    With PVReportProfile.Controls(I)
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
No_Container:
    End With
Next I
Exit Sub
Err_Handler:
    Select Case Err
        Case 438: Resume No_Container
        Case Else: MsgBox Err & "-RP31:" & Error$
    End Select
End Sub

Private Sub PrinterSettingsButton_Click()
On Error GoTo Err_Handler
    PrinterDialogBox.ShowPrinter
Exit Sub
Err_Handler:
    If Err = 32755 Then Exit Sub ' Cancel Printer Setting
End Sub





Private Sub ScaleButton05_Click()
On Error GoTo Err_Handler
    RenderScale = 0.5
    Call RenderForm
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RP32:" & Error$
    End Select
End Sub

Private Sub ScaleButton10_Click()
On Error GoTo Err_Handler
    RenderScale = 1
    Call RenderForm
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RP33:" & Error$
    End Select
End Sub

Private Sub ScaleButton15_Click()
On Error GoTo Err_Handler
    RenderScale = 1.5
    Call RenderForm
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RP34:" & Error$
    End Select
End Sub

Private Sub ScaleButton20_Click()
On Error GoTo Err_Handler
    RenderScale = 2
    Call RenderForm
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RP35:" & Error$
    End Select
End Sub

'===========================
'Declare the Function to select printer
'===========================
 
Function SelectPrinter(ByVal printer_name As String) As Boolean
On Error GoTo Err_Handler
    
    Dim I As Integer
 
    SelectPrinter = False
    For I = 0 To Printers.Count - 1
        ' if the specified printer is found, select it and return True
        If Printers(I).DeviceName = printer_name Then
            Set Printer = Printers(I)
            SelectPrinter = True
            Exit For
        End If
    Next I
    
Exit Function
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RP36:" & Error$
    End Select
End Function
 


Private Sub CmboPrinterList_Click()
On Error GoTo Err_Handler
    Dim I As Integer
    
    Call SelectPrinter(CmboPrinterList.text)
    
    If Me.CmboPrinterList.Enabled = True Then
        Me.picReportPage.Cls
        
        For I = 1 To Me.FloatingText.Count - 1 'We dont want the text to be rendered on the preview yet
            Me.FloatingText(I).Visible = False 'when changes printers
        Next I
        
        Me.RestoreOriginalState
        Me.InitialiseForm
        
        For I = 1 To Me.FloatingText.Count - 1 'Even thou we dont want the text to be rendered, we still
            Me.FloatingText(I).Visible = True  'want to be able to see them after printer select changed
        Next I
    End If
    
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RP37:" & Error$
    End Select
End Sub

Sub ShuntControlsPerContainer()
On Error GoTo Err_Handler

Dim I As Integer
Dim S As Integer
Dim ControlType As String

'Draw renderings first that are marked back

For S = 0 To 1 'Loop for the different Shunts
    For I = 1 To PVReportProfile.Controls.Count - 1
        
        With PVReportProfile.Controls(I)
            If TypeName(.Container) = "PictureBox" Then
                If .Tag <> "Container" Then
                    If TypeName(PVReportProfile.Controls(I)) = "TextBox" Or _
                       TypeName(PVReportProfile.Controls(I)) = "Label" Or _
                       TypeName(PVReportProfile.Controls(I)) = "Shape" Or _
                       TypeName(PVReportProfile.Controls(I)) = "Image" Or _
                       TypeName(PVReportProfile.Controls(I)) = "PictureBox" Then
                        If .Left > ControlsPositionShunt(S).Left And _
                           .Left < ControlsPositionShunt(S).Left + ControlsPositionShunt(S).width And _
                           .Top > ControlsPositionShunt(S).Top And _
                           .Top < ControlsPositionShunt(S).Top + ControlsPositionShunt(S).height Then
                            .Top = .Top + ShuntY(S)
                            .Left = .Left + ShuntX(S)
                        End If
                    End If
                    If TypeName(PVReportProfile.Controls(I)) = "Line" Then
                        If .x1 > ControlsPositionShunt(S).Left And _
                           .x1 < ControlsPositionShunt(S).Left + ControlsPositionShunt(S).width And _
                           .y1 > ControlsPositionShunt(S).Top And _
                           .y1 < ControlsPositionShunt(S).Top + ControlsPositionShunt(S).height Then
                            .y1 = .y1 + ShuntY(S)
                            .y2 = .y2 + ShuntY(S)
                            .x1 = .x1 + ShuntX(S)
                            .x2 = .x2 + ShuntX(S)
                        End If
                    End If
                End If
            End If
No_Container:
        End With
    Next I
Next S
Exit Sub
Err_Handler:
    Select Case Err
        Case 438: Resume No_Container
        Case Else: MsgBox Err & "-RP38:" & Error$
    End Select
End Sub

Sub RestoreOriginalState()
On Error GoTo Err_Handler
    Dim I As Long

    For I = 1 To Me.Controls.Count - 1
        If Me.Controls(I).name <> "FloatingText" Then
            Me.Controls(I).Visible = OriginalStateVisible(I)
            Me.Controls(I).Tag = OriginalStateTag(I)
            Me.Controls(I).Left = OriginalStateLeft(I)
            Me.Controls(I).Top = OriginalStateTop(I)
            Me.Controls(I).x1 = OriginalStateX1(I)
            Me.Controls(I).y1 = OriginalStateY1(I)
            Me.Controls(I).x2 = OriginalStateX2(I)
            Me.Controls(I).y2 = OriginalStateY2(I)
        End If
    Next I
Exit Sub
Err_Handler:
    Select Case Err
        Case 9 'Subscript out of range
            'Addition of text will cause this error
            Exit Sub
        Case 438, 382: Resume Next
        Case Else: MsgBox Err & "-RP39:" & Error$
    End Select

End Sub


Private Sub UserTitle_Change()
On Error GoTo Err_Handler

'Dim FileSaveFail As Boolean

'UserTitleProfile = Trim(Me.UserTitle.text) 'PCN4433

    Dim UserTittleCaption As String
    Dim TittleParse As String
        
    TittleParse = "Profile_" & ScreenDrawing.ImageGraphState(0).GraphType & "="
    Call INI_WriteBack(MyFile, TittleParse, Me.UserTitle.text)

Exit Sub
Err_Handler:
    Select Case Err

        Case Else: MsgBox Err & "-RP40:" & Error$
    End Select
End Sub
