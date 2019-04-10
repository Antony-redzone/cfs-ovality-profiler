VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form PVReportSingle 
   BorderStyle     =   0  'None
   Caption         =   "Summary Report"
   ClientHeight    =   12600
   ClientLeft      =   150
   ClientTop       =   -2865
   ClientWidth     =   20715
   Icon            =   "PVReportSingle.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12600
   ScaleWidth      =   20715
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PageFramePictureBox 
      BackColor       =   &H00FFC0C0&
      Height          =   11775
      Left            =   120
      ScaleHeight     =   11715
      ScaleWidth      =   18915
      TabIndex        =   2
      Top             =   720
      Width           =   18975
      Begin VB.PictureBox picReportPage 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   11904
         Left            =   120
         ScaleHeight     =   11850
         ScaleWidth      =   18675
         TabIndex        =   3
         Tag             =   "Paper"
         Top             =   120
         Width           =   18735
         Begin VB.TextBox UserTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2760
            MousePointer    =   3  'I-Beam
            TabIndex        =   54
            Top             =   840
            Visible         =   0   'False
            Width           =   11175
         End
         Begin VB.TextBox CommentsTextBox 
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
            Height          =   1095
            Left            =   480
            MousePointer    =   3  'I-Beam
            MultiLine       =   -1  'True
            TabIndex        =   6
            Top             =   2640
            Width           =   10365
         End
         Begin VB.TextBox FloatingTextDefault 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2520
            TabIndex        =   5
            Text            =   "Default Text Setting"
            Top             =   360
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox FloatingText 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   4320
            TabIndex        =   4
            Top             =   360
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label TunningSettingsLabel 
            BackColor       =   &H00FFFFFF&
            Caption         =   "4.23.36.19.04"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   7800
            TabIndex        =   84
            Top             =   10680
            Width           =   2295
         End
         Begin VB.Line KeyLine 
            Visible         =   0   'False
            X1              =   15975
            X2              =   15975
            Y1              =   2400
            Y2              =   3720
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
            Index           =   1
            Left            =   16530
            TabIndex        =   83
            Top             =   2385
            Visible         =   0   'False
            Width           =   255
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
            Index           =   1
            Left            =   17880
            TabIndex        =   82
            Top             =   2445
            Visible         =   0   'False
            Width           =   255
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
            Index           =   0
            Left            =   14130
            TabIndex        =   81
            Top             =   2385
            Visible         =   0   'False
            Width           =   255
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
            Index           =   0
            Left            =   15480
            TabIndex        =   80
            Top             =   2445
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label PVKey_Flat3D_Value7_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   15360
            TabIndex        =   79
            Top             =   2700
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label PVKey_Flat3D_Value5_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   15360
            TabIndex        =   78
            Top             =   3180
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label PVKey_Flat3D_Value6_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   15360
            TabIndex        =   77
            Top             =   2940
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label PVKey_Flat3D_Value4_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   15360
            TabIndex        =   76
            Top             =   3435
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label PVKey_Flat3D_Value0_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   17760
            TabIndex        =   75
            Top             =   3180
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label PVKey_Flat3D_Value1_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   17760
            TabIndex        =   74
            Top             =   2940
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label PVKey_Flat3D_Value2_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   17760
            TabIndex        =   73
            Top             =   2700
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label PVKey_Flat3D_Value0 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   16440
            TabIndex        =   72
            Top             =   3180
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label PVKey_Flat3D_Value1 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   16440
            TabIndex        =   71
            Top             =   2940
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label PVKey_Flat3D_Value2 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   16440
            TabIndex        =   70
            Top             =   2700
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label PVKey_Flat3D_Value7 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   14040
            TabIndex        =   69
            Top             =   2700
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label PVKey_Flat3D_Value5 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   14040
            TabIndex        =   68
            Top             =   3180
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label PVKey_Flat3D_Value6 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   14040
            TabIndex        =   67
            Top             =   2940
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label PVKey_Flat3D_Value4 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   14040
            TabIndex        =   66
            Top             =   3435
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Shape PVKey_Flat3D_Color4 
            Height          =   270
            Left            =   13680
            Top             =   3405
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Shape FlatKeyFrame 
            Height          =   1335
            Left            =   11070
            Top             =   2400
            Visible         =   0   'False
            Width           =   4815
         End
         Begin VB.Shape PVKey_Flat3D_Color2 
            FillColor       =   &H00EEE0B5&
            FillStyle       =   0  'Solid
            Height          =   270
            Left            =   16080
            Top             =   2670
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Shape PVKey_Flat3D_Color1 
            FillColor       =   &H00CC9B5A&
            FillStyle       =   0  'Solid
            Height          =   270
            Left            =   16080
            Top             =   2910
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Shape PVKey_Flat3D_Color0 
            FillColor       =   &H006F4928&
            FillStyle       =   0  'Solid
            Height          =   270
            Left            =   16080
            Top             =   3150
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Shape PVKey_Flat3D_Color5 
            FillColor       =   &H0014FFFF&
            FillStyle       =   0  'Solid
            Height          =   270
            Left            =   13680
            Top             =   3150
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Shape PVKey_Flat3D_Color6 
            FillColor       =   &H000096FF&
            FillStyle       =   0  'Solid
            Height          =   270
            Left            =   13680
            Top             =   2910
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Shape PVKey_Flat3D_Color7 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   270
            Left            =   13680
            Top             =   2670
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Label PVKey_Flat3D_Value4_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   14640
            TabIndex        =   65
            Top             =   3435
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label PVKey_Flat3D_Value6_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   14640
            TabIndex        =   64
            Top             =   2940
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label PVKey_Flat3D_Value5_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   14640
            TabIndex        =   63
            Top             =   3180
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label PVKey_Flat3D_Value7_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   14640
            TabIndex        =   62
            Top             =   2700
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label PVKey_Flat3D_Value2_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   17040
            TabIndex        =   61
            Top             =   2700
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label PVKey_Flat3D_Value1_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   17040
            TabIndex        =   60
            Top             =   2940
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label PVKey_Flat3D_Value0_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   17040
            TabIndex        =   59
            Top             =   3180
            Visible         =   0   'False
            Width           =   840
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
            Left            =   2160
            TabIndex        =   58
            Top             =   3840
            Width           =   855
         End
         Begin VB.Label FractileLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "90% - Fractile:"
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
            TabIndex        =   57
            Top             =   3840
            Width           =   1575
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
            Left            =   15240
            TabIndex        =   56
            Top             =   7320
            Width           =   615
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
            Left            =   15240
            TabIndex        =   55
            Top             =   6360
            Width           =   615
         End
         Begin VB.Image ControlsPositionShunt 
            Height          =   11415
            Index           =   0
            Left            =   4800
            Top             =   120
            Width           =   5655
         End
         Begin VB.Image ControlsPositionShunt 
            Height          =   11175
            Index           =   1
            Left            =   10785
            Top             =   360
            Width           =   7935
         End
         Begin VB.Label YScaleLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   53
            Top             =   4200
            Width           =   600
         End
         Begin VB.Line GraphYDevisionLine 
            BorderColor     =   &H00808080&
            Index           =   0
            X1              =   1080
            X2              =   15240
            Y1              =   4320
            Y2              =   4320
         End
         Begin VB.Line GraphYDevisionLine 
            BorderColor     =   &H00808080&
            Index           =   1
            X1              =   1200
            X2              =   15240
            Y1              =   4620
            Y2              =   4620
         End
         Begin VB.Line GraphYDevisionLine 
            BorderColor     =   &H00808080&
            Index           =   2
            X1              =   1200
            X2              =   15240
            Y1              =   4920
            Y2              =   4920
         End
         Begin VB.Line GraphYDevisionLine 
            BorderColor     =   &H00808080&
            Index           =   3
            X1              =   1200
            X2              =   15240
            Y1              =   5235
            Y2              =   5235
         End
         Begin VB.Line GraphYDevisionLine 
            BorderColor     =   &H00808080&
            Index           =   4
            X1              =   1080
            X2              =   15240
            Y1              =   5550
            Y2              =   5550
         End
         Begin VB.Line GraphYDevisionLine 
            BorderColor     =   &H00808080&
            Index           =   5
            X1              =   1200
            X2              =   15240
            Y1              =   5850
            Y2              =   5850
         End
         Begin VB.Line GraphYDevisionLine 
            BorderColor     =   &H00808080&
            Index           =   6
            X1              =   1200
            X2              =   15240
            Y1              =   6165
            Y2              =   6165
         End
         Begin VB.Line GraphYDevisionLine 
            BorderColor     =   &H00808080&
            Index           =   7
            X1              =   1200
            X2              =   15240
            Y1              =   6465
            Y2              =   6465
         End
         Begin VB.Line GraphYDevisionLine 
            BorderColor     =   &H00808080&
            Index           =   8
            X1              =   1080
            X2              =   15240
            Y1              =   6780
            Y2              =   6780
         End
         Begin VB.Line GraphYDevisionLine 
            BorderColor     =   &H00808080&
            Index           =   9
            X1              =   1200
            X2              =   15240
            Y1              =   7080
            Y2              =   7080
         End
         Begin VB.Line GraphYDevisionLine 
            BorderColor     =   &H00808080&
            Index           =   10
            X1              =   1200
            X2              =   15240
            Y1              =   7395
            Y2              =   7395
         End
         Begin VB.Line GraphYDevisionLine 
            BorderColor     =   &H00808080&
            Index           =   11
            X1              =   1200
            X2              =   15240
            Y1              =   7695
            Y2              =   7695
         End
         Begin VB.Line GraphYDevisionLine 
            BorderColor     =   &H00808080&
            Index           =   12
            X1              =   1080
            X2              =   15240
            Y1              =   8010
            Y2              =   8010
         End
         Begin VB.Line GraphYDevisionLine 
            BorderColor     =   &H00808080&
            Index           =   13
            X1              =   1200
            X2              =   15240
            Y1              =   8310
            Y2              =   8310
         End
         Begin VB.Line GraphYDevisionLine 
            BorderColor     =   &H00808080&
            Index           =   14
            X1              =   1200
            X2              =   15240
            Y1              =   8625
            Y2              =   8625
         End
         Begin VB.Line GraphYDevisionLine 
            BorderColor     =   &H00808080&
            Index           =   15
            X1              =   1200
            X2              =   15240
            Y1              =   8925
            Y2              =   8925
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
            Height          =   495
            Left            =   5040
            TabIndex        =   52
            Top             =   3735
            Width           =   10815
            WordWrap        =   -1  'True
         End
         Begin VB.Label CleanFlowSystemsWebAddressLabel 
            Caption         =   "www.redzone.com"
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
            Left            =   7800
            TabIndex        =   51
            Top             =   10440
            Width           =   2295
         End
         Begin VB.Label YScaleLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   480
            TabIndex        =   49
            Top             =   9120
            Width           =   600
         End
         Begin VB.Label YScaleLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   480
            TabIndex        =   48
            Top             =   7890
            Width           =   600
         End
         Begin VB.Label YScaleLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   47
            Top             =   6660
            Width           =   600
         End
         Begin VB.Label YScaleLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   46
            Top             =   5430
            Width           =   600
         End
         Begin VB.Label GraphUnitLabel 
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
            Left            =   15600
            TabIndex        =   45
            Top             =   9360
            Width           =   135
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
            Left            =   14880
            TabIndex        =   44
            Top             =   2160
            Width           =   735
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
            Left            =   14880
            TabIndex        =   43
            Top             =   1920
            Width           =   495
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
            Left            =   12720
            TabIndex        =   42
            Top             =   3240
            Width           =   735
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
            Left            =   12720
            TabIndex        =   41
            Top             =   3000
            Width           =   735
         End
         Begin VB.Label LowerLimitLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "Lower Limit="
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
            Left            =   11040
            TabIndex        =   40
            Top             =   3240
            Width           =   1575
         End
         Begin VB.Label UpperLimitLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "Upper limit ="
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
            Left            =   11040
            TabIndex        =   39
            Top             =   3000
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
            Left            =   11040
            TabIndex        =   38
            Top             =   2760
            Width           =   1575
         End
         Begin VB.Shape UperLimitData 
            Height          =   1095
            Left            =   10920
            Top             =   2640
            Width           =   2655
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
            TabIndex        =   37
            Top             =   2400
            Width           =   1215
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
            TabIndex        =   35
            Top             =   10440
            Width           =   1575
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
            TabIndex        =   34
            Top             =   10440
            Width           =   375
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
            Left            =   14040
            TabIndex        =   33
            Top             =   1440
            Width           =   1275
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
            Left            =   14040
            TabIndex        =   32
            Top             =   2160
            Width           =   735
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
            Left            =   14040
            TabIndex        =   31
            Top             =   1920
            Width           =   735
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
            Left            =   1920
            TabIndex        =   30
            Top             =   1680
            Width           =   6615
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
            Left            =   14040
            TabIndex        =   29
            Top             =   1680
            Width           =   2415
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
            Left            =   7800
            TabIndex        =   28
            Top             =   2160
            Width           =   4335
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
            Left            =   7800
            TabIndex        =   27
            Top             =   1920
            Width           =   4335
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
            Left            =   7800
            TabIndex        =   26
            Top             =   1440
            Width           =   3495
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
            Left            =   1920
            TabIndex        =   25
            Top             =   2160
            Width           =   3135
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
            Left            =   1920
            TabIndex        =   24
            Top             =   1920
            Width           =   3135
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
            Left            =   1920
            TabIndex        =   23
            Top             =   1440
            Width           =   3615
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
            TabIndex        =   22
            Top             =   1440
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
            Left            =   13080
            TabIndex        =   21
            Top             =   1440
            Width           =   855
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
            Left            =   12600
            TabIndex        =   20
            Top             =   2160
            Width           =   1335
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
            Left            =   12600
            TabIndex        =   19
            Top             =   1920
            Width           =   1335
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
            TabIndex        =   18
            Top             =   1680
            Width           =   1335
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
            Left            =   12360
            TabIndex        =   17
            Top             =   1680
            Width           =   1575
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
            Left            =   6120
            TabIndex        =   16
            Top             =   2160
            Width           =   1575
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
            Left            =   6120
            TabIndex        =   15
            Top             =   1920
            Width           =   1575
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
            Left            =   6120
            TabIndex        =   14
            Top             =   1440
            Width           =   1575
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
            TabIndex        =   13
            Top             =   2160
            Width           =   1335
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
            TabIndex        =   12
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Line HeaderBreakLine 
            X1              =   480
            X2              =   15840
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Line FooterBreakLine 
            X1              =   480
            X2              =   15840
            Y1              =   10080
            Y2              =   10080
         End
         Begin VB.Label lblTitle 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "Summary Report"
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
            Left            =   2550
            TabIndex        =   9
            Top             =   360
            Width           =   11295
         End
         Begin VB.Image LogoImage 
            Height          =   855
            Left            =   480
            Stretch         =   -1  'True
            Top             =   360
            Width           =   2175
         End
         Begin VB.Image CLPLogoImage 
            Height          =   690
            Left            =   600
            Picture         =   "PVReportSingle.frx":038A
            Top             =   10200
            Width           =   2010
         End
         Begin VB.Line Line2 
            X1              =   1200
            X2              =   1200
            Y1              =   4320
            Y2              =   9360
         End
         Begin VB.Image GraphXScaleContainer 
            Height          =   495
            Left            =   1200
            Top             =   9240
            Width           =   14025
         End
         Begin VB.Line GraphYDevisionLine 
            Index           =   16
            X1              =   1200
            X2              =   15240
            Y1              =   9240
            Y2              =   9240
         End
         Begin VB.Image GraphContainer 
            Appearance      =   0  'Flat
            Height          =   4950
            Left            =   1200
            Stretch         =   -1  'True
            Top             =   4320
            Width           =   14025
         End
         Begin VB.Label UnitLabel 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            Left            =   240
            TabIndex        =   8
            Top             =   3960
            Width           =   615
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
            TabIndex        =   7
            Top             =   10200
            Width           =   5055
         End
         Begin VB.Label CopyrightLabel 
            Caption         =   "Copyright 2015"
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
            Left            =   12480
            TabIndex        =   50
            Top             =   10200
            Width           =   1455
         End
         Begin VB.Label PageLabel 
            Alignment       =   1  'Right Justify
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
            Left            =   15000
            TabIndex        =   36
            Top             =   10200
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
            Left            =   14640
            TabIndex        =   11
            Top             =   10440
            Width           =   1260
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
            Left            =   13680
            TabIndex        =   10
            Top             =   10440
            Width           =   855
         End
         Begin VB.Shape GraphBackgroundShape 
            BackColor       =   &H00DDDDA2&
            FillColor       =   &H00DDDDA2&
            FillStyle       =   0  'Solid
            Height          =   5535
            Left            =   480
            Tag             =   "Back"
            Top             =   4200
            Width           =   15375
         End
      End
   End
   Begin VB.ComboBox CmboPrinterList 
      Height          =   315
      Left            =   5160
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Text            =   "Select a Printer"
      Top             =   120
      Width           =   3135
   End
   Begin MSComDlg.CommonDialog CommonDialogBox 
      Left            =   120
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Printer Settings"
   End
   Begin VB.PictureBox scrPageScroll 
      Height          =   8655
      Left            =   18840
      ScaleHeight     =   8595
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComDlg.CommonDialog FloatingTextDialog 
      Left            =   120
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image ControlsReport 
      Height          =   720
      Index           =   0
      Left            =   120
      MousePointer    =   1  'Arrow
      Picture         =   "PVReportSingle.frx":0BEF
      Tag             =   "DrawText"
      ToolTipText     =   "Text"
      Top             =   -60
      Width           =   720
   End
   Begin VB.Image ControlsReport 
      Height          =   720
      Index           =   1
      Left            =   960
      MousePointer    =   1  'Arrow
      Picture         =   "PVReportSingle.frx":28B9
      Tag             =   "Move"
      ToolTipText     =   "Move Report or Text"
      Top             =   -60
      Width           =   720
   End
   Begin VB.Image ControlsReport 
      Height          =   720
      Index           =   2
      Left            =   1800
      MousePointer    =   1  'Arrow
      Picture         =   "PVReportSingle.frx":4583
      Tag             =   "ZoomIn"
      ToolTipText     =   "Zoom In"
      Top             =   -60
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsReport 
      Height          =   720
      Index           =   3
      Left            =   4320
      MousePointer    =   1  'Arrow
      Picture         =   "PVReportSingle.frx":624D
      Tag             =   "Print"
      ToolTipText     =   "Print"
      Top             =   -60
      Width           =   720
   End
   Begin VB.Image ControlsReport 
      Height          =   720
      Index           =   4
      Left            =   2640
      MousePointer    =   1  'Arrow
      Picture         =   "PVReportSingle.frx":7F17
      Tag             =   "ZoomOut"
      ToolTipText     =   "Zoom Out"
      Top             =   -60
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsReport 
      Height          =   720
      Index           =   5
      Left            =   3480
      MousePointer    =   1  'Arrow
      Picture         =   "PVReportSingle.frx":9BE1
      Tag             =   "SaveReportToPVD"
      ToolTipText     =   "Save Report To PVD"
      Top             =   -60
      Width           =   720
   End
   Begin VB.Image CloseReport 
      Height          =   480
      Left            =   12960
      MousePointer    =   1  'Arrow
      Picture         =   "PVReportSingle.frx":B8AB
      ToolTipText     =   "Close Report"
      Top             =   90
      Width           =   480
   End
   Begin VB.Image ControlHighlight 
      Height          =   570
      Left            =   10680
      Picture         =   "PVReportSingle.frx":C575
      Top             =   30
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Image ControlHighlightDepressed 
      Height          =   585
      Left            =   11760
      Picture         =   "PVReportSingle.frx":E107
      Top             =   30
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Image ControlsBackPanel 
      Height          =   615
      Left            =   0
      MousePointer    =   1  'Arrow
      Picture         =   "PVReportSingle.frx":FD51
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
Attribute VB_Name = "PVReportSingle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

Public PreviewStartFrame As Long
Public PreviewEndFrame As Long
Dim ShuntY(1) As Single
Dim ShuntX(1) As Single
'Dim ReportMouseDown As Boolean
'Public PrintPreviewAction As String

Sub PrintPVSingleReport()
On Error GoTo Err_Handler
    Dim i As Integer
    
    Call Printer.KillDoc
    

    ScreenDrawingType = 1
    ScreenDrawingOrientation = 1
    
    If ScreenDrawing.ImageGraphState(0).GraphType = "Ovality" Then
        Me.Explination.Tag = "Visible"
    End If

    Call DrawPVGraphsReport
    Call RenderToPrinter.RenderReport(Me, Printer, 1)
    Call DrawPVGraphsReport
    
    Call RenderToPrinter.RenderSingleTextBox(Me.UserTitle, Printer, False) 'PCN4277
    Call RenderToPrinter.RenderSingleTextBox(Me.CommentsTextBox, Printer, True)  'PCN4412
   
    For i = 1 To Me.FloatingText.Count - 1
        Call RenderToPrinter.RenderSingleTextBox(Me.FloatingText(i), Printer, True) 'PCN4412
    Next i
   
   
   
    Call Printer.EndDoc
    ScreenDrawingType = 0
    ScreenDrawingOrientation = 0
    

    
Exit Sub
Err_Handler:
MsgBox Err & "-RS1:" & Error$
End Sub

Private Sub CloseReport_Click()
On Error GoTo Err_Handler

Unload Me
    
Exit Sub
Err_Handler:
    MsgBox Err & "-RS2:" & Error$
End Sub

Private Sub CommentsTextBox_Change()
On Error GoTo Err_Handler
Dim SaveFailed As Boolean

If PVDFileName <> "" Then
    PipelineInfo.Comments = PVReportSingle.CommentsTextBox.text
    Call SaveToFilePipeAndConfigInfo("PipelineInfo", SaveFailed)
End If

Exit Sub
Err_Handler:
   MsgBox Err & "-RS3:" & Error$
End Sub


Private Sub ControlsBackPanel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

If Me.ControlHighlight.Visible Then Me.ControlHighlight.Visible = False
If Me.ControlHighlightDepressed.Visible Then Me.ControlHighlightDepressed.Visible = False
    
Exit Sub
Err_Handler:
    MsgBox Err & "-RS4:" & Error$
End Sub

Private Sub ControlsReport_Click(Index As Integer)
On Error GoTo Err_Handler

Call ExecuteReportButton(Me, Index)
    
Exit Sub
Err_Handler:
    MsgBox Err & "-RS5:" & Error$
End Sub

Private Sub ControlsReport_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


Me.ControlHighlightDepressed.Left = Me.ControlHighlight.Left
Me.ControlHighlightDepressed.Visible = True
Me.ControlHighlight.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-RS6:" & Error$
End Sub

Private Sub ControlsReport_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ControlHighlight.Left = Me.ControlsReport(Index).Left - 100

If Me.ControlHighlight.Visible = False Then Me.ControlHighlight.Visible = True

    
Exit Sub
Err_Handler:
    MsgBox Err & "-RS7:" & Error$
End Sub

Private Sub ControlsReport_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ControlHighlight.Visible = True
Me.ControlHighlightDepressed.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-RS8:" & Error$
End Sub


Private Sub FloatingText_Change(Index As Integer)
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingText_Change(Me, Index)
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RS9:" & Error$
    End Select
End Sub

Private Sub FloatingText_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingText_KeyPress(Me, Index, KeyAscii)
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RS10:" & Error$
    End Select
End Sub

Private Sub FloatingText_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingText_MouseDown(Me, Index, Button, Shift, X, Y)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RS11:" & Error$
    End Select
End Sub

Private Sub FloatingText_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingText_MouseMove(Me, Index, Button, Shift, X, Y)
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RS12:" & Error$
    End Select
End Sub

Private Sub FloatingTextBackgroundColourMenu_Click()
On Error GoTo Err_Handler
    
    Call RenderToPrinter.FloatingTextBackgroundColourMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RS13:" & Error$
    End Select
End Sub

Private Sub FloatingTextDefaultMenu_Click()
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingTextDefaultMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RS14:" & Error$
    End Select
End Sub

Private Sub FloatingTextDeleteAllMenu_Click()
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingTextDeleteAllMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RS15:" & Error$
    End Select
End Sub

Private Sub FloatingTextDeleteMenu_Click()
On Error GoTo Err_Handler
    
    Call RenderToPrinter.FloatingTextDeleteMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RS16:" & Error$
    End Select
End Sub

Private Sub FloatingTextFontMenu_Click()
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingTextFontMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RS17:" & Error$
    End Select
End Sub

Sub IntialiseForm()
On Error GoTo Err_Handler


    Call PositionReportControls
    Call FillOutPrintForm
    Call MarkForPrinting
    
    Call GraphSpecificSettings

    ZOrder 0

'    PVGraphOvalityXScale = 8
'    PVGraphOvalityXOffset = -25
    'If flat render after graph is drawn, otherwise render before graph is drawn

    
    
    ScreenDrawingType = 2
    ScreenDrawingOrientation = 1
    
    Call DrawPVGraphsReport
    Call RenderToPrinter.RenderReport(Me, picReportPage, 1)
    Call DrawPVGraphsReport
    
    ScreenDrawingType = 0
    ScreenDrawingOrientation = 0
    PVReportSingle.CommentsTextBox.Visible = True
    UserTitle.Visible = True

    'Set mouse icon for move
    PrintPreviewAction = "MoveAll"
    Call SetupReportMouseIcon(Me, 108)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RS18:" & Error$
    End Select
    'Resume 'ANT
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

    Me.UserTitle.Font.Charset = LanguageCharset
    Me.FloatingText(0).Font.Charset = LanguageCharset
    Me.CommentsLabel.Font.Charset = LanguageCharset
    Me.CommentsTextBox.Font.Charset = LanguageCharset
    
    Me.CmboPrinterList.Enabled = False
    Call GetPrinterList(Me)
    Set ImageGraphState(6).PictureImage = PrecisionVisionGraph.PrinterReportImage 'PCN4271

    Me.Left = 0
    Me.width = ClearLineProfilerV6.width - 200
    Me.Top = 0
    Me.height = ClearLineProfilerV6.height - 500
'    Me.Show
    
    PageFramePictureBox.Left = 0
    PageFramePictureBox.width = Me.width - PageFramePictureBox.Left
    PageFramePictureBox.height = Me.height
    PageFramePictureBox.Top = 650
    Me.ControlsBackPanel.width = Me.width
    Me.CloseReport.Left = Me.width - 750
    
    Call ConvertLanguage(Me, Language) 'PCN4171

    Call GetPipeDetailsLabels(Me)
    
    Me.PageLabel.Caption = DisplayMessage("Page") & " 1/1"

    Set PrintPreviewForm = Me
    PreviewStartFrame = 0
    PreviewEndFrame = PVDataNoOfLines

    Dim i As Integer
    Me.CmboPrinterList.Enabled = True

    Me.Controls(82).Visible = True
    
    i = Me.Controls.Count
    ReDim OriginalStateVisible(i)
    ReDim OriginalStateTag(i)
    ReDim OriginalStateLeft(i)
    ReDim OriginalStateTop(i)
    ReDim OriginalStateX1(i)
    ReDim OriginalStateY1(i)
    ReDim OriginalStateX2(i)
    ReDim OriginalStateY2(i)
    ReDim OriginalStateWidth(i)
    ReDim OriginalStateHeight(i)
        
    
    For i = 1 To Me.Controls.Count - 1
    
        OriginalStateVisible(i) = Me.Controls(i).Visible
        OriginalStateTag(i) = Me.Controls(i).Tag
        OriginalStateLeft(i) = Me.Controls(i).Left
        OriginalStateTop(i) = Me.Controls(i).Top
        OriginalStateX1(i) = Me.Controls(i).x1
        OriginalStateY1(i) = Me.Controls(i).y1
        OriginalStateX2(i) = Me.Controls(i).x2
        OriginalStateY2(i) = Me.Controls(i).y2
        OriginalStateWidth(i) = Me.Controls(i).width
        OriginalStateHeight(i) = Me.Controls(i).height
        
    Next i

    Call Me.IntialiseForm
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Exit Sub
Err_Handler:
    
    Select Case Err
        Case 438, 387, 393: Resume Next
        Case Else: MsgBox Err & "-RS19:" & Error$
    End Select
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
    PageFramePictureBox.width = Me.width - PageFramePictureBox.Left
    PageFramePictureBox.height = Me.height
    PageFramePictureBox.Top = 720
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RS20:" & Error$
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
Dim i As Integer

On Error GoTo ManualOrientation

Printer.Orientation = vbPRORLandscape


'Detect the CURRENT page setup of the deault printer

ManualOrientationSet: 'PCN4367

Printer.PrintQuality = vbPRPQHigh
CommonDialogBox.Orientation = cdlLandscape

picReportPage.width = Printer.width
picReportPage.height = Printer.height
picReportPage.Left = 50
picReportPage.Top = 50

Call SetControlsPerPaperSize

 'PCN4367 ManualOrientationSet:

lblTitle.Left = (picReportPage.width / 2) - (lblTitle.width / 2)

Exit Sub

ManualOrientation:
On Error GoTo Err_Handler

Dim originalheight
Dim originalwidth

originalheight = Printer.height
originalwidth = Printer.width

'If printer page is allready landscape size then no need to set to landscape size
If originalwidth > originalheight Then GoTo ManualOrientationSet ' PCN4367

'Swap orientation of printer size
Printer.height = originalwidth
Printer.width = originalheight

'Set the printpreview page to reflect printer page size
picReportPage.width = Printer.width
picReportPage.height = Printer.height

GoTo ManualOrientationSet

Exit Sub
Err_Handler:
MsgBox Err & "-RS21:" & Error$

End Sub



Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler
Dim FileSaveFail As Boolean
Dim i As Integer

Call SaveToFilePipeObs(FileSaveFail)

'PCN4531
For i = 1 To Me.FloatingText.Count - 1
     Unload Me.FloatingText(i)
Next i
PipelineDetails.GeneralComments.text = Trim(PipelineInfo.Comments)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RS22:" & Error$
    End Select
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
MsgBox Err & "-RS23:" & Error$
End Sub

Private Sub picReportPage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    If ReportMouseDown Then
        picReportPage.Left = picReportPage.Left + X - ReportMouseX
        picReportPage.Top = picReportPage.Top + Y - ReportMouseY
    End If
Exit Sub
Err_Handler:
MsgBox Err & "-RS24:" & Error$
End Sub

Private Sub picReportPage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    ReportMouseDown = False
Exit Sub
Err_Handler:
MsgBox Err & "-RS25:" & Error$
End Sub

Private Sub FillOutPrintForm()
On Error GoTo Err_Handler
        Dim i As Long
        Dim ControlType As String
        Dim YScaleLabel0 As Double
        Dim YScaleLabel1 As Double
        Dim YScaleLabel2 As Double
        Dim YScaleLabel3 As Double
        Dim YScaleLabel4 As Double
        Dim DisplayUnits As String
        
        Dim LeftLimit As Double
        Dim RightLimit As Double
        Dim Limit1Top As Single
        Dim Limit2Top As Single
        Dim GraphInfoIndex As Integer
        Dim FractileOne As Single
        Dim FractileTwo As Single
        
        Dim FlatKeyData(7) As Single
        
            Dim TunningSettingString As String
    
    TunningSettingsLabel = AutoTune.ColourFeedback.text & "."
    TunningSettingsLabel = TunningSettingsLabel.Caption & "." & AutoTune.PickupLevelFeedback.text
    TunningSettingsLabel = TunningSettingsLabel.Caption & "." & AutoTune.ContrastFeedback.text
    TunningSettingsLabel = TunningSettingsLabel.Caption & "." & AutoTune.LaserWidthFeedback.text
    TunningSettingsLabel = TunningSettingsLabel.Caption & "." & AutoTune.ZoneFeedback.text

    TunningSettingsLabel.Caption = TunningSettingsLabel
        
        
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
        
'        PVGraphOvalityXScale = 8
'        PVGraphOvalityXOffset = -25
        Call PrecisionVisionGraph.GetGeneralPVGraphData(ScreenDrawing.ImageGraphState(0).GraphType)

'PCN4253 ' The me.limitvalue1 and me.limitvalue2 added. '''''''''''''''''''''''''''''''''
    Me.LimitValue1 = Format(ConvertUnitByGraph(PVXScaleLimitPerL, 0, DisplayUnits), "###0.0")
    Me.LimitValue2 = Format(ConvertUnitByGraph(PVXScaleLimitPerR, 0, DisplayUnits), "###0.0")
    LeftLimit = SafeCDbl(Me.LimitValue1) 'PCN4161                                                   '
    RightLimit = SafeCDbl(Me.LimitValue2) 'PCN4161                                                 '
    Limit1Top = PVGraphXScaleTwipFromUnit(0, PVXScaleLimitPerL)                         '
    Limit2Top = PVGraphXScaleTwipFromUnit(0, PVXScaleLimitPerR)                         '
    Limit1Top = Limit1Top * (Me.GraphContainer.height / PrecisionVisionGraph.PVGraphImage(0).width)
    Limit2Top = Limit2Top * (Me.GraphContainer.height / PrecisionVisionGraph.PVGraphImage(0).width)
    Limit1Top = Me.GraphContainer.height - Limit1Top                                    '
    Limit2Top = Me.GraphContainer.height - Limit2Top                                    '
    Me.LimitValue1.Top = Limit1Top + Me.GraphContainer.Top                              '
    Me.LimitValue2.Top = Limit2Top + Me.GraphContainer.Top - Me.LimitValue1.height      '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    With PVReportSingle
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
        
        ''''''''''''''''''''v "Pipe " & removed from this point int he graph title
        .lblTitle.Caption = PrecisionVisionGraph.Label_GraphName(0) & " " & DisplayMessage("Summary Report")
        
        .CommentsTextBox.text = PipelineInfo.Comments
        
        .UnitLabel = PrecisionVisionGraph.PVXScaleLabelUnits(0)
        

    
        'By default make all lable no background
        For i = 0 To .Controls.Count - 1
            ControlType = TypeName(.Controls(i))
            Select Case ControlType
                Case "Label": .Controls(i).BackStyle = 0
            End Select
        Next i
        
        If MeasurementUnits = "mm" Then
            .PipeLengthUnitLabel = "m"
            .PipeDiameterUnitLabel = "mm"
            .GraphUnitLabel = "m" 'PCN4206
        Else
            .PipeLengthUnitLabel = "ft"
            .PipeDiameterUnitLabel = "in"
            .GraphUnitLabel = "ft" 'PCN4206
        End If
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
        
        
Me.PVKey_Flat3D_Value7_Unit(0) = PVGraphsKeyForm.PVKey_Flat3D_Value7_Unit(0) 'PCN4920 (1) is the radial reading
Me.PVKey_Flat3D_Value6_Unit(0) = PVGraphsKeyForm.PVKey_Flat3D_Value6_Unit(0) 'PCN4920 (1) is the radial reading
Me.PVKey_Flat3D_Value5_Unit(0) = PVGraphsKeyForm.PVKey_Flat3D_Value5_Unit(0) 'PCN4920 (1) is the radial reading
Me.PVKey_Flat3D_Value4_Unit(0) = PVGraphsKeyForm.PVKey_Flat3D_Value4_Unit(0) 'PCN4920 (1) is the radial reading
Me.PVKey_Flat3D_Value2_Unit(0) = PVGraphsKeyForm.PVKey_Flat3D_Value2_Unit(0) 'PCN4920 (1) is the radial reading
Me.PVKey_Flat3D_Value1_Unit(0) = PVGraphsKeyForm.PVKey_Flat3D_Value1_Unit(0) 'PCN4920 (1) is the radial reading
Me.PVKey_Flat3D_Value0_Unit(0) = PVGraphsKeyForm.PVKey_Flat3D_Value0_Unit(0) 'PCN4920 (1) is the radial reading

Me.PVKey_Flat3D_Value7_Unit(1) = PVGraphsKeyForm.PVKey_Flat3D_Value7_Unit(1) 'PCN4920 (1) is the radial reading
Me.PVKey_Flat3D_Value6_Unit(1) = PVGraphsKeyForm.PVKey_Flat3D_Value6_Unit(1) 'PCN4920 (1) is the radial reading
Me.PVKey_Flat3D_Value5_Unit(1) = PVGraphsKeyForm.PVKey_Flat3D_Value5_Unit(1) 'PCN4920 (1) is the radial reading
Me.PVKey_Flat3D_Value4_Unit(1) = PVGraphsKeyForm.PVKey_Flat3D_Value4_Unit(1) 'PCN4920 (1) is the radial reading
Me.PVKey_Flat3D_Value2_Unit(1) = PVGraphsKeyForm.PVKey_Flat3D_Value2_Unit(1) 'PCN4920 (1) is the radial reading
Me.PVKey_Flat3D_Value1_Unit(1) = PVGraphsKeyForm.PVKey_Flat3D_Value1_Unit(1) 'PCN4920 (1) is the radial reading
Me.PVKey_Flat3D_Value0_Unit(1) = PVGraphsKeyForm.PVKey_Flat3D_Value0_Unit(1) 'PCN4920 (1) is the radial reading


Me.PVKey_Flat3D_Value7 = PVGraphsKeyForm.PVKey_Flat3D_Value7.text
Me.PVKey_Flat3D_Value6 = PVGraphsKeyForm.PVKey_Flat3D_Value6.Caption
Me.PVKey_Flat3D_Value5 = PVGraphsKeyForm.PVKey_Flat3D_Value5.Caption
Me.PVKey_Flat3D_Value4 = PVGraphsKeyForm.PVKey_Flat3D_Value4.Caption
Me.PVKey_Flat3D_Value2 = PVGraphsKeyForm.PVKey_Flat3D_Value2.Caption
Me.PVKey_Flat3D_Value1 = PVGraphsKeyForm.PVKey_Flat3D_Value1.Caption
Me.PVKey_Flat3D_Value0 = PVGraphsKeyForm.PVKey_Flat3D_Value0.text


        
        

            

    End With

Exit Sub
Err_Handler:
MsgBox Err & "-RS26:" & Error$
'Resume 'ANT
End Sub

Sub GraphSpecificSettings()
On Error GoTo Err_Handler
    Dim i As Integer
    Dim ColourKeysLeftShift As Single
    
    If ScreenDrawing.ImageGraphState(0).GraphType <> "Flat" Then
        Me.LimitLinesLabel.Left = Me.LimitLinesLabel.Left + 2300
        Me.UpperLimitLabel.Left = Me.UpperLimitLabel.Left + 2300
        Me.LowerLimitLabel.Left = Me.LowerLimitLabel.Left + 2300
        Me.LowerLimitData.Left = Me.LowerLimitData.Left + 2300
        Me.UpperLimitData.Left = Me.UpperLimitData.Left + 2300
        Me.CommentsTextBox.width = Me.CommentsTextBox.width + 2300
        Me.UperLimitData.Left = Me.UperLimitData.Left + 2300
 
    End If
    If ScreenDrawing.ImageGraphState(0).GraphType = "Ovality" Then
        PVReportSingle.UnitLabel = PrecisionVisionGraph.PVXScaleLabelUnits(0)
        

        PVReportSingle.Explination.Visible = True
        PVReportSingle.Explination.Tag = ""
        
        Me.LimitValue2.Visible = False: Me.LimitValue2.Tag = "" 'PCN4253
        
        'PCN4324 if ovality display is less than zero, dont display
        For i = 0 To 4
            If Me.YScaleLabel(i).Caption < 0 Then Me.YScaleLabel(i).Visible = False: Me.YScaleLabel(i).Tag = ""
        Next i
        
        Me.RadiusLabel(0).Tag = ""
        Me.RadiusLabel(1).Tag = ""
        
        Me.DiameterLabel(0).Tag = ""
        Me.DiameterLabel(1).Tag = ""
        
        Me.RadiusLabel(0).Visible = False
        Me.RadiusLabel(1).Visible = False
        
        Me.DiameterLabel(0).Visible = False
        Me.DiameterLabel(1).Visible = False
        
    ElseIf ScreenDrawing.ImageGraphState(0).GraphType = "Flat" Then
        With PVReportSingle
        
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
        Me.LowerLimitLabel.Visible = False: Me.LowerLimitLabel.Tag = ""
        Me.UpperLimitData.Visible = False: UpperLimitData.Tag = ""
        Me.LowerLimitData = False: LowerLimitData.Tag = ""
        Me.UperLimitData.Visible = False: Me.UperLimitData.Tag = ""
        
        
        Me.FlatKeyFrame.Visible = True
        Me.FlatKeyFrame.Tag = Visible
        
        
        ColourKeysLeftShift = Me.PVKey_Flat3D_Color7.Left - Me.LimitLinesLabel.Left - 150
        
        
        Me.RadiusLabel(0).Left = Me.RadiusLabel(0).Left - ColourKeysLeftShift
        Me.RadiusLabel(1).Left = Me.RadiusLabel(1).Left - ColourKeysLeftShift
        
        Me.DiameterLabel(0).Left = Me.DiameterLabel(0).Left - ColourKeysLeftShift
        Me.DiameterLabel(1).Left = Me.DiameterLabel(1).Left - ColourKeysLeftShift
        
        Me.RadiusLabel(0).Tag = "Visible"
        Me.RadiusLabel(1).Tag = "Visible"
        
        Me.DiameterLabel(0).Tag = "Visible"
        Me.DiameterLabel(1).Tag = "Visible"
        
        Me.KeyLine.Visible = True
        Me.KeyLine.Tag = "Visible"
        
'        For i = 0 To 4
'            Me.KeyLine(i).x1 = Me.KeyLine(i).x1 - ColourKeysLeftShift
'            Me.KeyLine(i).x2 = Me.KeyLine(i).x2 - ColourKeysLeftShift
'        Next i
        Me.KeyLine.x1 = Me.KeyLine.x1 - ColourKeysLeftShift
        Me.KeyLine.x2 = Me.KeyLine.x2 - ColourKeysLeftShift
        
        Me.PVKey_Flat3D_Color7.Left = Me.PVKey_Flat3D_Color7.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Color6.Left = Me.PVKey_Flat3D_Color6.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Color5.Left = Me.PVKey_Flat3D_Color5.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Color4.Left = Me.PVKey_Flat3D_Color4.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Color2.Left = Me.PVKey_Flat3D_Color2.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Color1.Left = Me.PVKey_Flat3D_Color1.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Color0.Left = Me.PVKey_Flat3D_Color0.Left - ColourKeysLeftShift
        
        Me.PVKey_Flat3D_Value7_Unit(0).Left = Me.PVKey_Flat3D_Value7_Unit(0).Left - ColourKeysLeftShift 'PCN4920 (0) radial (1) diameter
        Me.PVKey_Flat3D_Value6_Unit(0).Left = Me.PVKey_Flat3D_Value6_Unit(0).Left - ColourKeysLeftShift 'PCN4920 (0) radial (1) diameter
        Me.PVKey_Flat3D_Value5_Unit(0).Left = Me.PVKey_Flat3D_Value5_Unit(0).Left - ColourKeysLeftShift 'PCN4920 (0) radial (1) diameter
        Me.PVKey_Flat3D_Value4_Unit(0).Left = Me.PVKey_Flat3D_Value4_Unit(0).Left - ColourKeysLeftShift 'PCN4920 (0) radial (1) diameter
        Me.PVKey_Flat3D_Value2_Unit(0).Left = Me.PVKey_Flat3D_Value2_Unit(0).Left - ColourKeysLeftShift 'PCN4920 (0) radial (1) diameter
        Me.PVKey_Flat3D_Value1_Unit(0).Left = Me.PVKey_Flat3D_Value1_Unit(0).Left - ColourKeysLeftShift 'PCN4920 (0) radial (1) diameter
        Me.PVKey_Flat3D_Value0_Unit(0).Left = Me.PVKey_Flat3D_Value0_Unit(0).Left - ColourKeysLeftShift 'PCN4920 (0) radial (1) diameter
        
        Me.PVKey_Flat3D_Value7_Unit(1).Left = Me.PVKey_Flat3D_Value7_Unit(1).Left - ColourKeysLeftShift 'PCN4920 (0) radial (1) diameter
        Me.PVKey_Flat3D_Value6_Unit(1).Left = Me.PVKey_Flat3D_Value6_Unit(1).Left - ColourKeysLeftShift 'PCN4920 (0) radial (1) diameter
        Me.PVKey_Flat3D_Value5_Unit(1).Left = Me.PVKey_Flat3D_Value5_Unit(1).Left - ColourKeysLeftShift 'PCN4920 (0) radial (1) diameter
        Me.PVKey_Flat3D_Value4_Unit(1).Left = Me.PVKey_Flat3D_Value4_Unit(1).Left - ColourKeysLeftShift 'PCN4920 (0) radial (1) diameter
        Me.PVKey_Flat3D_Value2_Unit(1).Left = Me.PVKey_Flat3D_Value2_Unit(1).Left - ColourKeysLeftShift 'PCN4920 (0) radial (1) diameter
        Me.PVKey_Flat3D_Value1_Unit(1).Left = Me.PVKey_Flat3D_Value1_Unit(1).Left - ColourKeysLeftShift 'PCN4920 (0) radial (1) diameter
        Me.PVKey_Flat3D_Value0_Unit(1).Left = Me.PVKey_Flat3D_Value0_Unit(1).Left - ColourKeysLeftShift 'PCN4920 (0) radial (1) diameter
        
        Me.PVKey_Flat3D_Value7.Left = Me.PVKey_Flat3D_Value7.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value6.Left = Me.PVKey_Flat3D_Value6.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value5.Left = Me.PVKey_Flat3D_Value5.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value4.Left = Me.PVKey_Flat3D_Value4.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value2.Left = Me.PVKey_Flat3D_Value2.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value1.Left = Me.PVKey_Flat3D_Value1.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value0.Left = Me.PVKey_Flat3D_Value0.Left - ColourKeysLeftShift
        
        Me.PVKey_Flat3D_Color7.Visible = True: Me.PVKey_Flat3D_Color7.Tag = "Visible"
        Me.PVKey_Flat3D_Color6.Visible = True: Me.PVKey_Flat3D_Color6.Tag = "Visible"
        Me.PVKey_Flat3D_Color5.Visible = True: Me.PVKey_Flat3D_Color5.Tag = "Visible"
        Me.PVKey_Flat3D_Color4.Visible = True: Me.PVKey_Flat3D_Color4.Tag = "Visible"
        Me.PVKey_Flat3D_Color2.Visible = True: Me.PVKey_Flat3D_Color2.Tag = "Visible"
        Me.PVKey_Flat3D_Color1.Visible = True: Me.PVKey_Flat3D_Color1.Tag = "Visible"
        Me.PVKey_Flat3D_Color0.Visible = True: Me.PVKey_Flat3D_Color0.Tag = "Visible"
        
        Me.PVKey_Flat3D_Value7_Unit(0).Visible = True: Me.PVKey_Flat3D_Value7_Unit(0).Tag = "Visible" 'PCN4920 (0) radial (1) diameter
        Me.PVKey_Flat3D_Value6_Unit(0).Visible = True: Me.PVKey_Flat3D_Value6_Unit(0).Tag = "Visible" 'PCN4920 (0) radial (1) diameter
        Me.PVKey_Flat3D_Value5_Unit(0).Visible = True: Me.PVKey_Flat3D_Value5_Unit(0).Tag = "Visible" 'PCN4920 (0) radial (1) diameter
        Me.PVKey_Flat3D_Value4_Unit(0).Visible = True: Me.PVKey_Flat3D_Value4_Unit(0).Tag = "Visible" 'PCN4920 (0) radial (1) diameter
        Me.PVKey_Flat3D_Value2_Unit(0).Visible = True: Me.PVKey_Flat3D_Value2_Unit(0).Tag = "Visible" 'PCN4920 (0) radial (1) diameter
        Me.PVKey_Flat3D_Value1_Unit(0).Visible = True: Me.PVKey_Flat3D_Value1_Unit(0).Tag = "Visible" 'PCN4920 (0) radial (1) diameter
        Me.PVKey_Flat3D_Value0_Unit(0).Visible = True: Me.PVKey_Flat3D_Value0_Unit(0).Tag = "Visible" 'PCN4920 (0) radial (1) diameter
        
        Me.PVKey_Flat3D_Value7_Unit(1).Visible = True: Me.PVKey_Flat3D_Value7_Unit(1).Tag = "Visible" 'PCN4920 (0) radial (1) diameter
        Me.PVKey_Flat3D_Value6_Unit(1).Visible = True: Me.PVKey_Flat3D_Value6_Unit(1).Tag = "Visible" 'PCN4920 (0) radial (1) diameter
        Me.PVKey_Flat3D_Value5_Unit(1).Visible = True: Me.PVKey_Flat3D_Value5_Unit(1).Tag = "Visible" 'PCN4920 (0) radial (1) diameter
        Me.PVKey_Flat3D_Value4_Unit(1).Visible = True: Me.PVKey_Flat3D_Value4_Unit(1).Tag = "Visible" 'PCN4920 (0) radial (1) diameter
        Me.PVKey_Flat3D_Value2_Unit(1).Visible = True: Me.PVKey_Flat3D_Value2_Unit(1).Tag = "Visible" 'PCN4920 (0) radial (1) diameter
        Me.PVKey_Flat3D_Value1_Unit(1).Visible = True: Me.PVKey_Flat3D_Value1_Unit(1).Tag = "Visible" 'PCN4920 (0) radial (1) diameter
        Me.PVKey_Flat3D_Value0_Unit(1).Visible = True: Me.PVKey_Flat3D_Value0_Unit(1).Tag = "Visible" 'PCN4920 (0) radial (1) diameter
        
        
        Me.PVKey_Flat3D_Value7.Visible = True: Me.PVKey_Flat3D_Value7.Tag = "Visible"
        Me.PVKey_Flat3D_Value6.Visible = True: Me.PVKey_Flat3D_Value6.Tag = "Visible"
        Me.PVKey_Flat3D_Value5.Visible = True: Me.PVKey_Flat3D_Value5.Tag = "Visible"
        Me.PVKey_Flat3D_Value4.Visible = True: Me.PVKey_Flat3D_Value4.Tag = "Visible"
        Me.PVKey_Flat3D_Value2.Visible = True: Me.PVKey_Flat3D_Value2.Tag = "Visible"
        Me.PVKey_Flat3D_Value1.Visible = True: Me.PVKey_Flat3D_Value1.Tag = "Visible"
        Me.PVKey_Flat3D_Value0.Visible = True: Me.PVKey_Flat3D_Value0.Tag = "Visible"
        
        If MedianFlat And PVDFileName <> "" Then
            Me.PVKey_Flat3D_Value7_Unit(0).Visible = False: Me.PVKey_Flat3D_Value7_Unit(0).Tag = "" 'PCN4920 (0) radial (1) diameter
            Me.PVKey_Flat3D_Value6_Unit(0).Visible = False: Me.PVKey_Flat3D_Value6_Unit(0).Tag = "" 'PCN4920 (0) radial (1) diameter
            Me.PVKey_Flat3D_Value5_Unit(0).Visible = False: Me.PVKey_Flat3D_Value5_Unit(0).Tag = "" 'PCN4920 (0) radial (1) diameter
            Me.PVKey_Flat3D_Value4_Unit(0).Visible = False: Me.PVKey_Flat3D_Value4_Unit(0).Tag = "" 'PCN4920 (0) radial (1) diameter
            Me.PVKey_Flat3D_Value2_Unit(0).Visible = False: Me.PVKey_Flat3D_Value2_Unit(0).Tag = "" 'PCN4920 (0) radial (1) diameter
            Me.PVKey_Flat3D_Value1_Unit(0).Visible = False: Me.PVKey_Flat3D_Value1_Unit(0).Tag = "" 'PCN4920 (0) radial (1) diameter
            Me.PVKey_Flat3D_Value0_Unit(0).Visible = False: Me.PVKey_Flat3D_Value0_Unit(0).Tag = "" 'PCN4920 (0) radial (1) diameter
            RadiusLabel(1).Left = RadiusLabel(1).Left + 700
            PVKey_Flat3D_Value0.Left = PVKey_Flat3D_Value0.Left + 700
            PVKey_Flat3D_Value1.Left = PVKey_Flat3D_Value1.Left + 700
            PVKey_Flat3D_Value2.Left = PVKey_Flat3D_Value2.Left + 700
            PVKey_Flat3D_Color0.Left = PVKey_Flat3D_Color0.Left + 700
            PVKey_Flat3D_Color1.Left = PVKey_Flat3D_Color1.Left + 700
            PVKey_Flat3D_Color2.Left = PVKey_Flat3D_Color2.Left + 700
            KeyLine.x1 = KeyLine.x1 + 700
            KeyLine.x2 = KeyLine.x2 + 700
            PVKey_Flat3D_Value4_Unit(1).Left = PVKey_Flat3D_Value4_Unit(1).Left + 700
            PVKey_Flat3D_Value5_Unit(1).Left = PVKey_Flat3D_Value5_Unit(1).Left + 700
            PVKey_Flat3D_Value6_Unit(1).Left = PVKey_Flat3D_Value6_Unit(1).Left + 700
            PVKey_Flat3D_Value7_Unit(1).Left = PVKey_Flat3D_Value7_Unit(1).Left + 700
            DiameterLabel(0).Left = DiameterLabel(0).Left + 700
            PVKey_Flat3D_Value4.Left = PVKey_Flat3D_Value4.Left + 1400
            PVKey_Flat3D_Value5.Left = PVKey_Flat3D_Value5.Left + 1400
            PVKey_Flat3D_Value6.Left = PVKey_Flat3D_Value6.Left + 1400
            PVKey_Flat3D_Value7.Left = PVKey_Flat3D_Value7.Left + 1400
            RadiusLabel(0).Left = RadiusLabel(0).Left + 1400
            PVKey_Flat3D_Color4.Left = PVKey_Flat3D_Color4.Left + 1400
            PVKey_Flat3D_Color5.Left = PVKey_Flat3D_Color5.Left + 1400
            PVKey_Flat3D_Color6.Left = PVKey_Flat3D_Color6.Left + 1400
            PVKey_Flat3D_Color7.Left = PVKey_Flat3D_Color7.Left + 1400
            FlatKeyFrame.Left = FlatKeyFrame.Left + 1400
            FlatKeyFrame.width = FlatKeyFrame.width - 1400
            CommentsTextBox.width = CommentsTextBox.width + 1400
        End If
        
    End If
    
    If MedianFlat And PVDFileName <> "" Then 'PCN5186
        Me.FractileLabel.Visible = False: Me.FractileLabelData.Visible = False 'PCN4296
        Me.FractileLabel.Tag = "": Me.FractileLabelData.Tag = ""
    End If
    
    If ScreenDrawing.ImageGraphState(0).GraphType <> "Ovality" Then
        PVReportSingle.Explination.Visible = False
        PVReportSingle.Explination.Tag = ""
        
        
    End If
    
    
    
'PCN6458     If ScreenDrawing.ImageGraphState(0).GraphType = "Inclination" Then
'PCN6458         Me.LimitLinesLabel.Visible = False: Me.LimitLinesLabel.Tag = ""
'PCN6458 '        Me.UpperLimitLabel.Visible = False: Me.UpperLimitLabel.Tag = ""
'PCN6458 '        Me.LowerLimitLabel.Visible = False: Me.LowerLimitLabel.Tag = ""
'PCN6458 '        Me.UpperLimitData.Visible = False: UpperLimitData.Tag = ""
'PCN6458 '        Me.LowerLimitData = False: LowerLimitData.Tag = ""
'PCN6458         Me.UperLimitData.Visible = False: Me.UperLimitData.Tag = ""
'PCN6458         LimitValue1.Visible = False: Me.LimitValue1.Tag = ""
'PCN6458         Me.LimitValue2.Visible = False: Me.LimitValue2.Tag = ""
'PCN6458         PVReportSingle.Explination.Visible = True
'PCN6458         PVReportSingle.Explination.Tag = "Visible"
'PCN6458         PVReportSingle.Explination.Caption = DisplayMessage("The accuracy is dependant on the precision of the provided sea level heights and is an indicator only")
'PCN6458     End If
    
    
    Dim UserTittleCaption As String
    Dim TittleParse As String
        
    TittleParse = "Summary_" & ScreenDrawing.ImageGraphState(0).GraphType & "="
    Call GetINI_ParameterInfoOnly(MyFile, TittleParse, UserTittleCaption)
    Me.UserTitle.text = UserTittleCaption
    
    If Me.LimitValue2.Top + Me.LimitValue2.height < Me.GraphContainer.Top Then Me.LimitValue2.Visible = False: Me.LimitValue2.Tag = ""
    If Me.LimitValue1.Top > Me.GraphContainer.Top + Me.GraphContainer.height Then Me.LimitValue1.Visible = False: Me.LimitValue1.Tag = ""

    
Exit Sub
Err_Handler:
MsgBox Err & "-RS27:" & Error$
End Sub

Sub MarkForPrinting()
On Error GoTo Err_Handler
Dim i As Integer
Dim ControlType As String

'Draw renderings first that are marked back
For i = 1 To PVReportSingle.Controls.Count - 1
    
    With PVReportSingle.Controls(i)
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
    
Next i
Exit Sub
Err_Handler:
    Select Case Err
        Case 438: Resume No_Container
        Case Else: MsgBox Err & "-RS28:" & Error$
    End Select
End Sub

'===========================
'Declare the Function to select printer
'===========================
 
Function SelectPrinter(ByVal printer_name As String) As Boolean
On Error GoTo Err_Handler
    
    Dim i As Integer
 
    SelectPrinter = False
    For i = 0 To Printers.Count - 1
        ' if the specified printer is found, select it and return True
        If Printers(i).DeviceName = printer_name Then
            Set Printer = Printers(i)
            SelectPrinter = True
            Exit For
        End If
    Next i
    
Exit Function
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RS29:" & Error$
    End Select
End Function
 



Private Sub CmboPrinterList_Click()
On Error GoTo Err_Handler
    Dim i As Integer
    Call SelectPrinter(CmboPrinterList.text)
    If CmboPrinterList.Enabled = True Then
        Me.picReportPage.Cls
        
        For i = 1 To Me.FloatingText.Count - 1 'We dont want the text to be rendered on the preview yet
            Me.FloatingText(i).Visible = False 'when changes printers
        Next i
        
        Call Me.RestoreOriginalState
        Me.CmboPrinterList.Visible = True 'I dont know why but its always false at form load
        Call Me.IntialiseForm
        Me.CmboPrinterList.Enabled = True
        
        For i = 1 To Me.FloatingText.Count - 1 'Even thou we dont want the text to be rendered, we still
            Me.FloatingText(i).Visible = True  'want to be able to see them after printer select changed
        Next i
    End If
    
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RS30:" & Error$
    End Select
End Sub

Sub SetControlsPerPaperSize()
On Error GoTo Err_Handler
Dim i As Integer

If Printer.PaperSize = 1 Then
    ShuntX(1) = -975
    ShuntX(0) = -487
    UserTitle.width = UserTitle.width - 975
    CommentsTextBox.width = CommentsTextBox.width - 975
    GraphBackgroundShape.width = GraphBackgroundShape.width - 975
    GraphXScaleContainer.width = GraphXScaleContainer.width - 975
    GraphContainer.width = GraphContainer.width - 975
    For i = 0 To 16: GraphYDevisionLine(i).x2 = GraphYDevisionLine(i).x2 - 975:   Next i
    HeaderBreakLine.x2 = HeaderBreakLine.x2 - 975
    FooterBreakLine.x2 = FooterBreakLine.x2 - 975
    
    
Else
    ShuntX(0) = 0
    ShuntX(1) = 0
End If

Call ShuntControlsPerContainer



Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RS31:" & Error$
    End Select
End Sub


Sub ShuntControlsPerContainer()
On Error GoTo Err_Handler

Dim i As Integer
Dim S As Integer
Dim ControlType As String

'Draw renderings first that are marked back

For S = 0 To 1 'Loop for the different Shunts
    For i = 1 To PVReportSingle.Controls.Count - 1
        
        With PVReportSingle.Controls(i)
            If TypeName(.Container) = "PictureBox" Then
                If .Tag <> "Container" Then
                    If TypeName(PVReportSingle.Controls(i)) = "TextBox" Or _
                       TypeName(PVReportSingle.Controls(i)) = "Label" Or _
                       TypeName(PVReportSingle.Controls(i)) = "Shape" Or _
                       TypeName(PVReportSingle.Controls(i)) = "Image" Or _
                       TypeName(PVReportSingle.Controls(i)) = "PictureBox" Then
                        If .Left > ControlsPositionShunt(S).Left And _
                           .Left < ControlsPositionShunt(S).Left + ControlsPositionShunt(S).width And _
                           .Top > ControlsPositionShunt(S).Top And _
                           .Top < ControlsPositionShunt(S).Top + ControlsPositionShunt(S).height Then
                            .Top = .Top + ShuntY(S)
                            .Left = .Left + ShuntX(S)
                        End If
                    End If
                    If TypeName(PVReportSingle.Controls(i)) = "Line" Then
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
    Next i
Next S
Exit Sub
Err_Handler:
    Select Case Err
        Case 438: Resume No_Container
        Case Else: MsgBox Err & "-RS32:" & Error$
    End Select
    
End Sub

Sub RestoreOriginalState()
On Error GoTo Err_Handler
    Dim i As Long

    For i = 1 To Me.Controls.Count - 1
        If Me.Controls(i).name <> "FloatingText" Then
            Me.Controls(i).Visible = OriginalStateVisible(i)
            Me.Controls(i).Tag = OriginalStateTag(i)
            Me.Controls(i).Left = OriginalStateLeft(i)
            Me.Controls(i).Top = OriginalStateTop(i)
            Me.Controls(i).x1 = OriginalStateX1(i)
            Me.Controls(i).y1 = OriginalStateY1(i)
            Me.Controls(i).x2 = OriginalStateX2(i)
            Me.Controls(i).y2 = OriginalStateY2(i)
            Me.Controls(i).width = OriginalStateWidth(i)
            Me.Controls(i).height = OriginalStateHeight(i)
        End If
    Next i
Exit Sub
Err_Handler:
    Select Case Err
        Case 9 'Subscript out of range
            'Addition of text will cause this error
            Exit Sub
        Case 438, 382, 383: Resume Next
        Case Else: MsgBox Err & "-RS33:" & Error$
    End Select

End Sub

Private Sub UserTitle_Change()
On Error GoTo Err_Handler

    Dim UserTittleCaption As String
    Dim TittleParse As String
        
    TittleParse = "Summary_" & ScreenDrawing.ImageGraphState(0).GraphType & "="
    Call INI_WriteBack(MyFile, TittleParse, Me.UserTitle.text)
    'Me.UserTitle.text = UserTittleCaption


Exit Sub
Err_Handler:
    Select Case Err

        Case Else: MsgBox Err & "-RS34:" & Error$
    End Select
End Sub
