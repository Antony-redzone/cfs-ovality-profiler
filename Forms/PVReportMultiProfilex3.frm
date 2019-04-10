VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form PVReportMultiProfilex3 
   BorderStyle     =   0  'None
   Caption         =   "Observations Report"
   ClientHeight    =   12600
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   21135
   Icon            =   "PVReportMultiProfilex3.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12600
   ScaleWidth      =   21135
   ShowInTaskbar   =   0   'False
   Begin VB.HScrollBar PageHScroll 
      Height          =   255
      Left            =   8640
      Max             =   2
      Min             =   1
      MousePointer    =   1  'Arrow
      TabIndex        =   53
      Top             =   240
      Value           =   1
      Width           =   1335
   End
   Begin VB.CommandButton ScaleButton05 
      Caption         =   "0.5"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton ScaleButton10 
      Caption         =   "1"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton ScaleButton15 
      Caption         =   "1.5"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
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
   Begin VB.PictureBox PageFramePictureBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   14895
      Left            =   600
      ScaleHeight     =   14865
      ScaleWidth      =   18345
      TabIndex        =   2
      Top             =   720
      Width           =   18375
      Begin VB.PictureBox picReportNextPage 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Index           =   0
         Left            =   120
         ScaleHeight     =   1035
         ScaleWidth      =   11355
         TabIndex        =   52
         Tag             =   "Paper"
         Top             =   12120
         Width           =   11415
      End
      Begin VB.PictureBox picReportPage1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   11904
         Left            =   0
         ScaleHeight     =   11850
         ScaleWidth      =   17955
         TabIndex        =   7
         Tag             =   "Paper"
         Top             =   0
         Width           =   18015
         Begin ClearLineProfiler.PVDGraphControl PVDGraphFlatOvality 
            Height          =   2655
            Left            =   240
            TabIndex        =   69
            Top             =   7440
            Visible         =   0   'False
            Width           =   12375
            _ExtentX        =   21828
            _ExtentY        =   4683
         End
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
            Left            =   2520
            MousePointer    =   3  'I-Beam
            TabIndex        =   56
            Top             =   720
            Visible         =   0   'False
            Width           =   12015
         End
         Begin VB.TextBox FloatingTextDefault 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2640
            TabIndex        =   9
            Text            =   "Default Text Setting"
            Top             =   360
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox FloatingText 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   2640
            TabIndex        =   8
            Top             =   0
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Line KeyLine2 
            Visible         =   0   'False
            X1              =   16920
            X2              =   16920
            Y1              =   7560
            Y2              =   9720
         End
         Begin VB.Line KeyLine1 
            Visible         =   0   'False
            X1              =   16125
            X2              =   16125
            Y1              =   7560
            Y2              =   9720
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
            Left            =   17280
            TabIndex        =   85
            Top             =   7620
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
            Left            =   15600
            TabIndex        =   84
            Top             =   7560
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label PVKey_Flat3D_Value4 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   15480
            TabIndex        =   83
            Top             =   8640
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label PVKey_Flat3D_Value6 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   15480
            TabIndex        =   82
            Top             =   8160
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label PVKey_Flat3D_Value5 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   15480
            TabIndex        =   81
            Top             =   8400
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label PVKey_Flat3D_Value7 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   15480
            TabIndex        =   80
            Top             =   7920
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label PVKey_Flat3D_Value2 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   15480
            TabIndex        =   79
            Top             =   8880
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label PVKey_Flat3D_Value1 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   15480
            TabIndex        =   78
            Top             =   9120
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label PVKey_Flat3D_Value0 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   15480
            TabIndex        =   77
            Top             =   9360
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Shape ColourKeyBoarder 
            Height          =   2175
            Left            =   15000
            Top             =   7560
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.Shape PVKey_Flat3D_Color4 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   270
            Left            =   15120
            Top             =   8625
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Shape PVKey_Flat3D_Color2 
            FillColor       =   &H00EEE0B5&
            FillStyle       =   0  'Solid
            Height          =   270
            Left            =   15120
            Top             =   8865
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Shape PVKey_Flat3D_Color1 
            FillColor       =   &H00CC9B5A&
            FillStyle       =   0  'Solid
            Height          =   270
            Left            =   15120
            Top             =   9105
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Shape PVKey_Flat3D_Color0 
            FillColor       =   &H006F4928&
            FillStyle       =   0  'Solid
            Height          =   270
            Left            =   15120
            Top             =   9345
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Shape PVKey_Flat3D_Color5 
            FillColor       =   &H0014FFFF&
            FillStyle       =   0  'Solid
            Height          =   270
            Left            =   15120
            Top             =   8385
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Shape PVKey_Flat3D_Color6 
            FillColor       =   &H000096FF&
            FillStyle       =   0  'Solid
            Height          =   270
            Left            =   15120
            Top             =   8145
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Shape PVKey_Flat3D_Color7 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   270
            Left            =   15120
            Top             =   7905
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Label PVKey_Flat3D_Value4_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   16200
            TabIndex        =   68
            Top             =   8625
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label PVKey_Flat3D_Value6_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   16200
            TabIndex        =   67
            Top             =   8145
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label PVKey_Flat3D_Value5_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   16200
            TabIndex        =   66
            Top             =   8385
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label PVKey_Flat3D_Value7_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   16200
            TabIndex        =   65
            Top             =   7905
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label PVKey_Flat3D_Value2_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   16200
            TabIndex        =   64
            Top             =   8865
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label PVKey_Flat3D_Value1_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   16200
            TabIndex        =   63
            Top             =   9105
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label PVKey_Flat3D_Value0_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   16200
            TabIndex        =   62
            Top             =   9345
            Visible         =   0   'False
            Width           =   960
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
            Left            =   1440
            TabIndex        =   61
            Top             =   10320
            Width           =   615
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
            Left            =   240
            TabIndex        =   60
            Top             =   10320
            Width           =   1095
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
            Left            =   15120
            TabIndex        =   59
            Top             =   8880
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
            Left            =   15120
            TabIndex        =   58
            Top             =   8040
            Width           =   615
         End
         Begin VB.Label lblTitle 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "Observations Report"
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
            TabIndex        =   57
            Top             =   240
            Width           =   12015
         End
         Begin VB.Image controlspositionshunt 
            Height          =   4455
            Index           =   4
            Left            =   12960
            Tag             =   "Container"
            Top             =   7320
            Width           =   5055
         End
         Begin VB.Image controlspositionshunt 
            Height          =   1455
            Index           =   3
            Left            =   9720
            Tag             =   "Container"
            Top             =   960
            Width           =   6855
         End
         Begin VB.Image controlspositionshunt 
            Height          =   4815
            Index           =   0
            Left            =   11280
            Tag             =   "Container"
            Top             =   2400
            Width           =   5295
         End
         Begin VB.Image controlspositionshunt 
            Height          =   4815
            Index           =   1
            Left            =   5880
            Tag             =   "Container"
            Top             =   2400
            Width           =   10695
         End
         Begin VB.Image controlspositionshunt 
            Height          =   4815
            Index           =   2
            Left            =   480
            Tag             =   "Container"
            Top             =   2400
            Width           =   16095
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
            TabIndex        =   55
            Top             =   10800
            Width           =   5055
         End
         Begin VB.Line LineDown 
            Index           =   2
            X1              =   16320
            X2              =   16320
            Y1              =   6480
            Y2              =   7080
         End
         Begin VB.Line LineDown 
            Index           =   1
            X1              =   10920
            X2              =   10920
            Y1              =   6480
            Y2              =   7080
         End
         Begin VB.Line ProfileSliceRubberBand 
            Index           =   2
            X1              =   12480
            X2              =   16320
            Y1              =   7320
            Y2              =   7080
         End
         Begin VB.Line ProfileSlice 
            Index           =   2
            X1              =   12480
            X2              =   12480
            Y1              =   10080
            Y2              =   7320
         End
         Begin VB.Line ProfileSliceRubberBand 
            Index           =   1
            X1              =   6960
            X2              =   10920
            Y1              =   7320
            Y2              =   7080
         End
         Begin VB.Line ProfileSlice 
            Index           =   1
            X1              =   6960
            X2              =   6960
            Y1              =   10080
            Y2              =   7320
         End
         Begin VB.Line ProfileSlice 
            Index           =   0
            X1              =   1560
            X2              =   1560
            Y1              =   10080
            Y2              =   7320
         End
         Begin VB.Line ProfileSliceRubberBand 
            Index           =   0
            X1              =   1560
            X2              =   5520
            Y1              =   7320
            Y2              =   7080
         End
         Begin VB.Line LineDown 
            Index           =   0
            X1              =   5520
            X2              =   5520
            Y1              =   6480
            Y2              =   7080
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
            Left            =   14640
            TabIndex        =   51
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
            Left            =   14640
            TabIndex        =   50
            Top             =   1920
            Width           =   495
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
            Left            =   13680
            TabIndex        =   49
            Top             =   1440
            Width           =   1335
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
            Left            =   13680
            TabIndex        =   48
            Top             =   2160
            Width           =   855
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
            Left            =   13680
            TabIndex        =   47
            Top             =   1920
            Width           =   855
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
            Left            =   1560
            TabIndex        =   46
            Top             =   1680
            Width           =   5535
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
            Left            =   13680
            TabIndex        =   45
            Top             =   1680
            Width           =   3135
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
            Left            =   7320
            TabIndex        =   44
            Top             =   2160
            Width           =   4695
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
            Left            =   7320
            TabIndex        =   43
            Top             =   1920
            Width           =   4695
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
            Left            =   7320
            TabIndex        =   42
            Top             =   1440
            Width           =   4095
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
            Left            =   1560
            TabIndex        =   41
            Top             =   2160
            Width           =   3375
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
            Left            =   1560
            TabIndex        =   40
            Top             =   1920
            Width           =   3375
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
            Left            =   1560
            TabIndex        =   39
            Top             =   1440
            Width           =   3375
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
            Left            =   240
            TabIndex        =   38
            Top             =   1440
            Width           =   1215
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
            Left            =   12960
            TabIndex        =   37
            Top             =   1440
            Width           =   615
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
            Left            =   12240
            TabIndex        =   36
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
            Left            =   12240
            TabIndex        =   35
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
            Left            =   240
            TabIndex        =   34
            Top             =   1680
            Width           =   1215
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
            Left            =   12240
            TabIndex        =   33
            Top             =   1680
            Width           =   1335
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
            Left            =   5520
            TabIndex        =   32
            Top             =   2160
            Width           =   1695
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
            Left            =   5520
            TabIndex        =   31
            Top             =   1920
            Width           =   1695
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
            Left            =   5520
            TabIndex        =   30
            Top             =   1440
            Width           =   1695
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
            Left            =   240
            TabIndex        =   29
            Top             =   2160
            Width           =   1215
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
            Left            =   240
            TabIndex        =   28
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Shape Shape2 
            Height          =   2055
            Left            =   840
            Top             =   7680
            Width           =   14055
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
            Left            =   240
            TabIndex        =   27
            Top             =   7560
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
            Left            =   240
            TabIndex        =   26
            Top             =   9600
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
            Index           =   2
            Left            =   240
            TabIndex        =   25
            Top             =   8580
            Width           =   600
         End
         Begin VB.Image GraphXScaleContainer 
            Height          =   375
            Left            =   840
            Top             =   9720
            Width           =   14055
         End
         Begin VB.Image GraphContainer 
            Appearance      =   0  'Flat
            Height          =   2055
            Left            =   840
            Stretch         =   -1  'True
            Top             =   7680
            Width           =   14055
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
            Height          =   375
            Left            =   4680
            TabIndex        =   24
            Top             =   10110
            Width           =   11895
            WordWrap        =   -1  'True
         End
         Begin VB.Label GraphUnitLbl 
            Alignment       =   2  'Center
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
            Left            =   15000
            TabIndex        =   23
            Top             =   9840
            Width           =   255
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
            Left            =   240
            TabIndex        =   22
            Top             =   9090
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
            Index           =   1
            Left            =   240
            TabIndex        =   21
            Top             =   8070
            Width           =   600
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
            TabIndex        =   20
            Top             =   7200
            Width           =   495
         End
         Begin VB.Image PVYScaleImage 
            Height          =   3585
            Index           =   2
            Left            =   11040
            Top             =   2520
            Width           =   405
         End
         Begin VB.Image PVXScaleImage 
            Height          =   300
            Index           =   2
            Left            =   11520
            Top             =   6150
            Width           =   4815
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
            Left            =   11160
            TabIndex        =   19
            Top             =   6510
            Width           =   5175
            WordWrap        =   -1  'True
         End
         Begin VB.Image PVProfileImage 
            Height          =   3615
            Index           =   2
            Left            =   11520
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   4815
         End
         Begin VB.Image PVXScaleImage 
            Height          =   300
            Index           =   1
            Left            =   6120
            Top             =   6150
            Width           =   4815
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
            Left            =   5760
            TabIndex        =   18
            Top             =   6510
            Width           =   5175
            WordWrap        =   -1  'True
         End
         Begin VB.Image PVYScaleImage 
            Height          =   3585
            Index           =   1
            Left            =   5640
            Top             =   2520
            Width           =   405
         End
         Begin VB.Image PVProfileImage 
            Height          =   3615
            Index           =   1
            Left            =   6120
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   4815
         End
         Begin VB.Image PVXScaleImage 
            Height          =   300
            Index           =   0
            Left            =   720
            Top             =   6150
            Width           =   4815
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
            Left            =   360
            TabIndex        =   17
            Top             =   6510
            Width           =   5175
            WordWrap        =   -1  'True
         End
         Begin VB.Image PVYScaleImage 
            Height          =   3585
            Index           =   0
            Left            =   240
            Top             =   2520
            Width           =   405
         End
         Begin VB.Image PVProfileImage 
            Height          =   3615
            Index           =   0
            Left            =   720
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   4815
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
            Left            =   6960
            TabIndex        =   16
            Top             =   11040
            Width           =   2295
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
            Left            =   3360
            TabIndex        =   13
            Top             =   11040
            Width           =   1935
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
            Left            =   2880
            TabIndex        =   12
            Top             =   11040
            Width           =   375
         End
         Begin VB.Line HeaderBreakLine 
            X1              =   240
            X2              =   16200
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Line FooterBreakLine 
            X1              =   240
            X2              =   16320
            Y1              =   10560
            Y2              =   10560
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
            Left            =   14880
            TabIndex        =   11
            Top             =   11040
            Width           =   1335
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
            TabIndex        =   10
            Top             =   11040
            Width           =   1215
         End
         Begin VB.Image LogoImage 
            Height          =   855
            Left            =   240
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2175
         End
         Begin VB.Image CLPLogoImage 
            Height          =   690
            Left            =   360
            Picture         =   "PVReportMultiProfilex3.frx":038A
            Top             =   10800
            Width           =   2010
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
         Begin VB.Shape PVProfileGraphBoarder 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   3975
            Index           =   2
            Left            =   11040
            Tag             =   "Back"
            Top             =   2520
            Width           =   5295
         End
         Begin VB.Shape PVProfileGraphBoarder 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   3975
            Index           =   1
            Left            =   5640
            Tag             =   "Back"
            Top             =   2520
            Width           =   5295
         End
         Begin VB.Shape PVProfileGraphBoarder 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   3975
            Index           =   0
            Left            =   240
            Tag             =   "Back"
            Top             =   2520
            Width           =   5295
         End
         Begin VB.Shape GraphBackgroundShape 
            BackColor       =   &H00DDDDA2&
            FillColor       =   &H00DDDDA2&
            FillStyle       =   0  'Solid
            Height          =   2655
            Left            =   240
            Tag             =   "Back"
            Top             =   7440
            Width           =   15015
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
            Left            =   13440
            TabIndex        =   15
            Top             =   10680
            Width           =   1455
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
            TabIndex        =   14
            Top             =   10680
            Width           =   855
         End
         Begin VB.Label PVKey_Flat3D_Value0_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   17160
            TabIndex        =   76
            Top             =   9360
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label PVKey_Flat3D_Value1_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   17160
            TabIndex        =   75
            Top             =   9120
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label PVKey_Flat3D_Value2_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   17160
            TabIndex        =   74
            Top             =   8880
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label PVKey_Flat3D_Value7_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   17160
            TabIndex        =   73
            Top             =   7920
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label PVKey_Flat3D_Value5_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   17160
            TabIndex        =   72
            Top             =   8400
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label PVKey_Flat3D_Value6_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   17160
            TabIndex        =   71
            Top             =   8160
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label PVKey_Flat3D_Value4_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   17160
            TabIndex        =   70
            Top             =   8640
            Visible         =   0   'False
            Width           =   960
         End
      End
   End
   Begin VB.CommandButton AllFramesButtons 
      Caption         =   "All Frames"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox CmboPrinterList 
      Height          =   315
      Left            =   5160
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Text            =   "Select a Printer"
      Top             =   120
      Width           =   3135
   End
   Begin MSComDlg.CommonDialog FloatingTextDialog 
      Left            =   120
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
   Begin VB.Image CloseReport 
      Height          =   480
      Left            =   12240
      MousePointer    =   1  'Arrow
      Picture         =   "PVReportMultiProfilex3.frx":0BEF
      ToolTipText     =   "Close Report"
      Top             =   90
      Width           =   480
   End
   Begin VB.Image ControlsReport 
      Height          =   720
      Index           =   5
      Left            =   3480
      MousePointer    =   1  'Arrow
      Picture         =   "PVReportMultiProfilex3.frx":18B9
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
      Picture         =   "PVReportMultiProfilex3.frx":3583
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
      Picture         =   "PVReportMultiProfilex3.frx":524D
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
      Picture         =   "PVReportMultiProfilex3.frx":6F17
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
      Picture         =   "PVReportMultiProfilex3.frx":8BE1
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
      Picture         =   "PVReportMultiProfilex3.frx":A8AB
      Tag             =   "DrawText"
      ToolTipText     =   "Text"
      Top             =   -60
      Width           =   720
   End
   Begin VB.Label NoOfPagesLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Page 1 of 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8640
      MousePointer    =   1  'Arrow
      TabIndex        =   54
      Top             =   30
      Width           =   1335
   End
   Begin VB.Image ControlHighlight 
      Height          =   570
      Left            =   9960
      Picture         =   "PVReportMultiProfilex3.frx":C575
      Top             =   30
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Image ControlHighlightDepressed 
      Height          =   585
      Left            =   11040
      Picture         =   "PVReportMultiProfilex3.frx":E107
      Top             =   30
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Image ControlsBackPanel 
      Height          =   615
      Left            =   0
      MousePointer    =   1  'Arrow
      Picture         =   "PVReportMultiProfilex3.frx":FD51
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12615
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
Attribute VB_Name = "PVReportMultiProfilex3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim ReportMouseDown As Boolean
Public PreviewStartFrame As Long
Public PreviewEndFrame As Long
'Public PrintPreviewAction As String
Public picReportPage As PictureBox
Public NumberOfExtraPages As Integer
Private NumOfObs As Integer
Dim ShuntY(4) As Single
Dim ShuntX(4) As Single



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
        Case Else: MsgBox Err & "-RMP1:" & Error$
    End Select
End Sub


Sub PrintMultiProfileReport()
On Error GoTo Err_Handler
    Call Printer.EndDoc
   
'    Printer.PaperSize = vbPRPSA4
    Printer.Orientation = vbPRORLandscape
    Printer.PrintQuality = vbPRPQHigh
    CommonDialogBox.Orientation = cdlLandscape
   
   
   
    Dim Pg As Integer
    Dim i As Integer
    
    PreviewStartFrame = 1
    PreviewEndFrame = PVDataNoOfLines
    ScreenDrawingType = 1
    ScreenDrawingOrientation = 1
    
    For i = 1 To Me.FloatingText.Count - 1
        Me.FloatingText(i).Visible = False
        Me.FloatingText(i).Tag = ""
    Next i
    
    Set picReportPage = picReportPage1
    'Call RestoreOriginalState 'Restore was never setup, now it is its shouldn't be restored here
    Call PrintPage(picReportPage1, 1)
    For i = 1 To Me.FloatingText.Count - 1
        If PVReportMultiProfilex3.FloatingText(i).Container.name = "picReportPage1" Then
            Call RenderToPrinter.RenderSingleTextBox(Me.FloatingText(i), Printer, True)
        End If
    Next i
    
    For Pg = 1 To NumberOfExtraPages
        ScreenDrawingType = 1
        ScreenDrawingOrientation = 1
    
        Printer.NewPage
        Set picReportPage = picReportNextPage(Pg)
        'Call RestoreOriginalState 'same as above
        Call PrintPage(picReportNextPage(Pg), Pg + 1)
        For i = 1 To PVReportMultiProfilex3.FloatingText.Count - 1
            If PVReportMultiProfilex3.FloatingText(i).Container.name = "picReportNextPage" Then
                If PVReportMultiProfilex3.FloatingText(i).Container.Index = Pg Then
                    Call RenderToPrinter.RenderSingleTextBox(Me.FloatingText(i), Printer, True)
                End If
            End If
        Next i
    Next Pg
    

    picReportPage1.Visible = True
    For Pg = 1 To NumberOfExtraPages
        picReportNextPage(Pg).Visible = True
    Next Pg
    
    Call Printer.EndDoc
    Me.PageFramePictureBox.Visible = True
    For i = 1 To Me.FloatingText.Count - 1
        Me.FloatingText(i).Visible = True
    Next i
    



Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP2:" & Error$
    End Select
End Sub
Sub PrintPage(Page As Picture, ByVal PageNo As Integer)
On Error GoTo Err_Handler
    Dim ReportNumber As Integer
    Dim ObsNo As Integer
    Dim ProfilerSlice As Single
    Dim ProfileSlicePosition As Single
    
    RenderScale = 1

    PageLabel.Caption = "Page: " & PageNo & "/" & NumberOfExtraPages + 1
    For ReportNumber = 0 To 2
        ObsNo = ((PageNo - 1) * 3) + ReportNumber + 1
        If ObsNo > NumOfObs Then
            ProfileSliceRubberBand(ReportNumber).Visible = False
            ProfileSliceRubberBand(ReportNumber).Tag = ""
            ProfileSlice(ReportNumber).Visible = False
            ProfileSlice(ReportNumber).Tag = ""
            ObservationsLabel(ReportNumber).Visible = False
            ObservationsLabel(ReportNumber).Tag = ""
            PVProfileGraphBoarder(ReportNumber).Visible = False
            PVProfileGraphBoarder(ReportNumber).Tag = ""
            LineDown(ReportNumber).Visible = False
            LineDown(ReportNumber).Tag = ""
            PVProfileImage(ReportNumber).Visible = False
            PVProfileImage(ReportNumber).Tag = "" 'PCN4414 was False suppose to be ""
            
        Else
            'If PrecisionVisionGraph.ObsPictureStorage(ObsNo).Tag = SnapShot Then
                  PVReportMultiProfilex3.PVProfileImage(ReportNumber).Picture = PrecisionVisionGraph.ObsPictureStorage(ObsNo).Picture
            'End If
            
            ProfileSliceRubberBand(ReportNumber).Tag = "Visible" 'PCN4414
            ProfileSlice(ReportNumber).Tag = "Visible" 'PCN4414
            ObservationsLabel(ReportNumber).Tag = "Visible" 'PCN4414
            PVProfileGraphBoarder(ReportNumber).Tag = "Back" 'PCN4414 inportant this one sent to back
            LineDown(ReportNumber).Tag = "Visible" 'PCN4414
            PVProfileImage(ReportNumber).Tag = "Visible" 'PCN4414
            
            ProfileSlicePosition = GraphContainer.width / (CSng(PreviewEndFrame - PreviewStartFrame))
            ProfileSlicePosition = (PipeObservations(ObsNo).PipeObsFrameNo - PreviewStartFrame) * ProfileSlicePosition
            If ProfileSlicePosition < 0 Then ProfileSlicePosition = -600
            If ProfileSlicePosition > GraphContainer.width Then ProfileSlicePosition = GraphContainer.width + 300
            
            ProfileSlicePosition = ProfileSlicePosition + GraphContainer.Left
            
            ProfileSliceRubberBand(ReportNumber).x1 = ProfileSlicePosition
            ProfileSlice(ReportNumber).x1 = ProfileSlicePosition
            ProfileSlice(ReportNumber).x2 = ProfileSlicePosition
            
            ObservationsLabel(ReportNumber) = Observations.RemoveCrLf(PipeObservations(ObsNo).PipeObs)  'PCN4389
        End If
    Next ReportNumber
    
        
'    PreviewStartFrame = 1
'    PreviewEndFrame = PVDataNoOfLines
    'Me.UserTitle.Tag = "Visible" PCN4277
    If ScreenDrawing.ImageGraphState(0).GraphType = "Ovality" Then
        Me.Explination.Tag = "Visible"
    End If
    
    
        If FlatOvality Then
         Call PVDGraphFlatOvality.PrintGraph(Printer, 1, Me.PVDGraphFlatOvality.Left, Me.PVDGraphFlatOvality.Top)
    Else
        Call DrawPVGraphsReport
        Call RenderToPrinter.RenderReport(Me, Printer, 1)
        Call RenderToPrinter.RenderSingleTextBox(Me.UserTitle, Printer, False) 'PCN4277
        Call DrawPVGraphsReport
    End If
    

    
    For ReportNumber = 0 To 2
        ObsNo = ((PageNo - 1) * 3) + ReportNumber + 1
        If ObsNo > NumOfObs Then Exit For
        Set ScreenDrawing.ReportDummyGraphImage = PVProfileImage(ReportNumber)
'        If PrecisionVisionGraph.ObsPictureStorage(ObsNo).Tag <> SnapShot Then
'            Call DrawProfilesStartToFinish(Printer, True, PipeObservations(ObsNo).PipeObsFrameNo) 'PCN3691
            DrawSF = ReportDummyGraphImage.width / ClearLineScreen.MainScreen.width
            Set ScreenDrawing.ReportGraphImageX = PVReportMultiProfilex3.PVXScaleImage(ReportNumber)
            Set ScreenDrawing.ReportGraphImageY = PVReportMultiProfilex3.PVYScaleImage(ReportNumber)
            Call DrawMainScale(Printer)
'            Call ScreenDrawing.OnscreenTextDrawAll(Printer)
'        End If
        
    Next ReportNumber
    
    ScreenDrawingType = 0
    ScreenDrawingOrientation = 0
    ScreenDrawingType = 0
    ScreenDrawingOrientation = 0
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP3:" & Error$
    End Select
End Sub




Private Sub CloseReport_Click()
On Error GoTo Err_Handler

Unload Me
    
Exit Sub
Err_Handler:
    MsgBox Err & "-RMP4:" & Error$
End Sub


Private Sub ControlsBackPanel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


If Me.ControlHighlight.Visible Then Me.ControlHighlight.Visible = False
If Me.ControlHighlightDepressed.Visible Then Me.ControlHighlightDepressed.Visible = False
    
Exit Sub
Err_Handler:
    MsgBox Err & "-RMP5:" & Error$
End Sub

Private Sub ControlsReport_Click(Index As Integer)
On Error GoTo Err_Handler
Dim i

For i = 1 To NumberOfExtraPages
    picReportNextPage(i).AutoRedraw = False
Next i


Call ExecuteReportButton(Me, Index)

    
Exit Sub
Err_Handler:
    MsgBox Err & "-RMP6:" & Error$
End Sub

Private Sub ControlsReport_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


Me.ControlHighlightDepressed.Left = Me.ControlHighlight.Left
Me.ControlHighlightDepressed.Visible = True
Me.ControlHighlight.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-RMP7:" & Error$
End Sub

Private Sub ControlsReport_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ControlHighlight.Left = Me.ControlsReport(Index).Left - 100

If Me.ControlHighlight.Visible = False Then Me.ControlHighlight.Visible = True

    
Exit Sub
Err_Handler:
    MsgBox Err & "-RMP8:" & Error$
End Sub

Private Sub ControlsReport_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ControlHighlight.Visible = True
Me.ControlHighlightDepressed.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-RMP9:" & Error$
End Sub

Private Sub FloatingText_Change(Index As Integer)
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingText_Change(Me, Index)
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP10:" & Error$
    End Select
End Sub

Private Sub FloatingText_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingText_KeyPress(Me, Index, KeyAscii)
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP11:" & Error$
    End Select
End Sub

Private Sub FloatingText_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingText_MouseDown(Me, Index, Button, Shift, X, Y)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP12:" & Error$
    End Select
End Sub

Private Sub FloatingText_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingText_MouseMove(Me, Index, Button, Shift, X, Y)
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP13:" & Error$
    End Select
End Sub

Private Sub FloatingTextBackgroundColourMenu_Click()
On Error GoTo Err_Handler
    
    Call RenderToPrinter.FloatingTextBackgroundColourMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP14:" & Error$
    End Select
End Sub


Private Sub FloatingTextDefaultMenu_Click()
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingTextDefaultMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP15:" & Error$
    End Select
End Sub

Private Sub FloatingTextDeleteAllMenu_Click()
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingTextDeleteAllMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP16:" & Error$
    End Select
End Sub

Private Sub FloatingTextDeleteMenu_Click()
On Error GoTo Err_Handler
    
    Call RenderToPrinter.FloatingTextDeleteMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP17:" & Error$
    End Select
End Sub

Private Sub FloatingTextFontMenu_Click()
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingTextFontMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP18:" & Error$
    End Select
End Sub

Sub InitialiseForm()
On Error GoTo Err_Handler

    Call PositionReportControls
    Call FillOutPrintForm
    Call MarkForPrinting
    
    Call GraphSpecificSettings
    

    If MeasurementUnits = "mm" Then
        Me.PipeLengthUnitLabel = "m"
        Me.PipeDiameterUnitLabel = "mm"
        Me.GraphUnitLbl = "m"
    Else
        Me.PipeLengthUnitLabel = "ft"
        Me.PipeDiameterUnitLabel = "in"
        Me.GraphUnitLbl = "ft"
    End If

    

   Call Me.RenderForm
    'Set mouse icon for move
    PrintPreviewAction = "MoveAll"
    Call SetupReportMouseIcon(Me, 108)
    Me.PageFramePictureBox.Visible = True


Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP19:" & Error$
    
    End Select
End Sub



Private Sub Form_Load()
On Error GoTo Err_Handler

    Dim i As Long
    Dim Pg As Integer
    
    Me.UserTitle.Font.Charset = LanguageCharset
    Me.FloatingText(0).Font.Charset = LanguageCharset
    
    Me.CmboPrinterList.Enabled = False
    Me.PageFramePictureBox.Visible = False

    Me.picReportPage1.Font.Charset = LanguageCharset
    
    Me.Left = 0
    Me.width = ClearLineProfilerV6.width - 200
    Me.Top = 0
    Me.height = ClearLineProfilerV6.height - 500
    
    Me.lblTitle.Font.Charset = LanguageCharset
    
    Call GetPipeDetailsLabels(Me)
    
    If CLPScreenMode = ThreeD Then
            Call ClearLineScreen.D3D_CaptureImage(LocToSave & "Snapshot.bmp")
            PVReportMultiProfilex3.PVProfileImage(0).Picture = LoadPicture(LocToSave & "Snapshot.bmp")
    End If

    Call GetPrinterList(Me)
    Set ImageGraphState(6).PictureImage = PrecisionVisionGraph.PrinterReportImage 'PCN4271
    Me.Show
    
    
    PageFramePictureBox.Left = 0
    PageFramePictureBox.width = Me.width - PageFramePictureBox.Left
    PageFramePictureBox.height = Me.height
    PageFramePictureBox.Top = 650
    Me.ControlsBackPanel.width = Me.width
    Me.CloseReport.Left = Me.width - 750
    
    Call ConvertLanguage(Me, Language) 'PCN4171
'    Me.PageLabel.Caption = DisplayMessage("Page") & " 1/1"
    
    RenderScale = 1
    Set PrintPreviewForm = Me
    
    PreviewStartFrame = 1
    PreviewEndFrame = PVDataNoOfLines
    
    NumOfObs = PrecisionVisionGraph.ObsPictureStorage.Count - 1
    NumberOfExtraPages = (NumOfObs - 2) / 3
    If NumberOfExtraPages < 0 Then NumberOfExtraPages = 0
    
    'Set up the total number of pages
    PageHScroll.Max = 1 + NumberOfExtraPages
    Call SetTheNoOfPages
    
    For Pg = 1 To NumberOfExtraPages
        Load picReportNextPage(Pg)
        picReportNextPage(Pg).Visible = True
        picReportNextPage(Pg).Tag = "Paper"
    Next Pg
        
    
    picReportNextPage(0).Visible = False
    picReportNextPage(0).Enabled = False
    picReportNextPage(0).Tag = ""
    
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
    Me.CmboPrinterList.Enabled = True
    Call Me.InitialiseForm
   
Exit Sub
Err_Handler:
    Select Case Err
        Case 438, 382, 383, 393: Resume Next
        Case Else: MsgBox Err & "-RMP20:" & Error$
    
    End Select

End Sub

Private Sub RenderPage(Page As PictureBox, ByVal PageNo As Integer)
On Error GoTo Err_Handler
    Dim ReportNumber As Integer
    Dim ObsNo As Integer
    Dim ProfileSlicePosition As Single
    
    Page.Cls
    Page.width = Printer.width * RenderScale
    Page.height = Printer.height * RenderScale
    
    PageLabel.Caption = "Page: " & PageNo & "/" & NumberOfExtraPages + 1
    For ReportNumber = 0 To 2
        ObsNo = ((PageNo - 1) * 3) + ReportNumber + 1
        If ObsNo > NumOfObs Then
            ProfileSliceRubberBand(ReportNumber).Visible = False
            ProfileSliceRubberBand(ReportNumber).Tag = ""
            ProfileSlice(ReportNumber).Visible = False
            ProfileSlice(ReportNumber).Tag = ""
            ObservationsLabel(ReportNumber).Visible = False
            ObservationsLabel(ReportNumber).Tag = ""
            PVProfileGraphBoarder(ReportNumber).Visible = False
            PVProfileGraphBoarder(ReportNumber).Tag = ""
            LineDown(ReportNumber).Visible = False
            LineDown(ReportNumber).Tag = ""
            PVProfileImage(ReportNumber).Visible = False
            PVProfileImage(ReportNumber).Tag = ""
            
        Else
            'If PrecisionVisionGraph.ObsPictureStorage(ObsNo).Tag = SnapShot Then
                  PVReportMultiProfilex3.PVProfileImage(ReportNumber).Picture = PrecisionVisionGraph.ObsPictureStorage(ObsNo).Picture
            'End If
            'PVProfileImage(ReportNumber).Tag = "Visible"
            
            ProfileSlicePosition = GraphContainer.width / (CSng(PreviewEndFrame - PreviewStartFrame))
            ProfileSlicePosition = (PipeObservations(ObsNo).PipeObsFrameNo - PreviewStartFrame) * ProfileSlicePosition
            If ProfileSlicePosition < 0 Then ProfileSlicePosition = -600
            If ProfileSlicePosition > GraphContainer.width Then ProfileSlicePosition = GraphContainer.width + 300
            
            ProfileSlicePosition = ProfileSlicePosition + GraphContainer.Left
            
            ProfileSliceRubberBand(ReportNumber).x1 = ProfileSlicePosition
            ProfileSlice(ReportNumber).x1 = ProfileSlicePosition
            ProfileSlice(ReportNumber).x2 = ProfileSlicePosition
            
            ObservationsLabel(ReportNumber) = Observations.RemoveCrLf(PipeObservations(ObsNo).PipeObs)  'PCN4389
        End If
    Next ReportNumber
    
    ScreenDrawingType = 2
    ScreenDrawingOrientation = 1

    If FlatOvality Then
         Call PVDGraphFlatOvality.PrintGraph(Printer, 1, Me.PVDGraphFlatOvality.Left, Me.PVDGraphFlatOvality.Top)
    Else
        Call DrawPVGraphsReport
    End If
    
    Call RenderToPrinter.RenderReport(Me, Page, RenderScale)
    Call DrawPVGraphsReport
    
    For ReportNumber = 0 To 2
        ObsNo = ((PageNo - 1) * 3) + ReportNumber + 1
        If ObsNo > NumOfObs Then Exit For
        Set ScreenDrawing.ReportDummyGraphImage = PVProfileImage(ReportNumber)
            DrawSF = ReportDummyGraphImage.width / ClearLineScreen.MainScreen.width
            Set ScreenDrawing.ReportGraphImageX = PVReportMultiProfilex3.PVXScaleImage(ReportNumber)
            Set ScreenDrawing.ReportGraphImageY = PVReportMultiProfilex3.PVYScaleImage(ReportNumber)
            Call DrawMainScale(Page)
    Next ReportNumber
    ScreenDrawingType = 0
    ScreenDrawingOrientation = 0

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP21:" & Error$
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
Dim Pg As Integer

On Error GoTo ManualOrientation

Printer.Orientation = vbPRORLandscape


ManualOrientationSet:

Printer.PrintQuality = vbPRPQHigh
CommonDialogBox.Orientation = cdlLandscape

'Detect the CURRENT page setup of the deault printer
picReportPage1.width = Printer.width
picReportPage1.height = Printer.height
picReportPage1.Left = 50
picReportPage1.Top = 50

Call SetControlsPerPaperSize

If NumberOfExtraPages >= 1 Then
    picReportNextPage(1).width = picReportPage1.width
    picReportNextPage(1).width = picReportPage1.width
    picReportNextPage(1).height = picReportPage1.height
    picReportNextPage(1).Left = picReportPage1.Left
    picReportNextPage(1).Top = picReportPage1.height + 100
End If
    
For Pg = 2 To NumberOfExtraPages
    picReportNextPage(Pg).width = picReportPage1.width
    picReportNextPage(Pg).width = picReportPage1.width
    picReportNextPage(Pg).height = picReportPage1.height
    picReportNextPage(Pg).Left = picReportPage1.Left
    picReportNextPage(Pg).Top = picReportNextPage(Pg - 1).height + picReportNextPage(Pg - 1).Top + 100
Next Pg

Exit Sub
ManualOrientation:
On Error GoTo Err_Handler

Dim originalheight
Dim originalwidth

originalheight = Printer.height
originalwidth = Printer.width

If originalwidth > originalheight Then GoTo ManualOrientationSet

Printer.height = originalwidth
Printer.width = originalheight

GoTo ManualOrientationSet
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP22:" & Error$
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
        Case Else: MsgBox Err & "-RMP23:" & Error$
    End Select
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

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP24:" & Error$
    End Select
End Sub

Private Sub PageFramePictureBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


If Me.ControlHighlight.Visible Then Me.ControlHighlight.Visible = False
If Me.ControlHighlightDepressed.Visible Then Me.ControlHighlightDepressed.Visible = False
    
Exit Sub
Err_Handler:
    MsgBox Err & "-RMP25:" & Error$
End Sub

Private Sub PageHScroll_Change()
On Error GoTo Err_Handler
Dim ReportPageSpacing As Integer
Dim PageIndex As Integer

If Me.ControlHighlight.Visible Then Me.ControlHighlight.Visible = False
If Me.ControlHighlightDepressed.Visible Then Me.ControlHighlightDepressed.Visible = False

If NumberOfExtraPages = 0 Then Exit Sub

'ReportPageSpacing = picReportNextPage(1).Top - picReportPage1.Top

picReportPage1.Top = 50 - ((PageHScroll.value - 1) * (picReportPage1.height + 100))
picReportPage1.Left = 50

For PageIndex = 1 To NumberOfExtraPages
    picReportNextPage(PageIndex).Top = picReportPage1.Top + (PageIndex * (picReportPage1.height + 100))
    picReportNextPage(PageIndex).Left = picReportPage1.Left
Next PageIndex
Call SetTheNoOfPages


Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP26:" & Error$
    End Select
End Sub

Private Sub picReportNextPage_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Call ReportPageMouseDown(Me, picReportNextPage(Index), Button, Shift, X, Y) 'PCN4193

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP27:" & Error$
    End Select
End Sub

Private Sub picReportNextPage_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
Dim Pg As Integer
    If ReportMouseDown Then
        picReportPage1.Left = picReportPage1.Left + X - ReportMouseX
        picReportPage1.Top = picReportPage1.Top + Y - ReportMouseY
        For Pg = 1 To NumberOfExtraPages
            picReportNextPage(Pg).Left = picReportNextPage(Pg).Left + X - ReportMouseX
            picReportNextPage(Pg).Top = picReportNextPage(Pg).Top + Y - ReportMouseY
        Next Pg
    End If
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP28:" & Error$
    End Select
End Sub

Private Sub picReportNextPage_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    
ReportMouseDown = False
If PrintPreviewAction = "MoveAll" Then
    Call SetupReportMouseIcon(PVReportMultiProfilex3, 108)
End If

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP29:" & Error$
    End Select
End Sub

Private Sub picReportPage1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    
Call ReportPageMouseDown(Me, picReportPage1, Button, Shift, X, Y) 'PCN4193
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP30:" & Error$
    End Select

End Sub

Private Sub picReportPage1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
Dim Pg As Integer
    If ReportMouseDown Then
        picReportPage1.Left = picReportPage1.Left + X - ReportMouseX
        picReportPage1.Top = picReportPage1.Top + Y - ReportMouseY
        For Pg = 1 To NumberOfExtraPages
            picReportNextPage(Pg).Left = picReportNextPage(Pg).Left + X - ReportMouseX
            picReportNextPage(Pg).Top = picReportNextPage(Pg).Top + Y - ReportMouseY
        Next Pg
    End If
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP31:" & Error$
    End Select
End Sub

Private Sub picReportPage1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    
ReportMouseDown = False
If PrintPreviewAction = "MoveAll" Then
    Call SetupReportMouseIcon(PVReportMultiProfilex3, 108)
End If

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP32:" & Error$
    End Select
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
    Dim CapacityVal As Double
    Dim OvalityVal As Double
    Dim DeltaMinVal As Double
    Dim DeltaMaxVal As Double
    Dim XDiameterVal As Double
    Dim YDiameterVal As Double
    Dim DiameterMedianVal As Double

    
    Dim LeftLimit As Double
    Dim RightLimit As Double
    Dim Limit1Top As Single
    Dim Limit2Top As Single
    
    Dim GraphInfoIndex As Integer
    Dim FractileOne As Single
    Dim FractileTwo As Single
    
    YScaleLabel0 = PVGraphXScaleUnitFromTwip(0, PrecisionVisionGraph.PVXScale(0).width / 4 * 4)
    YScaleLabel1 = PVGraphXScaleUnitFromTwip(0, PrecisionVisionGraph.PVXScale(0).width / 4 * 3)
    YScaleLabel2 = PVGraphXScaleUnitFromTwip(0, PrecisionVisionGraph.PVXScale(0).width / 4 * 2)
    YScaleLabel3 = PVGraphXScaleUnitFromTwip(0, PrecisionVisionGraph.PVXScale(0).width / 4 * 1)
    YScaleLabel4 = PVGraphXScaleUnitFromTwip(0, PrecisionVisionGraph.PVXScale(0).width / 4 * 0)
    
    YScaleLabel0 = Format(ConvertUnitByGraph(YScaleLabel0, 0, DisplayUnits), "###0.0#")
    YScaleLabel1 = Format(ConvertUnitByGraph(YScaleLabel1, 0, DisplayUnits), "###0.0#")
    YScaleLabel2 = Format(ConvertUnitByGraph(YScaleLabel2, 0, DisplayUnits), "###0.0#")
    YScaleLabel3 = Format(ConvertUnitByGraph(YScaleLabel3, 0, DisplayUnits), "###0.0#")
    YScaleLabel4 = Format(ConvertUnitByGraph(YScaleLabel4, 0, DisplayUnits), "###0.0#")
    
 '   PVGraphOvalityXScale = 8
'    PVGraphOvalityXOffset = -25
    Call PrecisionVisionGraph.GetGeneralPVGraphData(ScreenDrawing.ImageGraphState(0).GraphType)
    
'PCN4253 ' The me.limitvalue1 and me.limitvalue2 added. '''''''''''''''''''''''''''''''''
    Me.LimitValue1 = Format(ConvertUnitByGraph(PVXScaleLimitPerL, 0, DisplayUnits), "###0.0")
    Me.LimitValue2 = Format(ConvertUnitByGraph(PVXScaleLimitPerR, 0, DisplayUnits), "###0.0")
    LeftLimit = SafeCDbl(Me.LimitValue1)  'PCN4161                                                  '
    RightLimit = SafeCDbl(Me.LimitValue2)  'PCN4161                                                '
    Limit1Top = PVGraphXScaleTwipFromUnit(0, PVXScaleLimitPerL)                         '
    Limit2Top = PVGraphXScaleTwipFromUnit(0, PVXScaleLimitPerR)                         '
    Limit1Top = Limit1Top * (Me.GraphContainer.height / PrecisionVisionGraph.PVGraphImage(0).width)
    Limit2Top = Limit2Top * (Me.GraphContainer.height / PrecisionVisionGraph.PVGraphImage(0).width)
    Limit1Top = Me.GraphContainer.height - Limit1Top                                    '
    Limit2Top = Me.GraphContainer.height - Limit2Top                                    '
    Me.LimitValue1.Top = Limit1Top + Me.GraphContainer.Top                              '
    Me.LimitValue2.Top = Limit2Top + Me.GraphContainer.Top - Me.LimitValue1.height      '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    With PVReportMultiProfilex3
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
        
        .lblTitle.Caption = PrecisionVisionGraph.Label_GraphName(0) & " " & DisplayMessage("Observations Report")
        

        .ObservationsLabel(0).Caption = ""
        

    
        'By default make all lable no background
        For i = 0 To .Controls.Count - 1
            ControlType = TypeName(.Controls(i))
            Select Case ControlType
                Case "Label": .Controls(i).BackStyle = 0
            End Select
        Next i
        .UnitLabel = PrecisionVisionGraph.PVXScaleLabelUnits(0)

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
        
            Me.FractileLabelData = Me.FractileLabelData + ",    " + DisplayMessage("Exceeded limits") + ": " + Format(GraphInfoContainer(GraphInfoIndex).PercentExceeded, "###0.0") + "%"
        
        End If
    End With
    
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
    
    
    
    Call ScreenDrawing.FormTopMost(PVGraphsKeyForm.hwnd) 'PCN2990
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP33:" & Error$
    End Select
End Sub

Sub GraphSpecificSettings()
On Error GoTo Err_Handler
    Dim i As Integer
    Dim ColourKeysLeftShift As Single
    
    If ScreenDrawing.ImageGraphState(0).GraphType <> "Flat" Then
        Me.GraphContainer.width = Me.GraphContainer.width + 850
        Me.Shape2.width = Me.Shape2.width + 850
        Me.GraphUnitLbl.Left = Me.GraphUnitLbl.Left + 1150
        Me.GraphXScaleContainer.width = Me.GraphXScaleContainer.width + 850
        Me.GraphBackgroundShape.width = Me.GraphBackgroundShape.width + 850
        Me.LimitValue1.Left = Me.LimitValue1.Left + 850
        Me.LimitValue2.Left = Me.LimitValue2.Left + 850
    End If
    
    If ScreenDrawing.ImageGraphState(0).GraphType = "Ovality" Then
        PVReportMultiProfilex3.UnitLabel = PrecisionVisionGraph.PVXScaleLabelUnits(0)
        

        PVReportMultiProfilex3.Explination.Visible = True
        PVReportMultiProfilex3.Explination.Tag = ""
        
        Me.LimitValue2.Visible = False: Me.LimitValue2.Tag = "" 'PCN4253
        
        'PCN4324 if ovality display is less than zero, dont display
        For i = 0 To 4
            If Me.YScaleLabel(i) < 0 Then Me.YScaleLabel(i).Visible = False: Me.YScaleLabel(i).Tag = ""
        Next i
        
    ElseIf ScreenDrawing.ImageGraphState(0).GraphType = "Flat" Then
        With PVReportMultiProfilex3
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
        
        ColourKeysLeftShift = Me.PVKey_Flat3D_Color7.Left - Me.FooterBreakLine.x2 + Me.ColourKeyBoarder.width
        
        Me.DiameterLabel.Left = Me.DiameterLabel.Left - ColourKeysLeftShift
        Me.RadiusLabel.Left = Me.RadiusLabel.Left - ColourKeysLeftShift
        
        'PCN9420.................................................
        Me.KeyLine1.x1 = Me.KeyLine1.x1 - ColourKeysLeftShift   '
        Me.KeyLine1.x2 = Me.KeyLine1.x2 - ColourKeysLeftShift   '
        Me.KeyLine2.x1 = Me.KeyLine2.x1 - ColourKeysLeftShift   '
        Me.KeyLine2.x2 = Me.KeyLine2.x2 - ColourKeysLeftShift   '
        '........................................................
        
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
        
        Me.ColourKeyBoarder.Left = Me.ColourKeyBoarder.Left - ColourKeysLeftShift
        
        Me.GraphUnitLbl.Left = Me.GraphUnitLbl.Left - ColourKeysLeftShift + 1700
        
'        ColourKeysLeftShift = Me.GraphUnitLbl.Left - Printer.width + 6200
'        Me.GraphXScaleContainer.width = Me.GraphXScaleContainer.width - ColourKeysLeftShift
'        Me.GraphContainer.width = Me.GraphContainer.width - ColourKeysLeftShift
'        Me.GraphBackgroundShape.width = Me.GraphBackgroundShape.width - ColourKeysLeftShift
'        Me.Shape2.width = Me.Shape2.width - ColourKeysLeftShift

        

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
        Me.ColourKeyBoarder.Visible = True: Me.ColourKeyBoarder.Tag = "Visible"
        
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
            ColourKeyBoarder.width = ColourKeyBoarder.width - 700
            ColourKeyBoarder.Left = ColourKeyBoarder.Left + 700
            Me.GraphContainer.width = Me.GraphContainer.width + 700
            Me.Shape2.width = Me.Shape2.width + 700
            Me.GraphUnitLbl.Left = Me.GraphUnitLbl.Left + 700
            Me.GraphXScaleContainer.width = Me.GraphXScaleContainer.width + 700
            Me.GraphBackgroundShape.width = Me.GraphBackgroundShape.width + 700
        
        End If
        
        If FlatOvality Then
            Me.PVDGraphFlatOvality.Visible = True: Me.PVDGraphFlatOvality.Tag = "Visible"
        
            Me.GraphBackgroundShape.Visible = False: Me.GraphBackgroundShape.Tag = ""
            Me.Shape2.Visible = False: Me.Shape2.Tag = ""
            Me.GraphContainer.Visible = False: Me.GraphContainer.Tag = ""
            Me.GraphXScaleContainer.Visible = False: Me.GraphXScaleContainer.Tag = ""
            For i = 0 To 4
                Me.YScaleLabel(i).Visible = False: Me.YScaleLabel(i).Tag = ""
            Next i
        End If
        
    End If
    
    If MedianFlat And PVDFileName <> "" Then 'PCN5186
        Me.FractileLabel.Visible = False: Me.FractileLabelData.Visible = False 'PCN4296
        Me.FractileLabel.Tag = "": Me.FractileLabelData.Tag = ""
    End If
    
    If ScreenDrawing.ImageGraphState(0).GraphType <> "Ovality" Then
        PVReportMultiProfilex3.Explination.Visible = False
        PVReportMultiProfilex3.Explination.Tag = ""
    End If
    
    If ScreenDrawing.ImageGraphState(0).GraphType <> "Flat" Then
        Me.PVKey_Flat3D_Color7.Visible = False: Me.PVKey_Flat3D_Color7.Tag = ""
        Me.PVKey_Flat3D_Color6.Visible = False: Me.PVKey_Flat3D_Color6.Tag = ""
        Me.PVKey_Flat3D_Color5.Visible = False: Me.PVKey_Flat3D_Color5.Tag = ""
        Me.PVKey_Flat3D_Color4.Visible = False: Me.PVKey_Flat3D_Color4.Tag = ""
        Me.PVKey_Flat3D_Color2.Visible = False: Me.PVKey_Flat3D_Color2.Tag = ""
        Me.PVKey_Flat3D_Color1.Visible = False: Me.PVKey_Flat3D_Color1.Tag = ""
        Me.PVKey_Flat3D_Color0.Visible = False: Me.PVKey_Flat3D_Color0.Tag = ""
        Me.ColourKeyBoarder.Visible = False: Me.ColourKeyBoarder.Tag = ""
        
        Me.PVKey_Flat3D_Value7.Visible = False: Me.PVKey_Flat3D_Value7.Tag = ""
        Me.PVKey_Flat3D_Value6.Visible = False: Me.PVKey_Flat3D_Value6.Tag = ""
        Me.PVKey_Flat3D_Value5.Visible = False: Me.PVKey_Flat3D_Value5.Tag = ""
        Me.PVKey_Flat3D_Value4.Visible = False: Me.PVKey_Flat3D_Value4.Tag = ""
        Me.PVKey_Flat3D_Value2.Visible = False: Me.PVKey_Flat3D_Value2.Tag = ""
        Me.PVKey_Flat3D_Value1.Visible = False: Me.PVKey_Flat3D_Value1.Tag = ""
        Me.PVKey_Flat3D_Value0.Visible = False: Me.PVKey_Flat3D_Value0.Tag = ""
        
        Me.PVKey_Flat3D_Value7_Unit(0).Visible = False: Me.PVKey_Flat3D_Value7_Unit(0).Tag = "" 'PCN4920 (1) is diameter, (0) is radius
        Me.PVKey_Flat3D_Value6_Unit(0).Visible = False: Me.PVKey_Flat3D_Value6_Unit(0).Tag = "" 'PCN4920 (1) is diameter, (0) is radius
        Me.PVKey_Flat3D_Value5_Unit(0).Visible = False: Me.PVKey_Flat3D_Value5_Unit(0).Tag = "" 'PCN4920 (1) is diameter, (0) is radius
        Me.PVKey_Flat3D_Value4_Unit(0).Visible = False: Me.PVKey_Flat3D_Value4_Unit(0).Tag = "" 'PCN4920 (1) is diameter, (0) is radius
        Me.PVKey_Flat3D_Value2_Unit(0).Visible = False: Me.PVKey_Flat3D_Value2_Unit(0).Tag = "" 'PCN4920 (1) is diameter, (0) is radius
        Me.PVKey_Flat3D_Value1_Unit(0).Visible = False: Me.PVKey_Flat3D_Value1_Unit(0).Tag = "" 'PCN4920 (1) is diameter, (0) is radius
        Me.PVKey_Flat3D_Value0_Unit(0).Visible = False: Me.PVKey_Flat3D_Value0_Unit(0).Tag = "" 'PCN4920 (1) is diameter, (0) is radius
        
        Me.PVKey_Flat3D_Value7_Unit(1).Visible = False: Me.PVKey_Flat3D_Value7_Unit(1).Tag = "" 'PCN4920 (1) is diameter, (0) is radius
        Me.PVKey_Flat3D_Value6_Unit(1).Visible = False: Me.PVKey_Flat3D_Value6_Unit(1).Tag = "" 'PCN4920 (1) is diameter, (0) is radius
        Me.PVKey_Flat3D_Value5_Unit(1).Visible = False: Me.PVKey_Flat3D_Value5_Unit(1).Tag = "" 'PCN4920 (1) is diameter, (0) is radius
        Me.PVKey_Flat3D_Value4_Unit(1).Visible = False: Me.PVKey_Flat3D_Value4_Unit(1).Tag = "" 'PCN4920 (1) is diameter, (0) is radius
        Me.PVKey_Flat3D_Value2_Unit(1).Visible = False: Me.PVKey_Flat3D_Value2_Unit(1).Tag = "" 'PCN4920 (1) is diameter, (0) is radius
        Me.PVKey_Flat3D_Value1_Unit(1).Visible = False: Me.PVKey_Flat3D_Value1_Unit(1).Tag = "" 'PCN4920 (1) is diameter, (0) is radius
        Me.PVKey_Flat3D_Value0_Unit(1).Visible = False: Me.PVKey_Flat3D_Value0_Unit(1).Tag = "" 'PCN4920 (1) is diameter, (0) is radius
    End If
    

    
    Dim UserTittleCaption As String
    Dim TittleParse As String
        
    TittleParse = "Observations_" & ScreenDrawing.ImageGraphState(0).GraphType & "="
    Call GetINI_ParameterInfoOnly(MyFile, TittleParse, UserTittleCaption)
    Me.UserTitle.text = UserTittleCaption
    
    If Me.LimitValue2.Top + Me.LimitValue2.height < Me.GraphContainer.Top Then Me.LimitValue2.Visible = False: Me.LimitValue2.Tag = ""
    If Me.LimitValue1.Top > Me.GraphContainer.Top + Me.GraphContainer.height Then Me.LimitValue1.Visible = False: Me.LimitValue1.Tag = ""

    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP34:" & Error$
    End Select
    
End Sub

Sub MarkForPrinting()
On Error GoTo Err_Handler

Dim i As Integer
Dim ControlType As String

'Draw renderings first that are marked back
For i = 1 To PVReportMultiProfilex3.Controls.Count - 1
    
    With PVReportMultiProfilex3.Controls(i)
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
        Case Else: MsgBox Err & "-RMP35:" & Error$
    End Select
End Sub






Private Sub ScaleButton05_Click()
On Error GoTo Err_Handler
    RenderScale = 0.5
    Call RenderForm
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP36:" & Error$
    End Select
End Sub

Private Sub ScaleButton10_Click()
On Error GoTo Err_Handler
    RenderScale = 1
    Call RenderForm
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP37:" & Error$
    End Select
End Sub

Private Sub ScaleButton15_Click()
On Error GoTo Err_Handler
    RenderScale = 1.5
    Call RenderForm
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP38:" & Error$
    End Select
End Sub

Private Sub ScaleButton20_Click()
On Error GoTo Err_Handler
    RenderScale = 2
    Call RenderForm
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP39:" & Error$
    End Select
End Sub

Sub RenderForm()
On Error GoTo Err_Handler
    Dim Pg As Integer
    
    Set picReportPage = picReportPage1
    picReportPage.Font.Charset = LanguageCharset
    
    PreviewStartFrame = 1
    PreviewEndFrame = PVDataNoOfLines
    ScreenDrawingType = 2
    ScreenDrawingOrientation = 1
    
    Call RenderPage(picReportPage1, 1)
    For Pg = 1 To NumberOfExtraPages
        Set picReportPage = picReportNextPage(Pg)
'        Call RestoreOriginalState 'The restore was never setup now it is, it shouldn't be restored here
        Call RenderPage(picReportNextPage(Pg), Pg + 1)
        picReportNextPage(Pg).Top = (picReportNextPage(Pg).height + 300) * Pg
    Next Pg

    picReportPage1.Visible = True
    For Pg = 1 To NumberOfExtraPages
        picReportNextPage(Pg).Visible = True
    Next Pg
    
    UserTitle.Visible = True
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP40:" & Error$
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
        Case Else: MsgBox Err & "-RMP41:" & Error$
    End Select
End Function
 

 

Private Sub CmboPrinterList_Click()
On Error GoTo Err_Handler
    
If Me.ControlHighlight.Visible Then Me.ControlHighlight.Visible = False
If Me.ControlHighlightDepressed.Visible Then Me.ControlHighlightDepressed.Visible = False

Call SelectPrinter(CmboPrinterList.text)
    Dim i As Integer
    
    If Me.CmboPrinterList.Enabled = True Then

    For i = 1 To Me.FloatingText.Count - 1 'We dont want the text to be rendered on the preview yet
        Me.FloatingText(i).Visible = False 'when changes printers
    Next i
    
    Call Me.RestoreOriginalState
    Call Me.InitialiseForm
    
    For i = 1 To Me.FloatingText.Count - 1 'Even thou we dont want the text to be rendered, we still
        Me.FloatingText(i).Visible = True  'want to be able to see them after printer select changed
    Next i
    
End If

    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP42:" & Error$
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
        Case 438, 382, 383, 393: Resume Next
        Case Else: MsgBox Err & "-RMP43:" & Error$
    End Select


End Sub

Sub SetTheNoOfPages()
On Error GoTo Err_Handler
Dim NoOfPages As Integer
Dim CurrentPage As Integer

CurrentPage = PageHScroll.value

If NumberOfExtraPages > 0 Then
    NoOfPagesLabel.Caption = DisplayMessage("Page") & " " & CurrentPage & " " & DisplayMessage(" of ") & (NumberOfExtraPages + 1)
    NoOfPagesLabel.Visible = True
    PageHScroll.Visible = True
Else
    NoOfPagesLabel.Visible = False
    PageHScroll.Visible = False
    PageHScroll.value = 1
End If


Exit Sub
Err_Handler:
    MsgBox Err & "-RMP44:" & Error$
End Sub

Sub SetControlsPerPaperSize()
On Error GoTo Err_Handler
Dim i As Integer
Dim ShuntAmount As Single
Dim PlusMoveForFlatKey As Single

ShuntAmount = 735

If ScreenDrawing.ImageGraphState(0).GraphType = "Flat" Then
    PlusMoveForFlatKey = -2000

End If


If Printer.PaperSize = 1 Then
    ShuntAmount = 735
    ShuntX(0) = -(ShuntAmount / 3)
    ShuntX(1) = -(ShuntAmount / 3)
    ShuntX(2) = -(ShuntAmount / 3)
    ShuntX(3) = -ShuntAmount
    ShuntX(4) = -ShuntAmount
    

    
    UserTitle.width = UserTitle.width - ShuntAmount

'    For I = 0 To 16: GraphYDevisionLine(I).X2 = GraphYDevisionLine(I).X2 - 975:   Next I
    HeaderBreakLine.x2 = HeaderBreakLine.x2 - ShuntAmount
    FooterBreakLine.x2 = FooterBreakLine.x2 - ShuntAmount
    For i = 0 To 2
        PVProfileGraphBoarder(i).width = PVProfileGraphBoarder(i).width - (ShuntAmount / 3) '
        'LineDown(I).X1 = LineDown(I).X1 - (ShuntAmount / 3)
        'LineDown(I).X2 = LineDown(I).X1
        ObservationsLabel(i).width = ObservationsLabel(i).width - (ShuntAmount / 3)
        ProfileSliceRubberBand(i).x2 = ProfileSliceRubberBand(i).x2 - ((ShuntAmount / 3) * (i + 1))
    Next i
    
    
Else
    ShuntAmount = 0
    ShuntX(0) = 0
    ShuntX(1) = 0
    ShuntX(2) = 0
    ShuntX(3) = 0
    ShuntX(4) = 0
End If

    Shape2.width = Shape2.width - ShuntAmount + PlusMoveForFlatKey
    GraphBackgroundShape.width = GraphBackgroundShape.width - ShuntAmount + PlusMoveForFlatKey
    GraphXScaleContainer.width = GraphXScaleContainer.width - ShuntAmount + PlusMoveForFlatKey
    GraphContainer.width = GraphContainer.width - ShuntAmount + PlusMoveForFlatKey

Call ShuntControlsPerContainer

GraphUnitLbl.Left = GraphUnitLbl.Left + PlusMoveForFlatKey

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RMP45:" & Error$
    End Select

End Sub


Sub ShuntControlsPerContainer()
On Error GoTo Err_Handler

Dim i As Integer
Dim S As Integer
Dim ControlType As String

'Draw renderings first that are marked back

For S = 0 To 4 'Loop for the different Shunts
    For i = 1 To Me.Controls.Count - 1
        
        With Me.Controls(i)
            If TypeName(.Container) = "PictureBox" Then
                If .Tag <> "Container" Then
                    If TypeName(Me.Controls(i)) = "TextBox" Or _
                       TypeName(Me.Controls(i)) = "Label" Or _
                       TypeName(Me.Controls(i)) = "Shape" Or _
                       TypeName(Me.Controls(i)) = "Image" Or _
                       TypeName(Me.Controls(i)) = "PictureBox" Then
                        If .Left > controlspositionshunt(S).Left And _
                           .Left < controlspositionshunt(S).Left + controlspositionshunt(S).width And _
                           .Top > controlspositionshunt(S).Top And _
                           .Top < controlspositionshunt(S).Top + controlspositionshunt(S).height Then
                            .Top = .Top + ShuntY(S)
                            .Left = .Left + ShuntX(S)
                        End If
                    End If
                    If TypeName(Me.Controls(i)) = "Line" Then
                        If .x1 > controlspositionshunt(S).Left And _
                           .x1 < controlspositionshunt(S).Left + controlspositionshunt(S).width And _
                           .y1 > controlspositionshunt(S).Top And _
                           .y1 < controlspositionshunt(S).Top + controlspositionshunt(S).height Then
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
        Case Else: MsgBox Err & "-RMP46:" & Error$
    End Select
    
End Sub

Private Sub UserTitle_Change()
On Error GoTo Err_Handler

    Dim UserTittleCaption As String
    Dim TittleParse As String
        
    TittleParse = "Observations_" & ScreenDrawing.ImageGraphState(0).GraphType & "="
    Call INI_WriteBack(MyFile, TittleParse, Me.UserTitle.text)

'Dim FileSaveFail As Boolean
'
'UserTitleObservations = Trim(Me.UserTitle.text) 'PCN4433

Exit Sub
Err_Handler:
    Select Case Err

        Case Else: MsgBox Err & "-RMP47:" & Error$
    End Select
End Sub
