VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form PVReport4in1 
   BorderStyle     =   0  'None
   Caption         =   "Analysis Report"
   ClientHeight    =   14940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   22005
   Icon            =   "PVReport4in1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MousePointer    =   1  'Arrow
   ScaleHeight     =   14940
   ScaleWidth      =   22005
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox CmboPrinterList 
      Enabled         =   0   'False
      Height          =   315
      Left            =   5160
      MousePointer    =   1  'Arrow
      TabIndex        =   101
      Text            =   "Select a Printer"
      Top             =   120
      Width           =   3135
   End
   Begin VB.HScrollBar PageHScroll 
      Height          =   255
      Left            =   10680
      Max             =   2
      Min             =   1
      MousePointer    =   1  'Arrow
      TabIndex        =   100
      Top             =   240
      Value           =   1
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.VScrollBar NoGraphsDisplayVScroll 
      Height          =   255
      Left            =   10080
      Max             =   -1
      Min             =   -7
      MousePointer    =   1  'Arrow
      TabIndex        =   86
      Top             =   240
      Value           =   -1
      Width           =   375
   End
   Begin VB.CommandButton NoOfGraphsUp 
      Caption         =   "^"
      Height          =   315
      Left            =   -240
      TabIndex        =   84
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton NoGraphsDown 
      Caption         =   "v"
      Height          =   315
      Left            =   -240
      TabIndex        =   83
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton ScaleButton05 
      Caption         =   "0.5"
      Height          =   255
      Left            =   -240
      TabIndex        =   44
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton ScaleButton10 
      Caption         =   "1"
      Height          =   255
      Left            =   -240
      TabIndex        =   43
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton ScaleButton15 
      Caption         =   "1.5"
      Height          =   255
      Left            =   -240
      TabIndex        =   42
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton ScaleButton20 
      Caption         =   "2"
      Height          =   255
      Left            =   -240
      TabIndex        =   41
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PageFramePictureBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   15795
      Left            =   0
      ScaleHeight     =   15765
      ScaleWidth      =   21060
      TabIndex        =   0
      Top             =   600
      Width           =   21090
      Begin VB.PictureBox picReportPagePg2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   15700
         Left            =   17520
         ScaleHeight     =   15675
         ScaleWidth      =   11505
         TabIndex        =   52
         Tag             =   "Paper"
         Top             =   240
         Width           =   11535
         Begin VB.Label LimitValue1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "xxxx.xx"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   10680
            TabIndex        =   116
            Top             =   12480
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
            Index           =   6
            Left            =   10680
            TabIndex        =   115
            Top             =   13440
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
            Index           =   5
            Left            =   10680
            TabIndex        =   114
            Top             =   9120
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
            Index           =   5
            Left            =   10680
            TabIndex        =   113
            Top             =   10080
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
            Index           =   4
            Left            =   10680
            TabIndex        =   112
            Top             =   5640
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
            Index           =   4
            Left            =   10680
            TabIndex        =   111
            Top             =   6600
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
            Index           =   3
            Left            =   10680
            TabIndex        =   110
            Top             =   1920
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
            Index           =   3
            Left            =   10680
            TabIndex        =   109
            Top             =   2880
            Width           =   615
         End
         Begin VB.Image ControlsPositionShunt 
            Height          =   855
            Index           =   1
            Left            =   120
            Tag             =   "Shunt"
            Top             =   14760
            Width           =   11295
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
            Index           =   6
            Left            =   11040
            TabIndex        =   96
            Top             =   14400
            Width           =   255
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
            Index           =   5
            Left            =   11040
            TabIndex        =   95
            Top             =   11280
            Width           =   255
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
            Index           =   4
            Left            =   11040
            TabIndex        =   94
            Top             =   7920
            Width           =   255
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
            Index           =   3
            Left            =   11040
            TabIndex        =   93
            Top             =   4440
            Width           =   255
         End
         Begin VB.Image GraphContainerArray 
            Appearance      =   0  'Flat
            Height          =   2415
            Index           =   3
            Left            =   1320
            Stretch         =   -1  'True
            Top             =   1800
            Width           =   9375
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
            Index           =   19
            Left            =   600
            TabIndex        =   81
            Top             =   4080
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
            Index           =   18
            Left            =   600
            TabIndex        =   80
            Top             =   3480
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
            Index           =   17
            Left            =   600
            TabIndex        =   79
            Top             =   2880
            Width           =   585
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
            Index           =   16
            Left            =   600
            TabIndex        =   78
            Top             =   2280
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
            Index           =   15
            Left            =   600
            TabIndex        =   77
            Top             =   1680
            Width           =   600
         End
         Begin VB.Shape Shape2 
            Height          =   2415
            Index           =   3
            Left            =   1320
            Top             =   1800
            Width           =   9375
         End
         Begin VB.Image GraphXScaleContainerArray 
            Height          =   375
            Index           =   3
            Left            =   1320
            Top             =   4200
            Width           =   9375
         End
         Begin VB.Label UnitLabel 
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
            Index           =   6
            Left            =   1200
            TabIndex        =   76
            Top             =   11760
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
            Index           =   34
            Left            =   600
            TabIndex        =   75
            Top             =   14280
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
            Index           =   33
            Left            =   600
            TabIndex        =   74
            Top             =   13680
            Width           =   600
         End
         Begin VB.Shape ShapeArray 
            BackColor       =   &H00000000&
            BackStyle       =   1  'Opaque
            BorderWidth     =   2
            FillStyle       =   0  'Solid
            Height          =   15
            Index           =   1
            Left            =   3000
            Shape           =   1  'Square
            Top             =   11280
            Visible         =   0   'False
            Width           =   15
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
            Index           =   1
            Left            =   9240
            TabIndex        =   72
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
            Index           =   1
            Left            =   10200
            TabIndex        =   71
            Top             =   15120
            Width           =   855
         End
         Begin VB.Line FooterBreakLine 
            Index           =   1
            X1              =   600
            X2              =   11280
            Y1              =   14850
            Y2              =   14850
         End
         Begin VB.Line HeaderBreakLine 
            Index           =   1
            X1              =   600
            X2              =   11280
            Y1              =   1440
            Y2              =   1440
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
            Index           =   1
            Left            =   4320
            TabIndex        =   70
            Top             =   14880
            Width           =   2295
         End
         Begin VB.Label PageLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "Page: 2"
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
            Left            =   9600
            TabIndex        =   69
            Top             =   14880
            Width           =   1215
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
            Index           =   32
            Left            =   600
            TabIndex        =   67
            Top             =   13080
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
            Index           =   31
            Left            =   600
            TabIndex        =   66
            Top             =   12480
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
            Index           =   30
            Left            =   600
            TabIndex        =   65
            Top             =   11880
            Width           =   600
         End
         Begin VB.Image GraphContainerArray 
            Appearance      =   0  'Flat
            Height          =   2415
            Index           =   6
            Left            =   1320
            Stretch         =   -1  'True
            Top             =   12000
            Width           =   9375
         End
         Begin VB.Image GraphXScaleContainerArray 
            Height          =   375
            Index           =   6
            Left            =   1320
            Top             =   14400
            Width           =   9375
         End
         Begin VB.Shape Shape2 
            Height          =   2415
            Index           =   6
            Left            =   1320
            Top             =   12000
            Width           =   9375
         End
         Begin VB.Image GraphXScaleContainerArray 
            Height          =   375
            Index           =   5
            Left            =   1320
            Top             =   11040
            Width           =   9375
         End
         Begin VB.Image GraphXScaleContainerArray 
            Height          =   375
            Index           =   4
            Left            =   1320
            Top             =   7680
            Width           =   9375
         End
         Begin VB.Shape Shape2 
            Height          =   2415
            Index           =   5
            Left            =   1320
            Top             =   8640
            Width           =   9375
         End
         Begin VB.Shape Shape2 
            Height          =   2535
            Index           =   4
            Left            =   1320
            Top             =   5160
            Width           =   9375
         End
         Begin VB.Image GraphContainerArray 
            Appearance      =   0  'Flat
            Height          =   2415
            Index           =   5
            Left            =   1320
            Stretch         =   -1  'True
            Top             =   8640
            Width           =   9375
         End
         Begin VB.Image GraphContainerArray 
            Appearance      =   0  'Flat
            Height          =   2535
            Index           =   4
            Left            =   1320
            Stretch         =   -1  'True
            Top             =   5160
            Width           =   9375
         End
         Begin VB.Label UnitLabel 
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
            Index           =   5
            Left            =   1200
            TabIndex        =   64
            Top             =   8400
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
            Index           =   29
            Left            =   600
            TabIndex        =   63
            Top             =   10920
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
            Index           =   28
            Left            =   600
            TabIndex        =   62
            Top             =   10320
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
            Index           =   27
            Left            =   600
            TabIndex        =   61
            Top             =   9720
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
            Index           =   26
            Left            =   600
            TabIndex        =   60
            Top             =   9120
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
            Index           =   25
            Left            =   600
            TabIndex        =   59
            Top             =   8520
            Width           =   600
         End
         Begin VB.Label UnitLabel 
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
            Index           =   4
            Left            =   1200
            TabIndex        =   58
            Top             =   4920
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
            Index           =   24
            Left            =   600
            TabIndex        =   57
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
            Index           =   23
            Left            =   600
            TabIndex        =   56
            Top             =   6930
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
            Index           =   22
            Left            =   600
            TabIndex        =   55
            Top             =   6300
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
            Index           =   21
            Left            =   600
            TabIndex        =   54
            Top             =   5670
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
            Index           =   20
            Left            =   600
            TabIndex        =   53
            Top             =   5040
            Width           =   600
         End
         Begin VB.Shape GraphBackgroundShape 
            BackColor       =   &H00DDDDA2&
            FillColor       =   &H00DDDDA2&
            FillStyle       =   0  'Solid
            Height          =   3060
            Index           =   6
            Left            =   600
            Tag             =   "Back"
            Top             =   11760
            Width           =   10695
         End
         Begin VB.Shape GraphBackgroundShape 
            BackColor       =   &H00DDDDA2&
            FillColor       =   &H00DDDDA2&
            FillStyle       =   0  'Solid
            Height          =   3270
            Index           =   5
            Left            =   600
            Tag             =   "Back"
            Top             =   8325
            Width           =   10695
         End
         Begin VB.Shape GraphBackgroundShape 
            BackColor       =   &H00DDDDA2&
            FillColor       =   &H00DDDDA2&
            FillStyle       =   0  'Solid
            Height          =   3300
            Index           =   4
            Left            =   600
            Tag             =   "Back"
            Top             =   4875
            Width           =   10695
         End
         Begin VB.Image CLPLogoImage 
            Height          =   690
            Index           =   1
            Left            =   600
            Picture         =   "PVReport4in1.frx":038A
            Top             =   14880
            Width           =   2010
         End
         Begin VB.Label CopyrightLabel 
            Alignment       =   2  'Center
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
            Index           =   1
            Left            =   4800
            TabIndex        =   68
            Top             =   15120
            Width           =   1455
         End
         Begin VB.Image LogoImage 
            Height          =   855
            Index           =   1
            Left            =   600
            Stretch         =   -1  'True
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label UnitLabel 
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
            Index           =   3
            Left            =   1200
            TabIndex        =   82
            Top             =   1560
            Width           =   255
         End
         Begin VB.Shape GraphBackgroundShape 
            BackColor       =   &H00DDDDA2&
            FillColor       =   &H00DDDDA2&
            FillStyle       =   0  'Solid
            Height          =   3180
            Index           =   3
            Left            =   600
            Tag             =   "Back"
            Top             =   1560
            Width           =   10695
         End
         Begin VB.Label lblTitle 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "Capacity Analysis Report"
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
            Index           =   1
            Left            =   480
            TabIndex        =   73
            Top             =   600
            Width           =   11055
         End
      End
      Begin VB.PictureBox picReportPagePg1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   15700
         Left            =   0
         ScaleHeight     =   15675
         ScaleWidth      =   16785
         TabIndex        =   1
         Tag             =   "Paper"
         Top             =   0
         Width           =   16815
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
            TabIndex        =   119
            Top             =   960
            Visible         =   0   'False
            Width           =   6495
         End
         Begin VB.TextBox FloatingText 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   2640
            TabIndex        =   99
            Top             =   120
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox FloatingTextDefault 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2640
            TabIndex        =   98
            Text            =   "Default Text Setting"
            Top             =   480
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox CommentsTextBox 
            Appearance      =   0  'Flat
            Height          =   735
            Index           =   0
            Left            =   600
            MousePointer    =   3  'I-Beam
            TabIndex        =   2
            Top             =   13920
            Width           =   10575
         End
         Begin VB.Line KeyLine 
            Visible         =   0   'False
            X1              =   16560
            X2              =   16560
            Y1              =   13440
            Y2              =   14640
         End
         Begin VB.Shape FlatKeyFrame 
            Height          =   1335
            Left            =   11520
            Top             =   13440
            Visible         =   0   'False
            Width           =   4815
         End
         Begin VB.Label PVKey_Flat3D_Value0_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   15000
            TabIndex        =   149
            Top             =   14115
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label PVKey_Flat3D_Value1_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   15000
            TabIndex        =   148
            Top             =   13875
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label PVKey_Flat3D_Value2_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   15000
            TabIndex        =   147
            Top             =   13635
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label PVKey_Flat3D_Value7_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   12600
            TabIndex        =   146
            Top             =   13635
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label PVKey_Flat3D_Value5_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   12600
            TabIndex        =   145
            Top             =   14115
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label PVKey_Flat3D_Value6_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   12600
            TabIndex        =   144
            Top             =   13875
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label PVKey_Flat3D_Value4_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   12600
            TabIndex        =   143
            Top             =   14370
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Shape PVKey_Flat3D_Color7 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   270
            Left            =   11640
            Top             =   13605
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Shape PVKey_Flat3D_Color6 
            FillColor       =   &H000096FF&
            FillStyle       =   0  'Solid
            Height          =   270
            Left            =   11640
            Top             =   13845
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Shape PVKey_Flat3D_Color5 
            FillColor       =   &H0014FFFF&
            FillStyle       =   0  'Solid
            Height          =   270
            Left            =   11640
            Top             =   14085
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Shape PVKey_Flat3D_Color0 
            FillColor       =   &H006F4928&
            FillStyle       =   0  'Solid
            Height          =   270
            Left            =   14040
            Top             =   14085
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Shape PVKey_Flat3D_Color1 
            FillColor       =   &H00CC9B5A&
            FillStyle       =   0  'Solid
            Height          =   270
            Left            =   14040
            Top             =   13845
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Shape PVKey_Flat3D_Color2 
            FillColor       =   &H00EEE0B5&
            FillStyle       =   0  'Solid
            Height          =   270
            Left            =   14040
            Top             =   13605
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Shape PVKey_Flat3D_Color4 
            Height          =   270
            Left            =   11640
            Top             =   14340
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Label PVKey_Flat3D_Value4 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   12000
            TabIndex        =   142
            Top             =   14370
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label PVKey_Flat3D_Value6 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   12000
            TabIndex        =   141
            Top             =   13875
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label PVKey_Flat3D_Value5 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   12000
            TabIndex        =   140
            Top             =   14115
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label PVKey_Flat3D_Value7 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   12000
            TabIndex        =   139
            Top             =   13635
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label PVKey_Flat3D_Value2 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   14400
            TabIndex        =   138
            Top             =   13635
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label PVKey_Flat3D_Value1 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   14400
            TabIndex        =   137
            Top             =   13875
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label PVKey_Flat3D_Value0 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   14400
            TabIndex        =   136
            Top             =   14115
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label PVKey_Flat3D_Value2_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   15720
            TabIndex        =   135
            Top             =   13635
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label PVKey_Flat3D_Value1_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   15720
            TabIndex        =   134
            Top             =   13875
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label PVKey_Flat3D_Value0_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   15720
            TabIndex        =   133
            Top             =   14115
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label PVKey_Flat3D_Value4_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   13320
            TabIndex        =   132
            Top             =   14370
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label PVKey_Flat3D_Value6_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   13320
            TabIndex        =   131
            Top             =   13875
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label PVKey_Flat3D_Value5_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   13320
            TabIndex        =   130
            Top             =   14115
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label PVKey_Flat3D_Value7_Unit 
            Caption         =   "0mm"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   13320
            TabIndex        =   129
            Top             =   13635
            Visible         =   0   'False
            Width           =   840
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
            Left            =   13440
            TabIndex        =   128
            Top             =   13395
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
            Left            =   12090
            TabIndex        =   127
            Top             =   13335
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
            Left            =   15840
            TabIndex        =   126
            Top             =   13395
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
            Index           =   1
            Left            =   14490
            TabIndex        =   125
            Top             =   13335
            Visible         =   0   'False
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
            Index           =   10
            Left            =   600
            TabIndex        =   124
            Top             =   10320
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
            Index           =   11
            Left            =   600
            TabIndex        =   123
            Top             =   10950
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
            Index           =   12
            Left            =   600
            TabIndex        =   122
            Top             =   11580
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
            Index           =   13
            Left            =   600
            TabIndex        =   121
            Top             =   12210
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
            Index           =   14
            Left            =   600
            TabIndex        =   120
            Top             =   12840
            Width           =   600
         End
         Begin VB.Image LogoImage 
            Height          =   855
            Index           =   0
            Left            =   600
            Stretch         =   -1  'True
            Top             =   480
            Width           =   2175
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
            Left            =   600
            TabIndex        =   118
            Top             =   2760
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
            Left            =   1800
            TabIndex        =   117
            Top             =   2760
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
            Index           =   2
            Left            =   10680
            TabIndex        =   108
            Top             =   10800
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
            Index           =   2
            Left            =   10680
            TabIndex        =   107
            Top             =   11760
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
            Index           =   1
            Left            =   10680
            TabIndex        =   106
            Top             =   7200
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
            Index           =   1
            Left            =   10680
            TabIndex        =   105
            Top             =   8160
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
            Index           =   0
            Left            =   10680
            TabIndex        =   104
            Top             =   3840
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
            Index           =   0
            Left            =   10680
            TabIndex        =   103
            Top             =   4800
            Width           =   615
         End
         Begin VB.Image ControlsPositionShunt 
            Height          =   2055
            Index           =   0
            Left            =   0
            Tag             =   "Shunt"
            Top             =   13560
            Width           =   11415
         End
         Begin VB.Line HeaderBreakLine 
            Index           =   0
            X1              =   600
            X2              =   11280
            Y1              =   1440
            Y2              =   1440
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
            Index           =   0
            Left            =   5400
            TabIndex        =   31
            Top             =   15120
            Width           =   2295
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
            Left            =   3000
            TabIndex        =   97
            Top             =   14880
            Width           =   5055
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
            Index           =   2
            Left            =   10680
            TabIndex        =   92
            Top             =   13200
            Width           =   255
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
            Index           =   1
            Left            =   10680
            TabIndex        =   91
            Top             =   9600
            Width           =   255
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
            Index           =   0
            Left            =   10680
            TabIndex        =   90
            Top             =   6120
            Width           =   255
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
            TabIndex        =   89
            Top             =   15120
            Width           =   735
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
            TabIndex        =   88
            Top             =   15120
            Width           =   1695
         End
         Begin VB.Image GraphXScaleContainerArray 
            Height          =   375
            Index           =   1
            Left            =   1320
            Top             =   9360
            Width           =   9375
         End
         Begin VB.Label UnitLabel 
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
            Index           =   2
            Left            =   1200
            TabIndex        =   51
            Top             =   10200
            Width           =   255
         End
         Begin VB.Image GraphContainerArray 
            Appearance      =   0  'Flat
            Height          =   2535
            Index           =   2
            Left            =   1320
            Stretch         =   -1  'True
            Top             =   10440
            Width           =   9375
         End
         Begin VB.Shape Shape2 
            Height          =   2535
            Index           =   2
            Left            =   1320
            Top             =   10440
            Width           =   9375
         End
         Begin VB.Image GraphXScaleContainerArray 
            Height          =   375
            Index           =   2
            Left            =   1320
            Top             =   12960
            Width           =   9375
         End
         Begin VB.Shape GraphBackgroundShape 
            BackColor       =   &H00DDDDA2&
            FillColor       =   &H00DDDDA2&
            FillStyle       =   0  'Solid
            Height          =   3300
            Index           =   2
            Left            =   600
            Tag             =   "Back"
            Top             =   10200
            Width           =   10695
         End
         Begin VB.Shape Shape2 
            Height          =   2535
            Index           =   1
            Left            =   1320
            Top             =   6840
            Width           =   9375
         End
         Begin VB.Image GraphContainerArray 
            Appearance      =   0  'Flat
            Height          =   2535
            Index           =   1
            Left            =   1320
            Stretch         =   -1  'True
            Top             =   6840
            Width           =   9375
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
            Left            =   600
            TabIndex        =   40
            Top             =   3120
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
            Left            =   600
            TabIndex        =   39
            Top             =   5760
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
            Left            =   600
            TabIndex        =   38
            Top             =   4500
            Width           =   600
         End
         Begin VB.Image GraphXScaleContainerArray 
            Height          =   375
            Index           =   0
            Left            =   1320
            Top             =   5880
            Width           =   9375
         End
         Begin VB.Image GraphContainerArray 
            Appearance      =   0  'Flat
            Height          =   2535
            Index           =   0
            Left            =   1320
            Stretch         =   -1  'True
            Top             =   3360
            Width           =   9375
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
            Index           =   0
            Left            =   4200
            TabIndex        =   37
            Top             =   2520
            Width           =   7095
            WordWrap        =   -1  'True
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
            Index           =   0
            Left            =   8280
            TabIndex        =   36
            Top             =   14880
            Width           =   1455
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
            Index           =   0
            Left            =   10800
            TabIndex        =   35
            Top             =   2280
            Width           =   495
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
            Index           =   0
            Left            =   10800
            TabIndex        =   34
            Top             =   2040
            Width           =   495
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
            Index           =   0
            Left            =   600
            TabIndex        =   33
            Top             =   13680
            Width           =   1215
         End
         Begin VB.Label PageLabel 
            Caption         =   "Page: 1"
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
            Left            =   10320
            TabIndex        =   32
            Top             =   14880
            Width           =   855
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
            Index           =   0
            Left            =   10080
            TabIndex        =   30
            Top             =   1560
            Width           =   1215
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
            Index           =   0
            Left            =   10080
            TabIndex        =   29
            Top             =   2280
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
            Index           =   0
            Left            =   10080
            TabIndex        =   28
            Top             =   2040
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
            Index           =   0
            Left            =   1800
            TabIndex        =   27
            Top             =   1800
            Width           =   4335
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
            Index           =   0
            Left            =   10080
            TabIndex        =   26
            Top             =   1800
            Width           =   1215
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
            Index           =   0
            Left            =   6000
            TabIndex        =   25
            Top             =   2280
            Width           =   2535
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
            Index           =   0
            Left            =   6000
            TabIndex        =   24
            Top             =   2040
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
            Index           =   0
            Left            =   6000
            TabIndex        =   23
            Top             =   1560
            Width           =   3015
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
            Index           =   0
            Left            =   1800
            TabIndex        =   22
            Top             =   2280
            Width           =   2775
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
            Index           =   0
            Left            =   1800
            TabIndex        =   21
            Top             =   2040
            Width           =   2775
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
            Index           =   0
            Left            =   1800
            TabIndex        =   20
            Top             =   1560
            Width           =   2775
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
            Index           =   0
            Left            =   600
            TabIndex        =   19
            Top             =   1560
            Width           =   1095
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
            Index           =   0
            Left            =   9120
            TabIndex        =   18
            Top             =   1560
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
            Index           =   0
            Left            =   8640
            TabIndex        =   17
            Top             =   2280
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
            Index           =   0
            Left            =   8640
            TabIndex        =   16
            Top             =   2040
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
            Index           =   0
            Left            =   600
            TabIndex        =   15
            Top             =   1800
            Width           =   1095
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
            Index           =   0
            Left            =   8280
            TabIndex        =   14
            Top             =   1800
            Width           =   1695
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
            Index           =   0
            Left            =   4560
            TabIndex        =   13
            Top             =   2280
            Width           =   1335
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
            Index           =   0
            Left            =   4560
            TabIndex        =   12
            Top             =   2040
            Width           =   1335
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
            Index           =   0
            Left            =   4560
            TabIndex        =   11
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Line FooterBreakLine 
            Index           =   0
            X1              =   600
            X2              =   11160
            Y1              =   14760
            Y2              =   14760
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
            Index           =   0
            Left            =   10320
            TabIndex        =   8
            Top             =   15240
            Width           =   855
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
            Index           =   0
            Left            =   9360
            TabIndex        =   7
            Top             =   15240
            Width           =   855
         End
         Begin VB.Label lblTitle 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "Capacity Analysis Report"
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
            Index           =   0
            Left            =   600
            TabIndex        =   6
            Tag             =   "back"
            Top             =   480
            Width           =   10815
         End
         Begin VB.Image CLPLogoImage 
            Height          =   690
            Index           =   0
            Left            =   600
            Picture         =   "PVReport4in1.frx":0BEF
            Top             =   14880
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
            Left            =   600
            TabIndex        =   5
            Top             =   5130
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
            Left            =   600
            TabIndex        =   4
            Top             =   3870
            Width           =   600
         End
         Begin VB.Label UnitLabel 
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
            Index           =   0
            Left            =   1200
            TabIndex        =   3
            Top             =   3120
            Width           =   255
         End
         Begin VB.Shape Shape2 
            Height          =   2535
            Index           =   0
            Left            =   1320
            Top             =   3360
            Width           =   9375
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
            Index           =   9
            Left            =   600
            TabIndex        =   50
            Top             =   9240
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
            Index           =   8
            Left            =   600
            TabIndex        =   49
            Top             =   8610
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
            Index           =   7
            Left            =   600
            TabIndex        =   48
            Top             =   7980
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
            Index           =   6
            Left            =   600
            TabIndex        =   47
            Top             =   7350
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
            Index           =   5
            Left            =   600
            TabIndex        =   46
            Top             =   6720
            Width           =   600
         End
         Begin VB.Label UnitLabel 
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
            Index           =   1
            Left            =   1200
            TabIndex        =   45
            Top             =   6600
            Width           =   255
         End
         Begin VB.Shape GraphBackgroundShape 
            BackColor       =   &H00DDDDA2&
            FillColor       =   &H00DDDDA2&
            FillStyle       =   0  'Solid
            Height          =   3300
            Index           =   1
            Left            =   600
            Tag             =   "Back"
            Top             =   6600
            Width           =   10695
         End
         Begin VB.Shape GraphBackgroundShape 
            BackColor       =   &H00DDDDA2&
            FillColor       =   &H00DDDDA2&
            FillStyle       =   0  'Solid
            Height          =   3300
            Index           =   0
            Left            =   600
            Tag             =   "Back"
            Top             =   3120
            Width           =   10695
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
            Index           =   0
            Left            =   600
            TabIndex        =   10
            Top             =   2280
            Width           =   1095
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
            Index           =   0
            Left            =   600
            TabIndex        =   9
            Top             =   2040
            Width           =   1095
         End
      End
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
   Begin MSComDlg.CommonDialog FloatingTextDialog 
      Left            =   120
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   10680
      MousePointer    =   1  'Arrow
      TabIndex        =   102
      Top             =   30
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image ControlsReport 
      Height          =   720
      Index           =   0
      Left            =   120
      MousePointer    =   1  'Arrow
      Picture         =   "PVReport4in1.frx":1454
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
      Picture         =   "PVReport4in1.frx":311E
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
      Picture         =   "PVReport4in1.frx":4DE8
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
      Picture         =   "PVReport4in1.frx":6AB2
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
      Picture         =   "PVReport4in1.frx":877C
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
      Picture         =   "PVReport4in1.frx":A446
      Tag             =   "SaveReportToPVD"
      ToolTipText     =   "Save Report To PVD"
      Top             =   -60
      Width           =   720
   End
   Begin VB.Image CloseReport 
      Height          =   480
      Left            =   14280
      MousePointer    =   1  'Arrow
      Picture         =   "PVReport4in1.frx":C110
      ToolTipText     =   "Close Report"
      Top             =   90
      Width           =   480
   End
   Begin VB.Label TotalNoOfGraphsLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No Of Panels"
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
      Left            =   8520
      MousePointer    =   1  'Arrow
      TabIndex        =   87
      Top             =   30
      Width           =   1935
   End
   Begin VB.Label NoGraphsDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      Enabled         =   0   'False
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
      Left            =   8520
      MousePointer    =   1  'Arrow
      TabIndex        =   85
      Top             =   240
      Width           =   1575
   End
   Begin VB.Image ControlHighlight 
      Height          =   570
      Left            =   12000
      Picture         =   "PVReport4in1.frx":CDDA
      Top             =   30
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Image ControlHighlightDepressed 
      Height          =   585
      Left            =   13080
      Picture         =   "PVReport4in1.frx":E96C
      Top             =   30
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Image ControlsBackPanel 
      Height          =   615
      Left            =   0
      MousePointer    =   1  'Arrow
      Picture         =   "PVReport4in1.frx":105B6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14775
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
Attribute VB_Name = "PVReport4in1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public GraphContainer As Image
Public GraphXScaleContainer As Image
Public picReportPage As PictureBox

'Dim ReportMouseDown As Boolean
Public PreviewStartFrame As Long
Public PreviewEndFrame As Long
'Public RenderScale As Single

Public NoOfGraphPanels As Integer 'The number of graph panels

Dim ShuntY(2) As Single
Dim ShuntX(2) As Single

Private Sub AllFramesButtons_Click()
On Error GoTo Err_Handler

    GraphStartFrame = 1
    GraphEndFrame = PVDataNoOfLines
    Call PositionReportControls
    Call FillOutPrintForm
    Call GraphSpecificSettings
    Call RenderPage(picReportPagePg1)
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R4in11:" & Error$
    End Select
End Sub

Private Sub CloseReport_Click()
On Error GoTo Err_Handler

Unload Me
    
Exit Sub
Err_Handler:
    MsgBox Err & "-R4in12:" & Error$
End Sub


Sub Print4in1Report()
On Error GoTo Err_Handler
    Dim i As Integer

    Printer.EndDoc
    RenderScale = 1
    ScreenDrawingType = 1
    ScreenDrawingOrientation = 1
    
   
    Call PageToPrint(picReportPagePg1)
    Call RenderToPrinter.RenderSingleTextBox(Me.CommentsTextBox(0), Printer, True)
'    If ScreenDrawing.ImageGraphState(0).GraphType <> "Ovality" Then
'        Call RenderToPrinter.RenderSingleTextBox(Me, Printer, False)
'    End If
    
    
    For i = 1 To Me.FloatingText.Count - 1
        If Me.FloatingText(i).Container.name = "picReportPagePg1" Then
            Call RenderToPrinter.RenderSingleTextBox(Me.FloatingText(i), Printer, True)
        End If
    Next i
    
    If NoOfGraphPanels > 3 Then
        Printer.NewPage
        Call PageToPrint(picReportPagePg2)
        For i = 1 To Me.FloatingText.Count - 1
            If Me.FloatingText(i).Container.name = "picReportPagePg2" Then
                Call RenderToPrinter.RenderSingleTextBox(Me.FloatingText(i), Printer, True)
            End If
        Next i
    End If
    Printer.EndDoc
    
    ScreenDrawingType = 0
    ScreenDrawingOrientation = 0

    Call RepositionAndRender
    'If Printer.PaperSize <> 1 Then Me.CommentsTextBox(0).Top = 13920 + 735
    If Printer.PaperSize <> 1 Then
        Me.CommentsTextBox(0).Top = 13920 + 235
    Else
        Me.CommentsTextBox(0).Top = 13920 - 500
    End If
    

    Me.PageFramePictureBox.Visible = True
    For i = 1 To Me.FloatingText.Count - 1
        Me.FloatingText(i).Visible = True
    Next i
    
    
Exit Sub
Err_Handler:
    Select Case Err
        Case 428: MsgBox "Printer Error": Exit Sub 'Printer error
        Case Else: MsgBox Err & "-R4in13:" & Error$
    End Select

End Sub



Sub PageToPrint(PageToRender As PictureBox)
On Error GoTo Err_Handler
    Dim i As Integer
    Dim J As Integer
    Dim NG As Long
    Dim Pg1NG As Long
    Dim NoFr As Long
    Dim StartGraph As Long
    Dim EndGraph As Long
    
    

    
    Call RestoreOriginalState
    
    If PageToRender.name = "picReportPagePg1" Then
        ControlsPositionShunt(0).Tag = "Shunt": ControlsPositionShunt(1).Tag = "Egnore"
    Else
        ControlsPositionShunt(0).Tag = "Egnore": ControlsPositionShunt(1).Tag = "Shunt"
    End If
    
    Call Me.SetControlsPerPaperSize
    



    For i = 1 To PVReport4in1.Controls.Count - 1
        With PVReport4in1.Controls(i)
            If .name <> PageToRender.name And _
               .Container.name <> "PVReport4in1" And _
               .name <> "PageFramePictureBox" And _
               .name <> "FloatingTextDefault" And _
               .name <> "FloatingText" Then
                If .Container.name <> PageToRender.name Then
                   .Visible = False
                   .Tag = ""
                Else
                    .Visible = True

                End If

            End If
        End With
    Next i
    PageToRender.Visible = True
    
    Set picReportPage = PageToRender

    PageToRender.Cls
    PageToRender.width = Printer.width * RenderScale
    PageToRender.height = Printer.height * RenderScale
    
 
    'Call PositionReportControls
    Call FillOutPrintForm
    Call GraphSpecificSettings
    
    FloatingTextDefault.Visible = False
    FloatingTextDefault.Tag = ""
    
    FloatingText(0).Visible = False
    FloatingText(0).Tag = ""
 

    
    If PageToRender.name = "picReportPagePg2" Then
        PVReport4in1.Explination(0).Visible = False
        PVReport4in1.Explination(0).Tag = ""
    End If
    
    Call MarkForPrinting
    Me.UserTitle.Tag = "" 'PCN4277
'    Call RenderToPrinter.RenderReport(Me, Printer, RenderScale)


        
    NG = NoOfGraphPanels
    NoFr = PVDataNoOfLines / NG
    
    PreviewStartFrame = 0
    PreviewEndFrame = NoFr
    
    If PageToRender.name = "picReportPagePg1" Then
         EndGraph = NG
         StartGraph = 0
         If NG > 3 Then EndGraph = 3
         PreviewStartFrame = 0
         PreviewEndFrame = NoFr
    Else
        If NG <= 3 Then ScreenDrawingType = 0:  ScreenDrawingOrientation = 0: Exit Sub
        StartGraph = 3
        EndGraph = NG
        PreviewStartFrame = 3 * (NoFr + 1)
        PreviewEndFrame = 4 * (NoFr + 1)
        
        'PCN4417
        Me.PVKey_Flat3D_Color7.Visible = False: Me.PVKey_Flat3D_Color7.Tag = ""
        Me.PVKey_Flat3D_Color6.Visible = False: Me.PVKey_Flat3D_Color6.Tag = ""
        Me.PVKey_Flat3D_Color5.Visible = False: Me.PVKey_Flat3D_Color5.Tag = ""
        Me.PVKey_Flat3D_Color4.Visible = False: Me.PVKey_Flat3D_Color4.Tag = ""
        Me.PVKey_Flat3D_Color2.Visible = False: Me.PVKey_Flat3D_Color2.Tag = ""
        Me.PVKey_Flat3D_Color1.Visible = False: Me.PVKey_Flat3D_Color1.Tag = ""
        Me.PVKey_Flat3D_Color0.Visible = False: Me.PVKey_Flat3D_Color0.Tag = ""
        
        Me.PVKey_Flat3D_Value7_Unit(0).Visible = False: Me.PVKey_Flat3D_Value7_Unit(0).Tag = ""
        Me.PVKey_Flat3D_Value6_Unit(0).Visible = False: Me.PVKey_Flat3D_Value6_Unit(0).Tag = ""
        Me.PVKey_Flat3D_Value5_Unit(0).Visible = False: Me.PVKey_Flat3D_Value5_Unit(0).Tag = ""
        Me.PVKey_Flat3D_Value4_Unit(0).Visible = False: Me.PVKey_Flat3D_Value4_Unit(0).Tag = ""
        Me.PVKey_Flat3D_Value2_Unit(0).Visible = False: Me.PVKey_Flat3D_Value2_Unit(0).Tag = ""
        Me.PVKey_Flat3D_Value1_Unit(0).Visible = False: Me.PVKey_Flat3D_Value1_Unit(0).Tag = ""
        Me.PVKey_Flat3D_Value0_Unit(0).Visible = False: Me.PVKey_Flat3D_Value0_Unit(0).Tag = ""
        
        Me.PVKey_Flat3D_Value7_Unit(1).Visible = False: Me.PVKey_Flat3D_Value7_Unit(1).Tag = ""
        Me.PVKey_Flat3D_Value6_Unit(1).Visible = False: Me.PVKey_Flat3D_Value6_Unit(1).Tag = ""
        Me.PVKey_Flat3D_Value5_Unit(1).Visible = False: Me.PVKey_Flat3D_Value5_Unit(1).Tag = ""
        Me.PVKey_Flat3D_Value4_Unit(1).Visible = False: Me.PVKey_Flat3D_Value4_Unit(1).Tag = ""
        Me.PVKey_Flat3D_Value2_Unit(1).Visible = False: Me.PVKey_Flat3D_Value2_Unit(1).Tag = ""
        Me.PVKey_Flat3D_Value1_Unit(1).Visible = False: Me.PVKey_Flat3D_Value1_Unit(1).Tag = ""
        Me.PVKey_Flat3D_Value0_Unit(1).Visible = False: Me.PVKey_Flat3D_Value0_Unit(1).Tag = ""
        
        'Me.KeyShapeBoarder.Visible = False: Me.KeyShapeBoarder.Tag = ""
        Me.FlatKeyFrame.Visible = False: Me.FlatKeyFrame.Tag = ""
        
    End If

        
    'PCN???? The graphs that were not to be displayed were marked tag."Visible" this makes sure they are
    'not rendered                                   '
    For i = EndGraph To 6                           '
        GraphContainerArray(i).Tag = ""             '
        Me.LimitValue1(i).Tag = "": Me.LimitValue1(i).Visible = False
        Me.LimitValue2(i).Tag = "": Me.LimitValue2(i).Visible = False
        Me.GraphUnitLabel(i).Tag = "": Me.GraphUnitLabel(i).Visible = False
        Me.UnitLabel(i).Tag = "": Me.UnitLabel(i).Visible = False
        For J = 0 To 4
            Me.YScaleLabel((i * 5) + J).Tag = "": Me.YScaleLabel((i * 5) + J).Visible = False
        Next J
        Me.Shape2(i).Tag = "": Me.Shape2(i).Visible = False
    Next i                                          '
    '''''''''''''''''''''''''''''''''''''''''''''''''
        
    For i = StartGraph To EndGraph - 1
        If PreviewEndFrame >= PVDataNoOfLines Then PreviewEndFrame = PVDataNoOfLines
        Set GraphContainer = GraphContainerArray(i)
        Set GraphXScaleContainer = GraphXScaleContainerArray(i)
        
        
        Call DrawPVGraphsReport
        Call RenderToPrinter.RenderReport(Me, Printer, RenderScale)
        
        Call DrawPVGraphsReport
        GraphContainerArray(i).Visible = False
        GraphContainerArray(i).Tag = "" '
        GraphBackgroundShape(i).Visible = False
        GraphBackgroundShape(i).Tag = False
        
        PreviewStartFrame = PreviewStartFrame + NoFr
        PreviewEndFrame = PreviewEndFrame + NoFr
    Next i

    'Me.UserTitle.Tag = "Visible" PCN4277
    Call RenderToPrinter.RenderReport(Me, Printer, RenderScale)
    
    'PCNBLALBA 'PCNANT
    If PageToRender.name = "picReportPagePg1" Then
        Me.UserTitle.Left = LogoImage(0).Left + LogoImage(0).width
        Me.UserTitle.width = ((Me.lblTitle(0).width / 2) + Me.lblTitle(0).Left - Me.UserTitle.Left) * 2
        'Me.lblTitle(0).width
        
        
        'Me.lblTitle(0).Left: Me.UserTitle.width = Me.lblTitle(0).width
        
    End If
    If PageToRender.name = "picReportPagePg2" Then Me.UserTitle.Left = Me.lblTitle(1).Left: Me.UserTitle.width = Me.lblTitle(1).width
    
    
    Call RenderToPrinter.RenderSingleTextBox(Me.UserTitle, Printer, False)

    Me.UserTitle.Left = Me.lblTitle(0).Left
    Me.UserTitle.width = Me.lblTitle(0).width
    


    

Exit Sub
Err_Handler:
    Select Case Err
        Case 438, 387: Resume Next
        Case 428: MsgBox "Printer Error": Exit Sub 'Printer error
        Case Else: MsgBox Err & "-R4in14:" & Error$
        
    End Select
End Sub



Private Sub CommentsTextBox_Change(Index As Integer)
On Error GoTo Err_Handler
Dim SaveFailed As Boolean

If PVDFileName <> "" Then
    PipelineInfo.Comments = CommentsTextBox(0).text
    Call SaveToFilePipeAndConfigInfo("PipelineInfo", SaveFailed)
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-R4in15:" & Error$
End Sub

Private Sub ControlsBackPanel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


If Me.ControlHighlight.Visible Then Me.ControlHighlight.Visible = False
If Me.ControlHighlightDepressed.Visible Then Me.ControlHighlightDepressed.Visible = False
    
Exit Sub
Err_Handler:
    MsgBox Err & "-R4in16:" & Error$
End Sub

Private Sub ControlsReport_Click(Index As Integer)
On Error GoTo Err_Handler

Call ExecuteReportButton(Me, Index)
    
Exit Sub
Err_Handler:
    MsgBox Err & "-R4in17:" & Error$
End Sub

Private Sub ControlsReport_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


Me.ControlHighlightDepressed.Left = Me.ControlHighlight.Left
Me.ControlHighlightDepressed.Visible = True
Me.ControlHighlight.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-R4in18:" & Error$
End Sub

Private Sub ControlsReport_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ControlHighlight.Left = Me.ControlsReport(Index).Left - 100

If Me.ControlHighlight.Visible = False Then Me.ControlHighlight.Visible = True

    
Exit Sub
Err_Handler:
    MsgBox Err & "-R4in19:" & Error$
End Sub

Private Sub ControlsReport_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ControlHighlight.Visible = True
Me.ControlHighlightDepressed.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-R4in110:" & Error$
End Sub

Private Sub FloatingText_Change(Index As Integer)
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingText_Change(Me, Index)
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R4in111:" & Error$
    End Select
End Sub

Private Sub FloatingText_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingText_KeyPress(Me, Index, KeyAscii)
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R4in112:" & Error$
    End Select
End Sub

Private Sub FloatingText_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingText_MouseDown(Me, Index, Button, Shift, X, Y)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R4in113:" & Error$
    End Select
End Sub

Private Sub FloatingText_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingText_MouseMove(Me, Index, Button, Shift, X, Y)
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R4in114:" & Error$
    End Select
End Sub

Private Sub FloatingTextBackgroundColourMenu_Click()
On Error GoTo Err_Handler
    
    Call RenderToPrinter.FloatingTextBackgroundColourMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R4in115:" & Error$
    End Select
End Sub


Private Sub FloatingTextDefaultMenu_Click()
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingTextDefaultMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R4in116:" & Error$
    End Select
End Sub

Private Sub FloatingTextDeleteAllMenu_Click()
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingTextDeleteAllMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R4in117:" & Error$
    End Select
End Sub

Private Sub FloatingTextDeleteMenu_Click()
On Error GoTo Err_Handler
    
    Call RenderToPrinter.FloatingTextDeleteMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R4in118:" & Error$
    End Select
End Sub

Private Sub FloatingTextFontMenu_Click()
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingTextFontMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R4in119:" & Error$
    End Select
End Sub

Sub IntitialiseForm()
On Error GoTo Err_Handler

    Me.PageFramePictureBox.Visible = False
    Dim i As Long
    
    Set ImageGraphState(6).PictureImage = PrecisionVisionGraph.PrinterReportImage 'PCN4271
    

    
    RenderScale = 1
    Set PrintPreviewForm = Me
    
    Call RepositionAndRender
    'If Printer.PaperSize <> 1 Then Me.CommentsTextBox(0).Top = 13920 + 735
    If Printer.PaperSize <> 1 Then
        Me.CommentsTextBox(0).Top = 13920 + 235
    Else
        Me.CommentsTextBox(0).Top = 13920 - 500
    End If
    
    'AutoSetNoOfGraphs
    Call AutoSetNoOfGraphs
    
    'Set mouse icon for move
    PrintPreviewAction = "MoveAll"
    Call SetupReportMouseIcon(Me, 108)
    
    'If Printer.PaperSize <> 1 Then Me.CommentsTextBox(0).Top = 13920 + 735
    If Printer.PaperSize <> 1 Then
        Me.CommentsTextBox(0).Top = 13920 + 235
    Else
        Me.CommentsTextBox(0).Top = 13920 - 500
    End If
    
    Me.PageFramePictureBox.Visible = True
    Me.CmboPrinterList.Enabled = True

Exit Sub
Err_Handler:
    Select Case Err
        Case 438: Resume Next
        Case Else: MsgBox Err & "-R4in120:" & Error$
    End Select
End Sub



Public Sub Form_Load()
On Error GoTo Err_Handler
    Dim i As Integer
    
    Me.UserTitle.Font.Charset = LanguageCharset
    Me.FloatingText(0).Font.Charset = LanguageCharset
    
    Me.PageFramePictureBox.Visible = False
    Call GetPrinterList(Me)
    
    Me.Left = 0
    Me.width = ClearLineProfilerV6.width - 200
    Me.Top = 0
    Me.height = ClearLineProfilerV6.height - 500
    Me.Show
    
    Call ConvertLanguage(Me, Language) 'PCN4171

    Call GetPipeDetailsLabels(Me)
    Me.PageLabel(0).Caption = DisplayMessage("Page") & " 1"
    Me.PageLabel(1).Caption = DisplayMessage("Page") & " 2"
    
    PageFramePictureBox.Left = 0
    PageFramePictureBox.width = Me.width - PageFramePictureBox.Left
    PageFramePictureBox.height = Me.height
    PageFramePictureBox.Top = 650
    Me.ControlsBackPanel.width = Me.width
    Me.CloseReport.Left = Me.width - 750
    
    
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
    

    'Call AutoSetNoOfGraphs
    Call Me.IntitialiseForm
    
    
    
''    Me.PageFramePictureBox.Visible = False
''    Dim I As Long
''
''    Set ImageGraphState(6).PictureImage = PrecisionVisionGraph.PrinterReportImage 'PCN4271
''
''    I = PVReport4in1.Controls.Count
''    ReDim OriginalStateVisible(I)
''    ReDim OriginalStateTag(I)
''
''    Me.Left = 0
''    Me.width = ClearLineProfilerV6.width - 200
''    Me.Top = 0
''    Me.height = ClearLineProfilerV6.height - 500
''    Me.Show
''
''    Call ConvertLanguage(Me, Language) 'PCN4171
''
''    Call GetPipeDetailsLabels(Me)
''    Me.PageLabel(0).Caption = DisplayMessage("Page") & " 1"
''    Me.PageLabel(1).Caption = DisplayMessage("Page") & " 2"
''
''
''    For I = 1 To PVReport4in1.Controls.Count - 1
''        OriginalStateVisible(I) = PVReport4in1.Controls(I).Visible
''        OriginalStateTag(I) = PVReport4in1.Controls(I).Tag
''    Next I
''
''    PageFramePictureBox.Left = 0
''    PageFramePictureBox.width = Me.width - PageFramePictureBox.Left
''    PageFramePictureBox.height = Me.height
''    PageFramePictureBox.Top = 650
''    Me.ControlsBackPanel.width = Me.width
''    Me.CloseReport.Left = Me.width - 750
''
''    RenderScale = 1
''    Set PrintPreviewForm = Me
''
''    Call RepositionAndRender
''
''    'AutoSetNoOfGraphs
''    Call AutoSetNoOfGraphs
''
''    'Set mouse icon for move
''    PrintPreviewAction = "MoveAll"
''    Call SetupReportMouseIcon(Me, 108)
''    Me.PageFramePictureBox.Visible = True
''    Me.CmboPrinterList.Enabled = True
    'Me.UserTitle.text = Trim(Right(PipeObservations(97).PipeObs, Len(PipeObservations(97).PipeObs) - 13))
    
'    Me.UserTitle.text = UserTitleAnalysis 'PCN4433
    
    
Exit Sub
Err_Handler:
    Select Case Err
        Case 438, 382, 383, 393: Resume Next
        Case Else: MsgBox Err & "-R4in121:" & Error$
    
    End Select
End Sub

Private Sub RenderPage(PageToRender As PictureBox)
On Error GoTo Err_Handler
    Dim i As Integer
    Dim J As Integer
    Dim NG As Long
    Dim Pg1NG As Long
    Dim NoFr As Long
    Dim StartGraph As Long
    Dim EndGraph As Long
    
    Call RestoreOriginalState
    
    If PageToRender.name = "picReportPagePg1" Then
        ControlsPositionShunt(0).Tag = "Shunt": ControlsPositionShunt(1).Tag = "Egnore"
    Else
        ControlsPositionShunt(0).Tag = "Egnore": ControlsPositionShunt(1).Tag = "Shunt"
    End If
    
    Call Me.SetControlsPerPaperSize
    
    
    For i = 1 To PVReport4in1.Controls.Count - 1
        With PVReport4in1.Controls(i)
            If .name <> PageToRender.name And _
               .Container.name <> "PVReport4in1" And _
               .name <> "PageFramePictureBox" Then
                If .Container.name <> PageToRender.name Then
                   .Visible = False
                   .Tag = ""
                Else
                    .Visible = True
                    
                End If
                
            End If
        End With
    Next i
    PageToRender.Visible = True
    
    Set picReportPage = PageToRender

    PageToRender.Cls
    PageToRender.width = Printer.width * RenderScale
    PageToRender.height = Printer.height * RenderScale
    

    Call PositionReportControls
    Call FillOutPrintForm
    Call GraphSpecificSettings
    
    Me.UserTitle.Visible = False
    Me.UserTitle.Tag = ""
    
    Me.CommentsTextBox(0).Visible = False
    Me.CommentsTextBox(0).Tag = ""
        
    FloatingTextDefault.Visible = False
    FloatingTextDefault.Tag = ""
    
    FloatingText(0).Visible = False
    FloatingText(0).Tag = ""
    
    If PageToRender.name = "picReportPagePg2" Then
        PVReport4in1.Explination(0).Visible = False
        PVReport4in1.Explination(0).Tag = ""
    End If

    Call MarkForPrinting

    'Call RenderToPrinter.RenderReport(Me, PageToRender, RenderScale)

    ScreenDrawingType = 2
    ScreenDrawingOrientation = 1
        
    If NoOfGraphPanels <= 0 Then NoOfGraphPanels = 1
    NG = NoOfGraphPanels
    NoFr = PVDataNoOfLines / NG
    
    PreviewStartFrame = 0
    PreviewEndFrame = NoFr
    
    
    'Call DrawPVGraphsReport
    
    
    

    
    If PageToRender.name = "picReportPagePg1" Then
         EndGraph = NG
         StartGraph = 0
         If NG > 3 Then EndGraph = 3
         PreviewStartFrame = 0
         PreviewEndFrame = NoFr
    Else
        If NG <= 3 Then ScreenDrawingType = 0:  ScreenDrawingOrientation = 0: Exit Sub
        StartGraph = 3
        EndGraph = NG
        PreviewStartFrame = 3 * (NoFr + 1)
        PreviewEndFrame = 4 * (NoFr + 1)
        
        'PCN4417
        Me.PVKey_Flat3D_Color7.Visible = False: Me.PVKey_Flat3D_Color7.Tag = ""
        Me.PVKey_Flat3D_Color6.Visible = False: Me.PVKey_Flat3D_Color6.Tag = ""
        Me.PVKey_Flat3D_Color5.Visible = False: Me.PVKey_Flat3D_Color5.Tag = ""
        Me.PVKey_Flat3D_Color4.Visible = False: Me.PVKey_Flat3D_Color4.Tag = ""
        Me.PVKey_Flat3D_Color2.Visible = False: Me.PVKey_Flat3D_Color2.Tag = ""
        Me.PVKey_Flat3D_Color1.Visible = False: Me.PVKey_Flat3D_Color1.Tag = ""
        Me.PVKey_Flat3D_Color0.Visible = False: Me.PVKey_Flat3D_Color0.Tag = ""
        
        Me.PVKey_Flat3D_Value7_Unit(0).Visible = False: Me.PVKey_Flat3D_Value7_Unit(0).Tag = ""
        Me.PVKey_Flat3D_Value6_Unit(0).Visible = False: Me.PVKey_Flat3D_Value6_Unit(0).Tag = ""
        Me.PVKey_Flat3D_Value5_Unit(0).Visible = False: Me.PVKey_Flat3D_Value5_Unit(0).Tag = ""
        Me.PVKey_Flat3D_Value4_Unit(0).Visible = False: Me.PVKey_Flat3D_Value4_Unit(0).Tag = ""
        Me.PVKey_Flat3D_Value2_Unit(0).Visible = False: Me.PVKey_Flat3D_Value2_Unit(0).Tag = ""
        Me.PVKey_Flat3D_Value1_Unit(0).Visible = False: Me.PVKey_Flat3D_Value1_Unit(0).Tag = ""
        Me.PVKey_Flat3D_Value0_Unit(0).Visible = False: Me.PVKey_Flat3D_Value0_Unit(0).Tag = ""
        
        Me.PVKey_Flat3D_Value7_Unit(1).Visible = False: Me.PVKey_Flat3D_Value7_Unit(1).Tag = ""
        Me.PVKey_Flat3D_Value6_Unit(1).Visible = False: Me.PVKey_Flat3D_Value6_Unit(1).Tag = ""
        Me.PVKey_Flat3D_Value5_Unit(1).Visible = False: Me.PVKey_Flat3D_Value5_Unit(1).Tag = ""
        Me.PVKey_Flat3D_Value4_Unit(1).Visible = False: Me.PVKey_Flat3D_Value4_Unit(1).Tag = ""
        Me.PVKey_Flat3D_Value2_Unit(1).Visible = False: Me.PVKey_Flat3D_Value2_Unit(1).Tag = ""
        Me.PVKey_Flat3D_Value1_Unit(1).Visible = False: Me.PVKey_Flat3D_Value1_Unit(1).Tag = ""
        Me.PVKey_Flat3D_Value0_Unit(1).Visible = False: Me.PVKey_Flat3D_Value0_Unit(1).Tag = ""
        
        'Me.KeyShapeBoarder.Visible = False: Me.KeyShapeBoarder.Tag = ""
        Me.FlatKeyFrame.Visible = False: Me.FlatKeyFrame.Tag = ""
        
        
    End If
    
    'PCN???? The graphs that were not to be displayed were marked tag."Visible" this makes sure they are
    'not rendered                                   '
    For i = EndGraph To 6                           '
        GraphContainerArray(i).Tag = ""             '
        Me.LimitValue1(i).Tag = "": Me.LimitValue1(i).Visible = False
        Me.LimitValue2(i).Tag = "": Me.LimitValue2(i).Visible = False
        Me.GraphUnitLabel(i).Tag = "": Me.GraphUnitLabel(i).Visible = False
        Me.UnitLabel(i).Tag = "": Me.UnitLabel(i).Visible = False
        For J = 0 To 4
            Me.YScaleLabel((i * 5) + J).Tag = "": Me.YScaleLabel((i * 5) + J).Visible = False
        Next J
        Me.Shape2(i).Tag = "": Me.Shape2(i).Visible = False
        
    Next i                                          '
    '''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    For i = StartGraph To EndGraph - 1
'     If i <> 0 Then
        If PreviewEndFrame >= PVDataNoOfLines Then PreviewEndFrame = PVDataNoOfLines
        Set GraphContainer = GraphContainerArray(i)
        Set GraphXScaleContainer = GraphXScaleContainerArray(i)

        
        Call DrawPVGraphsReport
        Call RenderToPrinter.RenderReport(Me, PageToRender, RenderScale)
        
        Call DrawPVGraphsReport
        GraphContainerArray(i).Visible = False
        GraphContainerArray(i).Tag = "" '
        GraphBackgroundShape(i).Visible = False
        GraphBackgroundShape(i).Tag = False
        
        PreviewStartFrame = PreviewStartFrame + NoFr
        PreviewEndFrame = PreviewEndFrame + NoFr
'    End If
        
    Next i
    
    
    Call RenderToPrinter.RenderReport(Me, PageToRender, RenderScale)

    

    
    'Call RenderToPrinter.RenderReport(Me, PageToRender, RenderScale)
    
    ScreenDrawingType = 0
    ScreenDrawingOrientation = 0
Exit Sub
Err_Handler:
    Select Case Err
        Case 438, 387: Resume Next
        Case Else: MsgBox Err & "-R4in122:" & Error$

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

Printer.Orientation = vbPRORPortrait

ManualOrientationSet:

Printer.PrintQuality = vbPRPQHigh
CommonDialogBox.Orientation = cdlPortrait



Exit Sub
ManualOrientation:
On Error GoTo Err_Handler

Dim originalheight
Dim originalwidth

originalheight = Printer.height
originalwidth = Printer.width

If originalheight > originalwidth Then GoTo ManualOrientationSet


Printer.height = originalwidth
Printer.width = originalheight


GoTo ManualOrientationSet
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R4in123:" & Error$
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
        Case Else: MsgBox Err & "-R4in124:" & Error$
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
PipelineDetails.GeneralComments.text = Trim(PipelineInfo.Comments)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R4in125:" & Error$
    End Select
End Sub

Private Sub PageFramePictureBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


If Me.ControlHighlight.Visible Then Me.ControlHighlight.Visible = False
If Me.ControlHighlightDepressed.Visible Then Me.ControlHighlightDepressed.Visible = False
    
Exit Sub
Err_Handler:
    MsgBox Err & "-R4in126:" & Error$
End Sub

Private Sub PageHScroll_Change()
On Error GoTo Err_Handler
Dim ReportPageSpacing As Integer

ReportPageSpacing = picReportPagePg2.Left - picReportPagePg1.Left
If PageHScroll.value = 1 Then
    picReportPagePg1.Left = 100
    picReportPagePg2.Left = picReportPagePg1.Left + ReportPageSpacing
    Call SetNoOfPages
ElseIf PageHScroll.value = 2 Then
    ReportPageSpacing = picReportPagePg2.Left - picReportPagePg1.Left
    picReportPagePg2.Left = 200
    picReportPagePg1.Left = picReportPagePg2.Left - ReportPageSpacing
    Call SetNoOfPages
End If

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R4in127:" & Error$
    End Select
End Sub

Private Sub NoGraphsDisplayVScroll_Change()
On Error GoTo Err_Handler
Dim NG As Integer
Dim MaxNoOfGraphPanels As Integer

'MaxNoOfGraphPanels = 7

NG = -1 * NoGraphsDisplayVScroll.value
If NG > 7 Or NG < 1 Then Exit Sub
NoOfGraphPanels = NG
NoGraphsDisplay.Caption = NoOfGraphPanels

RenderScale = 1
Call RepositionAndRender
'If Printer.PaperSize <> 1 Then Me.CommentsTextBox(0).Top = 13920 + 735
    If Printer.PaperSize <> 1 Then
        Me.CommentsTextBox(0).Top = 13920 + 235
    Else
        Me.CommentsTextBox(0).Top = 13920 - 500
    End If

'Call PageHScroll_Change
Call SetNoOfPages
Me.PageFramePictureBox.Visible = True 'PCN????

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R4in128:" & Error$
    End Select
End Sub

Private Sub picReportPagePg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Call ReportPageMouseDown(Me, picReportPagePg1, Button, Shift, X, Y) 'PCN4193

Exit Sub
Err_Handler:
    Select Case Err
        Case 438
            Resume Next
        Case Else: MsgBox Err & "-R4in129:" & Error$
    End Select
End Sub

Private Sub picReportPagePg2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


Call ReportPageMouseDown(Me, picReportPagePg2, Button, Shift, X, Y) 'PCN4193

Exit Sub
Err_Handler:
    Select Case Err
        Case 438
            Resume Next
        Case Else: MsgBox Err & "-R4in130:" & Error$
    End Select
End Sub

Private Sub picReportPagePg1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    If ReportMouseDown Then
        picReportPagePg1.Left = picReportPagePg1.Left + X - ReportMouseX
        picReportPagePg1.Top = picReportPagePg1.Top + Y - ReportMouseY
        picReportPagePg2.Left = picReportPagePg2.Left + X - ReportMouseX
        picReportPagePg2.Top = picReportPagePg2.Top + Y - ReportMouseY
        
    End If
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R4in131:" & Error$
    End Select
End Sub

Private Sub picReportPagePg2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    If ReportMouseDown Then
        picReportPagePg1.Left = picReportPagePg1.Left + X - ReportMouseX
        picReportPagePg1.Top = picReportPagePg1.Top + Y - ReportMouseY
        picReportPagePg2.Left = picReportPagePg2.Left + X - ReportMouseX
        picReportPagePg2.Top = picReportPagePg2.Top + Y - ReportMouseY
        
    End If
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R4in132:" & Error$
    End Select
End Sub

Private Sub picReportPagePg1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

ReportMouseDown = False
If PrintPreviewAction = "MoveAll" Then
    Call SetupReportMouseIcon(PVReport4in1, 108)
End If

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R4in133:" & Error$
    End Select
End Sub

Private Sub picReportPagePg2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

ReportMouseDown = False
If PrintPreviewAction = "MoveAll" Then
    Call SetupReportMouseIcon(PVReport4in1, 108)
End If

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R4in134:" & Error$
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
    Dim ProfileSlicePosition As Single
    
    Dim LeftLimit As Double
    Dim RightLimit As Double
    Dim Limit1Top(6) As Single
    Dim Limit2Top(6) As Single
    
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
    
    LeftLimit = ConvertUnitByGraph(PVXScaleLimitPerL, 0, DisplayUnits)                                                 '
    RightLimit = ConvertUnitByGraph(PVXScaleLimitPerR, 0, DisplayUnits)                                             '
    
    For i = 0 To 6
        Me.LimitValue1(i) = Format(LeftLimit, "###0.0")
        Me.LimitValue2(i) = Format(RightLimit, "###0.0")
    
        Limit1Top(i) = PVGraphXScaleTwipFromUnit(0, PVXScaleLimitPerL)                         '
        Limit2Top(i) = PVGraphXScaleTwipFromUnit(0, PVXScaleLimitPerR)                         '
        Limit1Top(i) = Limit1Top(i) * (Me.GraphContainerArray(i).height / PrecisionVisionGraph.PVGraphImage(0).width)
        Limit2Top(i) = Limit2Top(i) * (Me.GraphContainerArray(i).height / PrecisionVisionGraph.PVGraphImage(0).width)
        Limit1Top(i) = Me.GraphContainerArray(i).height - Limit1Top(i)                                    '
        Limit2Top(i) = Me.GraphContainerArray(i).height - Limit2Top(i)                                    '
        Me.LimitValue1(i).Top = Limit1Top(i) + Me.GraphContainerArray(i).Top                              '
        Me.LimitValue2(i).Top = Limit2Top(i) + Me.GraphContainerArray(i).Top - Me.LimitValue1(i).height      '
    Next i
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    With PVReport4in1
        .SiteIDData(0).Caption = Trim(PipelineInfo.SiteID)
        .StartNodeData(0).Caption = Trim(PipelineInfo.StartName)
        .FinishNodeData(0).Caption = Trim(PipelineInfo.FinishName)
        .StartLocationData(0).Caption = Trim(PipelineInfo.StartLocation)
        .FinishLocationData(0).Caption = Trim(PipelineInfo.FinishLocation)
        .AssetNoData(0).Caption = Trim(PipelineInfo.AssetNo)
        .CityData(0).Caption = Trim(PipelineInfo.City)
        .PipeLengthData(0).Caption = Trim(PipelineInfo.PipeLength)
        .PipeMaterialData(0).Caption = Trim(PipelineInfo.Material)
        .PipeDiameterData(0).Caption = Trim(PipelineInfo.IntDiameter)
        .DateLabelData(0).Caption = CStr(PipelineInfo.Date)
        .PrintedData(0).Caption = CStr(Date)
        .PrintedData(1).Caption = .PrintedData(0).Caption
        .PhData.Caption = PhoneNo
        .LogoImage(0).Picture = LoadPicture(CompanyLogoPath)
        .CompanyNameLabel = CompanyName
        
        For i = 0 To 6
            .YScaleLabel((i * 5) + 0) = YScaleLabel0
            .YScaleLabel((i * 5) + 1) = YScaleLabel1
            .YScaleLabel((i * 5) + 2) = YScaleLabel2
            .YScaleLabel((i * 5) + 3) = YScaleLabel3
            .YScaleLabel((i * 5) + 4) = YScaleLabel4
        
            .UnitLabel(i) = PrecisionVisionGraph.PVXScaleLabelUnits(0)
        Next i
            
'        .UpperLimitData = RightLimit
'        .LowerLimitData = LeftLimit
        
        .lblTitle(0).Caption = PrecisionVisionGraph.Label_GraphName(0) & " " & DisplayMessage("Analysis Report")
        .lblTitle(1).Caption = .lblTitle(0).Caption
        
        .CommentsTextBox(0).text = PipelineInfo.Comments
        
        'By default make all lable no background
        For i = 0 To .Controls.Count - 1
            ControlType = TypeName(.Controls(i))
            Select Case ControlType
                Case "Label": .Controls(i).BackStyle = 0
            End Select
        Next i
        
        If MeasurementUnits = "mm" Then
            .PipeLengthUnitLabel(0) = "m"
            .PipeDiameterUnitLabel(0) = "mm"
            i = 0
            For i = 0 To 6
                .GraphUnitLabel(i) = "m"
            Next i
        Else
            .PipeLengthUnitLabel(0) = "ft"
            .PipeDiameterUnitLabel(0) = "in"
            i = 0
            For i = 0 To 6
                .GraphUnitLabel(i) = "ft"
            Next i
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
        Case Else: MsgBox Err & "-R4in135:" & Error$
    End Select
    
End Sub

Sub GraphSpecificSettings()
On Error GoTo Err_Handler
    Dim i As Integer
    Dim ColourKeysLeftShift As Single
    Dim ColourKeysTopShift As Single
    
    If ScreenDrawing.ImageGraphState(0).GraphType = "Ovality" Then
 
        With PVReport4in1
        For i = 0 To 6

            .YScaleLabel((i * 5) + 4) = ""
        
            .UnitLabel(i) = PrecisionVisionGraph.PVXScaleLabelUnits(0)
            Me.LimitValue2(i).Visible = False: Me.LimitValue2(i).Tag = "" 'PCN4253
        Next i
        End With

        PVReport4in1.Explination(0).Visible = True
        PVReport4in1.Explination(0).Tag = ""
        
        'PCN4324 if ovality display is less than zero, dont display
        

        
        For i = 0 To 34
            If Trim(Me.YScaleLabel(i).Caption) < 0 Then Me.YScaleLabel(i).Visible = False: Me.YScaleLabel(i).Tag = ""
        Next i
        
        
    ElseIf ScreenDrawing.ImageGraphState(0).GraphType = "Flat" Then
        With PVReport4in1
        For i = 0 To 6
            .UnitLabel(i).Visible = False
            .UnitLabel(i).Tag = ""
            .YScaleLabel((i * 5) + 0) = 6
            .YScaleLabel((i * 5) + 1) = 3
            .YScaleLabel((i * 5) + 2) = 12
            .YScaleLabel((i * 5) + 3) = 9
            .YScaleLabel((i * 5) + 4) = 6
        
            .UnitLabel(i) = PrecisionVisionGraph.PVXScaleLabelUnits(0)
            
            Me.LimitValue1(i).Visible = False: Me.LimitValue1(i).Tag = "" 'PCN4253
            Me.LimitValue2(i).Visible = False: Me.LimitValue2(i).Tag = "" 'PCN4253
            
            
        Next i
        End With
        Me.FractileLabel.Visible = False: Me.FractileLabelData.Visible = False 'PCN4296
        Me.FractileLabel.Tag = "": Me.FractileLabelData.Tag = ""

        ColourKeysLeftShift = (FlatKeyFrame.Left + FlatKeyFrame.width) - (Me.CommentsTextBox(0).Left + CommentsTextBox(0).width) + 60
        ColourKeysTopShift = Me.FlatKeyFrame.Top - Me.CommentsLabel(0).Top + 170
        
        'ColourKeysLeftShift = (KeyShapeBoarder.Left + KeyShapeBoarder.width) - (Me.CommentsTextBox(0).Left + CommentsTextBox(0).width) + 60
        'ColourKeysTopShift = Me.KeyShapeBoarder.Top - Me.CommentsLabel(0).Top + 170
        
        Me.PVKey_Flat3D_Color7.Left = Me.PVKey_Flat3D_Color7.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Color7.Top = Me.PVKey_Flat3D_Color7.Top - ColourKeysTopShift
        
        Me.PVKey_Flat3D_Color6.Left = Me.PVKey_Flat3D_Color6.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Color6.Top = Me.PVKey_Flat3D_Color6.Top - ColourKeysTopShift
        
        Me.PVKey_Flat3D_Color5.Left = Me.PVKey_Flat3D_Color5.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Color5.Top = Me.PVKey_Flat3D_Color5.Top - ColourKeysTopShift
        
        Me.PVKey_Flat3D_Color4.Left = Me.PVKey_Flat3D_Color4.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Color4.Top = Me.PVKey_Flat3D_Color4.Top - ColourKeysTopShift
        
        Me.PVKey_Flat3D_Color2.Left = Me.PVKey_Flat3D_Color2.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Color2.Top = Me.PVKey_Flat3D_Color2.Top - ColourKeysTopShift
        
        Me.PVKey_Flat3D_Color1.Left = Me.PVKey_Flat3D_Color1.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Color1.Top = Me.PVKey_Flat3D_Color1.Top - ColourKeysTopShift
        
        Me.PVKey_Flat3D_Color0.Left = Me.PVKey_Flat3D_Color0.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Color0.Top = Me.PVKey_Flat3D_Color0.Top - ColourKeysTopShift
        
        
        Me.PVKey_Flat3D_Value7.Left = Me.PVKey_Flat3D_Value7.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value7.Top = Me.PVKey_Flat3D_Value7.Top - ColourKeysTopShift
        
        Me.PVKey_Flat3D_Value6.Left = Me.PVKey_Flat3D_Value6.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value6.Top = Me.PVKey_Flat3D_Value6.Top - ColourKeysTopShift
        
        Me.PVKey_Flat3D_Value5.Left = Me.PVKey_Flat3D_Value5.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value5.Top = Me.PVKey_Flat3D_Value5.Top - ColourKeysTopShift
        
        Me.PVKey_Flat3D_Value4.Left = Me.PVKey_Flat3D_Value4.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value4.Top = Me.PVKey_Flat3D_Value4.Top - ColourKeysTopShift
        
        Me.PVKey_Flat3D_Value2.Left = Me.PVKey_Flat3D_Value2.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value2.Top = Me.PVKey_Flat3D_Value2.Top - ColourKeysTopShift
        
        Me.PVKey_Flat3D_Value1.Left = Me.PVKey_Flat3D_Value1.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value1.Top = Me.PVKey_Flat3D_Value1.Top - ColourKeysTopShift
        
        Me.PVKey_Flat3D_Value0.Left = Me.PVKey_Flat3D_Value0.Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value0.Top = Me.PVKey_Flat3D_Value0.Top - ColourKeysTopShift
        
        
        Me.PVKey_Flat3D_Value7_Unit(0).Left = Me.PVKey_Flat3D_Value7_Unit(0).Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value7_Unit(0).Top = Me.PVKey_Flat3D_Value7_Unit(0).Top - ColourKeysTopShift
        
        Me.PVKey_Flat3D_Value6_Unit(0).Left = Me.PVKey_Flat3D_Value6_Unit(0).Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value6_Unit(0).Top = Me.PVKey_Flat3D_Value6_Unit(0).Top - ColourKeysTopShift
        
        Me.PVKey_Flat3D_Value5_Unit(0).Left = Me.PVKey_Flat3D_Value5_Unit(0).Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value5_Unit(0).Top = Me.PVKey_Flat3D_Value5_Unit(0).Top - ColourKeysTopShift
        
        Me.PVKey_Flat3D_Value4_Unit(0).Left = Me.PVKey_Flat3D_Value4_Unit(0).Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value4_Unit(0).Top = Me.PVKey_Flat3D_Value4_Unit(0).Top - ColourKeysTopShift
        
        Me.PVKey_Flat3D_Value2_Unit(0).Left = Me.PVKey_Flat3D_Value2_Unit(0).Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value2_Unit(0).Top = Me.PVKey_Flat3D_Value2_Unit(0).Top - ColourKeysTopShift
        
        Me.PVKey_Flat3D_Value1_Unit(0).Left = Me.PVKey_Flat3D_Value1_Unit(0).Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value1_Unit(0).Top = Me.PVKey_Flat3D_Value1_Unit(0).Top - ColourKeysTopShift
        
        Me.PVKey_Flat3D_Value0_Unit(0).Left = Me.PVKey_Flat3D_Value0_Unit(0).Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value0_Unit(0).Top = Me.PVKey_Flat3D_Value0_Unit(0).Top - ColourKeysTopShift
        
        Me.PVKey_Flat3D_Value7_Unit(1).Left = Me.PVKey_Flat3D_Value7_Unit(1).Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value7_Unit(1).Top = Me.PVKey_Flat3D_Value7_Unit(1).Top - ColourKeysTopShift
        
        Me.PVKey_Flat3D_Value6_Unit(1).Left = Me.PVKey_Flat3D_Value6_Unit(1).Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value6_Unit(1).Top = Me.PVKey_Flat3D_Value6_Unit(1).Top - ColourKeysTopShift
        
        Me.PVKey_Flat3D_Value5_Unit(1).Left = Me.PVKey_Flat3D_Value5_Unit(1).Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value5_Unit(1).Top = Me.PVKey_Flat3D_Value5_Unit(1).Top - ColourKeysTopShift
        
        Me.PVKey_Flat3D_Value4_Unit(1).Left = Me.PVKey_Flat3D_Value4_Unit(1).Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value4_Unit(1).Top = Me.PVKey_Flat3D_Value4_Unit(1).Top - ColourKeysTopShift
        
        Me.PVKey_Flat3D_Value2_Unit(1).Left = Me.PVKey_Flat3D_Value2_Unit(1).Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value2_Unit(1).Top = Me.PVKey_Flat3D_Value2_Unit(1).Top - ColourKeysTopShift
        
        Me.PVKey_Flat3D_Value1_Unit(1).Left = Me.PVKey_Flat3D_Value1_Unit(1).Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value1_Unit(1).Top = Me.PVKey_Flat3D_Value1_Unit(1).Top - ColourKeysTopShift
        
        Me.PVKey_Flat3D_Value0_Unit(1).Left = Me.PVKey_Flat3D_Value0_Unit(1).Left - ColourKeysLeftShift
        Me.PVKey_Flat3D_Value0_Unit(1).Top = Me.PVKey_Flat3D_Value0_Unit(1).Top - ColourKeysTopShift
        
        'Me.KeyShapeBoarder.Left = Me.KeyShapeBoarder.Left - ColourKeysLeftShift
        'Me.KeyShapeBoarder.Top = Me.KeyShapeBoarder.Top - ColourKeysTopShift
        
        'ANT' Me.FlatKeyFrame.Left = Me.FlatKeyFrame.Left - ColourKeysLeftShift
        'ANT' Me.FlatKeyFrame.Top = Me.FlatKeyFrame.Top - ColourKeysTopShift
'        Me.KeyLine.x1 = Me.KeyLine.x1 - ColourKeysLeftShift
'        Me.KeyLine.x2 = Me.KeyLine.x2 - ColourKeysLeftShift
'        Me.KeyLine.y1 = Me.KeyLine.y1 - ColourKeysTopShift
'        Me.KeyLine.y2 = Me.KeyLine.y2 - ColourKeysTopShift
        
        
        Me.RadiusLabel(0).Left = Me.RadiusLabel(0).Left - ColourKeysLeftShift
        Me.RadiusLabel(1).Left = Me.RadiusLabel(1).Left - ColourKeysLeftShift
        Me.RadiusLabel(0).Top = Me.RadiusLabel(0).Top - ColourKeysTopShift
        Me.RadiusLabel(1).Top = Me.RadiusLabel(1).Top - ColourKeysTopShift
        
        Me.DiameterLabel(0).Left = Me.DiameterLabel(0).Left - ColourKeysLeftShift
        Me.DiameterLabel(1).Left = Me.DiameterLabel(1).Left - ColourKeysLeftShift
        Me.DiameterLabel(0).Top = Me.DiameterLabel(0).Top - ColourKeysTopShift
        Me.DiameterLabel(1).Top = Me.DiameterLabel(1).Top - ColourKeysTopShift
        
'        If Printer.PaperSize = 1 Then
'            ColourKeysLeftShift = ColourKeysLeftShift - 500
'        Else
'            ColourKeysLeftShift = ColourKeysLeftShift - 900
'        End If
        
         Me.FlatKeyFrame.Visible = False: Me.FlatKeyFrame.Tag = ""
        
        Me.RadiusLabel(0).Visible = True: RadiusLabel(0).Tag = "Visible"
        Me.RadiusLabel(1).Visible = True: RadiusLabel(1).Tag = "Visible"
        
        Me.DiameterLabel(0).Visible = True: DiameterLabel(0).Tag = "Visible"
        Me.DiameterLabel(1).Visible = True: DiameterLabel(1).Tag = "Visible"
                
        Me.CommentsTextBox(0).width = Me.CommentsTextBox(0).width - ColourKeysLeftShift
        
        
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
        
        Me.PVKey_Flat3D_Value7_Unit(0).Visible = True: Me.PVKey_Flat3D_Value7_Unit(0).Tag = "Visible"
        Me.PVKey_Flat3D_Value6_Unit(0).Visible = True: Me.PVKey_Flat3D_Value6_Unit(0).Tag = "Visible"
        Me.PVKey_Flat3D_Value5_Unit(0).Visible = True: Me.PVKey_Flat3D_Value5_Unit(0).Tag = "Visible"
        Me.PVKey_Flat3D_Value4_Unit(0).Visible = True: Me.PVKey_Flat3D_Value4_Unit(0).Tag = "Visible"
        Me.PVKey_Flat3D_Value2_Unit(0).Visible = True: Me.PVKey_Flat3D_Value2_Unit(0).Tag = "Visible"
        Me.PVKey_Flat3D_Value1_Unit(0).Visible = True: Me.PVKey_Flat3D_Value1_Unit(0).Tag = "Visible"
        Me.PVKey_Flat3D_Value0_Unit(0).Visible = True: Me.PVKey_Flat3D_Value0_Unit(0).Tag = "Visible"
        
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

            Me.KeyLine.Visible = False: Me.KeyLine.Tag = ""
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
            RadiusLabel(0).Left = RadiusLabel(0).Left + 700
            RadiusLabel(1).Left = RadiusLabel(1).Left + 700
            
            'ColourKeyBoarder.width = ColourKeyBoarder.width - 700
            'ColourKeyBoarder.Left = ColourKeyBoarder.Left + 700
'            FlatKeyFrame.width = FlatKeyFrame.width - 700
'            FlatKeyFrame.Left = FlatKeyFrame.Left + 700
            'Me.GraphContainer.width = Me.GraphContainer.width + 700
            'Me.Shape2.width = Me.Shape2.width + 700
            'Me.GraphUnitLbl.Left = Me.GraphUnitLbl.Left + 700
            'Me.GraphXScaleContainer.width = Me.GraphXScaleContainer.width + 700
            'Me.GraphBackgroundShape.width = Me.GraphBackgroundShape.width + 700
        
        End If
        
       
    End If
    
    If MedianFlat And PVDFileName <> "" Then 'PCN5186
        Me.FractileLabel.Visible = False: Me.FractileLabelData.Visible = False 'PCN4296
        Me.FractileLabel.Tag = "": Me.FractileLabelData.Tag = ""
    End If
    
    If ScreenDrawing.ImageGraphState(0).GraphType <> "Ovality" Then
        PVReport4in1.Explination(0).Visible = False
        PVReport4in1.Explination(0).Tag = ""
    End If
    
    If ScreenDrawing.ImageGraphState(0).GraphType <> "Flat" Then
        Me.PVKey_Flat3D_Color7.Visible = False: Me.PVKey_Flat3D_Color7.Tag = ""
        Me.PVKey_Flat3D_Color6.Visible = False: Me.PVKey_Flat3D_Color6.Tag = ""
        Me.PVKey_Flat3D_Color5.Visible = False: Me.PVKey_Flat3D_Color5.Tag = ""
        Me.PVKey_Flat3D_Color4.Visible = False: Me.PVKey_Flat3D_Color4.Tag = ""
        Me.PVKey_Flat3D_Color2.Visible = False: Me.PVKey_Flat3D_Color2.Tag = ""
        Me.PVKey_Flat3D_Color1.Visible = False: Me.PVKey_Flat3D_Color1.Tag = ""
        Me.PVKey_Flat3D_Color0.Visible = False: Me.PVKey_Flat3D_Color0.Tag = ""
        
      
        Me.PVKey_Flat3D_Value7_Unit(0).Visible = False: Me.PVKey_Flat3D_Value7_Unit(0).Tag = ""
        Me.PVKey_Flat3D_Value6_Unit(0).Visible = False: Me.PVKey_Flat3D_Value6_Unit(0).Tag = ""
        Me.PVKey_Flat3D_Value5_Unit(0).Visible = False: Me.PVKey_Flat3D_Value5_Unit(0).Tag = ""
        Me.PVKey_Flat3D_Value4_Unit(0).Visible = False: Me.PVKey_Flat3D_Value4_Unit(0).Tag = ""
        Me.PVKey_Flat3D_Value2_Unit(0).Visible = False: Me.PVKey_Flat3D_Value2_Unit(0).Tag = ""
        Me.PVKey_Flat3D_Value1_Unit(0).Visible = False: Me.PVKey_Flat3D_Value1_Unit(0).Tag = ""
        Me.PVKey_Flat3D_Value0_Unit(0).Visible = False: Me.PVKey_Flat3D_Value0_Unit(0).Tag = ""
        
        Me.PVKey_Flat3D_Value7_Unit(1).Visible = False: Me.PVKey_Flat3D_Value7_Unit(1).Tag = ""
        Me.PVKey_Flat3D_Value6_Unit(1).Visible = False: Me.PVKey_Flat3D_Value6_Unit(1).Tag = ""
        Me.PVKey_Flat3D_Value5_Unit(1).Visible = False: Me.PVKey_Flat3D_Value5_Unit(1).Tag = ""
        Me.PVKey_Flat3D_Value4_Unit(1).Visible = False: Me.PVKey_Flat3D_Value4_Unit(1).Tag = ""
        Me.PVKey_Flat3D_Value2_Unit(1).Visible = False: Me.PVKey_Flat3D_Value2_Unit(1).Tag = ""
        Me.PVKey_Flat3D_Value1_Unit(1).Visible = False: Me.PVKey_Flat3D_Value1_Unit(1).Tag = ""
        Me.PVKey_Flat3D_Value0_Unit(1).Visible = False: Me.PVKey_Flat3D_Value0_Unit(1).Tag = ""
        
        'Me.KeyShapeBoarder.Visible = False: Me.KeyShapeBoarder.Tag = ""
        Me.FlatKeyFrame.Visible = False: Me.FlatKeyFrame.Tag = ""
        
    End If
    
'    If ScreenDrawing.ImageGraphState(0).GraphType = "Inclination" Then
'        Me.LimitLinesLabel.Visible = False: Me.LimitLinesLabel.Tag = ""
'        Me.UpperLimitLabel.Visible = False: Me.UpperLimitLabel.Tag = ""
'        Me.LowerLimitLabel.Visible = False: Me.LowerLimitLabel.Tag = ""
'        Me.UpperLimitData.Visible = False: UpperLimitData.Tag = ""
'        Me.LowerLimitData = False: LowerLimitData.Tag = ""
'        Me.UperLimitData.Visible = False: Me.UperLimitData.Tag = ""
'    End If
     
     Dim UserTittleCaption As String
    Dim TittleParse As String
        
    TittleParse = "Analysis_" & ScreenDrawing.ImageGraphState(0).GraphType & "="
    Call GetINI_ParameterInfoOnly(MyFile, TittleParse, UserTittleCaption)
    Me.UserTitle.text = UserTittleCaption
    
    For i = 0 To 6
        If Me.LimitValue2(i).Top + Me.LimitValue2(i).height < Me.GraphContainerArray(i).Top Then Me.LimitValue2(i).Visible = False: Me.LimitValue2(i).Tag = ""
        If Me.LimitValue1(i).Top > Me.GraphContainerArray(i).Top + Me.GraphContainerArray(i).height Then Me.LimitValue1(i).Visible = False: Me.LimitValue2(i).Tag = ""
    Next i
    
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R4in136:" & Error$

    End Select
   
End Sub

Sub MarkForPrinting()
On Error GoTo Err_Handler

Dim i As Integer
Dim ControlType As String



'Draw renderings first that are marked back
For i = 1 To PVReport4in1.Controls.Count - 1
    
    With PVReport4in1.Controls(i)
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
        Case Else: MsgBox Err & "-R4in137:" & Error$
    End Select
End Sub




Private Sub ScaleButton05_Click()
On Error GoTo Err_Handler
    RenderScale = 0.5
    Call RepositionAndRender
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R4in138:" & Error$
    End Select
End Sub

Private Sub ScaleButton10_Click()
On Error GoTo Err_Handler
    RenderScale = 1
    Call RepositionAndRender
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R4in139:" & Error$
    End Select
End Sub

Private Sub ScaleButton15_Click()
On Error GoTo Err_Handler
    RenderScale = 1.5
    Call RepositionAndRender
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R4in140:" & Error$
    End Select
End Sub

Private Sub ScaleButton20_Click()
On Error GoTo Err_Handler
    RenderScale = 2
    Call RepositionAndRender
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R4in141:" & Error$
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
        Case Else: MsgBox Err & "-R4in142:" & Error$
    End Select
End Function
 


Private Sub CmboPrinterList_Click()
On Error GoTo Err_Handler
    Dim i As Integer

    Call SelectPrinter(CmboPrinterList.text)
    If CmboPrinterList.Enabled = True Then
        For i = 1 To Me.FloatingText.Count - 1 'We dont want the text to be rendered on the preview yet
            Me.FloatingText(i).Visible = False 'when changes printers
        Next i
        
        Call Me.RestoreOriginalState
        Call Me.IntitialiseForm
        
        For i = 1 To Me.FloatingText.Count - 1 'Even thou we dont want the text to be rendered, we still
            Me.FloatingText(i).Visible = True  'want to be able to see them after printer select changed
        Next i
    End If
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R4in143:" & Error$
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
        Case Else: MsgBox Err & "-R4in144:" & Error$
    End Select


End Sub

Sub RepositionAndRender()
On Error GoTo Err_Handler

    Call RenderPage(picReportPagePg1)
    Call RenderPage(picReportPagePg2)
    


    
    picReportPagePg1.Visible = True
    If NoOfGraphPanels > 3 Then
        picReportPagePg2.Visible = True
    Else
        picReportPagePg2.Visible = False
    End If
    
'    CommentsTextBox(0).Visible = True
'    UserTitle.Visible = True

    Me.UserTitle.Visible = True
    Me.CommentsTextBox(0).Visible = True
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R4in145:" & Error$
    End Select
End Sub

Sub AutoSetNoOfGraphs()
On Error GoTo Err_Handler
Dim TotalDistance As Double
Dim NoOfGraphsAsDbl As Double
Dim NoOfGraphs As Integer

'PCN????
'TotalDistance = Abs(ConfigInfo.DistanceStart - ConfigInfo.DistanceFinish) 'sometimes old PVD dont have finnish distance
TotalDistance = Abs(PVDistances(1) - PVDistances(PVDataNoOfLines - 1)) '-1 just in case its last is not good
NoOfGraphsAsDbl = Round(TotalDistance / 30, 0)
NoOfGraphs = CInt(NoOfGraphsAsDbl)

If NoOfGraphs <= 0 Then NoOfGraphs = 1
If NoOfGraphs > 7 Then NoOfGraphs = 7 'PCN4278 thanks geoff, we only have 7 graphs, not 8 :):)
NoOfGraphPanels = NoOfGraphs
NoGraphsDisplay.Caption = NoOfGraphPanels

Call SetNoOfPages


RenderScale = 1
Call RepositionAndRender
    
Exit Sub
Err_Handler:
    MsgBox Err & "-R4in146:" & Error$

End Sub

Sub SetNoOfPages()
On Error GoTo Err_Handler
Dim NoOfPages As Integer
Dim CurrentPage As Integer

CurrentPage = PageHScroll.value

If NoOfGraphPanels > 3 Then
    NoOfPagesLabel.Caption = DisplayMessage("Page") & " " & CurrentPage & " " & DisplayMessage(" of ") & "2"
    NoOfPagesLabel.Visible = True
    PageHScroll.Visible = True
Else
'    NoOfPagesLabel.Caption = DisplayMessage("Page") & " " & CurrentPage & " " & DisplayMessage(" of ") & "1"
    NoOfPagesLabel.Visible = False
    PageHScroll.Visible = False
    PageHScroll.value = 1
End If


Exit Sub
Err_Handler:
    MsgBox Err & "-R4in147:" & Error$
End Sub

Sub ScaleGraphs(ByVal YSc As Single, ByVal XSc As Single)
On Error GoTo Err_Handler

Dim i As Integer
Dim J As Integer
Dim YOffset As Single
Dim YShift As Single
Dim XOffset As Single
Dim XShift As Single


For i = 0 To 6
    YOffset = Me.GraphBackgroundShape(i).Top
    YShift = (YOffset - Me.GraphBackgroundShape(0).Top) * (1 - YSc)
    
    Me.GraphBackgroundShape(i).Top = Me.GraphBackgroundShape(i).Top - YShift
    Me.GraphBackgroundShape(i).height = Me.GraphBackgroundShape(i).height * YSc
    
    Me.GraphXScaleContainerArray(i).Top = ((Me.GraphXScaleContainerArray(i).Top - YOffset) * YSc) + YOffset - YShift
    Me.GraphXScaleContainerArray(i).height = Me.GraphXScaleContainerArray(i).height * YSc
    
    Me.GraphContainerArray(i).Top = ((Me.GraphContainerArray(i).Top - YOffset) * YSc) + YOffset - YShift
    Me.GraphContainerArray(i).height = Me.GraphContainerArray(i).height * YSc
    
    Me.Shape2(i).height = Me.Shape2(i).height * YSc
    Me.Shape2(i).Top = ((Me.Shape2(i).Top - YOffset) * YSc) + YOffset - YShift
    
    For J = 0 To 4
        Me.YScaleLabel((i * 5) + J).Top = ((Me.YScaleLabel((i * 5) + J).Top - YOffset) * YSc) + YOffset - YShift
    Next J
    
    Me.UnitLabel(i).Top = ((Me.UnitLabel(i).Top - YOffset) * YSc) + YOffset - YShift
    Me.GraphUnitLabel(i).Top = ((Me.GraphUnitLabel(i).Top - YOffset) * YSc) + YOffset - YShift
Next i

For i = 0 To 6
    XOffset = Me.GraphBackgroundShape(i).Left
    XShift = 0
    
    Me.GraphBackgroundShape(i).Left = Me.GraphBackgroundShape(i).Left - XShift
    Me.GraphBackgroundShape(i).width = Me.GraphBackgroundShape(i).width * XSc
    
    Me.GraphXScaleContainerArray(i).Left = ((Me.GraphXScaleContainerArray(i).Left - XOffset) * XSc) + XOffset - XShift
    Me.GraphXScaleContainerArray(i).width = Me.GraphXScaleContainerArray(i).width * XSc
    
    Me.GraphContainerArray(i).Left = ((Me.GraphContainerArray(i).Left - XOffset) * XSc) + XOffset - XShift
    Me.GraphContainerArray(i).width = Me.GraphContainerArray(i).width * XSc
    
    Me.Shape2(i).width = Me.Shape2(i).width * XSc
    Me.Shape2(i).Left = ((Me.Shape2(i).Left - XOffset) * XSc) + XOffset - XShift
    
    For J = 0 To 4
        Me.YScaleLabel((i * 5) + J).Left = ((Me.YScaleLabel((i * 5) + J).Left - XOffset) * XSc) + XOffset - XShift
    Next J
    
    Me.UnitLabel(i).Left = ((Me.UnitLabel(i).Left - XOffset) * XSc) + XOffset - XShift
    Me.GraphUnitLabel(i).Left = ((Me.GraphUnitLabel(i).Left - XOffset) * XSc) + XOffset - XShift
    Me.LimitValue2(i).Left = ((Me.LimitValue1(i).Left - XOffset) * XSc) + XOffset - XShift
    Me.LimitValue1(i).Left = Me.LimitValue2(i).Left
    
Next i

Exit Sub
Err_Handler:
   MsgBox Err & "-R4in148:" & Error$
End Sub

Sub SetControlsPerPaperSize()
On Error GoTo Err_Handler
Dim i As Integer
Dim ShuntAmount As Single

    picReportPagePg1.width = Printer.width * RenderScale
    picReportPagePg1.height = Printer.height * RenderScale
    picReportPagePg1.Left = 100
    picReportPagePg1.Top = 100

    picReportPagePg2.width = picReportPagePg1.width
    picReportPagePg2.height = picReportPagePg1.height
    picReportPagePg2.Left = picReportPagePg1.width + 200
    picReportPagePg2.Top = 100

'ShuntAmount = 735
ShuntAmount = 235

If Printer.PaperSize = 1 Then
    ShuntY(0) = -500
    ShuntY(1) = -500
    ShuntX(0) = 0
    ShuntX(1) = 0
    
    
    
Else
    ShuntY(0) = ShuntAmount
    ShuntY(1) = ShuntAmount
    ShuntX(0) = 0
    ShuntX(1) = 0
    
End If

Call ShuntControlsPerContainer

If Printer.PaperSize = 1 Then
    'Call Me.ScaleGraphs(1, 1.05)    'letter
    Call Me.ScaleGraphs(0.95, 1#)
    Call Me.MoveLabels(0)
Else
    'Call Me.ScaleGraphs(1.05, 1#)   'A4
    Call Me.ScaleGraphs(1#, 0.96)
    Call Me.MoveLabels(-450)
End If




Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R4in149:" & Error$
    End Select

End Sub

Sub MoveLabels(ByVal ShiftAmount As Single)
On Error GoTo Err_Handler

Dim i As Integer


i = 0
 
    DateLabel(i).Left = DateLabel(i).Left + ShiftAmount
    DateLabelData(i).Left = DateLabelData(i).Left + ShiftAmount
    
    PipeMaterialLabel(i).Left = PipeMaterialLabel(i).Left + ShiftAmount
    PipeMaterialData(i).Left = PipeMaterialData(i).Left + ShiftAmount
    
    PipeLengthLabel(i).Left = PipeLengthLabel(i).Left + ShiftAmount
    PipeLengthUnitLabel(i).Left = PipeLengthUnitLabel(i).Left + ShiftAmount
    PipeLengthData(i).Left = PipeLengthData(i).Left + ShiftAmount
    
    PipeDiameterLabel(i).Left = PipeDiameterLabel(i).Left + ShiftAmount
    PipeDiameterUnitLabel(i).Left = PipeDiameterUnitLabel(i).Left + ShiftAmount
    PipeDiameterData(i).Left = PipeDiameterData(i).Left + ShiftAmount
    
    If ScreenDrawing.ImageGraphState(0).GraphType <> "Flat" Then
        CommentsTextBox(i).width = CommentsTextBox(i).width + ShiftAmount
    Else
        CommentsTextBox(i).width = CommentsTextBox(i).width + (ShiftAmount / 2)
    End If
    
For i = 0 To 1
    HeaderBreakLine(i).x2 = HeaderBreakLine(i).x2 + ShiftAmount
    CopyrightLabel(i).Left = CopyrightLabel(i).Left + ShiftAmount
    PageLabel(i).Left = PageLabel(i).Left + ShiftAmount
    PrintedLabel(i).Left = PrintedLabel(i).Left + ShiftAmount
    PrintedData(i).Left = PrintedData(i).Left + ShiftAmount
    FooterBreakLine(i).x2 = FooterBreakLine(i).x2 + ShiftAmount
Next i

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R4in150:" & Error$
    End Select
End Sub

Sub ShuntControlsPerContainer()
On Error GoTo Err_Handler

Dim i As Integer
Dim S As Integer
Dim ControlType As String

'Draw renderings first that are marked back

For S = 0 To 1 'Loop for the different Shunts
    For i = 1 To Me.Controls.Count - 1
        
        With Me.Controls(i)
            If TypeName(.Container) = "PictureBox" Then
                If .Tag <> "Shunt" And .Tag <> "Egnore" And ControlsPositionShunt(S).Tag = "Shunt" Then
                    If TypeName(Me.Controls(i)) = "TextBox" Or _
                       TypeName(Me.Controls(i)) = "Label" Or _
                       TypeName(Me.Controls(i)) = "Shape" Or _
                       TypeName(Me.Controls(i)) = "Image" Or _
                       TypeName(Me.Controls(i)) = "PictureBox" Then
                        If .Left > ControlsPositionShunt(S).Left And _
                           .Left < ControlsPositionShunt(S).Left + ControlsPositionShunt(S).width And _
                           .Top > ControlsPositionShunt(S).Top And _
                           .Top < ControlsPositionShunt(S).Top + ControlsPositionShunt(S).height Then
                            .Top = .Top + ShuntY(S)
                            .Left = .Left + ShuntX(S)
                        End If
                    End If
                    If TypeName(Me.Controls(i)) = "Line" Then
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
        Case Else: MsgBox Err & "-R4in151:" & Error$
    End Select
    
End Sub



'PCN4433
Private Sub UserTitle_Change()
On Error GoTo Err_Handler

    Dim UserTittleCaption As String
    Dim TittleParse As String
        
    TittleParse = "Analysis_" & ScreenDrawing.ImageGraphState(0).GraphType & "="
    Call INI_WriteBack(MyFile, TittleParse, Me.UserTitle.text)

'Dim FileSaveFail As Boolean

'UserTitleAnalysis = Trim(Me.UserTitle.text) 'PCN4433

Exit Sub
Err_Handler:
    Select Case Err

        Case Else: MsgBox Err & "-R4in152:" & Error$
    End Select
End Sub

