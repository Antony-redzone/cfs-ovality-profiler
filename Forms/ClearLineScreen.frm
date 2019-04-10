VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ClearLineScreen 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Profiler Screen"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15120
   ControlBox      =   0   'False
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
   Icon            =   "ClearLineScreen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   617
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1008
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox DistanceTracker 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   30
      TabIndex        =   18
      Text            =   "0"
      Top             =   11940
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer ProfilerTimer 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   11400
      Top             =   7980
   End
   Begin VB.PictureBox MainScreen 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   480
      MousePointer    =   2  'Cross
      ScaleHeight     =   509
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   680
      TabIndex        =   4
      Top             =   0
      Width           =   10260
      Begin VB.PictureBox DimenResults 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   120
         ScaleHeight     =   270
         ScaleWidth      =   720
         TabIndex        =   13
         Top             =   6750
         Width           =   750
         Begin VB.Label DimenMeasure 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   240
            TabIndex        =   14
            Top             =   0
            Width           =   3000
         End
         Begin VB.Image DimenImage 
            Appearance      =   0  'Flat
            Height          =   240
            Left            =   0
            Picture         =   "ClearLineScreen.frx":014A
            Top             =   0
            Width           =   240
         End
      End
      Begin VB.TextBox OnScreenText 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         HideSelection   =   0   'False
         Index           =   0
         Left            =   0
         MouseIcon       =   "ClearLineScreen.frx":04D4
         MousePointer    =   3  'I-Beam
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox OnScreenTextBoxPicture 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   24
         Top             =   360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox OnscreenTextDefault 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFF7EF&
         Height          =   375
         Left            =   0
         TabIndex        =   25
         Text            =   "Default Text"
         Top             =   720
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.PictureBox SnapShotScreen 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2640
         Left            =   0
         MousePointer    =   2  'Cross
         ScaleHeight     =   176
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   153
         TabIndex        =   5
         Top             =   0
         Width           =   2295
         Begin VB.Shape VideoImageDistIgnoreBoarder 
            BorderColor     =   &H00C0C000&
            BorderStyle     =   4  'Dash-Dot
            Height          =   15
            Left            =   0
            Top             =   0
            Width           =   15
         End
         Begin VB.Shape VideoImageIgnoreBoarder 
            BorderColor     =   &H00C0C000&
            BorderStyle     =   4  'Dash-Dot
            Height          =   15
            Left            =   0
            Top             =   0
            Width           =   15
         End
      End
      Begin VB.PictureBox PVScreen 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00F0F0FF&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2880
         Left            =   0
         ScaleHeight     =   192
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   246
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   3690
         Begin VB.Label PVFrameStatus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   420
            TabIndex        =   9
            Top             =   7350
            Width           =   1185
         End
         Begin VB.Image PVFrameImage 
            Height          =   240
            Left            =   120
            Picture         =   "ClearLineScreen.frx":0D9E
            Top             =   7350
            Width           =   240
         End
      End
      Begin VB.PictureBox MagnifyPictContainer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2280
         Left            =   1800
         ScaleHeight     =   152
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   203
         TabIndex        =   26
         Top             =   3840
         Width           =   3039
         Begin VB.PictureBox MagnifyValuePicBox 
            Appearance      =   0  'Flat
            BackColor       =   &H00B36A36&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   2040
            Left            =   2700
            ScaleHeight     =   134
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   21
            TabIndex        =   29
            Top             =   240
            Width           =   345
            Begin VB.PictureBox MagnifySliderOld 
               Appearance      =   0  'Flat
               BackColor       =   &H00400000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   135
               Left            =   0
               ScaleHeight     =   7
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   23
               TabIndex        =   30
               Top             =   -1500
               Width           =   375
            End
            Begin VB.Image MagnifySlider 
               Height          =   240
               Left            =   60
               MousePointer    =   1  'Arrow
               Picture         =   "ClearLineScreen.frx":1128
               Top             =   0
               Width           =   240
            End
            Begin VB.Image MagScaleBackgroundSelected 
               Height          =   2010
               Left            =   0
               MousePointer    =   1  'Arrow
               Picture         =   "ClearLineScreen.frx":14B2
               Top             =   0
               Visible         =   0   'False
               Width           =   315
            End
            Begin VB.Image SliderScale 
               Height          =   2010
               Left            =   0
               MousePointer    =   1  'Arrow
               Picture         =   "ClearLineScreen.frx":3674
               Top             =   0
               Width           =   315
            End
         End
         Begin VB.PictureBox MagnifyMoveBar 
            Appearance      =   0  'Flat
            BackColor       =   &H00B36A36&
            FillStyle       =   0  'Solid
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            MouseIcon       =   "ClearLineScreen.frx":5836
            MousePointer    =   99  'Custom
            ScaleHeight     =   15
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   201
            TabIndex        =   28
            Top             =   0
            Width           =   3045
            Begin VB.Label MagnifyLabel 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "x88"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   -15
               TabIndex        =   31
               Top             =   -15
               Width           =   375
            End
            Begin VB.Image CloseImage 
               Height          =   240
               Left            =   2700
               MousePointer    =   1  'Arrow
               Picture         =   "ClearLineScreen.frx":6100
               Stretch         =   -1  'True
               Top             =   0
               Width           =   345
            End
         End
         Begin VB.PictureBox MagnifyImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C0FF&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   2040
            Left            =   0
            ScaleHeight     =   134
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   179
            TabIndex        =   27
            Top             =   240
            Width           =   2720
         End
      End
      Begin VB.PictureBox AreaResults 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   870
         TabIndex        =   15
         Top             =   7080
         Visible         =   0   'False
         Width           =   900
         Begin VB.Label AreaMeasure 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   10
            Width           =   570
         End
         Begin VB.Image AreaImage 
            Appearance      =   0  'Flat
            Height          =   240
            Left            =   0
            Picture         =   "ClearLineScreen.frx":6442
            Top             =   0
            Width           =   240
         End
      End
      Begin VB.PictureBox VideoScreen 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   7695
         Left            =   0
         ScaleHeight     =   513
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   449
         TabIndex        =   17
         Top             =   0
         Width           =   6735
         Begin VB.PictureBox EmbededSnapshot 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000001&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1095
            Left            =   4560
            ScaleHeight     =   73
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   81
            TabIndex        =   32
            Top             =   2520
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.PictureBox MainScreen3D 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FF80FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   6840
         ScaleHeight     =   247
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   207
         TabIndex        =   12
         Top             =   3120
         Visible         =   0   'False
         Width           =   3135
         Begin VB.Timer ThreeDtimer 
            Enabled         =   0   'False
            Interval        =   33
            Left            =   360
            Top             =   1320
         End
      End
      Begin VB.PictureBox PVScreenPicInPic 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2250
         Left            =   7380
         ScaleHeight     =   146
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   181
         TabIndex        =   10
         Top             =   60
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.PictureBox DummyScreen 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   6045
         Left            =   0
         ScaleHeight     =   403
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   585
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   8775
      End
      Begin VB.Label PVDimensionLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "47"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   2520
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.TextBox ScaleSetting 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   315
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "X"
      ToolTipText     =   "Scale settings"
      Top             =   7785
      Width           =   425
   End
   Begin VB.PictureBox PVXScale 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   680
      TabIndex        =   2
      Top             =   7785
      Width           =   10260
      Begin VB.Line XMarker1st 
         BorderColor     =   &H00FF8080&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   17
      End
      Begin VB.Line XMarker2nd 
         BorderColor     =   &H0000C000&
         X1              =   5
         X2              =   5
         Y1              =   0
         Y2              =   17
      End
   End
   Begin VB.PictureBox PVXScaleWindow 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7800
      Left            =   0
      ScaleHeight     =   7740
      ScaleWidth      =   360
      TabIndex        =   0
      Top             =   0
      Width           =   425
      Begin VB.PictureBox PVYScale 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   8295
         Left            =   0
         ScaleHeight     =   553
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   1
         Top             =   310
         Width           =   425
         Begin VB.Image SecretButtonHere 
            Height          =   255
            Left            =   0
            Top             =   7200
            Width           =   375
         End
         Begin VB.Line YMarker2nd 
            BorderColor     =   &H0000C000&
            X1              =   24
            X2              =   0
            Y1              =   5
            Y2              =   5
         End
         Begin VB.Line YMarker1st 
            BorderColor     =   &H00FF8080&
            X1              =   24
            X2              =   0
            Y1              =   0
            Y2              =   0
         End
      End
      Begin VB.Label Y_Units 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   15
         TabIndex        =   33
         Top             =   15
         Width           =   325
      End
      Begin VB.Shape Y_UnitsBackground 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   300
         Left            =   0
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox VideoFrameSliderBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   30
      ScaleHeight     =   465
      ScaleWidth      =   9390
      TabIndex        =   19
      Top             =   8115
      Width           =   9390
      Begin VB.PictureBox VideoRecordMarkerStartAdjuster 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   105
         Left            =   360
         MousePointer    =   99  'Custom
         Picture         =   "ClearLineScreen.frx":69CC
         ScaleHeight     =   105
         ScaleWidth      =   180
         TabIndex        =   22
         Top             =   0
         Width           =   180
      End
      Begin VB.PictureBox VideoRecordMarkerStopAdjuster 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   105
         Left            =   7470
         MousePointer    =   99  'Custom
         Picture         =   "ClearLineScreen.frx":6B0A
         ScaleHeight     =   105
         ScaleWidth      =   180
         TabIndex        =   21
         Top             =   0
         Width           =   180
      End
      Begin VB.PictureBox VideoSliderMarker 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2970
         Picture         =   "ClearLineScreen.frx":6C48
         ScaleHeight     =   180
         ScaleWidth      =   90
         TabIndex        =   20
         Top             =   85
         Width           =   90
      End
      Begin VB.Shape FrameBlank 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   240
         Top             =   135
         Width           =   8880
      End
      Begin VB.Line VideoRecordMarkerStart 
         BorderColor     =   &H0000FF00&
         X1              =   525
         X2              =   525
         Y1              =   375
         Y2              =   0
      End
      Begin VB.Line VideoRecordMarkerStop 
         BorderColor     =   &H000000FF&
         X1              =   7470
         X2              =   7470
         Y1              =   375
         Y2              =   0
      End
      Begin VB.Image Image5 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   0
         Picture         =   "ClearLineScreen.frx":6D7A
         Top             =   90
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog OnScreenTextDialog 
      Left            =   11520
      Top             =   10080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Shape PanelEdgeLine 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   3495
      Left            =   10800
      Top             =   0
      Width           =   75
   End
   Begin VB.Label AVITimeLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   9420
      TabIndex        =   11
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Shape ViewIndicator 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   1515
      Index           =   3
      Left            =   0
      Top             =   0
      Width           =   75
   End
   Begin VB.Menu BrushSizeArray 
      Caption         =   "Brush Size"
      Visible         =   0   'False
      Begin VB.Menu BrushSize1 
         Caption         =   "&1"
         Index           =   1
      End
      Begin VB.Menu BrushSize2 
         Caption         =   "&2"
         Index           =   2
      End
      Begin VB.Menu BrushSize3 
         Caption         =   "&3"
         Index           =   3
      End
      Begin VB.Menu BrushSize4 
         Caption         =   "&4"
         Index           =   4
      End
      Begin VB.Menu BrushSize5 
         Caption         =   "&5"
         Index           =   5
      End
      Begin VB.Menu BrushSize6 
         Caption         =   "&6"
         Index           =   6
      End
      Begin VB.Menu BrushSize7 
         Caption         =   "&7"
         Index           =   7
      End
      Begin VB.Menu BrushSize8 
         Caption         =   "&8"
         Index           =   8
      End
      Begin VB.Menu BrushSize9 
         Caption         =   "&9"
         Index           =   9
      End
      Begin VB.Menu BrushSize10 
         Caption         =   "&10"
         Index           =   10
      End
   End
   Begin VB.Menu EraseSizeArray 
      Caption         =   "EraseSize"
      Visible         =   0   'False
      Begin VB.Menu EraseSize1 
         Caption         =   "&1"
         Index           =   1
      End
   End
   Begin VB.Menu DotSizeArray 
      Caption         =   "Dot Size"
      Visible         =   0   'False
      Begin VB.Menu DotSize1 
         Caption         =   "&1"
         Index           =   1
      End
   End
   Begin VB.Menu BrushColor 
      Caption         =   "Brush Color"
      Visible         =   0   'False
   End
   Begin VB.Menu Tools 
      Caption         =   "Tools"
      Visible         =   0   'False
      Begin VB.Menu Modify2 
         Caption         =   "M&odify"
      End
      Begin VB.Menu Delete2 
         Caption         =   "Dele&te"
      End
   End
   Begin VB.Menu Option1 
      Caption         =   "Options"
      Visible         =   0   'False
      Begin VB.Menu autosnap 
         Caption         =   "Auto &Snap"
         Checked         =   -1  'True
      End
      Begin VB.Menu autotriangle 
         Caption         =   "Auto &Triangle "
         Checked         =   -1  'True
         Enabled         =   0   'False
      End
      Begin VB.Menu autotermination 
         Caption         =   "Auto &Termination"
         Checked         =   -1  'True
      End
      Begin VB.Menu attachdrawing 
         Caption         =   "Attach &Drawing"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu LineKind 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu TerPL 
         Caption         =   "&Terminate Line"
         Index           =   1
      End
   End
   Begin VB.Menu CircleKind 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu NormalCircle 
         Caption         =   "&Circle"
      End
      Begin VB.Menu EggShapes 
         Caption         =   "E&gg Shape"
      End
      Begin VB.Menu EllipseCircle 
         Caption         =   "&Ellipse"
      End
   End
   Begin VB.Menu CalcMenu 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu Length1 
         Caption         =   "&Length"
         Index           =   1
      End
   End
   Begin VB.Menu FillingMethod 
      Caption         =   "Filling Method"
      Visible         =   0   'False
      Begin VB.Menu SolidFilling 
         Caption         =   "Solid"
      End
      Begin VB.Menu PatternFilling 
         Caption         =   "Tiling Pattern"
      End
   End
   Begin VB.Menu OnScreenTextMenu 
      Caption         =   "OnScreenText"
      Visible         =   0   'False
      Begin VB.Menu OnscreenTextFontMenu 
         Caption         =   "Font"
      End
      Begin VB.Menu OnScreenTextBackgroundColourMenu 
         Caption         =   "Background Colour"
      End
      Begin VB.Menu OnscreenTextDefaultMenu 
         Caption         =   "Rest to default"
      End
      Begin VB.Menu Blank 
         Caption         =   ""
      End
      Begin VB.Menu OnscreenTextDeleteMenu 
         Caption         =   "Delete"
      End
      Begin VB.Menu OnscreenTextDeleteAllMenu 
         Caption         =   "Delete All"
      End
   End
End
Attribute VB_Name = "ClearLineScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Testing
Option Explicit

' PCN3640 move text box etc from clearline screen form to proper vairables
'Public CalLen As Double             '
Public Area_Pixel As Double         '
Public Area_Real As Double          '
Public CurrentFile1 As String       '
Public dispx1 As Double             '
Public dispy1 As Double             '
Public dispx2 As Double             '
Public dispy2 As Double             '
Public dispx3 As Double             '
Public dispy3 As Double             '
Public Distance1 As Double          '
Public Length2 As Double            '
Public Area2 As Double              '
Public StartAngle As Double         '
Public EndAngle As Double           '
Public WLStartAngle As Double       '
Public AngPercent As Double         '
Public AreaRealPercent As Double    '
Public Len_Pixel As Double          '
Public Circum_Pixel As Double       '
Public Circum_Real As Double        '
Public CircumRealPercent As Double  '
Public Rds2 As Double               '
Public ApproxArea As Boolean        '
Public ShapeNo As Integer           '
Public CoorX As Double              '
Public CoorY As Double              '
'Public CalLength_tmp As Double      '
Public CurrentShp As Integer        '
Public Chk4AutoTriangle As Boolean  '
Public Chk4AutoTermination As Boolean
Public Chk4SnapEnabled As Boolean   '

Dim MTMode As String 'Measurement Tool mode, eg Line, Multiline, Rotate, etc
'Dim DrawShapeType As String ' Draw Shape Type, Circle, Egg or oval 'PCN3055 Defined in Startup
Dim OutPipeRadius As Long 'Expected Internal
Dim IntPipeRadius As Long 'Outside Radius of pipeline
Dim ScreenCentreX As Integer 'Centre coordinates of the Main Screen
Dim ScreenCentreY As Integer 'Centre coordinates of the Main Screen
Dim VideoPipeCentreX As Integer
Dim VideoPipeCentreY As Integer 'Centre offset to Main Screen center for the pipe centre
Dim LineX1 As Long
Dim LineY1 As Long
Dim LineX2 As Long
Dim LineY2 As Long
Dim Scrolling As Boolean 'PCNLS220103
'Dim BegValue As Integer  'PCNLS220103
Dim BegValue As Long  'PCNLS220103 'PCNGL170303-3
Public AVITime As Long 'String * 5 'PCNLS220103
Dim AVITimeVar As Long 'String * 5 'PCNLS220103
Public VideoFrameNo As Long 'Keeps track of the current video frame number (used in the C code) 'PCNGL220103
Dim SliderFrame As Long  'PCNLS240103
Dim EndOfAVI As Boolean 'PCNLS290103
Dim PicInPicMouseDownX As Integer  'PCN1863
Dim PicInPicMouseDownY As Integer  'PCN1863
Dim countL As Integer 'PCNLS190303


Dim Radius(0 To 179) As Long
Dim Colour(0 To 179) As Long
Dim CurrentAVITime As Double 'PCNGL150103
'vvvv PCN2240 **************************************
Public ButtonNumber As Integer
'^^^^ **********************************************

'PCN3569.........................................
Private Const WM_NCLBUTTONDOWN As Long = &HA1&
Private Const HTCAPTION As Long = 2&
Private Declare Function ReleaseCapture Lib "user32" () As Long
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, wParam As Any, lParam As Any) As Long
'.................................................

Public CurrentMouseMoveX As Single
Public CurrentMouseMoveY As Single
Public LastMouseMoveX As Single
Public LastMouseMoveY As Single
Public LeftMouseDown As Boolean
Public RightMouseDown As Boolean

'PCN3621 Magnifing glass, (25 August 2005, Antony) '
Dim MagnifyAction As String
Dim MagnifyMouseDownX As Single
Dim MagnifyMouseDownY As Single

Dim MagnifyScale As Single
Dim MagnifyXCentre As Single
Dim MagnifyYCentre As Single
''''''''''''''''''''''''''''''''''''''''''''''''''''

Public SonarMeasure As String   'PCN4419
Public SonarX1 As Long         'PCN4419
Public SonarY1 As Long         'PCN4419
Public SonarX2 As Long         'PCN4419
Public SonarY2 As Long         'PCN4419

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Filename:  ClearlineScreen.frm , ClearlineScreen.frx
'    Date:  01.02.2001
' Updated:  26.06.2001
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'RGB(RedValue, GreenValue, BlueValue), 0 <= RedValue or GreenValue or BlueValue <= 255, RedValue and GreenValue and BlueValue are integers.
'           RedValue     GreenValue   BlueValue
'Black      0            0            0
'Blue       0            0            255
'Green      0            255          0
'Cyan       0            255          255
'Red        255          0            0
'Magenta    255          0            255
'Yellow     255          255          0
'White      255          255          255

'QBColor(ColorValue) 0 <= ColorValue <= 15, ColorValue is integer.
'ColorValue     Color       ColorValue      Color
'   0           Black           8           Gray
'   1           Blue            9           Light Blue
'   2           Green           10          Light Green
'   3           Cyan            11          Light Cyan
'   4           Red             12          Light Red
'   5           Magenta         13          Light Magenta
'   6           Yellow          14          Light Yellow
'   7           White           15          Bright White
'
' GiveColor(c) method uses c values as below.
'   c           GiveColor
'   0           vbBlack
'   1           vbBlue
'   2           vbGreen
'   3           vbCyan
'   4           vbRed
'   5           vbMagenta
'   6           vbYellow
'   7           vbWhite
'
'NormalDrawingColor : Blue
'SelectedObjectColor: Cyan
'SnapCircle, Movement, Track : Yellow
'ModifyCircle : Magenta
'ChosenModifyCircle : Red
'AutoTriangleLines & Angle : White
'FillColor : Green
'Color considered when fill : Blue
'Color considered when selection : Blue, Cyan
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Find `RGB(` - 200 times used. => O.K.

'vvvv PCN1931 *************************************
'These variables are duplicated in the Startup function
'Public NormalDrawingColor As Long
'Public ExtraObjectColor As Long 'This color is only for objects to be hidden when area is being filled(e.g. lines & arc for Angle)
'Public TempDrawingColor As Long
'Public SelectedObjectColor As Long
'Public RotatingObjectColor As Long
'Public ModifyingObjectColor As Long
'Public MovingObjectColor As Long
'Public ChosenModiCircleColor As Long
'Public ModiCircleColor As Long
'Public JointCircleColor As Long
'Public AreaFillingColor As Long
'Public SelectionBoundaryColor As Long
'Public EDBoundaryColor As Long
'Public DotDrawingColor As Long
'Public IPRepaintColor As Long ' Repainting Color after Image Processing
'Public IPPaintColor As Long ' Painting Color after Image Processing
'Public BrushColor1 As Long
'Public BrushColor2 As Long 'When Var. BrushC = 0, BrushColor1 is used and 1, BrushColor2
'Public TextSizeIndicatorColor As Long 'Loaded from the INI
'^^^^ ***************************************************
'Find vbBlack - 6 times used. => O.K.
'Find vbBlue - 7 times used. => O.K.
'Find vbGreen - 6 times used. => O.K.
'Find vbCyan - 4 times used. => O.K.
'Find vbRed - 3 times used. => O.K.
'Find vbMagenta - 3 times used. => O.K.
'Find vbYellow - 3 times used. => O.K.
'Find vbWhite - 6 times used. => O.K.
'Find ForeColor => O.K.
'Find BackColor => O.K.
'Find Color assignment using integer (e.g:7(white)) - ? times used.
'Replace them with different variable names.
'Allow user to change these variables during runtime.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public ObsFlag As Boolean
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Const SnapShotScreen_Width As Long = 1024 '349
Const SnapShotScreen_Height As Long = 768 '285
Const TotLabel As Long = 1 '120000 'for memory on 011001
Const TotOneObj As Integer = 100
Const OneObjs As Integer = 1500
Const TotJointObj As Integer = 100



Dim CaptImg As String
Public CurrentFile As String
Public CurrentPath As String
Public ChangeFlag As Boolean
Dim InitialLoadFlag As Boolean
Dim PolyLineFlag As Boolean
Dim PolyLineStarted As Boolean
Dim OrthoLineFlag As Boolean
Dim OrthoStart As Boolean
Dim CalFlag As Boolean
Dim FirstCircleFlag As Boolean ' If close object is the first circle among circles.
Dim FirstCircleObjectNo As Integer
Dim StartedOnCircle As Boolean

Dim NoCircleFlag As Boolean
Dim TotLength As Double ' length of polylines
Dim TotArea As Double ' Integrated Aea

'text
Public TextStartX As Integer
Dim TextStartY As Integer

'erase
Dim EraseStartFlag As Boolean

Dim MultiDrawingFlag As Boolean

'UNDO, REDO
Const MaxUndo As Integer = 1
Const MaxRedo As Integer = 1
Dim UndoAction(MaxUndo) As Integer ' Index
Dim RedoAction(MaxRedo) As Integer ' Index
Dim UndoObj(MaxUndo, 4, 3) As Integer ' Index, Point1~4, XYcoordinate
Dim RedoObj(MaxRedo, 4, 3) As Integer ' Index, Point1~4, XYcoordinate
Dim DelObjStack4Undo(MaxUndo, 700, 4, 3)  'Index(UndoStack), CurrentShape, Point1~4, XYcoordinate
Dim DelObjStack4Redo(MaxRedo, 700, 4, 3)  'Index(RedoStack), CurrentShape, Point1~4, XYcoordinate
Dim TotUndo As Integer
Dim TotRedo As Integer

Dim PolyStart(2) As Double
Dim PolyStartLineNo As Integer

'Finding Objects for modification, and selection and deselection (before deletion/movement)
'Dim SelDesDisabled As Boolean
'Dim FoundDot(NoDot) As Integer
'Dim FoundLine(NoLine) As Integer
'Dim FoundCircle(NoCircle) As Integer
'Dim FoundArc(NoArc) As Integer
'Dim FoundCurve(NoCurve) As Integer
'Dim FoundText(NoText) As Integer
'Dim TotFoundDot As Integer
'Dim TotFoundLine As Integer
'Dim TotFoundCircle As Integer
'Dim TotFoundArc As Integer
'Dim TotFoundCurve As Integer
'Dim TotFoundText As Integer

'Edge Detection
Dim picw As Integer
Dim pich As Integer
Dim OrigImg(SnapShotScreen_Width, SnapShotScreen_Height) As Long ' Store Data of Original Image
Dim CurrentImg(SnapShotScreen_Width, SnapShotScreen_Height) As Long ' Store Data of Current Image
Public OrigHeight As Integer
Dim OrigWidth As Integer
Dim ZoomedImg(SnapShotScreen_Width, SnapShotScreen_Height) As Long ' Store Data of Current Image
Dim Seg(SnapShotScreen_Width, SnapShotScreen_Height) As Long 'for segmentation
Dim seg2(SnapShotScreen_Width, SnapShotScreen_Height) As Long 'for filling area
Dim Edges4Fill(SnapShotScreen_Width, SnapShotScreen_Height) As Long 'for filling area
Dim Edges4R(SnapShotScreen_Width, SnapShotScreen_Height) As Long
Dim Edges4G(SnapShotScreen_Width, SnapShotScreen_Height) As Long
Dim Edges4B(SnapShotScreen_Width, SnapShotScreen_Height) As Long
Dim BW4R(SnapShotScreen_Width, SnapShotScreen_Height) As Long
Dim BW4G(SnapShotScreen_Width, SnapShotScreen_Height) As Long
Dim BW4B(SnapShotScreen_Width, SnapShotScreen_Height) As Long
Dim BW4RGB(SnapShotScreen_Width, SnapShotScreen_Height) As Long
Dim LabelCounter As Long
Dim MaxSegArea As Long
Dim Labels(TotLabel) As Long ' For tidying up
Dim EDStart(2) As Integer
Dim EDEnd(2) As Integer
Dim pvX As Integer
Dim pvY As Integer
Dim EDStartFlag As Boolean
Dim BeforePressed As Boolean

Dim lex As Integer
Dim rex As Integer
Dim uex As Integer
Dim dex As Integer
Dim ley As Integer
Dim rey As Integer
Dim uey As Integer
Dim dey As Integer

Dim PaintNow As Boolean
Dim DotFlag As Boolean
Dim BrushNow As Boolean
Dim TotPoint As Integer
Dim SaveX As Double
Dim SaveY As Double
Dim SaveX2 As Double
Dim SaveY2 As Double
Dim StoreX As Double
Dim StoreY As Double
 
'Dim X4Line(NoLine), Y4Line(NoLine), X4Line2(NoLine), Y4Line2(NoLine) As Double
'Dim X4Circle(NoCircle), Y4Circle(NoCircle), X4Circle2(NoCircle), Y4Circle2(NoCircle)  As Double
'Dim EggShape(NoEggShape, 1, 1) As Double
'Dim TotEggShape As Integer
'Dim TextAttr(NoText, 3) As Long ' x, y, size, colorvalue
'Dim Texts(NoText, 1) As String 'content, fontname
'Dim TextFontType(NoText, 1) As Boolean 'bold, italic
Dim Texts_Undo(700, 1) As String 'content, fontname
Dim TextFontType_Undo(700, 1) As Boolean 'bold, italic
Dim Texts_Redo(700, 1) As String 'content, fontname
Dim TextFontType_Redo(700, 1) As Boolean 'bold, italic

'Dim TotText As Integer
'Dim EggComponents(NoEggShape, 3) As Double
'Dim X4Arc(NoArc), Y4Arc(NoArc), X4Arc2(NoArc), Y4Arc2(NoArc), X4Arc3(NoArc), Y4Arc3(NoArc), TotArc As Double
'Dim StartAng4Arc(NoArc), EndAng4Arc(NoArc) As Double

Dim directionFlag As Boolean
Dim asin As Double
Dim asin2 As Double
Dim EposX1 As Double
Dim EPosX2 As Double
Dim EPosX3 As Double
Dim EPosX4 As Double
Dim EposY1 As Double
Dim EposY2 As Double
Dim EposY3 As Double
Dim EPosY4 As Double
Dim MposX As Double
Dim MposY As Double

Dim SelStartX As Double
Dim SelStartY As Double
Dim SelEndX As Double
Dim SelEndY As Double
Dim SelStarted As Boolean

'Global Var. for ScanFill
Dim left_endx As Integer
Dim left_endy As Integer
Dim right_endx As Integer
Dim right_endy As Integer
Dim up_endx As Integer
Dim up_endy As Integer
Dim down_endx As Integer
Dim down_endy As Integer

'global var. for move
Dim mx1 As Double
Dim mx2 As Double
Dim mx3 As Double
Dim my1 As Double
Dim my2 As Double
Dim my3 As Double
Dim MoveFlag As Boolean
Dim StartMoveX As Double
Dim StartMoveY As Double
Dim FirstJointSuccess As Boolean
Dim SecondJointSuccess As Boolean
Dim SuccessX As Double
Dim SuccessY As Double
Dim MovingTrack As Integer


Dim JointObjectInfo(TotJointObj) As String  ' 10 objects can be joined in one point 'may not necessary if so,deletelater
    ' (shape(Dot, Line..), index of the shape)
Dim JointObjectCoordinate(TotJointObj, 2) As Double
Dim JointExist As Boolean
Dim CurrentJoinObj As String
Dim CurrentJoinObjNo As Integer
'Dim TotObjNo As Integer

'Capability to consider several objects one object.
'Dim TotOneObjNo As Integer
'Dim OneObjNo(TotOneObj) As Integer
'Dim OneObjInfo(TotOneObj, OneObjs, 2) As Integer '(Obj no., Objs, shape and shapeNo.)
  ' Undo for One Object Group.
Dim TotOneObjNo_Undo As Integer
Dim OneObjNo_Undo(TotOneObj) As Integer
Dim OneObjInfo_Undo(TotOneObj, OneObjs, 2) As Integer
Dim OneObjIndex_Undo(TotOneObj) As Integer
  ' Redo for One Object Group.
Dim TotOneObjNo_Redo As Integer
Dim OneObjNo_Redo(TotOneObj) As Integer
Dim OneObjInfo_Redo(TotOneObj, OneObjs, 2) As Integer
Dim OneObjIndex_Redo(TotOneObj) As Integer

Dim ApproxAreaFlag As Boolean
Dim ApproxAreaStartedFlag As Boolean
Dim ApproxStartX As Double
Dim ApproxStartY As Double

Dim CheckForDLLError As Boolean 'PCN1970
Private Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'Private Declare Sub startthreed Lib "viewpipe.dll" (ByVal hinst As Long, ByVal hwnd1 As Long)
'Private Declare Sub closethreed Lib "viewpipe.dll" ()

'C++ Functions Declared in Laserlib.dll used for playing media files
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                                            'ByVal hinst As Long, ByVal hwind As Long, ByVal FName As String, width As Long, height As Long
Private Declare Sub getversion Lib "laserlib.dll" (ByRef version As Double) 'PCN1970
Private Declare Sub initialise Lib "laserlib.dll" (ByVal hinst As Long, ByVal hwnd1 As Long, ByVal mediatype As String, ByVal name As String, ByRef width As Long, ByRef height As Long, ByRef xadju As Double, ByRef yadju As Double, ByRef realheight As Long, ByRef realwidth As Long, ByVal CaptureDevice As Long, ByVal originalwidth As Long, ByVal originalheight As Long, ByVal NoVideoSync As Long) 'PCN2289 'PCN2395 capture device from string to long
Private Declare Sub uninitialise Lib "laserlib.dll" ()
Private Declare Sub uninitialiseLive Lib "laserlib.dll" ()
Private Declare Function getprofiledata Lib "laserlib.dll" (ByVal Frame As Long, Radius As Long, Colour As Long, Time As Double, X As Long, Y As Long, r As Long, loosing As Long) As Long
Private Declare Sub videopause Lib "laserlib.dll" ()
Private Declare Sub videorun Lib "laserlib.dll" ()
Private Declare Sub videostep Lib "laserlib.dll" () 'PCN2668 (4 March 2004, Antony van Iersel)
Private Declare Sub videoframeadvance Lib "laserlib.dll" () 'PCN2865 Antony, frame advnace added to laserlib, more reliable than using step that is used for recording PVD
Private Declare Sub videoframerewind Lib "laserlib.dll" () 'PCN2865 Antony, frame advnace added to laserlib, more reliable than using step that is used for recording PVD
Private Declare Sub grabsnapshot Lib "laserlib.dll" (ByVal SnapName As String, ByVal Registered As Long, ByVal watermark As String, ByVal FishEyeON As Long) ' PCN4596 now fisheye off is only on a snap
Private Declare Sub timeseek Lib "laserlib.dll" (ByVal Time As Double)
Private Declare Sub writetofileC Lib "laserlib.dll" (ByVal Frame As Long, ByVal Time As Double)
Private Declare Sub getTime Lib "laserlib.dll" (ByRef VideoLength As Double)
Private Declare Sub resizewindow Lib "laserlib.dll" ()

Private Declare Sub setrecprofstat Lib "laserlib.dll" (ByVal i As Long) 'pass it one to set to true, 0 for false
Private Declare Sub refreshframe Lib "laserlib.dll" ()
Private Declare Sub resetcounters Lib "laserlib.dll" ()
Private Declare Sub getcenter Lib "laserlib.dll" (X As Single, Y As Single)
Private Declare Sub getcurrenttime Lib "laserlib.dll" (ByRef CurrentTime As Double)
Private Declare Sub setwindow Lib "laserlib.dll" (ByVal i As Long)
Private Declare Sub Optimize Lib "laserlib.dll" Alias "optimize" (ByVal i As Long)
Private Declare Sub emptybuffer Lib "laserlib.dll" ()
Private Declare Sub setdistancerectangle Lib "laserlib.dll" (ByVal XBottom As Long, ByVal YBottom As Long, ByVal XTop As Long, ByVal YTop As Long, ByVal Units As Long) 'PCN2639 ' PCN2874
Private Declare Sub cleardistancerectangle Lib "laserlib.dll" () 'PCN2639
Private Declare Sub showrectangle Lib "laserlib.dll" ()
Private Declare Sub setwaterlevelbool Lib "laserlib.dll" (ByVal i As Long, ByVal J As Long)
Private Declare Sub getgroupedprofiledataxy Lib "laserlib.dll" (ByRef xcoordinate As Single, ByRef ycoordinate As Single, _
                                                                ByRef XCentre As Single, ByRef YCentre As Single, _
                                                                ByRef Time As Double, _
                                                                ByRef NumFrames As Long, _
                                                                ByRef Distance As Long) 'PCN2891 xcoordinate and xcoordinate are 10x pixel co-ordinates of the video Media (3:4 ratio)
                                                                'PCN3219 added centres
                                                                
Private Declare Sub gettotalnumframes Lib "laserlib.dll" (ByRef numberOfFrames As Long) 'Double 'returns the total number of frames /1000
Private Declare Sub refscreen Lib "laserlib.dll" ()
Private Declare Sub setdistancecounter Lib "laserlib.dll" (ByVal Distance As Long, ByVal Direction As Long) 'PCN2639
Private Declare Sub hough_getprocesstime Lib "laserlib.dll" (ByRef time_in_millisec As Double)  'PCN2735

'vvvv PCN2240 *****************************************************************
'3D module library

'PCN3194
Private Declare Sub Hough_InitialiseSingleImage Lib "laserlib.dll" (ByRef ImagePointer As Byte, _
                                                                 ByVal imageWidth As Long, _
                                                                 ByVal imageHeight As Long) 'PCN3194

'PCNant????
Private Declare Sub hough_debugcoordxy1 Lib "laserlib.dll" (ByVal X As Long, ByVal Y As Long)
Private Declare Sub hough_debugcoordxy2 Lib "laserlib.dll" (ByVal X As Long, ByVal Y As Long)
Private Declare Sub hough_debugcoordxy3 Lib "laserlib.dll" (ByVal X As Long, ByVal Y As Long)
Private Declare Sub hough_debugcoordxy4 Lib "laserlib.dll" (ByVal X As Long, ByVal Y As Long)
Private Declare Sub hough_debugcoordxy5 Lib "laserlib.dll" (ByVal X As Long, ByVal Y As Long)

Private Declare Sub hough_IsVideoRunning Lib "laserlib.dll" (ByRef IsVideoRunning As Long)
Private Declare Sub hough_setlaserpoint Lib "laserlib.dll" (ByVal XCoord As Long, ByVal YCoord As Long)



Private Declare Sub d3d_getversion Lib "threedim.dll" (ByRef d3d_version As Double) 'PCN1970
Private Declare Sub clearline_getversion Lib "Clearline.dll" (ByRef dll_version As Double)
Private Declare Sub sonar_getversion Lib "SonarProfiler.dll" (ByRef dll_version As Double)


'PCN2240 ' PCN2693 Needed for colour calculations (Antony van Iersel, 15 March 2004)
'PCN2860 'PCN2988 'PCN2988 'PCN3111 0 = metric, 1 = imperial
Private Declare Sub d3d_initialise Lib "threedim.dll" (ByVal hInstance As Long, _
                                                       ByVal hwnd1 As Long, _
                                                       ByVal vertex_mode As String, _
                                                       ByRef DataX As Single, _
                                                       ByRef DataY As Single, _
                                                       ByRef pvCentreX As Single, _
                                                       ByRef pvCentreY As Single, _
                                                       ByVal dataSize As Long, _
                                                       ByVal Path As String, _
                                                       ByRef pvColourRed As Long, _
                                                       ByRef pvColourGreen As Long, _
                                                       ByRef pvColourBlue As Long, _
                                                       ByVal ExpRad As Double, _
                                                       ByVal PVCalculationsMultiplier As Long, _
                                                       ByVal PVDataXYMultiplier As Long, _
                                                       ByVal Units As Long) 'PCN3540 Ovlaity change from long to single
                                                       
Private Declare Sub d3d_destroy Lib "threedim.dll" ()
Private Declare Sub d3d_refresh Lib "threedim.dll" ()
Private Declare Sub d3d_left_button_down Lib "threedim.dll" (ByVal X As Long, ByVal Y As Long)
Private Declare Sub d3d_right_button_down Lib "threedim.dll" (ByVal X As Long, ByVal Y As Long)
Private Declare Sub d3d_mousemove_and_down Lib "threedim.dll" (ByVal X As Long, ByVal Y As Long)
Private Declare Sub d3d_keydown Lib "threedim.dll" (ByVal Key As Long)
Private Declare Sub d3d_laser_focus Lib "threedim.dll" (ByVal focus As Long)
Private Declare Sub d3d_camselect Lib "threedim.dll" (ByVal cameranumber As Long)
Private Declare Sub d3d_scene_on_off Lib "threedim.dll" (ByVal scenenumber As Long)
Private Declare Sub d3d_play_speed Lib "threedim.dll" (ByVal playspeed As Long)
Private Declare Sub d3d_pipe_scale Lib "threedim.dll" (ByVal X As Long)
Private Declare Sub d3d_reset Lib "threedim.dll" ()
Private Declare Sub d3d_zoom_speed Lib "threedim.dll" (ByVal zoomspeed As Long)
Private Declare Sub d3d_rotate_pipe_z Lib "threedim.dll" (ByVal deg As Long)
Private Declare Sub d3d_rotate_pipe_y Lib "threedim.dll" (ByVal deg As Long)
Private Declare Sub d3d_next_pipe_texture Lib "threedim.dll" ()
Private Declare Sub d3d_capture_window Lib "threedim.dll" (ByVal SnapName As String, ByVal hwnd1 As Long) 'PCN2240
Private Declare Sub d3d_pipe_colour_limits Lib "threedim.dll" () 'PCN2337
Private Declare Sub d3d_export_stl Lib "threedim.dll" (ByVal ExportFileName As String)   ' PCN2376
Private Declare Sub d3d_setlanguage Lib "threedim.dll" (ByVal line As Long, ByVal text As String) ' PCN2473 (Antony van Iersel, 10 March 2004)
Private Declare Sub d3d_directxversion Lib "threedim.dll" (ByRef Ver As Long) 'PCN3141
Private Declare Sub d3d_pvgraphtype Lib "threedim.dll" (ByVal GraphType As String, ByVal minLimit As Double, ByVal maxLimit As Double) 'PCN2337 Antony 3 November 2003, Set the 3D Pipe Shading Limits and Type, PCN2693 pass as double not long

'^^^^ *************************************************************************



'Stuff to change the priority and therefore speed up the VB drawing PCNLS170403
Public Enum PRIORITY_CLASS
   NORMAL_PRIORITY_CLASS = &H20
   IDLE_PRIORITY_CLASS = &H40
   HIGH_PRIORITY_CLASS = &H80
   REALTIME_PRIORITY_CLASS = &H100
End Enum

Private Const PROCESS_DUP_HANDLE = &H40

Private Declare Function OpenProcess Lib "kernel32" _
   (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long
    
Private Declare Function CloseHandle Lib "kernel32" _
   (ByVal hObject As Long) As Long
   
Private Declare Function GetCurrentProcessId _
   Lib "kernel32" () As Long
   
Private Declare Function SetPriorityClass Lib "kernel32" _
  (ByVal hProcess As Long, ByVal dwPriorityClass As Long) _
   As Long

Dim DimX As Double
Dim DimY As Double 'PCN376
Dim MidLineFlag As Boolean
'PCN267
Dim CalLine(1, 1) As Integer

Dim ResizeFlag As Boolean

'vvvv PCN2930 *********************************************************
Dim VideoMarkerLastX As Long 'Last mouse X position - Start marker
Dim VideoMarkerMoveDir As String 'Indication of mouse move direction
Dim VideoFrameSlider As SliderParameters
'^^^^ *****************************************************************

'vvvv PCN4368 ************************************
Public Event MouseLeave()
'Public Event MouseHover()
Dim WithEvents MouseTrackMagnifyMoveBar As clsTrackInfo
Attribute MouseTrackMagnifyMoveBar.VB_VarHelpID = -1
'^^^^ ********************************************

Public Function VideoScreenActivate(ErrorStr As String) 'PCN3005
'****************************************************************************************
'Name    : VideoScreenActivate
'Created : Sep 9 2004
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    : contains the code for the video button click event
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler

    If mediatype = "" Then 'PCN2783
        ErrorStr = "Must load a Media file first!"
        'MsgBox DisplayMessage(ErrorStr)
        ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage(ErrorStr): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
        Exit Function
    End If
 
    
    'vvvv PCN2240 ****************************

    
    ThreeDtimer.Enabled = False
    '^^^^ ************************************
    Call ProfilerPause
    PVScreen.Visible = False
    
    If mediatype = Video Then
        'PCN3194 this was moved into the Video only condition was just above '(19 August 2005, Antony)
        '
        'Enable AVI Play buttons 'PCNGL1812022
''        ControlToolbar.Buttons.Item(9).Visible = True 'was 23 AVIPlayPause 'PCN2681
''        ControlToolbar.Buttons.Item(10).Visible = True 'was 24 AVIStepForward 'PCN2681
''        ControlToolbar.Buttons.Item(11).Visible = True 'was 25 AVIStepRewind 'PCN2681
''        ControlToolbar.Buttons.Item(6).Visible = False 'was 20 PVPlayPause 'PCN2681
''        ControlToolbar.Buttons.Item(7).Visible = False 'was 21 PVStepForward 'PCN2681
''        ControlToolbar.Buttons.Item(8).Visible = False 'was 22 PVStepRewind 'PCN2681
        '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        'VideoTimer.Enabled = True 'PCN3194
        If VideoFileName = "" And mediatype = Video Then 'PCNGL140103
            'MsgBox DisplayMessage("Must load a Media file first!"), vbExclamation 'PCN2111
            ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Must load a Media file first!"): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
        ElseIf CheckAVIInitialised Then

            ClearLineTitle.TitleBarCaption.Caption = DisplayMessage(Video) & " - " & VideoFileName 'PCN2759
            ClearLineTitle.TitleBarCaption.ToolTipText = DisplayMessage(Video) & " - " & VideoFileName 'PCN4171
            If CLPScreenMode = PV Then
                CLPScreenMode = Video
                Call GotoPVProfile(PVFrameNo)
            End If
            CLPScreenMode = Video
            
            PVScreen.MousePointer = 99 'PCNGL210103
            PVScreen.MouseIcon = LoadResPicture(MainScreenMouseIcon, vbResIcon) 'PCNGL210103
            VideoScreen.ZOrder 0 'PCNGL190103
     '       Call SetupMTButtonsForVideo 'PCNGL300103
            Call RefreshVideoScreen
    '                Else 'PCN1863

            Call SetupMainScreenForVideo
        End If
    ElseIf mediatype = "Live" Then
        Call SetupMainScreenForVideo  'PCNLS290103
        Call ControlsScreen.SetupForCalibration 'PCNLIVE
    End If
    ' Turn off Picture in Picture
    PVScreenPicInPic.Visible = False
    PicInPicMode = "OFF"
   

Exit Function
Err_Handler:
Select Case Err
    Case 75
        PVScreen.MousePointer = 0
        Resume Next
    Case Else
        MsgBox Err & "-CLS1:" & Error$
End Select
End Function

Private Sub BtnLoadDebug_Click()
    DebugForm.ZOrder 0
End Sub

Sub ProfilerStartRecord()
On Error GoTo Err_Handler
Dim ErrorStr As String

    If PVRecording Then Exit Sub
    If PVDFileName <> "" Then Exit Sub
    'If mediatype <> Video Then MsgBox DisplayMessage("There must be a media file loaded before recording can take place."), vbInformation 'PCN2111
    If mediatype <> Video Then ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("There must be a media file loaded before recording can take place."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Call RecordPVProfile_Start(ErrorStr)

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS2:" & Error$
End Sub

Sub ProfilerStopRecord()
On Error GoTo Err_Handler

Call RecordPVProfile_Stop

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS3:" & Error$
End Sub

Sub ProfilerPausePlay()
On Error GoTo Err_Handler

If ProfilerTimer.Enabled Then
    If CLPScreenAction = "DrawMaskBox" Or _
       CLPScreenAction = "DrawWaterLevel" Or _
       CLPScreenAction = "DrawHorCalibrationLine" Or _
       CLPScreenAction = "DrawCalibrationLine" Then
       Call ClearLineScreen.TakeASnapShot
    End If
    Call ProfilerPause
Else
    Call ProfilerRun
End If


    
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS4:" & Error$
End Sub

Sub ProfilerPause()
On Error GoTo Err_Handler

    Call videopause 'Make sure the video stops running, regardless if the video is initiated or not, CPP will handle that
    ProfilerTimer.Enabled = False 'Stop all playing in the profiler

    Call ControlsScreen.ControlsViewSetup
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS5:" & Error$
End Sub

Sub ProfilerRun()
On Error GoTo Err_Handler
    If mediatype <> Video And PVDFileName = "" Then Exit Sub
    ProfilerTimer.Enabled = True
    If CLPScreenMode = SnapShot Or CLPScreenMode = Video Then VideoScreen.ZOrder 0: CLPScreenMode = Video
    If PVDFileName <> "" Then Call ScreenDrawing.DrawProfilesStartToFinish(PVScreen, False) 'Clear Permanent profile
    Call ControlsScreen.ControlsViewSetup 'PCN4171

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS6:" & Error$
End Sub




Private Sub CloseImage_Click()
On Error GoTo Err_Handler
    
Me.MagnifyPictContainer.Visible = False
Call ControlsScreen.SetupControlZoomOnSnap

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS7:" & Error$
End Sub

Private Sub Form_Activate()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Form_Activate Sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    26/11/02     Building initial framework
'
'Description:
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

'SonarConfig.Show
'Y_Units.Caption = MeasurementUnits 'PCNGL150103
''Call DrawPVProfile_Setup(ClearLineScreen.PVScreen, StopTimer) 'PCNGL310103
'If CentreLineX = 0 Then 'PCNGL310103
'    CentreLineX = Round(ClearLineScreen.MainScreen.width / 2)
'End If
'If CentreLineY = 0 Then
'    CentreLineY = Round(ClearLineScreen.MainScreen.height / 2)
'End If
'PVDrawScreenRatio = ConfigInfo.Ratio 'ClearLineScreen.Ratio 'PCN3035
'DrawingCentreX = CentreLineX
'DrawingCentreY = CentreLineY
'Call DrawMainScale(ClearLineScreen.PVScreen)   'PCNGL300103

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS8:" & Error$
End Sub

Private Sub Form_Load()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : Form_Load
'Created : 8 November 2002,
'Updated : 18 November 2003, PCN2402
'Prg By  : Geoff Logan
'Param   :
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
ConvertLanguage Me, Language 'PCN2111
''PopupCalibrationTypes.width = PopupCalibrationTypes.ButtonWidth 'PCN2138
''PopupDrawShapes.width = PopupDrawShapes.ButtonWidth 'PCN2138

Dim ScreenRatio As Double  'PCNGL240103

'Position form 'PCNGL060103
ClearLineScreen.width = 10900
ClearLineScreen.Left = 0
ClearLineScreen.Top = 360
ClearLineScreen.height = 8450

Me.BackColor = RGB(172, 196, 231) 'PCN4171
Me.VideoFrameSliderBar.BackColor = RGB(172, 196, 231) 'PCN4171
Me.PanelEdgeLine.Top = 0
Me.PanelEdgeLine.Left = MainScreen.Left + MainScreen.width + 10  'PCN4171
Me.PanelEdgeLine.height = ClearLineScreen.height

'Resize screen for 800 * 600 operating system PCN1876


'vvvv ******* Setup MainScreen for default video size (768x576 or aspect ratio of 0.75) 'PCNGL2401032
VideoAspectRatio = VideoAspectRatio768x576 'As default
MainScreen.height = MainScreen.width * VideoAspectRatio
'Setup Video screen
Call VideoScreenSetupAspectRatio
'^^^^ *************************************************************************

'VideoSnapShotMode = SnapShot
'CLPScreenMode = Video 'PCNGL291202
CLPScreenMode = SnapShot 'PCN4043

MainScreenMouseIcon = 116 'Default mouse
''ControlToolbar.Buttons.Item(10).Enabled = True 'PCN2372 'PCN2681
''ControlToolbar.Buttons.Item(11).Enabled = True 'PCN2372 'PCN2681
'If hough_anycapturedevices And VideoCaptureDevice > 0 Then
' ControlToolbar.Buttons.Item(1).Enabled = True
'Else
' ControlToolbar.Buttons.Item(1).Enabled = False
'End If
            
ReDim RefCirclePriority(3) 'PCN1835

'vvvv Setup Picture In Picture screen ********************************** 'PCNGL240103
ScreenRatio = ClearLineScreen.PVScreenPicInPic.width / ClearLineScreen.MainScreen.width
ClearLineScreen.PVScreenPicInPic.height = ClearLineScreen.MainScreen.height * ScreenRatio + 10
Dim curSelect As StdPicture 'PCN1863
Set curSelect = LoadResPicture(108, vbResIcon) 'Move icon 'PCN1863
PVScreenPicInPic.MousePointer = 99 'PCN1863
PVScreenPicInPic.MouseIcon = curSelect 'Move icon 'PCN1863
'^^^^ *************************************************************************

'vvvv PCN1970 ******************************
' Check the version of the LaserLib.dll
'Whenever changing the version of a DLL, ensure the number complies with the following guidelines.
'The DLL may be changed more or less often than the VB version.
'Do we want to update a user with a new version of the VB software every time we change the DLL version? Probably not.
'So the VB software will except DLL version with the same major version number. That is if the VB DLL version is 1.0, the VB will accept the DLL version 1.0 to 1.9. The VB will not DLL versions <1.0 or >1.9.
'E.g.: ClearLine Profiler's LaserLib.dll version = 1.0. Then ClearLine will accept LaserLib.dll version from 1.0 to 1.9
'Therefore, for a VB software with a DLL version number of 1.0, ALL DLLs with versions 1.0 to 1.9 MUST work on this VB software.
'If the change in the DLL means it will not work on ALL VB software of the same major DLL version, then the DLL's version MUST increase the major DLL version.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Dim DLLVersion As Double

CheckForDLLError = True
Call getversion(DLLVersion)
CheckForDLLError = False
Dim strTemp As String 'Temporary message box string
If DLLVersion <> LaserLibVersion Then
    strTemp = DisplayMessage("LaserLib.dll VERSION ERROR. Expecting ")
    strTemp = strTemp & LaserLibVersion & DisplayMessage(", laserlib.dll is currently ")
    strTemp = strTemp & Format(DLLVersion, "###0.0") & DisplayMessage(" - This application may not work as designed.") 'PCN2111
    MsgBox strTemp, vbCritical 'PCN2111
End If
'^^^^ **************************************

'vvvv PCN1970 ******************************
' Check the version of the ThreeDim.dll
CheckForDLLError = True
DLLVersion = 0#
Call d3d_getversion(DLLVersion)
CheckForDLLError = False
If DLLVersion <> ThreeDimVersion Then
    strTemp = DisplayMessage("ThreeDim.dll VERSION ERROR. Expecting ")
    strTemp = strTemp & ThreeDimVersion & DisplayMessage(", ThreeDim.dll is currently ")
    strTemp = strTemp & Format(DLLVersion, "###0.0") & DisplayMessage(" - This application may not work as designed.") 'PCN2111
    MsgBox strTemp, vbCritical 'PCN2111
End If
'^^^^ **************************************

CheckForDLLError = True
Call clearline_getversion(DLLVersion)
If DLLVersion <> ClearLineVersion Then
    strTemp = DisplayMessage("ClearLine.dll VERSION ERROR. Expecting ")
    strTemp = strTemp & ClearLineVersion & DisplayMessage(", ClearLine.dll is currently ")
    strTemp = strTemp & Format(DLLVersion, "###0.0") & DisplayMessage(" - This application may not work as designed.") 'PCN2111
    MsgBox strTemp, vbCritical 'PCN2111
End If

'Remove sonar alltogether
'If SoftwareConfiguration <> "Reader" Then
'    CheckForDLLError = True
'    Call sonar_getversion(DLLVersion)
'    If DLLVersion <> SonarVersion Then
'        strTemp = DisplayMessage("SonarProfiler.dll VERSION ERROR. Expecting ")
'        strTemp = strTemp & SonarVersion & DisplayMessage(", SonarProfiler.dll is currently ")
'        strTemp = strTemp & Format(DLLVersion, "###0.0") & DisplayMessage(" - This application may not work as designed.") 'PCN2111
'        MsgBox strTemp, vbCritical 'PCN2111
'    End If
'End If

Dim DirectXVersion As Long
'vvvv PCN3141 ******************************
' Check the version of currently installed direct x
CheckForDLLError = True
DLLVersion = 0#
Call d3d_directxversion(DirectXVersion)
CheckForDLLError = False
If DirectXVersion < DirectXNeeded Then
    strTemp = DisplayMessage("Direct X VERSION ERROR. Expecting Microsoft Direct X ")
    strTemp = strTemp & DirectXNeeded & DisplayMessage(", or above. Microsoft Direct X is currently ")
    strTemp = strTemp & DirectXVersion & DisplayMessage(" - This application may not work as designed.") 'PCN2111
    MsgBox strTemp, vbCritical 'PCN2111
End If
'^^^^ **************************************

'Initialize three D  PCNLS120603
'Call preinitializethreed(App.hInstance)

'vvvv PCN2240 ***************************************
' 3D setup
D3D_Control = "Z-axis"
D3D_MapOn = True
D3D_WaterOnOff = False
D3D_Scale = 2
D3D_CameraView = "External"
'^^^^ ***********************************************

'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
Call VideoFrameSliderSetup 'PCN2930
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Call OnscreenTextResetAllToDefault 'PCN3569

    MagnifyAction = "None"

    'PCN4353 ''''''''''''''''''''''''''''''''''''''''
    MagnifyScale = 2                                '
    MagnifySlider.Top = 25                          '
    Me.MagnifySlider.height = 16                    '
    Me.MagScaleBackgroundSelected.height = 38       '
    Me.MagScaleBackgroundSelected.Visible = True    '
    '''''''''''''''''''''''''''''''''''''''''''''''''
    
PVScreen.width = MainScreen.width
PVScreen.height = MainScreen.height

Me.EmbededSnapshot.width = MainScreen.width
Me.EmbededSnapshot.height = MainScreen.height

SnapShotScreen.width = MainScreen.width
SnapShotScreen.height = MainScreen.height

VideoScreen.width = MainScreen.width
VideoScreen.height = MainScreen.height

MainScreen3D.width = MainScreen.width
MainScreen3D.height = MainScreen.height

CentreLineX = MainScreen.width / 2
CentreLineY = MainScreen.height / 2

'vvvv PCN4205 ********************************
Y_Units.Caption = MeasurementUnits
'^^^^ ****************************************

'vvvv PCN4328 ************************************
'Initilise the mouse leave event on the key's drag bar.
Set MouseTrackMagnifyMoveBar = New clsTrackInfo
MouseTrackMagnifyMoveBar.hwnd = MagnifyMoveBar.hwnd

StartTrack MouseTrackMagnifyMoveBar
'^^^^ ********************************************


Exit Sub
Err_Handler:
    'vvvv PCN1970 ******************************************
    If CheckForDLLError = True Then
        If Err = 453 Then 'Function is not in DLL
            DLLVersion = 0#
            Resume Next
        Else
            'MsgBox DisplayMessage("DLL ERROR - ") & Error$, vbCritical 'PCN2111
            ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("DLL ERROR - ") & Error$: ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
        End If
    Else
        MsgBox Err & "-CLS9:" & Error$
    End If
    '^^^^ **************************************************
End Sub





Private Sub MagnifyImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call SnapShotScreen_MouseDown(Button, Shift, MagnifyToSnapshotX(X), MagnifyToSnapshotY(Y))
End Sub

Private Sub MagnifyImage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SnapShotScreen_MouseMove(Button, Shift, MagnifyToSnapshotX(X), MagnifyToSnapshotY(Y))
End Sub

Private Sub MagnifyImage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SnapShotScreen_MouseUp(Button, Shift, MagnifyToSnapshotX(X), MagnifyToSnapshotY(Y))
End Sub

Private Sub MouseTrackMagnifyMoveBar_MouseLeave() 'PCN4368
On Error GoTo Err_Handler

Me.MagnifyMoveBar.BackColor = &HB36A36

RaiseEvent MouseLeave
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS10:" & Error$
End Sub

Private Sub MagnifyMoveBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
Dim curSelect As StdPicture

Set curSelect = LoadResPicture(109, vbResIcon)
MagnifyMoveBar.MouseIcon = curSelect
    
MagnifyMouseDownX = X
MagnifyMouseDownY = Y
MagnifyAction = "Move"



Exit Sub
Err_Handler:
    Select Case Err
    Case Else
        MsgBox Err & "-CLS11:" & Error$
    End Select
    
End Sub

Private Sub MagnifyMoveBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    Dim MoveX As Single
    Dim MoveY As Single
    
Me.MagnifyMoveBar.BackColor = &HC0C000 'PCN4368
    
    If MagnifyAction = "Move" Then
        With Me.MagnifyPictContainer

            MoveX = .Left - MagnifyMouseDownX + X
            MoveY = .Top - MagnifyMouseDownY + Y
            
            If MoveX < 0 Then MoveX = 0
            If MoveY < 0 Then MoveY = 0
            If (MoveX + .width) > Me.SnapShotScreen.width Then MoveX = Me.SnapShotScreen.width - .width
            If (MoveY + .height) > Me.SnapShotScreen.height Then MoveY = Me.SnapShotScreen.height - .height
            .Left = MoveX
            .Top = MoveY
            
        End With
        Call MagnifyUpdateImage
        Call SnapShotScreen_MouseMove(Button, Shift, MagnifyToSnapshotX(MagnifyImage.width / 2), MagnifyToSnapshotY(MagnifyImage.height / 2))
        DoEvents
    End If
Exit Sub
Err_Handler:
    Select Case Err
    Case Else
        MsgBox Err & "-CLS12:" & Error$
    End Select
End Sub





Private Sub MagnifyMoveBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
Dim curSelect As StdPicture

Set curSelect = LoadResPicture(108, vbResIcon)
MagnifyMoveBar.MouseIcon = curSelect

MagnifyAction = "None"

Exit Sub
Err_Handler:
    Select Case Err
    Case Else
        MsgBox Err & "-CLS13:" & Error$
    End Select
End Sub

Private Sub MagnifyPictContainer_GotFocus()
On Error GoTo Err_Handler
    Call MagnifyUpdateImage
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS14:" & Error$
End Sub

Private Sub MagnifySlider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    MagnifyAction = "Move"
    MagnifyMouseDownX = X
    MagnifyMouseDownY = Y
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS15:" & Error$
End Sub

Private Sub MagnifySlider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    Dim MoveY As Single
    Dim IconOffset As Integer
    
    If MagnifyAction = "Move" Then
        IconOffset = 13
        MoveY = Me.MagnifySlider.Top + ((MagnifyMouseDownY + Y) / 15)
        
        If MoveY < IconOffset Then MoveY = IconOffset
        If MoveY + Me.MagnifySlider.height > Me.SliderScale.height Then
            MoveY = Me.SliderScale.height - Me.MagnifySlider.height
        End If
        Me.MagnifySlider.Top = MoveY - IconOffset 'Account for Icon offset
        'Setup background
        Me.MagScaleBackgroundSelected.height = MoveY
        Me.MagScaleBackgroundSelected.Visible = True
        
        'Translate slider to magnification
        MagnifyScale = (MagnifySlider.Top / SliderScale.height * 6) + 1
        Me.MagnifyUpdateImage
    End If

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS16:" & Error$
End Sub

Private Sub MagnifySlider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    
    MagnifyAction = "None"

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS17:" & Error$
End Sub


Private Sub MainScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : MainScreen_MouseMove
'Created :
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    
Call DrawMainScaleRedLines(X, Y)

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS18:" & Error$
End Sub


Private Sub MainScreen3D_KeyDown(KeyCode As Integer, Shift As Integer)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : MainScreen3D_KeyDown
'Created : 7/10/2003, PCN2240
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

If ThreeDRunning = True Then
    Call d3d_keydown(KeyCode)
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS19:" & Error$
End Sub

Private Sub MainScreen3D_KeyPress(KeyAscii As Integer)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : MainScreen3D_KeyPress
'Created : 7/10/2003, PCN2240
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

If ThreeDRunning = True Then
    Call d3d_keydown(KeyAscii)
End If

Exit Sub
Err_Handler:
     MsgBox Err & "-CLS20:" & Error$
End Sub

Private Sub MainScreen3D_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : MainScreen3D_MouseDown
'Created : 7/10/2003, PCN2240
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim curSelect As StdPicture 'PCN2240
Dim x1 As Long
Dim y1 As Long

x1 = X
y1 = Y
If (Button = 1) Then 'left
    If ThreeDRunning = True Then
        Set curSelect = LoadResPicture(109, vbResIcon) 'PCN2240 Hand icon
        MainScreen3D.MousePointer = 99 'PCN2240
        MainScreen3D.MouseIcon = curSelect 'PCN2240
        Call d3d_left_button_down(x1, y1)
        ButtonNumber = 1 'PCN2240
    Else
        TrackingPosition.MouseX = X
        TrackingPosition.MouseY = Y
        TrackingPosition.MouseButton1Down = True
    End If
ElseIf (Button = 2) Then 'right
    If ThreeDRunning = True Then
        Set curSelect = LoadResPicture(119, vbResIcon) 'PCN2240 ZoomIn icon
        MainScreen3D.MousePointer = 99 'PCN2240
        MainScreen3D.MouseIcon = curSelect 'PCN2240
        Call d3d_right_button_down(x1, y1)
        ButtonNumber = 2 'PCN2240
    Else
        TrackingPosition.MouseX = X
        TrackingPosition.MouseY = Y
        TrackingPosition.MouseButton2Down = True
    End If
End If
LeftMouseDown = True

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS21:" & Error$
End Sub

Private Sub MainScreen3D_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : MainScreen3D_MouseMove
'Created : 7/10/2003, PCN2240
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim curSelect As StdPicture 'PCN2240
Dim x1, y1 As Long
    
x1 = X
y1 = Y
If ThreeDRunning = True And LeftMouseDown = True Then
    Call d3d_mousemove_and_down(x1, y1)
    'Check the direction of movement
    If Y > D3D_MouseY2 And ButtonNumber = 2 Then
        Set curSelect = LoadResPicture(120, vbResIcon) 'PCN2240 Hand icon
        MainScreen3D.MousePointer = 99 'PCN2240
        MainScreen3D.MouseIcon = curSelect 'PCN2240
    ElseIf Y < D3D_MouseY2 And ButtonNumber = 2 Then
        Set curSelect = LoadResPicture(119, vbResIcon) 'PCN2240 Hand icon
        MainScreen3D.MousePointer = 99 'PCN2240
        MainScreen3D.MouseIcon = curSelect 'PCN2240
    ElseIf ButtonNumber = 2 Then
        Set curSelect = LoadResPicture(121, vbResIcon) 'PCN2240 Hand icon
        MainScreen3D.MousePointer = 99 'PCN2240
        MainScreen3D.MouseIcon = curSelect 'PCN2240
    End If

End If
D3D_MouseY2 = Y

If Not ThreeDRunning Then
    Call TrackingPosition.MouseMove(X, Y)
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS22:" & Error$
End Sub

Private Sub MainScreen3D_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : MainScreen3D_MouseUp
'Created : 7/10/2003, PCN2240
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim curSelect As StdPicture 'PCN2240

LeftMouseDown = False
Set curSelect = LoadResPicture(108, vbResIcon) 'PCN2240 Hand icon
MainScreen3D.MousePointer = 99 'PCN2240
MainScreen3D.MouseIcon = curSelect 'PCN2240
ButtonNumber = 0 'PCN2240

TrackingPosition.MouseButton2Down = False
TrackingPosition.MouseButton1Down = False

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS23:" & Error$
End Sub




Private Sub OnScreenText_Change(Index As Integer)
On Error GoTo Err_Handler
    Call SetTextBoxWidthAndHeight(OnScreenText(Index))
    OnScreenText(Index).Refresh
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS24:" & Error$
End Sub


Private Sub OnScreenText_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo Err_Handler

    If KeyAscii = 13 Then
        If OnScreenText(Index).text = "" Then OnscreenTextDelete (Index)
    End If
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS25:" & Error$
End Sub

Private Sub OnscreenTextDelete(Index As Integer)
On Error GoTo Err_Handler
    Dim NumberOfTextBoxes As Integer

    NumberOfTextBoxes = OnScreenText.Count - 1
    OnScreenText(Index).Left = OnScreenText(NumberOfTextBoxes).Left
    OnScreenText(Index).Top = OnScreenText(NumberOfTextBoxes).Top
    OnScreenText(Index).text = OnScreenText(NumberOfTextBoxes).text
    Unload OnScreenText(NumberOfTextBoxes)

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS26:" & Error$
End Sub

Public Sub OnscreenTextMoveAll(X As Single, Y As Single)
On Error GoTo Err_Handler
    Dim NumberOfTextBoxes As Integer
    Dim Count As Integer

    NumberOfTextBoxes = OnScreenText.Count - 1
    For Count = 1 To NumberOfTextBoxes
        OnScreenText(Count).Left = OnScreenText(Count).Left + X
        OnScreenText(Count).Top = OnScreenText(Count).Top + Y
    Next Count
    
Exit Sub
Err_Handler:
    MsgBox Err & " - " & Error$

End Sub
Public Sub OnscreenTextMove(Index As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    OnScreenText(Index).Left = OnScreenText(Index).Left + X
    OnScreenText(Index).Top = OnScreenText(Index).Top + Y
    'OnScreenText(Index).left = X
    'OnScreenText(Index).Top = Y
    
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS27:" & Error$
End Sub
Sub OnscreenTextResetAll()
On Error GoTo Err_Handler
    Dim NumberOfTextBoxes As Integer
    Dim Count As Integer

    NumberOfTextBoxes = OnScreenText.Count - 1
    For Count = 1 To NumberOfTextBoxes
        OnScreenText(Count).BackColor = OnScreenText(0).BackColor
        OnScreenText(Count).BorderStyle = OnScreenText(0).BorderStyle
        OnScreenText(Count).Font = OnScreenText(0).Font
        OnScreenText(Count).FontBold = OnScreenText(0).FontBold
        OnScreenText(Count).FontItalic = OnScreenText(0).FontItalic
        OnScreenText(Count).FontName = OnScreenText(0).FontName
        OnScreenText(Count).FontSize = OnScreenText(0).FontSize
        OnScreenText(Count).FontStrikethru = OnScreenText(0).FontStrikethru
        OnScreenText(Count).FontUnderline = OnScreenText(0).FontUnderline
        OnScreenText(Count).ForeColor = OnScreenText(0).ForeColor
        Call SetTextBoxWidthAndHeight(OnScreenText(Count))
        OnScreenText(Count).Refresh
        
    Next Count
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS28:" & Error$
End Sub
Sub OnscreenTextResetAllToDefault()
On Error GoTo Err_Handler
    Dim NumberOfTextBoxes As Integer
    Dim Count As Integer

    NumberOfTextBoxes = OnScreenText.Count - 1
    For Count = 0 To NumberOfTextBoxes
        OnScreenText(Count).BackColor = OnscreenTextDefault.BackColor
        OnScreenText(Count).BorderStyle = OnscreenTextDefault.BorderStyle
        OnScreenText(Count).Font = OnscreenTextDefault.Font
        OnScreenText(Count).FontBold = OnscreenTextDefault.FontBold
        OnScreenText(Count).FontItalic = OnscreenTextDefault.FontItalic
        OnScreenText(Count).FontName = OnscreenTextDefault.FontName
        OnScreenText(Count).FontSize = OnscreenTextDefault.FontSize
        OnScreenText(Count).FontStrikethru = OnscreenTextDefault.FontStrikethru
        OnScreenText(Count).FontUnderline = OnscreenTextDefault.FontUnderline
        OnScreenText(Count).ForeColor = OnscreenTextDefault.ForeColor
        Call SetTextBoxWidthAndHeight(OnScreenText(Count))
        OnScreenText(Count).Refresh
        
    Next Count
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS29:" & Error$
End Sub
Public Sub OnscreenTextDeleteAll()
On Error GoTo Err_Handler
    Dim NumberOfTextBoxes As Integer
    Dim Count As Integer

    NumberOfTextBoxes = OnScreenText.Count - 1
    For Count = 1 To NumberOfTextBoxes
        Call OnscreenTextDelete(1)
    Next Count
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS30:" & Error$
End Sub

Public Sub OnscreenTextHide()
On Error GoTo Err_Handler
    Dim Count As Integer
    Dim NumberOfTextBoxes As Integer

    NumberOfTextBoxes = OnScreenText.Count - 1
    For Count = 1 To NumberOfTextBoxes
        OnScreenText(Count).Visible = False
    Next Count

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS31:" & Error$
End Sub

Public Sub OnscreenTextShow()
On Error GoTo Err_Handler
    
    Dim Count As Integer
    Dim NumberOfTextBoxes As Integer

    NumberOfTextBoxes = OnScreenText.Count - 1
    For Count = 1 To NumberOfTextBoxes
        OnScreenText(Count).Visible = True
        OnScreenText(Count).ZOrder 0
    Next Count

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS32:" & Error$
End Sub



Private Sub OnScreenText_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
 'Call SetTextBoxWidthAndHeight(OnScreenText(Index))
End Sub

Private Sub OnScreenText_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    If Button = vbLeftButton Then
        If CLPScreenAction = "MoveAll" Then
            OnScreenText(Index).MousePointer = 99
            OnScreenText(Index).MouseIcon = LoadResPicture(123, vbResIcon) 'grab text icon
            Call ReleaseCapture
            Call SendMessage(OnScreenText(Index).hwnd, WM_NCLBUTTONDOWN, ByVal HTCAPTION, ByVal 0&)

        End If
    ElseIf Button = vbRightButton Then
        'The following three lines disables the default edit popup menu from http://www.devx.com/vb2themax/Tip/18376
        OnScreenText(Index).Enabled = False ' disable the textbox
        DoEvents                            ' (this DoEvents seems to be optional)
        OnScreenText(Index).Enabled = True  ' re-enable the control, so that it doesn't appear as grayed
        OnScreenText(0).Tag = Index
        PopupMenu OnScreenTextMenu 'show your custom menu
        Call OnscreenTextResetAll
    
    End If
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS33:" & Error$
End Sub

Private Sub OnScreenText_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    If CLPScreenAction = "MoveAll" Then
        OnScreenText(Index).MousePointer = 99
        OnScreenText(Index).MouseIcon = LoadResPicture(122, vbResIcon) 'Move holding text icon
    Else
        OnScreenText(Index).MousePointer = vbIbeam
    End If
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS34:" & Error$
End Sub

Private Sub OnScreenText_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS35:" & Error$
End Sub

Private Sub OnScreenTextBackgroundColourMenu_Click()
On Error GoTo Err_Handler
    OnScreenTextDialog.CancelError = True
    OnScreenTextDialog.Flags = cdlCCRGBInit
    OnScreenTextDialog.Color = OnScreenText(0).BackColor
    OnScreenTextDialog.ShowColor
    
    OnScreenText(0).BackColor = OnScreenTextDialog.Color
    Call OnscreenTextResetAll

Exit Sub
Err_Handler:
    Select Case Err
        Case 32755: Exit Sub 'Cancel
        Case Else: MsgBox Err & "-CLS36:" & Error$
    End Select
End Sub

Private Sub OnscreenTextDefaultMenu_Click()
    Call OnscreenTextResetAllToDefault
End Sub

Private Sub OnscreenTextDeleteAllMenu_Click()
On Error GoTo Err_Handler
    Dim Index As Integer
    
    Index = OnScreenText(0).Tag
    Call OnscreenTextDeleteAll
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS37:" & Error$
End Sub

Private Sub OnscreenTextDeleteMenu_Click()
On Error GoTo Err_Handler
    Dim Index As Integer
    
    Index = OnScreenText(0).Tag
    Call OnscreenTextDelete(Index)
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS38:" & Error$
End Sub

Private Sub OnscreenTextFontMenu_Click()
On Error GoTo Err_Handler
    OnScreenTextDialog.CancelError = True

  ' Set the Flags property
    OnScreenTextDialog.Flags = cdlCFEffects Or cdlCFBoth
    OnScreenTextDialog.FontName = OnScreenText(0).Font.name
    OnScreenTextDialog.FontSize = OnScreenText(0).Font.Size
    OnScreenTextDialog.FontBold = OnScreenText(0).Font.Bold
    OnScreenTextDialog.FontItalic = OnScreenText(0).Font.Italic
    OnScreenTextDialog.FontUnderline = OnScreenText(0).Font.Underline
    OnScreenTextDialog.FontStrikethru = OnScreenText(0).FontStrikethru
    OnScreenTextDialog.Color = OnScreenText(0).ForeColor

    OnScreenTextDialog.ShowFont
    
    OnScreenText(0).Font.name = OnScreenTextDialog.FontName
    OnScreenText(0).Font.Size = OnScreenTextDialog.FontSize
    OnScreenText(0).Font.Bold = OnScreenTextDialog.FontBold
    OnScreenText(0).Font.Italic = OnScreenTextDialog.FontItalic
    OnScreenText(0).Font.Underline = OnScreenTextDialog.FontUnderline
    OnScreenText(0).FontStrikethru = OnScreenTextDialog.FontStrikethru
    OnScreenText(0).ForeColor = OnScreenTextDialog.Color
Exit Sub
Err_Handler:
    Select Case Err
        Case 32755: Exit Sub 'Cancel
        Case Else: MsgBox Err & "-CLS39:" & Error$
    End Select
End Sub

Sub OnscreenTextDrawAll(screenctrl As Control, DrawSF As Double)
On Error GoTo Err_Handler

    Dim NumberOfOnscreenText As Integer
    Dim Count As Integer
    Dim XCoordA, YCoordA, XCoordB, YCoordB As Integer
    Dim TextScreenWidth, TextScreenHeight As Integer
    
        
    NumberOfOnscreenText = ClearLineScreen.OnScreenText.Count - 1
    For Count = 1 To NumberOfOnscreenText
        XCoordA = ClearLineScreen.OnScreenText(Count).Left
        YCoordA = ClearLineScreen.OnScreenText(Count).Top
        XCoordB = XCoordA + ClearLineScreen.OnScreenText(Count).width
        YCoordB = YCoordA + ClearLineScreen.OnScreenText(Count).height
        XCoordA = (XCoordA - 1) * DrawSF
        YCoordA = (YCoordA - 2) * DrawSF
        XCoordB = (XCoordB - 1) * DrawSF
        YCoordB = (YCoordB - 2) * DrawSF
        
        
        screenctrl.Line (XCoordA, YCoordA)-(XCoordB, YCoordB), ClearLineScreen.OnScreenText(Count).BackColor, BF
        screenctrl.Line (XCoordA, YCoordA)-(XCoordB, YCoordB), vbBlack, B
        
        screenctrl.Font = ClearLineScreen.OnScreenText(Count).Font
        screenctrl.FontBold = ClearLineScreen.OnScreenText(Count).FontBold
        screenctrl.FontItalic = ClearLineScreen.OnScreenText(Count).FontItalic
        screenctrl.FontName = ClearLineScreen.OnScreenText(Count).FontName
        screenctrl.FontSize = ClearLineScreen.OnScreenText(Count).FontSize * DrawSF
        screenctrl.FontStrikethru = ClearLineScreen.OnScreenText(Count).FontStrikethru
        screenctrl.FontUnderline = ClearLineScreen.OnScreenText(Count).FontUnderline
        screenctrl.ForeColor = ClearLineScreen.OnScreenText(Count).ForeColor
        
        TextScreenWidth = screenctrl.TextWidth(ClearLineScreen.OnScreenText(Count).text)
        TextScreenWidth = TextScreenWidth
        
        screenctrl.CurrentX = (((XCoordB - XCoordA) - TextScreenWidth) / 2) + XCoordA
        screenctrl.CurrentY = YCoordA
                
        screenctrl.Print ClearLineScreen.OnScreenText(Count).text
        
    Next Count
    screenctrl.FillStyle = vbFSTransparent
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS40:" & Error$
End Sub

Sub Calibrate(CalibrationType As String)  'PCN1825
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Calibrate function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    6/2/03
'   V1 Antony van Iersel 10 March 2006
'   PCN4046
'Description:
'       Sets up calibration of the PVScreen or SnapShotScreen
'
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Dim CalibrationTypeLengthInInches As Double 'PCN1825
Call ScreenDrawing.SetupDrawingRoutine("DrawCalibrationLine")

Select Case CalibrationType 'PCN1825

    Case "C600"
        CalibrationTypeLength = 600
        CalibrationTypeLengthInInches = 24
    Case "C500"
        CalibrationTypeLength = 500
        CalibrationTypeLengthInInches = 20
    Case "C400"
        CalibrationTypeLength = 400
        CalibrationTypeLengthInInches = 16
    Case "C300"
        CalibrationTypeLength = 300
        CalibrationTypeLengthInInches = 12
    Case "C200"
        CalibrationTypeLength = 200
        CalibrationTypeLengthInInches = 8
    '^^^^ **********************************
    Case "CalSnapOn"
        'To be developed
        CalibrationTypeLength = "" 'This should be held in the INI file - Stored as mm
        Exit Sub      'PCN3176
    Case "CalManual"
        'No action
        CalibrationTypeLength = 0
        Exit Sub      'PCN3176
    Case Else
        Exit Sub      'PCN3176
End Select

'vvvv PCN3176 **************************
If MeasurementUnits <> "mm" And CalibrationTypeLength <> 0 Then  'Check if the length needs to be converted to inches
     CalibrationTypeLength = CalibrationTypeLengthInInches
End If
'^^^^ **********************************
    
Exit Sub
Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & "-CLS41:" & Error$
End Select

End Sub




Private Sub ProfilerTimer_Timer()
On Error GoTo Err_Handler

If (mediatype = Video And PVDFileName = "") Or PVRecording = True Then
    Call VideoRecordingLoop: Exit Sub
ElseIf mediatype = Video And PVDFileName <> "" And CLPScreenMode = Video Then
    Call VideoPlaybackLoop: Exit Sub
ElseIf PVDFileName <> "" Then
    Call PVDPlaybackLoop
Else
    Call ProfilerPause 'If none of these conditions are true, then something is amiss and stop the video/pvd playback
End If

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-CLS42:" & Error$
    End Select
End Sub

Private Sub PVDPlaybackLoop()
On Error GoTo Err_Handler

If PVFrameNo >= PVDataNoOfLines Then
    Call GotoPVProfile(PVDataNoOfLines, True)
    Call ProfilerPause
    Call MoveSlider(PVTimes(PVFrameNo))
    Exit Sub
End If
PVFrameNo = PVFrameNo + 1
Call GotoPVProfile(PVFrameNo, False)
Call MoveSlider(PVTimes(PVFrameNo))

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS43:" & Error$
End Sub


Private Sub PVScreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PVScreen_MouseDown Sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    16/11/02     Building initial framework
'
'Description:
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

    Call DrawMouseDown(ClearLineScreen.PVScreen, Button, Shift, X, Y)

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS44:" & Error$
End Sub

Private Sub PVScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PVScreen_MouseMove Sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    16/11/02     Building initial framework
'
'Description:
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Call DrawMouseMove(ClearLineScreen.PVScreen, ButtonNumber, Shift, X, Y)
Call DrawMainScaleRedLines(X, Y)

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS45:" & Error$
End Sub

Function DrawNormalLine(X, Y)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'DrawNormalLine Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    16/11/02     Building initial framework
'
'Description:
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim PipeFill As Integer

'vvvvvv  Dimensioning Tool **** For concept proving 16/11/02
If CLPScreenMode = Video Then
    MainScreen.Circle (6100, 3550), 3150, RGB(100, 180, 22) ' Expected internal diameter of host pipe
    MainScreen.Circle (6100, 3550), 3650, RGB(100, 180, 22) ' Expected external diameter of host pipe
    For PipeFill = 3200 To 3600 Step 75
        'MainScreen.Circle (6100, 3550), PipeFill, RGB(100, 180, 22) ' Expected external diameter of host pipe
        MainScreen.Circle (6100, 3550), PipeFill, RGB(50, 90, 10) ' Expected external diameter of host pipe
    Next PipeFill
    'redraw Limits
    MainScreen.Circle (6100, 3550), 2750, RGB(215, 115, 250) ' Expected internal diameter of host pipe
    For PipeFill = 0 To 2700 Step 175
        'MainScreen.Circle (6100, 3550), PipeFill, RGB(240, 190, 255) ' Expected external diameter of host pipe
        MainScreen.Circle (6100, 3550), PipeFill, RGB(70, 0, 100) ' Expected external diameter of host pipe
    Next PipeFill
ElseIf CLPScreenMode = PV Then
    PVDimensionLabel.Left = LineX1 + (X - LineX1) / 2
    PVDimensionLabel.Top = LineY1 + (Y - LineY1) / 2
    PVDimensionLabel.Caption = (X - LineX1)
    LineX2 = X
    LineY2 = Y
    PVDimensionLabel.Visible = True
    'Draw dimension line
    MainScreen.Cls
    MainScreen.Line (LineX1, LineY1)-(LineX2, LineY2)
    'redraw shapes
    'MainScreen.Circle (5050, 4100), 2350, RGB(0, 255, 0) 'Green
    'MainScreen.Circle (5050, 4100), 2850, RGB(0, 255, 0) 'Green
    'MainScreen.Circle (5050, 4100), 2350, RGB(190, 115, 25) ' Expected internal diameter of host pipe
    'MainScreen.Circle (5050, 4100), 2850, RGB(190, 115, 25) ' Expected external diameter of host pipe
    'MainScreen.Circle (5050, 4100), 2350, RGB(180, 100, 22) ' Expected internal diameter of host pipe
    'MainScreen.Circle (5050, 4100), 2850, RGB(180, 100, 22) ' Expected external diameter of host pipe
    MainScreen.Circle (5050, 4100), 2350, RGB(100, 180, 22) ' Expected internal diameter of host pipe
    MainScreen.Circle (5050, 4100), 2850, RGB(100, 180, 22) ' Expected external diameter of host pipe
    For PipeFill = 2400 To 2800 Step 50
        MainScreen.Circle (5050, 4100), PipeFill, RGB(225, 245, 160) ' Expected external diameter of host pipe
    Next PipeFill
    'redraw Limits
    MainScreen.Circle (5050, 4100), 2100, RGB(215, 115, 250) ' Expected internal diameter of host pipe
    For PipeFill = 0 To 2000 Step 175
        MainScreen.Circle (5050, 4100), PipeFill, RGB(240, 190, 255) ' Expected external diameter of host pipe
    Next PipeFill
End If
'^^^^^^  Dimensioning Tool **** For concept proving 16/11/02



Exit Function
Err_Handler:
    MsgBox Err & "-CLS46:" & Error$
End Function

Private Sub PVScreen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PVScreen_MouseUp Sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    16/11/02     Building initial framework
'
'Description:
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Call DrawMouseUp(ClearLineScreen.PVScreen, Button, Shift, X, Y)

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS47:" & Error$
End Sub




Function SetupMouseIcon(MouseIconID) 'PCNGL021202
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'SetupMouseIcon Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    16/11/02     Building initial framework
'
'Description:
'
'Purpose:
'   MouseIconName is the file name of the mouse icon
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim curSelect As StdPicture

'vvvv PCN4252 **********************
If MouseIconID = 0 Then
    CLPScreenDrawState = ""
    CLPScreenAction = ""
    PVScreen.MousePointer = 2
    SnapShotScreen.MousePointer = 2
    Exit Function
End If
'^^^^ ******************************

MainScreenMouseIcon = MouseIconID
Set curSelect = LoadResPicture(MainScreenMouseIcon, vbResIcon)
SnapShotScreen.MousePointer = 99
SnapShotScreen.MouseIcon = curSelect
PVScreen.MousePointer = 99
PVScreen.MouseIcon = curSelect

Exit Function
Err_Handler:
    Select Case Err
        Case 53 'File not found
            PVScreen.MousePointer = 2
            Resume Next
        Case 75 'File not found
            PVScreen.MousePointer = 2
            Resume Next
        Case Else
            MsgBox Err & "-CLS48:" & Error$
    End Select
End Function










Sub SetGrabedImage(GrabedImage)
On Error GoTo Err_Handler
Dim pth As String
'    MainScreen.Picture = GrabedImage          ' Load image.
'    CurrentFile1 = "GrabedImage"
'    CurrentFile = "GrabedImage"
''   StoreOriginalImage
'    OrigWidth = MainScreen.ScaleWidth
'    OrigHeight = MainScreen.ScaleHeight
'    Clear_Click2
'    Form_Resize
    PVScreen.Picture = GrabedImage
    CurrentFile1 = "GrabedImage" 'added to loadimage rtn
    CurrentFile = "GrabedImage" 'added to loadimage rtn
    ResizeFlag = False
    OrigWidth = PVScreen.ScaleWidth
    OrigHeight = PVScreen.ScaleHeight
    'SnapShotScreen.Picture = GrabedImage 'removed from loadimage rtn
    ResizeFlag = True
'    Clear_Click2 'added to loadimage rtn
    Call LoadDummyScreenImage
Exit Sub
Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & "-CLS50:" & Error$
End Select
End Sub



Public Sub LoadDummyScreenImage()
On Error GoTo Err_Handler


SnapShotScreen.AutoRedraw = True 'PCN2124
DummyScreen.AutoRedraw = True 'PCN2124
DummyScreen.Picture = LoadPicture(SnapshotFilename)
Dim propx As Double
Dim propy As Double
If MainScreen.ScaleHeight > 0 And MainScreen.ScaleWidth > 0 And ResizeFlag Then
    SnapShotScreen.height = MainScreen.height 'PCNGL2401032
    SnapShotScreen.width = MainScreen.width 'PCNGL2401032
    'vvvv PCN1965 *************************************
    If ConfigInfo.MediaWidth > 0 And ConfigInfo.MediaHeight > 0 Then 'PCN1965
        propx = SnapShotScreen.ScaleHeight / SnapShotScreen.ScaleWidth
        propy = ConfigInfo.MediaHeight / ConfigInfo.MediaWidth 'PCN1965
        If propx > propy Then
            pich = Int(SnapShotScreen.ScaleWidth * propy)
            picw = Int(SnapShotScreen.ScaleWidth)
        Else
            picw = Int(SnapShotScreen.ScaleHeight / propy)
            pich = Int(SnapShotScreen.ScaleHeight)
        End If
        If DummyScreen.Picture <> 0 Then
            If YAdjust = 0 Then
                SnapShotScreen.PaintPicture DummyScreen.Picture, 0, 0, picw, pich 'PCN1965
            Else
                SnapShotScreen.PaintPicture DummyScreen.Picture, 0, 0, picw, pich  'PCN1965
            End If
            SnapShotScreen.Picture = SnapShotScreen.Image 'PCN3194
        End If
        SnapShotScreen.ZOrder 0
        SnapShotScreen.Visible = True
    Else
        propx = SnapShotScreen.ScaleHeight / SnapShotScreen.ScaleWidth
        pich = Int(SnapShotScreen.ScaleWidth * propx)
        picw = Int(SnapShotScreen.ScaleWidth)
        If DummyScreen.Picture <> 0 Then
            If YAdjust = 0 Then
                SnapShotScreen.PaintPicture DummyScreen.Picture, 0, 0, picw, pich 'PCN1965
            Else
                SnapShotScreen.PaintPicture DummyScreen.Picture, 0, 0, picw, pich  'PCN1965
            End If
            SnapShotScreen.Picture = SnapShotScreen.Image
        End If
    End If

    '^^^^ *********************************************
End If

SnapShotScreen.AutoRedraw = False 'PCN2124

Exit Sub
Err_Handler:
Select Case Err
    Case 481 'Invalid picture
        Resume Next
    Case Else
        MsgBox Err & "-CLS51:" & Error$
        'Resume 'PCNLS
End Select
End Sub



Sub Initialization()
On Error GoTo Err_Handler

''Setup centre line co-ordinates of the PV screen
'CentreLineX = Round(ClearLineScreen.MainScreen.width / 2)
'CentreLineY = Round(ClearLineScreen.MainScreen.height / 2)
'
'PaintNow = False
'BrushNow = False
'
'    TotPoint = 0
'    If CLPScreenAction = "MoveAll" And DrawAutoSnap = False Then 'PCNGL210103
'        Chk4SnapEnabled = 1
'        DrawAutoSnap = True 'PCNGL210103
'    End If
'    Chk4SnapEnabled = True
'
'
'    MoveFlag = False
'    If CurrentShape = -8 Or CurrentShape = -11 Or CurrentShape = -4 Or CurrentShape = -5 Or _
'       CurrentShape = 9 Or CurrentShape = -10 Or CurrentShp = -16 Or CurrentShp = -17 Or _
'       CurrentShape = -18 Or CLPScreenAction = "DrawText" Or CLPScreenAction = "Pattern Fill" Then
'       SnapShotScreen.AutoRedraw = True
'    End If
'    If CLPScreenAction = "DrawText" Then
'        SnapShotScreen.AutoRedraw = True
'    End If
'    CurrentShape = -9
'                        ' -24:Pattern Fill
'                        ' -23: Include a line for Dimension
'                        ' -22:Filesave, -21:TextW, -20:Rotate,
'                        ' -18:Track mouse icon for edge detection,
'                        ' -17: partial Detection, -16:save/load file(canvas2)->Not used.,
'                        ' -15:Undo, -14:Redo,
'                        ' -12:Multiselect,-11:Area, -10:Length,
'                        ' -9:nothing, -8:Edge/Repaint, -7 : move, -6: delete, -5:Brush,
'                        ' -4:fill, -3:Modify, -2:desel=multidesel, -1:sel(toggle), 0:dot,1:line,
'                        ' 2:circle, 3:arc, 4:curve, 5:Egg Shape, 9:erase
'
'    CurrentShp = 0 'Abe - test
'                        ' -25:an Automized procedure(SHP)
'                        ' -23:Clock From To(SHP) or Dimension , -24:Angle From To(SHP)->being not used,
'                        ' -16,17:Unknown - stop SnapShotScreen.cls on mousemove
'
'
'    '150102Gauge.MessageBox = "" 'PCN 31
'    JointExist = False
'    FirstJointSuccess = False
'    SecondJointSuccess = False
'
'    If MainScreenMouseIcon = 0 Then 'PCNGL021202
'        SnapShotScreen.MousePointer = 2  'Cross
'        PVScreen.MousePointer = 2  'Cross
'    Else
'        SnapShotScreen.MousePointer = 99
'        SnapShotScreen.MouseIcon = LoadResPicture(MainScreenMouseIcon, vbResIcon)
'        PVScreen.MousePointer = 99
'        'PVScreen.MouseIcon = LoadPicture(App.Path & MainScreenMouseIcon)
'        PVScreen.MouseIcon = LoadResPicture(MainScreenMouseIcon, vbResIcon)
'
'    End If
'
'    EDStartFlag = False
'
'
'
'    StartedOnCircle = False
'    OrthoLineFlag = False
'    If ApproxAreaFlag Then
'        ApproxAreaFlag = False
'        If Not ApproxAreaStartedFlag Then
'            TotOneObjNo = TotOneObjNo - 1
'        Else
'            ApproxAreaStartedFlag = False
'        End If
'    End If
'    CalFlag = False
'    ObsFlag = False
'    BeforeFill = False

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS52:" & Error$
End Sub



Function LetSavePictureOnly(PicName As String)
On Error GoTo Err_Handler
    SavePicture SnapShotScreen, PicName
    LetSavePictureOnly = True
Exit Function
Err_Handler:
Select Case Err
    Case 380 ' invalid property value
        'MsgBox "No picture on this picture box. Please capture an image first.", vbInformation, "Image Control" 'PCNGL140103
        LetSavePictureOnly = False
        Exit Function
    Case Else
        MsgBox Err & "-CLS53:" & Error$
End Select
End Function


Public Sub Select_Click()
On Error GoTo Err_Handler
    Initialization
    CLPScreenAction = "Select"
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS54:" & Error$
End Sub

''Public Sub Modify_Click()
''On Error GoTo Err_Handler
''
''Initialization
''CLPScreenAction = "Modify"
''ModifyStartFlag = False
''
''Exit Sub
''Err_Handler:
''    MsgBox Err & " - " & error$
''End Sub

Public Sub NormalLine_Click()
On Error GoTo Err_Handler
    Initialization
    LineBtnClick
    TotLength = 0
    MidLineFlag = False 'PCNGL290103
    
Exit Sub
Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & "-CLS55:" & Error$
End Select
End Sub

Public Sub PolyLine_Click()
On Error GoTo Err_Handler

   Initialization
   LineBtnClick
   PolyLineFlag = True
   PolyLineStarted = True
'  Chk4AutoTriangle = 0
   TotLength = 0
    MidLineFlag = False 'PCNGL290103
   
Exit Sub
Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & "-CLS56:" & Error$
End Select
End Sub


Public Sub Move_Click()
On Error GoTo Err_Handler

    Initialization
    CLPScreenAction = "MoveAll"
    'disable autosnap.
    'Chk4SnapEnabled.Value = 0
    'autosnap.Checked = False
    '''
    JointExist = False
    FirstJointSuccess = False
    SecondJointSuccess = False
    SnapShotScreen.MousePointer = 99
    SnapShotScreen.MouseIcon = LoadResPicture(MainScreenMouseIcon, vbResIcon) 'PCNGL291102
    PVScreen.MousePointer = 99 'PCNGL291102
    PVScreen.MouseIcon = LoadResPicture(MainScreenMouseIcon, vbResIcon) 'PCNGL291102

Exit Sub
Err_Handler:
Select Case Err
    Case 53 ' Can't find mouse icon
        Resume Next
    Case Else
        MsgBox Err & "-CLS57:" & Error$
End Select
End Sub


Public Sub Text_Click()
On Error GoTo Err_Handler

''    If Not isopen("PipelineDetails") Then
    PipelineDetails.Show

    DoEvents 'PCNGL040603-2

'    PipelineDetails.SetFocus
'    End If
    Initialization
    CLPScreenAction = "DrawText"
    SnapShotScreen.MousePointer = 3 ' I beam

    SnapShotScreen.AutoRedraw = False
       
    
   
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS58:" & Error$
End Sub

Public Sub Calibration_Click()
On Error GoTo Err_Handler

    Initialization
    LineBtnClick
    TotLength = 0
    CalFlag = True
    CalDrawFlag = True
    If Len(PipelineDetails.InternalDiameterExpected.text) = 0 Then 'PCN 31
        'MsgBox DisplayMessage("Please specify internal pipe diameter first."), vbExclamation 'PCN2111
        ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Please specify internal pipe diameter first."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
        PipelineDetails.ZOrder (0) 'PNN1865 "PNSLS200203
        Exit Sub
    End If
    
Exit Sub
Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & "-CLS59:" & Error$
End Select
End Sub

Private Sub LineBtnClick()
On Error GoTo Err_Handler

    CLPScreenAction = "DrawLine"
    PolyLineFlag = False
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS60:" & Error$
End Sub

Function ColorValue(colr)
    Select Case colr
    Case "Black"
        ColorValue = vbBlack '0
    Case "Blue"
        ColorValue = vbBlue '4
    Case "Green"
        ColorValue = vbGreen '2
    Case "Red"
        ColorValue = vbRed '1
    Case "Cyan"
        ColorValue = vbCyan '6
    Case "Magenta"
        ColorValue = vbMagenta '5
    Case "White"
        ColorValue = vbWhite '7
    Case "Yellow"
        ColorValue = vbYellow '3
    Case Else
        ColorValue = SnapShotScreen.ForeColor
    End Select
End Function

Public Sub LoadImage(Filenm)
On Error GoTo Err_Handler
    PipelineDetails.Visible = True
'    DummyScreen.Picture = LoadPicture(Filenm) 'PCN2124
    ResizeFlag = False
'    OrigWidth = DummyScreen.ScaleWidth 'PCN2124
'    OrigHeight = DummyScreen.ScaleHeight 'PCN2124
'    Debug.Print DummyScreen.ScaleWidth 'PCNGL2401032
'    Debug.Print DummyScreen.ScaleHeight
'PCN 1966 LS 8/7/03 (uncomment the following line)
'    SnapShotScreen.Picture = LoadPicture(Filenm) 'PCN2124
    ResizeFlag = True
    SnapshotFilename = Filenm
    Call LoadDummyScreenImage
    
Exit Sub
Err_Handler:
Select Case Err
    Case 481 'Invalid picture, unable to load
        'MsgBox DisplayMessage("Picture file is invalid, unable to open."), vbExclamation 'PCN2111
        ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Picture file is invalid, unable to open."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Case Else
        MsgBox Err & "-CLS61:" & Error$
End Select
End Sub



Private Sub SecretButtonHere_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Registered = False Or SoftwareConfiguration = "Reader" Then Exit Sub
    
    If Shift = 7 And Button = 4 Then DebugForm.ZOrder 0 'PCN4259
    If Shift = 7 And Button = 2 Then DebugForm.ZOrder 0
End Sub

Private Sub SnapShotScreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Call DrawMouseDown(ClearLineScreen.SnapShotScreen, Button, Shift, X, Y)

Exit Sub
Err_Handler:
Select Case Err
    Case 53 ' Could not find the mouse icon
        Resume Next
    Case Else
    MsgBox Err & "-CLS62:" & Error$
End Select
End Sub
Private Sub SnapShotScreen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Call DrawMouseUp(ClearLineScreen.SnapShotScreen, Button, Shift, X, Y)

Exit Sub
Err_Handler:
Select Case Err
    Case 53 ' Could not find the mouse icon
        Resume Next
    Case Else
    MsgBox Err & "-CLS63:" & Error$
End Select
End Sub

Private Sub SnapShotScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Call DrawMouseMove(ClearLineScreen.SnapShotScreen, Button, Shift, X, Y)
Call DrawMainScaleRedLines(X, Y)

Exit Sub
Err_Handler:
Select Case Err
    Case 53 ' Could not find the mouse icon
        Resume Next
    Case Else
    MsgBox Err & "-CLS64:" & Error$
End Select
End Sub ' End of 'MouseMove'

Function Arccos(X) As Double
On Error GoTo Err_Handler
If Sqr(-X * X + 1) <> 0 Then
    Arccos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
Else
    Arccos = Atn(0) + 2 * Atn(1)
End If

Exit Function
Err_Handler:
Select Case Err
    Case 11 ' divide by zero
'       ab = ab + 0.001
        MsgBox Err & "-CLS65:" & Error$ 'PCN2111
    Case 5 '
        MsgBox Err & "-CLS66:" & Error$
        'MsgBox "Temporary message after mouseup. Please check your drawing is as intended. " & Error$
    Case Else
        MsgBox Err & "-CLS67:" & Error$
End Select
End Function

Function CalAsin(sx, sy, X, Y) As Double
On Error GoTo Err_Handler
Dim a As Double
Dim b As Double
Dim ab As Double
Dim asin As Double
          
          a = Sqr((sx - X) ^ 2 + (sy - Y) ^ 2)
          b = sx - X
          If a = 0 Then
            ab = 0
          Else
            ab = b / a
          End If
          If ab = 1 Then
            asin = PI / 2
          ElseIf ab = -1 Then
            asin = 3 * PI / 2
          Else
            If a = 0 Then
                asin = PI / 2
            Else
                asin = Atn(ab / Sqr(-ab * ab + 1))
            End If
          End If

          If Y < sy And X >= sx Then ' 1
            asin = PI / 2 + asin
          ElseIf Y >= sy And X >= sx Then ' 2
            asin = 3 * PI / 2 - asin
          ElseIf Y >= sy And X < sx Then ' 3
            asin = 3 * PI / 2 - asin
          ElseIf Y < sy And X < sx Then ' 4
            asin = PI / 2 + asin
          End If

          CalAsin = asin

Exit Function

Err_Handler:

Select Case Err
    Case 11 ' divide by zero
'       ab = ab + 0.001
        'MsgBox "Divide by zero" 'PCN2111
        MsgBox Err & "-CLS68:" & Error$
        Resume Next
    Case 5 '
        'MsgBox "Temporary message after mouseup. Please check your drawing is as intended. " & error$ 'PCN2111
        Resume Next
    Case Else
        MsgBox Err & "-CLS69:" & Error$
End Select

End Function




Private Sub PVScreenPicInPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'PCN1863
On Error GoTo Err_Handler

PVScreenPicInPic.MousePointer = 99 'PCN1863
PVScreenPicInPic.MouseIcon = LoadResPicture(109, vbResIcon) 'Move holding icon -'PCN1863
PicInPicMouseDownY = -Y 'PCN1863
PicInPicMouseDownX = -X 'PCN1863

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS70:" & Error$
End Sub

Private Sub PVScreenPicInPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

If Button = 1 Then  'PCN1863
    PVScreenPicInPic.Top = PVScreenPicInPic.Top + PicInPicMouseDownY + Y  'PCN1863
    PVScreenPicInPic.Left = PVScreenPicInPic.Left + PicInPicMouseDownY + X    'PCN1863
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS71:" & Error$
End Sub

Private Sub PVScreenPicInPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

PVScreenPicInPic.MousePointer = 99 'PCN1863
PVScreenPicInPic.MouseIcon = LoadResPicture(108, vbResIcon) 'Move icon -'PCN1863

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS72:" & Error$
End Sub


' FISH-EYE PCN2290 --------------------------------------------------v
Private Sub PVYScale_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'FISH-EYE PCN2290
On Error GoTo Err_Handler
'PCNANT???? not used anymore
'    Dim PixelVal As Long
'    Dim PixelVal2 As Long
'    PixelVal = fTwipsToPixels(7680, DIRECTION_HORIZONTAL)
'    PixelVal2 = fTwipsToPixels(350, DIRECTION_HORIZONTAL)
'    If Button = 1 Then
'        YMarker1st.Y1 = Y
'        YMarker1st.Y2 = Y
'        Call sethline((PixelVal - Y - PixelVal2) / PixelVal)
'    Else
'        YMarker2nd.Y1 = Y
'        YMarker2nd.Y2 = Y
'        Call sethline2nd((PixelVal - Y - PixelVal2) / PixelVal)
'    End If
'    If AVIPlay = False Then
'        If MediaType <> "Live" Then 'PCN2451
'            Call refscreen 'fisheyeinterface
'        End If
'    End If
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS73:" & Error$
    
End Sub
Private Sub PVXScale_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'FISH-EYE PCN2290
On Error GoTo Err_Handler
'PCNANT???? not used anymore
'    If Button = 1 Then
'        XMarker1st.X1 = X
'        XMarker1st.X2 = X
'        Call setvline(X / PVXScale.width)
'    Else
'        XMarker2nd.X1 = X
'        XMarker2nd.X2 = X
'        Call setvline2nd(X / PVXScale.width)
'    End If
'    If AVIPlay = False Then
'        If MediaType <> "Live" Then 'PCN2451
'            Call refscreen 'fisheyeinterface
'        End If
'    End If
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS74:" & Error$
    
End Sub
' FISH-EYE PCN2290 --------------------------------------------------^
Private Sub ScaleSetting_Click()
On Error GoTo Err_Handler 'PCNGL300103

Call DrawMainScale(ClearLineScreen.PVScreen)

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS75:" & Error$
End Sub


Private Sub TerPL_Click(Index As Integer)
On Error GoTo Err_Handler
'''   LineBtnClick
''    If Not ApproxAreaFlag Then
''        If CurrentShp = -25 Then
''            Initialization
''            CurrentShp = -25
''        Else
''            Initialization
''        End If
''        CLPScreenAction = "DrawLine"
''    Else
''        CurrentShape = -9
''    End If
'''    If Not PolyLineFlag Then
'''        PipelineDetails.Len_Real = 0 'PCN 31
'''        PipelineDetails.LenRealPercent = 0 'PCN 31
'''    Else
''        PolyLineStarted = True
''        '010202
''        'If Len(Ratio) > 0 Then
''        '    PipelineDetails.Len_Real = Round(TotLength * Ratio, Digits) 'PCN 31
''        '    If Len(PipelineDetails.InternalDiameterExpected) > 0 Then 'PCN 31
''        '        PipelineDetails.LenRealPercent = Round(100 * PipelineDetails.Len_Real / PipelineDetails.InternalDiameterExpected, 1) 'PCN 31
''        '    End If
''        'End If
''    End If
''    SnapShotScreen.Cls
''    Call DrawAll(ClearLineScreen.SnapShotScreen)
''    Call DrawAll(ClearLineScreen.PVScreen)
''
Exit Sub
Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & "-CLS76:" & Error$
End Select
End Sub

Sub Sgmt3()
On Error GoTo Err_Handler
    Dim H As Long
    Dim i As Long
    Dim J As Long
    Dim k As Long
    Dim L As Long
    Dim m As Long
    Dim equiv_no As Long
    Dim equiv(SnapShotScreen_Height * 20) As Long
    Dim equiv2(SnapShotScreen_Height * 20) As Long
    
    ' Top-down pass
    For i = 1 To SnapShotScreen.ScaleWidth - 1 '0
        equiv_no = 0
        For J = 1 To SnapShotScreen.ScaleHeight - 1 '0
            If seg2(i, J) <> 0 Then
              If seg2(i + 1, J) <> 0 And seg2(i, J) <> seg2(i + 1, J) Then
                equiv(equiv_no) = seg2(i, J)
                equiv2(equiv_no) = seg2(i + 1, J)
                equiv_no = equiv_no + 1
               End If
'             If seg2(i + 1, j - 1) <> 0 And seg2(i, j) <> seg2(i + 1, j - 1) Then
'                equiv(equiv_no) = seg2(i, j)
'                equiv2(equiv_no) = seg2(i + 1, j - 1)
'                equiv_no = equiv_no + 1
'              End If
              If seg2(i, J + 1) <> 0 And seg2(i, J) <> seg2(i, J + 1) Then
                equiv2(equiv_no) = seg2(i, J)
                equiv(equiv_no) = seg2(i, J + 1)
                equiv_no = equiv_no + 1
              End If
            End If
        
            If equiv_no > 0 Then
              For k = 0 To equiv_no - 1 '0
                For m = 1 To SnapShotScreen.ScaleHeight - 1
                  For L = i To i + 1
                    If seg2(L, m) = equiv2(k) Then
                      seg2(L, m) = equiv(k)
                    End If
                  Next L
                Next m
              Next k
            End If
            equiv_no = 0
        
        Next J
    Next i

    ' Bottom-up pass
    For i = SnapShotScreen.ScaleWidth - 1 To 2 Step -1 '1
        equiv_no = 0
        For J = SnapShotScreen.ScaleHeight - 1 To 2 Step -1 '1
            If seg2(i, J) <> 0 Then
              If seg2(i - 1, J) <> 0 And seg2(i, J) <> seg2(i - 1, J) Then
                equiv(equiv_no) = seg2(i, J)
                equiv2(equiv_no) = seg2(i - 1, J)
                equiv_no = equiv_no + 1
              End If
'              If seg2(i - 1, j + 1) <> 0 And seg2(i, j) <> seg2(i - 1, j + 1) Then
'                equiv(equiv_no) = seg2(i, j)
'                equiv2(equiv_no) = seg2(i - 1, j + 1)
'                equiv_no = equiv_no + 1
'              End If
              If seg2(i, J - 1) <> 0 And seg2(i, J) <> seg2(i, J - 1) Then
                equiv2(equiv_no) = seg2(i, J)
                equiv(equiv_no) = seg2(i, J - 1)
                equiv_no = equiv_no + 1
              End If
            End If
            
            If equiv_no > 0 Then
              For k = 0 To equiv_no - 1
                For m = 1 To SnapShotScreen.ScaleHeight - 1
                  For L = i To i - 1 Step -1
                    If seg2(L, m) = equiv2(k) Then
                        seg2(L, m) = equiv(k)
                    End If
                  Next L
                Next m
              Next k
            End If
            equiv_no = 0
        
        Next J
    Next i
    
Exit Sub
Err_Handler:
Select Case Err
    Case 6 'overflow
        'MsgBox "Overflow - " & error$
        MsgBox Err & "-CLS77:" & Error$ 'PCN2111
    Case Else
        MsgBox Err & "-CLS78:" & Error$
End Select
End Sub



Function DrawPVGrid(PVScreenCntrl As Control, PVGridSize, CentreLineX, CentreLineY)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'DrawPVGrid function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    8/11/02     Building initial framework
'
'Description:
'   The Precision Vision PVScreen display window has a background grid
'   and centre lines to provide a visual measurement effect.
'Purpose:
'   PVGridSize = Precision Vision grid line spacing
'   CentreLineX, CentreLineY are the PVScreen centre lines
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim PVGridLine As Integer 'Position of grid line
Dim PVGridXYPos1 As Integer
Dim PVGridXYPos2 As Integer
Dim GridColour As Long
Dim RedReferenceRadius As Single
Dim BlueReferenceRadius As Single




If PVGridSize = 0 Then Exit Function



If ScreenDrawing.ImageGraphState(0).GraphType = "Flat" And DrawShapeType = "Circle" And (Not MedianFlat) Then 'PCN4370
    RedReferenceRadius = ((ExpectedDiameter / 2) * ((100 + Flat3dLimitR) / 100)) / PVDrawScreenRatio
    BlueReferenceRadius = ((ExpectedDiameter / 2) * ((100 + Flat3dLimitL) / 100)) / PVDrawScreenRatio
    



    PVScreenCntrl.BackColor = &HF0F0FF
    PVScreenCntrl.FillStyle = vbFSSolid
    
    PVScreenCntrl.FillColor = RGB(255, 255, 255)
    PVScreenCntrl.Circle (CentreLineX + ShapeCentreX, CentreLineY + ShapeCentreY), RedReferenceRadius, RGB(255, 255, 255)
    
    PVScreenCntrl.FillColor = RGB(240, 240, 255)
    PVScreenCntrl.Circle (CentreLineX + ShapeCentreX, CentreLineY + ShapeCentreY), BlueReferenceRadius, RGB(240, 240, 255)
Else
    PVScreenCntrl.BackColor = vbWhite
End If
PVGridLine = 1

PVGridXYPos1 = (PVGridSize * PVGridLine)


If ScreenDrawingType >= 1 Then
    GridColour = RGB(255, 255, 255)
Else
    GridColour = RGB(235, 235, 235)
End If

Do While MainScreen.width > PVGridXYPos1
    'Draw X axis grid lines
    Call RenderLine(PVScreenCntrl, CentreLineX - PVGridXYPos1, 0, CentreLineX - PVGridXYPos1, MainScreen.width, GridColour) '235
    Call RenderLine(PVScreenCntrl, CentreLineX + PVGridXYPos1, 0, CentreLineX + PVGridXYPos1, MainScreen.width, GridColour)
    'Draw Y axis grid lines
    Call RenderLine(PVScreenCntrl, 0, CentreLineY - PVGridXYPos1, MainScreen.width, CentreLineY - PVGridXYPos1, GridColour)
    Call RenderLine(PVScreenCntrl, 0, CentreLineY + PVGridXYPos1, MainScreen.width, CentreLineY + PVGridXYPos1, GridColour)
    PVGridLine = PVGridLine + 1
    PVGridXYPos1 = (PVGridSize * PVGridLine)
Loop
'Draw centre lines
Call RenderLine(PVScreenCntrl, CentreLineX, 0, CentreLineX, MainScreen.width, RGB(255, 200, 170)) '170
Call RenderLine(PVScreenCntrl, 0, CentreLineY, MainScreen.width, CentreLineY, RGB(255, 200, 170))



Exit Function
Err_Handler:
    MsgBox Err & "-CLS79:" & Error$

End Function

Sub DrawMouseDown(screenctrl As Control, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'DrawMouseDown function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    8/11/02     Building initial framework
'
'Description:
'
'
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

LastMouseMoveX = CurrentMouseMoveX
LastMouseMoveY = CurrentMouseMoveY
CurrentMouseMoveX = X
CurrentMouseMoveY = Y


If Button = 1 Then CLPScreenDrawAction = "LeftClick": LeftMouseDown = True
If Button = 2 Then CLPScreenDrawAction = "RightClick": RightMouseDown = True
If CLPScreenAction = "DrawLine" Then Call ScreenDrawing.DrawActionSingleLine(screenctrl)
If CLPScreenAction = "DrawPolyLine" Then Call ScreenDrawing.DrawActionPolyLine(screenctrl)
If CLPScreenAction = "DrawCalibrationLine" Then Call ScreenDrawing.DrawActionCalibrationLine(screenctrl)
If CLPScreenAction = "DrawOrthogonalLine" Then Call ScreenDrawing.DrawActionOrthogonalLine(screenctrl)
If CLPScreenAction = "MoveAll" Then Call ScreenDrawing.DrawActionMoveAll(screenctrl)
If CLPScreenAction = "DrawWaterLevel" Then Call ScreenDrawing.DrawActionWaterLevel(screenctrl)
If CLPScreenAction = "DrawText" Then Call ScreenDrawing.DrawActionDrawText(screenctrl)
If CLPScreenAction = "DrawMaskBox" Then Call ScreenDrawing.DrawActionMaskBox(screenctrl)
If CLPScreenAction = "DrawHorCalibrationLine" Then Call ScreenDrawing.DrawActionHorCalibrationLine(screenctrl)
If CLPScreenAction = "DrawDimension" Then Call ScreenDrawing.DrawActionDimension(screenctrl)
If CLPScreenAction = "DrawIntegratedArea" Then Call ScreenDrawing.DrawActionIntegratedArea(screenctrl)
    
Exit Sub
Err_Handler:
Select Case Err
    Case 53 ' Could not find the mouse icon
        Resume Next
    Case Else
    MsgBox Err & "-CLS80:" & Error$

End Select
End Sub

Sub DrawMouseUp(screenctrl As Control, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error GoTo Err_Handler

If Button = 1 Then CLPScreenDrawAction = "LeftUp": LeftMouseDown = False
If Button = 2 Then CLPScreenDrawAction = "RightUp": RightMouseDown = False

If CLPScreenAction = "MoveAll" Then Call ScreenDrawing.DrawActionMoveAll(screenctrl)
If CLPScreenAction = "DrawOrthogonalLine" Then Call ScreenDrawing.DrawActionOrthogonalLine(screenctrl)
If CLPScreenAction = "DrawDimension" Then Call ScreenDrawing.DrawActionDimension(screenctrl)


Exit Sub
Err_Handler:
Select Case Err
    Case 53 ' Could not find the mouse icon
        Resume Next
    Case Else
    MsgBox Err & "-CLS81:" & Error$

End Select
End Sub



Sub DrawMouseMove(screenctrl As Control, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error GoTo Err_Handler

    LastMouseMoveX = CurrentMouseMoveX
    LastMouseMoveY = CurrentMouseMoveY
    CurrentMouseMoveX = X
    CurrentMouseMoveY = Y
    
    CLPScreenDrawAction = "Move"
    If CLPScreenAction = "DrawLine" Then Call ScreenDrawing.DrawActionSingleLine(screenctrl)
    If CLPScreenAction = "DrawPolyLine" Then Call ScreenDrawing.DrawActionPolyLine(screenctrl)
    If CLPScreenAction = "DrawCalibrationLine" Then Call ScreenDrawing.DrawActionCalibrationLine(screenctrl)
    If CLPScreenAction = "DrawOrthogonalLine" Then Call ScreenDrawing.DrawActionOrthogonalLine(screenctrl)
    If CLPScreenAction = "MoveAll" Then Call ScreenDrawing.DrawActionMoveAll(screenctrl)
    If CLPScreenAction = "DrawWaterLevel" Then Call ScreenDrawing.DrawActionWaterLevel(screenctrl)
    If CLPScreenAction = "DrawMaskBox" Then Call ScreenDrawing.DrawActionMaskBox(screenctrl)
    If CLPScreenAction = "DrawHorCalibrationLine" Then Call ScreenDrawing.DrawActionHorCalibrationLine(screenctrl)
    If CLPScreenAction = "DrawDimension" Then Call ScreenDrawing.DrawActionDimension(screenctrl)
    If CLPScreenAction = "DrawIntegratedArea" Then Call ScreenDrawing.DrawActionIntegratedArea(screenctrl)
    
    
    

Exit Sub
Exit Sub
Err_Handler:
Select Case Err
    Case 11 ' Divide by zero
        MsgBox Err & " - " & Error$
    Case 5 '
       'MsgBox Err & "Temporary message after mousemove. Please check your drawing is as intended."
    Case 6 'overflow
    Case 53 'Could not find the mouse icon
    Case Else
        MsgBox Err & "-CLS82:" & Error$
End Select
End Sub



'*****************************************************************************************
'*****************************************************************************************
'*****************************************************************************************




Function AutoMoveShapes(screenctrl As Control, MoveOffsetX, MoveOffsetY)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'AutoMoveShapes Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    28/11/02     Building initial framework
'
'Description:
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Exit Function
Err_Handler:
Select Case Err

    Case Else
        MsgBox Err & "-CLS83:" & Error$
End Select
End Function


Function InitVideoSlider()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'InitVideoSlider function  Louise Shrimpton louiseS@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    24/01/03     Building initial framework
'
'Description:
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    
    GetAVITime 'This populates the AVITime time, which is the length of the video
    If mediatype = Video Then
        VideoFrameSlider.Max = (AVITime) * 100 'PCN2930
    Else
        VideoFrameSlider.Max = 10000 'PCN2930
    End If
    SliderFrame = 0
    VideoFrameSlider.value = SliderFrame 'PCN2930
    Call VideoFrameSliderMove 'PCN2930
    'vvvv PCN2782 *********************
'    AVITimeVar = 0
'    Call FormatTime
    '^^^^ *****************************

Exit Function
Err_Handler:
    MsgBox Err & "-CLS84:" & Error$
End Function

Public Sub MoveSlider(Time As Double)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'MoveSlider  Louise Shrimpton louiseS@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    24/01/03     Building initial framework
'
'Description:
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

If mediatype = Video Then
    If AVITime = 0 Then Exit Sub 'PCN????
    VideoFrameSlider.value = (Time / AVITime) * VideoFrameSlider.Max 'PCN2930
    SliderFrame = VideoFrameSlider.value 'PCN2930
    Call VideoFrameSliderMove 'PCN2930
End If

Exit Sub
Err_Handler:
Select Case Err
    Case 6: Resume Next
    Case Else: MsgBox Err & "-CLS85:" & Error$
End Select
End Sub




Private Sub FormatTime()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'FormatTime  Louise Shrimpton louiseS@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    04/02/03     Building initial framework
'
'Description:  Formats the time in hours, minutes and seconds
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

    Dim sec As Long
    Dim Min As Long
    Dim hr As Long
    sec = AVITime Mod 60
    Min = AVITime \ 60
    hr = Min \ 60
    Min = Min Mod 60
    Dim Total, MinStr, SecStr, Current As String
    If sec < 10 Then
        SecStr = "0" & sec
    Else
        SecStr = sec
    End If
    If Min < 10 Then
        MinStr = "0" & Min
    Else
        MinStr = Min
    End If
    If hr = 0 Then
        Total = Min & ":" & SecStr
    Else
        Total = hr & ":" & MinStr & ":" & SecStr
    End If
    
    If VideoFrameSlider.Max = 0 Then
        AVITimeVar = 0
    Else
        AVITimeVar = Round((VideoFrameSlider.value / VideoFrameSlider.Max) * AVITime, 1)
    End If
    
    sec = AVITimeVar Mod 60
    Min = AVITimeVar \ 60
    hr = Min \ 60
    Min = Min Mod 60
    If sec < 10 Then
        SecStr = "0" & sec
    Else
        SecStr = sec
    End If
    If Min < 10 Then
        MinStr = "0" & Min
    Else
        MinStr = Min
    End If
    If hr = 0 Then
        Current = Min & ":" & SecStr
    Else
        Current = hr & ":" & MinStr & ":" & SecStr
    End If
    AVITimeLabel.Caption = Current & " / " & Total
Exit Sub
Err_Handler:
    Select Case Err
        Case 6: Resume Next
        Case Else: MsgBox Err & "-CLS86:" & Error$
    End Select
End Sub





Private Sub ThreeDtimer_Timer()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : ThreeDtimer_Timer()
'Created : 22 September 2003, PCN2240
'Updated :
'Prg By  : Louise Shrimpton, Michelle Lindsay
'Param   : (None)
'Desc    :
'
'Usage   :
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    Call d3d_refresh
    If zoomspeed < 0 And ButtonNumber = 10 Then
        Call d3d_zoom_speed(zoomspeed)
        zoomspeed = zoomspeed - 1
        If zoomspeed < -100 Then
            zoomspeed = -100
        End If
    ElseIf zoomspeed > 0 And ButtonNumber = 11 Then
        Call d3d_zoom_speed(zoomspeed)
        zoomspeed = zoomspeed + 1
        If zoomspeed > 100 Then
            zoomspeed = 100
        End If
    End If
    If laserspeed <> 0 Then
        'Check that PVFrameNo is not greater than total number of PVFrames
        If PVFrameNo + laserspeed < PVDataNoOfLines And (PVFrameNo + laserspeed) > 0 Then
            PVFrameNo = PVFrameNo + laserspeed
            Call d3d_laser_focus(PVFrameNo)
            'Move PVGraph marker
            PrecisionVisionGraph.MoveMarker (PVFrameNo)
        ElseIf PVFrameNo < PVDataNoOfLines And laserspeed > 0 Then
            PVFrameNo = PVDataNoOfLines
            Call d3d_laser_focus(PVFrameNo)
            'Move PVGraph marker
            PrecisionVisionGraph.MoveMarker (PVFrameNo)
        ElseIf PVFrameNo > 1 And laserspeed < 0 Then
            PVFrameNo = 1
            Call d3d_laser_focus(PVFrameNo)
            'Move PVGraph marker
            PrecisionVisionGraph.MoveMarker (PVFrameNo)
        End If
    End If
    'vvvv PCN4171 *****************************
    If ButtonNumber = 12 Then
        Call d3d_rotate_pipe_z(5) 'RotateXAntiClock
    ElseIf ButtonNumber = 13 Then
        Call d3d_rotate_pipe_z(-5) 'RotateXClockwise
    ElseIf ButtonNumber = 14 Then
        Call d3d_rotate_pipe_y(5) 'RotateYAntiClock
    ElseIf ButtonNumber = 15 Then
        Call d3d_rotate_pipe_y(-5) 'RotateYClockwise
    End If
    '^^^^ *************************************

Exit Sub
Err_Handler:
   MsgBox Err & "-CLS87:" & Error$
End Sub




Private Sub VideoFrameSliderBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
Dim MarkerCentreOffset As Integer 'Defines the offset required to ensure marker is aligned to the correct position on the slider.
Dim ErrorStr As String
Dim NewLeftPos As String

If VideoSliderMarker.Visible Then 'PCN4237
    If CLPScreenAction = "DrawMaskBox" Or _
       CLPScreenAction = "DrawWaterLevel" Or _
       CLPScreenAction = "DrawHorCalibrationLine" Or _
       CLPScreenAction = "DrawCalibrationLine" Then
        Call ClearLineScreen.TakeASnapShot
        Call ControlsScreen.ControlsViewSetup
    Else
        Call VideoScreenActivate(ErrorStr)
    End If
    
    MarkerCentreOffset = 30
    
    NewLeftPos = (X - MarkerCentreOffset)
    If NewLeftPos >= VideoFrameSlider.FrameLeft And NewLeftPos <= (VideoFrameSlider.FrameLeft + VideoFrameSlider.Framewidth - MarkerCentreOffset) Then
        VideoSliderMarker.Left = NewLeftPos
        VideoFrameSlider.MarkerPosition = VideoSliderMarker.Left + MarkerCentreOffset
        VideoFrameSlider.value = (VideoFrameSlider.Max - VideoFrameSlider.Min) * ((VideoFrameSlider.MarkerPosition - VideoFrameSlider.FrameLeft) / VideoFrameSlider.Framewidth) 'PCN2955
        
        DoEvents
        Call VideoFrameSliderVideoScroll
    End If
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS88:" & Error$
End Sub

Private Sub VideoRecordMarkerStartAdjuster_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : VideoRecordMarkerStartAdjuster_MouseDown
'Created : 7 July 2004, PCN2930
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : The Start video marker is moved left or right on the Video Frame Slider by the Start Adjuster.
'          To reset the move direction flag.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

VideoMarkerMoveDir = ""
VideoMarkerLastX = X

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS89:" & Error$
End Sub

Private Sub VideoRecordMarkerStartAdjuster_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : VideoRecordMarkerStartAdjuster_MouseDown
'Created : 7 July 2004, PCN2930
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : The Start video marker is moved left or right on the Video Frame Slider by the Start Adjuster.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim NextMarkerPosition As Long

If Button = 1 And PVRecording <> True And PVDFileName = "" Then 'Left mouse press 'Disable when recording
    'Adjust Start marker
    If VideoMarkerLastX <> 0 And VideoMarkerMoveDir = "" Then
        'Must await for Marker to move into new position
        VideoMarkerLastX = X
        VideoMarkerMoveDir = "B" 'In both directions
        Exit Sub
    End If

    'Reset the start distance if not automatic distance
    If DistanceMethod <> "AutomaticCounter" Then
        ConfigInfo.DistanceStart = InvalidData
    End If
    
    'Adjust markers
    NextMarkerPosition = VideoRecordMarkerStart.x1 + (X - VideoMarkerLastX)
    If NextMarkerPosition <= VideoFrameSlider.FrameLeft Then
        VideoRecordMarkerStartAdjuster.Left = VideoFrameSlider.FrameLeft - VideoRecordMarkerStartAdjuster.width
        VideoRecordMarkerStart.x1 = VideoFrameSlider.FrameLeft
    ElseIf NextMarkerPosition <= VideoRecordMarkerStop.x1 Then
        VideoRecordMarkerStartAdjuster.Left = VideoRecordMarkerStartAdjuster.Left + (X - VideoMarkerLastX)
        VideoRecordMarkerStart.x1 = NextMarkerPosition
    Else
        VideoRecordMarkerStartAdjuster.Left = VideoRecordMarkerStop.x1 - VideoRecordMarkerStartAdjuster.width
        VideoRecordMarkerStart.x1 = VideoRecordMarkerStop.x1
    End If
    VideoRecordMarkerStart.x2 = VideoRecordMarkerStart.x1
    With VideoFrameSlider
    .MarkerStart = VideoRecordMarkerStart.x1 - .FrameLeft 'Set Record start position
    VideoMarkerLastX = 0 'Reset position
    End With
    
    DoEvents
    Call VideoRecordMarkerStartAdjuster_MouseUp(Button, Shift, X, Y)
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS90:" & Error$
End Sub

Private Sub VideoRecordMarkerStartAdjuster_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : VideoRecordMarkerStartAdjuster_MouseUp
'Created : 7 July 2004, PCN2930
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Redraw the slider
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

If PVRecording <> True And PVDFileName = "" Then  'Disable when and after recording
    Call VideoFrameSliderRedraw
    With VideoFrameSlider
    .value = (.Max - .Min) * (.MarkerStart / .Framewidth)  'Assume .Framewidth <> 0  'PCN2955
    End With
    Call GotoStartMarker 'VideoFrameSliderSetForStartRecord
    'LastDataTime = CurrentAVITime 'PCNANT????
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS91:" & Error$
End Sub

Private Sub VideoRecordMarkerStopAdjuster_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : VideoRecordMarkerStopAdjuster_MouseDown
'Created : 7 July 2004, PCN2930
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : The Stop video marker is moved left or right on the Video Frame Slider by the Start Adjuster.
'          To reset the move direction flag.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

VideoMarkerMoveDir = ""
VideoMarkerLastX = X
Call ProfilerPause

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS92:" & Error$
End Sub

Private Sub VideoRecordMarkerStopAdjuster_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : VideoRecordMarkerStopAdjuster_MouseDown
'Created : 7 July 2004, PCN2930
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : The Stop video marker is moved left or right on the Video Frame Slider by the Start Adjuster.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim NextMarkerPosition As Long
'Debug.Print "MouseMove " & X & Y
If Button = 1 And PVRecording <> True And PVDFileName = "" Then   'Left mouse press 'Disable when recording
    'Adjust Stop marker
    If VideoMarkerLastX <> 0 And VideoMarkerMoveDir = "" Then
        'Must await for Marker to move into new position
        VideoMarkerLastX = X
        VideoMarkerMoveDir = "B" 'In both directions
        Exit Sub
    End If

    'Reset the finish distance if not automatic distance
    If DistanceMethod <> "AutomaticCounter" Then
        ConfigInfo.DistanceFinish = InvalidData
    End If
    
    'Adjust markers
    NextMarkerPosition = VideoRecordMarkerStop.x1 + (X - VideoMarkerLastX)
    If NextMarkerPosition >= (VideoFrameSlider.FrameLeft + VideoFrameSlider.Framewidth) Then
        VideoRecordMarkerStopAdjuster.Left = VideoFrameSlider.FrameLeft + VideoFrameSlider.Framewidth + 1
        VideoRecordMarkerStop.x1 = VideoFrameSlider.FrameLeft + VideoFrameSlider.Framewidth
    ElseIf NextMarkerPosition >= VideoRecordMarkerStart.x1 Then
        VideoRecordMarkerStopAdjuster.Left = VideoRecordMarkerStopAdjuster.Left + (X - VideoMarkerLastX)
        VideoRecordMarkerStop.x1 = NextMarkerPosition
    Else
        VideoRecordMarkerStopAdjuster.Left = VideoRecordMarkerStart.x1 + 1
        VideoRecordMarkerStop.x1 = VideoRecordMarkerStart.x1
    End If
    VideoRecordMarkerStop.x2 = VideoRecordMarkerStop.x1
    
    With VideoFrameSlider
    .MarkerStop = VideoRecordMarkerStop.x1 - .FrameLeft  'Set Record stop position
    VideoMarkerLastX = 0 'Reset position
    End With
    DoEvents
    Call VideoRecordMarkerStopAdjuster_MouseUp(Button, Shift, X, Y)

    
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS93:" & Error$
End Sub

Private Sub VideoRecordMarkerStopAdjuster_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : VideoRecordMarkerStopAdjuster_MouseUp
'Created : 7 July 2004, PCN2930
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Redraw the slider
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

If PVRecording <> True And PVDFileName = "" Then  'Disable when and after recording
    Call VideoFrameSliderRedraw
    With VideoFrameSlider
    .value = (.Max - .Min) * (.MarkerStop / .Framewidth)  'Assume .Framewidth <> 0  'PCN2955
    End With
    Call GotoStopMarker 'VideoFrameSliderSetForStartRecord
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS94:" & Error$
End Sub

' PCNant????


Private Sub VideoScreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

If SonarMeasure = "Measure" Then
    If Button = 1 Then
        SonarX1 = X
        SonarY1 = Y
        SonarMeasure = "Start"
    Else
        SonarX2 = X
        SonarY2 = Y
        SonarMeasure = "Finish"
    End If
End If

        Call hough_setlaserpoint(CLng(X / 15 / Me.VideoScreen.width * ConfigInfo.MediaWidth), _
                                 CLng(ConfigInfo.MediaHeight - (Y / 15 / Me.VideoScreen.height * ConfigInfo.MediaHeight)))
        Call Me.RefreshVideoScreen
        Call DrawMainScaleRedLines(X, Y)
        
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS95:" & Error$
End Sub

Public Sub DrawSonarMeasureLine()
'****************************************************************************************
'Name    : DrawSonarMeasureLine
'Created : 9 Jan 2007, PCN4419
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    :
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
'Dim Measurement As Double
'Dim W As Long
'Dim H As Long
'
'W = SonarX1 - SonarX2
'H = SonarY1 - SonarY2
'
'Set ClearLineScreen.VideoScreen.Picture = SonarConfig.picCalibration.Picture
'VideoScreen.Line (SonarX1, SonarY1)-(SonarX2, SonarY2), vbWhite
'VideoScreen.Line (SonarX1 - 5, SonarY1 - 5)-(SonarX1 + 5, SonarY1 + 5), vbWhite, B
'VideoScreen.Line (SonarX2 - 5, SonarY2 - 5)-(SonarX2 + 5, SonarY2 + 5), vbWhite, B
'
'Measurement = Sqr((W * W) + (H * H)) * SonarConfig.mmPerPixels
'SonarConfig.lblMeasurement = "Distance: " & Round(Measurement, 2)

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS96:" & Error$
End Sub

Private Sub VideoScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'****************************************************************************************
'Name    : VideoScreen_MouseUp
'Created : 9 Jan 2007, PCN4419
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    :
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler

If SonarMeasure = "Start" Then
    SonarX1 = X
    SonarY1 = Y
    Call DrawSonarMeasureLine
ElseIf SonarMeasure = "Finish" Then
    SonarX2 = X
    SonarY2 = Y
    Call DrawSonarMeasureLine
ElseIf SonarMeasure = "Measure" Then
'    ClearLineScreen.MousePointer = vbCrosshair
'    ClearLineScreen.VideoScreen.MousePointer = vbCrosshair
    Screen.MousePointer = vbCrosshair
'Else
'    ClearLineScreen.MousePointer = vbHourglass
'    ClearLineScreen.VideoScreen.MousePointer = vbHourglass
    Screen.MousePointer = vbNormal
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS97:" & Error$
End Sub


Private Sub VideoScreen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

If SonarMeasure = "Start" Or SonarMeasure = "Finish" Then
    SonarMeasure = "Measure"
End If
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS98:" & Error$
End Sub

Private Sub VideoSliderMarker_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : VideoSliderMarker_MouseDown
'Created : 7 July 2004, PCN2930
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : The Video Slider marker is moved left or right on the Video Frame Slider by the Start Adjuster.
'          To reset the move direction flag.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim ErrorStr As String

VideoMarkerMoveDir = ""
VideoMarkerLastX = X

Call VideoScreenActivate(ErrorStr)

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS99:" & Error$
End Sub

Private Sub VideoSliderMarker_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : VideoSliderMarker_MouseMove
'Created : 7 July 2004, PCN2930
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : The Video Slider marker is moved left or right on the Video Frame Slider by the Start Adjuster.
'
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

If mediatype <> Video Then Exit Sub

If Button = 1 And PVRecording <> True Then  'Left mouse press 'Disable when recording
    'Adjust Start marker
    If VideoMarkerLastX <> 0 And VideoMarkerMoveDir = "" Then
        'Must await for Marker to move into new position
        VideoMarkerLastX = X
        VideoMarkerMoveDir = "B" 'In both directions
        Exit Sub
    End If

    'Adjust markers
    Dim MarkerCentreOffset As Integer 'Defines the offset required to ensure marker is aligned to the correct position on the slider.
    MarkerCentreOffset = 30
    If VideoSliderMarker.Left + (X - VideoMarkerLastX) <= (VideoFrameSlider.FrameLeft - MarkerCentreOffset) Then
        VideoSliderMarker.Left = VideoFrameSlider.FrameLeft - MarkerCentreOffset
    ElseIf VideoSliderMarker.Left + (X - VideoMarkerLastX) >= (VideoFrameSlider.FrameLeft + VideoFrameSlider.Framewidth - MarkerCentreOffset) Then
        VideoSliderMarker.Left = (VideoFrameSlider.FrameLeft + VideoFrameSlider.Framewidth - MarkerCentreOffset)
    Else
        VideoSliderMarker.Left = VideoSliderMarker.Left + (X - VideoMarkerLastX)
    End If
    
    VideoMarkerLastX = 0 'Reset position

    VideoFrameSlider.MarkerPosition = VideoSliderMarker.Left + MarkerCentreOffset
    VideoFrameSlider.value = (VideoFrameSlider.Max - VideoFrameSlider.Min) * ((VideoFrameSlider.MarkerPosition - VideoFrameSlider.FrameLeft) / VideoFrameSlider.Framewidth) 'PCN2955

    DoEvents
    Call VideoFrameSliderVideoScroll
    If IPD = True Then DistanceCounterTextUpdate 'PCN3744
    

End If

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS100:" & Error$
End Sub

Private Sub VideoSliderMarker_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : VideoSliderMarker_MouseDown
'Created : 13 September 2004, PCN2930
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : After the video slider bar is moved, goto the correct PVFrameNo is exists.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler



If CLPScreenAction = "DrawMaskBox" Or _
   CLPScreenAction = "DrawWaterLevel" Or _
   CLPScreenAction = "DrawHorCalibrationLine" Or _
   CLPScreenAction = "DrawCalibrationLine" Then
    Call ClearLineScreen.TakeASnapShot
    Call ControlsScreen.ControlsViewSetup
    Exit Sub
   
End If
   
If CLPScreenMode = PV Then Call GotoPVProfile(PVFrameNo, True)


Exit Sub
Err_Handler:
    MsgBox Err & "-CLS101:" & Error$
End Sub

Private Sub VideoTimer_Timer()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : VideoTimer_Timer
'Created : 20 June 2004, PCN2891
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : This is the timer that activates getting the frame data from the C code.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler



'    'Check if the PVDFile for the closest PVFrameNo for the current Play time' 'PCNGL170103
'    If PVFrameNo <= PVDataNoOfLines And AVIPlay = True And PVRecording = False Then 'PCNLS040203
'        Call GotoPVProfileForCurrentTime 'PCN2930
'
'
'    ElseIf AVIPlay = True And PVRecording = False Then  'PCN2966 also forgot to port over from 5.4.2
'        Call DistanceCounterTextUpdate 'PCN2966 also forgot to port over from 5.4.2
'    'vvvv PCN2970 ************************************
'    ElseIf AVIPlay = False And PVRecording = False And PVPlay = True Then
'        If PVFrameNo < PVDataNoOfLines Then
'            PVFrameNo = PVFrameNo + 1
'            Call GotoPVProfile(PVFrameNo, False)
'            Call DistanceCounterTextUpdate 'PCN2966
'        End If
'    '^^^^ ********************************************
'    End If
'End If
'
'    If DistanceMethod = "AutomaticCounter" Then
'        If Distance.CalculatedDist.Tag = False Then
'            Call DistanceCounterTextUpdate 'PCN2891 'PCN2930  'PCN2928
'        End If
'    End If

Exit Sub
Err_Handler:
Select Case Err
    Case Else: MsgBox Err & "-CLS102:" & Error$
End Select

End Sub
Public Sub VideoRecordingLoop()
On Error GoTo Err_Handler

Dim NumFrames As Long 'Numbero Of Frames In CPP Recording Buffer As Long

If PVRecording = True And IsCPPVideoRunning Then   'PCN1863 'PCN2930
    Call gettotalnumframes(NumFrames)
    If NumFrames > 400 Then
        Call setrecprofstat(0)
        Call ProfilerPause 'CPP not to record in the GetBlockOfPVData method.
        ConfigInfo.ProfileRecordingMethod = "XY"
        Call GetBlockOfPVData
        
    End If
End If
    
Call DistanceCounterTextUpdate 'PCN2966 When Playing distance counter to be updated every frame (10 Aug 2004)
Call MoveSlider(GetCurrentAVITime)

If Not IsCPPVideoRunning And ProfilerTimer.Enabled = True Then Call videorun 'Before making the CPP video run make
                                                                             'sure its suppose to be by checking timer enabled

Exit Sub
Err_Handler:
Select Case Err
    Case Else: MsgBox Err & "-CLS103:" & Error$
End Select
End Sub

Public Sub VideoPlaybackLoop()
On Error GoTo Err_Handler
Dim PVFrame As Long

    If Not IsCPPVideoRunning And ProfilerTimer.Enabled = True Then Call videorun 'Before making the CPP video run make
                                                                                 'sure its suppose to be by checking timer enabled
    PVFrameNo = GetPVFrameNoFromTime(GetCurrentAVITime)
    PVScaleMarkerStFrame = PVFrameNo
    PVScaleMarkerFnFrame = PVFrameNo
    Call PrecisionVisionGraph.RepositionYScaleMarkers
    Call PVGraphsKeyForm.PVGraphsKeyUpdate 'PCN2990
    Call MoveSlider(GetCurrentAVITime)

Exit Sub
Err_Handler:
Select Case Err
    Case Else: MsgBox Err & "-CLS104:" & Error$
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Form_unload function  Louise Shrimpton louiseS@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    16/12/02     Building initial framework
'
'Description:  Makes sure the program won't crash - need to stop the c code running
'
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    Call ProfilerPause
    If mediatype <> "" Then
        DoEvents 'PCN2877 speed up uninitialise (10 June 2004, Antony van Iersel) don't really no if it helps. seems to.
        If mediatype = "Live" Then
            Call UnitLiveFeed
        ElseIf mediatype = Video Then
            Call uninitialise
        End If
    End If
    If ThreeDRunning = True Then
        Unload3D
    End If
'    Call destroythreed
    
'Message to remind user to save the .pvd file if required PCN1895
Dim Ans As Integer
If (PVDSaved = False And LastRecordedFrame > 0 And Registered = True) Then 'ML260303
    'Ans = MsgBox(DisplayMessage("The .pvd file is unsaved, do you wish to save this file?"), vbYesNo + vbExclamation)  'PCN2111
    ProfilerMessageBox.MsgBoxYesNo (DisplayMessage("The .pvd file is unsaved, do you wish to save this file?"))
    Ans = PMBAnswer
    If Ans = vbYes Then
        Call SaveImageAndOrData
    End If
End If
        

'Message to remind unregistered user to register the product
If Registered = False And SoftwareConfiguration <> "Reader" Then 'ML170203 'PCN2212 'PCN3809
    'MsgBox DisplayMessage("Please remember to register the ClearLine Profiler software."), vbInformation 'PCN2111
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Please remember to register the ClearLine Profiler software."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
End If

'Reset Regional Number settings - Richard 23/03/10
Call StartUp.SetRegionalSettings(LOCALE_STHOUSAND, RegThousandSeperator)
Call StartUp.SetRegionalSettings(LOCALE_SDECIMAL, RegDecSymbol)

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS105:" & Error$
End Sub


Public Function SetupMainScreenForVideo() 'PCNGL301202
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'SetupMainScreenForVideo function  Geoff Logan geofflogan@cbsys.co.nz 'PCNGL2401032
'
'Revision history"
'   V0.0    Geoff Logan,    30/12/02     Building initial framework
'
'Description:
'               Setup the MainScreen for Video (including AVI,MPG or live video)
'               Can be called by the file open function or by pressing screen
'               configuration button.
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

VideoScreen.Visible = True
PVScreen.Visible = False 'PCNGL030103
'VideoSnapShotMode = Video 'PCN4043
CLPScreenMode = Video
Call ControlsScreen.ControlsViewSetup

PVDimensionLabel.Visible = False
'PVScreen.MousePointer = 0 PCNGL030103
SnapShotScreen.Visible = False
'ConfigToolBar1.Buttons.Item(1).Image = 1 '2 'PCNLS 1988
'ConfigToolBar1.Buttons.Item(1).ToolTipText = SnapShot
'VideoTimer.Enabled = True 'PCNGL291202
VideoScreen.Visible = True
VideoScreen.ZOrder 0 'PCNGL291202
'AVIPlay = False 'PCNGL261202
'PCNGL291202 Call initialise(App.hInstance, VideoScreen.hwnd, Filename) '"z:\\Louise_Testing_19_12_02\\houghlib\\test2.avi") '"z:\\Louise_Testing_19_12_02\\houghlib\\test3.avi") 'Filename) '"z:\\Louise_Testing_19_12_02\\houghlib\\test2.avi") '"z:\\Louise_Testing_19_12_02\\houghlib\\test3.avi")

'Call SetupMTButtonsForVideo 'PCNGL300103
Call ControlsScreen.ControlsViewSetup 'PCN4171

'PVGraphsKeyForm.Visible = False 'PCN2990 'PCN2818


Exit Function
Err_Handler:
    MsgBox Err & "-CLS106:" & Error$
End Function

Public Function SetupMainScreenForPV_Old() 'PCNGL301202
End Function

Public Function SetupMainScreenForPV() 'PCNGL301202 'PCN4171
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'SetupMainScreenForPV function  Geoff Logan geofflogan@cbsys.co.nz 'PCNGL2401032
'
'Revision history"
'   V0.0    Geoff Logan,    30/12/02     Building initial framework
'
'Description:
'               Setup the MainScreen for Precision Vision
'               Can be called by the file open function or by pressing screen
'               configuration button.
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

'CLPScreenMode = PV 'PCN1863
CLPScreenMode = PV 'PCN4171
Call ControlsScreen.ControlsViewSetup 'PCN4171

PVScreen.MousePointer = 99
PVScreen.MouseIcon = LoadResPicture(MainScreenMouseIcon, vbResIcon) 'PCNGL291102
PVFrameStatus.Visible = True
PVScreen.Visible = True
SnapShotScreen.Visible = False

PVScreen.ZOrder 0  'PCNGL261202
ClearLineScreen.PVScreen.Cls
Call DrawPVProfile_Setup(ClearLineScreen.PVScreen) 'PCN1858 'PCN1863 PCN3526
Call DrawMainScale(ClearLineScreen.PVScreen) 'PCN1858
Call DrawPVGrid(ClearLineScreen.PVScreen, MainScaleGrid / 8, CentreLineX, CentreLineY)

Call PVGraphsKeyForm.DisplayPVGraphsKey 'PCN2990 'PCN2818
Call SetDimenResultsSize(True) 'PCN4171
DimenResults.ZOrder 0
AreaResults.ZOrder 0

Exit Function
Err_Handler:
    MsgBox Err & "-CLS107:" & Error$
End Function


Public Function ChangePriority(dwPriorityClass As PRIORITY_CLASS) As Boolean

'USAGE: Change the Priority of the Current Process
'(i.e., the Process that represents the VB app
'that calls this function)

'PARAMETER: dwPriorityClass: Value as
'defined by the Enum PRIORITY_CLASS in
'declarations

'RETURNS: TRUE IF SUCCESSFUL, FALSE OTHEWISE

   Dim hProcess As Long
   Dim ret As Long, PID As Long
   Dim dw As Long
   PID = GetCurrentProcessId()
   ' Get the Handle of the Process
   hProcess = OpenProcess(PROCESS_DUP_HANDLE, True, PID)
    
   If hProcess = 0 Then Exit Function
 
   ' Change the priority
   ret = SetPriorityClass(hProcess, dwPriorityClass)
   If ret = 0 Then
        dw = Err.LastDllError
    End If


   ' Close the process
   Call CloseHandle(hProcess)

   ChangePriority = ret <> 0

End Function
        
Public Function InitVideo() 'PCNGL301202
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'InitVideo function  Louise Shrimpton louises@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    09/01/03     Building initial framework
'
'Description:
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim Status As Long
Status = 0
    Dim NoVideoSync As Long
    Call ProfilerPause
    Call CheckAVIInitialised
    If mediatype = Video Then
         VideoScreen.AutoRedraw = True
         VideoScreen.Appearance = 0 'PCN2475
         VideoScreen.AutoSize = False
         VideoScreen.BorderStyle = 0 'PCN2475
         VideoScreen.CausesValidation = True
         VideoScreen.ClipControls = True
         VideoScreen.ScaleMode = 1
         
         
         VideoScreen.Cls
        '!!!!!! MUST DISABLE VIDEO TIMER FIRST !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        DoEvents
        If PVDFileName <> "" Then
            NoVideoSync = 0
        Else
            NoVideoSync = 1
        End If
        '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
         Call initialise(App.hInstance, VideoScreen.hwnd, mediatype, VideoFileName, ConfigInfo.MediaWidth, ConfigInfo.MediaHeight, XAdjust, YAdjust, MediaOriginalHeight, MediaOriginalWidth, VideoCaptureDevice, ConfigInfo.FishEyeOriginalWidth, ConfigInfo.FishEyeOriginalHeight, NoVideoSync) 'PCNGL140103 'PCN2289 'PCN2395 capture device from string to long
         'VideoScreen.Refresh
         

         
         
         
      
'
'        End Select
         'This doesn't work.  Need to figure out how to set our process's security features so
 'that we have access to changing the priority. At the moment we don't.
        'If ChangePriority(HIGH_PRIORITY_CLASS) = False Then
        'End If
        'vvvv Setup the VideoScreen for the new Video resolution ****************** 'PCNGL280103
        If VideoFileName = "Error" Then
        End If
        LightInPipe = False
        UserDefinedLight = False
        If ConfigInfo.MediaWidth > 0 And ConfigInfo.MediaHeight > 0 Then
            VideoAspectRatio = ConfigInfo.MediaHeight / ConfigInfo.MediaWidth
        Else
            VideoAspectRatio = VideoAspectRatio768x576 'As default
        End If
        'Setup Video screen
        Call VideoScreenSetupAspectRatio
        'Tell the C code to resize its image drawing
        Call ResizeVideoWindow
        '^^^^ *********************************************************************
        Call setrecprofstat(Status)
    End If
''    ClearLineScreen.ControlToolbar.Buttons.Item(9).Image = 25 'PCN2681
    ClearLineScreen.InitVideoSlider
    If TuningStyle = "Automatic" Then
       ' Call AutoTune.LoadFileAutoTune
    End If

    Call RefreshVideoScreen 'PCN4323 one customer video is blank when loaded, hopefully refreshing the
                            'the screen fixes it, because after loaded and play or seek is done it works
                            'fine.
    
Exit Function
Err_Handler:
    MsgBox Err & "-CLS108:" & Error$
End Function

Public Sub InitImageProcessing() 'PCNGL301202
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCN3194
'Name:      InitImageProcessing
'Developer: Antony van Iersel
'Date:      19 August 2005
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    Dim sa As SAFEARRAY2D, bmp As BITMAP    'Array that is manipulated and then coppied to the passed image
    Dim pict() As Byte  'Pointer that is passed to the different drawnig subfunctions, so they know where to draw
    
    If mediatype = StillImage Then
        ConfigInfo.MediaWidth = DummyScreen.width
        ConfigInfo.MediaHeight = DummyScreen.height
             
         Call initialise(App.hInstance, _
                         VideoScreen.hwnd, _
                         mediatype, _
                         ImageFileName, _
                         ConfigInfo.MediaWidth, _
                         ConfigInfo.MediaHeight, _
                         XAdjust, _
                         YAdjust, _
                         MediaOriginalHeight, _
                         MediaOriginalWidth, _
                         VideoCaptureDevice, _
                         ConfigInfo.FishEyeOriginalWidth, _
                         ConfigInfo.FishEyeOriginalHeight, 0) 'PCNGL140103 'PCN2289 'PCN2395 capture device from string to long
         
        VideoAspectRatio = ConfigInfo.MediaHeight / ConfigInfo.MediaWidth
        Call VideoScreenSetupAspectRatio
    
        'Get the imformation of the image that it is passed to it, to see if its compatible for drawing on
        GetObjectAPI ClearLineScreen.DummyScreen.Picture, Len(bmp), bmp
        
        
        'The compatability is only colour depth, in this case the image has to be 24bit colour'''
        If bmp.bmBitsPixel <> 24 Then   ' if not then exit                                      '
            MsgBox " 24-bit bitmaps only", vbCritical                                           '
            Exit Sub                                                                            '
        End If                                                                                  '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
        ' have the local matrix array to point to the passed image bitmap
        With sa                                                         '
            .cbElements = 1                                             '
            .cDims = 2                                                  '
            .Bounds(0).lLbound = 0                                      '
            .Bounds(0).cElements = bmp.bmHeight                         '
            .Bounds(1).lLbound = 0                                      '
            .Bounds(1).cElements = bmp.bmWidthBytes                     '
            .pvData = bmp.bmBits                                        '
        End With                                                        '
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
        CopyMemory ByVal VarPtrArray(pict), VarPtr(sa), 4 'Point the pict pointer to the image arrray
    
        Call Hough_InitialiseSingleImage(pict(0, 0), _
                                  ConfigInfo.MediaWidth, _
                                  ConfigInfo.MediaHeight)
        CopyMemory ByVal VarPtrArray(pict), 0&, 4 'ANT
                                  
    End If
''    ClearLineScreen.ControlToolbar.Buttons.Item(9).Image = 25 'PCN2681
    ClearLineScreen.InitVideoSlider
    If TuningStyle = "Automatic" Then
       ' Call AutoTune.LoadFileAutoTune
    End If
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS109:" & Error$
End Sub
Public Function UnitVideoSlider() 'PCNGL301202
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'UnitVideoSlider function  Louise Shrimpton louises@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    28/01/03     Building initial framework
'
'Description:
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    
VideoFrameSlider.value = 0 'PCN2930
AVITimeVar = 0
AVITime = 0
'vvvv PCN2782 *********************
Call VideoFrameSliderMove
'AVITimeLabel.Caption = AVITimeVar & " : " & AVITime
'^^^^ *****************************
    
Exit Function
Err_Handler:
    MsgBox Err & "-CLS110:" & Error$
End Function

Public Function UnInitVideo() 'PCNGL301202
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'UnInitVideo function  Louise Shrimpton louises@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    09/01/03     Building initial framework
'
'Description:
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    Call ProfilerPause
    If mediatype = Video Or mediatype = StillImage Then 'If this is not true, code has serious problems!! ls210103
        Call UnitVideoSlider
        AVIInitialised = False

        '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        DoEvents 'PCN2877 speed up uninitialise (10 June 2004, Antony van Iersel) don't really no if it helps. seems to.
        Call uninitialise
    End If
    mediatype = ""
Exit Function
Err_Handler:
    MsgBox Err & "-CLS111:" & Error$
End Function



Function Snap()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Snap function  Louise Shrimpton louiseS@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    18/12/02     Building initial framework
'
'Description:   Gets the bitmap information from the C code and draws it on
'               the picture box PVScreen bit by bit.  This will change.
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Dim buffer(0 To 442368) As Long '442367) As Long
Dim SnapFileName As String
Dim WatermarkFileName As String
Dim reg As Long
'''Dim bRememberFisheyeState As Boolean 'PCN4356

'''bRememberFisheyeState = ConfigInfo.FishEyeFlag 'PCN4356

''''If CLPScreenAction = "DrawMaskBox" Then
'''''    FisheyeFunctions.FEOFF 'PCN4356 'PCN4596, Never ever ever turn fisheye off
''''    'PCN4293 when making a maskes, need to see all of video
''''    ClearLineScreen.DimenResults.Visible = False
''''    ClearLineScreen.AreaResults.Visible = False
''''Else
''''    If ConfigInfo.FishEyeFlag = True Or ConfigInfo.FishEyeHorDistortion <> 1 Then
''''        FisheyeFunctions.FEON
''''    End If
''''End If

If Registered = False And SoftwareConfiguration <> "Reader" Then 'PCN Viewer has gray bar
    reg = 0
Else
    reg = 1
End If
'    SnapFileName = "e:\\Documents and Settings\\LouiseS\\Desktop\\6_1_03ClearLineProfilerV4\\Snapshot13103Test.bmp"
SnapFileName = LocToSave & "Snapshot.bmp"
WatermarkFileName = LocToSave & "Watermark.bmp"

    If ConfigInfo.FishEyeFlag = True Or ConfigInfo.FishEyeHorDistortion <> 1 Then
        FisheyeFunctions.FEON
    End If
    

If mediatype = Video Or mediatype = "Live" Then 'PCNLS050203
    If CLPScreenAction = "DrawMaskBox" Then
        ClearLineScreen.DimenResults.Visible = False
        ClearLineScreen.AreaResults.Visible = False
        Call grabsnapshot(SnapFileName, reg, WatermarkFileName, 0) 'PCN4356 snap without fisheye eg 0
    Else
        
        Call grabsnapshot(SnapFileName, reg, WatermarkFileName, 1) 'PCN4356 snap with fisheye
    End If
End If



'If bRememberFisheyeState Then FisheyeFunctions.FEON
    
Exit Function
Err_Handler:
    MsgBox Err & "-CLS112:" & Error$
End Function




Public Function UnitLiveFeed()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'UnitLiveFeed function  Louise Shrimpton louises@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    15/01/03
'
'Description:
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    If mediatype = "Live" Then 'ls 220103
        Call uninitialiseLive
    End If
    mediatype = ""
Exit Function
Err_Handler:
    MsgBox Err & "-CLS113:" & Error$
End Function

Public Function InitLiveFeed() As Boolean
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'InitLiveFeed function  Louise Shrimpton louises@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    15/01/03
'
'Description:
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    Dim connect As String
    connect = "Live"
   ' Call initialise(App.hInstance, VideoScreen.hwnd, connect, MediaWidth, MediaHeight, OvalPipe) 'PCNGL140103
   ' Call initialise(App.hInstance, VideoScreen.hwnd, connect, MediaWidth, MediaHeight, XAdjust, YAdjust, MediaOriginalHeight, MediaOriginalWidth) 'PCNGL140103
'    Select Case CaptureDevice
'            Case "Default"
'                Call initialise(App.hInstance, VideoScreen.hwnd, connect, MediaWidth, MediaHeight, XAdjust, YAdjust, MediaOriginalHeight, MediaOriginalWidth, "default") 'PCNGL140103 'PCN2289
'            Case "Nvidia"
'                Call initialise(App.hInstance, VideoScreen.hwnd, connect, MediaWidth, MediaHeight, XAdjust, YAdjust, MediaOriginalHeight, MediaOriginalWidth, "nvidia") 'PCNGL140103, 'PCN2289
'         Case Else
     Call initialise(App.hInstance, _
                     VideoScreen.hwnd, _
                     connect, _
                     "", _
                     ConfigInfo.MediaWidth, _
                     ConfigInfo.MediaHeight, _
                     XAdjust, _
                     YAdjust, _
                     MediaOriginalHeight, _
                     MediaOriginalWidth, _
                     VideoCaptureDevice, _
                     ConfigInfo.FishEyeOriginalWidth, _
                     ConfigInfo.FishEyeOriginalHeight, False) 'PCNGL140103 'PCN2289 'PCN2395 capture device from string to long
               
'    End Select
    If connect = "erro" Then 'PCN2418 (21 Nov 2003, Antony van Iersel) C++ laserlib is now return erro instead of error, error didn't fit in "Live" string causing it notice there was a problem
        InitLiveFeed = False
        'MsgBox DisplayMessage("No Capture Device, Please install a capture card and/or drivers."), vbInformation 'PCN2418
        ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("No Capture Device, Please install a capture card and/or drivers."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Else
        mediatype = "Live"
        LightInPipe = False
        UserDefinedLight = False
    '    Call AutoTune.AutoTuneImageProcessing
     '   AutoTune.Visible = False
        InitLiveFeed = True
    End If
Exit Function
Err_Handler:
    MsgBox Err & "-CLS114:" & Error$
End Function

Public Sub SeekTime(t As Double)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'SeekTime function  Louise Shrimpton louises@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    22/01/03
'
'Description:
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    If mediatype = Video Then 'ls 220103
        Call ProfilerPause
        
        Call timeseek(t)
    End If
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS115:" & Error$
End Sub



Public Function GetAVITime() As Boolean
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'GetAVITime function  Louise Shrimpton louises@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    22/01/03
'
'Description:  Gets the TOTAL time from the C in 100 nanoseconds and converts it to
'              seconds
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    Dim RecievedTime As Double
    If mediatype = Video Then 'ls 220103
        Call getTime(RecievedTime)
        AVITime = RecievedTime
        GetAVITime = True
    Else
        GetAVITime = False
        
    End If
Exit Function
Err_Handler:
    MsgBox Err & "-CLS116:" & Error$
End Function







Public Function ResizeVideoWindow() As Boolean
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'ResizeVideoWindow Function  Louise Shrimpton louises@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    24/01/03
'
'Description:
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    If mediatype = Video Or mediatype = "Live" Then 'ls 220103
        Call resizewindow
        ResizeVideoWindow = True
    Else
        ResizeVideoWindow = False
    End If
Exit Function
Err_Handler:
    MsgBox Err & "-CLS117:" & Error$
End Function


Public Function RefreshVideoScreen() As Boolean
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'RefreshVideoScreen Function  Louise Shrimpton louises@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    24/01/03
'
'Description:
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

If mediatype = Video Or mediatype = "Live" Then
    Call refreshframe
ElseIf mediatype = StillImage Then
        Call ScreenDrawing.ProcessSingleImage
End If

Exit Function
Err_Handler:
    MsgBox Err & "-CLS118:" & Error$
End Function

Public Function VideoScreenSetupAspectRatio() 'PCNGL2401032
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'VideoScreenSetupAspectRatio Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    24/01/03
'
'Description:
'       'Video Screen aspect ratio, Image Height/Width , determined by C code.
'       VideoAspectRatio holds value determined by the C code. This is used to
'       dimension the VideoScreen picture box.
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim CheckAspectRatio As Double



VideoScreen.height = MainScreen.height
VideoScreen.width = MainScreen.width

If VideoScreen.height = 0 Then Exit Function
CheckAspectRatio = VideoScreen.height / VideoScreen.width

If CheckAspectRatio <= 0 Or VideoAspectRatio <= 0 Then
    Exit Function
ElseIf CheckAspectRatio >= VideoAspectRatio Then
    VideoScreen.height = VideoScreen.width * VideoAspectRatio
Else
    VideoScreen.width = VideoScreen.height / VideoAspectRatio
End If

Exit Function
Err_Handler:
    MsgBox Err & "-CLS119:" & Error$
End Function


Public Function EndOfMedia() 'PCNGL2401032
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'EndOfMedia Function  Louise Shrimpton louises@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    29/01/03
'
'Description:
' goto last frame, set slider to last position, make time finish, set finished variables
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

EndOfAVI = True
If CheckAVIInitialised And mediatype <> "" Then
    VideoFrameSlider.value = VideoFrameSlider.Max 'PCN2930
    SliderFrame = VideoFrameSlider.value  'PCN2930
    Call VideoFrameSliderMove 'PCN2930
    AVITimeLabel.Caption = AVITimeVar & " : " & AVITime
End If

Exit Function
Err_Handler:
    MsgBox Err & "-CLS120:" & Error$
End Function

Public Sub GetCentreOfProfile(X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'SeekFrame function  Louise Shrimpton louises@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    30/01/03
'
'Description:
'       Returns the centre X and Y of the profile currently being displayed
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim CheckAspectRatio As Double 'PCN1863
Dim PVFrame As Long
Dim CurrentTime As Double
CurrentTime = GetCurrentAVITime


If PVDFileName = "" Then X = CentreLineX: Y = CentreLineY: Exit Sub
If mediatype = "Live" Or mediatype = StillImage Then X = CentreLineX: Y = CentreLineY: Exit Sub
If mediatype = Video And (CurrentTime < PVTimes(1) Or _
                          CurrentTime > PVTimes(PVDataNoOfLines)) Then
    X = CentreLineX: Y = CentreLineY: Exit Sub
End If

If (GetPVDVer >= 6.3) Then
    PVFrame = GetPVFrameNoFromTime(CurrentTime)
    X = TD_PVCentreX(PVFrame - 1)
    Y = TD_PVCentreY(PVFrame - 1)
    X = X / PVDrawScreenRatio
    Y = Y / PVDrawScreenRatio
    X = CentreLineX - X
    Y = PVScreen.height - CentreLineY + Y
    Exit Sub
End If

X = CentreLineX
Y = CentreLineY
    
Exit Sub
Err_Handler:
   MsgBox Err & "-CLS121:" & Error$
End Sub

Function SetupMTButtonsForVideo()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'SetupMTButtonsForVideo function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    30/01/03
'
'Description:
'       Enables and Disables the buttons required for Video mode
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim i As Integer

Exit Function
Err_Handler:
    MsgBox Err & "-CLS122:" & Error$
End Function



Sub TakeASnapShot()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'TakeASnapShot function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    05/02/03
'
'Description:
'       Tells the C code to create a BMP image and loads this into SnapShotScreen
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim NewVideoCentreLineX As Long 'PCNGL170203
Dim NewVideoCentreLineY As Long 'PCNGL170203
Dim NewNoOfCircles As Integer 'PCNGL170203
Dim PVFrame As Long

If mediatype = "" Then Exit Sub

Call ProfilerPause

'PCN4593, needed the "And PVRecording" in there
If (CLPScreenMode = PV Or CLPScreenMode = SnapShot) And PVDFileName <> "" And PVRecording = False Then Call SeekTime(PVTimes(PVFrameNo))


Call GetCentreOfProfile(VideoCentreLineX, VideoCentreLineY) 'PCNGL300103 'PCNGL170203
DrawingCentreX = VideoCentreLineX
DrawingCentreY = VideoCentreLineY

ClearLineScreen.SnapShotScreen.AutoRedraw = True 'PCNGL3001032
ClearLineScreen.SnapShotScreen.Cls 'PCNGL3001032

If mediatype <> StillImage Then 'PCN3194
    Call Snap 'PCNGL170203
    ClearLineScreen.LoadImage (LocToSave & "Snapshot.bmp")
End If

ClearLineScreen.SnapShotScreen.AutoRedraw = False

'VideoSnapShotMode = SnapShot
CLPScreenMode = SnapShot
Call ControlsScreen.ControlsViewSetup 'PCN4171


PVScreen.MousePointer = 99
PVScreen.MouseIcon = LoadResPicture(MainScreenMouseIcon, vbResIcon)

SnapShotScreen.Visible = True
SnapShotScreen.ZOrder 0



DoEvents 'PCNGL170203

Call DrawProfilesStartToFinish(ClearLineScreen.SnapShotScreen, True)

''Call SetupMTButtonsForSnapShot 'PCNGL300103
Call SetDimenResultsSize(True)  'PCN4171
DimenResults.ZOrder 0
AreaResults.ZOrder 0


Exit Sub
Err_Handler:
    MsgBox Err & "-CLS123:" & Error$
End Sub
Private Sub SnapShotScreen_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Handler
    Key_Pressed SnapShotScreen, KeyAscii 'PCN2312
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS124:" & Error$
End Sub
Private Sub PVScreen_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Handler
    Key_Pressed PVScreen, KeyAscii 'PCN2312
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS125:" & Error$
End Sub

'PCN2312 ==========================================================================v
Private Sub Key_Pressed(Pic As PictureBox, KeyAscii As Integer)
On Error GoTo Err_Handler

   If KeyAscii = 27 Then  'ESC
        Initialization
        PolyLineFlag = False
        Pic.Cls
        CurrentShp = 0
'        ConText = 0
        Call DrawAll
        
    ElseIf CLPScreenAction = "DrawText" Then
    End If
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS126:" & Error$
    Resume Next

End Sub
'PCN2312 ==========================================================================^


Function EmptyCBuffer()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'EmptyCBuffer function  Louise Shrimpton louiseS@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    3/04/03
'
'Description:
'       Empties the buffer in the C
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    If mediatype = Video Or mediatype = "Live" Then
        Call emptybuffer
    End If
Exit Function
Err_Handler:
    MsgBox Err & "-CLS127:" & Error$
End Function


Public Sub OptimizeSpeed(i As Long)
    If mediatype = Video Then
        Call Optimize(i)
    End If
End Sub

Public Sub SetRectangle(XBottom As Long, YBottom As Long, XTop As Long, YTop As Long, RectType As String) 'PCN2639
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'SetRectangle function  Louise Shrimpton louiseS@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    24/04/03
'
'Description:
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

If RectType = "Distance" Then
    If MeasurementUnits = "in" Then 'PCN2874 if units inches then count set to inches esle meters
        Call setdistancerectangle(XBottom, YBottom, XTop, YTop, 1) 'PCN2874 Last variable is unit 0 for meters, 1 for feet
    Else
        Call setdistancerectangle(XBottom, YBottom, XTop, YTop, 0) 'PCN2874 Last variable is unit 0 for meters, 1 for feet
    End If
Else
    'Call setrectanglecoord(XBottom, YBottom, XTop, YTop)
End If
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS128:" & Error$
End Sub

Public Sub ShowRect()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'ShowRectangle function  Louise Shrimpton louiseS@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    24/04/03
'
'Description:
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    Call showrectangle
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS129:" & Error$
End Sub

Public Function GetCurrentAVITime() As Double
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'GetCurrentAVITime function  Louise Shrimpton louiseS@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    1/05/03
'
'Description:
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim CurrentTime As Double
    If mediatype = Video Then
        Call getcurrenttime(CurrentTime)
        GetCurrentAVITime = CurrentTime
    End If
Exit Function
Err_Handler:
    MsgBox Err & "-CLS130:" & Error$
End Function

'Private Declare Sub setwaterlevel Lib "laserlib.dll" (ByVal theta1 As Long, ByVal theta2 As Long)

Public Sub SetWaterLevelinPipe(ByVal Theta1 As Double, ByVal Theta2 As Double) 'PCN4715 theta1 and theta2 was by ref, not by val. It changed the water level setting on set
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'SetWaterLevelinPipe function  Louise Shrimpton louiseS@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    21/05/03
'
'Description:  Calls the C to set the Water level in the pipe
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    If Theta1 = 0 And Theta2 = 0 Then Exit Sub
    If Theta1 > (2 * PI) Or Theta2 > (2 * PI) Or Theta1 < 0 Or Theta2 < 0 Then Exit Sub
    Dim wlLeft As Integer
    Dim wlRight As Integer
    Dim temp As Integer
    Dim i As Integer
        
' Out of bounds, if outside 360 degs, (2PI) then don't set and return.
' t1=PI-(PI/2); t2=PI+(PI/2); For Testing, Just put some values in
Theta1 = Theta1 + (PI / 90) 'PCN4064, added one degree onto the water level, better to much than to less
Theta2 = Theta2 - (PI / 90) 'PCN4064, added one degree onto the water level, better to much than to less

    

'////////////////////////'//
  wlLeft = Theta1 * 90 / PI  '// PCN1939 convert angle form 6 o'clock to which profile point
  wlRight = Theta2 * 90 / PI '// to egnore when water is on. If t1 is 1/4 PI and t2 is 3/4 PI
'////////////////////////'// then bottom half of profile is egnored.
  wlLeft = wlLeft + 90:   If wlLeft > 180 Then wlLeft = wlLeft - 180
  wlRight = wlRight + 90: If wlRight > 180 Then wlRight = wlRight - 180
'
'  // If t2 is less the t1 then swap, always
  If wlLeft > wlRight Then ' // make sure t2 is right and t1 is left.
      temp = wlLeft
      wlLeft = wlRight
      wlRight = temp
  End If


    'Build new egnorelist
    For i = 0 To (NoOfProfileSegments - 1)
        WaterEgnoreList(i) = 0
        If Abs(wlLeft - wlRight) < (NoOfProfileSegments / 2) Then
            If (i > wlLeft) And (i < wlRight) Then WaterEgnoreList(i) = 1
        Else
            If (i > -1 And i < wlLeft) Or (i > wlRight And i < NoOfProfileSegments) Then WaterEgnoreList(i) = 1
        End If
    Next i
    Call SetWaterLevel(WaterEgnoreList(0))
    
    'If PVDataNoOfLines > 1 Then Call ScreenDrawing.RecalculatePVData
 
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS131:" & Error$
End Sub

Function SetupVideoIgnoreBoarder(IgnoreBoarderCtrl As Control, x1 As Long, y1 As Long, x2 As Long, y2 As Long)   'PCNGL240403-1 'PCN2639
On Error GoTo Err_Handler

'vvvv PCN2639 ******************
'vvvv PCNGL140503-3 **********************************************
If MediaOriginalHeight > y2 Then 'PCNGL280503-1
    IgnoreBoarderCtrl.Top = (MediaOriginalHeight - y2) / (MediaOriginalHeight / SnapShotScreen.height) 'PCNGL280503-1
End If
If IgnoreY2 > IgnoreY1 Then 'PCNGL280503-1
    IgnoreBoarderCtrl.height = (y2 - y1) / (MediaOriginalHeight / SnapShotScreen.height) 'PCNGL280503-1
End If
IgnoreBoarderCtrl.Left = x1 / (ConfigInfo.MediaWidth / SnapShotScreen.width)  'PCNGL280503-1
IgnoreBoarderCtrl.width = (x2 - x1) / (ConfigInfo.MediaWidth / SnapShotScreen.width) 'PCNGL280503-1
'^^^^ ****************************************************
'^^^^ **************************

Exit Function
Err_Handler:
Select Case Err
    Case 380 'Invalid property
        Exit Function
    Case Else
        MsgBox Err & "-CLS132:" & Error$
End Select
End Function








Function DrawTextMouseDown(screenctrl As Control, Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'DrawTextMouseDown function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    3/06/03     Building initial framework
'
'Description:
'
'
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    Dim NumberOfTextBoxes As Integer

    NumberOfTextBoxes = OnScreenText.Count
    Load OnScreenText(NumberOfTextBoxes)
    
    OnScreenText(NumberOfTextBoxes).Left = X
    OnScreenText(NumberOfTextBoxes).Top = Y
    
    OnScreenText(NumberOfTextBoxes).Visible = True
    OnScreenText(NumberOfTextBoxes).ZOrder 0
    OnScreenText(NumberOfTextBoxes).SetFocus
    Call SetTextBoxWidthAndHeight(OnScreenText(NumberOfTextBoxes))
'ConText = 0
TextStartX = X
TextStartY = Y


Exit Function
Err_Handler:
    Select Case Err
        Case 9: NumberOfTextBoxes = 0: Resume Next
        Case Else: MsgBox Err & "-CLS133:" & Error$
    End Select
End Function


Public Sub Load3D()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Load3D function  Louise Shrimpton
'
'Description:
' Performs safety checks and loads the 3d
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    On Error GoTo Err_Handler
    Dim i As Long
    Dim J As Integer
    Dim TDArraySize As Long
    Dim TDArrayLineNo As Long
    Dim TDProfilesPerPage As Long 'Defines the maximum number of profiles per page
    Dim FileLoadError As Boolean
    Dim TDPage As Integer
    Dim NumData As Long
    Dim Multiplier As Long
    Dim curSelect As StdPicture 'PCN2230
    Dim LoadingProgress_3D As Integer 'PCN2988
    Dim LoadingProgress_3D_Current As Integer 'PCN2988
    Dim ProgressIncrement As Integer 'PCN2988
    Dim StartFrameNo As Long 'PCN2988
    
    MainScreen3D.ZOrder 0 'PCN2230
    Set curSelect = LoadResPicture(108, vbResIcon) 'PCN2230 Hand icon
    MainScreen3D.MousePointer = 99 'PCN2230
    MainScreen3D.MouseIcon = curSelect 'PCN2230
    
   
    
    If ThreeDRunning = False Then
        Call CLPProgressBar.ProgressBarInitialise(DisplayMessage("Loading 3D data"))
        DoEvents
        If PVDFileName = "" Then
            Call CLPProgressBar.ProgressBarPosition(1#) 'PCN2401 'Can't load 3D module without PV data to process.
            Exit Sub
        End If

        'vvvv 'PCN2401 *********************************
        If IsOpen("CLPProgressBar") Then
            Call CLPProgressBar.ProgressBarPosition(1#)
            DoEvents
        End If
        
        Dim Path As String
        Dim ExpRad As Double 'The expected radius
        Dim Units3D As Long
                  
          

        Path = App.Path & "\"
        ExpRad = ExpectedDiameter / 2
        If MeasurementUnits = "mm" Then
            Units3D = 0
        Else
            Units3D = 1
        End If
        
    
        
        'PCN4197 at this stage its easier to just pass the original string then to
        ' change the cpp 3d code to except 0 or 1 as hardware or software.
                                                    '
        Dim ThreeDStyleToBePassed As String         '
        If ThreeDRenderingStyle = 0 Then            '
            ThreeDStyleToBePassed = "Hardware"      '
        Else                                        '
            ThreeDStyleToBePassed = "Software"      '
        End If                                      '
                                                    '
        '''''''''''''''''''''''''''''''''''''''''''''
  
        Call d3d_initialise(App.hInstance, _
                            MainScreen3D.hwnd, _
                            ThreeDStyleToBePassed, _
                            TD_PVDataX(0), _
                            TD_PVDataY(0), _
                            TD_PVCentreX(1), _
                            TD_PVCentreY(1), _
                            PVDataNoOfLines, _
                            Path, _
                            PVFlat3DRed(0, 0), _
                            PVFlat3DGreen(0, 0), _
                            PVFlat3DBlue(0, 0), _
                            ExpRad, _
                            PVCalculationsMultiplier, _
                            PVDataXYMultiplier, Units3D)

        Call d3d_laser_focus(PVFrameNo) 'PCN4087
 
        
      
        




        ThreeDtimer.Enabled = True
        ThreeDRunning = True
    Else
        ThreeDRunning = False
        ThreeDtimer.Enabled = False
        Unload3D
    End If
    
    'PCN???? 16 Dec 2010, Japanese crasshed because of the following lines, now commented out - Antony and Richard
' PCN2473 Antony van Iersel 11 March 2004
' Loading language for D3D into a language array. This is copied to C++
' which is then displayed in the 3D panels and messages
    ReDim D3DLanguageArray(20)
'    D3DLanguageArray(0) = DisplayMessage("Frame") & ":"
'    D3DLanguageArray(1) = DisplayMessage("Capacity") & ":"
'    D3DLanguageArray(2) = DisplayMessage("Ovality") & ":"
'    D3DLanguageArray(3) = DisplayMessage("Delta max") & ":"
'    D3DLanguageArray(4) = DisplayMessage("Delta min") & ":"
'    D3DLanguageArray(5) = DisplayMessage("Loading")
'    D3DLanguageArray(6) = DisplayMessage("Re-Painting")
'    D3DLanguageArray(7) = DisplayMessage("Sorry, Hardware Vertexing is not available on your Graphics Processing Unit")
'    D3DLanguageArray(8) = DisplayMessage("Cannot Initialise Hardware Vertexing")
'    D3DLanguageArray(9) = DisplayMessage("Please Try Mixed or Software Vertexing")
'    D3DLanguageArray(10) = DisplayMessage("Sorry, Mixed Vertexing is not available on your Graphics Processing Unit")
'    D3DLanguageArray(11) = DisplayMessage("Cannot Initialise Mixed Vertexing")
'    D3DLanguageArray(12) = DisplayMessage("Please Try Software Vertexing")
'    D3DLanguageArray(13) = DisplayMessage("Sorry, Software Vertexing is not available on your Graphics Processing Unit")
'    D3DLanguageArray(14) = DisplayMessage("This hardware cannot run the 3D Application")
    
' The copying of the language array to the C++ D3D language array
    Dim arrayIndex As Integer
    Dim d3dLanguageLines As Integer
    d3dLanguageLines = 14
    For arrayIndex = 0 To d3dLanguageLines
        Call d3d_setlanguage(arrayIndex, D3DLanguageArray(arrayIndex)) '(array line to be copied into, string to be copied)
    Next arrayIndex
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Call ControlsScreen.ExecuteControlsFixedButton(11) 'PCN4087 turn map off by default
Exit Sub
Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & "-CLS134:" & Error$
End Select
End Sub

Public Sub Unload3D()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Unload3D function  Louise Shrimpton
'
'Description:
' Performs safety checks and unloads the 3d
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    On Error GoTo Err_Handler
    If ThreeDRunning = True Then
        ThreeDtimer.Enabled = False 'PCN2240
        'Call uninitializethreed
        Call d3d_destroy
        Dim i As Integer
        For i = 0 To 1000 Step 1  'sleep
        Next i
''        ClearLineScreen.ThreeDToolbar1.Visible = False 'PCN2240
        ThreeDRunning = False
    End If
Exit Sub
Err_Handler:
Select Case Err
    Case 11 ' divide by zero
        MsgBox Err & " - " & Error$
'        ab = ab + 0.001
    Case 6 'overflow - this is caused when center point and a point in the circumference is same.
    Case 5 '
        'MsgBox "Temporary message after mouseup. Please check your drawing is as intended."
    Case 53 'Could not find the mouse icon
        Resume Next
    Case Else
        MsgBox Err & "-CLS135:" & Error$
End Select
End Sub

Public Function SetUnsetWaterLevel(IgnoreCenter As Boolean, IgnoreProfile As Boolean)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'SetUnsetWaterLevel function  Louise Shrimpton
'
'Description:
' sets and unsets the waterlevel function in the c
' useful for increasing speed, if the waterlevel function is not set in the c
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    Dim i, J As Long
    If IgnoreCenter = True Then
        i = 1
    Else
        i = 0
    End If
    If IgnoreProfile = True Then
        J = 1
    Else
        J = 0
    End If
    Call setwaterlevelbool(i, J)
Exit Function
Err_Handler:
Select Case Err
    Case 11 ' divide by zero
        MsgBox Err & " - " & Error$
'        ab = ab + 0.001
    Case 6 'overflow - this is caused when center point and a point in the circumference is same.
    Case 5 '
        'MsgBox "Temporary message after mouseup. Please check your drawing is as intended."
    Case 53 'Could not find the mouse icon
        Resume Next
    Case Else
        MsgBox Err & "-CLS136:" & Error$
End Select
End Function



Public Sub GetBlockOfPVData()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : GetBlockOfPVData
'Created : 20 June 2004, PCN2891
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Gets the latest block of PVData from the C and draws the PVGraphs.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
If mediatype <> Video Then Exit Sub

Dim numframesd As Double  'this is the total number of frames / 1000
Dim NumFrames As Long 'the number of frames when multiplied by 1000


Dim Allradius() As Long
Dim AllradiusX() As Single 'PCN6004
Dim AllradiusY() As Single 'PCN6004
Dim Alltime() As Double
Dim Alldistances() As Long 'PCN2639
Dim AllCentresX() As Single 'PCN6004 'PCN3219
Dim AllCentresY() As Single 'PCN6004 'PCN3219

Dim i As Long
Dim J As Integer
Dim PVDataAddressOffset As Long
Dim LenOfPVData As Integer
Dim PVDFileSaveFail As Boolean
Dim PVArrayBufferNo As Integer
Dim ErrorStatus As String
Dim FastDrawData() As Long
Dim GraphIndex As Integer
Dim PVDataStartFrameNo As Long 'Start FrameNo of the PVData save block
'''''''''''''''''''''''''''''''''''''''''''''''''''
    
Call gettotalnumframes(NumFrames)

If NumFrames < 2 Then Exit Sub
'All this data is retrieved from laserlib.dll ''''''''''''''
ReDim Allradius((NumFrames + 1) * NoOfProfileSegments)
ReDim AllradiusX((NumFrames + 1) * NoOfProfileSegments)
ReDim AllradiusY((NumFrames + 1) * NoOfProfileSegments)
ReDim Alltime(NumFrames + 1)
ReDim Alldistances(NumFrames + 1) 'PCN2639
ReDim AllCentresX(NumFrames + 1) 'PCN3219
ReDim AllCentresY(NumFrames + 1) 'PCN3219

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

RequestFrameNo = 0 ' LastRecordedFrame


Call getgroupedprofiledataxy(AllradiusX(0), AllradiusY(0), _
                             AllCentresX(0), AllCentresY(0), _
                             Alltime(0), _
                             NumFrames, _
                             Alldistances(0)) 'PCN2129 'PCN3219 added all centres
'Determine VideoScreenScale

Call VideoScreenScaleCalc



Dim temp As Double
'Now Draw the graphs
'vvvv PCN2112 ******************************************
'Get the PVData to file as fast as possible. Then draw only
'the currently selected graph. DON'T DRAW the PVPROFILE.
'Setup the PVD file header if the file has not yet been created.
PVDFileSaveFail = False
Call CreatePVDFileHeader(PVDataAddressOffset, LenOfPVData, PVDFileSaveFail)

PVDataStartFrameNo = NumTimesRecorded - 1

PVDataNoOfLines = PVDataNoOfLines + NumFrames
PVFrameNo = PVDataNoOfLines

Call ReDimensionIndicators



ReDim Preserve TD_PVDataX((PVDataNoOfLines + 10) * NoOfProfileSegments)
ReDim Preserve TD_PVDataY((PVDataNoOfLines + 10) * NoOfProfileSegments)

ReDim Preserve TD_PVCentreX(PVDataNoOfLines + 10) 'PCN3219
ReDim Preserve TD_PVCentreY(PVDataNoOfLines + 10) 'PCN3219


ReDim Preserve PVXDiameterFullData(PVDataNoOfLines + 10) 'PCN2703
ReDim Preserve PVYDiameterFullData(PVDataNoOfLines + 10) 'PCN2703
ReDim Preserve PVCapacityFullData(PVDataNoOfLines + 10)

ReDim Preserve GraphInfoContainer(PVOvality).DataSingle(PVDataNoOfLines + 10)


ReDim Preserve GraphInfoContainer(PVDebris).DataSingle(PVDataNoOfLines + 10) 'PCN4461
ReDim Preserve PVDiameterMedian(PVDataNoOfLines + 10)
'ReDim Preserve PVFractile(PVDataNoOfLines + 10) 'PCN4235

ReDim Preserve PVDeltaFullMax(PVDataNoOfLines + 10)
ReDim Preserve PVDeltaFullMin(PVDataNoOfLines + 10)
ReDim Preserve PVDeltaSegFullMax(PVDataNoOfLines + 10)
ReDim Preserve PVDeltaSegFullMin(PVDataNoOfLines + 10)

ReDim Preserve GraphInfoContainer(PVMaxDiameter).DataDouble(PVDataNoOfLines + 10)
ReDim Preserve GraphInfoContainer(PVMinDiameter).DataDouble(PVDataNoOfLines + 10) 'PCN4333
ReDim Preserve PVDiameterFullMin(PVDataNoOfLines + 10)
ReDim Preserve PVDiameterSegFullMax(PVDataNoOfLines + 10)
ReDim Preserve PVDiameterSegFullMin(PVDataNoOfLines + 10)

ReDim Preserve PVFlat3DRed(NoOfProfileSegments, PVDataNoOfLines + 10) As Long
ReDim Preserve PVFlat3DGreen(NoOfProfileSegments, PVDataNoOfLines + 10) As Long
ReDim Preserve PVFlat3DBlue(NoOfProfileSegments, PVDataNoOfLines + 10) As Long

ReDim Preserve PVTimes(PVDataNoOfLines + 10)
ReDim Preserve PVDistances(PVDataNoOfLines + 10)

ReDim Preserve PVShapeCentreX(PVDataNoOfLines + 10)
ReDim Preserve PVShapeCentreY(PVDataNoOfLines + 10)
ReDim Preserve GraphInfoContainer(PVDeflectionX).DataSingle(PVDataNoOfLines + 10) 'PCN5186
ReDim Preserve GraphInfoContainer(PVDeflectionY).DataSingle(PVDataNoOfLines + 10) 'PCN5186

ReDim Preserve SmoothDeflectionX(PVDataNoOfLines + 10)
ReDim Preserve SmoothDeflectionY(PVDataNoOfLines + 10)

'ReDim Preserve GraphInfoContainer(PVInclination).DataSingle(PVDataNoOfLines + 10) 'PCN6128 'PCN6458
'ReDim Preserve GraphInfoContainer(PVDesignGradient).DataSingle(PVDataNoOfLines + 10) 'PCN6178 'PCN6458

'^^^^ **************************************************************

Dim FileNo As Integer
FileNo = FreeFile

Open PVDFileName For Binary Access Write As #FileNo
PVDHeaderPVData.PVDCheck = PVDataNoOfLines
Put #FileNo, PVDFilePointers.PVDPointerPVData, PVDHeaderPVData
'Set the start of the block write for the PVData
PVDataAddressOffset = PVDataAddressOffsetCalc(Seek(FileNo), PVDataStartFrameNo) 'PCN2639


'PCN3526 '''''''''''''''''''''
Dim XYDataArrayIndex As Long 'Calculated index for the XY data for laserlib dll
Dim TD_XYDataOffset As Long  'Where was the last entry writen into the TD_Datax and y
                             '
Dim DataX As Single 'Temp place to place XY data used for populating PVData and
Dim DataY As Single 'calculating the TD_Data x and y
                             '
'Note: frames start at one, TD_PVDataX and TD_PVDataY start at 0, that is why -1
TD_XYDataOffset = PVDataStartFrameNo * NoOfProfileSegments
                             '
''''''''''''''''''''''''''''''

'i was initialised to 1 but now 0, and all use of i was i-1 now just i
For i = 0 To (NumFrames - 1) 'PCN2216 'PCN2501
    'PVFrameNo = NumTimesRecorded 'PCN2129
    'PVArrayBufferNo = PVFrameNo Mod (MaxFrameBufferNo + 1)
    
    PVArrayBufferNo = 1 'PCN3526
    

    'vvvv PCN2501 ***************************************
    
    For J = 1 To NoOfProfileSegments 'PCN2501
        
        'PCN3526 DataX added and everything tied to it except the old
        'PVData(j, 1, PVArrayBufferNo) = DataX etc
        
        XYDataArrayIndex = (i * NoOfProfileSegments) + J
        DataX = AllradiusX(XYDataArrayIndex) 'Retrieve X co-ordinates
        DataY = AllradiusY(XYDataArrayIndex) 'Retrieve Y co-ordinates
        
        pvData(J, 1, PVArrayBufferNo) = DataX
        pvData(J, 2, PVArrayBufferNo) = DataY
        
        TD_PVDataX(TD_XYDataOffset + XYDataArrayIndex) = DataX * ConfigInfo.Ratio * VideoScreenScale
        TD_PVDataY(TD_XYDataOffset + XYDataArrayIndex) = DataY * ConfigInfo.Ratio * VideoScreenScale
     

        'Calculate Radius - If required.
        pvData(J, 0, PVArrayBufferNo) = Int(PVDataTrueRadiusCalc(J, PVArrayBufferNo))

    Next J
    '^^^^ ***********************************************
    TD_PVCentreX(NumTimesRecorded) = AllCentresX(i) * ConfigInfo.Ratio * VideoScreenScale
    TD_PVCentreY(NumTimesRecorded) = AllCentresY(i) * ConfigInfo.Ratio * VideoScreenScale
 
    'load time data in for the current frame
    CurrentAVITime = Alltime(i) 'PCN2639 'PCN2959'was looking to far down the array for last time
    PVTimes(NumTimesRecorded) = CurrentAVITime
    'vvvv PCN2639 *****************************************
    
    
    'Determine Distance if required
    If ConfigInfo.DistanceStart > InvalidData Or IPD = True Then 'PCN3744
        'Check how many decimal places the counter requires
        
        If MeasurementUnits = "mm" Then 'PCN3599 if in imperial it still treated like metric same as above
            'PCN3744 added IPD
            If IPD = False Then PVDistances(NumTimesRecorded) = (Alldistances(i) / 10) 'PCN3272 fogot to adjust the index on the dump, same as all the rest of the array dumps (14 Jan 2005, Ant)
            If IPD = True Then PVDistances(NumTimesRecorded) = (Alldistances(i) / 100) 'PCN3272 fogot to adjust the index on the dump, same as all the rest of the array dumps (14 Jan 2005, Ant)
        Else
            'PCN3744 added IPD
            If IPD = False Then PVDistances(NumTimesRecorded) = (Alldistances(i))
            If IPD = True Then PVDistances(NumTimesRecorded) = (Alldistances(i) / 10)
        End If
        
        If PVDistances(NumTimesRecorded) = InvalidData Then
            If NumTimesRecorded = 0 Then
                PVDistances(NumTimesRecorded) = 0
            Else
                PVDistances(NumTimesRecorded) = PVDistances(NumTimesRecorded - 1)
            End If
        End If
        


        
    Else
        PVDistances(NumTimesRecorded) = NumTimesRecorded 'Not valid data
    End If
    '^^^^ *************************************************
    
    Call RapidSavePVDataToFile(NumTimesRecorded, PVArrayBufferNo, PVDataAddressOffset, LenOfPVData, PVDFileSaveFail, FileNo)
   
    NumTimesRecorded = NumTimesRecorded + 1

Next i
Close #FileNo
    
'    Call ScreenDrawing.PVCentreCalcCPP(PVDataStartFrameNo + 1, PVDataNoOfLines)
   
    Call ScreenDrawing.PVDiameterMedianCalcCPP(PVDataStartFrameNo + 1, PVDataNoOfLines)
    Call ScreenDrawing.PVCapacityCalcCPP(PVDataStartFrameNo + 1, PVDataNoOfLines)
    Call ScreenDrawing.PVOvalityCalcCPP(PVDataStartFrameNo + 1, PVDataNoOfLines)
    Call ScreenDrawing.PVXYDiameterCalcCPP(PVDataStartFrameNo + 1, PVDataNoOfLines)
'    Call ScreenDrawing.PVDeltaMaxMinCalcCPP(PVDataStartFrameNo + 1, PVDataNoOfLines)
    Call ScreenDrawing.PVDiameterMaxMinCalcCPP(PVDataStartFrameNo + 1, PVDataNoOfLines)
    Call ScreenDrawing.FixMinMax(PVDataStartFrameNo + 1, PVDataNoOfLines) 'PCN6524
'    Call ScreenDrawing.PVInclinationCalc(PVDataStartFrameNo + 1, PVDataNoOfLines) 'PCN6458
    
  
    'Call ScreenDrawing.PVDebrisCalcCPP(PVDataStartFrameNo + 1, PVDataNoOfLines) 'PCN4461

'    Call ScreenDrawing.PVFractileCalcCPP(PVDataStartFrameNo + 1, PVDataNoOfLines)
'
'
'
'
LastDataTime = CurrentAVITime
Dim RoundingDia As Integer
If MeasurementUnits = "mm" Then RoundingDia = 1
If MeasurementUnits <> "mm" Then RoundingDia = 2

If OptionsPage.SmoothCheck.value = vbChecked Then Call CPPSmoothGraphs

    If PipelineDetails.InternalDiameterExpected = "" Then
        ExpectedDiameter = CalculatePercentile(50, "MedianDiameter")
        ExpectedDiameter = Round(ExpectedDiameter, RoundingDia)
        If ExpectedDiameter <= 0 Then ExpectedDiameter = 200
    End If
    Call ScreenDrawing.PVFlat3DCalcCPP(PVDataStartFrameNo + 1, PVDataNoOfLines)
    
    Call ScreenDrawing.PVDeflectionCalcCPP   'PCN5186

Call GotoPVProfile(PVFrameNo, True)




'Call GotoPVGraphProfile(PVFrameNo)
'ClearLineScreen.PVFrameStatus.Caption = PVFrameNo & DisplayMessage(" of ") & PVDataNoOfLines 'PCN2111
'vvvv PCN2639 *****************************************************
'Redraw PVYScale
If DistanceMethod = "AutomaticCounter" Then
    If PVDataNoOfLines > 2 Then
        CameraSpeedInFrames = (PVDistances(PVDataNoOfLines - 1) - PVDistances(1)) / (PVDataNoOfLines - 2)
    Else
        'Call PrecisionVisionGraph.PVYScaleClearAll
        '^^^^ *****************************************************
    End If
End If
'^^^^ *************************************************************
DoEvents
    
Call Distance.RecalculateDistance

Call PrecisionVisionGraph.MoveGraph(PVFrameNo)
Call ScreenDrawing.DrawPVGraphs
    
If PVRecording = True Then Call ProfilerRun
Call setrecprofstat(1)
    
Exit Sub
Err_Handler:
Select Case Err
    Case 6: Resume Next 'ANT
    
    Case Else
        MsgBox Err & "-CLS137:" & Error$
        

End Select
End Sub

Public Sub CreatePVDFileHeader(PVDataAddressOffset As Long, LenOfPVData As Integer, PVDFileSaveFail As Boolean)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'CreatePVDFileHeader function  Geoff Logan 'PCN2112
'
'Description:
' Checks if the PVD file has been created. If it has not, this function will
' create it and setup the file header information, ready for data storage.
'
'Purpose:
'
'
On Error GoTo Err_Handler
Dim PVArrayBufferNo As Integer
Dim ErrorStatus As String
Dim GraphIndex As Integer
Dim i, J As Integer

ErrorStatus = "OK"
'PVArrayBufferNo = PVDataNoOfLines Mod (MaxFrameBufferNo + 1)
PVDFileSaveFail = False
If PVDFileName = "" Or Dir(PVDFileName) = "" Then
    PVDFileName = LocToSave & DefaultPVDFileName
    '********************************************************
    'vvvvv ***** Initialise file ****************************
    'Draw only 1 windows of data (for loading speed) 'PCNGL070103
    MaxDisplayedFrameNo = PrecisionVisionGraph.PVGraphWindow.height
    'From Abe's code. Only use temporary
    Const TextType As Integer = 1 ' 1: 180(Dia), 2:360(180 Dia, 180 Color)
    'Initialize some variables         'Testing ls 10/12/02
    If ImageGraphState(0).GraphType = Empty Then
        ImageGraphState(0).GraphType = "Capacity"
    End If
    Call InitilisePVProfile(MaxFrameBufferNo) 'PCNGL241202
    ErrorStatus = "Kill file"
    Kill PVDFileName
'    Call SaveInPVDFormat("4.0", PVDFileName, PVDFileSaveFail) 'PCN2639
'    Call SaveInPVDFormat("5.0", PVDFileName, PVDFileSaveFail) 'PCN2639 'PCN2891
    Call SaveInPVDFormat("6.X", PVDFileName, PVDFileSaveFail) 'PCN2891
End If

Dim FileNo As Integer
FileNo = FreeFile


If Not PVDFileSaveFail Then
    'Check whether a file is open
    Open PVDFileName For Binary Access Read Lock Write As #FileNo 'PCN2208
    'Load the File Main Header
    Get #FileNo, , PVDFileMainHeader
Else
    PVDFileMainHeader.PVDFileMHPointerAddress = 0
End If
'Read the file header pointers
If PVDFileMainHeader.PVDFileMHPointerAddress <> 0 Then 'PCNGL140103
    Get #FileNo, PVDFileMainHeader.PVDFileMHPointerAddress, PVDFilePointers
    Get #FileNo, PVDFilePointers.PVDPointerPVData, PVDHeaderPVData
    'PCN3274 If left(PVDHeaderPVData.PVDHeaderDescriptor, 8) = "[PVData]" And PVDHeaderPVData.PVDCheck <> 0 Then 'PCNGL130103
    If Left(PVDHeaderPVData.PVDHeaderDescriptor, 8) = "[PVData]" And PVDHeaderPVData.PVDCheck >= 0 Then 'PCNGL130103
        LenOfPVData = Len(pvData(1, 0, 1))
        'Goto the correct address for writing PVData
'        PVDataAddressOffset = Seek(1) + (PVDataNoOfLines) * LenOfPVData * NoOfProfileSegments  ' The calculation for the size of PVData file data block
'        PVDataAddressOffset = PVDataAddressOffset + (PVDataNoOfLines) * (Len(PVCapacityData(1)) + Len(PVOvalityData(1)) + 2 * Len(PVDelta(1)) + Len(AVIFrameTime(1))) ' The calculation for the size of PVData file data block 'PCNGL150103
        PVDataAddressOffset = PVDataAddressOffsetCalc(Seek(FileNo), PVDataNoOfLines) 'PCN2639
    Else
        PVDFileSaveFail = True
    End If
Else
    PVDFileSaveFail = True 'PCNGL140103
End If
Close #FileNo



Exit Sub
FileErr_Handler:
    Close #FileNo
Exit Sub
Err_Handler:
Select Case Err
    Case 9 'Out of subscript range.
        GoTo FileErr_Handler
    Case 53 'File not found (Kill statement error trap) 'PCNGL140103
        If ErrorStatus = "Kill file" Then Resume Next
        PVDFileSaveFail = True
    Case Else
        MsgBox Err & "-CLS138:" & Error$
End Select
End Sub



Public Sub RapidSavePVDataToFile(CurrentFrameNo As Long, PVArrayBufferNo As Integer, PVDataAddressOffset As Long, LenOfPVData As Integer, PVDFileSaveFail As Boolean, ByVal FileNo)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'RapidSavePVDataToFile function  Geoff Logan 'PCN2112
'
'Description:
' Saves to file the PVData from the C++ module as fast as possible.
'
'Purpose:
'
'***************************************************************
On Error GoTo Err_Handler
Dim ErrorStatus As String
Dim GraphIndex As Integer
Dim i, J As Integer


'vvvv PCN2891 ***** Save PVData to PVD file  **********************
For i = 1 To NoOfProfileSegments
    If Not PVDFileSaveFail Then
        'Save data to file
        
'PCN2980 (8 September 2005, Antony) We no longer save Radius, so its not needed
'        If ConfigInfo.ProfileRecordingMethod = "XY" Then''''''''''''''''
            'Save the PVData X                                          '
            Put #FileNo, PVDataAddressOffset, pvData(i, 1, PVArrayBufferNo)  '
            PVDataAddressOffset = PVDataAddressOffset + LenOfPVData     '
            'Save the PVData Y                                          '
            Put #FileNo, PVDataAddressOffset, pvData(i, 2, PVArrayBufferNo)  '
            PVDataAddressOffset = PVDataAddressOffset + LenOfPVData     '
'        Else                                                           '
'            'Save the PVData Radius                                    '
'            Put #1, PVDataAddressOffset, PVData(i, 0, PVArrayBufferNo) '
'            PVDataAddressOffset = PVDataAddressOffset + LenOfPVData    '
'        End If''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End If
    J = J + 1
Next i
'^^^^ ****************************************************************
'vvvv ***** Save the PVData and calculations to PVD file  **********************
If Not PVDFileSaveFail Then
    Put #FileNo, , TD_PVCentreX(CurrentFrameNo)
    Put #FileNo, , TD_PVCentreY(CurrentFrameNo)
    
'    Put #1, , PVCapacityData(PVArrayBufferNo)
'    Put #1, , PVOvalityData(PVArrayBufferNo)
'    Put #1, , PVDelta(0, PVArrayBufferNo)
'    Put #1, , PVDelta(1, PVArrayBufferNo)
    'vvvv **** Store current AVI frame time ***************************
    'To be used to more accurately link the PVD file Frame no to the AVI frames. PCNGL150103
    PVDataAddressOffset = PVDataAddressOffset + PVCalculationsBlockSize  'PCN2639
    Put #FileNo, PVDataAddressOffset, PVTimes(CurrentFrameNo)
    '^^^^ **************************************************************
    'vvvv PCN2639 ****************************
    Put #FileNo, , PVDistances(CurrentFrameNo)
    '^^^^ ************************************
'    PVDataAddressOffset = PVDataAddressOffset + (Len(PVCapacityData(1)) + Len(PVOvalityData(1)) + 2 * Len(PVDelta(1)) + Len(AVIFrameTime(1)) + Len(PVDistances(1))) 'PCN2639
    PVDataAddressOffset = PVDataAddressOffset + PVRelatedInfoBlockSize  'PCN2639
End If


Exit Sub
FileErr_Handler:
    Close #FileNo
Exit Sub
Err_Handler:
Select Case Err
    Case 9 'Out of subscript range.
        GoTo FileErr_Handler
    Case 53 'File not found (Kill statement error trap) 'PCNGL140103
        If ErrorStatus = "Kill file" Then Resume Next
        PVDFileSaveFail = True
    Case Else
        MsgBox Err & "-CLS139:" & Error$
End Select
End Sub

Public Sub StorePVCalcDataInArrays(PVArrayBufferNo As Integer)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : StorePVCalcDataInArrays
'Created : 17 September 2003, PCN2240
'Updated :
'Prg By  : Geoff Logan
'Param   : PVArrayBufferNo - set to 0 or 1 as per the current buffer number
'Desc    : This code stores the various calculations in the specific
'           arrays. The arrays will then be used by the 3D module and
'           PVGraph paging function.
'Usage   :
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

'^^^^ *******************************************
'PVCapacityFullData(PVFrameNo) = Round(PVCapacityData(PVArrayBufferNo) / PVCalculationsMultiplier, 2)
'PVCapacityFullData(PVFrameNo) = PVCapacityData(PVArrayBufferNo)
'PVOvalityFullData(PVFrameNo) = PVOvalityData(PVArrayBufferNo)
'PVDeltaFullMax(PVFrameNo) = PVDelta(0, PVArrayBufferNo)
'PVDeltaFullMin(PVFrameNo) = -PVDelta(1, PVArrayBufferNo)

''Calculate Flat3D
'Call PVFlat3DCalc(PVFrameNo, PVArrayBufferNo)

Exit Sub
Err_Handler:
Select Case Err
    Case 9 'Out of subscript range.
    Case Else
        MsgBox Err & "-CLS140:" & Error$
End Select
End Sub


Public Sub D3D_CaptureImage(SnapFileName As String)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : D3D_CaptureImage
'Created : 8 October 2003, PCN2240
'Updated :
'Prg By  : Geoff Logan
'Param   : (none)
'Desc    : This code creates a bmp file of the contians of the MainScreen3D
'           window. This is used for the reports and saving to file.
'Usage   :
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Call d3d_refresh 'PCN sometimes the 3d screen was pink, not refreshed, this is a force.
Call d3d_capture_window(SnapFileName, MainScreen3D.hwnd)


Exit Sub
Err_Handler:
    MsgBox Err & "-CLS141:" & Error$
End Sub


Public Sub D3D_ExportToFile(ExportFileName As String, ExportFormat As String)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : D3D_CaptureImage
'Created : 21 November 2003, PCN2376
'Updated :
'Prg By  : Geoff Logan
'Param   : (none)
'Desc    : Calls the ThreeDim.dll code to export the 3D pipe to file in STL format.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Select Case ExportFormat
    Case "STL"
        Call d3d_export_stl(ExportFileName)
    Case Else
End Select

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS142:" & Error$
End Sub




Public Sub DistanceCounterSet(DistanceStart As Double)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PVRecordButtonSet
'Created : 27 March 2003, PCN2639
'Updated :
'Prg By  : Geoff Logan
'Param   : DistanceCounter -
'          CountDir -
'Desc    : Calls the LaserLib.dll setdistancecounter function
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim DistanceCounter As Long
Dim ConvertToDistCount As Double
Dim CountDir As Long

' PCN2874 (Antony 10 Aug 2004) when in metric counter start is multiplied by 10
' else it is current distance               '
If MeasurementUnits = "mm" Then             '
    ConvertToDistCount = DistanceStart * 10 '
Else                                        '
    ConvertToDistCount = DistanceStart      '
End If                                      '
'''''''''''''''''''''''''''''''''''''''''''''
ConvertToDistCount = Int(ConvertToDistCount)
DistanceCounter = CLng(ConvertToDistCount)
If PVRecording Then Exit Sub

If CountDirection = "Up" Then
    CountDir = 1
Else
    CountDir = 0
    CountDirection = "Down"
End If

Call setdistancecounter(DistanceCounter, CountDir)

ConfigInfo.DistanceStart = DistanceStart
ConfigInfo.DistanceDirection = CountDirection
ConfigInfo.DistanceFinish = -1 'Not required



Exit Sub
Err_Handler:
    MsgBox Err & "-CLS143:" & Error$
End Sub

Public Sub ClearDistRectangle()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : ClearDistRectangle
'Created : 29 March 2003, PCN2639
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Calls the LaserLib.dll and clears the distance counter
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Call cleardistancerectangle
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS144:" & Error$
End Sub






Public Sub GotoPVProfile(ByVal FrameNo As Long, Optional PermanantOptional)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : GotoPVProfile
'Created : 10 June 2004, PCN2883
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Sets up the PV profile and supporting functions for Frame FrameNo.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim PVFileLoadError As Boolean
Dim Permanant As Boolean


If PVDFileName = "" Then Exit Sub 'There is no profile to go to :(
If IsMissing(PermanantOptional) Then
    Permanant = False
Else
    Permanant = PermanantOptional
End If



PVFrameNo = FrameNo
ShapeCentreX = PVShapeCentreX(PVFrameNo) / PVDrawScreenRatio
ShapeCentreY = PVShapeCentreY(PVFrameNo) / PVDrawScreenRatio

PVScaleMarkerStFrame = PVFrameNo
PVScaleMarkerFnFrame = PVFrameNo
Call PrecisionVisionGraph.RepositionYScaleMarkers


'Call PVGraphsKeyForm.PVGraphsKeyUpdate 'PCN2990
Call DistanceCounterTextUpdate  'PCN2928

If Not PVRecording Then
    If CLPScreenMode = PV Then
        If Permanant Then
            Call ScreenDrawing.DrawProfilesStartToFinish(PVScreen, True)
        Else
            Call RapidReadPVData(PVFrameNo)
            ClearLineScreen.PVScreen.AutoRedraw = False
            ClearLineScreen.PVScreen.Cls 'PCNGL200103
            Call DrawPVProfile(ClearLineScreen.PVScreen, 1) 'Goto frame PVFrameNo PCN3526
        End If
    ElseIf CLPScreenMode = Video Then
        If AVIInitialised Then
           Call SeekTime(PVTimes(FrameNo))  'PCNGL160103 'PCN2639
           Call MoveSlider(PVTimes(FrameNo))      'PCN2639
        End If
    ElseIf CLPScreenMode = ThreeD Then
        If ThreeDRunning = True And CLPScreenMode = ThreeD Then Call d3d_laser_focus(PVFrameNo)
    ElseIf CLPScreenMode = SnapShot Then
        Call Me.TakeASnapShot
    End If
End If


Call PVGraphsKeyForm.PVGraphsKeyUpdate 'PCN2990

Call ClearLineScreen.SetPVFrameStatus 'PCN4171

PVGraphsKeyForm.ZOrder 0 'PCN4320
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS145:" & Error$
End Sub


Function RecordPVProfile_Start_Old() 'PCN4171

End Function

Sub RecordPVProfile_Start(ErrorStr As String)  'PCN4171
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : RecordPVProfile_Start
'Created : 11 June 2004, PCN2884
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Sets up application to start recording PVD.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    
'Move PV Marker to right place
Call ProfilerPause 'Pause profiler video while setting things up

'If Len(PipelineDetails.InternalDiameterExpected) = 0 Then 'PCNGL301202
'    PipelineDetails.ZOrder 0 'PCNGL060103
'    'ErrorStr = "Please enter internal pipe diameter first and calibrate the image."
'    ErrorStr = "Please enter internal pipe diameter." 'PCN4308
'    MsgBox DisplayMessage(ErrorStr), vbExclamation 'PCN1907 'PCN2111
'    PipelineDetails.ZOrder (0) 'PNN1865 "PNSLS200203
'    PipelineDetails.InternalDiameterExpected.SetFocus 'PCNGL0601032
'    Exit Function
'End If

If PipelineDetails.InternalDiameterExpected.text = 0 Then PipelineDetails.InternalDiameterExpected.text = "" 'PCN4512

'Check that the Start and Finish distances have been entered
If ConfigInfo.DistanceStart = InvalidData And Not IPD Then
    Call DataEntryForm.SetUpStartFinishDistances("Start")
    ErrorStr = "The Start Distances are required."
    'MsgBox DisplayMessage(ErrorStr), vbExclamation 'PCN
    Exit Sub
ElseIf ConfigInfo.DistanceFinish = InvalidData And Not IPD Then
    Call DataEntryForm.SetUpStartFinishDistances("Finish")
    ErrorStr = "The Finish Distances are required."
    'MsgBox DisplayMessage(ErrorStr), vbExclamation
    Exit Sub
End If


        
'PCN3031 disable fisheye becuase recording has started
Call FisheyeFunctions.DisableFishEye

'Disallowing data entry into Pipeline Details fields while recording
Call PipeDetailsDataEntrySetup(False) 'PCN2884
'vvvv PCN2884 ***************************************************
'Enable other PV buttons
'Call SetupMTButtonsForVideo
Call ControlsScreen.ControlsViewSetup 'PCN4171
'^^^^ ***********************************************************
'vvvv PCN2639 **********************************************

If DistanceMethod = "AutomaticCounter" Then
 
    If CountDirection <> "Up" Then
        CountDirection = "Down"
    End If
    'Store current Distance Process Method in ConfigInfo
    ConfigInfo.DistanceProcessMethod = DistanceMethod
    'vvvv PCN4171 *****************************************
''    If isopen("Distance") Then
''        If Len(Distance.txtStartDistance) = 0 Then
''            DistanceStart = -1
''            ConfigInfo.DistanceStart = DistanceStart
''            ConfigInfo.DistanceDirection = CountDirection
''            ConfigInfo.DistanceFinish = -1 'Not required
''        Else
''            On Error GoTo DistanceError
''            DistanceStart = CDbl(Distance.txtStartDistance) 'PCN2898 removed the function Val eg was CDbl(Val(txtStartDistance)
''            On Error GoTo Err_Handler
''        End If
''    End If
    '^^^^ *************************************************
    
    'vvv PCN4171 *************************
    If IPD Then
        Call CheckForIPD
    End If
    If DistanceStart <> InvalidData Then
        Call DistanceCounterSet(DistanceStart) 'PCN2874 Also fogot to copy over chnages from 5.4.2
    End If
    '^^^^ ********************************
Else
    DistanceMethod = "StartFinishEstimate"
    ConfigInfo.DistanceProcessMethod = "StartFinishEstimate"
    ConfigInfo.DistanceDirection = ""
End If
'^^^^ ******************************************************

'Setup for Drawing the PV Profile 'PCNGL220103
If PicInPicMode = PV Then  'PCN1863
    
    Call DrawPVProfile_Setup(ClearLineScreen.PVScreenPicInPic) 'PCNGL220103 'PCN1863 PCN3526
Else
    Call DrawPVProfile_Setup(ClearLineScreen.PVScreen) 'PCNGL220103 PCN3526
End If

PrecisionVisionGraph.ZOrder 0 'PCNGL060103
If ImageGraphState(0).GraphType <> "" Then
    Call PrecisionVisionGraph.SetupPVGraphScreen(ImageGraphState(0).GraphType)
Else
    Call PrecisionVisionGraph.SetupPVGraphScreen("Capacity")
End If

If IsOpen("AutoTune") Then
    Call SetCheckBoxTick(AutoTune.ShowSettings, False)
End If

Call DeleteAll 'make sure the centre is centred before recording
Me.ProfilerTimer.Interval = 200 'The Timer needs to update ocasionally so no to spend so much time
                                'checking if its finnished. But when stoped needs to be reset to 200

Call SetupVideoDisplayAsNormal
'^^^^ **************************************

Call GotoStartMarker 'VideoFrameSliderSetForStartRecord 'PCN2930

OptionsPage.FishEyeCameraDropdown.Enabled = False
OptionsPage.SmoothCheck.Enabled = False
ReDim GraphInfoContainer(PVOvalitySmooth).DataSingle(0) 'PCN9999
ReDim GraphInfoContainer(PVXDiameterSmooth).DataSingle(0)
ReDim GraphInfoContainer(PVYDiameterSmooth).DataSingle(0)
ReDim GraphInfoContainer(PVMedianDiameterSmooth).DataSingle(0) 'PCN9999
ReDim GraphInfoContainer(PVMaxDiameterSmooth).DataSingle(0)
ReDim GraphInfoContainer(PVMinDiameterSmooth).DataSingle(0)
ReDim GraphInfoContainer(PVCapacitySmooth).DataSingle(0)
'ReDim GraphInfoContainer(PVInclinationSmooth).DataSingle(0) 'PCN6128 'PCN6458



Call emptybuffer 'PCN3178 The buffer in the C has to be cleared to make sure the frame it starts with is the proper starting place
NumTimesRecorded = 1 'Reset numoftimesrecored to 1
Call hough_processimageonoff(True)
Call setrecprofstat(1) 'Tell CPP to store video processing data
Call ProfilerRun
PVRecording = True



Exit Sub
'vvvv PCN2639 *******************
DistanceError:
    DistanceStart = -1
    Resume Next
'^^^^ ***************************
Err_Handler:
Select Case Err
    Case 13 'Invalid data
        DistanceStart = -1
        Resume Next
    Case Else
        MsgBox Err & "-CLS146:" & Error$
End Select
End Sub



Sub RecordPVProfile_Stop()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : RecordPVProfile_Stop
'Created : 11 June 2004, PCN2884
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Stops PVD recording and sets up application for post recording.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Dim i As Double
Dim SaveFailed As Boolean
Call setrecprofstat(0)
Call ProfilerPause
PVRecording = False
Call GetBlockOfPVData 'PCN2891
If PVDataNoOfLines < 2 Then Exit Sub

Me.ProfilerTimer.Interval = 40  'The timer when not recording needs to update faster so that controls
                                'and the PVD line on the graphs follow the video quicker
    
LastRecordedFrame = PVFrameNo  'PCNLS310103

Call PipeDetailsDataEntrySetup(True) 'PCN2884

'vvvv PCN4171 *****************************************************************
Call ControlsScreen.ControlsViewSetup 'PCN4171

'vvvv PCN2930 ********************************
VideoFrameSlider.MarkerStop = VideoFrameSlider.MarkerPosition - VideoFrameSlider.FrameLeft
Call VideoFrameSliderMarkerSetup
Call VideoFrameSliderRedraw
Call hough_processimageonoff(False)
'^^^^ ****************************************

DoEvents
'Check to see if there was water level set, and if there was then finetune the centre and data
'calculations
'If WLStartAngle <> 0 Or WLFinishAngle <> 0 Then

Call CLPProgressBar.ProgressBarInitialise(DisplayMessage("Processing Data"))
DoEvents
For i = 0 To (PVDataNoOfLines / 100)
    Call CLPProgressBar.ProgressBarPosition((i / PVDataNoOfLines * 100))
    DoEvents
    
    Call ScreenDrawing.PVCentreCalcCPP((i * 100), 99 + (i * 100)) 'PCN3194
    Call ScreenDrawing.PVDiameterMedianCalcCPP((i * 100), 99 + (i * 100)) 'PCN3540 'PCN4974 need to calculate median diameter incase its used with diameter flat.
    Call ScreenDrawing.PVFlat3DCalcCPP((i * 100), 99 + (i * 100)) 'PCN3513
    Call ScreenDrawing.PVCapacityCalcCPP((i * 100), 99 + (i * 100)) 'PCN3540
    Call ScreenDrawing.PVXYDiameterCalcCPP((i * 100), 99 + (i * 100)) 'PCN3540
    Call ScreenDrawing.PVDeltaMaxMinCalcCPP((i * 100), 99 + (i * 100)) 'pcn3540
    Call ScreenDrawing.PVDiameterMaxMinCalcCPP((i * 100), 99 + (i * 100)) 'PCN3540
    Call ScreenDrawing.FixMinMax((i * 100), 99 + (i * 100)) 'PCN6524
'    Call ScreenDrawing.PVInclinationCalc((I * 100), 99 + (I * 100)) 'PCN6128'PCN6458

'    Call ScreenDrawing.PVFractileCalcCPP((i * 100), 99 + (i * 100)) 'PCN3540
    Call ScreenDrawing.PVOvalityCalcCPP((i * 100), 99 + (i * 100)) 'PCN3540
    'Call ScreenDrawing.PVDebrisCalcCPP((i * 100), 99 + (i * 100)) 'PCN4461
Next i
 'Call ScreenDrawing.AdjustLayToFit
'PCN4296

'PCN6458 Call CPPSmoothInclination
Call ScreenDrawing.CPPFilterGraphs 'PCN4355

Call PageFunctions.SaveCentreCalculations
Call CLPProgressBar.ProgressBarPosition(1)

'This is to fix the time stamp, its out by one frame, and needs to be shifted up, has to be
'done here because its allready a legacy and is saved like this on recording, it is fixed on
'loading, but of course this hasn't yet been loaded yet.

Dim FrameNo As Long
For FrameNo = 1 To PVDataNoOfLines
    PVTimes(FrameNo - 1) = PVTimes(FrameNo)
Next FrameNo
PVTimes(PVDataNoOfLines) = 0
Call PageFunctions.FixTimeStampErrors
PipelineDetails.InternalDiameterExpected.text = Round(ExpectedDiameter, 1) 'PCNANT
Call PVGraphsKeyForm.DisplayPVGraphsKey 'PCN4339

OptionsPage.FishEyeCameraDropdown.Enabled = False
OptionsPage.SmoothCheck.Enabled = True

    If OptionsPage.SmoothCheck.value = vbChecked Then
        ReDim GraphInfoContainer(PVOvalitySmooth).DataSingle(PVDataNoOfLines) 'PCN9999
        ReDim GraphInfoContainer(PVXDiameterSmooth).DataSingle(PVDataNoOfLines)
        ReDim GraphInfoContainer(PVYDiameterSmooth).DataSingle(PVDataNoOfLines)
        ReDim GraphInfoContainer(PVMedianDiameterSmooth).DataSingle(PVDataNoOfLines) 'PCN9999
        ReDim GraphInfoContainer(PVMaxDiameterSmooth).DataSingle(PVDataNoOfLines)
        ReDim GraphInfoContainer(PVMinDiameterSmooth).DataSingle(PVDataNoOfLines)
        ReDim GraphInfoContainer(PVCapacitySmooth).DataSingle(PVDataNoOfLines)
'PCN6458       ReDim GraphInfoContainer(PVInclinationSmooth).DataSingle(PVDataNoOfLines) 'PCN6128
        

        
        Call CPPSmoothGraphs
        ' PCN5186 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Call ScreenDrawing.ClearAllGraphsAndRuler                             '
        ' Call ScreenDrawing.DrawPVGraphs 'PCN3373                              '
        ' Call ScreenDrawing.DrawPVXScale                                       '
        ' Call ScreenDrawing.PVDeflectionCalcCPP(0, PVDataNoOfLines) 'PCN5186   '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    Else
        ReDim GraphInfoContainer(PVOvalitySmooth).DataSingle(0) 'PCN9999
        ReDim GraphInfoContainer(PVXDiameterSmooth).DataSingle(0)
        ReDim GraphInfoContainer(PVYDiameterSmooth).DataSingle(0)
        ReDim GraphInfoContainer(PVMedianDiameterSmooth).DataSingle(0) 'PCN9999
        ReDim GraphInfoContainer(PVMaxDiameterSmooth).DataSingle(0)
        ReDim GraphInfoContainer(PVMinDiameterSmooth).DataSingle(0)
        ReDim GraphInfoContainer(PVCapacitySmooth).DataSingle(0)
'PCN6458        ReDim GraphInfoContainer(PVInclinationSmooth).DataSingle(0) 'PCN6128

        ' PCN5186 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Call ScreenDrawing.ClearAllGraphsAndRuler                             '
        ' Call ScreenDrawing.DrawPVGraphs 'PCN3373                              '
        ' Call ScreenDrawing.DrawPVXScale                                       '
        ' Call ScreenDrawing.PVDeflectionCalcCPP(0, PVDataNoOfLines) 'PCN5186   '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    End If
    'PCN5186 need to recalculate deflective values

   
    'Call ScreenDrawing.PVDeflectionCalcCPP   'PCN5186
'PCN6458     Call ScreenDrawing.UpDateInclinationGraph 'PCN6313
    Call ScreenDrawing.ClearAllGraphsAndRuler
    Call ScreenDrawing.DrawPVGraphs 'PCN3373
    Call ScreenDrawing.DrawPVXScale
    'Call ScreenDrawing.PVDeflectionCalcCPP   'PCN5186
    
Dim RoundingDia As Integer
If MeasurementUnits = "mm" Then RoundingDia = 1
If MeasurementUnits <> "mm" Then RoundingDia = 2

    If PipelineDetails.InternalDiameterExpected.text = "" Then
        ExpectedDiameter = CalculatePercentile(50, "MedianDiameter")
        ExpectedDiameter = Round(ExpectedDiameter, RoundingDia)
    End If
    
    PipelineDetails.InternalDiameterExpected.text = Round(ExpectedDiameter, RoundingDia) 'PCNANT
    Call CopyPipelineFormToPipeDetails
    Call SaveToFilePipeAndConfigInfo("PipelineInfo", SaveFailed) 'PCN1768

    Call ScreenDrawing.PVFlat3DCalcCPP(1, PVDataNoOfLines)

    Call CLPProgressBar.ProgressBarPosition(1)

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS147:" & Error$
End Sub

Sub LiveVideoConnect()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : LiveVideoConnect
'Created : 11 June 2004, PCN2884
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Live video connect setup.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    
If mediatype <> "" Then
    DoEvents 'PCN2877 speed up uninitialise (10 June 2004, Antony van Iersel) don't really no if it helps. seems to.
    Call uninitialise
    Call UnitVideoSlider
End If
If ThreeDRunning = True Then
    Unload3D
End If
Dim success As Boolean
success = InitLiveFeed
If success <> True Then 'PCNLS300103
    Exit Sub
End If
'Complete initialisation 'PCNGL220103
VideoFileName = ""
'Reset PV data and screens
Call PrecisionVisionGraph.ResetPVData
Call ClearLineScreen.Initialization 'PCNLS
Call InitilisePVProfile(1) 'PCNLS
''Call ClearLineScreen.ResetRecord 'PCN1792 'PCNLS
PipelineDetails.ZOrder (0) 'PCNML020503
'Set MainScreen for video
ClearLineScreen.PVScreen.Visible = False
'Disable Record button 'PCN2831
''ClearLineScreen.ControlToolbar.Buttons.Item(5).Enabled = False  'PCN2831
'Disable AVI Play buttons and make the PV Play button invisible
''ClearLineScreen.ControlToolbar.Buttons.Item(9).Visible = True 'PCNGL270103 'PCN2681
''ClearLineScreen.ControlToolbar.Buttons.Item(10).Visible = True 'PCNGL270103 'PCN2681
''ClearLineScreen.ControlToolbar.Buttons.Item(11).Visible = True 'PCNGL270103 'PCN2681
''ClearLineScreen.ControlToolbar.Buttons.Item(9).Enabled = False  'PCNGL270103 'PCN2681
''ClearLineScreen.ControlToolbar.Buttons.Item(10).Enabled = False 'PCNGL270103 'PCN2681
''ClearLineScreen.ControlToolbar.Buttons.Item(11).Enabled = False 'PCNGL270103 'PCN2681
''ClearLineScreen.ControlToolbar.Buttons.Item(6).Visible = False 'PCN2681
''ClearLineScreen.ControlToolbar.Buttons.Item(7).Visible = False 'PCN2681
''ClearLineScreen.ControlToolbar.Buttons.Item(8).Visible = False 'PCN2681
'                ClearLineScreen.ConfigToolBar1.Buttons.Item(3).Enabled = False 'PCN2295 'PCN2639
''ClearLineScreen.ConfigToolBar1.Enabled = True  'PCN2295 'PCN2639 'PCN2831
Call ClearLineScreen.SetupMainScreenForVideo

ClearLineScreen.VideoScreen.AutoRedraw = True
ClearLineScreen.VideoScreen.Visible = True
''ClearLineScreen.ControlToolbar.Buttons.Item(1).Image = 29 'Connected PCNGL270103 'PCN2681
PVDFileName = "" 'PCN2464
ConfigInfo.FishEyeDistortion = 0 'PCN3039 have to set this to -1 to force a fisheye calculation
Call FisheyeFunctions.SetDistortion(Fisheye.TFactor.value) 'PCN3039 Even thou the fish eye was set the mask was not yet
                                                      ' created, regardless if you chose to set fisheye or not.
    Fisheye.Visible = False 'PCN4243, when the above is called, because it calls gets the Fisheye.TFactor value it makes the form visible, not suppose to be, not used anymore
Call FishEyeLoadFileCheck(Video) 'PCN2527

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS148:" & Error$
End Sub

Sub LiveVideoDisconnect()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : LiveVideoDisconnect
'Created : 11 June 2004, PCN2884
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Live video disconnect setup.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    
'Then disconnect live video feed
Call UnitLiveFeed
''ClearLineScreen.ControlToolbar.Buttons.Item(1).Image = 1 'Disconnected PCNGL270103 'PCN2681
'Enable AVI Play buttons
''ClearLineScreen.ControlToolbar.Buttons.Item(9).Enabled = True  'PCNGL270103 AVIPlayPause 'PCN2681
''ClearLineScreen.ControlToolbar.Buttons.Item(10).Enabled = True 'PCNGL270103 AVIStepForward 'PCN2681
''ClearLineScreen.ControlToolbar.Buttons.Item(11).Enabled = True 'PCNGL270103 AVIStepRewind 'PCN2681
'''                ClearLineScreen.ConfigToolBar1.Buttons.Item(3).Enabled = True 'PCN2295 'PCN2639
''ClearLineScreen.ConfigToolBar1.Enabled = True 'PCN2295 'PCN2639
'''Disable Record button 'PCN2831
''ClearLineScreen.ControlToolbar.Buttons.Item(5).Enabled = False  'PCN2831
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS149:" & Error$
End Sub

Sub PipeDetailsDataEntrySetup(ToEnable As Boolean)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PipeDetailsDataEntrySetup
'Created : 11 June 2004, PCN2884
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : .
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    
'PCN1987 - Disallowing data entry into Pipeline Details fields while recording
PipelineDetails.AssetNo.Enabled = ToEnable
PipelineDetails.City.Enabled = ToEnable
PipelineDetails.SiteID.Enabled = ToEnable
PipelineDetails.sDate.Enabled = ToEnable
PipelineDetails.sTime.Enabled = ToEnable
PipelineDetails.StartNodeNo.Enabled = ToEnable
PipelineDetails.StartNodeLocation.Enabled = ToEnable
PipelineDetails.FinishNodeNo.Enabled = ToEnable
PipelineDetails.FinishNodeLocation.Enabled = ToEnable
PipelineDetails.InternalDiameterExpected.Enabled = ToEnable
PipelineDetails.OutsideDiameter.Enabled = ToEnable
PipelineDetails.PipeLength.Enabled = ToEnable
PipelineDetails.Material.Enabled = ToEnable
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS150:" & Error$
End Sub

Sub DistanceCounterTextUpdate()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : DistanceCounterTextUpdate
'Created : 20 June 2004, PCN2891
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Tests to see if the PVD has been setup with a Distance Method.
'          If it has, DistanceTracker.text and Distance.CalculatedDist
'          are updated with the current Distance Counter setting.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim distanceTest As Long 'PCN2928
    
'vvvv PCN2966 **************************************
'PCN3744 add to the conditiona below vvvvvvvvvv IPD = True

'If (ConfigInfo.DistanceStart >= 0 Or IPD = True) And DistanceMethod = "AutomaticCounter" And (PVDataNoOfLines < 2 Or PVRecording) Then 'PCN2928 'PCN2966
' PCN3884 was >=0 now > -1000
If (ConfigInfo.DistanceStart > -1000 Or IPD = True) And DistanceMethod = "AutomaticCounter" And (PVDataNoOfLines < 2 Or PVRecording) Then 'PCN2928 'PCN2966
    If IsOpen("Distance") Then
        Call getcounter(distanceTest)
        'PCN4171
''        'vvvv PCN2928 *****************************************
''        If MeasurementUnits = "mm" Then
''            'PCN3744 added IPD
''            If IPD = True Then Distance.CalculatedDist = Format(distanceTest / 100, "#0.00") 'PCNAVI060804 forgot to copy / 10
''            If IPD = False Then Distance.CalculatedDist = Format(distanceTest / 10, "#0.00") 'PCNAVI060804 forgot to copy / 10
''
''        Else
''            'PCN3744 added IPD
''            If IPD = True Then Distance.CalculatedDist = Format(distanceTest / 10, "#0.00")
''            If IPD = False Then Distance.CalculatedDist = Format(distanceTest, "#0.00")
''        End If
''        '^^^^ *************************************************
    End If


'vvvv PCN2639 **********************************
'ElseIf ConfigInfo.DistanceStart >= 0 And DistanceMethod <> "None" And PVDataNoOfLines > 1 Then 'PCN2928 'PCN2966
' PCN3884 was >=0 now > -1000
ElseIf ConfigInfo.DistanceStart > -1000 And DistanceMethod <> "None" And PVDataNoOfLines > 1 Then 'PCN2928 'PCN2966
'PCN4171
''    If isopen("Distance") Then
''        'vvvv PCN2928 *****************************************
''        If MeasurementUnits = "mm" Then
''            Distance.CalculatedDist = Format(PVDistances(PVFrameNo), "#0.00") 'PCNAVI060804 forgot to copy / 10
''        Else
''            Distance.CalculatedDist = Format(PVDistances(PVFrameNo), "#0.00")
''        End If
''        '^^^^ *************************************************
''    End If
'^^^^ ******************************************

End If
'^^^^ **********************************************
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS151:" & Error$
End Sub

Sub VideoScreenScaleCalc()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : VideoScreenScaleCalc
'Created : 20 June 2004, PCN2891
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Calculates the VideoScreenScale based on what fits in a 3/4 ratio (MainScreen.height / MainScreen.width).
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim CheckAspectRatio As Double
    
CheckAspectRatio = MainScreen.height / MainScreen.width 'PCN1863
If CheckAspectRatio <= 0 Or VideoAspectRatio <= 0 Then
    VideoScreenScale = 1
ElseIf ConfigInfo.MediaWidth = 0 Or ConfigInfo.MediaHeight = 0 Then
    VideoScreenScale = 1
ElseIf CheckAspectRatio >= VideoAspectRatio Then
    VideoScreenScale = MainScreen.width / ConfigInfo.MediaWidth 'PCN1863
Else
    VideoScreenScale = MainScreen.height / ConfigInfo.MediaHeight  'PCN1863
End If

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS152:" & Error$
End Sub

Sub VideoFrameSliderSetup()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : VideoFrameSliderSetup
'Created : 8 July 2004, PCN2930
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Builds the VideoFrame slider.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim PixX1 As Integer
Dim PixY1 As Integer
Dim PixX2 As Integer
Dim PixY2 As Integer
Dim ToRecordStart As Integer
Dim RecordedStart As Integer
Dim ToRecordStop As Integer
Dim RecordedStop As Integer
Dim i As Integer

FrameBlank.Visible = False 'These are just for show

With VideoFrameSlider

.FrameTop = 100
.FrameHeight = 200
.FrameLeft = 300
.Framewidth = 9000
.FrameRailHeight = 30
.FrameSpaceMajor = 100
.FrameMinorSpacing = .FrameSpaceMajor / 4
'Setup  Min and Value defaults
'.Max = .Framewidth
.Min = 0
.value = 0
'Find the values for .MarkerStart and .MarkerStop
'.MarkerStart = .Min
'.MarkerStop = .Max
Call VideoFrameSliderStartStop

PixX1 = .FrameLeft
PixX2 = .FrameLeft + .Framewidth
PixY1 = .FrameTop
PixY2 = .FrameTop

VideoFrameSliderBar.AutoRedraw = True
'Frame Blank Background
VideoFrameSliderBar.Line (.FrameLeft, .FrameTop)-(.FrameLeft + .Framewidth, .FrameTop + .FrameHeight), &HFFFFFF, BF 'White
VideoFrameSliderBar.Line (.FrameLeft, .FrameTop + 2 * .FrameRailHeight)-(.FrameLeft + .Framewidth, .FrameTop + .FrameHeight - .FrameRailHeight), &HE0E0E0, BF       'Grey


If (PVRecording Or PVDFileName = "" Or PVDataNoOfLines = 0) And (mediatype = Video Or mediatype = "Live") Then
    VideoRecordMarkerStart.Visible = True
    VideoRecordMarkerStartAdjuster.Visible = True
    VideoRecordMarkerStop.Visible = True
    VideoRecordMarkerStopAdjuster.Visible = True
    'To Record Frame Background - Yellow
    ToRecordStart = .MarkerStart + .FrameLeft
    ToRecordStop = .MarkerStop + .FrameLeft
    VideoFrameSliderBar.Line (ToRecordStart, .FrameTop + 2 * .FrameRailHeight)-(ToRecordStop, .FrameTop + .FrameHeight - .FrameRailHeight), &HFFFF&, BF 'Yellow
Else
    VideoRecordMarkerStart.Visible = False
    VideoRecordMarkerStartAdjuster.Visible = False
    VideoRecordMarkerStop.Visible = False
    VideoRecordMarkerStopAdjuster.Visible = False
    If mediatype = Video Or mediatype = "Live" Then
        'Recorded Frame Background - Red
        RecordedStart = .MarkerStart + .FrameLeft
        RecordedStop = .MarkerStop + .FrameLeft
        VideoFrameSliderBar.Line (RecordedStart, .FrameTop + 2 * .FrameRailHeight)-(RecordedStop, .FrameTop + .FrameHeight - .FrameRailHeight), &HFF&, BF  'Red
    End If
End If

Call VideoFrameSliderMarkerSetup


'Frame edging
VideoFrameSliderBar.Line (.FrameLeft, .FrameTop)-(.FrameLeft, .FrameTop + .FrameHeight), 0
VideoFrameSliderBar.Line (.FrameLeft + .Framewidth, .FrameTop)-(.FrameLeft + .Framewidth, .FrameTop + .FrameHeight), 0

VideoFrameSliderBar.Line (PixX1, PixY1)-(PixX2, PixY2), 0
PixY1 = .FrameTop + .FrameHeight
PixY2 = .FrameTop + .FrameHeight
VideoFrameSliderBar.Line (PixX1, PixY1)-(PixX2, PixY2), 0

PixY1 = .FrameTop + .FrameRailHeight
PixY2 = .FrameTop + .FrameRailHeight
VideoFrameSliderBar.Line (PixX1, PixY1)-(PixX2, PixY2), 0

PixY1 = .FrameTop + .FrameHeight - .FrameRailHeight
PixY2 = .FrameTop + .FrameHeight - .FrameRailHeight
VideoFrameSliderBar.Line (PixX1, PixY1)-(PixX2, PixY2), 0

'Frame Spacing
PixX1 = .FrameLeft + 30
PixX2 = .FrameLeft + 30
PixY1 = .FrameTop
PixY2 = .FrameTop + .FrameHeight

While PixX1 < (.FrameLeft + .Framewidth)
    VideoFrameSliderBar.Line (PixX1 + 15, PixY1 + 2 * .FrameRailHeight)-(PixX2 + 15, PixY2 - (.FrameRailHeight + 2)), &HFFFFFF
    VideoFrameSliderBar.Line (PixX1, PixY1)-(PixX2, PixY2), 0
    For i = 1 To 3
        PixX1 = PixX1 + .FrameMinorSpacing
        PixX2 = PixX2 + .FrameMinorSpacing
        If PixX1 < (.FrameLeft + .Framewidth) Then
            VideoFrameSliderBar.Line (PixX1, PixY1)-(PixX2, PixY1 + .FrameRailHeight), 0
            VideoFrameSliderBar.Line (PixX1, PixY2 - .FrameRailHeight)-(PixX2, PixY2), 0
        End If
    Next i
Wend

End With

Call VideoFrameSliderMove
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS153:" & Error$
End Sub


Sub VideoFrameSliderMove()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : VideoFrameSliderMove
'Created : 8 July 2004, PCN2930
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Moves the main position marker on the VideoFrame slider.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim MarkerCentreOffset As Integer 'Defines the offset required to ensure marker is aligned to the correct position on the slider.
Dim MarkerNewPosition As Double 'Defines position of marker in twips on the VideoFrameSliderBar

MarkerCentreOffset = 30

With VideoFrameSlider

If (.Max - .Min) > 0 Then
    MarkerNewPosition = .value * (.Framewidth / (.Max - .Min)) + .FrameLeft 'PCN2955
Else
    MarkerNewPosition = .FrameLeft
End If


VideoSliderMarker.Left = MarkerNewPosition - MarkerCentreOffset
VideoFrameSlider.MarkerPosition = MarkerNewPosition

Dim TheTime As Double

'Check to see if currently recording and at or past the Record Stop position
If PVRecording = True Then
    'If MarkerNewPosition > .MarkerStop + .FrameLeft - 100 Then '100 was negated just to make sure it stops
    'PCN4523
    If MarkerNewPosition > VideoRecordMarkerStopAdjuster.Left Then
        Call RecordPVProfile_Stop
    Else
        Call VideoFrameSliderRedraw
    End If
    TheTime = GetCurrentAVITime
    
    If TheTime > (AVITime - 0.08) Then Call RecordPVProfile_Stop
End If
    
End With


Call FormatTime
'^^^^ ***************************************************************************
   
Exit Sub
Err_Handler:
    Select Case Err
        Case 6: Resume Next
        Case Else: MsgBox Err & "-CLS154:" & Error$
    End Select
End Sub

Sub VideoFrameSliderRedraw()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : VideoFrameSliderRedraw
'Created : 8 July 2004, PCN2930
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Redraws the slider.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim PixX1 As Integer
Dim PixY1 As Integer
Dim PixX2 As Integer
Dim PixY2 As Integer
Dim i As Integer

With VideoFrameSlider

VideoFrameSliderBar.AutoRedraw = True
'Frame Blank Background
'VideoFrameSliderBar.Line (.FrameLeft, .FrameTop)-(.FrameLeft + .Framewidth, .FrameTop + .FrameHeight), &HFFFFFF, BF 'White
VideoFrameSliderBar.Line (.FrameLeft, .FrameTop + 2 * .FrameRailHeight)-(.FrameLeft + .Framewidth - 20, .FrameTop + .FrameHeight - .FrameRailHeight - 20), &HE0E0E0, BF       'Grey

'To Record Frame Background - Yellow
VideoFrameSliderBar.Line (.MarkerStart + .FrameLeft, .FrameTop + 2 * .FrameRailHeight)-(.MarkerStop + .FrameLeft, .FrameTop + .FrameHeight - .FrameRailHeight - 20), &HFFFF&, BF  'Yellow
If PVRecording Then
    'Recorded Frame Background - Red
    VideoFrameSliderBar.Line (.MarkerStart + .FrameLeft, .FrameTop + 2 * .FrameRailHeight)-(.MarkerPosition, .FrameTop + .FrameHeight - .FrameRailHeight - 20), &HFF&, BF  'Red
ElseIf PVDFileName <> "" Then
    'Recorded Frame Background - Red
    VideoFrameSliderBar.Line (.MarkerStart + .FrameLeft, .FrameTop + 2 * .FrameRailHeight)-(.MarkerStop + .FrameLeft, .FrameTop + .FrameHeight - .FrameRailHeight - 20), &HFF&, BF  'Red
End If
'Frame Spacing
PixX1 = .FrameLeft + 30
PixX2 = .FrameLeft + 30
PixY1 = .FrameTop
PixY2 = .FrameTop + .FrameHeight

While PixX1 < (.FrameLeft + .Framewidth)
    VideoFrameSliderBar.Line (PixX1 + 15, PixY1 + 2 * .FrameRailHeight)-(PixX2 + 15, PixY2 - (.FrameRailHeight + 2)), &HFFFFFF
    VideoFrameSliderBar.Line (PixX1, PixY1)-(PixX2, PixY2), 0
    For i = 1 To 3
        PixX1 = PixX1 + .FrameMinorSpacing
        PixX2 = PixX2 + .FrameMinorSpacing
    Next i
Wend

End With

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS155:" & Error$
End Sub

Sub VideoFrameSliderMarkerSetup()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : VideoFrameSliderMarkerSetup
'Created : 9 July 2004, PCN2930
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Sets up the position of the Start and Stop markers on the VideoFrame slider.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

With VideoFrameSlider

'Adjust Start markers
VideoRecordMarkerStart.x1 = .FrameLeft + .MarkerStart
VideoRecordMarkerStart.x2 = VideoRecordMarkerStart.x1
VideoRecordMarkerStartAdjuster.Left = VideoRecordMarkerStart.x1 - VideoRecordMarkerStartAdjuster.width
'Adjust Stop markers
VideoRecordMarkerStop.x1 = .FrameLeft + .MarkerStop
VideoRecordMarkerStop.x2 = VideoRecordMarkerStop.x1
VideoRecordMarkerStopAdjuster.Left = VideoRecordMarkerStop.x1
    
End With
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS156:" & Error$
End Sub

Sub VideoFrameSliderStartStop()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : VideoFrameSliderStartStop
'Created : 9 July 2004, PCN2930
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Determines the value of the MarkerStart and MarkerStop markers on the VideoFrame slider.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim MarkerNewValue As Double 'Defines value of marker in twips on the VideoFrameSliderBar
Dim TimeStart As Double
Dim TimeStop As Double

With VideoFrameSlider


'Determine the Start marker value
On Error GoTo SetDefaultValues
TimeStart = PVTimes(1)
'Determine the Stop marker value
TimeStop = PVTimes(PVDataNoOfLines - 1)

If TimeStop <> 0 And (TimeStop - TimeStart) > 0 And (.Max - .Min) > 0 Then
    MarkerNewValue = 100 * TimeStart * (.Framewidth / (.Max - .Min))
    .MarkerStart = MarkerNewValue
    MarkerNewValue = 100 * TimeStop * (.Framewidth / (.Max - .Min))
    .MarkerStop = MarkerNewValue
Else
    GoTo SetDefaultValues
End If

End With


Exit Sub
SetDefaultValues:
    VideoFrameSlider.MarkerStart = 0
    VideoFrameSlider.MarkerStop = VideoFrameSlider.Framewidth
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS557:" & Error$
End Sub
Sub GotoStartMarker()
On Error GoTo Err_Handler
Dim MarkerCentreOffset As Integer 'Defines the offset required to ensure marker is aligned to the correct position on the slider.
Dim MarkerNewPosition As Double 'Defines position of marker in twips on the VideoFrameSliderBar

MarkerCentreOffset = 30

With VideoFrameSlider
    MarkerNewPosition = .MarkerStart + .FrameLeft
    .value = (.Max - .Min) * (.MarkerStart / .Framewidth)  'Assume .Framewidth <> 0 'PCN2955
End With

VideoSliderMarker.Left = MarkerNewPosition - MarkerCentreOffset
VideoFrameSlider.MarkerPosition = MarkerNewPosition

Call VideoFrameSliderVideoScroll

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS158:" & Error$
End Sub

Sub GotoStopMarker()
On Error GoTo Err_Handler
Dim MarkerCentreOffset As Integer 'Defines the offset required to ensure marker is aligned to the correct position on the slider.
Dim MarkerNewPosition As Double 'Defines position of marker in twips on the VideoFrameSliderBar

MarkerCentreOffset = 30

With VideoFrameSlider
    MarkerNewPosition = .MarkerStop + .FrameLeft
    .value = (.Max - .Min) * (.MarkerStop / .Framewidth)  'Assume .Framewidth <> 0 'PCN2955
End With

VideoSliderMarker.Left = MarkerNewPosition - MarkerCentreOffset
VideoFrameSlider.MarkerPosition = MarkerNewPosition

Call VideoFrameSliderVideoScroll

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS159:" & Error$
End Sub

Sub VideoFrameSliderVideoScroll()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : VideoFrameSliderVideoScroll
'Created : 12 July 2004, PCN2930
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Scrolls the video after moving the Video Frame Slider.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

'Scroll video
SliderFrame = VideoFrameSlider.value
Call SeekTime(VideoFrameSlider.value / 100) 'approxframeseek(SliderFrame)
'Get the videoframeno
'VideoFrame = getcurrentframe PCN3289 not to be used (3 Feb 2005,Antony)

Call FormatTime


Exit Sub
Err_Handler:
    MsgBox Err & "-CLS160:" & Error$
End Sub


Sub SetTextBoxWidthAndHeight(OnScreenText As TextBox)
On Error GoTo Err_Handler


Dim TextHeight As Integer
Dim TextWidth As Integer
Dim TextString As String

TextString = OnScreenText.text + "WW"

OnScreenTextBoxPicture.Font = OnScreenText.Font ' copy the font to the picturebox
OnScreenTextBoxPicture.FontSize = OnScreenText.FontSize ' copy the fontsize

MainScreen.Font = OnScreenText.Font
MainScreen.FontBold = OnScreenText.FontBold
MainScreen.FontItalic = OnScreenText.FontItalic
MainScreen.FontName = OnScreenText.FontName
MainScreen.FontSize = OnScreenText.FontSize
MainScreen.FontStrikethru = OnScreenText.FontStrikethru
MainScreen.FontUnderline = OnScreenText.FontUnderline


TextHeight = MainScreen.TextHeight(TextString)  ' figure the text height
TextWidth = MainScreen.TextWidth(TextString)  ' figure the text width

OnScreenText.width = TextWidth  ' set labels width
OnScreenText.height = TextHeight ' set labels height

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS161:" & Error$
End Sub



Public Sub MagnifyUpdateImage()
On Error GoTo Err_Handler
    Dim SnapToMagScale As Single
    
    Dim MagScaleLeft As Single
    Dim MagScaleTop As Single
    Dim MagScaleCentreX As Single
    Dim MagScaleCentreY As Single
    Dim MagScaleWidth As Single
    Dim MagScaleHeight As Single
    Dim MagScaleBottom As Single
    Dim MagScaleRight As Single
    
    Dim ShiftX As Single
    Dim ShiftY As Single
    
    Dim MagnifierCentreX As Single
    Dim MagnifierCentreY As Single
    
    MagnifyScale = Round(MagnifyScale, 1)
    Me.MagnifyLabel = "x" & MagnifyScale
    
    MagScaleWidth = (Me.DummyScreen.width / Me.MainScreen.width * Me.MagnifyPictContainer.width) / MagnifyScale
    MagScaleHeight = (Me.DummyScreen.height / Me.MainScreen.height * Me.MagnifyPictContainer.height) / MagnifyScale
        

    ShiftX = Me.MagnifyPictContainer.width - (Me.MagnifyPictContainer.width / MagnifyScale)
    ShiftY = Me.MagnifyPictContainer.height - (Me.MagnifyPictContainer.height / MagnifyScale)
    ShiftX = ShiftX * (Me.MagnifyPictContainer.Left / (Me.MainScreen.width - Me.MagnifyPictContainer.width))
    ShiftY = ShiftY * (Me.MagnifyPictContainer.Top / (Me.MainScreen.height - Me.MagnifyPictContainer.height))
    
    MagScaleLeft = (Me.MagnifyPictContainer.Left + ShiftX) / (Me.MainScreen.width) * Me.DummyScreen.width
    MagScaleTop = (Me.MagnifyPictContainer.Top + ShiftY) / (Me.MainScreen.height) * Me.DummyScreen.height
    
    Call MagnifyImage.PaintPicture(DummyScreen.Picture, _
                                  0, 0, _
                                  MagnifyImage.width, MagnifyImage.height, _
                                  MagScaleLeft, MagScaleTop, _
                                  MagScaleWidth, MagScaleHeight)

Exit Sub
Err_Handler:
    Select Case Err
        Case 481: Me.MagnifyPictContainer.Visible = False: Exit Sub 'There is not picture to magnify
    Case Else
        MsgBox Err & "-CLS162:" & Error$
    End Select
End Sub

Function MagnifyToSnapshotX(ByVal X As Single) As Single
On Error GoTo Err_Handler
    Dim ShiftX As Single

    X = X / (Me.MagnifyImage.width + 2)
    ShiftX = Me.MagnifyPictContainer.width - (Me.MagnifyPictContainer.width / MagnifyScale)
    ShiftX = ShiftX * (Me.MagnifyPictContainer.Left / (Me.MainScreen.width - Me.MagnifyPictContainer.width))

    MagnifyToSnapshotX = Me.MagnifyPictContainer.width / MagnifyScale * X
    MagnifyToSnapshotX = Me.MagnifyPictContainer.Left + MagnifyToSnapshotX + ShiftX
      
 
Exit Function
Err_Handler:
    Select Case Err
    Case Else

        MsgBox Err & "-CLS163:" & Error$
    End Select
End Function

Function MagnifyToSnapshotY(ByVal Y As Single) As Single
On Error GoTo Err_Handler
    Dim ShiftY As Single

    Y = Y / (Me.MagnifyImage.height + 3)
    ShiftY = Me.MagnifyPictContainer.height - (Me.MagnifyPictContainer.height / MagnifyScale)
    ShiftY = ShiftY * (Me.MagnifyPictContainer.Top / (Me.MainScreen.height - Me.MagnifyPictContainer.height))
    
    MagnifyToSnapshotY = Me.MagnifyPictContainer.height / MagnifyScale * Y
    MagnifyToSnapshotY = Me.MagnifyPictContainer.Top + MagnifyToSnapshotY + ShiftY
    
    
Exit Function
Err_Handler:
    Select Case Err
    Case Else
        MsgBox Err & "-CLS164:" & Error$
    End Select
End Function

Sub UpdateDimensionRuler(ByVal value As Double)
On Error GoTo Err_Handler
Dim answer As String
Dim answerInPercent As String 'PCN4248 it has to check to see if there is a Internal Diameter if not
                              ' then dont display the percentage, at the moment it goes 100%
Dim AsPercentage As Double
Dim NewValue As Double

    NewValue = Abs(value)
    
    'PCN4248 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If ExpectedDiameter > 0 Then                                    '
        AsPercentage = 100 + ConvertRealToPer(NewValue, "Dia")      '
        answerInPercent = Format(AsPercentage, "#00.0") & "%"       '
    Else                                                            '
        answerInPercent = ""                                        '
    End If                                                          '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    answer = Format(NewValue, "#00.0") & MeasurementUnits & " " & answerInPercent 'Format(AsPercentage, "#00.0") & "%"
    DimenMeasure.Caption = answer

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-CLS165:" & Error$
    End Select
End Sub

Sub UpdateIntegratedArea(ByVal value As Double)
On Error GoTo Err_Handler
Dim answer As String
Dim AsPercentage As Double
Dim NewValue As Double
    answer = Format(value, "#00.0") & "%"
    AreaMeasure.Caption = answer

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-CLS166:" & Error$
    End Select
End Sub


Function IsCPPVideoRunning() As Boolean
On Error GoTo Err_Handler

Dim IsVideoRunning As Long
Call hough_IsVideoRunning(IsVideoRunning)

IsCPPVideoRunning = IIf(IsVideoRunning = 1, True, False)

Exit Function
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-CLS1667:" & Error$
    End Select
End Function

Sub ProfilerStepFoward()
On Error GoTo Err_Handler
    Call ProfilerPause
    If CLPScreenMode = PV Then
        If PVFrameNo < PVDataNoOfLines Then PVFrameNo = PVFrameNo + 1: Call GotoPVProfile(PVFrameNo, True)
        
    ElseIf mediatype = Video Then
         Call videoframeadvance
         Call VideoPlaybackLoop
         If CLPScreenMode = SnapShot Then Call Me.TakeASnapShot
         Call ControlsScreen.ControlsViewSetup
    
    End If
    
    
    

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS168:" & Error$
End Sub

Sub ProfilerStepRewind()
On Error GoTo Err_Handler
    Call ProfilerPause
    If CLPScreenMode = PV Then
        If PVFrameNo > 1 Then PVFrameNo = PVFrameNo - 1: Call GotoPVProfile(PVFrameNo, True)
    ElseIf mediatype = Video Then
        Call videoframerewind
        Call VideoPlaybackLoop
        If CLPScreenMode = SnapShot Then Call Me.TakeASnapShot
        Call ControlsScreen.ControlsViewSetup
    End If
    

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS169:" & Error$
End Sub

Sub SetVideoRecordMarkerStartAdjuster()
On Error GoTo Err_Handler
Dim NextMarkerPosition As Long
Dim MarkerCentreOffset As Integer 'Defines the offset required to ensure marker is aligned to the correct position on the slider.

MarkerCentreOffset = 30

If PVRecording <> True And PVDFileName = "" Then 'Left mouse press 'Disable when recording
    'Adjust markers
    NextMarkerPosition = VideoSliderMarker.Left + MarkerCentreOffset
    If NextMarkerPosition <= VideoFrameSlider.FrameLeft Then
        VideoRecordMarkerStartAdjuster.Left = VideoFrameSlider.FrameLeft - VideoRecordMarkerStartAdjuster.width
        VideoRecordMarkerStart.x1 = VideoFrameSlider.FrameLeft
    ElseIf NextMarkerPosition <= VideoRecordMarkerStop.x1 Then
        VideoRecordMarkerStartAdjuster.Left = NextMarkerPosition - 180
        VideoRecordMarkerStart.x1 = NextMarkerPosition
    Else
        VideoRecordMarkerStartAdjuster.Left = VideoRecordMarkerStop.x1 - VideoRecordMarkerStartAdjuster.width
        VideoRecordMarkerStart.x1 = VideoRecordMarkerStop.x1
    End If
    VideoRecordMarkerStart.x2 = VideoRecordMarkerStart.x1
    With VideoFrameSlider
    .MarkerStart = VideoRecordMarkerStart.x1 - .FrameLeft 'Set Record start position
    VideoMarkerLastX = 0 'Reset position
    End With
    
    DoEvents
    Call VideoFrameSliderRedraw
    With VideoFrameSlider
    .value = (.Max - .Min) * (.MarkerStart / .Framewidth)
    End With
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS170:" & Error$
End Sub

Sub SetVideoRecordMarkerStopAdjuster()
On Error GoTo Err_Handler
Dim NextMarkerPosition As Long
Dim MarkerCentreOffset As Integer 'Defines the offset required to ensure marker is aligned to the correct position on the slider.

MarkerCentreOffset = 30

If PVRecording <> True And PVDFileName = "" Then 'Left mouse press 'Disable when recording
    'Adjust markers
    NextMarkerPosition = VideoSliderMarker.Left + MarkerCentreOffset
    If NextMarkerPosition >= (VideoFrameSlider.FrameLeft + VideoFrameSlider.Framewidth) Then
        VideoRecordMarkerStopAdjuster.Left = VideoFrameSlider.FrameLeft + VideoFrameSlider.Framewidth + 1
        VideoRecordMarkerStop.x1 = VideoFrameSlider.FrameLeft + VideoFrameSlider.Framewidth
    ElseIf NextMarkerPosition >= VideoRecordMarkerStart.x1 Then
        VideoRecordMarkerStopAdjuster.Left = NextMarkerPosition
        VideoRecordMarkerStop.x1 = NextMarkerPosition
    Else
        VideoRecordMarkerStopAdjuster.Left = VideoRecordMarkerStart.x1 + 1
        VideoRecordMarkerStop.x1 = VideoRecordMarkerStart.x1
    End If
    VideoRecordMarkerStop.x2 = VideoRecordMarkerStop.x1
    
    With VideoFrameSlider
    .MarkerStop = VideoRecordMarkerStop.x1 - .FrameLeft  'Set Record stop position
    VideoMarkerLastX = 0 'Reset position
    End With
    DoEvents
    Call VideoFrameSliderRedraw
    With VideoFrameSlider
        .value = (.Max - .Min) * (.MarkerStop / .Framewidth)  'Assume .Framewidth <> 0
    End With
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS171:" & Error$
End Sub

Sub SetDimenResultsSize(ExpandResults As Boolean)  'PCN4171
On Error GoTo Err_Handler

If ExpandResults Then
    DimenResults.width = 130 '115
Else
    DimenResults.width = 50
    DimenMeasure.Caption = ""
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS172:" & Error$
End Sub

Sub SetPVFrameStatus()
On Error GoTo Err_Handler

ClearLineScreen.PVFrameStatus.Font.Charset = LanguageCharset
ClearLineScreen.PVFrameStatus.Caption = PVFrameNo & DisplayMessage(" of ") & PVDataNoOfLines
'Change width of PVFrameStatus caption
ClearLineScreen.PVFrameStatus.width = Len(ClearLineScreen.PVFrameStatus.Caption) * 7
ClearLineScreen.PVFrameStatus.Visible = True
ClearLineScreen.PVFrameStatus.ZOrder 0

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS173:" & Error$
End Sub

Sub UpdateThreeDimColour()
'PCN:       PCN4210
On Error GoTo Err_Handler
    Call d3d_pvgraphtype("Flat3D", 0, 0)
Exit Sub
Err_Handler:
    MsgBox Err & "-CLS174:" & Error$
End Sub

Sub SnapShotForEmbededFile()
On Error GoTo Err_Handler

    
    If PVDFileName = "" Or PVRecording Then Exit Sub
    'If mediatype = "" Or PVRecording Then Exit Sub
    
    EmbededSnapshot.Cls
    
    If CLPScreenMode = ThreeD Then
        'Call ClearLineScreen.D3D_CaptureImage(WindowsTempDirectory & "CBS\EmbedFile.bmp")
        Call ClearLineScreen.D3D_CaptureImage(WindowsTempDirectory & "EmbedFile.bmp") 'ID4601
    ElseIf CLPScreenMode = SnapShot Then
        Call ScreenDrawing.DrawProfilesStartToFinish(Me.SnapShotScreen, True)
        Me.PVScreen.AutoRedraw = True
        Call OnscreenTextDrawAll(Me.SnapShotScreen, 1) 'PCN????
        'SavePicture Me.SnapShotScreen.Image, WindowsTempDirectory & "CBS\EmbedFile.bmp"
        SavePicture Me.SnapShotScreen.Image, WindowsTempDirectory & "EmbedFile.bmp" 'ID4601
    ElseIf CLPScreenMode = PV Then
        Call ScreenDrawing.DrawProfilesStartToFinish(Me.PVScreen, True)
        Me.PVScreen.AutoRedraw = True
        Call OnscreenTextDrawAll(Me.PVScreen, 1) 'PCN????
        
        'SavePicture Me.PVScreen.Image, WindowsTempDirectory & "CBS\EmbedFile.bmp"
        SavePicture Me.PVScreen.Image, WindowsTempDirectory & "EmbedFile.bmp" 'ID4601
    ElseIf CLPScreenMode = Video Then
        'Call grabsnapshot(WindowsTempDirectory & "CBS\EmbedFile.bmp", 1, "", 0)
        Call grabsnapshot(WindowsTempDirectory & "EmbedFile.bmp", 1, "", 0) 'ID4601
    End If
    
    With PipelineDetails.JPGMake1
        '.InputFile = WindowsTempDirectory & "CBS\EmbedFile.bmp"
        .InputFile = WindowsTempDirectory & "EmbedFile.bmp" 'ID4601
        .Quality = 80
        .Quality = 150
        '.OutputFile = WindowsTempDirectory & "CBS\EmbedFile.jpg"
        .OutputFile = WindowsTempDirectory & "EmbedFile.jpg" 'ID4601
        .Go
    End With

Exit Sub
Err_Handler:
    MsgBox Err & "-CLS175:" & Error$
End Sub
