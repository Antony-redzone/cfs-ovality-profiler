VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form SonarConfig 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Sonar Config"
   ClientHeight    =   12120
   ClientLeft      =   1530
   ClientTop       =   1020
   ClientWidth     =   12840
   Icon            =   "SonarConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "SonarConfig.frx":058A
   ScaleHeight     =   808
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   856
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton FowardOnePreview 
      Caption         =   "FWD"
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
      Left            =   3840
      TabIndex        =   53
      Top             =   6960
      Width           =   495
   End
   Begin VB.CommandButton BackOnePreview 
      Caption         =   "RW"
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
      Left            =   3360
      TabIndex        =   52
      Top             =   6960
      Width           =   495
   End
   Begin VB.TextBox SingleFrameFile 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   51
      Text            =   "C:\ClearLineProfilerv6\Sonar\599.3s3"
      Top             =   7320
      Width           =   3015
   End
   Begin VB.CommandButton PreviewSingleFrameButton 
      Caption         =   "Preview Single Frame"
      Height          =   255
      Left            =   1320
      TabIndex        =   50
      Top             =   6960
      Width           =   1815
   End
   Begin VB.TextBox txtScale 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   150
      TabIndex        =   46
      Text            =   "1.0"
      Top             =   5880
      Width           =   495
   End
   Begin VB.TextBox txtX 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   150
      TabIndex        =   45
      Text            =   "0"
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox txtY 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   960
      TabIndex        =   44
      Text            =   "0"
      Top             =   5280
      Width           =   735
   End
   Begin MSComctlLib.Slider sldAngle 
      Height          =   255
      Left            =   1920
      TabIndex        =   42
      Top             =   3360
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      MousePointer    =   10
      LargeChange     =   72
      Max             =   360
      TickFrequency   =   10
   End
   Begin VB.TextBox txtCable 
      Height          =   285
      Left            =   1920
      TabIndex        =   40
      Text            =   "0"
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdMeasure 
      BackColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   120
      Picture         =   "SonarConfig.frx":122BA
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   7440
      Width           =   465
   End
   Begin VB.PictureBox picReg 
      BackColor       =   &H00808080&
      Height          =   4050
      Left            =   8400
      Picture         =   "SonarConfig.frx":126E4
      ScaleHeight     =   3990
      ScaleWidth      =   4275
      TabIndex        =   7
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton Register 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Register"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2760
         Width           =   2100
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   2280
         Width           =   4000
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txtProductNo 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00585700&
         Caption         =   "Sonar Registration"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   720
         TabIndex        =   14
         Top             =   120
         Width           =   2895
      End
      Begin VB.Label Label3 
         BackColor       =   &H00585700&
         Caption         =   "Registration Code"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00585700&
         Caption         =   "User Name"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00585700&
         Caption         =   "Product No"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   855
      End
   End
   Begin MSComctlLib.Slider sldOversamples 
      Height          =   255
      Left            =   1920
      TabIndex        =   27
      Top             =   1440
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      MousePointer    =   10
      Min             =   1
      Max             =   8
      SelStart        =   2
      Value           =   2
   End
   Begin VB.PictureBox picControls 
      BackColor       =   &H00808080&
      Height          =   4050
      Left            =   8400
      Picture         =   "SonarConfig.frx":15B58
      ScaleHeight     =   3990
      ScaleWidth      =   4275
      TabIndex        =   16
      Top             =   4200
      Width           =   4335
      Begin VB.ComboBox cmbArcSize 
         Height          =   315
         ItemData        =   "SonarConfig.frx":18FCC
         Left            =   120
         List            =   "SonarConfig.frx":18FCE
         TabIndex        =   21
         Text            =   "360"
         Top             =   480
         Width           =   1000
      End
      Begin VB.ComboBox cmbCentreAngle 
         Height          =   315
         ItemData        =   "SonarConfig.frx":18FD0
         Left            =   120
         List            =   "SonarConfig.frx":18FD2
         TabIndex        =   20
         Text            =   "180"
         Top             =   1200
         Width           =   1000
      End
      Begin VB.TextBox txtSampleRate 
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Text            =   "2.5"
         Top             =   1920
         Width           =   1000
      End
      Begin VB.ComboBox cmbStepSize 
         Height          =   315
         ItemData        =   "SonarConfig.frx":18FD4
         Left            =   1800
         List            =   "SonarConfig.frx":18FD6
         TabIndex        =   18
         Text            =   "0.9"
         Top             =   1200
         Width           =   1000
      End
      Begin VB.TextBox txtSamples 
         Height          =   285
         Left            =   1800
         TabIndex        =   17
         Text            =   "301"
         Top             =   600
         Width           =   1000
      End
      Begin VB.Label lblArcSize 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Arc Size"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblCentreAngle 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Centre Angle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblSampleRate 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "SampleRate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblSamples 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Samples"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   1800
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblStepSize 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Step Size"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   1800
         TabIndex        =   22
         Top             =   960
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   5160
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdSavebut 
      BackColor       =   &H0000FF00&
      Height          =   300
      Left            =   5760
      Picture         =   "SonarConfig.frx":18FD8
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   300
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   7635
      Left            =   4440
      Picture         =   "SonarConfig.frx":19562
      ScaleHeight     =   509
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   680
      TabIndex        =   3
      Top             =   5400
      Width           =   10200
      Begin VB.PictureBox picLogo 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   3840
         Left            =   2640
         Picture         =   "SonarConfig.frx":1AD61
         ScaleHeight     =   256
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   256
         TabIndex        =   34
         Top             =   1560
         Width           =   3840
      End
      Begin VB.PictureBox picPallette 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   15
         Left            =   2640
         Picture         =   "SonarConfig.frx":1B7B3
         ScaleHeight     =   15
         ScaleWidth      =   3840
         TabIndex        =   33
         Top             =   1320
         Width           =   3840
      End
   End
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   5160
      Top             =   360
   End
   Begin VB.PictureBox picConnecttemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   6720
      Picture         =   "SonarConfig.frx":1BAF5
      ScaleHeight     =   660
      ScaleWidth      =   660
      TabIndex        =   1
      Top             =   3600
      Width           =   660
   End
   Begin VB.PictureBox picCalibration 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   7635
      Left            =   4680
      Picture         =   "SonarConfig.frx":1C26D
      ScaleHeight     =   509
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   680
      TabIndex        =   6
      Top             =   7800
      Width           =   10200
   End
   Begin MSComctlLib.Slider sldMin 
      Height          =   255
      Left            =   1920
      TabIndex        =   29
      Top             =   3000
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      MousePointer    =   10
      Min             =   20
      Max             =   180
      SelStart        =   110
      TickFrequency   =   10
      Value           =   110
   End
   Begin MSComctlLib.Slider sldMax 
      Height          =   255
      Left            =   1920
      TabIndex        =   31
      Top             =   2640
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      MousePointer    =   10
      Min             =   50
      Max             =   255
      SelStart        =   245
      TickFrequency   =   5
      Value           =   245
   End
   Begin MSComctlLib.Slider sldPulseWidth 
      Height          =   255
      Left            =   1920
      TabIndex        =   35
      Top             =   1800
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      MousePointer    =   10
      Min             =   1
      Max             =   6
      SelStart        =   2
      Value           =   2
   End
   Begin MSComctlLib.Slider sldCalibrate 
      Height          =   255
      Left            =   1920
      TabIndex        =   37
      Top             =   4200
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      MousePointer    =   10
      Min             =   1
      Max             =   100
      SelStart        =   50
      TickFrequency   =   10
      Value           =   50
   End
   Begin VB.Image stophighlight 
      Appearance      =   0  'Flat
      Height          =   660
      Left            =   7680
      Picture         =   "SonarConfig.frx":1DA6C
      Top             =   2640
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image imgClosed 
      Appearance      =   0  'Flat
      Height          =   660
      Left            =   6720
      Picture         =   "SonarConfig.frx":1E2A5
      Top             =   2520
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image MouseOver 
      Appearance      =   0  'Flat
      Height          =   660
      Left            =   7560
      Picture         =   "SonarConfig.frx":1EACB
      Top             =   1680
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Calibration Trim"
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
      Left            =   360
      TabIndex        =   54
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Image picConnect 
      Height          =   660
      Left            =   315
      Picture         =   "SonarConfig.frx":1F2EB
      Top             =   330
      Width           =   660
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Scale"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   150
      TabIndex        =   49
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Centre X"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   150
      TabIndex        =   48
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "CentreY"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   960
      TabIndex        =   47
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Display Angle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   43
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cable Payout"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   41
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label lblMeasurement 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Distance:"
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
      Height          =   375
      Left            =   120
      TabIndex        =   39
      Top             =   7680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Pulse Width"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   36
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Max"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   32
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Threshold"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   30
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label lblAdvanced 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Advanced ->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3120
      TabIndex        =   28
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label lblSave 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblConnect 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1140
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.Image imgNotActive 
      Height          =   480
      Left            =   4440
      Picture         =   "SonarConfig.frx":1FAE5
      Top             =   960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgActive 
      Height          =   480
      Left            =   4440
      Picture         =   "SonarConfig.frx":207AF
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgClosedHighlited 
      Appearance      =   0  'Flat
      Height          =   660
      Left            =   6720
      Picture         =   "SonarConfig.frx":21479
      Top             =   240
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image imgOpenHighlited 
      Appearance      =   0  'Flat
      Height          =   660
      Left            =   6720
      Picture         =   "SonarConfig.frx":21C94
      Top             =   1680
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image imgOpen 
      Appearance      =   0  'Flat
      Height          =   660
      Left            =   6720
      Picture         =   "SonarConfig.frx":22488
      Top             =   960
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label lblOverSamples 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Range"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   615
   End
   Begin VB.Image savemouseover 
      Height          =   660
      Left            =   3525
      Picture         =   "SonarConfig.frx":22C82
      Top             =   345
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image cmdSaveMouseOver 
      Height          =   660
      Left            =   3525
      Picture         =   "SonarConfig.frx":23339
      Top             =   345
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image cmdSaveDisabled 
      Height          =   660
      Left            =   3525
      Picture         =   "SonarConfig.frx":239D3
      Top             =   345
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image cmdSave 
      Height          =   660
      Left            =   3525
      Picture         =   "SonarConfig.frx":240D8
      Top             =   330
      Width           =   660
   End
End
Attribute VB_Name = "SonarConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mmPerPixels As Double

Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Type SAFEARRAYBOUND             '
    cElements As Long                   '
    lLbound As Long                     '
End Type                                '
                                        '
Private Type SAFEARRAY1D                '
    cDims As Integer                    '
    fFeatures As Integer                '
    cbElements As Long                  '
    cLocks As Long                      '
    pvData As Long                      '
    Bounds(0 To 0) As SAFEARRAYBOUND    '
End Type                                '
                                        
Private Type SAFEARRAY2D                '
    cDims As Integer                    '
    fFeatures As Integer                '
    cbElements As Long                  '
    cLocks As Long                      '
    pvData As Long                      '
    Bounds(0 To 1) As SAFEARRAYBOUND    '
End Type                                '
                                        '
Private Type BITMAP                     '
    bmType As Long                      '
    bmWidth As Long                     '
    bmHeight As Long                    '
    bmWidthBytes As Long                '
    bmPlanes As Integer                 '
    bmBitsPixel As Integer              '
    bmBits As Long                      '
End Type                                '
                                        '
'''''''''''''''''''''''''''''''''''''''''

Private Declare Sub InitialiseDLL Lib "SonarProfiler.dll" ()
Private Declare Function StartSonarSweep Lib "SonarProfiler.dll" (ByVal width As Long, ByVal height As Long, ByVal Preview As Long, ByVal ArcSize As Long, ByVal CentreAngle As Long, ByVal SampleRate As Double, ByVal OverSamples As Long, ByVal samples As Long, ByVal StepSize As Double, ByVal PulseWidth As Long) As Long
'Private Declare Function StartSonarSweep Lib "SonarProfiler.dll" (ByVal width As Long, ByVal height As Long, ByVal Preview As Long) As Long
Private Declare Sub StopScanning Lib "SonarProfiler.dll" ()
Private Declare Sub StartScanning Lib "SonarProfiler.dll" (ByVal name As String)

'ANT'Private Declare Sub ReadProfileFromFile Lib "SonarProfiler.dll" (ByRef Img As Byte, ByVal bytes As Long, ByVal name As String)
Private Declare Sub ReadProfileFromFile Lib "SonarProfiler.dll" (ByRef Img As Byte, ByVal bytes As Long, ByVal name As String, ByVal width As Long, ByVal height As Long, ByRef Hours As Long, ByRef Mins As Long, ByRef MSecs As Long, ByRef Distance As Long)


Private Declare Sub drawlineBMP Lib "SonarProfiler.dll" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal Col As Byte, ByRef Img As Byte)
Private Declare Sub drawframe Lib "SonarProfiler.dll" (ByRef Img As Byte)
Private Declare Sub CheckForScanning Lib "SonarProfiler.dll" (ByRef Status As Long)
Private Declare Sub VBRay Lib "SonarProfiler.dll" (ByRef Img As Byte)
Private Declare Sub SetMaxAndMin Lib "SonarProfiler.dll" (ByVal Max As Long, ByVal Min As Long)
Private Declare Sub LoadPallette Lib "SonarProfiler.dll" (ByRef Pallette As Byte)
Private Declare Sub LoadLogo Lib "SonarProfiler.dll" (ByRef Logo As Byte, ByVal width As Long, ByVal height As Long)

Private Declare Sub SetCablePayoutStart Lib "SonarProfiler.dll" (ByVal CablePos As Long)
Private Declare Sub GetCablePayout Lib "SonarProfiler.dll" (ByRef CablePos As Long)

Private Declare Sub APISleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

Private Declare Sub SetScanAngle Lib "SonarProfiler.dll" (ByVal Angle As Double)
Private Declare Sub SetScale Lib "SonarProfiler.dll" (ByVal ScaleVal As Double)
Private Declare Sub SetSonarCentre Lib "SonarProfiler.dll" (ByVal X As Long, ByVal Y As Long)


Private Connected As Boolean
Private RefreshCount As Long

Private sa As SAFEARRAY2D
Private bmp As BITMAP    'Array that is manipulated and then coppied to the passed image
Private pict() As Byte  'Pointer that is passed to the different drawnig subfunctions, so they know where to draw

Private DataLoc As String
Private JpegLoc As String

Private Sub BackOnePreview_Click()
    Call StepPreview(-1)
End Sub

Private Sub cmdMeasure_Click()
'****************************************************************************************
'Name    : cmdMeasure_Click
'Created : 9 Jan 2007, PCN4419
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    :
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler

Set picCalibration.Picture = ClearLineScreen.VideoScreen.Picture
ClearLineScreen.SonarMeasure = "Measure"
If picConnect.Picture = imgClosed.Picture Then
    cmdMeasure.Visible = False
    lblMeasurement.Visible = True
Else
    Exit Sub
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-SC1:" & Error$
End Sub

Private Sub savemouseover_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    savemouseover.Visible = False
    cmdSaveMouseOver.Visible = True
End Sub

Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    savemouseover.Visible = True
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

tmrRefresh.Enabled = False
RefreshCount = 0

Call ValidateSonarReg

If Dir(App.Path & "\Sonar", vbDirectory) = "" Then MkDir App.Path & "\Sonar"
DataLoc = App.Path & "\Sonar\"


Me.Left = PVPageLeft
Me.Top = PVPageTop
Me.width = PVPageWidth
Me.height = PVPageHeight

Set picConnect.Picture = imgOpen.Picture
Connected = False

cmbArcSize.Clear
cmbArcSize.AddItem "30"
cmbArcSize.AddItem "60"
cmbArcSize.AddItem "90"
cmbArcSize.AddItem "120"
cmbArcSize.AddItem "150"
cmbArcSize.AddItem "180"
cmbArcSize.AddItem "210"
cmbArcSize.AddItem "240"
cmbArcSize.AddItem "270"
cmbArcSize.AddItem "360"
cmbArcSize.text = "360"

cmbCentreAngle.Clear
cmbCentreAngle.AddItem "30"
cmbCentreAngle.AddItem "60"
cmbCentreAngle.AddItem "90"
cmbCentreAngle.AddItem "120"
cmbCentreAngle.AddItem "150"
cmbCentreAngle.AddItem "180"
cmbCentreAngle.AddItem "210"
cmbCentreAngle.AddItem "240"
cmbCentreAngle.AddItem "270"
cmbCentreAngle.AddItem "300"
cmbCentreAngle.AddItem "330"
cmbCentreAngle.text = "180"

cmbStepSize.Clear
cmbStepSize.AddItem "0.9"
cmbStepSize.AddItem "1.8"
cmbStepSize.AddItem "2.7"
cmbStepSize.AddItem "3.6"
cmbStepSize.text = "0.9"

'cmbPulseSize.Clear
'cmbPulseSize.AddItem "2"
'cmbPulseSize.AddItem "4"
'cmbPulseSize.AddItem "8"
'cmbPulseSize.AddItem "12"
'cmbPulseSize.AddItem "16"
'cmbPulseSize.AddItem "20"
'cmbPulseSize.text = "2"

sldOversamples.value = 2
mmPerPixels = 1.2
sldMax.value = 245
sldMin.value = 110

GetObjectAPI picPallette.Picture, Len(bmp), bmp
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

Call LoadPallette(pict(0, 0))

GetObjectAPI picLogo.Picture, Len(bmp), bmp
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

Call LoadLogo(pict(0, 0), picLogo.width, picLogo.height)
Call InitialiseDLL
CopyMemory ByVal VarPtrArray(pict), 0&, 4 'ANT

Exit Sub
Err_Handler:
    MsgBox Err & "-SC2:" & Error$
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


If Connected = False Then
    Set picConnect.Picture = imgOpen.Picture
Else
    Set picConnect.Picture = imgClosed.Picture
End If

If savemouseover.Visible = True Then
    savemouseover.Visible = False
End If

If cmdSaveMouseOver.Visible = True Then
    cmdSaveMouseOver.Visible = False
End If


Exit Sub
Err_Handler:
    MsgBox Err & "-SC3:" & Error$
End Sub


Private Sub FowardOnePreview_Click()
    Call StepPreview(1)
End Sub




Private Sub lblMeasurement_DblClick()
On Error GoTo Err_Handler


cmdMeasure.Visible = True
lblMeasurement.Visible = False
ClearLineScreen.SonarMeasure = ""


Screen.MousePointer = vbNormal
ClearLineScreen.MousePointer = vbNormal
ClearLineScreen.VideoScreen.MousePointer = vbNormal


Exit Sub
Err_Handler:
    MsgBox Err & "-SC4:" & Error$
End Sub

Private Sub picConnect_Click()
Dim ret As Long
Dim PulseWidth As Long
On Error GoTo Err_Handler

'MsgBox "Dont forget to turn Sonar Reg back on!!!", vbCritical

If Registered = False Then
    'MsgBox DisplayMessage("Registration Failed. Please Check your Registration Code."), vbCritical
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Registration Failed. Please Check your Registration Code."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Exit Sub
End If

If SonarIsRegistered = False Then
    picReg.Left = 0
    Exit Sub
End If

If Connected = False Then
    lblMeasurement.Visible = False
    cmdMeasure.Visible = True
End If

Call SetScale(0)
Call SetSonarCentre(0, 0)

If Connected = False Then
    Call BackupCurrentData
    Set ClearLineScreen.VideoScreen.Picture = Me.picCanvas.Picture
    ClearLineScreen.VideoScreen.ZOrder 0
    ClearLineScreen.SonarX1 = ClearLineScreen.VideoScreen.width / 2
    ClearLineScreen.SonarX2 = ClearLineScreen.VideoScreen.width / 2
    ClearLineScreen.SonarY1 = ClearLineScreen.VideoScreen.height / 2
    ClearLineScreen.SonarY2 = ClearLineScreen.VideoScreen.height / 2
    
    cmdSaveDisabled.Visible = True
    cmdSave.Enabled = False
        
    Select Case sldPulseWidth.value
        Case 1
            PulseWidth = 2
        Case 2
            PulseWidth = 4
        Case 3
            PulseWidth = 8
        Case 4
            PulseWidth = 12
        Case 5
            PulseWidth = 16
        Case 6
            PulseWidth = 20
    End Select
    
    Call StartSonarSweep(Me.picCanvas.width, Me.picCanvas.height, 0, CLng(Me.cmbArcSize.text), CLng(Me.cmbCentreAngle.text), SafeCDbl(Me.txtSampleRate.text), CLng(2 ^ sldOversamples.value), CLng(Me.txtSamples.text), SafeCDbl(Me.cmbStepSize.text), PulseWidth) 'PCN4616
    'Call SetMaxAndMin(245, 110)
    
    Connected = True
    Set picConnect.Picture = imgClosed.Picture
    
    Call ClearTempDirs
    
    Call StartScanning(DataLoc)
    tmrRefresh.Enabled = True
    lblConnect.Caption = "Stop Recording"
Else
    Connected = False
    
    Set picConnect.Picture = imgOpen.Picture
    imgNotActive.Left = 96
    imgActive.Left = 500
    Call StopScanning
    tmrRefresh.Enabled = False
    lblConnect.Caption = "Record"
    
    cmdSaveDisabled.Visible = False
    cmdSave.Enabled = True
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-SC5:" & Error$
End Sub

Private Sub picConnect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

If Connected = False Then
    Set picConnect.Picture = imgOpenHighlited.Picture
Else
    Set picConnect.Picture = imgClosedHighlited.Picture
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-SC6:" & Error$
End Sub

Private Sub picConnect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

If Connected = False Then
    Set picConnect.Picture = MouseOver.Picture
Else
    Set picConnect.Picture = stophighlight.Picture
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-SC7:" & Error$
End Sub



Sub BackupCurrentData()
On Error GoTo Err_Handler

Dim BackupImages As String
Dim SourceImages As String


BackupImages = App.Path & "\Sonar\BackupImages"
If Dir(BackupImages, vbDirectory) = "" Then MkDir (BackupImages)
BackupImages = BackupImages & "\" & Year(Date) & "_" & _
                                    Month(Date) & "_" & _
                                    Day(Date) & "_" & _
                                    Hour(Time) & "_" & _
                                    Minute(Time) & "_" & _
                                    Second(Time)
If Dir(BackupImages, vbDirectory) = "" Then MkDir (BackupImages)


SourceImages = Dir(App.Path & "\sonar\*.3s3*")
Do While SourceImages <> ""
    'Check to see if the file exists in the target directory
    On Error Resume Next
    FileCopy App.Path & "\sonar\" & SourceImages, BackupImages & "\" & SourceImages
    On Error GoTo Err_Handler
    SourceImages = Dir()
Loop


Exit Sub
Err_Handler:
    MsgBox Err & "-SC8:" & Error$
End Sub


Private Sub Register_Click()
'****************************************************************************************
'Name    : tmrRefresh_Timer
'Created :
'Updated :
'Prg By  :
'Param   :
'Desc    :
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
Dim ArmClass As CArmadillo

Set ArmClass = New CArmadillo

If ArmClass.IsValidKey(txtUser, txtCode) Then
    
    Call ArmClass.InstallKey(txtUser, txtCode)
    
    If UCase(ArmClass.ClearLineRegType) = "SONAR" Then
        
        Dim ArmRegType As String
        ArmRegType = ArmClass.ClearLineRegType
        'MsgBox DisplayMessage("Congratulations, your ClearLine Profiler Software is now registered. Please shut-down the application and restart to allow new settings to take effect."), vbInformation
        ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Congratulations, your ClearLine Profiler Software is now registered. Please shut-down the application and restart to allow new settings to take effect."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
        picReg.Left = 550
        SonarIsRegistered = True
    End If
Else
    'MsgBox DisplayMessage("Registration Failed. Please Check your Registration Code."), vbCritical
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Registration Failed. Please Check your Registration Code."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    SonarIsRegistered = False
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-SC9:" & Error$
End Sub

Private Sub savemouseover_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
Dim DatName As String
Dim JpgName As String
Dim BmpName As String
Dim ret As Long
Dim Count As Long
Dim CountNumberFrames As Long

cmdSaveMouseOver.Visible = False

'ANT''''''''''''''''''''''''''''''''
Dim Hours As Long
Dim Minutes As Long
Dim MilliSeconds As Long
Dim Distance As Long
'''''''''''''''''''''''''''''''''''

'Call SetScanAngle(CDbl(sldAngle.value))
'Call SetScale(CDbl(txtScale.text))
'Call SetSonarCentre(CLng(txtX.text), CLng(txtY.text))



Dialog.DefaultExt = "avi"
Dialog.DialogTitle = "Save Avi in specified location"
Dialog.Filter = "*.avi"
Dialog.ShowSave

'Me.picCanvas.AutoSize = True
'Me.picCanvas.Picture = LoadPicture(App.Path & "\Bmp1920x1536.bmp")

GetObjectAPI Me.picCanvas.Picture, Len(bmp), bmp
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


JpegLoc = App.Path & "\Sonar\Video\"

If Dir(JpegLoc, vbDirectory) = "" Then MkDir (JpegLoc)


                                    


DatName = Dir(DataLoc & "*.3S3")


While DatName <> ""
    CountNumberFrames = CountNumberFrames + 1



    Count = CLng(Left(DatName, Len(DatName) - 3))
    
    'Call ReadProfileFromFile(pict(0, 0), 3, DataLoc & DatName)
    
    Call SetScanAngle(SafeCDbl(sldAngle.value)) 'PCN4616
    Call SetScale(SafeCDbl(txtScale.text)) 'PCN4616
    Call SetSonarCentre(CLng(txtX.text), CLng(txtY.text))
    
    Call ReadProfileFromFile(pict(0, 0), 3, DataLoc & DatName, picCanvas.width, picCanvas.height, Hours, Minutes, MilliSeconds, Distance)
    
    JpgName = Left(DatName, Len(DatName) - 3) & "jpg"
    Call SavePicture(SonarConfig.picCanvas, JpegLoc & "Temp.bmp")
    
    With PipelineDetails.JPGMake1
        .InputFile = JpegLoc & "Temp.bmp"
        .Quality = 100
        .OutputFile = JpegLoc & "IMG" & JpgName
        .Go
    End With

    picCanvas.Refresh
    DatName = Dir()
Wend

CopyMemory ByVal VarPtrArray(pict), 0&, 4 'ANT


Call RenderCalibrator(CountNumberFrames + 1)


Dim ffmpegCMD As String
ffmpegCMD = "\ffmpeg -i """ & JpegLoc & "IMG%d.jpg"" -sameq -vcodec mpeg4 """ & Dialog.FileName & """"
ret = Shell(App.Path & ffmpegCMD, vbNormalFocus)


Exit Sub
Err_Handler:
    MsgBox Err & "-SC10:" & Error$
End Sub

Private Sub sldAngle_Click()
'****************************************************************************************
'Name    : sldAngle_Change
'Created :
'Updated :
'Prg By  :
'Param   :
'Desc    :
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler

Call SetScanAngle(SafeCDbl(sldAngle.value)) 'PCN4616

Exit Sub
Err_Handler:
    MsgBox Err & "-SC11:" & Error$
End Sub

Private Sub sldCalibrate_Change()
'****************************************************************************************
'Name    : sldCalibrate_Change
'Created :
'Updated :
'Prg By  :
'Param   :
'Desc    :
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
    
    'mmPerPixels = 1.1 + (sldCalibrate.value * 0.002)
    Call Me.SetMMPerPixel
    Call ClearLineScreen.DrawSonarMeasureLine

Exit Sub
Err_Handler:
    MsgBox Err & "-SC12:" & Error$
End Sub

Private Sub sldMax_Change()
'****************************************************************************************
'Name    : sldMax_Change
'Created :
'Updated :
'Prg By  :
'Param   :
'Desc    :
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
    
    Call SetMaxAndMin(sldMax.value, sldMin.value)

Exit Sub
Err_Handler:
    MsgBox Err & "-SC13:" & Error$
End Sub

Private Sub sldMin_Change()
'****************************************************************************************
'Name    : sldMin_Change
'Created :
'Updated :
'Prg By  :
'Param   :
'Desc    :
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
    
    Call SetMaxAndMin(sldMax.value, sldMin.value)
    Call PreviewSingleFrameButton_Click

Exit Sub
Err_Handler:
    MsgBox Err & "-SC14:" & Error$
End Sub


Private Sub sldOversamples_Click()
On Error GoTo Err_Handler

    Call SetMMPerPixel
    
Exit Sub
Err_Handler:
    MsgBox Err & "-SC15:" & Error$
End Sub

Private Sub tmrRefresh_Timer()
'****************************************************************************************
'Name    : tmrRefresh_Timer
'Created :
'Updated :
'Prg By  :
'Param   :
'Desc    :
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
Dim Active As Long

RefreshCount = RefreshCount + 1

Call CheckForScanning(Active)
If Active = 0 Then
    tmrRefresh.Enabled = False
    Exit Sub
End If

If imgNotActive.Left = 96 Then
    If RefreshCount Mod 20 = 0 Then
        imgNotActive.Left = 500
        imgActive.Left = 96
    End If
Else
    If RefreshCount Mod 10 = 0 Then
        imgActive.Left = 500
        imgNotActive.Left = 96
    End If
End If

GetObjectAPI ClearLineScreen.VideoScreen.Picture, Len(bmp), bmp
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

Call VBRay(pict(0, 0))
ClearLineScreen.VideoScreen.Refresh
CopyMemory ByVal VarPtrArray(pict), 0&, 4 'ANT

Dim Val As Long
Call GetCablePayout(Val)
txtCable = Val

Exit Sub
Err_Handler:
    MsgBox Err & "-SC16:" & Error$
End Sub

Sub SetMMPerPixel()
On Error GoTo Err_Handler



mmPerPixels = 1.1 + (sldCalibrate.value * 0.002)
If sldOversamples.value = 2 Then
    mmPerPixels = mmPerPixels
ElseIf sldOversamples.value = 3 Then
    mmPerPixels = mmPerPixels * 2.08
Else
    mmPerPixels = mmPerPixels * 2.08 * 2.08
End If



Exit Sub
Err_Handler:
    MsgBox Err & "-SC17:" & Error$
End Sub

Private Sub RenderCalibrator(ByVal CountVal As Integer)
'****************************************************************************************
'Name    : RenderCalibrator
'Created :
'Updated :
'Prg By  :
'Param   :
'Desc    :
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
Dim RectLeft As Integer
Dim RectWidth As Integer
Dim RectTop As Integer
Dim RectHeight As Integer
Dim SquareWidth As Integer
Dim SquareHeight As Integer
Dim PixelsPerMM As Double
Dim CalibrationText As String
Dim Distance As Long

Call Me.SetMMPerPixel


'If sldOversamples.value = 2 Then
'    PixelsPerMM = 260 / 295
'ElseIf sldOversamples.value = 3 Then
'    PixelsPerMM = 125 / 295 'oversamples = 8
'Else
'    PixelsPerMM = 60 / 295 'oversamples = 16
'End If

PixelsPerMM = 1 / Me.mmPerPixels

If MeasurementUnits = "mm" Then
    Distance = (CLng(2 ^ sldOversamples.value) * 100)
    RectWidth = PixelsPerMM * Distance
    CalibrationText = "mm"
Else
    Distance = (CLng(2 ^ sldOversamples.value) * 5)
    RectWidth = PixelsPerMM * 25.4 * Distance
    CalibrationText = "in"
End If
CalibrationText = (Distance / SafeCDbl(txtScale.text)) & " " & CalibrationText 'PCN4616

RectLeft = picCalibration.width / 2 - (RectWidth / 2)

RectTop = picCalibration.height * 0.45
RectHeight = 50

SquareWidth = RectWidth * 0.1
SquareHeight = 25


picCalibration.Line (RectLeft, RectTop)-(RectLeft + RectWidth, RectTop + RectHeight), vbWhite, BF

picCalibration.Line (RectLeft, RectTop)-(RectLeft + SquareWidth, RectTop + SquareHeight), vbBlack, BF
picCalibration.Line (RectLeft - SquareWidth, RectTop)-(RectLeft - 1, RectTop + SquareHeight), vbWhite, BF
picCalibration.Line (RectLeft + RectWidth, RectTop)-(RectLeft + RectWidth - SquareWidth, RectTop + SquareHeight), vbBlack, BF
picCalibration.Line (RectLeft + RectWidth + 1, RectTop)-(RectLeft + RectWidth + SquareWidth, RectTop + SquareHeight), vbWhite, BF

picCalibration.CurrentX = RectLeft + (RectWidth / 2) - 10
picCalibration.CurrentY = RectTop + SquareHeight
picCalibration.Print CalibrationText

picCalibration.Refresh
DoEvents

Call SavePicture(picCalibration.Image, JpegLoc & "Temp.bmp")

Dim I As Integer

For I = CountVal To CountVal + 5
    With PipelineDetails.JPGMake1
        .InputFile = JpegLoc & "Temp.bmp"
        .Quality = 100
        .OutputFile = JpegLoc & "IMG" & I & ".jpg"
        .Go
    End With
Next I


Exit Sub
Err_Handler:
    MsgBox Err & "-SC18:" & Error$
End Sub

Private Sub ClearTempDirs()
'****************************************************************************************
'Name    : ClearTempDirs
'Created :
'Updated :
'Prg By  :
'Param   :
'Desc    :
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler

Call KillFilesSafely(App.Path & "\Sonar\", "*.txt")
Call KillFilesSafely(App.Path & "\Sonar\Video\", "*.jpg")
Call KillFilesSafely(App.Path & "\sonar\", "*.3s3")

Exit Sub
Err_Handler:
    MsgBox Err & "-SC19:" & Error$
End Sub


Sub KillFilesSafely(FilePath As String, FileName As String)
'****************************************************************************************
'Name    : KillFilesSafely
'Created : 1 March 2005 PCN3336
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    :
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
Dim PathAndFileCriteria As String

PathAndFileCriteria = FilePath & "\" & FileName
Kill PathAndFileCriteria


Exit Sub
ClearReadOnly:
    On Error GoTo Err_Handler
    Call FileReadOnlyClear(FilePath, FileName)
    Kill PathAndFileCriteria

Exit Sub
Err_Handler:
Select Case Err
    Case 53 'File not found
        Resume Next
    Case 75 'Path or File access error (This may mean the file is READYONLY)
        GoTo ClearReadOnly
    Case Else
End Select
End Sub

Sub FileReadOnlyClear(PathName As String, FileName As String)
'****************************************************************************************
'Name    : FileReadOnlyClear
'Created : 3 March 2005 PCN3336
'Updated :
'Prg By  : Geoff Logan
'Param   : PathAndFileCriteria - eg "C:\CBSTemp\*.*"
'Desc    : Clears the vbReadOnly attribute of readonly files.
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
Dim FileNameToSetRead As String
Dim PathAndFileName As String

PathAndFileName = PathName & FileName
'Clears the vbReadOnly attribute.
FileNameToSetRead = Dir(PathAndFileName, vbReadOnly)
While FileNameToSetRead <> ""
    FileNameToSetRead = PathName & FileNameToSetRead
    SetAttr FileNameToSetRead, GetAttr(FileNameToSetRead) And (Not vbReadOnly)
    FileNameToSetRead = Dir()
Wend

Exit Sub
Err_Handler:
    Select Case Err
        Case 6: Resume Next 'Overflow PCNVista A temp vista fix.
        Case Else: MsgBox Err & "-SC20:" & Error$
    End Select
End Sub

Private Sub ValidateSonarReg()
'****************************************************************************************
'Name    : ValidateSonarReg
'Created : 15 Sep 2006
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    :
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler

Dim ArmClass As CArmadillo
Dim RegType As String

Set ArmClass = New CArmadillo

txtProductNo = ArmClass.HardwareFingerPrint
txtUser = ArmClass.UserName
RegType = ArmClass.ClearLineRegType

If UCase(RegType) = "SONAR" Then
    SonarIsRegistered = True
Else
    SonarIsRegistered = False
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-SC21:" & Error$
End Sub


Private Sub txtCable_KeyPress(KeyAscii As Integer)
'****************************************************************************************
'Name    : txtCable_KeyPress
'Created : 11 Jan 2007, PCN4419
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    :
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
Dim Val As Long

If IsNumeric(txtCable.text & Chr(KeyAscii)) Then
    Val = CLng(txtCable.text & Chr(KeyAscii))
    Call SetCablePayoutStart(Val)
Else
    MsgBox "Please enter only numbers in the Cable Payout field", vbInformation, "Invaled Entry"
End If

Exit Sub
Err_Handler:
   MsgBox Err & "-SC22:" & Error$
End Sub

Private Sub PreviewSingleFrameButton_Click()
'****************************************************************************************
'Name    : PreviewSingleFrameButton_Click
'Created : 26 Jan 2006
'Updated :
'Prg By  : Antony van Iersel
'Param   :
'Desc    : Takes a sonar file name and directory and displays the Sonar Data as
'          as if it was creating the frame for, the sonar file name is retrieved from
'          the preview sonar file name textbox
'Usage   : To be able to quickly verify the sonar is drawing correctly
'****************************************************************************************
On Error GoTo Err_Handler

Dim DataLoc As String
Dim Count As Long
Dim Hours As Long
Dim Minutes As Long
Dim MilliSeconds As Long
Dim Distance As Long

ClearLineScreen.VideoScreen.Picture = LoadPicture(App.Path & "\Bmp684x515.bmp")
DataLoc = SingleFrameFile.text


GetObjectAPI ClearLineScreen.VideoScreen.Picture, Len(bmp), bmp
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

CopyMemory ByVal VarPtrArray(pict), VarPtr(sa), 4 'Point the pict pointer to the image arrray
    
Call SetMaxAndMin(sldMax.value, sldMin.value)
Call ReadProfileFromFile(pict(0, 0), 3, DataLoc, ClearLineScreen.VideoScreen.width, ClearLineScreen.VideoScreen.height, Hours, Minutes, MilliSeconds, Distance)
ClearLineScreen.VideoScreen.Refresh
CopyMemory ByVal VarPtrArray(pict), 0&, 4

ClearLineScreen.VideoScreen.ZOrder 0
ClearLineScreen.VideoScreen.Visible = True




Exit Sub
Err_Handler:
    MsgBox Err & "-SC23:" & Error$
End Sub

Sub StepPreview(ByVal amount As Integer)
On Error GoTo Err_Handler

Dim CurrentFrame As String
Dim StartNo As Integer
Dim EndNo As Integer


CurrentFrame = LCase(SingleFrameFile)
StartNo = InStr(1, CurrentFrame, "sonar") + 6
EndNo = InStr(1, CurrentFrame, ".")
CurrentFrame = Mid(CurrentFrame, StartNo, EndNo - StartNo)
CurrentFrame = App.Path & "\sonar\" & CStr(CInt(CurrentFrame) + amount) & ".3s3"
SingleFrameFile = CurrentFrame
Call PreviewSingleFrameButton_Click



Exit Sub
Err_Handler:
    MsgBox Err & "-SC24:" & Error$
End Sub
