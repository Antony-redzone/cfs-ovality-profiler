VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ControlsScreen 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   ClientHeight    =   12570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "ControlsScreen.frx":0000
   ScaleHeight     =   12570
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   Begin VB.Timer GeneralTimer 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10320
      Top             =   1800
   End
   Begin MSComctlLib.ImageList StoredReportsImageList 
      Left            =   7560
      Top             =   11400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":3C60E
            Key             =   "MultiProfile"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":3E2E8
            Key             =   "MultiLine"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":3FFC2
            Key             =   "Profile"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":41C9C
            Key             =   "Single"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ControlsCalibrationImageList 
      Left            =   3480
      Top             =   8520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":43976
            Key             =   "CalibrationV"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":45650
            Key             =   "CalibrationVNotSelected"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":4732A
            Key             =   "CalibrationCrack"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":49004
            Key             =   "CalibrationCrackNotSelected"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ControlsWizardImageList 
      Left            =   7800
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":4ACDE
            Key             =   "Calibration"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":4C9B8
            Key             =   "ImageProcessing"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":4E692
            Key             =   "SetMask"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":5036C
            Key             =   "SetWaterLevel"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":52046
            Key             =   "StepForward"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":53D20
            Key             =   "StepRewind"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":559FA
            Key             =   "StepRewindDisabled"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":576D4
            Key             =   "PVRecord"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":593AE
            Key             =   "PVRecordDisabled"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ControlsViewImageList 
      Left            =   9240
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":5B088
            Key             =   "3D"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":5CD62
            Key             =   "3DNotSelected"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":5EA3C
            Key             =   "SnapShot"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":60716
            Key             =   "SnapShotNotSelected"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":623F0
            Key             =   "PV"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":640CA
            Key             =   "PVNotSelected"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":65DA4
            Key             =   "Video"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":67A7E
            Key             =   "VideoNotSelected"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ControlsFixedImageList 
      Left            =   8640
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":69758
            Key             =   "Background"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":69ED2
            Key             =   "Depress"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":6A64C
            Key             =   "Record"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":6ADC6
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":6B540
            Key             =   "Play"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":6BCBA
            Key             =   "Pause"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":6C434
            Key             =   "IP_OFF"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":6CBAE
            Key             =   "IP_ON"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":6D328
            Key             =   "RefShapeOFF"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":6DC02
            Key             =   "RefShapeON"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":6E4DC
            Key             =   "3DMapOn"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":6EC56
            Key             =   "3DMapOff"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":6F3D0
            Key             =   "Zoom_Off"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsScreen.frx":6FCAA
            Key             =   "Zoom_On"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1mlReport 
      Caption         =   "1mi Project Report"
      Height          =   255
      Left            =   9120
      TabIndex        =   9
      Top             =   10440
      Width           =   1455
   End
   Begin VB.Label Label1kReport 
      Caption         =   "1km Project Report"
      Height          =   255
      Left            =   9120
      TabIndex        =   8
      Top             =   10080
      Width           =   1455
   End
   Begin VB.Image ControlsReports 
      Height          =   720
      Index           =   5
      Left            =   2880
      Picture         =   "ControlsScreen.frx":70584
      Tag             =   "1K"
      ToolTipText     =   "1Km Report"
      Top             =   5280
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsStoredReports 
      Height          =   720
      Index           =   7
      Left            =   6000
      Picture         =   "ControlsScreen.frx":7224E
      Top             =   11280
      Width           =   720
   End
   Begin VB.Image ControlsStoredReports 
      Height          =   720
      Index           =   6
      Left            =   5160
      Picture         =   "ControlsScreen.frx":73F18
      Top             =   11280
      Width           =   720
   End
   Begin VB.Image ControlsStoredReports 
      Height          =   720
      Index           =   5
      Left            =   4320
      Picture         =   "ControlsScreen.frx":75BE2
      Top             =   11280
      Width           =   720
   End
   Begin VB.Image ControlsStoredReports 
      Height          =   720
      Index           =   4
      Left            =   3480
      Picture         =   "ControlsScreen.frx":778AC
      Top             =   11280
      Width           =   720
   End
   Begin VB.Image ControlsStoredReports 
      Height          =   720
      Index           =   3
      Left            =   2640
      Picture         =   "ControlsScreen.frx":79576
      Top             =   11280
      Width           =   720
   End
   Begin VB.Image ControlsStoredReports 
      Height          =   720
      Index           =   2
      Left            =   1800
      Picture         =   "ControlsScreen.frx":7B240
      Top             =   11280
      Width           =   720
   End
   Begin VB.Image ControlsStoredReports 
      Height          =   720
      Index           =   1
      Left            =   960
      Picture         =   "ControlsScreen.frx":7CF0A
      Top             =   11280
      Width           =   720
   End
   Begin VB.Image ControlsStoredReports 
      Height          =   720
      Index           =   0
      Left            =   120
      Picture         =   "ControlsScreen.frx":7EBD4
      Top             =   11280
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsReports 
      Height          =   720
      Index           =   4
      Left            =   3480
      Picture         =   "ControlsScreen.frx":8089E
      Tag             =   "StoredReports"
      ToolTipText     =   "Stored Reports"
      Top             =   5280
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsFixed 
      Height          =   480
      Index           =   18
      Left            =   3000
      Picture         =   "ControlsScreen.frx":82568
      Tag             =   "Zoom"
      ToolTipText     =   "Zoom In"
      Top             =   120
      Width           =   480
   End
   Begin VB.Image ControlsRefShape 
      Height          =   720
      Index           =   3
      Left            =   2640
      Picture         =   "ControlsScreen.frx":82E32
      Tag             =   "RefShapeHorseShoe"
      Top             =   9480
      Width           =   720
   End
   Begin VB.Image ControlsSpecial 
      Height          =   720
      Index           =   3
      Left            =   2640
      Picture         =   "ControlsScreen.frx":84AFC
      Tag             =   "RefShapes"
      Top             =   10320
      Width           =   720
   End
   Begin VB.Image ControlsSpecial 
      Height          =   720
      Index           =   2
      Left            =   1800
      Picture         =   "ControlsScreen.frx":859C6
      Tag             =   "SetWaterLevel"
      Top             =   10320
      Width           =   720
   End
   Begin VB.Image ControlsSpecial 
      Height          =   720
      Index           =   1
      Left            =   960
      Picture         =   "ControlsScreen.frx":87690
      Tag             =   "SetVideoMask"
      Top             =   10320
      Width           =   720
   End
   Begin VB.Image ControlsSpecial 
      Height          =   720
      Index           =   0
      Left            =   120
      Picture         =   "ControlsScreen.frx":8935A
      Tag             =   "Calibration"
      Top             =   10320
      Width           =   720
   End
   Begin VB.Image ControlsRefShape 
      Height          =   720
      Index           =   2
      Left            =   1800
      Picture         =   "ControlsScreen.frx":8B024
      Tag             =   "RefShapeEgg"
      Top             =   9480
      Width           =   720
   End
   Begin VB.Image ControlsRefShape 
      Height          =   720
      Index           =   1
      Left            =   960
      Picture         =   "ControlsScreen.frx":8CCEE
      Tag             =   "RefShapeCircle"
      Top             =   9480
      Width           =   720
   End
   Begin VB.Image ControlsRefShape 
      Height          =   720
      Index           =   0
      Left            =   120
      Picture         =   "ControlsScreen.frx":8E9B8
      Tag             =   "ExternalDiameter"
      Top             =   9480
      Width           =   720
   End
   Begin VB.Image ControlsCalibration 
      Height          =   720
      Index           =   2
      Left            =   1800
      Picture         =   "ControlsScreen.frx":90682
      Tag             =   "CalibrationCrack"
      ToolTipText     =   "Crack Measurement Calibration"
      Top             =   8520
      Width           =   720
   End
   Begin VB.Image ControlsCalibration 
      Height          =   720
      Index           =   1
      Left            =   960
      Picture         =   "ControlsScreen.frx":9234C
      Tag             =   "CalibrationV"
      ToolTipText     =   "Calibration - Vertical"
      Top             =   8520
      Width           =   720
   End
   Begin VB.Image ControlsCalibration 
      Height          =   720
      Index           =   0
      Left            =   120
      Picture         =   "ControlsScreen.frx":94016
      Tag             =   "CalibrationH"
      ToolTipText     =   "Calibration - Horizontal"
      Top             =   8520
      Width           =   720
   End
   Begin VB.Image ControlsFixed 
      Height          =   480
      Index           =   17
      Left            =   600
      Picture         =   "ControlsScreen.frx":95CE0
      Tag             =   "Reports"
      ToolTipText     =   "Reports"
      Top             =   120
      Width           =   480
   End
   Begin VB.Image ControlsPVDRecording 
      Height          =   720
      Index           =   4
      Left            =   2640
      Picture         =   "ControlsScreen.frx":965AA
      Tag             =   "StopRecording"
      ToolTipText     =   "Stop Recording"
      Top             =   3000
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsFixed 
      Height          =   360
      Index           =   16
      Left            =   10200
      Picture         =   "ControlsScreen.frx":98274
      Tag             =   "Controls3DSnapShot"
      ToolTipText     =   "Take a snap-shot for reporting"
      Top             =   7920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image ControlsFixed 
      Height          =   360
      Index           =   15
      Left            =   9720
      Picture         =   "ControlsScreen.frx":989DE
      Tag             =   "Controls3DPause"
      ToolTipText     =   "Pause"
      Top             =   7920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image ControlsFixed 
      Height          =   360
      Index           =   14
      Left            =   9240
      Picture         =   "ControlsScreen.frx":99148
      Tag             =   "Controls3DPlay"
      ToolTipText     =   "Play 3D, press again to change speed"
      Top             =   7920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image ControlsFixed 
      Height          =   360
      Index           =   13
      Left            =   8760
      Picture         =   "ControlsScreen.frx":998B2
      Tag             =   "Controls3DRewind"
      ToolTipText     =   "Play 3D back, press again to change speed"
      Top             =   7920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image ControlsFixed 
      Height          =   360
      Index           =   12
      Left            =   8280
      Picture         =   "ControlsScreen.frx":9A01C
      Tag             =   "Controls3DReset"
      ToolTipText     =   "Reset tools"
      Top             =   7920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image ControlsFixed 
      Height          =   360
      Index           =   11
      Left            =   1080
      Picture         =   "ControlsScreen.frx":9A786
      Tag             =   "Controls3DMapOnOff"
      ToolTipText     =   "Turn map view on/off"
      Top             =   7920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image ControlsFixed 
      Height          =   360
      Index           =   10
      Left            =   1560
      Picture         =   "ControlsScreen.frx":9AEF0
      Tag             =   "Controls3DController"
      ToolTipText     =   "Control selection"
      Top             =   7920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image ControlsFixed 
      Height          =   360
      Index           =   9
      Left            =   2040
      Picture         =   "ControlsScreen.frx":9B65A
      Tag             =   "Controls3DZoom"
      ToolTipText     =   "Zoom in"
      Top             =   7920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image Controls3D 
      Height          =   720
      Index           =   10
      Left            =   6000
      Picture         =   "ControlsScreen.frx":9BDC4
      Tag             =   "RotateYClockwise"
      Top             =   7200
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Controls3D 
      Height          =   720
      Index           =   9
      Left            =   5160
      Picture         =   "ControlsScreen.frx":9DA8E
      Tag             =   "RotateYAntiClock"
      Top             =   7200
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Controls3D 
      Height          =   720
      Index           =   8
      Left            =   4320
      Picture         =   "ControlsScreen.frx":9F758
      Tag             =   "RotateXClockwise"
      ToolTipText     =   "Rotate pipe clockwise"
      Top             =   7200
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Controls3D 
      Height          =   720
      Index           =   7
      Left            =   3480
      Picture         =   "ControlsScreen.frx":A1422
      Tag             =   "RotateXAntiClock"
      ToolTipText     =   "Rotate pipe anti-clockwise"
      Top             =   7200
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Controls3D 
      Height          =   720
      Index           =   6
      Left            =   6000
      Picture         =   "ControlsScreen.frx":A30EC
      Tag             =   "ZoomOut"
      ToolTipText     =   "Zoom out"
      Top             =   6240
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Controls3D 
      Height          =   720
      Index           =   5
      Left            =   5160
      Picture         =   "ControlsScreen.frx":A4DB6
      Tag             =   "ZoomIn"
      ToolTipText     =   "Zoom in"
      Top             =   6240
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Controls3D 
      Height          =   720
      Index           =   4
      Left            =   4320
      Picture         =   "ControlsScreen.frx":A6A80
      Tag             =   "Scale"
      ToolTipText     =   "Adjust pipe scale"
      Top             =   6240
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Controls3D 
      Height          =   720
      Index           =   3
      Left            =   2640
      Picture         =   "ControlsScreen.frx":A874A
      Tag             =   "ViewMap"
      ToolTipText     =   "Map camera view"
      Top             =   6240
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Controls3D 
      Height          =   720
      Index           =   2
      Left            =   1800
      Picture         =   "ControlsScreen.frx":AA414
      Tag             =   "ViewFlyBy"
      ToolTipText     =   "FlyBy camera view"
      Top             =   6240
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Controls3D 
      Height          =   720
      Index           =   1
      Left            =   960
      Picture         =   "ControlsScreen.frx":AC0DE
      Tag             =   "ViewOutSide"
      ToolTipText     =   "External camera view"
      Top             =   6240
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Controls3D 
      Height          =   720
      Index           =   0
      Left            =   120
      Picture         =   "ControlsScreen.frx":ADDA8
      Tag             =   "ViewTanAndTilt"
      ToolTipText     =   "'Pan and Tilt' camera view"
      Top             =   6240
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsReports 
      Height          =   720
      Index           =   3
      Left            =   2160
      Picture         =   "ControlsScreen.frx":AFA72
      Tag             =   "MultiProfile"
      ToolTipText     =   "Observations Report"
      Top             =   5280
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsReports 
      Height          =   720
      Index           =   2
      Left            =   1440
      Picture         =   "ControlsScreen.frx":B173C
      Tag             =   "Profile"
      ToolTipText     =   "Profile Report"
      Top             =   5280
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsReports 
      Height          =   720
      Index           =   1
      Left            =   840
      Picture         =   "ControlsScreen.frx":B3406
      Tag             =   "MultiLine"
      ToolTipText     =   "Analysis Report"
      Top             =   5280
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsReports 
      Height          =   720
      Index           =   0
      Left            =   120
      Picture         =   "ControlsScreen.frx":B50D0
      Tag             =   "Single"
      ToolTipText     =   "Summary Report"
      Top             =   5280
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsPostRecording 
      Height          =   720
      Index           =   1
      Left            =   960
      Picture         =   "ControlsScreen.frx":B6D9A
      Tag             =   "SaveToFile"
      ToolTipText     =   "Save To File"
      Top             =   4200
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsPVDRecording 
      Height          =   720
      Index           =   3
      Left            =   1800
      Picture         =   "ControlsScreen.frx":B8A64
      Tag             =   "DisplayPipeDetails"
      ToolTipText     =   "Pipeline Details"
      Top             =   3000
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsPostRecording 
      Height          =   720
      Index           =   0
      Left            =   120
      Picture         =   "ControlsScreen.frx":BA72E
      Tag             =   "DisplayPipeDetails"
      ToolTipText     =   "Pipeline Details"
      Top             =   4200
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsPVDRecording 
      Height          =   720
      Index           =   2
      Left            =   3480
      Picture         =   "ControlsScreen.frx":BC3F8
      Tag             =   "SetWaterLevel"
      ToolTipText     =   "Set Water Level"
      Top             =   3000
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsPVDRecording 
      Height          =   720
      Index           =   1
      Left            =   960
      Picture         =   "ControlsScreen.frx":BE0C2
      Tag             =   "SetVideoMask"
      ToolTipText     =   "Set Mask"
      Top             =   3000
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsPVDRecording 
      Height          =   720
      Index           =   0
      Left            =   120
      Picture         =   "ControlsScreen.frx":BFD8C
      Tag             =   "IPSettings"
      ToolTipText     =   "Tuning settings"
      Top             =   3000
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsWizard 
      Height          =   720
      Index           =   6
      Left            =   4320
      Picture         =   "ControlsScreen.frx":C1A56
      Tag             =   "StartRecording"
      ToolTipText     =   "Record"
      Top             =   720
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsWizard 
      Height          =   720
      Index           =   5
      Left            =   2640
      Picture         =   "ControlsScreen.frx":C3720
      Tag             =   "SetFinishRecordingMarker"
      ToolTipText     =   "Set Finish Recording Marker"
      Top             =   720
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsWizard 
      Height          =   720
      Index           =   4
      Left            =   1800
      Picture         =   "ControlsScreen.frx":C53EA
      Tag             =   "SetStartRecordingMarker"
      ToolTipText     =   "Set Start Recording Marker"
      Top             =   720
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsWizard 
      Height          =   720
      Index           =   3
      Left            =   6000
      Picture         =   "ControlsScreen.frx":C70B4
      Tag             =   "SetWaterLevel"
      ToolTipText     =   "Set Water Level"
      Top             =   720
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsWizard 
      Height          =   720
      Index           =   2
      Left            =   5160
      Picture         =   "ControlsScreen.frx":C8D7E
      Tag             =   "SetVideoMask"
      ToolTipText     =   "Set Mask"
      Top             =   720
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsWizard 
      Height          =   720
      Index           =   1
      Left            =   3480
      Picture         =   "ControlsScreen.frx":CAA48
      Tag             =   "IPSettings"
      ToolTipText     =   "Tuning settings"
      Top             =   720
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsDraw 
      Height          =   480
      Index           =   8
      Left            =   6360
      Picture         =   "ControlsScreen.frx":CC712
      Tag             =   "DrawText"
      ToolTipText     =   "Text"
      Top             =   120
      Width           =   480
   End
   Begin VB.Image ControlsDraw 
      Height          =   480
      Index           =   7
      Left            =   5760
      Picture         =   "ControlsScreen.frx":CD3DC
      Tag             =   "Erase"
      ToolTipText     =   "Clear"
      Top             =   120
      Width           =   480
   End
   Begin VB.Image ControlsDraw 
      Height          =   480
      Index           =   6
      Left            =   7560
      Picture         =   "ControlsScreen.frx":CE0A6
      Tag             =   "Integrated"
      ToolTipText     =   "Integrated Area Calculation"
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ControlsDraw 
      Height          =   480
      Index           =   5
      Left            =   6960
      Picture         =   "ControlsScreen.frx":CED70
      Tag             =   "Perpendicular"
      ToolTipText     =   "Perpendicular line to main circle"
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ControlsDraw 
      Height          =   480
      Index           =   4
      Left            =   5160
      Picture         =   "ControlsScreen.frx":CF63A
      Tag             =   "MultiLine"
      ToolTipText     =   "PolyLine"
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ControlsDraw 
      Height          =   480
      Index           =   3
      Left            =   4560
      Picture         =   "ControlsScreen.frx":CFF04
      Tag             =   "LineDraw"
      ToolTipText     =   "Line"
      Top             =   120
      Width           =   480
   End
   Begin VB.Image ControlsDraw 
      Height          =   480
      Index           =   2
      Left            =   4080
      Picture         =   "ControlsScreen.frx":D07CE
      Tag             =   "Dimension"
      ToolTipText     =   "Dimension"
      Top             =   120
      Width           =   480
   End
   Begin VB.Image ControlsDraw 
      Height          =   495
      Index           =   1
      Left            =   3600
      Picture         =   "ControlsScreen.frx":D1498
      Tag             =   "Move"
      ToolTipText     =   "Move"
      Top             =   120
      Width           =   480
   End
   Begin VB.Image ControlsWizard 
      Height          =   720
      Index           =   0
      Left            =   120
      Picture         =   "ControlsScreen.frx":D1976
      Tag             =   "Calibration"
      ToolTipText     =   "Calibration"
      Top             =   1800
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsView 
      Height          =   720
      Index           =   3
      Left            =   9960
      Picture         =   "ControlsScreen.frx":D3640
      Tag             =   "Video"
      ToolTipText     =   "Video"
      Top             =   720
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsView 
      Height          =   720
      Index           =   2
      Left            =   9120
      Picture         =   "ControlsScreen.frx":D530A
      Tag             =   "PrecisionVision"
      ToolTipText     =   "Precision Vision"
      Top             =   720
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsView 
      Height          =   720
      Index           =   1
      Left            =   7440
      Picture         =   "ControlsScreen.frx":D6FD4
      Tag             =   "SnapShot"
      ToolTipText     =   "SnapShot"
      Top             =   720
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsView 
      Height          =   720
      Index           =   0
      Left            =   8280
      Picture         =   "ControlsScreen.frx":D8C9E
      Tag             =   "3D"
      ToolTipText     =   "3D"
      Top             =   720
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsFixed 
      Height          =   360
      Index           =   8
      Left            =   120
      Picture         =   "ControlsScreen.frx":DA968
      Tag             =   "SpecialButtons"
      ToolTipText     =   "Additional Tools"
      Top             =   120
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image ControlsFixed 
      Height          =   480
      Index           =   7
      Left            =   2520
      Picture         =   "ControlsScreen.frx":DB0D2
      Tag             =   "DisplayRefShape"
      ToolTipText     =   "Display Reference Shape"
      Top             =   105
      Width           =   480
   End
   Begin VB.Image ControlsFixed 
      Height          =   360
      Index           =   6
      Left            =   10200
      Picture         =   "ControlsScreen.frx":DB99C
      Tag             =   "DisplayPrecisionVision"
      ToolTipText     =   "Display Precision Vision Settings"
      Top             =   120
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image ControlsFixed 
      Height          =   360
      Index           =   5
      Left            =   7800
      Picture         =   "ControlsScreen.frx":DC106
      Tag             =   "SnapShot"
      ToolTipText     =   "Capture image"
      Top             =   2040
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image ControlsFixed 
      Height          =   360
      Index           =   4
      Left            =   9720
      Picture         =   "ControlsScreen.frx":DC870
      Tag             =   "StepRewind"
      ToolTipText     =   "Step Rewind"
      Top             =   120
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image ControlsFixed 
      Height          =   360
      Index           =   3
      Left            =   9240
      Picture         =   "ControlsScreen.frx":DCFDA
      Tag             =   "StepForward"
      ToolTipText     =   "Step Forward"
      Top             =   120
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image ControlsFixed 
      Height          =   360
      Index           =   2
      Left            =   8760
      Picture         =   "ControlsScreen.frx":DD744
      Tag             =   "PlayPause"
      ToolTipText     =   "Play/Pause"
      Top             =   120
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image ControlsFixed 
      Height          =   360
      Index           =   1
      Left            =   8280
      Picture         =   "ControlsScreen.frx":DDEAE
      Tag             =   "Record"
      Top             =   120
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image ControlsFixed 
      Height          =   720
      Index           =   0
      Left            =   6840
      Picture         =   "ControlsScreen.frx":DE618
      Stretch         =   -1  'True
      Top             =   60
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image ControlViewHighlight 
      Height          =   960
      Left            =   8200
      Picture         =   "ControlsScreen.frx":E13A2
      Top             =   645
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label ControlsStoredReportsLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
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
      Height          =   255
      Index           =   7
      Left            =   6000
      TabIndex        =   7
      Top             =   12000
      Width           =   720
   End
   Begin VB.Label ControlsStoredReportsLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "7"
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
      Height          =   255
      Index           =   6
      Left            =   5160
      TabIndex        =   6
      Top             =   12000
      Width           =   720
   End
   Begin VB.Label ControlsStoredReportsLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "6"
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
      Height          =   255
      Index           =   5
      Left            =   4320
      TabIndex        =   5
      Top             =   12000
      Width           =   720
   End
   Begin VB.Label ControlsStoredReportsLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "5"
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
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   4
      Top             =   12000
      Width           =   720
   End
   Begin VB.Label ControlsStoredReportsLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "4"
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
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   3
      Top             =   12000
      Width           =   720
   End
   Begin VB.Label ControlsStoredReportsLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
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
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   2
      Top             =   12000
      Width           =   720
   End
   Begin VB.Label ControlsStoredReportsLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "2"
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
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   12000
      Width           =   720
   End
   Begin VB.Label ControlsStoredReportsLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
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
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   12000
      Width           =   720
   End
   Begin VB.Image ControlHightlightDepressed 
      Height          =   960
      Left            =   5060
      Picture         =   "ControlsScreen.frx":E41E4
      Top             =   675
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Image ControlHightlight 
      Height          =   960
      Left            =   870
      Picture         =   "ControlsScreen.frx":E7026
      Top             =   645
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Image ControlDrawHighlightDepressed 
      Height          =   600
      Left            =   3480
      Picture         =   "ControlsScreen.frx":E9E68
      Top             =   75
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image ControlDrawHighlight 
      Height          =   600
      Left            =   2160
      Picture         =   "ControlsScreen.frx":EA541
      Top             =   75
      Visible         =   0   'False
      Width           =   540
   End
End
Attribute VB_Name = "ControlsScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub d3d_camselect Lib "threedim.dll" (ByVal cameranumber As Long)
Private Declare Sub d3d_scene_on_off Lib "threedim.dll" (ByVal scenenumber As Long)
Private Declare Sub d3d_play_speed Lib "threedim.dll" (ByVal playspeed As Long)
Private Declare Sub d3d_pipe_scale Lib "threedim.dll" (ByVal X As Long)
Private Declare Sub d3d_reset Lib "threedim.dll" ()
Private Declare Sub d3d_zoom_speed Lib "threedim.dll" (ByVal zoomspeed As Long)
Private Declare Sub d3d_rotate_pipe_z Lib "threedim.dll" (ByVal deg As Long)
Private Declare Sub d3d_rotate_pipe_y Lib "threedim.dll" (ByVal deg As Long)

Dim TimerTicks As Long

Public PipeScale3D As Integer


Private Sub Controls3D_Click(Index As Integer)
On Error GoTo Err_Handler

Call Execute3DButton(Index)

Exit Sub
Err_Handler:
    MsgBox Err & "-CS1:" & Error$
End Sub

Private Sub Controls3D_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


Me.ControlHightlightDepressed.Left = Me.ControlHightlight.Left
Me.ControlHightlightDepressed.Visible = True
Me.ControlHightlight.Visible = False

Call Execute3DButtonMouseDown(Index)

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS2:" & Error$
End Sub

Private Sub Controls3D_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ControlHightlight.Left = Me.Controls3D(Index).Left - 100

If Me.ControlHightlight.Visible = False Then Me.ControlHightlight.Visible = True

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS3:" & Error$
End Sub

Private Sub Controls3D_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ControlHightlight.Visible = True
Me.ControlHightlightDepressed.Visible = False

Call Execute3DButtonMouseUp(Index)

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS4:" & Error$
End Sub

Private Sub ControlsCalibration_Click(Index As Integer)
On Error GoTo Err_Handler

Call ExecuteCalibrationButton(Index)

Exit Sub
Err_Handler:
    MsgBox Err & "-CS5:" & Error$
End Sub

Private Sub ControlsCalibration_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


Me.ControlHightlightDepressed.Left = Me.ControlHightlight.Left
Me.ControlHightlightDepressed.Visible = True
Me.ControlHightlight.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS6:" & Error$
End Sub

Private Sub ControlsCalibration_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

If CalibrationMethodActioned <> "CalibrationH" And CalibrationMethodActioned <> "CalibrationV" And Index = 1 Then
    If Me.ControlHightlight.Visible = True Then Me.ControlHightlight.Visible = False
    Exit Sub
ElseIf Index = 2 Then
    Exit Sub
End If

Me.ControlHightlight.Left = Me.ControlsCalibration(Index).Left - 100

If Me.ControlHightlight.Visible = False Then Me.ControlHightlight.Visible = True

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS7:" & Error$
End Sub

Private Sub ControlsCalibration_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

'PCN4277
'Me.ControlHightlight.Visible = True
'Me.ControlHightlightDepressed.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS8:" & Error$
End Sub

Private Sub ControlsDraw_Click(Index As Integer)
On Error GoTo Err_Handler

Call ExecuteControlsDrawButton(Index)
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS9:" & Error$
End Sub

Private Sub ControlsDraw_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

'PCN4277
'Me.ControlHightlightDepressed.Left = Me.ControlHightlight.Left
'Me.ControlHightlightDepressed.Visible = True
'Me.ControlHightlight.Visible = False

If Me.ControlDrawHighlightDepressed.Left <> Me.ControlDrawHighlight.Left Then Me.ControlDrawHighlightDepressed.Left = Me.ControlDrawHighlight.Left
If Me.ControlDrawHighlightDepressed.Visible = False Then Me.ControlDrawHighlightDepressed.Visible = True
'Me.ControlDrawHightlight.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS10:" & Error$
End Sub

Private Sub ControlsDraw_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

'PCN4277
'Me.ControlHightlight.Left = Me.ControlsDraw(Index).Left - 100
If Me.ControlDrawHighlight.Left <> Me.ControlsDraw(Index).Left Then Me.ControlDrawHighlight.Left = Me.ControlsDraw(Index).Left


'PCN4277
'If Me.ControlHightlight.Visible = False Then Me.ControlHightlight.Visible = True
If Me.ControlDrawHighlight.Visible = False Then Me.ControlDrawHighlight.Visible = True
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS11:" & Error$
End Sub

Private Sub ControlsDraw_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

'PCN4277
'Me.ControlHightlight.Visible = True
'Me.ControlHightlightDepressed.Visible = False
'Me.ControlDrawHightlight.Visible = True
If Me.ControlDrawHighlightDepressed.Visible = True Then Me.ControlDrawHighlightDepressed.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Error$ & " - " & Err
End Sub

Private Sub ControlsFixed_DblClick(Index As Integer)
On Error GoTo Err_Handler

'Execute button - was require at this location to trap 'Double Click' events.
Call ExecuteControlsFixedButton(Index)

Exit Sub
Err_Handler:
    MsgBox Err & "-CS12:" & Error$
End Sub

Private Sub ControlsFixed_Click(Index As Integer)
On Error GoTo Err_Handler

Call ExecuteControlsFixedButton(Index)

Exit Sub
Err_Handler:
    MsgBox Err & "-CS13:" & Error$
End Sub

Private Sub ControlsFixed_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

'If Index > 0 Then
'    Set Me.ControlsFixed(0).Picture = Me.ControlsFixedImageList.ListImages("Depress").Picture
'End If
If Me.ControlDrawHighlightDepressed.Left <> Me.ControlDrawHighlight.Left Then Me.ControlDrawHighlightDepressed.Left = Me.ControlDrawHighlight.Left
If Me.ControlDrawHighlightDepressed.Visible = False Then Me.ControlDrawHighlightDepressed.Visible = True

Call ExecuteControlsFixedMouseDown(Index) 'PCN4335
   
  
Exit Sub
Err_Handler:
    MsgBox Err & "-CS14:" & Error$
End Sub

Private Sub ControlsFixed_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
Dim ShiftHighlight As Single

'If Index > 0 And Me.ControlsFixed(0).Left <> Me.ControlsFixed(Index).Left Then
'    Me.ControlsFixed(0).Left = Me.ControlsFixed(Index).Left
'End If
'If Me.ControlsFixed(0).Visible = False Then Me.ControlsFixed(0).Visible = True

'PCN4277
If Index < 9 Then ShiftHighlight = 90
If Index = 9 Then ShiftHighlight = 90
If Index = 10 Then ShiftHighlight = 90
If Index = 11 Then ShiftHighlight = 60
If Index = 7 Then ShiftHighlight = 45
If Index >= 12 And Index <= 16 Then ShiftHighlight = 90
If Index = 17 Then ShiftHighlight = 45
If Index = 18 Then ShiftHighlight = 45
If Me.ControlDrawHighlight.Left <> Me.ControlsFixed(Index).Left - ShiftHighlight Then Me.ControlDrawHighlight.Left = Me.ControlsFixed(Index).Left - ShiftHighlight
If Me.ControlDrawHighlight.Visible = False Then Me.ControlDrawHighlight.Visible = True


Exit Sub
Err_Handler:
    MsgBox Err & "-CS15:" & Error$
End Sub

Private Sub ControlsFixed_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

'If Index > 0 Then
'    Set Me.ControlsFixed(0).Picture = Me.ControlsFixedImageList.ListImages("Background").Picture
'End If
If Me.ControlDrawHighlightDepressed.Visible = True Then Me.ControlDrawHighlightDepressed.Visible = False
Call ExecuteControlsFixedMouseUp(Index) 'PCN4335
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS16:" & Error$
End Sub






Private Sub ControlsPostRecording_Click(Index As Integer)
On Error GoTo Err_Handler

Call ExecuteControlsPostRecordingButton(Index)

Exit Sub
Err_Handler:
    MsgBox Err & "-CS17:" & Error$
End Sub

Private Sub ControlsPostRecording_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


Me.ControlHightlightDepressed.Left = Me.ControlHightlight.Left
Me.ControlHightlightDepressed.Visible = True
Me.ControlHightlight.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS18:" & Error$
End Sub

Private Sub ControlsPostRecording_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ControlHightlight.Left = Me.ControlsPostRecording(Index).Left - 100

If Me.ControlHightlight.Visible = False Then Me.ControlHightlight.Visible = True

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS19:" & Error$
End Sub

Private Sub ControlsPostRecording_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ControlHightlight.Visible = True
Me.ControlHightlightDepressed.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS20:" & Error$
End Sub

Private Sub ControlsPVDRecording_Click(Index As Integer)
On Error GoTo Err_Handler

Call ExecuteControlsPVDRecordingButton(Index)


Exit Sub
Err_Handler:
    MsgBox Err & "-CS21:" & Error$
End Sub

Private Sub ControlsPVDRecording_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


Me.ControlHightlightDepressed.Left = Me.ControlHightlight.Left
Me.ControlHightlightDepressed.Visible = True
Me.ControlHightlight.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS22:" & Error$
End Sub

Private Sub ControlsPVDRecording_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ControlHightlight.Left = Me.ControlsPVDRecording(Index).Left - 100

If Me.ControlHightlight.Visible = False Then Me.ControlHightlight.Visible = True

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS23:" & Error$
End Sub

Private Sub ControlsPVDRecording_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ControlHightlight.Visible = True
Me.ControlHightlightDepressed.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS24:" & Error$
End Sub


Private Sub ControlsRefShape_Click(Index As Integer)
On Error GoTo Err_Handler

Call ExecuteRefShapeButton(Index)

Exit Sub
Err_Handler:
    MsgBox Err & "-CS25:" & Error$
End Sub

Private Sub ControlsRefShape_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


Me.ControlHightlightDepressed.Left = Me.ControlHightlight.Left
Me.ControlHightlightDepressed.Visible = True
Me.ControlHightlight.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS26:" & Error$
End Sub

Private Sub ControlsRefShape_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ControlHightlight.Left = Me.ControlsRefShape(Index).Left - 100

If Me.ControlHightlight.Visible = False Then Me.ControlHightlight.Visible = True

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS27:" & Error$
End Sub

Private Sub ControlsRefShape_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ControlHightlight.Visible = True
Me.ControlHightlightDepressed.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS28:" & Error$
End Sub

Private Sub ControlsReports_Click(Index As Integer)
On Error GoTo Err_Handler

Call ExecuteReportsButton(Index)

Exit Sub
Err_Handler:
    MsgBox Err & "-CS29:" & Error$
End Sub

Private Sub ControlsReports_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


Me.ControlHightlightDepressed.Left = Me.ControlHightlight.Left
Me.ControlHightlightDepressed.Visible = True
Me.ControlHightlight.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS30:" & Error$
End Sub

Private Sub ControlsReports_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ControlHightlight.Left = Me.ControlsReports(Index).Left - 100

If Me.ControlHightlight.Visible = False Then Me.ControlHightlight.Visible = True

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS31:" & Error$
End Sub

Private Sub ControlsReports_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ControlHightlight.Visible = True
Me.ControlHightlightDepressed.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS32:" & Error$
End Sub

Private Sub ControlsSpecial_Click(Index As Integer)
On Error GoTo Err_Handler

Call ExecuteSpecialButton(Index)

Exit Sub
Err_Handler:
    MsgBox Err & "-CS33:" & Error$
End Sub

Private Sub ControlsSpecial_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


Me.ControlHightlightDepressed.Left = Me.ControlHightlight.Left
Me.ControlHightlightDepressed.Visible = True
Me.ControlHightlight.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS34:" & Error$
End Sub

Private Sub ControlsSpecial_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ControlHightlight.Left = Me.ControlsSpecial(Index).Left - 100

If Me.ControlHightlight.Visible = False Then Me.ControlHightlight.Visible = True

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS35:" & Error$
End Sub

Private Sub ControlsSpecial_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ControlHightlight.Visible = True
Me.ControlHightlightDepressed.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS36:" & Error$
End Sub

Private Sub ControlsStoredReports_Click(Index As Integer)
On Error GoTo Err_Handler

Call ExecuteStoredReportsButton(Index)

Exit Sub
Err_Handler:
    MsgBox Err & "-CS37:" & Error$
End Sub

Private Sub ControlsStoredReports_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


Me.ControlHightlightDepressed.Left = Me.ControlHightlight.Left
Me.ControlHightlightDepressed.Visible = True
Me.ControlHightlight.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS38:" & Error$
End Sub

Private Sub ControlsStoredReports_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ControlHightlight.Left = Me.ControlsStoredReports(Index).Left - 100

If Me.ControlHightlight.Visible = False Then Me.ControlHightlight.Visible = True

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS39:" & Error$
End Sub

Private Sub ControlsStoredReports_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ControlHightlight.Visible = True
Me.ControlHightlightDepressed.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS40:" & Error$
End Sub

Sub ControlsView_Click(Index As Integer)
On Error GoTo Err_Handler
Dim ErrorStr As String

ClearLineScreen.ThreeDtimer = False ' PCN4213 allways off untill clicked SetupCLPScreenTo3D
CLPScreenAction = "" 'PCN4252
ControlsStoredReports(0).Visible = False: ControlsReports(0).Visible = False 'PCN4277

Call ClearLineScreen.SetupMouseIcon(116)
    
Select Case Me.ControlsView(Index).Tag
    Case "3D"
        Call SetupCLPScreenTo3D(ErrorStr)
        If ErrorStr <> "" Then Exit Sub
        CLPScreenMode = ThreeD
    Case "SnapShot"
        Call SetupCLPScreenToSnapShot(ErrorStr)
        If ErrorStr <> "" Then Exit Sub
        CLPScreenMode = SnapShot
    Case "PrecisionVision"
        Call SetupCLPScreenToPV(ErrorStr)
        If ErrorStr <> "" Then Exit Sub
        CLPScreenMode = PV
    Case "Video"
        Call ClearLineScreen.OnscreenTextHide
        Call ClearLineScreen.VideoScreenActivate(ErrorStr) 'PCN4171
        If ErrorStr <> "" Then Exit Sub
'        CLPScreenMode = Video
'        CLPScreenDrawAction = ""
'        CLPScreenDrawState = ""
'        CLPScreenAction = ""
'        Call ClearLineScreen.SetupMouseIcon(116)
End Select

Call ControlsViewSetup
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS41:" & Error$
End Sub

Private Sub ControlsView_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


Me.ControlViewHighlight.Left = Me.ControlsView(Index).Left - 100

If Me.ControlViewHighlight.Visible = False Then Me.ControlViewHighlight.Visible = True

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS42:" & Error$
End Sub


Private Sub ControlsWizard_Click(Index As Integer)
On Error GoTo Err_Handler

Call ExecuteControlsWizardButton(Index)
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS43:" & Error$
End Sub

Private Sub ControlsWizard_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


If ConfigInfo.DistanceStart = InvalidData And Index >= 5 Then
    Exit Sub
ElseIf ConfigInfo.DistanceFinish = InvalidData And Index >= 6 Then
    Exit Sub
End If

Me.ControlHightlightDepressed.Left = Me.ControlHightlight.Left
Me.ControlHightlightDepressed.Visible = True
Me.ControlHightlight.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS44:" & Error$
End Sub

Private Sub ControlsWizard_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

If ConfigInfo.DistanceStart = InvalidData And Index >= 5 Then
    Me.ControlHightlight.Visible = False
    Exit Sub
ElseIf ConfigInfo.DistanceFinish = InvalidData And Index >= 6 Then
    Me.ControlHightlight.Visible = False
    Exit Sub
End If

Me.ControlHightlight.Left = Me.ControlsWizard(Index).Left - 100

If Me.ControlHightlight.Visible = False Then Me.ControlHightlight.Visible = True

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS45:" & Error$
End Sub

Private Sub ControlsWizard_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ControlHightlight.Visible = True
Me.ControlHightlightDepressed.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS46:" & Error$
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
    
Me.Top = ClearLineScreen.height + ClearLineScreen.Top
Me.height = 1650
Me.Left = 0
Me.width = ClearLineScreen.width
Call ConvertLanguage(Me, Language)


Exit Sub
Err_Handler:
    MsgBox Err & "-CS47:" & Error$
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


If Me.ControlsFixed(0).Visible Then Me.ControlsFixed(0).Visible = False
If Me.ControlHightlight.Visible Then Me.ControlHightlight.Visible = False
If Me.ControlViewHighlight.Visible Then Me.ControlViewHighlight.Visible = False
If Me.ControlDrawHighlight.Visible Then Me.ControlDrawHighlight.Visible = False
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS48:" & Error$
End Sub



Sub ControlsViewSetup()
On Error GoTo Err_Handler

'PCN4277


Set Me.ControlsView(0).Picture = Me.ControlsViewImageList.ListImages("3DNotSelected").Picture
Set Me.ControlsView(1).Picture = Me.ControlsViewImageList.ListImages("SnapShotNotSelected").Picture
Set Me.ControlsView(2).Picture = Me.ControlsViewImageList.ListImages("PVNotSelected").Picture
Set Me.ControlsView(3).Picture = Me.ControlsViewImageList.ListImages("VideoNotSelected").Picture

Call HideAllControls

'vvvv PCN4319 ********************************
Call SetupControlsView   'PCN4320
'^^^^ ****************************************



'vvvv PCN3809 ********************************
If SoftwareConfiguration = "Reader" Then
    Call SetupForReaderConfiguration
    'Setup Controls fixed
    If ClearLineScreen.ProfilerTimer.Enabled Then
        Set Me.ControlsFixed(2).Picture = Me.ControlsFixedImageList.ListImages("Pause").Picture
    Else
        Set Me.ControlsFixed(2).Picture = Me.ControlsFixedImageList.ListImages("Play").Picture
    End If
    Exit Sub
End If
'^^^^ ****************************************

Select Case CLPScreenMode
    Case ThreeD
        Set Me.ControlsView(0).Picture = Me.ControlsViewImageList.ListImages("3D").Picture
        Call SetupFor3DWithZoom
    Case SnapShot
        Set Me.ControlsView(3).Picture = Me.ControlsViewImageList.ListImages(Video).Picture 'PCN4319
        If CLPScreenAction = "DrawMaskBox" Or CLPScreenAction = "DrawWaterLevel" Then
            CLPScreenMode = Video
            Call ControlsViewSetup
            CLPScreenMode = SnapShot 'PCN
            Exit Sub
        End If
        If PVDFileName = "" And Not PVRecording Then 'PCN4237
        'If CLPScreenAction = "DrawCalibrationLine" Or CLPScreenAction = "DrawHorCalibrationLine" Then
            CLPScreenMode = SnapShot
            Call SetupForCalibration
            Exit Sub
        End If
        
        Set Me.ControlsView(1).Picture = Me.ControlsViewImageList.ListImages(SnapShot).Picture
        Call SetupForSnapShot
    Case PV
        Set Me.ControlsView(2).Picture = Me.ControlsViewImageList.ListImages(PV).Picture
        Call SetupForPV
    Case Video
        If CLPScreenAction = "DrawCalibrationLine" Or CLPScreenAction = "DrawHorCalibrationLine" Then 'PCN4237
            CLPScreenMode = SnapShot
            Call SetupForCalibration
        Exit Sub
        End If
        Set Me.ControlsView(3).Picture = Me.ControlsViewImageList.ListImages(Video).Picture
        
        If mediatype = "Live" Then
            Call SetupForLive
        Else
            Call SetupForVideo
        End If
        
End Select

'Setup Controls fixed
If ClearLineScreen.ProfilerTimer.Enabled Then
    Set Me.ControlsFixed(2).Picture = Me.ControlsFixedImageList.ListImages("Pause").Picture
Else
    Set Me.ControlsFixed(2).Picture = Me.ControlsFixedImageList.ListImages("Play").Picture
End If

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS49:" & Error$
End Sub

Sub SetupForVideo() 'PCN4319
On Error GoTo Err_Handler

If PVDFileName <> "" And mediatype = "" Then
    Call SetupForPV
ElseIf mediatype = "" Or mediatype = "StillImage" Or mediatype = "Live" Then
    Exit Sub
ElseIf PVDFileName = "" And Not PVRecording Then 'PCN4252
    Call SetupForVideoRecordWizard
ElseIf PVRecording Then
    Call SetupForPVDRecording
ElseIf InStr(1, PVDFileName, "PVDRecording") <> 0 Then
    Call SetupForPostRecording
Else
    Call SetupForVideoWithPVD
End If


Exit Sub
Err_Handler:
    MsgBox Err & "-CS50:" & Error$
End Sub

Sub SetupForLive()
On Error GoTo Err_Handler


    Call Me.SetupForCalibration
    Call Me.SetupControlsDraw
    

Exit Sub
Err_Handler:
    MsgBox Err & "-CS51:" & Error$
End Sub

Sub SetupForVideoRecordWizard()
On Error GoTo Err_Handler
Dim ButtonIndex As Integer

'Setup wizard buttons
For ButtonIndex = 1 To 6
    Me.ControlsWizard(ButtonIndex).Top = 720 '840
    Me.ControlsWizard(ButtonIndex).Visible = True
Next ButtonIndex
'Setup video fixed buttons
Call SetupControlsVideo
Call Me.SetupControlsDraw 'PCN4277

If ConfigInfo.DistanceStart = InvalidData And Not IPD Then
    Me.ControlsWizard(5).Picture = Me.ControlsWizardImageList.ListImages("StepRewindDisabled").Picture
    Me.ControlsWizard(6).Picture = Me.ControlsWizardImageList.ListImages("PVRecordDisabled").Picture
ElseIf ConfigInfo.DistanceFinish = InvalidData And Not IPD Then
    Me.ControlsWizard(5).Picture = Me.ControlsWizardImageList.ListImages("StepRewind").Picture
    Me.ControlsWizard(6).Picture = Me.ControlsWizardImageList.ListImages("PVRecordDisabled").Picture
Else
    Me.ControlsWizard(5).Picture = Me.ControlsWizardImageList.ListImages("StepRewind").Picture
    Me.ControlsWizard(6).Picture = Me.ControlsWizardImageList.ListImages("PVRecord").Picture
End If

'PCN4277'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If CalibrationMethodActioned = "CalibrationH" Or CalibrationMethodActioned = "CalibrationV" Then
    Set Me.ControlsCalibration(1).Picture = Me.ControlsCalibrationImageList.ListImages("CalibrationV").Picture
Else                                                                        '
    Set Me.ControlsCalibration(1).Picture = Me.ControlsCalibrationImageList.ListImages("CalibrationVNotSelected").Picture
End If                                                                      '
ControlsCalibration(0).Top = 720:   ControlsCalibration(0).Visible = True   '
ControlsCalibration(1).Top = 720:   ControlsCalibration(1).Visible = True   '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If CLPScreenAction = "DrawCalibrationLine" Then
    If Me.ControlHightlightDepressed.Left <> Me.ControlsCalibration(0).Left - 100 Then Me.ControlHightlightDepressed.Left = Me.ControlsCalibration(0).Left - 100
    If Not Me.ControlHightlightDepressed.Visible Then Me.ControlHightlightDepressed.Visible = True
End If

If CLPScreenAction = "DrawHorCalibrationLine" Then
    If Me.ControlHightlightDepressed.Left - 100 <> Me.ControlsCalibration(1).Left Then Me.ControlHightlightDepressed.Left = Me.ControlsCalibration(1).Left - 100
    If Not Me.ControlHightlightDepressed.Visible Then Me.ControlHightlightDepressed.Visible = True
End If


Exit Sub
Err_Handler:
    MsgBox Err & "-CS52:" & Error$
End Sub

Sub SetupForVideoWithPVD()
On Error GoTo Err_Handler
Dim ButtonIndex As Integer



'Setup video fixed buttons
Call SetupControlsVideo
Call Me.SetupControlsDraw 'PCN4277
''Disable the IP display button
'Me.ControlsFixed(6).Visible = False


Exit Sub
Err_Handler:
    MsgBox Err & "-CS53:" & Error$
End Sub

Sub SetupForSnapShot()
On Error GoTo Err_Handler
Call Me.SetupControlsDraw 'PCN4277
'Dim ButtonIndex As Integer
'
'For ButtonIndex = 1 To 8
'    Me.ControlsDraw(ButtonIndex).Top = 120 'PCN4277 '720 '840
'    Me.ControlsDraw(ButtonIndex).Visible = True
'Next ButtonIndex
'
'If ScreenDrawing.ShowReferenceShape = False Then
'    Me.ControlsDraw(5).Visible = False
'    Me.ControlsDraw(6).Visible = False
'End If

'Setup video fixed buttons
Call SetupControlsVideo
'Enable the RefShape display button
Me.ControlsFixed(7).Visible = True
''Disable the IP display button
'Me.ControlsFixed(6).Visible = False

Call SetupControlZoomOnSnap

Exit Sub
Err_Handler:
    MsgBox Err & "-CS54:" & Error$
End Sub

Sub SetupForPV()
On Error GoTo Err_Handler
Call Me.SetupControlsDraw 'PCN4277

Call Me.SetupControlsVideo



Exit Sub
Err_Handler:
    MsgBox Err & "-CS55:" & Error$
End Sub

Sub SetupForLoadedImage()
On Error GoTo Err_Handler
    Call Me.SetupControlsDraw
Exit Sub
Err_Handler:
    MsgBox Err & "-CS56:" & Error$
End Sub

Sub SetupFor3DWithZoom()
On Error GoTo Err_Handler
Dim ButtonIndex As Integer

For ButtonIndex = 0 To 6
    Me.Controls3D(ButtonIndex).Top = 720 '840
    Me.Controls3D(ButtonIndex).Visible = True
Next ButtonIndex
'Hide the rotate buttons
For ButtonIndex = 7 To 10
    Me.Controls3D(ButtonIndex).Visible = False
Next ButtonIndex
'Hide Zoom button
Me.ControlsFixed(18).Visible = False

'Setup 3D fixed buttons
Call SetupControls3D
'Call Me.SetupControlsDraw 'PCN4277 now undo PCN4277 by commenting this out

Exit Sub
Err_Handler:
    MsgBox Err & "-CS57:" & Error$
End Sub

Sub SetupFor3DWithRotate()
On Error GoTo Err_Handler
Dim ButtonIndex As Integer

For ButtonIndex = 0 To 3
    Me.Controls3D(ButtonIndex).Top = 720 '840
    Me.Controls3D(ButtonIndex).Visible = True
Next ButtonIndex
'Hide the Zoom buttons
For ButtonIndex = 4 To 6
    Me.Controls3D(ButtonIndex).Visible = False
Next ButtonIndex
'Show the rotate buttons
For ButtonIndex = 7 To 10
    Me.Controls3D(ButtonIndex).Top = 720 '840
    Me.Controls3D(ButtonIndex).Visible = True
Next ButtonIndex
'Setup 3D fixed buttons
Call SetupControls3D

Exit Sub
Err_Handler:
    MsgBox Err & "-CS58:" & Error$
End Sub


Sub SetupForPVDRecording()
On Error GoTo Err_Handler
Dim ButtonIndex As Integer

For ButtonIndex = 0 To 4
    If ButtonIndex <> 2 Then
        Me.ControlsPVDRecording(ButtonIndex).Top = 720 '840
        Me.ControlsPVDRecording(ButtonIndex).Visible = True
    End If
Next ButtonIndex


Exit Sub
Err_Handler:
    MsgBox Err & "-CS59:" & Error$
End Sub

Sub SetupForPostRecording()
On Error GoTo Err_Handler
Dim ButtonIndex As Integer
Call Me.SetupControlsDraw 'PCN4277

For ButtonIndex = 0 To 1
    Me.ControlsPostRecording(ButtonIndex).Top = 720 '840
    Me.ControlsPostRecording(ButtonIndex).Visible = True
Next ButtonIndex


Exit Sub
Err_Handler:
    MsgBox Err & "-CS60:" & Error$
End Sub


Sub SetupForReports()
On Error GoTo Err_Handler
Dim ButtonIndex As Integer
Dim ErrorStr As String

'If PVDFileName = "" Or PVRecording Then
If mediatype = "" And PVDFileName = "" Or PVRecording Then 'PCN4552
    ErrorStr = "Must have a pvd file loaded"
    'MsgBox DisplayMessage(ErrorStr)
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage(ErrorStr): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Exit Sub
End If

Call HideAllControls
Call Me.SetupControlsDraw

'If PVDFileName <> "" And PVRecording = False Then Me.ControlsView(0).Visible = True: Me.ControlsView(2).Visible = True
'If mediatype <> "" Then Me.ControlsView(3).Visible = True
Call SetupControlsView   'PCN4320

'PCNLIVE
For ButtonIndex = 0 To 1
    ControlsCalibration(ButtonIndex).Visible = False
Next ButtonIndex

For ButtonIndex = 0 To 5
    Me.ControlsReports(ButtonIndex).Top = 720 '840
    Me.ControlsReports(ButtonIndex).Visible = True
Next ButtonIndex


Exit Sub
Err_Handler:
    MsgBox Err & "-CS61:" & Error$
End Sub

Sub SetupForStoredReports()
On Error GoTo Err_Handler
Dim ButtonIndex As Integer
Dim ErrorStr As String
Dim NoStoredReports As Integer
Dim CurrentReportType As Integer

If SoftwareConfiguration = "Reader" And PVDFileName = "" Then
    Exit Sub
ElseIf PVDFileName = "" Or PVRecording Then
    ErrorStr = "Must have a pvd file loaded"
    'MsgBox DisplayMessage(ErrorStr)
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage(ErrorStr): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Exit Sub
End If

Call HideAllControls

Call SetupControlsView   'PCN4320

NoStoredReports = UBound(StoredReportArray)

If NoStoredReports <= 0 Then Exit Sub

'Limit the reports to 7 at this stage
If NoStoredReports > 8 Then NoStoredReports = 8

For ButtonIndex = 0 To NoStoredReports - 1
    Me.ControlsStoredReports(ButtonIndex).Top = 720 '840
    CurrentReportType = StoredReportArray(ButtonIndex + 1).ReportType
    If CurrentReportType = 0 Then CurrentReportType = 1
    Set Me.ControlsStoredReports(ButtonIndex).Picture = Me.StoredReportsImageList.ListImages(CurrentReportType).Picture
    Me.ControlsStoredReports(ButtonIndex).ToolTipText = StoredReportArray(ButtonIndex + 1).Title
    Me.ControlsStoredReports(ButtonIndex).Tag = StoredReportArray(ButtonIndex + 1).ReportNumber
    Me.ControlsStoredReports(ButtonIndex).Visible = True
    Me.ControlsStoredReportsLabel(ButtonIndex).Top = 950
    Me.ControlsStoredReportsLabel(ButtonIndex).Visible = True
Next ButtonIndex


Exit Sub
Err_Handler:
    MsgBox Err & "-CS62:" & Error$
End Sub

Sub SetupForCalibration()
On Error GoTo Err_Handler
Dim ButtonIndex As Integer
Dim ErrorStr As String

Call HideAllControls
If Me.ControlsView(3).Visible = False Then Me.ControlsView(3).Visible = True

If CalibrationMethodActioned = "CalibrationH" Or CalibrationMethodActioned = "CalibrationV" Then
    Set Me.ControlsCalibration(1).Picture = Me.ControlsCalibrationImageList.ListImages("CalibrationV").Picture
Else
    Set Me.ControlsCalibration(1).Picture = Me.ControlsCalibrationImageList.ListImages("CalibrationVNotSelected").Picture
End If
'Crack calibration disabled
Set Me.ControlsCalibration(2).Picture = Me.ControlsCalibrationImageList.ListImages("CalibrationCrackNotSelected").Picture


For ButtonIndex = 0 To 1 'PCN4277 was 2 for crack calibration
    Me.ControlsCalibration(ButtonIndex).Top = 720 '840
    Me.ControlsCalibration(ButtonIndex).Visible = True
Next ButtonIndex

If mediatype = Video Then
    Call Me.SetupForVideo 'PCN4319
ElseIf mediatype = "StillImage" Or mediatype = "Live" Then 'PCN4406
     Call Me.SetupControlsDraw
End If





Exit Sub
Err_Handler:
    MsgBox Err & "-CS63:" & Error$
End Sub

Sub SetupForSpecialButtons()
On Error GoTo Err_Handler
Dim ButtonIndex As Integer
Dim ErrorStr As String

If mediatype = "" Then
    ErrorStr = "Must load a Media file first!"
    'MsgBox DisplayMessage(ErrorStr)
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage(ErrorStr): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Exit Sub
End If

Call HideAllControls

For ButtonIndex = 0 To 3
    Me.ControlsSpecial(ButtonIndex).Top = 720 '840
    Me.ControlsSpecial(ButtonIndex).Visible = True
Next ButtonIndex


Exit Sub
Err_Handler:
    MsgBox Err & "-CS64:" & Error$
End Sub

Sub SetupForRefShapeButtons()
On Error GoTo Err_Handler
Dim ButtonIndex As Integer
Dim ErrorStr As String

Call HideAllControls

For ButtonIndex = 0 To 3
    Me.ControlsRefShape(ButtonIndex).Top = 720 '840
    Me.ControlsRefShape(ButtonIndex).Visible = True
Next ButtonIndex


Exit Sub
Err_Handler:
    MsgBox Err & "-CS65:" & Error$
End Sub

Sub HideAllControls()
On Error GoTo Err_Handler
Dim ButtonIndex As Integer

'Video Record Wizard controls
For ButtonIndex = 0 To 6
    Me.ControlsWizard(ButtonIndex).Visible = False
Next ButtonIndex

'Recording
ButtonIndex = 0
For ButtonIndex = 0 To 4
    Me.ControlsPVDRecording(ButtonIndex).Visible = False
Next ButtonIndex

'Post recording
ButtonIndex = 0
For ButtonIndex = 0 To 1
    Me.ControlsPostRecording(ButtonIndex).Visible = False
Next ButtonIndex

'SnapShot and PV controls
ButtonIndex = 0
For ButtonIndex = 1 To 8
    Me.ControlsDraw(ButtonIndex).Visible = False
Next ButtonIndex

'Reports
ButtonIndex = 0
For ButtonIndex = 0 To 5
    Me.ControlsReports(ButtonIndex).Visible = False
Next ButtonIndex

'Stored Reports
ButtonIndex = 0
For ButtonIndex = 0 To 7
    Me.ControlsStoredReports(ButtonIndex).Visible = False
    Me.ControlsStoredReportsLabel(ButtonIndex).Visible = False
Next ButtonIndex

'Highlight features
Me.ControlHightlight.Visible = False
Me.ControlHightlightDepressed.Visible = False

'Calibration
ButtonIndex = 0
For ButtonIndex = 0 To 2
    Me.ControlsCalibration(ButtonIndex).Visible = False
Next ButtonIndex

'Special buttons
ButtonIndex = 0
For ButtonIndex = 0 To 3
    Me.ControlsSpecial(ButtonIndex).Visible = False
Next ButtonIndex

'Ref Shapes
ButtonIndex = 0
For ButtonIndex = 0 To 3
    Me.ControlsRefShape(ButtonIndex).Visible = False
Next ButtonIndex

'3D buttons
ButtonIndex = 0
For ButtonIndex = 0 To 10
    Me.Controls3D(ButtonIndex).Visible = False
Next ButtonIndex
ButtonIndex = 12
For ButtonIndex = 9 To 16 '12 To 16, did not include the 3d tools
    Me.ControlsFixed(ButtonIndex).Visible = False
Next ButtonIndex

'PCN4277
For ButtonIndex = 0 To 3
    Me.ControlsView(ButtonIndex).Visible = False
Next ButtonIndex


'Zoom display window
ClearLineScreen.MagnifyPictContainer.ZOrder 1 'Hide



Exit Sub
Err_Handler:
    MsgBox Err & "-CS66:" & Error$
End Sub

Sub SetupControlDisplayIP()
On Error GoTo Err_Handler
Dim DisplayIP As Long
        
Call getimageanalysis(DisplayIP)
If DisplayIP = 0 Then
    Set Me.ControlsFixed(6).Picture = Me.ControlsFixedImageList.ListImages("IP_OFF").Picture
Else
    Set Me.ControlsFixed(6).Picture = Me.ControlsFixedImageList.ListImages("IP_ON").Picture
End If
Me.ControlsFixed(6).Visible = True


Exit Sub
Err_Handler:
    MsgBox Err & "-CS67:" & Error$
End Sub

Sub SetDisplayIPSettings()
On Error GoTo Err_Handler
Dim DisplayIP As Long
Dim ErrorStr As String
        
If CLPScreenMode <> Video Then Call ClearLineScreen.VideoScreenActivate(ErrorStr)  'PCN4429
        
Call getimageanalysis(DisplayIP)
'Toggle the current setting
If DisplayIP <> 0 Then
    Call SetupVideoDisplayAsNormal
Else
    Call SetupVideoDisplayForPVTuning
End If
Call SetupControlDisplayIP
If IsOpen("AutoTune") Then Call AutoTune.SetupIPCheckBox


Exit Sub
Err_Handler:
    MsgBox Err & "-CS68:" & Error$

End Sub

Sub SetupControlZoomOnSnap()
On Error GoTo Err_Handler
    
If ClearLineScreen.MagnifyPictContainer.Visible Then
    Set Me.ControlsFixed(18).Picture = Me.ControlsFixedImageList.ListImages("Zoom_On").Picture
    If CLPScreenMode = SnapShot Or _
       CLPScreenAction = "DrawCalibrationLine" Or _
       CLPScreenAction = "DrawHorCalibrationLine" Then
        ClearLineScreen.MagnifyPictContainer.ZOrder 0
        Call ClearLineScreen.MagnifyUpdateImage
    End If
Else
    Set Me.ControlsFixed(18).Picture = Me.ControlsFixedImageList.ListImages("Zoom_Off").Picture
End If
Me.ControlsFixed(18).Visible = True


Exit Sub
Err_Handler:
    MsgBox Err & "-CS69:" & Error$
End Sub


Sub SetDisplayZoomOnSnap(ZoomState As String)
On Error GoTo Err_Handler
    

If ClearLineScreen.MagnifyPictContainer.Visible = False And ZoomState = "Toggle" Then
    ClearLineScreen.MagnifyPictContainer.Visible = True

ElseIf ClearLineScreen.MagnifyPictContainer.Visible = True And ZoomState = "Toggle" Then
    ClearLineScreen.MagnifyPictContainer.Visible = False
ElseIf ZoomState = "On" Then
    ClearLineScreen.MagnifyPictContainer.Visible = True
Else
    ClearLineScreen.MagnifyPictContainer.Visible = False
End If
Call SetupControlZoomOnSnap

Exit Sub
Err_Handler:
    MsgBox Err & "-CS70:" & Error$
End Sub


Sub SetupControlsVideo()
On Error GoTo Err_Handler
Dim ButtonIndex As Integer
Dim VideoIsAvailable As Boolean

If mediatype = Video Or PVDFileName <> "" Then
    VideoIsAvailable = True
Else
    VideoIsAvailable = False
End If

'Video Record Wizard controls
For ButtonIndex = 2 To 4 'PCN4277 was 5
    Me.ControlsFixed(ButtonIndex).Visible = VideoIsAvailable
Next ButtonIndex

'Disable the 3D controls
'Setup 3D controls
For ButtonIndex = 9 To 16
    Me.ControlsFixed(ButtonIndex).Visible = False
Next ButtonIndex

Me.ControlsFixed(6).Visible = VideoIsAvailable 'PCN4277
'Disable the RefShape display button
Me.ControlsFixed(7).Visible = True 'False
'Enable the IP display button
If VideoIsAvailable Then Call SetupControlDisplayIP 'PCN4277

'Show Zoom button
Me.ControlsFixed(18).Visible = True

Exit Sub
Err_Handler:
    MsgBox Err & "-CS71:" & Error$
End Sub

Sub SetupControls3D()
On Error GoTo Err_Handler
Dim ButtonIndex As Integer

'Disable the Video Record Wizard controls
For ButtonIndex = 1 To 7
    Me.ControlsFixed(ButtonIndex).Visible = False
Next ButtonIndex

'Setup 3D controls
For ButtonIndex = 9 To 16
    Me.ControlsFixed(ButtonIndex).Top = Me.ControlsFixed(0).Top
    Me.ControlsFixed(ButtonIndex).Visible = True
Next ButtonIndex

Exit Sub
Err_Handler:
    MsgBox Err & "-CS72:" & Error$
End Sub

Sub SetupControlsDraw()
'****************************************************************************************
'Name    : SetupControlsDraw
'PCN     : PCN4277
'Created : 2 October 2006
'Updated :
'Prg By  : Antony van Iersel
'Param   :
'Desc    : Setup for the drawing icons
'****************************************************************************************
On Error GoTo Err_Handler

Dim ButtonIndex As Integer

For ButtonIndex = 1 To 8
    Me.ControlsDraw(ButtonIndex).Top = 120 'PCN4277 '720 '840
    Me.ControlsDraw(ButtonIndex).Visible = True
    Me.ControlsDraw(4).Visible = False
Next ButtonIndex

If ScreenDrawing.ShowReferenceShape = False Then
    Me.ControlsDraw(5).Visible = False
    Me.ControlsDraw(6).Visible = False
End If


Exit Sub
Err_Handler:
    MsgBox Err & "-CS73:" & Error$
End Sub

Sub ExecuteControlsWizardButton(Index As Integer)
On Error GoTo Err_Handler
Dim ErrorStr As String

Call AutoTune.SetupSelectedTask("")

Select Case Me.ControlsWizard(Index).Tag
    Case "Calibration"
        Call SetupForCalibration
        
    Case "IPSettings"
        Call SetPVSettings
        
    Case "SetVideoMask"
        Call SetVideoMask
        
    Case "SetWaterLevel"
        If PVRecording Then Exit Sub 'PCN4216
        Call SetWaterLevel
        
    Case "SetStartRecordingMarker"
        Call ClearLineScreen.SetVideoRecordMarkerStartAdjuster
'        If ConfigInfo.DistanceStart = InvalidData Or ConfigInfo.DistanceFinish = InvalidData Then
        If Not IPD Then
            Call DataEntryForm.SetUpStartFinishDistances("Start")
        End If
    
    Case "SetFinishRecordingMarker"
        If ConfigInfo.DistanceStart <> InvalidData Or IPD Then
            Call ClearLineScreen.SetVideoRecordMarkerStopAdjuster
'           If ConfigInfo.DistanceStart = InvalidData Or ConfigInfo.DistanceFinish = InvalidData Then
            If Not IPD Then
                Call DataEntryForm.SetUpStartFinishDistances("Finish")
            End If
        End If
        
    Case "StartRecording"
        Call StartStopRecording(ErrorStr)
        If ErrorStr <> "" Then Exit Sub
        'Set new controls
        Call HideAllControls
        Call SetupForPVDRecording
        
        
End Select
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS74:" & Error$
End Sub

Sub ExecuteControlsPVDRecordingButton(Index As Integer)
On Error GoTo Err_Handler

Call AutoTune.SetupSelectedTask("")

Select Case Me.ControlsPVDRecording(Index).Tag
    Case "IPSettings"
        Call SetPVSettings
    
    Case "SetVideoMask"
        Call SetVideoMask
        
    Case "SetWaterLevel"
        Call SetWaterLevel

    Case "DisplayPipeDetails"
        Call ControlsMain.ControlsDisplaySetup("DisplayPipeDetails")
        
    Case "StopRecording"
        If PVRecording Then
            Call ClearLineScreen.RecordPVProfile_Stop
            Call ControlsViewSetup
        End If


End Select

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS75:" & Error$
End Sub


Sub ExecuteControlsPostRecordingButton(Index As Integer)
On Error GoTo Err_Handler

Select Case Me.ControlsPostRecording(Index).Tag
    Case "SaveToFile"
        Call SaveThePVDRecordingToFile
        
    Case "DisplayPipeDetails"
        Load PipelineDetails
        PipelineDetails.Show
        PipelineDetails.ZOrder 0

End Select

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS76:" & Error$
End Sub

Sub ExecuteControlsFixedButton(Index As Integer)
On Error GoTo Err_Handler
Dim SnapFileName As String 'The filename and directory of the BMP created when the 3D window is captured.
Dim ErrorStr As String

MainScreenMouseIcon = 116
ClearLineScreen.PVScreen.MousePointer = 2
ClearLineScreen.SnapShotScreen.MousePointer = 2 'PCN4252
If IsOpen("AutoTune") Then 'PCN4171
Call AutoTune.SetupSelectedTask("") 'PCN4252
End If
'Call ClearLineScreen.SetupMouseIcon(0) 'PCN4252 - Reset the mouse cursor 'PCN4277

Select Case Me.ControlsFixed(Index).Tag
    Case "Record", "Stop"
        Call StartStopRecording(ErrorStr)
        If ErrorStr <> "" Then Exit Sub

    Case "PlayPause"
        Call ClearLineScreen.ProfilerPausePlay

    Case "StepForward"
        Call ClearLineScreen.ProfilerStepFoward
        
    Case "StepRewind"
        Call ClearLineScreen.ProfilerStepRewind
    
    Case SnapShot
        Call ClearLineScreen.TakeASnapShot
        Call ControlsViewSetup
    
    Case "Zoom"
        Call SetDisplayZoomOnSnap("Toggle")
        
    Case "DisplayPrecisionVision"
        Call SetDisplayIPSettings
        
    Case "DisplayRefShape"
        If Len(PipelineDetails.InternalDiameterExpected) = 0 Then
            PipelineDetails.ZOrder 0
            ErrorStr = "Please specify internal pipe diameter first."
            'MsgBox DisplayMessage(ErrorStr), vbExclamation
            ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage(ErrorStr): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
            PipelineDetails.ZOrder (0)
            PipelineDetails.InternalDiameterExpected.SetFocus
            Exit Sub
        End If
        
        
        
        If ScreenDrawing.ShowReferenceShape Then
            If CLPScreenMode = Video Or CLPScreenMode = ThreeD Then
                ScreenDrawing.ShowReferenceShape = True
                Set Me.ControlsFixed(Index).Picture = Me.ControlsFixedImageList.ListImages("RefShapeON").Picture
                ClearLineScreen.TakeASnapShot
            Else
                ScreenDrawing.ShowReferenceShape = False
                Set Me.ControlsFixed(Index).Picture = Me.ControlsFixedImageList.ListImages("RefShapeOFF").Picture
            End If
        Else
            ScreenDrawing.ShowReferenceShape = True
            Set Me.ControlsFixed(Index).Picture = Me.ControlsFixedImageList.ListImages("RefShapeON").Picture
            If CLPScreenMode = Video Or CLPScreenMode = ThreeD Then Call ClearLineScreen.TakeASnapShot
        End If
        Call ScreenDrawing.DrawProfilesStartToFinish(ClearLineScreen.SnapShotScreen, True)
        Call ScreenDrawing.DrawProfilesStartToFinish(ClearLineScreen.PVScreen, True)
        
        'PCN4277
        If ControlsStoredReports(0).Visible = True Or ControlsReports(0).Visible = True Then
            SetupForReports
        Else
            ControlsViewSetup
        End If
        
    Case "Controls3DZoom"
        Call SetupFor3DWithZoom
        
    Case "Controls3DController"
        Call SetupFor3DWithRotate
    
    Case "Controls3DMapOnOff"
        Call d3d_scene_on_off(0)
        If D3D_MapOn Then
            'Map off
            Set Me.ControlsFixed(Index).Picture = Me.ControlsFixedImageList.ListImages("3DMapOff").Picture
            D3D_MapOn = False
        Else
            'map on
            Set Me.ControlsFixed(Index).Picture = Me.ControlsFixedImageList.ListImages("3DMapOn").Picture
            D3D_MapOn = True
        End If
    Case "Controls3DSnapShot"
        SnapFileName = LocToSave & "Snapshot.bmp"
        Call ClearLineScreen.D3D_CaptureImage(SnapFileName)
        ClearLineScreen.LoadImage (SnapFileName)
        Call ControlsView_Click(0) 'Once taken snapshot go back to three dim PCN4559
    
    Case "Controls3DReset"
        Call d3d_reset
    
    Case "Controls3DRewind"
        If laserspeed = -1 Then
            laserspeed = -2
        ElseIf laserspeed = -2 Then
            laserspeed = -3
        Else
            laserspeed = -1
        End If
        
    Case "Controls3DPlay"
        If laserspeed = 1 Then
            laserspeed = 2
        ElseIf laserspeed = 2 Then
            laserspeed = 3
        Else
            laserspeed = 1
        End If
    
    Case "Controls3DPause"
        laserspeed = 0
    
    Case "Reports":  Call SetupForReports
    
    Case "SpecialButtons"
        Call SetupForSpecialButtons
    
End Select
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS77:" & Error$
End Sub

'PCN4335
Sub ExecuteControlsFixedMouseDown(Index As Integer)
On Error GoTo Err_Handler
Dim ErrorStr As String

Select Case Me.ControlsFixed(Index).Tag
    Case "StepForward"
        TimerTicks = 0
        Me.GeneralTimer.Tag = "StepFoward"
        Me.GeneralTimer.Enabled = True
    Case "StepRewind"
        TimerTicks = 0
        Me.GeneralTimer.Tag = "StepRewind"
        Me.GeneralTimer.Enabled = True
End Select
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS78:" & Error$
End Sub

'PCN4335
Sub ExecuteControlsFixedMouseUp(Index As Integer)
On Error GoTo Err_Handler
Dim ErrorStr As String

Me.GeneralTimer.Enabled = False
Me.GeneralTimer.Tag = ""
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS79:" & Error$
End Sub


Sub ExecuteControlsDrawButton(Index As Integer)
On Error GoTo Err_Handler


If CLPScreenAction = "DrawMaskBox" Then CLPScreenAction = "": Call ClearLineScreen.TakeASnapShot

Select Case Me.ControlsDraw(Index).Tag
    Case "Move":          Call ScreenDrawing.SetupDrawingRoutine("MoveAll")
    Case "Dimension":     Call ScreenDrawing.SetupDrawingRoutine("Dimension")
    Case "LineDraw":      Call ScreenDrawing.SetupDrawingRoutine("DrawSingleLine")
    Case "MultiLine"
        PolyLineDistance = 0 'PCN4171
        Call ScreenDrawing.SetupDrawingRoutine("DrawPolyLine")
    Case "Perpendicular": Call ScreenDrawing.SetupDrawingRoutine("DrawOrthogonalLine")
    Case "Integrated":    Call ScreenDrawing.SetupDrawingRoutine("DrawIntegratedArea")
    Case "Erase":         Call ScreenDrawing.SetupDrawingRoutine("Delete")
    Case "DrawText":      Call ScreenDrawing.SetupDrawingRoutine("DrawText")
        
End Select


If CLPScreenMode = Video Or CLPScreenMode = ThreeD Then Call ClearLineScreen.TakeASnapShot
Call ControlsScreen.ControlsViewSetup
    
If mediatype = "StillImage" Then
    'PCN4406 ''''''''''''''''''''''''''''''''''''
    Call ClearLineScreen.SetDimenResultsSize(True)              '
    ClearLineScreen.DimenResults.ZOrder 0       '
    ClearLineScreen.AreaResults.ZOrder 0        '
    '''''''''''''''''''''''''''''''''''''''''''''
End If
Exit Sub

Err_Handler:
    MsgBox Err & "-CS80:" & Error$
End Sub

Sub ExecuteReportsButton(Index As Integer)
On Error GoTo Err_Handler
Dim ErrorStr As String

'PCN4552
If PVDFileName = "" And Me.ControlsReports(Index).Tag <> "Profile" Or PVRecording Then
    ErrorStr = "Must have a pvd file loaded"
    'MsgBox DisplayMessage(ErrorStr)
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage(ErrorStr): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Exit Sub
End If

Call ScreenDrawing.CalculateExceedingLimitsAndPercentile

Select Case Me.ControlsReports(Index).Tag
    Case "Single":          Load PVReportSingle
    Case "MultiLine":       Load PVReport4in1
    Case "Profile":
        Call ClearLineScreen.SnapShotForEmbededFile
        RenderScale = 1
        Load PVReportProfile
    Case "MultiProfile"
        'This report is a multi profile report and is based on the
        'Profiles at all of the Observations points.
        Load PVReportMultiProfilex3
    Case "StoredReports":   Call SetupForStoredReports
    Case "1K":              Load Confirm1KDialog: Confirm1KDialog.Visible = True: Confirm1KDialog.ZOrder 0
End Select

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS81:" & Error$
End Sub

Sub ExecuteStoredReportsButton(Index As Integer)
On Error GoTo Err_Handler
Dim ReportIndex As Integer


Dim I As Integer

For I = 1 To PrecisionVisionGraph.ReportsPictureStorage.Count - 1
    Unload PrecisionVisionGraph.ReportsPictureStorage(I)
Next I



ReportIndex = Index + 1
Unload PVReportStoredInPVD


Load PVReportStoredInPVD
Call PVReportStoredInPVD.SetupForStoredReport(ReportIndex)

PVReportStoredInPVD.Show

Exit Sub
Err_Handler:
    MsgBox Err & "-CS82:" & Error$
End Sub


Sub Execute3DButton(Index As Integer)
On Error GoTo Err_Handler

Select Case Me.Controls3D(Index).Tag
    Case "ViewTanAndTilt"
        Call d3d_camselect(1)

    Case "ViewOutSide"
        Call d3d_camselect(2)

    Case "ViewFlyBy"
        Call d3d_camselect(3)

    Case "ViewMap"
        Call d3d_camselect(4)

    Case "Scale"
        If PipeScale3D = 32 Then
            Call d3d_pipe_scale(2)
            PipeScale3D = 2
        ElseIf PipeScale3D = 2 Or PipeScale3D = 0 Then
            Call d3d_pipe_scale(8)
            PipeScale3D = 8
        Else
            Call d3d_pipe_scale(32)
            PipeScale3D = 32
        End If

''    Case "ZoomIn"
''        If zoomspeed <> -1 Then
''            zoomspeed = -1
''        Else
''            zoomspeed = 0
''        End If
''
''    Case "ZoomOut"
''        If zoomspeed <> 1 Then
''            zoomspeed = 1
''        Else
''            zoomspeed = 0
''        End If
''
''    Case "RotateXAntiClock"
''        Call d3d_rotate_pipe_z(5)
''
''    Case "RotateXClockwise"
''        Call d3d_rotate_pipe_z(-5)
''
''    Case "RotateYAntiClock"
''        Call d3d_rotate_pipe_y(5)
''
''    Case "RotateYClockwise"
''        Call d3d_rotate_pipe_y(-5)


End Select

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS83:" & Error$
End Sub

Sub Execute3DButtonMouseDown(Index As Integer)
On Error GoTo Err_Handler

Select Case Me.Controls3D(Index).Tag
    Case "ZoomIn"
        zoomspeed = -1
        ClearLineScreen.ButtonNumber = 10
        
    Case "ZoomOut"
        zoomspeed = 1
        ClearLineScreen.ButtonNumber = 11

    Case "RotateXAntiClock"
        ClearLineScreen.ButtonNumber = 12

    Case "RotateXClockwise"
        ClearLineScreen.ButtonNumber = 13

    Case "RotateYAntiClock"
        ClearLineScreen.ButtonNumber = 14

    Case "RotateYClockwise"
        ClearLineScreen.ButtonNumber = 15

End Select

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS84:" & Error$
End Sub

Sub Execute3DButtonMouseUp(Index As Integer)
On Error GoTo Err_Handler

zoomspeed = 0
ClearLineScreen.ButtonNumber = 0
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS85:" & Error$
End Sub

Sub ExecuteCalibrationButton(Index As Integer)
On Error GoTo Err_Handler

Select Case Me.ControlsCalibration(Index).Tag
    Case "CalibrationH": Call SetCalibration
    Case "CalibrationV"
        If CalibrationMethodActioned = "CalibrationH" Or CalibrationMethodActioned = "CalibrationV" Then
            Call SetVerticalCalibration
        End If
    Case "CalibrationCrack"

End Select

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS86:" & Error$
End Sub

Sub ExecuteSpecialButton(Index As Integer)
On Error GoTo Err_Handler

Select Case Me.ControlsSpecial(Index).Tag
    Case "Calibration"
        Call SetCalibration
        Call SetupForCalibration

    Case "SetVideoMask"
        Call SetVideoMask
        
    Case "SetWaterLevel"
        Call SetWaterLevel
    
    Case "RefShapes"
        Call SetupForRefShapeButtons
        
End Select

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS87:" & Error$
End Sub

Sub ExecuteRefShapeButton(Index As Integer)
On Error GoTo Err_Handler

Select Case Me.ControlsRefShape(Index).Tag
    Case "ExternalDiameter"
        Call DrawInternalDiameter("DrawOutsideDia") 'PCNGL310103
    
    Case "RefShapeCircle"
        DrawShapeType = "Circle"

    Case "RefShapeEgg"
        DrawShapeType = "Egg"

    Case "RefShapeHorseShoe"
        DrawShapeType = "HorseShoe"

End Select

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CS88:" & Error$
End Sub

Sub SetPVSettings()
On Error GoTo Err_Handler
Dim ErrorStr As String

Call ControlsMain.ControlsDisplaySetup("PVSettings")
Call AutoTune.SetupSelectedTask("Tuning")
Call ClearLineScreen.OnscreenTextHide
Call ClearLineScreen.VideoScreenActivate(ErrorStr) 'PCN4171
If ErrorStr = "" Then
    Call ControlsViewSetup
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-CS89:" & Error$
End Sub


Sub SetVideoMask()
On Error GoTo Err_Handler
'PCN4356 Dim bRememberFisheyeState As Boolean

Call ControlsMain.ControlsDisplaySetup("PVSettings")
Call AutoTune.SetupSelectedTask("Mask")
'PCN4356 bRememberFisheyeState = ConfigInfo.FishEyeFlag
'PCN4356 FisheyeFunctions.FEOFF

IgnoreAreaFlag = True
DelFlag = False
'Take a snap shot
Call ScreenDrawing.SetupDrawingRoutine("DrawMaskBox")
Call ClearLineScreen.TakeASnapShot

'PCN4293 when making a maskes, need to see all of video
ClearLineScreen.DimenResults.Visible = False
ClearLineScreen.AreaResults.Visible = False

DoEvents


'PCN4356 If bRememberFisheyeState Then FisheyeFunctions.FEON

Call AutoTune.SetupSelectedTask("Mask")

Exit Sub
Err_Handler:
    MsgBox Err & "-CS90:" & Error$
End Sub

Sub SetWaterLevel()
On Error GoTo Err_Handler

If PVRecording Then Exit Sub 'PCN4216
Call ControlsMain.ControlsDisplaySetup("PVSettings")
Call AutoTune.SetupSelectedTask("WaterLevel")

Call ScreenDrawing.SetupDrawingRoutine("DrawWaterLevel")

If expecteddiameter = 0 Then
    'MsgBox DisplayMessage("Please specify internal pipe diameter first."), vbExclamation
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Please specify internal pipe diameter first."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    PipelineDetails.ZOrder 0
    PipelineDetails.InternalDiameterExpected.SetFocus
    Exit Sub
End If

ScreenDrawing.ShowReferenceShape = True
Call ClearLineScreen.TakeASnapShot
'Call MsgBox(DisplayMessage("Please ensure the main reference shape is centered."), vbOKCancel + vbInformation)  'PCN2111
ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Please ensure the main reference shape is centered."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0


Exit Sub
Err_Handler:
    MsgBox Err & "-CS91:" & Error$
End Sub



Sub StartStopRecording(ErrorStr As String)
On Error GoTo Err_Handler


If PVRecording Then
    Call ClearLineScreen.RecordPVProfile_Stop
'    Call SetupForPostRecording
Else
    Call ClearLineScreen.RecordPVProfile_Start(ErrorStr)
    If ErrorStr <> "" Then Exit Sub
End If
Call ControlsViewSetup

Exit Sub
Err_Handler:
    MsgBox Err & "-CS92:" & Error$
End Sub

Sub SetupCLPScreenToPV(ErrorStr As String)
On Error GoTo Err_Handler
        
If PVDFileName = "" Then 'PCN2783
    ErrorStr = "Must have a pvd file loaded"
    'MsgBox DisplayMessage(ErrorStr)
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage(ErrorStr): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Exit Sub
End If
Call ClearLineScreen.ProfilerPause


DrawingCentreX = CentreLineX
DrawingCentreY = CentreLineY

' Turn off Picture in Picture
ClearLineScreen.PVScreenPicInPic.Visible = False
PicInPicMode = "OFF"
'Setup for PV mode
If CLPScreenMode <> PV Then 'PCNGL140203
    CLPScreenMode = PV 'PCN1863 'PCNGL140203
    ClearLineTitle.TitleBarCaption.Caption = DisplayMessage("Precision Vision") & " - " & PVDFileName 'PCN2759
    ClearLineTitle.TitleBarCaption.ToolTipText = DisplayMessage("Precision Vision") & " - " & PVDFileName 'PCN4171
    Call ClearLineScreen.SetupMainScreenForPV 'PCN1863
End If
''Call ClearLineScreen.SetupMTButtonsForPV  'PCNGL300103
Call ScreenDrawing.DrawProfilesStartToFinish(ClearLineScreen.PVScreen, True)
Call PVGraphsKeyForm.DisplayPVGraphsKey 'PCN2990 'PCN2818

Exit Sub
Err_Handler:
    MsgBox Err & "-CS93:" & Error$
End Sub

Sub SetupCLPScreenToSnapShot(ErrorStr As String)
On Error GoTo Err_Handler
        
If mediatype = "" Then 'PCN2783
    ErrorStr = "Must load a Media file first!"
    'MsgBox DisplayMessage(ErrorStr)
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage(ErrorStr): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Exit Sub
End If
Call ClearLineScreen.ProfilerPause

ClearLineScreen.PVScreen.Visible = False

'vvvv *****  'PCNGL140203
If CLPScreenMode <> SnapShot Then 'PCNGL140203
    'Relocate shapes to video centre line
    'ClearLineScreen.SnapShotScreen.Cls 'PCNLS1999 'PCN3219
    
End If
'^^^^ **************************************************
CLPScreenMode = SnapShot 'PCN4043
Call ControlsViewSetup 'PCN4171

ClearLineTitle.TitleBarCaption.Caption = DisplayMessage(SnapShot) & " - " & PVDFileName 'PCN2759
ClearLineTitle.TitleBarCaption.ToolTipText = DisplayMessage(SnapShot) & " - " & PVDFileName 'PCN4171
''Call ClearLineScreen.SetupMTButtonsForSnapShot 'PCNGL300103
ClearLineScreen.PVScreen.MousePointer = 99
ClearLineScreen.PVScreen.MouseIcon = LoadResPicture(MainScreenMouseIcon, vbResIcon) 'PCNGL291102
'Configure ClearLineScreen for Snap Shot mode
If ClearLineScreen.SnapShotScreen.Picture = 0 Then
    Call ClearLineScreen.Snap
    'PCN3576 'The folling moved from the function snap
    ClearLineScreen.LoadImage (LocToSave & "Snapshot.bmp") '("E:\Documents and Settings\LouiseS\Desktop\6_1_03ClearLineProfilerV4\Images\Test171202.bmp ")
    'vvvv Setup Screen for the SnapShot *************************************** 'PCNGL190103
    CLPScreenMode = SnapShot
    
    ClearLineScreen.PVScreen.MousePointer = 99
    ClearLineScreen.PVScreen.MouseIcon = LoadResPicture(MainScreenMouseIcon, vbResIcon)
    'Configure ClearLineScreen for Snap Shot mode
    ClearLineScreen.SnapShotScreen.Visible = True
    ClearLineScreen.SnapShotScreen.ZOrder 0 'PCNGL261202
    '^^^^ ********************************************************************
End If

ClearLineScreen.SnapShotScreen.Visible = True
ClearLineScreen.SnapShotScreen.ZOrder 0 'PCNGL261202
ClearLineScreen.SnapShotScreen.AutoRedraw = True 'PCNGL2901032

ClearLineScreen.SnapShotScreen.AutoRedraw = False 'PCNGL140203
Call ScreenDrawing.DrawProfilesStartToFinish(ClearLineScreen.SnapShotScreen, True)
'Call DrawAll(ClearLineScreen.SnapShotScreen) 'PCNGL2901032
' Turn off Picture in Picture
ClearLineScreen.PVScreenPicInPic.Visible = False
PicInPicMode = "OFF"
'Call ClearLineScreen.OnscreenTextShow 'PCN3569 'PCN4171
Call ClearLineScreen.SetDimenResultsSize(True)  'PCN4171
ClearLineScreen.DimenResults.ZOrder 0
ClearLineScreen.AreaResults.ZOrder 0


Exit Sub
Err_Handler:
    MsgBox Err & "-CS94:" & Error$
End Sub

Sub SetupCLPScreenTo3D(ErrorStr As String)
On Error GoTo Err_Handler
    Dim RegNow As Variant
        
If PVDFileName = "" Then
    ErrorStr = "Must have a pvd file loaded"
    'MsgBox DisplayMessage(ErrorStr)
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage(ErrorStr): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Exit Sub
End If
       
Call ClearLineScreen.ProfilerPause

Call ClearLineScreen.OnscreenTextHide 'PCN3569
ClearLineTitle.TitleBarCaption.Caption = DisplayMessage("ClearLine 3D") & " - " & PVDFileName 'PCN2759
ClearLineTitle.TitleBarCaption.ToolTipText = DisplayMessage("ClearLine 3D") & " - " & PVDFileName 'PCN4171

If ThreeDRunning = False Then
    If PVDFileName = "" Then 'PCN2374
        'MsgBox DisplayMessage("Must have a pvd file loaded")
        ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Must have a pvd file loaded"): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
        Exit Sub
    End If
    CLPScreenMode = ThreeD 'PCN2240
    
    PipeScale3D = 2
    laserspeed = 0
    ClearLineScreen.MainScreen3D.Picture = LoadPicture("")
    ClearLineScreen.MainScreen3D.Left = 0
    ClearLineScreen.MainScreen3D.Top = 0
    ClearLineScreen.MainScreen3D.width = ClearLineScreen.MainScreen.width
    ClearLineScreen.MainScreen3D.height = ClearLineScreen.MainScreen.height
    ClearLineScreen.MainScreen3D.BorderStyle = 0 'PCN2475
    ClearLineScreen.MainScreen3D.Visible = True
    ClearLineScreen.MainScreen3D.ZOrder 0
    DoEvents
    Call ClearLineScreen.Load3D 'GLTesting **********************************************************
    
Else
    If PVDFileName = "" Then 'PCN2374
        'MsgBox DisplayMessage("Must have a pvd file loaded")
        ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Must have a pvd file loaded"): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
        Exit Sub
    End If
    CLPScreenMode = ThreeD 'PCN2240
    ClearLineScreen.ThreeDtimer.Enabled = True
    ClearLineScreen.MainScreen3D.Visible = True 'PCN2240
    ClearLineScreen.MainScreen3D.ZOrder 0
End If

PVGraphsKeyForm.Visible = True 'PCN2990 'PCN2818


Exit Sub
Err_Handler:
    MsgBox Err & "-CS95:" & Error$
End Sub

Sub SaveThePVDRecordingToFile()
On Error GoTo Err_Handler

If PVRecording = True Then 'PCN2379
    'MsgBox DisplayMessage("Stop PVD recording before saving") 'PCN2762
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Stop PVD recording before saving"): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Exit Sub
End If
If Registered = False Then 'PCNML220103 'Testing ML040203
    'MsgBox DisplayMessage("Cannot save a .PVD file, please register the software to access this."), vbExclamation 'PCN2111
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Cannot save a .PVD file, please register the software to access this."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
Else
    Call SaveImageAndOrData 'PCNGL110103
    Call ControlsViewSetup
    Call ControlsMain.SetDisplayMainsFns   'PCN4298
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-CS96:" & Error$
End Sub

Sub SetCalibration()
On Error GoTo Err_Handler

Call ScreenDrawing.SetupDrawingRoutine("DrawCalibrationLine")
CalibrationTypeLength = 0
Call OptionsPage.FishEyeCameraDropdown_OnSelect
Call ClearLineScreen.TakeASnapShot

Call SetDisplayZoomOnSnap("On")

'If CLPScreenMode = Video Then
 '   CalibrationMethodActioned = ""
 '   Call SetVerticalCalibration
'End If


Exit Sub
Err_Handler:
    MsgBox Err & "-CS97:" & Error$
End Sub

Sub SetVerticalCalibration()
On Error GoTo Err_Handler

If CalibrationMethodActioned <> "CalibrationH" And CalibrationMethodActioned <> "CalibrationV" Then Exit Sub
Call ScreenDrawing.SetupDrawingRoutine("DrawHorCalibrationLine")
If CLPScreenMode = Video Then
    Call ClearLineScreen.TakeASnapShot
End If

CalibrationTypeLength = 0

Call SetDisplayZoomOnSnap("On")

Exit Sub
Err_Handler:
    MsgBox Err & "-CS98:" & Error$
End Sub

Sub SwitchToVideoMode()
On Error GoTo Err_Handler

Call ControlsView_Click(3)  'Index three is the video button

Exit Sub
Err_Handler:
    MsgBox Err & "-CS99:" & Error$
End Sub

Sub SetupForReaderConfiguration() 'PCN3809
On Error GoTo Err_Handler

Call SetupForStoredReports

'Hide all unused controls
'Hide fixed controls except the video controls
Me.ControlsFixed(1).Visible = False
For ButtonIndex = 2 To 4 'PCN4299
    Me.ControlsFixed(ButtonIndex).Visible = True
Next ButtonIndex
ButtonIndex = 4
For ButtonIndex = 5 To 18 'PCN4299
    Me.ControlsFixed(ButtonIndex).Visible = False
Next ButtonIndex
'Hide SnapShot button
Me.ControlsView(1).Visible = False

'vvvv PCN4299 *********************************************
'Show the Controls view buttons available for the reader
Me.ControlsView(0).Visible = True
Me.ControlsView(2).Visible = True
Me.ControlsView(3).Visible = True
'^^^^ *****************************************************

Me.ControlsDraw(2).Visible = True 'PCN4395

'Setup Controls view
Select Case CLPScreenMode
    Case ThreeD
        Set Me.ControlsView(0).Picture = Me.ControlsViewImageList.ListImages("3D").Picture
    Case PV
        Set Me.ControlsView(2).Picture = Me.ControlsViewImageList.ListImages("PV").Picture
    Case Video
        Set Me.ControlsView(3).Picture = Me.ControlsViewImageList.ListImages("Video").Picture
End Select

'vvvv PCN4321 *************************
ScreenDrawing.ShowReferenceShape = True
'^^^^ *********************************

Exit Sub
Err_Handler:
    MsgBox Err & "-CS100:" & Error$
End Sub


'PCN4335 holding in frame rewind or foward then will play in slowmation
Private Sub GeneralTimer_Timer()
On Error GoTo Err_Handler

TimerTicks = TimerTicks + 1
If TimerTicks < 5 Then Exit Sub

Select Case Me.GeneralTimer.Tag
    Case "StepFoward": Call ClearLineScreen.ProfilerStepFoward
    Case "StepRewind": Call ClearLineScreen.ProfilerStepRewind
End Select



Exit Sub
Err_Handler:
    MsgBox Err & "-CS101:" & Error$
End Sub

Sub SetupControlsView() 'PCN4320
On Error GoTo Err_Handler


If PVDFileName = "" And mediatype = "" Then
    'The Profiler is in startup mode without a video or PVD loaded.
    Exit Sub
ElseIf PVDFileName <> "" And PVRecording = False Then
    'If pvd is loaded then show 3D and precision vision
    Me.ControlsView(0).Visible = True: Me.ControlsView(2).Visible = True
End If

If mediatype = Video Or mediatype = "Live" Then
    Me.ControlsView(3).Visible = True 'If video or live present show video icon
End If

If mediatype = "Live" Then
    Call Me.SetupForLive
End If

If mediatype = "StillImage" Then
    Call Me.SetupForLoadedImage
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-CS102:" & Error$
End Sub

