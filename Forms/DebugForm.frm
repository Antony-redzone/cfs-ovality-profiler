VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form DebugForm 
   AutoRedraw      =   -1  'True
   Caption         =   "Debug Form"
   ClientHeight    =   14985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19230
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   999
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1282
   Begin VB.CommandButton CleanProfile 
      Caption         =   "CleanProfile"
      Height          =   495
      Left            =   480
      TabIndex        =   127
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton SmoothButton 
      Caption         =   "Smooth Graph"
      Height          =   375
      Left            =   480
      TabIndex        =   126
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox MetersPerGraphValue 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   124
      Text            =   "50"
      Top             =   4680
      Width           =   615
   End
   Begin VB.CheckBox UseLimitTick 
      Caption         =   "Use Limit"
      Height          =   255
      Left            =   2160
      TabIndex        =   123
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox LimitValue 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3360
      TabIndex        =   122
      Text            =   "0"
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton TwoInOneReport 
      Caption         =   "2 in 1 Report"
      Height          =   375
      Left            =   480
      TabIndex        =   121
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Profiling center calculation adjustment"
      Height          =   1575
      Left            =   360
      TabIndex        =   116
      Top             =   240
      Width           =   3855
      Begin VB.CommandButton UnlockDonutBtn 
         Caption         =   "Unlock Donut"
         Height          =   255
         Left            =   120
         TabIndex        =   118
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton LockDonutBtn 
         Caption         =   "Lock donut"
         Height          =   255
         Left            =   120
         TabIndex        =   117
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Post Centre caclulation adjustment"
      Height          =   1695
      Left            =   480
      TabIndex        =   111
      Top             =   1920
      Width           =   4575
      Begin VB.CommandButton CutWaterLevelButton 
         Height          =   195
         Left            =   240
         TabIndex        =   120
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton ReduceWaterLevelBtn 
         Height          =   195
         Left            =   240
         TabIndex        =   119
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox CutWaterLevelCheck 
         Caption         =   "Simple circular fill in, not using mirror top"
         Enabled         =   0   'False
         Height          =   195
         Left            =   600
         TabIndex        =   115
         Top             =   480
         Width           =   3375
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Left            =   1560
         TabIndex        =   113
         Top             =   1320
         Width           =   615
      End
      Begin VB.CheckBox ReduceWaterLevelCheck 
         Caption         =   "Reduce waterlevel effect from centre calc"
         Enabled         =   0   'False
         Height          =   255
         Left            =   600
         TabIndex        =   112
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label17 
         Caption         =   "Recalc waterlevel"
         Height          =   255
         Left            =   120
         TabIndex        =   114
         Top             =   1320
         Width           =   1455
      End
   End
   Begin VB.CheckBox ShiftCheckBox 
      Caption         =   "Shift On"
      Height          =   255
      Left            =   360
      TabIndex        =   110
      Top             =   9480
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CommandButton CalculateTotal 
      Caption         =   "Caclulate Total"
      Height          =   375
      Left            =   5520
      TabIndex        =   109
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox DebrisEndFrame 
      Height          =   285
      Left            =   6960
      TabIndex        =   108
      Text            =   "0"
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox DebrisStartFrame 
      Height          =   285
      Left            =   6960
      TabIndex        =   107
      Text            =   "0"
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox TotalDebris 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   106
      Top             =   840
      Width           =   1455
   End
   Begin VB.CheckBox DebrisOnCheck 
      Caption         =   "Debris On"
      Height          =   255
      Left            =   7080
      TabIndex        =   105
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton DebrisCalc 
      Caption         =   "Calculate Debris"
      Height          =   375
      Left            =   5520
      TabIndex        =   104
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton TurnOnCentreCalc 
      Caption         =   "Centre Calc On"
      Height          =   375
      Left            =   5520
      TabIndex        =   103
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton TurnOffCentreCalc 
      Caption         =   "Centre Calc Off"
      Height          =   375
      Left            =   5520
      TabIndex        =   102
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton PVReportButton 
      Caption         =   "1K Report"
      Height          =   375
      Left            =   480
      TabIndex        =   101
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox LabelResolutionTextBox 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3480
      TabIndex        =   100
      Text            =   "500"
      Top             =   8520
      Width           =   735
   End
   Begin VB.CommandButton DrawTrackingLabelResolution 
      Caption         =   "Label Resolution"
      Height          =   255
      Left            =   1920
      TabIndex        =   99
      Top             =   8520
      Width           =   1455
   End
   Begin VB.TextBox TrackingResolutionTextBox 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3480
      TabIndex        =   98
      Text            =   "1"
      Top             =   8160
      Width           =   735
   End
   Begin VB.CommandButton DrawTrackingResolution 
      Caption         =   "Draw Resolution"
      Height          =   255
      Left            =   1920
      TabIndex        =   97
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CheckBox EgnoreCentreHeight 
      Caption         =   "Height = 0 (Used for float)"
      Height          =   375
      Left            =   5520
      TabIndex        =   96
      Top             =   1920
      Width           =   2175
   End
   Begin VB.OptionButton CentreLevelOption 
      Caption         =   "Outer Edge"
      Height          =   375
      Index           =   1
      Left            =   5520
      TabIndex        =   95
      Top             =   1560
      Width           =   1215
   End
   Begin VB.OptionButton CentreLevelOption 
      Caption         =   "Standard"
      Height          =   375
      Index           =   0
      Left            =   5520
      TabIndex        =   94
      Top             =   1200
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox JustATextBox3 
      Height          =   375
      Left            =   120
      TabIndex        =   93
      Text            =   "Text1"
      Top             =   12120
      Width           =   2655
   End
   Begin VB.TextBox JustATextBox2 
      Height          =   375
      Left            =   120
      TabIndex        =   92
      Text            =   "Text1"
      Top             =   11640
      Width           =   2655
   End
   Begin VB.TextBox JustATextBox 
      Height          =   285
      Left            =   120
      TabIndex        =   91
      Text            =   "Text1"
      Top             =   11280
      Width           =   2655
   End
   Begin VB.CommandButton DrawTrackingAmountBtn 
      Caption         =   "Draw Amount"
      Height          =   375
      Left            =   1920
      TabIndex        =   90
      Top             =   7560
      Width           =   1095
   End
   Begin VB.TextBox EndLine 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3120
      TabIndex        =   89
      Text            =   "0"
      Top             =   7800
      Width           =   1095
   End
   Begin VB.TextBox StartLine 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3120
      TabIndex        =   88
      Text            =   "0"
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton TrackingLoadBtn 
      Caption         =   "Load Tracking"
      Height          =   375
      Left            =   2520
      TabIndex        =   87
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton BtnSetLaserTracking 
      Caption         =   "Set Laser Tracking"
      Height          =   495
      Left            =   3120
      TabIndex        =   86
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton SaveDistanceInfo 
      Caption         =   "Save distance information"
      Height          =   495
      Left            =   240
      TabIndex        =   85
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton FilterBtn 
      Caption         =   "FilterButton"
      Height          =   375
      Left            =   480
      TabIndex        =   84
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton CancelBtn 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   7680
      TabIndex        =   83
      Top             =   13680
      Width           =   1335
   End
   Begin VB.TextBox TrueDiameterOffsetValue 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8760
      TabIndex        =   73
      Text            =   "0"
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox CapacityOffsetValue 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8760
      TabIndex        =   71
      Text            =   "0"
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox RotateTo 
      Height          =   285
      Left            =   6840
      TabIndex        =   68
      Text            =   "1"
      Top             =   13200
      Width           =   735
   End
   Begin VB.TextBox RotateFrom 
      Height          =   285
      Left            =   4560
      TabIndex        =   67
      Text            =   "1"
      Top             =   13200
      Width           =   735
   End
   Begin VB.TextBox RotationPVDValue 
      Height          =   285
      Left            =   5280
      TabIndex        =   66
      Text            =   "0"
      Top             =   13680
      Width           =   855
   End
   Begin VB.CommandButton RotateRight 
      Caption         =   "Rotate Right"
      Height          =   255
      Left            =   4080
      TabIndex        =   65
      Top             =   13800
      Width           =   1215
   End
   Begin VB.CommandButton RotateLeft 
      Caption         =   "Rotate Left"
      Height          =   255
      Left            =   4080
      TabIndex        =   64
      Top             =   13560
      Width           =   1215
   End
   Begin VB.CommandButton SaveNewPVDData 
      Caption         =   "Save Rotated Data"
      Height          =   255
      Left            =   4080
      TabIndex        =   63
      Top             =   14040
      Width           =   3495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Sonar"
      Height          =   855
      Left            =   240
      TabIndex        =   62
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton AutoRotate 
      Caption         =   "AutoRotate"
      Height          =   255
      Left            =   6360
      TabIndex        =   61
      Top             =   13680
      Width           =   1215
   End
   Begin VB.CommandButton StoreCustomShape 
      Caption         =   "Store Custom Shape"
      Height          =   495
      Left            =   7800
      TabIndex        =   60
      Top             =   15720
      Width           =   1095
   End
   Begin VB.CommandButton LoadCustomShape 
      Caption         =   "Load Custom Shape"
      Height          =   495
      Left            =   4080
      TabIndex        =   59
      Top             =   15720
      Width           =   1095
   End
   Begin VB.TextBox EndAngle 
      Height          =   285
      Index           =   3
      Left            =   7920
      TabIndex        =   54
      Text            =   "0"
      Top             =   15360
      Width           =   975
   End
   Begin VB.TextBox StartAngle 
      Height          =   285
      Index           =   3
      Left            =   6960
      TabIndex        =   53
      Text            =   "0"
      Top             =   15360
      Width           =   975
   End
   Begin VB.TextBox Radius 
      Height          =   285
      Index           =   3
      Left            =   6000
      TabIndex        =   52
      Text            =   "0"
      Top             =   15360
      Width           =   975
   End
   Begin VB.TextBox ShapeOriginY 
      Height          =   285
      Index           =   3
      Left            =   5040
      TabIndex        =   51
      Text            =   "0"
      Top             =   15360
      Width           =   975
   End
   Begin VB.TextBox ShapeOriginX 
      Height          =   285
      Index           =   3
      Left            =   4080
      TabIndex        =   50
      Text            =   "0"
      Top             =   15360
      Width           =   975
   End
   Begin VB.TextBox EndAngle 
      Height          =   285
      Index           =   2
      Left            =   7920
      TabIndex        =   49
      Text            =   "0"
      Top             =   15120
      Width           =   975
   End
   Begin VB.TextBox StartAngle 
      Height          =   285
      Index           =   2
      Left            =   6960
      TabIndex        =   48
      Text            =   "0"
      Top             =   15120
      Width           =   975
   End
   Begin VB.TextBox Radius 
      Height          =   285
      Index           =   2
      Left            =   6000
      TabIndex        =   47
      Text            =   "0"
      Top             =   15120
      Width           =   975
   End
   Begin VB.TextBox ShapeOriginY 
      Height          =   285
      Index           =   2
      Left            =   5040
      TabIndex        =   46
      Text            =   "0"
      Top             =   15120
      Width           =   975
   End
   Begin VB.TextBox ShapeOriginX 
      Height          =   285
      Index           =   2
      Left            =   4080
      TabIndex        =   45
      Text            =   "0"
      Top             =   15120
      Width           =   975
   End
   Begin VB.TextBox EndAngle 
      Height          =   285
      Index           =   1
      Left            =   7920
      TabIndex        =   44
      Text            =   "0"
      Top             =   14880
      Width           =   975
   End
   Begin VB.TextBox StartAngle 
      Height          =   285
      Index           =   1
      Left            =   6960
      TabIndex        =   43
      Text            =   "0"
      Top             =   14880
      Width           =   975
   End
   Begin VB.TextBox Radius 
      Height          =   285
      Index           =   1
      Left            =   6000
      TabIndex        =   42
      Text            =   "0"
      Top             =   14880
      Width           =   975
   End
   Begin VB.TextBox ShapeOriginY 
      Height          =   285
      Index           =   1
      Left            =   5040
      TabIndex        =   41
      Text            =   "0"
      Top             =   14880
      Width           =   975
   End
   Begin VB.TextBox ShapeOriginX 
      Height          =   285
      Index           =   1
      Left            =   4080
      TabIndex        =   40
      Text            =   "0"
      Top             =   14880
      Width           =   975
   End
   Begin VB.TextBox EndAngle 
      Height          =   285
      Index           =   0
      Left            =   7920
      TabIndex        =   39
      Text            =   "0"
      Top             =   14640
      Width           =   975
   End
   Begin VB.TextBox StartAngle 
      Height          =   285
      Index           =   0
      Left            =   6960
      TabIndex        =   38
      Text            =   "0"
      Top             =   14640
      Width           =   975
   End
   Begin VB.TextBox Radius 
      Height          =   285
      Index           =   0
      Left            =   6000
      TabIndex        =   37
      Text            =   "0"
      Top             =   14640
      Width           =   975
   End
   Begin VB.TextBox ShapeOriginY 
      Height          =   285
      Index           =   0
      Left            =   5040
      TabIndex        =   36
      Text            =   "0"
      Top             =   14640
      Width           =   975
   End
   Begin VB.TextBox ShapeOriginX 
      Height          =   285
      Index           =   0
      Left            =   4080
      TabIndex        =   35
      Text            =   "0"
      Top             =   14640
      Width           =   975
   End
   Begin MSComctlLib.Slider SliderRed 
      Height          =   1455
      Left            =   5520
      TabIndex        =   32
      Top             =   5520
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   2566
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   1
      Max             =   9
   End
   Begin VB.CommandButton DumpBtn 
      Caption         =   "Dump XY"
      Height          =   495
      Left            =   7200
      TabIndex        =   31
      Top             =   7440
      Width           =   1575
   End
   Begin VB.TextBox LimitLineDrawnTextBox 
      Height          =   285
      Left            =   6600
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   6480
      Width           =   2175
   End
   Begin VB.TextBox LimitLineTextBox 
      Height          =   285
      Left            =   6600
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton BtnProcessDummyImage 
      Caption         =   "Process Image"
      Height          =   1335
      Left            =   9000
      Picture         =   "DebugForm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox ScaleText 
      Height          =   285
      Left            =   6720
      TabIndex        =   19
      Text            =   "1"
      Top             =   5040
      Width           =   615
   End
   Begin MSComctlLib.Slider GraphScaleSlider 
      Height          =   255
      Left            =   8040
      TabIndex        =   18
      Top             =   5400
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393216
      Min             =   1
      Max             =   30
      SelStart        =   5
      Value           =   5
   End
   Begin VB.TextBox Text5Debug 
      Height          =   285
      Left            =   7200
      TabIndex        =   8
      Text            =   "0"
      Top             =   11760
      Width           =   855
   End
   Begin VB.CommandButton SaveGrid 
      Caption         =   "Save Grid"
      Height          =   495
      Left            =   10680
      TabIndex        =   6
      Top             =   8160
      Width           =   1695
   End
   Begin VB.CommandButton Calibrate 
      Caption         =   "Callibrate"
      Height          =   495
      Left            =   9240
      TabIndex        =   5
      Top             =   8160
      Width           =   1335
   End
   Begin VB.CommandButton StartGrid 
      Caption         =   "Start Entering grid"
      Height          =   495
      Left            =   7200
      TabIndex        =   4
      Top             =   8160
      Width           =   1935
   End
   Begin VB.TextBox Text2Debug 
      Height          =   285
      Left            =   7200
      TabIndex        =   3
      Text            =   "0"
      Top             =   11400
      Width           =   855
   End
   Begin MSComctlLib.Slider Slider2Debug 
      Height          =   195
      Left            =   8040
      TabIndex        =   2
      Top             =   5160
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   344
      _Version        =   393216
      Max             =   100
      TickFrequency   =   5
   End
   Begin VB.TextBox Text1Debug 
      Height          =   285
      Left            =   7200
      TabIndex        =   1
      Text            =   "0"
      Top             =   11040
      Width           =   855
   End
   Begin MSComctlLib.Slider Slider1Debug 
      Height          =   195
      Left            =   8040
      TabIndex        =   0
      Top             =   4920
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   344
      _Version        =   393216
      Max             =   256
      TickFrequency   =   5
   End
   Begin VB.Frame Frame2 
      Caption         =   "Refernce Shape Control"
      Height          =   2295
      Left            =   3960
      TabIndex        =   7
      Top             =   10800
      Width           =   5175
      Begin VB.TextBox CurrentScale 
         Height          =   285
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   82
         Text            =   "0"
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton ScaleBtn 
         Caption         =   "Apply Scale"
         Height          =   255
         Left            =   1920
         TabIndex        =   81
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox ScaleValue 
         Height          =   285
         Left            =   960
         TabIndex        =   80
         Text            =   "0"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox CurrentStretch 
         Height          =   285
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   78
         Text            =   "0"
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton StretchBtn 
         Caption         =   "Apply Stretch"
         Height          =   255
         Left            =   1920
         TabIndex        =   77
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox StretchValue 
         Height          =   285
         Left            =   960
         TabIndex        =   75
         Text            =   "0"
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Apply Rotation"
         Height          =   255
         Left            =   1800
         TabIndex        =   17
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "ApplyY Pos"
         Height          =   255
         Left            =   1800
         TabIndex        =   16
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Apply X Pos"
         Height          =   255
         Left            =   1800
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text8Debug 
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Text            =   "0"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text7Debug 
         Height          =   285
         Left            =   960
         TabIndex        =   10
         Text            =   "0"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text6Debug 
         Height          =   285
         Left            =   960
         TabIndex        =   9
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "Scale"
         Height          =   255
         Left            =   240
         TabIndex        =   79
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Stretch"
         Height          =   255
         Left            =   240
         TabIndex        =   76
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Set Rot"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Set PosY"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Set PosX"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Dump time control"
      Height          =   1575
      Left            =   7320
      TabIndex        =   20
      Top             =   8760
      Width           =   3735
      Begin VB.TextBox LastTimeDumpTextBox 
         Height          =   375
         Left            =   2640
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox TimeGonePastTextBox 
         Height          =   285
         Left            =   2640
         TabIndex        =   24
         Text            =   "Text2"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox CurrentTimeTextBox 
         Height          =   285
         Left            =   2640
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label LastTimeDumpLabel 
         Caption         =   "Last time dump"
         Height          =   375
         Left            =   1320
         TabIndex        =   25
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Time gone past since last dump"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Current time from LaserLib"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Label Label18 
      Caption         =   "Meters Per Graph"
      Height          =   255
      Left            =   480
      TabIndex        =   125
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      Height          =   2175
      Left            =   1800
      Top             =   6720
      Width           =   2775
   End
   Begin VB.Label Label14 
      Caption         =   "True Diameter Offset"
      Height          =   255
      Left            =   7200
      TabIndex        =   74
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label CapacityOffsetLabel 
      Caption         =   "Capacity Offset by %"
      Height          =   255
      Left            =   7200
      TabIndex        =   72
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label13 
      Caption         =   "To"
      Height          =   255
      Left            =   6480
      TabIndex        =   70
      Top             =   13200
      Width           =   495
   End
   Begin VB.Label Label12 
      Caption         =   "From"
      Height          =   255
      Left            =   4080
      TabIndex        =   69
      Top             =   13200
      Width           =   735
   End
   Begin VB.Shape Shape1 
      Height          =   3375
      Left            =   3960
      Top             =   13080
      Width           =   5175
   End
   Begin VB.Label Label11 
      Caption         =   "End"
      Height          =   255
      Left            =   7920
      TabIndex        =   58
      Top             =   14400
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Start"
      Height          =   255
      Left            =   6960
      TabIndex        =   57
      Top             =   14400
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "Radius"
      Height          =   255
      Left            =   6000
      TabIndex        =   56
      Top             =   14400
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Origin Y"
      Height          =   255
      Left            =   5160
      TabIndex        =   55
      Top             =   14400
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Origin X"
      Height          =   255
      Left            =   4080
      TabIndex        =   34
      Top             =   14400
      Width           =   615
   End
   Begin VB.Shape ImageBlurSize 
      FillColor       =   &H0000C0C0&
      FillStyle       =   0  'Solid
      Height          =   150
      Left            =   5880
      Top             =   5520
      Width           =   150
   End
   Begin VB.Label RedValue 
      BackColor       =   &H008080FF&
      Caption         =   "0"
      Height          =   255
      Left            =   5640
      TabIndex        =   33
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "LimitLine"
      Height          =   255
      Left            =   5880
      TabIndex        =   29
      Top             =   6120
      Width           =   615
   End
End
Attribute VB_Name = "DebugForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub hough_debugslider1 Lib "laserlib.dll" (ByVal value As Long)
Private Declare Sub hough_debugslider2 Lib "laserlib.dll" (ByVal value As Long)
Private Declare Sub hough_debugslider3 Lib "laserlib.dll" (ByVal value As Long)
Private Declare Sub hough_debugslider4 Lib "laserlib.dll" (ByVal value As Long)
Private Declare Sub hough_debugbutton1 Lib "laserlib.dll" ()
Private Declare Sub hough_debugbutton3 Lib "laserlib.dll" ()
Private Declare Sub hough_SetColourAdjust Lib "laserlib.dll" (ByVal Red As Double, ByVal Green As Double, ByVal Blue As Double)

Private Declare Sub hough_centreOff Lib "laserlib.dll" ()
Private Declare Sub hough_centreOn Lib "laserlib.dll" ()

Private Declare Sub ThreeDim_Initialise Lib "OpenGLThreeDim.dll" (ByVal hwnd As Long)
Private Declare Sub ThreeDim_RenderScene Lib "OpenGLThreeDim.dll" ()



Private Declare Sub clearline_TestingDrawingLines Lib "Clearline.dll" (ByVal hwnd As Long)

Private Declare Sub autofisheye Lib "laserlib.dll" ()

Public BtnLaserOneDown As Boolean
Public BtnLaserTwoDown As Boolean
Public BtnLaserCentreDown As Boolean
Public BtnLaserLeftSideDown As Boolean
Public BtnLaserRightSideDown As Boolean

Private SelectedGraph As String
Private graphScale As Double
Dim CancelProc As Boolean





Public FishEyeGrid As Boolean

Public FeX, FeY As Integer

Private Sub Btn1Debug_Click()
    BtnLaserOneDown = True
    BtnLaserTwoDown = False
    BtnLaserCentreDown = False
    
    
    ClearLineScreen.VideoScreen.SetFocus
    
End Sub

Private Sub Btn2Debug_Click()

End Sub

Private Sub Btn1Button_Click()
    hough_debugbutton1
    ClearLineScreen.RefreshVideoScreen
    If Lab1Debug.Caption = "Off" Then
        Lab1Debug.Caption = "On"
    Else
        Lab1Debug.Caption = "Off"
    End If
    
    
    
End Sub

Private Sub Btn3Debug_Click()

    hough_debugbutton3
    ClearLineScreen.RefreshVideoScreen
    If Lab3Debug.Caption = "Off" Then
         Lab3Debug.Caption = "On"
    Else
        Lab3Debug.Caption = "Off"
    End If
    

End Sub

Private Sub AutoRotate_Click()
Dim i As Integer
Dim J As Integer

    Dim RotFrom As Long
    Dim RotTo As Long
    CancelProc = False
    
    RotFrom = CInt(RotateFrom.text)
    RotTo = CInt(RotateTo.text)



    If DrawShapeType = "Circle" Then
        MsgBox "Can not align cirular reference, please choose another"
        Exit Sub
    End If
    
    If RotFrom = RotTo Then
        For i = 0 To (PVDataNoOfLines / 10)
            Call CLPProgressBar.ProgressBarPosition((i / PVDataNoOfLines * 10))
            DoEvents
        
            Call ScreenDrawing.PVAutoRotateCPP((i * 10), 9 + (i * 10))
            Call ScreenDrawing.PVShapeCentreCalcCPP((i * 10), 9 + (i * 10))
            'Call ScreenDrawing.PVCentreCalcCPP((I * 10), 9 + (I * 10))
            Call ScreenDrawing.PVFlat3DCalcCPP((i * 10), 9 + (i * 10)) 'PCN3513
    
            ClearLineScreen.UpdateThreeDimColour
            If CancelProc = True Then Exit For
        Next i
    Else
        For i = RotFrom To RotTo
                If i Mod 10 = 0 Then
                    Call CLPProgressBar.ProgressBarPosition(i / RotTo)
                    DoEvents
                End If
                
                Call ScreenDrawing.PVAutoRotateCPP(i, i)
                Call ScreenDrawing.PVShapeCentreCalcCPP(i, i)
                'Call ScreenDrawing.PVCentreCalcCPP(I, I)
                Call ScreenDrawing.PVFlat3DCalcCPP(i, i) 'PCN3513
        
                ClearLineScreen.UpdateThreeDimColour
                If CancelProc = True Then Exit For
        Next i
        
    End If
    CancelProc = False
    Call CLPProgressBar.ProgressBarPosition(1)
    

    
End Sub



Private Sub BtnProcessDummyImage_Click()
Call ProcessSingleImage
End Sub



Private Sub CalculateTotal_Click()

Dim Yards As Double
Dim Total As Double
Dim i As Long

For i = CLng(Me.DebrisStartFrame.text) To CLng(Me.DebrisEndFrame.text) - 1
    Yards = Abs(PVDistances(i) - PVDistances(i + 1))
    Yards = Yards / 3
    If GraphInfoContainer(PVDebris).DataSingle(i) > 0 Then
        Total = Total + (GraphInfoContainer(PVDebris).DataSingle(i) * Yards)
    End If
Next i

    Me.TotalDebris = Round(Total, 2) & " cubic yards"
    

End Sub

Private Sub Calibrate_Click()

Call autofisheye

End Sub



Private Sub CancelBtn_Click()
    CancelProc = True
End Sub

Private Sub CapacityOffsetValue_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        CapacityDataOffset = CSng(CapacityOffsetValue.text)
        Call ScreenDrawing.ClearAllGraphsAndRuler
        Call ScreenDrawing.DrawPVGraphs 'PCN3373
        Call ScreenDrawing.DrawPVXScale
    End If
End Sub

Private Sub CentreLevelOption_Click(Index As Integer)
    If CentreLevelOption(0).value = True Then
        OuterEdgeCentre = 0
    Else
        OuterEdgeCentre = 1
    End If
End Sub

Private Sub Check1_Click()
    If Check1.value = 1 Then
        WaterLevelCentre = 2
    Else
        WaterLevelCentre = 0
    End If
    
End Sub

Private Sub CleanProfile_Click()
    Call PVCleanUpProfileCPP(1, PVDataNoOfLines)
End Sub

Private Sub Command1_Click()
    Dim Shift As Double
    Shift = SafeCDbl(Text6Debug.text)
    Shift = Shift + CentreLineX  'Displace by screen centre reference
'    Shift = Shift - X4Circle2(1) 'negate the current shape centre to get shift value
    
    Text1Debug = SafeCDbl(Text6Debug.text)
    Call ClearLineScreen.AutoMoveShapes(ClearLineScreen.PVScreen, Shift, 0)
    
    Call ScreenDrawing.DrawProfilesStartToFinish(ClearLineScreen.PVScreen, True) 'PCN1835
    Call ScreenDrawing.DrawAll
    
End Sub

Private Sub Command2_Click()
    Dim Shift As Double
    Shift = SafeCDbl(Text7Debug.text)
    Shift = Shift + CentreLineY
'    Shift = Shift - Y4Circle2(1)
    
    Text2Debug = SafeCDbl(Text7Debug.text)
    
    Call ClearLineScreen.AutoMoveShapes(ClearLineScreen.PVScreen, 0, Shift)
    Call ScreenDrawing.DrawProfilesStartToFinish(ClearLineScreen.PVScreen, True) 'PCN1835

    
    
End Sub

Private Sub Command3_Click()
    Text5Debug = SafeCDbl(Text8Debug.text)
    ShapeRotationAngle = SafeCDbl(Text8Debug.text) / 180 * PI
    
    Call ScreenDrawing.DrawProfilesStartToFinish(ClearLineScreen.PVScreen, True) 'PCN1835
    ClearLineScreen.SnapShotScreen.AutoRedraw = True
    Call ScreenDrawing.DrawAll
   
    
    
    

End Sub









Private Sub Command4_Click()
    Call ScreenDrawing.RecalculatePVData
    
End Sub

Private Sub Command5_Click()
'SonarConfig.Show
End Sub







Private Sub Command8_Click()

End Sub





Private Sub CutWaterLevelButton_Click()
    If WaterLevelCentre = 3 Then
        WaterLevelCentre = 0
        CutWaterLevelCheck.value = 0
    Else
        WaterLevelCentre = 3
        CutWaterLevelCheck.value = 1
        ReduceWaterLevelCheck.value = 0
    End If
End Sub

Private Sub DebrisCalc_Click()
    Dim i As Long
    
    Call CLPProgressBar.ProgressBarInitialise(DisplayMessage("Processing Data"))
    DoEvents
    
    For i = 0 To (PVDataNoOfLines / 100)
        Call CLPProgressBar.ProgressBarPosition((i / PVDataNoOfLines * 100))
        DoEvents
        Call ScreenDrawing.PVDebrisCalcCPP((i * 100), 99 + (i * 100)) ' pcn4461
    Next i
    
    Call CLPProgressBar.ProgressBarPosition(1)
End Sub

Private Sub DebrisOnCheck_Click()
    If DebrisOnCheck.value = 1 Then
        DebrisOn = True
    Else
        DebrisOn = False
        ReDim DrawingIntegratedArea(0)
        Call DrawProfilesStartToFinish(ClearLineScreen.PVScreen, True)
        Call DrawProfilesStartToFinish(ClearLineScreen.SnapShotScreen, True)
    End If
    
End Sub

Private Sub DrawTrackingAmountBtn_Click()
    If StartLine < 1 Then StartLine = 1
    If EndLine > TrackingPosition.NoCompassReadings Then EndLine = TrackingPosition.NoCompassReadings
    
    TrackingPosition.StartDrawCompassData = StartLine
    TrackingPosition.EndDrawCompassData = EndLine
    
    Call TrackingPosition.DisplayCompassData
    
    
End Sub

Private Sub DrawTrackingLabelResolution_Click()

    
    TrackingPosition.DrawingLabelResolution = LabelResolutionTextBox.text

    If TrackingPosition.DrawingLabelResolution < 1 Then TrackingPosition.DrawingLabelResolution = 1
    TrackingResolutionTextBox.text = TrackingPosition.DrawingCompassResolution
    
    Call TrackingPosition.DisplayCompassData
    
End Sub

Private Sub DrawTrackingResolution_Click()
    TrackingPosition.DrawingCompassResolution = TrackingResolutionTextBox.text
    If TrackingPosition.DrawingCompassResolution < 1 Then TrackingPosition.DrawingCompassResolution = 1
    TrackingResolutionTextBox.text = TrackingPosition.DrawingCompassResolution
    
    Call TrackingPosition.DisplayCompassData
End Sub

Private Sub DumpBtn_Click()
Dim FileNo
Dim myloop As Long
Dim Segment As Long
Dim Index As Long

FileNo = FreeFile

Open "C:\XYDump.txt" For Output As #FileNo

For myloop = 0 To PVDataNoOfLines
    For Segment = 0 To 179
        Index = Segment + (myloop * 180)
        Print #FileNo, "Index " & Index & "," & TD_PVDataX(Index) & "," & TD_PVDataY(Index)
    Next Segment
Next myloop





Close #FileNo
End Sub

'Private Sub Form_Load()
'    BtnLaserOneDown = False
'    BtnLaserTwoDown = False
'    BtnLaserCentreDown = False
'    BtnLaserLeftSideDown = False
'    BtnLaserRightSideDown = False
'
'    Slider4Debug.Value = 200
'
'    FishEyeGrid = False
'
'    ReDim FisheyeFunctions.FEGrid(4, 3, 2) As Integer '(x,y,0) is x (x,y,1) is y
'    ReferenceShapeTxtBox.text = ReferenceShape(SemiEllipticalType).name
'    width = 5940
'    height = 7800
'
'
'End Sub

'Private Sub SaveGrid_Click()
'
'
'Dim i, t As Integer
'
'
'
'Open "c:\grid.dat" For Output As #123
'
'For i = 0 To 3
'    For t = 0 To 4
'
'        Print #123, FisheyeFunctions.FEGrid(t, i, 0)
'        Print #123, FisheyeFunctions.FEGrid(t, i, 1)
'
'    Next t
'Next i
'
'Close #123

'End Sub

Private Sub SetMedianLimits_Click()
    DiameterMedianLimitR = 1800
    DiameterMedianLimitL = 600
    PVGraphDiaMedianXOffset = 50
    PVGraphDiaMedianXScale = 100
    
    
End Sub

Private Sub EgnoreCentreHeight_Click()
    If EgnoreCentreHeight.value = vbChecked Then
        WaterLevelCentre = True
    Else
        WaterLevelCentre = False
    End If
End Sub

Private Sub FilterBtn_Click()
Call CPPFilterGraphs
End Sub

Private Sub Form_Load()
    Me.width = 6090
    If OuterEdgeCentre = False Then
        CentreLevelOption(0).value = True
        CentreLevelOption(1).value = False
    Else
        CentreLevelOption(0).value = False
        CentreLevelOption(1).value = True
    End If
    
    If WaterLevelCentre = 2 Then
         ReduceWaterLevelCheck.value = 1
    Else
         ReduceWaterLevelCheck.value = 0
    End If
    
    If WaterLevelCentre = 3 Then
        CutWaterLevelCheck.value = 1
    Else
        CutWaterLevelCheck.value = 0
    End If

'    If WaterLevelCentre = False Then
'        EgnoreCentreHeight.value = vbUnchecked
'    Else
'        EgnoreCentreHeight.value = vbChecked
'    End If
    

    
    If DebrisOn Then
        DebrisOnCheck.value = 1
    Else
        DebrisOnCheck.value = 0
    End If
    
    If ShiftOn Then
        ShiftCheckBox.value = 1
    Else
        ShiftCheckBox.value = 0
    End If
    
    Me.DebrisStartFrame.text = 1
    Me.DebrisEndFrame = PVDataNoOfLines
End Sub



Private Sub Load2ndPVD_Click()
On Error GoTo Error_handler
    
    Dim ToOpenFileName As String
    
    'Me.CommonDialogLoad2ndDVD.ShowOpen
    'ToOpenFileName = Me.CommonDialogLoad2ndDVD.FileName
    
    If ToOpenFileName = "" Then Exit Sub
    Call PageFunctions.Open2ndPVDData(ToOpenFileName)
Exit Sub
Error_handler:
    Select Case Err
        Case Else: MsgBox Error$ & " - " & Err
    End Select
End Sub

Private Sub LoadCustomShape_Click()
Dim i As Integer
Dim ReferenceShapeNo As Integer

ReferenceShapeNo = ScreenDrawing.GetNumShapeType(DrawShapeType)

For i = 0 To 3
    With ReferenceShape(ReferenceShapeNo)
        ShapeOriginX(i).text = .Arcs(i).OriginX: ShapeOriginY(i).text = .Arcs(i).OriginY
        Radius(i).text = .Arcs(i).Radius
        StartAngle(i).text = .Arcs(i).StartAngle
        EndAngle(i).text = .Arcs(i).EndAngle
    End With
Next i
    
    
End Sub

Public Sub LockDonutBtn_Click()
    Call hough_lockdonut(ExpectedDiameter / VideoScreenScale * 1.5)
    LockedDonut = True
    Call ClearLineScreen.RefreshVideoScreen
    
    
End Sub

Private Sub PVReportButton_Click()
    Load PVReport1K
End Sub

Private Sub ReduceWaterLevelBtn_Click()
    If WaterLevelCentre = 2 Then
        ReduceWaterLevelCheck.value = 0
        WaterLevelCentre = 0
    Else
        ReduceWaterLevelCheck.value = 1
        CutWaterLevelCheck.value = 0
        WaterLevelCentre = 2
    End If
    

        
        
End Sub



Private Sub RotateLeft_Click()
    Dim RotationValue As Single
    Dim RotFrom As Long
    Dim RotTo As Long
    Dim i As Long
    CancelProc = False
    
    RotFrom = CInt(RotateFrom.text)
    RotTo = CInt(RotateTo.text)
    RotationValue = CSng(Me.RotationPVDValue.text)
    
    If RotFrom = RotTo Then
        Call RotatePVDData(RotationValue)
    Else
        For i = RotFrom To RotTo
            Call RotatePVDData(RotationValue, i)
            If CancelProc = True Then Exit For
        Next i
    End If
    CancelProc = False
    
End Sub

Private Sub RotateRight_Click()
    Dim RotationValue As Single
    Dim RotFrom As Long
    Dim RotTo As Long
    Dim i As Long
    
    RotFrom = CInt(RotateFrom.text)
    RotTo = CInt(RotateTo.text)
    RotationValue = CSng(Me.RotationPVDValue.text)
    
    If RotFrom = RotTo Then
        Call RotatePVDData(-RotationValue)
    Else
        For i = RotFrom To RotTo
            Call RotatePVDData(-RotationValue, i)
        Next i
    End If
End Sub

Private Sub SaveDistanceInfo_Click()
    Call PageFunctions.SaveDistanceCalculations
    
End Sub

Private Sub SaveNewPVDData_Click()
    Call PageFunctions.SaveFullPVDataToFile
End Sub

Private Sub Slider1_Change()
    SemiEllipticalType = Slider1.value
    ReferenceShapeTxtBox.text = ReferenceShape(SemiEllipticalType).Name
    
End Sub

Private Sub Slider1_Scroll()
     ReferenceShapeTxtBox.text = ReferenceShape(Slider1.value).Name
 End Sub

Private Sub ScaleBtn_Click()
    Dim ScaleAmount As Single
    Dim CurrentScale As Single
    Dim RefereceShapeNo As Integer
    Dim i As Integer
    
    If Not IsNumeric(Me.ScaleValue.text) Then Exit Sub
    ScaleAmount = CSng(Me.ScaleValue.text)
    CurrentScale = CSng(Me.CurrentScale.text)
    
    CurrentScale = ((100 - CurrentScale) / 100) * 2
    ScaleAmount = ((100 - ScaleAmount) / 100) * 2
    
    
    ReferenceShapeNo = ScreenDrawing.GetNumShapeType(DrawShapeType)
    

    With ReferenceShape(ReferenceShapeNo)
        .CentreOffsetX = .CentreOffsetX / CurrentScale: .CentreOffsetX = .CentreOffsetX * ScaleAmount
        .CentreOffsetY = .CentreOffsetY / CurrentScale: .CentreOffsetY = .CentreOffsetY * ScaleAmount
        For i = 0 To .NoArcs - 1
            With .Arcs(i)
                .OriginX = .OriginX / CurrentScale: .OriginX = .OriginX * ScaleAmount
                .OriginY = .OriginY / CurrentScale: .OriginY = .OriginY * ScaleAmount
                .Radius = .Radius / CurrentScale: .Radius = .Radius * ScaleAmount
            End With
        Next i
        
        For i = 0 To .NoLines - 1
            With .Lines(i)
                .EndX = .EndX / CurrentScale: .EndX = .EndX * ScaleAmount
                .EndY = .EndY / CurrentScale: .EndY = .EndY * ScaleAmount
                .StartX = .StartX / CurrentScale: .StartX = .StartX * ScaleAmount
                .StartY = .StartY / CurrentScale: .StartY = .StartY * ScaleAmount
            End With
        Next i
    End With

    
    
    Me.CurrentScale.text = Me.ScaleValue.text
    Call ScreenDrawing.DrawProfilesStartToFinish(ClearLineScreen.PVScreen, True)
End Sub

Private Sub ShiftCheckBox_Click()
    Dim a As Integer
    If ShiftCheckBox.value = 1 Then
        ShiftOn = True
        PrecisionVisionGraph.WaterShiftControls.Visible = True
        PrecisionVisionGraph.ReferenceShiftControls.Visible = True
        Observations.SortObs
    Else
        ShiftOn = False
        ReDim Observations.ReferenceShapeShiftObs(0)
        ReDim Observations.WaterLevelShiftObs(0)
        PrecisionVisionGraph.WaterShiftControls.Visible = False
        PrecisionVisionGraph.ReferenceShiftControls.Visible = False
    End If
End Sub

Private Sub Slider1Debug_Scroll()
    hough_debugslider1 (Slider1Debug.value)
    Text1Debug.text = Slider1Debug.value
    If Slider1Debug.value + 1 > Slider2Debug.value Then
        Slider2Debug.value = Slider1Debug.value + 1
        Text2Debug.text = Slider2Debug.value
        hough_debugslider2 (Slider2Debug.value)
    End If
 '   ClearLineScreen.RefreshVideoScreen
    
End Sub

Private Sub Slider2Debug_Scroll()
    hough_debugslider2 (Slider2Debug.value)
    Text2Debug.text = Slider2Debug.value
    If Slider2Debug.value - 1 < Slider1Debug.value Then
        Slider1Debug.value = Slider2Debug.value - 1
        Text1Debug.text = Slider1Debug.value
        hough_debugslider1 (Slider1Debug.value)

    End If
    ClearLineScreen.RefreshVideoScreen
End Sub

Private Sub Slider3Debug_Click()
    Text3Debug.text = Slider3Debug.value * -1
    hough_debugslider3 (Slider3Debug.value * -1)
    ClearLineScreen.RefreshVideoScreen
End Sub

Private Sub Slider4Debug_Click()
    Text4Debug.text = Slider4Debug.value
    hough_debugslider4 (Slider4Debug.value)
    ClearLineScreen.RefreshVideoScreen
End Sub

Private Sub SliderBlue_Change()
    Call AdjustColourFilter
End Sub

Private Sub SliderBlue_Scroll()
    Call AdjustColourFilter
End Sub

Private Sub SliderGreen_Change()
    AdjustColourFilter
End Sub
Private Sub SliderGreen_Scroll()
    AdjustColourFilter
End Sub
Private Sub SliderRed_Change()
    Call AdjustColourFilter
End Sub

Private Sub SliderRed_Scroll()
    Call AdjustColourFilter
End Sub

Private Sub SmoothButton_Click()
    'Call SmoothOutOvality
    Call CPPSmoothGraphs
End Sub

Private Sub StartGrid_Click()

    FishEyeGrid = True
    
    FeX = 0
    FeY = 0

End Sub

Private Sub StoreCustomShape_Click()
Dim i As Integer
Dim ReferenceShapeNo As Integer

ReferenceShapeNo = ScreenDrawing.GetNumShapeType(DrawShapeType)

For i = 0 To 3
    With ReferenceShape(ReferenceShapeNo)
        .Arcs(i).OriginX = ShapeOriginX(i).text: .Arcs(i).OriginY = ShapeOriginY(i).text
        .Arcs(i).Radius = Radius(i).text
        .Arcs(i).StartAngle = StartAngle(i).text
        .Arcs(i).EndAngle = EndAngle(i).text
    End With
Next i

    Call ScreenDrawing.DrawProfilesStartToFinish(ClearLineScreen.PVScreen, True)
End Sub



Private Sub StretchBtn_Click()
    Dim StretchAmount As Single
    Dim CurrentStretch As Single
    Dim RefereceShapeNo As Integer
    Dim i As Integer

    If Not IsNumeric(Me.StretchValue.text) Then Exit Sub
    StretchAmount = CSng(Me.StretchValue.text)
    CurrentStretch = CSng(Me.CurrentStretch.text)
    
    CurrentStretch = ((100 - CurrentStretch) / 100) * 2
    StretchAmount = ((100 - StretchAmount) / 100) * 2
    
    ReferenceShapeNo = ScreenDrawing.GetNumShapeType(DrawShapeType)
    If DrawShapeType = "Elliptical ASTM C507" Or DrawShapeType = "CupCake" Then
         With ReferenceShape(ReferenceShapeNo)
             .Arcs(1).OriginY = .Arcs(1).OriginY - CurrentStretch: .Arcs(1).OriginY = .Arcs(1).OriginY + StretchAmount
             .Lines(0).EndY = .Lines(0).EndY - CurrentStretch: .Lines(0).EndY = .Lines(0).EndY + StretchAmount
             .Lines(1).EndY = .Lines(1).EndY - CurrentStretch: .Lines(1).EndY = .Lines(1).EndY + StretchAmount
         End With
    End If
    If DrawShapeType = "BoxCulvert" Then
        With ReferenceShape(ReferenceShapeNo)
            For i = 3 To 6
                .Lines(i).EndY = .Lines(i).EndY - CurrentStretch: .Lines(i).EndY = .Lines(i).EndY + StretchAmount
                .Lines(i + 1).StartY = .Lines(i + 1).StartY - CurrentStretch: .Lines(i + 1).StartY = .Lines(i + 1).StartY + StretchAmount
            Next i
        End With
    End If
    If DrawShapeType = "Square" Then
           With ReferenceShape(ReferenceShapeNo)
            For i = 1 To 2
                .Lines(i).EndY = .Lines(i).EndY - CurrentStretch: .Lines(i).EndY = .Lines(i).EndY + StretchAmount
                .Lines(i + 1).StartY = .Lines(i + 1).StartY - CurrentStretch: .Lines(i + 1).StartY = .Lines(i + 1).StartY + StretchAmount
            Next i
        End With
    End If
    
    
    Me.CurrentStretch.text = Me.StretchValue.text
    Call ScreenDrawing.DrawProfilesStartToFinish(ClearLineScreen.PVScreen, True)
    
      
    
End Sub


Private Sub Text2Debug_Change()
 '       Y4Circle(1) = Val(Text2Debug.text) + CentreLineY
        
End Sub

Private Sub Text2Debug_Validate(Cancel As Boolean)
 '   X4Circle2(1) = Val(Text1Debug.text) + CentreLineX

End Sub

Private Sub Text5Debug_Validate(Cancel As Boolean)
'    ScreenDrawing.ShapeRotationAngle = Val(Text5Debug.text)
End Sub

Private Sub TextBox1Debug_Validate(Cancel As Boolean)
    Dim Number As Double
    Number = SafeCDbl(TextBox1Debug.text) 'PCN4161

    ClearLineScreen.RefreshVideoScreen
End Sub



Private Sub ThreeDimOpenGLWindow_Click()
    ThreeDimOpenGLWindow.AutoRedraw = False
    Call ThreeDim_Initialise(ThreeDimOpenGLWindow.hwnd)
    Call ThreeDim_RenderScene
    ThreeDimOpenGLWindow.AutoRedraw = True
End Sub

Public Sub AdjustColourFilter()

    ImageBlurSize.width = SliderRed.value * 10
    ImageBlurSize.height = SliderRed.value * 10
    RedValue = SliderRed.value

'    Dim Red As Double
'    Dim Green As Double
'    Dim Blue As Double
'
'    RedValue.Caption = Red
'    GreenValue.Caption = Green
'    BlueValue.Caption = Blue
'
'    Red = SliderRed.value
'    Green = SliderGreen.value
'    Blue = SliderBlue.value
'
    Call hough_SetColourAdjust(SliderRed.value, 0, 0)
'    ClearLineScreen.RefreshVideoScreen
    
End Sub

Private Sub TrueDiameterOffset_Change()

End Sub

Private Sub TrackingLoadBtn_Click()
    Call TrackingPosition.LoadAndProcessCompass
End Sub

Private Sub TrueDiameterOffsetValue_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        TrueDiameterOffset = CSng(TrueDiameterOffsetValue.text)
        Call ScreenDrawing.ClearAllGraphsAndRuler
        Call ScreenDrawing.DrawPVGraphs 'PCN3373
        Call ScreenDrawing.DrawPVXScale
    End If
End Sub


Private Sub TurnOffCentreCalc_Click()
   Call hough_centreOff

End Sub

Private Sub TurnOnCentreCalc_Click()
   Call hough_centreOn
End Sub

Private Sub TwoInOneReport_Click()
        Load PVReport2in1
End Sub

Private Sub UnlockDonutBtn_Click()
    Call hough_unlockdonut
    LockedDonut = False
    
    Call ClearLineScreen.RefreshVideoScreen
End Sub
