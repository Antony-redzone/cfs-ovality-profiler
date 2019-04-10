VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form AutoTune 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   9090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6735
   ControlBox      =   0   'False
   Icon            =   "AutoTune.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   Begin VB.Frame TuningFrame 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tuning"
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
      Height          =   3840
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   4095
      Begin VB.TextBox ZoneFeedback 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3425
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "0"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox LaserWidthFeedback 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2585
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "0"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox ContrastFeedback 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "0"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox PickupLevelFeedback 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1000
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "0"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox ColourFeedback 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C4EAAC&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   300
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "0"
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton DetailsExpand 
         BackColor       =   &H00FFFFFF&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3285
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   12
         Tag             =   "1"
         Top             =   3870
         Visible         =   0   'False
         Width           =   435
      End
      Begin MSComctlLib.Slider PVContrast 
         Height          =   1935
         Left            =   3360
         TabIndex        =   13
         Top             =   720
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   3413
         _Version        =   393216
         BorderStyle     =   1
         Orientation     =   1
         LargeChange     =   10
         Max             =   100
         SelStart        =   20
         TickStyle       =   2
         TickFrequency   =   10
         Value           =   20
      End
      Begin MSComctlLib.Slider SDX 
         Height          =   1935
         Left            =   2520
         TabIndex        =   14
         Top             =   720
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   3413
         _Version        =   393216
         BorderStyle     =   1
         Orientation     =   1
         LargeChange     =   10
         Max             =   100
         SelStart        =   15
         TickStyle       =   2
         TickFrequency   =   5
         Value           =   15
      End
      Begin MSComctlLib.Slider GT 
         Height          =   1935
         Left            =   960
         TabIndex        =   15
         Top             =   720
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   3413
         _Version        =   393216
         BorderStyle     =   1
         Orientation     =   1
         Min             =   1
         Max             =   200
         SelStart        =   5
         TickStyle       =   2
         TickFrequency   =   10
         Value           =   5
      End
      Begin MSComctlLib.Slider PVEnhance 
         Height          =   1935
         Left            =   240
         TabIndex        =   16
         ToolTipText     =   "Color"
         Top             =   720
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   3413
         _Version        =   393216
         BorderStyle     =   1
         Orientation     =   1
         LargeChange     =   10
         Max             =   3
         SelStart        =   1
         TickStyle       =   2
         Value           =   1
      End
      Begin MSComctlLib.Slider ContrastSlider 
         Height          =   1935
         Left            =   1680
         TabIndex        =   17
         Top             =   720
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   3413
         _Version        =   393216
         BorderStyle     =   1
         Orientation     =   1
         LargeChange     =   10
         Max             =   100
         TickStyle       =   2
         TickFrequency   =   5
      End
      Begin VB.Label TuningFrameLbl 
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
         TabIndex        =   29
         Top             =   0
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label PVEnhancelbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   675
         Left            =   120
         TabIndex        =   28
         Top             =   2760
         Width           =   675
      End
      Begin VB.Label PVContrast_lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Zone"
         Height          =   315
         Left            =   3240
         TabIndex        =   27
         Top             =   2760
         Width           =   795
      End
      Begin VB.Label SDX_lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Laser width"
         Height          =   630
         Left            =   2520
         TabIndex        =   26
         Top             =   2760
         Width           =   705
      End
      Begin VB.Label GT_lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Pickup Level"
         Height          =   615
         Left            =   840
         TabIndex        =   25
         Top             =   2760
         Width           =   795
      End
      Begin VB.Label Details_lbl 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Details ..."
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
         Height          =   390
         Left            =   315
         TabIndex        =   24
         Top             =   3870
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Image ShowSettings 
         Height          =   480
         Left            =   3255
         Picture         =   "AutoTune.frx":038A
         Tag             =   "0"
         Top             =   3270
         Width           =   480
      End
      Begin VB.Label ShowSettings_lbl 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Show settings in video image"
         ForeColor       =   &H00004000&
         Height          =   420
         Left            =   240
         TabIndex        =   23
         Top             =   3360
         Width           =   2925
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   2400
         X2              =   2400
         Y1              =   120
         Y2              =   3180
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   4080
         X2              =   2040
         Y1              =   3180
         Y2              =   3180
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   2040
         Y1              =   3180
         Y2              =   3180
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   135
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2280
         Width           =   135
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Contrast"
         Height          =   195
         Left            =   1680
         TabIndex        =   18
         Top             =   2760
         Width           =   795
      End
   End
   Begin VB.Frame DistanceInfoFrame 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Distance Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   6360
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox txtFinishDistance 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2640
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtStartDistance 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2640
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.Label FinishDistLabel 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Finish Distance"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   2400
         WordWrap        =   -1  'True
      End
      Begin VB.Label StartDistLabel 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Start Distance"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   2400
         WordWrap        =   -1  'True
      End
      Begin VB.Label DistFinishUnitLabel 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   3600
         TabIndex        =   8
         Top             =   720
         Width           =   255
      End
      Begin VB.Label DistStartUnitLabel 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   3600
         TabIndex        =   6
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Frame MaskFrame 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mask"
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
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   4095
      Begin VB.Label MaskFrameLbl 
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
         TabIndex        =   31
         Top             =   0
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Image cmdSetMaskSectors 
         Height          =   720
         Left            =   1560
         Picture         =   "AutoTune.frx":0FCC
         ToolTipText     =   "Set Mask Sectors for Centre"
         Top             =   180
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Image cmdClearMask 
         Height          =   720
         Left            =   2520
         Picture         =   "AutoTune.frx":2C96
         ToolTipText     =   "Clear Mask"
         Top             =   180
         Width           =   720
      End
      Begin VB.Image cmdSetMask 
         Height          =   720
         Left            =   600
         Picture         =   "AutoTune.frx":4960
         ToolTipText     =   "Set Mask"
         Top             =   180
         Width           =   720
      End
   End
   Begin VB.Frame WaterLevelFrame 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Water / Debris Mask"
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
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   4095
      Begin VB.CommandButton InvertWaterLevel 
         Height          =   255
         Left            =   2160
         TabIndex        =   2
         Top             =   1080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label WaterLevelFrameLbl 
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
         TabIndex        =   30
         Top             =   0
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Image WaterLevelRefShapeMove 
         Height          =   720
         Left            =   1560
         Picture         =   "AutoTune.frx":662A
         ToolTipText     =   "Move Reference Shape"
         Top             =   180
         Width           =   720
      End
      Begin VB.Image WaterLevelClear 
         Height          =   720
         Left            =   2520
         Picture         =   "AutoTune.frx":82F4
         ToolTipText     =   "Clear Water / Debris Mask"
         Top             =   180
         Width           =   720
      End
      Begin VB.Image SetWaterLevelBtn 
         Height          =   720
         Left            =   600
         Picture         =   "AutoTune.frx":9FBE
         ToolTipText     =   "Set Water / Debris Mask (Left click to add mask, Right click to remove)"
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Invert water level"
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Timer TuneTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4560
      Top             =   0
   End
   Begin VB.Shape SelectedTask 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   3930
      Left            =   90
      Top             =   2355
      Visible         =   0   'False
      Width           =   4170
   End
   Begin VB.Image FormBackground 
      Height          =   8655
      Left            =   0
      Picture         =   "AutoTune.frx":BC88
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5025
   End
End
Attribute VB_Name = "AutoTune"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim i As Integer 'PCN1916
Dim NumTests As Integer 'PCN1916
Private StartTime As Double  'The start time of the media file. When tuning has finished
                             'the file will be set back at this place


Private Sub GetData_Click()
   ClearLineScreen.GetBlockOfPVData
End Sub



Public Sub cmdClearMask_Click()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : cmdClearMask_Click
'Created : 25 March 2004, PCN2639
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    : Clear where the distance counter is located on the video screen.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

IgnoreX1 = 1
IgnoreX2 = 1
IgnoreY1 = 1
IgnoreY2 = 1
' BYABE FISH-EYE( PCN2290 )---------------------------v
Call INI_WriteBack(MyFile, "IgnoreX1=", IgnoreX1)
Call INI_WriteBack(MyFile, "IgnoreY1=", IgnoreY1)
Call INI_WriteBack(MyFile, "IgnoreX2=", IgnoreX2)
Call INI_WriteBack(MyFile, "IgnoreY2=", IgnoreY2)
' BYABE FISH-EYE( PCN2290 )---------------------------^
Call ClearLineScreen.SetRectangle(IgnoreX1, IgnoreY1, IgnoreX2, IgnoreY2, "Ingore1") 'PCN2639
'Draw boundary for no-go area
Call ClearLineScreen.SetupVideoIgnoreBoarder(ClearLineScreen.VideoImageIgnoreBoarder, IgnoreX1, IgnoreY1, IgnoreX2, IgnoreY2) 'PCN2639
Call ScreenDrawing.ClearMask


Exit Sub
Err_Handler:
    MsgBox Err & "-AT1:" & Error$
End Sub

Private Sub cmdSetMask_Click()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : cmdSetMask_Click
'Created : 25 March 2004, PCN2639
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    : By defining where the text, not usually the distance counter, is located on the video screen,
'          the image processing can ignore this part of the video and hence ensures
'          the profile is not distorted.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Call ControlsScreen.SetVideoMask 'PCN4596 crap, never should have done the below code

'''PCN 4596 this code should never have been here, it should have called a communal set mask
''Dim bRememberFisheyeState As Boolean 'PCN2907 9 September 2004 Antony
''bRememberFisheyeState = ConfigInfo.FishEyeFlag 'PCN2907 9 September 2004 Antony
''FisheyeFunctions.FEOFF 'PCN2907 9 September 2004 Antony
''IgnoreAreaFlag = True
''DelFlag = False
'''Take a snap shot
''Call ClearLineScreen.TakeASnapShot
''DoEvents
''Call ScreenDrawing.SetupDrawingRoutine("DrawMaskBox")
'''Call ClearLineScreen.SetupMouseIcon(118) 'Ignore Text Region
'''ClearLineScreen.Select_Click
''If bRememberFisheyeState Then FisheyeFunctions.FEON 'PCN2907 9 September 2004 Antony

Exit Sub
Err_Handler:
    MsgBox Err & "-AT2:" & Error$
End Sub

Private Sub ContrastSlider_Change()
    Call UpdateSliderFeedback
End Sub

Private Sub ContrastSlider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    
    
If mediatype = Video Or mediatype = "Live" Or mediatype = StillImage Then 'PCN3194
    Call setprofileoverlay(2) 'PCN3017
    Call setprofilecandidates(1) 'PCN3017
    ClearLineScreen.RefreshVideoScreen
End If
    
End Sub

Private Sub ContrastSlider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If mediatype = Video Or mediatype = "Live" Or mediatype = StillImage Then 'PCN3194
    If ShowSettings.Tag = 1 Then
        'Final profile ON
        Call setprofileoverlay(1)
    Else
        'Final profile OFF
        Call setprofileoverlay(0)
    End If
    Call setprofilecandidates(0) 'PCN3017
    ClearLineScreen.RefreshVideoScreen
End If
End Sub

Private Sub ContrastSlider_Scroll()
    Dim contrast As Double
    Dim brightness As Double
    
        Call UpdateSliderFeedback
    contrast = CDbl(ContrastSlider.value)
    brightness = CDbl(0)
    




If mediatype = Video Or mediatype = "Live" Or mediatype = StillImage Then 'PCN3194
    Call houge_AdjustContrastBright(contrast, brightness)
    Call setgradthreshold(ConfigInfo.IPGradThres) 'PCN2820
    ClearLineScreen.RefreshVideoScreen 'PCN3017
End If
End Sub

Private Sub DetailsExpand_Click()

On Error GoTo Err_Handler

'vvvv PCN3017 *******************
'Load ImageProcess
'ImageProcess.Visible = True
'ImageProcess.ZOrder 0
'Unload Me
'^^^^ ***************************
    
Exit Sub
Err_Handler:
    MsgBox Err & "-AT3:" & Error$
End Sub


Private Sub GT_Change()
    Call UpdateSliderFeedback
End Sub

'Private Sub Graphs_Click() ' PCN2612 - New layout, no longer used
' If Graphs.value = 1 Then
'        DrawGraphsNow = True
'    Else
'        DrawGraphsNow = False
'    End If
'End Sub

Private Sub GT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : GT_MouseDown
'Created : 6 May 2004, PCN2612
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim IPOn As Long 'PCN4429
Dim ErrorStr As String 'PCN4429

Call getimageanalysis(IPOn) 'PCN4429
'If IPOn And CLPScreenMode = SnapShot
If CLPScreenMode <> Video Then Call ClearLineScreen.VideoScreenActivate(ErrorStr)  'PCN4429

ConfigInfo.IPGradThres = GT.value 'PCN2820
'If ConfigInfo.IPProcessMethod <> "Type1" Then 'PCN2820
'    ConfigInfo.IPProcessMethod = "Type1" 'PCN2820
'    Call SetPVTuningProcessValues
'End If

'Turn on Selection filter while changing GT
ShowGreenX = 1
ShowGreenY = 1
ShowProfileCandidates = 1 'PCN3017
If mediatype = Video Or mediatype = "Live" Or mediatype = StillImage Then 'PCN3194
    'Final profile Exploded
    Call setprofileoverlay(2) 'PCN3017
'    Call setselectionfilter(ShowGreenX, ShowGreenY)
    Call setprofilecandidates(ShowProfileCandidates) 'PCN3017
    ClearLineScreen.RefreshVideoScreen
End If



Exit Sub
Err_Handler:
    MsgBox Err & "-AT4:" & Error$
End Sub

Private Sub GT_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : GT_MouseUp
'Created : 11 May 2004, PCN2612
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Call SavePVTuningToINI

'Turn off Selection filter
ShowGreenX = 0
ShowGreenY = 0

DoEvents
If mediatype = Video Or mediatype = "Live" Or mediatype = StillImage Then 'PCN3194
    ShowProfileCandidates = 0 'PCN3017
    If ShowSettings.Tag = 1 Then
        'Final profile ON
        Call setprofileoverlay(1)
    Else
        'Final profile OFF
        Call setprofileoverlay(0)
    End If
'    Call setselectionfilter(ShowGreenX, ShowGreenY)
    Call setprofilecandidates(ShowProfileCandidates) 'PCN3017
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-AT5:" & Error$
End Sub

Private Sub GT_Scroll()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : GT_Scroll
'Created : 11 May 2004, PCN2612
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    Call UpdateSliderFeedback
GT.text = GT.value
ConfigInfo.IPGradThres = GT.value 'PCN2820
If mediatype = Video Or mediatype = "Live" Or mediatype = StillImage Then 'PCN3194
    Call setgradthreshold(ConfigInfo.IPGradThres) 'PCN2820
    ClearLineScreen.RefreshVideoScreen 'PCN3017
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-AT6:" & Error$
End Sub

''PCN3219 always profile the water level.
''Private Sub ProfIgnore_Click()
'''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'''ProfIgnore_Click sub Louise Shrimpton
'''
'''Description:
'''   Sets the waterlevel ignore variables according to the checkbox's value
'''
'''Purpose:
'''
'''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''On Error GoTo Err_Handler
''
'''vvvv PCN2612 ************************************
''If ProfIgnore.Tag = 0 Then
''    WaterLevelIgnoreProfile = True
''    Call SetCheckBoxTick(ProfIgnore, True)
''Else
''    WaterLevelIgnoreProfile = False
''    Call SetCheckBoxTick(ProfIgnore, False)
''End If
'''^^^^ ********************************************
''Call ClearLineScreen.SetUnsetWaterLevel(WaterLevelIgnoreCenter, WaterLevelIgnoreProfile)
''
''Exit Sub
''Err_Handler:
''    MsgBox Err & "-" & error$
''End Sub


Private Sub Form_Load()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Form_Load sub Louise Shrimpton louiseS@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    18/03/03
'
'Description:
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

ConvertLanguage Me, Language 'PCN2111

'vvvv PCN4171 **********************************
'    AutoTune.width = PipelineDetails.width 'PCNGL240403-1
'    AutoTune.height = 8430 '7850 '6500 '3045 'PCNGL240403-1 'PCN1939
'    AutoTune.Top = 0 'PCNGL240403-1 'PCN4171
'    AutoTune.Left = PipelineDetails.Left  'PCNGL240403-1
Me.Left = PVPageLeft
Me.Top = PVPageTop
Me.width = PVPageWidth
Me.height = PVPageHeight

FormBackground.height = PVPageHeight
MaskFrame.BackColor = RGB(172, 196, 231)
WaterLevelFrame.BackColor = RGB(172, 196, 231) 'PCN4171
TuningFrame.BackColor = RGB(172, 196, 231) 'PCN4171 AC C4 EA
DistanceInfoFrame.BackColor = RGB(172, 196, 231) 'PCN4171
'^^^^ ******************************************

    ZoneFeedback.BackColor = RGB(172, 196, 231)
    LaserWidthFeedback.BackColor = RGB(172, 196, 231)
    ContrastFeedback.BackColor = RGB(172, 196, 231)
    PickupLevelFeedback.BackColor = RGB(172, 196, 231)
    ColourFeedback.BackColor = RGB(172, 196, 231)

AutoTune.Visible = True
AutoTune.ZOrder 0
'Call ClearLineScreen.GetAVITime  'get the current time into the variable
StartTime = ClearLineScreen.GetCurrentAVITime  'get the variable


Call ClearLineScreen.ProfilerPause

' PCN3215
Call SetupIPCheckBox 'PCN4171

''PCN3219 always profile the water level
''If WaterLevelIgnoreProfile = True Then
''    Call SetCheckBoxTick(ProfIgnore, True) 'PCN3215
''Else
''    Call SetCheckBoxTick(ProfIgnore, False) 'PCN2612
''End If

Call SetTuningSliders 'PCN2612

If LanguageCharset <> 0 Then
    Me.TuningFrameLbl.Font.Charset = LanguageCharset: Me.TuningFrameLbl.BackColor = RGB(172, 196, 231): Me.TuningFrameLbl.Caption = DisplayMessage(TuningFrame.Caption): Me.TuningFrameLbl.Visible = True
    Me.MaskFrameLbl.Font.Charset = LanguageCharset: Me.MaskFrameLbl.BackColor = RGB(172, 196, 231): Me.MaskFrameLbl.Caption = DisplayMessage(MaskFrame.Caption): Me.MaskFrameLbl.Visible = True
    Me.WaterLevelFrameLbl.Font.Charset = LanguageCharset: Me.WaterLevelFrameLbl.BackColor = RGB(172, 196, 231): Me.WaterLevelFrameLbl.Caption = DisplayMessage(WaterLevelFrame.Caption): Me.WaterLevelFrameLbl.Visible = True
End If
   
Exit Sub
Err_Handler:
    MsgBox Err & "-AT7:" & Error$
End Sub

Private Sub InvertWaterLevel_Click()
''Dim I As Integer
''
''For I = 0 To 180
''    If WaterEgnoreList(I) = 0 Then
''        WaterEgnoreList(I) = 1
''    Else
''        WaterEgnoreList(I) = 0
''    End If
''Next I
''
''Call SetWaterLevel(WaterEgnoreList(0))
''If PVDataNoOfLines > 1 Then Call ScreenDrawing.RecalculatePVData

End Sub



Private Sub PVContrast_Change()
    Call UpdateSliderFeedback
End Sub

'Private Sub ManualTune_Click() - PCN2612 - New layout, no longer used
''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''ManualTune_Click sub Louise Shrimpton louiseS@cbsys.co.nz
''
''Revision history"
''   V0.0    Louise Shrimpton,    18/03/03
''
''Description:
''       Loads the Manual Tuning form, and unloads itself
''Purpose:
''
''
''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'On Error GoTo Err_Handler
'    Load ImageProcess
'    ImageProcess.Visible = True
'    ImageProcess.ZOrder 0
'    'Call ImageProcess.SetManualTuningVariables
'    Unload Me
'Exit Sub
'Err_Handler:
'    MsgBox Err & "-" & error$
'End Sub



'Private Sub SetGraphs_Click() 'PCN2612 - No longer required
'    If Graphs.value = 1 Then
'        DrawGraphsNow = True
'    Else
'        DrawGraphsNow = False
'    End If
'End Sub


Private Sub PVContrast_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PVContrast_MouseDown
'Created : 6 May 2004, PCN2612
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim IPOn As Long 'PCN4429
Dim ErrorStr As String 'PCN4429

Call getimageanalysis(IPOn) 'PCN4429
'If IPOn And CLPScreenMode = SnapShot Then Call ClearLineScreen.VideoScreenActivate(ErrorStr) 'PCN4429
If CLPScreenMode <> Video Then Call ClearLineScreen.VideoScreenActivate(ErrorStr)  'PCN4429

ConfigInfo.IPZone = PVContrast.value 'PCN3017
If mediatype <> StillImage Then ClearLineScreen.RefreshVideoScreen
Call setprofileoverlay(1) 'PCN3017

Exit Sub
Err_Handler:
    MsgBox Err & "-AT8:" & Error$
End Sub

Private Sub PVContrast_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PVContrast_MouseUp
'Created : 11 May 2004, PCN2612
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

ConfigInfo.IPZone = PVContrast.value 'PCN2820
Call SavePVTuningToINI

If ShowSettings.Tag = 1 Then
    'Final profile ON
    Call setprofileoverlay(1)
Else
    'Final profile OFF
    Call setprofileoverlay(0)
End If


Exit Sub
Err_Handler:
    MsgBox Err & "-AT9:" & Error$
End Sub

Private Sub PVContrast_Scroll()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PVContrast_Scroll
'Created : 11 May 2004, PCN2612
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Call UpdateSliderFeedback
ConfigInfo.IPZone = PVContrast.value 'PCN2820
If mediatype = Video Or mediatype = "Live" Or mediatype = StillImage Then 'PCN3194
    'Call setprofilecontrast(ConfigInfo.IPZone) 'PCN2820 PCN3017
    Call hough_setinsidezone(CDbl(ConfigInfo.IPZone)) 'PCN3017
    Call hough_setoutsidezone(CDbl(ConfigInfo.IPZone * 2)) 'PCN3017
    If mediatype <> StillImage Then ClearLineScreen.RefreshVideoScreen

End If


Exit Sub
Err_Handler:
    MsgBox Err & "-ATA:" & Error$
End Sub

Private Sub PVEnhance_Change()
    Call UpdateSliderFeedback
End Sub

Private Sub PVEnhance_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
Dim IPOn As Long 'PCN4429
Dim ErrorStr As String 'PCN4429

Call getimageanalysis(IPOn) 'PCN4429
If IPOn And CLPScreenMode = SnapShot Then Call ClearLineScreen.VideoScreenActivate(ErrorStr) 'PCN4429

    Call setimageanalysis(1) 'PCN3017 set the video grayscale
    If mediatype <> StillImage Then ClearLineScreen.RefreshVideoScreen  'PCN3017 set the video grayscale

Exit Sub
Err_Handler:
    MsgBox Err & "-ATB:" & Error$
End Sub

Private Sub PVEnhance_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PVEnhance_MouseUp
'Created : 11 May 2004, PCN2612
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim IPOn As Long 'PCN4429
Dim ErrorStr As String 'PCN4429

Call getimageanalysis(IPOn) 'PCN4429
'If IPOn And CLPScreenMode = SnapShot Then Call ClearLineScreen.VideoScreenActivate(ErrorStr) 'PCN4429
If CLPScreenMode <> Video Then Call ClearLineScreen.VideoScreenActivate(ErrorStr)  'PCN4429

Call SavePVTuningToINI
    If ShowSettings.Tag = 0 Then 'PCN3017 set the video grayscale
        Call setimageanalysis(0) 'PCN3017 set the video grayscale
    End If
    ClearLineScreen.RefreshVideoScreen 'PCN3017 set the video grayscale

Exit Sub
Err_Handler:
    MsgBox Err & "-ATC:" & Error$
End Sub

Private Sub PVEnhance_Scroll()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PVEnhance_Scroll
'Created : 6 May 2004, PCN2612
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
   Call UpdateSliderFeedback
Select Case PVEnhance.value

 

    Case 0
        If Trim(IPEnhancementAndIPProcessMethod.IPEnhancement) = "Low" Then 'PCN2820
            Exit Sub
        Else
            IPEnhancementAndIPProcessMethod.IPEnhancement = "Low" 'PCN2820
            SetPVTuningProcessValues 'PCN3017 set the video grayscale
            If mediatype <> StillImage Then ClearLineScreen.RefreshVideoScreen 'PCN3017 set the video grayscale
        End If
    Case 1
        If Trim(IPEnhancementAndIPProcessMethod.IPEnhancement) = "High" Then 'PCN2820
            Exit Sub
        Else
            IPEnhancementAndIPProcessMethod.IPEnhancement = "High" 'PCN2820
            SetPVTuningProcessValues 'PCN3017 set the video grayscale
            If mediatype <> StillImage Then ClearLineScreen.RefreshVideoScreen 'PCN3017 set the video grayscale
        End If
    Case 2
        If Trim(IPEnhancementAndIPProcessMethod.IPEnhancement) = "Standard" Then 'PCN2820
            Exit Sub
        Else
            IPEnhancementAndIPProcessMethod.IPEnhancement = "Standard" 'PCN2820
            SetPVTuningProcessValues 'PCN3017 set the video grayscale
            If mediatype <> StillImage Then ClearLineScreen.RefreshVideoScreen 'PCN3017 set the video grayscale

        End If
    Case 3
        If Trim(IPEnhancementAndIPProcessMethod.IPEnhancement) = "Mixed" Then 'PCN2820
            Exit Sub
        Else
            IPEnhancementAndIPProcessMethod.IPEnhancement = "Mixed" 'PCN2820
            SetPVTuningProcessValues 'PCN3017 set the video grayscale
            If mediatype <> StillImage Then ClearLineScreen.RefreshVideoScreen 'PCN3017 set the video grayscale
        End If
    Case Else
        Exit Sub
End Select

DoEvents
If mediatype = Video Or mediatype = "Live" Or mediatype = StillImage Then 'PCN3194
    Call setvideofiltertype(PVEnhance.value)
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-ATD:" & Error$
End Sub


Private Sub SDX_Change()
    Call UpdateSliderFeedback
End Sub

Private Sub SDX_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : SDX_MouseDown
'Created : 6 May 2004, PCN2612
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim IPOn As Long 'PCN4429
Dim ErrorStr As String 'PCN4429

Call getimageanalysis(IPOn) 'PCN4429
'If IPOn And CLPScreenMode = SnapShot Then Call ClearLineScreen.VideoScreenActivate(ErrorStr) 'PCN4429
If CLPScreenMode <> Video Then Call ClearLineScreen.VideoScreenActivate(ErrorStr)  'PCN4429

SDX.text = Format(SDX.value, "#0.0") 'PCNAVI
ConfigInfo.IPStDX = SafeCDbl(SDX.text) 'PCNAVI
ConfigInfo.IPStDY = ConfigInfo.IPStDX

ShowLaserWidth = 1 'PCN3017
hough_showlaserwidth (ShowLaserWidth) 'PCN3017
Call setprofileoverlay(1) 'PCN3017
If mediatype <> StillImage Then ClearLineScreen.RefreshVideoScreen

Exit Sub
Err_Handler:
    MsgBox Err & "-ATE:" & Error$
End Sub

Private Sub SDX_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : SDX_MouseUp
'Created : 11 May 2004, PCN2612
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Call SavePVTuningToINI
'PCN3017
'If ShowSettings.Tag = 1 Then
'    'Final profile ON
'    Call setprofileoverlay(1)
'
'End If

ShowLaserWidth = 0 'PCN3017
hough_showlaserwidth (ShowLaserWidth) 'PCN3017
'vvvv PCN3017 *********************
If ShowSettings.Tag = 1 Then
    'Final profile ON
    Call setprofileoverlay(1)
Else
    'Final profile OFF
    Call setprofileoverlay(0)
End If
'^^^^ *****************************


Exit Sub
Err_Handler:
    MsgBox Err & "-ATF:" & Error$
End Sub

Private Sub SDX_Scroll()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : SDX_Scroll
'Created : 11 May 2004, PCN2612
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    Call UpdateSliderFeedback
'SDX.text = Format(SDX.value / 10, "#0.0")
'ConfigInfo.IPStDX = Val(SDX.text) 'PCN2820
'ConfigInfo.IPStDY = ConfigInfo.IPStDX 'PCN2820
'ConfigInfo.PVShapeCentreX = CalXParaYPara 'PCN2820
'ConfigInfo.PVShapeCentreY = ConfigInfo.PVShapeCentreX 'PCN2820
SDX.text = Format(SDX.value, "#0.0") 'PCNAVI
ConfigInfo.IPStDX = SafeCDbl(SDX.text) 'PCNAVI
ConfigInfo.IPStDY = ConfigInfo.IPStDX
If mediatype = Video Or mediatype = "Live" Or mediatype = StillImage Then 'PCN3194
    Call setstandarddeviation(ConfigInfo.IPStDX, 0, 0, 0) 'PCN2820
    If mediatype <> StillImage Then ClearLineScreen.RefreshVideoScreen
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-AT10:" & Error$
End Sub



Private Sub SetWaterLevelBtn_Click()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'SetWaterLevel_Click sub Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    22/05/03  'PCN1939
'
'Description:
'   By the user indicating where the initial water level is relative to the
'   pipe centre, the water level effect on the centre point calculations can
'   be ignored.

'Purpose:
'   To flag that the user wants to set the water level indicator.
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Call ControlsScreen.SetWaterLevel


''Call ScreenDrawing.SetupDrawingRoutine("DrawWaterLevel")
''
''
''If ExpectedDiameter = 0 Then
''    MsgBox DisplayMessage("Please specify internal pipe diameter first."), vbExclamation
''    PipelineDetails.ZOrder 0
''    PipelineDetails.InternalDiameterExpected.SetFocus
''    Exit Sub
''End If
''
''ScreenDrawing.ShowReferenceShape = True
''Call ClearLineScreen.TakeASnapShot
''Call MsgBox(DisplayMessage("Please ensure the main reference shape is centered."), vbOKCancel + vbInformation)  'PCN2111

Exit Sub
Err_Handler:
    MsgBox Err & "-AT11:" & Error$
End Sub


Private Sub ShowSettings_Click()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : ShowSettings_Click
'Created : 12 May 2004, PCN2612
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim OnOff As Long
Dim ErrorStr As String

If CLPScreenMode <> Video Then Call ClearLineScreen.VideoScreenActivate(ErrorStr)  'PCN4429

Call getimageanalysis(OnOff)

If OnOff Then
    Call SetupVideoDisplayForPVTuning
Else
    Call SetupVideoDisplayAsNormal
End If

Call SetupIPCheckBox

Call ControlsScreen.SetDisplayIPSettings


Exit Sub
Err_Handler:
    MsgBox Err & "-AT12:" & Error$
End Sub

Private Sub WaterLevelClear_Click() 'PCN2699
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'WaterLevelClear_Click() Function  Michelle Lindsay michellelindsay@cbsys.co.nz
'
'Revision history
'   V5.3.1    Michelle Lindsay,    11/03/2004     Building initial framework
'
'Description:
'Clearing the water level drawing from the screen and from c code
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim IgnoreCenter As Boolean
Dim IgnoreProfile As Boolean
Dim FileLoadError As Boolean
Dim i As Integer
    
IgnoreCenter = False
IgnoreProfile = False

Call ClearLineScreen.SetUnsetWaterLevel(IgnoreCenter, IgnoreProfile)
Call ScreenDrawing.DeleteWaterLevel
WLStartAngle = 0
WLFinishAngle = 0
For i = 0 To 180
    WaterEgnoreList(i) = 0
Next i

Call ScreenDrawing.RecalculatePVData
Call ScreenDrawing.DeleteWaterLevel
Call PageFunctions.SaveToFilePipeAndConfigInfo("ConfigInfo", FileLoadError) 'Saves the new water level
Call ClearLineScreen.RefreshVideoScreen

Exit Sub
Err_Handler:
    MsgBox Err & "-AT13:" & Error$
End Sub



Function SetTuningSliders()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : SetTuningSliders
'Created : 12 May 2004, PCN2612
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    : This function sets the Tuning sliders to the current settings
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler


GT.text = ConfigInfo.IPGradThres 'PCN2820
GT.value = ConfigInfo.IPGradThres 'PCN2820
SDX.text = ConfigInfo.IPStDX 'PCN2820
SDX.value = ConfigInfo.IPStDX 'PCN2820 'PCN3017 was times by 10
PVContrast.value = ConfigInfo.IPZone 'PCN2820

If Trim(IPEnhancementAndIPProcessMethod.IPEnhancement) = "Low" Then 'PCN2820
    PVEnhance.value = 0
ElseIf Trim(IPEnhancementAndIPProcessMethod.IPEnhancement) = "Standard" Then 'PCN2820
    PVEnhance.value = 2
ElseIf Trim(IPEnhancementAndIPProcessMethod.IPEnhancement) = "High" Then 'PCN2820
    PVEnhance.value = 1
Else
    PVEnhance.value = 3
End If

Exit Function
Err_Handler:
    MsgBox Err & "-AT14:" & Error$
End Function



Sub SetupIPCheckBox()
On Error GoTo Err_Handler
Dim OnOff As Long

Call getimageanalysis(OnOff)
If OnOff = 0 Then
    Call SetCheckBoxTick(ShowSettings, False) 'PCN2612
Else
    Call SetCheckBoxTick(ShowSettings, True)
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-AT15:" & Error$
End Sub

Sub SetupSelectedTask(TaskName As String)
On Error GoTo Err_Handler

SelectedTask.Visible = True

Select Case TaskName
    Case "Mask"
        SelectedTask.Left = MaskFrame.Left - 50
        SelectedTask.Top = MaskFrame.Top - 50
        SelectedTask.height = MaskFrame.height + 100
        SelectedTask.width = MaskFrame.width + 100

    Case "WaterLevel"
        SelectedTask.Left = WaterLevelFrame.Left - 50
        SelectedTask.Top = WaterLevelFrame.Top - 50
        SelectedTask.height = WaterLevelFrame.height + 100
        SelectedTask.width = WaterLevelFrame.width + 100
    
    Case "Tuning"
        SelectedTask.Left = TuningFrame.Left - 50
        SelectedTask.Top = TuningFrame.Top - 50
        SelectedTask.height = TuningFrame.height + 100
        SelectedTask.width = TuningFrame.width + 100
        Call SetupVideoDisplayForPVTuning
        Call SetupIPCheckBox
        Call ControlsScreen.SetupControlDisplayIP
    
    Case "DistanceInfo"
        SelectedTask.Left = DistanceInfoFrame.Left - 50
        SelectedTask.Top = DistanceInfoFrame.Top - 50
        SelectedTask.height = DistanceInfoFrame.height + 100
        SelectedTask.width = DistanceInfoFrame.width + 100
    
    Case Else
        SelectedTask.Visible = False

End Select

Exit Sub
Err_Handler:
    MsgBox Err & "-AT16:" & Error$
End Sub


Private Sub WaterLevelRefShapeMove_Click()
On Error GoTo Err_Handler

Call ScreenDrawing.SetupDrawingRoutine("MoveAll")

Exit Sub
Err_Handler:
    MsgBox Err & "-AT17:" & Error$
End Sub

Sub UpdateSliderFeedback()
    ZoneFeedback.text = PVContrast.value
    LaserWidthFeedback.text = SDX.value
    ContrastFeedback.text = ContrastSlider.value
    PickupLevelFeedback.text = GT.value
    ColourFeedback.text = PVEnhance.value
    ZoneFeedback.Refresh
    LaserWidthFeedback.Refresh
    ContrastFeedback.Refresh
    PickupLevelFeedback.Refresh
    ColourFeedback.Refresh
    
    
End Sub

