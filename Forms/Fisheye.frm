VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Fisheye 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lens Correction"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   8505
   Visible         =   0   'False
   Begin VB.TextBox CameraModelText 
      Height          =   285
      Left            =   1200
      TabIndex        =   25
      Top             =   6810
      Visible         =   0   'False
      Width           =   3015
   End
   Begin ClearLineProfiler.CBS_DropDownBox CameraDropdown 
      Height          =   375
      Left            =   360
      TabIndex        =   20
      Top             =   5400
      Width           =   3600
      _extentx        =   7832
      _extenty        =   661
   End
   Begin VB.CommandButton HiddenManualFecButton 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   3960
      TabIndex        =   24
      Top             =   6240
      Width           =   255
   End
   Begin VB.TextBox EditFecText 
      Height          =   285
      Left            =   1200
      TabIndex        =   23
      Text            =   "Manual FEC Setting password"
      Top             =   6240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox CameraPicture 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   640
      Left            =   160
      ScaleHeight     =   645
      ScaleWidth      =   300
      TabIndex        =   22
      Top             =   5300
      Width           =   300
   End
   Begin VB.Frame fmAuto 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Automatic Calibration"
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
      ForeColor       =   &H00800000&
      Height          =   2145
      Left            =   4320
      TabIndex        =   10
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton btnAutoCal 
         Caption         =   "Go!"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1215
         TabIndex        =   12
         ToolTipText     =   "Starts/Stops automatic calibration."
         Top             =   435
         Width           =   1410
      End
      Begin MSComctlLib.Slider sldFisheyeContrast 
         Height          =   630
         Left            =   300
         TabIndex        =   11
         Top             =   1335
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   1111
         _Version        =   393216
         BorderStyle     =   1
         Enabled         =   0   'False
         Max             =   255
         SelStart        =   150
         TickStyle       =   2
         TickFrequency   =   10
         Value           =   150
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1215
         TabIndex        =   14
         Top             =   435
         Width           =   1410
      End
      Begin VB.Label lblFisheyeContrast 
         BackStyle       =   0  'Transparent
         Caption         =   "Contrast"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   975
         Width           =   1305
      End
   End
   Begin VB.Frame lblCalibrationTitle 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Lens Correction"
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
      Height          =   4875
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4095
      Begin VB.TextBox VideoResolution 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00400000&
         Height          =   330
         Left            =   2820
         TabIndex        =   18
         Text            =   "0"
         Top             =   3120
         Width           =   915
      End
      Begin VB.TextBox FECResolution 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00400000&
         Height          =   330
         Left            =   2820
         TabIndex        =   16
         Text            =   "0"
         Top             =   2640
         Width           =   915
      End
      Begin VB.CommandButton btnLoad 
         Caption         =   "Load"
         Height          =   375
         Left            =   555
         TabIndex        =   5
         ToolTipText     =   "Loads a calibration setting."
         Top             =   4320
         Width           =   1410
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2070
         TabIndex        =   4
         ToolTipText     =   "Saves calibration setting."
         Top             =   4320
         Width           =   1410
      End
      Begin VB.TextBox YCentre 
         Enabled         =   0   'False
         Height          =   330
         Left            =   3120
         TabIndex        =   3
         Text            =   "0"
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox XCentre 
         Enabled         =   0   'False
         Height          =   330
         Left            =   3120
         TabIndex        =   2
         Text            =   "0"
         Top             =   1680
         Width           =   615
      End
      Begin MSComctlLib.Slider TFactor 
         Height          =   630
         Left            =   1800
         TabIndex        =   6
         Top             =   870
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1111
         _Version        =   393216
         BorderStyle     =   1
         Enabled         =   0   'False
         LargeChange     =   1
         Max             =   300
         SelStart        =   14
         TickStyle       =   2
         TickFrequency   =   10
         Value           =   14
      End
      Begin VB.Label lblVideoResolution 
         BackStyle       =   0  'Transparent
         Caption         =   "Video Resolution"
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   120
         TabIndex        =   19
         Top             =   3120
         Width           =   1605
      End
      Begin VB.Image DisplayFishEye 
         Height          =   480
         Left            =   3360
         Tag             =   "0"
         Top             =   3600
         Width           =   480
      End
      Begin VB.Label lblDisplayFisheye 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Show settings in video image"
         ForeColor       =   &H00004000&
         Height          =   420
         Left            =   120
         TabIndex        =   17
         Top             =   3720
         Width           =   2925
      End
      Begin VB.Label lblOriginalResolutionlbl 
         BackStyle       =   0  'Transparent
         Caption         =   "FEC Resolution"
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   120
         TabIndex        =   15
         Top             =   2640
         Width           =   1605
      End
      Begin VB.Label lblParameter 
         BackStyle       =   0  'Transparent
         Caption         =   "Parameter"
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   135
         TabIndex        =   9
         Top             =   1080
         Width           =   1590
      End
      Begin VB.Label adjy_lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Camera Center Y"
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   135
         TabIndex        =   8
         Top             =   2220
         Width           =   1605
      End
      Begin VB.Label adjx_lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Camera Center X"
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   135
         TabIndex        =   7
         Top             =   1740
         Width           =   1695
      End
      Begin VB.Label FishEyeON_lbl 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "OFF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   345
         TabIndex        =   1
         Top             =   375
         Width           =   2925
      End
      Begin VB.Image FishEyeON 
         Height          =   480
         Left            =   3285
         Tag             =   "1"
         Top             =   270
         Width           =   480
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   855
      Left            =   120
      TabIndex        =   21
      Top             =   5160
      Width           =   4095
   End
   Begin VB.Label CameraModelLabel 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Camera Model"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   6840
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "Fisheye"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub btnLoad_Click()
'****************************************************************************************
'Name    : btnLoad_Click
'Created : Sep 7 2004
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    : click event for Load FEC button
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler

'
'
    ClearLineProfilerV6.Dialog.Filter = "FishEye Calibration (*.fec)|*.fec||"
    ClearLineProfilerV6.Dialog.FileName = LocToSave & "*.fec"
    ClearLineProfilerV6.Dialog.ShowOpen 'PCN3185
    If ClearLineProfilerV6.Dialog.FileName <> LocToSave & "*.fec" Then
        Call FisheyeFunctions.FecLoadInformation("", ClearLineProfilerV6.Dialog.FileName)
    End If

Exit Sub
Err_Handler:
    Select Case Err
        Case 32755: Exit Sub 'PCN3951
        Case Else: MsgBox Err & "-F1:" & Error$
    End Select
End Sub

Private Sub btnSave_Click()
'****************************************************************************************
'Name    : btnSave_Click
'Created : Sep 7 2004
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    : click event for Save FEC button
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler

    Call FisheyeFunctions.SaveFEC

Exit Sub
Err_Handler:
MsgBox Err & "-F2:" & Error$
End Sub

Private Sub DisplayFishEye_Click()
'****************************************************************************************
'Name    : DisplayFishEye_Click
'Created : Sep 9 2004   'PCN3031
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    : click event for DisplayFishEye button object
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
    If FisheyeFunctions.FisheyeDisplayed = True Then
        FisheyeFunctions.FisheyeDisplayed = False
        Call INI_WriteBack(MyFile, "Fish_Displayed=", "False")
        Call SetCheckBoxTick(DisplayFishEye, False)
        Call FisheyeFunctions.LiveFishEyeOFF
        ClearLineScreen.RefreshVideoScreen
    Else
        FisheyeFunctions.FisheyeDisplayed = True
        Call INI_WriteBack(MyFile, "Fish_Displayed=", "True")
        Call SetCheckBoxTick(DisplayFishEye, True)
        Call FisheyeFunctions.LiveFishEyeON
        ClearLineScreen.RefreshVideoScreen
    End If
Exit Sub
Err_Handler:
MsgBox Err & "-F3:" & Error$
End Sub





Private Sub EditFecText_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Handler

    If KeyAscii = 13 Then
        If EditFecText.text = "CleanFlowPass" Then
            btnSave.Enabled = True
            YCentre.Enabled = True
            XCentre.Enabled = True
            TFactor.Enabled = True
            EditFecText.text = "Manual FEC unlocked"
            CameraModelLabel.Visible = True
            CameraModelText.Visible = True
            DoEvents
        Else
            btnSave.Enabled = False
            YCentre.Enabled = False
            XCentre.Enabled = False
            TFactor.Enabled = False
            EditFecText.text = "Incorect password!!"
            DoEvents
            Call Sleep(3000)
            EditFecText.text = "Manual FEC Setting password"
            EditFecText.Visible = False
            CameraModelLabel.Visible = False
            CameraModelText.Visible = False
        End If
    End If

Exit Sub
Err_Handler:
MsgBox Err & "-F4:" & Error$
End Sub




Private Sub Form_Load()
'****************************************************************************************
'Name    : Form_Load
'Created : Aug 27 2004      PCN3019
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    : Sets up the initial state of the profiler
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
Dim FEDistortion As String
Dim FE_X_Centre As String
Dim FE_Y_Centre As String
Me.Visible = False

ConvertLanguage Me, Language 'PCN3034


'   Set up the forms position
width = PipelineDetails.width
height = 8430
Top = 300
Left = PipelineDetails.Left
'ZOrder 0

'   Load the current fisheye settings

FECResolution.Locked = True
VideoResolution.Locked = True
'   txtResolution.ForeColor =

'PCN3860
DisplayFishEye.Enabled = IIf(mediatype = Video Or mediatype = "Live", True, False)

Call PopulateCameraDropDown(CameraDropdown) 'PCN3595

If PVDFileName <> "" And Dir(PVDFileName) <> "" And ConfigInfo.FishEyeFlag = False Then
    CameraDropdown.Enabled = False 'PCN3595
    Call FisheyeFunctions.FEOFF

    FishEyeON.Enabled = False
    Exit Sub 'This PVD does not have FishEye applied
ElseIf PVDFileName <> "" And Dir(PVDFileName) <> "" Then
    CameraDropdown.Enabled = False 'PCN3595
    CameraDropdown.text = ""
    Call FisheyeFunctions.FEON

    'vvvv PCN3083 ****************************
'    If CDbl(Right(ConfigInfo.PVDFileVersion, Len(ConfigInfo.PVDFileVersion) - 1)) < 6.1 Then
'        Fisheye.TFactor.value = ConfigInfo.FishEyeDistortion * 10
'        Call FisheyeFunctions.SetDistortion(ConfigInfo.FishEyeDistortion * 10)
'            ConfigInfo.FishEyeDistortion = ConfigInfo.FishEyeDistortion * 10
'    End If
    '^^^^ ************************************
    Call FisheyeFunctions.DisableFishEye
    FishEyeON.Enabled = False
Else
    CameraDropdown.Enabled = True
    FishEyeON.Enabled = False 'PCN3595

    If ConfigInfo.FishEyeFlag = True Then
        Call FisheyeFunctions.FEON
    Else
         Call FisheyeFunctions.FEOFF
    End If
End If

'vvvv PCN3031 **************************************
If FisheyeFunctions.FisheyeDisplayed = False Then
    Call SetCheckBoxTick(DisplayFishEye, False)
Else
    Call SetCheckBoxTick(DisplayFishEye, True)
End If
'^^^^ **********************************************

Fisheye.TFactor.value = ConfigInfo.FishEyeDistortion
Fisheye.XCentre.text = CStr(ConfigInfo.FishEyeCenterX)
Fisheye.YCentre.text = CStr(ConfigInfo.FishEyeCenterY)
Fisheye.FECResolution.text = CStr(ConfigInfo.FishEyeOriginalHeight) & "x" & CStr(ConfigInfo.FishEyeOriginalWidth)

Call getimagesize(height, width)
Fisheye.VideoResolution.text = CStr(height) & "x" & CStr(width)

'''    'PCN3073
'''    Call FisheyeFunctions.DisableFishEye


Exit Sub
Err_Handler:
MsgBox Err & "-F5:" & Error$
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call LiveFishEyeOFF
End Sub

Private Sub HiddenManualFecButton_Click()
On Error GoTo Err_Handler

    EditFecText.Visible = True

Exit Sub
Err_Handler:
MsgBox Err & "-F6:" & Error$
End Sub

Private Sub TFactor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'****************************************************************************************
'Name    : TransX_Click
'Created : Aug 24 2004
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    : Deactivates FishEye transfomation - created to conform to previous
'          code degsign
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler

    XCentre.Locked = True
    YCentre.Locked = True
    Call SetDistortion(TFactor.value)
    YCentre.Locked = False
    XCentre.Locked = False

Exit Sub
Err_Handler:
MsgBox Err & "-F7:" & Error$
End Sub



Private Sub XCentre_Validate(Cancel As Boolean)
'****************************************************************************************
'Name    : XCentre_Change
'Created : Aug 27 2004, PCN3019
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    : event called when the user changes the value in the XCEntre text box
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
Dim testvalue As Double

    testvalue = SafeCDbl(XCentre) 'PCN4161
    
    If testvalue < FisheyeFunctions.FishEyeCentreMin Or testvalue > FisheyeFunctions.FishEyeCentreMax Then
        Cancel = True
        Exit Sub
    End If
    
    YCentre.Locked = True
    TFactor.Enabled = False
    Call SetFishEyeCentre(SafeCDbl(XCentre.text), SafeCDbl(YCentre.text))
    TFactor.Enabled = True
    YCentre.Locked = False
    

Exit Sub
Err_Handler:
Select Case Err
    Case 13 'Type mismatch, cancel the change
        Cancel = True
    Case Else
        MsgBox Err & "-F8:" & Error$
        Cancel = True
End Select
End Sub
Private Sub YCentre_Validate(Cancel As Boolean)
'****************************************************************************************
'Name    : YCentre_Change
'Created : Aug 27 2004
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    : event called when the user changes the value in the YCEntre text box
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
Dim testvalue As Double

    testvalue = SafeCDbl(YCentre) 'PCN4161
    
    If testvalue < FisheyeFunctions.FishEyeCentreMin Or testvalue > FisheyeFunctions.FishEyeCentreMax Then
        Cancel = True
        Exit Sub
    End If
    
    XCentre.Locked = True
    TFactor.Enabled = False
    Call SetFishEyeCentre(SafeCDbl(XCentre.text), SafeCDbl(YCentre.text))
    TFactor.Enabled = True
    XCentre.Locked = False

Exit Sub
Err_Handler:
Select Case Err
    Case 13 'Type mismatch, cancel the change
        Cancel = True
    Case Else
        MsgBox Err & "-F9:" & Error$
        Cancel = True
End Select
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Drop down menu for camera select is below PCN3595 (21 Oct 2005, Antony)
'


Private Sub CameraDropdown_OnSelect()
On Error GoTo Error_handler
    Dim FishEyeFile As String

    If CameraDropdown.ItemSelected >= 0 Then
        FishEyeFile = TheFECFiles(CameraDropdown.ItemSelected)
        Call FisheyeFunctions.FecLoadInformation(App.Path & "\Fec Files\", FishEyeFile)
    End If

Exit Sub
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-F10:" & Error$, vbExclamation
    End Select
End Sub



