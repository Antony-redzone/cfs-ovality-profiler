VERSION 5.00
Begin VB.Form OptionsPage 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   9885
   ClientLeft      =   19995
   ClientTop       =   4755
   ClientWidth     =   7425
   ControlBox      =   0   'False
   Icon            =   "OptionsPage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9885
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   Begin VB.Frame AladdinHaspFrame 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Aladdin Hasp"
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
      Height          =   615
      Left            =   120
      TabIndex        =   21
      Top             =   6360
      Width           =   4065
      Begin VB.CheckBox HaspLockOption 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Check1"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   255
      End
      Begin VB.Label AladdinHaspFrameLbl 
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
         TabIndex        =   28
         Top             =   0
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   480
         Picture         =   "OptionsPage.frx":038A
         Stretch         =   -1  'True
         Top             =   175
         Width           =   615
      End
   End
   Begin ClearLineProfiler.CBS_DropDownBox FishEyeCameraDropdown 
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   3600
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   661
   End
   Begin VB.Frame FisheyeSettingsFrame 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Lens Correction Settings"
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
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   3360
      Width           =   4065
      Begin VB.Label FisheyeSettingsFrameLbl 
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
         TabIndex        =   25
         Top             =   0
         Visible         =   0   'False
         Width           =   45
      End
   End
   Begin VB.Frame CompanyDetailsFrame 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Company Details"
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
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4065
      Begin VB.CommandButton btnLoadLogo 
         BackColor       =   &H00E0E0E0&
         Height          =   615
         Left            =   1200
         Picture         =   "OptionsPage.frx":1E15
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1250
         Width           =   615
      End
      Begin VB.TextBox txtCompName 
         Height          =   315
         Left            =   1920
         TabIndex        =   0
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtPhoneNo 
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Top             =   735
         Width           =   2055
      End
      Begin VB.Label CompanyDetailsFrameLbl 
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
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Image CompanyLogo 
         BorderStyle     =   1  'Fixed Single
         Height          =   915
         Left            =   1920
         Stretch         =   -1  'True
         Top             =   1110
         Width           =   2055
      End
      Begin VB.Label lblCompanyName 
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblPhoneNo 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No."
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   120
         TabIndex        =   5
         Top             =   735
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Logo"
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   120
         TabIndex        =   4
         Top             =   1110
         Width           =   1740
      End
   End
   Begin VB.Frame VideoInputSettingsFrame 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Video Input Settings"
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
      Height          =   1425
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   4065
      Begin VB.ComboBox cboDeviceSelect 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Text            =   "None"
         Top             =   600
         Width           =   3735
      End
      Begin VB.CommandButton cmdDeviceSetting 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Input Select"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2400
         TabIndex        =   3
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label VideoInputSettingsFrameLbl 
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
         TabIndex        =   26
         Top             =   0
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Video Device Select"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label lblVideoSettings 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Video Input Select"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   2160
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame ThreeDRenderingSettingFrame 
      BackColor       =   &H00FFC0C0&
      Caption         =   "3D Rendering Setting"
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
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   5520
      Width           =   4065
      Begin VB.ComboBox cbo3DRender 
         Height          =   315
         ItemData        =   "OptionsPage.frx":3ADF
         Left            =   120
         List            =   "OptionsPage.frx":3AE1
         TabIndex        =   15
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label ThreeDRenderingSettingFrameLbl 
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
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   45
      End
   End
   Begin VB.Frame SystemSettingsFrame 
      BackColor       =   &H00FFC0C0&
      Caption         =   "System Settings"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   4065
      Begin VB.ComboBox cmbLanguage 
         Height          =   315
         ItemData        =   "OptionsPage.frx":3AE3
         Left            =   1920
         List            =   "OptionsPage.frx":3AE5
         TabIndex        =   18
         Top             =   600
         Width           =   2055
      End
      Begin VB.ComboBox cboUnits 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "OptionsPage.frx":3AE7
         Left            =   1920
         List            =   "OptionsPage.frx":3AF1
         TabIndex        =   17
         Top             =   255
         Width           =   2055
      End
      Begin VB.Label SystemSettingsFrameLbl 
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
         TabIndex        =   24
         Top             =   0
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Language"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Units_lbl 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Units"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   19
         Top             =   255
         Width           =   2580
      End
   End
   Begin VB.Frame GraphOptionsFrame 
      Caption         =   "Flat / XY Method"
      Height          =   1335
      Left            =   120
      TabIndex        =   29
      Top             =   6990
      Width           =   4095
      Begin VB.CheckBox SmoothCheck 
         BackColor       =   &H00ECCDB8&
         Caption         =   "Filter"
         CausesValidation=   0   'False
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
         Left            =   120
         TabIndex        =   32
         Top             =   960
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.OptionButton MedianDiameterOpt 
         BackColor       =   &H00ECCDB8&
         Caption         =   "Internal Diameter"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   31
         ToolTipText     =   "Uses 'Internal Diameter' from Pipeline Details"
         Top             =   600
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.OptionButton MedianDiameterOpt 
         BackColor       =   &H00ECCDB8&
         Caption         =   "Median Diameter (Deflection)"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   30
         ToolTipText     =   "Uses 'Median Diameter' per frame"
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Image FormBackground 
      Height          =   9135
      Left            =   0
      Picture         =   "OptionsPage.frx":3AFD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5025
   End
End
Attribute VB_Name = "OptionsPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim UnitsConverted As Boolean

''''Dim TheRefShapeFiles() As String  'PCN4171
''''Dim TheRefShapeFileLabels() As String  'PCN4171



Private Sub btnLoadLogo_Click()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : btnLoadLogo_Click
'Created : 7 September 2004, PCN3024
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Loads from file the Logo.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Call LoadLogo

Exit Sub
Err_Handler:
    MsgBox Err & "-O1:" & Error$
End Sub


Private Sub cbo3DRender_Click()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : cbo3DRender_Click()
'Created : 6th October 2003, PCN2266
'Updated : 6th October 2003, PCN2266
'Prg By  : Michelle Lindsay
'Param   : (None)
'Desc    : Allows the user to change the 3d rendering style of the computer between
'          software, hardware, and a mix of both.  Default is mixed.
'Usage   :
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
ThreeDRenderingStyle = cbo3DRender.ListIndex

Call INI_WriteBack(MyFile, "ThreeDRenderingStyle=", ThreeDRenderingStyle)
cbo3DRender.text = cbo3DRender.List(ThreeDRenderingStyle) 'PCN2266

'   MsgBox DisplayMessage("Please Press Apply to confirm Changes.") 'PCN2266 'PCN3024

Exit Sub
Err_Handler:
    MsgBox Err & "-O2:" & Error$
End Sub

Private Sub cbo3DRender_Validate(Cancel As Boolean)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : cbo3DRender_Validate
'Created : 13 September 2004, PCN3024
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Validate 3D Render and save to INI.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

ThreeDRenderingStyle = cbo3DRender.ListIndex

Call INI_WriteBack(MyFile, "ThreeDRenderingStyle=", ThreeDRenderingStyle)


Exit Sub
Err_Handler:
    MsgBox Err & "-O3:" & Error$


'Call INI_WriteBack(MyFile, "PaperSize=", cboPageSize)'Why was this here? (Geof, removed, 22 Sept 2004)

End Sub



Private Sub cboDeviceSelect_Change()
On Error GoTo Err_Handler

Exit Sub
Err_Handler:
    MsgBox Err & "-O4:" & Error$
    
End Sub

Private Sub cboDeviceSelect_Click()
On Error GoTo Err_Handler


Call INI_WriteBack(MyFile, "VideoCaptureDevice=", cboDeviceSelect.ListIndex)
VideoCaptureDevice = cboDeviceSelect.ListIndex

Call ControlsMain.ControlsDisplaySetup("Options") 'PCN4243
Exit Sub
Err_Handler:
    MsgBox Err & "-O5:" & Error$
End Sub

Private Sub cboDeviceSelect_DropDown()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : cboDeviceSelect_DropDown
'Created : 22 September 2004, PCN2395
'Updated :
'Prg By  : Antony van Iersel
'Param   :
'Desc    : Repopulate the combo box with latest capture devices
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim CurrentSelection As String

'Store the current setting
CurrentSelection = cboDeviceSelect.text
'Remove all items in the list
cboDeviceSelect.Clear
'restore current setting
cboDeviceSelect.text = CurrentSelection
'Regenerate the list
Call hough_getcapturedevices(cboDeviceSelect.hwnd)  'PCN2395 Multiple device select

Exit Sub
Err_Handler:
    MsgBox Err & "-O6:" & Error$
End Sub

Private Sub cboDeviceSelect_Validate(Cancel As Boolean)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : cboDeviceSelect_Validate
'Created : 22 September 2004, PCN2395
'Updated :
'Prg By  : Antony van Iersel
'Param   :
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Call INI_WriteBack(MyFile, "VideoCaptureDevice=", cboDeviceSelect.ListIndex)
VideoCaptureDevice = cboDeviceSelect.ListIndex

'If VideoCaptureDevice > 0 Then
'    ClearLineScreen.ControlToolbar.Buttons.Item(1).Enabled = True
'Else
'    ClearLineScreen.ControlToolbar.Buttons.Item(1).Enabled = False
'End If


Exit Sub
Err_Handler:
    MsgBox Err & "-O7:" & Error$
End Sub




Private Sub cmbLanguage_Click()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : cmbLanguage_Click()
'Created : 21st October 2003, PCN2301
'Updated : 21st October 2003, PCN2301
'Prg By  : Michelle Lindsay
'Param   : (None)
'Desc    : Allows the user to change the displayed language for the forms and buttons
'          as long as there is a corresponding text file for the selected language.
'Usage   :
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

'    MsgBox DisplayMessage("Please Press Apply to confirm Changes.") 'PCN3024
    
Exit Sub
Err_Handler:
    MsgBox Err & "-O8:" & Error$
End Sub

Private Sub cmbLanguage_Validate(Cancel As Boolean)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : cmbLanguage_Validate()
'Created : 7 September 2004, PCN3024
'Updated :
'Prg By  : Geoff Logan
'Param   : Cancel -
'Desc    : Validates the language.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

If Language <> cmbLanguage.text Then
    Call INI_WriteBack(MyFile, "Language=", cmbLanguage)
    'MsgBox DisplayMessage("Please close the ClearLine Profiler application and restart to apply the change."), vbExclamation
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Please close the ClearLine Profiler application and restart to apply the change."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
End If
    
    
Exit Sub
Err_Handler:
    MsgBox Err & "-O9:" & Error$
End Sub




Private Sub cmdDeviceSetting_Click()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'cmdVideoSettings_Click Sub  Antony van Iersel
'
'Revision history"
'   V0.0    Antony van Iersel,    16/10/03     Building initial framework
'           Antony van Iersel,    06/11/2003   Bringing code inline with C
'
'Description: PCN2330 Set the video input from capture card
'
'Purpose: To ensure user can preselect video input e.g. SVHS composite.
'         This ensures that the card will be on the right video input.
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

' PCN2430 (21 Nov 2003, Antony van Iersel) Michelle had allready created the condition,
' live string was needed in little case, "LIVE" is now "Live". That is all that needed change.
If mediatype <> "Live" Then                                                 ''''''''''''''''''
    'MsgBox DisplayMessage("Live video is not connected, device setting will not be changed") 'PCN2762
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Live video is not connected, device setting will not be changed"): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Exit Sub                                                                 '
End If                                                                       '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Call setdeviceinput

Exit Sub
Err_Handler:
    MsgBox Err & "-O10:" & Error$

End Sub


Private Sub DrawingScale_Validate(Cancel As Boolean)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : DrawingScale_Validate
'Created : 7 September 2004, PCN2970
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Call PrecisionVisionGraph.PVGraphReset
Call PrecisionVisionGraph.PVGraphResetFlat3D

Exit Sub
Err_Handler:
    MsgBox Err & "-O11:" & Error$
End Sub

Private Sub DrawingScaleX_Validate(Cancel As Boolean)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : DrawingScaleX_Validate
'Created : 12 August 2004, PCN2970
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Call PrecisionVisionGraph.PVGraphReset


Exit Sub
Err_Handler:
    MsgBox Err & "-O12:" & Error$
End Sub



Public Sub FishEyeCameraDropdown_OnSelect()
On Error GoTo Error_handler
    Dim FishEyeFile As String



    If FishEyeCameraDropdown.ItemSelected >= 0 Then
        FishEyeFile = TheFECFiles(FishEyeCameraDropdown.ItemSelected)
        Call FisheyeFunctions.FecLoadInformation(App.Path & "\Fec Files\", FishEyeFile)
        Call ClearLineScreen.TakeASnapShot 'PCN4453
    End If

Exit Sub
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-O13:" & Error$, vbExclamation
    End Select
End Sub

Private Sub FisheyeSettingsFrame_DblClick()
On Error GoTo Err_Handler

Load Fisheye
'Fisheye.Show

Exit Sub
Err_Handler:
    MsgBox Err & "-O14:" & Error$
End Sub


Private Sub Form_Activate()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : Form_Activate
'Created : 7 September 2004, PCN3024
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

'If a PVD has been loaded then disable the units control.
If PVDFileName = "" Then
    cboUnits.Enabled = True
Else
    cboUnits.Enabled = False
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-O15:" & Error$
End Sub


Private Sub cboUnits_Click()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'cboUnits_Click Sub  Louise Shrimpton louises@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    31/01/03     Building initial framework
'
'Description:
'
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    'MsgBox DisplayMessage("WARNING: If Units are changed, calibration will be reset to default.") 'PCN2111
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("WARNING: If Units are changed, calibration will be reset to default."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0

    If Left(cboUnits.text, 2) = MeasurementUnits Then
        Exit Sub
    ElseIf (cboUnits.text <> "mm" And cboUnits.text <> "in") Or PVDFileName <> "" Then 'PCN3024
        'MsgBox DisplayMessage("Invalid selection."), vbExclamation 'PCN2111
        ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Invalid selection."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
        'Cancel = True
        Exit Sub
    End If
    
    
    If Left(cboUnits.text, 2) = "mm" And MeasurementUnits = "in" And UnitsConverted = False Then
        CalLen_Global = CalLen_Global * 25.4 '(16 / 0.63)
        UnitsConverted = True
    ElseIf Left(cboUnits.text, 2) = "in" And MeasurementUnits = "mm" Then
        CalLen_Global = CalLen_Global * (1 / 25.4) '0.63 / 16)
        UnitsConverted = True
    End If
    
    
    'Re-initialise the application 'To be done
    Call DrawMainScale(ClearLineScreen.PVScreen) 'PCNGL300103 'PCN3691

    
    'vvvv PCN3024 *************************************
     MeasurementUnits = Left(cboUnits, 2)
     ConfigInfo.Units = Left(MeasurementUnits, 2)
     Call INI_WriteBack(MyFile, "MeasurementUnits=", Left(MeasurementUnits, 2))
     'Call INI_WriteBack(MyFile, "CalibrationDistance=", ClearLineScreen.CalLen)
     Call INI_WriteBack(MyFile, "CalibrationDistance=", CalLen_Global)
     UnitsConverted = False
    '^^^^ *********************************************
    
        
    Unload PipelineDetails
    Load PipelineDetails
    
    ClearLineScreen.Y_Units.Caption = MeasurementUnits
    
If MeasurementUnits = "mm" Then
    ControlsScreen.ControlsReports(5).ToolTipText = ControlsScreen.Label1kReport.Caption
Else
    ControlsScreen.ControlsReports(5).ToolTipText = ControlsScreen.Label1mlReport.Caption
End If
    



Exit Sub
Err_Handler:
    MsgBox Err & "-O16:" & Error$
End Sub

Private Sub cboUnits_Validate(Cancel As Boolean) 'PCNGL300103
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'cboUnits_Validate Sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    30/01/03     Building initial framework
'
'Description:
'       When the units are changed from/to inches to/from mm, then the complete program
'       needs to be re-initialised.
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
'Dim Resp As Variant

If Left(cboUnits.text, 2) = MeasurementUnits Then
    Exit Sub
ElseIf (cboUnits.text <> "mm" And cboUnits.text <> "in") Or PVDFileName <> "" Then 'PCN3024
    'MsgBox DisplayMessage("Invalid selection."), vbExclamation 'PCN2111
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Invalid selection."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Cancel = True
    Exit Sub
End If


If Left(cboUnits.text, 2) = "mm" And MeasurementUnits = "in" And UnitsConverted = False Then
    CalLen_Global = CalLen_Global * (16 / 0.63)
    UnitsConverted = True
ElseIf Left(cboUnits.text, 2) = "in" And MeasurementUnits = "mm" Then
    CalLen_Global = CalLen_Global * (0.63 / 16)
    UnitsConverted = True
End If


'Re-initialise the application 'To be done
Call DrawMainScale(ClearLineScreen.PVScreen) 'PCNGL300103 'PCN3691



'vvvv PCN3024 *************************************
 MeasurementUnits = Left(cboUnits, 2)
 ConfigInfo.Units = Left(MeasurementUnits, 2)
 Call INI_WriteBack(MyFile, "MeasurementUnits=", Left(MeasurementUnits, 2))
 'Call INI_WriteBack(MyFile, "CalibrationDistance=", ClearLineScreen.CalLen)
 Call INI_WriteBack(MyFile, "CalibrationDistance=", CalLen_Global)
 UnitsConverted = False
'^^^^ *********************************************

Unload PipelineDetails
Load PipelineDetails

Exit Sub
Err_Handler:
    MsgBox Err & "-O17:" & Error$

End Sub


Private Sub Form_Load()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : Form_Load
'Created : 12 November 2002,
'Updated : 18 November 2003, PCN2402 Tidy up of form layout for XP
'Prg By  : Geoff Logan
'Param   :
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim NoOfErrors As Integer
Dim INI_Value As String
Me.SmoothCheck.Caption = DisplayMessage("Filter")
Me.MedianDiameterOpt(0).Caption = DisplayMessage("Median Diameter (Deflection)")
Me.MedianDiameterOpt(1).Caption = DisplayMessage("Internal Diameter")

NoOfErrors = 0

'vvvv PCN4171 **********************************
Me.Left = PVPageLeft
Me.Top = PVPageTop
Me.width = PVPageWidth
Me.height = PVPageHeight
FormBackground.height = PVPageHeight
CompanyDetailsFrame.BackColor = RGB(172, 196, 231)
SystemSettingsFrame.BackColor = RGB(172, 196, 231)
VideoInputSettingsFrame.BackColor = RGB(172, 196, 231)
FisheyeSettingsFrame.BackColor = RGB(172, 196, 231)
ThreeDRenderingSettingFrame.BackColor = RGB(172, 196, 231)
AladdinHaspFrame.BackColor = RGB(172, 196, 231)
Me.GraphOptionsFrame.BackColor = RGB(172, 196, 231)
Me.SmoothCheck.BackColor = RGB(172, 196, 231)
Me.MedianDiameterOpt(0).BackColor = RGB(172, 196, 231)
Me.MedianDiameterOpt(1).BackColor = RGB(172, 196, 231)


'^^^^ ******************************************

ConvertLanguage Me, Language 'PCN2111

'If (IdentifyOperatingSystem = "Windows XP") Then 'ML120203
'    Frame1.Left = 0
'    Frame2.Left = 0
'End If
    

txtCompName = CompanyName
txtPhoneNo = PhoneNo
cboUnits.text = MeasurementUnits
cbo3DRender.Font.Charset = LanguageCharset
cbo3DRender.AddItem (DisplayMessage("Auto"))
cbo3DRender.AddItem (DisplayMessage("Software"))
cbo3DRender.text = cbo3DRender.List(ThreeDRenderingStyle) 'PCN2266



'cboCaptureDevice.text = CaptureDevice 'PCN2289 'PCN2398

'vvvv *********** Load a valid Company Logo *****************************'PCNGL150103
If CompanyLogoPath <> "" Then 'PCN4269
    If Dir(CompanyLogoPath) <> "" Then
        CompanyLogo.Picture = LoadPicture(CompanyLogoPath)
    End If
End If
'^^^^ *******************************************************************

'PCN2111 ------------------------------------------------v
Dim strFile As String
Dim strLine As String
Dim intPos As String

Dim FileNo As Integer
FileNo = FreeFile

strFile = Dir(ReadOnlyAppPath & "Language\Languages.txt") 'PCN2123
If strFile <> "" Then
    Open ReadOnlyAppPath & "Language\Languages.txt" For Input As #FileNo 'PCN2123
    Line Input #FileNo, strLine
    While Not EOF(FileNo)
        Line Input #FileNo, strLine
        If strLine <> "" Then
            intPos = InStr(strLine, ",")
            strFile = Left(strLine, intPos - 1)
            cmbLanguage.AddItem strFile
        End If
    Wend
    Close #FileNo
End If
cmbLanguage.text = Language
'--------------------------------------------------------^


Call hough_getcapturedevices(cboDeviceSelect.hwnd)  'PCN2395 Multiple device select
If VideoCaptureDevice > cboDeviceSelect.ListCount - 1 Then      '
    cboDeviceSelect.ListIndex = cboDeviceSelect.ListCount - 1   '
    VideoCaptureDevice = cboDeviceSelect.ListCount - 1          '
    Call INI_WriteBack(MyFile, "VideoCaptureDevice=", cboDeviceSelect.ListIndex) '
Else                                                            '
    cboDeviceSelect.ListIndex = VideoCaptureDevice              '
End If                                                          '
If mediatype = "Live" Then                                      '
    cboDeviceSelect.Enabled = False                             '
Else                                                            '
    cboDeviceSelect.Enabled = True                              '
End If                                                          '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'vvvv PCN4171 *************************************
'Fisheye camera selection moved from the Fisheye form
Call PopulateCameraDropDown(FishEyeCameraDropdown) 'PCN3595
If PVDFileName <> "" And Dir(PVDFileName) <> "" And ConfigInfo.FishEyeFlag = False Then
    FishEyeCameraDropdown.Enabled = False 'PCN3595
    Exit Sub 'This PVD does not have FishEye applied
ElseIf PVDFileName <> "" And Dir(PVDFileName) <> "" Then
    FishEyeCameraDropdown.Enabled = False 'PCN3595
    FishEyeCameraDropdown.text = ""
    Call FisheyeFunctions.FEON
    Call FisheyeFunctions.DisableFishEye
Else
    FishEyeCameraDropdown.Enabled = True
    If ConfigInfo.FishEyeFlag = True Then
        Call FisheyeFunctions.FEON
    Else
         Call FisheyeFunctions.FEOFF
    End If
End If
'^^^^ **********************************************

''''Call PopulateRefShapeDropDown(RefShape_DropDownBox) 'PCN4171 'PCN4344
''''

'vvvv PCN3865 **********************************************
Call GetINI_ParameterInfoOnly(MyFile, "HASPLock=", INI_Value)
INI_Value = UCase(INI_Value)
HaspLockOption = IIf(INI_Value = "FALSE", 0, 1)
'^^^^ ******************************************************

If LanguageCharset <> 0 Then
    Me.CompanyDetailsFrameLbl.Font.Charset = LanguageCharset: Me.CompanyDetailsFrameLbl.BackColor = RGB(172, 196, 231): Me.CompanyDetailsFrameLbl.Caption = DisplayMessage(CompanyDetailsFrame.Caption): Me.CompanyDetailsFrameLbl.Visible = True
    Me.SystemSettingsFrameLbl.Font.Charset = LanguageCharset: Me.SystemSettingsFrameLbl.BackColor = RGB(172, 196, 231): Me.SystemSettingsFrameLbl.Caption = DisplayMessage(SystemSettingsFrame.Caption): Me.SystemSettingsFrameLbl.Visible = True
    Me.FisheyeSettingsFrameLbl.Font.Charset = LanguageCharset: Me.FisheyeSettingsFrameLbl.BackColor = RGB(172, 196, 231): Me.FisheyeSettingsFrameLbl.Caption = DisplayMessage(FisheyeSettingsFrame.Caption): Me.FisheyeSettingsFrameLbl.Visible = True
    Me.VideoInputSettingsFrameLbl.Font.Charset = LanguageCharset: Me.VideoInputSettingsFrameLbl.BackColor = RGB(172, 196, 231): Me.VideoInputSettingsFrameLbl.Caption = DisplayMessage(VideoInputSettingsFrame.Caption): Me.VideoInputSettingsFrameLbl.Visible = True
    Me.ThreeDRenderingSettingFrameLbl.Font.Charset = LanguageCharset: Me.ThreeDRenderingSettingFrameLbl.BackColor = RGB(172, 196, 231): Me.ThreeDRenderingSettingFrameLbl.Caption = DisplayMessage(ThreeDRenderingSettingFrame.Caption): Me.ThreeDRenderingSettingFrameLbl.Visible = True
    Me.AladdinHaspFrameLbl.Font.Charset = LanguageCharset: Me.AladdinHaspFrameLbl.BackColor = RGB(172, 196, 231): Me.AladdinHaspFrameLbl.Caption = DisplayMessage(AladdinHaspFrame.Caption): Me.AladdinHaspFrameLbl.Visible = True
End If
Dim FlatTypeString As String

Call GetINI_ParameterInfoOnly(MyFile, "FlatType=", FlatTypeString)

If FlatTypeString = "Flat" Then
    Me.MedianDiameterOpt(1).value = True
ElseIf FlatTypeString = "Deflection" Then
    Me.MedianDiameterOpt(0).value = True
Else
    Me.MedianDiameterOpt(1).value = True
End If

    If MedianDiameterOpt(0).value = True Then MedianFlat = True: Call INI_WriteBack(MyFile, "FlatType=", "Deflection")
    If MedianDiameterOpt(1).value = True Then MedianFlat = False: Call INI_WriteBack(MyFile, "FlatType=", "Flat")
    Call PrecisionVisionGraph.UpdateGraphTitles

Exit Sub
CloseAllFiles:
    If NoOfErrors < 2 Then
        Close #FileNo
    End If
    
    

Exit Sub
Err_Handler:
    MsgBox Err & "-O18:" & Error$
    NoOfErrors = NoOfErrors + 1
    GoTo CloseAllFiles
End Sub




Private Sub HaspLockOption_Click()
On Error GoTo Err_Handler

    If HaspLockOption.value = 1 Then
        Call INI_WriteBack(MyFile, "HASPLock=", "true")
    Else
        Call INI_WriteBack(MyFile, "HASPLock=", "false")
    End If
     
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-O19:" & Error$
    End Select
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' PCN4974
' Antony van Iersel
' 4 August 2010
' Median Diamter Checkbox up, when changed redisplay the Flat approperiatly
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub MedianDiameterOpt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    
    If MedianDiameterOpt(0).value = True Then MedianFlat = True: PipelineDetails.RefShape_DropDownBox.Enabled = False: Call INI_WriteBack(MyFile, "FlatType=", "Deflection")
    If MedianDiameterOpt(1).value = True Then MedianFlat = False: PipelineDetails.RefShape_DropDownBox.Enabled = True: Call INI_WriteBack(MyFile, "FlatType=", "Flat")
        
    Call ScreenDrawing.PVFlat3DCalcCPP(1, PVDataNoOfLines)
    Call PrecisionVisionGraph.UpdateGraphTitles
    PrecisionVisionGraph.FlatTM2.Left = 1680
    If Language <> "English" Or MedianFlat Then 'PCN4974, added median flat
        PrecisionVisionGraph.FlatTM2.Left = 2130
    End If
    Call PVGraphsKeyForm.DisplayPVGraphsKey
    Call PrecisionVisionGraph.PopulatePVGraphSelectCaptions 'PCN5185
    Call ScreenDrawing.ClearAllGraphsAndRuler
    Call ScreenDrawing.DrawPVGraphs 'PCN3373
    Call ScreenDrawing.DrawPVXScale
    Call PVGraphsKeyForm.PVGraphsKeyUpdate 'PCN5186
    
    Exit Sub

Err_Handler:
    MsgBox Err & "-PD25:" & Error$
End Sub

Private Sub SmoothCheck_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.SmoothCheck.value = vbChecked Then
        ReDim GraphInfoContainer(PVOvalitySmooth).DataSingle(PVDataNoOfLines) 'PCN9999
        ReDim GraphInfoContainer(PVXDiameterSmooth).DataSingle(PVDataNoOfLines)
        ReDim GraphInfoContainer(PVYDiameterSmooth).DataSingle(PVDataNoOfLines)
        ReDim GraphInfoContainer(PVMedianDiameter).DataSingle(PVDataNoOfLines)
        ReDim GraphInfoContainer(PVMaxDiameterSmooth).DataSingle(PVDataNoOfLines)
        ReDim GraphInfoContainer(PVMinDiameterSmooth).DataSingle(PVDataNoOfLines)
        ReDim GraphInfoContainer(PVCapacitySmooth).DataSingle(PVDataNoOfLines)
'PCN6458         ReDim GraphInfoContainer(PVInclinationSmooth).DataSingle(PVDataNoOfLines) 'PCN6128
        
        
        Call CPPSmoothGraphs
         Call PVDeflectionCalcCPP   'PCN5186
        Call ScreenDrawing.ClearAllGraphsAndRuler
        Call ScreenDrawing.DrawPVGraphs 'PCN3373
        Call ScreenDrawing.DrawPVXScale
        
    Else
        ReDim GraphInfoContainer(PVOvalitySmooth).DataSingle(0) 'PCN9999
        ReDim GraphInfoContainer(PVXDiameterSmooth).DataSingle(0)
        ReDim GraphInfoContainer(PVYDiameterSmooth).DataSingle(0)
        ReDim GraphInfoContainer(PVMedianDiameterSmooth).DataSingle(0)
        ReDim GraphInfoContainer(PVMaxDiameterSmooth).DataSingle(0)
        ReDim GraphInfoContainer(PVMinDiameterSmooth).DataSingle(0)
        ReDim GraphInfoContainer(PVCapacitySmooth).DataSingle(0)
'PCN6458         ReDim GraphInfoContainer(PVInclinationSmooth).DataSingle(0) 'PCN6128
'PCN6458          Call PVDeflectionCalcCPP   'PCN5186
        Call ScreenDrawing.ClearAllGraphsAndRuler
        Call ScreenDrawing.DrawPVGraphs 'PCN3373
        Call ScreenDrawing.DrawPVXScale
    End If
    Call PVGraphsKeyForm.PVGraphsKeyUpdate 'PCN5186
   
    
    
    'PCN5186 need to recalculate deflection graph
End Sub

Private Sub txtCompName_Validate(Cancel As Boolean)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : txtCompName_Validate
'Created : 7 September 2004, PCN3024
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Validate Company Name and save to INI.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

CompanyName = txtCompName
Call INI_WriteBack(MyFile, "CompanyName=", CompanyName)

Exit Sub
Err_Handler:
    MsgBox Err & "-O20:" & Error$
End Sub

Private Sub txtPhoneNo_Validate(Cancel As Boolean)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : txtPhoneNo_Validate
'Created : 7 September 2004, PCN3024
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Validate PhoneNo and save to INI.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

PhoneNo = txtPhoneNo
Call INI_WriteBack(MyFile, "PhoneNo=", PhoneNo)


Exit Sub
Err_Handler:
    MsgBox Err & "-O21:" & Error$
End Sub


Function LoadLogo()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : LoadLogo
'Created : 7 September 2004, PCN3024
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Loads from file the Logo.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'On Error GoTo Err_Handler

'CommonDialog1.Filter = "JPEG (*.jpg;*.jpeg)|*.jpg;*.jpeg|Bitmap (*.bmp)|*.bmp"
'CommonDialog1.FileName = ""
'CommonDialog1.ShowOpen
'CompanyLogoPath = CommonDialog1.FileName
'Call INI_WriteBack(MyFile, "CompanyLogoPath=", CompanyLogoPath)
'CompanyLogo.Picture = LoadPicture(CommonDialog1.FileName)
'SavePicture CompanyLogo.Picture, LocToSave & "logo.bmp"
'If Dir(LocToSave & "logo.jpg") <> "" Then
'    Dim answer As Integer
'    'answer = MsgBox(DisplayMessage("Will you overwrite existing logo file?"), vbYesNo)  'PCN2111
'    ProfilerMessageBox.MsgBoxYesNo (DisplayMessage("Will you overwrite existing logo file?"))
'    answer = PMBAnswer
'
'    If answer = vbYes Then
'       PipelineDetails.JPGMake1.InputFile = LocToSave & "logo.bmp"
'       PipelineDetails.JPGMake1.OutputFile = LocToSave & "logo.jpg"
'       PipelineDetails.JPGMake1.Quality = 100
'       PipelineDetails.JPGMake1.Go
'    End If
'Else
'    PipelineDetails.JPGMake1.InputFile = LocToSave & "logo.bmp"
'    PipelineDetails.JPGMake1.OutputFile = LocToSave & "logo.jpg"
'    PipelineDetails.JPGMake1.Quality = 100
'    PipelineDetails.JPGMake1.Go
'End If
'
'
'Exit Function
'Err_Handler:
'    MsgBox Err & "-O22:" & Error$
End Function

