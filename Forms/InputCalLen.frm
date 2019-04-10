VERSION 5.00
Begin VB.Form InputCalLen 
   Caption         =   "Calibration"
   ClientHeight    =   675
   ClientLeft      =   5565
   ClientTop       =   4680
   ClientWidth     =   5190
   Icon            =   "InputCalLen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   675
   ScaleWidth      =   5190
   Begin VB.TextBox CalibrationLengthInput 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   4200
      TabIndex        =   1
      Top             =   180
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Enter calibration length:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   55
      TabIndex        =   0
      Top             =   210
      Width           =   4100
   End
End
Attribute VB_Name = "InputCalLen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub Form_Load() 'PCN1825
On Error GoTo Err_Handler
ConvertLanguage Me, Language 'PCN2111

If CalibrationTypeLength <> 0 Then InputCalLen.CalibrationLengthInput = CalibrationTypeLength 'PCN1825
Me.Show

Exit Sub
Err_Handler:
    MsgBox Err & " - " & error$
End Sub


Private Sub CalibrationLengthInput_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Handler

Dim AdjustmentValue As Double
Dim CurrentAdjustmentValue As Double
Dim CurrentFishScale As Double
Dim NewFishScale As Double
Dim currentheight As Long
Dim currentwidth As Long

Dim YScaleAdjustment As Double

If CLPScreenAction = "DrawHorCalibrationLine" Then
    If KeyAscii = 27 Then  'ESC
        Unload InputCalLen
    ElseIf KeyAscii = vbKeyReturn Then
        If CalibrationLengthInput = 0 Or CalLengthYScale_Global = 0 Then  'PCN1910 LS
            MsgBox DisplayMessage("Can't calibrate with a zero length."), vbExclamation 'PCN2111
        Else
            Call getimagesize(currentheight, currentwidth)
            Call getscalevalue(CurrentFishScale)
            Call hough_GetYFishScale(CurrentAdjustmentValue)
            YScaleAdjustment = CalibrationLengthInput / CalLengthYScale_Global
            YScaleAdjustment = (YScaleAdjustment - 1) + CurrentAdjustmentValue
            Call hough_SetYFishScale(YScaleAdjustment)
            Call setoriginalsize(ConfigInfo.MediaWidth, ConfigInfo.MediaHeight)
            Call calculatescale
            Call CreateFishEyeMask
            getscalevalue (NewFishScale)
            ConfigInfo.FishEyeRatio = NewFishScale
            Call INI_WriteBack(MyFile, "Fish_Ratio=", ConfigInfo.FishEyeRatio)
            CalLength_Global = NewFishScale / CurrentFishScale * CalLength_Global
            Call INI_WriteBack(MyFile, "CalibrationDistance=", CalibrationLengthInput)
            ConfigInfo.Ratio = CalLen_Global / CalLength_Global 'PCN3035 'PCN3640
            
            Call INI_WriteBack(MyFile, "CalibrationLineLength=", CalLength_Global)
            'Redraw the main scales
            Call DrawMainScale(ClearLineScreen.PVScreen) 'PCNGL300103 'PCN3691
            PVDrawScreenRatio = ConfigInfo.Ratio
            ConfigInfo.FishEyeHorDistortion = YScaleAdjustment 'PCN3687
            Call INI_WriteBack(MyFile, "Fish_DistortionHorizontal=", YScaleAdjustment) 'PCN3687
            
'        Call setoriginalsize(FEWidth, FEHeight)
'        Call calculatescale
'        Call CreateFishEyeMask
'        FEScale = getscalevalue
'        Call INI_WriteBack(MyFile, "Fish_Ratio=", FEScale)
            
            
            Unload InputCalLen
        End If
    End If
End If
    
If CLPScreenAction = "DrawCalibrationLine" Then
    If KeyAscii = 27 Then  'ESC
        Unload InputCalLen
    ElseIf KeyAscii = vbKeyReturn Then
        If CalibrationLengthInput = 0 Then   'PCN1910 LS
            MsgBox DisplayMessage("Can't calibrate with a zero length."), vbExclamation 'PCN2111
        Else
            CalLen_Global = CalibrationLengthInput
            Call INI_WriteBack(MyFile, "CalibrationDistance=", CalibrationLengthInput)
            ConfigInfo.Ratio = CalLen_Global / CalLength_Global 'PCN3035 'PCN3640
            Call INI_WriteBack(MyFile, "CalibrationLineLength=", CalLength_Global)
            
            'Redraw the main scales
            Call DrawMainScale(ClearLineScreen.PVScreen) 'PCNGL300103 'PCN3691
            PVDrawScreenRatio = ConfigInfo.Ratio
            
            Unload InputCalLen
        End If
    End If
End If
    
Exit Sub
Err_Handler:
    MsgBox Err & " - " & error$
    Resume
End Sub

