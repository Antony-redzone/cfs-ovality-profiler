Attribute VB_Name = "FisheyeFunctions"
Option Explicit


Private Declare Sub turnfisheyeon Lib "laserlib.dll" () ' Activates the FishEye
Private Declare Sub turnfisheyeoff Lib "laserlib.dll" () ' Deactivates the FishEye

Private Declare Sub setimagesize Lib "laserlib.dll" ()
Private Declare Sub settfactor Lib "laserlib.dll" (ByVal Factor As Double)
' Sets the Transformation factor for the fisheye
Private Declare Sub livefisheye Lib "laserlib.dll" (ByVal Status As Long)
Private Declare Sub setfecentre Lib "laserlib.dll" (ByVal xta As Long, ByVal yta As Long)
Private Declare Sub transformoneimage Lib "laserlib.dll" ()


Public Const FishEyeTFactorMax As Integer = 250
Public Const FishEyeTFactorMin As Integer = 0 'PCN3835, Now fish eye is always on, was 1 but 0 is now valid as off
Public Const FishEyeCentreMax As Integer = 99
Public Const FishEyeCentreMin As Integer = -99
Public Const FishEyeDistortionDivider As Integer = 400 'PCN3045
Public ViewerString As String

Public FisheyeDisplayed As Boolean ' PCN3031

Public TheFECFiles() As String 'PCN4171


Public Function InitializeFishEyeForPVD() 'PCN3005
'****************************************************************************************
'Name    : InitializeFishEyeForPVD
'Created : Sep 14 2004
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    : Initialises FishEye from the ConfigInfo data structure
'Usage   : Only called to initialize fisheye for pvd
'****************************************************************************************
On Error GoTo Err_Handler
Dim tempPVDConfigVersion As Double
tempPVDConfigVersion = CDbl(Right(ConfigInfo.PVDFileVersion, Len(ConfigInfo.PVDFileVersion) - 1))

    If ConfigInfo.FishEyeDistortion = 0 Then Call FEOFF ': Exit Function

    If ConfigInfo.FishEyeDistortion < FishEyeTFactorMin Or ConfigInfo.FishEyeDistortion > FishEyeTFactorMax Then GoTo Corrupt
    Call settfactor(1 + (ConfigInfo.FishEyeDistortion / FishEyeDistortionDivider))
    
    If ConfigInfo.FishEyeCenterX > FishEyeCentreMax Or ConfigInfo.FishEyeCenterX < FishEyeCentreMin Then GoTo Corrupt
    If ConfigInfo.FishEyeCenterY > FishEyeCentreMax Or ConfigInfo.FishEyeCenterY < FishEyeCentreMin Then GoTo Corrupt
    Call setfecentre(CLng(ConfigInfo.FishEyeCenterX), CLng(ConfigInfo.FishEyeCenterY))
    
    If ConfigInfo.FishEyeOriginalWidth = 0 Or ConfigInfo.FishEyeOriginalHeight = 0 Then
                ConfigInfo.FishEyeRatio = 228.8
                ConfigInfo.FishEyeOriginalWidth = 352
                ConfigInfo.FishEyeOriginalHeight = 263
    End If
    Call setoriginalsize(ConfigInfo.FishEyeOriginalWidth, ConfigInfo.FishEyeOriginalHeight)
    
    If tempPVDConfigVersion < 6.1 Then
        Call calculatescale
        Call getscalevalue(ConfigInfo.FishEyeRatio)
    Else
        Call setscalevalue(ConfigInfo.FishEyeRatio)
    End If
    If tempPVDConfigVersion >= 6.25 Then
        Call hough_SetYFishScale(ConfigInfo.FishEyeHorDistortion)
    Else
        Call hough_SetYFishScale(1)
    End If
    
    Call turnfisheyeon
    
    Call CreateFishEyeMask
    
    If isopen("Fisheye") Then
        Fisheye.TFactor.value = ConfigInfo.FishEyeDistortion
        Fisheye.XCentre.text = CStr(ConfigInfo.FishEyeCenterX)
        Fisheye.YCentre.text = CStr(ConfigInfo.FishEyeCenterY)
        Fisheye.FECResolution.text = CStr(ConfigInfo.FishEyeOriginalHeight) & "x" & CStr(ConfigInfo.FishEyeOriginalWidth)
        Dim width As Long
        Dim height As Long
        Call getimagesize(height, width)
        Fisheye.VideoResolution.text = CStr(height) & "x" & CStr(width)
    End If

    Call FEON

Exit Function
Corrupt:
MsgBox DisplayMessage("Can not read this file, loading default settings.")

Call FECLoadDefaultSettings

Exit Function
Err_Handler:
MsgBox Err & " - " & error$
End Function

Public Function InitializeFishEyeFromConfig() 'PCN3005
'****************************************************************************************
'Name    : InitializeFishEyeFromConfig
'Created : Aug 30 2004
'Updated : Sep 13 2004 save PVD setting to INI
'Prg By  : Eddie Jensen
'Param   :
'Desc    : Initialises FishEye from the ConfigInfo data structure
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler

    Call settfactor(1 + (ConfigInfo.FishEyeDistortion / FishEyeDistortionDivider))
    Call setfecentre(CLng(ConfigInfo.FishEyeCenterX), CLng(ConfigInfo.FishEyeCenterY))
    Call setoriginalsize(ConfigInfo.FishEyeOriginalWidth, ConfigInfo.FishEyeOriginalHeight)
    Call setscalevalue(ConfigInfo.FishEyeRatio)
    Call hough_SetYFishScale(ConfigInfo.FishEyeHorDistortion)
        
    
    Call turnfisheyeon
    Call CreateFishEyeMask
    
    If PVDFileName = "" Then 'PCN3863
    Call INI_WriteBack(MyFile, "Fish_OriginalWidth=", ConfigInfo.FishEyeOriginalWidth)
    Call INI_WriteBack(MyFile, "Fish_OriginalHeight=", ConfigInfo.FishEyeOriginalHeight)
    Call INI_WriteBack(MyFile, "Fish_CenterX=", ConfigInfo.FishEyeCenterX)
    Call INI_WriteBack(MyFile, "Fish_CenterY=", ConfigInfo.FishEyeCenterY)
    Call INI_WriteBack(MyFile, "Fish_Distortion=", ConfigInfo.FishEyeDistortion)
    Call INI_WriteBack(MyFile, "Fish_DistortionHorizontal=", ConfigInfo.FishEyeHorDistortion) 'PCN3687
    Call INI_WriteBack(MyFile, "Fish_Ratio=", ConfigInfo.FishEyeRatio)
    End If

    If isopen("Fisheye") Then
        Fisheye.TFactor.value = ConfigInfo.FishEyeDistortion
        Fisheye.XCentre.text = CStr(ConfigInfo.FishEyeCenterX)
        Fisheye.YCentre.text = CStr(ConfigInfo.FishEyeCenterY)
        Fisheye.FECResolution.text = CStr(ConfigInfo.FishEyeOriginalHeight) & "x" & CStr(ConfigInfo.FishEyeOriginalWidth)
        Dim width As Long
        Dim height As Long
        Call getimagesize(height, width)
        Fisheye.VideoResolution.text = CStr(height) & "x" & CStr(width)
    End If

Exit Function
Err_Handler:
MsgBox Err & " - " & error$
End Function
Public Function InitializeFishEyeFromINI() 'PCN3005
'****************************************************************************************
'Name    : InitializeFishEyeFromINI
'Created : Aug 30 2004
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    : Initialises FishEye from the INI
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
Dim FEDistortion As String
Dim FEDistortionHor As String
Dim FE_X_Centre As String
Dim FE_Y_Centre As String
Dim height As String
Dim width As String
Dim fe_scale As String
Dim currentwidth As Long
Dim currentheight As Long
Dim DisplayFish As String
Dim FecCameraModel As String 'PCN3595 (21 Oct 2005, Antony)
Dim FecFileName As String    'PCN3595 (21 Oct 2005, Antony)

' set the original
Call getimagesize(currentheight, currentwidth)

Call GetINI_ParameterInfoOnly(MyFile, "Fish_OriginalHeight=", height): ConfigInfo.FishEyeOriginalHeight = Val(height)
Call GetINI_ParameterInfoOnly(MyFile, "Fish_OriginalWidth=", width):   ConfigInfo.FishEyeOriginalWidth = Val(width)

Call setoriginalsize(CLng(width), CLng(height))

Call GetINI_ParameterInfoOnly(MyFile, "Fish_Ratio=", fe_scale): ConfigInfo.FishEyeRatio = Val(fe_scale)
Call GetINI_ParameterInfoOnly(MyFile, "Fish_Distortion=", FEDistortion): ConfigInfo.FishEyeDistortion = Val(FEDistortion)
Call GetINI_ParameterInfoOnly(MyFile, "Fish_DistortionHorizontal=", FEDistortionHor): ConfigInfo.FishEyeHorDistortion = Val(FEDistortionHor)
Call settfactor(1 + (ConfigInfo.FishEyeDistortion / FishEyeDistortionDivider))

Call GetINI_ParameterInfoOnly(MyFile, "Fish_CenterX=", FE_X_Centre): ConfigInfo.FishEyeCenterX = Val(FE_X_Centre)
Call GetINI_ParameterInfoOnly(MyFile, "Fish_CenterY=", FE_Y_Centre): ConfigInfo.FishEyeCenterY = Val(FE_Y_Centre)
Call setfecentre(Val(FE_X_Centre), Val(FE_Y_Centre))
Call setscalevalue(CDbl(fe_scale)) ' * height / currentheight))
Call hough_SetYFishScale(ConfigInfo.FishEyeHorDistortion)

Call turnfisheyeon

Call GetINI_ParameterInfoOnly(MyFile, "Fish_Displayed=", DisplayFish)
Call GetINI_ParameterInfoOnly(MyFile, "CameraModel=", FecCameraModel) 'PCN3595 (21 Oct 2005, Antony)
Call GetINI_ParameterInfoOnly(MyFile, "FecFileName=", FecFileName)    'PCN3595 (21 Oct 2005, Antony)
FisheyeDisplayed = IIf(DisplayFish = "True", True, False)


If isopen("Fisheye") Then
    Fisheye.FECResolution.text = height & "x"
    Fisheye.FECResolution.text = Fisheye.FECResolution.text & width
    Fisheye.TFactor.value = ConfigInfo.FishEyeDistortion
    Fisheye.XCentre.text = ConfigInfo.FishEyeCenterX
    Fisheye.YCentre.text = ConfigInfo.FishEyeCenterY
    Fisheye.VideoResolution.text = CStr(currentheight) & "x" & CStr(currentwidth)
    Fisheye.CameraDropdown.text = FecCameraModel 'PCN3595 if loaded show camera, (25 Oct 2005, Antony)
    Call SetCheckBoxTick(Fisheye.DisplayFishEye, FisheyeDisplayed)
End If

Exit Function
Err_Handler:
MsgBox Err & " - " & error$
End Function
Public Function EnableFishEye()
'****************************************************************************************
'Name    : EnableFishEye
'Created : Aug 30 2004
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    : Enables all the relevant fisheye objects on the fisheye form
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler

    If Not isopen("Fisheye") Then Exit Function
    Fisheye.lblParameter.Enabled = True
    Fisheye.adjx_lbl.Enabled = True
    Fisheye.adjy_lbl.Enabled = True
    Fisheye.TFactor.Enabled = True
    Fisheye.XCentre.Enabled = True
    Fisheye.YCentre.Enabled = True
    Fisheye.btnLoad.Enabled = True
    Fisheye.btnSave.Enabled = True

Exit Function
Err_Handler:
MsgBox Err & " - " & error$
End Function
Public Function DisableFishEye()
'****************************************************************************************
'Name    : EnableFishEye
'Created : Aug 30 2004
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    : Enables all the relevant fisheye objects on the fisheye form
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler

    If Not isopen("Fisheye") Then Exit Function
    Fisheye.lblParameter.Enabled = False
    Fisheye.adjx_lbl.Enabled = False
    Fisheye.adjy_lbl.Enabled = False
    Fisheye.TFactor.Enabled = False
    Fisheye.XCentre.Enabled = False
    Fisheye.YCentre.Enabled = False
    Fisheye.btnLoad.Enabled = False
    Fisheye.btnSave.Enabled = False
    
Exit Function
Err_Handler:
MsgBox Err & " - " & error$
End Function
Public Function FEON() 'PCN3005
'****************************************************************************************
'Name    : FEON_Click
'Created : Aug 24 2004
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    : Activates FishEye transfomation - created to conform to previous
'          code degsign
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
    
'    If ConfigInfo.FishEyeDistortion = 0 Then FEOFF: Exit Function 'PCN3595 even thou fish eye is always on if
'                                                                  'distortion in 0 then it is displayed as off
    
    Screen.MousePointer = vbHourglass
    'Call EnableFishEye 'PCN3595 just makes the slider and centre text boxes enabled, but this is now locked
    
    If ConfigInfo.FishEyeDistortion = 0 Then 'CPN3595
        Call FEOFF
    Else
        Call turnfisheyeon
    End If
    
    If isopen("Fisheye") Then
        Call SetCheckBoxTick(Fisheye.FishEyeON, True)
        Fisheye.FishEyeON_lbl.Caption = DisplayMessage("ON")
        Fisheye.FishEyeON_lbl.ForeColor = &HC000&
    End If
    
    ConfigInfo.FishEyeFlag = True
    Screen.MousePointer = vbDefault
    
Exit Function
Err_Handler:
MsgBox Err & " - " & error$
End Function

Public Function FEOFF() 'PCN3005
'****************************************************************************************
'Name    : FEOFF_Click
'Created : Aug 24 2004
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    : Deactivates FishEye transfomation - created to conform to previous
'          code degsign
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler

    Call DisableFishEye
    Call turnfisheyeoff
    
    If isopen("Fisheye") Then
        Call SetCheckBoxTick(Fisheye.FishEyeON, False)
'        Fisheye.FishEyeON_lbl.Caption = DisplayMessage("OFF")  'PCN3593
'        Fisheye.FishEyeON_lbl.ForeColor = &HFF&                'PCN3595
    End If

    ConfigInfo.FishEyeFlag = False

Exit Function
Err_Handler:
MsgBox Err & " - " & error$
End Function
Public Function SetFishEyeCentre(X As Long, Y As Long)
'****************************************************************************************
'Name    : XCentre_Change
'Created : Aug 27 2004
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    : event called when the user changes the value in either fisheye centre text box
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
Dim imageHeight As Long
Dim imageWidth As Long
    
    Call ClearLineScreen.ProfilerPause
    If ConfigInfo.FishEyeCenterX = X And ConfigInfo.FishEyeCenterY = Y Then Exit Function
    
    Screen.MousePointer = vbHourglass
    Call getimagesize(imageHeight, imageWidth)
    Fisheye.VideoResolution.text = CStr(imageHeight) & "x" & CStr(imageWidth)
    ConfigInfo.FishEyeOriginalWidth = imageWidth
    ConfigInfo.FishEyeOriginalHeight = imageHeight
    Call INI_WriteBack(MyFile, "Fish_OriginalWidth=", imageWidth)
    Call INI_WriteBack(MyFile, "Fish_OriginalHeight=", imageHeight)
    Fisheye.FECResolution.text = ConfigInfo.FishEyeOriginalHeight & "x" & ConfigInfo.FishEyeOriginalWidth
    Call setimagesize
    Call setfecentre(X, Y)
    Call settfactor(1 + (ConfigInfo.FishEyeDistortion / FishEyeDistortionDivider))
    Call calculatescale
    Call getscalevalue(ConfigInfo.FishEyeRatio)
    ConfigInfo.FishEyeCenterX = X
    Call INI_WriteBack(MyFile, "Fish_CenterX=", X)
    ConfigInfo.FishEyeCenterY = Y
    Call INI_WriteBack(MyFile, "Fish_CenterY=", Y)
    Call getscalevalue(ConfigInfo.FishEyeRatio)
    Call INI_WriteBack(MyFile, "Fish_Ratio=", ConfigInfo.FishEyeRatio)
    
    Call CreateFishEyeMask

'   If FisheyeDisplayed = False Then
'        LiveFishEyeON
'        Call ClearLineScreen.RefreshVideoScreen
'        Sleep (300)
'        LiveFishEyeOFF
'    Else
'        Call ClearLineScreen.RefreshVideoScreen
'        Sleep (300)
'    End If
    Screen.MousePointer = vbDefault

Exit Function
Err_Handler:
MsgBox Err & " - " & error$
End Function
Public Function SetDistortion(Distortion As Integer) 'PCN3005
'****************************************************************************************
'Name    : SetDistortion
'Created : Aug 30 2004
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    : universal function to change the distortion
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
Dim imageHeight As Long
Dim imageWidth As Long

    If Distortion = ConfigInfo.FishEyeDistortion Then Exit Function
    
    Call ClearLineScreen.ProfilerPause
    
    Screen.MousePointer = vbHourglass
    Call getimagesize(imageHeight, imageWidth)
    Fisheye.VideoResolution.text = CStr(imageHeight) & "x" & CStr(imageWidth)
    ConfigInfo.FishEyeOriginalWidth = imageWidth
    ConfigInfo.FishEyeOriginalHeight = imageHeight
    Call INI_WriteBack(MyFile, "Fish_OriginalWidth=", imageWidth)
    Call INI_WriteBack(MyFile, "Fish_OriginalHeight=", imageHeight)
    Fisheye.FECResolution.text = ConfigInfo.FishEyeOriginalHeight & "x" & ConfigInfo.FishEyeOriginalWidth
    Call setimagesize
    Call setfecentre(ConfigInfo.FishEyeCenterX, ConfigInfo.FishEyeCenterY)
    Call settfactor(1 + (Distortion / FishEyeDistortionDivider))
    Call calculatescale
    ConfigInfo.FishEyeDistortion = Distortion
    Call INI_WriteBack(MyFile, "Fish_Distortion=", ConfigInfo.FishEyeDistortion)
    Call getscalevalue(ConfigInfo.FishEyeRatio)
    Call INI_WriteBack(MyFile, "Fish_Ratio=", ConfigInfo.FishEyeRatio)

    Call CreateFishEyeMask
    
   
    Screen.MousePointer = vbDefault
Exit Function
Err_Handler:
MsgBox Err & " - " & error$
End Function

Public Function SaveFEC() 'PCN3005
'****************************************************************************************
'Name    : SaveFEC
'Created : Sep 7 2004
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    : Saves the FEC settings to file
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler

    ClearLineProfilerV6.Dialog.Filter = "FishEye Calibration (*.fec)|*.fec||"
    ClearLineProfilerV6.Dialog.FileName = LocToSave & "*.fec"
    ClearLineProfilerV6.Dialog.ShowSave
    If ClearLineProfilerV6.Dialog.FileName <> LocToSave & "*.fec" Then
        Open ClearLineProfilerV6.Dialog.FileName For Output As #100
        Print #100, "[Revision]"
        Print #100, "FECRevision=2.1"
        Print #100, "[FishEyeParameters]"
        Print #100, "Fish_Distortion=" & ConfigInfo.FishEyeDistortion
        Print #100, "Fish_CenterX=" & ConfigInfo.FishEyeCenterX
        Print #100, "Fish_CenterY=" & ConfigInfo.FishEyeCenterY
        Print #100, "Fish_OriginalWidth=" & ConfigInfo.FishEyeOriginalWidth
        Print #100, "Fish_OriginalHeight=" & ConfigInfo.FishEyeOriginalHeight
        Print #100, "Fish_Ratio=" & ConfigInfo.FishEyeRatio
        Print #100, "[Camera]"
        Print #100, "CameraModel=" & Fisheye.CameraModelText.text
        Print #100, "\\"
    End If
    Close #100

Exit Function
Err_Handler:
    Select Case Err
        Case 3951: Exit Function 'PCN3951
        Case Else: MsgBox Err & " - " & error$
    End Select
End Function
Sub FECLoadDefaultSettings()
'****************************************************************************************
'Name    : FECLoadDefaultSettings
'Created : Sep 8 2004
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    : Loads the default FEC settings
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler

        Call INI_WriteBack(MyFile, "Fish_Distortion=", 0)
        Call INI_WriteBack(MyFile, "Fish_DistortionHorizontal=", 1) 'PCN3687
        Call INI_WriteBack(MyFile, "Fish_CenterX=", 0)
        Call INI_WriteBack(MyFile, "Fish_CenterY=", 0)
        Call INI_WriteBack(MyFile, "Fish_OriginalWidth=", 704)
        Call INI_WriteBack(MyFile, "Fish_OriginalHeight=", 576)
        Call INI_WriteBack(MyFile, "Fish_Ratio=", 364)
        Call INI_WriteBack(MyFile, "FecFileName=", "")
        Call INI_WriteBack(MyFile, "CameraModel=", ViewerString) 'PCN3595 (21 Oct 2005, Antony) added CameraModel to FEC File

Exit Sub
Err_Handler:
MsgBox Err & " - " & error$
End Sub

Public Sub FecLoadInformation(FECPathName As String, ByVal FecFileName As String) 'PCN3595 (21 Oct 2005, Antony) split path and filename
'****************************************************************************************
'Name    : LoadFecInformation
'Created : Sep 8 2004
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    : Loads the FEC settings from file
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
Dim ConfigLine As String
Dim X As Long
Dim Y As Long
Dim Parameter As String
Dim value As String
Dim PathName As String
Dim FileName As String
Dim HaltApp As Boolean
Dim FECRev As String
Dim FECFileNo As Integer
Dim SectionHead As String
Dim SectionDetail As String

Dim FEDistortion As Integer
Dim FECentreX As Integer
Dim FECentreY As Integer
Dim FEWidth As Integer
Dim FEHeight As Integer
Dim FEScale As Double
Dim FecCameraModel As String 'PCN3595 (21 Oct 2005, Antony)
Dim FECFileNameAndPath As String


Call CLPProgressBar.ProgressBarInitialise(DisplayMessage(""))
DoEvents

FECFileNameAndPath = FECPathName & FecFileName



FECRev = ""
FECFileNo = 9

X = 1

   
Config_LineCnt = 0

Open FECFileNameAndPath For Input As #FECFileNo
Do While Not EOF(FECFileNo)
 Line Input #FECFileNo, ConfigLine
 Config_LineCnt = Config_LineCnt + 1
Loop
Close #FECFileNo

ReDim ConfigArray(Config_LineCnt)
  
Open FECFileNameAndPath For Input As #FECFileNo
Do While Not EOF(FECFileNo)
 Line Input #FECFileNo, ConfigLine
 ConfigArray(X) = ConfigLine
 X = X + 1
Loop
Close #FECFileNo

' Run Through Entire Array and Validate Paths and Files

For Y = 1 To Config_LineCnt
  SectionHead = ConfigArray(Y)
  If SectionHead = "[Revision]" Then
      Y = Y + 1
      SectionDetail = ConfigArray(Y)
      Do Until SectionDetail = "" Or Left(SectionDetail, 1) = "[" Or SectionDetail = "\\" 'End marker of data
        If SectionDetail <> "" Then Call GetParam(SectionDetail, Parameter, value, PathName, FileName)
        Select Case Parameter
            Case "FECRevision="
                FECRev = value
            Case Else
        End Select
        Y = Y + 1
        SectionDetail = ConfigArray(Y)
      Loop
  End If
  SectionHead = ConfigArray(Y)
  If SectionHead = "[FishEyeParameters]" Then
     Y = Y + 1
     SectionDetail = ConfigArray(Y)
     Do Until SectionDetail = "" Or Left(SectionDetail, 1) = "[" Or SectionDetail = "\\" 'End marker of data
      If SectionDetail <> "" Then Call GetParam(SectionDetail, Parameter, value, PathName, FileName)
      Select Case Parameter
        Case "Fish_Distortion=":            FEDistortion = Val(value)
        Case "Fish_CenterX=":               FECentreX = Val(value)
        Case "Fish_CenterY=":               FECentreY = Val(value)
        Case "Fish_OriginalWidth=":         FEWidth = ConfigInfo.MediaWidth
        Case "Fish_OriginalHeight=":        FEHeight = ConfigInfo.MediaHeight
        Case "Fish_Ratio=":                 FEScale = Val(value)
      End Select
      
      Y = Y + 1
      SectionDetail = ConfigArray(Y)
    Loop
  End If
  SectionHead = ConfigArray(Y)

Call CLPProgressBar.ProgressBarPosition(0.33)
DoEvents

'''''''''''''' PCN3595 (21 Oct 2005, Antony) Add CameraModel ''''
'
ConfigInfo.FishEyeHorDistortion = 1
Call hough_SetYFishScale(1)

'
  If SectionHead = "[Camera]" Then                              '
    Y = Y + 1                                                   '
    SectionDetail = ConfigArray(Y)                              '
     Do Until SectionDetail = "" Or Left(SectionDetail, 1) = "[" Or SectionDetail = "\\" 'End marker of data
      If SectionDetail <> "" Then Call GetParam(SectionDetail, Parameter, value, PathName, FileName)
      Select Case Parameter                                     '
        Case "CameraModel="                                     '
            FecCameraModel = value                              '
      End Select                                                '
      Y = Y + 1                                                 '
      SectionDetail = ConfigArray(Y)                            '
    Loop                                                        '
  End If                                                        '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   SectionDetail = "***" ' reset to NON blank
Next Y

' VALIDATE FISHEYE PARAMETERS
If IsNumeric(FEDistortion) = False Then GoTo Failed_Validation
'If FEDistortion < Fisheye.TFactor.Min Or FEDistortion > Fisheye.TFactor.Max Then GoTo Failed_Validation
If IsNumeric(FECentreX) = False Then GoTo Failed_Validation
If IsNumeric(FECentreY) = False Then GoTo Failed_Validation
If IsNumeric(FEWidth) = False Then GoTo Failed_Validation
If IsNumeric(FEHeight) = False Then GoTo Failed_Validation
If IsNumeric(FEScale) = False Then GoTo Failed_Validation
If FecCameraModel = "" Then FecCameraModel = FecFileName

Call CLPProgressBar.ProgressBarPosition(0.65)
DoEvents

'Check version of FEC file
'Select Case FECRev
    'vvvv PCN2850 ****************************************************
    '!!!!!!!!! Notes: When updating ...               !!!!!!!!!!!!!!!!
    '^^^^ ************************************************************
    If FECRev = "2.0" Or FECRev = "2.1" Then
        'Write back this information to the INI
        Call INI_WriteBack(MyFile, "Fish_Distortion=", FEDistortion)
        Call INI_WriteBack(MyFile, "Fish_CenterX=", FECentreX)
        Call INI_WriteBack(MyFile, "Fish_CenterY=", FECentreY)
        Call INI_WriteBack(MyFile, "Fish_OriginalWidth=", FEWidth)
        Call INI_WriteBack(MyFile, "Fish_OriginalHeight=", FEHeight)
        Call INI_WriteBack(MyFile, "Fish_Ratio=", FEScale)
        Call INI_WriteBack(MyFile, "FecFileName=", FecFileName)
        Call INI_WriteBack(MyFile, "CameraModel=", FecCameraModel) 'PCN3595 (21 Oct 2005, Antony) added CameraModel to FEC File
        Call INI_WriteBack(MyFile, "Fish_DistortionHorizontal=", 1) 'PCN3687
        'Load FEC information from the INI
        Call InitializeFishEyeFromINI
        Call setoriginalsize(FEWidth, FEHeight)
        Call calculatescale
        Call CreateFishEyeMask
        Call getscalevalue(FEScale)
        Call INI_WriteBack(MyFile, "Fish_Ratio=", FEScale)
        Call InitializeFishEyeFromINI
        
    End If
    If FECRev = "2.1" Or FECRev = "2.0" Then
    Else
        GoTo Failed_Validation
    End If
'End Select
   Call CLPProgressBar.ProgressBarPosition(1)

Exit Sub
Failed_Validation:
    MsgBox DisplayMessage("Can not read this file, loading default settings.")
    'Load the default FEC settings
    Call FECLoadDefaultSettings
    Call CLPProgressBar.ProgressBarPosition(1)
Exit Sub
Err_Handler:
Select Case Err
    Case 9: GoTo Failed_Validation 'Subscript out of range (end of file)
    '    Exit Function
    Case 75: GoTo Failed_Validation
    Case Else
        MsgBox Err & " - " & error$
End Select

End Sub


Function LiveFishEyeON()
    Call livefisheye(1)
End Function
Function LiveFishEyeOFF()
    Call livefisheye(0)
End Function

Public Sub CreateFishEyeMask()
    Call createmask
End Sub

Sub PopulateCameraDropDown(CameraControl As Control)
On Error GoTo error_handler
    
    Dim FecCameraModel As String
    

    
    ReDim TheFECFiles(0)
    Call LoadCameras(CameraControl)
    
    CameraControl.Font.Size = 10
    CameraControl.Font = "MS Sans Serif"
    CameraControl.Font.Bold = True
    
    Call GetINI_ParameterInfoOnly(MyFile, "CameraModel=", FecCameraModel)
    CameraControl.text = FecCameraModel
    
    If PVDFileName <> "" Then CameraControl.text = "" 'if loaded a PVD then unkown Fec for PVD
    
    
    
Exit Sub
error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-" & error$, vbExclamation
    End Select
End Sub

Function ParseCameraName(ByVal FecFileName As String) As String
On Error GoTo error_handler

    Dim FileNo
    Dim InputString As String
    Dim FecFileArray() As String
    Dim Parameter As String
    Dim value As String
    Dim VersionNo As Single
    Dim LineNo As Integer
    Dim i As Integer
    Dim PathName As String
    Dim FileName As String
    Dim ExtName As String
    
    LineNo = 0
    
    ReDim FecFileArray(0)
    
    FileNo = FreeFile
    
    Open App.Path & "\Fec Files\" & FecFileName For Input As #FileNo
        Do While Not EOF(FileNo)   ' Check for end of file.
        Line Input #1, InputString ' Read line of data.
        FecFileArray(LineNo) = InputString
        LineNo = LineNo + 1
        ReDim Preserve FecFileArray(LineNo)
    Loop
    Close #FileNo
    
    'Looking for version number
    For i = 0 To LineNo
        If FecFileArray(i) = "[Revision]" Then
            Call GetParam(FecFileArray(i + 1), Parameter, value)
            If Parameter = "FECRevision=" Then VersionNo = value
        End If
    Next i
    
    If VersionNo < 2.1 Then ParseCameraName = Left(FecFileName, Len(FecFileName) - 4): Exit Function
    For i = 0 To LineNo
        If FecFileArray(i) = "[Camera]" Then
            Call GetParam(FecFileArray(i + 1), Parameter, value)
            If Parameter = "CameraModel=" Then ParseCameraName = value: Exit Function
        End If
    Next i
    
    'PCN3861 put the whole path name in the minus the extension in the drop down, instead of the file name only
    Call SplitFilePath(FecFileName, PathName, FileName, ExtName)
    ParseCameraName = FileName
    
Exit Function
error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-" & error$, vbExclamation
    End Select
End Function

Sub LoadCameras(CameraControl As Control)
On Error GoTo error_handler
    
    Dim FecTextFiles As String
    Dim FecFileName As String
    Dim CameraName As String
    Dim CountFecFiles As Integer
    CountFecFiles = 0

    FecTextFiles = App.Path & "\Fec Files\*.fec"
    FecFileName = Dir(FecTextFiles)
    Do While FecFileName <> ""   ' Start the loop.
        CameraName = ParseCameraName(FecFileName)
        If CameraName <> "None" Then
            CameraControl.AddItem (CameraName)
            TheFECFiles(CountFecFiles) = FecFileName
            CountFecFiles = CountFecFiles + 1
            ReDim Preserve TheFECFiles(CountFecFiles)
        End If
        FecFileName = Dir   ' Get next entry.
    Loop

Exit Sub
error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-" & error$, vbExclamation
    End Select
End Sub

