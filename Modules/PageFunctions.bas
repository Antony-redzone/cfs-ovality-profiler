Attribute VB_Name = "PageFunctions"
 Option Explicit
 

 Private Declare Sub clearline_LoadPVD_Data Lib "Clearline.dll" (ByVal FileName As String, _
                                                                 ByVal PVDataStartAddress As Long, _
                                                                 ByVal PVDataBlockSize As Long, _
                                                                 ByVal XY As Long, _
                                                                 ByRef pvDataX As Single, _
                                                                 ByRef pvDataY As Single, _
                                                                 ByVal PVDataXYMultiplier As Double, _
                                                                 ByVal FromFrame As Long, _
                                                                 ByVal ToFrame As Long) 'PCN.l.;;;;;;;;;;;;;;;;l./l;3603
Private Declare Sub clearline_MoveFileData Lib "Clearline.dll" (ByVal FileName As String, _
                                                                ByVal FromFilePosition As Long, _
                                                                ByVal ToFilePosition As Long)
Private Declare Sub clearline_EmbedFileData Lib "Clearline.dll" (ByVal PVDFileName As String, _
                                                                 ByVal EmbFileName As String, _
                                                                 ByVal CurrentFileLocation As Long)
Private Declare Sub clearline_ExtractEmbedFile Lib "Clearline.dll" (ByVal PVDFileName As String, _
                                                                    ByVal EmbFileName As String, _
                                                                    ByVal CurrentFileLocation As Long, _
                                                                    ByVal FileLength As Long)

Private LoadingTimeStampError As Boolean
 

Private Const OFS_MAXPATHNAME = 256
Private Const OF_WRITE = &H1
Private Const FILE_BEGIN = 0

Private Type OFSTRUCT
        cBytes As Byte
        fFixedDisk As Byte
        nErrCode As Integer
        Reserved1 As Integer
        Reserved2 As Integer
        szPathName(OFS_MAXPATHNAME) As Byte
End Type
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, _
                                                        lpReOpenBuff As OFSTRUCT, _
                                                  ByVal wStyle As Long) As Long

Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long

Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, _
                                                        ByVal lDistanceToMove As Long, _
                                                        ByVal lpDistanceToMoveHigh As Long, _
                                                        ByVal dwMoveMethod As Long) As Long
                                                        
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public ProgressBarPercent As Integer 'PCN4241 For keeping track of the progress bar percentage

'Public UniMessageBox As UnicodeMsgBox

Public PMBAnswer As Integer





Sub OpenAnyFile(ToOpenFileName As String) 'PCN2133 (A Parameter added)
On Error GoTo Err_Handler
Dim CBSFilenm As String
Dim FileErrorReturned As Boolean 'PCNGL140103
Dim GraphIndex As Integer 'PCNGL140103
Dim LoadAVINow As Integer
Dim FishMsgbox As Variant 'PCN2392
Dim FileExtension As String


Dim Ans As Integer
If Len(ClearLineScreen.CurrentFile) > 0 And ClearLineScreen.ChangeFlag Then
    'Ans = MsgBox(DisplayMessage("Discard current drawing?"), vbYesNo + vbExclamation) 'PCN2111
    ProfilerMessageBox.MsgBoxYesNo (DisplayMessage("Discard current drawing?"))
    Ans = PMBAnswer
    If Ans = vbNo Then
        Exit Sub
    Else
        ClearLineScreen.ChangeFlag = False
    End If
End If

Call PrecisionVisionGraph.ObservationClose_Click 'PCN4588
Call PrecisionVisionGraph.ObservationClose_Click 'PCN4588 needs to be called twice, to be able to close
                                                 ' two levels of observations


If PVRecording = True Then
    'MsgBox DisplayMessage("To open a new file, stop the recording and the video, then open the new file."), vbExclamation 'PCN2111
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("To open a new file, stop the recording and the video, then open the new file."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Exit Sub
End If

Call ClearLineScreen.ProfilerPause
    
If PVDSaved = False And LastRecordedFrame > 0 Then 'PCN1895
    'MsgBox DisplayMessage("The .pvd file is not saved, click on Save button if you wish to save this file."), vbExclamation 'PCN2111
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("The .pvd file is not saved, click on Save button if you wish to save this file."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
End If

PipelineDetails.ZOrder 0 'PCN1777 'PCNLS190203

'PipelineDetails.CommonDialog1.Filter = "JPEG (*.jpg;*.jpeg)|*.jpg;*.jpeg|Bitmap (*.bmp)|*.bmp"
'PipelineDetails.CommonDialog1.Filter = "Precision Vision Files (*.pvd)|*.pvd|Image Files (*.jpg;*.bmp)|*.jpg;*.bmp"
'PCN3093vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
If Registered = True Then
    If ToOpenFileName = "" Then 'PCN2133 ---v
'    PipelineDetails.CommonDialog1.Filter = "Precision Vision Files (*.pvd)|*.pvd|Image Files (*.jpg;*.bmp)|*.jpg;*.bmp|AVI Files (*.avi)|*.avi|Mpeg Files (*.mpg;*.mpa;*.m2p;*.mp2)|*.mpg;*.mpa;*.m2p;*.mp2|VOB Files (*.VOB;*.vob)|*.VOB;*.vob" 'PCN1915 'PCN2871
'    PipelineDetails.CommonDialog1.FileName = ""
'    PipelineDetails.CommonDialog1.ShowOpen
'    ToOpenFileName = PipelineDetails.CommonDialog1.FileName
    End If '--------------------------------^
    FileExtension = UCase(Right(ToOpenFileName, 4))
Else
    If ToOpenFileName = "" Then 'PCN2133 ---v
        PipelineDetails.CommonDialog1.Filter = "Precision Vision Files (*.pvd)|*.pvd" 'PCN1915 'PCN2871"
        PipelineDetails.CommonDialog1.FileName = ""
        PipelineDetails.CommonDialog1.ShowOpen
        ToOpenFileName = PipelineDetails.CommonDialog1.FileName
    End If '--------------------------------^
    FileExtension = UCase(Right(ToOpenFileName, 4))
    'PCN3866, if the file extension is typed in it did load a mpg or other file
    If FileExtension = ".MPG" Or _
       FileExtension = ".MPA" Or _
       FileExtension = ".M2P" Or _
       FileExtension = ".MP2" Or _
       FileExtension = ".AVI" Or _
       FileExtension = ".VOB" Or _
       FileExtension = ".BMP" Or _
       FileExtension = ".JPG" Then
        ToOpenFileName = ""
        PipelineDetails.CommonDialog1.FileName = ""
    End If
End If
'End PCN3093^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^


If Len(ToOpenFileName) > 0 Then 'PCN2133

    
    Call CLPProgressBar.ProgressBarPosition(0.05) 'PCN4171
    DoEvents
    'Disable the controls toolbars
    ControlsScreen.Enabled = False
    ControlsMain.Enabled = False

    'Initialise variables
    Call PrecisionVisionGraph.ResetPVData 'PCNLS200203 'PCN2201
    'Disable Pic n Pic PCN2223
'    PVDFileName = LocToSave & DefaultPVDFileName ' Reset file name to default 'PCNGL170103
'   PCN1768 *** PVDFileName = LocToSave & DefaultPVDFileName ' Reset file name to default 'PCNGL170103
    PVDFileName = "" ' PCN1768
    'vvvv PCNGL200303-1 ************************************
    ' Remove the recording file DefaultPVDFileName
    Kill LocToSave & DefaultPVDFileName
    '^^^^ **************************************************
    
    ClearLineScreen.EmptyCBuffer 'PCNLS030403  Fixes problem of having PV data from old file
    VideoFileName = "" ' Reset file name 'PCNGL170103
    PVRecording = False  'PCNLS030203
    ExpectedDiameter = 0 'PCN1835
    ReDim StoredReportArray(0) 'PCNGL05100601
    ReDim PipeObservations(0)
    ReDim Observations.ReferenceShapeShiftObs(0)
    ReDim Observations.WaterLevelShiftObs(0)
    
    '''''''''''''''''''''''''''''''''''
    Unload PVReport1K
    Unload PVReport4in1
    Unload PVReportMultiProfilex3
    Unload PVReportProfile
    Unload PVReportSingle
    Unload PVReportStoredInPVD
    
    Call ObsInitPictureStorage 'PCN4514
    

    MaxCalculatedFrameNo = 0 'PCN2310 'PCN2970
    ' Turn off Picture in Picture
    ClearLineScreen.PVScreenPicInPic.Visible = False 'PCNGL240103
    PicInPicMode = "OFF" 'PCNGL240103
    'vvvv PCN2463 **************************
    'Distance variables
    'vvvv PCN2639 *******************************
    'read from ini Distance method
    Call GetINI_ParameterInfoOnly(MyFile, "DistanceMethod=", DistanceMethod)
    If DistanceMethod = "" Then
        DistanceMethod = "None"
    End If
    ConfigInfo.DistanceProcessMethod = DistanceMethod
    ConfigInfo.DistanceStart = InvalidData 'Not required 'PCN4171
    DistanceStart = InvalidData 'PCN4171
    '^^^^ ***************************************
'    CameraSpeedInFrames = 0
'    CameraSpeedInTime = 0
    CameraSpeedInFrames = 0.006 'm/frame from eg of [10m/min / (60sec/min * 25frames/sec)]
    CameraSpeedInTime = 0.015 'm/sec from eg of [10m/min / 60sec/min)]

    PrecisionVisionGraph.Y_Units.Visible = False
    '^^^^ **********************************
    'vvvv PCN3490 **********************************
    PrecisionVisionGraph.PipeDisplay.AutoRedraw = True
    PrecisionVisionGraph.PipeDisplay.Cls
    '^^^^ **********************************
    Call ControlsScreen.SetDisplayZoomOnSnap("Off") 'PCN4171
    'vvvv PCNXXXX *********************************
    Call ScreenDrawing.GraphSelect("Flat", 0) 'PCN4171
    PrecisionVisionGraph.SetLimitLines 'PCN2769
    Call PrecisionVisionGraph.SetupPVGraphScreen(ImageGraphState(0).GraphType)
    Call PrecisionVisionGraph.GetGeneralPVGraphData(ImageGraphState(0).GraphType)
    PrecisionVisionGraph.Label_GraphName(0) = PrecisionVisionGraph.GetContainerGraphLabel(0)
    PrecisionVisionGraph.Label_GraphNameShadow(0) = PrecisionVisionGraph.Label_GraphName(0)
    '^^^^ *****************************************
    Call ScreenDrawing.DeleteAll 'PCN4184
    ClearLineScreen.SetDimenResultsSize (False) 'PCN4184 'Reset DimenMeasure
    If mediatype = "Live" Then
        'ClearLineScreen.ControlToolbar.Buttons.Item(5).Enabled = True 'PCN2733
        Call ClearLineScreen.UnitLiveFeed
''        ClearLineScreen.ControlToolbar.Buttons.Item(1).Image = 1 'Disconnected PCNGL270103 'PCN2681 'PCN4171
        'Enable AVI Play buttons
''        ClearLineScreen.ControlToolbar.Buttons.Item(9).Enabled = True  'PCNGL270103 'PCN2681
''        ClearLineScreen.ControlToolbar.Buttons.Item(10).Enabled = True 'PCNGL270103 'PCN2681
''        ClearLineScreen.ControlToolbar.Buttons.Item(11).Enabled = True 'PCNGL270103 'PCN2681
        'ClearLineScreen.SnapShotScreen.Cls 'PCNLS 1999 'PCN3219
        'Disable Record button 'PCN2831
''        ClearLineScreen.ControlToolbar.Buttons.Item(5).Enabled = False  'PCN2831
    ElseIf mediatype <> "" Then  'Mpeg or AVI
        'ClearLineScreen.ControlToolbar.Buttons.Item(5).Enabled = True 'PCN2733
        Call ClearLineScreen.UnInitVideo
        Call ClearLineScreen.UnitVideoSlider
        'ClearLineScreen.SnapShotScreen.Cls 'PCNLS 1999 'PCN3219
    End If
    If ThreeDRunning = True Then
        ClearLineScreen.Unload3D
    End If
    mediatype = "" 'PCNLS220103
      'Reset the array storing times of PV data PCN 1988 LS 11/7/03
    Call ClearLineScreen.Initialization 'Measuring tools PCNGL190103
'    Call InitilisePVProfile(1) 'PCNGL170103 PCN2970
''    Call ClearLineScreen.ResetRecord 'PCN1792 'PCNLS190203
    Call InitialiseFieldsOnForms 'PCNGL170103
    'Enable the AVI record button 'PCNGL140103
''    ClearLineScreen.ControlToolbar.Buttons(5).Enabled = True 'PCN2681
''    ClearLineScreen.ConfigToolBar1.Buttons(1).Enabled = True
''    ClearLineScreen.ConfigToolBar1.Buttons(5).Enabled = True 'PCN2759
    'Check for Precision Vision Data file
    If UCase(Right(ToOpenFileName, 4)) = ".PVD" Then 'PCN2133
        Call OpenPVDFile(ToOpenFileName)
        'Call Open2ndPVDData(ToOpenFileName) 'PCN4380
        
        
    ElseIf UCase(Right(ToOpenFileName, 4)) = ".MPG" _
        Or UCase(Right(ToOpenFileName, 4)) = ".M2V" _
        Or UCase(Right(ToOpenFileName, 4)) = ".MPA" _
        Or UCase(Right(ToOpenFileName, 4)) = ".M2P" _
        Or UCase(Right(ToOpenFileName, 4)) = ".VOB" _
        Or UCase(Right(ToOpenFileName, 4)) = ".AVI" Then  'PCN1915 PCN2133
            Call OpenVideoFile(ToOpenFileName)
            
    ElseIf UCase(Right(ToOpenFileName, 4)) = ".JPG" _
        Or UCase(Right(ToOpenFileName, 4)) = ".BMP" Then
            Call OpenStillImageFile(ToOpenFileName)
    Else
        'Process image files
    End If
End If
If VideoFileName <> "" Then
'PCN 2146
    'ClearLineScreen.ControlToolbar.Buttons.Item(5).Enabled = True 'PCN2681 'PCN2733
End If


''Call ClearLineScreen.PVRecordButtonSet 'PCN2460

Call DrawMainScale(ClearLineScreen.PVScreen) 'PCN3691

Call ClearLineScreen.VideoFrameSliderSetup  'PCN2930
        
''PCN3513 no longer need background load (Antony, 12 may 2005)
''
'''vvvv PCN2970 ***************************************************
''If UCase(Right(ToOpenFileName, 4)) = ".PVD" Then
''    'Start loading the Flat3D and possible the Max/Min Diameter data
''    'in the background.
''    MaxDisplayedFrameNo = 0
''    DoEvents
''    Call PVDataCalcsBackgroundLoad
''End If
'''^^^^ ***********************************************************

If IsOpen("CLPProgressBar") Then 'PCN4171
    Call CLPProgressBar.ProgressBarPosition(1#)
End If

'Enable the controls toolbars
ControlsScreen.Enabled = True
ControlsMain.Enabled = True


Exit Sub
Err_Handler:
Select Case Err 'PCNGL200303-1
    Case 53 'File not found  'PCNGL200303-1
        Resume Next
    Case 32755 'Cancel the file open
        Exit Sub
    Case Else
        MsgBox Err & "-PF1:" & Error$
End Select

End Sub
Sub OpenVideoFile(ToOpenFileName As String)
On Error GoTo Err_Handler
    Dim CBSFilenm As String
    
    Call CLPProgressBar.ProgressBarPosition(0.2) 'PCN4171
    DoEvents
    AutoTune.TuningFrame.Enabled = True
    PossibleConfigInfoCurruption = False
    Call ClearLineScreen.ProfilerPause
    mediatype = Video
    
    'PCN6312
     DesignGradient = 0: PipelineDetails.DesignGradientTextBox.text = "0"
     SeaLevelEndHeight = 0: PipelineDetails.SeaLevelEndHeightTextBox.text = "0"
     SeaLevelStartHeight = 0: PipelineDetails.SeaLevelStartHeightTextBox.text = "0"
    ''''''''
    
''    ClearLineScreen.ControlToolbar.Buttons.Item(5).Enabled = True 'PCN2733
    CBSFilenm = Left(ToOpenFileName, Len(ToOpenFileName)) 'PCN2133
    VideoFileName = CBSFilenm 'PCNGL140103
    ClearLineTitle.TitleBarCaption.Caption = DisplayMessage(Video) & " - " & ToOpenFileName 'PCNGL210403-2 PCN2133 'PCN2759 'PCN4171
    ClearLineTitle.TitleBarCaption.ToolTipText = DisplayMessage(Video) & " - " & ToOpenFileName 'PCN4171
    'Reset PV data and screens 'PCNGL241202
    'vvvv PCN2639 *******************************************
'        PVDataNoOfLines = 1
'        RequestFrameNo = 1 'PCNGL060103
'        MaxDisplayedFrameNo = 0 'Initialise 'PCNGL070103
'        Call InitilisePVProfile(MaxFrameBufferNo) 'PCNGL140103
    Call PrecisionVisionGraph.ResetPVData
    
    '^^^^ ***************************************************
    PipelineDetails.ZOrder 0 'PCNGL170103 'PCNGL300103
    
    Call CLPProgressBar.ProgressBarPosition(0.4) 'PCN4171
    
    DoEvents 'PCNGL300103
    'Set MainScreen for video
    ClearLineScreen.PVScreen.Visible = False
    'Enable AVI Play buttons 'PCNGL1812022
    Call ClearLineScreen.SetupMainScreenForVideo 'PCNGL301202
    'Need to get first frame of the avi
    Call SetAVIInitialised 'PCNGL150103

    ClearLineScreen.VideoScreen.AutoRedraw = True
    ClearLineScreen.VideoScreen.Visible = True
    ClearLineScreen.InitVideo
    
    Call CLPProgressBar.ProgressBarPosition(0.5) 'PCN4171
    DoEvents
    
    'vvvv PCN2639 ************************************************************
    If IgnoreDistX1 > 0 And IgnoreDistY1 > 0 And IgnoreDistX2 > 0 And IgnoreDistY2 > 0 Then
        Call ClearLineScreen.SetRectangle(IgnoreDistX1, IgnoreDistY1, IgnoreDistX2, IgnoreDistY2, "Distance")
    End If
    '^^^^ ********************************************************************
    'FISH-EYE( PCN2290 ) -v
    If IgnoreX1 > 0 And IgnoreY1 > 0 And IgnoreX2 > 0 And IgnoreY2 > 0 Then
        Call ClearLineScreen.SetRectangle(IgnoreX1, IgnoreY1, IgnoreX2, IgnoreY2, "Ignore1") 'PCNGL280503-1 'PCN2639
    End If
    'FISH-EYE( PCN2290 ) -^

    Call CLPProgressBar.ProgressBarPosition(0.7) 'PCN4171
    DoEvents
    
    Call hough_processimageonoff(True)
    CalibrationMethodActioned = "" 'PCN4211 Reset video calibration
    Call SetupVideoDisplayAsNormal 'PCN2612
    ConfigInfo.FishEyeDistortion = 0 'PCN3039 Distortion have to be set to force a Fisheye calculation
'    Call FisheyeFunctions.SetDistortion(Fisheye.TFactor.value) 'PCN3039 Even thou the fish eye was set the mask was not yet
                                                      ' created, regardless if you chose to set fisheye or not.
    Call FishEyeLoadFileCheck(Video) 'PCN2392
    
    Call CLPProgressBar.ProgressBarPosition(0.8) 'PCN4171
    DoEvents
    
    LastDataTime = 0 'PCNANT????
    Call ScreenDrawing.ClearAllGraphsAndRuler 'PCN3402

    Call Observations.ClearAllObservationsAndDistanceSettings
    
    Call CLPProgressBar.ProgressBarPosition(0.85) 'PCN4171
    ReDim ScreenDrawing.DrawingMaskBox(0)
    DoEvents
    
    Call CheckForIPD 'PCN3744
    PVDrawScreenRatio = ConfigInfo.Ratio 'Set the PVDrawScreenRatio for the current callibrated measurements
    Call ClearLineScreen.VideoScreenScaleCalc
    Call hough_processimageonoff(False)
    ConfigInfo.DistanceStart = InvalidData 'Not required
    ConfigInfo.DistanceFinish = InvalidData 'Remove the last distance settings
    DrawingCentreX = CentreLineX
    DrawingCentreY = CentreLineY
    
    Call CLPProgressBar.ProgressBarPosition(0.95) 'PCN4171
    DoEvents
    
    'If IsOpen("OptionsPage") Then
        OptionsPage.FishEyeCameraDropdown.Enabled = True

        
    'End If
    ConfigInfo.WLFinishAngle = 0: WLFinishAngle = 0
    ConfigInfo.WLStartAngle = 0: WLStartAngle = 0
    ReDim WaterEgnoreList(180)

    'vvvv PCN4171 ************************
    CLPScreenMode = Video
    Call ControlsScreen.ControlsViewSetup
    Call ControlsMain.ControlsDisplaySetup("DisplayPipeDetails")
    '^^^^ ********************************
    
    'PCN4433' clear the report titles on new load of mpg
    UserTitleAnalysis = ""
    UserTitleSummary = ""
    UserTitleObservations = ""
    UserTitleProfile = ""

    
    
    

Exit Sub
Err_Handler:
Select Case Err 'PCNGL200303-1
    Case Else
        MsgBox Err & "-PF2:" & Error$
End Select
End Sub
Sub OpenStillImageFile(ToOpenFileName As String)
On Error GoTo Err_Handler
    Dim CBSFilenm As String
    mediatype = StillImage
    PossibleConfigInfoCurruption = False
    'PCN4406
    DrawingCentreX = CentreLineX
    DrawingCentreY = CentreLineY
    

    ClearLineScreen.LoadImage ToOpenFileName 'PCN2133
    ClearLineScreen.CurrentFile1 = ToOpenFileName 'PCN2133
    ClearLineScreen.CurrentFile = ToOpenFileName 'PCN2133
    PipelineDetails.Visible = True
    PipelineDetails.ZOrder 0 'PCN2373
    If UCase(Right(ToOpenFileName, 4)) = ".JPG" Then 'PCN2133
        CBSFilenm = Left(ToOpenFileName, Len(ToOpenFileName) - 3) & "CBS" 'PCN2133
    ElseIf UCase(Right(ToOpenFileName, 4)) = ".BMP" Then 'PCN2133
        CBSFilenm = Left(ToOpenFileName, Len(ToOpenFileName) - 3) & "CBS" 'PCN2133
    Else
        'MsgBox "Filename is wrong. Please check filename. Measurement information is saved in temporary.cbs file.", vbInformation 'PCN2111
       'PCN1972 LS 8/7/03
        'CBSFilenm = App.Path & "\CBS\Temporary.CBS"
        CBSFilenm = LocToSave & "Temporary.CBS"
    End If
    ImageFileName = ToOpenFileName
    'Call LoadPipelineDetails(CBSFilenm)
    'Call GetPipe_Information(CBSFilenm) 'ML200203 'Commented out as per GL25.02.03
    MaxDisplayedFrameNo = 0 'Initialise 'PCNGL070103
'Make it into snapshot mode****
'PCN
    'VideoSnapShotMode = SnapShot
    'CLPScreenMode = Video
    CLPScreenMode = SnapShot 'PCN4043
    Call ControlsScreen.ControlsViewSetup
    Call ControlsMain.ControlsDisplaySetup("DisplayPipeDetails") 'PCN4171
    
    'PCN4406 ''''''''''''''''''''''''''''''''''''
    Call ClearLineScreen.SetDimenResultsSize(True)              '
    ClearLineScreen.DimenResults.ZOrder 0       '
    ClearLineScreen.AreaResults.ZOrder 0        '
    '''''''''''''''''''''''''''''''''''''''''''''
    
    
''    Call ClearLineScreen.SetupMTButtonsForSnapShot 'PCNGL300103
    'PVScreen.MousePointer = 99
    'PVScreen.MouseIcon = LoadResPicture(MainScreenMouseIcon, vbResIcon) 'PCNGL291102
    'ConfigToolBar1.Buttons.Item(1).Image = 1
    ClearLineScreen.SnapShotScreen.Visible = True
    ClearLineScreen.SnapShotScreen.ZOrder 0 'PCNGL261202
    ClearLineScreen.SnapShotScreen.AutoRedraw = True 'PCNGL2901032
    'ClearLineScreen.SnapShotScreen.Cls 'PCNGL2901032 'PCN3219

    ClearLineScreen.SnapShotScreen.AutoRedraw = False 'PCNGL140203
    Call DrawAll   'PCNGL2901032
'    Call ClearLineScreen.InitImageProcessing 'PCN3194
Exit Sub
Err_Handler:
Select Case Err 'PCNGL200303-1
    Case Else
        MsgBox Err & "-PF3:" & Error$
End Select

End Sub

Sub OpenPVDFile(ToOpenFileName As String)
On Error GoTo Err_Handler
Dim GraphSet As Integer 'PCN2970
Dim CBSFilenm As String
Dim FileErrorReturned As Boolean 'PCNGL140103
Dim LoadAVINow As Integer
Dim i As Integer
Dim ErrorStr As String
   

    
    ClearLineScreen.ProfilerPause
    AutoTune.TuningFrame.Enabled = False
    PrecisionVisionGraph.ZOrder 0 'PCNLS200203

    Call ScreenDrawing.ClearAllGraphsAndRuler 'PCN3402
    For i = 0 To 180
        WaterEgnoreList(i) = 0
    Next i

    VideoFileName = "" ' Flag that this file is not an AVI

    '************ Setup Loading message **************************** 'PCNGL140103
    'PCN3373 Call SetPositionOfPVGraphBaseCover 'PCN2970
'    Call SetPVScreensHeights(GraphSet) 'PCN2970
    DoEvents
    '*******************************************************************
    'Process Precision Vision Data file
    CBSFilenm = ToOpenFileName 'PCN2133

    CLPScreenMode = PV 'PCN1863
    Call ControlsScreen.ControlsViewSetup
    ' Turn off Picture in Picture
    
    ClearLineScreen.PVScreenPicInPic.Visible = False
    PicInPicMode = "OFF"

    'Call PVProfileLoad(CBSFilenm) 'PCNGL140103
    'vvvv **************************************************** 'PCNGL140103
    'Load from file the PVD data and if applicable the AVI file
    PVDFileName = CBSFilenm 'PCNGL140103
    
    'vvvv PCN4241 **********************************
    'Check to see if the PVD is read only
    If ThisFileIsReadOnly(PVDFileName) And SoftwareConfiguration <> "Reader" Then
        'MsgBox DisplayMessage("Warning this PVD is Read ONLY. Unable to save changes."), vbExclamation
        ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Warning this PVD is Read ONLY. Unable to save changes."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    End If
    '^^^^ ******************************************

    Call LoadInPVDFormat(CBSFilenm, FileErrorReturned)
    'vvvv PCN2401 *******************************
'    'vvvv************PCN4195
'    If isopen("CLPProgressBar") Then
'        Call CLPProgressBar.ProgressBarPosition(1#)
'        DoEvents
'    End If
'    '^^^^************PCN4195
    '^^^^ *******************************************
    If FileErrorReturned Then Exit Sub
    'Setup to Draw PV Profile 'PCNGL220103
    
    PrecisionVisionGraph.StartupBackgroundImage.Visible = False
    
    Call DrawPVProfile_Setup(ClearLineScreen.PVScreen) 'PCNGL220103 PCN3526

    Call ProcessBarIncrement 'PCN4241

    Call ClearLineScreen.SetupMainScreenForPV 'PCN1858

    Call ProcessBarIncrement 'PCN4241

    Dim strTemp As String 'Temporary string for message box standardisation
    'If the PVD file has no AVI associated with it, ask if they want to load an
    'associated one LS050203
    If LoadVideo Then
        If VideoFileName = "notfound" Or Dir(VideoFileName) = "" Or VideoFileName = "" Then
    
            If VideoFileName <> "" Then
                If Registered = True Then 'PCN3093
                    strTemp = DisplayMessage("Could not find the associated Video File: ")  'PCN2111
                    strTemp = strTemp + Trim(ConfigInfo.VideoFileName) + " " + DisplayMessage("Press Yes to Locate a Video File, No to continue without a Video File loaded.") 'PCN2111
                    LoadAVINow = MsgBox(strTemp, vbYesNo)
                'vvvv PCN3809 ********************************
                ElseIf SoftwareConfiguration = "Reader" Then
                    strTemp = DisplayMessage("Could not find the associated Video File: ")
                    strTemp = strTemp + Trim(ConfigInfo.VideoFileName)
                    MsgBox strTemp, vbInformation
                End If
                '^^^^ ****************************************
            Else
                If Registered = True Then 'PCN3093
                    'LoadAVINow = MsgBox(DisplayMessage("There is no Video File associated with this PVD file. Would you like to load one now?"), vbYesNo)  'PCN2111
                    ProfilerMessageBox.MsgBoxYesNo (DisplayMessage("There is no Video File associated with this PVD file. Would you like to load one now?"))
                    LoadAVINow = PMBAnswer
                'vvvv PCN3809 ********************************
                ElseIf SoftwareConfiguration = "Reader" Then
                    'MsgBox DisplayMessage("There is no Video File associated with this PVD file."), vbInformation
                    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("There is no Video File associated with this PVD file."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
                '^^^^ ****************************************
                Else
                    MsgBox (DisplayMessage("File not able to be loaded on unregistered application"))
                End If
            End If
            If LoadAVINow = vbYes Then
                PipelineDetails.CommonDialog1.Filter = "AVI Files (*.avi)|*.avi|Mpeg Files (*.mpg;*.m2v;*.mpa;*.m2p;*.mp2)|*.mpg;*.m2v;*.mpa;*.m2p;*.mp2|VOB Files (*.VOB;*.vob)|*.VOB;*.vob" 'PCN1915 'PCN2871
                PipelineDetails.CommonDialog1.FileName = ""
                PipelineDetails.CommonDialog1.ShowOpen
                CBSFilenm = Left(PipelineDetails.CommonDialog1.FileName, Len(PipelineDetails.CommonDialog1.FileName))
                VideoFileName = CBSFilenm
            Else
                CBSFilenm = "" 'PCN2857
                VideoFileName = "" 'PCN2857
            End If
            If Len(VideoFileName) > 0 Then 'PCN2857
                ConfigInfo.VideoFileName = VideoFileName 'PCNLS070203
                Call SaveToFilePipeAndConfigInfo("ConfigInfo", FileErrorReturned)
                CLPScreenMode = Video 'PCN2832
                Call ControlsScreen.ControlsViewSetup
                Call ControlsMain.ControlsDisplaySetup("DisplayPVGraph") 'PCN4171
            Else
                'Disable video button
    ''            ClearLineScreen.ConfigToolBar1.Buttons(1).Enabled = False
    ''            ClearLineScreen.ConfigToolBar1.Buttons(5).Enabled = False 'PCN2759
             
                Call PVGraphsKeyForm.DisplayPVGraphsKey 'PCN2990 'PCN2818
    
            End If
        End If
    End If
    'Check if this PVD has an associated AVI file 'PCNGL140103
    
    ''''''' PCN???? I dont think this is suppose to be there, if there is a video
    'associated then it should load, regardless if its registered, reader or not
    '
    'If Registered = True Or SoftwareConfiguration = "Reader" Then 'PCN3093 'PCN3809
    If LoadVideo Then
        If Dir(VideoFileName) <> "" And VideoFileName <> "" Then
''            ClearLineScreen.ControlToolbar.Buttons.Item(5).Enabled = False 'PCN2733
            mediatype = Video 'PCNLS050203
            'Need to get first frame of the avi

            ClearLineScreen.VideoScreen.AutoRedraw = True
            ClearLineScreen.VideoScreen.Visible = True
            ClearLineScreen.InitVideo
            'vvvv PCN2639 ************************************************************
            If IgnoreDistX1 > 0 And IgnoreDistY1 > 0 And IgnoreDistX2 > 0 And IgnoreDistY2 > 0 Then
                Call ClearLineScreen.SetRectangle(IgnoreDistX1, IgnoreDistY1, IgnoreDistX2, IgnoreDistY2, "Distance")
            End If
            '^^^^ ********************************************************************
            'FISH-EYE( PCN2290 ) -v
            If IgnoreX1 > 0 And IgnoreY1 > 0 And IgnoreX2 > 0 And IgnoreY2 > 0 Then
                Call ClearLineScreen.SetRectangle(IgnoreX1, IgnoreY1, IgnoreX2, IgnoreY2, "Ignore1") 'PCNGL280503-1 'PCN2639
            End If

            Call ProcessBarIncrement 'PCN4241

            Call SetAVIInitialised 'PCNGL150103

            Call ProcessBarIncrement 'PCN4241

            Call hough_processimageonoff(False)
            Call SetupVideoDisplayAsNormal 'PCN2612 Setups video with displaying IP settings.
            'vvvv PCN2392 *****************************************
            If ConfigInfo.FishEyeFlag Or ConfigInfo.FishEyeHorDistortion <> 1 Then
                Call FishEyeLoadFileCheck("PVD")
            Else
                'PCN3005 changed this call from FEOFF_Click
                FisheyeFunctions.FEOFF
                'Fisheye.ZOrder 1
            End If
            '^^^^ *************************************************
            
        End If
    End If
Call ProcessBarIncrement 'PCN4241
    
    'Disable the AVI record button
''    ClearLineScreen.ControlToolbar.Buttons(5).Enabled = False 'PCN2681
    '************ Disable Loading message **************************** 'PCNGL181202 'PCNGL140103
    
    'PCN3373 Call SetPositionOfPVGraphBaseCover 'PCN2970
'    Call SetPVScreensHeights(GraphSet) 'PCN2970
    DoEvents
    
    PVFrameNo = 1

    
    'Enable Pic n Pic PCN2223
''    ClearLineScreen.ConfigToolBar1.Buttons.Item(4).Enabled = True
    
    'PCN3312 Force a call to the PVScreen button after load.
    Call ControlsScreen.SetupCLPScreenToPV(ErrorStr)
    
    Call ClearAllGraphsAndRuler
    Call PrecisionVisionGraph.MoveGraph(1)
    Call ClearLineScreen.GotoPVProfile(PVFrameNo, True)
    Call Observations.ObsDisplayALL
    
    If IsOpen("Fisheye") Then
        Fisheye.CameraDropdown.Enabled = False 'PCN3595 do not allow fec selection
        Fisheye.FishEyeON.Enabled = False
        Fisheye.CameraDropdown.text = ""
    End If
       

    Call ProcessBarIncrement 'PCN4241


    'Call set water level AGAIN (to be changed later) so that the video load if there was on will reflect the
    'water level setting
    If WLStartAngle <> 0 Or WLFinishAngle <> 0 Then
        Call ClearLineScreen.SetWaterLevelinPipe(WLStartAngle, WLFinishAngle)
        WaterLevelIgnoreCenter = True 'PCNLS190603
        Call ClearLineScreen.SetUnsetWaterLevel(WaterLevelIgnoreCenter, False) 'PCN3219 WaterLevelIgnoreProfile) 'PCNLS190603waterlevel
    End If
    
    If IsOpen("OptionsPage") Then
        OptionsPage.FishEyeCameraDropdown.Enabled = False 'PCN3595
        OptionsPage.FishEyeCameraDropdown.text = ""
    End If
    
    Call PrecisionVisionGraph.SetPVGraphY_Units
    
    Call ControlsMain.ControlsDisplaySetup("DisplayPVGraph") 'PCN4171


On Error GoTo Err_Handler
Exit Sub
Err_Handler:
Select Case Err 'PCNGL200303-1
    Case 52: Resume Next  'Cant access video file, go find it
    Case Else
        MsgBox Err & "-PF4:" & Error$

End Select
    
End Sub



Sub SaveImageAndOrData() 'PCNGL110103
On Error GoTo Err_Handler
Dim SaveFilterStr As String 'PCNGL140103
Dim FileErrorReturned As Boolean 'PCNGL140103
Dim InputNoOfSegments As Variant 'PCN2481 171203
Dim NoOfSegPerProfile As Integer 'PCN2481 171203
Dim DefaultExt As String 'PCN2505
Dim StartFrameNo As Long 'PCN3060
Dim FileLoadError As Boolean 'PCN3060
Dim answer As Integer

Dim DefaultFileName As String
Dim DefaultFilePath As String
Dim DefaultFileExt As String
SaveFilterStr = ""
DefaultExt = ".pvd" 'PCN2505
DefaultFileExt = ""
DefaultFilePath = ""
'Check that this is a fresh PVD recording 'PCNGL140103
If PVDFileName = LocToSave & DefaultPVDFileName Then
    SaveFilterStr = "Precision Vision Files (*.pvd)|*.pvd"
'    DefaultExt = ".pvd" 'PCN2505 'PCN2881
    DefaultExt = ".pvd" 'PCN2881
    Call SplitFilePath(VideoFileName, DefaultFilePath, DefaultFileName, DefaultFileExt) 'PCN3834
End If

If ClearLineScreen.LetSavePictureOnly(LocToSave & "temp.bmp") And PVDFileName <> LocToSave & DefaultPVDFileName Then 'PCNGL140103
    'vvvv PCN1956 ********************************
    'Ensure all drawings are saved. '
    ClearLineScreen.SnapShotScreen.AutoRedraw = True
    Call DrawAll
    ClearLineScreen.SnapShotScreen.AutoRedraw = False
    SavePicture ClearLineScreen.SnapShotScreen.Image, LocToSave & "temp2.bmp" 'PCNGL140103
    ClearLineScreen.SnapShotScreen.AutoRedraw = True
    ClearLineScreen.SnapShotScreen.Cls
    ClearLineScreen.SnapShotScreen.AutoRedraw = False
    If SaveFilterStr <> "" Then
        SaveFilterStr = SaveFilterStr & "|"
    End If
    SaveFilterStr = SaveFilterStr & "JPEG (*.jpg;*.jpeg)|*.jpg;*.jpeg |Bitmap(*.bmp)|*.bmp" 'PCN1781 'PCN260603 'PCN2453
    '^^^^ ****************************************
End If
'vvvv PCN2376 *******************************
If ThreeDRunning Then
    If SaveFilterStr <> "" Then
        SaveFilterStr = SaveFilterStr & "|"
    End If
    SaveFilterStr = SaveFilterStr & "Export 3D in STL format (*.stl)|*.stl"
End If
'^^^^ ***************************************
'vvvv PCN2481 ********************
If PVDFileName <> "" And PVDFileName <> LocToSave & DefaultPVDFileName Then
    If SaveFilterStr <> "" Then
        SaveFilterStr = SaveFilterStr & "|"
    End If
    SaveFilterStr = SaveFilterStr & DisplayMessage("Export PV data in Microsoft Excel CSV format") & " (*.csv)|*.csv"
    SaveFilterStr = SaveFilterStr & "|" & DisplayMessage("Export single graph in tab delimited text format") & "|*.csv"
    SaveFilterStr = SaveFilterStr & "|" & DisplayMessage("Export PV data in Tab Delimited Text format") & " (*.txt)|*.txt" 'PCN2481 171203
    DefaultExt = ".csv"
End If
'^^^^ ****************************
If SaveFilterStr <> "" Then
    PipelineDetails.CommonDialog1.Filter = SaveFilterStr
    If LCase(DefaultExt) = ".pvd" Then  'PCN3834
        PipelineDetails.CommonDialog1.FileName = DefaultFilePath & DefaultFileName & ".pvd" 'PCN3834
    ElseIf Len(strMediaFilePath) > 0 Then 'PCN2133
        PipelineDetails.CommonDialog1.FileName = Left(strMediaFilePath, Len(strMediaFilePath) - 4) 'PCN2133 'PCN3952
    Else
        PipelineDetails.CommonDialog1.FileName = "" 'PCN2133
    End If 'PCN2133
    PipelineDetails.CommonDialog1.DefaultExt = DefaultExt 'PCN2505
    PipelineDetails.CommonDialog1.ShowSave
    If Len(PipelineDetails.CommonDialog1.FileName) > 0 Then
    
    If Dir(PipelineDetails.CommonDialog1.FileName) <> "" Then
        'answer = MsgBox(DisplayMessage("File already exists. Will you overwrite?"), vbYesNo)  'PCN2111
        ProfilerMessageBox.MsgBoxYesNo (DisplayMessage("File already exists. Will you overwrite?"))
        answer = PMBAnswer
    End If
    
    If answer = vbNo Then Exit Sub
    
    Select Case UCase(Right(PipelineDetails.CommonDialog1.FileName, 4))
        Case ".PVD"
            If IsOpen("CLPProgressBar") Then
                Call CLPProgressBar.ProgressBarPosition(1#)
                DoEvents
            End If
            '^^^^ *****************************************
            'Move PVD recording file to new location
            
            FileCopy PVDFileName, PipelineDetails.CommonDialog1.FileName
            Kill PVDFileName
            PVDFileName = PipelineDetails.CommonDialog1.FileName 'PCN1768
            PVDSaved = True
            'PCN2185 --------------------------------------------------------------------------v
'            Dim InterfaceFileExists As String 'PCN2246
'            InterfaceFileExists = Dir(LocToSave & "CLPInterface.int") 'PCN2246
'            If Len(PVDFileName) > 0 And InterfaceFileExists = "CLPInterface.int" Then   'PCN2246
'                Call INI_WriteBack(LocToSave & "CLPInterface.int", "MediaFilePath=", PVDFileName) 'PCN2123, PCN2176
'                strMediaFilePath = PVDFileName
'            End If
            '---------------------------------------------------------------------------^
        Case ".JPG"
            SaveThis PipelineDetails.CommonDialog1.FileName
        Case ".BMP"
            SaveThis PipelineDetails.CommonDialog1.FileName
        Case ".STL" 'PCN2376
            Call ClearLineScreen.D3D_ExportToFile(PipelineDetails.CommonDialog1.FileName, "STL") 'PCN2376
        Case ".CSV" 'PCN2481
            
            
            'vvvv PCN2481 ********************
            'InputNoOfSegments = InputBox(DisplayMessage("How many points of the profile do you wish to export?") & " (6, 12, 18, 36 or 180)")
            'If InputNoOfSegments = "" Then Exit Function 'PCN2882
            FileErrorReturned = False
            ' NoOfSegPerProfile = CInt(InputNoOfSegments)
            ' If FileErrorReturned Then
            '   NoOfSegPerProfile = NoOfProfileSegments
            ' End If
            If PipelineDetails.CommonDialog1.FilterIndex = 2 Then
                Call ExportSingleGraph(PipelineDetails.CommonDialog1.FileName)
            Else
                Call ExportPVData(PipelineDetails.CommonDialog1.FileName, "CSV", 1, PVDataNoOfLines, 180) 'PCN2481
            End If
            '^^^^ ****************************
        Case ".TXT" 'PCN2481 171203
            'vvvv PCN2481 171203 ********************
            InputNoOfSegments = InputBox(DisplayMessage("How many points of the profile do you wish to export?") & " (6, 12, 18, 36 or 180)")
            If InputNoOfSegments = "" Then Exit Sub      'PCN2882
            FileErrorReturned = False
            NoOfSegPerProfile = CInt(InputNoOfSegments)
            If FileErrorReturned Then
                NoOfSegPerProfile = NoOfProfileSegments
            End If
            Call ExportPVData(PipelineDetails.CommonDialog1.FileName, "TXT", 1, PVDataNoOfLines, NoOfSegPerProfile) 'PCN2481
            '^^^^ ****************************
    End Select
    End If
Else
    'Save configuration and pipeline information 'PCNGL140103
    
End If
    
    
Exit Sub
Err_Handler:
Select Case Err
    Case 6 'Overflow 'PCN2481
        'FileErrorReturned = True 'PCN2481
        'Resume Next 'PCN2481
        Exit Sub      ' This is a error purposly thrown by cancel, so exit, DON'T CONTINUE....... :( :( PCN3951
    Case 13 'Type Mismatch 'PCN2481
        FileErrorReturned = True 'PCN2481
    Case 70: FileErrorReturned = True: Exit Sub      'PCN3762
    Case 32755: Exit Sub      'PCN3591 This is the the cancel thrown error not 6. So I dont know why six is what it is.
    Case Else
        MsgBox Err & "-PF5:" & Error$
End Select
End Sub

Function GetFileInformation(ExpectedHdr As String, tmp As String) As String

' MGR 23/10/02
' Compare First delimited entry in string to passed value and if matched then return subsequent values

On Error GoTo Err_Handler

Dim X As Long
Dim Y As String

' get first comma delimited entry

X = InStr(tmp, ",")
Y = Left(tmp, X)




Exit Function
Err_Handler:
    MsgBox Err & "-PF6:" & Error$
End Function

Sub SaveThis(Filename1)
On Error GoTo Err_Handler
    Dim answer As Integer 'PCN1916
    Dim i As Integer
    Dim J As Integer
    Dim fs
    Set fs = CreateObject("Scripting.FileSystemObject")

    answer = vbYes
    If UCase(Right(Filename1, 3)) = "JPG" Or _
        UCase(Right(Filename1, 3)) = "PEG" Then
'PCN2834 now checked at save dialog
'        If Dir(Filename1) <> "" Then
'            Answer = MsgBox(DisplayMessage("File already exists. Will you overwrite?"), vbYesNo) 'PCN2111
'        End If

'        If Answer = vbYes Then
            Dim jm As Control
            Set jm = PipelineDetails.JPGMake1
           'PCN 1972 LS 8/7/03
            'jm.InputFile = App.Path & "\temp.bmp"
            If Len(strMediaFilePath) = 0 Then 'PCN2133
                jm.InputFile = LocToSave & "temp.bmp"
                jm.Quality = 80
                jm.OutputFile = Filename1
                jm.Go
            End If 'PCN2133
            'PCN 1972 LS 8/7/03
            'jm.InputFile = App.Path & "\temp2.bmp"
            If Len(strMediaFilePath) > 0 Then 'PCN2133
                jm.InputFile = LocToSave & "temp2.bmp" 'PCN2133
                jm.Quality = 80 'PCN2133
                jm.OutputFile = Left(Filename1, InStr(UCase(Filename1), ".J") - 1) & ".jpg" 'PCN2133
                jm.Go 'PCN2133
            Else 'PCN2133
                jm.InputFile = LocToSave & "temp2.bmp"
                jm.Quality = 80
                jm.OutputFile = Left(Filename1, InStr(UCase(Filename1), ".J") - 1) & "_Drawing.jpg"
                jm.Go
            End If 'PCN2133
'        Else
'            Exit Function
'        End If
    ElseIf UCase(Right(Filename1, 3)) = "BMP" Then
'
'        If Dir(Filename1) <> "" Then
'            Answer = MsgBox(DisplayMessage("File already exists. Will you overwrite?"), vbYesNo)  'PCN2111
'        End If
'        If Answer = vbYes Then
        'PCN 1972 LS 8/7/03
            If Len(strMediaFilePath) = 0 Then 'PCN2133
                'fs.MoveFile App.Path & "\temp.bmp", Filename1
                fs.MoveFile LocToSave & "temp.bmp", Filename1
            End If
            If Len(strMediaFilePath) > 0 Then 'PCN2133
                fs.MoveFile LocToSave & "temp2.bmp", Left(Filename1, InStr(UCase(Filename1), ".BMP") - 1) & ".bmp" 'PCN2133
            Else 'PCN2133
                'fs.MoveFile App.Path & "\temp2.bmp", left(Filename1, InStr(UCase(Filename1), ".BMP") - 1) & "_Drawing.bmp"
                fs.MoveFile LocToSave & "temp2.bmp", Left(Filename1, InStr(UCase(Filename1), ".BMP") - 1) & "_Drawing.bmp"
            End If 'PCN2133
'        End If
    Else
        MsgBox "Filename extension must be JPG, JPEG, or BMP", vbExclamation
        Exit Sub
    End If
    
    'PCN2133 --------------------------------------------------------------------------v
    Dim Parameter As String
    Dim PathName As String
    Dim HaltApp As Boolean
    If Len(Filename1) > 0 And Dir(LocToSave & "CLPInterface.int") <> "" Then  'PCN2185 'PCN2406
        Call INI_WriteBack(LocToSave & "CLPInterface.int", "MediaFilePath=", Filename1) 'PCN2123, PCN2176
        Call ValidatePath(Filename1, PathName, strMediaFileName, HaltApp, Parameter)
        strMediaFilePath = Filename1
        'Call INI_WriteBack(LocToSave & "CLPInterface.int", "ObsComments=", Observations.Recommendations.text)'PCN2123, PCN2176
    End If '---------------------------------------------------------------------------^
    'PCN 1972 LS 8/7/03
    'fs.deletefile (App.Path & "\temp.bmp")
    fs.deletefile (LocToSave & "temp.bmp")

    'Save Measurement 'commented out as per GL25.02.03
    Dim CBSFilenm As String
    If UCase(Right(Filename1, 4)) = ".JPG" Then
        CBSFilenm = Left(Filename1, Len(Filename1) - 3) & "CBS"
    ElseIf UCase(Right(Filename1, 5)) = ".JPEG" Then
        CBSFilenm = Left(Filename1, Len(Filename1) - 4) & ".CBS"
    ElseIf UCase(Right(Filename1, 4)) = ".BMP" Then
        CBSFilenm = Left(Filename1, Len(Filename1) - 3) & "CBS"
    Else
        'MsgBox "Filename is wrong. Please check filename. Measurement information is saved in Temporary.CBS file in CBS folder under application folder.", vbInformation 'PCN2111
        'CBSFilenm = App.Path & "\CBS\Temporary.CBS"
        'PCN 1972 LS 8/7/03
        'CBSFilenm = LocToSave & "CBS\Temporary.CBS"
        CBSFilenm = LocToSave & "Temporary.CBS" 'PCN4601
    End If
    
    ArrayCnt = 0
    ImageDataFile = CBSFilenm  'PCNGL1812022
    ReDim PipeInfoArray(30)
    
    Call SaveToFilePipeAndConfigInfo(CBSFilenm, True) 'PCNGL1812022 'PCNGL130103 'ML200203
  
'    ClearLineScreen.SaveThisData (CBSFilenm) 'PCN1944
    
Exit Sub
Err_Handler:
Select Case Err
    Case 58 'file already exist
'        answer = MsgBox("File already exists. Overwrite?", vbYesNo)
'        If answer = vbYes Then
            fs.deletefile (Filename1)
            Resume
'        Else
'            Exit Sub
'        End If
    Case 53 'file not found (this error occurs when h/w required doesn't exist.)
        Resume Next
    Case Else
        MsgBox Err & "-PF7:" & Error$
End Select

End Sub

Sub LoadPipelineDetails_V40(PipeInfoAddress As Long, ByVal FileNo)  'PCNGL1812022 'PCNGL120103
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'LoadPipelineDetails_V40 Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff,    12/01/03     Building initial framework
'Description:
'       Loads the pipeline details from file.
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

'When the PipelineInfo is no longer V4.0 then translate V4.0 into PipelineInfo
Get #FileNo, PipeInfoAddress, PipelineInfo


Exit Sub
Err_Handler:
    MsgBox Err & "-PF8:" & Error$
End Sub



Sub LoadConfigInfo_V40(ConfigInfoAddress As Long, ByVal FileNo)  'PCNGL130103
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'LoadConfigInfo_V40 Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff,    13/01/03     Building initial framework
'Description:
'       Loads the configuration information from file.
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

'When the ConfigInfo is no longer V4.0 then translate V4.0 into ConfigInfo
Get #FileNo, ConfigInfoAddress, ConfigInfo
PipeObsBuffer = 10 'There is 10 line buffer available for Pipe Observations in the PVD file. This ensures the file does not have to be completely re-writen when up to 10 observations are added. 'PCN2928
ConfigInfo.FishEyeOriginalHeight = ConfigInfo.MediaHeight 'PCN3019
ConfigInfo.FishEyeOriginalWidth = ConfigInfo.MediaWidth 'PCN3019

Exit Sub
Err_Handler:
    MsgBox Err & "-PF9:" & Error$
End Sub

Sub LoadConfigInfo_V50(ConfigInfoAddress As Long, FileLoadError As Boolean, ByVal FileNo)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : LoadConfigInfo_V50
'Created : 19 March 2004, PCN2639
'Updated :
'Prg By  : Geoff Logan
'Param   : ConfigInfoAddress
'Desc    : Loads the configuration information from file.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Get #FileNo, ConfigInfoAddress, ConfigInfo 'Note: Even though this ConfigInfo is not the same as V5X, the new parameter

Select Case ConfigInfo.PVDFileVersion
    'vvvv PCN2850 ****************************************************
    '!!!!!!!!! Note: When updating ConfigInfo version !!!!!!!!!!!!!!!!
    '!!!!!!!!! Remember to also update                !!!!!!!!!!!!!!!!
    '!!!!!!!!! SaveToFilePipeAndConfigInfo and        !!!!!!!!!!!!!!!!
    '!!!!!!!!! SaveInPVDFormat_VXX                    !!!!!!!!!!!!!!!!
    '^^^^ ************************************************************
    'vvvv PCN2928 *********************************************
    Case "V5.3"
        PVCalculationsMultiplier = 100
        PVDataFrameBlockSize = PVDataFrameBlockSize_V50
        PVCalculationsBlockSize = PVCalculationsBlockSize_V50
        PVRelatedInfoBlockSize = PVRelatedInfoBlockSize_V50
        DistanceMethod = Trim(ConfigInfo.DistanceProcessMethod)
        DistanceStart = ConfigInfo.DistanceStart
        PipeObsBuffer = 100 'There is 100 line buffer available for Pipe Observations in the PVD file. This ensures the file does not have to be completely re-writen when up to 100 observations are added.
        
    '^^^^ *****************************************************
        ConfigInfo.ProfileRecordingMethod = "Radius" 'PCN2891 'PCN2952
    Case "V5.2"
        PVCalculationsMultiplier = 100 'PCN2829
        PVDataFrameBlockSize = PVDataFrameBlockSize_V50
        PVCalculationsBlockSize = PVCalculationsBlockSize_V50
        PVRelatedInfoBlockSize = PVRelatedInfoBlockSize_V50
        DistanceMethod = Trim(ConfigInfo.DistanceProcessMethod)
        DistanceStart = ConfigInfo.DistanceStart
        ConfigInfo.ProfileRecordingMethod = "Radius" 'PCN2891
        PipeObsBuffer = 10 'There is 10 line buffer available for Pipe Observations in the PVD file. This ensures the file does not have to be completely re-writen when up to 10 observations are added. 'PCN2928
    Case "V5.1"
        PVCalculationsMultiplier = 1 'PCN2829
        PVDataFrameBlockSize = PVDataFrameBlockSize_V50
        PVCalculationsBlockSize = PVCalculationsBlockSize_V50
        PVRelatedInfoBlockSize = PVRelatedInfoBlockSize_V50
        DistanceMethod = Trim(ConfigInfo.DistanceProcessMethod)
        DistanceStart = ConfigInfo.DistanceStart
        ConfigInfo.ProfileRecordingMethod = "Radius" 'PCN2891
        PipeObsBuffer = 10 'There is 10 line buffer available for Pipe Observations in the PVD file. This ensures the file does not have to be completely re-writen when up to 10 observations are added. 'PCN2928
    Case "V5.0"
        PVCalculationsMultiplier = 1 'PCN2829
        PVDataFrameBlockSize = PVDataFrameBlockSize_V50
        PVCalculationsBlockSize = PVCalculationsBlockSize_V50
        PVRelatedInfoBlockSize = PVRelatedInfoBlockSize_V50
        DistanceMethod = Trim(ConfigInfo.DistanceProcessMethod)
        DistanceStart = ConfigInfo.DistanceStart
        Call GetINI_ImageProcessInfo(MyFile, 2) 'PCN2820
        Call GetINI_LimitLineInfo(MyFile, 3) 'PCN2820
        ConfigInfo.ProfileRecordingMethod = "Radius" 'PCN2891
        PipeObsBuffer = 10 'There is 10 line buffer available for Pipe Observations in the PVD file. This ensures the file does not have to be completely re-writen when up to 10 observations are added. 'PCN2928
    Case Else
        FileLoadError = True 'Not this version
End Select

ConfigInfo.FishEyeOriginalHeight = ConfigInfo.MediaHeight 'PCN3019
ConfigInfo.FishEyeOriginalWidth = ConfigInfo.MediaWidth 'PCN3019


Exit Sub
Err_Handler:
    MsgBox Err & "-PF10:" & Error$
End Sub


Sub LoadConfigInfo_V6X(ConfigInfoAddress As Long, FileLoadError As Boolean, ByVal FileNo)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : LoadConfigInfo_V6X
'Created : 19 March 2004, PCN2891
'Updated :
'Prg By  : Geoff Logan
'Param   : ConfigInfoAddress
'Desc    : Loads the configuration information from file.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim PVDVer As Single

Get #FileNo, ConfigInfoAddress, ConfigInfo



PVDVer = GetPVDVer
Select Case ConfigInfo.PVDFileVersion
    'vvvv PCN2850 ****************************************************
    '!!!!!!!!! Note: When updating ConfigInfo version !!!!!!!!!!!!!!!!
    '!!!!!!!!! Remember to also update                !!!!!!!!!!!!!!!!
    '!!!!!!!!! SaveToFilePipeAndConfigInfo and        !!!!!!!!!!!!!!!!
    '!!!!!!!!! SaveInPVDFormat_VXX                    !!!!!!!!!!!!!!!!
    '^^^^ ************************************************************
    Case "V6.1", "V6.2", "6.25", "V6.3", "V6.4"   'PCN3019 'PCN3576 'PCN4006
        
        'Catching possible corrupted file 'ID5395, we wont corruption let happen
        'If Trim(ConfigInfo.ProfileRecordingMethod) <> "XY" Or Trim(ConfigInfo.FishEyeOriginalWidth) <> "720" Then
            'Get #FileNo, ConfigInfoAddress, ConfigInfo_currupting
            'PossibleConfigInfoCurruption = True
            'CopyFromCurruptionToGoodConfigInfo
        'End If 'Its allways XY
        
        PVCalculationsMultiplier = PVCalculationsMultiplier_V60
        PVDataFrameBlockSize = PVDataFrameBlockSize_V60
        PVCalculationsBlockSize = PVCalculationsBlockSize_V60
        If PVDVer >= 6.3 Then
            PVDataFrameBlockSize = PVDataFrameBlockSize_V70 'PCN4006
            ShapeCentreX = ConfigInfo.PVShapeCentreX 'PCN4336
            ShapeCentreY = ConfigInfo.PVShapeCentreY 'PCN4336
            If ShapeCentreX > ClearLineScreen.PVScreen.width / 2 Or _
               ShapeCentreY > ClearLineScreen.PVScreen.height / 2 Then ShapeCentreX = 0: ShapeCentreY = 0
        End If
        PVRelatedInfoBlockSize = PVRelatedInfoBlockSize_V60
        DistanceMethod = Trim(ConfigInfo.DistanceProcessMethod)
        DistanceStart = ConfigInfo.DistanceStart
        PVDrawScreenRatio = ConfigInfo.Ratio
        
        'Determine VideoScreenScale
'PCN3021 these variable assingments are no longer needed
'        MediaHeight = ConfigInfo.MediaHeight
'        MediaWidth = ConfigInfo.MediaWidth
        Call ClearLineScreen.VideoScreenScaleCalc 'This may not be required in new code.
        PipeObsBuffer = 100 'There is 100 line buffer available for Pipe Observations in the PVD file. This ensures the file does not have to be completely re-writen when up to 100 observations are added. 'PCN2928
    Case "V6.0"
        PVCalculationsMultiplier = PVCalculationsMultiplier_V60
        PVDataFrameBlockSize = PVDataFrameBlockSize_V60
        PVCalculationsBlockSize = PVCalculationsBlockSize_V60
        PVRelatedInfoBlockSize = PVRelatedInfoBlockSize_V60
        DistanceMethod = Trim(ConfigInfo.DistanceProcessMethod)
        DistanceStart = ConfigInfo.DistanceStart
        'Determine VideoScreenScale
'PCN3021 these variable assingments are no longer needed
'        MediaHeight = ConfigInfo.MediaHeight
'        MediaWidth = ConfigInfo.MediaWidth
        Call ClearLineScreen.VideoScreenScaleCalc 'This may not be required in new code.
        PipeObsBuffer = 100 'There is 100 line buffer available for Pipe Observations in the PVD file. This ensures the file does not have to be completely re-writen when up to 100 observations are added. 'PCN2928
        ConfigInfo.FishEyeOriginalHeight = ConfigInfo.MediaHeight 'PCN3019
        ConfigInfo.FishEyeOriginalWidth = ConfigInfo.MediaWidth 'PCN3019
    Case Else
        ConfigInfo.ProfileRecordingMethod = "Radius"
        FileLoadError = True 'Not this version
End Select

Exit Sub
Err_Handler:
    MsgBox Err & "-PF11:" & Error$
End Sub

Sub LoadConfigInfo_V41(ConfigInfoAddress As Long, FileLoadError As Boolean, ByVal FileNo)  'PCNGL130103 ' PCN2952
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : LoadConfigInfo_V41
'Created : 11 November 2003, PCN2392
'Updated : 7 August 2004, PCN2952
'Prg By  : Geoff Logan
'Param   : ConfigInfoAddress
'          FileLoadError - Set to true if incorrect version.
'Desc    : Loads the configuration information from file.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim strTemp1 As String 'PCN2952
Dim strTemp2 As String 'PCN2952
Dim StrLoc As Integer 'PCN2952

Get #FileNo, ConfigInfoAddress, ConfigInfo_V41 'PCN2639
'Check if V4.0 then if so, translate V4.0 into ConfigInfo
If Trim(ConfigInfo.PVDFileVersion) <> "V4.1" Then
    'vvvv PCN2952 **************************
    If Left(ConfigInfo.PVDFileVersion, 1) <> "V" Then
        'PCN3019
        Trim(ConfigInfo_V41.PVDFileVersion) = "V4.0" 'PCN3019
        ConfigInfo_V41.FishEyeDistortion = 0 'PCN3019
    Else
        'This version is greater than this application can read
        strTemp1 = DisplayMessage("ClearLine.ini VERSION ERROR. Expecting ")
        StrLoc = InStr(1, strTemp1, "ClearLine.ini")
        strTemp2 = Left(strTemp1, StrLoc - 1)
        strTemp1 = Right(strTemp1, Len(strTemp1) - (StrLoc + 12))
        strTemp1 = strTemp2 & "ClearLine PVD" & strTemp1 & PVDVersion
        strTemp1 = strTemp1 & ". " & DisplayMessage("This file is not loaded.")
        MsgBox strTemp1
        FileLoadError = True
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    '^^^^ **********************************
End If
Call LoadConfigInfo_V41IntoConfigInfo 'PCN3019
ConfigInfo.DistanceProcessMethod = "None" 'PCN2639
DistanceMethod = "None" 'PCN2639
PVDataFrameBlockSize = PVDataFrameBlockSize_V40 'PCN2639
PVCalculationsBlockSize = PVCalculationsBlockSize_V40 'PCN2639
PVRelatedInfoBlockSize = PVRelatedInfoBlockSize_V40 'PCN2639
Call GetINI_ImageProcessInfo(MyFile, 2) 'PCN2820
Call GetINI_LimitLineInfo(MyFile, 3) 'PCN2820
PVCalculationsMultiplier = 1 'PCN2829
ConfigInfo.ProfileRecordingMethod = "Radius" 'PCN2891
PipeObsBuffer = 10 'There is 10 line buffer available for Pipe Observations in the PVD file. This ensures the file does not have to be completely re-writen when up to 10 observations are added. 'PCN2928
ConfigInfo.FishEyeOriginalHeight = ConfigInfo.MediaHeight 'PCN3019
ConfigInfo.FishEyeOriginalWidth = ConfigInfo.MediaWidth 'PCN3019

Exit Sub
Err_Handler:
    MsgBox Err & "-PF12:" & Error$
End Sub
Sub LoadPipelineObs_V60(PipeObsAddress As Long, NoOfPipeObsLines As Long, ByVal FileNo)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCN     : 3576
'Name    : LoadPipelineObs_V60
'Created : 13 July 2005 July 2004
'Updated : Updated from 27 July 2004 LoadPipeLineObs_V50 PCN2928
'Prg By  : Geoff Logan
'Param   : PipeObsAddress - Pipe Obs PVD starting Address
'          NoOfPipeObsLines - No Of Pipe Obs Lines
'Desc    : Loads the pipeline observations from file.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim ObsLineNo As Integer
Dim LoadObs As String 'Used in check import data


ReDim PipeObservations(0)
Dim PipeObservationsUserTitle As PipeObservationType_V60
Dim FoundNoObsFlag As Boolean

NoOfPipeObservations = 0

FoundNoObsFlag = False

For ObsLineNo = 1 To PipeObsBuffer 'PCN2928
    
    If FoundNoObsFlag = False Then 'Run while obs are found in the PVD
        ReDim Preserve PipeObservations(ObsLineNo)
        Get #FileNo, , PipeObservations(ObsLineNo)
    
        If Left(PVDHeaderPipeObs.PVDHeaderDescriptor, 13) <> "[PipeObs] V60" Then
            PipeObservations(ObsLineNo).PipeObsSnapshotLength = 0
            PipeObservations(ObsLineNo).PipeObsSnapshotOffset = 0
        End If
    
        LoadObs = PipeObservations(ObsLineNo).PipeObs

        If InStr(1, LoadObs, "No Observation") <> 0 Then
            NoOfPipeObservations = ObsLineNo - 1
            ReDim Preserve PipeObservations(NoOfPipeObservations)
            FoundNoObsFlag = True
        End If
        If ObsLineNo = 96 Then FoundNoObsFlag = True: NoOfPipeObservations = 96
    Else
        Get #FileNo, , PipeObservationsUserTitle
        Select Case ObsLineNo
            Case 97: UserTitleAnalysis = Trim(Right(PipeObservationsUserTitle.PipeObs, Len(PipeObservationsUserTitle.PipeObs) - 14))
            Case 98: UserTitleObservations = Trim(Right(PipeObservationsUserTitle.PipeObs, Len(PipeObservationsUserTitle.PipeObs) - 14))
            Case 99: UserTitleSummary = Trim(Right(PipeObservationsUserTitle.PipeObs, Len(PipeObservationsUserTitle.PipeObs) - 14))
            Case 100: UserTitleProfile = Trim(Right(PipeObservationsUserTitle.PipeObs, Len(PipeObservationsUserTitle.PipeObs) - 14))
        End Select
    End If
Next ObsLineNo





Exit Sub
Err_Handler:
    MsgBox Err & "-PF13:" & Error$
End Sub
Sub LoadPipelineObs_V50(PipeObsAddress As Long, NoOfPipeObsLines As Long, ByVal FileNo)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : LoadPipelineObs_V50
'Created : 27 July 2004, PCN2928
'Updated :
'Prg By  : Geoff Logan
'Param   : PipeObsAddress - Pipe Obs PVD starting Address
'          NoOfPipeObsLines - No Of Pipe Obs Lines
'Desc    : Loads the pipeline observations from file.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim ObsLineNo As Integer
Dim LoadObs As String 'Used in check import data

ReDim PipeObservations(0)

For ObsLineNo = 1 To PipeObsBuffer - 4 'PCN2928 'PCN4433
    ReDim Preserve PipeObservations(ObsLineNo)
    Get #FileNo, , PipeObservations(ObsLineNo)
    LoadObs = PipeObservations(ObsLineNo).PipeObs
'    Debug.Print PipeObservations(ObsLineNo).PipeObsFrameNo & ", " & PipeObservations(ObsLineNo).PipeObsDist & ", " & PipeObservations(ObsLineNo).PipeObs
    If InStr(1, LoadObs, "No Observation") <> 0 Then
'        NoOfPipeObsLines = ObsLineNo - 1
        NoOfPipeObservations = ObsLineNo - 1
        ReDim Preserve PipeObservations(NoOfPipeObservations)
        Exit Sub
    End If
Next ObsLineNo



Exit Sub
Err_Handler:
    MsgBox Err & "-PF14:" & Error$
End Sub


Sub LoadPipelineObs_V40(PipeObsAddress As Long, NoOfPipeObsLines As Long, ByVal FileNo)  'PCNGL130103
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'LoadPipelineObs_V40 Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff,    13/01/03     Building initial framework
'Description:
'       Loads the pipeline observations from file.
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim ObsLineNo As Integer
Dim LoadObs As String 'Used in check import data

'When the PipeObservations is no longer V4.0 then translate V4.0 into PipeObservations
ReDim PipeObservations(NoOfPipeObsLines)
Get #FileNo, PipeObsAddress, PipeObservations(0) 'Get the first line
'Check if the first line contains valid data (ie not "No Observations")
LoadObs = PipeObservations(0).PipeObsDist
If InStr(1, LoadObs, "No Observation") <> 0 Then 'PCN3000
    NoOfPipeObsLines = 0
    Exit Sub
End If

For ObsLineNo = 1 To NoOfPipeObsLines
    Get #FileNo, , PipeObservations(ObsLineNo)
    LoadObs = PipeObservations(ObsLineNo).PipeObsDist
    If InStr(1, LoadObs, "No Observation") <> 0 Then 'PCN3000
        NoOfPipeObsLines = ObsLineNo
        Exit Sub
    End If
Next ObsLineNo


Exit Sub
Err_Handler:
    MsgBox Err & "-PF15:" & Error$
End Sub

Sub InitialiseFieldsOnForms()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'InitialiseFieldsOnForms Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff,    17/01/03     Building initial framework
'
'Description:
'       This clears the various form fields.
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

ClearLineTitle.TitleBarCaption.Caption = DisplayMessage("Precision Vision") 'PCN2111 'PCN2759
ClearLineTitle.TitleBarCaption.ToolTipText = DisplayMessage("Precision Vision") 'PCN4171

PipelineDetails.InternalDiameterExpected = ""
PipelineDetails.OutsideDiameter = "" 'PCN1837
PipelineDetails.PipeLength = ""
'PipelineDetails.Len_Real = 0
'PipelineDetails.LenRealPercent = 0

PipelineDetails.Material = ""

'PCN4131
'Observations.AssetNo = ""
'Observations.SiteID = ""
'Observations.StartNodeNo = ""
'Observations.FinishNodeNo = ""
'Observations.InternalDiameterExpected = ""
'Observations.OutsideDiameter = ""
'Observations.Recommendations = ""
'Observations.Distance = ""
'Observations.Observation = ""
'vvvv PCN4179 **********************************
PipelineInfo.AssetNo = ""
PipelineInfo.SiteID = ""
PipelineInfo.City = ""
PipelineInfo.Date = 0
PipelineInfo.Time = 0
PipelineInfo.StartName = ""
PipelineInfo.StartLocation = ""
PipelineInfo.FinishName = ""
PipelineInfo.FinishLocation = ""
PipelineInfo.Material = ""
PipelineInfo.PipeLength = 0
PipelineInfo.ExtDiameter = 0
PipelineInfo.IntDiameter = 0
PipelineInfo.Comments = ""
Call CopyPipeDetailsToPipelineForm
'^^^^ ******************************************

Exit Sub
Err_Handler:
    MsgBox Err & "-PF16:" & Error$

End Sub

Sub CopyPipeDetailsToPipelineForm() 'PCNGL130103
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'CopyPipeDetailsToPipelineForm Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff,    11/12/02     Building initial framework
'
'Description:
'       Copies the data in PipelineInfo onto the PipelineDetails forms
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

If PipelineInfo.IntDiameter <> 0 Then 'PCNGL140103
    PipelineDetails.InternalDiameterExpected = PipelineInfo.IntDiameter
    ExpectedDiameter = PipelineInfo.IntDiameter 'PCN3647
Else
    PipelineDetails.InternalDiameterExpected = "" 'PCNGL140103
    ExpectedDiameter = 0 'PCN3647
End If
If PipelineInfo.ExtDiameter <> 0 Then 'PCN1832
    PipelineDetails.OutsideDiameter = PipelineInfo.ExtDiameter 'PCN1832
Else
    PipelineDetails.OutsideDiameter = "" 'PCN1832
End If
If PipelineInfo.PipeLength <> 0 Then 'PCNGL140103 'PCN1832
    PipelineDetails.PipeLength = PipelineInfo.PipeLength
Else
    PipelineDetails.PipeLength = "" 'PCNGL140103
End If

If LanguageCharset <> 0 Then
    PipelineDetails.Material.Font.Charset = LanguageCharset
    PipelineDetails.SiteID.Font.Charset = LanguageCharset
    PipelineDetails.City.Font.Charset = LanguageCharset
    PipelineDetails.StartNodeNo.Font.Charset = LanguageCharset
    PipelineDetails.StartNodeLocation.Font.Charset = LanguageCharset
    PipelineDetails.FinishNodeNo.Font.Charset = LanguageCharset
    PipelineDetails.FinishNodeLocation.Font.Charset = LanguageCharset
    PipelineDetails.GeneralComments.Font.Charset = LanguageCharset
    PipelineDetails.AssetNo.Font.Charset = LanguageCharset
End If

PipelineDetails.Material = Trim(PipelineInfo.Material) 'PCNGL290103
PipelineDetails.SiteID = Trim(PipelineInfo.SiteID) 'PCNGL290103
PipelineDetails.City = Trim(PipelineInfo.City) 'PCNGL290103
If PipelineInfo.Date <> 0 Then 'PCNGL140103
    PipelineDetails.sDate = Format(PipelineInfo.Date, "Short Date") 'PCNGL140103
Else
    PipelineDetails.sDate = Format(Date, "Short Date") 'PCNGL140103
End If
If PipelineInfo.Time <> 0 Then 'PCNGL140103
    PipelineDetails.sTime = Format(PipelineInfo.Time, "Short time") 'PCNGL140103
Else
    PipelineDetails.sTime = Format(Time, "Short Time") 'PCNGL140103
End If
PipelineDetails.StartNodeNo = Trim(PipelineInfo.StartName) 'PCNGL290103
PipelineDetails.StartNodeLocation = Trim(PipelineInfo.StartLocation) 'PCNGL290103
PipelineDetails.FinishNodeNo = Trim(PipelineInfo.FinishName) 'PCNGL290103
PipelineDetails.FinishNodeLocation = Trim(PipelineInfo.FinishLocation) 'PCNGL290103

PipelineDetails.GeneralComments.text = Trim(PipelineInfo.Comments)  'PCN4171
PipelineDetails.AssetNo = Trim(PipelineInfo.AssetNo) 'PCNGL290103

Exit Sub
Err_Handler:
    MsgBox Err & "-PF17:" & Error$
End Sub


Sub CopyConfigInfoToForms() 'PCNGL130103
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'CopyConfigInfoToForms Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff,    11/12/02     Building initial framework
'
'Description:
'       Copies the data in ConfigInfo onto the required forms
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

'PCNGL150103 Move this code from the test function
'PipelineDetails.Angl1 = ConfigInfo.WLStartAngle
'PipelineDetails.Angl2 = ConfigInfo.WLFinishAngle
WLStartAngle = ConfigInfo.WLStartAngle
WLFinishAngle = ConfigInfo.WLFinishAngle


'ClearLineScreen.CalLen = ConfigInfo.CalDist
CalLen_Global = ConfigInfo.CalDist

'ClearLineScreen.CalLength_tmp = ConfigInfo.CalLineLength
CalLength_Global = ConfigInfo.CalLineLength

'PipelineDetails.Len_Real = ConfigInfo.LenReal
'PipelineDetails.LenRealPercent = ConfigInfo.LenRealPercent
'ClearLineScreen.Ratio = ConfigInfo.Ratio 'PCN3035

'vvvv PCN2973 **********************************************************
'The CalLineLength is sometimes being corrupted. Its value can be
'restored by using Ratio and CalDist.
Dim CalculatedCalLength_Global As Single

If ConfigInfo.Ratio <> 0 Then
    CalculatedCalLength_Global = CalLen_Global / ConfigInfo.Ratio
    'Check if the CalculatedCalLength_Global is same/similar to the CalLength_Global
    If Round(CalculatedCalLength_Global, 0) <> Round(CalLength_Global, 0) Then
        'ConfigInfo.CalLineLength may have been corrupted. Let's recalculate it.
        ConfigInfo.CalLineLength = CalculatedCalLength_Global
        'ClearLineScreen.CalLength_tmp = ConfigInfo.CalLineLength
        CalLength_Global = ConfigInfo.CalLineLength
    End If
End If
'^^^^ ******************************************************************

PipelineDetails.unit1.Caption = Trim(ConfigInfo.Units) 'PCNGL290103
PipelineDetails.unit2.Caption = Trim(ConfigInfo.Units) 'PCNGL290103
PipelineDetails.unit3.Caption = Trim(ConfigInfo.Units) 'PCNGL290103

'PCN4443 at last found where it was set.
If ConfigInfo.Units = "mm" Then
    PipelineDetails.Unit6.Caption = "m"
    PipelineDetails.lblPipeLengthUnit.Caption = "m"
    
Else
    PipelineDetails.Unit6.Caption = "ft"
    PipelineDetails.lblPipeLengthUnit.Caption = "ft"
End If

PipelineDetails.seaLevelUnitLabel(0).Caption = PipelineDetails.lblPipeLengthUnit.Caption 'PCN6128
PipelineDetails.seaLevelUnitLabel(1).Caption = PipelineDetails.lblPipeLengthUnit.Caption 'PCN6128

NoOfProfileSegments = ConfigInfo.NoOfProfileSegments 'PCNGL140103

VideoFileName = Trim(ConfigInfo.VideoFileName) 'PCNGL150103 'PCNGL280103
If Dir(VideoFileName) = "" Or VideoFileName = "" Then VideoFileName = FindVideoFile(PVDFileName, VideoFileName)

'PCN3021 these variable assingments are no longer needed
'MediaHeight = ConfigInfo.MediaHeight  'PCN1833
'MediaWidth = ConfigInfo.MediaWidth  'PCN1833

Exit Sub
Err_Handler:
    Select Case Err
        Case 52: VideoFileName = FindVideoFile(PVDFileName, VideoFileName): Exit Sub
        Case Else: MsgBox Err & "-PF18:" & Error$
    End Select
End Sub

Sub LoadFontInfo_V40(FontInfoAddress As Long, ByVal FileNo)
Get #FileNo, FontInfoAddress, FontInfo

End Sub

Sub SaveFontInfo_V40(ByVal FileNo)  'PCNGL120103
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'SaveFontInfo_V40 Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff,    12/01/03     Building initial framework
'Description:
'       Saves the Font details to file.
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

FontInfo.FontName = FontName
FontInfo.FontType = FontType
FontInfo.FontSize = FontSize
FontInfo.FontColour = FontColour

Put #FileNo, , FontInfo

Exit Sub
Err_Handler:
    MsgBox Err & "-PF19:" & Error$
End Sub

Sub SaveObservations_V40(NoOfObsLines As Integer, ByVal FileNo)  'PCNGL1812022 'PCNGL120103
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'SaveObservations_V40 Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff,    12/01/03     Building initial framework
'Description:
'       Saves the observations to file.
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim ObsLineNo As Integer
Dim BlankObsLineNo As Integer

If NoOfObsLines = 0 Then 'PCNGL130103
    'Put 10 lines of blank observations into the file. This ensures the file does not have to be completely re-writen when up to 10 observations are added.
    ReDim PipeObservations(10)
    PipeObservations(1).PipeObsDist = 0 'PCN2928
    PipeObservations(1).PipeObs = "No Observation" 'Used to confirm to the loading program that this observation is blank 'PCN2928
End If

For ObsLineNo = 1 To NoOfObsLines
    Put #FileNo, , PipeObservations(ObsLineNo)
Next ObsLineNo

'Fill in blank observations up to 10 obs
BlankObsLineNo = ObsLineNo
For ObsLineNo = BlankObsLineNo To 10
    ReDim Preserve PipeObservations(10)
    PipeObservations(ObsLineNo).PipeObsDist = 0
    PipeObservations(ObsLineNo).PipeObs = "No Observation" 'Used to confirm to the loading program that this observation is blank
    Put #FileNo, , PipeObservations(ObsLineNo)
Next ObsLineNo

Exit Sub
Err_Handler:
    MsgBox Err & "-PF20:" & Error$
End Sub

Sub SaveObservations_V50(NoOfObsLines As Integer, FileNo As Integer)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : SaveObservations_V50
'Created : 27 July 2004, PCN2928
'Updated :
'Prg By  : Geoff Logan
'Param   : NoOfObsLines - Number Of Obs Lines
'Desc    : Saves the observations to file.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim ObsLineNo As Integer
Dim BlankObsLineNo As Integer

If PipeObsBuffer < 10 Then Exit Sub      'The buffer has not been setup corrently. DON'T write to the PVD file.

ReDim Preserve PipeObservations(PipeObsBuffer)

If NoOfObsLines = 0 Then
    PipeObservations(1).PipeObsDist = 0
    PipeObservations(1).PipeObs = "No Observation" 'Used to confirm to the loading program that this observation is blank 'PCN2928
End If

If NoOfObsLines < PipeObsBuffer - 4 Then 'PCN4433
    For ObsLineNo = 1 To NoOfObsLines
        Put #FileNo, , PipeObservations(ObsLineNo)
    Next ObsLineNo
    
    'Fill in blank observations up to PipeObsBuffer observations
    BlankObsLineNo = ObsLineNo
    For ObsLineNo = BlankObsLineNo To PipeObsBuffer - 4 'PCN4433
        PipeObservations(ObsLineNo).PipeObsDist = 0
        PipeObservations(ObsLineNo).PipeObs = "No Observation" 'Used to confirm to the loading program that this observation is blank
        Put #FileNo, , PipeObservations(ObsLineNo)
    Next ObsLineNo
    
    ReDim Preserve PipeObservations(NoOfObsLines) 'Reset to original setting
Else
    'The NoOfObsLines have exceeded the maximum allowable PipeObsBuffer.
    NoOfObsLines = PipeObsBuffer - 4 'PCN4433
    For ObsLineNo = 1 To NoOfObsLines
        Put #FileNo, , PipeObservations(ObsLineNo)
    Next ObsLineNo
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-PF21:" & Error$
End Sub

Sub SaveObservations_V60(NoOfObsLines As Integer, FileNo As Integer)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : SaveObservations_V50
'Created : 27 July 2004, PCN2928
'Updated :
'Prg By  : Geoff Logan
'Param   : NoOfObsLines - Number Of Obs Lines
'Desc    : Saves the observations to file.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim ObsLineNo As Integer
Dim BlankObsLineNo As Integer
Dim PipeObsAddressData As Long
Dim PipeObsAddressHeader As Long
Dim PipeObservationsUserTitle As PipeObservationType_V60

If PipeObsBuffer < 10 Then Exit Sub      'The buffer has not been setup corrently. DON'T write to the PVD file.

'If the header came from a PVD file older than V6.2 eg any obs header older than V60 then
'convert it to
If Left(PVDHeaderPipeObs.PVDHeaderDescriptor, 13) <> "[PipeObs] V60" Then
    PipeObsAddressHeader = PVDFilePointers.PVDPointerPipeObs
    PipeObsAddressData = Seek(FileNo)
    Seek #FileNo, PipeObsAddressHeader
    
    PVDHeaderPipeObs.PVDHeaderDescriptor = "[PipeObs] V60"
    Put #FileNo, , PVDHeaderPipeObs
    Seek #FileNo, PipeObsAddressData
End If
    
    


ReDim Preserve PipeObservations(PipeObsBuffer)

If NoOfObsLines = 0 Then
    PipeObservations(1).PipeObsDist = 0
    PipeObservations(1).PipeObs = "No Observation" 'Used to confirm to the loading program that this observation is blank 'PCN2928
    PipeObservations(1).PipeObsFrameNo = 0        'PCN3576
    PipeObservations(1).PipeObsSnapshotOffset = 0 '   "
    PipeObservations(1).PipeObsSnapshotLength = 0 '   "
End If

If NoOfObsLines < PipeObsBuffer - 4 Then 'PCN4433
    For ObsLineNo = 1 To NoOfObsLines
        Put #FileNo, , PipeObservations(ObsLineNo)
    Next ObsLineNo
    
    'Fill in blank observations up to PipeObsBuffer observations
    BlankObsLineNo = ObsLineNo
    For ObsLineNo = BlankObsLineNo To PipeObsBuffer - 4 'PCN4433
        PipeObservations(ObsLineNo).PipeObsDist = 0
        PipeObservations(ObsLineNo).PipeObs = "No Observation" 'Used to confirm to the loading program that this observation is blank
        PipeObservations(ObsLineNo).PipeObsFrameNo = 0        'PCN3576
        PipeObservations(ObsLineNo).PipeObsSnapshotOffset = 0 '   "
        PipeObservations(ObsLineNo).PipeObsSnapshotLength = 0 '   "
        Put #FileNo, , PipeObservations(ObsLineNo)
    Next ObsLineNo
    
    ReDim Preserve PipeObservations(NoOfObsLines) 'Reset to original setting
Else
    'The NoOfObsLines have exceeded the maximum allowable PipeObsBuffer.
    NoOfObsLines = PipeObsBuffer - 4 'PCN4433
    For ObsLineNo = 1 To NoOfObsLines
        Put #FileNo, , PipeObservations(ObsLineNo)
    Next ObsLineNo
End If

''''''''''''' Tagging the report titles at the end of the obs, obs were 1 to 100, not the
''''''''''''' are 1 to 96 and the user titles are from 97 to 100

PipeObservationsUserTitle.PipeObsDist = 0
PipeObservationsUserTitle.PipeObs = "" 'Used to confirm to the loading program that this observation is blank
PipeObservationsUserTitle.PipeObsFrameNo = 0        'PCN3576
PipeObservationsUserTitle.PipeObsSnapshotOffset = 0 '   "
PipeObservationsUserTitle.PipeObsSnapshotLength = 0 '   "

PipeObservationsUserTitle.PipeObs = "No Observation" & " " & Trim(UserTitleAnalysis): Put #FileNo, , PipeObservationsUserTitle
PipeObservationsUserTitle.PipeObs = "No Observation" & " " & Trim(UserTitleObservations): Put #FileNo, , PipeObservationsUserTitle
PipeObservationsUserTitle.PipeObs = "No Observation" & " " & Trim(UserTitleSummary): Put #FileNo, , PipeObservationsUserTitle
PipeObservationsUserTitle.PipeObs = "No Observation" & " " & Trim(UserTitleProfile): Put #FileNo, , PipeObservationsUserTitle


Exit Sub
Err_Handler:
    MsgBox Err & "-PF22:" & Error$
End Sub

Sub SaveInPVDFormat(PVDFileFormatVersion As String, SaveToFileName As String, FileErrorReturned As Boolean)  'PCNGL1812022 'PCNGL110103
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : GetAllFormsInfo
'Created : 18 December 2002, PCNGL110301
'Updated : 22 March 2004, PCN2639
'Prg By  : Geoff Logan
'Param   : PVDFileFormatVersion - version number
'          FileLoadError - Returns a true value if this file creation process was not successful.
'Desc    : Saves the Precision Vision data to file, in binary format.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
'Dim SaveMsg As Variant PCN1916 - Not used
Dim DrawFrameNo As Long
Dim FrameBufferNo  As Integer 'The size of the PVData frame buffer 'PCNGL140103


''Ask user to determine the Save file's name and path
'SaveToFileName = PVDFileName 'PCNGL140103

Call CopyPipelineFormToPipeDetails 'PCN1833
Call CopyFormToConfigInfo 'PCN1833

Select Case PVDFileFormatVersion
    Case "6.X" 'PCN2891
        'Save in the current file format
        Call SaveInPVDFormat_V6X(SaveToFileName, FileErrorReturned) 'PCN2891
    Case "5.0" 'PCN2639
        'Save in the previous file format
        Call SaveInPVDFormat_V50(SaveToFileName, FileErrorReturned) 'PCN2639
    Case "4.0"
        'Save in previous file format 'PCN2639
        Call SaveInPVDFormat_V40(SaveToFileName, FileErrorReturned)
    Case Else
        'Save in the current file format
        Call SaveInPVDFormat_V6X(SaveToFileName, FileErrorReturned) 'PCN2639 'PCN3000
End Select


 
Exit Sub
Err_Handler:
    Select Case Err
        Case 9 ' out of range
        Case Else
            MsgBox Err & "-PF23:" & Error$
    End Select
End Sub

Sub SaveInPVDFormat_V40(FileName As String, FileLoadError As Boolean) 'PCNGL140103    'PCNGL1812022 'PCNGL110103
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'SavePrecisionVisionData Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff,    11/01/03     Building initial framework
'
'Description:
'       Saves the Precision Vision data to file, in binary format, version 4.0 (PCNGL110301)
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim PVSegmentNo As Integer
Dim PVDataLineNo As Long
'Dim NoOfPipeObservations As Integer 'PCNGL130103 PCN2928

Trim(ConfigInfo.PVDFileVersion) = "V4.1" 'PCN2639
PipeObsBuffer = 10 'PCN2928

'Convert ConfigInfo to ConfigInfo_V41 'PCN2639
Call ConvertConfigInfoToConfigInfo_V41 'PCN2639

FileLoadError = False 'PCNGL140103

PVDataFrameBlockSize = PVDataFrameBlockSize_V40 'PCN2639
PVCalculationsBlockSize = PVCalculationsBlockSize_V40 'PCN2639
PVRelatedInfoBlockSize = PVRelatedInfoBlockSize_V40 'PCN2639

Dim FileNo As Integer
FileNo = FreeFile


Open FileName For Binary Access Write As #FileNo 'ID5384 #1

'Compile and save the File Main Header
'PVDFileMainHeader is;
With PVDFileMainHeader
    'vvvv PCN2207 *******************************************
    'If this application is unregistered then the PVD file
    'PVDRecording.PVD should not be allowed to be read by changing
    'its name and loading the file using the openfile method.
    
'    If RegType = "Purchased" Then  'PCN2207 testing to ADDED once new registration added.
    If Registered Then
        .PVDFileMHAppName = App.Title
        .PVDFileMHVersionMajor = App.Major
        .PVDFileMHVersionMinor = App.Minor
        .PVDFileMHVersionRev = App.Revision
    Else
        .PVDFileMHAppName = "Unregistered"
        .PVDFileMHVersionMajor = -100
        .PVDFileMHVersionMinor = -100
        .PVDFileMHVersionRev = -100
    End If
    '^^^^ ***************************************************
    .PVDFileMHPointerAddress = Len(PVDFileMainHeader) + 1 'PCNGL120103
    .PVDFileMHNoOfPointers = PVDFileOutPutNoOfPointers ' PVDFileOutPutNoOfPointers is a constant declared in the Startup module
    '.PVDFileMHRecordMode = RecordMode 'Taken out as per GL 010403
End With
Put #FileNo, , PVDFileMainHeader

'Determine file header pointers and CheckSums then write the header pointers to block 2
'File header pointers (must be the second block of data in any PVD file, there are no fix number of pointers)
PVDFilePointers.PVDPointerConfigInfo = 0    'To determine when block is about to be write, then update this line
PVDFilePointers.PVDPointerPipeInfo = 1      'To determine when block is about to be write, then update this line
PVDFilePointers.PVDPointerPipeObs = 2      'To determine when block is about to be write, then update this line 'PCNGL130103
PVDFilePointers.PVDPointerFontInfo = 3      'To determine when block is about to be write, then update this line
PVDFilePointers.PVDPointerDrawInfo = 4      'To determine when block is about to be write, then update this line
PVDFilePointers.PVDPointerPVData = 5        'To determine when block is about to be write, then update this line
Put #FileNo, , PVDFilePointers 'PCNGL120103

'File header CheckSum (These are check numbers used in auditing the length of a block of header data)
PVDHeaderConfigInfo.PVDCheck = 0  'PCNGL130103
PVDHeaderPipeInfo.PVDCheck = 0
PVDHeaderPipeObs.PVDCheck = 0 'PCNGL130103
PVDHeaderFontInfo.PVDCheck = 0
PVDHeaderDrawInfo.PVDCheck = 0
PVDHeaderPVData.PVDCheck = 0
    
    
'Write to file the System Configuration Information 'PCNGL130103
PVDHeaderConfigInfo.PVDHeaderDescriptor = "[ConfigInfo]"
PVDHeaderConfigInfo.PVDCheck = ConfigInfoNoOfLines
PVDFilePointers.PVDPointerConfigInfo = Seek(FileNo) 'Set the Configuration info pointer
Put #FileNo, , PVDHeaderConfigInfo
Put #FileNo, , ConfigInfo_V41 'PCN2639
'Note: At a later date, may need to check that ConfigInfo is currently V4.0

'Write to file the Pipe Information
PVDHeaderPipeInfo.PVDHeaderDescriptor = "[PipeInfo]"
PVDHeaderPipeInfo.PVDCheck = PipeInfoNoOfLines
PVDFilePointers.PVDPointerPipeInfo = Seek(FileNo) 'Set the Pipe info pointer
Put #FileNo, , PVDHeaderPipeInfo
Put #FileNo, , PipelineInfo 'PCNGL130103
'Note: At a later date, may need to check that PipelineInfo is currently V4.0


'Write to file the Pipe Observations 'PCNGL130103
PVDHeaderPipeObs.PVDHeaderDescriptor = "[PipeObs]"
NoOfPipeObservations = UBound(PipeObservations)
PVDHeaderPipeObs.PVDCheck = NoOfPipeObservations
PVDFilePointers.PVDPointerPipeObs = Seek(FileNo) 'Set the Font Info pointer
Put #FileNo, , PVDHeaderPipeObs
Call SaveObservations_V50(NoOfPipeObservations, 1) 'PCN2928


'Write to file the Font Information
PVDHeaderFontInfo.PVDHeaderDescriptor = "[FontInfo]"
PVDHeaderFontInfo.PVDCheck = FontInfoNoOfLines
PVDFilePointers.PVDPointerFontInfo = Seek(FileNo) 'Set the Font Info pointer
Put #FileNo, , PVDHeaderFontInfo
Call SaveFontInfo_V40(FileNo)


'Write to file the Drawing Information
PVDHeaderDrawInfo.PVDHeaderDescriptor = "[DrawInfo]"
PVDHeaderDrawInfo.PVDCheck = 0 'Are there any lines to draw?
PVDFilePointers.PVDPointerDrawInfo = Seek(FileNo) 'Set the Draw Info pointer
Put #FileNo, , PVDHeaderDrawInfo
'Determine and write drawing data


'Write to file the PVData
PVDHeaderPVData.PVDHeaderDescriptor = "[PVData]"
PVDHeaderPVData.PVDCheck = PVDataNoOfLines
PVDFilePointers.PVDPointerPVData = Seek(FileNo) 'Set the PVData pointer
Put #FileNo, , PVDHeaderPVData


If PVDataNoOfLines <> 0 Then 'PCNGL130103
    For PVDataLineNo = 0 To MaxFrameBufferNo 'PCNGL140103
        For PVSegmentNo = 1 To NoOfProfileSegments
            Put #FileNo, , pvData(PVSegmentNo, 0, PVDataLineNo)
        Next PVSegmentNo
        Put #FileNo, , pvCapacityData(PVDataLineNo) 'PCNGL1301032
        Put #FileNo, , PVOvalityData(PVDataLineNo) 'PCNGL1301032
        Put #FileNo, , PVDelta(0, PVDataLineNo) 'PCNGL1301032
        Put #FileNo, , PVDelta(1, PVDataLineNo) 'PCNGL1301032
        'DoEvents
    Next PVDataLineNo
End If
    
'Re-write the pointers with the correct settings
Put #FileNo, PVDFileMainHeader.PVDFileMHPointerAddress, PVDFilePointers 'PCNGL120103

    
' Close before reopening in another mode.
Close #FileNo

    
Exit Sub      'PCNGL130103
FileLoadErr_handler:
    Close #FileNo
    'MsgBox DisplayMessage("Save failed: Err=9,SaveInPVDFormat_V40"), vbInformation 'PCN2111
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Save failed: Err=9,SaveInPVDFormat_V40"): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    
    
Exit Sub
Err_Handler:
    Select Case Err
        Case 9 ' out of range
            GoTo FileLoadErr_handler
        Case Else
            MsgBox Err & "-PF24:" & Error$
    End Select
End Sub

Sub SaveInPVDFormat_V50(FileName As String, FileLoadError As Boolean)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : SaveInPVDFormat_V50
'Created : 22 March 2004, PCN2639
'Updated :
'Prg By  : Geoff Logan
'Param   : FileName - PVD file name to be created
'          FileLoadError - Returns a true value if this file creation process was not successful.
'Desc    : Saves the Precision Vision data to file, in binary format, version 5.0
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim PVSegmentNo As Integer
Dim PVDataLineNo As Long

'ConfigInfo.PVDFileVersion = "V5.0" 'PCN2820
'vvvv PCN2829 *******************************
'ConfigInfo.PVDFileVersion = "V5.1" 'PCN2820
'ConfigInfo.PVDFileVersion = "V5.2" 'PCN2928
Trim(ConfigInfo.PVDFileVersion) = "V5.3" 'PCN2928
PVCalculationsMultiplier = 100
PipeObsBuffer = 100 'PCN2928
'^^^^ ***************************************

FileLoadError = False 'PCNGL140103

ConfigInfo.DistanceProcessMethod = DistanceMethod
PVDataFrameBlockSize = PVDataFrameBlockSize_V50
PVCalculationsBlockSize = PVCalculationsBlockSize_V50
PVRelatedInfoBlockSize = PVRelatedInfoBlockSize_V50


Dim FileNo As Integer
FileNo = FreeFile
Open FileName For Binary Access Write As #FileNo

'Compile and save the File Main Header
'PVDFileMainHeader is;
With PVDFileMainHeader
    'vvvv PCN2207 *******************************************
    'If this application is unregistered then the PVD file
    'PVDRecording.PVD should not be allowed to be read by changing
    'its name and loading the file using the openfile method.
    
'    If RegType = "Purchased" Then  'PCN2207 testing to ADDED once new registration added.
    If Registered Then
        .PVDFileMHAppName = App.Title
        .PVDFileMHVersionMajor = App.Major
        .PVDFileMHVersionMinor = App.Minor
        .PVDFileMHVersionRev = App.Revision
    Else
        .PVDFileMHAppName = "Unregistered"
        .PVDFileMHVersionMajor = -100
        .PVDFileMHVersionMinor = -100
        .PVDFileMHVersionRev = -100
    End If
    '^^^^ ***************************************************
    .PVDFileMHPointerAddress = Len(PVDFileMainHeader) + 1 'PCNGL120103
    .PVDFileMHNoOfPointers = PVDFileOutPutNoOfPointers ' PVDFileOutPutNoOfPointers is a constant declared in the Startup module
    '.PVDFileMHRecordMode = RecordMode 'Taken out as per GL 010403
End With
Put #FileNo, , PVDFileMainHeader

'Determine file header pointers and CheckSums then write the header pointers to block 2
'File header pointers (must be the second block of data in any PVD file, there are no fix number of pointers)
PVDFilePointers.PVDPointerConfigInfo = 0    'To determine when block is about to be write, then update this line
PVDFilePointers.PVDPointerPipeInfo = 1      'To determine when block is about to be write, then update this line
PVDFilePointers.PVDPointerPipeObs = 2      'To determine when block is about to be write, then update this line 'PCNGL130103
PVDFilePointers.PVDPointerFontInfo = 3      'To determine when block is about to be write, then update this line
PVDFilePointers.PVDPointerDrawInfo = 4      'To determine when block is about to be write, then update this line
PVDFilePointers.PVDPointerPVData = 5        'To determine when block is about to be write, then update this line
Put #FileNo, , PVDFilePointers 'PCNGL120103

'File header CheckSum (These are check numbers used in auditing the length of a block of header data)
PVDHeaderConfigInfo.PVDCheck = 0  'PCNGL130103
PVDHeaderPipeInfo.PVDCheck = 0
PVDHeaderPipeObs.PVDCheck = 0 'PCNGL130103
PVDHeaderFontInfo.PVDCheck = 0
PVDHeaderDrawInfo.PVDCheck = 0
PVDHeaderPVData.PVDCheck = 0
    
    
'Write to file the System Configuration Information 'PCNGL130103
PVDHeaderConfigInfo.PVDHeaderDescriptor = "[ConfigInfo]"
PVDHeaderConfigInfo.PVDCheck = ConfigInfoNoOfLines
PVDFilePointers.PVDPointerConfigInfo = Seek(FileNo) 'Set the Configuration info pointer
Put #FileNo, , PVDHeaderConfigInfo
Put #FileNo, , ConfigInfo 'Leave as is unless this becomes an older version and needs conversion.
'Note: At a later date, may need to check that ConfigInfo is currently V4.0

'Write to file the Pipe Information
PVDHeaderPipeInfo.PVDHeaderDescriptor = "[PipeInfo]"
PVDHeaderPipeInfo.PVDCheck = PipeInfoNoOfLines
PVDFilePointers.PVDPointerPipeInfo = Seek(FileNo) 'Set the Pipe info pointer
Put #FileNo, , PVDHeaderPipeInfo
Put #FileNo, , PipelineInfo 'PCNGL130103
'Note: At a later date, may need to check that PipelineInfo is currently V4.0


'Write to file the Pipe Observations 'PCNGL130103
PVDHeaderPipeObs.PVDHeaderDescriptor = "[PipeObs]"
NoOfPipeObservations = UBound(PipeObservations)
PVDHeaderPipeObs.PVDCheck = NoOfPipeObservations
PVDFilePointers.PVDPointerPipeObs = Seek(FileNo) 'Set the Font Info pointer
Put #FileNo, , PVDHeaderPipeObs
Call SaveObservations_V50(NoOfPipeObservations, 1) 'PCN2928


'Write to file the Font Information
PVDHeaderFontInfo.PVDHeaderDescriptor = "[FontInfo]"
PVDHeaderFontInfo.PVDCheck = FontInfoNoOfLines
PVDFilePointers.PVDPointerFontInfo = Seek(FileNo) 'Set the Font Info pointer
Put #FileNo, , PVDHeaderFontInfo
Call SaveFontInfo_V40(FileNo)


'Write to file the Drawing Information
PVDHeaderDrawInfo.PVDHeaderDescriptor = "[DrawInfo]"
PVDHeaderDrawInfo.PVDCheck = 0 'Are there any lines to draw?
PVDFilePointers.PVDPointerDrawInfo = Seek(FileNo) 'Set the Draw Info pointer
Put #FileNo, , PVDHeaderDrawInfo
'Determine and write drawing data


'Write to file the PVData
PVDHeaderPVData.PVDHeaderDescriptor = "[PVData]"
PVDHeaderPVData.PVDCheck = PVDataNoOfLines
PVDFilePointers.PVDPointerPVData = Seek(FileNo) 'Set the PVData pointer
Put #FileNo, , PVDHeaderPVData


If PVDataNoOfLines <> 0 Then 'PCNGL130103
    For PVDataLineNo = 0 To MaxFrameBufferNo 'PCNGL140103
        For PVSegmentNo = 1 To NoOfProfileSegments
            Put #FileNo, , pvData(PVSegmentNo, 0, PVDataLineNo)
        Next PVSegmentNo
        Put #FileNo, , pvCapacityData(PVDataLineNo) 'PCNGL1301032
        Put #FileNo, , PVOvalityData(PVDataLineNo) 'PCNGL1301032
        Put #FileNo, , PVDelta(0, PVDataLineNo) 'PCNGL1301032
        Put #FileNo, , PVDelta(1, PVDataLineNo) 'PCNGL1301032
        'DoEvents
    Next PVDataLineNo
End If
    
'Re-write the pointers with the correct settings
Put #FileNo, PVDFileMainHeader.PVDFileMHPointerAddress, PVDFilePointers 'PCNGL120103

    
' Close before reopening in another mode.
Close #FileNo

    
Exit Sub      'PCNGL130103
FileLoadErr_handler:
    Close #FileNo
    'MsgBox DisplayMessage("Save failed: Err=9,SaveInPVDFormat_V50"), vbInformation 'PCN2111
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Save failed: Err=9,SaveInPVDFormat_V50"): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    
    
Exit Sub
Err_Handler:
    Select Case Err
        Case 9 ' out of range
            GoTo FileLoadErr_handler
        Case Else
            MsgBox Err & "-PF25:" & Error$
    End Select
End Sub

Sub SaveInPVDFormat_V6X(FileName As String, FileLoadError As Boolean)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : SaveInPVDFormat_V6X
'Created : 20 June 2004, PCN2891
'Updated :
'Prg By  : Geoff Logan
'Param   : FileName - PVD file name to be created
'          FileLoadError - Returns a true value if this file creation process was not successful.
'Desc    : Saves the Precision Vision data to file, in binary format, version 5.0
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim PVSegmentNo As Integer
Dim PVDataLineNo As Long
Dim NoOfPipeObservations As Integer 'PCNGL130103

'ConfigInfo.PVDFileVersion = "V5.0" 'PCN2820
'vvvv PCN2829 *******************************
'ConfigInfo.PVDFileVersion = "V5.1" 'PCN2820
'ConfigInfo.PVDFileVersion = "V5.2" ' PCN2891
'ConfigInfo.PVDFileVersion = "V6.0" 'PCN2891 'PCN3019
'ConfigInfo.PVDFileVersion = "V6.1" 'PCN3019
ConfigInfo.PVDFileVersion = PVDVersion 'PCN3576
PVCalculationsMultiplier = PVCalculationsMultiplier_V60 'PCN2891
PipeObsBuffer = 100 'PCN2928 'PCN3000
'^^^^ ***************************************

ConfigInfo.ProfileRecordingMethod = "XY" 'New method (the old method is "Radius")

FileLoadError = False 'PCNGL140103

ConfigInfo.DistanceProcessMethod = DistanceMethod
'PVDataFrameBlockSize = PVDataFrameBlockSize_V60
PVDataFrameBlockSize = PVDataFrameBlockSize_V70
PVCalculationsBlockSize = PVCalculationsBlockSize_V60
PVRelatedInfoBlockSize = PVRelatedInfoBlockSize_V60

Dim FileNo As Integer
FileNo = FreeFile
Open FileName For Binary Access Write As #FileNo

'Compile and save the File Main Header
'PVDFileMainHeader is;
With PVDFileMainHeader
    'vvvv PCN2207 *******************************************
    'If this application is unregistered then the PVD file
    'PVDRecording.PVD should not be allowed to be read by changing
    'its name and loading the file using the openfile method.
    
'    If RegType = "Purchased" Then  'PCN2207 testing to ADDED once new registration added.
    If Registered Then
        .PVDFileMHAppName = App.Title
        .PVDFileMHVersionMajor = App.Major
        .PVDFileMHVersionMinor = App.Minor
        .PVDFileMHVersionRev = App.Revision
    Else
        .PVDFileMHAppName = "Unregistered"
        .PVDFileMHVersionMajor = -100
        .PVDFileMHVersionMinor = -100
        .PVDFileMHVersionRev = -100
    End If
    '^^^^ ***************************************************
    .PVDFileMHPointerAddress = Len(PVDFileMainHeader) + 1 'PCNGL120103
    .PVDFileMHNoOfPointers = PVDFileOutPutNoOfPointers ' PVDFileOutPutNoOfPointers is a constant declared in the Startup module
    '.PVDFileMHRecordMode = RecordMode 'Taken out as per GL 010403
End With
Put #FileNo, , PVDFileMainHeader

'Determine file header pointers and CheckSums then write the header pointers to block 2
'File header pointers (must be the second block of data in any PVD file, there are no fix number of pointers)
PVDFilePointers.PVDPointerConfigInfo = 0    'To determine when block is about to be write, then update this line
PVDFilePointers.PVDPointerPipeInfo = 1      'To determine when block is about to be write, then update this line
PVDFilePointers.PVDPointerPipeObs = 2      'To determine when block is about to be write, then update this line 'PCNGL130103
PVDFilePointers.PVDPointerFontInfo = 3      'To determine when block is about to be write, then update this line
PVDFilePointers.PVDPointerDrawInfo = 4      'To determine when block is about to be write, then update this line
PVDFilePointers.PVDPointerPVData = 5        'To determine when block is about to be write, then update this line
Put #FileNo, , PVDFilePointers 'PCNGL120103

'File header CheckSum (These are check numbers used in auditing the length of a block of header data)
PVDHeaderConfigInfo.PVDCheck = 0  'PCNGL130103
PVDHeaderPipeInfo.PVDCheck = 0
PVDHeaderPipeObs.PVDCheck = 0 'PCNGL130103
PVDHeaderFontInfo.PVDCheck = 0
PVDHeaderDrawInfo.PVDCheck = 0
PVDHeaderPVData.PVDCheck = 0
    
    
'Write to file the System Configuration Information 'PCNGL130103
PVDHeaderConfigInfo.PVDHeaderDescriptor = "[ConfigInfo]"
PVDHeaderConfigInfo.PVDCheck = ConfigInfoNoOfLines
PVDFilePointers.PVDPointerConfigInfo = Seek(FileNo) 'Set the Configuration info pointer
Put #FileNo, , PVDHeaderConfigInfo
Put #FileNo, , ConfigInfo 'Leave as is unless this becomes an older version and needs conversion.
'Note: At a later date, may need to check that ConfigInfo is currently V4.0

'Write to file the Pipe Information
PVDHeaderPipeInfo.PVDHeaderDescriptor = "[PipeInfo]"
PVDHeaderPipeInfo.PVDCheck = PipeInfoNoOfLines
PVDFilePointers.PVDPointerPipeInfo = Seek(FileNo) 'Set the Pipe info pointer
Put #FileNo, , PVDHeaderPipeInfo
Put #FileNo, , PipelineInfo 'PCNGL130103
'Note: At a later date, may need to check that PipelineInfo is currently V4.0


'Write to file the Pipe Observations 'PCNGL130103
PVDHeaderPipeObs.PVDHeaderDescriptor = "[PipeObs]"
NoOfPipeObservations = UBound(PipeObservations)
PVDHeaderPipeObs.PVDCheck = NoOfPipeObservations
PVDFilePointers.PVDPointerPipeObs = Seek(FileNo) 'Set the Font Info pointer
Put #FileNo, , PVDHeaderPipeObs
Call SaveObservations_V60(NoOfPipeObservations, 1) 'PCN2928 'PCN3000 'PCN3576


'Write to file the Font Information
PVDHeaderFontInfo.PVDHeaderDescriptor = "[FontInfo]"
PVDHeaderFontInfo.PVDCheck = FontInfoNoOfLines
PVDFilePointers.PVDPointerFontInfo = Seek(FileNo) 'Set the Font Info pointer
Put #FileNo, , PVDHeaderFontInfo
Call SaveFontInfo_V40(FileNo)


'Write to file the Drawing Information
PVDHeaderDrawInfo.PVDHeaderDescriptor = "[DrawInfo]"
PVDHeaderDrawInfo.PVDCheck = 0 'Are there any lines to draw?
PVDFilePointers.PVDPointerDrawInfo = Seek(FileNo) 'Set the Draw Info pointer
Put #FileNo, , PVDHeaderDrawInfo
'Determine and write drawing data


'Write to file the PVData
PVDHeaderPVData.PVDHeaderDescriptor = "[PVData]"
PVDHeaderPVData.PVDCheck = PVDataNoOfLines
PVDFilePointers.PVDPointerPVData = Seek(FileNo) 'Set the PVData pointer
Put #FileNo, , PVDHeaderPVData


If PVDataNoOfLines <> 0 Then 'PCNGL130103
    For PVDataLineNo = 0 To MaxFrameBufferNo 'PCNGL140103
        For PVSegmentNo = 1 To NoOfProfileSegments
            'vvvv PCN2891 **************************************************
            If Trim(ConfigInfo.ProfileRecordingMethod) = "XY" Then
                'Save the PVData X
                Put #FileNo, , pvData(PVSegmentNo, 1, PVDataLineNo)
                'Save the PVData Y
                Put #FileNo, , pvData(PVSegmentNo, 2, PVDataLineNo)
            Else
                'Save the PVData Radius
                Put #FileNo, , pvData(PVSegmentNo, 0, PVDataLineNo)
            End If
            '^^^^ ***********************************************************
        Next PVSegmentNo
        Put #FileNo, , pvCapacityData(PVDataLineNo) 'PCNGL1301032
        Put #FileNo, , PVOvalityData(PVDataLineNo) 'PCNGL1301032
        Put #FileNo, , PVDelta(0, PVDataLineNo) 'PCNGL1301032
        Put #FileNo, , PVDelta(1, PVDataLineNo) 'PCNGL1301032
        'DoEvents
    Next PVDataLineNo
End If
    
'Re-write the pointers with the correct settings
Put #FileNo, PVDFileMainHeader.PVDFileMHPointerAddress, PVDFilePointers 'PCNGL120103

    
' Close before reopening in another mode.
Close #FileNo

    
Exit Sub      'PCNGL130103
FileLoadErr_handler:
    Close #FileNo
    'MsgBox DisplayMessage("Save failed: Err=9,SaveInPVDFormat_V50"), vbInformation 'PCN2111
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Save failed: Err=9,SaveInPVDFormat_V50"): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    
    
Exit Sub
Err_Handler:
    Select Case Err
        Case 9 ' out of range
            GoTo FileLoadErr_handler
        Case Else
            MsgBox Err & "-PF26:" & Error$
    End Select
End Sub


Sub LoadInPVDFormat(FileName As String, FileLoadError As Boolean)     'PCNGL140103
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'LoadInPVDFormat Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff,    14/01/03     Building initial framework
'
'Description:
'       Loads the Precision Vision data from file, in binary format,
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim DrawFrameNo As Long
Dim FrameBufferNo  As Integer 'The size of the PVData frame buffer 'PCNGL140103
Dim ErrorType As String 'PCN3964
Dim PVDataAddressOffset As Long
Dim PVDVer As Single

'vvvv PCN2891 *******************************************
Call LoadInPVDFormat_V6X(FileName, FileLoadError, ErrorType) 'PCN3964
If ErrorType = "Wrong Units" Then Exit Sub      'PCN3964
If FileLoadError Then
    'Check to see if the file is in format V5.X
    Call LoadInPVDFormat_V50(FileName, FileLoadError, ErrorType) 'PCN3964
    If ErrorType = "Wrong Units" Then Exit Sub      'PCN3964
    If FileLoadError Then
        'Check to see if the file is in format V4.X
        Call LoadInPVDFormat_V40(FileName, FileLoadError, ErrorType) 'PCN3964
        If ErrorType = "Wrong Units" Then Exit Sub      'PCN3964
        If FileLoadError Then
            'MsgBox DisplayMessage("Can't read file, ") & FileName, vbExclamation  'PCN2952
            ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Can't read file, ") & FileName: ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
End If




PipelineDetails.SeaLevelStartHeightTextBox.text = Trim(FontInfo.FontType)
SeaLevelStartHeight = SafeCDbl(Trim(FontInfo.FontType))
PipelineDetails.SeaLevelEndHeightTextBox.text = Trim(FontInfo.FontColour)
SeaLevelEndHeight = SafeCDbl(Trim(FontInfo.FontColour))

DesignGradient = FontInfo.FontSize / 1000 'PCN6165
PipelineDetails.DesignGradientTextBox.text = DesignGradient 'PCN6165


'^^^^ ****************************************************
'Call ScreenDrawing.SetupPVCentre 'PCN4484

'Draw Y scale
Call ClearLineScreen.VideoScreenScaleCalc ' PCN3513


' Move to first frame 'PCNGL171202
PVFrameNo = 1 'PCNGL171202
'Move start and finish Y Markers
Call PrecisionVisionGraph.SetupYScaleMarkers(PrecisionVisionGraph.PVYScaleZeroMarker(0).y1) ' PCN2970

'Update PV frame status bar
Call ClearLineScreen.SetPVFrameStatus 'PCN4171
    
'Draw capacity graph
'vvvv PCN2164 *****************************************************
Dim PVDataStartAddress As Long
Dim PVGraphDataAddressOffset As Long
Dim FileNumber As Integer

FileNumber = 1
Call GetPVDPointerPVDataFromFile(FileName, PVDataStartAddress, FileLoadError) 'PCN2164
If FileLoadError Then Exit Sub
Open FileName For Binary Access Read Lock Write As #FileNumber
'^^^^ *************************************************************
'vvvv PCN2401 *****************************************************
Dim LoadingProgress As Integer
''Dim LoadingProgress_Current As Integer 'PCN4241
Dim ProgressIncrement As Integer

If PVDataNoOfLines >= 40000 Then
    ProgressIncrement = 2
ElseIf PVDataNoOfLines <= 1000 Then
    ProgressIncrement = 10
Else
    ProgressIncrement = 2 + 8 * (40000 - PVDataNoOfLines) / 39000
End If

Call CLPProgressBar.ProgressBarInitialise(DisplayMessage("Loading data"))
DoEvents
ProgressBarPercent = 0 'PCN4241
'^^^^ *****************************************************************
'vvvv PCN2703******************************************************
ReDim PVXDiameterFullData(PVDataNoOfLines) As Double  'PCN2703
ReDim PVYDiameterFullData(PVDataNoOfLines) As Double  'PCN2703
'^^^^ **************************************************************
'vvvv PCN2962**********************************************************
'ReDim PVMinMaxSegNosFullData(PVDataNoOfLines, 3) As Long 'PCN2966
Dim TDArraySize As Long
TDArraySize = NoOfProfileSegments * (PVDataNoOfLines + 1)
ReDim TD_PVDataX(TDArraySize)
ReDim TD_PVDataY(TDArraySize)

'PCN3219 centres
Call ScreenDrawing.SetupPVCentre 'PCN4484
ReDim TD_PVCentreX(PVDataNoOfLines)
ReDim TD_PVCentreY(PVDataNoOfLines)

'XY
ReDim PVXDiameterFullData(PVDataNoOfLines)
ReDim PVYDiameterFullData(PVDataNoOfLines)
'Flat3D
ReDim PVFlat3DRed(NoOfProfileSegments, PVDataNoOfLines) As Long
ReDim PVFlat3DGreen(NoOfProfileSegments, PVDataNoOfLines) As Long
ReDim PVFlat3DBlue(NoOfProfileSegments, PVDataNoOfLines) As Long
'Median
'ReDim PVDiameterMedian(0) 'PCN3489
'ReDim PVFractile(PVDataNoOfLines) 'PCN4235
ReDim PVDiameterMedian(PVDataNoOfLines)
ReDim PVCapacityFullData(PVDataNoOfLines) 'PCN3540
ReDim GraphInfoContainer(PVOvality).DataSingle(PVDataNoOfLines)
ReDim GraphInfoContainer(PVDeflectionX).DataSingle(PVDataNoOfLines) 'PCN5186
ReDim GraphInfoContainer(PVDeflectionY).DataSingle(PVDataNoOfLines) 'PCN5186
ReDim SmoothDeflectionX(PVDataNoOfLines)
ReDim SmoothDeflectionY(PVDataNoOfLines)
'PCN6458 ReDim GraphInfoContainer(PVInclination).DataSingle(PVDataNoOfLines) 'PCN6128
'PCN6458 ReDim GraphInfoContainer(PVDesignGradient).DataSingle(PVDataNoOfLines) 'PCN6178




ReDim GraphInfoContainer(PVDebris).DataSingle(PVDataNoOfLines) 'PCN4461
'ReDim PVOvalityOrigFullData(PVDataNoOfLines)


'Delta PCN3540 (Antony, 4 August 2005)
ReDim PVDeltaFullMax(PVDataNoOfLines)
ReDim PVDeltaFullMin(PVDataNoOfLines)
ReDim PVDeltaSegFullMax(PVDataNoOfLines)
ReDim PVDeltaSegFullMin(PVDataNoOfLines)

ReDim GraphInfoContainer(PVMaxDiameter).DataDouble(PVDataNoOfLines)
ReDim GraphInfoContainer(PVMinDiameter).DataDouble(PVDataNoOfLines) 'PCN4333
ReDim PVDiameterFullMin(PVDataNoOfLines)
ReDim PVDiameterSegFullMax(PVDataNoOfLines)
ReDim PVDiameterSegFullMin(PVDataNoOfLines)

ReDim PVDistances(PVDataNoOfLines)
ReDim PVTimes(PVDataNoOfLines)



'ReDim PVShapeCentreX(PVDataNoOfLines)
'ReDim PVShapeCentreY(PVDataNoOfLines)


'''' PCN3441 (6 April 2005, Antony van Iersel)
''
''PrecisionVisionGraph.PVGraphScreen(0).Top = 0

'PrecisionVisionGraph.PVGraphScreen(0).Visible = True PCN3373
'^^^^ ***************************************************

'PCN3219 check to see if there is water level or not, and if there is set the egnore points


PVDVer = GetPVDVer
If PVDVer = 6.3 And (WLStartAngle <> 0 Or WLFinishAngle <> 0) Then
    'This is a one of call if PVD has been recorded in ver 6.3 but the centre calculation has
    'not yet been saved
    
    Call ClearLineScreen.SetWaterLevelinPipe(WLStartAngle, WLFinishAngle)
    WaterLevelIgnoreCenter = True 'PCNLS190603
    Call ClearLineScreen.SetUnsetWaterLevel(WaterLevelIgnoreCenter, False) 'PCN3219 WaterLevelIgnoreProfile) 'PCNLS190603waterlevel
    Call ScreenDrawing.RecalculatePVData
ElseIf PVDVer < 6.4 Then
    ConfigInfo.WLStartAngle = 0 And ConfigInfo.WLFinishAngle = 0
    
    WLStartAngle = 0
    WLFinishAngle = 0
    Call ScreenDrawing.DeleteWaterLevel
    WaterLevelIgnoreCenter = False
    Call ClearLineScreen.SetUnsetWaterLevel(WaterLevelIgnoreCenter, False)
ElseIf WLStartAngle <> 0 Or WLFinishAngle <> 0 Then
    Call ClearLineScreen.SetWaterLevelinPipe(WLStartAngle, WLFinishAngle)
    WaterLevelIgnoreCenter = True 'PCNLS190603
    Call ClearLineScreen.SetUnsetWaterLevel(WaterLevelIgnoreCenter, False) 'PCN3219 WaterLevelIgnoreProfile) 'PCNLS190603waterlevel
End If


LoadingTimeStampError = False

For DrawFrameNo = 0 To PVDataNoOfLines - 1 'PCN2962
    PVFrameNo = DrawFrameNo + 1 'PCN2255 LS 'PCN2962
    'vvvv PCN2164 ************************************************
    PVGraphDataAddressOffset = PVDataStartAddress + (DrawFrameNo + 1) * PVDataFrameBlockSize 'PCN2639
    PVGraphDataAddressOffset = PVGraphDataAddressOffset + (DrawFrameNo) * (PVCalculationsBlockSize + PVRelatedInfoBlockSize)   'PCN2639
    Call RapidReadPVGraphsDataFromFile(1, PVFrameNo, PVGraphDataAddressOffset, FileLoadError) 'PCN2164
 

    '^^^^ ************************************
    'vvvv PCN2401 **********************************************
    LoadingProgress = 70 * PVFrameNo / PVDataNoOfLines  'PCN2962
    If LoadingProgress Mod ProgressIncrement = 0 And LoadingProgress > (ProgressBarPercent + 1) Then 'PCN4241
        Call CLPProgressBar.ProgressBarPosition(LoadingProgress / 100)
        DoEvents
        ProgressBarPercent = LoadingProgress 'PCN4241
    End If
    '^^^^ **********************************************************
Next DrawFrameNo

Close #FileNumber 'PCN2164

'Call PageFunctions.SaveToFilePipeAndConfigInfo("ConfigInfo", FileLoadError) 'ID5395

Call FixTimeStampErrors
Call LoadFullPVDataFromFile
'Call LoadPVDataFromFile_For_3D(PVDFileName, 0, PVDataNoOfLines, FileLoadError)

'Call ScreenDrawing.PVCentreCalcCPP(0, PVDataNoOfLines) 'PCN3194
Call ScreenDrawing.PVDiameterMedianCalcCPP(0, PVDataNoOfLines) 'PCN3540 'PCN4974 need to calculate median diameter incase its used with diameter flat.
Call ScreenDrawing.PVFlat3DCalcCPP(0, PVDataNoOfLines) 'PCN3513
Call ScreenDrawing.PVCapacityCalcCPP(0, PVDataNoOfLines) 'PCN3540
Call ScreenDrawing.PVXYDiameterCalcCPP(0, PVDataNoOfLines) 'PCN3540
Call ScreenDrawing.PVDeltaMaxMinCalcCPP(0, PVDataNoOfLines) 'pcn3540
Call ScreenDrawing.PVDiameterMaxMinCalcCPP(0, PVDataNoOfLines) 'PCN3540
Call ScreenDrawing.FixMinMax((0), PVDataNoOfLines) 'PCN6524

Call ScreenDrawing.PVOvalityCalcCPP(0, PVDataNoOfLines) 'PCN3540
'Call ScreenDrawing.PVInclinationCalc(0, PVDataNoOfLines) 'PCN6128





'Call ScreenDrawing.PVDebrisCalcCPP(0, PVDataNoOfLines) ' PCN4461






Call ScreenDrawing.CPPFilterGraphs 'PCN4355
    

    
    
    If OptionsPage.SmoothCheck.value = vbChecked Then
        ReDim GraphInfoContainer(PVOvalitySmooth).DataSingle(PVDataNoOfLines) 'PCN9999
        ReDim GraphInfoContainer(PVXDiameterSmooth).DataSingle(PVDataNoOfLines)
        ReDim GraphInfoContainer(PVYDiameterSmooth).DataSingle(PVDataNoOfLines)
        ReDim GraphInfoContainer(PVMedianDiameterSmooth).DataSingle(PVDataNoOfLines) 'PCN9999
        ReDim GraphInfoContainer(PVMaxDiameterSmooth).DataSingle(PVDataNoOfLines)
        ReDim GraphInfoContainer(PVMinDiameterSmooth).DataSingle(PVDataNoOfLines)
        ReDim GraphInfoContainer(PVCapacitySmooth).DataSingle(PVDataNoOfLines)
'PCN6458         ReDim GraphInfoContainer(PVInclinationSmooth).DataSingle(PVDataNoOfLines) 'PCN6128
        
        Call CPPSmoothGraphs
 
    Else
        ReDim GraphInfoContainer(PVOvalitySmooth).DataSingle(0) 'PCN9999
        ReDim GraphInfoContainer(PVXDiameterSmooth).DataSingle(0)
        ReDim GraphInfoContainer(PVYDiameterSmooth).DataSingle(0)
        ReDim GraphInfoContainer(PVMedianDiameterSmooth).DataSingle(0) 'PCN9999
        ReDim GraphInfoContainer(PVMaxDiameterSmooth).DataSingle(0)
        ReDim GraphInfoContainer(PVMinDiameterSmooth).DataSingle(0)
        ReDim GraphInfoContainer(PVCapacitySmooth).DataSingle(0)
'PCN6458         ReDim GraphInfoContainer(PVInclinationSmooth).DataSingle(0) 'PCN6128

    End If
    
    Call ScreenDrawing.PVDeflectionCalcCPP   'PCN5186

Dim i As Long
'For i = 0 To PVDataNoOfLines
'    PVOvalityOrigFullData(i) = PVOvalityFullData(i)
'Next i


Call ProcessBarIncrement 'PCN4241

If ConfigInfo.DistanceStart > InvalidData And PVDataNoOfLines > 0 Then 'PCN3884 was -1 now -1000
    'Find the average speed CameraSpeedInFrames
    CameraSpeedInFrames = (PVDistances(PVDataNoOfLines - 1) - PVDistances(1)) / (PVDataNoOfLines - 1)
'    CameraSpeedInTime = 0.015 'm/sec from eg of [10m/min / 60sec/min)]
    CameraSpeedInTime = 0
End If

If ConfigInfo.DistanceStart = -1 Then ConfigInfo.DistanceStart = InvalidData 'PCN3448
If ConfigInfo.DistanceFinish = -1 Then ConfigInfo.DistanceFinish = InvalidData 'PCN3448
'Call DrawPVYScale(1) 'PCN2639 PCN1850 Just draw the first page of PVYScale. Draw when ViewIndicator is moved.
'vvvv PCN2970 *******************************************



'PCN3373 Call SetPositionOfPVGraphBaseCover 'PCN2970
'Redimension the ViewIndicators 'PCN2970
Call Observations.SortObs


Call ProcessBarIncrement 'PCN4241

Call Observations.ObsInitPictureStorage

Call ProcessBarIncrement 'PCN4241

Call Distance.RecalculateDistance
'PCN6458 Call ScreenDrawing.PVInclinationCalc(0, PVDataNoOfLines) 'PCN6128 'PCN6178 had to move to after distance calc, for design fradient to work
 'Call ScreenDrawing.AdjustLayToFit
Call ProcessBarIncrement 'PCN4241

Call ReDimensionIndicators 'PCN2970
'PCN6458 Call CPPSmoothInclination
Exit Sub
Err_Handler:
    Select Case Err
        Case 9: Exit Sub      ' out of range
        Case 6: Resume Next
        Case Else
            MsgBox Err & "-PF27:" & Error$

    End Select
End Sub

Sub LoadInPVDFormat_V6X(FileName As String, FileLoadError As Boolean, ErrorType As String) 'PCN3964
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : LoadInPVDFormat_V6X
'Created : 19 March 2004, PCN2891
'Updated :
'Prg By  : Geoff Logan
'Param   :  FileName - PVD file name
'           FileLoadError - File Load Error flag
'Desc    : Loads the Precision Vision data from file, in binary format, version 6.X
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim PVSegmentNo As Integer
Dim PVDataLineNo As Integer 'PCN2891
Dim PipeInfoNoOfLines As Integer
Dim NodeInfoNoOfLines As Integer
Dim DrawInfoNoOfLines As Integer

FileLoadError = False
ErrorType = "Unknown"  'PCN3964

If Dir(FileName) = "" Or FileName = "" Then
    'MsgBox DisplayMessage("File was not found."), vbInformation
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("File was not found."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    FileLoadError = True
    Exit Sub
End If

ClearLineTitle.TitleBarCaption.Caption = DisplayMessage("Precision Vision") & " - " & FileName 'PCN2561
ClearLineTitle.TitleBarCaption.ToolTipText = DisplayMessage("Precision Vision") & " - " & FileName 'PCN4171

Dim FileNo As Integer
FileNo = FreeFile
Open FileName For Binary Access Read Lock Write As #FileNo

'Load the File Main Header
Get #FileNo, , PVDFileMainHeader

'vvvv PCN2207 ********************************************
'Check to see that this file was produced by a registered
'copy of the application. Do not open this file if it is
'not registered.
If Left(PVDFileMainHeader.PVDFileMHAppName, 12) = "Unregistered" Or PVDFileMainHeader.PVDFileMHVersionMajor = -100 Or PVDFileMainHeader.PVDFileMHVersionMinor = -100 Or PVDFileMainHeader.PVDFileMHVersionRev = -100 Then
    'MsgBox DisplayMessage("Can't read file, ") & FileName, vbExclamation
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Can't read file, ") & FileName: ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Close #FileNo
    FileLoadError = True
    Exit Sub
End If
'^^^^ ****************************************************

'Determine file header pointers and CheckSums then read the header pointers to block 2
'File header pointers (must be the second block of data in any PVD file, there are no fix number of pointers)
If PVDFileMainHeader.PVDFileMHPointerAddress = 0 Then
    'MsgBox DisplayMessage("Can't read file, ") & FileName
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Can't read file, ") & FileName: ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Close #FileNo
    FileLoadError = True
    Exit Sub
End If
Get #FileNo, PVDFileMainHeader.PVDFileMHPointerAddress, PVDFilePointers


'Read from file the System Configuration Information
PVDHeaderConfigInfo.PVDHeaderDescriptor = ""
PVDHeaderConfigInfo.PVDCheck = 0
Get #FileNo, PVDFilePointers.PVDPointerConfigInfo, PVDHeaderConfigInfo

Call LoadConfigInfo_V6X(Seek(FileNo), FileLoadError, FileNo)
If FileLoadError Then
    'Not in Version 6.X format
    Close #FileNo
    Exit Sub
End If

'vvvv PCN3809 ******************************
If SoftwareConfiguration = "Reader" Then
    MeasurementUnits = ConfigInfo.Units
    ClearLineScreen.Y_Units.Caption = MeasurementUnits
End If
'^^^^ **************************************

'PCN1919
If (ConfigInfo.Units = "mm" And MeasurementUnits = "in") Or (ConfigInfo.Units = "in" And MeasurementUnits = "mm") Then
    'If LoadVideo = True Then MsgBox DisplayMessage("The PVD file is in a different measurement unit, please alter your measurement units before loading this file."), vbExclamation 'PCN2111
    If LoadVideo = True Then ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("The PVD file is in a different measurement unit, please alter your measurement units before loading this file."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    PVDLoadError = True
    Close #FileNo
    Screen.MousePointer = vbDefault
'PCN 1984 LS 10/7/03
    FileLoadError = True
    ErrorType = "Wrong Units" 'PCN3964
    OptionsPage.Show
    PVDFileName = "" 'PCN3964
    OptionsPage.ZOrder 0 'PCN3964
    Exit Sub
ElseIf (ConfigInfo.Units <> "mm" And MeasurementUnits <> "in") Then 'PCN2207
    'vvvv PCN2207 ********************************************
    'MsgBox DisplayMessage("Can't read file, ") & FileName, vbExclamation  'PCN2111
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Can't read file, ") & FileName: ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Close #FileNo
    FileLoadError = True
    Exit Sub
    '^^^^ ****************************************************
End If

'Read from file the Pipe Information
PVDHeaderPipeInfo.PVDHeaderDescriptor = ""
PVDHeaderPipeInfo.PVDCheck = 0
Get #FileNo, PVDFilePointers.PVDPointerPipeInfo, PVDHeaderPipeInfo
Call LoadPipelineDetails_V40(Seek(FileNo), FileNo) 'PCNGL1201032

'PCN4799
'Read from file the font info which actually is camera model
'So take note: For now fontinfo.fontname is actually camera model
PVDHeaderFontInfo.PVDHeaderDescriptor = ""
PVDHeaderFontInfo.PVDCheck = 0
Get #FileNo, PVDFilePointers.PVDPointerFontInfo, PVDHeaderFontInfo
Call LoadFontInfo_V40(Seek(FileNo), FileNo)



'Read from file the Pipe Observations 'PCNGL130103
PVDHeaderPipeObs.PVDHeaderDescriptor = ""
PVDHeaderPipeObs.PVDCheck = 0
Get #FileNo, PVDFilePointers.PVDPointerPipeObs, PVDHeaderPipeObs
'Call LoadPipelineObs_V40(Seek(1), PVDHeaderPipeObs.PVDCheck) 'PCN2928
'Call LoadPipelineObs_V50(Seek(1), PVDHeaderPipeObs.PVDCheck) 'PCN2928 'PCN3000
Call LoadPipelineObs_V60(Seek(FileNo), PVDHeaderPipeObs.PVDCheck, FileNo) 'PCN3576
'Call Observations.SortObs
'Call Observations.ObsInitPictureStorage



'Read from file the Font Information
PVDHeaderFontInfo.PVDHeaderDescriptor = ""
PVDHeaderFontInfo.PVDCheck = 0
Get #FileNo, PVDFilePointers.PVDPointerFontInfo, PVDHeaderFontInfo


'Read from file the Drawing Information
PVDHeaderDrawInfo.PVDHeaderDescriptor = ""
PVDHeaderDrawInfo.PVDCheck = 0 'Are there any lines to draw?
Get #FileNo, PVDFilePointers.PVDPointerDrawInfo, PVDHeaderDrawInfo
'Determine and read drawing data



'''Read from file the PVData
PVDHeaderPVData.PVDHeaderDescriptor = ""
PVDHeaderPVData.PVDCheck = 0

Get #FileNo, PVDFilePointers.PVDPointerPVData, PVDHeaderPVData
If Left(PVDHeaderPVData.PVDHeaderDescriptor, 8) = "[PVData]" And PVDHeaderPVData.PVDCheck >= 0 Then 'PCNGL130103 'PCN3274
    PVDataNoOfLines = PVDHeaderPVData.PVDCheck 'PCNGL130103
''    'Call InitilisePVProfile(PVDataNoOfLines) 'PCNGL130103
''    'For PVDataLineNo = 0 To PVDataNoOfLines
'''    Call InitilisePVProfile(MaxFrameBufferNo) 'PCNGL130103 'PCN2970
''    ReDim PVDistances(1) 'PCN2639
''    For PVDataLineNo = 0 To 1 'Get the first 2 profiles only 'PCNGL140103
''        For PVSegmentNo = 1 To NoOfProfileSegments
''            Get #1, , PVData(PVSegmentNo, 1, PVDataLineNo)
''            Get #1, , PVData(PVSegmentNo, 2, PVDataLineNo)
''            On Error Resume Next
''            PVData(PVSegmentNo, 0, PVDataLineNo) = Int(PVDataTrueRadiusCalc(PVSegmentNo, PVDataLineNo))
''            On Error GoTo Err_Handler
''        Next PVSegmentNo
''        Get #1, , PVCapacityData(PVDataLineNo) 'PCNGL1301032
''        Get #1, , PVOvalityData(PVDataLineNo) 'PCNGL1301032
''        Get #1, , PVDelta(0, PVDataLineNo) 'PCNGL1301032
''        Get #1, , PVDelta(1, PVDataLineNo) 'PCNGL1301032
''        'vvvv **** load AVI frame time *************************** 'PCNGL150103
''        'To be used to more accurately link the PVD file Frame no to the AVI frames. PCNGL150103
''        Get #1, , AVIFrameTime(PVDataLineNo)
''        '^^^^ ****************************************************
''        Get #1, , PVDistances(PVDataLineNo) 'PCN2639
''    Next PVDataLineNo
End If
'''

'
' Close before reopening in another mode.
Close #FileNo

'Update the forms with the loaded data 'PCNGL130103
Call CopyConfigInfoToForms
Call CopyPipeDetailsToPipelineForm
'Call CopyPipeDetailsToObsForm 'PCN4131
    
    
Exit Sub
Err_Handler:
    Select Case Err
        Case 9 ' out of range
        Case 52: Resume Next 'BadFile Name
        Case Else
            MsgBox Err & "-PF28:" & Error$
    End Select
End Sub


Sub LoadInPVDFormat_V50(FileName As String, FileLoadError As Boolean, ErrorType As String) 'PCN3964
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : LoadInPVDFormat_V50
'Created : 19 March 2004, PCN2639
'Updated :
'Prg By  : Geoff Logan
'Param   :  FileName - PVD file name
'           FileLoadError - File Load Error flag
'Desc    : Loads the Precision Vision data from file, in binary format, version 5.0
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim PVSegmentNo As Integer
Dim PVDataLineNo As Long
Dim PipeInfoNoOfLines As Integer
Dim NodeInfoNoOfLines As Integer
Dim DrawInfoNoOfLines As Integer

FileLoadError = False
ErrorType = "Unknown" 'PCN3964

If Dir(FileName) = "" Or FileName = "" Then
    'MsgBox DisplayMessage("File was not found."), vbInformation
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("File was not found."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    FileLoadError = True
    Exit Sub
End If

ClearLineTitle.TitleBarCaption.Caption = DisplayMessage("Precision Vision") & " - " & FileName 'PCN2561
ClearLineTitle.TitleBarCaption.ToolTipText = DisplayMessage("Precision Vision") & " - " & FileName 'PCN4171

Dim FileNo As Integer
FileNo = FreeFile
Open FileName For Binary Access Read Lock Write As #FileNo

'Load the File Main Header
Get #FileNo, , PVDFileMainHeader

'vvvv PCN2207 ********************************************
'Check to see that this file was produced by a registered
'copy of the application. Do not open this file if it is
'not registered.
If Left(PVDFileMainHeader.PVDFileMHAppName, 12) = "Unregistered" Or PVDFileMainHeader.PVDFileMHVersionMajor = -100 Or PVDFileMainHeader.PVDFileMHVersionMinor = -100 Or PVDFileMainHeader.PVDFileMHVersionRev = -100 Then
    'MsgBox DisplayMessage("Can't read file, ") & FileName, vbExclamation
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Can't read file, ") & FileName: ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Close #FileNo
    FileLoadError = True
    Exit Sub
End If
'^^^^ ****************************************************

'Determine file header pointers and CheckSums then read the header pointers to block 2
'File header pointers (must be the second block of data in any PVD file, there are no fix number of pointers)
If PVDFileMainHeader.PVDFileMHPointerAddress = 0 Then
    'MsgBox DisplayMessage("Can't read file, ") & FileName
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Can't read file, ") & FileName: ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Close #FileNo
    FileLoadError = True
    Exit Sub
End If
Get #FileNo, PVDFileMainHeader.PVDFileMHPointerAddress, PVDFilePointers


'Read from file the System Configuration Information
PVDHeaderConfigInfo.PVDHeaderDescriptor = ""
PVDHeaderConfigInfo.PVDCheck = 0
Get #FileNo, PVDFilePointers.PVDPointerConfigInfo, PVDHeaderConfigInfo

Call LoadConfigInfo_V50(Seek(FileNo), FileLoadError, FileNo)
If FileLoadError Then
    'Not in Version 5.0 format
    Close #FileNo
    Exit Sub
End If


'PCN1919
If (ConfigInfo.Units = "mm" And MeasurementUnits = "in") Or (ConfigInfo.Units = "in" And MeasurementUnits = "mm") Then
    'If LoadVideo = True Then MsgBox DisplayMessage("The PVD file is in a different measurement unit, please alter your measurement units before loading this file."), vbExclamation 'PCN2111
    If LoadVideo = True Then ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("The PVD file is in a different measurement unit, please alter your measurement units before loading this file."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Close #FileNo
'PCN 1984 LS 10/7/03
    FileLoadError = True
    ErrorType = "Wrong Units" 'PCN3964
    OptionsPage.Show
    PVDFileName = "" 'PCN3964
    OptionsPage.ZOrder 0 'PCN3964
    Exit Sub
ElseIf (ConfigInfo.Units <> "mm" And MeasurementUnits <> "in") Then 'PCN2207
    'vvvv PCN2207 ********************************************
    'MsgBox DisplayMessage("Can't read file, ") & FileName, vbExclamation  'PCN2111
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Can't read file, ") & FileName: ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    FileLoadError = True
    Exit Sub
    '^^^^ ****************************************************
End If

'Read from file the Pipe Information
PVDHeaderPipeInfo.PVDHeaderDescriptor = ""
PVDHeaderPipeInfo.PVDCheck = 0
Get #FileNo, PVDFilePointers.PVDPointerPipeInfo, PVDHeaderPipeInfo
Call LoadPipelineDetails_V40(Seek(FileNo), FileNo) 'PCNGL1201032


'Read from file the Pipe Observations 'PCNGL130103
PVDHeaderPipeObs.PVDHeaderDescriptor = ""
PVDHeaderPipeObs.PVDCheck = 0
Get #FileNo, PVDFilePointers.PVDPointerPipeObs, PVDHeaderPipeObs
Call LoadPipelineObs_V50(Seek(FileNo), PVDHeaderPipeObs.PVDCheck, FileNo) 'PCN2928



'Read from file the Font Information
PVDHeaderFontInfo.PVDHeaderDescriptor = ""
PVDHeaderFontInfo.PVDCheck = 0
Get #FileNo, PVDFilePointers.PVDPointerFontInfo, PVDHeaderFontInfo


'Read from file the Drawing Information
PVDHeaderDrawInfo.PVDHeaderDescriptor = ""
PVDHeaderDrawInfo.PVDCheck = 0 'Are there any lines to draw?
Get #FileNo, PVDFilePointers.PVDPointerDrawInfo, PVDHeaderDrawInfo
'Determine and read drawing data



'Read from file the PVData
PVDHeaderPVData.PVDHeaderDescriptor = ""
PVDHeaderPVData.PVDCheck = 0
Get #FileNo, PVDFilePointers.PVDPointerPVData, PVDHeaderPVData
If Left(PVDHeaderPVData.PVDHeaderDescriptor, 8) = "[PVData]" And PVDHeaderPVData.PVDCheck > 0 Then 'PCNGL130103 'PCN3274
    PVDataNoOfLines = PVDHeaderPVData.PVDCheck 'PCNGL130103
'    Call InitilisePVProfile(MaxFrameBufferNo) 'PCNGL130103 'PCN2970
    ReDim PVDistances(1) 'PCN2639
    For PVDataLineNo = 0 To 1 'Get the first 2 profiles only 'PCNGL140103
        For PVSegmentNo = 1 To NoOfProfileSegments
            Get #FileNo, , pvData(PVSegmentNo, 0, PVDataLineNo)
        Next PVSegmentNo
        Get #FileNo, , pvCapacityData(PVDataLineNo) 'PCNGL1301032
        Get #FileNo, , PVOvalityData(PVDataLineNo) 'PCNGL1301032
        Get #FileNo, , PVDelta(0, PVDataLineNo) 'PCNGL1301032
        Get #FileNo, , PVDelta(1, PVDataLineNo) 'PCNGL1301032
        'vvvv **** load AVI frame time *************************** 'PCNGL150103
        'To be used to more accurately link the PVD file Frame no to the AVI frames. PCNGL150103
        Get #FileNo, , AVIFrameTime(PVDataLineNo)
        '^^^^ ****************************************************
        Get #FileNo, , PVDistances(PVDataLineNo) 'PCN2639
    Next PVDataLineNo
End If

' Close before reopening in another mode.
Close #FileNo

'Update the forms with the loaded data 'PCNGL130103
Call CopyConfigInfoToForms
Call CopyPipeDetailsToPipelineForm
 'PCN4131 Call CopyPipeDetailsToObsForm
    
    
Exit Sub
Err_Handler:
    Select Case Err
        Case 9 ' out of range
        Case Else
            MsgBox Err & "-PF29:" & Error$
    End Select
End Sub

Sub LoadInPVDFormat_V40(FileName As String, FileLoadError As Boolean, ErrorType As String)      'PCNGL120103 'PCN3964
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'LoadInPVDFormat_V40 Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff,    12/01/03     Building initial framework
'
'Description:
'       Loads the Precision Vision data from file, in binary format, version 4.0 (PCNGL110301)
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim PVSegmentNo As Integer
Dim PVDataLineNo As Long
Dim PipeInfoNoOfLines As Integer
Dim NodeInfoNoOfLines As Integer
Dim DrawInfoNoOfLines As Integer

FileLoadError = False 'PCNGL140103
ErrorType = "Unknown"  'PCN3964

If Dir(FileName) = "" Or FileName = "" Then  'PCNGL160103
    'MsgBox DisplayMessage("File was not found."), vbInformation 'PCN2111
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("File was not found."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    FileLoadError = True 'PCNGL140103
    Exit Sub
End If

ClearLineTitle.TitleBarCaption.Caption = DisplayMessage("Precision Vision") & " - " & FileName 'PCNGL130103 'PCNGL140103 'PCN2111 'PCN2759
ClearLineTitle.TitleBarCaption.ToolTipText = DisplayMessage("Precision Vision") & " - " & FileName 'PCN4171

'Testing file reading speed over network and to local disk. Over the network in currently very slow.
'Open "C:\CleanFlow_UnderTest\ClearLineProfilerV4\TESTFILE.pvd" For Binary Access Read Lock Read As #1
'Open "Z:\CleanFlow\PreProduction\ClearLineProfilerV4\TESTFILE.pvd" For Binary Access Read Lock Read As #1

Dim FileNo As Integer
FileNo = FreeFile
Open FileName For Binary Access Read Lock Write As #FileNo 'PCN2208

'Load the File Main Header
Get #FileNo, , PVDFileMainHeader

'vvvv PCN2207 ********************************************
'Check to see that this file was produced by a registered
'copy of the application. Do not open this file if it is
'not registered.
If Left(PVDFileMainHeader.PVDFileMHAppName, 12) = "Unregistered" Or PVDFileMainHeader.PVDFileMHVersionMajor = -100 Or PVDFileMainHeader.PVDFileMHVersionMinor = -100 Or PVDFileMainHeader.PVDFileMHVersionRev = -100 Then
    'MsgBox DisplayMessage("Can't read file, ") & FileName, vbExclamation  'PCN2111
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Can't read file, ") & FileName: ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Close #FileNo
    FileLoadError = True
    Exit Sub
End If
'^^^^ ****************************************************

'Determine file header pointers and CheckSums then read the header pointers to block 2
'File header pointers (must be the second block of data in any PVD file, there are no fix number of pointers)
If PVDFileMainHeader.PVDFileMHPointerAddress = 0 Then 'PCNGL140103
    'MsgBox DisplayMessage("Can't read file, ") & FileName 'PCN2111
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Can't read file, ") & FileName: ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Close #FileNo
    FileLoadError = True 'PCNGL140103
    Exit Sub
End If
Get #FileNo, PVDFileMainHeader.PVDFileMHPointerAddress, PVDFilePointers


'Read from file the System Configuration Information 'PCNGL130103
PVDHeaderConfigInfo.PVDHeaderDescriptor = ""
PVDHeaderConfigInfo.PVDCheck = 0
Get #FileNo, PVDFilePointers.PVDPointerConfigInfo, PVDHeaderConfigInfo
'vvvv PCN2392 ***************************************
'Modification made for Fish Eye Distortion
'Call LoadConfigInfo_V40(Seek(1))
Call LoadConfigInfo_V41(Seek(FileNo), FileLoadError, FileNo) 'PCN2639
'^^^^ ***********************************************
'vvvv PCN2952 *********************
If FileLoadError Then
    'Not in Version 4.X format
    Close #FileNo
    Exit Sub
End If
'^^^^ *****************************

'PCN1919
If (ConfigInfo.Units = "mm" And MeasurementUnits = "in") Or (ConfigInfo.Units = "in" And MeasurementUnits = "mm") Then
    'If LoadVideo = True Then MsgBox DisplayMessage("The PVD file is in a different measurement unit, please alter your measurement units before loading this file."), vbExclamation 'PCN2111
    If LoadVideo = True Then ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("The PVD file is in a different measurement unit, please alter your measurement units before loading this file."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Close #FileNo
'PCN 1984 LS 10/7/03
    FileLoadError = True
    ErrorType = "Wrong Units" 'PCN3964
    OptionsPage.Show
    PVDFileName = "" 'PCN3964
    OptionsPage.ZOrder 0 'PCN3964
    Exit Sub
ElseIf (ConfigInfo.Units <> "mm" And MeasurementUnits <> "in") Then 'PCN2207
    'vvvv PCN2207 ********************************************
    'MsgBox DisplayMessage("Can't read file, ") & FileName, vbExclamation  'PCN2111
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Can't read file, ") & FileName: ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Close #FileNo
    FileLoadError = True
    Exit Sub
    '^^^^ ****************************************************
End If

'Read from file the Pipe Information
PVDHeaderPipeInfo.PVDHeaderDescriptor = ""
PVDHeaderPipeInfo.PVDCheck = 0
Get #FileNo, PVDFilePointers.PVDPointerPipeInfo, PVDHeaderPipeInfo
Call LoadPipelineDetails_V40(Seek(FileNo), FileNo) 'PCNGL1201032


'Read from file the Pipe Observations 'PCNGL130103
PVDHeaderPipeObs.PVDHeaderDescriptor = ""
PVDHeaderPipeObs.PVDCheck = 0
Get #FileNo, PVDFilePointers.PVDPointerPipeObs, PVDHeaderPipeObs
Call LoadPipelineObs_V40(Seek(FileNo), PVDHeaderPipeObs.PVDCheck, FileNo)


'Read from file the Font Information
PVDHeaderFontInfo.PVDHeaderDescriptor = ""
PVDHeaderFontInfo.PVDCheck = 0
Get #FileNo, PVDFilePointers.PVDPointerFontInfo, PVDHeaderFontInfo


'Read from file the Drawing Information
PVDHeaderDrawInfo.PVDHeaderDescriptor = ""
PVDHeaderDrawInfo.PVDCheck = 0 'Are there any lines to draw?
Get #FileNo, PVDFilePointers.PVDPointerDrawInfo, PVDHeaderDrawInfo
'Determine and read drawing data



'Read from file the PVData
PVDHeaderPVData.PVDHeaderDescriptor = ""
PVDHeaderPVData.PVDCheck = 0
Get #FileNo, PVDFilePointers.PVDPointerPVData, PVDHeaderPVData
If Left(PVDHeaderPVData.PVDHeaderDescriptor, 8) = "[PVData]" And PVDHeaderPVData.PVDCheck > 0 Then 'PCNGL130103 'PCN3274
    PVDataNoOfLines = PVDHeaderPVData.PVDCheck 'PCNGL130103
    'Call InitilisePVProfile(PVDataNoOfLines) 'PCNGL130103
    'For PVDataLineNo = 0 To PVDataNoOfLines
    Call InitilisePVProfile(MaxFrameBufferNo) 'PCNGL130103
    For PVDataLineNo = 0 To 1 'Get the first 2 profiles only 'PCNGL140103
        For PVSegmentNo = 1 To NoOfProfileSegments
            Get #FileNo, , pvData(PVSegmentNo, 0, PVDataLineNo)
        Next PVSegmentNo
        Get #FileNo, , pvCapacityData(PVDataLineNo) 'PCNGL1301032
        Get #FileNo, , PVOvalityData(PVDataLineNo) 'PCNGL1301032
        Get #FileNo, , PVDelta(0, PVDataLineNo) 'PCNGL1301032
        Get #FileNo, , PVDelta(1, PVDataLineNo) 'PCNGL1301032
        'vvvv **** load AVI frame time *************************** 'PCNGL150103
        'To be used to more accurately link the PVD file Frame no to the AVI frames. PCNGL150103
        Get #FileNo, , AVIFrameTime(PVDataLineNo)
        '^^^^ ****************************************************
        'DoEvents
    Next PVDataLineNo
End If

' Close before reopening in another mode.
Close #FileNo

'Update the forms with the loaded data 'PCNGL130103
Call CopyConfigInfoToForms
Call CopyPipeDetailsToPipelineForm
'PCN4131 Call CopyPipeDetailsToObsForm
    
    
Exit Sub
Err_Handler:
    Select Case Err
        Case 9 ' out of range
        Case Else
            MsgBox Err & "-PF30:" & Error$
    End Select
End Sub

Sub PVLandscapeReport(CurrentPVPage As Form)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PVLandscapeReport Function  Michelle Lindsay   michellelindsay@cbsys.co.nz
'
'Revision history"
'   V0.0    Michelle Lindsay,    09/01/03     Building initial framework
'
'Description:
'       Displays the PVLandscape report when the user clicks any PVLandscape
'       report button throughout the application.
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

'GraphReportingType = "PVLandscape"
'CurrentPVPage.PopupReportsToolbar.Visible = False
'If isopen("PVGraphReport") Then Unload PVGraphReport
'PVGraphReport.Show

Exit Sub
Err_Handler:
    Select Case Err
    Case 482 'Printer not connected
        Resume Next
    Case Else
    MsgBox Err & "-PF31:" & Error$
End Select
End Sub

Sub PVPortraitReport(CurrentPVPage As Form)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PVLandscapeReport Function  Michelle Lindsay   michellelindsay@cbsys.co.nz
'
'Revision history"
'   V0.0    Michelle Lindsay,    10/01/03     Building initial framework
'
'Description:
'       Displays the PVPortrait report when the user clicks any PVLandscape
'       report button throughout the application.
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

'GraphReportingType = "PVPortrait"
'CurrentPVPage.PopupReportsToolbar.Visible = False
'If isopen("LaserImageReportPVG") Then Unload LaserImageReportPVG
'LaserImageReportPVG.Show

Exit Sub
Err_Handler:
    Select Case Err
    Case 482 'Printer not connected
        Resume Next
    Case Else
        MsgBox Err & "-PF32:" & Error$
    End Select
End Sub

Sub CopyPipelineFormToPipeDetails() 'PCNGL130103 'PCN1833
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'CopyPipelineFormToPipeDetails Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff,    11/12/02     Building initial framework
'
'Description:
'       Testing
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

If Len(PipelineDetails.InternalDiameterExpected) <> 0 Then
    PipelineInfo.IntDiameter = PipelineDetails.InternalDiameterExpected
Else
    PipelineInfo.IntDiameter = 0
End If
If Len(PipelineDetails.PipeLength) <> 0 Then
    PipelineInfo.PipeLength = PipelineDetails.PipeLength
Else
    PipelineInfo.PipeLength = 0
End If
PipelineInfo.Material = PipelineDetails.Material

PipelineInfo.AssetNo = PipelineDetails.AssetNo

PipelineInfo.SiteID = PipelineDetails.SiteID

PipelineInfo.City = PipelineDetails.City

If Len(PipelineDetails.sDate) <> 0 Then 'PCNGL170103
    PipelineInfo.Date = PipelineDetails.sDate
Else
    PipelineInfo.Date = Date
End If

If Len(PipelineDetails.sTime) <> 0 Then 'PCNGL170103
    PipelineInfo.Time = PipelineDetails.sTime
Else
    PipelineInfo.Time = Time
End If

PipelineInfo.StartName = PipelineDetails.StartNodeNo

PipelineInfo.StartLocation = PipelineDetails.StartNodeLocation

PipelineInfo.FinishName = PipelineDetails.FinishNodeNo

PipelineInfo.FinishLocation = PipelineDetails.FinishNodeLocation

PipelineInfo.Comments = PipelineDetails.GeneralComments.text  'PCN4171

Exit Sub
Err_Handler:
    Select Case Err
        Case 13: Resume Next 'type mismatch error 'PCN3744 bad data loaded from interface file
        Case Else: MsgBox Err & "-PF33:" & Error$
    End Select
 End Sub


Sub CopyFormToConfigInfo() 'PCNGL1301032 'PCN1833
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'CopyFormToConfigInfo Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff,    13/12/02     Building initial framework
'
'Description:
'       Testing
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler


ConfigInfo.WLStartAngle = WLStartAngle
ConfigInfo.WLFinishAngle = WLFinishAngle

ConfigInfo.CalDist = CalLen_Global
ConfigInfo.CalLineLength = CalLength_Global

ConfigInfo.Units = PipelineDetails.unit1.Caption

ConfigInfo.NoOfProfileSegments = NoOfProfileSegments 'PCNGL140103

ConfigInfo.VideoFileName = VideoFileName 'PCNGL140103

Call GetINI_ParameterInfoOnly(MyFile, "CameraModel=", FontName) 'PCN4799 FontName is temporary used for camera model for now

Exit Sub
Err_Handler:
    MsgBox Err & "-PF34:" & Error$
    
End Sub

Function Validate_PipeLineInfo_V40(Validate As Control) As Boolean

End Function

Function SaveToFileFontInfo(FileSaveFail As Boolean) As Boolean  'PCNGL230103
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'SaveToFilePipeAndConfigInfo Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff,    23/01/03     Building initial framework
'
'Description:
'
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim SaveToAddress As Long 'Store the address where you wish to save the data
Dim FileNumber
FileSaveFail = False 'PCN1768
'FileSaveFail = True

If Dir(PVDFileName) = "" Or PVDFileName = "" Then Exit Function

'Check whether a file is open
FileNumber = FreeFile
Open PVDFileName For Binary Access Read Lock Write As #FileNumber 'PCN2208
'Load the File Main Header

            Get #FileNumber, PVDFilePointers.PVDPointerFontInfo, PVDHeaderPipeInfo
            If Left(PVDHeaderPipeInfo.PVDHeaderDescriptor, 10) = "[FontInfo]" And PVDHeaderPVData.PVDCheck <> 0 Then
                'Save Configuration Information to file
                SaveToAddress = Seek(FileNumber)
            Else
                FileSaveFail = True
            End If

Close #FileNumber
If FileSaveFail Then Exit Function  'PCN1768
If ThisFileIsReadOnly(PVDFileName) Then Exit Function 'PCN4241 - Check to see if the PVD is read only
FileNumber = FreeFile
Open PVDFileName For Binary Access Write As #FileNumber
        Put #FileNumber, SaveToAddress, FontInfo 'PCNGL250203


Close #FileNumber
FileSaveFail = False

Exit Function
FileErr_Handler:
    FileSaveFail = True
    Close #FileNumber

Exit Function
Err_Handler:
Select Case Err
    Case 52 'Bad filename or number 'PCN1863 I haven't worked out why I get this error
        Resume Next
    Case 63 'Bad record number 'PCN1863 I haven't worked out why I get this error
        Resume Next
    Case Else
        MsgBox Err & "-PF35:" & Error$
        GoTo FileErr_Handler
End Select
End Function

Sub SaveToFilePipeAndConfigInfo(WhatToSave As String, FileSaveFail As Boolean)   'PCNGL230103
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'SaveToFilePipeAndConfigInfo Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff,    23/01/03     Building initial framework
'
'Description:
'
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim SaveToAddress As Long 'Store the address where you wish to save the data
Dim FileNumber
FileSaveFail = False 'PCN1768
'FileSaveFail = True

If Dir(PVDFileName) = "" Or PVDFileName = "" Then Exit Sub

'Check whether a file is open
FileNumber = FreeFile
Open PVDFileName For Binary Access Read Lock Write As #FileNumber 'PCN2208
'Load the File Main Header

Get #FileNumber, , PVDFileMainHeader

'Read the file header pointers
If PVDFileMainHeader.PVDFileMHPointerAddress <> 0 Then
    Get #FileNumber, PVDFileMainHeader.PVDFileMHPointerAddress, PVDFilePointers
    Select Case WhatToSave
        Case "ConfigInfo"
            Get #FileNumber, PVDFilePointers.PVDPointerConfigInfo, PVDHeaderConfigInfo
            If Left(PVDHeaderConfigInfo.PVDHeaderDescriptor, 12) = "[ConfigInfo]" And PVDHeaderPVData.PVDCheck <> 0 Then
                'Save Configuration Information to file
                SaveToAddress = Seek(FileNumber)
            Else
                FileSaveFail = True
            End If
        Case "PipelineInfo"
            Get #FileNumber, PVDFilePointers.PVDPointerPipeInfo, PVDHeaderPipeInfo
            If Left(PVDHeaderPipeInfo.PVDHeaderDescriptor, 10) = "[PipeInfo]" And PVDHeaderPVData.PVDCheck <> 0 Then
                'Save Configuration Information to file
                SaveToAddress = Seek(FileNumber)
            Else
                FileSaveFail = True
            End If
        Case Else
            FileSaveFail = True
    End Select
Else
    FileSaveFail = True
End If
Close #FileNumber
If FileSaveFail Then Exit Sub       'PCN1768
If ThisFileIsReadOnly(PVDFileName) Then Exit Sub      'PCN4241 - Check to see if the PVD is read only
FileNumber = FreeFile
Open PVDFileName For Binary Access Write As #FileNumber
Select Case WhatToSave
    Case "ConfigInfo"
        'vvvv PCN2492 **********************
        'Check version
        'vvvv PCN3019 ******************************************
        If Trim(ConfigInfo.PVDFileVersion) = "V6.1" Or _
           Trim(ConfigInfo.PVDFileVersion) = "V6.2" Or _
           Trim(ConfigInfo.PVDFileVersion) = "6.25" Or _
           Trim(ConfigInfo.PVDFileVersion) = "V6.3" Or _
           Trim(ConfigInfo.PVDFileVersion) = "V6.4" Then 'PCN3219 'PCN3576 added "V6.2" 'PCN6004 add "V6.3" PVData file now singles

'ID5395
'           If PossibleConfigInfoCurruption Then
'                CopyFromGoodConfigInfoToCurruption
'                Put #FileNumber, SaveToAddress, ConfigInfo_currupting
'
'           Else
                Put #FileNumber, SaveToAddress, ConfigInfo
'            End If
'''''

        ElseIf Trim(ConfigInfo.PVDFileVersion) = "V6.0" Then
            Call ConvertConfigInfoToConfigInfo_V60
            Put #FileNumber, SaveToAddress, ConfigInfo_V60
        ElseIf Trim(ConfigInfo.PVDFileVersion) = "V5.3" Then
            'Note ConfigInfo_V53 does not exist.
            Call ConvertConfigInfoToConfigInfo_V52
            Put #FileNumber, SaveToAddress, ConfigInfo_V52
        ElseIf Trim(ConfigInfo.PVDFileVersion) = "V5.2" Then
            Call ConvertConfigInfoToConfigInfo_V52
            Put #FileNumber, SaveToAddress, ConfigInfo_V52
        '^^^^ **************************************************
        ElseIf Trim(ConfigInfo.PVDFileVersion) = "V5.1" Then 'PCN2850 'PCN3019
            'Copy ConfigInfo to ConfigInfo_V50 'PCN2850
            'Note ConfigInfo_V51 does not exist.
            Call ConvertConfigInfoToConfigInfo_V50 'PCN2850
            Put #FileNumber, SaveToAddress, ConfigInfo_V50 'PCN2850
        ElseIf Trim(ConfigInfo.PVDFileVersion) = "V5.0" Then 'PCN2850
            'Copy ConfigInfo to ConfigInfo_V50 'PCN2850
            Call ConvertConfigInfoToConfigInfo_V50 'PCN2850
            Put #FileNumber, SaveToAddress, ConfigInfo_V50 'PCN2850
        ElseIf Trim(ConfigInfo.PVDFileVersion) = "V4.1" Then 'PCN2639
            'Copy ConfigInfo to ConfigInfo_V41 'PCN2639
            Call ConvertConfigInfoToConfigInfo_V41 'PCN2639
            Put #FileNumber, SaveToAddress, ConfigInfo_V41 'PCN2639
        ElseIf Trim(ConfigInfo.PVDFileVersion) = "V4.0" Then
            'Copy ConfigInfo to ConfigInfo_V40
            Call ConvertConfigInfoToConfigInfo_V40
            Put #FileNumber, SaveToAddress, ConfigInfo_V40
        Else
'            Put #1, SaveToAddress, ConfigInfo 'PCN3019
        End If
        '^^^^ ******************************
    Case "PipelineInfo"
        Put #FileNumber, SaveToAddress, PipelineInfo 'PCNGL250203
        'Put #1, 0, PipelineInfo
    Case Else
        FileSaveFail = True
End Select


Close #FileNumber
FileSaveFail = False

Exit Sub
FileErr_Handler:
    FileSaveFail = True
    Close #FileNumber

Exit Sub
Err_Handler:
Select Case Err
    Case 52 'Bad filename or number 'PCN1863 I haven't worked out why I get this error
        Resume Next
    Case 63 'Bad record number 'PCN1863 I haven't worked out why I get this error
        Resume Next
    Case Else
        MsgBox Err & "-PF35:" & Error$
        GoTo FileErr_Handler
End Select
End Sub


Sub FishEyeLoadFileCheck(FileClass As String)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : FishEyeLoadFileCheck
'Created : 12 November 2003, PCN2392
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    : This function checks if there is a current FishEyeDistortion value
'           and if so, prompts user on its application to the video file being
'           loaded.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim FishMsgbox As String
Dim FecCameraModel As String 'PCN3595 fisheye always on (25 Oct 2005, Antony)
Dim ExtraInfo As String
Dim FecCeameraModelCurrent As String

Screen.MousePointer = vbHourglass  '

Call GetINI_ParameterInfoOnly(MyFile, "CameraModel=", FecCameraModel)

'^^^^ ************************************

If FileClass = "PVD" Then
    Call FisheyeFunctions.InitializeFishEyeForPVD
    Call FisheyeFunctions.FEON
    Call FisheyeFunctions.DisableFishEye
    If IsOpen("Fisheye") Then
        Fisheye.CameraDropdown.Enabled = False 'PCN3595 do not allow fec selection
        Fisheye.FishEyeON.Enabled = False
        Fisheye.CameraDropdown.text = ""
        
    End If
    
    'PCN4799 now the camera model is stored in the PVD, as fontinfo.fontname for now
    'and if there is no camera info then change to unkown
    
    'If GetPVDVer < 6.4 Then FecCameraModel = "unknown"
    FecCameraModel = Trim(FontInfo.FontName)
    If FecCameraModel = "" Then FecCameraModel = "unknown"
    Fisheye.CameraDropdown.text = FecCameraModel
    Fisheye.Visible = False
    
    Call GetINI_ParameterInfoOnly(MyFile, "CameraModel=", FecCeameraModelCurrent) 'PCN3595 (21 Oct 2005, Antony)
    
    If FecCameraModel = FecCeameraModelCurrent Then
        Call INI_WriteBack(MyFile, "Fish_DistortionHorizontal=", ConfigInfo.FishEyeHorDistortion)
        Call INI_WriteBack(MyFile, "CalibrationDistance=", ConfigInfo.CalDist)
        Call INI_WriteBack(MyFile, "CalibrationLineLength=", ConfigInfo.CalLineLength)
        Call INI_WriteBack(MyFile, "Fish_Ratio=", ConfigInfo.FishEyeRatio)
    End If
    
Else

    'FishMsgbox = MsgBox(DisplayMessage("Apply current lens distortion settings?"), vbInformation + vbYesNo) 'PCN2838
    
    'FishMsgbox = MsgBox(DisplayMessage("Apply current lens distortion settings? " & ), vbInformation + vbYesNo) 'PCN2838
    

    'If FishMsgbox = vbYes Then  'PCN2856 'PCN3595 removed because fisheye is now always on
                                        '
'    If ConfigInfo.MediaWidth > 1000 And ConfigInfo.MediaHeight > 900 Then
'            Call FisheyeFunctions.FEOFF 'PCNANT  if its a 3x3 video no fisheye
'            Fisheye.FishEyeON.Enabled = False   'PCNANT
'            Screen.MousePointer = vbDefault    '
'            Exit Function
'    End If
    
    
    Call FisheyeFunctions.FEON

    Call FisheyeFunctions.InitializeFishEyeFromINI
    
    'PCN4474
    If ConfigInfo.FishEyeDistortion = 0 And ConfigInfo.FishEyeHorDistortion = 1 Then
        FisheyeFunctions.FEOFF
    Else
        FisheyeFunctions.FEON
    End If
    
    If IsOpen("Fisheye") Then
        Fisheye.CameraDropdown.Enabled = True
        Fisheye.FishEyeON.Enabled = True    'PCN3595 do allow fec selection
        Fisheye.btnLoad.Enabled = True
    End If

    If ConfigInfo.FishEyeFlag = True Then 'PCN4474
        Call FisheyeFunctions.CreateFishEyeMask
    End If
    Screen.MousePointer = vbDefault    '

End If 'PCN4799 moved from below the fisheye setting information, so now it can display camera info

If ConfigInfo.FishEyeHorDistortion <> 1 Then
    ExtraInfo = Chr(13) & DisplayMessage("Vertical video adjustment") & " - " & Round(ConfigInfo.FishEyeHorDistortion, 5) * 100 & "%."
Else
    ExtraInfo = Chr(13) & ""
End If
    
        'FishMsgbox = MsgBox(DisplayMessage("Lens Distortion settings have been applied.") & _
            Chr(13) & DisplayMessage("Camera") & " - " & FecCameraModel & ExtraInfo, vbInformation + vbOKOnly) 'PCN3595
        ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Lens Distortion settings have been applied.") & _
            Chr(13) & DisplayMessage("Camera") & " - " & FecCameraModel & ExtraInfo: ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0



'End If 'PCN3595
Exit Sub
Err_Handler:
    MsgBox Err & "-PF36:" & Error$
End Sub

Sub ConvertConfigInfoToConfigInfo_V40()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : ConvertConfigInfoToConfigInfo_V40
'Created : 16 November 2003, PCN2492
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

ConfigInfo_V40.WLStartAngle = ConfigInfo.WLStartAngle
ConfigInfo_V40.WLFinishAngle = ConfigInfo.WLFinishAngle
ConfigInfo_V40.FishEyeHorDistortion = ConfigInfo.FishEyeHorDistortion 'PCN3687 was area that was never used
ConfigInfo_V40.VideoFileName = ConfigInfo.VideoFileName
ConfigInfo_V40.CalDist = ConfigInfo.CalDist
ConfigInfo_V40.CalLineLength = ConfigInfo.CalLineLength
ConfigInfo_V40.FileCountryCode = ConfigInfo.FileCountryCode
ConfigInfo_V40.FileLanguage = ConfigInfo.FileLanguage
ConfigInfo_V40.LenReal = ConfigInfo.LenReal
ConfigInfo_V40.LenRealPercent = ConfigInfo.LenRealPercent
ConfigInfo_V40.MediaHeight = ConfigInfo.MediaHeight
ConfigInfo_V40.MediaWidth = ConfigInfo.MediaWidth
ConfigInfo_V40.NoOfProfileSegments = ConfigInfo.NoOfProfileSegments
ConfigInfo_V40.Ratio = ConfigInfo.Ratio
ConfigInfo_V40.Units = ConfigInfo.Units

Exit Sub
Err_Handler:
    MsgBox Err & "-PF37:" & Error$
End Sub

Sub ConvertConfigInfoToConfigInfo_V41()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : ConvertConfigInfoToConfigInfo_V41
'Created : 23 March 2004, PCN2639
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

ConfigInfo_V41.PVDFileVersion = "V4.1" 'PCN2850
ConfigInfo_V41.WLStartAngle = ConfigInfo.WLStartAngle
ConfigInfo_V41.WLFinishAngle = ConfigInfo.WLFinishAngle
ConfigInfo_V41.FishEyeHorDistortion = ConfigInfo.FishEyeHorDistortion 'PCN3687 was area that was never used
ConfigInfo_V41.VideoFileName = ConfigInfo.VideoFileName
ConfigInfo_V41.CalDist = ConfigInfo.CalDist
ConfigInfo_V41.CalLineLength = ConfigInfo.CalLineLength
ConfigInfo_V41.FileCountryCode = ConfigInfo.FileCountryCode
ConfigInfo_V41.FileLanguage = ConfigInfo.FileLanguage
ConfigInfo_V41.LenReal = ConfigInfo.LenReal
ConfigInfo_V41.LenRealPercent = ConfigInfo.LenRealPercent
ConfigInfo_V41.MediaHeight = ConfigInfo.MediaHeight
ConfigInfo_V41.MediaWidth = ConfigInfo.MediaWidth
ConfigInfo_V41.NoOfProfileSegments = ConfigInfo.NoOfProfileSegments
ConfigInfo_V41.Ratio = ConfigInfo.Ratio
ConfigInfo_V41.Units = ConfigInfo.Units
ConfigInfo_V41.FishEyeCenterX = ConfigInfo.FishEyeCenterX
ConfigInfo_V41.FishEyeCenterY = ConfigInfo.FishEyeCenterY
ConfigInfo_V41.FishEyeDistortion = ConfigInfo.FishEyeDistortion
ConfigInfo_V41.FishEyeFlag = ConfigInfo.FishEyeFlag
ConfigInfo_V41.FishEyeRatio = ConfigInfo.FishEyeRatio

Exit Sub
Err_Handler:
    MsgBox Err & "-PF38:" & Error$
End Sub

Function PVDataAddressOffsetCalc(StartAddress As Long, FrameNo As Long) As Long
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PVDataAddressOffset
'Created : 22 March 2004, PCN2639
'Updated :
'Prg By  : Geoff Logan
'Param   : StartAddress - Start address of the PV data block
'          FrameNo - Frame number of desired PVData
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim AddressCalc As Long

AddressCalc = StartAddress + FrameNo * PVDataFrameBlockSize  'PCN2639
PVDataAddressOffsetCalc = AddressCalc + FrameNo * (PVCalculationsBlockSize + PVRelatedInfoBlockSize)   'PCN2639

Exit Function
Err_Handler:
    MsgBox Err & "-PF39:" & Error$
End Function


Public Sub RapidSavePVDistanceToFile(CurrentFrameNo As Long, PVDataStartAddress As Long, PVDFileSaveFail As Boolean, FileNo As Integer)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : RapidSavePVDistanceToFile
'Created : 23 March 2004, PCN2639
'Updated :
'Prg By  : Geoff Logan
'Param   : CurrentFrameNo -
'          PVDataAddressOffset -
'          PVDFileSaveFail -
'Desc    :
'Usage   :
'***************************************************************
On Error GoTo Err_Handler
Dim ErrorStatus As String
Dim PVAddressOffset As Long

If Trim(ConfigInfo.PVDFileVersion) <> "V4.0" And Trim(ConfigInfo.PVDFileVersion) <> "V4.0" Then
    PVAddressOffset = PVDataStartAddress + (CurrentFrameNo + 1) * PVDataFrameBlockSize
    PVAddressOffset = PVAddressOffset + (CurrentFrameNo) * (PVCalculationsBlockSize + PVRelatedInfoBlockSize)
    PVAddressOffset = PVAddressOffset + PVCalculationsBlockSize + Len(PVTimes(0))
    Put #FileNo, PVAddressOffset, PVDistances(CurrentFrameNo)
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
        MsgBox Err & "-PF40:" & Error$
End Select
End Sub

Sub ConvertConfigInfoToConfigInfo_V50()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : ConvertConfigInfoToConfigInfo_V50
'Created : 31 May 2004, PCN2850
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    : Copies the current ConfigInfo information into V5.0 of the ConfigInfo.
'Usage   : When a V5.0 PVD file is loaded, all saving to PVD must be in the same
'           Version format.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

ConfigInfo_V50.PVDFileVersion = Trim(ConfigInfo.PVDFileVersion) 'PCN2850
ConfigInfo_V50.WLStartAngle = ConfigInfo.WLStartAngle
ConfigInfo_V50.WLFinishAngle = ConfigInfo.WLFinishAngle
ConfigInfo_V50.FishEyeHorDistortion = ConfigInfo.FishEyeHorDistortion 'PCN3687 was area that was never used
ConfigInfo_V50.VideoFileName = ConfigInfo.VideoFileName
ConfigInfo_V50.CalDist = ConfigInfo.CalDist
ConfigInfo_V50.CalLineLength = ConfigInfo.CalLineLength
ConfigInfo_V50.FileCountryCode = ConfigInfo.FileCountryCode
ConfigInfo_V50.FileLanguage = ConfigInfo.FileLanguage
ConfigInfo_V50.LenReal = ConfigInfo.LenReal
ConfigInfo_V50.LenRealPercent = ConfigInfo.LenRealPercent
ConfigInfo_V50.MediaHeight = ConfigInfo.MediaHeight
ConfigInfo_V50.MediaWidth = ConfigInfo.MediaWidth
ConfigInfo_V50.NoOfProfileSegments = ConfigInfo.NoOfProfileSegments
ConfigInfo_V50.Ratio = ConfigInfo.Ratio
ConfigInfo_V50.Units = ConfigInfo.Units
ConfigInfo_V50.FishEyeCenterX = ConfigInfo.FishEyeCenterX
ConfigInfo_V50.FishEyeCenterY = ConfigInfo.FishEyeCenterY
ConfigInfo_V50.FishEyeDistortion = ConfigInfo.FishEyeDistortion
ConfigInfo_V50.FishEyeFlag = ConfigInfo.FishEyeFlag
ConfigInfo_V50.FishEyeRatio = ConfigInfo.FishEyeRatio
'V5.0 additions
ConfigInfo_V50.DistanceProcessMethod = Trim(ConfigInfo.DistanceProcessMethod)
ConfigInfo_V50.DistanceStart = ConfigInfo.DistanceStart
ConfigInfo_V50.DistanceDirection = Trim(ConfigInfo.DistanceDirection)
ConfigInfo_V50.DistanceFinish = ConfigInfo.DistanceFinish

Exit Sub
Err_Handler:
    MsgBox Err & "-PF41:" & Error$
End Sub

Public Sub RapidSaveOvalityDeltaToFile(CurrentFrameNo As Long, PVDataStartAddress As Long, PVDFileSaveFail As Boolean, FileNo As Integer)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : RapidSaveOvalityDeltaToFile
'Created : 23 March 2004, PCN Testing only for Phil to correct 2 PVDs.
'Updated :
'Prg By  : Geoff Logan
'Param   : CurrentFrameNo -
'          PVDataAddressOffset -
'          PVDFileSaveFail -
'Desc    :
'Usage   :
'***************************************************************
On Error GoTo Err_Handler
Dim ErrorStatus As String
Dim PVAddressOffset As Long

If Trim(ConfigInfo.PVDFileVersion) <> "V4.0" And Trim(ConfigInfo.PVDFileVersion) <> "V4.0" Then
    PVAddressOffset = PVDataStartAddress + (CurrentFrameNo + 1) * PVDataFrameBlockSize
    PVAddressOffset = PVAddressOffset + (CurrentFrameNo) * (PVCalculationsBlockSize + PVRelatedInfoBlockSize)
    
    Put #FileNo, PVAddressOffset, pvCapacityData(0)
    Put #FileNo, , PVOvalityData(0)
    Put #FileNo, , PVDelta(0, 0)
    Put #FileNo, , PVDelta(1, 0)
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
        MsgBox Err & "-PF42:" & Error$
End Select
End Sub

Sub CorrectPhilsData()
'31 May 2004 - NO PCN FOR THIS ONE. GL
'For conversion of Phils V5.1 PVD data (Ovality and Delta without 100x) Do not use.
On Error GoTo Err_Handler
Dim FileLoadError As Boolean
Dim PVAddressOffset As Long 'PCN2639
Dim PVDataStartAddress As Long 'PCN2639
Dim FileNo As Integer 'PCN2639
Dim CurrentFrame As Long
Dim PVGraphDataAddressOffset As Long

If PVDFileName = "" Or PVDataNoOfLines = 0 Then
    'MsgBox DisplayMessage("There is no recorded data to process")
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("There is no recorded data to process"): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Exit Sub
End If



'Find PVDataStartAddress
Call GetPVDPointerPVDataFromFile(PVDFileName, PVDataStartAddress, FileLoadError)
FileNo = 8

For CurrentFrame = 0 To PVDataNoOfLines



'Populate the PVDistances array - this may take some time.

    Open PVDFileName For Binary Access Read Lock Write As #FileNo

    PVGraphDataAddressOffset = PVDataStartAddress + (CurrentFrame + 1) * PVDataFrameBlockSize 'PCN2639
    PVGraphDataAddressOffset = PVGraphDataAddressOffset + (CurrentFrame) * (PVCalculationsBlockSize + PVRelatedInfoBlockSize)   'PCN2639
    Call RapidReadPVGraphsDataFromFile(FileNo, CurrentFrame, PVGraphDataAddressOffset, FileLoadError) 'PCN2164

    Close #FileNo

    'Correct values
    PVOvalityData(0) = PVOvalityData(0) * 100
    PVDelta(0, 0) = PVDelta(0, 0) * 100
    PVDelta(1, 0) = PVDelta(1, 0) * 100

    Open PVDFileName For Binary Access Write As #FileNo
    Call RapidSaveOvalityDeltaToFile(CurrentFrame, PVDataStartAddress, FileLoadError, FileNo)
    Close #FileNo
Next CurrentFrame


'MsgBox DisplayMessage("Finished Processing") 'PCNML220205
ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Finished Processing"): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0

TidyUp:
    Close #FileNo

Exit Sub
Err_Handler:
Select Case Err
    Case 6 'Overflow
        
        Resume Next
    Case 13 'Invalid data
        DistanceStart = InvalidData 'PCN3884 was -1
        Resume Next
    Case Else
        MsgBox Err & "-PF43:" & Error$
        GoTo TidyUp
End Select
End Sub



Sub RapidReadPVGraphsDataFromFile(OpenFileNo As Integer, FrameNo As Long, PVGraphDataAddressOffset As Long, FileLoadError As Boolean)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : RapidReadPVGraphsDataFromFile
'Created : 7 August 2003, PCN2164
'Updated : 24 March 2004, PCN2741 -  Added OpenFileNo
'Prg By  : Geoff Logan
'Param   :  FrameNo
'           PVGraphDataAddressOffset
'           FrameBufferNo
'           FileLoadError - returns a true value if an error occurs while loading.
'           OpenFileNo -
'Desc    : Gets, as fast as possible, the PVGraphs Data from the PVD file .
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim AddressOffset As Long 'PCN2639


Dim PVDVer As Single
Dim TimeIndex As Long

    
PVDVer = GetPVDVer
TimeIndex = FrameNo - 1

AddressOffset = PVGraphDataAddressOffset
If (PVDVer >= 6.3) Then
    Get #OpenFileNo, AddressOffset, TD_PVCentreX(FrameNo) 'PCNGL1301032 PCN3540
    Get #OpenFileNo, , TD_PVCentreY(FrameNo)
Else
    TD_PVCentreX(FrameNo) = 0
    TD_PVCentreY(FrameNo) = 0
End If
If TD_PVCentreX(FrameNo) > 10000 Or TD_PVCentreX(FrameNo) < -10000 Or _
       TD_PVCentreY(FrameNo) > 10000 Or TD_PVCentreY(FrameNo) < -10000 Then
        TD_PVCentreX(FrameNo) = 0
        TD_PVCentreY(FrameNo) = 0
End If
    
AddressOffset = AddressOffset + PVCalculationsBlockSize  'PCN2639
Get #OpenFileNo, AddressOffset, PVTimes(TimeIndex) 'PCNls
If FrameNo > 2 Then
    If PVTimes(TimeIndex) - PVTimes(TimeIndex - 1) > 1 Or _
       PVTimes(TimeIndex) - PVTimes(TimeIndex - 1) < 0 Then
       LoadingTimeStampError = True
    End If
End If

'PCN???? When version pvd ver les than 6.3, some of the times were recorded at index 2 instead of index 1
If TimeIndex = 2 Then
    If PVTimes(1) - PVTimes(0) > 1 Or _
       PVTimes(1) - PVTimes(0) < 0 Then
        PVTimes(0) = PVTimes(1) - (PVTimes(2) - PVTimes(1))
    End If
End If


If Trim(ConfigInfo.PVDFileVersion) <> "V4.0" And Trim(ConfigInfo.PVDFileVersion) <> "V4.1" Then 'PCN2639
    Get #OpenFileNo, , PVDistances(FrameNo) 'PCN2639
End If
If PVDistances(FrameNo) > 10000 Or PVDistances(FrameNo) < -10000 Then
    If FrameNo <> 0 Then PVDistances(FrameNo) = PVDistances(FrameNo - 1)
    If FrameNo = 0 Then PVDistances(FrameNo) = 0
End If
    



Exit Sub
Err_Handler:
    Select Case Err
        Case 9 'Out of subscript range
            Exit Sub
        Case Else
            MsgBox Err & "-PF44:" & Error$
    End Select
    FileLoadError = True
End Sub

Sub FixTimeStampErrors()
On Error GoTo Err_Handler

Dim StartTime As Double
Dim EndTime As Double
Dim TimeStep As Double
Dim FrameNo As Long
Dim TrueEnd As Long
Dim TDArraySize As Long

If PVDataNoOfLines < 1 Then Exit Sub
For TrueEnd = PVDataNoOfLines To 1 Step -1
    If PVTimes(TrueEnd) <> 0 Then Exit For
Next TrueEnd

If TrueEnd > 1 And TrueEnd < PVDataNoOfLines Then
    PVDataNoOfLines = TrueEnd
    TDArraySize = NoOfProfileSegments * (PVDataNoOfLines + 1)

    ReDim Preserve PVXDiameterFullData(PVDataNoOfLines) As Double
    ReDim Preserve PVYDiameterFullData(PVDataNoOfLines) As Double
    ReDim Preserve TD_PVDataX(TDArraySize)
    ReDim Preserve TD_PVDataY(TDArraySize)
    ReDim Preserve TD_PVCentreX(PVDataNoOfLines)
    ReDim Preserve TD_PVCentreY(PVDataNoOfLines)
    ReDim Preserve PVXDiameterFullData(PVDataNoOfLines)
    ReDim Preserve PVYDiameterFullData(PVDataNoOfLines)
    ReDim Preserve PVFlat3DRed(NoOfProfileSegments, PVDataNoOfLines) As Long
    ReDim Preserve PVFlat3DGreen(NoOfProfileSegments, PVDataNoOfLines) As Long
    ReDim Preserve PVFlat3DBlue(NoOfProfileSegments, PVDataNoOfLines) As Long
    ReDim Preserve PVDiameterMedian(PVDataNoOfLines)
    'ReDim Preserve PVFractile(PVDataNoOfLines) 'PCN4235
    ReDim Preserve PVCapacityFullData(PVDataNoOfLines) 'PCN3540
    
    ReDim Preserve GraphInfoContainer(PVOvality).DataSingle(PVDataNoOfLines)

    
    ReDim Preserve GraphInfoContainer(PVDebris).DataSingle(PVDataNoOfLines) 'PCN4461
    ReDim Preserve PVDeltaFullMax(PVDataNoOfLines)
    ReDim Preserve PVDeltaFullMin(PVDataNoOfLines)
    ReDim Preserve PVDeltaSegFullMax(PVDataNoOfLines)
    ReDim Preserve PVDeltaSegFullMin(PVDataNoOfLines)
    ReDim Preserve GraphInfoContainer(PVMaxDiameter).DataDouble(PVDataNoOfLines)
    ReDim Preserve GraphInfoContainer(PVMinDiameter).DataDouble(PVDataNoOfLines) 'PCN4333
    ReDim Preserve PVDiameterFullMin(PVDataNoOfLines)
    ReDim Preserve PVDiameterSegFullMax(PVDataNoOfLines)
    ReDim Preserve PVDiameterSegFullMin(PVDataNoOfLines)
    ReDim Preserve PVDistances(PVDataNoOfLines)
    ReDim Preserve PVTimes(PVDataNoOfLines)
    ReDim Preserve PVShapeCentreX(PVDataNoOfLines) 'PCN4484
    ReDim Preserve PVShapeCentreY(PVDataNoOfLines) 'PCN4484
    ReDim Preserve GraphInfoContainer(PVDeflectionX).DataSingle(PVDataNoOfLines) 'PCN5186
    ReDim Preserve GraphInfoContainer(PVDeflectionY).DataSingle(PVDataNoOfLines) 'PCN5186
    ReDim Preserve SmoothDeflectionX(PVDataNoOfLines)
    ReDim Preserve SmoothDeflectionY(PVDataNoOfLines)
'PCN6458     ReDim Preserve GraphInfoContainer(PVInclination).DataSingle(PVDataNoOfLines) 'PCN6128
'PCN6458     ReDim Preserve GraphInfoContainer(PVDesignGradient).DataSingle(PVDataNoOfLines) 'PCN6178
    
End If

If LoadingTimeStampError = True Then Exit Sub

StartTime = PVTimes(1)
EndTime = PVTimes(PVDataNoOfLines)
If EndTime <= 0 Then EndTime = PVTimes(PVDataNoOfLines - 1): PVTimes(PVDataNoOfLines) = EndTime  'If Last time is corupt then make secound last the last
TimeStep = (EndTime - StartTime) / PVDataNoOfLines

For FrameNo = 2 To PVDataNoOfLines - 1
    PVTimes(FrameNo) = ((FrameNo - 1) * TimeStep) + StartTime
Next FrameNo


Exit Sub
Err_Handler:
    
    Select Case Err
        Case Else: MsgBox Err & "-PF45:" & Error$
    End Select

End Sub

Sub SaveToFilePipeObs(FileSaveFail As Boolean)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : SaveToFilePipeObs
'Created : 23 July 2004, PCN2928
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Saves Pipes Obs back to the PVD file
'Usage   :
'***************************************************************
On Error GoTo Err_Handler
Dim SaveToAddress As Long 'Store the address where you wish to save the data
Dim ObsLineNo As Integer
Dim FileNo As Integer

If PipeObsBuffer = 0 Then Exit Sub

NoOfPipeObservations = UBound(PipeObservations)

'If Observations.ObservationsList.ListCount = 0 Then Exit Function 'PCN????
'If NoOfPipeObservations = 0 And Observations.ObservationsList.ListCount = 0 Then Exit Function
'If NoOfPipeObservations = 0 Then Exit Function 'PCV4131

FileSaveFail = False 'PCN1768

If Dir(PVDFileName) = "" Or PVDFileName = "" Then Exit Sub


FileNo = FreeFile ' PCN????
On Error GoTo FileErr_Handler
Open PVDFileName For Binary Access Write As #FileNo
On Error GoTo Err_Handler

'Write to file the Pipe Observations 'PCNGL130103
PVDHeaderPipeObs.PVDCheck = NoOfPipeObservations
Put #FileNo, PVDFilePointers.PVDPointerPipeObs, PVDHeaderPipeObs
Call SaveObservations_V60(NoOfPipeObservations, FileNo) 'PCN2928


Close #FileNo
FileSaveFail = False

Exit Sub
FileErr_Handler:
    FileSaveFail = True
    Close #FileNo

Exit Sub
Err_Handler:
Select Case Err
    Case 52 'Bad filename or number
        Resume Next
    Case 63 'Bad record number
        Resume Next
    Case Else
        MsgBox Err & "-PF46:" & Error$
        GoTo FileErr_Handler

End Select
End Sub



Sub LoadConfigInfo_V41IntoConfigInfo()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : LoadConfigInfo_V41IntoConfigInfo
'Created : 2 September 2004, PCN3019
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Loads the V41 ConfigInfo into the current ConfigInfo.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler


ConfigInfo.PVDFileVersion = Trim(ConfigInfo_V41.PVDFileVersion)
ConfigInfo.WLStartAngle = ConfigInfo_V41.WLStartAngle
ConfigInfo.WLFinishAngle = ConfigInfo_V41.WLFinishAngle
ConfigInfo.FishEyeHorDistortion = ConfigInfo_V41.FishEyeHorDistortion 'PCN3687 was area that was never used
ConfigInfo.VideoFileName = ConfigInfo_V41.VideoFileName
ConfigInfo.CalDist = ConfigInfo_V41.CalDist
ConfigInfo.CalLineLength = ConfigInfo_V41.CalLineLength
ConfigInfo.FileCountryCode = ConfigInfo_V41.FileCountryCode
ConfigInfo.FileLanguage = ConfigInfo_V41.FileLanguage
ConfigInfo.LenReal = ConfigInfo_V41.LenReal
ConfigInfo.LenRealPercent = ConfigInfo_V41.LenRealPercent
ConfigInfo.MediaHeight = ConfigInfo_V41.MediaHeight
ConfigInfo.MediaWidth = ConfigInfo_V41.MediaWidth
ConfigInfo.NoOfProfileSegments = ConfigInfo_V41.NoOfProfileSegments
ConfigInfo.Ratio = ConfigInfo_V41.Ratio
ConfigInfo.Units = ConfigInfo_V41.Units
ConfigInfo.FishEyeCenterX = ConfigInfo_V41.FishEyeCenterX
ConfigInfo.FishEyeCenterY = ConfigInfo_V41.FishEyeCenterY
ConfigInfo.FishEyeDistortion = ConfigInfo_V41.FishEyeDistortion
ConfigInfo.FishEyeFlag = ConfigInfo_V41.FishEyeFlag
ConfigInfo.FishEyeRatio = ConfigInfo_V41.FishEyeRatio



Exit Sub
Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & "-PF47:" & Error$
End Select
End Sub

Sub ConvertConfigInfoToConfigInfo_V52()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : ConvertConfigInfoToConfigInfo_V52
'Created : 2 September 2004, PCN3019
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    : Copies the current ConfigInfo information into V5.2 of the ConfigInfo.
'Usage   : When a V5.2 PVD file is loaded, all saving to PVD must be in the same
'           Version format.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

'vvvv PCN3331 *************************************
'Call ConvertConfigInfoToConfigInfo_V50
ConfigInfo_V52.PVDFileVersion = Trim(ConfigInfo.PVDFileVersion) 'PCN2850
ConfigInfo_V52.WLStartAngle = ConfigInfo.WLStartAngle
ConfigInfo_V52.WLFinishAngle = ConfigInfo.WLFinishAngle
ConfigInfo_V52.FishEyeHorDistortion = ConfigInfo.FishEyeHorDistortion 'PCN3687 was area that was never used
ConfigInfo_V52.VideoFileName = ConfigInfo.VideoFileName
ConfigInfo_V52.CalDist = ConfigInfo.CalDist
ConfigInfo_V52.CalLineLength = ConfigInfo.CalLineLength
ConfigInfo_V52.FileCountryCode = ConfigInfo.FileCountryCode
ConfigInfo_V52.FileLanguage = ConfigInfo.FileLanguage
ConfigInfo_V52.LenReal = ConfigInfo.LenReal
ConfigInfo_V52.LenRealPercent = ConfigInfo.LenRealPercent
ConfigInfo_V52.MediaHeight = ConfigInfo.MediaHeight
ConfigInfo_V52.MediaWidth = ConfigInfo.MediaWidth
ConfigInfo_V52.NoOfProfileSegments = ConfigInfo.NoOfProfileSegments
ConfigInfo_V52.Ratio = ConfigInfo.Ratio
ConfigInfo_V52.Units = ConfigInfo.Units
ConfigInfo_V52.FishEyeCenterX = ConfigInfo.FishEyeCenterX
ConfigInfo_V52.FishEyeCenterY = ConfigInfo.FishEyeCenterY
ConfigInfo_V52.FishEyeDistortion = ConfigInfo.FishEyeDistortion
ConfigInfo_V52.FishEyeFlag = ConfigInfo.FishEyeFlag
ConfigInfo_V52.FishEyeRatio = ConfigInfo.FishEyeRatio
'V5.0 additions
ConfigInfo_V52.DistanceProcessMethod = Trim(ConfigInfo.DistanceProcessMethod)
ConfigInfo_V52.DistanceStart = ConfigInfo.DistanceStart
ConfigInfo_V52.DistanceDirection = Trim(ConfigInfo.DistanceDirection)
ConfigInfo_V52.DistanceFinish = ConfigInfo.DistanceFinish
'^^^^ *********************************************
'V5.2 additions
ConfigInfo_V52.PVShapeCentreX = ConfigInfo.PVShapeCentreX 'PCN4336
ConfigInfo_V52.PVShapeCentreY = ConfigInfo.PVShapeCentreY 'PCN4336
ConfigInfo_V52.IPGradThres = ConfigInfo.IPGradThres
ConfigInfo_V52.IPStDX = ConfigInfo.IPStDX
ConfigInfo_V52.IPStDY = ConfigInfo.IPStDY
ConfigInfo_V52.IPProcessMethod = Trim(IPEnhancementAndIPProcessMethod.IPProcessMethod)
ConfigInfo_V52.IPZone = ConfigInfo.IPZone
ConfigInfo_V52.IPEnhancement = Trim(IPEnhancementAndIPProcessMethod.IPEnhancement)
ConfigInfo_V52.LimitCapacityL = ConfigInfo.LimitCapacityL
ConfigInfo_V52.LimitCapacityR = ConfigInfo.LimitCapacityR
ConfigInfo_V52.LimitOvality = ConfigInfo.LimitOvality
ConfigInfo_V52.LimitDeltaL = ConfigInfo.LimitDeltaL
ConfigInfo_V52.LimitDeltaR = ConfigInfo.LimitDeltaR
ConfigInfo_V52.LimitXYDiameterL = ConfigInfo.LimitXYDiameterL
ConfigInfo_V52.LimitXYDiameterR = ConfigInfo.LimitXYDiameterR



Exit Sub
Err_Handler:
    MsgBox Err & "-PF48:" & Error$
End Sub

Sub ConvertConfigInfoToConfigInfo_V60()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : ConvertConfigInfoToConfigInfo_V60
'Created : 2 September 2004, PCN3019
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    : Copies the current ConfigInfo information into V6.0 of the ConfigInfo.
'Usage   : When a V6.0 PVD file is loaded, all saving to PVD must be in the same
'           Version format.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

'vvvv PCN3331 *************************************
'Call ConvertConfigInfoToConfigInfo_V52
ConfigInfo_V60.PVDFileVersion = Trim(ConfigInfo.PVDFileVersion) 'PCN2850
ConfigInfo_V60.WLStartAngle = ConfigInfo.WLStartAngle
ConfigInfo_V60.WLFinishAngle = ConfigInfo.WLFinishAngle
ConfigInfo_V60.FishEyeHorDistortion = ConfigInfo.FishEyeHorDistortion 'PCN3687 was area that was never used
ConfigInfo_V60.VideoFileName = ConfigInfo.VideoFileName
ConfigInfo_V60.CalDist = ConfigInfo.CalDist
ConfigInfo_V60.CalLineLength = ConfigInfo.CalLineLength
ConfigInfo_V60.FileCountryCode = ConfigInfo.FileCountryCode
ConfigInfo_V60.FileLanguage = ConfigInfo.FileLanguage
ConfigInfo_V60.LenReal = ConfigInfo.LenReal
ConfigInfo_V60.LenRealPercent = ConfigInfo.LenRealPercent
ConfigInfo_V60.MediaHeight = ConfigInfo.MediaHeight
ConfigInfo_V60.MediaWidth = ConfigInfo.MediaWidth
ConfigInfo_V60.NoOfProfileSegments = ConfigInfo.NoOfProfileSegments
ConfigInfo_V60.Ratio = ConfigInfo.Ratio
ConfigInfo_V60.Units = ConfigInfo.Units
ConfigInfo_V60.FishEyeCenterX = ConfigInfo.FishEyeCenterX
ConfigInfo_V60.FishEyeCenterY = ConfigInfo.FishEyeCenterY
ConfigInfo_V60.FishEyeDistortion = ConfigInfo.FishEyeDistortion
ConfigInfo_V60.FishEyeFlag = ConfigInfo.FishEyeFlag
ConfigInfo_V60.FishEyeRatio = ConfigInfo.FishEyeRatio
'V5.0 additions
ConfigInfo_V60.DistanceProcessMethod = Trim(ConfigInfo.DistanceProcessMethod)
ConfigInfo_V60.DistanceStart = ConfigInfo.DistanceStart
ConfigInfo_V60.DistanceDirection = Trim(ConfigInfo.DistanceDirection)
ConfigInfo_V60.DistanceFinish = ConfigInfo.DistanceFinish
'V5.2 additions
ConfigInfo_V60.PVShapeCentreX = ConfigInfo.PVShapeCentreX 'PCN4336
ConfigInfo_V60.PVShapeCentreY = ConfigInfo.PVShapeCentreY 'PCN4336
ConfigInfo_V60.IPGradThres = ConfigInfo.IPGradThres
ConfigInfo_V60.IPStDX = ConfigInfo.IPStDX
ConfigInfo_V60.IPStDY = ConfigInfo.IPStDY
ConfigInfo_V60.IPProcessMethod = Trim(IPEnhancementAndIPProcessMethod.IPProcessMethod)
ConfigInfo_V60.IPZone = ConfigInfo.IPZone
ConfigInfo_V60.IPEnhancement = Trim(IPEnhancementAndIPProcessMethod.IPEnhancement)
ConfigInfo_V60.LimitCapacityL = ConfigInfo.LimitCapacityL
ConfigInfo_V60.LimitCapacityR = ConfigInfo.LimitCapacityR
ConfigInfo_V60.LimitOvality = ConfigInfo.LimitOvality
ConfigInfo_V60.LimitDeltaL = ConfigInfo.LimitDeltaL
ConfigInfo_V60.LimitDeltaR = ConfigInfo.LimitDeltaR
ConfigInfo_V60.LimitXYDiameterL = ConfigInfo.LimitXYDiameterL
ConfigInfo_V60.LimitXYDiameterR = ConfigInfo.LimitXYDiameterR
'^^^^ *********************************************
'V6.0 additions
ConfigInfo_V60.ProfileRecordingMethod = Trim(ConfigInfo.ProfileRecordingMethod)


Exit Sub
Err_Handler:
    MsgBox Err & "-PF49:" & Error$
End Sub

Public Sub RapidSavePVTimeToFile(CurrentFrameNo As Long, PVDataStartAddress As Long, PVDFileSaveFail As Boolean, FileNo As Integer)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : RapidSavePVTimeToFile
'Created : 8 Nov 2004, PCN3140
'Updated :
'Prg By  : Geoff Logan
'Param   : CurrentFrameNo -
'          PVDataAddressOffset -
'          PVDFileSaveFail -
'Desc    :
'Usage   :
'***************************************************************
On Error GoTo Err_Handler
Dim ErrorStatus As String
Dim PVAddressOffset As Long

If Trim(ConfigInfo.PVDFileVersion) <> "V4.0" And Trim(ConfigInfo.PVDFileVersion) <> "V4.0" Then
    PVAddressOffset = PVDataStartAddress + (CurrentFrameNo + 1) * PVDataFrameBlockSize
    PVAddressOffset = PVAddressOffset + (CurrentFrameNo) * (PVCalculationsBlockSize + PVRelatedInfoBlockSize)
    PVAddressOffset = PVAddressOffset + PVCalculationsBlockSize
    Put #FileNo, PVAddressOffset, PVTimes(CurrentFrameNo)
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
        MsgBox Err & "-PF50:" & Error$
End Select
End Sub


Public Sub UpdatePVTimeInPVDFile()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : RapidSavePVTimeToFile
'Created : 8 Nov 2004, PCN3140
'Updated :
'Prg By  : Geoff Logan
'Param   : CurrentFrameNo -
'          PVDataAddressOffset -
'          PVDFileSaveFail -
'Desc    :
'Usage   : Only for Debuging, it has not been tested for Public use
'***************************************************************
On Error GoTo Err_Handler
Dim ErrorStatus As String
Dim FileLoadError As Boolean
Dim PVAddressOffset As Long
Dim PVDataStartAddress As Long
Dim FileNo As Integer
Dim CurrentFrame As Long


CurrentFrame = 1 'PCN3140 Frame to manually change

PVTimes(CurrentFrame) = 0# 'PCN3140 The number to change it to

'Find PVDataStartAddress
Call GetPVDPointerPVDataFromFile(PVDFileName, PVDataStartAddress, FileLoadError)
FileNo = 8
Open PVDFileName For Binary Access Write As #FileNo

If Not FileLoadError And Trim(ConfigInfo.PVDFileVersion) <> "V4.0" And Trim(ConfigInfo.PVDFileVersion) <> "V4.1" Then
    Call RapidSavePVTimeToFile(CurrentFrame, PVDataStartAddress, FileLoadError, FileNo)
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
  '      PVDFileSaveFail = True
    Case Else
        MsgBox Err & "-PF51:" & Error$
End Select
End Sub

Public Sub EmbedFile(ByVal FileName As String, _
                     ByRef ReturnFileOffset As Long, _
                     ByRef ReturnFileLength As Long, _
                     ByRef PVDHeaderEmbedded As PVDHeaderEmbeddedType)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCN     : PCN3576
'Name    : EmbedFile
'Created : 12 July 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   : Filename - File to embed into PVD File
'          ReturnOffset - File offset that the embed file starts
'          ReturnLength - File size that is being embeded
'Desc    : This function will append a binary file byte for byte to the end of the current PVD file
'Usage   : initially used for embeding snapshots from observations
'***************************************************************
On Error GoTo Err_Handler

Dim FileNoFrom As Integer
Dim FileNoPVD As Integer
Dim DataByte As Byte
Dim FileLength As Long
Dim CurrentFileLocation As Long

Dir FileName    'Check to see if both source file and pvd file exist    '
If Dir(FileName) = "" Then Exit Sub                                     '
If Dir(PVDFileName) = "" Or PVDFileName = "" Then Exit Sub         '



FileNoFrom = FreeFile
Open FileName For Binary Access Read As FileNoFrom
FileLength = LOF(FileNoFrom)
PVDHeaderEmbedded.FileLength = LOF(FileNoFrom)
Close #FileNoFrom

FileNoPVD = FreeFile
On Error GoTo FileErr_Handler
Open PVDFileName For Binary Access Write As FileNoPVD
On Error GoTo Err_Handler

ReturnFileOffset = GetEndOfLastEmbeddedFile
If ReturnFileOffset = 0 Then ReturnFileOffset = LOF(FileNoPVD)

Seek #FileNoPVD, ReturnFileOffset
Put #FileNoPVD, ReturnFileOffset, PVDHeaderEmbedded

CurrentFileLocation = Loc(FileNoPVD)
Close #FileNoPVD

Call clearline_EmbedFileData(PVDFileName, FileName, CurrentFileLocation)
                            
                            


''Do While Not EOF(FileNoFrom)
''    Get FileNoFrom, , DataByte
''    Put FileNoPVD, , DataByte
''    FileLength = FileLength + 1
''Loop
'''
'''
'''
''Close FileNoFrom
''Close FileNoPVD

ReturnFileLength = PVDHeaderEmbedded.FileLength


Exit Sub
FileErr_Handler:
    Close #FileNoFrom
    Close #FileNoPVD
Exit Sub


Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & "-PF52:" & Error$

End Select
End Sub
Function GetEndOfLastEmbeddedFile() As Long
On Error GoTo Err_Handler

Dim NumberEmbeddedFiles As Integer
Dim PVDHeaderEmbedded As PVDHeaderEmbeddedType
Dim FilesIndex As Integer

NumberEmbeddedFiles = UBound(ListEmbeddedOwners)

If ListEmbeddedOwners(1).FileOffset = 0 Then GetEndOfLastEmbeddedFile = 0: Exit Function

For FilesIndex = NumberEmbeddedFiles To 1 Step -1
    If ListEmbeddedOwners(FilesIndex).FileOffset <> 0 Then
        GetEndOfLastEmbeddedFile = ListEmbeddedOwners(FilesIndex).FileOffset + _
                           ListEmbeddedOwners(FilesIndex).FileLength + _
                           Len(PVDHeaderEmbedded)
        Exit Function
    End If
Next FilesIndex

GetEndOfLastEmbeddedFile = 0

Exit Function
Err_Handler:
Select Case Err
    Case 9: GetEndOfLastEmbeddedFile = 0: Exit Function
    Case Else:    MsgBox Err & "-PF53:" & Error$
End Select
End Function

Public Sub EmbeddedFileExtract(ByVal FileName As String, _
                               ByVal FileOffset As Long, _
                               ByVal FileLength As Long, _
                               ByRef PVDHeaderEmbedded As PVDHeaderEmbeddedType)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCN     : PCN3576
'Name    : EmbeddedFileExtract
'Created : 12 July 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   : Filename - Filename to save the extracted embed file as
'          FileOffset - position the file bytes start in the PVD that need to be extracted
'          FileLength - Lenght of file that is to be extracted, (in bytes)
'Desc    : This function will extract an embeded file from the PVD file and save it as its
'          own seperate file which we hope we will be able to reload independantly
'Usage   : Initially used for embeding snapshots from observations
'***************************************************************
On Error GoTo Err_Handler

'Dim FileNoTo As Integer
Dim FileNoPVD As Integer

Dim DataByte As Byte
Dim Count As Long
Dim CurrentLocation As Long

If FileName = "" Then Exit Sub
If Dir(PVDFileName) = "" Or PVDFileName = "" Then Exit Sub

If Dir(FileName) <> "" Then
    Kill FileName
End If


'FileNoTo = FreeFile
'Open FileName For Binary Access Write As #FileNoTo

FileNoPVD = FreeFile
Open PVDFileName For Binary Access Read As #FileNoPVD

Seek #FileNoPVD, FileOffset
Get #FileNoPVD, , PVDHeaderEmbedded

CurrentLocation = Loc(FileNoPVD)
Close #FileNoPVD


Call clearline_ExtractEmbedFile(PVDFileName, FileName, CurrentLocation, FileLength)


'For Count = 1 To FileLength
'    If EOF(FileNoPVD) Then Exit For
'    Get #FileNoPVD, , DataByte
'    Put #FileNoTo, , DataByte
'Next Count

'Close #FileNoTo
'Close #FileNoPVD

Exit Sub
FileErr_Handler:
'    Close #FileNoTo
    Close #FileNoPVD
Exit Sub

Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & "-PF54:" & Error$
End Select
End Sub

Public Sub EmbeddedFileRemove(ByVal FileOffset As Long, ByVal FileLength As Long)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCN     : PCN3576
'Name    : EmbeddedFileRemove
'Created : 12 July 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   : FileOffset - file position of the embeded file that needs to be removed (in bytes)
'          FileLength - length of the embeded file that is to be removed (in bytes)
'Desc    : This function will remove embeded file from the PVD file and move any other embeded
'          file up to cover the whole of the removed file from the PVD
'Usage   : Initially used for embeding snapshots from observations
'***************************************************************
On Error GoTo Err_Handler
If FileOffset = 0 Or FileLength = 0 Then Exit Sub

Dim NumberEmbeddedFiles
Dim FilesIndex As Integer
Dim Swaped As Boolean
Dim CutSize As Long


If Dir(PVDFileName) = "" Or PVDFileName = "" Then Exit Sub

NumberEmbeddedFiles = UBound(ListEmbeddedOwners)

For FilesIndex = 1 To NumberEmbeddedFiles
    If ListEmbeddedOwners(FilesIndex).FileOffset = FileOffset Then Exit For
Next FilesIndex

CutSize = ListEmbeddedOwners(FilesIndex).FileLength + Len(ListEmbeddedOwners(FilesIndex).PVHeaderEmbedded)

'Call MoveFileData(ListEmbeddedOwners(FilesIndex).FileOffset + CutSize, _
                      ListEmbeddedOwners(FilesIndex).FileOffset)
Call clearline_MoveFileData(PVDFileName, ListEmbeddedOwners(FilesIndex).FileOffset + CutSize, _
                      ListEmbeddedOwners(FilesIndex).FileOffset)
                      

PipeObservations(ListEmbeddedOwners(FilesIndex).EmbeddedIndex).PipeObsSnapshotOffset = 0

For FilesIndex = FilesIndex + 1 To NumberEmbeddedFiles
    With PipeObservations(ListEmbeddedOwners(FilesIndex).EmbeddedIndex)
        If .PipeObsSnapshotOffset <> 0 Then .PipeObsSnapshotOffset = .PipeObsSnapshotOffset - CutSize
    End With
Next FilesIndex

Call ObsInitEmbeddedFiles 'Reinitialise EmbeddedFiles

    
Exit Sub
FileErr_Handler:
'    Close #FileNo
Exit Sub
Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & "-PF55:" & Error$
End Select
End Sub

Sub MoveFileData(ByVal FromFilePosition As Long, ByVal ToFilePosition As Long)
On Error GoTo Err_Handler

Dim FileNoPVDWrite As Integer
Dim FileNoPVDRead As Integer
Dim DataByte As Byte
Dim Count As Long
Dim MoveSize As Long
Dim NewEndOfFile As Long
Dim LengthOfFile As Long

Dim PVDHeaderEmbeddedWriteCheck As PVDHeaderEmbeddedType
Dim PVDHeaderEmbeddedReadCheck As PVDHeaderEmbeddedType

Dim hFile As Long, lpOfstruct As OFSTRUCT

MoveSize = FromFilePosition - ToFilePosition

FileNoPVDRead = FreeFile
Open PVDFileName For Binary Access Read As #FileNoPVDRead
LengthOfFile = LOF(FileNoPVDRead)

FileNoPVDWrite = FreeFile
Open PVDFileName For Binary Access Write As #FileNoPVDWrite

Seek #FileNoPVDRead, FromFilePosition
Seek #FileNoPVDWrite, ToFilePosition

Get #FileNoPVDRead, , PVDHeaderEmbeddedReadCheck
'Get #FileNoPVDWrite, , PVDHeaderEmbeddedWriteCheck

Seek #FileNoPVDRead, FromFilePosition
Seek #FileNoPVDWrite, ToFilePosition

Count = 0
'If FromFilePosition > LengthOfFile Then
    
    Do
    
        If EOF(FileNoPVDRead) Then Exit Do
        Get #FileNoPVDRead, , DataByte
        Put #FileNoPVDWrite, , DataByte
        Count = Count + 1
    Loop
'End If

Close #FileNoPVDRead
Close #FileNoPVDWrite

NewEndOfFile = ToFilePosition + Count
NewEndOfFile = LengthOfFile - (FromFilePosition - ToFilePosition) + 2


'Truncate file at NewEndOfFile distance
'hFile = OpenFile("PVDFileName", lpOfstruct, OF_WRITE)
'SetFilePointer hFile, NewEndOfFile, 0, FILE_BEGIN
'SetEndOfFile hFile
'CloseHandle hFile
'''''''''''''''''''''''''''''''''''''''''

Exit Sub
File_Error:
Close #FileNoPVDRead
Close #FileNoPVDWrite


Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-PF56:" & Error$

    End Select
    
End Sub

Function FindVideoFile(ByVal PVDName As String, ByVal VideoName As String) As String
On Error GoTo Err_Handler
    Dim PVDPathName As String
    Dim PVDFileName As String
    Dim PVDExtension As String
    
    Dim VideoPathName As String
    Dim VideoFileName As String
    Dim VideoExtension
    Dim TryFileName As String
    
    

    
    
    Call SplitFilePath(PVDName, PVDPathName, PVDFileName, PVDExtension)
    Call SplitFilePath(VideoName, VideoPathName, VideoFileName, VideoExtension)
    
    TryFileName = PVDPathName & VideoFileName & "." & VideoExtension
On Error GoTo BadFileName1
    If Dir(TryFileName) <> "" Then FindVideoFile = TryFileName: Exit Function
BadFileName1:
On Error GoTo Err_Handler
    
    TryFileName = PVDPathName & PVDFileName & "." & VideoExtension
On Error GoTo BadFileName2
    If Dir(TryFileName) <> "" Then FindVideoFile = TryFileName: Exit Function
BadFileName2:
On Error GoTo Err_Handler

    TryFileName = VideoPathName & PVDFileName & "." & VideoExtension
On Error GoTo BadFileName3
    If Dir(TryFileName) <> "" Then FindVideoFile = TryFileName: Exit Function
BadFileName3:
On Error GoTo Err_Handler
    FindVideoFile = VideoName
        
    
    

Exit Function
Err_Handler:
Select Case Err
    
    Case Else
        MsgBox Err & "-PF57:" & Error$
    End Select
End Function

Sub LoadFullPVDataFromFile()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : LoadFullPVDataFromFile
'Created : 9 June 2003
'Updated : 17 November 2003, PCN2401
'Prg By  : Geoff Logan
'Param   : PercentComplete - The percent complete to set the progress bar
'Desc    : Read the PVData data from file for 3D profile model.
'Usage   : Call ProgressBarPosition(0.55) is a standard example.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

    
    Dim PVDataStartAddress As Long
    Dim PVFileLoadError As Boolean
    Dim Scaler As Double
    Dim PVDataBlockSize As Long
    Dim XY As Long
    Dim PVDVer As Single
    
    PVDVer = GetPVDVer
    
   
    If PVDVer < 6.3 Then 'PCN4006

        If Trim(ConfigInfo.ProfileRecordingMethod) = "XY" Then
            Scaler = ConfigInfo.Ratio * VideoScreenScale / 10
        Else
            Scaler = 1
        End If
    Else 'PCN4006
        Scaler = ConfigInfo.Ratio * VideoScreenScale
    End If
    
    Call GetPVDPointerPVDataFromFile(PVDFileName, PVDataStartAddress, PVFileLoadError) 'PCN2164
    PVDataBlockSize = PVDataFrameBlockSize + PVCalculationsBlockSize + PVRelatedInfoBlockSize

    XY = 0
    If Trim(ConfigInfo.ProfileRecordingMethod) = "XY" Then XY = 1
    If PVDVer > 6.25 Then XY = 2 'PCN4168

    Call clearline_LoadPVD_Data(PVDFileName, _
                                PVDataStartAddress, _
                                PVDataBlockSize, _
                                XY, _
                                TD_PVDataX(1), _
                                TD_PVDataY(1), _
                                Scaler, _
                                0, _
                                PVDataNoOfLines) 'PCN3603
ExpectedDiameter = ExpectedDiameter


Exit Sub
Err_Handler:
    Select Case Err
        Case 6: Resume Next
        Case 9 'Out of subscript range
            Exit Sub
        Case Else
            MsgBox Err & "-PF58:" & Error$
    End Select
End Sub

Sub SplitFilePath(ByVal FileNameToSplit As String, ByRef PathName As String, ByRef FileName As String, ByRef FileExtension)
On Error GoTo Err_Handler
    
    Dim StringSearch As String
    Dim X As Long
    
    X = 0
    StringSearch = ""

'Split the path the the file
    Do Until (InStr(StringSearch, "\")) <> 0 Or (X > Len(FileNameToSplit))
      StringSearch = Right(FileNameToSplit, X)
      X = X + 1
    Loop
    If X <= Len(FileNameToSplit) Then
        PathName = Left(FileNameToSplit, Len(FileNameToSplit) - (X - 2))
        FileName = Right(FileNameToSplit, (X - 2))
    Else
        PathName = ""
        FileName = FileNameToSplit
    End If
        
    



'split the file from the extension
    X = 1
    StringSearch = ""
    Do Until InStr(StringSearch, ".") <> 0 Or X > Len(FileName)
      StringSearch = Right(FileNameToSplit, X)
      X = X + 1
    Loop
    If X <= Len(FileName) Then
        FileExtension = Right(FileName, (X - 2))
        FileName = Left(FileName, Len(FileName) - (X - 1))
    Else
        FileExtension = ""
    End If

Exit Sub
Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & "-PF59:" & Error$
End Select
End Sub

Sub CheckForIPD()
On Error GoTo Err_Handler
    Dim DistanceValue As Long
    Dim validIPD As Long
    
    Call hough_checkforIPD(validIPD)
    
     If validIPD = 1 Then
        IPD = True
        DistanceMethod = "AutomaticCounter"
        
        Call getcounter(DistanceValue)
        

        If MeasurementUnits = "mm" Then
            DistanceStart = Format(DistanceValue / 100, "#0.00") 'PCNAVI060804 forgot to copy / 10
        Else
            DistanceStart = Format(DistanceValue / 10, "#0.00")
        End If
    Else
        IPD = False
    End If

Exit Sub
Err_Handler:
Select Case Err
'    Case 6: DistanceStart = 0: Exit Sub
    Case Else
        MsgBox Err & "-PF60:" & Error$

End Select
End Sub

Sub SortEmbeddedList()
On Error GoTo Err_Handler

Dim NumberEmbeddedFiles As Integer
Dim FilesIndex As Integer
Dim Swaped As Boolean

NumberEmbeddedFiles = UBound(ListEmbeddedOwners)


Do
    Swaped = False
    For FilesIndex = 1 To NumberEmbeddedFiles - 1
        If (ListEmbeddedOwners(FilesIndex + 1).FileOffset <> 0 And ListEmbeddedOwners(FilesIndex).FileOffset > ListEmbeddedOwners(FilesIndex + 1).FileOffset) Or _
           (ListEmbeddedOwners(FilesIndex).FileOffset = 0 And ListEmbeddedOwners(FilesIndex + 1).FileOffset <> 0) Then
           Swaped = True
           
           ListEmbeddedOwners(0).EmbeddedType = ListEmbeddedOwners(FilesIndex).EmbeddedType
           ListEmbeddedOwners(0).EmbeddedIndex = ListEmbeddedOwners(FilesIndex).EmbeddedIndex
           ListEmbeddedOwners(0).FileLength = ListEmbeddedOwners(FilesIndex).FileLength
           ListEmbeddedOwners(0).FileOffset = ListEmbeddedOwners(FilesIndex).FileOffset
           ListEmbeddedOwners(0).PVHeaderEmbedded.Descriptor = ListEmbeddedOwners(FilesIndex).PVHeaderEmbedded.Descriptor
           ListEmbeddedOwners(0).PVHeaderEmbedded.FileLength = ListEmbeddedOwners(FilesIndex).PVHeaderEmbedded.FileLength
           ListEmbeddedOwners(0).PVHeaderEmbedded.Owner = ListEmbeddedOwners(FilesIndex).PVHeaderEmbedded.Owner
           ListEmbeddedOwners(0).PVHeaderEmbedded.PVDCheck = ListEmbeddedOwners(FilesIndex).PVHeaderEmbedded.PVDCheck
           
           ListEmbeddedOwners(FilesIndex).EmbeddedType = ListEmbeddedOwners(FilesIndex + 1).EmbeddedType
           ListEmbeddedOwners(FilesIndex).EmbeddedIndex = ListEmbeddedOwners(FilesIndex + 1).EmbeddedIndex
           ListEmbeddedOwners(FilesIndex).FileLength = ListEmbeddedOwners(FilesIndex + 1).FileLength
           ListEmbeddedOwners(FilesIndex).FileOffset = ListEmbeddedOwners(FilesIndex + 1).FileOffset
           ListEmbeddedOwners(FilesIndex).PVHeaderEmbedded.Descriptor = ListEmbeddedOwners(FilesIndex + 1).PVHeaderEmbedded.Descriptor
           ListEmbeddedOwners(FilesIndex).PVHeaderEmbedded.FileLength = ListEmbeddedOwners(FilesIndex + 1).PVHeaderEmbedded.FileLength
           ListEmbeddedOwners(FilesIndex).PVHeaderEmbedded.Owner = ListEmbeddedOwners(FilesIndex + 1).PVHeaderEmbedded.Owner
           ListEmbeddedOwners(FilesIndex).PVHeaderEmbedded.PVDCheck = ListEmbeddedOwners(FilesIndex + 1).PVHeaderEmbedded.PVDCheck
           
           ListEmbeddedOwners(FilesIndex + 1).EmbeddedType = ListEmbeddedOwners(0).EmbeddedType
           ListEmbeddedOwners(FilesIndex + 1).EmbeddedIndex = ListEmbeddedOwners(0).EmbeddedIndex
           ListEmbeddedOwners(FilesIndex + 1).FileLength = ListEmbeddedOwners(0).FileLength
           ListEmbeddedOwners(FilesIndex + 1).FileOffset = ListEmbeddedOwners(0).FileOffset
           ListEmbeddedOwners(FilesIndex + 1).PVHeaderEmbedded.Descriptor = ListEmbeddedOwners(0).PVHeaderEmbedded.Descriptor
           ListEmbeddedOwners(FilesIndex + 1).PVHeaderEmbedded.FileLength = ListEmbeddedOwners(0).PVHeaderEmbedded.FileLength
           ListEmbeddedOwners(FilesIndex + 1).PVHeaderEmbedded.Owner = ListEmbeddedOwners(0).PVHeaderEmbedded.Owner
           ListEmbeddedOwners(FilesIndex + 1).PVHeaderEmbedded.PVDCheck = ListEmbeddedOwners(0).PVHeaderEmbedded.PVDCheck
        End If
    Next FilesIndex
Loop Until Swaped = False

        

Exit Sub
Err_Handler:
    Select Case Err
        Case 9: ReDim ListEmbeddedOwners(0): Exit Sub
        Case Else: MsgBox Err & "-PF61:" & Error$
    End Select

End Sub

Sub SaveCentreCalculations()
If PVDFileName = "" Or PVRecording = True Then Exit Sub

Dim PVDVer As Single
PVDVer = GetPVDVer
If PVDVer < 6.3 Then Exit Sub 'If the PVD version is less than 6.3 then it doesn't have centre data

Dim AddressOffset As Long
Dim PVDataStartAddress As Long
Dim FileSaveError As Boolean
Dim FileNumber
Dim FrameNo As Long

FileNumber = FreeFile

Call GetPVDPointerPVDataFromFile(PVDFileName, PVDataStartAddress, FileSaveError) 'PCN2164
If FileSaveError Then Exit Sub
Open PVDFileName For Binary Access Write As #FileNumber

For FrameNo = 1 To PVDataNoOfLines
   AddressOffset = PVDataStartAddress + (FrameNo) * PVDataFrameBlockSize 'PCN2639
   AddressOffset = AddressOffset + (FrameNo - 1) * (PVCalculationsBlockSize + PVRelatedInfoBlockSize) 'PCN2639
   Put #FileNumber, AddressOffset, TD_PVCentreX(FrameNo) 'PCNGL1301032 PCN3540
   Put #FileNumber, , TD_PVCentreY(FrameNo)
Next FrameNo

Close #FileNumber
Exit Sub


Exit Sub
Err_Handler:
Select Case Err
    Case 53: Close #FileNumber: Exit Sub 'File not found (Kill statement error trap) 'PCNGL140103
    Case Else
        MsgBox Err & "-PF62:" & Error$
End Select
End Sub

Sub StoredReportStore(ByVal ReportNumber As Integer, ByVal ReportTitle As String, ByVal ReportType As Integer)
On Error GoTo Err_Handler
    Dim FileSaveFail As Boolean
    Dim NoStoredReports As Integer
    Dim NoObs As Integer
    Dim ReportFlag As String
    Dim NumberOfPages As String
    Dim PVDHeaderEmbedded As PVDHeaderEmbeddedType
    
    NoStoredReports = UBound(StoredReportArray)
'    If ReportNumber > NoStoredReports Then ReportNumber = NoStoredReports + 1 'Need to disable due to the cases where the Report number is greater than NoStoredReports
    NumberOfPages = Format(CountReportPages(ReportNumber) + 1, "##00") 'Count pages and add 1 for new page
    
    NoObs = UBound(PipeObservations)
    NoObs = NoObs + 1
    ReDim Preserve PipeObservations(NoObs)
    
    If ReportType < 0 Then ReportType = 0
    If ReportType > 9 Then ReportType = 9
    
    ReportFlag = "[[R]]" & NumberOfPages & "," & Format(ReportType, "#0") & "," & ReportTitle
    
    PVDHeaderEmbedded.Descriptor = "[EmbeddedFile]"
    PVDHeaderEmbedded.Owner = "Rep"
    PVDHeaderEmbedded.FileLength = 1
    PVDHeaderEmbedded.PVDCheck = 0
    
    
    With PipeObservations(NoObs)
        .PipeObs = ReportFlag
        .PipeObsDist = NumberOfPages 'Current page same as total pages
        .PipeObsFrameNo = ReportNumber
        'Call EmbedFile(WindowsTempDirectory & "cbs\embedfile.jpg", _ 'ID4601
        Call EmbedFile(WindowsTempDirectory & "embedfile.jpg", _
            .PipeObsSnapshotOffset, _
            .PipeObsSnapshotLength, _
            PVDHeaderEmbedded)
    End With

Call Observations.SortObs
Call SaveToFilePipeObs(FileSaveFail)

Exit Sub
Err_Handler:
Select Case Err
    Case Else: MsgBox Err & "-PF63:" & Error$
End Select
End Sub

Function CountReportPages(ByVal ReportNumber As Integer) As Integer
On Error GoTo Err_Handler
    Dim NoStoredReports As Integer
    Dim NumberOfPages As Integer
    Dim RepIndex As Integer
    
    NumberOfPages = 0
    NoStoredReports = UBound(StoredReportArray)
        
    For RepIndex = 1 To NoStoredReports
        If StoredReportArray(RepIndex).ReportNumber = ReportNumber Then NumberOfPages = NumberOfPages + 1
    Next RepIndex
    CountReportPages = NumberOfPages

Exit Function
Err_Handler:
Select Case Err
    Case Else: MsgBox Err & "-PF64:" & Error$
End Select

End Function

Sub StoredReportRetrieve(ByVal ReportNumber As Integer)
On Error GoTo Err_Handler
    Dim NoStoredReports As Integer
    Dim NumberOfPages As Integer
    Dim RepIndex As Integer
    Dim PageIndex As Integer
    Dim NumberRepImages As Integer
    Dim PVDHeaderEmbedded As PVDHeaderEmbeddedType
    
    NumberRepImages = PrecisionVisionGraph.ReportsPictureStorage.Count - 1
    
    For RepIndex = 1 To NumberRepImages
        Unload PrecisionVisionGraph.ReportsPictureStorage(RepIndex)
    Next RepIndex
    
    NoStoredReports = UBound(StoredReportArray)
    For RepIndex = 1 To NoStoredReports
        If StoredReportArray(RepIndex).ReportNumber = ReportNumber Then
            With StoredReportArray(RepIndex)
                For PageIndex = 1 To .NumberOfPages
                    Call EmbeddedFileExtract(WindowsTempDirectory & EmbeddedFileNameAndPath, _
                                             PipeObservations(.Page(PageIndex).EmbeddedIndex).PipeObsSnapshotOffset, _
                                             PipeObservations(.Page(PageIndex).EmbeddedIndex).PipeObsSnapshotLength, _
                                             PVDHeaderEmbedded)
                    Load PrecisionVisionGraph.ReportsPictureStorage(PageIndex)
                    
                    PrecisionVisionGraph.ReportsPictureStorage(PageIndex).Visible = True
                    PrecisionVisionGraph.ReportsPictureStorage(PageIndex).Left = PrecisionVisionGraph.ReportsPictureStorage(0).Left
                    PrecisionVisionGraph.ReportsPictureStorage(PageIndex).Top = PrecisionVisionGraph.ReportsPictureStorage(PageIndex - 1).Top - 500
                    
                    PrecisionVisionGraph.ReportsPictureStorage(PageIndex).Picture = LoadPicture(WindowsTempDirectory & EmbeddedFileNameAndPath)
                    PrecisionVisionGraph.ReportsTitleStorage = .Title
                Next PageIndex
            End With
        End If
    Next RepIndex
Exit Sub
Err_Handler:
Select Case Err
    Case 424: Resume Next
    Case Else: MsgBox Err & "-PF65:" & Error$
End Select



End Sub

Sub StoredReportDelete(ByVal ReportNumber As Integer)
On Error GoTo Err_Handler
    Dim NumberOfObservations As Integer
    Dim ObsIndex As Integer
    Dim DeletedReport As Boolean
    
    DeletedReport = True
    While DeletedReport = True
        DeletedReport = False
        NumberOfObservations = UBound(PipeObservations)
        For ObsIndex = NumberOfObservations To 1 Step -1
            If Left(PipeObservations(ObsIndex).PipeObs, 5) = "[[R]]" And PipeObservations(ObsIndex).PipeObsFrameNo = ReportNumber Then
                Call Observations.ObsDelete(ObsIndex)
                DeletedReport = True
                Exit For
            End If
        Next ObsIndex
    Wend
Exit Sub
Err_Handler:
Select Case Err
    Case Else: MsgBox Err & "-PF66:" & Error$
End Select
End Sub

Sub StoredReportDeleteAll()
On Error GoTo Err_Handler
    Dim NumberOfObservations As Integer
    Dim ObsIndex As Integer
    Dim DeletedReport As Boolean
    
    DeletedReport = True
    While DeletedReport = True
        DeletedReport = False
        NumberOfObservations = UBound(PipeObservations)
        For ObsIndex = NumberOfObservations To 1 Step -1
            If Left(PipeObservations(ObsIndex).PipeObs, 5) = "[[R]]" Then
                Observations.ObsDelete (ObsIndex)
                DeletedReport = True
                Exit For
            End If
        Next ObsIndex
    Wend
Exit Sub
Err_Handler:
Select Case Err
    Case Else: MsgBox Err & "-PF67:" & Error$
End Select
End Sub



Sub StoredReportInitialise()
On Error GoTo Err_Handler
    Dim NoStoredReports As Integer 'Number Of Stored Reports
    Dim NoOfPages As Integer
    Dim ObsIndex As Integer
    Dim NumberObs As Integer
    
    Dim ReportNoOfPages As Integer
    Dim ReportType As Integer
    Dim ReportTitle As String
    Dim ReportNumber As Integer
    Dim ReportPageNumber As Integer
    
    ReDim StoredReportArray(0)

    NoStoredReports = 0
    
    NumberObs = UBound(PipeObservations)
    
    For ObsIndex = 1 To NumberObs
        If Left(PipeObservations(ObsIndex).PipeObs, 5) = "[[R]]" Then
            Call ExtractReportFlagData(PipeObservations(ObsIndex).PipeObs, ReportNoOfPages, ReportType, ReportTitle)
            ReportNumber = PipeObservations(ObsIndex).PipeObsFrameNo
            ReportPageNumber = PipeObservations(ObsIndex).PipeObsDist
            
            'If ReportNumber > NoStoredReports Then 'Add new report to array
            If ReportPageNumber < 2 Then 'Add new report to array
                NoStoredReports = NoStoredReports + 1
                ReDim Preserve StoredReportArray(NoStoredReports)
                With StoredReportArray(NoStoredReports)
                     ReDim .Page(1)
                    
                    .ReportType = ReportType
                    .Title = ReportTitle
                    .NumberOfPages = ReportNoOfPages
                    .ReportNumber = ReportNumber
                    .Page(1).EmbeddedIndex = ObsIndex
                    .Page(1).PageNumber = ReportPageNumber
                End With
            Else 'Add new page to current report in array
                With StoredReportArray(NoStoredReports)
                    NoOfPages = UBound(.Page)
                    NoOfPages = NoOfPages + 1
                    ReDim Preserve .Page(NoOfPages)
                     .Page(NoOfPages).EmbeddedIndex = ObsIndex
                     .Page(NoOfPages).PageNumber = ReportPageNumber
                     .NumberOfPages = NoOfPages
                    Call SortReportPages(.Page(0))
                End With
            End If
        End If
    Next ObsIndex
Exit Sub
Err_Handler:
Select Case Err
    Case Else: MsgBox Err & "-PF68:" & Error$
End Select
End Sub

Sub ExtractReportFlagData(ByVal ObsFlag As String, _
                        ByRef ReportNoOfPages As Integer, _
                        ByRef ReportType As Integer, _
                        ByRef ReportTitle As String)
On Error GoTo Err_Handler

Dim R_NoOfPages As Integer
Dim R_Type As Integer
Dim R_Title As String

R_NoOfPages = Mid(ObsFlag, 6, 2)
R_Type = Mid(ObsFlag, 9, 1)
R_Title = Trim(Mid(ObsFlag, 11))
ReportNoOfPages = R_NoOfPages
ReportType = R_Type
ReportTitle = R_Title

Exit Sub
Err_Handler:
Select Case Err
    Case Else: MsgBox Err & "-PF69:" & Error$
End Select
End Sub

Sub SortReportPages(ByRef ReportPage As StoredReportPageIndex_V10)
On Error GoTo Err_Handler
    Dim NumberOfPages As Integer
    
    'NumberOfPages = UBound(ReportPage)
    

Exit Sub
Err_Handler:
Select Case Err
    Case Else: MsgBox Err & "-PF70:" & Error$
End Select
End Sub

Sub ProcessBarIncrement() 'PCN4241
On Error GoTo Err_Handler

ProgressBarPercent = ProgressBarPercent + 2
Call CLPProgressBar.ProgressBarPosition(ProgressBarPercent / 100)
DoEvents

Exit Sub
Err_Handler:
Select Case Err
Case 6:  Resume Next
    Case Else: MsgBox Err & "-PF71:" & Error$
End Select
End Sub

Function ThisFileIsReadOnly(FileName As String) As Boolean 'PCN4241
On Error GoTo Err_Handler
Dim FileAttributes  As Integer

FileAttributes = GetAttr(FileName)

If (FileAttributes And vbReadOnly) Then
    ThisFileIsReadOnly = True
Else
    ThisFileIsReadOnly = False
End If


Exit Function
Err_Handler:
Select Case Err
    Case 6: Resume Next 'Overflow PCNVista A temp vista fix.
    Case Else: MsgBox Err & "-PF72:" & Error$
End Select
End Function

Sub SaveFullPVDataToFile()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : LoadFullPVDataFromFile
'Created : 9 June 2003
'Updated : 17 November 2003, PCN2401
'Prg By  : Geoff Logan
'Param   : PercentComplete - The percent complete to set the progress bar
'Desc    : Read the PVData data from file for 3D profile model.
'Usage   : Call ProgressBarPosition(0.55) is a standard example.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

   
    Dim PVDataStartAddress As Long
    Dim PVFileLoadError As Boolean
    Dim Scaler As Double
    Dim PVDataBlockSize As Long
    Dim XY As Long
    Dim PVDVer As Single
    Dim WriteNumber
    Dim PVDataIndex As Long
    Dim PVDataAddressOffset As Long
    Dim Frame As Long
    Dim XData As Single
    Dim YData As Single
    Dim Segment As Integer
    
    PVDVer = GetPVDVer
    
    Scaler = ConfigInfo.Ratio * VideoScreenScale
    
    Call GetPVDPointerPVDataFromFile(PVDFileName, PVDataStartAddress, PVFileLoadError) 'PCN2164
    PVDataBlockSize = PVDataFrameBlockSize + PVCalculationsBlockSize + PVRelatedInfoBlockSize

    XY = 2 'PCN4168

    WriteNumber = FreeFile
    Open PVDFileName For Binary Access Write As #WriteNumber
    PVDataIndex = 1
    PVDataAddressOffset = PVDataStartAddress
    
    For Frame = 1 To PVDataNoOfLines
        
        Seek #WriteNumber, PVDataAddressOffset
        For Segment = 1 To 180
            XData = TD_PVDataX(PVDataIndex) / Scaler
            YData = TD_PVDataY(PVDataIndex) / Scaler
            Put #WriteNumber, , XData
            Put #WriteNumber, , YData
            PVDataIndex = PVDataIndex + 1
        Next Segment
        
        PVDataAddressOffset = PVDataAddressOffset + PVDataBlockSize
    Next Frame
    Close #WriteNumber
    Call CLPProgressBar.ProgressBarPosition(1)
    
Exit Sub
Err_Handler:
    Close #WriteNumber
    Select Case Err
        Case 9 'Out of subscript range
            Exit Sub
        Case Else
            MsgBox Err & "-PF73:" & Error$
    End Select
End Sub

'PCN4380

Sub Open2ndPVDData(ByVal FileName As String)
On Error GoTo Err_Handler

Dim ErrorType As String 'PCN3964
Dim PVDataAddressOffset As Long
Dim FileLoadError As Boolean


PVDFileName2nd = FileName

'Call LoadInPVDFormat_V6X(FileName, FileLoadError, ErrorType) 'PCN3964


'vvvv PCN2164 *****************************************************
Dim PVDataStartAddress As Long
Dim PVGraphDataAddressOffset As Long
Dim FileNumber As Integer
FileNumber = FreeFile


Call GetPVDDataAddressFromFile(FileName, PVDataStartAddress, FileLoadError)

If FileLoadError Then Exit Sub

Open FileName For Binary Access Read Lock Write As #FileNumber

Dim TDArraySize As Long

TDArraySize = NoOfProfileSegments * (PVDataNoOfLines2nd + 1)

ReDim TD_PVDataX2nd(TDArraySize)
ReDim TD_PVDataY2nd(TDArraySize)
ReDim TD_PVCentreX2nd(PVDataNoOfLines2nd)
ReDim TD_PVCentreY2nd(PVDataNoOfLines2nd)

ReDim PVDistances2nd(PVDataNoOfLines2nd)
ReDim PVTimes2nd(PVDataNoOfLines2nd)

LoadingTimeStampError = False

For PVFrameNo = 1 To PVDataNoOfLines2nd 'PCN2962
    PVGraphDataAddressOffset = PVDataStartAddress + (PVFrameNo) * PVDataFrameBlockSize 'PCN2639
    PVGraphDataAddressOffset = PVGraphDataAddressOffset + (PVFrameNo - 1) * (PVCalculationsBlockSize + PVRelatedInfoBlockSize) 'PCN2639
    Call Load2ndCentreDistanceTimeFromPVFileFile(1, PVFrameNo, PVGraphDataAddressOffset, FileLoadError) 'PCN2164
Next PVFrameNo

Close #FileNumber 'PCN2164

Call Load2ndFullPVDataFromFile

'Call FixTimeStampErrors
'Call LoadFullPVDataFromFile

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-PF74:" & Error$
    End Select
End Sub

Sub GetPVDDataAddressFromFile(ByVal FileName As String, PVGraphDataAddressOffset As Long, FileLoadError As Boolean)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : GetPVDDataAddressFromFile
'Created : 20 November 2006, PCN4380
'Updated :
'Prg By  : Antony
'Param   :  FileName - PVD file name
'           FileLoadError - returns a true value if an error occurs while loading.
'Desc    : Get the address in the file where the PVDData Starts, need no other info
'           or fills in no xtra pointer data, unlike the GetPVDPointerPVDataFromFile
'           , for this ones fill in other pointer data and needs global variables
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Dim PVDFileNumber
Dim FilePVDMainHeader As PVDFileMainHeaderType
Dim FilePVDPointers As PVDPointerType
Dim FileHeaderPVData As PVDHeaderType
Dim FileHeaderConfigInfo As PVDHeaderType

FileLoadError = False 'PCNGL140103

If Dir(FileName) = "" Or FileName = "" Then
    FileLoadError = True 'PCNGL140103
    Exit Sub
End If

'Check whether a file is open
PVDFileNumber = FreeFile
Open FileName For Binary Access Read Lock Write As #PVDFileNumber

'Load the File Main Header
Get #PVDFileNumber, , FilePVDMainHeader

'Determine file header pointers and CheckSums then read the header pointers to block 2
'File header pointers (must be the second block of data in any PVD file, there are no fix number of pointers)
If FilePVDMainHeader.PVDFileMHPointerAddress = 0 Then
    Close #PVDFileNumber
    FileLoadError = True
    Exit Sub
End If

Get #PVDFileNumber, FilePVDMainHeader.PVDFileMHPointerAddress, FilePVDPointers

'Read from file the capacity data
FileHeaderPVData.PVDHeaderDescriptor = ""
FileHeaderPVData.PVDCheck = 0
Get #PVDFileNumber, FilePVDPointers.PVDPointerPVData, FileHeaderPVData

If Left(FileHeaderPVData.PVDHeaderDescriptor, 8) <> "[PVData]" Or FileHeaderPVData.PVDCheck = 0 Then
    Close #PVDFileNumber
    FileLoadError = True
    Exit Sub
End If

PVGraphDataAddressOffset = Seek(PVDFileNumber) ' The start address of the PVData file data block

Get #PVDFileNumber, FilePVDPointers.PVDPointerConfigInfo, FileHeaderConfigInfo
Get #PVDFileNumber, , ConfigInfo2nd

Close #PVDFileNumber

PVDataNoOfLines2nd = FileHeaderPVData.PVDCheck


Exit Sub
FileErrorCleanup: 'PCNGL140103
    Close #PVDFileNumber 'PCN2980
    FileLoadError = True 'PCN2980
    
Exit Sub
Err_Handler:
    FileLoadError = True
    Select Case Err
        Case 9 'Out of subscript range
            Exit Sub
        Case 55 'File is already open 'PCNGL140103
            GoTo FileErrorCleanup
        Case Else
            MsgBox Err & "-PF75:" & Error$
    End Select
End Sub

Sub Load2ndFullPVDataFromFile()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : Load2ndFullPVDataFromFile
'Created : 21 November 2006
'Prg By  : Antony
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    
    Dim PVDataStartAddress As Long
    Dim PVFileLoadError As Boolean
    Dim Scaler As Double
    Dim PVDataBlockSize As Long
    Dim XY As Long

    Scaler = ConfigInfo2nd.Ratio * VideoScreenScale
    
    Call GetPVDDataAddressFromFile(PVDFileName2nd, PVDataStartAddress, PVFileLoadError) 'PCN2164
    
    PVDataBlockSize = PVDataFrameBlockSize + PVCalculationsBlockSize + PVRelatedInfoBlockSize

    XY = 2

    Call clearline_LoadPVD_Data(PVDFileName2nd, _
                                PVDataStartAddress, _
                                PVDataBlockSize, _
                                XY, _
                                TD_PVDataX2nd(1), _
                                TD_PVDataY2nd(1), _
                                Scaler, _
                                0, _
                                PVDataNoOfLines2nd) 'PCN3603

Exit Sub
Err_Handler:
    Select Case Err
        Case 6: Resume Next
        Case 9 'Out of subscript range
            Exit Sub
        Case Else
            MsgBox Err & "-PF76:" & Error$
    End Select
End Sub

Sub Load2ndCentreDistanceTimeFromPVFileFile(OpenFileNo As Integer, FrameNo As Long, PVGraphDataAddressOffset As Long, FileLoadError As Boolean)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : RapidReadPVGraphsDataFromFile
'Created : 7 August 2003, PCN2164
'Updated : 24 March 2004, PCN2741 -  Added OpenFileNo
'Prg By  : Geoff Logan
'Param   :  FrameNo
'           PVGraphDataAddressOffset
'           FrameBufferNo
'           FileLoadError - returns a true value if an error occurs while loading.
'           OpenFileNo -
'Desc    : Gets, as fast as possible, the PVGraphs Data from the PVD file .
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim AddressOffset As Long 'PCN2639


Dim PVDVer As Single
Dim TimeIndex As Long

TimeIndex = FrameNo - 1

AddressOffset = PVGraphDataAddressOffset

Get #OpenFileNo, AddressOffset, TD_PVCentreX2nd(FrameNo) 'PCNGL1301032 PCN3540
Get #OpenFileNo, , TD_PVCentreY2nd(FrameNo)

If TD_PVCentreX2nd(FrameNo) > 10000 Or TD_PVCentreX2nd(FrameNo) < -10000 Or _
       TD_PVCentreY2nd(FrameNo) > 10000 Or TD_PVCentreY2nd(FrameNo) < -10000 Then
        TD_PVCentreX2nd(FrameNo) = 0
        TD_PVCentreY2nd(FrameNo) = 0
End If
    
AddressOffset = AddressOffset + PVCalculationsBlockSize  'PCN2639
Get #OpenFileNo, AddressOffset, PVTimes2nd(TimeIndex) 'PCNls
If FrameNo > 2 Then
    If PVTimes2nd(TimeIndex) - PVTimes2nd(TimeIndex - 1) > 1 Or _
       PVTimes2nd(TimeIndex) - PVTimes2nd(TimeIndex - 1) < 0 Then
       LoadingTimeStampError = True
    End If
End If


If TimeIndex = 2 Then
    If PVTimes2nd(1) - PVTimes2nd(0) > 1 Or _
       PVTimes2nd(1) - PVTimes2nd(0) < 0 Then
        PVTimes2nd(0) = PVTimes2nd(1) - (PVTimes2nd(2) - PVTimes2nd(1))
    End If
End If

Get #OpenFileNo, , PVDistances2nd(FrameNo) 'PCN2639

If PVDistances2nd(FrameNo) > 10000 Or PVDistances2nd(FrameNo) < -10000 Then
    If FrameNo <> 0 Then PVDistances2nd(FrameNo) = PVDistances2nd(FrameNo - 1)
    If FrameNo = 0 Then PVDistances2nd(FrameNo) = 0
End If
    

Exit Sub
Err_Handler:
    Select Case Err
        Case 9 'Out of subscript range
            Exit Sub
        Case Else
            MsgBox Err & "-PF77:" & Error$
    End Select
    FileLoadError = True
End Sub

Sub SaveDistanceCalculations()
On Error GoTo Err_Handler

If PVDFileName = "" Or PVRecording = True Then Exit Sub


Dim AddressOffset As Long
Dim PVGraphDataStartAddress As Long
Dim PVGraphDataAddressOffset As Long
Dim FileSaveError As Boolean
Dim FileNumber
Dim FrameNo As Long

FileNumber = FreeFile

Call GetPVDDataAddressFromFile(PVDFileName, PVGraphDataStartAddress, FileSaveError) 'PCN2164
If FileSaveError Then Exit Sub
Open PVDFileName For Binary Access Write As #FileNumber

For FrameNo = 1 To PVDataNoOfLines
    PVGraphDataAddressOffset = PVGraphDataStartAddress + (FrameNo) * PVDataFrameBlockSize 'PCN2639
    PVGraphDataAddressOffset = PVGraphDataAddressOffset + PVCalculationsBlockSize + 8
    
   
   Put #FileNumber, PVGraphDataAddressOffset, PVDistances(FrameNo) 'PCNGL1301032 PCN3540
   
Next FrameNo

Close #FileNumber
Exit Sub


Exit Sub
Err_Handler:
Select Case Err
    Case 53: Close #FileNumber: Exit Sub 'File not found (Kill statement error trap) 'PCNGL140103
    Case Else
        MsgBox Err & "-PF78:" & Error$
End Select
End Sub

'ID5395
''Sub CopyFromCurruptionToGoodConfigInfo()
''On Error GoTo Err_Handler
''
''    ConfigInfo.Units = Trim(ConfigInfo_currupting.Units)
''    ConfigInfo.FileCountryCode = ConfigInfo_currupting.FileCountryCode
''    ConfigInfo.FileLanguage = ConfigInfo_currupting.FileLanguage
''    ConfigInfo.CalDist = ConfigInfo_currupting.CalDist
''    ConfigInfo.CalLineLength = ConfigInfo_currupting.CalLineLength
''    ConfigInfo.Ratio = ConfigInfo_currupting.Ratio
''    ConfigInfo.NoOfProfileSegments = ConfigInfo_currupting.NoOfProfileSegments
''    ConfigInfo.LenReal = ConfigInfo_currupting.LenReal
''    ConfigInfo.LenRealPercent = ConfigInfo_currupting.LenRealPercent
''    ConfigInfo.WLStartAngle = ConfigInfo_currupting.WLStartAngle
''    ConfigInfo.WLFinishAngle = ConfigInfo_currupting.WLFinishAngle
''    ConfigInfo.FishEyeHorDistortion = ConfigInfo_currupting.FishEyeHorDistortion
''    ConfigInfo.VideoFileName = Trim(ConfigInfo_currupting.VideoFileName)
''    ConfigInfo.MediaWidth = ConfigInfo_currupting.MediaWidth
''    ConfigInfo.MediaHeight = ConfigInfo_currupting.MediaHeight
''    ConfigInfo.PVDFileVersion = Trim(ConfigInfo_currupting.PVDFileVersion)
''    ConfigInfo.FishEyeFlag = ConfigInfo_currupting.FishEyeFlag
''    ConfigInfo.FishEyeDistortion = ConfigInfo_currupting.FishEyeDistortion
''    ConfigInfo.FishEyeRatio = ConfigInfo_currupting.FishEyeRatio
''    ConfigInfo.FishEyeCenterX = ConfigInfo_currupting.FishEyeCenterX
''    ConfigInfo.FishEyeCenterY = ConfigInfo_currupting.FishEyeCenterY
''    ConfigInfo.DistanceProcessMethod = Trim(ConfigInfo_currupting.DistanceProcessMethod)
''    ConfigInfo.DistanceStart = ConfigInfo_currupting.DistanceStart
''    ConfigInfo.DistanceDirection = Trim(ConfigInfo_currupting.DistanceDirection)
''    ConfigInfo.DistanceFinish = ConfigInfo_currupting.DistanceFinish
''    ConfigInfo.PVShapeCentreX = ConfigInfo_currupting.PVShapeCentreX
''    ConfigInfo.PVShapeCentreY = ConfigInfo_currupting.PVShapeCentreY
''    ConfigInfo.IPGradThres = ConfigInfo_currupting.IPGradThres
''    ConfigInfo.IPStDX = ConfigInfo_currupting.IPStDX
''    ConfigInfo.IPStDY = ConfigInfo_currupting.IPStDY
''    ConfigInfo.IPProcessMethod = Trim(ConfigInfo_currupting.IPProcessMethod)
''    ConfigInfo.IPZone = ConfigInfo_currupting.IPZone
''    ConfigInfo.IPEnhancement = Trim(ConfigInfo_currupting.IPEnhancement)
''    ConfigInfo.LimitCapacityL = ConfigInfo_currupting.LimitCapacityL
''    ConfigInfo.LimitCapacityR = ConfigInfo_currupting.LimitCapacityR
''    ConfigInfo.LimitOvality = ConfigInfo_currupting.LimitOvality
''    ConfigInfo.LimitDeltaL = ConfigInfo_currupting.LimitDeltaL
''    ConfigInfo.LimitDeltaR = ConfigInfo_currupting.LimitDeltaR
''    ConfigInfo.LimitXYDiameterL = ConfigInfo_currupting.LimitXYDiameterL
''    ConfigInfo.LimitXYDiameterR = ConfigInfo_currupting.LimitXYDiameterR
''    ConfigInfo.ProfileRecordingMethod = Trim(ConfigInfo_currupting.ProfileRecordingMethod)
''    ConfigInfo.FishEyeOriginalWidth = ConfigInfo_currupting.FishEyeOriginalWidth
''    ConfigInfo.FishEyeOriginalHeight = ConfigInfo_currupting.FishEyeOriginalHeight
''
''
''Exit Sub
''Err_Handler:
''Select Case Err
''
''    Case Else
''        MsgBox Err & "-PF79:" & Error$
''End Select
''End Sub

'ID5395
''Sub CopyFromGoodConfigInfoToCurruption()
''On Error GoTo Err_Handler
''
''ConfigInfo_currupting.Units = ConfigInfo.Units
''ConfigInfo_currupting.FileCountryCode = ConfigInfo.FileCountryCode
''ConfigInfo_currupting.FileLanguage = ConfigInfo.FileLanguage
''ConfigInfo_currupting.CalDist = ConfigInfo.CalDist
''ConfigInfo_currupting.CalLineLength = ConfigInfo.CalLineLength
''ConfigInfo_currupting.Ratio = ConfigInfo.Ratio
''ConfigInfo_currupting.NoOfProfileSegments = ConfigInfo.NoOfProfileSegments
''ConfigInfo_currupting.LenReal = ConfigInfo.LenReal
''ConfigInfo_currupting.LenRealPercent = ConfigInfo.LenRealPercent
''ConfigInfo_currupting.WLStartAngle = ConfigInfo.WLStartAngle
''ConfigInfo_currupting.WLFinishAngle = ConfigInfo.WLFinishAngle
''ConfigInfo_currupting.FishEyeHorDistortion = ConfigInfo.FishEyeHorDistortion
''ConfigInfo_currupting.VideoFileName = ConfigInfo.VideoFileName
''ConfigInfo_currupting.MediaWidth = ConfigInfo.MediaWidth
''ConfigInfo_currupting.MediaHeight = ConfigInfo.MediaHeight
''ConfigInfo_currupting.PVDFileVersion = Trim(ConfigInfo.PVDFileVersion)
''ConfigInfo_currupting.FishEyeFlag = ConfigInfo.FishEyeFlag
''ConfigInfo_currupting.FishEyeDistortion = ConfigInfo.FishEyeDistortion
''ConfigInfo_currupting.FishEyeRatio = ConfigInfo.FishEyeRatio
''ConfigInfo_currupting.FishEyeCenterX = ConfigInfo.FishEyeCenterX
''ConfigInfo_currupting.FishEyeCenterY = ConfigInfo.FishEyeCenterY
''ConfigInfo_currupting.DistanceProcessMethod = Trim(ConfigInfo.DistanceProcessMethod)
''ConfigInfo_currupting.DistanceStart = ConfigInfo.DistanceStart
''ConfigInfo_currupting.DistanceDirection = Trim(ConfigInfo.DistanceDirection)
''ConfigInfo_currupting.DistanceFinish = ConfigInfo.DistanceFinish
''ConfigInfo_currupting.PVShapeCentreX = ConfigInfo.PVShapeCentreX
''ConfigInfo_currupting.PVShapeCentreY = ConfigInfo.PVShapeCentreY
''ConfigInfo_currupting.IPGradThres = ConfigInfo.IPGradThres
''ConfigInfo_currupting.IPStDX = ConfigInfo.IPStDX
''ConfigInfo_currupting.IPStDY = ConfigInfo.IPStDY
''ConfigInfo_currupting.IPProcessMethod = Trim(ConfigInfo.IPProcessMethod)
''ConfigInfo_currupting.IPZone = ConfigInfo.IPZone
''ConfigInfo_currupting.IPEnhancement = Trim(ConfigInfo.IPEnhancement)
''ConfigInfo_currupting.LimitCapacityL = ConfigInfo.LimitCapacityL
''ConfigInfo_currupting.LimitCapacityR = ConfigInfo.LimitCapacityR
''ConfigInfo_currupting.LimitOvality = ConfigInfo.LimitOvality
''ConfigInfo_currupting.LimitDeltaL = ConfigInfo.LimitDeltaL
''ConfigInfo_currupting.LimitDeltaR = ConfigInfo.LimitDeltaR
''ConfigInfo_currupting.LimitXYDiameterL = ConfigInfo.LimitXYDiameterL
''ConfigInfo_currupting.LimitXYDiameterR = ConfigInfo.LimitXYDiameterR
''ConfigInfo_currupting.ProfileRecordingMethod = Trim(ConfigInfo.ProfileRecordingMethod)
''ConfigInfo_currupting.FishEyeOriginalWidth = ConfigInfo.FishEyeOriginalWidth
''ConfigInfo_currupting.FishEyeOriginalHeight = ConfigInfo.FishEyeOriginalHeight
''
''
''
''Exit Sub
''Err_Handler:
''Select Case Err
''
''    Case Else
''        MsgBox Err & "-PF79:" & Error$
''End Select
''End Sub
