Attribute VB_Name = "Observations"
'vvvv PCN3490 *********************************************************
Public ObsIndexArray() As Integer 'Used to store the Observation index at or near the current mouse position.
Const ObsMaxNumberInBubble As Integer = 5 'Maximum number of Observations for the ObsBubble.
Public AddNewObsFlag As Boolean
Public ObsOpenIndex As Integer 'Current obs that is open for viewing and editing
Public LastYObsPositionClicked As Single
Public LastObservationType As String ' "Observation" and "Distance"
'^^^^ *****************************************************************

'PCN4484 ''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Const GRADUAL As String = "Gradual"
Public Const EDGE As String = "Edge"

Public Type ShiftData                                       '
'                                                           '
    ShiftType As String 'This is either gradual or edge     '
    a As Single 'X offset, or A angle of water              '
    b As Single 'Y offset, or B angle of water              '
    Frame As Long ' Frame number                            '
    Distance As Double ' Distance
End Type                                                    '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public ReferenceShapeShiftObs() As ShiftData
Public WaterLevelShiftObs() As ShiftData

Function ObsDisplay(PosY As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PVObsDisplay
'Created : 20 Feb 2005, PCN3490
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Draws the markers for the Observation on the PipeDisplay.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim CurrentAutoRedraw As Boolean

With PrecisionVisionGraph
    CurrentAutoRedraw = .PipeDisplay.AutoRedraw
    .PipeDisplay.AutoRedraw = True
    .PipeDisplay.Line (400, PosY)-(600, PosY), RGB(0, 0, 145)
    .PipeDisplay.Circle (600, PosY), 35, RGB(100, 100, 255)
    .PipeDisplay.Circle (600, PosY), 25, RGB(150, 150, 255)
    .PipeDisplay.Circle (600, PosY), 15, RGB(150, 150, 255)
    .PipeDisplay.Circle (600, PosY), 40, RGB(0, 0, 145)
    .PipeDisplay.Circle (600, PosY), 8, RGB(0, 0, 145)
    .PipeDisplay.AutoRedraw = CurrentAutoRedraw
End With

Exit Function
Err_Handler:
    MsgBox Err & "-OBS1:" & Error$
End Function

Sub ClearObsMarkings()
On Error GoTo Err_Handler
  
 Dim CurrentAutoRedraw As Boolean

With PrecisionVisionGraph
    CurrentAutoRedraw = .PipeDisplay.AutoRedraw
    .PipeDisplay.AutoRedraw = True
    .PipeDisplay.Line (400, 0)-(700, 8115), PrecisionVisionGraph.PipeDisplay.BackColor, BF
    .PipeDisplay.AutoRedraw = CurrentAutoRedraw
End With

Call ScreenDrawing.ClearAllGraphsAndRuler
Call ScreenDrawing.DrawPVGraphs 'PCN3373


Exit Sub
Err_Handler:
MsgBox Err & "-OBS2:" & Error$
End Sub

Function ObsDisplayALL()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PVObsDisplayALL
'Created : 20 Feb 2005, PCN3490
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Draws the markers for the ALL Observations on the PipeDisplay.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim ObsIndex As Integer
Dim RulerIndex As Integer
Dim ObsPosOnPipeDisplay As Single
Dim RulerPosition As Single

If PVDataNoOfLines < 1 Then Exit Function

For ObsIndex = 1 To NoOfPipeObservations
    If Trim(PipeObservations(ObsIndex).PipeObs) = "<<<--I-->>>" Or _
       Left(PipeObservations(ObsIndex).PipeObs, 5) = "[[R]]" Or _
       Left(PipeObservations(ObsIndex).PipeObs, 6) = "[[SH]]" Or _
       Left(PipeObservations(ObsIndex).PipeObs, 6) = "[[WT]]" Then Exit For
    ObsPosOnPipeDisplay = ObsCalcPos(ObsIndex)
    If ObsPosOnPipeDisplay >= 0 Then 'Tests calculation was valid.
        Call ObsDisplay(ObsPosOnPipeDisplay)
    End If
Next ObsIndex

For RulerIndex = 1 To PrecisionVisionGraph.RulerMark.Count - 1
    Unload PrecisionVisionGraph.RulerMark(RulerIndex)
Next RulerIndex

For RulerIndex = ObsIndex To NoOfPipeObservations
    If Left(PipeObservations(RulerIndex).PipeObs, 5) = "[[R]]" Or _
       Left(PipeObservations(ObsIndex).PipeObs, 6) = "[[SH]]" Or _
       Left(PipeObservations(ObsIndex).PipeObs, 6) = "[[WT]]" Then Exit For
    Load PrecisionVisionGraph.RulerMark(RulerIndex - ObsIndex + 1)
    Set PrecisionVisionGraph.RulerMark(RulerIndex - ObsIndex + 1).Picture = PrecisionVisionGraph.RulerMark(0)
    PrecisionVisionGraph.RulerMark(RulerIndex - ObsIndex + 1).Visible = True
    PrecisionVisionGraph.RulerMark(RulerIndex - ObsIndex + 1).AutoSize = False
    PrecisionVisionGraph.RulerMark(RulerIndex - ObsIndex + 1).ZOrder 0
    PrecisionVisionGraph.RulerMark(RulerIndex - ObsIndex + 1).width = PrecisionVisionGraph.RulerMark(0).width
    PrecisionVisionGraph.RulerMark(RulerIndex - ObsIndex + 1).height = PrecisionVisionGraph.RulerMark(0).height
Next RulerIndex

For RulerIndex = ObsIndex To NoOfPipeObservations
    If Left(PipeObservations(RulerIndex).PipeObs, 5) = "[[R]]" Or _
       Left(PipeObservations(ObsIndex).PipeObs, 6) = "[[SH]]" Or _
       Left(PipeObservations(ObsIndex).PipeObs, 6) = "[[WT]]" Then Exit For
    RulerPosition = ObsCalcPos(RulerIndex)
    If RulerPosition >= 0 Then
        PrecisionVisionGraph.RulerMark(RulerIndex - ObsIndex + 1).Top = RulerPosition + PrecisionVisionGraph.PipeDisplay.Top
    End If
Next RulerIndex
Call PrecisionVisionGraph.AdjustRulerWidths

    
Exit Function
Err_Handler:
    MsgBox Err & "-OBS3:" & Error$

    
End Function

Sub ObsBubbleWithEyeSetup()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PVObsBubbleWithEyeSetup
'Created : 18 March 2005, PCN3490
'Updated :
'Prg By  : Geoff Logan
'Param   : ObsIndex - Observation Index
'Desc    : Sets up the Observation bubble at the correct possition.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler


Call ObsBubbleDisplay(1, UBound(ObsIndexArray))

Exit Sub
Err_Handler:
    MsgBox Err & "-OBS4:" & Error$
    
End Sub

Sub ObsBubbleDisplay(ByVal ObsStartIndex As Integer, ByVal ObsEndIndex As Integer)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PVObsBubbleWithEyeSetup
'Created : 18 March 2005, PCN3490
'Updated :
'Prg By  : Geoff Logan
'Param   : ObsIndex - Observation Index
'Desc    : Sets up the Observation bubble at the correct possition.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim ObsPosOnPipeDisplay As Single
Dim ObsInLines As Integer 'Used to calculate the number of line the obs are in the ObservationText
Dim ObsInLinesTotal As Integer
Dim ObsBubbleShiftOffset As Single
Dim arrayIndex As Integer
Dim ObsStr As String
Dim Units As String
Dim PipeObs As String
Dim NumberSelectedObs As Integer
Dim LowerIndexForLoad As Integer



If PVDataNoOfLines < 1 Then Exit Sub

ObsPosOnPipeDisplay = ObsCalcPos(ObsIndexArray(ObsStartIndex))
If ObsPosOnPipeDisplay < 0 Then Exit Sub      'Tests calculation was valid.
If ObsStartIndex < 1 Then ObsStartIndex = 1
If ObsEndIndex > UBound(ObsIndexArray) Then ObsEndIndex = UBound(ObsIndexArray)



NumberSelectedObs = ObsEndIndex - ObsStartIndex + 1
With PrecisionVisionGraph
    If SoftwareConfiguration = "Reader" Then .ProfileReportObsButtonImage.Visible = False 'PCN4379
    LowerIndexForLoad = .ObservationText.Count + 1
 
    For arrayIndex = LowerIndexForLoad To NumberSelectedObs
        Load .ObservationText(arrayIndex): .ObservationText(arrayIndex).ZOrder 0
        Load .ObservationEye(arrayIndex): .ObservationEye(arrayIndex).ZOrder 0
        Load .PVObsIndexNo(arrayIndex): .PVObsIndexNo(arrayIndex).ZOrder
    Next arrayIndex


    
    .PVObsEye.Top = ObsPosOnPipeDisplay - (.PVObsEye.height / 2)
    .PVObsEye.Visible = True
    
    For arrayIndex = ObsStartIndex To ObsEndIndex
        ObsIndex = ObsIndexArray(arrayIndex)
        .PVObsIndexNo(arrayIndex).text = ObsIndex
        
        Units = IIf(MeasurementUnits = "mm", "m", "ft")
        PipeObs = IIf(NumberSelectedObs = 1 And LastObservationType <> "Distance", _
                      Chr(13) & Chr(10) & Trim(PipeObservations(ObsIndex).PipeObs), _
                      "")
        .ObservationText(arrayIndex).Caption = _
            Format(PVDistances(PipeObservations(ObsIndex).PipeObsFrameNo), "#0.0") & _
            Units & _
            PipeObs
            
        .ObservationText(arrayIndex).Top = (arrayIndex * 290) - 200
        .ObservationEye(arrayIndex).Top = .ObservationText(arrayIndex).Top
        .ObservationText(arrayIndex).height = 2000

        .ObservationText(arrayIndex).Visible = True
        .ObservationEye(arrayIndex).Visible = True
    Next arrayIndex
    'Make invisible the unused controls

    For ObsIndex = arrayIndex To .ObservationText.Count
        .ObservationText(ObsIndex).Visible = False
        .ObservationEye(ObsIndex).Visible = False
    Next ObsIndex
    
    If LastObservationType = "Distance" Then
        .ObservationBubble.height = (NumberSelectedObs * 290) + 145
    ElseIf NumberSelectedObs > 1 Then
        .ObservationBubble.height = (NumberSelectedObs * 290) + 145
     Else
        If SoftwareConfiguration <> "Reader" Then
            .DeleteObs.Visible = True 'PCN4322
            .ProfileReportObsButtonImage.Visible = True 'PCN4379
        End If
        ObsInLinesTotal = GetLineCount(.ObservationText(1))
        .ObservationBubble.height = ObsInLinesTotal * 290 + 145
    End If
    
    .ObservationClose.Picture = LoadResPicture(109, vbResBitmap)
    .Observation.height = .ObservationBubble.height
    
    ObsBubbleShiftOffset = ObsPosOnPipeDisplay + .Observation.height - 1000 - .PipeMain.height
    
    If LastObservationType = "Distance" Then
        ObsBubbleShiftOffset = -170
        .SpeakObs.Visible = False
        .SpeakObsBlank.Visible = False
        .PVObsEye.Visible = False
        .SpeakObs.Visible = False
        .Observation.ZOrder 0
    ElseIf ObsBubbleShiftOffset < 0 Then
        ObsBubbleShiftOffset = 0
    End If
    
    .SpeakObsBlank.Top = 250 + ObsBubbleShiftOffset
    .SpeakObs.Top = ObsPosOnPipeDisplay - 160
    
    If LastObservationType <> "Distance" Then
        .SpeakObs.Visible = True
        .SpeakObs.ZOrder 0
        .SpeakObsBlank.Visible = True
        .PVObsEye.Visible = True
        .SpeakObs.Visible = True
    End If

    .Observation.Top = .SpeakObs.Top + 240 - ObsBubbleShiftOffset
    .Observation.Visible = True
    
    'Make invisible the Observation Entry controls.
    .PVObsDistance.Visible = False
    .PVObsDistanceUnit.Visible = False
    .PVObsEntry.Visible = False
End With
    
If NumberSelectedObs = 1 And LastObservationType <> "Distance" Then
    ObsOpenIndex = ObsIndexArray(ObsStartIndex)
    Call ObsPreviewDisplay
Else
    PrecisionVisionGraph.ObsPreviewPictureBox.Visible = False
    PrecisionVisionGraph.ObsPreviewPictureBox.ZOrder 0
End If

If NumberSelectedObs = 1 Then
    Call PrecisionVisionGraph.MoveGraph(PipeObservations(ObsIndexArray(1)).PipeObsFrameNo)
    Call ClearLineScreen.GotoPVProfile(PipeObservations(ObsIndexArray(1)).PipeObsFrameNo, True)
    If SoftwareConfiguration <> "Reader" Then
        PrecisionVisionGraph.DeleteObs.Visible = True 'PCN4322
        PrecisionVisionGraph.ProfileReportObsButtonImage.Visible = True 'PCN4379
    End If
Else
    PrecisionVisionGraph.DeleteObs.Visible = False
    PrecisionVisionGraph.ProfileReportObsButtonImage.Visible = False
End If
        

Exit Sub
Err_Handler:
    MsgBox Err & "-OBS5:" & Error$

End Sub

Sub ObsPreviewDisplay()
On Error GoTo Err_Handler

Dim SelectedObs As Integer
Dim FrameNo As Long

With PrecisionVisionGraph
    FrameNo = PipeObservations(ObsOpenIndex).PipeObsFrameNo
   
    Select Case .ObsPictureStorage(ObsOpenIndex).Tag
        Case PV: Call ScreenDrawing.DrawProfilesStartToFinish(.ObsPreview, True, FrameNo)
        Case SnapShot
            Call .ObsPreview.PaintPicture(.ObsPictureStorage(ObsOpenIndex).Picture, _
                0, 0, .ObsPreview.width, .ObsPreview.height) ', _
                '0, 0, .ObsPictureStorage(ObsOpenIndex).width, .ObsPictureStorage(ObsOpenIndex).height)
        Case Else: Call ScreenDrawing.DrawProfilesStartToFinish(.ObsPreview, True, FrameNo)
    End Select
    Call ClearLineScreen.GotoPVProfile(FrameNo, True)
    
    .ObsPreviewPictureBox.Top = .Observation.Top + .Observation.height
    .ObsPreviewPictureBox.Left = .Observation.Left
    If .ObsPreviewPictureBox.Top + .ObsPreviewPictureBox.height > .PVGraphWindow.height + .PVGraphWindow.Top Then
        .ObsPreviewPictureBox.Top = .Observation.Top - .ObsPreviewPictureBox.height
    End If
    If ShiftOn Then
        Call CheckForShapeShift(FrameNo) 'PCN4485
        Call CheckForWaterShift(FrameNo) 'PCN4486
    End If
    .ObsPreviewPictureBox.Visible = True
 
End With


Exit Sub
Err_Handler:
    MsgBox Err & "-OBS6:" & Error$
    
End Sub

Sub CheckForShapeShift(ByVal FrameNo As Long)
On Error GoTo Err_Handler
    Dim NumberSH  As Integer
    Dim Index As Integer
    Dim ShapeShiftFound As Boolean
    
    ShapeShiftFound = False
        
    NumberSH = UBound(ReferenceShapeShiftObs)
    For Index = 1 To NumberSH
        If ReferenceShapeShiftObs(Index).Frame = FrameNo Then
            ShapeShiftFound = True
            PrecisionVisionGraph.AdjustRefShapeCentre.value = 1
            If ReferenceShapeShiftObs(Index).ShiftType = EDGE Then
                PrecisionVisionGraph.ReferenceShiftType(1).value = True
            Else
                PrecisionVisionGraph.ReferenceShiftType(0).value = True
            End If
            
        End If
    Next Index
    
NoShapeShiftFound:
    
    If Not ShapeShiftFound Then
        PrecisionVisionGraph.AdjustRefShapeCentre.Tag = "Egnore"
        PrecisionVisionGraph.AdjustRefShapeCentre.value = 0
        PrecisionVisionGraph.ReferenceShiftType(0).Visible = False
        PrecisionVisionGraph.ReferenceShiftType(1).Visible = False
    Else
        PrecisionVisionGraph.ReferenceShiftType(0).Visible = True
        PrecisionVisionGraph.ReferenceShiftType(1).Visible = True
    End If
        
    
Exit Sub
Err_Handler:
    Select Case Err
        Case 9: GoTo NoShapeShiftFound ' There is no ReferenceShapeShiftObs
        Case Else: MsgBox Err & "-OBS7:" & Error$
    End Select
End Sub
Sub CheckForWaterShift(ByVal FrameNo As Long)
On Error GoTo Err_Handler
    Dim NumberWT As Integer
    Dim Index As Integer
    Dim WaterShiftFound As Boolean
    
    WaterShiftFound = False
    
    NumberWH = UBound(WaterLevelShiftObs)
    For Index = 1 To NumberWH
        If WaterLevelShiftObs(Index).Frame = FrameNo Then
            WaterShiftFound = True
            PrecisionVisionGraph.AdjustWaterCentre.value = 1
            If WaterLevelShiftObs(Index).ShiftType = EDGE Then
                PrecisionVisionGraph.WaterShiftType(1).value = True
            Else
                PrecisionVisionGraph.WaterShiftType(0).value = True
            End If
        End If
    Next Index

NoWaterShiftFound:
    
    If Not WaterShiftFound Then
        PrecisionVisionGraph.AdjustWaterCentre.value = 0
        PrecisionVisionGraph.WaterShiftType(0).Visible = False
        PrecisionVisionGraph.WaterShiftType(1).Visible = False
    Else
        PrecisionVisionGraph.WaterShiftType(0).Visible = True
        PrecisionVisionGraph.WaterShiftType(1).Visible = True
    End If

Exit Sub
Err_Handler:
    Select Case Err
        Case 9: GoTo NoWaterShiftFound ' There is no ReferenceShapeShiftObs
        Case Else: MsgBox Err & "-OBS8:" & Error$
    End Select
End Sub

Function ObsCalcPos(ByVal ObsIndex As Integer) As Single
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PVObsCalcPos
'Created : 18 March 2005, PCN3490
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Translates the Obs frame no to a position along the Display Pipe.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

If PVDataNoOfLines < 1 Then
    ObsCalcPos = -1 'Invalid calculation
    Exit Function
End If

ObsCalcPos = YCalcFromFrame(PipeObservations(ObsIndex).PipeObsFrameNo)

Exit Function
Err_Handler:
    MsgBox Err & "-OBS9:" & Error$
    ObsCalcPos = -1 'Invalid calculation
End Function


Function YCalcFromFrame(ByVal Frame As Long) As Single 'PCN4465 was not suppose to be integer, now long
On Error GoTo Err_Handler

If PVDataNoOfLines < 1 Then
    YCalcFromFrame = -1 'Invalid calculation
    Exit Function
End If

YCalcFromFrame = (Frame / PVDataNoOfLines) * PrecisionVisionGraph.PipeMain.height
YCalcFromFrame = YCalcFromFrame + PrecisionVisionGraph.PipeMain.Top

Exit Function
Err_Handler:
    MsgBox Err & "-OBS10:" & Error$
    YCalcFromFrame = -1 'Invalid calculation
End Function

Function FrameCalcFromY(ByVal Y As Single) As Single
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PVObsCalcPos
'Created : 18 March 2005, PCN3490
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Translates the Obs frame no to a position along the Display Pipe.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

If PVDataNoOfLines < 1 Then
    FrameCalcFromY = -1 'Invalid calculation
    Exit Function
End If

Y = Y - PrecisionVisionGraph.PipeMain.Top
FrameCalcFromY = (Y / PrecisionVisionGraph.PipeMain.height) * PVDataNoOfLines

Exit Function
Err_Handler:
    MsgBox Err & "-OBS11:" & Error$
    FrameCalcFromY = -1 'Invalid calculation
End Function

'Function PVObsMouseOverMe(Y As Single, ByRef ObsIndexArray As Integer) As Boolean
Function ObsMouseOverMe(Y As Single) As Integer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PVObsMouseOverMe
'Created : 18 March 2005, PCN3490
'Updated :
'Prg By  : Geoff Logan
'Param   : ObsType either "Observation" or "Distance"
'Desc    : Determines if the mouse is currently over an observation.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim ObsIndex As Integer
Dim ObsPosOnPipeDisplay As Single
Dim YPosAsFramesUpper As Single
Dim YPosAsFramesLower As Single
Dim YPosAsFrames As Single

Dim arrayIndex As Integer
Dim Closest As Integer
Dim ObsStep As Single

ObsMouseOverMe = 0 'Observation was not found
If PVDataNoOfLines < 1 Then Exit Function

ObsStep = 100
arrayIndex = 0
Closest = 1

YPosAsFramesLower = FrameCalcFromY(Y - ObsStep)
YPosAsFramesUpper = FrameCalcFromY(Y + ObsStep)
YPosAsFrames = FrameCalcFromY(Y)

For ObsIndex = 1 To NoOfPipeObservations
    'Only display Observations not distance markers
    If Left(PipeObservations(ObsIndex).PipeObs, 5) <> "[[R]]" And _
       Left(PipeObservations(ObsIndex).PipeObs, 6) <> "[[SH]]" And _
       Left(PipeObservations(ObsIndex).PipeObs, 6) <> "[[WT]]" And _
        (LastObservationType <> "Distance" And Trim(PipeObservations(ObsIndex).PipeObs) <> "<<<--I-->>>") Or _
       (LastObservationType = "Distance" And Trim(PipeObservations(ObsIndex).PipeObs) = "<<<--I-->>>") Then
        With PipeObservations(ObsIndex)
            If .PipeObsFrameNo >= 0 Then   'Tests calculation was valid.
                'Test to see in Y is close to ObsPosOnPipeDisplay
                If (.PipeObsFrameNo > YPosAsFramesLower) And (.PipeObsFrameNo < YPosAsFramesUpper) Then
                    'Mouse is over observation.
                    arrayIndex = arrayIndex + 1
                    ReDim Preserve ObsIndexArray(arrayIndex)
                    ObsIndexArray(arrayIndex) = ObsIndex
                    If Abs(.PipeObsFrameNo - YPosAsFrames) < Abs(PipeObservations(Closest).PipeObsFrameNo - YPosAsFrames) Then
                        Closest = ObsIndex
                    End If
                    ObsMouseOverMe = Closest
                End If
            End If
        End With
    End If
Next ObsIndex


Exit Function
Err_Handler:
    MsgBox Err & "-OBS12:" & Error$
    ObsMouseOverMe = 0 'Observation was not found
End Function

Function ObsAddNewOrUpdateSetup(FrameNo As Long, Optional PipeObsIndexNo)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PVObsAddNewOrUpdateSetup
'Created : 19 March 2005, PCN3490
'Updated :
'Prg By  : Geoff Logan
'Param   : FrameNo - Frame number
'Desc    : Sets up the observation for data entry or update at FrameNo.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim DisplayPipeHeight As Single
Dim ObsPosOnPipeDisplay As Single
Dim ObsBubbleShiftOffset As Single
Dim ObsIndex As Integer

Dim FileSaveFail As Boolean

If PVDataNoOfLines < 1 Then
    Exit Function
End If

DisplayPipeHeight = PrecisionVisionGraph.PipeMain.height - 90 'As per the setup of the ViewIndicators

'Make invisible the unused controls
ObsIndex = 1
PrecisionVisionGraph.ObservationEye(ObsIndex).Visible = True
'PrecisionVisionGraph.ObservationText(ObsIndex).Caption = "Fr " & FrameNo
If IsMissing(PipeObsIndexNo) Then
    If ConfigInfo.DistanceStart >= 0 Then
        PrecisionVisionGraph.PVObsDistance.text = Round(PVDistances(FrameNo), 2)
    End If
    'Setup New flag
    AddNewObsFlag = True
    'Reset PVObsIndexNo
    PrecisionVisionGraph.PVObsIndexNo(1) = "" 'PCN4131 was me on the start of it
    ObsPosOnPipeDisplay = YCalcFromFrame(FrameNo)
ElseIf PipeObsIndexNo = 0 Then
    If ConfigInfo.DistanceStart >= 0 Then
        PrecisionVisionGraph.PVObsDistance.text = Round(PVDistances(PipeObservations(PipeObsIndexNo).PipeObsFrameNo), 2)
    End If
    'Setup New flag
    AddNewObsFlag = True
    'Reset PVObsIndexNo
    PrecisionVisionGraph.PVObsIndexNo(1) = ""
Else
    PrecisionVisionGraph.PVObsDistance.text = Round(PVDistances(PipeObservations(PipeObsIndexNo).PipeObsFrameNo), 2)
    PrecisionVisionGraph.PVObsEntry.text = Trim(PipeObservations(PipeObsIndexNo).PipeObs)
    'Setup Update flag
    AddNewObsFlag = False
    PrecisionVisionGraph.PVObsIndexNo(1) = PipeObsIndexNo
    ObsPosOnPipeDisplay = ObsCalcPos(PipeObsIndexNo)
End If
If MeasurementUnits = "mm" Then
    PrecisionVisionGraph.PVObsDistanceUnit.Caption = "m"
Else
    PrecisionVisionGraph.PVObsDistanceUnit.Caption = "ft"
End If
PrecisionVisionGraph.PVObsEntry.height = 1500
PrecisionVisionGraph.PVObsEntry.Visible = True
PrecisionVisionGraph.PVObsDistance.Top = 120: PrecisionVisionGraph.PVObsDistance.Visible = True
PrecisionVisionGraph.PVObsDistanceUnit.Top = 120: PrecisionVisionGraph.PVObsDistanceUnit.Visible = True
PrecisionVisionGraph.ObservationText(ObsIndex).Visible = True
For ObsIndex = 1 To PrecisionVisionGraph.ObservationText.Count
    PrecisionVisionGraph.ObservationText(ObsIndex).Visible = False
    PrecisionVisionGraph.ObservationEye(ObsIndex).Visible = False
Next ObsIndex
'Setup height of Bubble
PrecisionVisionGraph.ObservationBubble.height = 2200
PrecisionVisionGraph.ObservationClose.Picture = LoadResPicture(109, vbResBitmap)
PrecisionVisionGraph.Observation.height = PrecisionVisionGraph.ObservationBubble.height

PrecisionVisionGraph.SpeakObs.Top = ObsPosOnPipeDisplay - 160
ObsBubbleShiftOffset = ObsPosOnPipeDisplay + PrecisionVisionGraph.Observation.height - 1000 - PrecisionVisionGraph.PipeMain.height
PrecisionVisionGraph.PVObsEye.Top = ObsPosOnPipeDisplay - (PrecisionVisionGraph.PVObsEye.height / 2)
PrecisionVisionGraph.PVObsEye.Visible = True
If ObsBubbleShiftOffset < 0 Then ObsBubbleShiftOffset = 0
PrecisionVisionGraph.SpeakObsBlank.Top = 250 + ObsBubbleShiftOffset
PrecisionVisionGraph.Observation.Top = PrecisionVisionGraph.SpeakObs.Top + 240 - ObsBubbleShiftOffset
PrecisionVisionGraph.Observation.Visible = True
PrecisionVisionGraph.SpeakObs.Visible = True
PrecisionVisionGraph.SpeakObs.ZOrder 0

If IsMissing(PipeObsIndexNo) Then
    Dim NumberImages As Integer
    Dim LookForNewOb As Integer
    NoOfPipeObservations = NoOfPipeObservations + 1
    ReDim Preserve PipeObservations(NoOfPipeObservations)
    
    PipeObservations(NoOfPipeObservations).PipeObs = ""
    PipeObservations(NoOfPipeObservations).PipeObsFrameNo = FrameNo
    PipeObservations(NoOfPipeObservations).PipeObsSnapshotLength = 0
    PipeObservations(NoOfPipeObservations).PipeObsSnapshotOffset = -1
    
    NumberImages = PrecisionVisionGraph.ObsPictureStorage.Count
    Load PrecisionVisionGraph.ObsPictureStorage(NumberImages)
    
    Call SortObs
    
    Call ClearObsMarkings
    Call ObsDisplayALL
    
    'Find just added observation back after sort routine
    
    Call ObservationUpdate
    PrecisionVisionGraph.PVObsEntry.text = ""
    For LookForNewOb = 1 To NoOfPipeObservations
        If PipeObservations(LookForNewOb).PipeObsSnapshotOffset = -1 Then
            PipeObservations(LookForNewOb).PipeObsSnapshotOffset = 0
            Exit For
        End If
    Next LookForNewOb
    
    ReDim ObsIndexArray(1)
    ObsIndexArray(1) = LookForNewOb
    ObsOpenIndex = LookForNewOb
    Call SaveToFilePipeObs(FileSaveFail)
    Call Observations.AttachObsPicture
    
    
    
Else
    Call ObsPreviewDisplay
End If

PrecisionVisionGraph.PVObsEntry.SetFocus 'PCN6026

Exit Function
Err_Handler:
    MsgBox Err & "-OBS13:" & Error$

    
End Function

Function ObservationAddNew()
'****************************************************************************************
'Name    : PVObservationAddNew
'Created : 21 March 2005, PCN3490
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Add a new observation to the Pipe Observations
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
Dim CheckDistValue As Double
Dim FileSaveFail As Boolean

If Not AddNewObsFlag Then Exit Function
'Ensure NoOfPipeObservations does not exceed the PipeObsBuffer
If NoOfPipeObservations >= PipeObsBuffer - 4 Then Exit Function 'PCN4433

On Error GoTo InvalidData
CheckDistValue = SafeCDbl(PVObsDistance.text) 'PCN4161
On Error GoTo Err_Handler
'Check there is a valid observation
If Len(PVObsEntry) = 0 Then Exit Function
'Store this new setting in the Pipe Observation array
NoOfPipeObservations = NoOfPipeObservations + 1
ReDim Preserve PipeObservations(NoOfPipeObservations)
PipeObservations(NoOfPipeObservations).PipeObs = PVObsEntry
PipeObservations(NoOfPipeObservations).PipeObsFrameNo = PVFrameNo
PipeObservations(NoOfPipeObservations).PipeObsDist = CheckDistValue
'Store NoOfPipeObservations in PVObsIndexNo
PrecisionVisionGraph.PVObsIndexNo(1) = NoOfPipeObservations
'Write back to the PVD
Call SaveToFilePipeObs(FileSaveFail)

'Reset AddNewObsFlag
AddNewObsFlag = 0
DoEvents

InvalidData:
    'Do nothing
Exit Function
Err_Handler:
    MsgBox Err & "-OBS14:" & Error$
End Function

Function ObservationUpdate()
'****************************************************************************************
'Name    : PVObservationUpdate
'Created : 22 March 2005, PCN3490
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Update an observation in the Pipe Observations (PVD).
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
Dim CheckPVObsIndexNo As Integer
Dim CheckDistValue As Double
Dim FileSaveFail As Boolean


On Error GoTo InvalidData
CheckPVObsIndexNo = CInt(PrecisionVisionGraph.PVObsIndexNo(1))
If CheckPVObsIndexNo = 0 Then Exit Function
CheckDistValue = SafeCDbl(PrecisionVisionGraph.PVObsDistance.text) 'PCN4161
On Error GoTo Err_Handler
'Check there is a valid observation
If Len(PrecisionVisionGraph.PVObsEntry) = 0 Then Exit Function
'Update values in Pipe Observations
PipeObservations(CheckPVObsIndexNo).PipeObsDist = CheckDistValue
PipeObservations(CheckPVObsIndexNo).PipeObs = PrecisionVisionGraph.PVObsEntry
'Write back to the PVD
Call SaveToFilePipeObs(FileSaveFail)

InvalidData:
    'Do nothing
Exit Function
Err_Handler:
    MsgBox Err & "-OBS15:" & Error$
End Function





Sub ObsInitEmbeddedFiles()
On Error GoTo Err_Handler

Dim NumberObs As Integer
Dim ObsIndex As Integer
Dim NumberEmbeddedFiles

NumberEmbeddedFiles = 1

NumberObs = UBound(PipeObservations)

For ObsIndex = 1 To NumberObs
    If Trim(PipeObservations(ObsIndex).PipeObs) <> "<<<--I-->>>" Then ' PCN4233 Was = "<<<--I-->> then exit for, but now just egnore becuase
                                                                      ' we have the imbedded reports after distance to list
        ReDim Preserve ListEmbeddedOwners(ObsIndex)
        
        With PipeObservations(ObsIndex)
            ListEmbeddedOwners(ObsIndex).EmbeddedIndex = ObsIndex
            ListEmbeddedOwners(ObsIndex).EmbeddedType = "" ' PCN4233 Dont know what it is untill loaded header, "Obs" or "Rep"
            ListEmbeddedOwners(ObsIndex).FileLength = .PipeObsSnapshotLength
            ListEmbeddedOwners(ObsIndex).FileOffset = .PipeObsSnapshotOffset
            If ListEmbeddedOwners(ObsIndex).FileLength < 0 Then ListEmbeddedOwners(ObsIndex).FileLength = 0
            If ListEmbeddedOwners(ObsIndex).FileOffset < 0 Then ListEmbeddedOwners(ObsIndex).FileOffset = 0
        End With
        With ListEmbeddedOwners(ObsIndex)
            If .FileLength <> 0 And .FileOffset <> 0 And Dir(PVDFileName) <> "" And PVDFileName <> "" Then
                FileNoPVD = FreeFile
                Open PVDFileName For Binary Access Read As #FileNoPVD
                Seek #FileNoPVD, .FileOffset
                Get #FileNoPVD, , .PVHeaderEmbedded
                Close #FileNoPVD
            End If
        End With
        ListEmbeddedOwners(ObsIndex).EmbeddedType = ListEmbeddedOwners(ObsIndex).PVHeaderEmbedded.Owner
        
    End If
Next ObsIndex
        
Call PageFunctions.SortEmbeddedList
        
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-OBS16:" & Error$
    End Select

End Sub

Sub ObsInitPictureStorage()
On Error GoTo Err_Handler

Dim ObsIndex As Integer
Dim NumberObs As Integer
Dim RulerIndex As Integer
Dim NumberRulers As Integer
Dim PVDHeaderEmbedded As PVDHeaderEmbeddedType



NumberObs = UBound(PipeObservations)


For ObsIndex = PrecisionVisionGraph.ObsPictureStorage.Count - 1 To 1 Step -1
    Unload PrecisionVisionGraph.ObsPictureStorage(ObsIndex)
Next ObsIndex

'For RulerIndex = PrecisionVisionGraph.RulerMark.Count - 1 To 1 Step -1
'    Unload PrecisionVisionGraph.RulerMark(RulerIndex)
'Next RulerIndex

For ObsIndex = 1 To (NumberObs)
    If Trim(PipeObservations(ObsIndex).PipeObs) = "<<<--I-->>>" Or _
        Left(PipeObservations(ObsIndex).PipeObs, 5) = "[[R]]" Or _
        Left(PipeObservations(ObsIndex).PipeObs, 6) = "[[SH]]" Then Exit For
    Load PrecisionVisionGraph.ObsPictureStorage(ObsIndex)
    If PipeObservations(ObsIndex).PipeObsSnapshotOffset <> 0 Then
        Call PageFunctions.EmbeddedFileExtract(WindowsTempDirectory & EmbeddedFileNameAndPath, _
                                                PipeObservations(ObsIndex).PipeObsSnapshotOffset, _
                                                PipeObservations(ObsIndex).PipeObsSnapshotLength, _
                                                PVDHeaderEmbedded)
        PrecisionVisionGraph.ObsPictureStorage(ObsIndex).Picture = LoadPicture(WindowsTempDirectory & EmbeddedFileNameAndPath)
        PrecisionVisionGraph.ObsPictureStorage(ObsIndex).Tag = SnapShot
    Else
        PrecisionVisionGraph.ObsPictureStorage(ObsIndex).Tag = PV
    End If
NextObs:
Next ObsIndex


Exit Sub
Err_Handler:
    Select Case Err
        Case 481
            PrecisionVisionGraph.ObsPictureStorage(ObsIndex).Tag = "": Resume NextObs
        Case Else: MsgBox Err & "-OBS17:" & Error$
    
    End Select
End Sub

'I appoligise for the long function, but its simple and neat.


Sub SortObs()
On Error GoTo Err_Handler

Dim Swaped As Boolean
Dim ObsIndex As Integer
Dim NumberObs As Integer
Dim DistanceIndexStart As Integer
Dim ReportIndexStart As Integer
Dim ShapeCentreIndexStart As Integer
Dim WaterCentreIndexStart As Integer
Swaped = True

'Shift Distances to end
NumberObs = UBound(PipeObservations)
DistanceIndexStart = NumberObs + 1
ReportIndexStart = NumberObs + 1
ShapeCentreIndexStart = NumberObs + 1
WaterCentreIndexStart = NumberObs + 1


'Remove leading spaces from obs'


'Sort observations from (distances , reports, shape centre and water height) 'PCN4484 added shape and water

While Swaped = True
    Swaped = False
    For ObsIndex = NumberObs To 2 Step -1
        If ((Trim(PipeObservations(ObsIndex).PipeObs) <> "<<<--I-->>>" And _
             Left(PipeObservations(ObsIndex).PipeObs, 5) <> "[[R]]") And _
             Left(PipeObservations(ObsIndex).PipeObs, 6) <> "[[SH]]" And _
             Left(PipeObservations(ObsIndex).PipeObs, 6) <> "[[WT]]") And _
          (Trim(PipeObservations(ObsIndex - 1).PipeObs) = "<<<--I-->>>" Or _
           Left(PipeObservations(ObsIndex - 1).PipeObs, 5) = "[[R]]" Or _
           Left(PipeObservations(ObsIndex - 1).PipeObs, 6) = "[[SH]]" Or _
           Left(PipeObservations(ObsIndex - 1).PipeObs, 6) = "[[WT]]") Then
            Call SwapObs(ObsIndex, ObsIndex - 1)
            Swaped = True
        End If
    Next ObsIndex
Wend
'Sort distance from reports, shape centre and water height
Swaped = True
While Swaped = True
    Swaped = False
    For ObsIndex = NumberObs To 2 Step -1
        If Trim(PipeObservations(ObsIndex).PipeObs) = "<<<--I-->>>" And _
           (Left(PipeObservations(ObsIndex - 1).PipeObs, 5) = "[[R]]" Or _
            Left(PipeObservations(ObsIndex - 1).PipeObs, 6) = "[[SH]]" Or _
            Left(PipeObservations(ObsIndex - 1).PipeObs, 6) = "[[WT]]") Then
            Call SwapObs(ObsIndex, ObsIndex - 1)
            Swaped = True
        End If
    Next ObsIndex
Wend

'Sort reports from shape centre and water height
Swaped = True
While Swaped = True
    Swaped = False
    For ObsIndex = NumberObs To 2 Step -1
        If Left(PipeObservations(ObsIndex - 1).PipeObs, 5) = "[[R]]" And _
           (Left(PipeObservations(ObsIndex - 1).PipeObs, 6) = "[[SH]]" Or _
            Left(PipeObservations(ObsIndex - 1).PipeObs, 6) = "[[WT]]") Then
            Call SwapObs(ObsIndex, ObsIndex - 1)
            Swaped = True
        End If
    Next ObsIndex
Wend

'Sort shape centre from water height
Swaped = True
While Swaped = True
    Swaped = False
    For ObsIndex = NumberObs To 2 Step -1
        If Left(PipeObservations(ObsIndex - 1).PipeObs, 6) = "[[SH]]" And _
           (Left(PipeObservations(ObsIndex - 1).PipeObs, 6) = "[[WT]]") Then
            Call SwapObs(ObsIndex, ObsIndex - 1)
            Swaped = True
        End If
    Next ObsIndex
Wend

'Sort observations by frame
Swaped = True
While Swaped = True
    Swaped = False
    For ObsIndex = 1 To (NumberObs - 1)
        If Trim(PipeObservations(ObsIndex).PipeObs) = "<<<--I-->>>" Or _
           Left(PipeObservations(ObsIndex).PipeObs, 5) = "[[R]]" Or _
           Left(PipeObservations(ObsIndex).PipeObs, 6) = "[[SH]]" Or _
           Left(PipeObservations(ObsIndex).PipeObs, 6) = "[[WT]]" _
                Then DistanceIndexStart = ObsIndex: Exit For
    
        If Trim(PipeObservations(ObsIndex + 1).PipeObs) = "<<<--I-->>>" Or _
           Left(PipeObservations(ObsIndex + 1).PipeObs, 5) = "[[R]]" Or _
           Left(PipeObservations(ObsIndex + 1).PipeObs, 6) = "[[SH]]" Or _
           Left(PipeObservations(ObsIndex + 1).PipeObs, 6) = "[[WT]]" _
                Then DistanceIndexStart = ObsIndex + 1: Exit For
                
        If PipeObservations(ObsIndex + 1).PipeObsFrameNo < PipeObservations(ObsIndex).PipeObsFrameNo Then
            Call SwapObs(ObsIndex, ObsIndex + 1)
            Swaped = True
        End If
    Next ObsIndex
Wend

'Sort distances by frame
Swaped = True
While Swaped = True
    Swaped = False
    For ObsIndex = DistanceIndexStart To (NumberObs - 1)
        If Left(PipeObservations(ObsIndex).PipeObs, 5) = "[[R]]" Or _
           Left(PipeObservations(ObsIndex).PipeObs, 6) = "[[SH]]" Or _
           Left(PipeObservations(ObsIndex).PipeObs, 6) = "[[WT]]" Then ReportIndexStart = ObsIndex: Exit For
           
        If Left(PipeObservations(ObsIndex + 1).PipeObs, 5) = "[[R]]" Or _
           Left(PipeObservations(ObsIndex + 1).PipeObs, 6) = "[[SH]]" Or _
           Left(PipeObservations(ObsIndex + 1).PipeObs, 6) = "[[WT]]" Then ReportIndexStart = ObsIndex + 1: Exit For
        If PipeObservations(ObsIndex + 1).PipeObsFrameNo < PipeObservations(ObsIndex).PipeObsFrameNo Then
            Call SwapObs(ObsIndex, ObsIndex + 1)
            Swaped = True
        End If
    Next ObsIndex
Wend

'Sort stored reports by report no
Swaped = True
While Swaped = True
    Swaped = False
    For ObsIndex = ReportIndexStart To (NumberObs - 1)
        If Left(PipeObservations(ObsIndex).PipeObs, 6) = "[[SH]]" Or _
           Left(PipeObservations(ObsIndex).PipeObs, 6) = "[[WT]]" Then ShapeCentreIndexStart = ObsIndex: Exit For
        
        If Left(PipeObservations(ObsIndex + 1).PipeObs, 6) = "[[SH]]" Or _
           Left(PipeObservations(ObsIndex + 1).PipeObs, 6) = "[[WT]]" Then ShapeCentreIndexStart = ObsIndex + 1: Exit For
        
        If PipeObservations(ObsIndex + 1).PipeObsFrameNo < PipeObservations(ObsIndex).PipeObsFrameNo Then
            Call SwapObs(ObsIndex, ObsIndex + 1)
            Swaped = True
        End If
    Next ObsIndex
Wend

'Sort stored reference shape centres by frame
Swaped = True
While Swaped = True
    Swaped = False
    For ObsIndex = ShapeCentreIndexStart To (NumberObs - 1)
        If Left(PipeObservations(ObsIndex).PipeObs, 6) = "[[WT]]" Then WaterCentreIndexStart = ObsIndex: Exit For
        If Left(PipeObservations(ObsIndex + 1).PipeObs, 6) = "[[WT]]" Then WaterCentreIndexStart = ObsIndex + 1: Exit For
        
        If PipeObservations(ObsIndex + 1).PipeObsFrameNo < PipeObservations(ObsIndex).PipeObsFrameNo Then
            Call SwapObs(ObsIndex, ObsIndex + 1)
            Swaped = True
        End If
    Next ObsIndex
Wend

'Sort stored water centres by frame
Swaped = True
While Swaped = True
    Swaped = False
    For ObsIndex = ShapeCentreIndexStart To (NumberObs - 1)
        If PipeObservations(ObsIndex + 1).PipeObsFrameNo < PipeObservations(ObsIndex).PipeObsFrameNo Then
            Call SwapObs(ObsIndex, ObsIndex + 1)
            Swaped = True
        End If
    Next ObsIndex
Wend

Call Observations.ObsInitEmbeddedFiles
Call PageFunctions.StoredReportInitialise

If ShiftOn Then
    Call Observations.PVShapeCentreInitialise 'PCN4484
    Call Observations.PVWaterLevelInitialise 'PCN4484
End If



Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-OBS18:" & Error$
    
    End Select
End Sub



Sub SwapObs(ByVal a As Integer, ByVal b As Integer)
On Error GoTo Err_Handler



Dim t As PipeObservationType_V60


    Call ObsAEqB(t, PipeObservations(a))
    Set PrecisionVisionGraph.ObsTempPicture.Picture = PrecisionVisionGraph.ObsPictureStorage(a).Picture
    PrecisionVisionGraph.ObsTempPicture.Tag = PrecisionVisionGraph.ObsPictureStorage(a).Tag
        
    Call ObsAEqB(PipeObservations(a), PipeObservations(b))
    Set PrecisionVisionGraph.ObsPictureStorage(a).Picture = PrecisionVisionGraph.ObsPictureStorage(b).Picture
    PrecisionVisionGraph.ObsPictureStorage(a).Tag = PrecisionVisionGraph.ObsPictureStorage(b).Tag
    
    Call ObsAEqB(PipeObservations(b), t)
    Set PrecisionVisionGraph.ObsPictureStorage(b).Picture = PrecisionVisionGraph.ObsTempPicture.Picture
    PrecisionVisionGraph.ObsPictureStorage(b).Tag = PrecisionVisionGraph.ObsTempPicture.Tag


    
Exit Sub
Err_Handler:
    Select Case Err
        Case 340: Resume Next 'ANT
        Case Else: MsgBox Err & "-OBS19:" & Error$
    End Select
End Sub

Sub ObsAEqB(ByRef a As PipeObservationType_V60, ByRef b As PipeObservationType_V60)
On Error GoTo Err_Handler

    a.PipeObs = b.PipeObs
    a.PipeObsDist = b.PipeObsDist
    a.PipeObsFrameNo = b.PipeObsFrameNo
    a.PipeObsSnapshotLength = b.PipeObsSnapshotLength
    a.PipeObsSnapshotOffset = b.PipeObsSnapshotOffset
    
    

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-OBS20:" & Error$
    End Select
End Sub

Function GetLineCount(object As Variant) As Integer
On Error GoTo Err_Handler

With PrecisionVisionGraph.CalculateNoLines

On Error GoTo TryText
    .text = object.Caption
    GoTo SetTextBox

TryText:
    On Error GoTo Err_Handler
    .text = object.text
  
    
    
SetTextBox:
    
    .width = object.width
    .Font = object.Font
    .Font.Bold = object.Font.Bold
    .Font.Charset = object.Font.Charset
    .Font.Italic = object.Font.Italic
    .Font.name = object.Font.name
    .Font.Size = object.Font.Size
    .Font.Strikethrough = object.Font.Strikethrough
    .Font.Underline = object.Font.Underline
    .Font.Weight = object.Font.Weight
    
    
    
    GetLineCount = SendMessage(.hwnd, EM_GETLINECOUNT, ByVal 1&, ByVal 0&)

End With
Exit Function
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-OBS21:" & Error$
    End Select
End Function

Sub ObsEye_Click(Index As Integer)
On Error GoTo Err_Handler

Dim ObsSelected As Integer

If UBound(ObsIndexArray) = 1 Then Exit Sub

ObsSelected = ObsIndexArray(Index)

ReDim ObsIndexArray(1)
ObsIndexArray(1) = ObsSelected

Call ObsBubbleDisplay(1, 1)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-OBS22:" & Error$

    End Select
End Sub

Sub AttachObsPicture()
On Error GoTo Err_Handler

    Dim SelectedObs As Integer
    
   
    Call ClearLineScreen.SnapShotForEmbededFile
    With PrecisionVisionGraph.ObsPictureStorage(ObsOpenIndex)
        '.Picture = LoadPicture(WindowsTempDirectory & "CBS\EmbedFile.jpg")
        .Picture = LoadPicture(WindowsTempDirectory & "EmbedFile.jpg") 'PCN4601
        .Tag = SnapShot
        Call EmbedObsPicture
    End With
        
    




    
    Call ObsPreviewDisplay
        
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-OBS23:" & Error$
    End Select
End Sub

Sub EmbedObsPicture()
On Error GoTo Err_Handler

Dim PVDHeaderEmbedded As PVDHeaderEmbeddedType
Dim FileSaveFail As Boolean


'Check to see if allready embededfile for Obs
With PrecisionVisionGraph.ObsPictureStorage(ObsOpenIndex)
    If PipeObservations(ObsOpenIndex).PipeObsSnapshotLength <> 0 Then
        Call PageFunctions.EmbeddedFileRemove(PipeObservations(ObsOpenIndex).PipeObsSnapshotOffset, _
                                              PipeObservations(ObsOpenIndex).PipeObsSnapshotLength)
    End If
End With



 
PVDHeaderEmbedded.Descriptor = "[EmbeddedFile]"
PVDHeaderEmbedded.Owner = "Obs"
PVDHeaderEmbedded.FileLength = 1
PVDHeaderEmbedded.PVDCheck = 0

'Call EmbedFile(WindowsTempDirectory & "cbs\embedfile.jpg", _ 'ID4601
Call EmbedFile(WindowsTempDirectory & "embedfile.jpg", _
               PipeObservations(ObsOpenIndex).PipeObsSnapshotOffset, _
               PipeObservations(ObsOpenIndex).PipeObsSnapshotLength, _
               PVDHeaderEmbedded)
Call SaveToFilePipeObs(FileSaveFail)

Call ObsInitEmbeddedFiles

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-OBS24:" & Error$
    End Select
End Sub

Sub ObsDelete(ByVal Index)
On Error GoTo Err_Handler

Dim FileSaveFail As Boolean
Dim TotalObs As Integer
Dim TotalObsStorage As Integer
Dim ObsIndex As Integer
Dim IsItReport As String



TotalObs = UBound(PipeObservations)
TotalObsStorage = PrecisionVisionGraph.ObsPictureStorage.Count


If Index > TotalObs Then Exit Sub
IsItReport = Left(PipeObservations(Index).PipeObs, 5)

If PipeObservations(Index).PipeObsSnapshotOffset <> 0 Then
    Call PageFunctions.EmbeddedFileRemove(PipeObservations(Index).PipeObsSnapshotOffset, _
                                          PipeObservations(Index).PipeObsSnapshotLength)
End If

If Index < 1 Then Exit Sub

For ObsIndex = Index To TotalObs - 1
    Call ObsAEqB(PipeObservations(ObsIndex), PipeObservations(ObsIndex + 1))
    If ObsIndex + 1 < TotalObsStorage Then
        Set PrecisionVisionGraph.ObsPictureStorage(ObsIndex).Picture = PrecisionVisionGraph.ObsPictureStorage(ObsIndex + 1).Picture
        PrecisionVisionGraph.ObsPictureStorage(ObsIndex).Tag = PrecisionVisionGraph.ObsPictureStorage(ObsIndex + 1).Tag
    End If
Next ObsIndex

ReDim Preserve PipeObservations(TotalObs - 1)

If IsItReport <> "[[R]]" Then If LastObservationType <> "Distance" Then Unload PrecisionVisionGraph.ObsPictureStorage(TotalObsStorage - 1)

NoOfPipeObservations = TotalObs - 1

Call SortObs 'Call ObsInitEmbeddedFiles PCN4233, now the whole lot is sorted again
Call SaveToFilePipeObs(FileSaveFail)

Call ClearObsMarkings
Call ObsDisplayALL
ReDim ObsIndexArray(0)
If IsItReport <> "[[R]]" Then If LastObservationType = "Distance" Then Call Distance.RecalculateDistance



Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-OBS25:" & Error$
    End Select
End Sub

Sub ClearAllObservationsAndDistanceSettings()
On Error GoTo Err_Handler
    Dim RulerIndex As Integer
    
    ReDim PipeObservations(0)
    NoOfPipeObservations = 0
    
    For RulerIndex = 1 To PrecisionVisionGraph.RulerMark.Count - 1
        Unload PrecisionVisionGraph.RulerMark(RulerIndex)
    Next RulerIndex
     
    
    

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-OBS26:" & Error$
    End Select
End Sub



Sub ClearEmptyObservations()
'****************************************************************************************
'Name    : ClearEmptyObservations
'PCN     : PCN4191
'Created : 16 August 2006, (about a week before this date)
'Updated :
'Prg By  : Antony van Iersel
'Param   :
'Desc    : It goes thru the observations array and deletes any with no comments
'Usage   : This is normally called when closing the observation entry box, and then
'          Checked to see if anything was actually entered.
'****************************************************************************************
On Error GoTo Err_Handler
Dim ObsDeleted As Boolean
Dim ObsIndex As Integer

ObsDeleted = True
TotalObs = UBound(PipeObservations)
If TotalObs < 1 Then Exit Sub
    
While ObsDeleted = True 'Loops thru until there are no obs with no text
    ObsDeleted = False                      'This will be set if there is an empty obs
    TotalObs = UBound(PipeObservations)     'Everytime we delete a obs the total has to be updated
    For ObsIndex = 1 To TotalObs            'Go thru all the obs and check for empty pipeobs
        If Trim(PipeObservations(ObsIndex).PipeObs) = "" Then 'If it is empty then delete obs
            ObsDeleted = True                                 'and flag to loop thru again
            Call ObsDelete(ObsIndex)
            Exit For
        End If
    Next ObsIndex
Wend

    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-OBS27:" & Error$
    End Select
End Sub

Function RemoveCrLf(ByVal TheString As String) As String
On Error GoTo Err_Handler
Dim WhereInCr As Integer
Dim WhereInLf As Integer

WhereInCr = InStr(TheString, vbCr)
If WhereInCr <> 0 Then
    TheString = Left(TheString, WhereInCr - 1) & Right(TheString, Len(TheString) - WhereInCr - 1)
    Call RemoveCrLf(TheString)
End If

RemoveCrLf = TheString


Exit Function
Err_Handler:
    MsgBox Err & "-OBS28:" & Error$
End Function

Sub PVShapeCentreInitialise()
On Error GoTo Err_Handler

Dim ShiftType As String
Dim a As Single
Dim b As Single


Dim TotalObs As Integer
Dim ObsIndex As Integer
Dim NumberSH As Integer

TotalObs = UBound(PipeObservations)
If TotalObs < 1 Then Exit Sub

NumberSH = 0
For ObsIndex = 1 To TotalObs
    If Left(PipeObservations(ObsIndex).PipeObs, 6) = "[[SH]]" Then
        NumberSH = NumberSH + 1
        ReDim Preserve ReferenceShapeShiftObs(NumberSH)
        ReferenceShapeShiftObs(NumberSH).Distance = PipeObservations(ObsIndex).PipeObsDist
        ReferenceShapeShiftObs(NumberSH).Frame = PipeObservations(ObsIndex).PipeObsFrameNo
        Call ExtractPVShapeWaterData(PipeObservations(ObsIndex).PipeObs, _
             ShiftType, _
             a, b)
        ReferenceShapeShiftObs(NumberSH).ShiftType = ShiftType
        ReferenceShapeShiftObs(NumberSH).a = a
        ReferenceShapeShiftObs(NumberSH).b = b
    End If
Next ObsIndex

Exit Sub
Err_Handler:
    MsgBox Err & "-OBS29:" & Error$

End Sub

Sub PVWaterLevelInitialise()
On Error GoTo Err_Handler
Dim ShiftType As String
Dim a As Single
Dim b As Single

Dim TotalObs As Integer
Dim ObsIndex As Integer
Dim NumberWT As Integer


TotalObs = UBound(PipeObservations)
If TotalObs < 1 Then Exit Sub

NumberWT = 0
For ObsIndex = 1 To TotalObs
    If Left(PipeObservations(ObsIndex).PipeObs, 6) = "[[WT]]" Then
        NumberWT = NumberWT + 1
        ReDim Preserve WaterLevelShiftObs(NumberWT)
        WaterLevelShiftObs(NumberWT).Distance = PipeObservations(ObsIndex).PipeObsDist
        WaterLevelShiftObs(NumberWT).Frame = PipeObservations(ObsIndex).PipeObsFrameNo
        Call ExtractPVShapeWaterData(PipeObservations(ObsIndex).PipeObs, _
             ShiftType, _
             a, b)
        WaterLevelShiftObs(NumberWT).ShiftType = ShiftType
        WaterLevelShiftObs(NumberWT).a = a
        WaterLevelShiftObs(NumberWT).b = b
    End If
Next ObsIndex
Exit Sub
Err_Handler:
    MsgBox Err & "-OBS30:" & Error$
End Sub

Sub ExtractPVShapeWaterData(ByVal ShiftData As String, _
                             ByRef ShiftType As String, _
                             ByRef a As Single, _
                             ByRef b As Single)
On Error GoTo Err_Handler

Dim ShiftString As String
Dim IndexEnd As Integer
Dim IndexStart As Integer
Dim ExtractedString As String

ShiftString = ShiftData
IndexEnd = InStr(8, ShiftString, ")")
ShiftType = Mid(ShiftString, 9, IndexEnd - 9)

IndexStart = IndexEnd + 2
IndexEnd = InStr(IndexStart, ShiftString, ")")
ExtractedString = Mid(ShiftString, IndexStart, IndexEnd - IndexStart)
a = CSng(ExtractedString)

IndexStart = IndexEnd + 2
IndexEnd = InStr(IndexStart, ShiftString, ")")
ExtractedString = Mid(ShiftString, IndexStart, IndexEnd - IndexStart)
b = CSng(ExtractedString)

Exit Sub
Err_Handler:
    MsgBox Err & "-OBS31:" & Error$
End Sub

Sub AddRefShapeCentreChange()
On Error GoTo Err_Handler

    Dim TotalObs As Integer
    Dim ObsIndex As Integer
    
    Dim ObsString As String
    Dim ShiftType As String
    Dim a As String
    Dim b As String
    Dim FrameNo As Long
    Dim Distance As Double
    
    
    FrameNo = PipeObservations(ObsOpenIndex).PipeObsFrameNo
    For ObsIndex = 1 To UBound(ReferenceShapeShiftObs)
        If ReferenceShapeShiftObs(ObsIndex).Frame = FrameNo Then Exit Sub
    Next ObsIndex
    
NoShapeShiftFound:
    
    TotalObs = UBound(PipeObservations)
    TotalObs = TotalObs + 1
    ReDim Preserve PipeObservations(TotalObs)
    
    PrecisionVisionGraph.ReferenceShiftType(1).value = True
    
    If PrecisionVisionGraph.ReferenceShiftType(0).value = True Then
        ShiftType = GRADUAL
    Else
        ShiftType = EDGE
    End If
    
    a = CStr(ShapeCentreX)
    b = CStr(ShapeCentreY)
    
    Distance = PVDistances(FrameNo)
     
    
    PipeObservations(TotalObs).PipeObs = "[[SH]] (" & _
                                         ShiftType & _
                                         ")(" & _
                                         a & ")(" & b & ")"
    PipeObservations(TotalObs).PipeObsFrameNo = FrameNo
    PipeObservations(TotalObs).PipeObsDist = Distance
    
    
    Call SortObs
                                            
    PrecisionVisionGraph.ReferenceShiftType(0).Visible = True
    PrecisionVisionGraph.ReferenceShiftType(1).Visible = True
    
    

Exit Sub
Err_Handler:
    Select Case Err
        Case 9: GoTo NoShapeShiftFound ' There is no ReferenceShapeShiftObs
        Case Else: MsgBox Err & "-OBS32:" & Error$
    End Select
End Sub

Sub DelRefShapeCentreChange()
On Error GoTo Err_Handler
    Dim FrameNo As Long

    FrameNo = PipeObservations(ObsOpenIndex).PipeObsFrameNo
    For ObsIndex = 1 To UBound(ReferenceShapeShiftObs)
        If ReferenceShapeShiftObs(ObsIndex).Frame = FrameNo Then Exit Sub
    Next ObsIndex
Exit Sub
Err_Handler:
    Select Case Err
        Case 9: Exit Sub ' There is no ReferenceShapeShiftObs
        Case Else: MsgBox Err & "-OBS33:" & Error$
    End Select
End Sub

Sub AddWaterLevelChange()
On Error GoTo Err_Handler

    Dim TotalObs As Integer
    Dim ObsIndex As Integer
    
    Dim ObsString As String
    Dim ShiftType As String
    Dim a As String
    Dim b As String
    Dim FrameNo As Long
    Dim Distance As Double
    
    FrameNo = PipeObservations(ObsOpenIndex).PipeObsFrameNo
    For ObsIndex = 1 To UBound(WaterLevelShiftObs)
        If WaterLevelShiftObs(ObsIndex).Frame = FrameNo Then Exit Sub
    Next ObsIndex
    
NoWaterShiftFound:
    
    TotalObs = UBound(PipeObservations)
    TotalObs = TotalObs + 1
    ReDim Preserve PipeObservations(TotalObs)
     PrecisionVisionGraph.WaterShiftType(1).value = True
    If PrecisionVisionGraph.WaterShiftType(0).value = True Then
        ShiftType = GRADUAL
    Else
        ShiftType = EDGE
    End If
    
    a = CStr(ShapeCentreX)
    b = CStr(ShapeCentreY)

    Distance = PVDistances(FrameNo)
    
    PipeObservations(TotalObs).PipeObs = "[[WT]] (" & _
                                         ShiftType & _
                                         ")(" & _
                                         a & ")(" & b & ")"
    PipeObservations(TotalObs).PipeObsFrameNo = FrameNo
    PipeObservations(TotalObs).PipeObsDist = Distance
    
    
    Call SortObs
                                            
    PrecisionVisionGraph.WaterShiftType(0).Visible = True
    PrecisionVisionGraph.WaterShiftType(1).Visible = True
    
    

Exit Sub
Err_Handler:
    Select Case Err
        Case 9: GoTo NoWaterShiftFound ' There is no ReferenceShapeShiftObs
        Case Else: MsgBox Err & "-OBS34:" & Error$
    End Select
End Sub

Sub DelWaterLevelChange()
On Error GoTo Err_Handler
    Dim FrameNo As Long

    FrameNo = PipeObservations(ObsOpenIndex).PipeObsFrameNo
    For ObsIndex = 1 To UBound(WaterLevelShiftObs)
        If WaterLevelShiftObs(ObsIndex).Frame = FrameNo Then Exit Sub
    Next ObsIndex
Exit Sub
Err_Handler:
    Select Case Err
        Case 9: Exit Sub ' There is no ReferenceShapeShiftObs
        Case Else: MsgBox Err & "-OBS35:" & Error$
    End Select
End Sub
Sub SetShiftWaterLevel(ByVal ShiftType As String, ByVal ObsIndex As Integer)
On Error GoTo Err_Handler
    Dim FrameNo As Long
    Dim a As String
    Dim b As String
    Dim ShiftIndex As Integer
    Dim FoundMatchingObs As Boolean

    FrameNo = PipeObservations(ObsIndex).PipeObsFrameNo
    
    FoundMatchingObs = False
    For ShiftIndex = 1 To UBound(WaterLevelShiftObs)
        If WaterLevelShiftObs(ShiftIndex).Frame = FrameNo Then
            FoundMatchingObs = True
            Exit For
        End If
    Next ShiftIndex

    If Not FoundMatchingObs Then Exit Sub
    If WaterLevelShiftObs(ShiftIndex).ShiftType = ShiftType Then Exit Sub
    
    FoundMatchingObs = False
    For ObsIndex = 1 To UBound(PipeObservations)
        If Left(PipeObservations(ObsIndex).PipeObs, 6) = "[[WT]]" And _
           PipeObservations(ObsIndex).PipeObsFrameNo = FrameNo Then
           FoundMatchingObs = True
           Exit For
        End If
    Next ObsIndex

    If Not FoundMatchingObs Then Exit Sub
    
    
    a = WaterLevelShiftObs(ShiftIndex).a
    b = WaterLevelShiftObs(ShiftIndex).b

    PipeObservations(ObsIndex).PipeObs = "[[WT]] (" & _
                                         ShiftType & _
                                         ")(" & _
                                         a & ")(" & b & ")"
    Call SortObs
    
    
    

Exit Sub
Err_Handler:
    Select Case Err
        Case 9: Exit Sub ' There is no ReferenceShapeShiftObs
        Case Else: MsgBox Err & "-OBS36:" & Error$
    End Select
End Sub

Sub SetShiftRefereneShape(ByVal ShiftType As String, ByVal ObsIndex As Integer)
On Error GoTo Err_Handler

   Dim FrameNo As Long
    Dim a As String
    Dim b As String
    Dim ShiftIndex As Integer
    Dim FoundMatchingObs As Boolean
    

    FrameNo = PipeObservations(ObsIndex).PipeObsFrameNo
    FoundMatchingObs = False
    For ShiftIndex = 1 To UBound(ReferenceShapeShiftObs)
        If ReferenceShapeShiftObs(ShiftIndex).Frame = FrameNo Then
            FoundMatchingObs = True
            Exit For
        End If
    Next ShiftIndex

    If Not FoundMatchingObs Then Exit Sub
    If ReferenceShapeShiftObs(ShiftIndex).ShiftType = ShiftType Then Exit Sub
    
    FoundMatchingObs = False
    For ObsIndex = 1 To UBound(PipeObservations)
        If Left(PipeObservations(ObsIndex).PipeObs, 6) = "[[SH]]" And _
           PipeObservations(ObsIndex).PipeObsFrameNo = FrameNo Then
            FoundMatchingObs = True
            Exit For
        End If
    Next ObsIndex

    If Not FoundMatchingObs Then Exit Sub

    a = ReferenceShapeShiftObs(ShiftIndex).a
    b = ReferenceShapeShiftObs(ShiftIndex).b

    PipeObservations(ObsIndex).PipeObs = "[[SH]] (" & _
                                         ShiftType & _
                                         ")(" & _
                                         a & ")(" & b & ")"
                                         
    Call SortObs
    
    
    

Exit Sub
Err_Handler:
    Select Case Err
        Case 9: Exit Sub ' There is no ReferenceShapeShiftObs
        Case Else: MsgBox Err & "-OBS37:" & Error$
    End Select
End Sub

Sub RecalculateReferenceShapeShift()
On Error GoTo Err_Handler

Dim StartFrame As Long
Dim EndFrame As Long
Dim StartFrameX As Single
Dim StartFrameY As Single
Dim EndFrameX As Single
Dim EndFrameY As Single

Dim ShiftType As String


Dim ShiftIndex As Integer
Dim TotalShift As Integer

TotalShift = UBound(ReferenceShapeShiftObs)
If TotalShift < 1 Then Call ScreenDrawing.MoveAllFramesPVShapeCentre(ShapeCentreX, ShapeCentreY): Exit Sub ' There is no ReferenceShapeShiftObs

StartFrame = 1
StartFrameX = PVShapeCentreX(1)
StartFrameY = PVShapeCentreY(1)

EndFrame = ReferenceShapeShiftObs(1).Frame - 1
EndFrameX = ReferenceShapeShiftObs(1).a
EndFrameY = ReferenceShapeShiftObs(1).b
ShiftType = ReferenceShapeShiftObs(1).ShiftType

Call MoveFramesPVShapeCentre(StartFrame, _
                             StartFrameX, _
                             StartFrameY, _
                             EndFrame, _
                             EndFrameX, _
                             EndFrameY, _
                             ShiftType)


For ShiftIndex = 1 To TotalShift - 1

    StartFrame = ReferenceShapeShiftObs(ShiftIndex).Frame
    StartFrameX = ReferenceShapeShiftObs(ShiftIndex).a
    StartFrameY = ReferenceShapeShiftObs(ShiftIndex).b
    ShiftType = ReferenceShapeShiftObs(ShiftIndex).ShiftType


    EndFrame = ReferenceShapeShiftObs(ShiftIndex + 1).Frame - 1
    EndFrameX = ReferenceShapeShiftObs(ShiftIndex + 1).a
    EndFrameY = ReferenceShapeShiftObs(ShiftIndex + 1).b
    
    
Call MoveFramesPVShapeCentre(StartFrame, _
                             StartFrameX, _
                             StartFrameY, _
                             EndFrame, _
                             EndFrameX, _
                             EndFrameY, _
                             ShiftType)

Next ShiftIndex

StartFrame = ReferenceShapeShiftObs(TotalShift).Frame
StartFrameX = ReferenceShapeShiftObs(TotalShift).a
StartFrameY = ReferenceShapeShiftObs(TotalShift).b

EndFrame = PVDataNoOfLines
EndFrameX = PVShapeCentreX(PVDataNoOfLines)
EndFrameY = PVShapeCentreY(PVDataNoOfLines)

ShiftType = ReferenceShapeShiftObs(TotalShift).ShiftType

Call MoveFramesPVShapeCentre(StartFrame, _
                             StartFrameX, _
                             StartFrameY, _
                             EndFrame, _
                             EndFrameX, _
                             EndFrameY, _
                             ShiftType)

Exit Sub
Err_Handler:
    Select Case Err
        Case 9: Call ScreenDrawing.MoveAllFramesPVShapeCentre(ShapeCentreX, ShapeCentreY): Exit Sub ' There is no ReferenceShapeShiftObs
        Case Else: MsgBox Err & "-OBS38:" & Error$
    End Select
End Sub

Sub MoveFramesPVShapeCentre(ByVal StartFrame As Long, _
                            ByVal StartFrameX As Single, _
                            ByVal StartFrameY As Single, _
                            ByVal EndFrame As Long, _
                            ByVal EndFrameX As Single, _
                            ByVal EndFrameY As Single, _
                            ByVal ShiftType As String)
On Error GoTo Err_Handler
    Dim FrameNo As Long
    Dim XStep As Single
    Dim YStep As Single
    Dim Step As Long
    
    
    
    
    If ShiftType = EDGE Then
        For FrameNo = StartFrame To EndFrame
            PVShapeCentreX(FrameNo) = StartFrameX
            PVShapeCentreY(FrameNo) = StartFrameY
        Next FrameNo
        Exit Sub
    End If
    
    XStep = (StartFrameX - EndFrameX) / (EndFrame - StartFrame)
    YStep = (StartFrameY - EndFrameY) / (EndFrame - StartFrame)
    
    Step = 0
    For FrameNo = StartFrame To EndFrame
        PVShapeCentreX(FrameNo) = StartFrameX - (Step * XStep)
        PVShapeCentreY(FrameNo) = StartFrameY - (Step * YStep)
        Step = Step + 1
    Next FrameNo
        
                            
Exit Sub
Err_Handler:
    Select Case Err
        Case 9: Call ScreenDrawing.MoveAllFramesPVShapeCentre(ShapeCentreX, ShapeCentreY): Exit Sub ' There is no ReferenceShapeShiftObs
        Case Else: MsgBox Err & "-OBS39:" & Error$
    End Select
End Sub
                            
Sub ShiftTheCorrectFrame(ByVal X As Single, ByVal Y As Single)
On Error GoTo Err_Handler
    Dim TotalShift As Integer
    Dim ShiftIndex As Integer
    

    TotalShift = UBound(ReferenceShapeShiftObs)
    If PVFrameNo < ReferenceShapeShiftObs(1).Frame Then
        PVShapeCentreX(1) = X
        PVShapeCentreY(1) = Y
        Exit Sub
    End If
    
    If PVFrameNo >= ReferenceShapeShiftObs(TotalShift).Frame Then
        ReferenceShapeShiftObs(TotalShift).a = X
        ReferenceShapeShiftObs(TotalShift).b = Y
        Exit Sub
    End If
    
    For ShiftIndex = 2 To TotalShift
        If PVFrameNo >= ReferenceShapeShiftObs(ShiftIndex - 1).Frame And _
            PVFrameNo < ReferenceShapeShiftObs(ShiftIndex).Frame Then
            ReferenceShapeShiftObs(ShiftIndex - 1).a = X
            ReferenceShapeShiftObs(ShiftIndex - 1).b = Y
            Exit Sub
        End If
    Next ShiftIndex
    
    

Exit Sub
Err_Handler:
    Select Case Err
        Case 9: PVShapeCentreX(1) = X
                PVShapeCentreY(1) = Y
                Exit Sub
        Case Else: MsgBox Err & "-OBS40:" & Error$
    End Select
End Sub

