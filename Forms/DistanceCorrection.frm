VERSION 5.00
Begin VB.Form DistanceCorrection 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   1050
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   1350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1050
   ScaleWidth      =   1350
   Begin VB.CommandButton ReCalculate 
      Caption         =   "Re-Calculate"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox NewDistanceText 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "DistanceCorrection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ReCalculate_Click()
'****************************************************************************************
'Name    : RecalculateDist_Click
'Created : 9 December 2003, PCN2463
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    : Validates data entry, recalculates CameraSpeed values and re-draws the
'           PVGraph y scale.
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
Dim DistanceFinish As Double
Dim TimeStart As Double
Dim TimeFinish As Double
Dim FileLoadError As Boolean
Dim CurrentTime As Double 'PCN2639
Dim CurrentFrame As Long 'PCN2639
Dim FrameFinish As Long 'PCN2639
Dim PVAddressOffset As Long 'PCN2639
Dim PVDataStartAddress As Long 'PCN2639
Dim FileNo As Integer 'PCN2639
Dim FixedPointStart As Integer 'PCN2928 'The closest Fixed Point (PipeObservation index) BEFORE a given PVFrameNo, which may not be the start of the PVD.
Dim FixedPointFinish As Integer 'PCN2928 'The closest Fixed Point (PipeObservation index) AFTER a given PVFrameNo, which may not be the finish of the PVD.
Dim FixedPointStartTime As Double 'PCN2928 'The closest Fixed Point (PipeObservation index) BEFORE a given PVFrameNo, which may not be the start of the PVD.
Dim FixedPointFinishTime As Double 'PCN2928 'The closest Fixed Point (PipeObservation index) AFTER a given PVFrameNo, which may not be the finish of the PVD.
Dim GoBackToFrameNo As Long 'PCN3216


If PVDFileName = "" Or PVDataNoOfLines = 0 Then
    MsgBox DisplayMessage("There is no recorded data to process")
    Exit Sub
End If

If Len(txtStartDistance) = 0 Then 'PCN2639
    DistanceStart = InvalidData 'PCN2639 'PCN3884 was -1
    Exit Sub 'PCN2639
End If

DistanceStart = CDbl(txtStartDistance) 'PCN2511 'PCN2898 removed the function Val eg was CDbl(Val(txtStartDistance)
DistanceFinish = CDbl(txtEndDistance) 'PCN2511'PCN2898 removed the function Val eg was CDbl(Val(txtEndDistance)

If DistanceMethod = "AutomaticCounter" Then Exit Sub 'Required until code ready.

'RecalculateDist.
'DoEvents 'PCN???? Why is this here????
If DistanceStart > InvalidData And DistanceFinish > InvalidData Then   'PCN2511 'PCN3884 was >=0
    GoBackToFrameNo = PVFrameNo 'PCN3216
    CameraSpeedInFrames = (DistanceFinish - DistanceStart) / PVDataNoOfLines
    DistanceMethod = "AvgFrame"
    'vvvv PCN2489 *******************************
    'vvvv PCN2511 *************************
    'Simple validation of the PVTimes
    If PVTimes(1) < PVTimes(2) Then
        TimeStart = PVTimes(1)
    Else
        TimeStart = PVTimes(2) 'PCN2489
    End If
    DistanceStartTime = TimeStart
    If PVTimes(PVDataNoOfLines - 1) < PVTimes(PVDataNoOfLines) Then
'        TimeFinish = PVTimes(PVDataNoOfLines)
        FrameFinish = PVDataNoOfLines 'PCN2639
    Else
'        TimeFinish = PVTimes(PVDataNoOfLines - 1)
        FrameFinish = PVDataNoOfLines - 1 'PCN2639
    End If
    TimeFinish = PVTimes(FrameFinish) 'PCN2639
    '^^^^ ***************************************
    If TimeStart < TimeFinish And TimeFinish <> 0 Then
        CameraSpeedInTime = (DistanceFinish - DistanceStart) / (TimeFinish - TimeStart)
        DistanceMethod = "AvgTime"
    End If
    '^^^^ ***************************************
    'vvvv PCN2928 *********************************
    If NoOfPipeObservations > 0 Then
        'Setup DistanceCounterFixedPt array
        Call DistanceFixedPtSetup(DistanceFinish)
        'Reset marker points
        FixedPointStart = 1
        FixedPointFinish = FixedPointStart
        'Reset  points times
        FixedPointStartTime = TimeStart
        FixedPointFinishTime = FixedPointStartTime
    End If
    '^^^^ *****************************************
    'vvvv PCN2639 ********************************
    'Find PVDataStartAddress
    Call GetPVDPointerPVDataFromFile(PVDFileName, PVDataStartAddress, FileLoadError)
    FileNo = 8
    Open PVDFileName For Binary Access Write As #FileNo
    
    'Populate the PVDistances array - this may take some time.
    For CurrentFrame = 1 To FrameFinish 'PCN2928
        If DistanceMethod = "AvgTime" Then
            CurrentTime = PVTimes(CurrentFrame)
            'vvvv PCN2928 *********************************
            'Include Fixed Distance Points
            If NoOfPipeObservations = 0 Then
                'There are no Fixed Distance Points
                PVDistances(CurrentFrame) = Round(CameraSpeedInTime * (CDbl(CurrentTime) - DistanceStartTime) + DistanceStart, 2)
            Else
                Call DistanceFixedPtPVDistanceCalc(CurrentFrame, FixedPointStart, FixedPointFinish, FrameFinish, FixedPointStartTime, FixedPointFinishTime, TimeFinish, CurrentTime)
            End If
            '^^^^ *****************************************
        ElseIf DistanceMethod = "AvgFrame" Then
            PVDistances(CurrentFrame) = Round(CameraSpeedInFrames * CDbl(CurrentFrame) + DistanceStart, 2)
        End If
'        If CurrentFrame > 150 Then
'            Debug.Print
'        End If
        If Not FileLoadError And ConfigInfo.PVDFileVersion <> "V4.0" And ConfigInfo.PVDFileVersion <> "V4.1" Then
            Call RapidSavePVDistanceToFile(CurrentFrame, PVDataStartAddress, FileLoadError, FileNo)
        End If
    Next CurrentFrame
    Close #FileNo
    '^^^^ ****************************************
    'Setup PrecisionVisionGraph form
    PrecisionVisionGraph.PDisplayStartDist = Format(DistanceStart, "###0.0")
    PrecisionVisionGraph.PDisplayFinishDist = Format(DistanceFinish, "###0.0")
    PrecisionVisionGraph.PDisplayFinishDist.Visible = True
    If ConfigInfo.Units = "mm" Then
        PrecisionVisionGraph.Y_Units.Caption = "   m"
    Else
        PrecisionVisionGraph.Y_Units.Caption = "   ft"
    End If
    'Redraw PVScale Y-Axis
    'vvvv PCN2971 *********************************************
'    PrecisionVisionGraph.PVYScale(0).Cls 'PCN2970
'    PrecisionVisionGraph.PVYScale(1).Cls 'PCN2970
    'Call PrecisionVisionGraph.PVYScaleClearAll
    
    'PCN3216 overides pcn2971, do not send back to top of pvgraph. Stay at current frame
    Call ScreenDrawing.ClearAllGraphsAndRuler
    
'''' PCN3441 (6 April 2005, Antony van Iersel)
''
''    Call DrawPVYScaleNumbersAndMarkers(PrecisionVisionGraph.PVYScale(0), 1, 0)
    '^^^^ *****************************************************
    DoEvents
    CurrentPVYScalePageNo = 1
    'vvvv PCN2478 ********************************
'    PipeDisplayMoveLastY = PrecisionVisionGraph.ViewIndicator.Item(0).Top
'    Call PrecisionVisionGraph.PipeDisplay_MouseMove(1, 0, 25, 25) 'Move viewindicator to top
    '^^^^ ****************************************
    MsgBox DisplayMessage("Calculation complete") 'PCN2463 'PCN2639
    
    'vvvv PCN2639 **********************************************
    If DistanceMethod = "None" Then
        ConfigInfo.DistanceProcessMethod = "None"
        ConfigInfo.DistanceStart = InvalidData 'PCN3884 was -1
        ConfigInfo.DistanceDirection = ""
        ConfigInfo.DistanceFinish = InvalidData 'PCN3884 was -1
    Else
        ConfigInfo.DistanceProcessMethod = DistanceMethod
        ConfigInfo.DistanceStart = DistanceStart
        ConfigInfo.DistanceDirection = CountDirection
        ConfigInfo.DistanceFinish = DistanceFinish
    End If
    'Save config to PVD
    Call SaveToFilePipeAndConfigInfo("ConfigInfo", FileLoadError)
    '^^^^ ******************************************************
    
    'PCN4131
    'vvvv PCN2928 **********************************************
    'If isopen("Observations") And NoOfPipeObservations > 0 Then
    '    Call Observations.ObservationsListLoad
    'End If
    '^^^^ ******************************************************
    
    'vvvv PCN3216 **********************************************
    PVFrameNo = GoBackToFrameNo
    Call ClearLineScreen.GotoPVProfile(PVFrameNo, True)
    '^^^^ ******************************************************
    
ElseIf DistanceStart > InvalidData Then 'PCN3884 was -1
    MsgBox "'" & lblStartDistance.Caption & "' " & DisplayMessage("must be less than") & " '" & lblFinishDistance.Caption & "'"
Else
    MsgBox DisplayMessage("Please input valid data") 'PCN2639
End If

TidyUp:
    Close #FileNo

Exit Sub
Err_Handler:
Select Case Err
    Case 13 'Invalid data
        DistanceStart = InvalidData 'PCN3884 was -1
        Resume Next
    Case Else
        MsgBox Err & "-" & error$
        GoTo TidyUp
End Select
End Sub
