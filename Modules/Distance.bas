Attribute VB_Name = "Distance"
Option Explicit

Private Type DistanceChangeNode
    Frame As Long
    Distance As Double
    time As Double
End Type

Sub RecalculateDistance()
On Error GoTo Err_Handler

Dim FileLoadError As Boolean
Dim ObsIndex As Integer
Dim DistanceChange() As DistanceChangeNode
Dim NumberOfDistances

If PVDFileName = "" Or PVDataNoOfLines = 0 Then Exit Sub
If DistanceMethod = "AutomaticCounter" Then Exit Sub 'Required until code ready.
If ConfigInfo.DistanceStart = InvalidData Or ConfigInfo.DistanceFinish = InvalidData Then Exit Sub

'Get Distance Markers
NumberOfDistances = 0
ReDim DistanceChange(0)

DistanceChange(0).Frame = 1
DistanceChange(0).time = PVTimes(1)
DistanceChange(0).Distance = ConfigInfo.DistanceStart

For ObsIndex = 1 To UBound(PipeObservations)
    If Trim(PipeObservations(ObsIndex).PipeObs) = "<<<--I-->>>" Then
        NumberOfDistances = NumberOfDistances + 1
        ReDim Preserve DistanceChange(NumberOfDistances)
        With DistanceChange(NumberOfDistances)
            .Frame = PipeObservations(ObsIndex).PipeObsFrameNo
            .Distance = PipeObservations(ObsIndex).PipeObsDist
            .time = PVTimes(.Frame)
        End With
    End If
Next ObsIndex
 
NumberOfDistances = NumberOfDistances + 1

ReDim Preserve DistanceChange(NumberOfDistances)
DistanceChange(NumberOfDistances).Frame = PVDataNoOfLines
DistanceChange(NumberOfDistances).time = PVTimes(PVDataNoOfLines)
DistanceChange(NumberOfDistances).Distance = ConfigInfo.DistanceFinish

For ObsIndex = 0 To NumberOfDistances - 1
    Call CalculateDistanceChange(DistanceChange(ObsIndex), DistanceChange(ObsIndex + 1))
Next ObsIndex

Call RefreshDistanceDisplays
Call PageFunctions.SaveToFilePipeAndConfigInfo("ConfigInfo", FileLoadError)

Exit Sub
Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & "-D1:" & error$
 
End Select
End Sub

 Sub DistanceAdd(ByVal Dist As Double, ByVal FrameNo As Long)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : DistanceAdd
'Created :
'Updated :
'Prg By  : Antony
'Param   :
'Desc    :
'Usage   :
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim CheckDistValue As Double
Dim FileSaveFail As Boolean

If DistanceMethod = "AutomaticCounter" Then Exit Sub
If FrameNo = 0 Then Exit Sub
If PVDataNoOfLines < 1 Then Exit Sub
If PVDataNoOfLines > PVDataNoOfLines Then Exit Sub


If FrameNo = 1 Then
    ConfigInfo.DistanceStart = Dist
ElseIf FrameNo = PVDataNoOfLines Then
    ConfigInfo.DistanceFinish = Dist
ElseIf ChangeCurrentDistance(Dist, FrameNo) = False Then
If Dist = PVDistances(FrameNo) Then Exit Sub
    If NoOfPipeObservations >= PipeObsBuffer - 4 Then Exit Sub 'Ensure NoOfPipeObservations does not exceed the PipeObsBuffer
    NoOfPipeObservations = NoOfPipeObservations + 1 'Store this new setting in the Pipe Observation array
    ReDim Preserve PipeObservations(NoOfPipeObservations)
    PipeObservations(NoOfPipeObservations).PipeObs = "<<<--I-->>>"
    PipeObservations(NoOfPipeObservations).PipeObsFrameNo = PVFrameNo
    PipeObservations(NoOfPipeObservations).PipeObsDist = Dist
    PipeObservations(NoOfPipeObservations).PipeObsSnapshotLength = 0 ' PCN3576
    PipeObservations(NoOfPipeObservations).PipeObsSnapshotOffset = 0 ' PCN3576
    Call Observations.SortObs
    Call ObsDisplayALL
    Call PageFunctions.SaveToFilePipeObs(FileSaveFail)
End If

Call RecalculateDistance


Exit Sub
Err_Handler:
    MsgBox Err & "-D2:" & error$
End Sub

Private Sub CalculateDistanceChange(ByRef FromChange As DistanceChangeNode, ByRef ToChange As DistanceChangeNode)
On Error GoTo Err_Handler
Dim FromFrame As Long
Dim ToFrame As Long
Dim DistanceMultiply As Double
Dim CurrentFrame As Long


If ToChange.time = FromChange.time Then Exit Sub

FromFrame = FromChange.Frame + 1
ToFrame = ToChange.Frame - 1

DistanceMultiply = (ToChange.Distance - FromChange.Distance) / (ToChange.time - FromChange.time)
For CurrentFrame = FromFrame To ToFrame
    PVDistances(CurrentFrame) = ((PVTimes(CurrentFrame) - FromChange.time) * DistanceMultiply) + FromChange.Distance
Next CurrentFrame

PVDistances(FromChange.Frame) = FromChange.Distance
PVDistances(ToChange.Frame) = ToChange.Distance

Exit Sub
Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & "-D3:" & error$
End Select
End Sub

Sub RefreshDistanceDisplays()
On Error GoTo Err_Handler
Dim ObsIndex As Integer

If PrecisionVisionGraph.PVObsDistance.Visible = True Then
    PrecisionVisionGraph.PVObsDistance.text = Round(PVDistances(PipeObservations(Observations.ObsOpenIndex).PipeObsFrameNo), 2)
End If

Call ScreenDrawing.ClearAllGraphsAndRuler
Call ScreenDrawing.DrawPVGraphs 'PCN3373

If Observations.LastObservationType = "Distance" Then
    For ObsIndex = 1 To UBound(ObsIndexArray)
        PrecisionVisionGraph.ObservationText(ObsIndex).Caption = Round(PVDistances(PipeObservations(ObsIndexArray(ObsIndex)).PipeObsFrameNo), 2)
    Next ObsIndex
End If
Exit Sub
Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & "-D4:" & error$
End Select
End Sub

Function ChangeCurrentDistance(ByVal Dist As Double, ByVal FrameNo As Long) As Boolean
On Error GoTo Err_Handler

Dim NumberObs As Integer
Dim ObsIndex As Integer

NumberObs = UBound(PipeObservations)
ChangeCurrentDistance = False
For ObsIndex = 1 To NumberObs
    If Trim(PipeObservations(ObsIndex).PipeObs) = "<<<--I-->>>" And PipeObservations(ObsIndex).PipeObsFrameNo = FrameNo Then
        PipeObservations(ObsIndex).PipeObsDist = Dist
        ChangeCurrentDistance = True
        Exit For
    End If
Next ObsIndex

Exit Function
Err_Handler:
    MsgBox Err & "-D5:" & error$
    ChangeCurrentDistance = False
End Function


