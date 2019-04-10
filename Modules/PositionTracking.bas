Attribute VB_Name = "TrackingPosition"
Option Explicit

Public CompassVersion As Single
Public DrawingCompassResolution As Integer
Public DrawingLabelResolution As Integer


Public Type TrakingPositionType_V1_0       '
    Directions(4) As Integer
    Time As Double
End Type

Public Type CompassData_V1_0
    Direction As Single
    Time As Double
    X As Single
    Y As Single
    Distance As Single
End Type


Public CompassData() As CompassData_V1_0
Public NoCompassReadings As Long
Public CompassCentreX As Single
Public CompassCentreY As Single
Public CompassScale As Single

Public MouseButton2Down
Public MouseButton1Down
Public MouseX
Public MouseY

Public StartDrawCompassData As Long
Public EndDrawCompassData As Long

Public RawTrackingData() As TrakingPositionType_V1_0

Sub LoadAndProcessCompass()
On Error GoTo Err_Handler
    Dim FileName As String
    Dim PathName As String
    Dim ExtName As String
    Dim CompassFileName As String
    
    
    CompassFileName = PVDFileName
    Call SplitFilePath(CompassFileName, PathName, FileName, ExtName)
    CompassFileName = PathName & FileName & ".3c3"
    If Dir(CompassFileName) = "" Then Exit Sub
    
    ClearLineScreen.MainScreen3D.ZOrder 0
    ClearLineScreen.MainScreen3D.BackColor = vbWhite
    ClearLineScreen.MainScreen3D.Visible = True
    ClearLineScreen.MainScreen3D.Left = 0
    ClearLineScreen.MainScreen3D.Top = 0

    Call LoadData(CompassFileName)
    Call ProcessData
    Call PopulateDistance
    Call CalculatePoints
    
    CompassCentreX = 0
    CompassCentreY = 0
    CompassScale = 2
    DrawingCompassResolution = 1
    DrawingLabelResolution = 500
    
    
    Call DisplayCompassData
    
Exit Sub
Err_Handler:
Select Case Err
    Case Else: MsgBox Err & "-T1:" & Error$
End Select
End Sub

Sub LoadData(ByVal FileName As String)
On Error GoTo Err_Handler

    Dim FileNo
    Dim InputString As String
    Dim NumberLines As Long
    Dim ParsedString As String
    
    FileNo = FreeFile
    Open FileName For Input As #FileNo
    
    Line Input #FileNo, InputString
    CompassVersion = SafeCDbl(Right(InputString, Len(InputString) - 1)) 'PCN4161
    
    Do While Not EOF(FileNo)   ' Check for end of file.
        Line Input #FileNo, InputString ' Read line of data.
        NumberLines = NumberLines + 1
        On Error GoTo NextLine
        ReDim Preserve RawTrackingData(NumberLines)
        RawTrackingData(NumberLines).Directions(1) = CInt(Left(InputString, 1))
        ParsedString = Mid(InputString, 3, 1): RawTrackingData(NumberLines).Directions(2) = ParsedString
        ParsedString = Mid(InputString, 5, 1): RawTrackingData(NumberLines).Directions(3) = ParsedString
        ParsedString = Mid(InputString, 7, 1): RawTrackingData(NumberLines).Directions(4) = ParsedString
        ParsedString = Mid(InputString, 9): RawTrackingData(NumberLines).Time = SafeCDbl(ParsedString) / 80000 'PCN4161
ResumeNextLine:
    On Error GoTo Err_Handler
    Loop

    Close #FileNo
Exit Sub

NextLine:
    NumberLines = NumberLines - 1
GoTo ResumeNextLine
Err_Handler:
Select Case Err
    
    Case Else: MsgBox Err & "-T2:" & Error$: Close #FileNo
End Select
End Sub

Sub ProcessData()
On Error GoTo Err_Handler


Dim StartDirection As Integer
Dim i As Long
Dim AverageDirection As Single
Dim CentreDirection As Single
Dim DistanceRatio As Single


Dim DirectionSlice(4) As Single

NoCompassReadings = UBound(RawTrackingData)
StartDrawCompassData = 1: DebugForm.StartLine = 1
EndDrawCompassData = NoCompassReadings: DebugForm.EndLine = NoCompassReadings
DistanceRatio = 700 / NoCompassReadings




ReDim CompassData(NoCompassReadings)

For i = 1 To NoCompassReadings
    If i = 44991 Then
        i = i
    End If

    With RawTrackingData(i)
    
        DirectionSlice(1) = ((.Directions(1) - 1) * 45)
        DirectionSlice(2) = (((.Directions(2) - 1) * 45) - 11.25)
        DirectionSlice(3) = (((.Directions(3) - 1) * 45) - 22.5)
        DirectionSlice(4) = (((.Directions(4) - 1) * 45) - 33.75)
    
    End With
        
    If DirectionSlice(4) < 0 Then DirectionSlice(4) = DirectionSlice(4) + 360
    
    
    If (DirectionSlice(3) - DirectionSlice(2)) > 180 Then DirectionSlice(3) = DirectionSlice(3) - 360
    If (DirectionSlice(2) - DirectionSlice(3)) > 180 Then DirectionSlice(2) = DirectionSlice(2) - 360
           
    CentreDirection = (DirectionSlice(2) + DirectionSlice(3)) / 2
   
    If (CentreDirection - DirectionSlice(1)) > 180 Then DirectionSlice(1) = DirectionSlice(1) + 360
    If (DirectionSlice(4) - CentreDirection) > 180 Then DirectionSlice(4) = DirectionSlice(4) - 360

    AverageDirection = DirectionSlice(1) + DirectionSlice(2) + DirectionSlice(3) '+ DirectionSlice(4)
    AverageDirection = AverageDirection / 3
    
    If AverageDirection < 0 Then AverageDirection = AverageDirection + 360
    If AverageDirection >= 360 Then AverageDirection = AverageDirection - 360
    CompassData(i).Direction = AverageDirection
    CompassData(i).Time = RawTrackingData(i).Time
    CompassData(i).Distance = i * DistanceRatio
    
Next i

Exit Sub
Err_Handler:
Select Case Err
    Case Else: MsgBox Err & "-T3:" & Error$
End Select
End Sub

Sub DisplayCompassData()
On Error GoTo Err_Handler

Dim PointX As Single
Dim PointY As Single
Dim PointBX As Single
Dim PointBY As Single
Dim LabelWidth As Single
Dim LabelHeight As Single
Dim DistanceString As String
Dim i As Long


    Dim LabelDirection As Boolean
    If CompassScale = 0 Then Exit Sub
    
    ClearLineScreen.MainScreen3D.Cls

    ClearLineScreen.MainScreen3D.CurrentX = CentreLineX + CompassCentreX
    ClearLineScreen.MainScreen3D.CurrentY = CentreLineY + CompassCentreY
           
    For i = StartDrawCompassData To EndDrawCompassData Step DrawingCompassResolution
        PointX = CentreLineX + CompassCentreX + (CompassData(i).X * CompassScale)
        PointY = CentreLineY + CompassCentreY + (CompassData(i).Y * CompassScale)
        ClearLineScreen.MainScreen3D.Line -(PointX, PointY)
        DistanceString = CStr(Round(CompassData(i).Distance, 2))
        LabelWidth = ClearLineScreen.MainScreen3D.TextWidth(DistanceString)
        LabelHeight = ClearLineScreen.MainScreen3D.TextHeight(DistanceString)
        
    

        
        If DrawingLabelResolution / CompassScale > 1 Then
            If (i - 1) Mod (DrawingLabelResolution / CompassScale) = 0 Then
                If LabelDirection Then
                    ClearLineScreen.MainScreen3D.Line -(PointX + 60, PointY)
                    ClearLineScreen.MainScreen3D.CurrentX = PointX + 60
                    ClearLineScreen.MainScreen3D.CurrentY = PointY - (LabelHeight / 2)
                    ClearLineScreen.MainScreen3D.Print DistanceString
                    'ClearLineScreen.MainScreen3D.Print DistanceString
    
                Else
                    ClearLineScreen.MainScreen3D.Line -(PointX - 60, PointY)
                    ClearLineScreen.MainScreen3D.CurrentX = PointX - 60 - LabelWidth
                    ClearLineScreen.MainScreen3D.CurrentY = PointY - (LabelHeight / 2)
                    ClearLineScreen.MainScreen3D.Print DistanceString
                    'ClearLineScreen.MainScreen3D.Print DistanceString
                End If
                LabelDirection = Not LabelDirection
            End If
        End If
        ClearLineScreen.MainScreen3D.CurrentX = PointX
        ClearLineScreen.MainScreen3D.CurrentY = PointY
    Next i

Exit Sub
Err_Handler:
Select Case Err
    Case 6: Exit Sub
    Case Else: MsgBox Err & "-T4:" & Error$
End Select

End Sub

Sub MouseMove(ByVal X As Single, ByVal Y As Single)
On Error GoTo Err_Handler

    If MouseButton1Down Then
        CompassCentreX = CompassCentreX - (MouseX - X)
        CompassCentreY = CompassCentreY - (MouseY - Y)
        Call DisplayCompassData
        MouseX = X
        MouseY = Y
    End If
    If MouseButton2Down Then
        CompassCentreX = CompassCentreX / CompassScale
        CompassCentreY = CompassCentreY / CompassScale

        
        CompassScale = CompassScale - ((MouseY - Y) / 2)
        If CompassScale < 0 Then CompassScale = CompassScale + ((MouseY - Y) / 2)
        
        CompassCentreX = CompassCentreX * CompassScale
        CompassCentreY = CompassCentreY * CompassScale
        
        Call DisplayCompassData
        MouseX = X
        MouseY = Y
    End If
    
    Exit Sub
Err_Handler:
Select Case Err
    Case Else: MsgBox Err & "-T5:" & Error$
End Select
End Sub

Sub CalculatePoints()
On Error GoTo Err_Handler
    Dim PointAX As Single
    Dim PointAY As Single
    Dim PointBX As Single
    Dim PointBY As Single
    Dim i As Long
    
    


    PointAX = 0
    PointAY = 0
    
           
    
    
    For i = 1 To NoCompassReadings
        PointBX = Sin(CompassData(i).Direction / 180 * PI) * (CompassData(i).Distance - CompassData(i - 1).Distance)
        PointBY = Cos(CompassData(i).Direction / 180 * PI) * (CompassData(i).Distance - CompassData(i - 1).Distance)
        PointBX = PointBX + PointAX
        PointBY = PointBY + PointAY
        CompassData(i).X = PointBX
        CompassData(i).Y = PointBY
        
        PointAX = PointBX
        PointAY = PointBY
    Next i
Exit Sub
Err_Handler:
Select Case Err
    Case Else: MsgBox Err & "-T6:" & Error$
End Select
End Sub

Sub PopulateDistance()
On Error GoTo Err_Handler

    Dim i As Long
    For i = 1 To NoCompassReadings
        CompassData(i).Distance = GetDistanceFromTime(CompassData(i).Time)
                                  
    Next i
    
    
    

Exit Sub
Err_Handler:
Select Case Err
    Case Else: MsgBox Err & "-T7:" & Error$
End Select
End Sub

Function GetDistanceFromTime(ByVal TheTime As Double) As Single
On Error GoTo Err_Handler
    Dim Distance1 As Single
    Dim Time1 As Double
    Dim Distance2 As Single
    Dim Time2 As Single
    Dim FrameNo As Long
    Dim DistToTimeRatio As Double
    
    FrameNo = GetPVFrameNoFromTime(TheTime)
    If FrameNo >= PVDataNoOfLines Then GetDistanceFromTime = PVDistances(FrameNo): Exit Function
    If (PVTimes(FrameNo) = TheTime) Then GetDistanceFromTime = PVDistances(FrameNo): Exit Function
    
    
    If (PVTimes(FrameNo) < TheTime) Then
        Time1 = PVTimes(FrameNo):     Distance1 = PVDistances(FrameNo)
        Time2 = PVTimes(FrameNo + 1): Distance2 = PVDistances(FrameNo + 1)
    Else
        Time2 = PVTimes(FrameNo):     Distance1 = PVDistances(FrameNo)
        Time1 = PVTimes(FrameNo + 1): Distance2 = PVDistances(FrameNo + 1)
    End If
    
    DistToTimeRatio = (TheTime - Time1) / (Time2 - Time1)
    GetDistanceFromTime = ((Distance2 - Distance1) * DistToTimeRatio) + Distance1
Exit Function
Err_Handler:
Select Case Err
    Case Else: MsgBox Err & "-T8:" & Error$
End Select
End Function
