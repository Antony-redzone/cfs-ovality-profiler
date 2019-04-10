Attribute VB_Name = "ScreenDrawing"
Option Explicit

Private Declare Sub clearline_CalculateFlat3d Lib "clearline.dll" (ByRef pvDataX As Single, _
                                                              ByRef pvDataY As Single, _
                                                              ByRef pvCentreX As Single, _
                                                              ByRef pvCentreY As Single, _
                                                              ByRef EgnoreList As Long, _
                                                              ByVal PVDataXYMultiplier As Long, _
                                                              ByRef pvColourRed As Long, _
                                                              ByRef pvColourGreen As Long, _
                                                              ByRef pvColourBlue As Long, _
                                                              ByVal PVExptededDia As Double, _
                                                              ByVal pvRedLimit As Double, _
                                                              ByVal pvBlueLimit As Double, _
                                                              ByVal FromFrame As Long, _
                                                              ByVal ToFrame As Long, _
                                                              ByVal OffsetX As Double, _
                                                              ByVal offsetY As Double, _
                                                              ByRef Shape As ReferenceShape_V10, _
                                                              ByVal ShapeRotation As Double) 'PCN3513 'PCN3567 offsets x and y added

Private Declare Sub clearline_CalculateCapacity Lib "clearline.dll" (ByRef pvDataX As Single, _
                                                              ByRef pvDataY As Single, _
                                                              ByRef pvCentreX As Single, _
                                                              ByRef pvCentreY As Single, _
                                                              ByRef EgnoreList As Long, _
                                                              ByRef PVCapacityData As Single, _
                                                              ByVal PVDataXYMultiplier As Long, _
                                                              ByVal PVExptededDia As Double, _
                                                              ByVal FromFrame As Long, _
                                                              ByVal ToFrame As Long) 'PCN3513
                                                              
Private Declare Sub clearline_CalculateOvality Lib "clearline.dll" (ByRef pvDataX As Single, _
                                                              ByRef pvDataY As Single, _
                                                              ByRef pvCentreX As Single, _
                                                              ByRef pvCentreY As Single, _
                                                              ByRef EgnoreList As Long, _
                                                              ByRef PVOvalityData As Single, _
                                                              ByVal PVDataXYMultiplier As Long, _
                                                              ByVal FromFrame As Long, _
                                                              ByVal ToFrame As Long) 'PCN3513
                                                              
Private Declare Sub clearline_CalculateDebugOvality Lib "clearline.dll" (ByRef pvDataX As Single, _
                                                              ByRef pvDataY As Single, _
                                                              ByRef pvCentreX As Single, _
                                                              ByRef pvCentreY As Single, _
                                                              ByRef EgnoreList As Long, _
                                                              ByRef PVOvalityData As Single, _
                                                              ByVal PVDataXYMultiplier As Long, _
                                                              ByVal Frame As Long, _
                                                              ByVal hwnd As Long, _
                                                              ByVal width As Single, _
                                                              ByVal height As Single, _
                                                              ByVal Ratio As Double) 'PCN3513
                                                              
Private Declare Sub clearline_CalculateXYDiameter Lib "clearline.dll" (ByRef pvDataX As Single, _
                                                              ByRef pvDataY As Single, _
                                                              ByRef pvCentreX As Single, _
                                                              ByRef pvCentreY As Single, _
                                                              ByRef EgnoreList As Long, _
                                                              ByRef PVXDiamterData As Double, _
                                                              ByRef PVYDiamterData As Double, _
                                                              ByVal PVDataXYMultiplier As Long, _
                                                              ByVal FromFrame As Long, _
                                                              ByVal ToFrame As Long) 'PCN3513
                                                              
Private Declare Sub clearline_CalculateDeltaMaxMin Lib "clearline.dll" (ByRef pvDataX As Single, _
                                                              ByRef pvDataY As Single, _
                                                              ByRef pvCentreX As Single, _
                                                              ByRef pvCentreY As Single, _
                                                              ByRef EgnoreList As Long, _
                                                              ByRef PVDeltaFullMax As Double, _
                                                              ByRef PVDeltaFullMin As Double, _
                                                              ByRef PVDeltaSegFullMax As Integer, _
                                                              ByRef PVDeltaSegFullMin As Integer, _
                                                              ByVal PVDataXYMultiplier As Long, _
                                                              ByVal FromFrame As Long, _
                                                              ByVal ToFrame As Long) 'PCN3513

Private Declare Sub clearline_CalculateDiameterMaxMin Lib "clearline.dll" (ByRef pvDataX As Single, _
                                                              ByRef pvDataY As Single, _
                                                              ByRef pvCentreX As Single, _
                                                              ByRef pvCentreY As Single, _
                                                              ByRef EgnoreList As Long, _
                                                              ByRef PVDiamterFullMax As Double, _
                                                              ByRef PVDiamterFullMin As Double, _
                                                              ByRef PVDiamterSegFullMax As Integer, _
                                                              ByRef PVDiamterSegFullMin As Integer, _
                                                              ByVal PVDataXYMultiplier As Long, _
                                                              ByVal FromFrame As Long, _
                                                              ByVal ToFrame As Long) 'PCN3513
                                                              
Private Declare Sub clearline_CalculateDiameterMedian Lib "clearline.dll" (ByRef pvDataX As Single, _
                                                              ByRef pvDataY As Single, _
                                                              ByRef pvCentreX As Single, _
                                                              ByRef pvCentreY As Single, _
                                                              ByRef EgnoreList As Long, _
                                                              ByRef PVDiameterMedian As Double, _
                                                              ByVal PVDataXYMultiplier As Long, _
                                                              ByVal FromFrame As Long, _
                                                              ByVal ToFrame As Long) 'PCN3513
                                                              
'Private Declare Sub clearline_CalculateFractile Lib "clearline.dll" (ByRef pvDataX As Single, _
'                                                              ByRef pvDataY As Single, _
'                                                              ByRef pvCentreX As Single, _
'                                                              ByRef pvCentreY As Single, _
'                                                              ByRef EgnoreList As Long, _
'                                                              ByRef PVFractile As Single, _
'                                                              ByVal PVDataXYMultiplier As Long, _
'                                                              ByVal FromFrame As Long, _
'                                                              ByVal ToFrame As Long) 'PCN3513
                                                              
Private Declare Sub Hough_ProcessSingleImage Lib "laserlib.dll" (ByRef ImagePointer As Byte, _
                                                                 ByVal imageWidth As Long, _
                                                                 ByVal imageHeight As Long) 'PCN3194
                                                                 
Private Declare Sub Clearline_CalculateCentre Lib "clearline.dll" Alias "clearline_CalculateCentre" (ByRef pvDataX As Single, _
                                                                ByRef pvDataY As Single, _
                                                                ByRef pvCentreX As Single, _
                                                                ByRef pvCentreY As Single, _
                                                                ByRef EgnoreList As Long, _
                                                                ByVal FromFrame As Long, _
                                                                ByVal ToFrame As Long) 'PCN3513
                                                              
Private Declare Sub clearline_RefShapeDistCalc Lib "clearline.dll" (ByRef Shape As ReferenceShape_V10, _
                                                                   ByVal X As Single, _
                                                                   ByVal Y As Single, _
                                                                   ByRef OrthoX As Double, _
                                                                   ByRef OrthoY As Double, _
                                                                   ByRef OrthoDistance As Double, _
                                                                   ByVal ShapeRadius As Double, _
                                                                   ByVal ShapeCentreX As Double, _
                                                                   ByVal ShapeCentreY As Double, _
                                                                   ByVal ShapeRotation As Double)
                                                                   
Private Declare Sub clearline_AutoRotate Lib "clearline.dll" (ByRef Shape As ReferenceShape_V10, _
                                                              ByVal ShapeRadius As Double, _
                                                              ByVal ShapeCentreX As Double, _
                                                              ByVal ShapeCentreY As Double, _
                                                              ByVal ShapeRotation As Double, _
                                                              ByRef pvDataX As Single, _
                                                              ByRef pvDataY As Single, _
                                                              ByRef pvCentreX As Single, _
                                                              ByRef pvCentreY As Single, _
                                                              ByVal FromFrame As Long, _
                                                              ByVal ToFrame As Long, _
                                                              ByRef EgnoreList As Long, _
                                                              ByVal hwnd As Long, _
                                                              ByVal width As Single, _
                                                              ByVal height As Single, _
                                                              ByVal Ratio As Double)
                                                                   
Private Declare Sub clearLine_SmoothOutGraphSingle Lib "clearline.dll" (ByRef DataGraph As Single, ByVal numberFrames As Long)
Private Declare Sub setrectanglecoord Lib "laserlib.dll" (ByVal XBottom As Single, ByVal YBottom As Single, ByVal XTop As Single, ByVal YTop As Single)
Private Declare Sub hough_clearrectanglecoord Lib "laserlib.dll" ()

                                                                 
Public Count As Integer
'Public OvalityMax(2000) As Double  'PCN???? Testing
'Public OvalityMean(2000) As Double 'PCN????
'Public OvalityValue(2000) As Double 'PCN????
'Public RawPVData(2000, 180) As Double 'PCN????

' X4Circle2(1) these were the centres of the
' Y4Circle2(1) reference circles. Now ShapeCentreX and Y
'PCN3055 (10 December 2004, Antony) '''''''
Dim NormaliseType As String ' Normalise calculation either for "Graph" or "Drawing"
Public ShapeCentreX As Double  ' Centre of the the shape for Normalise 'PCN3567 made public
Public ShapeCentreY As Double  ' calculations for either X or Y coordinate 'PCN3567 made public
Public DrawingCentreX As Double ' Relative drawing centre for drawing all objects on the screen, this changes
Public DrawingCentreY As Double ' between PV Screen and video screen. This is to keep the drawing lined up with the video
Dim ShapeRadius As Double   ' Raidus of shape.
Public ShapeRotationAngle As Double 'Rotation of shapes at screen centre coordinates in Radians
Public ShapeRotationAngleLast As Double 'Store the current of rotation needed to stop the angle reseting to zero
Public ShowReferenceShape As Boolean 'Turns the Shape off or on at will
Public ShowOutsideShape As Boolean ' Turns the external diameter off or on at will

'PCN3373 All the extra bits and pieces to keep track of the
'Graphs, GraphState was original concept to know what stae
'the graph was on its last draw so that it could be erased
'''''''''''''''''''''''''''''''''''''''''
                                        '
Public Type ImageGraphStateType         '
    XScale As Double                    '
    LeftLimitLine As Double             '
    RightLimitLine As Double            '
    CentreOffset As Double              '
    Left As Double                      '
    Right As Double                     '
    PreviousStartFrame As Long          '
    PreviousEndFrame As Long            '
    PreviousGraphType As String         '
    PictureImage As Image               '
    PreviouspSpeed As Long              '
    PreviousUnits As String             '
    GraphType As String                 '
End Type                                '
'''''''''''''''''''''''''''''''''''''''''
                                                    
'''''''''''''''''''''''''''''''''''''
Public Type GraphContainerType      '
    GraphType As String             '
    PVXScaleUnits As String
End Type                            '
'''''''''''''''''''''''''''''''''''''
                                                 
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Public ImageGraphState(6) As ImageGraphStateType    '
Public ImageRulerState As ImageGraphStateType       '
Public GraphInfoContainer(7) As GraphContainerType      '    'PCN4171                                                '
''Public PVGraphOrder(6) As String  'PCN4171
''Public PVXScaleUnits(6) As String 'PCN2990 'PCN4171
'''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public ReportDummyGraphImage As Image 'PCN3691          '
Public ReportGraphImageX As Image 'PCN3691              '
Public ReportGraphImageY As Image 'PCN3691              '
Public PrintPreviewForm As Form 'PCN3691                '
                                                        '
'PCN3691 It cant be a string because a concern of speed, 0 for standard graphs, 1 for reports
Public ScreenDrawingType As Integer                     '
Public ScreenDrawingOrientation As Integer              ' 'PCN3691
Public DrawSF As Double 'Draw All scale factor          'PCN1835 'Moved to globel for PCN3691
Public PrinterClipOn As Boolean                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim DiameterForMedian() As Variant 'PCN3489

' 'PCN2990 to make a form appear ontop all times.
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1


'PCN4046
Public DrawingActionLine As ShapeLine_V10
Public DrawingMaskBox() As ShapeLine_V10
Public OrthogonalLine As ShapeLine_V10
Public CalibrationLine As ShapeLine_V10
Public WaterLevelLine(4) As ShapeLine_V10
Public DrawingSingleLines() As ShapeLine_V10
Public DrawingPolyLines() As ShapePolyLine_V10
Public DrawingIntegratedArea() As ShapeLine_V10

'vvvv PCN4171 *************************
Public zoomspeed As Integer
Public laserspeed As Integer
Public D3D_Control As String
Public D3D_MapOnOff As Boolean
Public D3D_WaterOnOff As Boolean
Public D3D_Scale As Long
Public D3D_MouseX2 As Long
Public D3D_MouseY2 As Long
Public D3D_CameraView As String
'^^^^ ********************************
Public PolyLineDistance As Double 'PCN4171




Public Sub FormMakeNormal(hwnd As Long)
On Error GoTo Err_Handler:
    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Exit Sub
Err_Handler:
    Select Case Err
        Case Else
            MsgBox Err & " - " & error$
    End Select
End Sub
Public Sub FormTopMost(hwnd As Long)
On Error GoTo Err_Handler:
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Exit Sub
Err_Handler:
    Select Case Err
        Case Else
            MsgBox Err & " - " & error$
    End Select
End Sub


Sub DrawAll(screenctrl As Control)
On Error GoTo Err_Handler
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else
            MsgBox Err & " - " & error$
    End Select
End Sub




Sub DrawInternalDiameter(InternalExteralPipe)
On Error GoTo Err_Handler

Exit Sub
Err_Handler:
    MsgBox Err & " - " & error$
End Sub





Sub DrawPVProfile(PicCntrlName As Control, FrameNo As Integer)    'PCNGL231202 PCN3526
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'DrawPVProfile Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    27/11/02     Building initial framework
'           GL              23/12/02    Added ability for fn to be used to
'                                       draw PVScreen or PVPicInPic picture boxes
'
'Description:
'   Reads the contents of the Prevision Vision profiler data array and displays
'   the profile within the PVScreen.
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler


If Len(PipelineDetails.InternalDiameterExpected.text) = 0 Then Exit Sub
If FrameNo > PVDataNoOfLines Then Exit Sub
If CLPScreenMode = SnapShot Or CLPScreenMode = Video Then Exit Sub

    Call DrawPVProfile_XY(PicCntrlName, FrameNo) 'PCN3526                   '
    If ImageGraphState(0).GraphType = "Delta" Then                           '
        Call PVDeltaCalcShowMaxMin(1)                                       '
    End If                                                                  '

Exit Sub
Err_Handler:
    Select Case Err
        Case 9
            Exit Sub
        Case Else
            MsgBox Err & error$
    End Select
End Sub



Sub DrawCurve(screenctrl As Control, X1, Y1, X2, Y2, x3, y3, colr)
On Error GoTo Err_Handler
Err_Handler:
    MsgBox Err & " - " & error$
End Sub

Sub DrawYAxisNumber_Twips(CtrlName As Control, No, X1, Y1) 'PCNGL160103
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'DrawYAxisNumber_Twips Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    04/12/02     Building initial framework
'   V0.1    Geoff           16/01/03    Renamed function to DrawYAxisNumber_Twips
'
'Description:
'       This function has been designed to be used for drawing the numbers on the
'       y-axis scales. However, this function could be used where ever vertical
'       numbers are required.
'Purpose:
'       CtrlName is the name of the picturebox control
'       No is the number required
'       x1 and y1 are the top left start co-ordinates of the number.
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim XBit As Integer
Dim YBit As Integer
Dim BitSize As Integer

BitSize = 15 'Approximate size of pixel in terms of twips

If No < 0 Then Exit Sub      'PCN2511

For XBit = 0 To 7
    For YBit = 0 To 4
        If NumberPic(No, XBit, YBit) = 1 Then
            'CtrlName.PSet (X1 + XBit, Y1 + YBit) 'PCNGL091202
            CtrlName.Line (X1 + BitSize * XBit, Y1 + BitSize * YBit)-(X1 + BitSize * (XBit + 1), Y1 + BitSize * (YBit + 1)), 0 'PCNGL210103 Ensured colour of numbers is black
        End If
    Next YBit
Next XBit

Exit Sub
Err_Handler:
    MsgBox Err & error$
End Sub


Sub InitialiseNumberPicArray()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'InitialiseNumberPicArray Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    04/12/02     Building initial framework
'
'Description:
'
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

'The X Scale
'Load number '9'
NumberPic(9, 0, 0) = 0
NumberPic(9, 1, 0) = 1
NumberPic(9, 2, 0) = 1
NumberPic(9, 3, 0) = 1
NumberPic(9, 4, 0) = 1
NumberPic(9, 5, 0) = 1
NumberPic(9, 6, 0) = 1
NumberPic(9, 7, 0) = 0
NumberPic(9, 0, 1) = 1
NumberPic(9, 1, 1) = 0
NumberPic(9, 2, 1) = 0
NumberPic(9, 3, 1) = 0
NumberPic(9, 4, 1) = 1
NumberPic(9, 5, 1) = 0
NumberPic(9, 6, 1) = 0
NumberPic(9, 7, 1) = 1
NumberPic(9, 0, 2) = 1
NumberPic(9, 1, 2) = 0
NumberPic(9, 2, 2) = 0
NumberPic(9, 3, 2) = 0
NumberPic(9, 4, 2) = 1
NumberPic(9, 5, 2) = 0
NumberPic(9, 6, 2) = 0
NumberPic(9, 7, 2) = 1
NumberPic(9, 0, 3) = 1
NumberPic(9, 1, 3) = 0
NumberPic(9, 2, 3) = 0
NumberPic(9, 3, 3) = 0
NumberPic(9, 4, 3) = 1
NumberPic(9, 5, 3) = 0
NumberPic(9, 6, 3) = 0
NumberPic(9, 7, 3) = 1
NumberPic(9, 0, 4) = 0
NumberPic(9, 1, 4) = 1
NumberPic(9, 2, 4) = 1
NumberPic(9, 3, 4) = 1
NumberPic(9, 4, 4) = 0
NumberPic(9, 5, 4) = 0
NumberPic(9, 6, 4) = 1
NumberPic(9, 7, 4) = 0

'Load number '8'
NumberPic(8, 0, 0) = 0
NumberPic(8, 1, 0) = 1
NumberPic(8, 2, 0) = 1
NumberPic(8, 3, 0) = 0
NumberPic(8, 4, 0) = 1
NumberPic(8, 5, 0) = 1
NumberPic(8, 6, 0) = 1
NumberPic(8, 7, 0) = 0

NumberPic(8, 0, 1) = 1
NumberPic(8, 1, 1) = 0
NumberPic(8, 2, 1) = 0
NumberPic(8, 3, 1) = 1
NumberPic(8, 4, 1) = 0
NumberPic(8, 5, 1) = 0
NumberPic(8, 6, 1) = 0
NumberPic(8, 7, 1) = 1

NumberPic(8, 0, 2) = 1
NumberPic(8, 1, 2) = 0
NumberPic(8, 2, 2) = 0
NumberPic(8, 3, 2) = 1
NumberPic(8, 4, 2) = 0
NumberPic(8, 5, 2) = 0
NumberPic(8, 6, 2) = 0
NumberPic(8, 7, 2) = 1

NumberPic(8, 0, 3) = 1
NumberPic(8, 1, 3) = 0
NumberPic(8, 2, 3) = 0
NumberPic(8, 3, 3) = 1
NumberPic(8, 4, 3) = 0
NumberPic(8, 5, 3) = 0
NumberPic(8, 6, 3) = 0
NumberPic(8, 7, 3) = 1

NumberPic(8, 0, 4) = 0
NumberPic(8, 1, 4) = 1
NumberPic(8, 2, 4) = 1
NumberPic(8, 3, 4) = 0
NumberPic(8, 4, 4) = 1
NumberPic(8, 5, 4) = 1
NumberPic(8, 6, 4) = 1
NumberPic(8, 7, 4) = 0

'Load number '7'
NumberPic(7, 0, 0) = 1
NumberPic(7, 1, 0) = 0
NumberPic(7, 2, 0) = 0
NumberPic(7, 3, 0) = 0
NumberPic(7, 4, 0) = 0
NumberPic(7, 5, 0) = 0
NumberPic(7, 6, 0) = 0
NumberPic(7, 7, 0) = 0

NumberPic(7, 0, 1) = 1
NumberPic(7, 1, 1) = 1
NumberPic(7, 2, 1) = 1
NumberPic(7, 3, 1) = 0
NumberPic(7, 4, 1) = 0
NumberPic(7, 5, 1) = 0
NumberPic(7, 6, 1) = 0
NumberPic(7, 7, 1) = 0

NumberPic(7, 0, 2) = 1
NumberPic(7, 1, 2) = 0
NumberPic(7, 2, 2) = 0
NumberPic(7, 3, 2) = 1
NumberPic(7, 4, 2) = 1
NumberPic(7, 5, 2) = 0
NumberPic(7, 6, 2) = 0
NumberPic(7, 7, 2) = 0

NumberPic(7, 0, 3) = 1
NumberPic(7, 1, 3) = 0
NumberPic(7, 2, 3) = 0
NumberPic(7, 3, 3) = 0
NumberPic(7, 4, 3) = 0
NumberPic(7, 5, 3) = 1
NumberPic(7, 6, 3) = 1
NumberPic(7, 7, 3) = 1

NumberPic(7, 0, 4) = 1
NumberPic(7, 1, 4) = 0
NumberPic(7, 2, 4) = 0
NumberPic(7, 3, 4) = 0
NumberPic(7, 4, 4) = 0
NumberPic(7, 5, 4) = 0
NumberPic(7, 6, 4) = 0
NumberPic(7, 7, 4) = 0


'Load number '6'
NumberPic(6, 0, 0) = 0
NumberPic(6, 1, 0) = 1
NumberPic(6, 2, 0) = 0
NumberPic(6, 3, 0) = 0
NumberPic(6, 4, 0) = 1
NumberPic(6, 5, 0) = 1
NumberPic(6, 6, 0) = 1
NumberPic(6, 7, 0) = 0

NumberPic(6, 0, 1) = 1
NumberPic(6, 1, 1) = 0
NumberPic(6, 2, 1) = 0
NumberPic(6, 3, 1) = 1
NumberPic(6, 4, 1) = 0
NumberPic(6, 5, 1) = 0
NumberPic(6, 6, 1) = 0
NumberPic(6, 7, 1) = 1

NumberPic(6, 0, 2) = 1
NumberPic(6, 1, 2) = 0
NumberPic(6, 2, 2) = 0
NumberPic(6, 3, 2) = 1
NumberPic(6, 4, 2) = 0
NumberPic(6, 5, 2) = 0
NumberPic(6, 6, 2) = 0
NumberPic(6, 7, 2) = 1

NumberPic(6, 0, 3) = 1
NumberPic(6, 1, 3) = 0
NumberPic(6, 2, 3) = 0
NumberPic(6, 3, 3) = 1
NumberPic(6, 4, 3) = 0
NumberPic(6, 5, 3) = 0
NumberPic(6, 6, 3) = 0
NumberPic(6, 7, 3) = 1

NumberPic(6, 0, 4) = 0
NumberPic(6, 1, 4) = 1
NumberPic(6, 2, 4) = 1
NumberPic(6, 3, 4) = 1
NumberPic(6, 4, 4) = 1
NumberPic(6, 5, 4) = 1
NumberPic(6, 6, 4) = 1
NumberPic(6, 7, 4) = 0

'Load number '5'
NumberPic(5, 0, 0) = 1
NumberPic(5, 1, 0) = 0
NumberPic(5, 2, 0) = 0
NumberPic(5, 3, 0) = 0
NumberPic(5, 4, 0) = 1
NumberPic(5, 5, 0) = 1
NumberPic(5, 6, 0) = 1
NumberPic(5, 7, 0) = 0

NumberPic(5, 0, 1) = 1
NumberPic(5, 1, 1) = 0
NumberPic(5, 2, 1) = 0
NumberPic(5, 3, 1) = 1
NumberPic(5, 4, 1) = 0
NumberPic(5, 5, 1) = 0
NumberPic(5, 6, 1) = 0
NumberPic(5, 7, 1) = 1

NumberPic(5, 0, 2) = 1
NumberPic(5, 1, 2) = 0
NumberPic(5, 2, 2) = 0
NumberPic(5, 3, 2) = 1
NumberPic(5, 4, 2) = 0
NumberPic(5, 5, 2) = 0
NumberPic(5, 6, 2) = 0
NumberPic(5, 7, 2) = 1

NumberPic(5, 0, 3) = 1
NumberPic(5, 1, 3) = 1
NumberPic(5, 2, 3) = 0
NumberPic(5, 3, 3) = 1
NumberPic(5, 4, 3) = 0
NumberPic(5, 5, 3) = 0
NumberPic(5, 6, 3) = 0
NumberPic(5, 7, 3) = 1

NumberPic(5, 0, 4) = 0
NumberPic(5, 1, 4) = 0
NumberPic(5, 2, 4) = 1
NumberPic(5, 3, 4) = 1
NumberPic(5, 4, 4) = 0
NumberPic(5, 5, 4) = 0
NumberPic(5, 6, 4) = 1
NumberPic(5, 7, 4) = 0

'Load number '4'
NumberPic(4, 0, 0) = 0
NumberPic(4, 1, 0) = 0
NumberPic(4, 2, 0) = 0
NumberPic(4, 3, 0) = 0
NumberPic(4, 4, 0) = 0
NumberPic(4, 5, 0) = 1
NumberPic(4, 6, 0) = 0
NumberPic(4, 7, 0) = 0

NumberPic(4, 0, 1) = 1
NumberPic(4, 1, 1) = 1
NumberPic(4, 2, 1) = 1
NumberPic(4, 3, 1) = 1
NumberPic(4, 4, 1) = 1
NumberPic(4, 5, 1) = 1
NumberPic(4, 6, 1) = 1
NumberPic(4, 7, 1) = 1

NumberPic(4, 0, 2) = 0
NumberPic(4, 1, 2) = 1
NumberPic(4, 2, 2) = 0
NumberPic(4, 3, 2) = 0
NumberPic(4, 4, 2) = 0
NumberPic(4, 5, 2) = 1
NumberPic(4, 6, 2) = 0
NumberPic(4, 7, 2) = 0

NumberPic(4, 0, 3) = 0
NumberPic(4, 1, 3) = 0
NumberPic(4, 2, 3) = 1
NumberPic(4, 3, 3) = 1
NumberPic(4, 4, 3) = 0
NumberPic(4, 5, 3) = 1
NumberPic(4, 6, 3) = 0
NumberPic(4, 7, 3) = 0

NumberPic(4, 0, 4) = 0
NumberPic(4, 1, 4) = 0
NumberPic(4, 2, 4) = 0
NumberPic(4, 3, 4) = 0
NumberPic(4, 4, 4) = 1
NumberPic(4, 5, 4) = 1
NumberPic(4, 6, 4) = 0
NumberPic(4, 7, 4) = 0

'Load number '3'
NumberPic(3, 0, 0) = 0
NumberPic(3, 1, 0) = 1
NumberPic(3, 2, 0) = 1
NumberPic(3, 3, 0) = 0
NumberPic(3, 4, 0) = 1
NumberPic(3, 5, 0) = 1
NumberPic(3, 6, 0) = 1
NumberPic(3, 7, 0) = 0

NumberPic(3, 0, 1) = 1
NumberPic(3, 1, 1) = 0
NumberPic(3, 2, 1) = 0
NumberPic(3, 3, 1) = 1
NumberPic(3, 4, 1) = 0
NumberPic(3, 5, 1) = 0
NumberPic(3, 6, 1) = 0
NumberPic(3, 7, 1) = 1

NumberPic(3, 0, 2) = 1
NumberPic(3, 1, 2) = 0
NumberPic(3, 2, 2) = 0
NumberPic(3, 3, 2) = 1
NumberPic(3, 4, 2) = 0
NumberPic(3, 5, 2) = 0
NumberPic(3, 6, 2) = 0
NumberPic(3, 7, 2) = 1

NumberPic(3, 0, 3) = 1
NumberPic(3, 1, 3) = 0
NumberPic(3, 2, 3) = 0
NumberPic(3, 3, 3) = 0
NumberPic(3, 4, 3) = 0
NumberPic(3, 5, 3) = 0
NumberPic(3, 6, 3) = 0
NumberPic(3, 7, 3) = 1

NumberPic(3, 0, 4) = 0
NumberPic(3, 1, 4) = 1
NumberPic(3, 2, 4) = 0
NumberPic(3, 3, 4) = 0
NumberPic(3, 4, 4) = 0
NumberPic(3, 5, 4) = 0
NumberPic(3, 6, 4) = 1
NumberPic(3, 7, 4) = 0

'Load number '2'
NumberPic(2, 0, 0) = 0
NumberPic(2, 1, 0) = 1
NumberPic(2, 2, 0) = 1
NumberPic(2, 3, 0) = 1
NumberPic(2, 4, 0) = 0
NumberPic(2, 5, 0) = 0
NumberPic(2, 6, 0) = 0
NumberPic(2, 7, 0) = 1

NumberPic(2, 0, 1) = 1
NumberPic(2, 1, 1) = 0
NumberPic(2, 2, 1) = 0
NumberPic(2, 3, 1) = 0
NumberPic(2, 4, 1) = 1
NumberPic(2, 5, 1) = 0
NumberPic(2, 6, 1) = 0
NumberPic(2, 7, 1) = 1

NumberPic(2, 0, 2) = 1
NumberPic(2, 1, 2) = 0
NumberPic(2, 2, 2) = 0
NumberPic(2, 3, 2) = 0
NumberPic(2, 4, 2) = 0
NumberPic(2, 5, 2) = 1
NumberPic(2, 6, 2) = 0
NumberPic(2, 7, 2) = 1

NumberPic(2, 0, 3) = 1
NumberPic(2, 1, 3) = 0
NumberPic(2, 2, 3) = 0
NumberPic(2, 3, 3) = 0
NumberPic(2, 4, 3) = 0
NumberPic(2, 5, 3) = 0
NumberPic(2, 6, 3) = 1
NumberPic(2, 7, 3) = 1

NumberPic(2, 0, 4) = 0
NumberPic(2, 1, 4) = 1
NumberPic(2, 2, 4) = 0
NumberPic(2, 3, 4) = 0
NumberPic(2, 4, 4) = 0
NumberPic(2, 5, 4) = 0
NumberPic(2, 6, 4) = 0
NumberPic(2, 7, 4) = 1

'Load number '1'
NumberPic(1, 0, 0) = 0
NumberPic(1, 1, 0) = 0
NumberPic(1, 2, 0) = 0
NumberPic(1, 3, 0) = 0
NumberPic(1, 4, 0) = 0
NumberPic(1, 5, 0) = 0
NumberPic(1, 6, 0) = 0
NumberPic(1, 7, 0) = 0

NumberPic(1, 0, 1) = 1
NumberPic(1, 1, 1) = 1
NumberPic(1, 2, 1) = 1
NumberPic(1, 3, 1) = 1
NumberPic(1, 4, 1) = 1
NumberPic(1, 5, 1) = 1
NumberPic(1, 6, 1) = 1
NumberPic(1, 7, 1) = 1

NumberPic(1, 0, 2) = 0
NumberPic(1, 1, 2) = 1
NumberPic(1, 2, 2) = 0
NumberPic(1, 3, 2) = 0
NumberPic(1, 4, 2) = 0
NumberPic(1, 5, 2) = 0
NumberPic(1, 6, 2) = 0
NumberPic(1, 7, 2) = 0

NumberPic(1, 0, 3) = 0
NumberPic(1, 1, 3) = 0
NumberPic(1, 2, 3) = 1
NumberPic(1, 3, 3) = 0
NumberPic(1, 4, 3) = 0
NumberPic(1, 5, 3) = 0
NumberPic(1, 6, 3) = 0
NumberPic(1, 7, 3) = 0

NumberPic(1, 0, 4) = 0
NumberPic(1, 1, 4) = 0
NumberPic(1, 2, 4) = 0
NumberPic(1, 3, 4) = 0
NumberPic(1, 4, 4) = 0
NumberPic(1, 5, 4) = 0
NumberPic(1, 6, 4) = 0
NumberPic(1, 7, 4) = 0

'Load number '0'
NumberPic(0, 0, 0) = 0
NumberPic(0, 1, 0) = 1
NumberPic(0, 2, 0) = 1
NumberPic(0, 3, 0) = 1
NumberPic(0, 4, 0) = 1
NumberPic(0, 5, 0) = 1
NumberPic(0, 6, 0) = 1
NumberPic(0, 7, 0) = 0

NumberPic(0, 0, 1) = 1
NumberPic(0, 1, 1) = 0
NumberPic(0, 2, 1) = 0
NumberPic(0, 3, 1) = 0
NumberPic(0, 4, 1) = 0
NumberPic(0, 5, 1) = 0
NumberPic(0, 6, 1) = 0
NumberPic(0, 7, 1) = 1

NumberPic(0, 0, 2) = 1
NumberPic(0, 1, 2) = 0
NumberPic(0, 2, 2) = 0
NumberPic(0, 3, 2) = 0
NumberPic(0, 4, 2) = 0
NumberPic(0, 5, 2) = 0
NumberPic(0, 6, 2) = 0
NumberPic(0, 7, 2) = 1

NumberPic(0, 0, 3) = 1
NumberPic(0, 1, 3) = 0
NumberPic(0, 2, 3) = 0
NumberPic(0, 3, 3) = 0
NumberPic(0, 4, 3) = 0
NumberPic(0, 5, 3) = 0
NumberPic(0, 6, 3) = 0
NumberPic(0, 7, 3) = 1

NumberPic(0, 0, 4) = 0
NumberPic(0, 1, 4) = 1
NumberPic(0, 2, 4) = 1
NumberPic(0, 3, 4) = 1
NumberPic(0, 4, 4) = 1
NumberPic(0, 5, 4) = 1
NumberPic(0, 6, 4) = 1
NumberPic(0, 7, 4) = 0


'Load number Nothing
NumberPic(10, 0, 0) = 0
NumberPic(10, 1, 0) = 0
NumberPic(10, 2, 0) = 0
NumberPic(10, 3, 0) = 0
NumberPic(10, 4, 0) = 0
NumberPic(10, 5, 0) = 0
NumberPic(10, 6, 0) = 0
NumberPic(10, 7, 0) = 0

NumberPic(10, 0, 1) = 0
NumberPic(10, 1, 1) = 0
NumberPic(10, 2, 1) = 0
NumberPic(10, 3, 1) = 0
NumberPic(10, 4, 1) = 0
NumberPic(10, 5, 1) = 0
NumberPic(10, 6, 1) = 0
NumberPic(10, 7, 1) = 0

NumberPic(10, 0, 2) = 0
NumberPic(10, 1, 2) = 0
NumberPic(10, 2, 2) = 0
NumberPic(10, 3, 2) = 0
NumberPic(10, 4, 2) = 0
NumberPic(10, 5, 2) = 0
NumberPic(10, 6, 2) = 0
NumberPic(10, 7, 2) = 0

NumberPic(10, 0, 3) = 0
NumberPic(10, 1, 3) = 0
NumberPic(10, 2, 3) = 0
NumberPic(10, 3, 3) = 0
NumberPic(10, 4, 3) = 0
NumberPic(10, 5, 3) = 0
NumberPic(10, 6, 3) = 0
NumberPic(10, 7, 3) = 0

NumberPic(10, 0, 4) = 0
NumberPic(10, 1, 4) = 0
NumberPic(10, 2, 4) = 0
NumberPic(10, 3, 4) = 0
NumberPic(10, 4, 4) = 0
NumberPic(10, 5, 4) = 0
NumberPic(10, 6, 4) = 0
NumberPic(10, 7, 4) = 0

'Load number '.' 'PCNGL090103
NumberPic(11, 0, 0) = 0
NumberPic(11, 1, 0) = 0
NumberPic(11, 2, 0) = 0
NumberPic(11, 3, 0) = 0
NumberPic(11, 4, 0) = 0
NumberPic(11, 5, 0) = 0
NumberPic(11, 6, 0) = 0
NumberPic(11, 7, 0) = 0

NumberPic(11, 0, 1) = 0
NumberPic(11, 1, 1) = 0
NumberPic(11, 2, 1) = 0
NumberPic(11, 3, 1) = 0
NumberPic(11, 4, 1) = 0
NumberPic(11, 5, 1) = 0
NumberPic(11, 6, 1) = 0
NumberPic(11, 7, 1) = 0

NumberPic(11, 0, 2) = 0
NumberPic(11, 1, 2) = 0
NumberPic(11, 2, 2) = 0
NumberPic(11, 3, 2) = 0
NumberPic(11, 4, 2) = 0
NumberPic(11, 5, 2) = 0
NumberPic(11, 6, 2) = 0
NumberPic(11, 7, 2) = 1

NumberPic(11, 0, 3) = 0
NumberPic(11, 1, 3) = 0
NumberPic(11, 2, 3) = 0
NumberPic(11, 3, 3) = 0
NumberPic(11, 4, 3) = 0
NumberPic(11, 5, 3) = 0
NumberPic(11, 6, 3) = 0
NumberPic(11, 7, 3) = 0

NumberPic(11, 0, 4) = 0
NumberPic(11, 1, 4) = 0
NumberPic(11, 2, 4) = 0
NumberPic(11, 3, 4) = 0
NumberPic(11, 4, 4) = 0
NumberPic(11, 5, 4) = 0
NumberPic(11, 6, 4) = 0
NumberPic(11, 7, 4) = 0

'Load number 'k' 'PCNGL090103
NumberPic(12, 0, 0) = 0
NumberPic(12, 1, 0) = 0
NumberPic(12, 2, 0) = 0
NumberPic(12, 3, 0) = 0
NumberPic(12, 4, 0) = 0
NumberPic(12, 5, 0) = 0
NumberPic(12, 6, 0) = 0
NumberPic(12, 7, 0) = 0

NumberPic(12, 0, 1) = 0
NumberPic(12, 1, 1) = 0
NumberPic(12, 2, 1) = 1
NumberPic(12, 3, 1) = 0
NumberPic(12, 4, 1) = 0
NumberPic(12, 5, 1) = 0
NumberPic(12, 6, 1) = 0
NumberPic(12, 7, 1) = 1

NumberPic(12, 0, 2) = 0
NumberPic(12, 1, 2) = 0
NumberPic(12, 2, 2) = 0
NumberPic(12, 3, 2) = 1
NumberPic(12, 4, 2) = 0
NumberPic(12, 5, 2) = 1
NumberPic(12, 6, 2) = 1
NumberPic(12, 7, 2) = 0

NumberPic(12, 0, 3) = 0
NumberPic(12, 1, 3) = 0
NumberPic(12, 2, 3) = 0
NumberPic(12, 3, 3) = 0
NumberPic(12, 4, 3) = 1
NumberPic(12, 5, 3) = 0
NumberPic(12, 6, 3) = 0
NumberPic(12, 7, 3) = 0

NumberPic(12, 0, 4) = 1
NumberPic(12, 1, 4) = 1
NumberPic(12, 2, 4) = 1
NumberPic(12, 3, 4) = 1
NumberPic(12, 4, 4) = 1
NumberPic(12, 5, 4) = 1
NumberPic(12, 6, 4) = 1
NumberPic(12, 7, 4) = 1

'vvvv PCN2777 Load number '-' *****************
NumberPic(13, 0, 0) = 0
NumberPic(13, 1, 0) = 0
NumberPic(13, 2, 0) = 0
NumberPic(13, 3, 0) = 0
NumberPic(13, 4, 0) = 0
NumberPic(13, 5, 0) = 0
NumberPic(13, 6, 0) = 0
NumberPic(13, 7, 0) = 0

NumberPic(13, 0, 1) = 0
NumberPic(13, 1, 1) = 0
NumberPic(13, 2, 1) = 0
NumberPic(13, 3, 1) = 0
NumberPic(13, 4, 1) = 1
NumberPic(13, 5, 1) = 0
NumberPic(13, 6, 1) = 0
NumberPic(13, 7, 1) = 0

NumberPic(13, 0, 2) = 0
NumberPic(13, 1, 2) = 0
NumberPic(13, 2, 2) = 0
NumberPic(13, 3, 2) = 0
NumberPic(13, 4, 2) = 1
NumberPic(13, 5, 2) = 0
NumberPic(13, 6, 2) = 0
NumberPic(13, 7, 2) = 0

NumberPic(13, 0, 3) = 0
NumberPic(13, 1, 3) = 0
NumberPic(13, 2, 3) = 0
NumberPic(13, 3, 3) = 0
NumberPic(13, 4, 3) = 1
NumberPic(13, 5, 3) = 0
NumberPic(13, 6, 3) = 0
NumberPic(13, 7, 3) = 0

NumberPic(13, 0, 4) = 0
NumberPic(13, 1, 4) = 0
NumberPic(13, 2, 4) = 0
NumberPic(13, 3, 4) = 0
NumberPic(13, 4, 4) = 0
NumberPic(13, 5, 4) = 0
NumberPic(13, 6, 4) = 0
NumberPic(13, 7, 4) = 0
'^^^^ ************************************



'The X Scale
'Load the number 9
NumberPicX(9, 4, 0) = 0
NumberPicX(9, 4, 1) = 1
NumberPicX(9, 4, 2) = 1
NumberPicX(9, 4, 3) = 1
NumberPicX(9, 4, 4) = 1
NumberPicX(9, 4, 5) = 1
NumberPicX(9, 4, 6) = 1
NumberPicX(9, 4, 7) = 0
NumberPicX(9, 3, 0) = 1
NumberPicX(9, 3, 1) = 0
NumberPicX(9, 3, 2) = 0
NumberPicX(9, 3, 3) = 0
NumberPicX(9, 3, 4) = 1
NumberPicX(9, 3, 5) = 0
NumberPicX(9, 3, 6) = 0
NumberPicX(9, 3, 7) = 1
NumberPicX(9, 2, 0) = 1
NumberPicX(9, 2, 1) = 0
NumberPicX(9, 2, 2) = 0
NumberPicX(9, 2, 3) = 0
NumberPicX(9, 2, 4) = 1
NumberPicX(9, 2, 5) = 0
NumberPicX(9, 2, 6) = 0
NumberPicX(9, 2, 7) = 1
NumberPicX(9, 1, 0) = 1
NumberPicX(9, 1, 1) = 0
NumberPicX(9, 1, 2) = 0
NumberPicX(9, 1, 3) = 0
NumberPicX(9, 1, 4) = 1
NumberPicX(9, 1, 5) = 0
NumberPicX(9, 1, 6) = 0
NumberPicX(9, 1, 7) = 1
NumberPicX(9, 0, 0) = 0
NumberPicX(9, 0, 1) = 1
NumberPicX(9, 0, 2) = 1
NumberPicX(9, 0, 3) = 1
NumberPicX(9, 0, 4) = 0
NumberPicX(9, 0, 5) = 0
NumberPicX(9, 0, 6) = 1
NumberPicX(9, 0, 7) = 0
'Load number '8'
NumberPicX(8, 4, 0) = 0
NumberPicX(8, 4, 1) = 1
NumberPicX(8, 4, 2) = 1
NumberPicX(8, 4, 3) = 0
NumberPicX(8, 4, 4) = 1
NumberPicX(8, 4, 5) = 1
NumberPicX(8, 4, 6) = 1
NumberPicX(8, 4, 7) = 0

NumberPicX(8, 3, 0) = 1
NumberPicX(8, 3, 1) = 0
NumberPicX(8, 3, 2) = 0
NumberPicX(8, 3, 3) = 1
NumberPicX(8, 3, 4) = 0
NumberPicX(8, 3, 5) = 0
NumberPicX(8, 3, 6) = 0
NumberPicX(8, 3, 7) = 1

NumberPicX(8, 2, 0) = 1
NumberPicX(8, 2, 1) = 0
NumberPicX(8, 2, 2) = 0
NumberPicX(8, 2, 3) = 1
NumberPicX(8, 2, 4) = 0
NumberPicX(8, 2, 5) = 0
NumberPicX(8, 2, 6) = 0
NumberPicX(8, 2, 7) = 1

NumberPicX(8, 1, 0) = 1
NumberPicX(8, 1, 1) = 0
NumberPicX(8, 1, 2) = 0
NumberPicX(8, 1, 3) = 1
NumberPicX(8, 1, 4) = 0
NumberPicX(8, 1, 5) = 0
NumberPicX(8, 1, 6) = 0
NumberPicX(8, 1, 7) = 1

NumberPicX(8, 0, 0) = 0
NumberPicX(8, 0, 1) = 1
NumberPicX(8, 0, 2) = 1
NumberPicX(8, 0, 3) = 0
NumberPicX(8, 0, 4) = 1
NumberPicX(8, 0, 5) = 1
NumberPicX(8, 0, 6) = 1
NumberPicX(8, 0, 7) = 0

'Load number '7'
NumberPicX(7, 4, 0) = 1
NumberPicX(7, 4, 1) = 0
NumberPicX(7, 4, 2) = 0
NumberPicX(7, 4, 3) = 0
NumberPicX(7, 4, 4) = 0
NumberPicX(7, 4, 5) = 0
NumberPicX(7, 4, 6) = 0
NumberPicX(7, 4, 7) = 0

NumberPicX(7, 3, 0) = 1
NumberPicX(7, 3, 1) = 1
NumberPicX(7, 3, 2) = 1
NumberPicX(7, 3, 3) = 0
NumberPicX(7, 3, 4) = 0
NumberPicX(7, 3, 5) = 0
NumberPicX(7, 3, 6) = 0
NumberPicX(7, 3, 7) = 0

NumberPicX(7, 2, 0) = 1
NumberPicX(7, 2, 1) = 0
NumberPicX(7, 2, 2) = 0
NumberPicX(7, 2, 3) = 1
NumberPicX(7, 2, 4) = 1
NumberPicX(7, 2, 5) = 0
NumberPicX(7, 2, 6) = 0
NumberPicX(7, 2, 7) = 0

NumberPicX(7, 1, 0) = 1
NumberPicX(7, 1, 1) = 0
NumberPicX(7, 1, 2) = 0
NumberPicX(7, 1, 3) = 0
NumberPicX(7, 1, 4) = 0
NumberPicX(7, 1, 5) = 1
NumberPicX(7, 1, 6) = 1
NumberPicX(7, 1, 7) = 1

NumberPicX(7, 0, 0) = 1
NumberPicX(7, 0, 1) = 0
NumberPicX(7, 0, 2) = 0
NumberPicX(7, 0, 3) = 0
NumberPicX(7, 0, 4) = 0
NumberPicX(7, 0, 5) = 0
NumberPicX(7, 0, 6) = 0
NumberPicX(7, 0, 7) = 0


'Load number '6'
NumberPicX(6, 4, 0) = 0
NumberPicX(6, 4, 1) = 1
NumberPicX(6, 4, 2) = 0
NumberPicX(6, 4, 3) = 0
NumberPicX(6, 4, 4) = 1
NumberPicX(6, 4, 5) = 1
NumberPicX(6, 4, 6) = 1
NumberPicX(6, 4, 7) = 0

NumberPicX(6, 3, 0) = 1
NumberPicX(6, 3, 1) = 0
NumberPicX(6, 3, 2) = 0
NumberPicX(6, 3, 3) = 1
NumberPicX(6, 3, 4) = 0
NumberPicX(6, 3, 5) = 0
NumberPicX(6, 3, 6) = 0
NumberPicX(6, 3, 7) = 1

NumberPicX(6, 2, 0) = 1
NumberPicX(6, 2, 1) = 0
NumberPicX(6, 2, 2) = 0
NumberPicX(6, 2, 3) = 1
NumberPicX(6, 2, 4) = 0
NumberPicX(6, 2, 5) = 0
NumberPicX(6, 2, 6) = 0
NumberPicX(6, 2, 7) = 1

NumberPicX(6, 1, 0) = 1
NumberPicX(6, 1, 1) = 0
NumberPicX(6, 1, 2) = 0
NumberPicX(6, 1, 3) = 1
NumberPicX(6, 1, 4) = 0
NumberPicX(6, 1, 5) = 0
NumberPicX(6, 1, 6) = 0
NumberPicX(6, 1, 7) = 1

NumberPicX(6, 0, 0) = 0
NumberPicX(6, 0, 1) = 1
NumberPicX(6, 0, 2) = 1
NumberPicX(6, 0, 3) = 1
NumberPicX(6, 0, 4) = 1
NumberPicX(6, 0, 5) = 1
NumberPicX(6, 0, 6) = 1
NumberPicX(6, 0, 7) = 0

'Load number '5'
NumberPicX(5, 4, 1) = 1
NumberPicX(5, 4, 1) = 0
NumberPicX(5, 4, 2) = 0
NumberPicX(5, 4, 3) = 0
NumberPicX(5, 4, 4) = 1
NumberPicX(5, 4, 5) = 1
NumberPicX(5, 4, 6) = 1
NumberPicX(5, 4, 7) = 0

NumberPicX(5, 3, 0) = 1
NumberPicX(5, 3, 1) = 0
NumberPicX(5, 3, 2) = 0
NumberPicX(5, 3, 3) = 1
NumberPicX(5, 3, 4) = 0
NumberPicX(5, 3, 5) = 0
NumberPicX(5, 3, 6) = 0
NumberPicX(5, 3, 7) = 1

NumberPicX(5, 2, 0) = 1
NumberPicX(5, 2, 1) = 0
NumberPicX(5, 2, 2) = 0
NumberPicX(5, 2, 3) = 1
NumberPicX(5, 2, 4) = 0
NumberPicX(5, 2, 5) = 0
NumberPicX(5, 2, 6) = 0
NumberPicX(5, 2, 7) = 1

NumberPicX(5, 1, 0) = 1
NumberPicX(5, 1, 1) = 1
NumberPicX(5, 1, 2) = 0
NumberPicX(5, 1, 3) = 1
NumberPicX(5, 1, 4) = 0
NumberPicX(5, 1, 5) = 0
NumberPicX(5, 1, 6) = 0
NumberPicX(5, 1, 7) = 1

NumberPicX(5, 0, 0) = 0
NumberPicX(5, 0, 1) = 0
NumberPicX(5, 0, 2) = 1
NumberPicX(5, 0, 3) = 1
NumberPicX(5, 0, 4) = 0
NumberPicX(5, 0, 5) = 0
NumberPicX(5, 0, 6) = 1
NumberPicX(5, 0, 7) = 0

'Load number '4'
NumberPicX(4, 4, 0) = 0
NumberPicX(4, 4, 1) = 0
NumberPicX(4, 4, 2) = 0
NumberPicX(4, 4, 3) = 0
NumberPicX(4, 4, 4) = 0
NumberPicX(4, 4, 5) = 1
NumberPicX(4, 4, 6) = 0
NumberPicX(4, 4, 7) = 0

NumberPicX(4, 3, 0) = 1
NumberPicX(4, 3, 1) = 1
NumberPicX(4, 3, 2) = 1
NumberPicX(4, 3, 3) = 1
NumberPicX(4, 3, 4) = 1
NumberPicX(4, 3, 5) = 1
NumberPicX(4, 3, 6) = 1
NumberPicX(4, 3, 7) = 1

NumberPicX(4, 2, 0) = 0
NumberPicX(4, 2, 1) = 1
NumberPicX(4, 2, 2) = 0
NumberPicX(4, 2, 3) = 0
NumberPicX(4, 2, 4) = 0
NumberPicX(4, 2, 5) = 1
NumberPicX(4, 2, 6) = 0
NumberPicX(4, 2, 7) = 0

NumberPicX(4, 1, 0) = 0
NumberPicX(4, 1, 1) = 0
NumberPicX(4, 1, 2) = 1
NumberPicX(4, 1, 3) = 1
NumberPicX(4, 1, 4) = 0
NumberPicX(4, 1, 5) = 1
NumberPicX(4, 1, 6) = 0
NumberPicX(4, 1, 7) = 0

NumberPicX(4, 0, 0) = 0
NumberPicX(4, 0, 1) = 0
NumberPicX(4, 0, 2) = 0
NumberPicX(4, 0, 3) = 0
NumberPicX(4, 0, 4) = 1
NumberPicX(4, 0, 5) = 1
NumberPicX(4, 0, 6) = 0
NumberPicX(4, 0, 7) = 0

'Load number '3'
NumberPicX(3, 4, 0) = 0
NumberPicX(3, 4, 1) = 1
NumberPicX(3, 4, 2) = 1
NumberPicX(3, 4, 3) = 0
NumberPicX(3, 4, 4) = 1
NumberPicX(3, 4, 5) = 1
NumberPicX(3, 4, 6) = 1
NumberPicX(3, 4, 7) = 0

NumberPicX(3, 3, 0) = 1
NumberPicX(3, 3, 1) = 0
NumberPicX(3, 3, 2) = 0
NumberPicX(3, 3, 3) = 1
NumberPicX(3, 3, 4) = 0
NumberPicX(3, 3, 5) = 0
NumberPicX(3, 3, 6) = 0
NumberPicX(3, 3, 7) = 1

NumberPicX(3, 2, 0) = 1
NumberPicX(3, 2, 1) = 0
NumberPicX(3, 2, 2) = 0
NumberPicX(3, 2, 3) = 1
NumberPicX(3, 2, 4) = 0
NumberPicX(3, 2, 5) = 0
NumberPicX(3, 2, 6) = 0
NumberPicX(3, 2, 7) = 1

NumberPicX(3, 1, 0) = 1
NumberPicX(3, 1, 1) = 0
NumberPicX(3, 1, 2) = 0
NumberPicX(3, 1, 3) = 0
NumberPicX(3, 1, 4) = 0
NumberPicX(3, 1, 5) = 0
NumberPicX(3, 1, 6) = 0
NumberPicX(3, 1, 7) = 1

NumberPicX(3, 0, 0) = 0
NumberPicX(3, 0, 1) = 1
NumberPicX(3, 0, 2) = 0
NumberPicX(3, 0, 3) = 0
NumberPicX(3, 0, 4) = 0
NumberPicX(3, 0, 5) = 0
NumberPicX(3, 0, 6) = 1
NumberPicX(3, 0, 7) = 0

'Load number '2'
NumberPicX(2, 4, 0) = 0
NumberPicX(2, 4, 1) = 1
NumberPicX(2, 4, 2) = 1
NumberPicX(2, 4, 3) = 1
NumberPicX(2, 4, 4) = 0
NumberPicX(2, 4, 5) = 0
NumberPicX(2, 4, 6) = 0
NumberPicX(2, 4, 7) = 1

NumberPicX(2, 3, 0) = 1
NumberPicX(2, 3, 1) = 0
NumberPicX(2, 3, 2) = 0
NumberPicX(2, 3, 3) = 0
NumberPicX(2, 3, 4) = 1
NumberPicX(2, 3, 5) = 0
NumberPicX(2, 3, 6) = 0
NumberPicX(2, 3, 7) = 1

NumberPicX(2, 2, 0) = 1
NumberPicX(2, 2, 1) = 0
NumberPicX(2, 2, 2) = 0
NumberPicX(2, 2, 3) = 0
NumberPicX(2, 2, 4) = 0
NumberPicX(2, 2, 5) = 1
NumberPicX(2, 2, 6) = 0
NumberPicX(2, 2, 7) = 1

NumberPicX(2, 1, 0) = 1
NumberPicX(2, 1, 1) = 0
NumberPicX(2, 1, 2) = 0
NumberPicX(2, 1, 3) = 0
NumberPicX(2, 1, 4) = 0
NumberPicX(2, 1, 5) = 0
NumberPicX(2, 1, 6) = 1
NumberPicX(2, 1, 7) = 1

NumberPicX(2, 0, 0) = 0
NumberPicX(2, 0, 1) = 1
NumberPicX(2, 0, 2) = 0
NumberPicX(2, 0, 3) = 0
NumberPicX(2, 0, 4) = 0
NumberPicX(2, 0, 5) = 0
NumberPicX(2, 0, 6) = 0
NumberPicX(2, 0, 7) = 1

'Load number '1'
NumberPicX(1, 4, 0) = 0
NumberPicX(1, 4, 1) = 0
NumberPicX(1, 4, 2) = 0
NumberPicX(1, 4, 3) = 0
NumberPicX(1, 4, 4) = 0
NumberPicX(1, 4, 5) = 0
NumberPicX(1, 4, 6) = 0
NumberPicX(1, 4, 7) = 0

NumberPicX(1, 3, 0) = 1
NumberPicX(1, 3, 1) = 1
NumberPicX(1, 3, 2) = 1
NumberPicX(1, 3, 3) = 1
NumberPicX(1, 3, 4) = 1
NumberPicX(1, 3, 5) = 1
NumberPicX(1, 3, 6) = 1
NumberPicX(1, 3, 7) = 1

NumberPicX(1, 2, 0) = 0
NumberPicX(1, 2, 1) = 1
NumberPicX(1, 2, 2) = 0
NumberPicX(1, 2, 3) = 0
NumberPicX(1, 2, 4) = 0
NumberPicX(1, 2, 5) = 0
NumberPicX(1, 2, 6) = 0
NumberPicX(1, 2, 7) = 0

NumberPicX(1, 1, 0) = 0
NumberPicX(1, 1, 1) = 0
NumberPicX(1, 1, 2) = 1
NumberPicX(1, 1, 3) = 0
NumberPicX(1, 1, 4) = 0
NumberPicX(1, 1, 5) = 0
NumberPicX(1, 1, 6) = 0
NumberPicX(1, 1, 7) = 0

NumberPicX(1, 0, 0) = 0
NumberPicX(1, 0, 1) = 0
NumberPicX(1, 0, 2) = 0
NumberPicX(1, 0, 3) = 0
NumberPicX(1, 0, 4) = 0
NumberPicX(1, 0, 5) = 0
NumberPicX(1, 0, 6) = 0
NumberPicX(1, 0, 7) = 0

'Load number '0'
NumberPicX(0, 4, 0) = 0
NumberPicX(0, 4, 1) = 1
NumberPicX(0, 4, 2) = 1
NumberPicX(0, 4, 3) = 1
NumberPicX(0, 4, 4) = 1
NumberPicX(0, 4, 5) = 1
NumberPicX(0, 4, 6) = 1
NumberPicX(0, 4, 7) = 0

NumberPicX(0, 3, 0) = 1
NumberPicX(0, 3, 1) = 0
NumberPicX(0, 3, 2) = 0
NumberPicX(0, 3, 3) = 0
NumberPicX(0, 3, 4) = 0
NumberPicX(0, 3, 5) = 0
NumberPicX(0, 3, 6) = 0
NumberPicX(0, 3, 7) = 1

NumberPicX(0, 2, 0) = 1
NumberPicX(0, 2, 1) = 0
NumberPicX(0, 2, 2) = 0
NumberPicX(0, 2, 3) = 0
NumberPicX(0, 2, 4) = 0
NumberPicX(0, 2, 5) = 0
NumberPicX(0, 2, 6) = 0
NumberPicX(0, 2, 7) = 1

NumberPicX(0, 1, 0) = 1
NumberPicX(0, 1, 1) = 0
NumberPicX(0, 1, 2) = 0
NumberPicX(0, 1, 3) = 0
NumberPicX(0, 1, 4) = 0
NumberPicX(0, 1, 5) = 0
NumberPicX(0, 1, 6) = 0
NumberPicX(0, 1, 7) = 1

NumberPicX(0, 0, 0) = 0
NumberPicX(0, 0, 1) = 1
NumberPicX(0, 0, 2) = 1
NumberPicX(0, 0, 3) = 1
NumberPicX(0, 0, 4) = 1
NumberPicX(0, 0, 5) = 1
NumberPicX(0, 0, 6) = 1
NumberPicX(0, 0, 7) = 0


'Load number Nothing
NumberPicX(10, 0, 0) = 0
NumberPicX(10, 0, 1) = 0
NumberPicX(10, 0, 2) = 0
NumberPicX(10, 0, 3) = 0
NumberPicX(10, 0, 4) = 0
NumberPicX(10, 0, 5) = 0
NumberPicX(10, 0, 6) = 0
NumberPicX(10, 0, 7) = 0

NumberPicX(10, 1, 0) = 0
NumberPicX(10, 1, 1) = 0
NumberPicX(10, 1, 2) = 0
NumberPicX(10, 1, 3) = 0
NumberPicX(10, 1, 4) = 0
NumberPicX(10, 1, 5) = 0
NumberPicX(10, 1, 6) = 0
NumberPicX(10, 1, 7) = 0

NumberPicX(10, 2, 0) = 0
NumberPicX(10, 2, 1) = 0
NumberPicX(10, 2, 2) = 0
NumberPicX(10, 2, 3) = 0
NumberPicX(10, 2, 4) = 0
NumberPicX(10, 2, 5) = 0
NumberPicX(10, 2, 6) = 0
NumberPicX(10, 2, 7) = 0

NumberPicX(10, 3, 0) = 0
NumberPicX(10, 3, 1) = 0
NumberPicX(10, 3, 2) = 0
NumberPicX(10, 3, 3) = 0
NumberPicX(10, 3, 4) = 0
NumberPicX(10, 3, 5) = 0
NumberPicX(10, 3, 6) = 0
NumberPicX(10, 3, 7) = 0

NumberPicX(10, 4, 0) = 0
NumberPicX(10, 4, 1) = 0
NumberPicX(10, 4, 2) = 0
NumberPicX(10, 4, 3) = 0
NumberPicX(10, 4, 4) = 0
NumberPicX(10, 4, 5) = 0
NumberPicX(10, 4, 6) = 0
NumberPicX(10, 4, 7) = 0

'vvvv PCN2463 Load number '.' *****************
NumberPicX(11, 0, 0) = 0
NumberPicX(11, 0, 1) = 0
NumberPicX(11, 0, 2) = 0
NumberPicX(11, 0, 3) = 0
NumberPicX(11, 0, 4) = 0
NumberPicX(11, 0, 5) = 0
NumberPicX(11, 0, 6) = 0
NumberPicX(11, 0, 7) = 0

NumberPicX(11, 1, 0) = 0
NumberPicX(11, 1, 1) = 0
NumberPicX(11, 1, 2) = 0
NumberPicX(11, 1, 3) = 0
NumberPicX(11, 1, 4) = 0
NumberPicX(11, 1, 5) = 0
NumberPicX(11, 1, 6) = 0
NumberPicX(11, 1, 7) = 0

NumberPicX(11, 2, 0) = 0
NumberPicX(11, 2, 1) = 0
NumberPicX(11, 2, 2) = 0
NumberPicX(11, 2, 3) = 0
NumberPicX(11, 2, 4) = 0
NumberPicX(11, 2, 5) = 0
NumberPicX(11, 2, 6) = 0
NumberPicX(11, 2, 7) = 1

NumberPicX(11, 3, 0) = 0
NumberPicX(11, 3, 1) = 0
NumberPicX(11, 3, 2) = 0
NumberPicX(11, 3, 3) = 0
NumberPicX(11, 3, 4) = 0
NumberPicX(11, 3, 5) = 0
NumberPicX(11, 3, 6) = 0
NumberPicX(11, 3, 7) = 0

NumberPicX(11, 4, 0) = 0
NumberPicX(11, 4, 1) = 0
NumberPicX(11, 4, 2) = 0
NumberPicX(11, 4, 3) = 0
NumberPicX(11, 4, 4) = 0
NumberPicX(11, 4, 5) = 0
NumberPicX(11, 4, 6) = 0
NumberPicX(11, 4, 7) = 0
'^^^^ ************************************

'vvvv PCN2777 Load number '-' *****************
NumberPicX(13, 0, 0) = 0
NumberPicX(13, 0, 1) = 0
NumberPicX(13, 0, 2) = 0
NumberPicX(13, 0, 3) = 0
NumberPicX(13, 0, 4) = 0
NumberPicX(13, 0, 5) = 0
NumberPicX(13, 0, 6) = 0
NumberPicX(13, 0, 7) = 0

NumberPicX(13, 1, 0) = 0
NumberPicX(13, 1, 1) = 0
NumberPicX(13, 1, 2) = 0
NumberPicX(13, 1, 3) = 0
NumberPicX(13, 1, 4) = 1
NumberPicX(13, 1, 5) = 0
NumberPicX(13, 1, 6) = 0
NumberPicX(13, 1, 7) = 0

NumberPicX(13, 2, 0) = 0
NumberPicX(13, 2, 1) = 0
NumberPicX(13, 2, 2) = 0
NumberPicX(13, 2, 3) = 0
NumberPicX(13, 2, 4) = 1
NumberPicX(13, 2, 5) = 0
NumberPicX(13, 2, 6) = 0
NumberPicX(13, 2, 7) = 0

NumberPicX(13, 3, 0) = 0
NumberPicX(13, 3, 1) = 0
NumberPicX(13, 3, 2) = 0
NumberPicX(13, 3, 3) = 0
NumberPicX(13, 3, 4) = 1
NumberPicX(13, 3, 5) = 0
NumberPicX(13, 3, 6) = 0
NumberPicX(13, 3, 7) = 0

NumberPicX(13, 4, 0) = 0
NumberPicX(13, 4, 1) = 0
NumberPicX(13, 4, 2) = 0
NumberPicX(13, 4, 3) = 0
NumberPicX(13, 4, 4) = 0
NumberPicX(13, 4, 5) = 0
NumberPicX(13, 4, 6) = 0
NumberPicX(13, 4, 7) = 0
'^^^^ ************************************



Exit Sub
Err_Handler:
    MsgBox Err & error$
End Sub





Sub DrawPVYScaleNumbersAndMarkers(ImageControlName As Control, PVYScaleNoOfPage As Integer, ByVal StartFrameNo As Long)    'PCNGL161202 'PCNGL080103 'PCNGL090103 'PCN2971
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : DrawPVYScaleNumbersAndMarkers
'Created :
'Updated : 8 January 2003, Building initial framework
'          18 January 2003, Changed name to DrawPVYScaleNumbersAndMarkers
'          23 March 2004, PCN2639
'Prg By  : Geoff Logan
'Param   : ImageControlName - Image control name
'          PVYScaleNoOfPage - The PVScale is divided into pages the height of the PVWindow
'          StartFrameNo - Frame at which to start drawing scale
'Desc    :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim LocMark As Long 'This is the location on the Y scale of the 4 second mark
Dim ScaleValue As Double  'The value to be displayed at this scale position  'PCNGL090103
Dim ScaleDigitCode(8) As Integer 'Holds the digits within ScaleValue 'PCNGL090103 'PCN2463 extended to 7
Dim ScaleDigitOffset As Integer 'The y position offset between digits 'PCNGL090103
Dim ScaleDigitID As Integer 'The digits position within the ScaleValue, e.g. ScaleDigitID = 2 for "234" represents "3"
Dim ScaleNoOfDigits As Integer 'A count of the number of digits
Dim ScaleDigitStartPosOffset As Integer 'Used to ensure ScaleValue is positioned in the centre of the scale marker
Dim ScaleStep As Long
Dim ZeroOffset As Integer 'PCNGL091202
Dim ScalePos As Long 'PCNGL161202
Dim ScaleMinorMarkerOffset As Long 'PCN2462
Dim MinorMarkStep As Integer
'Dim YScaleUnitFactor As Integer  'Denotes when the scale changes from single to k units 'PCNGL080103  'PCN2463
Dim NoneZeroDigitFlag As Boolean 'Used to ignore the right hand digits of ScaleValue that are zero
Dim ScaleStart As Long
Dim ScaleFinish As Long
Dim ScaleKPosCorrection As Integer 'PCNGL100103
Dim ReportPageOffset As Long 'PCNGL150103
Dim FileLoadError As Boolean 'PCN2463
Dim ScaleBackColour As Long 'PCN2639
Dim GraphSet As Integer 'PCN2971



Exit Sub
Err_Handler:
Select Case Err 'PCN2639
    Case 9 'Out of subscript range
        ScaleValue = -2000
        Resume Next
    Case Else
        MsgBox Err & error$
End Select
End Sub




Sub DrawYLines(Dist As Integer)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'DrawYLines Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    04/12/02     Building initial framework
'
'Description:
'
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim quarter As Long  'This is to avoid losing the small fraction that is lost with integer
Dim whole As Long
Dim I As Long 'PCNGL121202

Exit Sub
Err_Handler:
    MsgBox Err & error$
End Sub


Sub DrawPVXScale()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'DrawPVXScale Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    13/12/02     Building initial framework
'
'Description:
'
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    
Dim LeftScale As Double
Dim CentreScale As Double
Dim RightScale As Double
Dim DisplayUnits As String
Dim GraphType As String

Dim Index As Integer
       

PrecisionVisionGraph.PVScaleClock(5).Visible = True 'PCNGL181202
'PCN3402 (11 April 2005, Antony van Iersel)



    
For Index = 0 To 5
    GraphType = ImageGraphState(Index).GraphType
    If GraphType = "Flat" Then
        PrecisionVisionGraph.PVXScaleLabelUnits(Index).Visible = True
        PrecisionVisionGraph.PVXScaleLabelL(Index).Visible = True
        PrecisionVisionGraph.PVXScaleLabelR(Index).Visible = True
        PrecisionVisionGraph.PVXScaleLabelC(Index).Visible = True
        PrecisionVisionGraph.PVXScaleLabelUnits(Index).Visible = False 'PCNGL181202
        PrecisionVisionGraph.PVScaleClock(Index).Visible = True  'PCNGL181202
        PrecisionVisionGraph.PVXScaleLabelL(Index) = 9 'PCN2122
        PrecisionVisionGraph.PVXScaleLabelR(Index) = 3 'PCN2122
        PrecisionVisionGraph.PVXScaleLabelC(Index) = 12 'PCNGL030103
    Else
    
        PrecisionVisionGraph.PVScaleClock(Index).Visible = False
        PrecisionVisionGraph.PVXScaleLabelUnits(Index).Visible = True
    
        LeftScale = PVGraphXScaleUnitFromTwip(Index, PrecisionVisionGraph.PVXScale(Index).width / 4)
        CentreScale = PVGraphXScaleUnitFromTwip(Index, PrecisionVisionGraph.PVXScale(Index).width / 2) 'ZeroMarker(Index).X1)
        RightScale = PVGraphXScaleUnitFromTwip(Index, PrecisionVisionGraph.PVXScale(Index).width / 4 * 3)
        
            
        PrecisionVisionGraph.PVXScale6(Index).X1 = PVGraphXScaleTwipFromUnit(Index, RightScale)
        PrecisionVisionGraph.PVXScale6(Index).X2 = PVGraphXScaleTwipFromUnit(Index, RightScale)
        
        PrecisionVisionGraph.PVXScale2(Index).X1 = PVGraphXScaleTwipFromUnit(Index, LeftScale)
        PrecisionVisionGraph.PVXScale2(Index).X2 = PVGraphXScaleTwipFromUnit(Index, LeftScale)
        
        PrecisionVisionGraph.PVXScaleZeroMarker(Index).X1 = PVGraphXScaleTwipFromUnit(Index, CentreScale)
        PrecisionVisionGraph.PVXScaleZeroMarker(Index).X2 = PVGraphXScaleTwipFromUnit(Index, CentreScale)
        
        PrecisionVisionGraph.PVXScaleLabelUnits(Index).Visible = True
        
        LeftScale = ConvertUnitByGraph(LeftScale, Index, DisplayUnits) 'PCN2829
        CentreScale = ConvertUnitByGraph(CentreScale, Index, DisplayUnits)
        RightScale = ConvertUnitByGraph(RightScale, Index, DisplayUnits) 'PCN2829
            
        PrecisionVisionGraph.PVXScaleLabelUnits(Index).Caption = DisplayUnits
            
        Call AlignLabel(PrecisionVisionGraph.PVXScaleLabelL(Index), LeftScale, Index) 'PCN2829
        Call AlignLabel(PrecisionVisionGraph.PVXScaleLabelC(Index), CentreScale, Index)
        Call AlignLabel(PrecisionVisionGraph.PVXScaleLabelR(Index), RightScale, Index) 'PCN2829
        
        PrecisionVisionGraph.PVXScaleLabelL(Index).Visible = True
        PrecisionVisionGraph.PVXScaleLabelR(Index).Visible = True
        PrecisionVisionGraph.PVXScaleLabelC(Index).Visible = True
    End If
Next Index

Call PrecisionVisionGraph.RepositionPVXLimitMarkers

Exit Sub
Err_Handler:
    MsgBox Err & error$

End Sub



Sub ReDimensionIndicators()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'ReDimensionIndicators Function  Louise Shrimpton louiseS@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff,    12/12/02     Building initial framework
'
'Description:
'   Check if PVGraphScreen.Height is greater than PVGraphWindow.Height
'   If so, proportion the ViewIndicators accordingly
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim ViewIndicatorRatio As Double
Dim GraphIndex As Integer
Dim ViewIndicatorHeight As Double 'PCN1855

'vvvv PCN2514 *******************************************************
'If PrecisionVisionGraph.PVGraphScreen(0).height > PrecisionVisionGraph.PVGraphWindow.height Then
'If the length of the PVGraph (not the PVGraphScreen) is greater than
'the PVGraphWindow then recalculate the size of the View Indicator.
'PVGraphTotalHeight = (PrecisionVisionGraph.PVGraphBaseCover.Top - PrecisionVisionGraph.PVGraphScreen(0).Top)
PVGraphTotalHeight = (PVDataNoOfLines + 2) * PVGraphYRatio
If PVGraphTotalHeight > PrecisionVisionGraph.PVGraphWindow.height Then
'    ViewIndicatorRatio = PrecisionVisionGraph.PVGraphWindow.height / (1.24 * PrecisionVisionGraph.PVGraphScreen(0).height)    'PCN161202 'PCNGL160503-1
'    ViewIndicatorHeight = ViewIndicatorRatio * (PrecisionVisionGraph.PipeMain.height + 400) 'PCN1601202 'PCN1855
    ViewIndicatorRatio = PrecisionVisionGraph.PVGraphWindow.height / (1.24 * PVGraphTotalHeight) 'PCN2970
    ViewIndicatorHeight = ViewIndicatorRatio * (PrecisionVisionGraph.PipeMain.height + 400)
'^^^^ ***************************************************************
Else
    ViewIndicatorHeight = PrecisionVisionGraph.PipeMain.height - 90 'PCN1855
End If
PrecisionVisionGraph.ViewIndicator(0).height = ViewIndicatorHeight 'PCN1855
PrecisionVisionGraph.ViewIndicator(1).height = ViewIndicatorHeight 'PCN1855
PrecisionVisionGraph.ViewIndicator(2).height = ViewIndicatorHeight 'PCN1855
PrecisionVisionGraph.ViewIndicator(4).height = ViewIndicatorHeight 'PCN1855
PrecisionVisionGraph.ViewIndicator(5).height = ViewIndicatorHeight 'PCN1855

'Move start and finish Y Markers

'For GraphIndex = 0 To 4
    PrecisionVisionGraph.YMarkerSt(0).Y1 = PrecisionVisionGraph.YMarkerSt(0).Y1 + PVGraphYRatio
    PrecisionVisionGraph.YMarkerSt(0).Y2 = PrecisionVisionGraph.YMarkerSt(0).Y1
    PrecisionVisionGraph.YMarkerFh(0).Y1 = PrecisionVisionGraph.YMarkerFh(0).Y1 + PVGraphYRatio
    PrecisionVisionGraph.YMarkerFh(0).Y2 = PrecisionVisionGraph.YMarkerFh(0).Y1
'Next GraphIndex

'vvvv PCN2970 **************************************************
' Position YScale makers
PrecisionVisionGraph.YScaleMarkerSt(0).Y1 = PrecisionVisionGraph.YMarkerSt(0).Y1
PrecisionVisionGraph.YScaleMarkerSt(0).Y2 = PrecisionVisionGraph.YMarkerSt(0).Y1
PrecisionVisionGraph.YScaleMarkerFh(0).Y1 = PrecisionVisionGraph.YMarkerFh(0).Y1
PrecisionVisionGraph.YScaleMarkerFh(0).Y2 = PrecisionVisionGraph.YMarkerFh(0).Y1
PrecisionVisionGraph.PVScaleMarkerStAdjuster(0).Top = PrecisionVisionGraph.YScaleMarkerSt(0).Y1 - PrecisionVisionGraph.PVScaleMarkerStAdjuster(0).height
PrecisionVisionGraph.PVScaleMarkerFhAdjuster(0).Top = PrecisionVisionGraph.YScaleMarkerFh(0).Y1 + 1
'^^^^ **********************************************************

Exit Sub
Err_Handler:
    Select Case Err
        Case 9 'Out of subscript range
            Exit Sub
        Case Else
            MsgBox Err & error$
    End Select
End Sub

Sub PVFlat3DCalcCPP(ByVal FromFrame As Long, ByVal ToFrame As Long)
On Error GoTo Err_Handler

Dim ExpRad As Integer
Dim Egnore(180) As Long
Dim I As Integer




If ExpectedDiameter <= 0 Then Exit Sub  'PCNGL100301
If FromFrame < 1 Then FromFrame = 1
If ToFrame > PVDataNoOfLines Then ToFrame = PVDataNoOfLines

'For I = 0 To 180
'    Egnore(I) = WaterEgnoreList(I)
'Next I

'Update the flat graph if it is selected


Call clearline_CalculateFlat3d(TD_PVDataX(0), _
                           TD_PVDataY(0), _
                           TD_PVCentreX(1), _
                           TD_PVCentreY(1), _
                           Egnore(0), _
                           PVDataXYMultiplier, _
                           PVFlat3DRed(0, 0), _
                           PVFlat3DGreen(0, 0), _
                           PVFlat3DBlue(0, 0), _
                           ExpectedDiameter, _
                           Flat3dLimitR, _
                           Flat3dLimitL, _
                           FromFrame, _
                           ToFrame, _
                           -ShapeCentreX * PVDrawScreenRatio, _
                           ShapeCentreY * PVDrawScreenRatio, _
                           ReferenceShape(GetNumShapeType(DrawShapeType)), _
                           ShapeRotationAngle) 'PCN3513 'PCN3567 Offset X and Y added
        
If ImageGraphState(0).GraphType = "Flat" Then
        Call DrawGraphImage(ImageGraphState(0), ImageGraphState(0).GraphType, 0, GraphStartFrame, GraphEndFrame, PVGraphGeneralXScale, PVGraphGeneralXOffset, PVXScaleLimitPerL, PVXScaleLimitPerR)
End If

If FromFrame = 1 And ToFrame = PVDataNoOfLines Then
    ClearLineScreen.UpdateThreeDimColour
End If

Exit Sub
Err_Handler:
Select Case Err 'PCNGL270103
    Case 9 'Out of subscript range
        Exit Sub
    Case 6: Resume Next ' overflow
    Case Else
        MsgBox Err & error$
End Select
End Sub
Sub PVCapacityCalcCPP(ByVal FromFrame As Long, ByVal ToFrame As Long)
On Error GoTo Err_Handler

Dim ExpRad As Integer
Dim ExpDia As Double

If ExpectedDiameter <= 0 Then Exit Sub 'PCNGL100301

If FromFrame < 1 Then FromFrame = 1
If ToFrame > PVDataNoOfLines Then ToFrame = PVDataNoOfLines
Call clearline_CalculateCapacity(TD_PVDataX(0), _
                           TD_PVDataY(0), _
                           TD_PVCentreX(1), _
                           TD_PVCentreY(1), _
                           WaterEgnoreList(0), _
                           PVCapacityFullData(1), _
                           PVDataXYMultiplier, _
                           ExpectedDiameter, _
                           FromFrame, _
                           ToFrame)  'PCN3513


Exit Sub
Err_Handler:
Select Case Err 'PCNGL270103
    Case 9 'Out of subscript range
        Exit Sub
    Case Else
        MsgBox Err & error$
End Select
End Sub

Sub PVOvalityCalcCPP(ByVal FromFrame As Long, ByVal ToFrame As Long)
On Error GoTo Err_Handler
Dim I As Long

If FromFrame < 1 Then FromFrame = 1
If ToFrame > PVDataNoOfLines Then ToFrame = PVDataNoOfLines
Call clearline_CalculateOvality(TD_PVDataX(0), _
                           TD_PVDataY(0), _
                           TD_PVCentreX(1), _
                           TD_PVCentreY(1), _
                           WaterEgnoreList(0), _
                           PVOvalityFullData(1), _
                           PVDataXYMultiplier, _
                           FromFrame, _
                           ToFrame)  'PCN3513

Exit Sub
Err_Handler:
Select Case Err 'PCNGL270103
    Case 9 'Out of subscript range
        Exit Sub
    Case Else
        MsgBox Err & error$
End Select
End Sub

Sub PVDebugOvalityCalcCPP(ByVal Frame As Long)
On Error GoTo Err_Handler

If Frame < 1 Then Frame = 1
If Frame > PVDataNoOfLines Then Frame = PVDataNoOfLines
Call clearline_CalculateDebugOvality(TD_PVDataX(0), _
                           TD_PVDataY(0), _
                           TD_PVCentreX(1), _
                           TD_PVCentreY(1), _
                           WaterEgnoreList(0), _
                           PVOvalityFullData(1), _
                           PVDataXYMultiplier, _
                           Frame, _
                           ClearLineScreen.PVScreen.hwnd, _
                           ClearLineScreen.PVScreen.width, _
                           ClearLineScreen.PVScreen.height, _
                           PVDrawScreenRatio)  'PCN3513
Exit Sub
Err_Handler:
Select Case Err 'PCNGL270103
    Case 9 'Out of subscript range
        Exit Sub
    Case Else
        MsgBox Err & error$
End Select
End Sub

Sub PVXYDiameterCalcCPP(ByVal FromFrame As Long, ByVal ToFrame As Long)
On Error GoTo Err_Handler
Dim Egnore(180) As Long
Dim I As Integer

If FromFrame < 1 Then FromFrame = 1
If ToFrame > PVDataNoOfLines Then ToFrame = PVDataNoOfLines

'For I = 0 To 180
'    Egnore(I) = WaterEgnoreList(I)
'Next I

Call clearline_CalculateXYDiameter(TD_PVDataX(0), _
                           TD_PVDataY(0), _
                           TD_PVCentreX(1), _
                           TD_PVCentreY(1), _
                           Egnore(0), _
                           PVXDiameterFullData(1), _
                           PVYDiameterFullData(1), _
                           PVDataXYMultiplier, _
                           FromFrame, _
                           ToFrame)  'PCN3513
Exit Sub
Err_Handler:
Select Case Err 'PCNGL270103
    Case 9 'Out of subscript range
        Exit Sub
    Case Else
        MsgBox Err & error$
End Select
End Sub

Sub PVDeltaMaxMinCalcCPP(ByVal FromFrame As Long, ByVal ToFrame As Long)
On Error GoTo Err_Handler
Dim Egnore(180) As Long
Dim I As Integer

If FromFrame < 1 Then FromFrame = 1
If ToFrame > PVDataNoOfLines Then ToFrame = PVDataNoOfLines

'For I = 0 To 180
'    Egnore(I) = WaterEgnoreList(I)
'Next I

Call clearline_CalculateDeltaMaxMin(TD_PVDataX(0), _
                           TD_PVDataY(0), _
                           TD_PVCentreX(1), _
                           TD_PVCentreY(1), _
                           Egnore(0), _
                           PVDeltaFullMax(1), _
                           PVDeltaFullMin(1), _
                           PVDeltaSegFullMax(1), _
                           PVDeltaSegFullMin(1), _
                           PVDataXYMultiplier, _
                           FromFrame, _
                           ToFrame)  'PCN3513
Exit Sub
Err_Handler:
Select Case Err 'PCNGL270103
    Case 9 'Out of subscript range
        Exit Sub
    Case Else
        MsgBox Err & error$
End Select
End Sub

Sub PVDiameterMaxMinCalcCPP(ByVal FromFrame As Long, ByVal ToFrame As Long)
On Error GoTo Err_Handler
Dim Egnore(180) As Long
Dim I As Integer

If FromFrame < 1 Then FromFrame = 1
If ToFrame > PVDataNoOfLines Then ToFrame = PVDataNoOfLines

'For I = 0 To 180
'    Egnore(I) = WaterEgnoreList(I)
'Next I

Call clearline_CalculateDiameterMaxMin(TD_PVDataX(0), _
                           TD_PVDataY(0), _
                           TD_PVCentreX(1), _
                           TD_PVCentreY(1), _
                           Egnore(0), _
                           PVDiameterFullMax(1), _
                           PVDiameterFullMin(1), _
                           PVDiameterSegFullMax(1), _
                           PVDiameterSegFullMin(1), _
                           PVDataXYMultiplier, _
                           FromFrame, _
                           ToFrame)  'PCN3513
Exit Sub
Err_Handler:
Select Case Err 'PCNGL270103
    Case 9 'Out of subscript range
        Exit Sub
    Case Else
        MsgBox Err & error$
End Select
End Sub

Sub PVDiameterMedianCalcCPP(ByVal FromFrame As Long, ByVal ToFrame As Long)
On Error GoTo Err_Handler

Dim Egnore(180) As Long
Dim I As Integer

If FromFrame < 1 Then FromFrame = 1
If ToFrame > PVDataNoOfLines Then ToFrame = PVDataNoOfLines

'For I = 0 To 180
'    Egnore(I) = WaterEgnoreList(I)
'Next I

Call clearline_CalculateDiameterMedian(TD_PVDataX(0), _
                           TD_PVDataY(0), _
                           TD_PVCentreX(1), _
                           TD_PVCentreY(1), _
                           Egnore(0), _
                           PVDiameterMedian(1), _
                           PVDataXYMultiplier, _
                           FromFrame, _
                           ToFrame)  'PCN3513
Exit Sub
Err_Handler:
Select Case Err 'PCNGL270103
    Case 9 'Out of subscript range
        Exit Sub
    Case Else
        MsgBox Err & error$
End Select
End Sub

'Sub PVFractileCalcCPP(ByVal FromFrame As Long, ByVal ToFrame As Long)
'On Error GoTo Err_Handler
'
'Dim Egnore(180) As Long
'Dim i As Integer
'
'If FromFrame < 1 Then FromFrame = 1
'If ToFrame > PVDataNoOfLines Then ToFrame = PVDataNoOfLines
'
''For I = 0 To 180
''    Egnore(I) = WaterEgnoreList(I)
''Next I
'
'Call clearline_CalculateFractile(TD_PVDataX(0), _
'                           TD_PVDataY(0), _
'                           TD_PVCentreX(1), _
'                           TD_PVCentreY(1), _
'                           Egnore(0), _
'                           PVFractile(1), _
'                           PVDataXYMultiplier, _
'                           FromFrame, _
'                           ToFrame)  'PCN3513
'Exit Sub
'Err_Handler:
'Select Case Err 'PCNGL270103
'    Case 9 'Out of subscript range
'        Exit Sub
'    Case Else
'        MsgBox Err & error$
'End Select
'End Sub

Sub PVCentreCalcCPP(ByVal FromFrame As Long, ByVal ToFrame As Long)
On Error GoTo Err_Handler

Dim ExpRad As Integer
Dim ExpDia As Double


If ExpectedDiameter <= 0 Then Exit Sub 'PCNGL100301

If FromFrame < 1 Then FromFrame = 1
If ToFrame > PVDataNoOfLines Then ToFrame = PVDataNoOfLines
Call Clearline_CalculateCentre(TD_PVDataX(0), _
                           TD_PVDataY(0), _
                           TD_PVCentreX(1), _
                           TD_PVCentreY(1), _
                           WaterEgnoreList(0), _
                           FromFrame, _
                           ToFrame)  'PCN3513
Exit Sub
Err_Handler:
Select Case Err 'PCNGL270103
    Case 9 'Out of subscript range
        Exit Sub
    Case Else
        MsgBox Err & error$
End Select
End Sub

Sub PVAutoRotateCPP(ByVal FromFrame As Long, ByVal ToFrame As Long, Optional Feedback)
On Error GoTo Err_Handler

If PVDataNoOfLines <= 1 Then Exit Sub
If FromFrame < 1 Then FromFrame = 1
If ToFrame > PVDataNoOfLines Then ToFrame = PVDataNoOfLines


If IsMissing(Feedback) Then
Call clearline_AutoRotate(ReferenceShape(GetNumShapeType(DrawShapeType)), _
                          ExpectedDiameter / 2, _
                          ShapeCentreX * PVDrawScreenRatio, _
                          ShapeCentreY * PVDrawScreenRatio, _
                          ShapeRotationAngle, _
                          TD_PVDataX(0), _
                          TD_PVDataY(0), _
                          TD_PVCentreX(1), _
                          TD_PVCentreY(1), _
                          FromFrame, _
                          ToFrame, _
                          WaterEgnoreList(0), _
                          0, _
                          ClearLineScreen.PVScreen.width, _
                          ClearLineScreen.PVScreen.height, _
                          PVDrawScreenRatio)
Else
Call clearline_AutoRotate(ReferenceShape(GetNumShapeType(DrawShapeType)), _
                          ExpectedDiameter / 2, _
                          ShapeCentreX * PVDrawScreenRatio, _
                          ShapeCentreY * PVDrawScreenRatio, _
                          ShapeRotationAngle, _
                          TD_PVDataX(0), _
                          TD_PVDataY(0), _
                          TD_PVCentreX(1), _
                          TD_PVCentreY(1), _
                          FromFrame, _
                          ToFrame, _
                          WaterEgnoreList(0), _
                          ClearLineScreen.PVScreen.hwnd, _
                          ClearLineScreen.PVScreen.width, _
                          ClearLineScreen.PVScreen.height, _
                          PVDrawScreenRatio)
End If
                                                              
Exit Sub
Err_Handler:
Select Case Err 'PCNGL270103
    Case 9 'Out of subscript range
        Exit Sub
    Case Else
        MsgBox Err & error$
End Select
End Sub

Sub SmoothOutOvality()
On Error GoTo Err_Handler

Dim ExpRad As Integer
Dim ExpDia As Double

If PVDataNoOfLines <= 1 Then Exit Sub

Call clearLine_SmoothOutGraphSingle(PVOvalityFullData(0), PVDataNoOfLines)
Exit Sub
Err_Handler:
Select Case Err 'PCNGL270103
    Case 9 'Out of subscript range
        Exit Sub
    Case Else
        MsgBox Err & error$
End Select
End Sub

Function PVFlat3DCalcNewShape(Frame As Long, FrameBufferNo As Integer)  'PCNGL170103 'PCN2970
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PVFlat3DCalc Function  Louise Shrimpton LouiseS@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    11/12/02     Building initial framework
'
'Description:
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
'Want to start at Data 0 and step through the array looking at the radius and the expected
'radius
Dim ExpRad As Integer
Dim GradeOfColour As Integer
Dim I As Integer
Dim NormalisedPercent As Single 'PCN2680
Dim DeltaLimitPercent As Single 'PCN2680
Dim DeltaLimitPerL As Single 'PCN2680
Dim DeltaLimitPerR As Single 'PCN2680
Dim ShapeNormalised As Double

If ExpectedDiameter = 0 Then Exit Function 'PCNGL100301
ExpRad = ExpectedDiameter / 2

'PCN3098
'PCN6004 cancles out the above pcn, there is no longer in difference between storage of mm and inch
'If MeasurementUnits = "In" Or MeasurementUnits = "in" Or MeasurementUnits = "inches" Then
'    ExpRad = (ExpRad * 100)
'End If

Call PrecisionVisionGraph.GetPVXLimits_Delta(DeltaLimitPerL, DeltaLimitPerR) 'PCN2680

For I = 1 To 180
    ShapeNormalised = ShapeTypeNormalisedPercent(ExpRad, I, FrameBufferNo)
    If pvData(I, 0, FrameBufferNo) = 0 Or ShapeNormalised = -999999 Then  ' It is probably undetected (should be black anyway)
        PVFlat3DRed(I, Frame) = 0
        PVFlat3DGreen(I, Frame) = 0
        PVFlat3DBlue(I, Frame) = 0
    Else    'PCNant????
        
        
        If ShapeNormalised < 0 Then  ' it is more Green
            On Error Resume Next 'PCN2777
            'NormalisedPercent = Round((100 * (ExpRad + ShapeNormalised) / ExpRad), 1) 'PCNGL140103 'PCN2680
             NormalisedPercent = -100 * (ShapeNormalised / (ExpRad * 2))
            
            On Error GoTo Err_Handler
            DeltaLimitPercent = Abs(DeltaLimitPerL)
     '      If NormalisedPercent > DeltaLimitPercent Then  'PCN2680
     '       DeltaLimitPercent = 100
            If NormalisedPercent > DeltaLimitPercent Then  'PCN2680
                'Blue
                PVFlat3DRed(I, Frame) = 40 '65 '0
                PVFlat3DGreen(I, Frame) = 73 '120 '0
                PVFlat3DBlue(I, Frame) = 111 '180 '255
            ElseIf NormalisedPercent > (2 * DeltaLimitPercent / 3) Then
                'Aqua (light Blue)
                PVFlat3DRed(I, Frame) = 90 '100 '120 '150 '85
                PVFlat3DGreen(I, Frame) = 155 '160 '173 '190 '150
                PVFlat3DBlue(I, Frame) = 204 '206 '213 '220 '200
            ElseIf NormalisedPercent > (DeltaLimitPercent / 3) Then
                'Green
                PVFlat3DRed(I, Frame) = 181 '200 '225 '130
                PVFlat3DGreen(I, Frame) = 224 '232 '242 '255
                PVFlat3DBlue(I, Frame) = 238 '242 '248 '158
            Else
                'White
                PVFlat3DRed(I, Frame) = 255
                PVFlat3DGreen(I, Frame) = 255
                PVFlat3DBlue(I, Frame) = 255
            End If
        ElseIf ShapeNormalised > 0 Then ' it is more blue
            On Error Resume Next
            'NormalisedPercent = 100 * (PVData(i, 0, FrameBufferNo) - ExpRad) / ExpRad
            NormalisedPercent = 100 * (ShapeNormalised / (ExpRad * 2))
            
            On Error GoTo Err_Handler
            DeltaLimitPercent = Abs(DeltaLimitPerR)
            If NormalisedPercent > DeltaLimitPercent Then  'PCN2680
                'Red
                PVFlat3DRed(I, Frame) = 255
                PVFlat3DGreen(I, Frame) = 0
                PVFlat3DBlue(I, Frame) = 0
            ElseIf NormalisedPercent > (2 * DeltaLimitPercent / 3) Then
                'Orange '  Yello - 1st set of numbers water care, secound set original
                PVFlat3DRed(I, Frame) = 255 '255 'PCN3055
                PVFlat3DGreen(I, Frame) = 255 '100 'PCN3055
                PVFlat3DBlue(I, Frame) = 0 ' 0
            ElseIf NormalisedPercent > (DeltaLimitPercent / 3) Then
                'Yellow 'Green '1st set of numbere Water Care, secound set original
                PVFlat3DRed(I, Frame) = 100  '255 '
                PVFlat3DGreen(I, Frame) = 255 '255 '
                PVFlat3DBlue(I, Frame) = 0 ' 0
            Else
                'White
                PVFlat3DRed(I, Frame) = 255
                PVFlat3DGreen(I, Frame) = 255
                PVFlat3DBlue(I, Frame) = 255
            End If
        Else  ' it is exactly the expected radius so it is white
            PVFlat3DRed(I, Frame) = 255
            PVFlat3DGreen(I, Frame) = 255
            PVFlat3DBlue(I, Frame) = 255
        End If
    End If
Next I
    
'Flag that the PVGraph has been calculated 'PCN2970
MaxCalculatedFrameNo = Frame 'PCN2970
   
Exit Function
Err_Handler:
    MsgBox Err & error$
End Function

Sub AlignLabel(LabelName As Label, NumNew As Double, Index As Integer)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'DrawVertNumber Function  Louise Shrimpton LouiseS@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    11/12/02     Building initial framework
'
'Description: This function takes a label and an integer.  This will
'       re-align the label to allow for the number of digits and
'       will place the new number in the label's caption
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim Center As Integer 'PCN1916
Dim NumTemp As Integer
Dim NumOld As Integer
Dim DigitsNew As Integer
Dim DigitsOld As Integer
Dim UnitsIndex As Integer 'PCN4207

UnitsIndex = GetGraphInfoIndex(Index) 'PCN4207

If GraphInfoContainer(UnitsIndex).PVXScaleUnits = "Real" Then 'PCN4207
    If MeasurementUnits = "mm" Then
        LabelName = Format(NumNew, "#0")
    Else
        LabelName = Format(NumNew, "#0.0")
    End If
Else
    ''If NumNew > 9 Then LabelName = Format(NumNew, "#0.0")
    ''If NumNew > 99 Then LabelName = Format(NumNew, "#0.0")
    If NumNew > 999 Then
        LabelName = Format(NumNew, "#0000")
    Else
        LabelName = Format(NumNew, "#0.0")
    End If
End If

Exit Sub
Err_Handler:
    MsgBox Err & error$
End Sub


Sub ClearAllPVGraphs()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'ClearAllPVGraphs Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    24/12/02     Building initial framework
'
'Description:
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim GraphIndex As Integer


Call PrecisionVisionGraph.PipeDisplay_MouseMove(1, 0, 25, 25) 'Move viewindicator to top 'PCNGL181202

Exit Sub
Err_Handler:
    MsgBox Err & error$
End Sub

Sub InitilisePVProfile(FrameNo As Integer) 'PCNGL241202 'PCNGL140103
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'InitilisePVProfile Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    24/12/02     Building initial framework
'
'Description:
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
If mediatype = "Live" Then Exit Sub      'PCN???? if Live then dont need to IntiialisePVProfile

Dim GraphIndex As Integer
Dim StandardPVGraphYRatio As Integer 'PCNGL0601032
Dim MaxOnePageGraphHeight As Long 'PCNGL0601032
Dim FrameNoPVGraphYRatio As Long 'PCNGL140103

'PCNGL140103 Changed these arrays to integer
ReDim pvData(NoOfProfileSegments, 2, FrameNo) 'PCNGL060103 'PCNGL110103 'PCN4006 removed as interger
ReDim PVCapacityData(FrameNo) As Integer 'PCNGL021202
ReDim PVOvalityData(FrameNo) As Integer  'PCNGL101202
ReDim PVDelta(2, FrameNo) As Integer 'PCNLS111202 'PCNGL060103
ReDim PVFlat3DRed(NoOfProfileSegments, PVDataNoOfLines) As Long 'PCNGL060103 'PCNGL110103 'PCN2970
ReDim PVFlat3DGreen(NoOfProfileSegments, PVDataNoOfLines) As Long 'PCNGL060103 'PCNGL110103 'PCN2970
ReDim PVFlat3DBlue(NoOfProfileSegments, PVDataNoOfLines) As Long 'PCNGL060103 'PCNGL110103 'PCN2970
ReDim AVIFrameTime(FrameNo) As Double 'PCNGL150103

'StandardPVGraphYRatio = 50 'PCNGL0601032
StandardPVGraphYRatio = PVGraphYRatio '2 '25 'PCNGL0601032 'PCNGL170303-2
MaxOnePageGraphHeight = 150000
FrameNoPVGraphYRatio = FrameNo * StandardPVGraphYRatio
If FrameNo >= MaxOnePageGraphHeight Then 'PCNGL0601032
    PVGraphYRatio = 1
'ElseIf FrameNo * StandardPVGraphYRatio < MaxOnePageGraphHeight Then 'PCNGL060103
ElseIf FrameNoPVGraphYRatio < MaxOnePageGraphHeight Then 'PCNGL060103
    PVGraphYRatio = StandardPVGraphYRatio
Else
    PVGraphYRatio = Int(MaxOnePageGraphHeight / FrameNo) 'PCNGL070103
End If
    
    'PCNGL070103 'Load PVGraphs form 'PCNGL181202 'PCNGL060103
    'Load PrecisionVisionGraph
    'PrecisionVisionGraph.Show
    'PrecisionVisionGraph.ZOrder 0
    'DoEvents
    
    '**** Clear all PV Graphs **********************************************
    PVFrameNo = 0 ' Move to first frame 'PCN121202 'PCN3526 set it to 0 because there
                  ' is not PVD frames yet was 1
    LastRecordedFrame = 0 'PCNLS080203
    'LastRecordedMainFrame = 0 'PCNLS080203 PCN3289 not used and not to be used (3 Feb 2005, Antony)
    PipeDisplayMoveLastY = PrecisionVisionGraph.ViewIndicator.Item(0).Top 'PCNGL181202
    'Call ClearAllPVGraphs   'PCNGL241202 'PCNGL140103
    Call ClearAllGraphsAndRuler   'PCNGL241202 'PCNGL140103
    '***********************************************************************
    
    'DoEvents 'PCNGL070103
    
Exit Sub
Err_Handler:
    MsgBox Err & error$
    
End Sub



Sub DrawProfilesStartToFinish(PVScreenCtrl As Control, Permanant As Boolean, Optional FrameNo) 'Moved from PrecisionVisionGraph form 'PCNGL030103
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'DrawProfilesStartToFinish function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    5/12/02     Building initial framework

'   v0.1 Antony van Iersel 18 May 2005
'   PCN3526, no longer loading from a file, loading from a array. (RapidReadPVDData)
'
'
'Description:
'       Once the start or finish adjuster has been moved, the PV Profile needs to be
'       redrawn from the start to finish markers
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim StartFrameNo As Long
Dim FinishFrameNo As Long
Dim CurrentFrameNo As Long


'Determine the relevant PVFrameNo for the current Start Marker


If ScreenDrawingType >= 1 Then
    DrawSF = ReportDummyGraphImage.width / ClearLineScreen.MainScreen.width  'PCN1835
Else
    PVScreenCtrl.AutoRedraw = True 'PCNGL200103
    PVScreenCtrl.Cls 'PCNGL200103
    DrawSF = PVScreenCtrl.width / ClearLineScreen.MainScreen.width
End If


If CLPScreenMode = PV Then Call ClearLineScreen.DrawPVGrid(PVScreenCtrl, MainScaleGrid / 8, CentreLineX, CentreLineY) 'PCNGL200103 'PCN1858
Call DrawAllReferenceShapes(PVScreenCtrl)
Call DrawSingleLines(PVScreenCtrl)
Call DrawPolyLines(PVScreenCtrl)
Call DrawWaterLevel(PVScreenCtrl)
Call DrawIntegratedArea(PVScreenCtrl)
Call DrawMaskBoxes(PVScreenCtrl, , , True)
'    Call PVGraphsKeyUpdate 'PCN2818 'PCN2990
Call PVGraphsKeyForm.PVGraphsKeyUpdate 'PCN2990
Call ClearLineScreen.OnscreenTextShow


Call ClearLineScreen.DistanceCounterTextUpdate  'PCN2928

If PVDataNoOfLines < 1 Then Exit Sub 'No profiles to draw
If CLPScreenMode = SnapShot Or CLPScreenMode = Video Then Exit Sub 'Dont need to draw on anyof these

StartFrameNo = PVScaleMarkerStFrame
FinishFrameNo = PVScaleMarkerFnFrame

If StartFrameNo < 1 Then StartFrameNo = 1
If StartFrameNo > PVDataNoOfLines Then StartFrameNo = PVDataNoOfLines
If FinishFrameNo < StartFrameNo Then FinishFrameNo = StartFrameNo
If FinishFrameNo > PVDataNoOfLines Then FinishFrameNo = PVDataNoOfLines

If Not Permanant Then ClearLineScreen.PVScreen.AutoRedraw = False

If IsMissing(FrameNo) Then
    For CurrentFrameNo = StartFrameNo To FinishFrameNo
        PVFrameNo = CurrentFrameNo
        Call RapidReadPVData(CurrentFrameNo)
        Call DrawPVProfile(PVScreenCtrl, 1)  'Goto frame PVFrameNo
    Next CurrentFrameNo
Else
    If FrameNo < 1 Then FrameNo = 1
    If FrameNo > PVDataNoOfLines Then FrameNo = PVDataNoOfLines
    Call RapidReadPVData(FrameNo)
    Call DrawPVProfile(PVScreenCtrl, 1)
End If

If Permanant Then ClearLineScreen.PVScreen.AutoRedraw = False

'Update PV frame status bar
If StartFrameNo = FinishFrameNo Then
    ClearLineScreen.PVFrameStatus.Caption = StartFrameNo & DisplayMessage(" of ") & PVDataNoOfLines 'PCN2111
Else
    ClearLineScreen.PVFrameStatus.Caption = StartFrameNo & DisplayMessage(" to ") & FinishFrameNo & DisplayMessage(" of ") & PVDataNoOfLines 'PCN2111
End If
'Change width of PVFrameStatus caption
ClearLineScreen.PVFrameStatus.width = Len(ClearLineScreen.PVFrameStatus.Caption) * 7 'PCNGL160103
    
Exit Sub
Err_Handler:
    MsgBox error$
    

End Sub


Sub DrawMainScale(PVScreenCtrl As Control) 'PCNGL300103
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'DrawMainScale Function
'
'Revision history"
'   V0.0    Louise,    7/1/02     Building initial framework
'
'Description:  This will initially draw and continue to redraw the Main Scale
'           on the Clearline screen, whenever it is called.
'
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim incre As Integer
Dim ScaleCentreLineX As Double
Dim ScaleCentreLineY As Double

'Call InitialiseNumberPicArray PCN3219

If CLPScreenMode = SnapShot Then 'PCNGL300103
    ScaleCentreLineX = VideoCentreLineX
    ScaleCentreLineY = VideoCentreLineY 'PCNGL300103
Else
    ScaleCentreLineX = CentreLineX
    ScaleCentreLineY = CentreLineY 'PCNGL300103
End If

If ScaleCentreLineX = 0 Then
    ScaleCentreLineX = ClearLineScreen.MainScreen.width / 2 'PCNGL140103
End If
If ScaleCentreLineY = 0 Then 'PCNGL150103
    ScaleCentreLineY = ClearLineScreen.MainScreen.height / 2
End If


If ScreenDrawingType >= 1 Then 'PCN3691
    Set ScreenDrawing.ReportDummyGraphImage = ScreenDrawing.ReportGraphImageX
    incre = DrawXScale(PVScreenCtrl, ScaleCentreLineX, ConfigInfo.Ratio) 'PCN3035
Else
    ClearLineScreen.PVXScale.Cls
    incre = DrawXScale(ClearLineScreen.PVXScale, ScaleCentreLineX, ConfigInfo.Ratio) 'PCN3035
End If

'MainScaleGrid = incre 'Used for setup PVScreen grid 'PCNGL200103
'If Len(ClearLineScreen.Ratio) <> 0 And ClearLineScreen.Ratio <> "0" Then 'PCN1858
If ConfigInfo.Ratio <> 0 Then  'PCN1858 'PCN3035
    MainScaleGrid = incre / ConfigInfo.Ratio 'Used for setup PVScreen grid 'PCNGL200103 'PCN3035
Else
    MainScaleGrid = incre / 0.9
End If


If ScreenDrawingType >= 1 Then
    Set ScreenDrawing.ReportDummyGraphImage = ScreenDrawing.ReportGraphImageY
    Call DrawYScale(PVScreenCtrl, ScaleCentreLineY, ConfigInfo.Ratio, incre) 'PCN3035
Else
    ClearLineScreen.PVYScale.Cls
    Call DrawYScale(ClearLineScreen.PVYScale, ScaleCentreLineY, ConfigInfo.Ratio, incre) 'PCN3035
End If


Exit Sub
Err_Handler:
    Select Case Err
        Case 9 'Out of subscript range
            Exit Sub
        Case Else
            MsgBox Err & error$
    End Select
End Sub

Sub DrawReferenceShape(screenctrl As Control, ByVal ShapeType As Integer, ByVal CentreX As Single, ByVal CentreY As Single, MainRadius As Double, ShapeColour As Long)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim Unit As Double

Dim XCoord As Double
Dim YCoord As Double
Dim XCoordOffset As Double
Dim YCoordOffset As Double
Dim Radius As Double
Dim ArcStart As Double
Dim ArcEnd As Double
Dim I As Integer
Dim numberOfArcs As Double

CentreX = CentreX + DrawingCentreX
CentreY = CentreY + DrawingCentreY

Dim DegToRad As Double

If UBound(ReferenceShape) < ShapeType Then Exit Sub
If MainRadius = 0 Then Exit Sub

DegToRad = PI / 180

Unit = MainRadius / PVDrawScreenRatio


XCoordOffset = ReferenceShape(ShapeType).CentreOffsetX * Unit
YCoordOffset = ReferenceShape(ShapeType).CentreOffsetY * Unit * -1

numberOfArcs = ReferenceShape(ShapeType).NoArcs - 1

If numberOfArcs >= 0 Then
    For I = 0 To numberOfArcs
        XCoord = (ReferenceShape(ShapeType).Arcs(I).OriginX * Unit)
        YCoord = (ReferenceShape(ShapeType).Arcs(I).OriginY * Unit)
        Radius = ReferenceShape(ShapeType).Arcs(I).Radius * Unit
        ArcStart = ReferenceShape(ShapeType).Arcs(I).StartAngle * DegToRad
        ArcEnd = ReferenceShape(ShapeType).Arcs(I).EndAngle * DegToRad
        XCoord = XCoord + CentreX + XCoordOffset
        YCoord = (YCoord * -1) + CentreY + YCoordOffset
        
        Call RotateCoordXY(XCoord, YCoord, CDbl(CentreX), CDbl(CentreY), ShapeRotationAngle)
        ArcStart = ArcStart - ShapeRotationAngle
        ArcEnd = ArcEnd - ShapeRotationAngle
        If ArcEnd < 0 Then ArcEnd = ArcEnd + (PI * 2)
        If ArcStart < 0 Then ArcStart = ArcStart + (PI * 2)
    '    ScreenCtrl.Circle (XCoord, YCoord), Radius, ShapeColour, ArcStart, ArcEnd
        Call RenderCircle(screenctrl, XCoord, YCoord, Radius, ShapeColour, ArcStart, ArcEnd)
    Next I
End If

Exit Sub
Err_Handler:
Select Case Err
    Case 11 ' divide by zero
        MsgBox Err & " - " & error$
'        ab = ab + 0.001
    Case 6 'overflow - this is caused when center point and a point in the circumference is same.
    Case 5 '
        'MsgBox "Temporary message after mouseup. Please check your drawing is as intended."
    Case 53 'Could not find the mouse icon
        Resume Next
    Case Else
        MsgBox Err & " - " & error$

End Select
End Sub


'
Function DrawXScale(PicCntrlName As Control, PVDrawCentreX As Double, PVDrawScreenRatio As Double) As Integer   'PCNGL161202 'PCNGL150103
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'DrawXScale Function  Louise Shrimpton louiseS@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    11/1/02     Building initial framework
'
'Description:  This will draw the X scale on the ClearLine Screen (main screen)
'
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim lengthMM As Integer  'This is half the width of the x axis
Dim increm As Integer   'This is the increment to increment the scale
Dim buffer As Double    'This is to do with layout of the numbers
Dim Middle As Integer   'The middle of the X Axis
Dim NoneZeroDigitFlag As Boolean
Dim ScreenRatio As Single 'PCNGL140103
Dim NumberScaleOffset As Integer 'PCNGL150103
Dim ScaleHeight As Single
Dim ScaleWidth As Single


If ScreenDrawingType >= 1 Then
    ScaleHeight = ClearLineScreen.PVXScale.height
    ScaleWidth = ClearLineScreen.PVXScale.width
Else
    'ClearLineScreen.PVYScale.AutoRedraw = True
    PicCntrlName.AutoRedraw = True 'PCNGL150103
    ScaleHeight = PicCntrlName.height
    ScaleWidth = PicCntrlName.width
End If


Middle = PVDrawCentreX 'The middle of the PVScreen is CentreLineX 'PCNGL140103

'If PicCntrlName.name <> "PVXScale" Then
'    NumberScaleOffset = -15
'Else
    NumberScaleOffset = -20
'End If

    'Call DrawXAxisPixNumber(PicCntrlName, 0, Middle - 2, PicCntrlName.height + NumberScaleOffset)
    Call DrawXAxisPixNumber(PicCntrlName, 0, Middle - 2, ScaleHeight + NumberScaleOffset)
    
    If Len(PVDrawScreenRatio) = 0 Or PVDrawScreenRatio = 0 Then
        ScreenRatio = 0.9
        lengthMM = 310 'Set as default
    Else  '
        ScreenRatio = PVDrawScreenRatio
        If MeasurementUnits = "mm" Then
            lengthMM = ScaleWidth * 0.5 * ScreenRatio
            buffer = 4.5 'Initial testing ls 11/01/03
        Else
            lengthMM = ScaleWidth * 0.5 * ScreenRatio
            buffer = 4.5 * (0.63 / 16) 'Initial testing ls 11/01/03
        End If
    End If
    If lengthMM <= buffer * 25 Then
        increm = 25
    ElseIf lengthMM <= buffer * 50 Then
        increm = 50
    ElseIf lengthMM <= buffer * 100 Then
        increm = 100
    ElseIf lengthMM <= buffer * 200 Then
        increm = 200
    ElseIf lengthMM <= buffer * 250 Then
        increm = 250
    ElseIf lengthMM <= buffer * 500 Then
        increm = 500
    ElseIf lengthMM <= buffer * 1000 Then
        increm = 1000
    Else
        ScreenRatio = 0.9 * (0.63 / 16)
        lengthMM = 310
        increm = 100 'Set as default 'PCNLS180203 EMERGENCY FIX, default
        'MsgBox "Pipe too big, contact your software supplier", vbInformation 'PCNGL140103
    End If
    'DrawXScale = increm

    If MeasurementUnits = "inches" Or MeasurementUnits = "in" Then  'PCNLS16_01_03
        increm = increm * (0.63 / 16)   'convert to inches
        'ScreenRatio = ScreenRatio * (0.63 / 16)
    End If
    
    DrawXScale = increm

'Need to change the y coord (2nd parameter) to be equal to the middle of the number
'Also, negative signs for left hand side

'The right side of the scale
Dim I As Integer
    For I = 1 To 5
        Call DrawXNumber(PicCntrlName, I * increm, (Middle + ((increm * I) / ScreenRatio)))  'PCNGL140103
        Call DrawXLines(PicCntrlName, (Middle + ((increm * (I - 1)) / ScreenRatio)), (Middle + ((increm * I) / ScreenRatio)))  'PCNGL140103
    Next I
    
'The left side of the scale
    For I = 1 To 5
        Call DrawXNumber(PicCntrlName, I * increm, (Middle - ((increm * I) / ScreenRatio)))  'PCNGL140103
        Call DrawXLines(PicCntrlName, (Middle - ((increm * I) / ScreenRatio)), (Middle - ((increm * (I - 1)) / ScreenRatio)))  'PCNGL140103
    Next I

Exit Function
Err_Handler:
    MsgBox Err & error$
End Function


Sub DrawXNumber(PicCntrlName As Control, Number As Integer, ByVal YCoord As Integer)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'DrawXNumber Function  Louise Shrimpton louiseS@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    11/1/02     Building initial framework
'
'Description:
'
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim NoneZeroDigitFlag As Boolean

Dim ScaleDigitID As Integer
    'Determine the digits within ScaleValue
    NoneZeroDigitFlag = False
    
Dim ScaleDigitCode(6) As Integer 'Holds the digits within ScaleValue 'PCNGL090103
Dim ScaleNoOfDigits As Integer 'A count of the number of digits
Dim NumberScaleOffset As Integer 'PCNGL150103


Dim ScaleHeight As Single

If ScreenDrawingType >= 1 Then
    ScaleHeight = ClearLineScreen.PVXScale.height
Else
    ScaleHeight = PicCntrlName.height
End If
    

'If PicCntrlName.name <> "PVXScale" Then
'    NumberScaleOffset = -15
'Else
    NumberScaleOffset = -20
'End If
    
    
    For ScaleDigitID = 6 To 1 Step -1
        ScaleDigitCode(ScaleDigitID) = Int(Number / (10 ^ (ScaleDigitID - 1))) Mod 10
        If ScaleDigitCode(ScaleDigitID) > 0 Or NoneZeroDigitFlag Then 'Ignore right hand digits that are zero
            NoneZeroDigitFlag = True
            ScaleNoOfDigits = ScaleNoOfDigits + 1
        Else
            ScaleDigitCode(ScaleDigitID) = 10 'Nothing
            If ScaleDigitID = 1 Then 'ScaleValue = nothing, so ScaleDigitCode(0) should be nothing 'PCNGL100301
                ScaleDigitCode(0) = 10 'Nothing 'PCNGL100103
            End If
        End If
    Next ScaleDigitID
    'Draw the y scale numbers
    ScaleDigitID = 0
    'Determine middle digit of ScaleValue and then the start position of the first digit
    Dim SizeNum As Integer
    Dim Left As Integer
    Dim a As Integer
    SizeNum = 8
 'Here we need to find the position of the digits.  The middle is given by YCoord
 'Therefore, if there is:
 '1 digit, place it at YCoord - (0.5 * SizeNum)
 '2 digits, place them at YCoord, YCoord - SizeNum
 '3 digits, place at YCoord - (1.5 * SizeNum), YCoord - SizeNum, YCoord + (1.5 *SizeNum)
 '4 digits, YCoord - (2*SizeNum), YCoord-SizeNum, YCoord, YCoord + SizeNum
 '5 digits, YCoord - (2.5 * SizeNum), YCoord - (1.5 * SizeNum), YCoord - (0.5 * SizeNum), YCoord + (0.5 * SizeNum), YCoord - (1.5 * SizeNum)
    If ScaleNoOfDigits = 1 Then
        Left = YCoord - 2 '(0.5 * SizeNum)
    ElseIf ScaleNoOfDigits = 2 Then
        Left = YCoord - SizeNum
    ElseIf ScaleNoOfDigits = 3 Then
        Left = YCoord - (1.5 * SizeNum)
    ElseIf ScaleNoOfDigits = 4 Then
        Left = YCoord - (2 * SizeNum)
    ElseIf ScaleNoOfDigits = 5 Then
        Left = YCoord - (2.5 * SizeNum)
    ElseIf ScaleNoOfDigits = 2 Then
        Left = YCoord - (3 * SizeNum)
    End If
    
    For ScaleDigitID = ScaleNoOfDigits To 1 Step -1
        Call DrawXAxisPixNumber(PicCntrlName, ScaleDigitCode(ScaleDigitID), Left + a, ScaleHeight + NumberScaleOffset)
        a = a + SizeNum
    Next ScaleDigitID

Exit Sub
Err_Handler:
    MsgBox Err & error$
End Sub


Sub DrawXLines(PicCntrlName As Control, Bottom As Integer, Top As Integer)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'DrawLine Function  Louise Shrimpton louiseS@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    11/1/02     Building initial framework
'
'Description: Draws Lines from the top to the bottom, with a big one at the bottom
'           and the top
'
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim quarter As Long  'This is to avoid losing the small fraction that is lost with integer
Dim whole As Long
Dim I As Long 'PCNGL121202

Dim ScaleHeight As Single

If ScreenDrawingType >= 1 Then
    ScaleHeight = ClearLineScreen.PVXScale.height
Else
    ScaleHeight = PicCntrlName.ScaleHeight
End If


    quarter = (Top - Bottom) \ 4      'division
    whole = quarter * 4
    
    For I = Bottom To Top Step quarter 'PCNGL161202
        If I < (Top - 5) Then  'This works because there will never be more than 4 quarters
            Call RenderLine(PicCntrlName, I, ScaleHeight, I, ScaleHeight - 4, 0) 'PCN3691 'PCNGL161202 'PCNGL210103
            'PicCntrlName.Line (i, PicCntrlName.ScaleHeight)-(i, PicCntrlName.ScaleHeight - 4), 0 'PCNGL161202 'PCNGL210103
        End If
        If quarter = 0 Then
            Exit Sub
        End If
    Next I
    I = 0 'PCNGL161202
    For I = Bottom To Top Step whole  'PCNGL161202
        If I < (Top - 6) Then  'This works because there will never be more than 5 steps.
            Call RenderLine(PicCntrlName, I, ScaleHeight, I, ScaleHeight - 7, 0) 'PCN3691 'PCNGL161202 'PCNGL210103
            'PicCntrlName.Line (i, PicCntrlName.ScaleHeight)-(i, PicCntrlName.ScaleHeight - 7), 0 'PCNGL161202 'PCNGL210103
        End If
        If whole = 0 Then
            Exit Sub
        End If
    Next I
    

Exit Sub
Err_Handler:
    MsgBox Err & error$
End Sub



Sub DrawYScale(PicCntrlName As Control, PVDrawCentreY As Double, PVDrawScreenRatio As Double, incre As Integer) 'PCNGL161202 'PCNGL150103
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'DrawYScale Function  Louise Shrimpton louiseS@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    13/1/02
'
'Description:  This will draw the Y scale on the ClearLine Screen (main screen)
'
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim lengthMM As Integer  'This is half the width of the y axis
Dim increm As Integer   'This is the increment to increment the scale
Dim buffer As Double    'This is to do with layout of the numbers
Dim Middle As Integer   'The middle of the Y Axis
Dim NoneZeroDigitFlag As Boolean
Dim ScreenRatio As Single 'PCNGL140103

If ScreenDrawingType = 0 Then
    PicCntrlName.AutoRedraw = True
    Middle = PVDrawCentreY - 21   'The middle of the PVScreen is PVDrawCentreX 'PCNGL140103
Else
    Middle = PVDrawCentreY
End If

'If PicCntrlName.name <> "PVYScale" Then 'PCNGL150103
'    Middle = PVDrawCentreY   'The middle of the PVScreen is PVDrawCentreX 'PCNGL140103
'Else
'    Middle = PVDrawCentreY - 21   'The middle of the PVScreen is PVDrawCentreX 'PCNGL140103
'End If
buffer = 4.5 'Initial testing ls 13/01/03

    
Call DrawYAxisPixNumber(PicCntrlName, 0, 10, Middle - 2)

    'Call DrawYAxisPixNumber(PicCntrlName, 0, middle, 3)

    If Len(PVDrawScreenRatio) = 0 Or PVDrawScreenRatio = 0 Then
        ScreenRatio = 0.9
        lengthMM = 310 'Set as default
    Else
        ScreenRatio = PVDrawScreenRatio
        If MeasurementUnits = "mm" Then
            lengthMM = ClearLineScreen.PVScreen.width * 0.5 * ScreenRatio
        Else
            lengthMM = ClearLineScreen.PVScreen.width * 0.5 * ScreenRatio * (16 / 0.63)
        End If
    End If
    
    If lengthMM <= buffer * 25 Then
        increm = 25
    ElseIf lengthMM <= buffer * 50 Then
        increm = 50
    ElseIf lengthMM <= buffer * 100 Then
        increm = 100
    ElseIf lengthMM <= buffer * 200 Then
        increm = 200
    ElseIf lengthMM <= buffer * 250 Then
        increm = 250
    ElseIf lengthMM <= buffer * 500 Then
        increm = 500
    ElseIf lengthMM <= buffer * 1000 Then
        increm = 1000
    Else
        ScreenRatio = 0.9 * (0.63 / 16)
        lengthMM = 310
        increm = 100 'Set as default 'PCNLS180203 EMERGENCY FIX, default
        'MsgBox "Pipe too big, contact your software supplier", vbInformation 'PCNGL140103
    End If
    increm = incre

    If MeasurementUnits = "inches" Or MeasurementUnits = "in" Then  'PCNLS16_01_03
        'ScreenRatio = ScreenRatio * (0.63 / 16)
    End If

'Need to change the y coord (2nd parameter) to be equal to the middle of the number
'Also, negative signs for left hand side

'The bottom of the scale
Dim I As Integer
    For I = 1 To 5
        Call DrawYNumber(PicCntrlName, I * increm, (Middle + ((increm * I) / ScreenRatio))) 'PCNGL140103
       Call DrawMainYLines(PicCntrlName, (Middle + ((increm * (I - 1)) / ScreenRatio)), (Middle + ((increm * I) / ScreenRatio))) 'PCNGL140103
    Next I
    
'The top of the scale
    For I = 1 To 5
        Call DrawYNumber(PicCntrlName, I * increm, (Middle - ((increm * I) / ScreenRatio))) 'PCNGL140103
        Call DrawMainYLines(PicCntrlName, (Middle - ((increm * I) / ScreenRatio)), (Middle - ((increm * (I - 1)) / ScreenRatio))) 'PCNGL140103
    Next I

Exit Sub
Err_Handler:
    MsgBox Err & error$

End Sub

Sub DrawYNumber(PicCntrlName As Control, Number As Integer, ByVal YCoord As Integer) 'PCNGL150103
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'DrawYNumber Function  Louise Shrimpton louiseS@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    11/1/02     Building initial framework
'
'Description:
'
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim NoneZeroDigitFlag As Boolean
Dim ScaleDigitID As Integer
    'Determine the digits within ScaleValue
    NoneZeroDigitFlag = False
    
Dim ScaleDigitCode(6) As Integer 'Holds the digits within ScaleValue 'PCNGL090103
Dim ScaleNoOfDigits As Integer 'A count of the number of digits

    For ScaleDigitID = 6 To 1 Step -1
        ScaleDigitCode(ScaleDigitID) = Int(Number / (10 ^ (ScaleDigitID - 1))) Mod 10
        If ScaleDigitCode(ScaleDigitID) > 0 Or NoneZeroDigitFlag Then 'Ignore right hand digits that are zero
            NoneZeroDigitFlag = True
            ScaleNoOfDigits = ScaleNoOfDigits + 1
        Else
            ScaleDigitCode(ScaleDigitID) = 10 'Nothing
            If ScaleDigitID = 1 Then 'ScaleValue = nothing, so ScaleDigitCode(0) should be nothing 'PCNGL100301
                ScaleDigitCode(0) = 10 'Nothing 'PCNGL100103
            End If
        End If
    Next ScaleDigitID
    'Draw the y scale numbers
    ScaleDigitID = 0
    'Determine middle digit of ScaleValue and then the start position of the first digit
    Dim SizeNum As Integer
    Dim Left As Integer
    Dim a As Integer
    SizeNum = 8
 'Here we need to find the position of the digits.  The middle is given by YCoord
 'Therefore, if there is:
 '1 digit, place it at YCoord - (0.5 * SizeNum)
 '2 digits, place them at YCoord, YCoord - SizeNum
 '3 digits, place at YCoord - (1.5 * SizeNum), YCoord - SizeNum, YCoord + (1.5 *SizeNum)
 '4 digits, YCoord - (2*SizeNum), YCoord-SizeNum, YCoord, YCoord + SizeNum
 '5 digits, YCoord - (2.5 * SizeNum), YCoord - (1.5 * SizeNum), YCoord - (0.5 * SizeNum), YCoord + (0.5 * SizeNum), YCoord - (1.5 * SizeNum)
    If ScaleNoOfDigits = 1 Then
        Left = YCoord - 2 '(0.5 * SizeNum)  ls16_01_03
    ElseIf ScaleNoOfDigits = 2 Then
        Left = YCoord - SizeNum
    ElseIf ScaleNoOfDigits = 3 Then
        Left = YCoord - (1.5 * SizeNum)
    ElseIf ScaleNoOfDigits = 4 Then
        Left = YCoord - (2 * SizeNum)
    ElseIf ScaleNoOfDigits = 5 Then
        Left = YCoord - (2.5 * SizeNum)
    ElseIf ScaleNoOfDigits = 6 Then 'PCNGL081203
        Left = YCoord - (3 * SizeNum)
    End If
    
'Call DrawYAxisPixNumber(PicCntrlName, 0, 10, middle - 2)
    
    For ScaleDigitID = 1 To ScaleNoOfDigits Step 1
        Call DrawYAxisPixNumber(PicCntrlName, ScaleDigitCode(ScaleDigitID), 10, Left + a) 'PCNGL150103
        a = a + SizeNum
    Next ScaleDigitID

Exit Sub
Err_Handler:
    MsgBox Err & error$
End Sub


Sub DrawMainYLines(PicCntrlName As Control, Bottom As Integer, Top As Integer)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'DrawMainYLines Function  Louise Shrimpton louiseS@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    13/1/02
'
'Description: Draws Lines from the top to the bottom, with a big one at the bottom
'           and the top
'
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim quarter As Long  'This is to avoid losing the small fraction that is lost with integer
Dim whole As Long
Dim I As Long 'PCNGL121202

    quarter = (Top - Bottom) \ 4      'division
    If quarter = 0 Then quarter = 1 'PCNGL150103
    whole = quarter * 4
    
    For I = Bottom To Top Step quarter
        If I < (Top - 5) Then  'This works because there will never be more than 4 quarters
            Call RenderLine(PicCntrlName, 0, I, 4, I, 0) 'PCNGL150103 'PCNGL230103 added colour = 0
            'PicCntrlName.Line (0, i)-(4, i), 0 'PCNGL150103 'PCNGL230103 added colour = 0
        End If
    Next I
    I = 0 'PCNGL161202
    For I = Bottom To Top Step whole
        If I < (Top - 6) Then  'This works because there will never be more than 5 steps.
            Call RenderLine(PicCntrlName, 0, I, 7, I, 0) 'PCNGL150103 'PCNGL230103 added colour = 0
            'PicCntrlName.Line (0, i)-(7, i), 0 'PCNGL150103 'PCNGL230103 added colour = 0
        End If
    Next I
    

Exit Sub
Err_Handler:
    MsgBox Err & error$
End Sub

Sub DrawYAxisPixNumber(CtrlName As Control, No, X1, Y1)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'DrawYAxisPixNumber Function  Louise Shrimpton louiseS@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    13/01/03
'
'Description:
'       This function has been designed to be used for drawing the numbers on the
'       y-axis scales. However, this function could be used where ever vertical
'       numbers are required.
'Purpose:
'       CtrlName is the name of the picturebox control
'       No is the number required
'       x1 and y1 are the top left start co-ordinates of the number.
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim XBit As Integer
Dim YBit As Integer
Dim BitSize As Integer

BitSize = 15 'Approximate size of pixel in terms of twips

If No < 0 Then Exit Sub      'PCN2511

For XBit = 0 To 7
    For YBit = 0 To 4
        If NumberPic(No, XBit, YBit) = 1 Then
            Call RenderPoint(CtrlName, X1 + XBit, Y1 + YBit, 0)
            'CtrlName.PSet (X1 + XBit, Y1 + YBit), 0 'PCNGL210103 Ensured colour of numbers is black 'PCNGL091202
            'CtrlName.Line (X1 + BitSize * XBit, Y1 + BitSize * YBit)-(X1 + BitSize * (XBit + 1), Y1 + BitSize * (YBit + 1))
        End If
    Next YBit
Next XBit

Exit Sub
Err_Handler:
    MsgBox Err & error$
End Sub

Sub DrawXAxisPixNumber(CtrlName As Control, No, X1, Y1)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'DrawVertNumber Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    04/12/02     Building initial framework
'
'Description:
'       This function has been designed to be used for drawing the numbers on the
'       y-axis scales. However, this function could be used where ever vertical
'       numbers are required.
'Purpose:
'       CtrlName is the name of the picturebox control
'       No is the number required
'       x1 and y1 are the top left start co-ordinates of the number.
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim XBit As Integer
Dim YBit As Integer
Dim BitSize As Integer

BitSize = 15 'Approximate size of pixel in terms of twips

If No < 0 Then Exit Sub      'PCN2511

For XBit = 0 To 4
    For YBit = 0 To 7
        If NumberPicX(No, XBit, YBit) = 1 Then
            Call RenderPoint(CtrlName, X1 + XBit, Y1 + YBit, 0) 'PCN3691
            'CtrlName.PSet (X1 + XBit, Y1 + YBit), 0 'PCNGL210103 Ensured colour of numbers is black 'PCNGL091202
            'CtrlName.Line (X1 + BitSize * XBit, Y1 + BitSize * YBit)-(X1 + BitSize * (XBit + 1), Y1 + BitSize * (YBit + 1))
        End If
    Next YBit
Next XBit

Exit Sub
Err_Handler:
    MsgBox Err & error$
End Sub




Sub DrawMainScaleRedLines(X As Single, Y As Single)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'DrawMainScaleRedLines Function  Louise Shrimpton louises@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise,    14/01/03
'
'Description:
'       Draw little red lines on the MainScreen scale to mark where the
'       mouse is on the screen
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    ClearLineScreen.PVXScale.AutoRedraw = False
    ClearLineScreen.PVYScale.AutoRedraw = False
    ClearLineScreen.PVXScale.Cls
    ClearLineScreen.PVYScale.Cls
    ClearLineScreen.PVYScale.Line (0, Y - 21)-(ClearLineScreen.PVYScale.width, Y - 21), RGB(255, 128, 128)
    ClearLineScreen.PVXScale.Line (X, 0)-(X, ClearLineScreen.PVXScale.height), RGB(255, 128, 128)
    ClearLineScreen.PVXScale.AutoRedraw = True
    ClearLineScreen.PVYScale.AutoRedraw = True


Exit Sub
Err_Handler:
    Select Case Err
        Case 9 'Out of subscript range
            Exit Sub
        Case Else
            MsgBox Err & error$
    End Select
End Sub

Sub DrawXAxisNumber_Twips(CtrlName As Control, No, X1, Y1)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'DrawXAxisNumber_Twips Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    16/01/03     Building initial framework
'
'Description:
'       This function has been designed to be used for drawing the numbers on the
'       y-axis scales. However, this function could be used where ever vertical
'       numbers are required.
'Purpose:
'       CtrlName is the name of the picturebox control
'       No is the number required
'       x1 and y1 are the top left start co-ordinates of the number.
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim XBit As Integer
Dim YBit As Integer
Dim BitSize As Integer

BitSize = 15 'Approximate size of pixel in terms of twips

If No < 0 Then Exit Sub      'PCN2511

For XBit = 0 To 4
    For YBit = 0 To 7
        If NumberPicX(No, XBit, YBit) = 1 Then 'PCNGL160103
            CtrlName.Line (X1 + BitSize * XBit, Y1 + BitSize * YBit)-(X1 + BitSize * (XBit + 1), Y1 + BitSize * (YBit + 1)), 0
        End If
    Next YBit
Next XBit

Exit Sub
Err_Handler:
    MsgBox Err & error$
End Sub


Function GetPVFrameNoFromTime(ByVal VideoTime As Double) As Long
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'GetPVFrameNoFromTime Function  Louise Shrimpton louises@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise,    15/01/03
'
'Description:
'       Takes the time and returns the nearest PVFrame number
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim CheckFrameNo As Long
Dim FileLoadError As Boolean
Dim found As Boolean
Dim t As Double
Dim t1 As Double
Dim t2 As Double
Dim Diff As Double
Dim sampletime As Double
Dim upperbound As Long
Dim lowerbound As Long
Dim I As Integer
Dim error As Double

FileLoadError = False
found = False
CheckFrameNo = 0
Diff = 0.1

t2 = -2
t1 = -1
t = 0

error = 0.02 'PCN2930

upperbound = PVDataNoOfLines
lowerbound = 1

If PVTimes(upperbound - 1) < VideoTime Then 'There is no time at PVFrameNo = PVDataNoOfLines
    GetPVFrameNoFromTime = PVDataNoOfLines
    Exit Function
ElseIf PVTimes(lowerbound) > VideoTime Then
    GetPVFrameNoFromTime = 1
    Exit Function
End If

While found = False
    CheckFrameNo = (upperbound - lowerbound) / 2 + lowerbound
    sampletime = PVTimes(CheckFrameNo)
    If (sampletime >= (VideoTime - error)) And (sampletime <= (VideoTime + error)) Then
        found = True
    ElseIf sampletime > VideoTime Then
        If upperbound = CheckFrameNo Or upperbound = CheckFrameNo + 1 Or upperbound = CheckFrameNo + 2 Then found = True
        upperbound = CheckFrameNo
    ElseIf sampletime < VideoTime Then
        If lowerbound = CheckFrameNo Or lowerbound = CheckFrameNo - 1 Or lowerbound = CheckFrameNo + 2 Then found = True
        lowerbound = CheckFrameNo
    End If
'    found = True
Wend

'PCN2930''''' Antony 14 September 2004 once found rough frame look at neighbour for if better
Dim FinalFrameNo As Long
Dim FinalFrameTime As Double

FinalFrameNo = CheckFrameNo
FinalFrameTime = Abs(PVTimes(FinalFrameNo) - VideoTime)

If CheckFrameNo > 0 And CheckFrameNo <= PVDataNoOfLines Then
    sampletime = Abs(PVTimes(CheckFrameNo - 1) - VideoTime)
    If sampletime < FinalFrameTime Then
        FinalFrameNo = CheckFrameNo - 1
        FinalFrameTime = sampletime
    End If
    sampletime = Abs(PVTimes(CheckFrameNo + 1) - VideoTime)
    If sampletime < FinalFrameTime Then
        FinalFrameNo = CheckFrameNo + 1
        FinalFrameTime = sampletime
    End If
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


GetPVFrameNoFromTime = FinalFrameNo
    
    

Exit Function
Err_Handler:
    Select Case Err
        Case 9 'Out of subscript range
            Exit Function
        Case Else
            MsgBox Err & error
    End Select
End Function

Sub DrawAllCircles(screenctrl As Control, NoOfCircles As Long, DrawSF As Double)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'DrawAllCircles Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff,    22/01/03      Cut code from DrawAll function, will use also for PicInPic
'
'Description:
'       Is called by the DrawAll function to Draw in the selected picture box, all
'       measuring tool circles
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Exit Sub
Err_Handler:
    Select Case Err
        Case Else
            MsgBox Err & error$
    End Select
End Sub


Sub DrawAllReferenceShapes(screenctrl As Control)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'DrawRefCircles Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff,    04/02/03      Cut code from DrawAllCircles function, will use also for PicInPic 'PCN1835
'
'Description:
'       Is NOT called by the DrawAll function. Reference circles are fixed (unless manually
'       moved) and are drawn at the same time the PVProfile is drawn or the Snap-shot function is called.
'
'Purpose:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

    If ShowReferenceShape Then
        Call DrawReferenceShape(screenctrl, GetNumShapeType(DrawShapeType), ShapeCentreX, ShapeCentreY, ExpectedDiameter / 2, vbGreen)
    End If
    ShowOutsideShape = ShowReferenceShape
    If ShowOutsideShape Then
        Call DrawReferenceShape(screenctrl, GetNumShapeType(DrawShapeType), ShapeCentreX, ShapeCentreY, PipelineInfo.ExtDiameter / 2, &HFFAA)
    End If
Exit Sub
Err_Handler:
    Select Case Err
        Case Else
            MsgBox Err & error$


    End Select
End Sub
Function GetNumShapeType(ByVal DrawShapeString As String)
On Error GoTo Err_Handler

    Select Case DrawShapeString
        Case "Circle": GetNumShapeType = 0
        Case "SemiElliptical": GetNumShapeType = 1
        Case "Egg": GetNumShapeType = 2
        Case "Bullet": GetNumShapeType = 3
        Case "Barn": GetNumShapeType = 4
    End Select

Exit Function
Err_Handler:
    Select Case Err
        Case Else
            MsgBox Err & error$
    End Select
End Function

Sub DrawPVProfile_Setup(PicCntrlName As Control) 'PCN3526
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'DrawPVProfile Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    22/01/03     Building initial framework
'
'Description:
'       Sets up the scales, control size and format for the required
'       PVProfile picture box
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim InternalRadius  As Double 'PCNGL071202
Dim incre As Integer 'PCNGL150103
Dim DrawSF As Double 'PCN1863
Dim XOffset As Single
Dim YOffset As Single

If ExpectedDiameter = 0 Then 'PCN 31
    Exit Sub
Else
    InternalRadius = ExpectedDiameter / 2 'PCNGL071202
End If

'*** Setup for required screen (PVScreen or PVScreenPicInPic) 'PCNGL231202
If PicCntrlName.name = "PVScreen" Then 'PCNGL231202
    PVDrawScreenRatio = ConfigInfo.Ratio 'PCN3035
    DrawingCentreX = CentreLineX
    DrawingCentreY = CentreLineY
    XOffset = 0
    YOffset = 0
Else

    DrawSF = PicCntrlName.width / ClearLineScreen.MainScreen.width 'PCNGL301202 'PCN1863
    PicCntrlName.height = ClearLineScreen.MainScreen.height + 10  'pcngl210103 'PCN1863
    DrawingCentreX = PicCntrlName.width / 2
    DrawingCentreY = PicCntrlName.height / 2 - 5  'PCNGL240103 <- removed because of DrallAll is now added to report 'PCNGL210103
    PVDrawScreenRatio = ConfigInfo.Ratio / DrawSF 'PCNGL301202 'PCN1863 'PCN3035
    
    '^^^ ****************************
    'Draw centre lines
    'vvvv PCN2363 ********************************
    Dim CurrentAutodrawSetting As Boolean
    
    CurrentAutodrawSetting = PicCntrlName.AutoRedraw
    '^^^^ ********************************************
    PicCntrlName.AutoRedraw = True
    PicCntrlName.Line (DrawingCentreX + XOffset, 0 + YOffset)-(DrawingCentreX + XOffset, (2 * DrawingCentreY) + YOffset), RGB(255, 200, 170)
    PicCntrlName.Line (0 + XOffset, DrawingCentreY + YOffset)-((2 * DrawingCentreX) + XOffset, DrawingCentreY + YOffset), RGB(255, 200, 170)
    'Draw the X and Y scales
    incre = DrawXScale(PicCntrlName, DrawingCentreX, PVDrawScreenRatio)
    Call DrawYScale(PicCntrlName, DrawingCentreY, PVDrawScreenRatio, incre)


    PicCntrlName.AutoRedraw = CurrentAutodrawSetting 'PCN2363

End If
    

Exit Sub
Err_Handler:
    Select Case Err
        Case 9
            Exit Sub
        Case Else
            MsgBox Err & error$
    
    End Select
End Sub



Sub DrawPVProfile_XY(PicCntrlName As Control, FrameNo As Integer)    'PCNGL220103 PCN3526
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : DrawPVProfile_XY
'Created : 20 June 2004, PCN2891
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Draws the actual Profile using X and Y co-ordinate methods.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim value As Double
Dim X As Double, Y As Double
Dim SaveX As Double, SaveY As Double
Dim Ang As Integer
Dim StartAng As Double
Dim EndAng As Double
Dim Angle As Double
Dim InternalRadius  As Double 'PCNGL071202
Dim shp As Integer 'I have no idea what this is for! GL
Dim ProfileColor As Long 'PCN1916
Dim incre As Integer 'PCNGL150103
Dim CheckAspectRatio As Double 'Temporary variable for checking the current video to screen ratio 'PCN1863
Dim InternalRad As Long 'Temp variable for internal radius reference calculation 'PCNGL160503-1

'vvvv PCN2703 *************************
Dim FeatureIndicatorX1 As Integer
Dim FeatureIndicatorY1 As Integer
Dim FeatureIndicatorX2 As Integer
Dim FeatureIndicatorY2 As Integer
Dim FeatureIndicatorX3 As Integer
Dim FeatureIndicatorY3 As Integer
Dim FeatureIndicatorX4 As Integer
Dim FeatureIndicatorY4 As Integer
'^^^^ *********************************
'vvvv PCN3053 ***********************************
Dim DrawSF As Double 'Draw scale factor

Dim DataX As Single 'PCN3219
Dim DataY As Single 'PCN3219
Dim Colour1 As Boolean 'PCN3219
Dim Colour2 As Boolean 'PCN3219

If ExpectedDiameter = 0 Then 'PCN 31 'PCN3647
    Exit Sub
Else
    InternalRadius = ExpectedDiameter / 2 'PCNGL071202 'PCN3647
End If
Angle = 270
Ang = 1
shp = 1


'Check if requested frame is greater then last loaded frame
If FrameNo > PVDataNoOfLines Then Exit Sub

If ScreenDrawingType = 0 Then
'ProfileColor = NormalDrawingColor 'PCN1931
    ProfileColor = vbBlue
    PicCntrlName.DrawWidth = 1
Else
    ProfileColor = vbBlack
    PicCntrlName.DrawWidth = 4
End If

If NoOfProfileSegments = 0 Then
    NoOfProfileSegments = 180
End If

StartAng = (270 - 2) * PI / NoOfProfileSegments 'PCN1834

If StartAng > 2 * PI Then StartAng = StartAng - 2 * PI
'Convert PVData from mm or 1/100 of inch to pixels for MainScreen

Call ClearLineScreen.VideoScreenScaleCalc 'PCN2891

DataX = pvData(1, 1, FrameNo) 'PCN3219
DataY = pvData(1, 2, FrameNo) 'PCN3219

Colour1 = False
Colour2 = False

'If PVData(1, 1, FrameNo) <> 0 Or PVData(1, 2, FrameNo) <> 0 Then 'PCN3219
If DataX <> 0 Or DataY <> 0 Then
    
    If DataX <= -10000 Then DataX = DataX + 20000: Colour1 = True
    If DataX >= 10000 Then DataX = DataX - 20000: Colour1 = True
    
    
    X = DataX / PVDataXYMultiplier / (PVDrawScreenRatio) + DrawingCentreX   'PCN2988 'PCN3053
    Y = DrawingCentreY - DataY / PVDataXYMultiplier / (PVDrawScreenRatio)  'PCN2988 'PCN3053
    SaveX = X
    SaveY = Y
Else
    SaveX = 0
    SaveY = 0 'PCN3219
    Colour1 = False
End If
For Ang = NoOfProfileSegments To 1 Step -1 'PCN1834
    Colour2 = False
    DataX = pvData(Ang, 1, FrameNo) 'PCN3219
    DataY = pvData(Ang, 2, FrameNo) 'PCN3219

    'If PVData(Ang, 1, FrameNo) <> 0 Or PVData(Ang, 2, FrameNo) <> 0 Then
    If DataX <> 0 Or DataY <> 0 Then
        
'PCN6004 no more difference in storage of inches or mm
'        If MeasurementUnits <> "mm" Then 'PCNGL310103
'            value = value / 100   'PCNGL310103
'        End If
        'X = PVData(Ang, 1, FrameNo) * VideoScreenScale  / PVDataXYMultiplier + PVDrawCentreX  'PCN3053
        'Y = PVDrawCentreY - PVData(Ang, 2, FrameNo) * VideoScreenScale  / PVDataXYMultiplier 'PCN3053
        
        If DataX <= -10000 Then DataX = DataX + 20000: Colour2 = True
        If DataX >= 10000 Then DataX = DataX - 20000: Colour2 = True
        
        X = DataX / PVDataXYMultiplier / (PVDrawScreenRatio) + DrawingCentreX  'PCN3053
        Y = DrawingCentreY - pvData(Ang, 2, FrameNo) / PVDataXYMultiplier / (PVDrawScreenRatio)   'PCN3053
        
  '      PicCntrlName.PSet (X, Y), vbBlue
        If SaveX <> 0 Or SaveY <> 0 Then 'PCN3219 added SaveY
            If Colour1 Or Colour2 Then
                PicCntrlName.DrawWidth = 2
                Call RenderLine(PicCntrlName, X, Y, SaveX, SaveY, vbRed)
            Else
                PicCntrlName.DrawWidth = 1
                Call RenderLine(PicCntrlName, X, Y, SaveX, SaveY, ProfileColor)
            End If
        End If
       
        SaveX = X
        SaveY = Y
    Else
        SaveX = 0
        SaveY = 0
        Colour2 = False
    End If
    Colour1 = Colour2
    Angle = Angle + 2
Next Ang
PicCntrlName.DrawWidth = 1
'Call DrawWaterLevel
    
'vvvv PCN2703 *********************************************
If ImageGraphState(0).GraphType = "XYDiameter" Then
    Call DrawPVIndicatorXYDiameter(PicCntrlName, FeatureIndicatorX1, FeatureIndicatorY1, FeatureIndicatorX2, FeatureIndicatorY2, FeatureIndicatorX3, FeatureIndicatorY3, FeatureIndicatorX4, FeatureIndicatorY4)
End If
'^^^^ *****************************************************


Exit Sub
Err_Handler:
    Select Case Err
        Case 9: Exit Sub
        Case 6: Resume Next
        Case Else: MsgBox Err & error$
    End Select
End Sub



Sub PVDeltaCalcShowMaxMin(FrameNo As Integer) 'PCN1938
On Error GoTo Err_Handler
    Dim SegMax As Integer
    Dim SegMin As Integer
    Dim X As Single
    Dim Y As Single
    
    
    
    
    SegMax = PVDeltaSegFullMax(PVFrameNo)
    SegMin = PVDeltaSegFullMin(PVFrameNo)

    If SegMax <> -1 Then
        X = pvData(SegMax, 1, FrameNo) / PVDataXYMultiplier / (PVDrawScreenRatio) + DrawingCentreX  'PCN3053
        Y = DrawingCentreY - pvData(SegMax, 2, FrameNo) / PVDataXYMultiplier / (PVDrawScreenRatio)  'PCN3053
        ClearLineScreen.PVScreen.Line (X, Y)-(DrawingCentreX, DrawingCentreY), vbBlue
    End If
    
    If SegMin <> -1 Then
        X = pvData(SegMin, 1, FrameNo) / PVDataXYMultiplier / (PVDrawScreenRatio) + DrawingCentreX  'PCN3053
        Y = DrawingCentreY - pvData(SegMin, 2, FrameNo) / PVDataXYMultiplier / (PVDrawScreenRatio)  'PCN3053
        ClearLineScreen.PVScreen.Line (X, Y)-(DrawingCentreX, DrawingCentreY), vbGreen
    End If

Exit Sub
Err_Handler:
    Select Case Err
        Case 9 'Out of subscript range
            Exit Sub
        Case Else
            MsgBox Err & error$
    End Select
End Sub

''Sub WaterLevelSetupAndDraw(screenctrl As Control, X As Single, Y As Single, SaveX As Double, SaveY As Double, WL_LineDistance As Double)
'''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'''WaterLevelSetupAndDraw Function  Geoff Logan geofflogan@cbsys.co.nz
'''
'''Revision history"
'''   V0.0    Geoff Logan,    28/05/03     Building initial framework  'PCN1939
'''
'''Description:
'''    Set the water level for image processing purposes.
'''    The initial water level is manually set to provide the
'''    image processing a method of filtering out the affect of
'''    water level on the profile process.
'''   The first mouse down is the start angle and the second is the finish, clockwise.
'''
'''Purpose:
'''
'''
'''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''On Error GoTo Err_Handler
''
''Dim WLStep As Integer
''Dim WL_i As Integer
''Dim WLDrawCentreX As Double
''Dim WLDrawCentreY As Double
''Dim WLDrawRadius As Double
''Dim WLTheta As Double
''Dim CalculationsCentreX As Double
''Dim CalculationsCentreY As Double
''
'''Draw water level
'''The first mouse down is the start angle and the second is the finish, clockwise.
''WLStep = 10
''WLDrawCentreX = (SaveX - X) / (2 * WLStep)
''WLDrawCentreY = (SaveY - Y) / (2 * WLStep)
''WLDrawRadius = WL_LineDistance / (2 * WLStep)
''If Abs(SaveX - X) > 0.00000001 Then
''    WLTheta = Atn((SaveY - Y) / (SaveX - X))
''Else
''    WLTheta = PI / 2
''End If
''WLDrawStAng = PI - WLTheta
''WLDrawFhAng = 2 * PI - WLTheta
'''Check orientation of waves
''If X > SaveX Then
''    WLDrawStAng = WLDrawStAng + PI
''    WLDrawFhAng = WLDrawFhAng + PI
''End If
'''Check 0 & 2PI limits
''If WLDrawStAng < 0 Then
''    WLDrawStAng = WLDrawStAng + 2 * PI
''ElseIf WLDrawStAng > 2 * PI Then
''    WLDrawStAng = WLDrawStAng - 2 * PI
''End If
''If WLDrawFhAng < 0 Then
''    WLDrawFhAng = WLDrawFhAng + 2 * PI
''ElseIf WLDrawFhAng > 2 * PI Then
''    WLDrawFhAng = WLDrawFhAng - 2 * PI
''End If
'''Now draw water level
''screenctrl.AutoRedraw = True
''For WL_i = 1 To 2 * WLStep Step 2
''    If X > SaveX Then
''        screenctrl.Circle (X + WL_i * WLDrawCentreX, Y + WL_i * WLDrawCentreY), WLDrawRadius, RGB(150, 150, 150), WLDrawStAng, WLDrawFhAng
''    Else
''        screenctrl.Circle (X + WL_i * WLDrawCentreX, Y + WL_i * WLDrawCentreY), WLDrawRadius, RGB(150, 150, 150), WLDrawStAng, WLDrawFhAng
''    End If
''Next WL_i
''screenctrl.AutoRedraw = False
''
'''Set the start and finish WL angles
''VideoCentreLineX = screenctrl.width / 2
''VideoCentreLineY = screenctrl.height / 2
''
''
''
''If screenctrl.name = "SnapShotScreen" And VideoCentreLineX <> 0 Then
''    CalculationsCentreX = VideoCentreLineX ' + ((X4Circle2(1) - CentreLineX))
''    CalculationsCentreY = VideoCentreLineY ' + ((Y4Circle2(1) - CentreLineX))
''Else
''    CalculationsCentreX = (screenctrl.width / 2) '+ ((X4Circle2(1) - CentreLineX))
''    CalculationsCentreY = (screenctrl.height / 2) '+ ((Y4Circle2(1) - CentreLineX))
''
''
''End If
''
''    If Abs(SaveX - CalculationsCentreX) < 0.00001 Then
''        If SaveY < CalculationsCentreY Then
''            WLStartAngle = 0
''        Else
''            WLStartAngle = PI
''        End If
''    ElseIf SaveX > CalculationsCentreX Then
''        WLStartAngle = Atn((SaveY - CalculationsCentreY) / (SaveX - CalculationsCentreX)) + PI / 2
''    Else
''        WLStartAngle = Atn((SaveY - CalculationsCentreY) / (SaveX - CalculationsCentreX)) + 3 * PI / 2
''    End If
''    If Abs(X - CalculationsCentreX) < 0.000001 Then
''        If Y < CalculationsCentreY Then
''            WLFinishAngle = 0
''        Else
''            WLFinishAngle = PI
''        End If
''    ElseIf X > CalculationsCentreX Then
''        WLFinishAngle = Atn((Y - CalculationsCentreY) / (X - CalculationsCentreX)) + PI / 2
''    Else
''        WLFinishAngle = Atn((Y - CalculationsCentreY) / (X - CalculationsCentreX)) + 3 * PI / 2
''    End If
''
''
'''Debug.Print "Start = " & 180 * WLStartAngle / PI & ", Finish = " & 180 * WLFinishAngle / PI
''
''Exit Sub
''Err_Handler:
''    MsgBox Err & error$
''End Sub


Sub RetrievePVXYData(ByVal FrameNo As Long, ByVal PointIndex As Long, ByRef XData As Double, ByRef YData As Double)
'PCN3513
'Name    : RetrievePVXYData
'Created : 16 May 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   : FrameNo, frame that the xy profile point is wanted from.
'          PointIndex, which profile point to retrieve
'          XData, return value of x
'          YData, return value of y
'Desc    : This function returns the XY values in td_pvdata by
'          giving the frame and profile point
'Usage   : Because the td_pvdatax and y are long array, manual array indexing is
'          needed. This does it for you and error checks
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'DataToText Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff,    30/07/03     Building initial framework
'
'Description:
'       Dump data to file
'Purpose:
'


On Error GoTo Err_Handler

Dim XYArrayIndex As Long

If (PVDataNoOfLines < 1) Or _
   (FrameNo > PVDataNoOfLines) Or _
   (PointIndex < 1) Or _
   (PointIndex > NoOfProfileSegments) Then
    XData = 0
    YData = 0
    Exit Sub
End If
    
XYArrayIndex = PointIndex + (NoOfProfileSegments * (FrameNo - 1))

XData = TD_PVDataX(XYArrayIndex)
YData = TD_PVDataY(XYArrayIndex)

If XData <> 0 Or YData <> 0 Then
    YData = YData + (TD_PVCentreY(FrameNo)) 'PCN3219 added centre offset
    XData = XData + (TD_PVCentreX(FrameNo))  'PCN3219 added centre offset
End If

Exit Sub
Err_Handler:
Select Case Err
    Case 6: Resume Next
    Case Else
        MsgBox Err & " - " & error$
End Select
    
 
    
End Sub

Sub DataToText(FrameNo As Long, PVArrayBufferNo As Integer, Value1 As Integer, Value2 As Double, Value3 As Double, Value4 As Double)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : GetAllFormsInfo
'Created : 30 July 2003, PCN1782
'Updated :
'Prg By  : Geoff Logan
'Param   :  FrameNo, is the PVFrameNo
'           PVArrayBufferNo for the current PVData
'           Value1 to Value4 are what the programmer wants to save to file for testing
'Desc    : This function is purely for testing code only. It is for saving the PVData and other data to a text file for further analysis, calculation varification
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'DataToText Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff,    30/07/03     Building initial framework
'
'Description:
'       Dump data to file
'Purpose:
'
On Error GoTo Err_Handler
Dim I As Integer

Open LocToSave & "ProfileData" & FrameNo & ".txt" For Output As #2
    Print #2, "PVFrameNo, PVData, MaxSegNo, q2, MaxDia, MeanDia"
    For I = 0 To 180
        Print #2, I & ", " & pvData(I, 0, PVArrayBufferNo) & ", " & Value1 & ", " & Value2 & ", " & Value3 & ", " & Value4
    Next I
Close #2


Exit Sub
Err_Handler:
    MsgBox Err & error$
End Sub

Sub GetPVDPointerPVDataFromFile(FileName As String, PVGraphDataAddressOffset As Long, FileLoadError As Boolean)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : GetPVDPointerPVDataFromFile
'Created : 7 August 2003, PCN2164
'Updated :
'Prg By  : Geoff Logan
'Param   :  FileName - PVD file name
'           FileLoadError - returns a true value if an error occurs while loading.
'Desc    : Gets the PVDPointerPVData from the PVD file .
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

FileLoadError = False 'PCNGL140103

If Dir(FileName) = "" Or FileName = "" Then
    FileLoadError = True 'PCNGL140103
    Exit Sub
End If

'Check whether a file is open
Open FileName For Binary Access Read Lock Write As #7 'PCN2980

'Load the File Main Header
Get #7, , PVDFileMainHeader 'PCN2980

'Determine file header pointers and CheckSums then read the header pointers to block 2
'File header pointers (must be the second block of data in any PVD file, there are no fix number of pointers)
If PVDFileMainHeader.PVDFileMHPointerAddress = 0 Then 'PCNGL140103
    Close #7 'PCN2980
    FileLoadError = True 'PCNGL140103
    Exit Sub
End If
Get #7, PVDFileMainHeader.PVDFileMHPointerAddress, PVDFilePointers 'PCN2980

'Read from file the capacity data
PVDHeaderPVData.PVDHeaderDescriptor = ""
PVDHeaderPVData.PVDCheck = 0
Get #7, PVDFilePointers.PVDPointerPVData, PVDHeaderPVData 'PCN2980

If Left(PVDHeaderPVData.PVDHeaderDescriptor, 8) <> "[PVData]" Or PVDHeaderPVData.PVDCheck = 0 Then 'PCNGL130103
    Close #7 'PCN2980
    FileLoadError = True
    Exit Sub
End If

PVGraphDataAddressOffset = Seek(7) ' The start address of the PVData file data block


Close #7 'PCN2980

Exit Sub
FileErrorCleanup: 'PCNGL140103
    Close #7 'PCN2980
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
            MsgBox Err & error$
    End Select
End Sub


Sub RapidReadPVDataFromFile(FileNumber As Integer, PVDataAddressOffset As Long, FrameBufferNo As Integer, FileLoadError As Boolean)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : RapidReadPVDataFromFile
'Created : 7 August 2003, PCN2164
'Updated :
'Prg By  : Geoff Logan
'Param   :  FileNumber - The number of the file in which the PVData is to be retrieved
'           PVDataAddressOffset
'           FrameBufferNo
'           FileLoadError - returns a true value if an error occurs while loading.
'Desc    : Gets, as fast as possible, the PVGraphs Data from the PVD file .
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim PVSegmentNo As Integer
Dim LenOfPVData As Integer
Dim XData As Long
Dim YData As Long



LenOfPVData = Len(pvData(1, 0, 1))

For PVSegmentNo = 1 To NoOfProfileSegments
    'vvvv PCN2891 *******************************************
'    If ConfigInfo.ProfileRecordingMethod = "XY" Then
'        'Get the PVData X
'        Get #FileNumber, PVDataAddressOffset, PVData(PVSegmentNo, 1, FrameBufferNo)
'        PVDataAddressOffset = PVDataAddressOffset + LenOfPVData
'        'Get the PVData Y
'        Get #FileNumber, PVDataAddressOffset, PVData(PVSegmentNo, 2, FrameBufferNo)
'        PVDataAddressOffset = PVDataAddressOffset + LenOfPVData
'        'Calculate Radius - If required.
'        PVData(PVSegmentNo, 0, FrameBufferNo) = Int(PVDataRadiusCalc(PVSegmentNo, FrameBufferNo))
'    Else
        'Get the PVData Radius
        Get #FileNumber, PVDataAddressOffset, pvData(PVSegmentNo, 0, FrameBufferNo)
        
        'PCN3513 (Antony, 16 may 2005) reconstonstruct the XY Data from old raidus PVD
        '                                                                       '
        PVDataAddressOffset = PVDataAddressOffset + LenOfPVData                 '
'        Call ConvertTDPVRadiusToXY((PVSegmentNo * 2) + 180, _
'                              PVData(PVSegmentNo, 0, FrameBufferNo), _
'                              XData, _
'                              YData)
        Call ConvertTDPVRadiusToXY((PVSegmentNo * 2), _
                              pvData(PVSegmentNo, 0, FrameBufferNo), _
                              XData, _
                              YData)
                              
                              
        pvData(PVSegmentNo, 1, FrameBufferNo) = XData / ConfigInfo.Ratio / VideoScreenScale
        pvData(PVSegmentNo, 2, FrameBufferNo) = YData / -ConfigInfo.Ratio / VideoScreenScale
        '                                                                       '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
'    End If
    '^^^^ ****************************************************
Next PVSegmentNo

Exit Sub
Err_Handler:
    Select Case Err
        Case 9 'Out of subscript range
            FileLoadError = True
            Exit Sub
        Case 52 'Bad filename or number
            FileLoadError = True
            Exit Sub
        Case Else
            FileLoadError = True
            MsgBox Err & error$
    End Select
End Sub

Sub RapidReadPVData(ByVal FrameNo As Long)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCN3526
'Name    : RapidReadPVData
'Created : 18 May 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   :  FrameBuffer, Frame to Load
'Desc    : Instead of loading data everytime form from, load it from a prefeched array
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Dim PVSegmentNo As Integer
Dim XData As Double
Dim YData As Double


For PVSegmentNo = 1 To NoOfProfileSegments
    Call RetrievePVXYData(FrameNo, PVSegmentNo, XData, YData)
    pvData(PVSegmentNo, 1, 1) = XData
    pvData(PVSegmentNo, 2, 1) = YData
    pvData(PVSegmentNo, 0, 1) = Int(PVDataTrueRadiusCalc(PVSegmentNo, 1))
Next PVSegmentNo

Exit Sub
Err_Handler:
    Select Case Err
    Case 9: Exit Sub 'Out of subscript range
    Case 6: Resume Next
    Case Else
        MsgBox Err & " - " & error$
End Select
 
End Sub

Sub ExportSingleGraph(ExportFileName As String)
On Error GoTo Err_Handler

Dim FileName As String
Dim FileNo
Dim ExportSingleGraphVersion As String
Dim Header As String
Dim LinePrint As String
Dim GraphString As String
Dim FrameNo As Long


ExportSingleGraphVersion = "SG1.0"
FileName = ExportFileName & ".CSV"
GraphString = "Ovality"
Header = "Frame," & GraphString & "," & ExportSingleGraphVersion

FileNo = FreeFile

Open ExportFileName For Output As FileNo

Print #FileNo, Header

For FrameNo = 1 To PVDataNoOfLines
    LinePrint = FrameNo & "," & PVOvalityFullData(FrameNo)
    Print #FileNo, LinePrint
Next FrameNo

Close #FileNo

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & error$: Close #FileNo
    End Select
End Sub


Sub ExportPVData(ExportFileName As String, ExportFormat As String, StartFrameNo As Long, FinishFrameNo As Long, NoOfSegPerProfile As Integer)   'Moved from PrecisionVisionGraph form 'PCNGL030103
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : ExportPVD
'Created : 11 December 2003, PCN2481
'Updated :
'Prg By  : Geoff Logan
'Param   : StartFrameNo -
'           FinishFrameNo -
'           NoOfSegPerProfile - The number of segments per profile to be exported
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim CurrentFrameNo As Long
Dim PVFileLoadError As Boolean 'PCNGL140103
Dim PVDataStartAddress As Long
Dim PVDataAddressOffset As Long
Dim LenOfPVData As Integer
Dim InputFileNumber As Integer
Dim OutputFileNumber As Integer
Dim ExportingProgress As Integer
Dim ProgressIncrement As Integer
Dim ExportingProgress_Current As Integer
Dim DistanceConverted As Double 'PCN2481

''Determine the relevant PVFrameNo for the current Start Marker
'StartFrameNo = Int((PrecisionVisionGraph.YScaleMarkerSt.Y1 - PrecisionVisionGraph.PVYScaleZeroMarker.Y1) / PVGraphYRatio) 'PCNGL030103 'PCN1828
''Determine the relevant PVFrameNo for the current Finish Marker
'FinishFrameNo = Int((PrecisionVisionGraph.YScaleMarkerFh.Y1 - PrecisionVisionGraph.PVYScaleZeroMarker.Y1) / PVGraphYRatio) 'PCNGL030103 'PCN1828
'PVFrameNo = FinishFrameNo

If StartFrameNo < 1 Then StartFrameNo = 1
If PVFrameNo < 1 Then
    PVFrameNo = 1
    FinishFrameNo = 1
End If

If NoOfSegPerProfile < 12 Then
    NoOfSegPerProfile = 6
ElseIf NoOfSegPerProfile < 18 Then
    NoOfSegPerProfile = 12
ElseIf NoOfSegPerProfile < 36 Then
    NoOfSegPerProfile = 18
ElseIf NoOfSegPerProfile < NoOfProfileSegments Then
    NoOfSegPerProfile = 36
Else
    NoOfSegPerProfile = NoOfProfileSegments
End If

ProgressIncrement = 5
Call CLPProgressBar.ProgressBarInitialise(DisplayMessage("---->>>>"))
DoEvents


InputFileNumber = 7
OutputFileNumber = 8
LenOfPVData = Len(pvData(1, 0, 1))
Call GetPVDPointerPVDataFromFile(PVDFileName, PVDataStartAddress, PVFileLoadError) 'PCN2164
If PVFileLoadError Then Exit Sub
Open PVDFileName For Binary Access Read Lock Write As InputFileNumber

Select Case ExportFormat
    Case "CSV"
        Open ExportFileName For Output As OutputFileNumber
    Case "TXT" 'PCN2481 171203
        Open ExportFileName For Output As OutputFileNumber

End Select



For CurrentFrameNo = StartFrameNo To FinishFrameNo
    
'    PVDataAddressOffset = PVDataStartAddress + (CurrentFrameNo) * LenOfPVData * NoOfProfileSegments  ' The calculation for the size of PVData file data block
'    PVDataAddressOffset = PVDataAddressOffset + (CurrentFrameNo) * (Len(PVCapacityData(1)) + Len(PVOvalityData(1)) + 2 * Len(PVDelta(1)) + Len(AVIFrameTime(1))) ' The calculation for the size of PVData file data block
    
'PCN3526 ''''''''''''''''''''''''''''''''''''''''''''''''
'                                                       '
'    PVDataAddressOffset = PVDataAddressOffsetCalc(PVDataStartAddress, CurrentFrameNo) 'PCN2639
'    Call RapidReadPVDataFromFile(InputFileNumber, PVDataAddressOffset, 1, PVFileLoadError) 'PCN2164
'    PVDataAddressOffset = Seek(InputFileNumber)
'    Call RapidReadPVGraphsDataFromFile(InputFileNumber, CurrentFrameNo, PVDataAddressOffset, 1, PVFileLoadError)
'                                                       '
    Call RapidReadPVData(CurrentFrameNo)                '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Calculate Distance
'    DistanceConverted = Round(CameraSpeedInTime * (PVTimes(CurrentFrameNo) - DistanceStartTime) + DistanceStart, 2) 'PCN3180
    DistanceConverted = PVDistances(CurrentFrameNo) 'PCN3180
    ' Write Frame data to file
    Select Case ExportFormat
        Case "CSV"
            If CurrentFrameNo = StartFrameNo Then
                Call ExportInDelimitedFormat_Header(OutputFileNumber, ", ", "Frame No", NoOfSegPerProfile, "VideoTime", "Distance") 'PCN2481 171203
            End If
            Call ExportInDelimitedFormat(OutputFileNumber, ", ", CStr(CurrentFrameNo), NoOfSegPerProfile, PVTimes(CurrentFrameNo), DistanceConverted) 'PCN2481 171203
        Case "TXT" 'PCN2481 171203
            If CurrentFrameNo = StartFrameNo Then
                Call ExportInDelimitedFormat_Header(OutputFileNumber, Chr(9), "Frame No", NoOfSegPerProfile, "VideoTime", "Distance") 'PCN2481 171203
            End If
            Call ExportInDelimitedFormat(OutputFileNumber, Chr(9), CStr(CurrentFrameNo), NoOfSegPerProfile, PVTimes(CurrentFrameNo), DistanceConverted) 'PCN2481 171203
    
    End Select
    ExportingProgress = 100 * (CurrentFrameNo - StartFrameNo) / (FinishFrameNo - StartFrameNo)
    If ExportingProgress Mod ProgressIncrement = 0 And ExportingProgress > (ExportingProgress_Current + 1) Then
        Call CLPProgressBar.ProgressBarPosition(ExportingProgress / 100)
        DoEvents
        ExportingProgress_Current = ExportingProgress
    End If
Next CurrentFrameNo

Close #OutputFileNumber



Close #InputFileNumber 'PCN2164

Call CLPProgressBar.ProgressBarPosition(1#)
DoEvents
    
Exit Sub
Err_Handler:
    MsgBox error$
End Sub

Sub ExportInDelimitedFormat(OutputFileNumber As Integer, DelimitType As Variant, LineRef As String, NoOfSegPerProfile As Integer, Optional PVTime As Double, Optional DistConverted As Double)    'Moved from PrecisionVisionGraph form 'PCNGL030103 'PCNGL151203
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : ExportInDelimitedFormat, (was ExportInCSVFormat)
'Created : 11 December 2003, PCN2481
'Updated : 17 December 2003, PCN2481
'Prg By  : Geoff Logan
'Param   : OutputFileNumber - Number of file output.
'           LineRef - Line reference, e.g. PVFrameNo or Distance
'           NoOfSegPerProfile - Determines how many data points are outputed to
'                               file per profile.
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim PVSegmentNo As Integer
Dim ExportData As String
Dim SegStep As Integer


If LineRef <> "" Then
    ExportData = LineRef
Else
    ExportData = "?"
End If


ExportData = ExportData & DelimitType & Format(DistConverted, "#0.00") 'PCN2481 171203
    
ExportData = ExportData & DelimitType & CStr(PVTime)  'PCN2481 171203

If NoOfSegPerProfile < 1 Or NoOfSegPerProfile > NoOfProfileSegments Then Exit Sub

If ConfigInfo.ProfileRecordingMethod = "XY" Then ExportData = ExportData & DelimitType & "X-->"
SegStep = NoOfProfileSegments / NoOfSegPerProfile
For PVSegmentNo = NoOfProfileSegments To 1 Step -SegStep
'    ExportData = ExportData & ", " & PVData(PVSegmentNo, 0, 1)
    'vvvv PCN3155 **************************************************************
    If ConfigInfo.ProfileRecordingMethod = "XY" Then
'        ExportData = ExportData & DelimitType & Format(PVDataRadiusCalc(PVSegmentNo, 1), "#0.0")
   '     ExportData = ExportData & DelimitType & Format(PVData(PVSegmentNo, 1, 1) / PVDataXYMultiplier * ConfigInfo.Ratio * VideoScreenScale, "#0.00")
    ExportData = ExportData & DelimitType & Format(pvData(PVSegmentNo, 1, 1) / PVDataXYMultiplier, "#0.00")
    Else
        ExportData = ExportData & DelimitType & pvData(PVSegmentNo, 0, 1)
    End If
    '^^^^ **********************************************************************
Next PVSegmentNo

'vvvv PCN2741 ********************************************
'Include Capacity, Ovality, DeltaMin, DeltaMax
ExportData = ExportData & DelimitType & (Format(PVCapacityFullData(LineRef), "#0.00")) 'PCNANT
ExportData = ExportData & DelimitType & (Format(PVOvalityFullData(LineRef), "#0.00")) 'PCNANT
ExportData = ExportData & DelimitType & (Format(PVDeltaFullMax(LineRef), "#0.00")) 'PCNANT
ExportData = ExportData & DelimitType & (Format(PVDeltaFullMin(LineRef), "#0.00")) 'PCNANT
'^^^^ ****************************************************

Print #OutputFileNumber, ExportData
If ConfigInfo.ProfileRecordingMethod = "XY" Then
    ExportData = DelimitType & DelimitType & DelimitType & "Y-->"
    For PVSegmentNo = NoOfProfileSegments To 1 Step -SegStep
            ExportData = ExportData & DelimitType & Format(pvData(PVSegmentNo, 2, 1) / PVDataXYMultiplier, "#0.00")
    Next PVSegmentNo
    Print #OutputFileNumber, ExportData
End If
    
Exit Sub
Err_Handler:
    MsgBox error$

End Sub

Sub ExportInDelimitedFormat_Header(OutputFileNumber As Integer, DelimitType As Variant, LineRef As String, NoOfSegPerProfile As Integer, Optional PVTimeHeader As String, Optional DistHeader As String) 'Moved from PrecisionVisionGraph form 'PCNGL030103 'PCNGL151203
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : ExportInDelimitedFormat_Header (was ExportInCSVFormat)
'Created : 11 December 2003, PCN2481
'Updated : 17 December 2003, PCN2481
'Prg By  : Geoff Logan
'Param   : OutputFileNumber - Number of file output.
'           LineRef - Line reference, e.g. PVFrameNo or Distance
'           NoOfSegPerProfile - Determines how many data points are outputed to
'                               file per profile.
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim PVSegmentNo As Integer
Dim ExportData As String
Dim SegStep As Integer
Dim SegAngle As Double
Dim XData As Double
Dim YData As Double
Dim I As Double

ExportData = ",,,WaterLevel"
For I = 0 To 179
    ExportData = ExportData & "," & WaterEgnoreList(I)
Next I

Print #OutputFileNumber, ExportData

If LineRef <> "" Then
    ExportData = LineRef
Else
    ExportData = "?"
End If

ExportData = ExportData & DelimitType & DistHeader 'PCN2481 171203

ExportData = ExportData & DelimitType & PVTimeHeader 'PCN2481 171203

If NoOfSegPerProfile < 1 Or NoOfSegPerProfile > NoOfProfileSegments Then Exit Sub
SegStep = NoOfProfileSegments / NoOfSegPerProfile

If ConfigInfo.ProfileRecordingMethod = "XY" Then
    For PVSegmentNo = 1 To NoOfProfileSegments
        SegAngle = PVSegmentNo
        
        If PVSegmentNo = 1 Then
            ExportData = ExportData & DelimitType & ",Point " & Format(SegAngle, "#0") & " (1 = 6o'clk -- 45 = 9o'clk) " 'PCN2481 171203
        Else
            ExportData = ExportData & DelimitType & "Point " & Format(SegAngle, "#0")   'PCN2481 171203
        End If
    Next PVSegmentNo
Else
    For PVSegmentNo = NoOfProfileSegments To 1 Step -SegStep
        SegAngle = (PVSegmentNo / NoOfProfileSegments) * 360
        
        If PVSegmentNo = NoOfProfileSegments Then
            ExportData = ExportData & DelimitType & Format(SegAngle, "#0") & "deg (360deg = 6o'clk -- 90deg = 9o'clk) " 'PCN2481 171203
        Else
            ExportData = ExportData & DelimitType & Format(SegAngle, "#0") & "deg" 'PCN2481 171203
        End If
    Next PVSegmentNo
End If

ExportData = ExportData & DelimitType & "Capacity , Ovality, DeltaMin, DeltaMax"  'PCN2741

Print #OutputFileNumber, ExportData

    
Exit Sub
Err_Handler:
    MsgBox error$
End Sub

Sub DrawPVIndicatorXYDiameter(PicCntrlName As Control, FeatureIndicatorX1 As Integer, FeatureIndicatorY1 As Integer, FeatureIndicatorX2 As Integer, FeatureIndicatorY2 As Integer, FeatureIndicatorX3 As Integer, FeatureIndicatorY3 As Integer, FeatureIndicatorX4 As Integer, FeatureIndicatorY4 As Integer)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : DrawPVIndicatorXYDiameter
'Created : 6 May 2004, PCN2703
'Updated :
'Prg By  : Geoff Logan
'Param   : PVGraphType -
'          FrameNo -
'          StartFrameNo -
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

If ImageGraphState(0).GraphType = "XYDiameter" Then
    If FeatureIndicatorX1 <> 0 And FeatureIndicatorY1 <> 0 And FeatureIndicatorX2 <> 0 And FeatureIndicatorY2 <> 0 Then
        PicCntrlName.Line (FeatureIndicatorX1, FeatureIndicatorY1)-(FeatureIndicatorX2, FeatureIndicatorY2), RGB(128, 0, 200) 'Purple
        Call DrawPVIndicatorArrows(PicCntrlName, FeatureIndicatorX1, FeatureIndicatorY1, FeatureIndicatorX2, FeatureIndicatorY2, RGB(128, 0, 200)) 'PCN2962
    End If
    If FeatureIndicatorX3 <> 0 And FeatureIndicatorY3 <> 0 And FeatureIndicatorX4 <> 0 And FeatureIndicatorY4 <> 0 Then
        PicCntrlName.Line (FeatureIndicatorX3, FeatureIndicatorY3)-(FeatureIndicatorX4, FeatureIndicatorY4), RGB(0, 128, 0) 'Green graph line
        Call DrawPVIndicatorArrows(PicCntrlName, FeatureIndicatorX3, FeatureIndicatorY3, FeatureIndicatorX4, FeatureIndicatorY4, RGB(0, 128, 0)) 'PCN2962
    End If
End If


Exit Sub
Err_Handler:
    Select Case Err
        Case 9 'Out of subscript range
            Exit Sub
        Case Else
            MsgBox Err & error$
    End Select
End Sub







Function PVDataAngleBetweenSegments(PVSegment As Integer, PVArrayBufferNo As Integer) As Double
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : PVDataAngleBetweenSegments
'Created : 22 June 2004, PCN2891
'Updated :
'Prg By  : Geoff Logan
'Param   : PVSegment - the current PV Segment
'          PVArrayBufferNo - PVData Array BufferNo
'Desc    : Determine the angle between segments.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim CurrentPVDataValue As Double
Dim CurrentAngle As Double
Dim NextPVDataValue As Double
Dim NextAngle As Double

'CurrentPVDataValue = PVDataRadiusCalc(PVSegment, PVArrayBufferNo)
CurrentPVDataValue = Sqr(pvData(PVSegment, 1, PVArrayBufferNo) ^ 2 + pvData(PVSegment, 2, PVArrayBufferNo) ^ 2)
If CurrentPVDataValue = 0 Then
    PVDataAngleBetweenSegments = 2 / 180 'PCN3097 at 6oclkock gives error supose to be radians
    Exit Function
Else
    'Calc current angle
'    CurrentAngle = ArcSin(PVData(PVSegment, 2, PVArrayBufferNo) / CurrentPVDataValue) * 180 / PI
    CurrentAngle = ArcSin(pvData(PVSegment, 2, PVArrayBufferNo) / CurrentPVDataValue)
    'Calc next angle
    If PVSegment < NoOfProfileSegments Then
        NextPVDataValue = Sqr(pvData(PVSegment + 1, 1, PVArrayBufferNo) ^ 2 + pvData(PVSegment + 1, 2, PVArrayBufferNo) ^ 2)
    Else
        NextPVDataValue = Sqr(pvData(1, 1, PVArrayBufferNo) ^ 2 + pvData(1, 2, PVArrayBufferNo) ^ 2)
    End If
    If NextPVDataValue = 0 Then
        PVDataAngleBetweenSegments = 2 / 180  'PCN3097 at 6oclkock gives error supose to be radians
        Exit Function
    Else
'        NextAngle = ArcSin(PVData(PVSegment + 1, 2, PVArrayBufferNo) / NextPVDataValue) * 180 / PI
        If PVSegment < NoOfProfileSegments Then
            NextAngle = ArcSin(pvData(PVSegment + 1, 2, PVArrayBufferNo) / NextPVDataValue)
        Else
            NextAngle = ArcSin(pvData(1, 2, PVArrayBufferNo) / NextPVDataValue)
        End If
        PVDataAngleBetweenSegments = Abs(NextAngle - CurrentAngle)
    End If
End If

Exit Function
Err_Handler:
    MsgBox Err & "-" & error$
End Function


Function ArcSin(X As Double) As Double
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : ArcSin
'Created : 22 June 2004, PCN2891
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Inv SIN.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

If Abs(X) = 1 Then
    ArcSin = X * PI / 2
Else
    ArcSin = Atn(X / Sqr(-X * X + 1))
End If

Exit Function
Err_Handler:
    If Err = 5 Then ArcSin = 0: Exit Function
    MsgBox Err & "-" & error$

End Function


Sub DrawPVIndicatorArrows(PicCntrlName As Control, PointX1 As Integer, PointY1 As Integer, PointX2 As Integer, PointY2 As Integer, ArrowColour As Long)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : DrawPVIndicatorArrows
'Created : 6 August 2004, PCN2962
'Updated :
'Prg By  : Geoff Logan
'Param   : PVGraphType -
'          FrameNo -
'          StartFrameNo -
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim Theta As Double
Dim NewX As Double
Dim NewY As Double
Dim ArrowAngle As Double
Dim ThetaOffset As Double

ArrowAngle = 0.19 'radians (11 degs)

If PointY2 = DrawingCentreY Or PointX2 = DrawingCentreX Then
    'Don't draw arrows if at the centre point
    Exit Sub
End If

'Determine angle of arrow for X1,X2, Y1 and Y2
If Abs(PointX1 - PointX2) > 0.00000001 Then
    Theta = Atn((PointY1 - PointY2) / (PointX1 - PointX2))
    If (PointX2 - PointX1) < 0 Then
        Theta = Theta + PI
    End If
Else
    Theta = PI / 2
End If


'draw arrows
NewX = PointX1 + (15 * Cos(Theta + ArrowAngle))
NewY = PointY1 + (15 * Sin(Theta + ArrowAngle))
PicCntrlName.Line (PointX1, PointY1)-(NewX, NewY), ArrowColour
NewX = PointX1 + (15 * Cos(Theta - ArrowAngle))
NewY = PointY1 + (15 * Sin(Theta - ArrowAngle))
PicCntrlName.Line (PointX1, PointY1)-(NewX, NewY), ArrowColour
NewX = PointX2 - (15 * Cos(Theta + ArrowAngle))
NewY = PointY2 - (15 * Sin(Theta + ArrowAngle))
PicCntrlName.Line (NewX, NewY)-(PointX2, PointY2), ArrowColour
NewX = PointX2 - (15 * Cos(Theta - ArrowAngle))
NewY = PointY2 - (15 * Sin(Theta - ArrowAngle))
PicCntrlName.Line (NewX, NewY)-(PointX2, PointY2), ArrowColour


Exit Sub
Err_Handler:
    Select Case Err
        Case 9 'Out of subscript range
            Exit Sub
        Case Else
            MsgBox Err & error$
    End Select
End Sub

Sub ConvertTDPVRadiusToXY(PVSegmentNo As Integer, PVRadius As Single, ByRef TDX As Long, ByRef TDY As Long)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : ConvertTDPVRadiusToXY
'Created : 18 August 2004, PCN2988
'Updated :
'Prg By  : Antony and Geoff
'Param   : PVSegmentNo - PV segment number
'          PVRadius - PVData in radius format (older versions)
'          TDX - 3D PVData X
'          TDY - 3D PVData Y
'Desc    : Converts the PVData (radius format) to PVData XY format for the 3D rendering
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim SegmentTheta As Double
Dim SegmentPos As Integer

SegmentPos = (180 - PVSegmentNo) + 270 'Adjust offset and direction
SegmentTheta = (SegmentPos) * PI / NoOfProfileSegments
'If SegmentTheta > 2 * PI Then SegmentTheta = SegmentTheta - 2 * PI

TDX = Cos(-SegmentTheta) * (PVRadius * PVDataXYMultiplier) 'Note: The PVDataXYMultiplier is not used here since the TD requires the accuracy.
TDY = Sin(-SegmentTheta) * (PVRadius * PVDataXYMultiplier) 'Note: The PVDataXYMultiplier is not used here since the TD requires the accuracy.

Exit Sub
Err_Handler:
    MsgBox Err & "-" & error$
End Sub






Function CalAsin(ByVal sx As Double, ByVal sy As Double, ByVal X As Double, ByVal Y As Double) As Double 'PCN3***
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : CalAsin
'Created : 4 October 2004, PCN3***
'Updated :
'Prg By  : Geoff Logan
'Param   : sx - Centre
'          sy - Centre
'          X - Point to find angle
'          Y - Point to find angle
'Desc    : Moved Abe's CalASin to this module.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim a As Double
Dim b As Double
Dim ab As Double
Dim asin As Double
          
          a = Sqr((sx - X) ^ 2 + (sy - Y) ^ 2)
          b = sx - X
          If a = 0 Then
            ab = 0
          Else
            ab = b / a
          End If
          If ab = 1 Then
            asin = PI / 2
          ElseIf ab = -1 Then
            asin = 3 * PI / 2
          Else
            If a = 0 Then
                asin = PI / 2
            Else
                asin = Atn(ab / Sqr(-ab * ab + 1))
            End If
          End If

          If Y < sy And X >= sx Then ' 1
            asin = PI / 2 + asin
          ElseIf Y >= sy And X >= sx Then ' 2
            asin = 3 * PI / 2 - asin
          ElseIf Y >= sy And X < sx Then ' 3
            asin = 3 * PI / 2 - asin
          ElseIf Y < sy And X < sx Then ' 4
            asin = PI / 2 + asin
          End If

          CalAsin = asin

Exit Function

Err_Handler:

Select Case Err
    Case 11 ' divide by zero
'       ab = ab + 0.001
        'MsgBox "Divide by zero" 'PCN2111
        MsgBox Err & " - " & error$
        Resume Next
    Case 5 '
        'MsgBox "Temporary message after mouseup. Please check your drawing is as intended. " & error$ 'PCN2111
        Resume Next
    Case Else
        MsgBox Err & " - " & error$
End Select

End Function


Function ShapeTypeNormalisedPercent(ExpRad As Integer, ProfileNo As Integer, FrameBufferNo As Integer) As Double
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : ShapeTypeNormalisedPercent
'Created : 6 December 2004 PCN3***
'Updated :
'Prg By  : Antony van Iersel
'Param   : ExpRad - Expected radius of shape
'          ProfileNo - Selected Profile point to retrieve Normalised Percent
'          FrameNumberBuffer - have no Idea :(
'Desc    : Calulates the normalised Percent of the different type of shapes
'Usage   : Used when calculating the PVFlat3DCalc
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    
    Dim X As Single
    Dim Y As Single
    Dim OrthoX As Double
    Dim OrthoY As Double
    Dim Normalised As Double
    
    NormaliseType = "Graph"
'    DrawShapeType = "SemiElliptical"
    
    X = pvData(ProfileNo, 1, FrameBufferNo) / PVDataXYMultiplier * ConfigInfo.Ratio * VideoScreenScale
    'On the reference shape, down is more positive, not negative.
    Y = -pvData(ProfileNo, 2, FrameBufferNo) / PVDataXYMultiplier * ConfigInfo.Ratio * VideoScreenScale
    If X = 0 And Y = 0 Then: ShapeTypeNormalisedPercent = -999999: Exit Function
    
    
'    ShapeTypeNormalisedPercent = 100 * (ExpRad - PVData(ProfileNo, 0, FrameBufferNo)) / ExpRad 'PCNGL140103 'PCN2680
'    Call ProfileRefShapeDistCalc(X, Y, OrthoX, OrthoY, Normalised)
'   ShapeTypeNormalisedPercent = (100 * (ExpRad - (ExpRad - Normalised))) / ExpRad
'    ShapeTypeNormalisedPercent = Abs(ExpRad - Normalised)
    ShapeTypeNormalisedPercent = Normalised
Exit Function
Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & " - " & error$
End Select
End Function

Function ProfileRefShapeDistCalcArc(X As Single, Y As Single, Radius As Double, ArcStart As Double, ArcEnd As Double, OrthoX As Double, OrthoY As Double, OrthoDistance As Double) As Boolean
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : ProfileRefShapeDistCalcArc
'Created : 13 December 2004 2004, PCN3055
'Updated :
'Prg By  : Antony van Iersel
'Param   : X - Current cursor X position
'          Y - Current cursor Y position
'          Radius - radius of arc
'          ArcStart - starting angle of arc (Anti Clockwise, 0 radians East)
'          ArcEnd - ending angle of arc (Anti Clockwise, 0 radians East)
'          OrthoX - The orthoganal X from the reference shape to (X,Y)
'          OrthoY - The orthoganal Y from the reference shape to (X,Y)
'          OrthoDistance - The orthoganal distance from the reference shape to (X,Y)
'Desc    : Finds the distance from the current point to the arc tangent,
'          Returns True or False if inside the arc,
'          Sets Distance from Current X,Y to Ortho X,Y and -ve or +ve if inside or outside expected radius
'          Sets OrthoX, Ortho Y, the point sitting on the arc,
'Usage   : Used to find the normalised value of a arbatury point from an arbatury arc
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Dim NormalisedAngle As Double 'Angle to find ortho x,y also to check to see if inside arc
Dim DistanceToOrtho As Double 'Distance to ortho x,y
Dim DistanceToXY As Double 'Distance to the passed x,y coordinates
Dim DistanceToEnd As Double
Dim DistanceToStart As Double
Dim EndX, EndY, StartX, StartY As Double

NormalisedAngle = CalAsin(0, 0, X, Y)
If ArcEnd = 0 Then
    ArcEnd = ArcEnd + (2 * PI)
End If
If ArcEnd < ArcStart Then: ArcEnd = ArcEnd + (2 * PI): NormalisedAngle = NormalisedAngle + (2 * PI)
'ArcStart = ArcStart + (2 * PI): ArcEnd = ArcEnd + (2 * PI) ' Move the angles by 360deg so arc doesn't pass over 0deg

'360deg added to angle to make sure the arc doesn't pass through 0deg then
'check if between arc start and end. If so then return false and
If (NormalisedAngle < ArcStart) Or (NormalisedAngle > ArcEnd) Then
    StartX = Cos(ArcStart) * Radius
    StartY = Sin(ArcStart) * Radius * -1
    EndX = Cos(ArcEnd) * Radius
    EndY = Sin(ArcEnd) * Radius * -1
    DistanceToStart = Sqr((StartX - X) ^ 2 + (StartY - Y) ^ 2)
    DistanceToEnd = Sqr((EndX - X) ^ 2 + (EndY - Y) ^ 2)
    If DistanceToStart < DistanceToEnd Then
        OrthoX = StartX
        OrthoY = StartY
        OrthoDistance = DistanceToStart
    Else
        OrthoX = EndX
        OrthoY = EndY
        OrthoDistance = DistanceToEnd
    End If
Else
    OrthoX = Cos(NormalisedAngle) * Radius
    OrthoY = Sin(NormalisedAngle) * Radius * -1
    DistanceToOrtho = Sqr(OrthoX ^ 2 + OrthoY ^ 2)
    DistanceToXY = Sqr(X ^ 2 + Y ^ 2)
    OrthoDistance = DistanceToXY - Radius
End If                                      '
'''''''''''''''''''''''''''''''''''''''''''''
    

ProfileRefShapeDistCalcArc = True

Exit Function
Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & " - " & error$
End Select

End Function



Sub RotateCoordXY(X As Double, Y As Double, CenX As Double, CenY As Double, Angle As Double)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCN3055
'Name    : RotateCoordXY
'Created : 15 December 2004
'Updated :
'Prg By  : Antony van Iersel
'Param   : X - X coordinate to rotate
'          Y - Y coordinate to rotate
'          CenX - Centre of the rotation
'          CenY - Centre of the rotation
'          Angle - Radians to rotate the coordinate
'Desc    : Rotates the X and Y value around a given centre by the angle given
'          The angle is in radians that rotates in a Anti-Clockwise direction
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Dim ShiftedX As Double
Dim ShiftedY As Double

' Shift the coordinates relative to the centre to be shifted
ShiftedX = X - CenX
ShiftedY = Y - CenY

'  newpoint.x = xold * cos(angle) - yold * sin(angle)
'  newpoint.Y = xold * Sin(Angle) + yold * Cos(Angle)

' X '         = cos(theta)*x - sin(theta)*y
' Y '         = sin(theta)*x + cos(theta)*y
X = (ShiftedX * Cos(Angle)) - (ShiftedY * Sin(Angle))
Y = (ShiftedX * Sin(Angle)) + (ShiftedY * Cos(Angle))

X = X + CenX
Y = Y + CenY

Exit Sub
Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & " - " & error$
End Select

End Sub

Function PVDataTrueRadiusCalc(j As Integer, PVArrayBufferNo As Integer, Optional X As Variant, Optional Y As Variant) As Double
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCNant????
'Name    : PVDataTrueRadiusCalc
'Created : 5 January 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   : j - the current PV Segment
'          PVArrayBufferNo - PVData Array BufferNo
'          X - (Optional) PVData for the X co-ordinate. Used when PVData is not populated with valid data, eg on loading PVD.
'          Y - (Optional) PVData for the Y co-ordinate. Used when PVData is not populated with valid data, eg on loading PVD.
'Desc    : Calculates the PVData radius from PVDataX and PVDataY.
'          Same as PVDataRadiusCalc except there is no exception of shape. The original did a reference shape
'          calculation if the shape was not circle, this one calculates the radius regardless of shape.
'Usage   : Note: Make sure VideoScreenScale is calculated before
'          running this function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim PVData_X As Double
Dim PVData_Y As Double
Dim PVData_Rx As Double
Dim PVData_Ry As Double
'vvvv PCN3*** **********************
Dim OrthoX As Double
Dim OrthoY As Double
'^^^^ ******************************

If j > NoOfProfileSegments Then
    PVDataTrueRadiusCalc = 0
    Exit Function
End If

If IsMissing(X) Or IsMissing(Y) Then
    PVData_X = pvData(j, 1, PVArrayBufferNo)
    PVData_Y = pvData(j, 2, PVArrayBufferNo)
Else
    PVData_X = X
    PVData_Y = Y
End If

PVData_Rx = ConfigInfo.Ratio * VideoScreenScale * (PVData_X / PVDataXYMultiplier)   'PCN2988
PVData_Ry = ConfigInfo.Ratio * VideoScreenScale * (PVData_Y / PVDataXYMultiplier)  'PCN2988


    PVDataTrueRadiusCalc = Sqr(PVData_Rx ^ 2 + PVData_Ry ^ 2)
    
Exit Function
Err_Handler:
Select Case Err
    Case 6: Resume Next
    Case Else
        MsgBox Err & " - " & error$
End Select
End Function

Function AreaOf3Points(ByVal AX As Double, ByVal AY As Double, _
                       ByVal BX As Double, ByVal BY As Double, _
                       ByVal cx As Double, ByVal cy As Double) As Double
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCNant????
'Name    : AreaOf3Points
'Created : 5 January 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   : a coordinante AX,AY
'          b coordinante BX,BY
'          c coordinante CX,CY
'Desc    : Calculates the area of three given coordinates
'Usage   : Used when calculating the area of a profile
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    If AX < 0 Then AX = AX - AX: BX = BX - AX: cx = cx - AX
    If BX < 0 Then AX = AX - BX: BX = BX - BX: cx = cx - BX
    If cx < 0 Then AX = AX - cx: BX = BX - cx: cx = cx - cx

    If AY < 0 Then AY = AY - AY: BY = BY - AY: cy = cy - AY
    If BY < 0 Then AY = AY - BY: BY = BY - BY: cy = cy - BY
    If cy < 0 Then AY = AY - cy: BY = BY - cy: cy = cy - cy


    AreaOf3Points = Abs(((AX * BY) + (AY * cy) + (BX * cy) - (cx * BY) - (cy * AX) - (AY * BX)) / 2)
Exit Function
Err_Handler:
    MsgBox Err & "-" & error$
End Function


Sub DrawPVGraphs()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCNant????
'Name    : DrawPVGraphs
'Created : 8 August 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   :
'Desc    : Draws the Graphs, deciding which one where
'Usage   : When ever the graphs are needed to be updated are draw, eg when you move the scroll bar
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    Dim I As Integer
 
    Call DrawGraphImage(ImageRulerState, "Ruler", 0, GraphStartFrame, GraphEndFrame, 0, 0, 0, 0)
    If PrecisionVisionGraph.PVGraphExpandBtn.Tag = "0" Then
        Call PrecisionVisionGraph.GetGeneralPVGraphData(ImageGraphState(0).GraphType)
        Call DrawGraphImage(ImageGraphState(0), ImageGraphState(0).GraphType, 0, GraphStartFrame, GraphEndFrame, PVGraphGeneralXScale, PVGraphGeneralXOffset, PVXScaleLimitPerL, PVXScaleLimitPerR)
    Else
        For I = 0 To 5
            Call PrecisionVisionGraph.GetGeneralPVGraphData(ImageGraphState(I).GraphType)
            Call DrawGraphImage(ImageGraphState(I), ImageGraphState(I).GraphType, 0, GraphStartFrame, GraphEndFrame, PVGraphGeneralXScale, PVGraphGeneralXOffset, PVXScaleLimitPerL, PVXScaleLimitPerR)
        Next I
    End If

    
Exit Sub
Err_Handler:
    MsgBox Err & "-" & error$

End Sub

Sub DrawPVGraphsReport()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCNant3691
'Name    : DrawPVGraphsReport
'Created : 13 September 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   :
'Desc    : Draws the graph reports, deciding which one goes where, sets up the graph drawing to report
'        : draws the graph or graphs then sets the graph drawing back to standard screen graphs
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    Dim I As Integer
    Dim SaveScreenDrawingType As Integer
    
    SaveScreenDrawingType = ScreenDrawingType
    
    
    
    Set ReportDummyGraphImage = PrintPreviewForm.GraphContainer
    Call PrecisionVisionGraph.GetGeneralPVGraphData(ImageGraphState(0).GraphType)
    If ImageGraphState(0).GraphType = "Flat" Then ScreenDrawingType = 0
    Call DrawGraphImage(ImageGraphState(6), ImageGraphState(0).GraphType, 0, PrintPreviewForm.PreviewStartFrame, PrintPreviewForm.PreviewEndFrame, PVGraphGeneralXScale, PVGraphGeneralXOffset, PVXScaleLimitPerL, PVXScaleLimitPerR)
    
    If ImageGraphState(0).GraphType = "Flat" Then
        Set ReportDummyGraphImage.Picture = ImageGraphState(6).PictureImage
        'ReportDummyGraphImage.Visible = True
        ScreenDrawingType = SaveScreenDrawingType
    End If
    
    Set ReportDummyGraphImage = PrintPreviewForm.GraphXScaleContainer
    Call DrawGraphImage(ImageRulerState, "Ruler", 0, PrintPreviewForm.PreviewStartFrame, PrintPreviewForm.PreviewEndFrame, 0, 0, 0, 0)

Exit Sub
Err_Handler:
    MsgBox Err & "-" & error$
    'Resume 'Ant
End Sub

Sub DrawGraphImage(ImageGraphState As ImageGraphStateType, _
                   ByVal GraphType As String, _
                   ByVal DrawSpeed As Integer, _
                   ByVal StartFrame As Long, _
                   ByVal EndFrame As Long, _
                   ByVal XScale As Double, _
                   ByVal CentreOffset As Double, _
                   ByVal LeftLimitLine As Double, _
                   ByVal RightLimitLine As Double)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCN3373
'Name    : RefreshPicture
'Created : 9 March 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   : ImageGraphState - Type to hold the image pointer, and history of the last thing drawn on that image
'          GraphType       - The graph or ruler that is to be drawn on the passed image
'          DrawSpeed       - 0 for Bitmap Draw, 1 for picturebox draw (eg vb lines etc) PCN3691
'          StartFrame      - The first frame of the graph to be drawn
'          EndFrame        - The last frame of the graph to be drawn
'          XScale          - Scale width of the graph that is to be drawn
'          CentreOffset    - Offset of the graph that is drawn, relation to the centre
'Desc    : Direct memory access to picture boxes for fast realtime drawing
'Usage   : Used for Graphs
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

    Dim pict() As Byte  'Pointer that is passed to the different drawnig subfunctions, so they know where to draw
    Dim sa As SAFEARRAY2D, bmp As BITMAP    'Array that is manipulated and then coppied to the passed image
    Dim Units As String
    'Get the imformation of the image that it is passed to it, to see if its compatible for drawing on
    GetObjectAPI ImageGraphState.PictureImage.Picture, Len(bmp), bmp
                                                                            
    'The compatability is only colour depth, in this case the image has to be 24bit colour'''
'    If bmp.bmBitsPixel <> 24 Then   ' if not then exit                                      '
'        MsgBox " 24-bit bitmaps only", vbCritical                                           '
'        Exit Sub                                                                            '
'    End If                                                                                  '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ' have the local matrix array to point to the passed image bitmap
    With sa                                                         '
        .cbElements = 1                                             '
        .cDims = 2                                                  '
        .Bounds(0).lLbound = 0                                      '
        .Bounds(0).cElements = bmp.bmHeight                         '
        .Bounds(1).lLbound = 0                                      '
        .Bounds(1).cElements = bmp.bmWidthBytes                     '
        .pvData = bmp.bmBits                                        '
    End With                                                        '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    CopyMemory ByVal VarPtrArray(pict), VarPtr(sa), 4 'Point the pict pointer to the image arrray
    Dim Left As Double
    Dim Right As Double

Left = 0
Right = 200


XScale = XScale / 100

'CentreOffset = 180 / 100 * CentreOffset

If Trim(ConfigInfo.DistanceProcessMethod) <> "None" And ConfigInfo.DistanceStart <> -1 Then
    Units = MeasurementUnits
Else
    Units = "Frames"
End If

If XScale = 0 Then XScale = 1
'Erase the image with the previous drawn data, that includes grids, lines, graph etc
If GraphType = "Clear" Then
    Call ClearImage(pict)
    CopyMemory ByVal VarPtrArray(pict), 0&, 4
    ImageGraphState.PictureImage.Refresh
    Exit Sub
End If

Dim ExpDia As Double
Dim PrinterWidth As Integer


If ExpectedDiameter = 0 Then 'PCN3647
    ExpDia = 0
Else
    ExpDia = ExpectedDiameter 'PCN3647
End If

With ImageGraphState
    'If the previous graph that was drawn was flat and the current one isnot then clear image
    'for a completly clean image for the new graph to draw on


        If .PreviousGraphType = "Flat" And GraphType <> "Flat" Then Call ClearImage(pict)
        
        'If the last image drawn was the ruler on this image erase it
        If .PreviousGraphType = "Ruler" And ScreenDrawingType = 0 Then
            Call DrawRuler(pict, "Ruler", .PreviousUnits, .PreviousStartFrame, .PreviousEndFrame, 255, 255, 255)
            'Call DrawDistanceAndObsMarkings(pict, .PreviousStartFrame, .PreviousEndFrame, True)
        Else
            'Note: You do not have to erase the grid lines because they are always in the same place
            'Erase the image depending on the previously drawn graph. Condition for each graph.
            If ScreenDrawingType = 0 Then
                If .PreviousGraphType = "Ovality" Then
                    
                    'Call DrawLineGraphSingle(pict, PVOvalityOrigFullData, 0, .PreviousStartFrame, .PreviousEndFrame, .XScale, .Left, .Right, .CentreOffset, 255, 255, 255, True, "Ovality")
                    Call DrawLineGraphSingle(pict, PVOvalityFullData, 0, .PreviousStartFrame, .PreviousEndFrame, .XScale, .Left, .Right, .CentreOffset, 255, 255, 255, True, "Ovality")
                ElseIf .PreviousGraphType = "MedianDiameter" Then
                    Call DrawLineGraphDouble(pict, PVDiameterMedian, 0, .PreviousStartFrame, .PreviousEndFrame, .XScale, .Left, .Right, .CentreOffset, 255, 255, 255, True)
                ElseIf .PreviousGraphType = "Capacity" Then
                    Call DrawLineGraphSingle(pict, PVCapacityFullData, 0, .PreviousStartFrame, .PreviousEndFrame, .XScale, .Left, .Right, .CentreOffset, 255, 255, 255, True, "Capacity")
                ElseIf .PreviousGraphType = "Delta" Then
                    Call DrawLineGraphDouble(pict, PVDeltaFullMax, 0, .PreviousStartFrame, .PreviousEndFrame, .XScale, .Left, .Right, .CentreOffset, 255, 255, 255, True)
                    Call DrawLineGraphDouble(pict, PVDeltaFullMin, 0, .PreviousStartFrame, .PreviousEndFrame, .XScale, .Left, .Right, .CentreOffset, 255, 255, 255, True)
                ElseIf .PreviousGraphType = "XYDiameter" Then
                    Call DrawLineGraphDouble(pict, PVXDiameterFullData, 0, .PreviousStartFrame, .PreviousEndFrame, .XScale, .Left, .Right, .CentreOffset, 255, 255, 255, True)
                    Call DrawLineGraphDouble(pict, PVYDiameterFullData, 0, .PreviousStartFrame, .PreviousEndFrame, .XScale, .Left, .Right, .CentreOffset, 255, 255, 255, True)
                ElseIf .PreviousGraphType = "MaxMinDiameter" Then                                             '
                    Call DrawLineGraphDouble(pict, PVDiameterFullMax, 0, .PreviousStartFrame, .PreviousEndFrame, .XScale, .Left, .Right, .CentreOffset, 255, 255, 255, True)
                    Call DrawLineGraphDouble(pict, PVDiameterFullMin, 0, .PreviousStartFrame, .PreviousEndFrame, .XScale, .Left, .Right, .CentreOffset, 255, 255, 255, True)
                ElseIf .PreviousGraphType = "MaxDiameter" Then                                             '
                    Call DrawLineGraphDouble(pict, PVDiameterFullMax, 0, .PreviousStartFrame, .PreviousEndFrame, .XScale, .Left, .Right, .CentreOffset, 255, 255, 255, True)
'                ElseIf .PreviousGraphType = "Fractile" Then
'                    Call DrawLineGraphSingle(pict, PVFractile, 0, .PreviousStartFrame, .PreviousEndFrame, .XScale, .Left, .Right, .CentreOffset, 255, 255, 255, True, "Fractile")
                End If
                
                Call DrawLimitLine(pict, (.Right + .Left) / 2, .XScale, .Left, .Right, .CentreOffset, 255, 255, 255)
                Call DrawLimitLine(pict, .LeftLimitLine, .XScale, .Left, .Right, .CentreOffset, 255, 255, 255)
                Call DrawLimitLine(pict, .RightLimitLine, .XScale, .Left, .Right, .CentreOffset, 255, 255, 255)
                
            End If

            If GraphType = "Ovality" Then
                Left = -100: Right = 100
            ElseIf GraphType = "MedianDiameter" Then
                Left = 0: Right = ExpectedDiameter * 2
            ElseIf GraphType = "Capacity" Then
                Left = -100: Right = 100
            ElseIf GraphType = "Delta" Then
                Left = 0: Right = ExpectedDiameter
            ElseIf GraphType = "XYDiameter" Then
                Left = 0: Right = ExpectedDiameter * 2
            ElseIf GraphType = "MaxMinDiameter" Then
                Left = 0: Right = ExpectedDiameter * 2
            ElseIf GraphType = "MaxDiameter" Then
                Left = 0: Right = ExpectedDiameter * 2
'            ElseIf GraphType = "Fractile" Then
'                Left = 0: Right = ExpectedDiameter * 2 'PCN4235
            End If
        End If
 
        
If ScreenDrawingType = 0 Then
        'Store the current state of the image so as to know what is was when it needs to erase it.
        .PreviousGraphType = GraphType      'What graph or ruler was drawn
        .PreviousStartFrame = StartFrame    'Start frame that was drawn
        .PreviousEndFrame = EndFrame        'End frame that was drawn
        .PreviouspSpeed = DrawSpeed      'What speed the graph was drawn
        .XScale = XScale                'What scale the graph was drawn
        .Left = Left
        .Right = Right
        .CentreOffset = CentreOffset    'What centre offset the graph was drawn
        .PreviousUnits = Units
        .RightLimitLine = RightLimitLine
        .LeftLimitLine = LeftLimitLine
End If
End With

'Now the passed image has been erased, draw the new image.

'Draw the new ruler, or draw the grid on the graphs''
If GraphType = "Ruler" Then
    If ScreenDrawingType >= 1 Then
        Printer.DrawStyle = vbSolid
        Printer.DrawWidth = 3: PrintPreviewForm.DrawWidth = 1
        Call DrawRuler(pict, "Ruler", Units, StartFrame, EndFrame, 128, 128, 128)
    End If
    Call DrawRuler(pict, "Ruler", Units, StartFrame, EndFrame, 0, 0, 0)
    'Call DrawDistanceAndObsMarkings(pict, StartFrame, EndFrame, False)
Else '
    If ScreenDrawingType >= 1 Then
        Printer.DrawStyle = vbSolid
        Printer.DrawWidth = 1
        If GraphType <> "Flat" Then Call DrawRuler(pict, "Grid", Units, StartFrame, EndFrame, 128, 128, 128)
        Printer.DrawStyle = vbSolid
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''
    If ScreenDrawingType >= 1 Then
        PrinterWidth = Printer.DrawWidth
        Printer.DrawWidth = 3: PrintPreviewForm.DrawWidth = 1
        PrintPreviewForm.DrawMode = vbCopyPen
    End If
    Call DrawLimitLine(pict, (Right + Left) / 2, XScale, Left, Right, CentreOffset, 200, 200, 200)
    If ScreenDrawingType >= 1 Then Printer.DrawWidth = 5: PrintPreviewForm.DrawWidth = 1
    Call DrawLimitLine(pict, LeftLimitLine, XScale, Left, Right, CentreOffset, 255, 69, 0)
    Call DrawLimitLine(pict, RightLimitLine, XScale, Left, Right, CentreOffset, 255, 69, 0)
   
    
    'Note on line DrawLineGraph parameters - pict  - image pointer to be drawn on
    '                                        graph - which line graph is to be drawn
    '                                        DrawingSpeed - for scrolling, 0 draw all frames, 1 draw one to one, 2 draw two to one etc
    '                                        StartFrame - first frame to draw from for passed graph
    '                                        EndFrame - last frame to draw to from passed graph
    '                                        XScale - x axis scale of the graph to be drawn
    '                                        CentreOffset - how far off the centre should the graph be drawn
    '                                        Red,Green,Blue - Colour of the graph to be drawn
    '
    'Draw the apropraite graph, depending on what type was passed thru.''
    
  
    If GraphType = "Flat" Then Call DrawFlatGraph(pict, StartFrame, EndFrame)
    If GraphType = "Ovality" Then
        Call DrawLineGraphSingle(pict, PVOvalityFullData, DrawSpeed, StartFrame, EndFrame, XScale, Left, Right, CentreOffset, 0, 0, 255, False, "Ovality")
        'Call DrawLineGraphSingle(pict, PVOvalityOrigFullData, DrawSpeed, Startframe, EndFrame, XScale, Left, Right, CentreOffset, 125, 125, 125, False, "Ovality")
    End If
    If GraphType = "Capacity" Then Call DrawLineGraphSingle(pict, PVCapacityFullData, DrawSpeed, StartFrame, EndFrame, XScale, Left, Right, CentreOffset, 0, 0, 255, False, "Capacity")
    If GraphType = "MedianDiameter" Then Call DrawLineGraphDouble(pict, PVDiameterMedian, DrawSpeed, StartFrame, EndFrame, XScale, Left, Right, CentreOffset, 0, 0, 255, False)
    If GraphType = "Delta" Then                                         '
        Call DrawLineGraphDouble(pict, PVDeltaFullMax, DrawSpeed, StartFrame, EndFrame, XScale, Left, Right, CentreOffset, 0, 0, 255, False)
        Call DrawLineGraphDouble(pict, PVDeltaFullMin, DrawSpeed, StartFrame, EndFrame, XScale, Left, Right, CentreOffset, 0, 255, 0, False)
    End If                                                              '
    If GraphType = "XYDiameter" Then                                            '
        Call DrawLineGraphDouble(pict, PVXDiameterFullData, DrawSpeed, StartFrame, EndFrame, XScale, Left, Right, CentreOffset, 0, 0, 255, False)
        Call DrawLineGraphDouble(pict, PVYDiameterFullData, DrawSpeed, StartFrame, EndFrame, XScale, Left, Right, CentreOffset, 255, 0, 0, False)
    End If
    If GraphType = "MaxMinDiameter" Then                                            '
        Call DrawLineGraphDouble(pict, PVDiameterFullMax, DrawSpeed, StartFrame, EndFrame, XScale, Left, Right, CentreOffset, 0, 0, 255, False)
        Call DrawLineGraphDouble(pict, PVDiameterFullMin, DrawSpeed, StartFrame, EndFrame, XScale, Left, Right, CentreOffset, 255, 0, 0, False)
    End If
    If GraphType = "MaxDiameter" Then                                            '
        Call DrawLineGraphDouble(pict, PVDiameterFullMax, DrawSpeed, StartFrame, EndFrame, XScale, Left, Right, CentreOffset, 0, 0, 255, False)
    End If
'    If GraphType = "Fractile" Then Call DrawLineGraphSingle(pict, PVFractile, DrawSpeed, StartFrame, EndFrame, XScale, Left, Right, CentreOffset, 0, 0, 255, False, "Fractile")
    
    
    If ScreenDrawingType >= 1 Then Printer.DrawWidth = PrinterWidth: PrintPreviewForm.DrawWidth = 1

    '
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' clear the temporary array descriptor
' without destroying the local temporary array
CopyMemory ByVal VarPtrArray(pict), 0&, 4

'Tell vb something has changed with this image and refresh it.
ImageGraphState.PictureImage.Refresh

Exit Sub
Err_Handler:
    MsgBox error$
End Sub

Sub DrawDistanceAndObsMarkings(ByRef pict() As Byte, ByVal StartFrame As Long, ByVal EndFrame As Long, ByVal Clear As Boolean)
    Dim ObsIndex As Integer
    Dim NumberObs As Integer
    Dim Frame As Long
    Dim GraphMark As Double
    Dim rowupper As Long
    
    rowupper = UBound(pict, 2)  'Retrieve number of rows in passed image

    NumberObs = UBound(PipeObservations)
    
    For ObsIndex = 1 To NumberObs
        Frame = PipeObservations(ObsIndex).PipeObsFrameNo
        If Frame > StartFrame And Frame < EndFrame Then
            GraphMark = (Frame - StartFrame) / (EndFrame - StartFrame) * 100
            If Clear = True Then
                Call DrawMark(pict, GraphMark, 255, 255, 255, 3)
            Else
                If Trim(PipeObservations(ObsIndex).PipeObs) = "<<<--I-->>>" Then
                    Call DrawMark(pict, GraphMark, 217, 212, 130, 3, True)
                Else
                    Call DrawMark(pict, GraphMark, 130, 212, 217, 3, True)
                End If
            End If
        End If
    Next ObsIndex
    
    
    
  '  If Clear = True Then Call DrawMark(pict, 40, 255, 255, 255, 3)
  '  If Clear <> True Then Call DrawMark(pict, 40, 217, 212, 130, 3, True)

End Sub

Sub DrawMark(ByRef pict() As Byte, _
             ByVal Mark As Double, _
             ByVal Red As Byte, _
             ByVal Green As Byte, _
             ByVal Blue As Byte, _
             Optional Weight As Byte, _
             Optional Transparent)
On Error GoTo Err_Handler

Dim ColumnUpperDiv3 As Long  'Because Red,Green,Blue are columns then divide that by three for pixel coordinate column
Dim rowupper As Long 'Number of rows in image
Dim XScaleWidth As Double
Dim XScaleHeight As Double
Dim CentreMark As Long
Dim Y As Long
Dim X As Long
Dim Column As Long
Dim cpRed As Integer
Dim cpGreen As Integer
Dim cpBlue As Integer

XScaleWidth = 100
XScaleHeight = 100

If Not IsMissing(Transparent) Then
    Red = 255 - Red
    Blue = 255 - Blue
    Green = 255 - Green
End If


ColumnUpperDiv3 = (UBound(pict, 1) - 3) / 3 'Retrieve and calculate number of Pixels columns in image
rowupper = UBound(pict, 2)  'Retrieve number of rows in passed image

CentreMark = CDbl(rowupper) - (Mark / XScaleHeight * CDbl(rowupper))

If IsMissing(Transparent) Then
    For Y = CentreMark - 1 To CentreMark + 1
        If Y >= 0 And Y < rowupper Then
            For X = 0 To ColumnUpperDiv3 - 3
                Column = X * 3  'Calculate the starting colour column from the given X coordinate
                pict(Column + 0, Y) = Blue
                pict(Column + 1, Y) = Green
                pict(Column + 2, Y) = Red
            Next X
        End If
    Next Y
Else
    For Y = CentreMark - 1 To CentreMark + 1
        If Y >= 0 And Y < rowupper Then
            For X = 0 To ColumnUpperDiv3 - 3
                Column = X * 3  'Calculate the starting colour column from the given X coordinate
                cpBlue = CInt(pict(Column + 0, Y)) - Blue
                cpGreen = CInt(pict(Column + 1, Y)) - Green
                cpRed = CInt(pict(Column + 2, Y)) - Red
                If cpBlue < 0 Then cpBlue = 0
                If cpGreen < 0 Then cpGreen = 0
                If cpRed < 0 Then cpRed = 0
                pict(Column + 0, Y) = CByte(cpBlue)
                pict(Column + 1, Y) = CByte(cpGreen)
                pict(Column + 2, Y) = CByte(cpRed)
            Next X
        End If
    Next Y
End If


Exit Sub
Err_Handler:
Select Case Err
    Case 6: Exit Sub
    Case 340: Exit Sub 'control element does not exist
    Case Else
        MsgBox Err & " - " & error$
End Select
End Sub

Sub ClearImage(ByRef pict() As Byte)
'PCN3373
'Name    : ClearImage
'Created : 18 March 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   : pict()       - image pointer to draw on
'Desc    : Direct memory access to picture boxes for fast realtime drawing
'Usage   : Clear the image that is passed to it
'Note    : Sometimes drawing the last thing that was drawn on it white doesn't make sense
'          Example, drawing the whole flat graph as white, seems silly when you can do
'          a blanket fill with a colour. Does the same thing really.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

On Error GoTo Err_Handler

Dim r As Integer, c As Integer
Dim rowupper As Integer
Dim ColumnUpper As Integer



'If DrawSpeed = 0 Then 'PCN3691 is the graph drawn by vb on a picture box, if not then
    ColumnUpper = UBound(pict, 1)   'Number of columns to fill
    rowupper = UBound(pict, 2)      'Number of rows to fill
    
    For c = 0 To ColumnUpper  'Loop thru all the columns
        For r = 0 To rowupper 'Loop thru all the rows
            pict(c, r) = 255    'Filling them with white, really have to change this
        Next                    'to be filled with any background colour
    Next '''''''''''''''''''''''''
'Else
'    PictureBoxImage.Cls 'If using VB to draw then clear using cls.
'End If

Exit Sub
Err_Handler:
    MsgBox error$

End Sub
Sub DrawRuler(ByRef pict() As Byte, ByVal RuleGrid As String, ByVal Units As String, ByVal StartFrame As Long, ByVal EndFrame As Long, ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCN3373
'Name    : DrawRuler
'Created : 18 March 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   : pict()       - image pointer to draw on
'          RuleGrid     - switch to either tell function to draw the Ruler or the grid across the passed image
'                         "Ruler" for Ruler, or "Grid" for grid
'          StartFrame   - The Start frame to calculate the ruler markings and numbers
'          EndFrame     - The End frame to calculate the ruler marking and numbers
'          Red, Green, Blue - The colour to draw the ruler or grid
'Desc    : Direct memory access to picture boxes for fast realtime drawing
'Usage   : Used graph Ruler (Ruler) and graph (Grid)
'Note    :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

On Error GoTo Err_Handler

Dim RowEnd As Long  'The end for the image rows to draw ruler on
Dim CurrentRow As Long  'Current row to retrieve frame number or distance to be drawn on ruler
Dim r As Integer    'Row that is being drawn on
Dim GraphScale As Double    'Scale - to calculate the relationship between the image and graph frames,
                            '        if 520 Frames to be drawn and image has 520 lines relationship
                            '        is 1 to 1, if frames were 1040 then relationship is 2 to 1 etc

Dim rowupper As Integer     'Number of rows in the passed image
Dim ColumnUpper As Integer  'Number of columns in the passed image * 3, One for each colour, red, green, blue
Dim MajourDevision As Integer   'Number of majour devisions down the ruler and grid
Dim MinourDevision As Integer   'Number of minour devisions down the ruler and grid
Dim DrawDistance As Boolean 'To draw the ruler as distance or not (not is frames)
Dim RulerNumber As Double 'RulerNumber is frames or distance if DrawDistance is true
Dim TopOfMarker As Double
Dim BottomOfMinorMarker As Double
Dim BottomOfMajourMarker As Double
Dim BottomOfText As Double


If StartFrame < 1 Then StartFrame = 1   'Can't have start frame less the 0 then exit
If EndFrame <= StartFrame Then Exit Sub 'If the end frame eg no frame to display is equal or les then exit
If PVDataNoOfLines = 0 Then Exit Sub    'If no frames in array to display then exit

ColumnUpper = UBound(pict, 1)   'Retrieve number of columns in image passed.
rowupper = UBound(pict, 2)  'Retrieve number of rows in image passed
GraphScale = (EndFrame - StartFrame) / rowupper 'Calculate the relationship scale between the number of frames and the image rows

MinourDevision = rowupper / 35  'Number of minour devisions
MajourDevision = MinourDevision * 5 'Ration of Minour devision to minour devisions

'If after adding the scale and the end frame is more than the number of frames available, make them then last
If EndFrame > PVDataNoOfLines Then EndFrame = PVDataNoOfLines - 1
If EndFrame <= 1 Then Exit Sub 'Note to fix up, this is the reason when you scroll a scaled graph
                               'it just goes blank at a certain position
                               
'If units are frames the make DrawDistance false, else DrawDistance is true and ruler is in distance not frames
If Units = "Frames" Then
    DrawDistance = False
Else
    DrawDistance = True
End If

'If ScreenDrawingType >= 1 Then 'PCN3691 this is now determined on orientation on graph type
If ScreenDrawingOrientation = 1 Then
    TopOfMarker = 100
    BottomOfMinorMarker = 90
    BottomOfMajourMarker = 80
    BottomOfText = 40
    
Else
    TopOfMarker = 0
    BottomOfMinorMarker = 25
    BottomOfMajourMarker = 50
    BottomOfText = 60
End If


RowEnd = (EndFrame - StartFrame) / GraphScale 'Work out how many lines to draw down the image
'When switch set to ruler then draw the numbers of the ruler and the markings
If RuleGrid = "Ruler" Then
    For r = (rowupper - RowEnd) To rowupper 'Scan tru all the rows of the image to decide which ones need a line or number
        If r Mod MajourDevision = 0 Or r = rowupper Then      'If it does need a line (eg a given devision) then continue
            
            'Calculate which fame is going to be needed to look up frame number or distance
            CurrentRow = StartFrame + ((rowupper - r) * GraphScale)
            
            'If drawing distances then retrieve the distance number for that particular frame else just the frame number
            If DrawDistance Then
                RulerNumber = PVDistances(CurrentRow)
                If CurrentRow = PVDataNoOfLines Then RulerNumber = PVDistances(PVDataNoOfLines - 1)
            Else
                RulerNumber = CurrentRow
            End If
            
' PCN4164
'            If ScreenDrawingType >= 1 Then
'                RulerNumber = Round(RulerNumber, 1)
'            Else
'                RulerNumber = Round(RulerNumber, 1)
'            End If
            'RulerNumber = Format(RulerNumber, "##0.0") 'PCN4164
            
            PrinterClipOn = False
            Call DrawNumber(pict, RulerNumber, BottomOfText, CDbl(r - 2) / rowupper * 100, Red, Green, Blue) 'Draw that number at the row position
            PrinterClipOn = True
            
            Call DrawLine(pict, TopOfMarker, CDbl(r) / rowupper * 100, BottomOfMajourMarker, CDbl(r) / rowupper * 100, Red, Green, Blue) 'Draw the majour line on the ruler
        ElseIf r Mod MinourDevision = 0 Then 'If not a mojour line then check if minour line and draw
            Call DrawLine(pict, TopOfMarker, CDbl(r) / rowupper * 100, BottomOfMinorMarker, CDbl(r) / rowupper * 100, Red, Green, Blue)   'Draw minour line
        End If
    Next
'When switch set to grid then draw the grid across the graph, not the ruler or numbers
Else    'Same as above but instead of numbers and a little line for the Ruler, its a line right across the graph image
    For r = (rowupper - RowEnd) To rowupper
        If r Mod MajourDevision = 0 Then    'Double line for majour
            Call DrawLine(pict, 0, CDbl(r) / rowupper * 100, ColumnUpper, CDbl(r) / rowupper * 100, Red, Green, Blue)
            If ScreenDrawingType = 0 Then
                Call DrawLine(pict, 0, CDbl(r + 1) / rowupper * 100, ColumnUpper, CDbl(r + 1) / rowupper * 100, Red, Green, Blue)
            End If
        ElseIf r Mod MinourDevision = 0 Then    'Single line for minour
            Call DrawLine(pict, 0, CDbl(r) / rowupper * 100, ColumnUpper, CDbl(r) / rowupper * 100, Red, Green, Blue)
        End If
    Next
End If

Exit Sub
Err_Handler:
    MsgBox error$

End Sub

Sub DrawNumber(ByRef pict() As Byte, ByVal No As Double, ByVal CoordX As Double, ByVal CoordY As Double, ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCN3373
'Name    : DrawNumber
'Created : 18 March 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   : pict()  - image pointer to draw on
'          No      - Number to draw
'          CoordX  - X coordinate for the first digit of the number
'          CoordY  - Y coordinate for the first digit of the number
'          Red, Green, Blue - The colour to draw the graph, Background colour will clrear drawing
'Desc    : Direct memory access to picture boxes for fast realtime drawing
'Usage   : Used mainly on the graph Ruler
'Note    :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

On Error GoTo Err_Handler
Dim NumberString As String  'The string version of the double number passed
Dim DigitPosition   As Double 'Position of the digit currently being processed
Dim LengthOfNumber  As Double   'Length the number, thats all digits and decimal point
Dim CharacterToDisplay As String    'Character to display of the digit currently being processed
Dim Digit As Integer    'Integer version the character to display
Dim Offset As Double   'Used to centre the alignment of the number being displayed
Dim Spacing As Double  'Space between each digit to be displayed
Dim YScaleHeight As Double
Dim XScaleWidth As Double
Dim ColumnUpperDiv3 As Long  'Because Red,Green,Blue are columns then divide that by three for pixel coordinate column
Dim rowupper As Long 'Number of rows in image
Dim CoordYB As Double
Dim CoordXB As Double

YScaleHeight = 100
XScaleWidth = 100

NumberString = Format(No, "0.0")  'Convert double to string 'PCN
LengthOfNumber = Len(NumberString)  'Retrieve number of digits plus decimal if needed

Spacing = 1.5   'Set the spacing between the digits being drawn
If ScreenDrawingType >= 1 Then Spacing = 0.8
Offset = ((LengthOfNumber + 1) * Spacing) / 2   'Calulate the offset needed to align to centre
                                                                                        
    ColumnUpperDiv3 = (UBound(pict, 1) - 3) / 3 'Retrieve and calculate the number of columns in passed image
    rowupper = UBound(pict, 2)  'Retrieve the number of rows in the passed image
                                                                                      
For DigitPosition = 1 To LengthOfNumber 'Loop thru all the digits needed to be drawn'''''
    CharacterToDisplay = Mid(NumberString, DigitPosition, 1)    'Extract the current digit being prcesed. One digit at a time
    If CharacterToDisplay >= "0" And CharacterToDisplay <= "9" Then 'If the Digit is not a decimal then continue drawing it
        Digit = CInt(CharacterToDisplay)    'Convert to decimal to pass the DrawCharacter routine
    Else
        Select Case CharacterToDisplay
            Case ".": Digit = 11
            Case "k": Digit = 12
            Case "-": Digit = 13
            Case Else: Digit = 10 'Space
        End Select
    End If
    


 '   If ScreenDrawingType >= 1 Then 'PCN3691 this is decided on orientation now not Drawingtype
    If ScreenDrawingOrientation = 1 Then
        CoordYB = 100 - CoordY
        CoordXB = 100 - CoordX
'        DigitPosition = DigitPosition / XScaleWidth * ReportDummyGraphImage.width
'        Offset = Offset / XScaleWidth * ReportDummyGraphImage.width
'        Spacing = Spacing / XScaleWidth * ReportDummyGraphImage.width

    Else
        CoordYB = CoordY
        CoordXB = CoordX
'        DigitPosition = DigitPosition / XScaleWidth * CDbl(ColumnUpperDiv3)
'        Offset = Offset / XScaleWidth * CDbl(ColumnUpperDiv3)
'        Spacing = Spacing / XScaleWidth * CDbl(ColumnUpperDiv3)

    End If

    'CoordY = CoordY + (DigitPosition * Spacing) - Offset
    
    'Draw the actuall digit, with the Coordinate x y, and colour'
    Call DrawCharacter(pict, Digit, CoordXB, CoordYB + (DigitPosition * Spacing) - Offset, Red, Green, Blue)
Next ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Exit Sub
Err_Handler:
    MsgBox error$

End Sub

Sub DrawCharacter(ByRef pict() As Byte, ByVal Character As Integer, ByVal CoordX As Double, ByVal CoordY As Double, ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCN3373
'Name    : DrawCharacter
'Created : 18 March 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   : pict()   - image pointer to draw on
'          Character to draw
'          CoordX  - X coordinate for the character position
'          CoordY  - Y coordinate for the character position
'          Red, Green, Blue - The colour to draw the graph, Background colour will clrear drawing
'Desc    : Direct memory access to picture boxes for fast realtime drawing
'Usage   : Used mainly on the graph Ruler
'Note    : Character codes are as follows
'          Code      - 0 1 2 3 4 5 6 7 8 9 10 11 12
'          Character - 0 1 2 3 4 5 6 7 8 9  s .  k
'                                           p
'                                           a
'                                           c
'                                           e
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

On Error GoTo Err_Handler

Dim a As Integer    'Address coordinate of character pixel 0 to 7 (Column)
Dim b As Integer    'Address coordinate of character pixel 0 to 4 (Row)
Dim Column As Integer   'Column of pixel to be drawn on image (Not character pixel index)
Dim Row As Integer      'Row of pixel to be drawn on image (Not character pixel index)
Dim ColumnUpperDiv3 As Integer 'red, green, blue are a column each, div by 3 gives number pixel columns in image
Dim rowupper As Integer 'Number of rows in passed image
Dim YScaleHeight As Double
Dim XScaleWidth As Double
Dim RowDble As Double
Dim ColumnDble As Double
Dim CoordXB As Double
Dim CoordYB As Double

XScaleWidth = 100
YScaleHeight = 100

ColumnUpperDiv3 = (UBound(pict, 1) - 3) / 3 'Retrieve and calculate the number of columns in passed image
rowupper = UBound(pict, 2)  'Retrieve the number of rows in the passed image

If ScreenDrawingType >= 1 Then
    If ScreenDrawingOrientation = 1 Then 'PCN3691
        CoordXB = CoordY / XScaleWidth * ReportDummyGraphImage.width
        CoordYB = CoordX / YScaleHeight * ReportDummyGraphImage.height
    Else
        CoordXB = CoordX / XScaleWidth * ReportDummyGraphImage.width
        CoordYB = CoordY / YScaleHeight * ReportDummyGraphImage.height
    End If
Else
    CoordX = CoordX / XScaleWidth * CDbl(ColumnUpperDiv3)
    CoordY = CoordY / YScaleHeight * CDbl(rowupper)
End If

If ScreenDrawingType = 0 Then
    For a = 0 To 7  'Loop thru every column index of character''
        For b = 0 To 4  'Loop thru every row index of character'
            If NumberPic(Character, a, b) = 1 Then  'If indexed character pixel is 1 then draw it
                Column = a + CoordX     'Workout the offest along the x coordinate to place character pixel
                Row = (4 - b) + CoordY  'Workout the offset along teh y coordinate to place caharcter pixel
                'If out of the passed image area (out of bounds) then skip to next pixel
                If CoordX > 0 And CoordX < ColumnUpperDiv3 - 8 And CoordY > 0 And CoordY < rowupper - 4 Then
                    Column = Column * 3             'Image colour index, three columns per pixel
                    pict(Column + 0, Row) = Blue    'Blue is the first
                    pict(Column + 1, Row) = Green   'Green is the secound
                    pict(Column + 2, Row) = Red     'and red is the last of the three columns
                End If                                          '
            End If                                              '
        Next                                                    '
    Next ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Else
    Dim AX As Double
    Dim AY As Double
    Dim DotWidth As Single
    
    For a = 0 To 7  'Loop thru every column index of character''
        For b = 0 To 4  'Loop thru every row index of character'
            If NumberPic(Character, a, b) = 1 Then  'If indexed character pixel is 1 then draw it
                AX = a * 15
                AY = b * 15
                ColumnDble = (-AY) + CoordXB        'Workout the offest along the x coordinate to place character pixel
                RowDble = (AX - (8 * 15)) + CoordYB 'Workout the offset along the y coordinate to place caharcter pixel
                'If out of the passed image area (out of bounds) then skip to next pixel
                'If CoordXB > 0 And CoordXB < ReportDummyGraphImage.width - (8 * 15) And CoordYB > 0 And CoordYB < ReportDummyGraphImage.height - (4 * 15) Then
                'PrintPreviewForm.picReportPage.PSet (ColumnDble, RowDble), RGB(Red, Green, Blue)
                If ScreenDrawingType <> 1 Then
                    DotWidth = 0
                Else
                    DotWidth = 7
                End If
                
                Call DrawPrinterPoint(ColumnDble, RowDble, DotWidth, DotWidth, _
                                      RGB(Red, Green, Blue), _
                                      ReportDummyGraphImage.Left, _
                                      ReportDummyGraphImage.Top, _
                                      ReportDummyGraphImage.width, _
                                      ReportDummyGraphImage.height)
                'End If                                          '
            End If                                              '
        Next                                                    '
    Next ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End If

Exit Sub
Err_Handler:
    Select Case Err
     Case 340: Exit Sub ' control element doesnot exist
     Case Else: MsgBox error$ & "error " & Err
    End Select
End Sub


             


Sub DrawLine(ByRef pict() As Byte, _
             ByVal CoordAX As Double, ByVal CoordAY As Double, _
             ByVal CoordBX As Double, ByVal CoordBY As Double, _
             ByVal Red As Byte, _
             ByVal Green As Byte, _
             ByVal Blue As Byte)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCN3373
'Name    : DrawLine
'Created : 18 March 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   : pict()   - image pointer to draw on
'          CoordAX  - X coordinate for the line start
'          CoordAY  - Y coordinate for the line start
'          CoordBX  - X coordinate for the line end
'          CoordBY  - Y coordinate for the line end
'          Red, Green, Blue - The colour to draw the graph, Background colour will clrear drawing
'Desc    : Direct memory access to picture boxes for fast realtime drawing
'Usage   : Used by the line graphs eg Delta, Capacity, Ovality etc
'Note    :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

On Error GoTo Err_Handler

Dim X As Long    'Final pixel coordinate to be drawn along the X axis
Dim Y As Long    'Final pixel coordinate to be drawn along the Y axis

Dim a As Double 'The absolute X coordinate at anyone time along the drawn line
Dim b As Double 'The absolute Y coordinate at anyone time along the drawn line
Dim XSize As Double 'Relative size of line along the X axis
Dim YSize As Double 'Relative size of line along the Y axis
Dim Largest As Double   'Which is larger the X axis size or y Axis size
Dim StepX As Double 'Absolute step between pixels, if Largest then 1 to 1
Dim StepY As Double 'Absolute step between pixels, if Largest then 1 to 1
Dim I As Double 'Loop thru all the pixels to draw
Dim Column As Long 'Column coordinate for Image picture given, (Note: red,green,blue are considered as 3 column)
Dim XScaleWidth As Double
Dim XScaleHeight As Double

Dim ColumnUpperDiv3 As Long  'Because Red,Green,Blue are columns then divide that by three for pixel coordinate column
Dim rowupper As Long 'Number of rows in image

XScaleWidth = 100
XScaleHeight = 100

Dim AX As Double
Dim AY As Double
Dim BY As Double
Dim BX As Double

Dim XO, YO As Double

If ScreenDrawingType >= 1 Then
    If ScreenDrawingOrientation = 1 Then
        BX = (100 - CoordBY) / XScaleWidth * ReportDummyGraphImage.width
        AX = (100 - CoordAY) / XScaleWidth * ReportDummyGraphImage.width
        BY = (100 - CoordBX) / XScaleHeight * ReportDummyGraphImage.height
        AY = (100 - CoordAX) / XScaleHeight * ReportDummyGraphImage.height
    Else
        BX = (CoordBX) / XScaleWidth * ReportDummyGraphImage.width
        AX = (CoordAX) / XScaleWidth * ReportDummyGraphImage.width
        BY = (CoordBY) / XScaleHeight * ReportDummyGraphImage.height
        AY = (CoordAY) / XScaleHeight * ReportDummyGraphImage.height
    End If
'    ReportDummyGraphImage.Line (AX, AY)-(Bx, By), RGB(Red, Green, Blue)
    
    Call DrawPrinterLine(AX, AY, BX, BY, RGB(Red, Green, Blue), _
                         ReportDummyGraphImage.Left, _
                         ReportDummyGraphImage.Top, _
                         ReportDummyGraphImage.width, _
                         ReportDummyGraphImage.height)
    Exit Sub
End If

ColumnUpperDiv3 = (UBound(pict, 1) - 3) / 3 'Retrieve and calculate number of Pixels columns in image
rowupper = UBound(pict, 2)  'Retrieve number of rows in passed image

CoordBX = CoordBX / XScaleWidth * CDbl(ColumnUpperDiv3)
CoordAX = CoordAX / XScaleWidth * CDbl(ColumnUpperDiv3)
CoordBY = CoordBY / XScaleHeight * CDbl(rowupper)
CoordAY = CoordAY / XScaleHeight * CDbl(rowupper)


CoordBX = Round(CoordBX, 0)
CoordBY = Round(CoordBY, 0)
CoordAX = Round(CoordAX, 0)
CoordAY = Round(CoordAY, 0)


XSize = CoordBX - CoordAX   'Whats the distance between start x coord and end x coord
YSize = CoordBY - CoordAY   'Whats the distance between start y coord and end y coord

'Which is the absolute largest of the two, y coord distance or x coord distance
If Abs(XSize) > Abs(YSize) Then     '
    Largest = Abs(XSize)            '
Else                                '
    Largest = Abs(YSize)            '
End If                              '
'''''''''''''''''''''''''''''''''''''

If Largest = 0 Then Exit Sub    'If they are on top of each other, then no line to draw and exit

StepX = XSize / Largest 'Divide the drawing steps into the largest distance to draw for X axis
StepY = YSize / Largest 'Divide the drawing steps into the lasgest distance to draw for Y axis

a = CoordAX 'Keep track of the x coord positions in absolute terms, eg double
b = CoordAY 'Keep track of teh y coord positions in absolute terms, eg double

'Loop thru the line spanning across the largest distance needed to draw the line

For I = 0 To Largest - 1
    X = a   'Round the X coordinate to the nearest pixel
    Y = b   'Round teh Y coordinate to the nearest pixel
    
    'If the pixel to be drawn is outside the image bounds then move onto next pixel''
    If X > ColumnUpperDiv3 Or X < 0 Or Y > rowupper Or Y < 0 Then                   '
        a = a + StepX                                                               '
        b = b + StepY                                                               '
    Else    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Column = X * 3  'Calculate the starting colour column from the given X coordinate
        pict(Column + 0, Y) = Blue  'Blue is the first in the three columns for the pixel
        pict(Column + 1, Y) = Green 'Green is the secound
        pict(Column + 2, Y) = Red   'and Red is the thrid in the three columns for the pixel
        a = a + StepX   'Move onto the next pixel along the X coordiante
        b = b + StepY   'Move onto the next pixel along teh Y coordinate
    End If
Next I

Exit Sub
Err_Handler:
Select Case Err
    Case 6: Exit Sub
    Case 340: Exit Sub 'control element does not exist
    Case Else
        MsgBox Err & " - " & error$
End Select
End Sub

Sub DrawFlatGraph(ByRef pict() As Byte, ByVal StartFrame As Long, ByVal EndFrame As Long)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCN3373
'Name    : DrawFlatGraph
'Created : 18 March 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   : pict()       - image pointer to draw on
'          StartFrame   - Start frame to display from the graph.
'          EndFrame     - The ending frame to display from the graph
'Desc    : Direct memory access to picture boxes for fast realtime drawing, set up for all the majour graphs
'Usage   : Used to display any graph on any image.
'Note    :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

On Error GoTo Err_Handler

Dim ColumnEnd As Long   'Number of columns in flat graph and image
Dim RowEnd As Long  'Number of rows in the passed image

Dim RedValue As Integer     'Retrieved Red value from flat graph
Dim BlueValue As Integer    'Retrieved Blue value from flat graph
Dim GreenValue As Integer   'Retrieved Green value from flat graph
Dim ctimesby3 As Long   'Number of columns divided by 3, one colour for each column
Dim CurrentRow As Long  'Current Row retrieved from the flat graph
Dim CurrentColumn As Long
Dim r As Integer, c As Integer  'Image row to fill in
Dim r2 As Integer               'In theory this suppose to continue r after the graph ran out and filled in blank
Dim GraphScale As Double    'Relationship scale between Graph frames to display and image size to fill
Dim rowupper As Integer     'Number or rows in image
Dim ColumnUpper As Integer  'Number of columns in image
Dim PrintScale As Double


If StartFrame < 0 Then StartFrame = 0   'Start frame can never be less than 0 exit
If EndFrame <= StartFrame Then Exit Sub 'End is equal or less then exit aswell
If PVDataNoOfLines = 0 Then Exit Sub 'If there is no data then exit

rowupper = UBound(pict, 2)  'Retrieve number of rows in passed image
ColumnUpper = UBound(pict, 1) / 3



If ScreenDrawingOrientation = 1 Then
    GraphScale = (EndFrame - StartFrame) / ColumnUpper
Else
    GraphScale = (EndFrame - StartFrame) / rowupper 'Calculate Scale to fit the total number of
End If                                               'frames to be drawn into the image provided

'Make sure the end frame is less or equal to number of frames in graph
If EndFrame > PVDataNoOfLines Then EndFrame = PVDataNoOfLines - GraphScale
If EndFrame <= 0 Then Exit Sub 'This is suppose to make sure that the graph is not drawn
                               'over the edge of the number of frames available, unfortuantly
                               'it stops the whole graph from drawing when this condition is
                               'meet, need better way
If ScreenDrawingOrientation = 1 Then
    RowEnd = rowupper 'Retrieve number of columns that are to be draw, divided
    ColumnEnd = (EndFrame - StartFrame) / GraphScale 'Stop drawing lines when the end of graph is met
Else
    RowEnd = (EndFrame - StartFrame) / GraphScale 'Stop drawing lines when the end of graph is met
    ColumnEnd = ColumnUpper 'Retrieve number of columns that are to be draw, divided
End If                                   'by number of colours per pixel

If ScreenDrawingOrientation = 1 Then
    For r = 0 To RowEnd  'Loop thru the columns to draw the flat data
        For c = 0 To ColumnEnd - 1 'Loop thru the rows up to the last row or last frame, which ever comes first
            ctimesby3 = c * 3   'Index the column per colour.
            CurrentColumn = StartFrame + (c * GraphScale) 'Calculate frame data as CurrentRow
            RedValue = PVFlat3DRed(r, CurrentColumn)       'Get red value from flat graph
            GreenValue = PVFlat3DGreen(r, CurrentColumn)   'Get green value from flat graph
            BlueValue = PVFlat3DBlue(r, CurrentColumn)     'Get blue value from flat graph
                                                '''''''''
            pict(ctimesby3 + 0, r) = BlueValue  'Draw Blue value of pixel from the image corisponding from flat graph
            pict(ctimesby3 + 1, r) = GreenValue 'Draw Green value of pixel from the image corisponding from flat graph
            pict(ctimesby3 + 2, r) = RedValue   'Draw Red value of pixel from the image corisponding from flat graph
        Next c                                   '
    Next r ''''''''''''''''''''''''''''''''''''''''

Else
    For c = 0 To ColumnEnd - 1  'Loop thru the columns to draw the flat data
        ctimesby3 = c * 3   'Index the column per colour.
        For r = (rowupper - RowEnd) To rowupper   'Loop thru the rows up to the last row or last frame, which ever comes first
            CurrentRow = StartFrame + ((rowupper - r) * GraphScale)  'Calculate frame data as CurrentRow
            RedValue = PVFlat3DRed(c, CurrentRow)       'Get red value from flat graph
            GreenValue = PVFlat3DGreen(c, CurrentRow)   'Get green value from flat graph
            BlueValue = PVFlat3DBlue(c, CurrentRow)     'Get blue value from flat graph
                                                '''''''''
            pict(ctimesby3 + 0, r) = BlueValue  'Draw Blue value of pixel from the image corisponding from flat graph
            pict(ctimesby3 + 1, r) = GreenValue 'Draw Green value of pixel from the image corisponding from flat graph
            pict(ctimesby3 + 2, r) = RedValue   'Draw Red value of pixel from the image corisponding from flat graph
        Next                                    '
    Next ''''''''''''''''''''''''''''''''''''''''
    For c = 0 To ColumnUpper    '
        For r2 = 0 To (rowupper - RowEnd) '
            pict(c, r2) = 128              '
        Next                               '
    Next                        '
    '''''''''''''''''''''''''''''

End If
'Fill in what the graph doesn't cover with white, should really be a selected bacground colour

Exit Sub
Err_Handler:
Select Case Err
    Case 9 ' Subscript out of range
        Exit Sub
    Case Else
        MsgBox Err & " - " & error$
    
End Select
End Sub

Sub DrawLineGraphDouble(ByRef pict() As Byte, _
                        ByRef Graph() As Double, _
                        ByVal DrawingSpeed As Integer, _
                        ByVal StartFrame As Long, _
                        ByVal EndFrame As Long, _
                        ByVal XScale As Double, _
                        ByVal Left As Double, _
                        ByVal Right As Double, _
                        ByVal CentreOffset As Double, _
                        ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte, _
                        ByVal Clear As Boolean)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCN3373
'Name    : DrawLingGraphDouble
'Created : 18 March 2005
'Updated : 29 March 2005
'           This was only going to be DrawLineGraph, but delta is double values,
'           so had to make an exception for double type array
'Prg By  : Antony van Iersel
'Param   : pict()       - image pointer to draw on
'          graph()      - linegraph to draw, eg XY, DeltaMax, DeltaMin, Capactiy etc
'          DrawingSpeed - Speed of drawing, 0 draw all frames, 1 draw only one row of image at a time, 2 draw every secound etc
'          StartFrame   - Start frame to display from the graph.
'          EndFrame     - The ending frame to display from the graph
'          XScale       - Scalling of the Graph width.
'          CentreOffset - The offset from the center of the image to draw graphe
'          Red, Green, Blue - The colour to draw the graph, Background colour will clrear drawing
'Desc    : Direct memory access to picture boxes for fast realtime drawing, set up for all the majour graphs
'Usage   : Used to display any graph on any image.
'Note    :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

On Error GoTo Err_Handler

Dim CurrentRow As Double       'Row coordinate for start of the line to be draw
Dim CurrentColumn As Double    'Column coordinate for start of line draw
Dim NextRow As Double      'Row coordinate for end of the line to be drawn
Dim NextColumn As Double   'Column coordinate for the end of the line to be drawn

Dim CurrentFrame As Long 'Current frame to lookup the start value for the line to be drawn
Dim ScaleStep As Double 'The scale that makes the image rows (height) fit the number of frames to be drawn
                        'from start frame to end frame. eg if startframe = 0 and end frame = 520
                        'and there are 520 row on the picture image then scale is one to one, if
                        'endframe was 1040 then scale is one to two then every row has two graphlines

Dim rowupper As Long 'Number of rows in the given picture image
Dim Centre As Double    'The centre of the picture image horizontally, used to centre graph
Dim GraphRawWidth As Double
Dim GraphRawScale As Double

If PVDataNoOfLines = 0 Then Exit Sub    'If There is no Data exit


If Right = 0 Then Right = 100


GraphRawWidth = Right - Left
GraphRawScale = GraphRawWidth / 100


rowupper = UBound(pict, 2)  'Retrieve number of rows in the picture image

'Bounds checking, cant be less than frame zero, or greater than total frames, -1 is added
'to account for the drawing of the line from current to current+1      ''''''''''''''''''
If StartFrame < 0 Then StartFrame = 0                                  '


If (EndFrame - StartFrame) = 0 Then Exit Sub    'If there are no frames to be drawn exit

ScaleStep = rowupper / (EndFrame - StartFrame) 'Fit the scale, number of image rows x scale should equal the
                                               'number of frames from the graph to be displayed

If EndFrame > PVDataNoOfLines - 1 Then EndFrame = PVDataNoOfLines - 1  '

Dim CoordA As Double
Dim CoordB As Double


'Loop thru all the frames that need to be displayed from the graph, index by CurrentFrame
For CurrentFrame = StartFrame To EndFrame                                               '
    CurrentRow = (CurrentFrame - StartFrame) * ScaleStep    'Image row for start of the line
    NextRow = (CurrentFrame + 1 - StartFrame) * ScaleStep   'Image row for end of the line
                                                                                        
    CoordA = Graph(CurrentFrame)
    CoordB = Graph(CurrentFrame + 1)
    If (CoordA <> InvalidData) And (CoordB <> InvalidData) Then
    
        CoordA = (CoordA - Left) / GraphRawScale
        CoordB = (CoordB - Left) / GraphRawScale
        CoordA = ((CoordA - 50) / XScale) + 50
        CoordB = ((CoordB - 50) / XScale) + 50
    CurrentRow = rowupper - CurrentRow
    NextRow = rowupper - NextRow
    
    CurrentRow = CurrentRow / rowupper * 100
    NextRow = NextRow / rowupper * 100
                                                                                      '
    'To calulate the column (x) take the graph value, div by X scale and offset by centre and adjusted offset
        CurrentColumn = CoordA + CentreOffset 'Image column for the start of the line
        NextColumn = CoordB + CentreOffset  'Image column for the end of the line
        Call DrawLine(pict, CurrentColumn, CurrentRow, NextColumn, NextRow, Red, Green, Blue) 'Finally draw the line of the graph
    End If
Next CurrentFrame                                                                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Exit Sub
Err_Handler:
Select Case Err
    Case 9 ' Subscript out of range
        Exit Sub
    Case Else
        MsgBox Err & " - " & error$

End Select

End Sub

Sub DrawLineGraphSingle(ByRef pict() As Byte, _
                        ByRef Graph() As Single, _
                        ByVal DrawingSpeed As Integer, _
                        ByVal StartFrame As Long, _
                        ByVal EndFrame As Long, _
                        ByVal XScale As Double, _
                        ByVal Left As Double, _
                        ByVal Right As Double, _
                        ByVal CentreOffset As Double, _
                        ByVal Red_ As Byte, ByVal Green_ As Byte, ByVal Blue_ As Byte, _
                        ByVal Clear As Boolean, _
                        ByVal DrawType As String)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCN3373
'Name    : DrawLingGraphDouble
'Created : 18 March 2005
'Updated : 29 March 2005
'           This was only going to be DrawLineGraph, but delta is double values,
'           so had to make an exception for double type array
'Prg By  : Antony van Iersel
'Param   : pict()       - image pointer to draw on
'          graph()      - linegraph to draw, eg XY, DeltaMax, DeltaMin, Capactiy etc
'          DrawingSpeed - Speed of drawing, 0 draw all frames, 1 draw only one row of image at a time, 2 draw every secound etc
'          StartFrame   - Start frame to display from the graph.
'          EndFrame     - The ending frame to display from the graph
'          XScale       - Scalling of the Graph width.
'          CentreOffset - The offset from the center of the image to draw graphe
'          Red, Green, Blue - The colour to draw the graph, Background colour will clrear drawing
'Desc    : Direct memory access to picture boxes for fast realtime drawing, set up for all the majour graphs
'Usage   : Used to display any graph on any image.
'Note    :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

On Error GoTo Err_Handler

Dim CurrentRow As Double       'Row coordinate for start of the line to be draw
Dim CurrentColumn As Double    'Column coordinate for start of line draw
Dim NextRow As Double      'Row coordinate for end of the line to be drawn
Dim NextColumn As Double   'Column coordinate for the end of the line to be drawn

Dim CurrentFrame As Long 'Current frame to lookup the start value for the line to be drawn
Dim ScaleStep As Double 'The scale that makes the image rows (height) fit the number of frames to be drawn
                        'from start frame to end frame. eg if startframe = 0 and end frame = 520
                        'and there are 520 row on the picture image then scale is one to one, if
                        'endframe was 1040 then scale is one to two then every row has two graphlines

Dim rowupper As Long 'Number of rows in the given picture image
Dim Centre As Double    'The centre of the picture image horizontally, used to centre graph
Dim GraphRawWidth As Double
Dim GraphRawScale As Double
Dim CoordA As Double
Dim CoordB As Double

Dim Red As Byte
Dim Green As Byte
Dim Blue As Byte


If PVDataNoOfLines = 0 Then Exit Sub    'If There is no Data exit

If Right = 0 Then Right = 100
GraphRawWidth = Right - Left
GraphRawScale = GraphRawWidth / 100
Centre = 0

rowupper = UBound(pict, 2)  'Retrieve number of rows in the picture image
'Centre = UBound(pict, 1) / 3 / 2   'Calculate the centre horizontally, 3 is bytes per colour, 2 is centre

'Bounds checking, cant be less than frame zero, or greater than total frames, -1 is added
'to account for the drawing of the line from current to current+1      ''''''''''''''''''
If StartFrame < 0 Then StartFrame = 0                                  '


If (EndFrame - StartFrame) = 0 Then Exit Sub    'If there are no frames to be drawn exit

ScaleStep = rowupper / (EndFrame - StartFrame) 'Fit the scale, number of image rows x scale should equal the
                                               'number of frames from the graph to be displayed

If EndFrame > PVDataNoOfLines - 1 Then EndFrame = PVDataNoOfLines - 1  '

'Loop thru all the frames that need to be displayed from the graph, index by CurrentFrame
For CurrentFrame = StartFrame To EndFrame                                               '
    CurrentRow = (CurrentFrame - StartFrame) * ScaleStep    'Image row for start of the line
    NextRow = (CurrentFrame + 1 - StartFrame) * ScaleStep   'Image row for end of the line
    CoordA = Graph(CurrentFrame)
    CoordB = Graph(CurrentFrame + 1)
    
    If CoordA <> -100000 And CoordB <> -100000 Then                                                                                  '
        If DrawType = "Ovality" Then
            If Not Clear And _
                  (Graph(CurrentFrame) < 0 Or Graph(CurrentFrame + 1) < 0) Then
                Red = Blue_ / 4 * 3
                Green = Blue / 4 * 3
                Blue = Blue_
            Else
                Red = Red_
                Green = Green_
                Blue = Blue_
            End If
            CoordA = Abs(CoordA)
            CoordB = Abs(CoordB)
        Else
            Red = Red_
            Green = Green_
            Blue = Blue_
        End If
            
        CoordA = (CoordA - Left) / GraphRawScale
        CoordB = (CoordB - Left) / GraphRawScale
        CoordA = ((CoordA - 50) / XScale) + 50
        CoordB = ((CoordB - 50) / XScale) + 50
        
        'To calulate the column (x) take the graph value, div by X scale and offset by centre and adjusted offset
        CurrentColumn = CoordA + CentreOffset 'Image column for the start of the line
        NextColumn = CoordB + CentreOffset   'Image column for the end of the line
        
        CurrentRow = rowupper - CurrentRow: CurrentRow = CurrentRow / rowupper * 100
        NextRow = rowupper - NextRow: NextRow = NextRow / rowupper * 100
        
        
        Call DrawLine(pict, CurrentColumn, CurrentRow, NextColumn, NextRow, Red, Green, Blue) 'Finally draw the line of the graph
    End If
Next CurrentFrame                                                                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Exit Sub
Err_Handler:
Select Case Err
    Case 9 ' Subscript out of range
        Exit Sub
    Case 6: Resume Next
    Case Else
        MsgBox Err & " - " & error$
End Select

End Sub


Sub DrawLimitLine(ByRef pict() As Byte, _
                  ByVal LimitLine As Double, _
                  ByVal XScale As Double, _
                  ByVal Left As Double, _
                  ByVal Right As Double, _
                  ByVal CentreOffset As Double, _
                  ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCN3373
'Name    : DrawLimitLine
'Created : 18 March 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   : pict()       - image pointer to draw on
'          XScale       - Scalling of the Graph width.
'          CentreOffset - The offset from the center of the image to draw graphe
'          Red, Green, Blue - The colour to draw the graph, Background colour will clrear drawing
'Desc    : Direct memory access to picture boxes for fast realtime drawing, set up for all the majour graphs and lines and rulers
'Usage   : Used to display limit lines
'Note    :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

On Error GoTo Err_Handler

    Dim rowupper As Long 'Number of rows in the given picture image
    Dim GraphRawWidth As Double 'Width of the unscaled left and right values of the graph
    Dim GraphRawScale As Double 'Scaling to convert the graph values to 0 to 100
    Dim Column As Double 'Final scale value to draw the limit line to pass to linedraw

    If PVDataNoOfLines = 0 Then Exit Sub    'If There is no Data exit
    
    GraphRawWidth = Right - Left
    If GraphRawWidth <= 0 Then Exit Sub 'If the graph width to draw from is zero in width then exit
    GraphRawScale = GraphRawWidth / 100 'Scaling value calculated, 100 is the scale width of the draw line function

    rowupper = UBound(pict, 2)  'Retrieve number of rows in the picture image
 
    Column = LimitLine  'Draw position of limit line before scaleing
    Column = (Column - Left) / GraphRawScale    'Scale the line along the x axis
    Column = ((Column - 50) / XScale) + 50      'relative to the center of the graph
    Column = Column + CentreOffset 'Graph x offset adjust
   
    Call DrawLine(pict, Column, 0, Column, rowupper, Red, Green, Blue) 'Finally draw the limitline of the graph
    
    Exit Sub
Err_Handler:
    Select Case Err
        Case 9 ' Subscript out of range
            Exit Sub
        Case Else
            MsgBox Err & " - " & error$
      
    End Select

End Sub

Sub RefreshGraphXScale(ImageGraphState As ImageGraphStateType, _
                       ByVal XScale As Double, _
                       ByVal XOffset As Double)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCN3373
'Name    : RefreshGraphXScale
'Created : 1 April 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   : ImageGraphState - Type to hold the image pointer, and history of the last thing drawn on that image
'Desc    : Direct memory access to picture boxes for fast realtime drawing
'Usage   : To refresh current pointed graph with a new x scale
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    Call DrawGraphImage(ImageGraphState, _
                        ImageGraphState.PreviousGraphType, _
                        ImageGraphState.PreviouspSpeed, _
                        ImageGraphState.PreviousStartFrame, _
                        ImageGraphState.PreviousEndFrame, _
                        XScale, _
                        XOffset, _
                        ImageGraphState.LeftLimitLine, _
                        ImageGraphState.RightLimitLine)
Exit Sub
Err_Handler:
    MsgBox error$

End Sub
Function PVGraphFrameFromY(ByVal Y As Long) As Long
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCN3373
'Name    : PVGraphFrameFromY
'Created : 30 March 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   : Y position on image to calculate equivalent frame number
'Return  : Return the Frame number represented by the Y position on the image
'Desc    : When you have the Y coordinate of the graph or ruler image, return
'        : the equivalent frame number
'Usage   : Used when the ruler or graph is clicked with the mouse the the frame
'          number at that position is needed
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    Dim imageHeight As Double
    Dim numberOfFrames As Double
    Dim Frame As Long
    Dim FromGraph As Integer
    
    If PrecisionVisionGraph.PVGraphExpandBtn.Tag = "0" Then
        FromGraph = 0
    Else
        FromGraph = 1
    End If
    
    
    
    
    imageHeight = ImageGraphState(FromGraph).PictureImage.height
    numberOfFrames = ImageGraphState(FromGraph).PreviousEndFrame - ImageGraphState(FromGraph).PreviousStartFrame
    
    If numberOfFrames <= 0 Then
        PVGraphFrameFromY = 0
        Exit Function
    End If
    
    If Y = 0 Then
        PVGraphFrameFromY = ImageGraphState(FromGraph).PreviousStartFrame
        Exit Function
    End If
    
    Frame = (numberOfFrames / (imageHeight / Y)) + ImageGraphState(FromGraph).PreviousStartFrame
    If Frame > PVDataNoOfLines Then Frame = PVDataNoOfLines - 1
    If Frame < 0 Then Frame = 1

    PVGraphFrameFromY = Frame
Exit Function
Err_Handler:
    MsgBox error$

End Function

Function PVGraphYFromFrame(ByVal Frame As Long) As Long
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCN3373
'Name    : PVGraphFrameFromY
'Created : 30 March 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   : Y position on image to calculate equivalent frame number
'Return  : Return the Frame number represented by the Y position on the image
'Desc    : When you have the Y coordinate of the graph or ruler image, return
'        : the equivalent frame number
'Usage   : Used when the ruler or graph is clicked with the mouse the the frame
'          number at that position is needed
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    Dim imageHeight As Double
    Dim numberOfFrames As Double
    Dim FromGraph As Integer
            
    If PrecisionVisionGraph.PVGraphExpandBtn.Tag = "0" Then
        FromGraph = 0
    Else
        FromGraph = 1
    End If
            
            
    imageHeight = ImageGraphState(FromGraph).PictureImage.height
    numberOfFrames = ImageGraphState(FromGraph).PreviousEndFrame - ImageGraphState(FromGraph).PreviousStartFrame
    Frame = Frame - ImageGraphState(FromGraph).PreviousStartFrame
    
    
    If numberOfFrames <= 0 Or Frame = 0 Then
        PVGraphYFromFrame = 0
        Exit Function
    End If
    
    PVGraphYFromFrame = imageHeight / (numberOfFrames / Frame)

Exit Function
Err_Handler:
    MsgBox error$

End Function

Sub ClearAllGraphsAndRuler()
'PCN3373
'Name    : ClearImage
'Created : 18 March 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   : pict()       - image pointer to draw on
'Desc    : Direct memory access to picture boxes for fast realtime drawing
'Usage   : Clear the image that is passed to it
'Note    : Sometimes drawing the last thing that was drawn on it white doesn't make sense
'          Example, drawing the whole flat graph as white, seems silly when you can do
'          a blanket fill with a colour. Does the same thing really.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

On Error GoTo Err_Handler
Dim I As Integer


For I = 0 To 5
    Call DrawGraphImage(ImageGraphState(I), "Clear", 0, 0, 0, 0, 0, 0, 0)
Next I
Call DrawGraphImage(ImageRulerState, "Clear", 0, 0, 0, 0, 0, 0, 0)

Exit Sub
Err_Handler:
Select Case Err
    Case 6: Resume Next
    Case Else
        MsgBox Err & " - " & error$
End Select

End Sub

Function PVGraphXScaleUnitFromTwip(ByVal Index As Integer, ByVal Twip As Double) As Double
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCN3402
'Name    : PVGraphXScaleUnitFromTwip
'Created : 11 April March 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   : Index - which XScale are we looking at
'          Twip - Twip position to convert to percentage
'Return  : Percentage conversion
'Desc    : When you have the x coordinate on any PVXScale you can retrieve its
'          percentage value
'Usage   : When clicking on the PVXScale Graph, a percentage is needed for limit
'          lines and scale numbers
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    Dim Left As Double
    Dim Right As Double
    Dim XScale As Double
    Dim XOffset As Double
    Dim Unit As Double
    Dim GraphRawWidth As Double
    Dim GraphRawScale As Double
    
    Dim XScaleWidth As Double
    
    XScaleWidth = PrecisionVisionGraph.PVXScale(Index).width
        
    With ImageGraphState(Index)
        Left = .Left
        Right = .Right
        XScale = .XScale
        XOffset = .CentreOffset
    End With
    
    GraphRawWidth = Right - Left
    If GraphRawWidth = 0 Then GraphRawWidth = 1
    GraphRawScale = 100 / GraphRawWidth
    
    Unit = Twip / XScaleWidth * 100
    Unit = Unit - XOffset
    
    Unit = Unit - 50
    Unit = Unit * XScale
    Unit = Unit / GraphRawScale + (GraphRawWidth / 2) + Left
'    Unit = Unit + (Right - Left) / 2
    
    
    
    PVGraphXScaleUnitFromTwip = Unit
    
    
  
Exit Function
Err_Handler:
    MsgBox error$

End Function

Function PVGraphXScaleTwipFromUnit(ByVal Index As Integer, ByVal Unit As Double) As Double
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCN3402
'Name    : PVGraphXScaleUnitFromTwip
'Created : 11 April March 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   : Index - which XScale are we looking at
'          Percent - Percent value to convert to twip position
'Return  : Position of XScaleGraph in twips
'Desc    : When you have the percentage value of any PVXScale you can retrieve position
'          along the image as a twip
'Usage   : When drawing the limit lines on the PVXGraph scale need to no absolute
'           position in twips
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
        
    Dim Left As Double
    Dim Right As Double
    Dim XScale As Double
    Dim XOffset As Double
    Dim GraphRawWidth As Double
    Dim XScaleWidth As Double
    Dim GraphRawScale As Double
    
    XScaleWidth = PrecisionVisionGraph.PVXScale(Index).width
        
    With ImageGraphState(Index)
        Left = .Left
        Right = .Right
        XScale = .XScale
        XOffset = .CentreOffset
    End With
         
    GraphRawWidth = Right - Left
    If GraphRawWidth = 0 Then GraphRawWidth = 1
    GraphRawScale = GraphRawWidth / 100
    
    Unit = (Unit - Left) / GraphRawScale
    Unit = ((Unit - 50) / XScale) + 50
    
    Unit = Unit + XOffset 'Image column for the start of the line
    
    PVGraphXScaleTwipFromUnit = Unit / 100 * XScaleWidth
        
Exit Function
Err_Handler:
    MsgBox error$


End Function

Sub SetAndStoreLimitLines(Index As Integer, Button As Integer, X As Single)
On Error GoTo Err_Handler
Dim Limit As Double

If ImageGraphState(Index).GraphType = "Flat" Then Exit Sub

Call PrecisionVisionGraph.GetGeneralPVGraphData(ImageGraphState(Index).GraphType)

Limit = PVGraphXScaleUnitFromTwip(Index, X)


If ImageGraphState(Index).GraphType = "Ovality" Then
    Call DataEntryForm.SetDataEntryType("LimitLineChangeBoth", Limit, Index)
Else
    If Button = 1 Then
        Call DataEntryForm.SetDataEntryType("LimitLineChangeLeft", Limit, Index)
    Else
        Call DataEntryForm.SetDataEntryType("LimitLineChangeRight", Limit, Index)
    End If
End If

Exit Sub

'If imagegraphstate(Index).GraphType = "Ovality" Then
'    PVXScaleLimitPerL = Limit
'    PVXScaleLimitPerR = Limit
'Else
'    If Button = 1 Then
'        PVXScaleLimitPerL = Limit
'    Else
'        PVXScaleLimitPerR = Limit
'    End If
'End If
'
'Call PrecisionVisionGraph.SetFromPVXLimits(Index)
'Call PrecisionVisionGraph.RepositionPVXLimitMarkers   'PCN2680
'Call PrecisionVisionGraph.StoreLimitLinesInINI
'
''Call d3d_pvgraphtype(PVGraphType, PVXScaleLimitPerL, PVXScaleLimitPerR) 'PCN2337
'Call DrawPVGraphs


Exit Sub
Err_Handler:
    MsgBox error$
End Sub

Public Function Median(ByVal NumericArray As Variant, HistogramRequired As Boolean, ErrorInCalc As String, Optional HistogramResults As Variant) As Double
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : Median
'Created : 16 April 2005, PCN3489
'Updated :
'Prg By  :
'Param   :
'Desc    : Finds the median number in the array NumericArray.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'******************************************************'
'INPUT:   An Array of Numbers
'RETURNS: The statistical median of that array.
'         If invalid data is passed, i.e., a value that
'         is not an array, or the Array contains non-numeric
'         data, an error is raised
'EXAMPLE:
'        Dim vNumbers as Variant
'        dim dblMedian as double
'        vNumbers = array(4, 9, 1, 5, 3, 1, 7)
'        dblMedian = Median(vNumbers)
'****************************************************
On Error GoTo Err_Handler
Dim arrLngAns As Variant
Dim lngElement1 As Long
Dim lngElement2 As Long
Dim dblSum As Double
Dim dblAns As Double
Dim lngElementCount As Long

ErrorInCalc = ""

'sort array
arrLngAns = BubbleSortArray(NumericArray, ErrorInCalc)
If Not IsArray(arrLngAns) Then
'    Err.Raise 30000, , "Invalid Data Passed to function"
    ErrorInCalc = "Invalid Data Passed to function"
    Exit Function
End If

If UBound(arrLngAns) <= 2 Then
    ErrorInCalc = "Too few numbers to determine a meaningful median"
    Exit Function
End If
'Calculate the mean

lngElementCount = (UBound(arrLngAns) - LBound(arrLngAns)) + 1

    If UBound(arrLngAns) Mod 2 = 0 Then
        lngElement1 = (UBound(arrLngAns) / 2) + _
           (LBound(arrLngAns) / 2)
        
    Else
        lngElement1 = Int(UBound(arrLngAns) / 2) + _
           Int(LBound(arrLngAns) / 2) + 1
    End If
    
    If lngElementCount Mod 2 <> 0 Then
        dblAns = arrLngAns(lngElement1)
    Else
        lngElement2 = lngElement1 + 1
        dblSum = arrLngAns(lngElement1) + arrLngAns(lngElement2)
        dblAns = dblSum / 2
    End If

Median = dblAns

If HistogramRequired Then
    'Determine distribution - This is independant of the mean calculation but uses the same sorted numbers
    HistogramResults = HistogramDistribution(arrLngAns, dblAns, ErrorInCalc)
End If


Exit Function
Err_Handler:
    Select Case Err
        Case Else
'            MsgBox Err & error$
            ErrorInCalc = Err & error$
    End Select
End Function

Private Function BubbleSortArray(ByVal NumericArray As Variant, ErrorInCalc As String) As Variant
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : BubbleSortArray
'Created : 16 April 2005, PCN3489
'Updated :
'Prg By  :
'Param   :
'Desc    : Sort given array ready for finding the median number in the array NumericArray.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'http://www.freevbcode.com/ShowCode.Asp?ID=580
On Error GoTo Err_Handler
Dim vAns As Variant
Dim vTemp As Variant
Dim bSorted As Boolean
Dim lCtr As Long
Dim lCount As Long
Dim lStart As Long


vAns = NumericArray
  
If Not IsArray(vAns) Then
    BubbleSortArray = vbEmpty
    Exit Function
End If

On Error GoTo ErrorHandler

lStart = LBound(vAns)
lCount = UBound(vAns)

    bSorted = False
   
    Do While Not bSorted
      bSorted = True

      For lCtr = lCount - 1 To lStart Step -1
        If vAns(lCtr + 1) < vAns(lCtr) Then
          DoEvents
          bSorted = False
           vTemp = vAns(lCtr)
           vAns(lCtr) = vAns(lCtr + 1)
           vAns(lCtr + 1) = vTemp
         End If
      Next lCtr
      
    Loop
    
BubbleSortArray = vAns
Exit Function

ErrorHandler:
BubbleSortArray = vbEmpty
Exit Function
Err_Handler:
    Select Case Err
        Case Else
'            MsgBox Err & error$
            ErrorInCalc = Err & error$
    End Select
End Function


''Sub MedianDiameterForInspection(ErrorInCalc As String)
'''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'''Name    : MedianDiameterForInspection
'''Created : 16 April 2005, PCN3489
'''Updated :
'''Prg By  : Geoff Logan
'''Param   :
'''Desc    : Calculates the median diameter for the entire inspection.
'''Usage   :
'''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''On Error GoTo Err_Handler
''Dim NoOfMedians As Long
''Dim HistogramResults As Variant
''Dim MaxMinArray As Variant
''Dim MaxMinMedian As Double
''Dim HistYScaleFactor As Double
''
''ErrorInCalc = ""
''
''NoOfMedians = UBound(PVDiameterMedian)
''If NoOfMedians > 2 Then
''    With PrecisionVisionGraph
''        PVDiameterMedianForInspection = Median(PVDiameterMedian, True, ErrorInCalc, HistogramResults)
''        .HistogramOfDiameterGraph.AutoRedraw = True
''        .HistogramOfDiameterGraph.Cls
''        'Draw Title
'''        .HistogramOfDiameterGraph.CurrentX = 400
'''        .HistogramOfDiameterGraph.CurrentY = 10
'''        .HistogramOfDiameterGraph.FontSize = 12
'''        .HistogramOfDiameterGraph.FontBold = True
'''        .HistogramOfDiameterGraph.Print DisplayMessage("Diameter Histogram")
''        .DiameterHistogramTitle.Caption = DisplayMessage("Diameter Histogram")
''        .DiameterHistogramSubTitle.Caption = DisplayMessage("Diameter Histogram")
''        'Draw key
'''        .HistogramOfDiameterGraph.Line (20, 40)-(100, 80), 0, B
'''        .HistogramOfDiameterGraph.Line (25, 49)-(15, 50), RGB(255, 0, 0), BF 'Median
'''        .HistogramOfDiameterGraph.Line (25, 29)-(15, 70), RGB(0, 0, 255), BF 'Max
'''        .HistogramOfDiameterGraph.Line (25, 39)-(15, 90), RGB(0, 128, 0), BF 'Min
'''        .HistogramOfDiameterGraph.FontSize = 8
'''        .HistogramOfDiameterGraph.FontBold = False
'''        .HistogramOfDiameterGraph.CurrentX = 20
'''        .HistogramOfDiameterGraph.CurrentY = 50
'''        .HistogramOfDiameterGraph.Print DisplayMessage("Median Diameter")
''        .HistogramKeyMedian.Caption = DisplayMessage("Median Diameter")
'''        .HistogramOfDiameterGraph.CurrentY = 70
'''        .HistogramOfDiameterGraph.Print DisplayMessage("Max Diameter")
''        .HistogramKeyMax.Caption = DisplayMessage("Maximum")
'''        .HistogramOfDiameterGraph.CurrentY = 90
'''        .HistogramOfDiameterGraph.Print DisplayMessage("Min Diameter")
''        .HistogramKeyMin.Caption = DisplayMessage("Minimum")
''        'Draw the X Scale
''        Call DrawHistogramResultsXScale(HistogramResults, PVDiameterMedianForInspection, .HistogramOfDiameterGraph, ErrorInCalc)
''
''
''        'Determine Y scale based on the size of the Histogram at the median
''        If HistogramResults(150) <> 0 Then
''            HistYScaleFactor = HistogramResults(150) / (0.9 * 640) 'Set HistogramResults(150) bucket count at 90% of HistogramOfDiameter height when set to 640
''        Else
''            HistYScaleFactor = 1000 / (0.9 * 640) 'Set 1000 bucket count at 90% of HistogramOfDiameter height when set to 640
''        End If
''        'Draw the Histogram for the Median diameter.
''        Call DrawHistogramResults(HistogramResults, PVDiameterMedianForInspection, .HistogramOfDiameterGraph, RGB(255, 0, 0), HistYScaleFactor, ErrorInCalc)
''
''        'Calc and draw the Histogram for the Min diameter.
''        HistogramResults = Empty
''        MaxMinArray = MedianArrayForMaxMin(PVYDiameterFullData, ErrorInCalc)
''        MaxMinMedian = Median(MaxMinArray, True, ErrorInCalc, HistogramResults)
''        Call DrawHistogramResults(HistogramResults, PVDiameterMedianForInspection, .HistogramOfDiameterGraph, RGB(0, 128, 0), HistYScaleFactor, ErrorInCalc)
''
''        'Calc and draw the Histogram for the Max diameter.
''        HistogramResults = Empty
''        MaxMinArray = MedianArrayForMaxMin(PVXDiameterFullData, ErrorInCalc)
''        MaxMinMedian = Median(MaxMinArray, True, ErrorInCalc, HistogramResults)
''        Call DrawHistogramResults(HistogramResults, PVDiameterMedianForInspection, .HistogramOfDiameterGraph, RGB(0, 0, 255), HistYScaleFactor, ErrorInCalc)
''
''
''    End With
''End If
''
'''Display Median and Histogram results
''
''
''Exit Sub
''Err_Handler:
''Select Case Err
''    Case Else
''        ErrorInCalc = Err & " - " & error$
''End Select
''End Sub


Function HistogramDistribution(ByVal NumericArray As Variant, MedianOfArray As Double, ErrorInCalc As String) As Variant
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : HistogramDistribution
'Created : 16 April 2005, PCN3489
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Calculates the distribution of an array of numbers.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim HistArray As Variant
Dim NumArrayIndex As Long
Dim NumArrayIndexStart As Long
Dim NumArrayUBound As Long
Dim HistBucketIndex As Integer
Dim HistNoOfBuckets As Integer
Dim HistBucketSize As Double
Dim HistBucketIndexMaxValue As Double
Dim HistBucketStartValue As Double

ErrorInCalc = ""

If Not IsArray(NumericArray) Then
    HistogramDistribution = vbEmpty
    ErrorInCalc = "NumericArray is not an array"
    Exit Function
End If

HistBucketSize = MedianOfArray * 0.0025 'The default bucket size is 0.25% of the median
HistNoOfBuckets = 300
HistBucketStartValue = MedianOfArray - ((HistNoOfBuckets / 2) * HistBucketSize)
ReDim HistArray(HistNoOfBuckets)

NumArrayIndexStart = 1
NumArrayUBound = UBound(NumericArray)

For HistBucketIndex = 1 To HistNoOfBuckets
    HistBucketIndexMaxValue = HistBucketStartValue + (HistBucketIndex * HistBucketSize)
    If NumArrayIndex < NumArrayUBound Then
        For NumArrayIndex = NumArrayIndexStart To NumArrayUBound
            If NumericArray(NumArrayIndex) > HistBucketIndexMaxValue Then
                NumArrayIndexStart = NumArrayIndex
                Exit For
            Else
                HistArray(HistBucketIndex) = HistArray(HistBucketIndex) + 1 'Added this number to this Histogram bucket.
            End If
        Next NumArrayIndex
    End If
    If IsEmpty(HistArray(HistBucketIndex)) Then
        HistArray(HistBucketIndex) = 0
    End If
Next HistBucketIndex

HistogramDistribution = HistArray
HistArray = vbEmpty

Exit Function
Err_Handler:
HistogramDistribution = vbEmpty
Select Case Err
    Case Else
        ErrorInCalc = Err & " - " & error$
End Select
End Function

Sub DrawHistogramResults(ByVal HistResultsArray As Variant, MedianOfArray As Double, TargetPicBox As Control, DrawColour As Long, HistYScaleFactor As Double, ErrorInCalc As String)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : DrawHistogramResults
'Created : 16 April 2005, PCN3489
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Draws the distribution of an array of numbers.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim HistArray As Variant
'Dim NumArrayIndex As Long
'Dim NumArrayIndexStart As Long
'Dim NumArrayUBound As Long
Dim HistBucketIndex As Integer
Dim HistNoOfBuckets As Integer
Dim HistBucketSize As Double
Dim HistBucketIndexMaxValue As Double
Dim HistBucketStartValue As Double
Dim PicLineY1 As Long
Dim PicLineY2 As Long
Dim HistBucketDrawWidth As Integer
'Dim HistYScaleFactor As Double

ErrorInCalc = ""

If Not IsArray(HistResultsArray) Then
    ErrorInCalc = "HistResultsArray is not an array"
    Exit Sub
End If

HistBucketSize = MedianOfArray * 0.0025 'The default bucket size is 0.25% of the median
HistNoOfBuckets = 300
HistBucketStartValue = MedianOfArray - ((HistNoOfBuckets / 2) * HistBucketSize)

'TargetPicBox.Visible = True
'TargetPicBox.ZOrder 0
TargetPicBox.AutoRedraw = True
'TargetPicBox.Cls

HistBucketDrawWidth = 3
''Determine Y scale based on the size of the Histogram at the median
'HistBucketIndexMaxValue = HistResultsArray(HistNoOfBuckets / 2)
'HistYScaleFactor = HistBucketIndexMaxValue / (0.8 * TargetPicBox.height)

'Draw the base line
PicLineY1 = TargetPicBox.height - 30
TargetPicBox.Line (0, PicLineY1 + 1)-(HistBucketDrawWidth * HistNoOfBuckets, PicLineY1 + 1), 0
HistBucketIndexMaxValue = HistBucketDrawWidth * HistNoOfBuckets / 2
TargetPicBox.Line (HistBucketIndexMaxValue, PicLineY1 + 1)-(HistBucketIndexMaxValue, PicLineY1 + 5), 0

For HistBucketIndex = 1 To HistNoOfBuckets
'    HistBucketIndexMaxValue = HistBucketStartValue + (HistBucketIndex * HistBucketSize)
    If HistResultsArray(HistBucketIndex) > 0 Then
        PicLineY2 = PicLineY1 - HistResultsArray(HistBucketIndex) / HistYScaleFactor
        If DrawColour = RGB(255, 0, 0) Then 'Main draw colour
            TargetPicBox.Line ((HistBucketDrawWidth * HistBucketIndex), PicLineY1)-((HistBucketDrawWidth * HistBucketIndex - HistBucketDrawWidth), PicLineY2), DrawColour, BF
        ElseIf DrawColour = RGB(0, 0, 255) Then
            TargetPicBox.Line ((HistBucketDrawWidth * HistBucketIndex), PicLineY1)-((HistBucketDrawWidth * HistBucketIndex), PicLineY2), DrawColour
        Else
            TargetPicBox.Line ((HistBucketDrawWidth * HistBucketIndex - HistBucketDrawWidth), PicLineY1)-((HistBucketDrawWidth * HistBucketIndex - HistBucketDrawWidth), PicLineY2), DrawColour
        End If
    End If
Next HistBucketIndex


Exit Sub
Err_Handler:
Select Case Err
    Case Else
        ErrorInCalc = Err & " - " & error$
End Select
End Sub


Function MedianArrayForMaxMin(ByVal NumericArray As Variant, ErrorInCalc As String) As Variant
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : MedianArrayForMaxMin
'Created : 16 April 2005, PCN3489
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Extracts valid numbers from the NumericArray for use in determining
'          the median of the array.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim ValidDataArray() As Variant
Dim UBoundOfArray As Long
Dim arrayIndex As Long
Dim ValidDataArrayIndex As Long
Dim ValidDataLimit As Double
Dim ExpDiameter As Double 'The expected diameter

ErrorInCalc = ""

If Not IsArray(NumericArray) Then
    ErrorInCalc = "MedianArrayForMaxMin.NumericArray is not an array"
    Exit Function
End If

ValidDataLimit = -100
UBoundOfArray = UBound(NumericArray)
ValidDataArrayIndex = 0

ExpDiameter = ExpectedDiameter 'PCN3647

'No difference in storage between inches and mm
'If MeasurementUnits = "In" Or MeasurementUnits = "in" Or MeasurementUnits = "inches" Then
'    ExpDiameter = (ExpDiameter * 100)
'End If


For arrayIndex = 1 To UBoundOfArray
    If NumericArray(arrayIndex) > ValidDataLimit Then
        ValidDataArrayIndex = ValidDataArrayIndex + 1
        ReDim Preserve ValidDataArray(ValidDataArrayIndex)
        ValidDataArray(ValidDataArrayIndex) = ExpDiameter * (1 + (NumericArray(arrayIndex) / 100))
    End If

Next arrayIndex

MedianArrayForMaxMin = ValidDataArray

Exit Function
Err_Handler:
Select Case Err
    Case Else
        ErrorInCalc = Err & " - " & error$
End Select
End Function


Sub DrawHistogramResultsXScale(ByVal HistResultsArray As Variant, MedianOfArray As Double, TargetPicBox As Control, ErrorInCalc As String)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : DrawHistogramResultsXScale
'Created : 16 April 2005, PCN3489
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Draws the Xscale for distribution of an array of numbers.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim HistArray As Variant
Dim HistBucketIndex As Integer
Dim HistNoOfBuckets As Integer
Dim HistBucketSize As Double
Dim HistBucketIndexMaxValue As Double
Dim HistBucketStartValue As Double
Dim PicLineY1 As Long
Dim PicLineY2 As Long
Dim HistBucketDrawWidth As Integer
Dim HistXScaleValue As String
Dim HistXScaleMarker As Long

ErrorInCalc = ""

If Not IsArray(HistResultsArray) Then
    ErrorInCalc = "HistResultsArray is not an array"
    Exit Sub
End If

HistBucketSize = MedianOfArray * 0.0025 'The default bucket size is 0.25% of the median
HistNoOfBuckets = 300
HistBucketStartValue = MedianOfArray - ((HistNoOfBuckets / 2) * HistBucketSize)

TargetPicBox.AutoRedraw = True

HistBucketDrawWidth = 3
''Determine Y scale based on the size of the Histogram at the median
'HistBucketIndexMaxValue = HistResultsArray(HistNoOfBuckets / 2)
'HistYScaleFactor = HistBucketIndexMaxValue / (0.8 * TargetPicBox.height)

'Draw the base line
PicLineY1 = TargetPicBox.height - 30
TargetPicBox.Line (0, PicLineY1 + 1)-(HistBucketDrawWidth * HistNoOfBuckets, PicLineY1 + 1), 0
HistBucketIndexMaxValue = HistBucketDrawWidth * HistNoOfBuckets / 2
'TargetPicBox.Line (HistBucketIndexMaxValue, PicLineY1 + 1)-(HistBucketIndexMaxValue, PicLineY1 + 5), 0


TargetPicBox.CurrentY = PicLineY1
TargetPicBox.Font.name = "Arial"
TargetPicBox.Font.Size = 8
TargetPicBox.ForeColor = 0
TargetPicBox.Font.Bold = False
TargetPicBox.Font.Italic = False
HistXScaleValue = Format(MedianOfArray, "#.0")
TargetPicBox.CurrentX = HistBucketIndexMaxValue - (Len(HistXScaleValue) + 5)
'TargetPicBox.Print HistXScaleValue
'Draw other XScale markers
HistBucketStartValue = MedianOfArray - (MedianOfArray * 0.0025) * 3 * HistNoOfBuckets / 8
For HistBucketIndex = 0 To 6
    HistBucketIndexMaxValue = HistBucketStartValue + (MedianOfArray * 0.0025) * (HistBucketIndex * HistNoOfBuckets / 8)
    HistXScaleValue = Format(HistBucketIndexMaxValue, "#.0")
'    HistXScaleMarker = HistBucketDrawWidth * HistNoOfBuckets * (0.25 + (HistBucketIndex / 8))
    HistXScaleMarker = HistBucketDrawWidth * HistNoOfBuckets * (HistBucketIndex + 1) / 8
    TargetPicBox.Line (HistXScaleMarker, PicLineY1 + 1)-(HistXScaleMarker, PicLineY1 + 5), 0
    TargetPicBox.CurrentX = HistXScaleMarker - (Len(HistXScaleValue) + 5)
    TargetPicBox.CurrentY = PicLineY1 + 3
    TargetPicBox.Print HistXScaleValue
'    HistBucketIndexMaxValue = MedianOfArray + (MedianOfArray * 0.0025) * HistNoOfBuckets / 8
'    TargetPicBox.Line (HistBucketIndexMaxValue, PicLineY1 + 1)-(HistBucketIndexMaxValue, PicLineY1 + 5), 0
'    HistXScaleValue = Format(HistBucketIndexMaxValue, "#.0")
'    TargetPicBox.CurrentX = HistBucketDrawWidth * 3 * HistNoOfBuckets / 8 - (Len(HistXScaleValue) + 5)
'    TargetPicBox.Print HistXScaleValue
Next HistBucketIndex


Exit Sub
Err_Handler:
Select Case Err
    Case Else
        ErrorInCalc = Err & " - " & error$
End Select
End Sub

'PCN3217
Sub ToggleReportsPopUp(ByRef CurrentForm As Form)
On Error GoTo Err_Handler
        If CurrentForm.PopupReportsToolbar.Visible = True Then
           CurrentForm.PopupReportsToolbar.Visible = False
        Else
            CurrentForm.PopupReportsToolbar.Visible = True
            CurrentForm.PopupViewToolbar.Visible = False
        End If
Exit Sub
Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & error$
End Select
        
End Sub

'PCN3217
Sub ToggleViewSelectionPopUp(ByRef CurrentForm As Form)
On Error GoTo Err_Handler
        If CurrentForm.PopupViewToolbar.Visible = True Then
           CurrentForm.PopupViewToolbar.Visible = False
        Else
            CurrentForm.PopupViewToolbar.Visible = True
            CurrentForm.PopupReportsToolbar.Visible = False
        End If
Exit Sub
Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & error$
End Select
        
End Sub

Sub GraphSelect(ByVal GraphType As String, Container As Integer)
On Error GoTo Err_Handler

ImageGraphState(Container).GraphType = GraphType
    
PrecisionVisionGraph.FlatTM2.Visible = False 'PCN4171

Call PrecisionVisionGraph.UpdateGraphTitles

Call PrecisionVisionGraph.SetupPVGraphScreen(ImageGraphState(Container).GraphType)
Call PrecisionVisionGraph.GetGeneralPVGraphData(ImageGraphState(Container).GraphType) 'PCN2337 'PCN????
'Call PrecisionVisionGraph.d3d_pvgraphtype(imagegraphstate(0).GraphType, PVXScaleLimitPerL, PVXScaleLimitPerR) 'PCN2337
Call PVGraphsKeyForm.DisplayPVGraphsKey 'PCN2990 'PCN2818
If GraphType = "Flat" Then
    'Call PVFlat3DCalcCPP(1, PVDataNoOfLines)
    PrecisionVisionGraph.FlatTM2.Left = 1680
    PrecisionVisionGraph.FlatTM2.Visible = True ' PCN4171
End If
Call ClearAllGraphsAndRuler
Call DrawPVGraphs 'PCN3373
Call DrawPVXScale
Call PrecisionVisionGraph.XScaleLabel_Click(Container)
Call PrecisionVisionGraph.XScaleLabel_Click(Container)

Exit Sub
Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & error$
End Select
End Sub


''Sub OnscreenTextDrawAll(screenctrl As Control)
''On Error GoTo Err_Handler
''
''    Dim NumberOfOnscreenText As Integer
''    Dim Count As Integer
''    Dim XCoordA, YCoordA, XCoordB, YCoordB As Integer
''    Dim TextScreenWidth, TextScreenHeight As Integer
''
''
''    NumberOfOnscreenText = ClearLineScreen.OnScreenText.Count - 1
''    For Count = 1 To NumberOfOnscreenText
''        XCoordA = ClearLineScreen.OnScreenText(Count).Left
''        YCoordA = ClearLineScreen.OnScreenText(Count).Top
''        XCoordB = XCoordA + ClearLineScreen.OnScreenText(Count).width
''        YCoordB = YCoordA + ClearLineScreen.OnScreenText(Count).height
''        If ScreenDrawingType <> 0 Then
''            XCoordA = (((XCoordA - 1) * DrawSF) + ReportDummyGraphImage.Left) * RS
''            YCoordA = (((YCoordA - 2) * DrawSF) + ReportDummyGraphImage.Top) * RS
''            XCoordB = (((XCoordB - 1) * DrawSF) + ReportDummyGraphImage.Left) * RS
''            YCoordB = (((YCoordB - 2) * DrawSF) + ReportDummyGraphImage.Top) * RS
''        Else
''            XCoordA = (XCoordA - 1) * DrawSF
''            YCoordA = (YCoordA - 2) * DrawSF
''            XCoordB = (XCoordB - 1) * DrawSF
''            YCoordB = (YCoordB - 2) * DrawSF
''        End If
''
''        screenctrl.Line (XCoordA, YCoordA)-(XCoordB, YCoordB), ClearLineScreen.OnScreenText(Count).BackColor, BF
''        screenctrl.Line (XCoordA, YCoordA)-(XCoordB, YCoordB), vbBlack, B
''
''        screenctrl.Font = ClearLineScreen.OnScreenText(Count).Font
''        screenctrl.FontBold = ClearLineScreen.OnScreenText(Count).FontBold
''        screenctrl.FontItalic = ClearLineScreen.OnScreenText(Count).FontItalic
''        screenctrl.FontName = ClearLineScreen.OnScreenText(Count).FontName
''        If ScreenDrawingType <> 0 Then
''            screenctrl.FontSize = ClearLineScreen.OnScreenText(Count).FontSize * DrawSF / 15 * RS
''            TextScreenWidth = TextScreenWidth * DrawSF / 15 * RS
''        Else
''            screenctrl.FontSize = ClearLineScreen.OnScreenText(Count).FontSize * DrawSF
''            TextScreenWidth = TextScreenWidth * DrawSF
''        End If
''        screenctrl.FontStrikethru = ClearLineScreen.OnScreenText(Count).FontStrikethru
''        screenctrl.FontUnderline = ClearLineScreen.OnScreenText(Count).FontUnderline
''        screenctrl.ForeColor = ClearLineScreen.OnScreenText(Count).ForeColor
''
''        TextScreenWidth = screenctrl.TextWidth(ClearLineScreen.OnScreenText(Count).text)
''
''
''        screenctrl.CurrentX = (((XCoordB - XCoordA) - TextScreenWidth) / 2) + XCoordA
''        screenctrl.CurrentY = YCoordA
''
''        screenctrl.Print ClearLineScreen.OnScreenText(Count).text
''
''    Next Count
''    screenctrl.FillStyle = vbFSTransparent
''
''
''Exit Sub
''Err_Handler:
''    MsgBox Err & " - " & error$
''    Resume
''End Sub

Sub ProcessSingleImage()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PCN3194
'Name    : ProcessSingleImage()
'Created : 18 August 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   :
'Desc    : To process a single image be it jpg or bmp, this image is from the video window
'Usage   : Used for Graphs
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
    Dim sa As SAFEARRAY2D, bmp As BITMAP    'Array that is manipulated and then coppied to the passed image
    Dim pict() As Byte  'Pointer that is passed to the different drawnig subfunctions, so they know where to draw

    GetObjectAPI ClearLineScreen.DummyScreen.Picture, Len(bmp), bmp
    
    'The compatability is only colour depth, in this case the image has to be 24bit colour'''
    If bmp.bmBitsPixel <> 24 Then   ' if not then exit                                      '
        MsgBox " 24-bit bitmaps only", vbCritical                                           '
        Exit Sub                                                                            '
    End If                                                                                  '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' have the local matrix array to point to the passed image bitmap
    With sa                                                         '
        .cbElements = 1                                             '
        .cDims = 2                                                  '
        .Bounds(0).lLbound = 0                                      '
        .Bounds(0).cElements = bmp.bmHeight                         '
        .Bounds(1).lLbound = 0                                      '
        .Bounds(1).cElements = bmp.bmWidthBytes                     '
        .pvData = bmp.bmBits                                        '
    End With                                                        '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    CopyMemory ByVal VarPtrArray(pict), VarPtr(sa), 4 'Point the pict pointer to the image arrray

    Call Hough_ProcessSingleImage(pict(0, 0), _
                                  ConfigInfo.MediaWidth, _
                                  ConfigInfo.MediaHeight)

    ClearLineScreen.SnapShotScreen.AutoRedraw = True
    ClearLineScreen.SnapShotScreen.PaintPicture ClearLineScreen.DummyScreen.Picture, 0, 0, _
                                            ClearLineScreen.SnapShotScreen.width, _
                                            ClearLineScreen.SnapShotScreen.height
    ' clear the temporary array descriptor
    ' without destroying the local temporary array
    CopyMemory ByVal VarPtrArray(pict), 0&, 4

    'SavePicture ClearLineScreen.DummyScreen.Picture, "C:\\HiResImageProcess.bmp"


Exit Sub
Err_Handler:
    MsgBox error$

End Sub

Sub DrawPrinterLine(ByVal AX As Double, ByVal AY As Double, _
                    ByVal BX As Double, ByVal BY As Double, _
                    ByVal Colour As Long, _
                    ByVal Left As Double, _
                    ByVal Top As Double, _
                    ByVal width As Double, _
                    ByVal height As Double)
On Error GoTo Err_Handler

Dim Bottom As Double
Dim Right As Double
Dim StartPoint As Byte
Dim EndPoint As Byte

Bottom = Top + height
Right = Left + width

AX = AX + Left
BX = BX + Left
AY = AY + Top
BY = BY + Top

If AY < Top And BY < Top Then Exit Sub
If AY > Bottom And BY > Bottom Then Exit Sub
If AX < Left And BX < Left Then Exit Sub
If AX > Right And BX > Right Then Exit Sub

Dim Slope As Double
Dim TX As Double
Dim TY As Double

If AY > BY Then TX = AX: TY = AY: AX = BX: AY = BY: BX = TX: BY = TY

If AY < Top And BY > Top Then
    If AX = BX Then
        AY = Top
    Else
        Slope = (BX - AX) / (BY - AY)
        AX = (Top - AY) * Slope + AX
        AY = Top
    End If
End If

If AY < Bottom And BY > Bottom Then
    If AX = BX Then
        BY = Bottom
    Else
        Slope = (BX - AX) / (BY - AY)
        BX = (Bottom - AY) * Slope + AX
        BY = Bottom
    End If
End If

If AX < BX Then
    If AX < Left Then AX = Left
    If BX > Right Then BX = Right
End If

If BX < AX Then
    If BX < Left Then BX = Left
    If AX > Right Then AX = Right
End If

If ScreenDrawingType = 2 Then
    PrintPreviewForm.picReportPage.DrawWidth = 1
    PrintPreviewForm.picReportPage.Line (AX * RS, AY * RS)-(BX * RS, BY * RS), Colour
End If
If ScreenDrawingType = 1 Then
    Printer.Line (AX * RS, AY * RS)-(BX * RS, BY * RS), Colour
End If

Exit Sub
Err_Handler:
    MsgBox error$
    
End Sub

Sub DrawPrinterPoint(ByVal AX As Double, ByVal AY As Double, ByVal XSize As Double, ByVal YSize As Double, _
                     ByVal Colour As Long, _
                     ByVal Left As Double, _
                     ByVal Top As Double, _
                     ByVal width As Double, _
                     ByVal height As Double)
On Error GoTo Err_Handler

    AX = AX + Left
    AY = AY + Top
     
    If PrinterClipOn Then
        If AX < Left Or (AX + XSize) > (Left + width) Then Exit Sub
        If AY < Top Or (AY + YSize) > (Top + height) Then Exit Sub
    End If
    
    If ScreenDrawingType = 2 Then
        PrintPreviewForm.picReportPage.Line (AX * RS, AY * RS)-((AX + XSize) * RS, (AY + YSize) * RS), Colour, BF
    End If
    If ScreenDrawingType = 1 Then
        Printer.FillColor = Colour
        Printer.Line (AX * RS, AY * RS)-((AX + XSize) * RS, (AY + YSize) * RS), Colour, BF
    End If
    
                     
Exit Sub
Err_Handler:
    MsgBox error$
End Sub

Sub DrawPrinterCircle(ByVal X As Single, ByVal Y As Single, _
                     ByVal Size As Single, _
                     ByVal Colour As Long, _
                     ByVal StartArc As Single, _
                     ByVal EndArc As Single, _
                     ByVal Left As Double, _
                     ByVal Top As Double, _
                     ByVal width As Double, _
                     ByVal height As Double)
On Error GoTo Err_Handler

    X = X + Left
    Y = Y + Top
     
    If PrinterClipOn Then
        If X < Left Or (X + Size) > (Left + width) Then Exit Sub
        If Y < Top Or (Y + Size) > (Top + height) Then Exit Sub
    End If
    
    If ScreenDrawingType = 2 Then
        If StartArc = 0 And EndArc = 0 Then
            PrintPreviewForm.picReportPage.Circle (X * RS, Y * RS), Size * RS, Colour
        Else
            PrintPreviewForm.picReportPage.Circle (X * RS, Y * RS), Size * RS, Colour, StartArc, EndArc
        End If
    End If
    If ScreenDrawingType = 1 Then
        If StartArc = 0 And EndArc = 0 Then
            Printer.Circle (X * RS, Y * RS), Size * RS, Colour
        Else
            Printer.Circle (X * RS, Y * RS), Size * RS, Colour, StartArc, EndArc
        End If
    End If
    
                     
Exit Sub
Err_Handler:
    MsgBox error$
End Sub

Sub RecalculatePVData()
On Error GoTo Err_Handler
    
    Dim FileLoadError As Boolean
    Dim PVDVer As Single
    Dim msg, Style, Title, Help, Ctxt, Response, MyString
    Dim I As Long
        
    If PVDFileName = "" Or PVRecording Then Exit Sub
    PVDVer = GetPVDVer
    If PVDVer < 6.3 Then
        MsgBox DisplayMessage("Wrong PVD version, unable to recalculate data") & vbCritical
        Exit Sub
    End If
    
    msg = DisplayMessage("Do you want to recalculate PV Data?")   ' Define message.
    
    Style = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
    Title = DisplayMessage("Recalculate PV Data")   ' Define title.
    Response = MsgBox(msg, Style, Title)
             
    If Response = vbNo Then 'reset to config values, or 0 if its ver 6.3
        If PVDVer = 6.3 Then
            WLStartAngle = 0
            WLFinishAngle = 0
            ConfigInfo.WLStartAngle = WLStartAngle
            ConfigInfo.WLFinishAngle = WLFinishAngle
            If PVDVer = 6.3 Then ConfigInfo.PVDFileVersion = "V6.4"
            
            Call ScreenDrawing.DeleteWaterLevel
            WaterLevelIgnoreCenter = False
            Call ClearLineScreen.SetUnsetWaterLevel(WaterLevelIgnoreCenter, False)
        Else
            WLStartAngle = ConfigInfo.WLStartAngle
            WLFinishAngle = ConfigInfo.WLFinishAngle
        End If
        Exit Sub
    End If
    
    ConfigInfo.WLStartAngle = WLStartAngle
    ConfigInfo.WLFinishAngle = WLFinishAngle
    
    If PVDVer = 6.3 Then ConfigInfo.PVDFileVersion = "V6.4"
    
    Call CLPProgressBar.ProgressBarInitialise(DisplayMessage("Processing Data"))
    DoEvents
    
    Call LoadFullPVDataFromFile
    For I = 0 To (PVDataNoOfLines / 100)
        Call CLPProgressBar.ProgressBarPosition((I / PVDataNoOfLines * 100))
        DoEvents
    
        Call ScreenDrawing.PVCentreCalcCPP((I * 100), 99 + (I * 100)) 'PCN3194
        Call ScreenDrawing.PVFlat3DCalcCPP((I * 100), 99 + (I * 100)) 'PCN3513
        Call ScreenDrawing.PVCapacityCalcCPP((I * 100), 99 + (I * 100)) 'PCN3540
        Call ScreenDrawing.PVXYDiameterCalcCPP((I * 100), 99 + (I * 100)) 'PCN3540
        Call ScreenDrawing.PVDeltaMaxMinCalcCPP((I * 100), 99 + (I * 100)) 'pcn3540
        Call ScreenDrawing.PVDiameterMaxMinCalcCPP((I * 100), 99 + (I * 100)) 'PCN3540
        Call ScreenDrawing.PVDiameterMedianCalcCPP((I * 100), 99 + (I * 100)) 'PCN3540
    '    Call ScreenDrawing.PVFractileCalcCPP((i * 100), 99 + (i * 100)) 'PCN3540
        Call ScreenDrawing.PVOvalityCalcCPP((I * 100), 99 + (I * 100)) 'PCN3540
        ClearLineScreen.UpdateThreeDimColour
    Next I
    Call PageFunctions.SaveCentreCalculations
    Call CLPProgressBar.ProgressBarPosition(1)
    DoEvents
Exit Sub
Err_Handler:
    MsgBox error$
End Sub



Sub RenderLine(PicCtrlName As Control, _
                    ByVal AX As Single, ByVal AY As Single, _
                    ByVal BX As Single, ByVal BY As Single, _
                    ByVal Colour As Long)
On Error GoTo Err_Handler

If ScreenDrawingType >= 1 Then
    Call DrawPrinterLine(AX * DrawSF, AY * DrawSF, _
                        BX * DrawSF, BY * DrawSF, _
                        Colour, _
                        ReportDummyGraphImage.Left, _
                        ReportDummyGraphImage.Top, _
                        ReportDummyGraphImage.width, _
                        ReportDummyGraphImage.height)
Else
    PicCtrlName.Line (AX * DrawSF, AY * DrawSF)-(BX * DrawSF, BY * DrawSF), Colour
End If

Exit Sub
Err_Handler:
Select Case Err
    Case 6: Resume Next
    Case Else
        MsgBox Err & " - " & error$
End Select
End Sub

Sub RenderBox(PicCtrlName As Control, _
                    ByVal AX As Single, ByVal AY As Single, _
                    ByVal BX As Single, ByVal BY As Single, _
                    ByVal Colour As Long)
On Error GoTo Err_Handler

If ScreenDrawingType >= 1 Then
    Call DrawPrinterLine(AX * DrawSF, AY * DrawSF, _
                        BX * DrawSF, AY * DrawSF, _
                        Colour, _
                        ReportDummyGraphImage.Left, _
                        ReportDummyGraphImage.Top, _
                        ReportDummyGraphImage.width, _
                        ReportDummyGraphImage.height)
    Call DrawPrinterLine(BX * DrawSF, AY * DrawSF, _
                        BX * DrawSF, BY * DrawSF, _
                        Colour, _
                        ReportDummyGraphImage.Left, _
                        ReportDummyGraphImage.Top, _
                        ReportDummyGraphImage.width, _
                        ReportDummyGraphImage.height)
    Call DrawPrinterLine(BX * DrawSF, AY * DrawSF, _
                        AX * DrawSF, BY * DrawSF, _
                        Colour, _
                        ReportDummyGraphImage.Left, _
                        ReportDummyGraphImage.Top, _
                        ReportDummyGraphImage.width, _
                        ReportDummyGraphImage.height)
    Call DrawPrinterLine(AX * DrawSF, BY * DrawSF, _
                        AX * DrawSF, AY * DrawSF, _
                        Colour, _
                        ReportDummyGraphImage.Left, _
                        ReportDummyGraphImage.Top, _
                        ReportDummyGraphImage.width, _
                        ReportDummyGraphImage.height)

Else
    PicCtrlName.Line (AX, AY)-(BX, AY), Colour
    PicCtrlName.Line (BX, AY)-(BX, BY), Colour
    PicCtrlName.Line (BX, BY)-(AX, BY), Colour
    PicCtrlName.Line (AX, BY)-(AX, AY), Colour
End If

Exit Sub
Err_Handler:
    MsgBox error$
End Sub


Sub RenderPoint(PicCtrlName As Control, _
                ByVal X As Single, _
                ByVal Y As Single, _
                ByVal Colour As Long)
On Error GoTo Err_Handler

Dim DotWidth As Single

If ScreenDrawingType <> 1 Then
    DotWidth = 0
Else
    DotWidth = 10
End If

If ScreenDrawingType >= 1 Then
    Call DrawPrinterPoint(X * DrawSF, Y * DrawSF, DotWidth, DotWidth, _
                          Colour, _
                          ReportDummyGraphImage.Left, _
                          ReportDummyGraphImage.Top, _
                          ReportDummyGraphImage.width, _
                          ReportDummyGraphImage.height)
Else
    PicCtrlName.PSet (X, Y), Colour
End If

Exit Sub
Err_Handler:
    MsgBox error$

End Sub
                
Sub RenderCircle(PicCtrlName As Control, _
                 ByVal X As Single, _
                 ByVal Y As Single, _
                 ByVal Size As Single, _
                 ByVal Colour As Long, _
                 Optional ByVal StartArc As Single, _
                 Optional ByVal EndArc As Single)
On Error GoTo Err_Handler
    If ScreenDrawingType >= 1 Then
    If ScreenDrawingType = 1 Then PicCtrlName.DrawWidth = 6
    Call DrawPrinterCircle(X * DrawSF, Y * DrawSF, _
                        Size * DrawSF, _
                        Colour, _
                        StartArc, EndArc, _
                        ReportDummyGraphImage.Left, _
                        ReportDummyGraphImage.Top, _
                        ReportDummyGraphImage.width, _
                        ReportDummyGraphImage.height)
Else
    PicCtrlName.Circle (X * DrawSF, Y * DrawSF), Size * DrawSF, Colour, StartArc, EndArc
End If
               
Exit Sub
Err_Handler:
    MsgBox error$
End Sub

Sub SetupDrawingRoutine(ByVal Action As String)
On Error GoTo Err_Handler

    CLPScreenDrawState = ""
    CLPScreenActionPrevious = CLPScreenAction
    Select Case Action
        Case "DrawSingleLine"
            CLPScreenAction = "DrawLine": Call ClearLineScreen.SetupMouseIcon(110)
            ClearLineScreen.DimenResults.Visible = True: ClearLineScreen.AreaResults.Visible = False
            
        Case "DrawPolyLine"
            CLPScreenAction = "DrawPolyLine":  Call ClearLineScreen.SetupMouseIcon(112)
            ClearLineScreen.DimenResults.Visible = True: ClearLineScreen.AreaResults.Visible = False
            
        Case "DrawCalibrationLine"
            CLPScreenAction = "DrawCalibrationLine": Call ClearLineScreen.SetupMouseIcon(110)
            ClearLineScreen.DimenResults.Visible = True: ClearLineScreen.AreaResults.Visible = False
            
        Case "DrawOrthogonalLine"
            CLPScreenAction = "DrawOrthogonalLine": Call ClearLineScreen.SetupMouseIcon(111)
            ClearLineScreen.DimenResults.Visible = True: ClearLineScreen.AreaResults.Visible = False
            
        Case "MoveAll": CLPScreenAction = "MoveAll": Call ClearLineScreen.SetupMouseIcon(108)
        Case "Delete": Call DeleteAll
        Case "DrawWaterLevel": CLPScreenAction = "DrawWaterLevel": Call ClearLineScreen.SetupMouseIcon(117)
        Case "DrawText": CLPScreenAction = "DrawText": ClearLineScreen.SnapShotScreen.MousePointer = 3
        Case "DrawMaskBox": CLPScreenAction = "DrawMaskBox": Call ClearLineScreen.SetupMouseIcon(118) 'Ignore Text Region
        Case "DrawHorCalibrationLine"
            CLPScreenAction = "DrawHorCalibrationLine": Call ClearLineScreen.SetupMouseIcon(110)
            ClearLineScreen.DimenResults.Visible = True: ClearLineScreen.AreaResults.Visible = False
            
        Case "Dimension"
            CLPScreenAction = "DrawDimension": Call ClearLineScreen.SetupMouseIcon(104)
            ClearLineScreen.DimenResults.Visible = True: ClearLineScreen.AreaResults.Visible = False
            
        Case "DrawIntegratedArea"
            CLPScreenAction = "DrawIntegratedArea": Call ClearLineScreen.SetupMouseIcon(101)
            ClearLineScreen.AreaResults.Visible = True:             ClearLineScreen.DimenResults.Visible = False
            

    End Select
Exit Sub
Err_Handler:
    MsgBox error$
End Sub

Sub DrawActionWaterLevel(screenctrl As Control)
On Error GoTo Err_Handler

Dim LineDistance As Double
Dim OrthoX As Double
Dim OrthoY As Double
Dim CurrentX As Single
Dim CurrentY As Single
Dim DistToCentre As Double
Dim AngleAtHyp As Double
Dim ExpectedRadius As Double
Dim AngleOfCur As Double
Dim WLStartX As Double, WLStartY As Double
Dim WLEndX As Double, WLEndY As Double
Dim WLStartXExt As Double, WLStartYExt As Double
Dim WLEndXExt As Double, WLEndYExt As Double
Dim I As Integer
Dim ErrorStr As String
Dim FileLoadError As Boolean




ExpectedRadius = ExpectedDiameter / 2

If CLPScreenDrawAction = "LeftClick" And CLPScreenDrawState = "RubberBand" Then
    CLPScreenDrawState = ""
    With WaterLevelLine(2)
        WLStartAngle = CalAsin(0, 0, .StartX - ShapeCentreX, -.StartY + ShapeCentreY) + (PI / 2)
        WLFinishAngle = CalAsin(0, 0, .EndX - ShapeCentreX, -.EndY + ShapeCentreY) + (PI / 2)
    End With
    If WLStartAngle > (2 * PI) Then WLStartAngle = WLStartAngle - (2 * PI)
    If WLFinishAngle > (2 * PI) Then WLFinishAngle = WLFinishAngle - (2 * PI)
    
    'If the water level been set is nonsince then exit
    Call DeleteWaterLevel   'Dont need to draw once set
    
    If WLStartAngle = 0 And WLFinishAngle = 0 Then
        screenctrl.AutoRedraw = True
        Call ScreenDrawing.DrawProfilesStartToFinish(screenctrl, True)
        screenctrl.AutoRedraw = False
        Exit Sub
    End If
       
    Call ControlsScreen.SwitchToVideoMode 'Switch to video mode to show water setting
    
    Call ClearLineScreen.SetWaterLevelinPipe(WLStartAngle, WLFinishAngle) 'set the ignore profile point and pass to c++
    WaterLevelIgnoreCenter = True 'PCNLS190603
    Call ClearLineScreen.SetUnsetWaterLevel(WaterLevelIgnoreCenter, False) 'PCN3219 WaterLevelIgnoreProfile) 'PCNLS190603waterlevel
    
    'There is no function to blindly turn on overlay tuning, there is a toggle thou
    Call SetupVideoDisplayForPVTuning                           '
    Call ControlsScreen.SetupControlDisplayIP                   '
    If isopen("AutoTune") Then Call AutoTune.SetupIPCheckBox    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If PVDFileName <> "" Then
        Call RecalculatePVData 'If there is PVD data then it will ask for changes to be set
    Else
        ConfigInfo.WLStartAngle = WLStartAngle
        ConfigInfo.WLFinishAngle = WLFinishAngle
    End If
    Call PageFunctions.SaveToFilePipeAndConfigInfo("ConfigInfo", FileLoadError)
    
    Call ClearLineScreen.SetWaterLevelinPipe(WLStartAngle, WLFinishAngle) 'set the ignore profile point and pass to c++
    WaterLevelIgnoreCenter = True 'PCNLS190603
    Call ClearLineScreen.SetUnsetWaterLevel(WaterLevelIgnoreCenter, False) 'PCN3219 WaterLevelIgnoreProfile) 'PCNLS190603waterlevel

    
    'There is no function to blindly turn on overlay tuning, there is a toggle thou
    Call ClearLineScreen.RefreshVideoScreen 'PVD calculation is cancled need refresh
                                            'to show previous water level setting
    CLPScreenDrawAction = ""
    CLPScreenDrawState = ""
    CLPScreenAction = ""
    Call ClearLineScreen.SetupMouseIcon(116)
    
ElseIf CLPScreenDrawAction = "LeftClick" And CLPScreenDrawState = "" Then
         
    For I = 1 To 4
        With WaterLevelLine(I)
            .StartX = 0
            .StartY = 0
            .EndX = 0
            .EndY = 0
            .Colour = 0
        End With
    Next I

        
    WaterLevelLine(1).Colour = &HCCD2FF
    WaterLevelLine(2).Colour = &HCCD2FF
    WaterLevelLine(3).Colour = &HCCD2FF
    WaterLevelLine(4).Colour = &HCCD2FF
    
    screenctrl.AutoRedraw = True
    Call ScreenDrawing.DrawProfilesStartToFinish(screenctrl, True)
    screenctrl.AutoRedraw = False
    CLPScreenDrawState = "RubberBand"


ElseIf CLPScreenDrawAction = "Move" And CLPScreenDrawState = "RubberBand" Then
    
    CurrentX = ClearLineScreen.CurrentMouseMoveX - DrawingCentreX 'CentreLineX
    CurrentY = ClearLineScreen.CurrentMouseMoveY - DrawingCentreY 'CentreLineY
    
    Call RefShapeDistCalc(CurrentX, CurrentY, OrthoX, OrthoY, LineDistance)
    LineDistance = -Abs(LineDistance)
    DistToCentre = ExpectedRadius + LineDistance
    AngleOfCur = -CalAsin(0, 0, CurrentX - ShapeCentreX, CurrentY - ShapeCentreY) + (PI / 2 + PI)
    If ExpectedRadius <> 0 And LineDistance < 0 Then AngleAtHyp = ArcSin(DistToCentre / ExpectedRadius)
    
    WLStartX = CSng((Cos((PI - AngleAtHyp) + AngleOfCur) * ExpectedRadius) / PVDrawScreenRatio)  ' - DrawingCentreX
    WLStartY = CSng((Sin((PI - AngleAtHyp) + AngleOfCur) * ExpectedRadius) / PVDrawScreenRatio)  ' - DrawingCentreY
    WLEndX = CSng((Cos(AngleAtHyp + AngleOfCur) * ExpectedRadius) / PVDrawScreenRatio) ' - DrawingCentreX
    WLEndY = CSng((Sin(AngleAtHyp + AngleOfCur) * ExpectedRadius) / PVDrawScreenRatio) ' - DrawingCentreY

    WLStartXExt = CSng((Cos((PI - AngleAtHyp) + AngleOfCur) * ExpectedDiameter) / PVDrawScreenRatio)  ' - DrawingCentreX
    WLStartYExt = CSng((Sin((PI - AngleAtHyp) + AngleOfCur) * ExpectedDiameter) / PVDrawScreenRatio)  ' - DrawingCentreY
    WLEndXExt = CSng((Cos(AngleAtHyp + AngleOfCur) * ExpectedDiameter) / PVDrawScreenRatio) ' - DrawingCentreX
    WLEndYExt = CSng((Sin(AngleAtHyp + AngleOfCur) * ExpectedDiameter) / PVDrawScreenRatio) ' - DrawingCentreY
    
    With WaterLevelLine(2)
        .StartX = (WLStartX) + ShapeCentreX '/ PVDrawScreenRatio) - DrawingCentreX
        .StartY = (WLStartY) + ShapeCentreY ' / PVDrawScreenRatio) - DrawingCentreY
        .EndX = (WLEndX) + ShapeCentreX '/ PVDrawScreenRatio) - DrawingCentreX
        .EndY = (WLEndY) + ShapeCentreY '/ PVDrawScreenRatio) - DrawingCentreY
    End With
    
    With WaterLevelLine(3)
        .StartX = 0 + ShapeCentreX 'DrawingCentreX
        .StartY = 0 + ShapeCentreY 'DrawingCentreY
        .EndX = WLStartXExt + ShapeCentreX
        .EndY = WLStartYExt + ShapeCentreY
    End With
    
    With WaterLevelLine(4)
        .StartX = 0 + ShapeCentreX 'DrawingCentreX
        .StartY = 0 + ShapeCentreY 'DrawingCentreY
        .EndX = WLEndXExt + ShapeCentreX
        .EndY = WLEndYExt + ShapeCentreY
    End With
    
    
    With WaterLevelLine(1)
        .StartX = CurrentX
        .StartY = CurrentY
        .EndX = CSng(OrthoX / PVDrawScreenRatio) - DrawingCentreX ' + ShapeCentreX 'CentreLineX -
        .EndY = CSng(OrthoY / PVDrawScreenRatio) - DrawingCentreY  '+ ShapeCentreY  'CentreLineY -
    End With
    
    
    
    screenctrl.Cls
    Call DrawWaterLevel(screenctrl)
    
End If

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & error$
    End Select
End Sub

Sub DrawActionSingleLine(screenctrl As Control)
On Error GoTo Err_Handler

Dim LineDistance As Double
Dim LineNumber As Long

Dim NumberLines As Long
NumberLines = UBound(DrawingSingleLines)

If CLPScreenDrawAction = "LeftClick" And CLPScreenDrawState = "RubberBand" Then
    CLPScreenDrawState = ""
    
    NumberLines = NumberLines + 1
    ReDim Preserve DrawingSingleLines(NumberLines)
    DrawingSingleLines(NumberLines) = DrawingActionLine
    DrawingSingleLines(NumberLines).Colour = vbGreen

    Call DrawProfilesStartToFinish(screenctrl, True)
ElseIf CLPScreenDrawAction = "LeftClick" And CLPScreenDrawState = "" Then

    With DrawingActionLine
        .StartX = ClearLineScreen.CurrentMouseMoveX - DrawingCentreX
        .StartY = ClearLineScreen.CurrentMouseMoveY - DrawingCentreY
        .EndX = ClearLineScreen.CurrentMouseMoveX - DrawingCentreX
        .EndY = ClearLineScreen.CurrentMouseMoveY - DrawingCentreY
        .Colour = vbBlue
    End With
    CLPScreenDrawState = "RubberBand"
    screenctrl.AutoRedraw = True
    Call DrawPVProfile(screenctrl, 1)  'Goto frame PVFrameNo
    screenctrl.AutoRedraw = False

ElseIf CLPScreenDrawAction = "Move" And CLPScreenDrawState = "RubberBand" Then
    With DrawingActionLine
        .EndX = ClearLineScreen.CurrentMouseMoveX - DrawingCentreX
        .EndY = ClearLineScreen.CurrentMouseMoveY - DrawingCentreY

    End With
    
    screenctrl.Cls
    Call DrawShapeLine(screenctrl, DrawingActionLine)
    
    LineDistance = GetLineDistance(DrawingActionLine)
    Call ClearLineScreen.UpdateDimensionRuler(LineDistance)
    
 End If

Exit Sub
Err_Handler:
    Select Case Err
        Case 9: NumberLines = -1: Resume Next
        Case Else: MsgBox Err & error$
    End Select
End Sub

Sub DrawActionMaskBox(screenctrl As Control)
On Error GoTo Err_Handler

Dim BoxNumber As Long


If CLPScreenDrawAction = "LeftClick" And CLPScreenDrawState = "RubberBand" Then
    CLPScreenDrawState = ""
    
    Call AddMask(screenctrl, DrawingActionLine)
    Call DrawProfilesStartToFinish(screenctrl, True)

    
ElseIf CLPScreenDrawAction = "LeftClick" And CLPScreenDrawState = "" Then

    With DrawingActionLine
        .StartX = ClearLineScreen.CurrentMouseMoveX - CentreLineX
        .StartY = ClearLineScreen.CurrentMouseMoveY - CentreLineY
        .EndX = ClearLineScreen.CurrentMouseMoveX - CentreLineX
        .EndY = ClearLineScreen.CurrentMouseMoveY - CentreLineY
        .Colour = vbBlue
    End With
    CLPScreenDrawState = "RubberBand"
    screenctrl.AutoRedraw = True
    Call DrawPVProfile(screenctrl, 1)  'Goto frame PVFrameNo
    screenctrl.AutoRedraw = False

ElseIf CLPScreenDrawAction = "Move" And CLPScreenDrawState = "RubberBand" Then
    With DrawingActionLine
        .EndX = ClearLineScreen.CurrentMouseMoveX - CentreLineX
        .EndY = ClearLineScreen.CurrentMouseMoveY - CentreLineY

    End With
    
    screenctrl.Cls
    'Call DrawShapeLine(screenctrl, DrawingActionLine)
    Call DrawShapeBox(screenctrl, DrawingActionLine, , True)
 End If

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & error$
    End Select
End Sub
Sub DrawActionIntegratedArea(screenctrl As Control)
On Error GoTo Err_Handler


Dim LineNumber As Long
Dim IntegratedArea As Double
Dim IntegratedAreaPercentage As Double

Dim NumberLines As Long
NumberLines = UBound(DrawingIntegratedArea)

If CLPScreenDrawAction = "RightClick" And CLPScreenDrawState = "RubberBand" Then
    CLPScreenDrawState = ""
    IntegratedArea = GetIntegratedArea
    IntegratedAreaPercentage = IntegratedArea / (PI * ((ExpectedDiameter / 2) * (ExpectedDiameter / 2))) * 100
    Call ClearLineScreen.UpdateIntegratedArea(IntegratedAreaPercentage)
 
'    NumberLines = NumberLines + 1
'    ReDim Preserve DrawingIntegratedArea(NumberLines)
'    DrawingIntegratedArea(NumberLines) = DrawingActionLine
'    DrawingIntegratedArea(NumberLines).colour = vbGreen
'
    Call DrawProfilesStartToFinish(screenctrl, True)
ElseIf CLPScreenDrawAction = "LeftClick" And CLPScreenDrawState = "RubberBand" Then
    NumberLines = NumberLines + 1
    ReDim Preserve DrawingIntegratedArea(NumberLines)
    DrawingIntegratedArea(NumberLines) = DrawingActionLine
    DrawingIntegratedArea(NumberLines).Colour = &HFFD200
    
    With DrawingActionLine
        .StartX = ClearLineScreen.CurrentMouseMoveX - DrawingCentreX
        .StartY = ClearLineScreen.CurrentMouseMoveY - DrawingCentreY
        .EndX = ClearLineScreen.CurrentMouseMoveX - DrawingCentreX
        .EndY = ClearLineScreen.CurrentMouseMoveY - DrawingCentreY
        .Colour = vbBlue
    End With
    CLPScreenDrawState = "RubberBand"
    screenctrl.AutoRedraw = True
    Call DrawProfilesStartToFinish(screenctrl, True)
    screenctrl.AutoRedraw = False
    
    IntegratedArea = GetIntegratedArea
    IntegratedAreaPercentage = IntegratedArea / (PI * ((ExpectedDiameter / 2) * (ExpectedDiameter / 2))) * 100
    Call ClearLineScreen.UpdateIntegratedArea(IntegratedAreaPercentage)
    
ElseIf CLPScreenDrawAction = "LeftClick" And CLPScreenDrawState = "" Then

    With DrawingActionLine
        .StartX = ClearLineScreen.CurrentMouseMoveX - DrawingCentreX
        .StartY = ClearLineScreen.CurrentMouseMoveY - DrawingCentreY
        .EndX = ClearLineScreen.CurrentMouseMoveX - DrawingCentreX
        .EndY = ClearLineScreen.CurrentMouseMoveY - DrawingCentreY
        .Colour = vbBlue
    End With
    CLPScreenDrawState = "RubberBand"
    screenctrl.AutoRedraw = True
    Call DrawProfilesStartToFinish(screenctrl, True)
    screenctrl.AutoRedraw = False

ElseIf CLPScreenDrawAction = "Move" And CLPScreenDrawState = "RubberBand" Then
    With DrawingActionLine
        .EndX = ClearLineScreen.CurrentMouseMoveX - DrawingCentreX
        .EndY = ClearLineScreen.CurrentMouseMoveY - DrawingCentreY
    End With
    screenctrl.Cls
    Call DrawIntegratedLine(screenctrl, DrawingActionLine, &HFFD200)
    IntegratedArea = GetSingleLineIntegratedArea(DrawingActionLine)
    IntegratedArea = IntegratedArea + GetIntegratedArea
    IntegratedAreaPercentage = IntegratedArea / (PI * ((ExpectedDiameter / 2) * (ExpectedDiameter / 2))) * 100
    Call ClearLineScreen.UpdateIntegratedArea(IntegratedAreaPercentage)
    
 End If

Exit Sub
Err_Handler:
    Select Case Err
        Case 9: NumberLines = -1: Resume Next
        Case Else: MsgBox Err & error$
    End Select
End Sub





Sub DrawActionPolyLine(screenctrl As Control)
On Error GoTo Err_Handler
Dim LineDistance As Double
Dim LineNumber As Long

Dim NumberPolyLines As Long
NumberPolyLines = UBound(DrawingPolyLines)

If CLPScreenDrawAction = "LeftClick" And CLPScreenDrawState = "RubberBand" Then
   
    CLPScreenDrawState = "RubberBand"
    PolyLineDistance = PolyLineDistance + GetLineDistance(DrawingActionLine)
    With DrawingPolyLines(NumberPolyLines)
        .NoLines = .NoLines + 1
        .Lines(.NoLines) = DrawingActionLine
        .Lines(.NoLines).Colour = vbGreen
    End With
    With DrawingActionLine
        .StartX = .EndX
        .StartY = .EndY
    End With
    
    screenctrl.AutoRedraw = True
    Call DrawProfilesStartToFinish(screenctrl, True)
    screenctrl.AutoRedraw = False

ElseIf CLPScreenDrawAction = "RightClick" And CLPScreenDrawState = "RubberBand" Then
    
    CLPScreenDrawState = ""
        Call DrawProfilesStartToFinish(screenctrl, True)
ElseIf CLPScreenDrawAction = "LeftClick" And CLPScreenDrawState = "" Then
    If NumberPolyLines < 128 Then
        NumberPolyLines = NumberPolyLines + 1
        ReDim Preserve DrawingPolyLines(NumberPolyLines)
        DrawingPolyLines(NumberPolyLines).NoLines = 0
        With DrawingActionLine
            .StartX = ClearLineScreen.CurrentMouseMoveX - DrawingCentreX
            .StartY = ClearLineScreen.CurrentMouseMoveY - DrawingCentreY
            .EndX = ClearLineScreen.CurrentMouseMoveX - DrawingCentreX
            .EndY = ClearLineScreen.CurrentMouseMoveY - DrawingCentreY
            .Colour = vbBlue
        End With
        CLPScreenDrawState = "RubberBand"
        screenctrl.AutoRedraw = True
        Call DrawProfilesStartToFinish(screenctrl, True)  'Goto frame PVFrameNo
        screenctrl.AutoRedraw = False
    End If
ElseIf CLPScreenDrawAction = "Move" And CLPScreenDrawState = "RubberBand" Then
    With DrawingActionLine
        .EndX = ClearLineScreen.CurrentMouseMoveX - DrawingCentreX
        .EndY = ClearLineScreen.CurrentMouseMoveY - DrawingCentreY

    End With
    
    screenctrl.Cls
    Call DrawShapeLine(screenctrl, DrawingActionLine)
    
    LineDistance = PolyLineDistance + GetLineDistance(DrawingActionLine)
    Call ClearLineScreen.UpdateDimensionRuler(LineDistance)
    
End If

Exit Sub
Err_Handler:
    Select Case Err
        Case 9: NumberPolyLines = -1: Resume Next
        Case Else: MsgBox Err & error$
    End Select
End Sub

Sub MoveSingleLines(ByVal AmountX As Single, ByVal AmountY As Single, Optional LineNumber)
On Error GoTo Err_Handler

Dim NumberLines As Long
Dim I As Long
NumberLines = UBound(DrawingSingleLines)

If IsMissing(LineNumber) Then
    For I = 0 To NumberLines
        Call MoveShapeLine(DrawingSingleLines(I), AmountX, AmountY)
    Next I
Else
    If (LineNumber >= 0 And LineNumber <= NumberLines) Then
        Call MoveShapeLine(DrawingSingleLines(LineNumber), AmountX, AmountY)
    End If
End If

Exit Sub
Err_Handler:
    Select Case Err
        Case 9: NumberLines = -1: Resume Next
        Case Else: MsgBox Err & error$
    End Select
End Sub


Sub MoveIntegratedArea(ByVal AmountX As Single, ByVal AmountY As Single)
On Error GoTo Err_Handler

Dim NumberLines As Long
Dim I As Long
NumberLines = UBound(DrawingIntegratedArea)

For I = 0 To NumberLines
    Call MoveShapeLine(DrawingIntegratedArea(I), AmountX, AmountY)
Next I

Exit Sub
Err_Handler:
    Select Case Err
        Case 9: NumberLines = -1: Resume Next
        Case Else: MsgBox Err & error$
    End Select
End Sub

Sub MovePolyLines(ByVal AmountX As Single, ByVal AmountY As Single, Optional PolyLineNumber)
On Error GoTo Err_Handler

Dim NumberPolyLines
Dim I As Long
Dim j As Long

NumberPolyLines = UBound(DrawingPolyLines)

If IsMissing(PolyLineNumber) Then
    For j = 0 To NumberPolyLines
        For I = 1 To DrawingPolyLines(j).NoLines
            Call MoveShapeLine(DrawingPolyLines(j).Lines(I), AmountX, AmountY)
        Next I
    Next j
Else
    If (PolyLineNumber >= 0 And PolyLineNumber <= NumberPolyLines) Then
        For I = 1 To DrawingPolyLines(PolyLineNumber).NoLines
            Call MoveShapeLine(DrawingPolyLines(PolyLineNumber).Lines(I), AmountX, AmountY)
        Next I
    End If
End If
    
Exit Sub
Err_Handler:
    Select Case Err
        Case 9: NumberPolyLines = -1: Resume Next
        Case Else: MsgBox Err & error$
    End Select
End Sub



Sub MoveShapeLine(TheLine As ShapeLine_V10, ByVal AmountX As Single, ByVal AmountY As Single)
On Error GoTo Err_Handler
    With TheLine
        .EndX = .EndX + AmountX
        .EndY = .EndY + AmountY
        .StartX = .StartX + AmountX
        .StartY = .StartY + AmountY
    End With
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & error$
    End Select
End Sub

Sub DrawSingleLines(screenctrl As Control, Optional LineNumber, Optional Colour)
On Error GoTo Err_Handler
Dim I As Long
Dim AX As Single
Dim AY As Single
Dim BX As Single
Dim BY As Single
Dim NumberLines As Long


NumberLines = UBound(DrawingSingleLines)

If IsMissing(LineNumber) Then
    For I = 0 To NumberLines
        Call DrawShapeLine(screenctrl, DrawingSingleLines(I), Colour)
    Next I
Else
    If (LineNumber >= 0 And LineNumber <= NumberLines) Then
        Call DrawShapeLine(screenctrl, DrawingSingleLines(LineNumber), Colour)
    End If
End If
    
Exit Sub
Err_Handler:
    Select Case Err
        Case 9: Exit Sub
        Case Else: MsgBox Err & error$
    End Select
End Sub

Sub DrawIntegratedArea(screenctrl As Control, Optional Colour)
On Error GoTo Err_Handler


Dim I As Long
Dim AX As Single
Dim AY As Single
Dim BX As Single
Dim BY As Single
Dim NumberLines As Long

NumberLines = UBound(DrawingIntegratedArea)

For I = 1 To NumberLines
    Call DrawIntegratedLine(screenctrl, DrawingIntegratedArea(I), Colour)
Next I

Exit Sub
Err_Handler:
    Select Case Err
        Case 9: Exit Sub
        Case Else: MsgBox Err & error$
    End Select
End Sub

Sub DrawIntegratedLine(screenctrl As Control, TheLine As ShapeLine_V10, Optional Colour)
On Error GoTo Err_Handler

Dim OrthoX As Double
Dim OrthoY As Double
Dim CurrentX As Single
Dim CurrentY As Single
Dim LineOne As ShapeLine_V10
Dim LineTwo As ShapeLine_V10
Dim LineDistance As Double

LineOne.StartX = TheLine.StartX: LineTwo.StartX = TheLine.EndX
LineOne.StartY = TheLine.StartY: LineTwo.StartY = TheLine.EndY

Call RefShapeDistCalc(LineOne.StartX, _
                      LineOne.StartY, _
                      OrthoX, OrthoY, LineDistance)
LineOne.EndX = (OrthoX / PVDrawScreenRatio) - DrawingCentreX
LineOne.EndY = (OrthoY / PVDrawScreenRatio) - DrawingCentreY
LineOne.Colour = TheLine.Colour

        
Call RefShapeDistCalc(LineTwo.StartX, _
                      LineTwo.StartY, _
                      OrthoX, OrthoY, LineDistance)
LineTwo.EndX = (OrthoX / PVDrawScreenRatio) - DrawingCentreX
LineTwo.EndY = (OrthoY / PVDrawScreenRatio) - DrawingCentreY
LineTwo.Colour = TheLine.Colour

Call DrawShapeLine(screenctrl, LineOne, Colour)
Call DrawShapeLine(screenctrl, LineTwo, Colour)
Call DrawShapeLine(screenctrl, TheLine, Colour)


Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & error$
    End Select
End Sub

Sub DrawMaskBoxes(screenctrl As Control, Optional BoxNumber, Optional Colour, Optional Fixed)
On Error GoTo Err_Handler
Dim I As Long
Dim AX As Single
Dim AY As Single
Dim BX As Single
Dim BY As Single
Dim NumberBoxes As Long

Dim Box As ShapeLine_V10

NumberBoxes = UBound(DrawingMaskBox)



If IsMissing(BoxNumber) Then
    For I = 0 To NumberBoxes
        Box = DrawingMaskBox(I)
        Box.StartX = Box.StartX / 100 * screenctrl.width - CentreLineX
        Box.StartY = (Box.StartY / 100 * screenctrl.height - CentreLineY) * -1
        Box.EndX = Box.EndX / 100 * screenctrl.width - CentreLineX
        Box.EndY = (Box.EndY / 100 * screenctrl.height - CentreLineY) * -1
        
        Call DrawShapeBox(screenctrl, Box, Colour, Fixed)
    Next I
Else
    If (BoxNumber >= 0 And BoxNumber <= NumberBoxes) Then
        Box = DrawingMaskBox(BoxNumber)
        Box.StartX = Box.StartX / 100 * screenctrl.width - CentreLineX
        Box.StartY = (Box.StartY / 100 * screenctrl.height - CentreLineY) * -1
        Box.EndX = Box.EndX / 100 * screenctrl.width - CentreLineX
        Box.EndY = (Box.EndY / 100 * screenctrl.height - CentreLineY) * -1
        
        Call DrawShapeBox(screenctrl, Box, Colour, Fixed)
    End If
End If
    

    
Exit Sub
Err_Handler:
    Select Case Err
        Case 9: Exit Sub
        Case Else: MsgBox Err & error$
    End Select
End Sub

Sub DrawPolyLines(screenctrl As Control)
On Error GoTo Err_Handler
Dim I As Long
Dim j As Long
Dim AX As Single
Dim AY As Single
Dim BX As Single
Dim BY As Single
Dim NumberPolyLines As Long
Dim Colour As Long

NumberPolyLines = UBound(DrawingPolyLines)

For j = 0 To NumberPolyLines
    With DrawingPolyLines(j)
        For I = 1 To .NoLines
            Call DrawShapeLine(screenctrl, .Lines(I))
        Next I
    End With
Next j
  
    

Exit Sub
Err_Handler:
    Select Case Err
        Case 9: Exit Sub
        Case Else: MsgBox Err & error$
    End Select
End Sub

Sub DrawActionCalibrationLine(screenctrl As Control)
On Error GoTo Err_Handler

Dim LineDistance As Double

    If CLPScreenDrawAction = "LeftClick" And CLPScreenDrawState = "RubberBand" Then
        CLPScreenDrawState = ""
        LineDistance = GetLineDistance(CalibrationLine) / PVDrawScreenRatio
        Call DrawProfilesStartToFinish(screenctrl, True)
        CalLength_Global = LineDistance
        Load DataEntryForm
        Call DataEntryForm.SetDataEntryType("CalibrationChange", CalLengthYScale_Global)
    ElseIf CLPScreenDrawAction = "LeftClick" And CLPScreenDrawState = "" Then
        
        With CalibrationLine
            .StartX = ClearLineScreen.CurrentMouseMoveX - DrawingCentreX
            .StartY = ClearLineScreen.CurrentMouseMoveY - DrawingCentreY
            .EndX = ClearLineScreen.CurrentMouseMoveX - DrawingCentreX
            .EndY = ClearLineScreen.CurrentMouseMoveY - DrawingCentreY
            .Colour = vbYellow
        End With
        CLPScreenDrawState = "RubberBand"
        screenctrl.AutoRedraw = True
        Call DrawPVProfile(screenctrl, 1)  'Goto frame PVFrameNo
        screenctrl.AutoRedraw = False
        

    ElseIf CLPScreenDrawAction = "Move" And CLPScreenDrawState = "RubberBand" Then
        With CalibrationLine
            .EndX = ClearLineScreen.CurrentMouseMoveX - DrawingCentreX
            .EndY = ClearLineScreen.CurrentMouseMoveY - DrawingCentreY
        End With
        
        screenctrl.Cls
        Call DrawShapeLine(screenctrl, CalibrationLine)
        
        LineDistance = GetLineDistance(CalibrationLine)
        Call ClearLineScreen.UpdateDimensionRuler(LineDistance)
        
    End If
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & error$
    End Select
End Sub

Sub DrawActionHorCalibrationLine(screenctrl As Control)
On Error GoTo Err_Handler

Dim LineDistance As Double

    If CLPScreenDrawAction = "LeftClick" And CLPScreenDrawState = "RubberBand" Then
        CLPScreenDrawState = ""
        LineDistance = GetLineDistance(DrawingActionLine)
        Call DrawProfilesStartToFinish(screenctrl, True)
        CalLengthYScale_Global = LineDistance
        Load DataEntryForm
        Call DataEntryForm.SetDataEntryType("CalibrationChange", CalLengthYScale_Global)
    ElseIf CLPScreenDrawAction = "LeftClick" And CLPScreenDrawState = "" Then
        
        With DrawingActionLine
            .StartX = ClearLineScreen.CurrentMouseMoveX - DrawingCentreX
            .StartY = ClearLineScreen.CurrentMouseMoveY - DrawingCentreY
            .EndX = ClearLineScreen.CurrentMouseMoveX - DrawingCentreX
            .EndY = ClearLineScreen.CurrentMouseMoveY - DrawingCentreY
            .Colour = vbYellow
        End With
        CLPScreenDrawState = "RubberBand"
        screenctrl.AutoRedraw = True
        Call DrawPVProfile(screenctrl, 1)  'Goto frame PVFrameNo
        screenctrl.AutoRedraw = False
        

    ElseIf CLPScreenDrawAction = "Move" And CLPScreenDrawState = "RubberBand" Then
        With DrawingActionLine
            .EndX = ClearLineScreen.CurrentMouseMoveX - DrawingCentreX
            .EndY = ClearLineScreen.CurrentMouseMoveY - DrawingCentreY
        End With
        
        screenctrl.Cls
        Call DrawShapeLine(screenctrl, DrawingActionLine)
        
        LineDistance = GetLineDistance(DrawingActionLine)
        Call ClearLineScreen.UpdateDimensionRuler(LineDistance)
        
    End If
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & error$
    End Select
End Sub

Sub DrawActionDimension(screenctrl As Control)
On Error GoTo Err_Handler

Dim LineDistance As Double

    If CLPScreenDrawAction = "LeftClick" And CLPScreenDrawState = "RubberBand" Then
        CLPScreenDrawState = ""
        LineDistance = GetLineDistance(DrawingActionLine)
        Call DrawProfilesStartToFinish(screenctrl, True)
    ElseIf CLPScreenDrawAction = "LeftClick" And CLPScreenDrawState = "" Then
        
        With DrawingActionLine
            .StartX = ClearLineScreen.CurrentMouseMoveX - DrawingCentreX
            .StartY = ClearLineScreen.CurrentMouseMoveY - DrawingCentreY
            .EndX = ClearLineScreen.CurrentMouseMoveX - DrawingCentreX
            .EndY = ClearLineScreen.CurrentMouseMoveY - DrawingCentreY
            .Colour = vbYellow
        End With
        CLPScreenDrawState = "RubberBand"
        screenctrl.AutoRedraw = True
        Call DrawPVProfile(screenctrl, 1)  'Goto frame PVFrameNo
        screenctrl.AutoRedraw = False
        

    ElseIf CLPScreenDrawAction = "Move" And CLPScreenDrawState = "RubberBand" Then
        With DrawingActionLine
            .EndX = ClearLineScreen.CurrentMouseMoveX - DrawingCentreX
            .EndY = ClearLineScreen.CurrentMouseMoveY - DrawingCentreY
        End With
        
        screenctrl.Cls
        Call DrawShapeLine(screenctrl, DrawingActionLine)
        
        LineDistance = GetLineDistance(DrawingActionLine)
        Call ClearLineScreen.UpdateDimensionRuler(LineDistance)
        
    End If
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & error$
    End Select
End Sub

Sub DrawActionOrthogonalLine(screenctrl As Control)
On Error GoTo Err_Handler

Dim LineDistance As Double
Dim OrthoX As Double
Dim OrthoY As Double
Dim CurrentX As Single
Dim CurrentY As Single

    If CLPScreenDrawAction = "LeftUp" And CLPScreenDrawState = "RubberBand" Then
        CLPScreenDrawState = ""
    ElseIf CLPScreenDrawAction = "LeftClick" And CLPScreenDrawState = "" Then
        CurrentX = ClearLineScreen.CurrentMouseMoveX - DrawingCentreX 'CentreLineX
        CurrentY = ClearLineScreen.CurrentMouseMoveY - DrawingCentreY 'CentreLineY
        
        Call RefShapeDistCalc(CurrentX, CurrentY, OrthoX, OrthoY, LineDistance)
         With OrthogonalLine
            .StartX = CurrentX
            .StartY = CurrentY
            .EndX = CSng(OrthoX / PVDrawScreenRatio) - DrawingCentreX ' + ShapeCentreX 'CentreLineX -
            .EndY = CSng(OrthoY / PVDrawScreenRatio) - DrawingCentreY '+ ShapeCentreY  'CentreLineY -
            .Colour = &HFFD200
        End With
        CLPScreenDrawState = "RubberBand"
        screenctrl.AutoRedraw = True
        Call DrawPVProfile(screenctrl, 1)  'Goto frame PVFrameNo
        screenctrl.AutoRedraw = False
        Call DrawShapeLine(screenctrl, OrthogonalLine)
        Call ClearLineScreen.UpdateDimensionRuler(LineDistance)
        
        

    ElseIf CLPScreenDrawAction = "Move" And CLPScreenDrawState = "RubberBand" Then
        
        CurrentX = ClearLineScreen.CurrentMouseMoveX - DrawingCentreX 'CentreLineX
        CurrentY = ClearLineScreen.CurrentMouseMoveY - DrawingCentreY 'CentreLineY
        
        Call RefShapeDistCalc(CurrentX, CurrentY, OrthoX, OrthoY, LineDistance)

        With OrthogonalLine
            .StartX = CurrentX
            .StartY = CurrentY
            .EndX = CSng(OrthoX / PVDrawScreenRatio) - DrawingCentreX ' + ShapeCentreX 'CentreLineX -
            .EndY = CSng(OrthoY / PVDrawScreenRatio) - DrawingCentreY '+ ShapeCentreY  'CentreLineY -
        End With
        
        screenctrl.Cls
        Call DrawShapeLine(screenctrl, OrthogonalLine)
        Call ClearLineScreen.UpdateDimensionRuler(LineDistance)
        
    End If

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & error$
    End Select
End Sub

Sub RefShapeDistCalc(ByVal CurrentX, ByVal CurrentY, ByRef OrthoX As Double, ByRef OrthoY As Double, ByRef LineDistance As Double)
On Error GoTo Err_Handler

        Call clearline_RefShapeDistCalc(ReferenceShape(GetNumShapeType(DrawShapeType)), _
                                        (CurrentX + DrawingCentreX) * PVDrawScreenRatio, _
                                        (CurrentY + DrawingCentreY) * PVDrawScreenRatio, _
                                        OrthoX, _
                                        OrthoY, _
                                        LineDistance, _
                                        ExpectedDiameter / 2, _
                                        (DrawingCentreX + ShapeCentreX) * PVDrawScreenRatio, _
                                        (DrawingCentreY + ShapeCentreY) * PVDrawScreenRatio, _
                                        ShapeRotationAngle)
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & error$
    End Select
End Sub

Sub DrawActionMoveAll(screenctrl As Control)
On Error GoTo Err_Handler

Dim AmountMovedX As Single
Dim AmountMovedY As Single
Dim LineDistance As Double

Dim OrthoX As Double
Dim OrthoY As Double
Dim CurrentX As Single
Dim CurrentY As Single

If CLPScreenDrawAction = "Move" And CLPScreenDrawState = "" And ShowReferenceShape Then
        CurrentX = ClearLineScreen.CurrentMouseMoveX - DrawingCentreX
        CurrentY = ClearLineScreen.CurrentMouseMoveY - DrawingCentreY
        Call RefShapeDistCalc(CurrentX, CurrentY, OrthoX, OrthoY, LineDistance)
        LineDistance = LineDistance / PVDrawScreenRatio
        
        If LineDistance < 4 And LineDistance > -4 Then
           Call ClearLineScreen.SetupMouseIcon(122)
           CLPScreenItemSelect = "ReferenceShape"
        Else
           Call ClearLineScreen.SetupMouseIcon(108)
           CLPScreenItemSelect = ""
        End If
ElseIf CLPScreenDrawAction = "Move" And CLPScreenDrawState = "Moving" Then

        AmountMovedX = ClearLineScreen.CurrentMouseMoveX - ClearLineScreen.LastMouseMoveX
        AmountMovedY = ClearLineScreen.CurrentMouseMoveY - ClearLineScreen.LastMouseMoveY
        
        If CLPScreenItemSelect <> "ReferenceShape" Then
            Call MoveSingleLines(AmountMovedX, AmountMovedY)
            Call MoveIntegratedArea(AmountMovedX, AmountMovedY)
            Call MovePolyLines(AmountMovedX, AmountMovedY)
            Call ClearLineScreen.OnscreenTextMoveAll(AmountMovedX, AmountMovedY)
            
            
        End If
        If ShowReferenceShape Then
            Call MoveReferenceShape(AmountMovedX, AmountMovedY)
            Call MoveWaterLevel(AmountMovedX, AmountMovedY)
        End If
        
        
        screenctrl.Cls
        Call DrawSingleLines(screenctrl)
        Call DrawPolyLines(screenctrl)
        Call DrawAllReferenceShapes(screenctrl)
        Call DrawWaterLevel(screenctrl)
        Call DrawIntegratedArea(screenctrl)
        
ElseIf CLPScreenDrawAction = "LeftClick" And CLPScreenDrawState = "" Then
    screenctrl.AutoRedraw = True
        Call DrawPVProfile(screenctrl, 1)  'Goto frame PVFrameNo
    screenctrl.AutoRedraw = False
    Call ClearLineScreen.SetupMouseIcon(109)
    CLPScreenDrawState = "Moving"
ElseIf CLPScreenDrawAction = "LeftUp" And CLPScreenDrawState = "Moving" Then
    Call DrawProfilesStartToFinish(screenctrl, True)
    If ShowReferenceShape And CLPScreenMode = PV Then Call ScreenDrawing.PVFlat3DCalcCPP(0, PVDataNoOfLines)
    CLPScreenDrawState = ""
    Call ClearLineScreen.SetupMouseIcon(108)
    If CLPScreenActionPrevious = "DrawWaterLevel" Then Call ScreenDrawing.SetupDrawingRoutine("DrawWaterLevel")
End If
       
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & error$
    End Select
End Sub

Sub DrawActionDrawText(screenctrl As Control)
On Error GoTo Err_Handler

Dim NumberOfTextBoxes As Integer
Dim CurrentX As Single
Dim CurrentY As Single



If CLPScreenDrawAction = "LeftClick" And CLPScreenDrawState = "" Then

    CurrentX = ClearLineScreen.CurrentMouseMoveX '- DrawingCentreX 'CentreLineX
    CurrentY = ClearLineScreen.CurrentMouseMoveY '- DrawingCentreY 'CentreLineY
    With ClearLineScreen
        NumberOfTextBoxes = .OnScreenText.Count
        Load .OnScreenText(NumberOfTextBoxes)
        
        .OnScreenText(NumberOfTextBoxes).Left = CurrentX
        .OnScreenText(NumberOfTextBoxes).Top = CurrentY
        
        .OnScreenText(NumberOfTextBoxes).Visible = True
        .OnScreenText(NumberOfTextBoxes).ZOrder 0
        .OnScreenText(NumberOfTextBoxes).SetFocus
        Call .SetTextBoxWidthAndHeight(.OnScreenText(NumberOfTextBoxes))
    
        '.ConText = 0
    End With
'    TextStartX = X
'    TextStartY = Y
End If

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & error$
    End Select
End Sub

Sub MoveReferenceShape(AmountMovedX, AmountMovedY)
On Error GoTo Err_Handler
Dim CurrentTime As Double

    CurrentTime = ClearLineScreen.GetCurrentAVITime
    
    ShapeCentreX = ShapeCentreX + AmountMovedX: ShapeCentreY = ShapeCentreY + AmountMovedY
    If CLPScreenMode = PV Then Call PVFlat3DCalcCPP(GraphStartFrame, GraphEndFrame): Exit Sub
    If CLPScreenMode <> PV And GetPVDVer < 3 Then Exit Sub
    If mediatype = Video And (CurrentTime < PVTimes(1) Or _
                          CurrentTime > PVTimes(PVDataNoOfLines)) Then Exit Sub
    
    Call PVFlat3DCalcCPP(GraphStartFrame, GraphEndFrame)
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & error$
    End Select
End Sub

Sub DrawShapeLine(screenctrl As Control, TheLine As ShapeLine_V10, Optional PassedColour)
On Error GoTo Err_Handler

Dim AX As Single
Dim AY As Single
Dim BX As Single
Dim BY As Single
Dim Colour As Long


With TheLine
    AX = .StartX + DrawingCentreX
    AY = .StartY + DrawingCentreY
    BX = .EndX + DrawingCentreX
    BY = .EndY + DrawingCentreY
    If IsMissing(PassedColour) Then
        Colour = .Colour
    Else
        Colour = CLng(PassedColour)
    End If
    
End With
Call RenderLine(screenctrl, AX, AY, BX, BY, Colour)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & error$
    End Select

End Sub

Sub DrawShapeBox(screenctrl As Control, TheBox As ShapeLine_V10, Optional PassedColour, Optional Fixed)
On Error GoTo Err_Handler

Dim AX As Single
Dim AY As Single
Dim BX As Single
Dim BY As Single
Dim Colour As Long


With TheBox
    If IsMissing(Fixed) Then
        AX = .StartX + DrawingCentreX
        AY = .StartY + DrawingCentreY
        BX = .EndX + DrawingCentreX
        BY = .EndY + DrawingCentreY
        
    Else
        AX = .StartX + CentreLineX
        AY = .StartY + CentreLineY
        BX = .EndX + CentreLineX
        BY = .EndY + CentreLineY
    End If
    
    If IsMissing(PassedColour) Then
        Colour = .Colour
    Else
        Colour = CLng(PassedColour)
    End If
    
End With
'Call RenderLine(screenctrl, AX, AY, BX, BY, Colour)
Call RenderBox(screenctrl, AX, AY, BX, BY, Colour)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & error$
    End Select

End Sub

Function GetLineDistance(TheLine As ShapeLine_V10)
On Error GoTo Err_Handler
    Dim answer As Double
    Dim XLine As Double
    Dim YLine As Double
    
    XLine = Abs(TheLine.EndX - TheLine.StartX)
    YLine = Abs(TheLine.EndY - TheLine.StartY)
    answer = Sqr((XLine * XLine) + (YLine * YLine))
    GetLineDistance = answer * PVDrawScreenRatio
Exit Function
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & error$
    End Select
End Function

Function GetIntegratedArea() As Double
On Error GoTo Err_Handler

    Dim NumberLines As Integer
    Dim TotalArea As Double
    Dim I As Integer
    
    NumberLines = UBound(DrawingIntegratedArea)
    If NumberLines < 1 Then GetIntegratedArea = 0: Exit Function
    
    TotalArea = 0
    For I = 1 To NumberLines
        TotalArea = TotalArea + GetSingleLineIntegratedArea(DrawingIntegratedArea(I))
    Next I
    GetIntegratedArea = TotalArea
    
Exit Function
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & error$
        GetIntegratedArea = 0
    End Select
End Function

Function GetSingleLineIntegratedArea(TheLine As ShapeLine_V10)
On Error GoTo Err_Handler

Dim OrthoX As Double
Dim OrthoY As Double
Dim CurrentX As Single
Dim CurrentY As Single
Dim LineOne As ShapeLine_V10
Dim LineTwo As ShapeLine_V10
Dim LineThree As ShapeLine_V10
Dim LineDistanceOne As Double
Dim LineDistanceTwo As Double
Dim TotalAreaOfCircle As Double
Dim SegmentArea As Double
Dim LineAngleOne As Double
Dim LineAngleTwo As Double
Dim AreaOfTriangleSegment As Double
Dim AngleOfSegment As Double

LineOne.StartX = TheLine.StartX: LineTwo.StartX = TheLine.EndX
LineOne.StartY = TheLine.StartY: LineTwo.StartY = TheLine.EndY

'TotalAreaOfCircle = 2 * PI * (ExpectedDiameter / 2) * (ExpectedDiameter / 2)
TotalAreaOfCircle = PI * (ExpectedDiameter / 2) * (ExpectedDiameter / 2)

If TotalAreaOfCircle = 0 Then GetSingleLineIntegratedArea = 0: Exit Function

Call RefShapeDistCalc(LineOne.StartX, _
                      LineOne.StartY, _
                      OrthoX, OrthoY, LineDistanceOne)
LineOne.EndX = OrthoX - (DrawingCentreX * PVDrawScreenRatio)
LineOne.EndY = OrthoY - (DrawingCentreY * PVDrawScreenRatio)
LineAngleOne = CalAsin(0, 0, LineOne.EndX, LineOne.EndY)

Call RefShapeDistCalc(LineTwo.StartX, _
                      LineTwo.StartY, _
                      OrthoX, OrthoY, LineDistanceTwo)
LineTwo.EndX = OrthoX - (DrawingCentreX * PVDrawScreenRatio)
LineTwo.EndY = OrthoY - (DrawingCentreY * PVDrawScreenRatio)
LineAngleTwo = CalAsin(0, 0, LineTwo.EndX, LineTwo.EndY)

AngleOfSegment = Abs(LineAngleOne - LineAngleTwo)
If AngleOfSegment > PI Then AngleOfSegment = Abs(AngleOfSegment - (2 * PI))



LineThree.StartX = TheLine.StartX * PVDrawScreenRatio
LineThree.StartY = TheLine.StartY * PVDrawScreenRatio
LineThree.EndX = TheLine.EndX * PVDrawScreenRatio
LineThree.EndY = TheLine.EndY * PVDrawScreenRatio

SegmentArea = AngleOfSegment / (2 * PI) * TotalAreaOfCircle

AreaOfTriangleSegment = AreaOf3Points(0, 0, _
                                      LineThree.StartX, LineThree.StartY, _
                                      LineThree.EndX, LineThree.EndY)

If LineDistanceOne <= 0 And LineDistanceTwo <= 0 Then
    GetSingleLineIntegratedArea = SegmentArea - AreaOfTriangleSegment
End If
If LineDistanceOne > 0 And LineDistanceTwo > 0 Then
    GetSingleLineIntegratedArea = AreaOfTriangleSegment - SegmentArea
End If






Exit Function
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & error$
        GetSingleLineIntegratedArea = 0
    End Select
End Function

Sub DeleteAll()
On Error GoTo Err_Handler

    ReDim DrawingSingleLines(0)
    ReDim DrawingPolyLines(0)
    ReDim DrawingIntegratedArea(0)
    Call ClearLineScreen.OnscreenTextDeleteAll
    If ShapeCentreX <> 0 Or ShapeCentreY <> 0 Then
        ShapeCentreX = 0: ShapeCentreY = 0
        Call ScreenDrawing.PVFlat3DCalcCPP(0, PVDataNoOfLines)
    End If
    
'screenctrl.AutoRedraw = True 'PCNGL3001032
'screenctrl.Cls 'PCNGL3001032
'screenctrl.AutoRedraw = False
'DoEvents 'PCNGL170203

    Call DrawProfilesStartToFinish(ClearLineScreen.PVScreen, True)
    Call DrawProfilesStartToFinish(ClearLineScreen.SnapShotScreen, True)
    
    Call ClearLineScreen.UpdateDimensionRuler(0)
    Call ClearLineScreen.UpdateIntegratedArea(0)
    Call ClearLineScreen.SetupMouseIcon(116)
    CLPScreenAction = ""
    


    
    

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & error$
    End Select
End Sub
Sub MoveWaterLevel(AmountMovedX, AmountMovedY)
On Error GoTo Err_Handler

Dim I
For I = 1 To 4

    WaterLevelLine(I).StartX = WaterLevelLine(I).StartX + AmountMovedX
    WaterLevelLine(I).StartY = WaterLevelLine(I).StartY + AmountMovedY
    WaterLevelLine(I).EndX = WaterLevelLine(I).EndX + AmountMovedX
    WaterLevelLine(I).EndY = WaterLevelLine(I).EndY + AmountMovedY
Next I

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & error$
    End Select
End Sub

Sub DeleteWaterLevel()
On Error GoTo Err_Handler

Dim I
For I = 1 To 4
    WaterLevelLine(I).StartX = 0
    WaterLevelLine(I).StartY = 0
    WaterLevelLine(I).EndX = 0
    WaterLevelLine(I).EndY = 0
Next

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & " - " & error
    End Select
End Sub


Sub DrawWaterLevel(screenctrl As Control)
On Error GoTo Err_Handler
Dim I
For I = 1 To 4
    Call DrawShapeLine(screenctrl, WaterLevelLine(I))
Next I

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & error$
    End Select
End Sub

Sub ClearMask()
On Error GoTo Err_Handler

ReDim DrawingMaskBox(0)
Call hough_clearrectanglecoord
Call DrawProfilesStartToFinish(ClearLineScreen.SnapShotScreen, True)



Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & error$
    End Select

End Sub

Sub AddMask(screenctrl As Control, ByRef Box As ShapeLine_V10)
On Error GoTo Err_Handler

    Dim NumberBoxes As Long
    NumberBoxes = UBound(DrawingMaskBox)
    
    NumberBoxes = NumberBoxes + 1
    ReDim Preserve DrawingMaskBox(NumberBoxes)
    
    Box.StartX = (Box.StartX + CentreLineX) / screenctrl.width * 100
    Box.StartY = ((Box.StartY * -1) + CentreLineY) / screenctrl.height * 100
    Box.EndX = (Box.EndX + CentreLineX) / screenctrl.width * 100
    Box.EndY = ((Box.EndY * -1) + CentreLineY) / screenctrl.height * 100
    
    DrawingMaskBox(NumberBoxes) = Box
    DrawingMaskBox(NumberBoxes).Colour = vbRed
    
    Call setrectanglecoord(Box.StartX, Box.StartY, Box.EndX, Box.EndY)

Exit Sub
Err_Handler:
    Select Case Err
        Case 9: NumberBoxes = -1: Resume Next
        Case Else: MsgBox Err & error$
    End Select
End Sub

Sub FlipShape(ByVal Index As Integer)
On Error GoTo Err_Handler

Dim PartNo As Integer
Dim TempAngle As Double

With ReferenceShape(Index)
    .CentreOffsetY = .CentreOffsetY * -1
    For PartNo = 0 To .NoArcs - 1
        .Arcs(PartNo).OriginY = .Arcs(PartNo).OriginY * -1
        TempAngle = .Arcs(PartNo).EndAngle
        .Arcs(PartNo).EndAngle = 360# - .Arcs(PartNo).StartAngle
        .Arcs(PartNo).StartAngle = 360# - TempAngle
    Next PartNo
End With

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & error$
    End Select
    
End Sub

Sub RotatePVDData(ByVal Rot As Double, Optional FrameNo)

Dim XYArrayIndex As Long
Dim I As Long
Dim XData As Double
Dim YData As Double
Dim TheFrame As Double

Rot = Rot / 180 * PI 'Convert to radians

If IsMissing(FrameNo) Then
    TheFrame = PVFrameNo
Else
    If FrameNo < 1 Or FrameNo > PVDataNoOfLines Then Exit Sub
    TheFrame = FrameNo
End If
    
    

For I = 1 To 180
    XYArrayIndex = I + (NoOfProfileSegments * (TheFrame - 1))

    XData = TD_PVDataX(XYArrayIndex)
    YData = TD_PVDataY(XYArrayIndex)

    If XData <> 0 Or YData <> 0 Then
        YData = YData + (TD_PVCentreY(TheFrame)) 'PCN3219 added centre offset
        XData = XData + (TD_PVCentreX(TheFrame))  'PCN3219 added centre offset
    End If
    Call RotateCoordXY(XData, YData, 0, 0, Rot)
    TD_PVDataX(XYArrayIndex) = XData - (TD_PVCentreX(TheFrame))
    TD_PVDataY(XYArrayIndex) = YData - (TD_PVCentreY(TheFrame))
Next I
    Call DrawProfilesStartToFinish(ClearLineScreen.PVScreen, True)
    Call DrawProfilesStartToFinish(ClearLineScreen.SnapShotScreen, True)

End Sub