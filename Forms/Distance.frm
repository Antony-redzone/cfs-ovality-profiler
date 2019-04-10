VERSION 5.00
Begin VB.Form DistanceForm 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Distance Information"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4350
   Icon            =   "Distance.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6075
   ScaleWidth      =   4350
   Begin VB.Frame frmMask 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mask: Distance - Counter - Text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   60
      Width           =   4095
      Begin VB.CommandButton cmdHelp 
         Height          =   615
         Left            =   3240
         Picture         =   "Distance.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Information/Help"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdClearMask 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clear Mask"
         Height          =   375
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdSetMask 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Set Mask"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame frmDistance 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Distance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   4020
      Left            =   120
      TabIndex        =   0
      Top             =   1260
      Width           =   4095
      Begin VB.CommandButton btnCountDirection 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Count Direction"
         Height          =   495
         Left            =   330
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1785
         Width           =   1860
      End
      Begin VB.PictureBox CountPicBox 
         BackColor       =   &H0080FFFF&
         Height          =   870
         Left            =   2430
         ScaleHeight     =   810
         ScaleWidth      =   1155
         TabIndex        =   14
         Top             =   1650
         Width           =   1215
         Begin VB.Image CountUpArrow 
            Height          =   645
            Left            =   105
            Picture         =   "Distance.frx":2054
            Top             =   1000
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.Image CountTractor 
            Height          =   645
            Left            =   735
            Picture         =   "Distance.frx":300A
            Top             =   90
            Width           =   255
         End
         Begin VB.Image CountDownArrow 
            Height          =   645
            Left            =   105
            Picture         =   "Distance.frx":3908
            Top             =   75
            Width           =   450
         End
      End
      Begin VB.TextBox CalculatedDist 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2430
         TabIndex        =   12
         Top             =   2625
         Width           =   1215
      End
      Begin VB.CommandButton RecalculateDistButton 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Recalculate Distance Scale"
         Height          =   495
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3210
         Width           =   2415
      End
      Begin VB.TextBox txtEndDistance 
         Height          =   285
         Left            =   2430
         TabIndex        =   2
         Top             =   3765
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtStartDistance 
         Height          =   285
         Left            =   2430
         TabIndex        =   1
         Top             =   1260
         Width           =   1215
      End
      Begin VB.Label ProcessMethod_lbl 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Process Method"
         Height          =   390
         Left            =   150
         TabIndex        =   18
         Top             =   255
         Width           =   3390
      End
      Begin VB.Image DistanceMethodType1 
         Height          =   480
         Left            =   3450
         Picture         =   "Distance.frx":48BE
         Tag             =   "1"
         Top             =   345
         Width           =   480
      End
      Begin VB.Image DistanceMethodType2 
         Height          =   480
         Left            =   3450
         Picture         =   "Distance.frx":5500
         Tag             =   "0"
         Top             =   705
         Width           =   480
      End
      Begin VB.Label DistanceMethodType1_lbl 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Automatic Counter"
         Height          =   285
         Left            =   510
         TabIndex        =   17
         Top             =   495
         Width           =   3015
      End
      Begin VB.Label DistanceMethodType2_lbl 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Start/Finish Estimate"
         Height          =   285
         Left            =   510
         TabIndex        =   16
         Top             =   795
         Width           =   3015
      End
      Begin VB.Label CalculatedDist_lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Current Distance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   420
         Left            =   150
         TabIndex        =   13
         Top             =   2685
         Width           =   2415
      End
      Begin VB.Label lblUnit2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "m"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   3750
         TabIndex        =   11
         Top             =   2700
         Width           =   255
      End
      Begin VB.Label lblUnit 
         BackColor       =   &H00C0C0C0&
         Caption         =   "m"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   3750
         TabIndex        =   10
         Top             =   1245
         Width           =   255
      End
      Begin VB.Label lblFinishDistance 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Distance at Analysis End"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   8
         Top             =   3765
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lblStartDistance 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Distance at Analysis Start"
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   150
         TabIndex        =   7
         Top             =   1260
         Width           =   2415
      End
   End
End
Attribute VB_Name = "DistanceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Type DistanceChangeNode
    Frame As Long
    Distance As Double
    time As Double
End Type


Private Sub btnCountDirection_Click()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : btnCountDirection_Click
'Created : 5 March 2004, PCN
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

If CountDirection = "Up" Then
    Call SetCounterDirection("Down")
Else
    Call SetCounterDirection("Up")
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-" & error$
End Sub



Private Sub CalculatedDist_Change()
On Error GoTo Err_Handler
    If Not IsNumeric(CalculatedDist.text) Then Exit Sub
    If DistanceMethod = "AutomaticCounter" Then
        Call ClearLineScreen.DistanceCounterSet(CDbl(CalculatedDist.text))
    End If
Exit Sub
Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & "-" & error$
End Select
End Sub

Private Sub CalculatedDist_GotFocus()
    CalculatedDist.Tag = True
End Sub
 
Public Sub DistanceAdd(ByVal Dist As Double, ByVal FrameNo As Long)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : CalculatedDist_Validate
'Created : 27 July 2004, PCN2928
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Store this new setting in the Pipe Observation array.
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
    If NoOfPipeObservations >= PipeObsBuffer Then Exit Sub 'Ensure NoOfPipeObservations does not exceed the PipeObsBuffer
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
    MsgBox Err & "-" & error$
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
    MsgBox Err & "-" & error$
    ChangeCurrentDistance = False
End Function

Private Sub cmdClearMask_Click()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : cmdClearMask_Click
'Created : 23 March 2004, PCN2639
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    : Clear where the distance counter is located on the video screen.
'Usage   :
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

IgnoreDistX1 = 1
IgnoreDistX2 = 1
IgnoreDistY1 = 1
IgnoreDistY2 = 1
Call INI_WriteBack(MyFile, "IgnoreDistX1=", IgnoreDistX1)
Call INI_WriteBack(MyFile, "IgnoreDistY1=", IgnoreDistY1)
Call INI_WriteBack(MyFile, "IgnoreDistX2=", IgnoreDistX2)
Call INI_WriteBack(MyFile, "IgnoreDistY2=", IgnoreDistY2)
'Call ClearLineScreen.SetRectangle(IgnoreDistX1, IgnoreDistY1, IgnoreDistX2, IgnoreDistY2, "Distance") 'PCN2639
Call ClearLineScreen.ClearDistRectangle 'PCN2639
'Draw boundary for Distance Counter area
Call ClearLineScreen.SetupVideoIgnoreBoarder(ClearLineScreen.VideoImageDistIgnoreBoarder, IgnoreDistX1, IgnoreDistY1, IgnoreDistX2, IgnoreDistY2)

Exit Sub
Err_Handler:
    MsgBox Err & "-" & error$
End Sub

Private Sub cmdHelp_Click()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'cmdHelp_Click sub Michelle Lindsay michellelindsay@cbsys.co.nz
'
'Revision history"
'   V0.0    Michelle Lindsay,   09/12/03    Placing on a new form
'
'Description:
'   Adding a button for user to go to the Helpfile position that
'   is applicable for the current form and command.
'
'Purpose:
'   To enable the help to be easily accessible and intuitive.
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Dim hwndHelp As Long
        If HelpFilename <> "" And Dir(ReadOnlyAppPath & "Language\" & HelpFilename) = "" Then 'Check whether that file exists actually.'PCN2167 7/8/03 by Abe
            MsgBox HelpFilename & " " & DisplayMessage("file for the language") & _
                "(" & Language & ") " & _
                DisplayMessage("does not exist. Create this file first. The default language(English) is loaded for Help file."), , "Clear Line Profiler" 'PCN2167 7/8/03 by Abe, PCN2171
        End If
        If HelpFilename <> "" And Dir(ReadOnlyAppPath & "Language\" & HelpFilename) <> "" Then  'PCN2167 7/8/03 by Abe ---------v
            hwndHelp = HtmlHelp(hwnd, ReadOnlyAppPath & "Language\" & HelpFilename, HH_DISPLAY_TOPIC, 1000)
        Else '-------------------------------------------------------------------------------------------^
            'PCN1972 LS 8/7/03
            'hwndHelp = HtmlHelp(hwnd, App.Path & "\HelpFile.chm", HH_DISPLAY_TOPIC, 0)
            hwndHelp = HtmlHelp(hwnd, LocToSave & "HelpFile.chm::/WaterLevel.htm", HH_DISPLAY_TOPIC, 1000)
        End If
Exit Sub

Err_Handler:
    MsgBox Err & "-" & error$
End Sub



Private Sub cmdSetMask_Click()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'DefineDistLocation_Click sub Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    24/04/03
'           Michelle Lindsay,   09/12/03    Placing on a new form
'
'Description:
'   By defining where the distance counter is located on the video screen,
'   the image processing can ignore this part of the video and hence ensures
'   the profile is not distorted.

'Purpose:
'   To flag that the user wants to define the distance counter location.
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Dim RememberFishEyeState As Boolean ' PCN2907 rememers the to turn fisheye back on if needed

RememberFishEyeState = ConfigInfo.FishEyeFlag 'PCN2907 Remember Fisheye state
FisheyeFunctions.FEOFF 'PCN2907 9 September 2004 Antony

IgnoreDistAreaFlag = True
DelFlag = False

'Take a snap shot
Call ClearLineScreen.TakeASnapShot
DoEvents
Call ClearLineScreen.SetupMouseIcon(118) 'Ignore Text Region PCNGL280503-2
ClearLineScreen.Select_Click
'Unload Distance PCN2639
If RememberFishEyeState Then FisheyeFunctions.FEON 'PCN2907 9 September 2004 Antony
Exit Sub
Err_Handler:
    MsgBox Err & "-" & error$
End Sub

Private Sub CountDownArrow_Click()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : CountDownArrow_Click
'Created : 5 March 2004, PCN
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Call SetCounterDirection("Up")

Exit Sub
Err_Handler:
    MsgBox Err & "-" & error$
End Sub

Private Sub CountPicBox_Click()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : CountTractor_Click
'Created : 5 March 2004, PCN
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

If CountDirection = "Up" Then
    Call SetCounterDirection("Down")
Else
    Call SetCounterDirection("Up")
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-" & error$
End Sub

Private Sub CountTractor_Click()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : CountTractor_Click
'Created : 5 March 2004, PCN
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

If CountDirection = "Up" Then
    Call SetCounterDirection("Down")
Else
    Call SetCounterDirection("Up")
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-" & error$
End Sub

Private Sub CountUpArrow_Click()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : CountDownArrow_Click
'Created : 5 March 2004, PCN
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Call SetCounterDirection("Down")

Exit Sub
Err_Handler:
    MsgBox Err & "-" & error$
End Sub

Private Sub DistanceMethodType1_Click()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : DistanceMethodType1_Click
'Created : 15 March 2004, PCN2612
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

If DistanceMethodType1.Tag = 0 Then
    Call SetCheckBoxTick(DistanceMethodType1, True)
    Call SetCheckBoxTick(DistanceMethodType2, False)
    DistanceMethod = "AutomaticCounter"
    Call SetupControlLayout(DistanceMethod)
    Call INI_WriteBack(MyFile, "DistanceMethod=", DistanceMethod)
Else
    Call SetCheckBoxTick(DistanceMethodType1, False)
    Call SetCheckBoxTick(DistanceMethodType2, False)
    DistanceMethod = "None"
    Call SetupControlLayout(DistanceMethod)
    Call INI_WriteBack(MyFile, "DistanceMethod=", DistanceMethod)
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-" & error$
End Sub

Private Sub DistanceMethodType2_Click()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : DistanceMethodType2_Click
'Created : 15 March 2004, PCN2612
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

If DistanceMethodType2.Tag = 0 Then
    Call SetCheckBoxTick(DistanceMethodType2, True)
    Call SetCheckBoxTick(DistanceMethodType1, False)
    DistanceMethod = "StartFinishEstimate"
    Call SetupControlLayout(DistanceMethod)
    Call INI_WriteBack(MyFile, "DistanceMethod=", DistanceMethod)
Else
    Call SetCheckBoxTick(DistanceMethodType1, False)
    Call SetCheckBoxTick(DistanceMethodType2, False)
    DistanceMethod = "None"
    Call SetupControlLayout(DistanceMethod)
    Call INI_WriteBack(MyFile, "DistanceMethod=", DistanceMethod)
End If


Exit Sub
Err_Handler:
    MsgBox Err & "-" & error$
End Sub

Private Sub Form_Deactivate()
    CalculatedDist.Tag = False
End Sub

Private Sub Form_Load()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Form_Load sub Michelle Lindsay michellelindsay@cbsys.co.nz
'
'Revision history"
'   V0.0    Michelle Lindsay,   09/12/03
'           Michelle Lindsay,   11/12/03    Adding units for distance figures
'
'Description:
'   Ensures that the Distance form is displayed in the correct position on the
'   main form.
'
'Purpose:
'   Positions the form on loading.
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
ConvertLanguage Me, Language 'PCN2111 PCN2506

width = PipelineDetails.width
height = 8430
Top = 300
Left = PipelineDetails.Left
Visible = True
ZOrder 0
If MeasurementUnits = "mm" Then 'PCN2463
    lblUnit.Caption = "m"
    lblUnit2.Caption = "m"
Else
    lblUnit.Caption = "ft"
    lblUnit2.Caption = "ft"
        
End If

'vvvv PCN2639 ************************************
If DistanceMethod = "AutomaticCounter" Then
    Call SetCheckBoxTick(DistanceMethodType1, True)
    Call SetCheckBoxTick(DistanceMethodType2, False)
    Call SetupControlLayout(DistanceMethod)
    'If ConfigInfo.DistanceStart >= 0 Or DistanceStart >= 0 Then 'PCN3884
    If ConfigInfo.DistanceStart > InvalidData Or DistanceStart > InvalidData Then
        txtStartDistance = DistanceStart
        Call SetCounterDirection(Trim(ConfigInfo.DistanceDirection))
    End If
    If PVDFileName = "" Or PVDataNoOfLines = 0 Or PVRecording Then
        RecalculateDistButton.Enabled = False
    Else
        RecalculateDistButton.Enabled = True
    End If
ElseIf DistanceMethod = "StartFinishEstimate" Or DistanceMethod = "AvgFrame" Or DistanceMethod = "AvgTime" Then 'PCNANT???? First Or was and And
    Call SetCheckBoxTick(DistanceMethodType1, False)
    Call SetCheckBoxTick(DistanceMethodType2, True)
    Call SetupControlLayout(DistanceMethod)
    'If ConfigInfo.DistanceStart >= 0 Or DistanceStart >= 0 Then 'PCN3884
    If ConfigInfo.DistanceStart > InvalidData Or DistanceStart > -InvalidData Then
        txtStartDistance = ConfigInfo.DistanceStart
    End If
    If ConfigInfo.DistanceFinish > InvalidData Then 'PCN3884 was >=0
        txtEndDistance = ConfigInfo.DistanceFinish
    End If
    If PVDFileName = "" Or PVDataNoOfLines = 0 Or PVRecording Then
        RecalculateDistButton.Enabled = False
    Else
        RecalculateDistButton.Enabled = True
    End If
Else
    Call SetCheckBoxTick(DistanceMethodType1, False)
    Call SetCheckBoxTick(DistanceMethodType2, False)
    Call SetupControlLayout("None")
    RecalculateDistButton.Enabled = False
End If
'^^^^ ********************************************

Call ClearLineScreen.DistanceCounterTextUpdate  'PCN2928
Distance.CalculatedDist.Tag = False

Exit Sub
Err_Handler:
    MsgBox Err & "-" & error$
End Sub


Public Sub RecalculateDistance()
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
        MsgBox Err & "-" & error$
 
End Select
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
        MsgBox Err & "-" & error$
End Select
End Sub





Function SetCounterDirection(Direction As String)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : SetCounterDirection
'Created : 5 March 2004, PCN
'Updated :
'Prg By  : Geoff Logan
'Param   : Direction - set count direction
'Desc    : This function sets the count direction variable and graphs
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

If Direction = "Up" Then
    CountDownArrow.Visible = False
    CountUpArrow.Top = CountDownArrow.Top
    CountUpArrow.Left = CountDownArrow.Left
    CountUpArrow.Visible = True
    CountDirection = "Up"
Else
    CountUpArrow.Visible = False
    CountDownArrow.Visible = True
    CountDirection = "Down"
End If

''Set Distance Counter analyzer
'If Len(txtStartDistance) = 0 Then
'    DistanceStart = -1
'Else
'    DistanceStart = CDbl(txtStartDistance) 'PCN2898 removed the function Val eg was CDbl(Val(txtStartDistance)
'End If
'If DistanceStart >= 0 Then
''    Dim ConvertToDistCount As Double '''''''''''''''
''    Dim DistanceCounter As Long                    '
''                                                   '
''    ConvertToDistCount = DistanceStart * 10        '
''    ConvertToDistCount = Int(ConvertToDistCount)   '
''    DistanceCounter = CLng(ConvertToDistCount)     ''''''''
'    Call ClearLineScreen.DistanceCounterSet(DistanceStart) 'PCN2874 Also fogot to copy over chnages from 5.4.2
'End If


Exit Function
Err_Handler:
Select Case Err
    Case 13 'Invalid data
        DistanceStart = InvalidData 'PCN3884 was -1
        Resume Next
    Case Else
        MsgBox Err & "-" & error$
End Select
End Function


Function SetupControlLayout(ProcessMethod As String)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : SetupControlLayout
'Created : 15 March 2004, PCN2639
'Updated :
'Prg By  : Geoff Logan
'Param   : ProcessMethod - Name of process
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Select Case ProcessMethod
    Case "AutomaticCounter"
        CountPicBox.Visible = True
        CalculatedDist.Visible = True
        CalculatedDist_lbl.Visible = True
        btnCountDirection.Visible = True: If IPD Then btnCountDirection.Visible = False
        CountPicBox.Enabled = True: If IPD Then CountPicBox.Visible = False
        CalculatedDist.Enabled = True: If IPD Then CalculatedDist.Enabled = False
        CalculatedDist_lbl.Enabled = True
        btnCountDirection.Enabled = True
        txtEndDistance.Visible = False
        lblFinishDistance.Visible = False
        RecalculateDistButton.Top = 3225
        lblUnit2.Top = 2700
        lblUnit2.ForeColor = &H8000&
        lblUnit2.FontBold = True
        lblStartDistance.Enabled = True
        txtStartDistance.Enabled = True: If IPD Then txtStartDistance.Enabled = False
        lblUnit.Enabled = True
        lblUnit2.Enabled = True
        If PVDFileName = "" Or PVDataNoOfLines = 0 Or PVRecording Then
            RecalculateDistButton.Enabled = False
        Else
            RecalculateDistButton.Enabled = True
        End If
    Case "StartFinishEstimate"
        Call SetupControlLayoutStartFinishEst 'PCN2928
        
    Case "AvgTime"
        Call SetupControlLayoutStartFinishEst 'PCN2928
        
    Case "AvgFrame"
        Call SetupControlLayoutStartFinishEst 'PCN2928
    
    Case Else
        CountPicBox.Visible = True
        CalculatedDist.Visible = True
        CalculatedDist_lbl.Visible = True
        btnCountDirection.Visible = True
        CountPicBox.Enabled = False
        CalculatedDist.Enabled = False
        CalculatedDist_lbl.Enabled = False
        btnCountDirection.Enabled = False
        txtEndDistance.Visible = False
        lblFinishDistance.Visible = False
        RecalculateDistButton.Top = 3225
        lblUnit2.Top = 2700
        lblUnit2.ForeColor = &H8000&
        lblUnit2.FontBold = True
        lblStartDistance.Enabled = False
        txtStartDistance.Enabled = False
        lblUnit.Enabled = False
        lblUnit2.Enabled = False
        RecalculateDistButton.Enabled = False
End Select

Exit Function
Err_Handler:
    MsgBox Err & "-" & error$
End Function


Private Sub txtStartDistance_Change()
On Error GoTo Err_Handler
    If Not IsNumeric(txtStartDistance.text) Then Exit Sub
    If DistanceMethod = "AutomaticCounter" Then
        DistanceStart = CDbl(txtStartDistance)
        Call ClearLineScreen.DistanceCounterSet(CDbl(DistanceStart))
    End If
Exit Sub
Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & "-" & error$
End Select

End Sub

Private Sub txtStartDistance_LostFocus()
'****************************************************************************************
'Name    : txtStartDistance_LostFocus
'Created : 27 March 2004, PCN2639
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    : Validates data entry when txtStartDistance has not yet been validated
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
Dim FileLoadError As Boolean

If Len(txtStartDistance) <> 0 Then
    DistanceStart = CDbl(txtStartDistance) 'PCN2898 removed the function Val eg was CDbl(Val(txtStartDistance)
    If DistanceStart > InvalidData Then 'PCN3884 was >=0
        txtStartDistance = DistanceStart
    Else
        DistanceStart = ""
    End If
End If



Exit Sub
Err_Handler:
Select Case Err
    Case 13 'Invalid data
        DistanceStart = InvalidData 'PCN3884 was -1
        Resume Next
    Case Else
        MsgBox Err & "-" & error$
End Select
End Sub

Private Sub txtStartDistance_Validate(Cancel As Boolean)
'****************************************************************************************
'Name    : txtStartDistance_Validate
'Created : 23 March 2004, PCN2639
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    : Validates data entry
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
Dim FileLoadError As Boolean

If Len(txtStartDistance) = 0 Then
'   DistanceStart = -1 'PCN3744 needed to egnore it, not change the DistanceStart
    Exit Sub
End If

DistanceStart = CDbl(txtStartDistance) 'PCN2898 removed the function Val eg was CDbl(Val(txtStartDistance)

If DistanceStart > InvalidData Then 'PCN3884 was >= 0
    'If PVD has not yet been recorded and DistanceMethod is automatic then store the DistanceStart
    If (PVDFileName = "" Or PVDataNoOfLines = 0) And DistanceMethod = "AutomaticCounter" Then
        ConfigInfo.DistanceStart = DistanceStart
        Call ClearLineScreen.DistanceCounterSet(DistanceStart) 'PCN2874 Also fogot to copy over chnages from 5.4.2 '<---|
    Else
        'Otherwise can only change the stored ConfigInfo start when the recalculation button is pressed.
        'Set Distance Counter analyzer
'        Dim ConvertToDistCount As Double '''''''''''''''
'        Dim DistanceCounter As Long                    '
                                                        '
'        ConvertToDistCount = DistanceStart * 10        '
'        ConvertToDistCount = Int(ConvertToDistCount)   '
'        DistanceCounter = CLng(ConvertToDistCount)     ''''''''
    '    Call ClearLineScreen.DistanceCounterSet(DistanceStart) 'PCN2874 Also fogot to copy over chnages from 5.4.2
    End If
Else
    MsgBox DisplayMessage("Please input valid data") 'PCN2639
End If

Exit Sub
Err_Handler:
Select Case Err
    Case 13 'Invalid data
        DistanceStart = InvalidData 'PCN3884
        Resume Next
    Case Else
        MsgBox Err & "-" & error$
End Select
End Sub

'Function DoubleValidation(Value As String, ConvertedValue)
''****************************************************************************************
''Name    : DataEntryValidation
''Created : 26 March 2004, PCN2639
''Updated :
''Prg By  : Geoff Logan
''Param   : Value - the value to be tested
''Desc    : Validates data entry
''Usage   :
''****************************************************************************************
'On Error GoTo Err_Handler
'
'DistanceStart = CDbl(Val(txtStartDistance))
'
'If DistanceStart >= 0 Then
'    'If PVD has not yet been recorded and DistanceMethod is automatic then store the DistanceStart
'    If (PVDFileName = "" Or PVDataNoOfLines = 0) And DistanceMethod = "AutomaticCounter" Then
'        ConfigInfo.DistanceStart = DistanceStart
'    Else
'        'Otherwise can only change the stored ConfigInfo start when the recalculation button is pressed.
'    End If
'Else
'    MsgBox DisplayMessage("Invalid data")
'End If
'
'Exit Function
'Err_Handler:
'Select Case Err
'    Case 13 'Invalid data
'        DistanceStart = -1
'        Resume Next
'    Case Else
'        MsgBox Err & "-" & error$
'End Select
'End Function

Function DistanceFixedPtSetup(DistanceFinish As Double)
'****************************************************************************************
'Name    : DistanceFixedPtSetup
'Created : 29 July 2004, PCN2928
'Updated :
'Prg By  : Geoff Logan
'Param   : DistanceFinish - Estimate finish distance
'Desc    : Sets up the DistanceCounterFixedPt array with the DistanceStart, DistanceFinish
'          and distance information from the PipeObservations array.
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
Dim FixedPointNo As Integer 'PCN2928 The Fixed Point index number
ReDim DistanceCounterFixedPt(NoOfPipeObservations + 2)

'Set PVD Start as a Fixed Point
DistanceCounterFixedPt(1).FixedPtFrameNo = 1
DistanceCounterFixedPt(1).FixedPtDist = DistanceStart
'Set PVD Finish as a Fixed Point
DistanceCounterFixedPt(2).FixedPtFrameNo = PVDataNoOfLines
DistanceCounterFixedPt(2).FixedPtDist = DistanceFinish
'Load all Observations
For FixedPointNo = 3 To NoOfPipeObservations + 2

    DistanceCounterFixedPt(FixedPointNo).FixedPtFrameNo = PipeObservations(FixedPointNo - 2).PipeObsFrameNo
    DistanceCounterFixedPt(FixedPointNo).FixedPtDist = PipeObservations(FixedPointNo - 2).PipeObsDist

Next FixedPointNo

Exit Function
Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & "-" & error$
End Select
End Function

Function DistanceFixedPtPVDistanceCalc(CurrentFrame As Long, FixedPointStart As Integer, FixedPointFinish As Integer, FrameFinish As Long, FixedPointStartTime As Double, FixedPointFinishTime As Double, TimeFinish As Double, CurrentTime As Double)
'****************************************************************************************
'Name    : DistanceFixedPtPVDistanceCalc
'Created : 29 July 2004, PCN2928
'Updated :
'Prg By  : Geoff Logan
'Param   : CurrentFrame - Specifies which PVFrameNo requires PVDistance calculation.
'Desc    : Determines PVDistances(CurrentFrame) based on the known (fixed) distance points.
'Usage   : Used only for the Start/Finish Distance Estimation method (time and frame based).
'****************************************************************************************
On Error GoTo Err_Handler
Dim CorrectedCameraSpeed As Double 'Adjusts camera speed setting to the two known points FixedPointStartDist and FixedPointFinishDist, which may not be the start/finish of the PVD.
Dim FrameNoDifference As Long
Dim NewStartFrameNo As Long
Dim NewFinishDiff As Long
Dim FixedPointNo As Integer 'The Fixed Point index number

If CurrentFrame > DistanceCounterFixedPt(FixedPointFinish).FixedPtFrameNo Then
    'Find the Start and Finish points for this frame
    FixedPointStart = FixedPointFinish
    NewStartFrameNo = DistanceCounterFixedPt(FixedPointStart).FixedPtFrameNo
    FixedPointStartTime = FixedPointFinishTime
    NewFinishDiff = FrameFinish 'Reset
    
    For FixedPointNo = 1 To NoOfPipeObservations + 2
        If DistanceCounterFixedPt(FixedPointNo).FixedPtFrameNo > NewStartFrameNo Then
            FrameNoDifference = DistanceCounterFixedPt(FixedPointNo).FixedPtFrameNo - NewStartFrameNo
            If FrameNoDifference <= NewFinishDiff Then
                NewFinishDiff = FrameNoDifference
                FixedPointFinish = FixedPointNo 'Set new finish value
            End If
        End If
    Next FixedPointNo
    If FixedPointFinish = 2 Then
        FixedPointFinishTime = TimeFinish
    Else
        FixedPointFinishTime = PVTimes(DistanceCounterFixedPt(FixedPointFinish).FixedPtFrameNo)
    End If
End If

If (FixedPointFinishTime - FixedPointStartTime) = 0 Then
    CorrectedCameraSpeed = CameraSpeedInTime
Else
    CorrectedCameraSpeed = (DistanceCounterFixedPt(FixedPointFinish).FixedPtDist - DistanceCounterFixedPt(FixedPointStart).FixedPtDist) / (FixedPointFinishTime - FixedPointStartTime)
End If

PVDistances(CurrentFrame) = Round(CorrectedCameraSpeed * (CDbl(CurrentTime) - FixedPointStartTime) + DistanceCounterFixedPt(FixedPointStart).FixedPtDist, 2)
    

Exit Function
Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & "-" & error$
End Select
End Function

Function SetupControlLayoutStartFinishEst()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : SetupControlLayoutStartFinishEst
'Created : 2 Sept 2004, PCN2928
'Updated :
'Prg By  : Geoff Logan
'Param   : ProcessMethod - Name of process
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

btnCountDirection.Visible = False
CountPicBox.Visible = False
'vvvv PCN2928 ***********************************
'Setup for Fixed Pt correction
RecalculateDistButton.Top = 3225
CalculatedDist.Visible = True
CalculatedDist_lbl.Visible = True
CalculatedDist.Enabled = True
CalculatedDist_lbl.Enabled = True
'^^^^ *******************************************
txtEndDistance.Top = 1600
txtEndDistance.Visible = True
lblFinishDistance.Top = 1600
lblFinishDistance.Visible = True
lblUnit2.Top = 1600
lblUnit2.ForeColor = &H8000000D
lblUnit2.FontBold = False
lblStartDistance.Enabled = True
txtStartDistance.Enabled = True
lblUnit.Enabled = True
lblUnit2.Enabled = True

If PVDFileName = "" Or PVDataNoOfLines = 0 Or PVRecording Then
    RecalculateDistButton.Enabled = False
Else
    RecalculateDistButton.Enabled = True
End If

Exit Function
Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & "-" & error$
End Select
End Function

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
        MsgBox Err & "-" & error$
End Select
End Sub


