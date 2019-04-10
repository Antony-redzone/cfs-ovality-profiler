VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form DataEntryForm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7515
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6450
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PipeDetailsChangePicture 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   3945
      TabIndex        =   18
      Top             =   5040
      Width           =   3975
      Begin RichTextLib.RichTextBox NewLabelPipeDetail 
         Height          =   345
         Left            =   1560
         TabIndex        =   21
         Top             =   120
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   609
         _Version        =   393217
         BackColor       =   15728639
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"DataEntry.frx":0000
      End
      Begin VB.TextBox NewLabelPipeDetailOld 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2280
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   2300
      End
      Begin VB.Label NewPipeDetailTextLabel 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "New Label"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   300
         Left            =   120
         TabIndex        =   20
         Top             =   165
         Width           =   1335
      End
   End
   Begin MSComctlLib.ImageList CalibrationImageList 
      Left            =   5040
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataEntry.frx":0082
            Key             =   "CalibrationH"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataEntry.frx":041C
            Key             =   "CalibrationV"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataEntry.frx":07B6
            Key             =   "CalibrationCrack"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox CalibrationChangePicture 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   4305
      TabIndex        =   15
      Top             =   4200
      Width           =   4335
      Begin VB.TextBox NewCalibration 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00EFFFFF&
         Height          =   315
         Left            =   3240
         TabIndex        =   16
         Top             =   120
         Width           =   975
      End
      Begin VB.Image CalImage 
         Height          =   240
         Left            =   120
         Picture         =   "DataEntry.frx":0B50
         Top             =   160
         Width           =   240
      End
      Begin VB.Label CalibrationLabel 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Enter calibration length"
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
         Height          =   300
         Left            =   360
         TabIndex        =   17
         Top             =   165
         Width           =   2775
      End
   End
   Begin VB.PictureBox LimitLineChangePicture 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   3105
      TabIndex        =   11
      Top             =   2760
      Width           =   3135
      Begin VB.TextBox NewLimitLine 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00EFFFFF&
         Height          =   315
         Left            =   1680
         TabIndex        =   12
         Top             =   120
         Width           =   975
      End
      Begin VB.Label LimitLineLabel 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Limit Line"
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
         Height          =   300
         Left            =   120
         TabIndex        =   14
         Top             =   160
         Width           =   1455
      End
      Begin VB.Label NewLimitLineUnits 
         BackColor       =   &H00C0FFFF&
         Caption         =   "%"
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
         Height          =   300
         Left            =   2760
         TabIndex        =   13
         Top             =   165
         Width           =   375
      End
   End
   Begin VB.PictureBox DistanceStartEndChangePictureBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   3825
      TabIndex        =   6
      Top             =   1440
      Width           =   3855
      Begin VB.TextBox CurrentStartEndDistance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0F0F0&
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox NewStartEndDistance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00EFFFFF&
         Height          =   315
         Left            =   2520
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label CurrentStartEndLabel 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Current start distance"
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
         Height          =   300
         Left            =   0
         TabIndex        =   10
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label NewStartEndlabel 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "New start distance"
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
         Left            =   0
         TabIndex        =   9
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.PictureBox DistanceChangePictureBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   3825
      TabIndex        =   1
      Top             =   240
      Width           =   3855
      Begin VB.TextBox NewDistance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00EFFFFF&
         Height          =   315
         Left            =   2520
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox CurrentDistance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0F0F0&
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label NewDistanceLabel 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "New Distance"
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
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label CurrentDistanceLabel 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.PictureBox DragBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00B36A36&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   3825
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.Image CloseImage 
         Height          =   240
         Left            =   2760
         Picture         =   "DataEntry.frx":0EDA
         Top             =   0
         Width           =   240
      End
   End
End
Attribute VB_Name = "DataEntryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LastMouseDownX As Single
Dim LastMouseDownY As Single
Dim Action As String
Dim DataEntryType As String
Dim WorkingFrame As Long
Dim PVGraphIndex As Integer
Dim CloseOnDeactivate As Boolean

Dim DataEntryBox As TextBox
Dim RichDataEntryBox As RichTextBox
Dim PipeDetailAnsi As String

'vvvv PCN4368 ************************************
Public Event MouseLeave()
'Public Event MouseHover()
Dim WithEvents MouseTrackDragBar As clsTrackInfo
Attribute MouseTrackDragBar.VB_VarHelpID = -1
'^^^^ ********************************************


Private Sub DragBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    LastMouseDownX = X
    LastMouseDownY = Y
    If Button = 1 Then Action = "Move"
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-DE1:" & Error$
    End Select
End Sub


Private Sub Form_Initialize()
    Me.Visible = False
End Sub

Private Sub MouseTrackDragBar_MouseLeave() 'PCN4368
On Error GoTo Err_Handler

Me.DragBar.BackColor = &HB36A36

RaiseEvent MouseLeave
Exit Sub
Err_Handler:
    MsgBox Err & "-DE2:" & Error$
End Sub

Private Sub DragBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.DragBar.BackColor = &HC0C000 'PCN4368

    If Action = "Move" Then
        Me.Left = Me.Left + X - LastMouseDownX
        Me.Top = Me.Top + Y - LastMouseDownY
    End If
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-DE3:" & Error$
    End Select
End Sub

Private Sub DragBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    Action = ""
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-DE4:" & Error$
    End Select
End Sub



Private Sub Form_Deactivate()
On Error GoTo Err_Handler
    If CloseOnDeactivate Then Call CloseDataEntry
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-DE5:" & Error$
    End Select
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

CloseOnDeactivate = False
Me.Visible = False
Call ConvertLanguage(Me, Language) 'PCN4171

'vvvv PCN4328 ************************************
'Initilise the mouse leave event on the key's drag bar.
Set MouseTrackDragBar = New clsTrackInfo
MouseTrackDragBar.hwnd = DragBar.hwnd

StartTrack MouseTrackDragBar
'^^^^ ********************************************

    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-DE6:" & Error$
    End Select
End Sub



Private Sub CloseImage_Click()
On Error GoTo Err_Handler
    
If LanguageCharset <> 0 Then
    PipeDetailAnsi = RichDataEntryBox.text
End If
Call CloseDataEntry
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-DE7:" & Error$
    End Select
End Sub



Private Sub CloseDataEntry()



On Error GoTo Err_Handler
Select Case DataEntryType
    Case "DistanceChange": Call DistanceChangeConfirm
    Case "DistanceStartChange": Call DistanceChangeConfirm
    Case "DistanceEndChange": Call DistanceChangeConfirm
    Case "LimitLineChangeLeft", "LimitLineChangeRight", "LimitLineChangeBoth": Call LimitLineChangeConfirm
    Case "CalibrationChange": Call CalibrationChangeConfirm
    Case "AssetNoLabelChange", _
         "AssetNoTextChange", _
         "SiteIDLabelChange", _
         "SiteIDTextChange", _
         "CityLabelChange", _
         "CityTextChange", _
         "DateLabelChange", _
         "DateTextChange", _
         "TimeLabelChange", _
         "TimeTextChange", _
         "StNodeLabelChange", _
         "StNodeTextChange", _
         "StNodeLocLabelChange", _
         "StNodeLocTextChange", _
         "FhNodeLabelChange", _
         "FhNodeTextChange", _
         "FhNodeLocLabelChange", _
         "FhNodeLocTextChange", _
         "IntDiaLabelChange", _
         "OutDiaLabelChange", _
         "PipeLenLabelChange", _
         "MaterialLabelChange", _
         "MaterialTextChange":
         
        If LCase(CharacterType) <> "unicode" Then
            Call PipelineDetailsLabelChangeConfirm 'PCN4171
        Else
            Call PipelineDetailsLabelChangeConfirmAnsi
        End If
        
        
End Select

Unload Me
        
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-DE8:" & Error$
    End Select
End Sub

Public Sub SetDataEntryType(ByVal DataType As String, Optional DataValue, Optional GraphIndex)
On Error GoTo Err_Handler
    CloseOnDeactivate = False
    Me.Visible = False
    DataEntryType = DataType
    Select Case DataType
        Case "DistanceChange": Set DataEntryBox = NewDistance: Call DistanceChangeSetUp
            DoEvents 'Make sure all the windows are active before setting this otherwise
            'This will cause the data entry box to close
            CloseOnDeactivate = True
            Me.Visible = True
            Me.ZOrder 0
        Case "DistanceStartChange": Set DataEntryBox = NewStartEndDistance: Call DistanceStartChangeSetUp
            DoEvents 'Make sure all the windows are active before setting this otherwise
            'This will cause the data entry box to close
            CloseOnDeactivate = True
            Me.Visible = True
            Me.ZOrder 0
        Case "DistanceEndChange": Set DataEntryBox = NewStartEndDistance: Call DistanceEndChangeSetUp
            DoEvents 'Make sure all the windows are active before setting this otherwise
            'This will cause the data entry box to close
            CloseOnDeactivate = True
            Me.Visible = True
            Me.ZOrder 0
        Case "LimitLineChangeLeft", "LimitLineChangeRight", "LimitLineChangeBoth"
            Set DataEntryBox = NewLimitLine
            PVGraphIndex = GraphIndex
            Call LimitLineChangeSetUp(DataValue)
            DoEvents 'Make sure all the windows are active before setting this otherwise
            'This will cause the data entry box to close
            CloseOnDeactivate = True
            Me.Visible = True
            Me.ZOrder 0
        Case "CalibrationChange"
            Set DataEntryBox = NewCalibration
'            PVGraphIndex = GraphIndex
            Call CalibrationChangeSetUp(DataValue)
            DoEvents 'Make sure all the windows are active before setting this otherwise
            'This will cause the data entry box to close
            CloseOnDeactivate = True
            Me.Visible = True
            Me.ZOrder 0
            Me.NewCalibration.SetFocus 'PCN6026
        Case "AssetNoLabelChange", "SiteIDLabelChange", "CityLabelChange", "DateLabelChange", "TimeLabelChange", "StNodeLabelChange", "StNodeLocLabelChange", "FhNodeLabelChange", "FhNodeLocLabelChange", "IntDiaLabelChange", "OutDiaLabelChange", "PipeLenLabelChange", "MaterialLabelChange"
            Set RichDataEntryBox = NewLabelPipeDetail
            RichDataEntryBox.Font.Charset = LanguageCharset
            Call PipeDetailsLabelChangeSetUp(DataValue)
            DoEvents 'Make sure all the windows are active before setting this otherwise
            'This will cause the data entry box to close
            CloseOnDeactivate = True
            Me.Visible = True
            Me.ZOrder 0
            RichDataEntryBox.SetFocus
    End Select

    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-DE9:" & Error$
    End Select
End Sub

Sub DistanceChangeSetUp()
On Error GoTo Err_Handler
Dim ErrorStr As String 'PCN4171

    Call ClearLineScreen.VideoScreenActivate(ErrorStr) 'PCN4171
    DistanceChangePictureBox.Left = 0
    DistanceChangePictureBox.Top = DragBar.height
    DragBar.width = DistanceChangePictureBox.width
    CloseImage.Left = DragBar.width - CloseImage.width
    
    Me.width = DragBar.width
    Me.height = DragBar.height + DistanceChangePictureBox.height
    
    CurrentDistanceLabel.Caption = DisplayMessage("Current distance")
    NewDistanceLabel.Caption = DisplayMessage("New distance")
    
    DistanceChangePictureBox.Visible = True
    CurrentDistance.text = Round(PVDistances(PVFrameNo), 2)
    WorkingFrame = PVFrameNo
    'Me.Visible
    'Me.ZOrder 0
    
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-DE10:" & Error$
    End Select
End Sub

Sub DistanceStartChangeSetUp()
On Error GoTo Err_Handler
Dim ErrorStr As String 'PCN4171

    Call ClearLineScreen.VideoScreenActivate(ErrorStr) 'PCN4171
    DistanceStartEndChangePictureBox.Left = 0
    DistanceStartEndChangePictureBox.Top = DragBar.height
    DragBar.width = DistanceStartEndChangePictureBox.width
    CloseImage.Left = DragBar.width - CloseImage.width
    
    Me.width = DragBar.width
    Me.height = DragBar.height + DistanceStartEndChangePictureBox.height
    
    CurrentStartEndLabel.Caption = DisplayMessage("Current start distance")
    NewStartEndlabel.Caption = DisplayMessage("New start distance")
    
    DistanceStartEndChangePictureBox.Visible = True
    CurrentStartEndDistance.text = ConfigInfo.DistanceStart
    If LanguageCharset <> 0 Then
        CurrentStartEndDistance.Font.Charset = LanguageCharset
    End If
    If ConfigInfo.DistanceStart = InvalidData Then CurrentStartEndDistance.text = DisplayMessage("Not set.")
    
    
    WorkingFrame = 1
    'Me.Visible
    'Me.ZOrder 0

    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-DE11:" & Error$
        'Resume 'ANT
    End Select
End Sub

Sub DistanceEndChangeSetUp()
On Error GoTo Err_Handler
Dim ErrorStr As String 'PCN4171

    Call ClearLineScreen.VideoScreenActivate(ErrorStr) 'PCN4171
    DistanceStartEndChangePictureBox.Left = 0
    DistanceStartEndChangePictureBox.Top = DragBar.height
    DragBar.width = DistanceStartEndChangePictureBox.width
    CloseImage.Left = DragBar.width - CloseImage.width
    
    Me.width = DragBar.width
    Me.height = DragBar.height + DistanceStartEndChangePictureBox.height
    
    CurrentStartEndLabel.Caption = DisplayMessage("Current finish distance")
    NewStartEndlabel.Caption = DisplayMessage("New finish distance")
    
    DistanceStartEndChangePictureBox.Visible = True
    CurrentStartEndDistance.text = ConfigInfo.DistanceFinish
    If LanguageCharset <> 0 Then
        CurrentStartEndDistance.Font.Charset = LanguageCharset
    End If
    If ConfigInfo.DistanceFinish = InvalidData Then CurrentStartEndDistance.text = DisplayMessage("Not set.")
    WorkingFrame = PVDataNoOfLines
    'Me.Visible
    'Me.ZOrder 0
    
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-DE12:" & Error$
    End Select
End Sub

Sub LimitLineChangeSetUp(ByVal DataValue As Double)
On Error GoTo Err_Handler
Dim ErrorStr As String 'PCN4171
Dim UnitsIndex As Integer
Dim NoOfGraphs As Integer
Dim UnitType As String
Dim ConvertedDataValue As Double
Dim DisplayUnits As String

'Call ClearLineScreen.VideoScreenActivate(ErrorStr) 'PCN4171
LimitLineChangePicture.Left = 0
LimitLineChangePicture.Top = DragBar.height
DragBar.width = LimitLineChangePicture.width
CloseImage.Left = DragBar.width - CloseImage.width

Me.width = DragBar.width
Me.height = DragBar.height + LimitLineChangePicture.height

LimitLineLabel.Caption = DisplayMessage("Limit Line")

LimitLineChangePicture.Visible = True

'vvvv PCN4207 ************************************
''NoOfGraphs = UBound(PVGraphOrder)
''For UnitsIndex = 0 To NoOfGraphs
''    If PVGraphOrder(UnitsIndex) = ImageGraphState(0).PreviousGraphType Then
''        UnitType = PVXScaleUnits(UnitsIndex)
''        Exit For
''    End If
''Next UnitsIndex

UnitsIndex = GetGraphInfoIndex(0)


'PCN6458 If ImageGraphState(0).GraphType = "Inclination" Then
'PCN6458     If MeasurementUnits = "mm" Then
'PCN6458         NewLimitLine.text = Format(DataValue, "#0.0")
'PCN6458     Else
'PCN6458         NewLimitLine.text = Format(DataValue, "#0.00")
'PCN6458     End If
'PCN6458     NewLimitLineUnits.Caption = MeasurementUnits
If ImageGraphState(0).GraphType = "XYDiameter" And MedianFlat Then 'PCN5186 added the next four lines
    ConvertedDataValue = ConvertRealToPerByGraph(DataValue, 0, DisplayUnits)
    NewLimitLine.text = Format(ConvertedDataValue, "#0.0")
    NewLimitLineUnits.Caption = "%"
ElseIf GraphInfoContainer(UnitsIndex).PVXScaleUnits = "Real" Then ' Or PVXScaleUnits(UnitsIndex) = "Area" Then
    If MeasurementUnits = "mm" Then
        NewLimitLine.text = Format(DataValue, "#0")
    Else
        NewLimitLine.text = Format(DataValue, "#0.0")
    End If
    NewLimitLineUnits.Caption = MeasurementUnits
Else
    ConvertedDataValue = ConvertRealToPerByGraph(DataValue, 0, DisplayUnits)
    NewLimitLine.text = Format(ConvertedDataValue, "#0.0")
    NewLimitLineUnits.Caption = "%"
End If
'^^^^ ********************************************

WorkingFrame = PVDataNoOfLines

Me.Top = PrecisionVisionGraph.height - 2000
Me.Left = PrecisionVisionGraph.Left + 500
'Me.Visible = True
'Me.ZOrder 0

    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-DE13:" & Error$
    End Select
End Sub


Sub CalibrationChangeSetUp(ByVal DataValue As Double)
On Error GoTo Err_Handler
Dim ErrorStr As String 'PCN4171

'Call ClearLineScreen.VideoScreenActivate(ErrorStr) 'PCN4171 cancel out this change, we want the profiler
                                                    'snapshot mode not videomode
CalibrationChangePicture.Left = 0
CalibrationChangePicture.Top = DragBar.height
DragBar.width = CalibrationChangePicture.width
CloseImage.Left = DragBar.width - CloseImage.width

Me.width = DragBar.width
Me.height = DragBar.height + CalibrationChangePicture.height

CalibrationLabel.Caption = DisplayMessage("Enter calibration length")

If CLPScreenAction = "DrawCalibrationLine" Then
    Set Me.CalImage.Picture = Me.CalibrationImageList.ListImages("CalibrationH").Picture
ElseIf CLPScreenAction = "DrawHorCalibrationLine" Then
    Set Me.CalImage.Picture = Me.CalibrationImageList.ListImages("CalibrationV").Picture
Else
    Set Me.CalImage.Picture = Me.CalibrationImageList.ListImages("CalibrationCrack").Picture
End If

CalibrationChangePicture.Visible = True

WorkingFrame = PVDataNoOfLines

Me.Top = 3 * ClearLineScreen.height / 4
Me.Left = ClearLineScreen.width / 2 - 2000
'Me.Visible = True
'Me.ZOrder 0

    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-DE14:" & Error$
    End Select
End Sub

Sub PipeDetailsLabelChangeSetUp(ByVal DataValue As String) 'PCN4171
On Error GoTo Err_Handler
Dim ErrorStr As String

PipeDetailsChangePicture.Left = 0
PipeDetailsChangePicture.Top = DragBar.height
DragBar.width = PipeDetailsChangePicture.width
CloseImage.Left = DragBar.width - CloseImage.width
    
Me.width = DragBar.width
Me.height = DragBar.height + PipeDetailsChangePicture.height
Me.Top = PipeDetailsChangeTop
Me.Left = PipelineDetails.Left - Me.width
    
NewPipeDetailTextLabel.Caption = DisplayMessage("New Label")
    
PipeDetailsChangePicture.Visible = True
'NewLabelPipeDetail.Font.Charset = 128
If LCase(CharacterType) <> "unicode" Then
    NewLabelPipeDetail.text = DataValue
Else
    PipeDetailAnsi = DataValue
End If

WorkingFrame = PVDataNoOfLines
'Me.Visible
'Me.ZOrder 0
    
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-DE15:" & Error$
    End Select
End Sub

Function PipeDetailsChangeTop() As Integer 'PCN4171
Dim ErrorStr As String

Select Case DataEntryType
    Case "AssetNoLabelChange"
        PipeDetailsChangeTop = PipelineDetails.AssetNo.Top + PipelineDetails.AssetInfo.Top - 450
    Case "AssetNoTextChange"
        PipeDetailsChangeTop = PipelineDetails.AssetNo.Top + PipelineDetails.AssetInfo.Top - 450
    Case "SiteIDLabelChange"
        PipeDetailsChangeTop = PipelineDetails.SiteID.Top + PipelineDetails.AssetInfo.Top - 450
    Case "SiteIDTextChange"
        PipeDetailsChangeTop = PipelineDetails.SiteID.Top + PipelineDetails.AssetInfo.Top - 450
    Case "CityLabelChange"
        PipeDetailsChangeTop = PipelineDetails.City.Top + PipelineDetails.AssetInfo.Top - 450
    Case "CityTextChange"
        PipeDetailsChangeTop = PipelineDetails.City.Top + PipelineDetails.AssetInfo.Top - 450
    Case "DateLabelChange"
        PipeDetailsChangeTop = PipelineDetails.sDate.Top + PipelineDetails.AssetInfo.Top - 450
    Case "DateTextChange"
        PipeDetailsChangeTop = PipelineDetails.sDate.Top + PipelineDetails.AssetInfo.Top - 450
    Case "TimeLabelChange"
        PipeDetailsChangeTop = PipelineDetails.sTime.Top + PipelineDetails.AssetInfo.Top - 450
    Case "TimeTextChange"
        PipeDetailsChangeTop = PipelineDetails.sTime.Top + PipelineDetails.AssetInfo.Top - 450
    Case "StNodeLabelChange"
        PipeDetailsChangeTop = PipelineDetails.StartNodeNo.Top + PipelineDetails.StartNodeFrame.Top - 450
    Case "StNodeTextChange"
        PipeDetailsChangeTop = PipelineDetails.StartNodeNo.Top + PipelineDetails.StartNodeFrame.Top - 450
    Case "StNodeLocLabelChange"
        PipeDetailsChangeTop = PipelineDetails.StNodeLoc_lbl.Top + PipelineDetails.StartNodeFrame.Top - 450
    Case "StNodeLocTextChange"
        PipeDetailsChangeTop = PipelineDetails.StNodeLoc_lbl.Top + PipelineDetails.StartNodeFrame.Top - 450
    Case "FhNodeLabelChange"
        PipeDetailsChangeTop = PipelineDetails.FinishNodeNo.Top + PipelineDetails.FinishNodeFrame.Top - 450
    Case "FhNodeTextChange"
        PipeDetailsChangeTop = PipelineDetails.FinishNodeNo.Top + PipelineDetails.FinishNodeFrame.Top - 450
    Case "FhNodeLocLabelChange"
        PipeDetailsChangeTop = PipelineDetails.FhNodeLoc_lbl.Top + PipelineDetails.FinishNodeFrame.Top - 450
    Case "FhNodeLocTextChange"
        PipeDetailsChangeTop = PipelineDetails.FhNodeLoc_lbl.Top + PipelineDetails.FinishNodeFrame.Top - 450
    Case "IntDiaLabelChange"
        PipeDetailsChangeTop = PipelineDetails.InternalDiameterExpected.Top + PipelineDetails.PipeDataFrame.Top - 450
    Case "OutDiaLabelChange"
        PipeDetailsChangeTop = PipelineDetails.OutsideDiameter.Top + PipelineDetails.PipeDataFrame.Top - 450
    Case "PipeLenLabelChange"
        PipeDetailsChangeTop = PipelineDetails.PipeLength.Top + PipelineDetails.PipeDataFrame.Top - 450
    Case "MaterialLabelChange"
        PipeDetailsChangeTop = PipelineDetails.Material.Top + PipelineDetails.PipeDataFrame.Top - 450
    Case "MaterialTextChange"
        PipeDetailsChangeTop = PipelineDetails.Material.Top + PipelineDetails.PipeDataFrame.Top - 450
    Case Else
        PipeDetailsChangeTop = 0
End Select

Exit Function
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-DE16:" & Error$
    End Select
End Function
Sub DistanceChangeConfirm()
On Error GoTo Err_Handler

Dim NumberObs As Integer
Dim ObsIndex As Integer
Dim Dist As Double




If DataEntryBox.text = "" Then Exit Sub
If Not IsNumeric(DataEntryBox.text) Then Exit Sub
    
Dist = SafeCDbl(DataEntryBox.text) 'PCN4161

If DataEntryType = "DistanceStartChange" Then ConfigInfo.DistanceStart = Dist
If DataEntryType = "DistanceEndChange" Then ConfigInfo.DistanceFinish = Dist

Call Distance.DistanceAdd(Dist, WorkingFrame)

'Refresh the ControlsScreen button display
Call ControlsScreen.ControlsViewSetup

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-DE17:" & Error$
    End Select
End Sub

Sub LimitLineChangeConfirm()
On Error GoTo Err_Handler
Dim NumberObs As Integer
Dim ObsIndex As Integer
Dim Limit As Double
Dim ConvertedLimit As Double 'PCN4207
Dim DisplayUnits As String 'PCN4207
Dim NoOfGraphs As Integer 'PCN4207
Dim UnitsIndex As Integer 'PCN4207

If NewLimitLine.text = "" Then Exit Sub
If Not IsNumeric(NewLimitLine.text) Then Exit Sub
    
Limit = SafeCDbl(NewLimitLine.text) 'PCN4161

'vvvv PCN4207 *********************************************
''NoOfGraphs = UBound(PVGraphOrder)
''For UnitsIndex = 0 To NoOfGraphs
''    If PVGraphOrder(UnitsIndex) = ImageGraphState(0).PreviousGraphType Then
''        Exit For
''    End If
''Next UnitsIndex
UnitsIndex = GetGraphInfoIndex(0)

'PCN5186 do nothing if not real, following line added
If ImageGraphState(0).GraphType = "XYDiameter" And MedianFlat Then

ElseIf GraphInfoContainer(UnitsIndex).PVXScaleUnits <> "Real" Then
    If GraphInfoContainer(UnitsIndex).GraphType <> "Capacity" And GraphInfoContainer(UnitsIndex).GraphType <> "Ovality" Then 'PCN4274 needed to egnore ovality and capacity
        ConvertedLimit = ConvertPerToReal(Limit, "Dia")
    Else
        ConvertedLimit = Limit
    End If
    Limit = ConvertedLimit
'End If
'^^^^ ******************************************************
Else
'    limit = convertRealToPer(Limit,"Dia
End If

'Limit = ConvertRealToPerByGraph(Limit, 0, DisplayUnits) 'PCN4407

Select Case DataEntryType
    Case "LimitLineChangeLeft": PVXScaleLimitPerL = Limit
    Case "LimitLineChangeRight": PVXScaleLimitPerR = Limit
    Case "LimitLineChangeBoth": PVXScaleLimitPerL = Limit: PVXScaleLimitPerR = Limit
End Select

Call PrecisionVisionGraph.SetFromPVXLimits(PVGraphIndex) 'PCN4407
Call PrecisionVisionGraph.RepositionPVXLimitMarkers   'PCN2680
Call PrecisionVisionGraph.StoreLimitLinesInINI

Call DrawPVGraphs


Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-DE18:" & Error$
    End Select
End Sub

Sub CalibrationChangeConfirm()
On Error GoTo Err_Handler
Dim AdjustmentValue As Double
Dim CurrentAdjustmentValue As Double
Dim CurrentFishScale As Double
Dim NewFishScale As Double
Dim VideoHeight As Long
Dim VideoWidth As Long
Dim YScaleAdjustment As Double
Dim NewCalAsDbl As Double


If NewCalibration.text = "" Then Exit Sub
If Not IsNumeric(NewCalibration.text) Then Exit Sub
    
NewCalAsDbl = SafeCDbl(NewCalibration.text) 'PCN4161


If CLPScreenAction = "DrawHorCalibrationLine" Then
    If NewCalAsDbl = 0 Or CalLengthYScale_Global = 0 Then  'PCN1910 LS
        'MsgBox DisplayMessage("Can't calibrate with a zero length."), vbExclamation 'PCN2111
        ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Can't calibrate with a zero length."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Else
        Dim ScaleRatio As Double
        Call getscalevalue(CurrentFishScale)
        Call calculatescale
        Call getscalevalue(ScaleRatio)
        ScaleRatio = ScaleRatio / CurrentFishScale
        Call hough_GetYFishScale(CurrentAdjustmentValue)
        YScaleAdjustment = NewCalAsDbl / CalLengthYScale_Global
        YScaleAdjustment = YScaleAdjustment * CurrentAdjustmentValue
        Call hough_SetYFishScale(YScaleAdjustment)
        Call calculatescale
        Call getscalevalue(NewFishScale)
        NewFishScale = NewFishScale / ScaleRatio
        Call setscalevalue(NewFishScale)
        Call CreateFishEyeMask
        
        ConfigInfo.FishEyeRatio = NewFishScale
        Call INI_WriteBack(MyFile, "Fish_Ratio=", ConfigInfo.FishEyeRatio)
        CalLength_Global = NewFishScale / CurrentFishScale * CalLength_Global
        Call INI_WriteBack(MyFile, "CalibrationDistance=", CalLen_Global) 'PCN???? calibration was saving a vertical adjustment opps. (NewCalAsDouble) not len_global
        ConfigInfo.Ratio = CalLen_Global / CalLength_Global 'PCN3035 'PCN3640
        
        Call INI_WriteBack(MyFile, "CalibrationLineLength=", CalLength_Global)
        'Redraw the main scales
        Call DrawMainScale(ClearLineScreen.PVScreen) 'PCNGL300103 'PCN3691
        PVDrawScreenRatio = ConfigInfo.Ratio
        ConfigInfo.FishEyeHorDistortion = YScaleAdjustment 'PCN3687
        Call INI_WriteBack(MyFile, "Fish_DistortionHorizontal=", YScaleAdjustment) 'PCN3687
        Call ClearLineScreen.TakeASnapShot
        CalibrationMethodActioned = "CalibrationV" 'PCN4211

 
    End If
End If
    
If CLPScreenAction = "DrawCalibrationLine" Then
    If NewCalAsDbl = 0 Then   'PCN1910 LS
        'MsgBox DisplayMessage("Can't calibrate with a zero length."), vbExclamation 'PCN2111
        ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Can't calibrate with a zero length."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Else
        CalLen_Global = NewCalAsDbl
        Call INI_WriteBack(MyFile, "CalibrationDistance=", NewCalAsDbl)
        ConfigInfo.Ratio = CalLen_Global / CalLength_Global 'PCN3035 'PCN3640
        Call INI_WriteBack(MyFile, "CalibrationLineLength=", CalLength_Global)
        
        'Redraw the main scales
        Call DrawMainScale(ClearLineScreen.PVScreen) 'PCNGL300103 'PCN3691
        PVDrawScreenRatio = ConfigInfo.Ratio
        CalibrationMethodActioned = "CalibrationH"
        'Call ControlsScreen.SetupForCalibration 'PCN4211
     

    End If
End If
    
CLPScreenAction = ""
Call ClearLineScreen.SetupMouseIcon(116)
Call ControlsScreen.SetupForCalibration 'PCN4211


Exit Sub
Err_Handler:
    MsgBox Err & "-DE19:" & Error$
End Sub

Sub PipelineDetailsLabelChangeConfirm()
On Error GoTo Err_Handler


Select Case DataEntryType
    Case "AssetNoLabelChange"
        PipelineDetails.AssetNo_lbl.Caption = NewLabelPipeDetail.text
        Call INI_WriteBack(MyFile, "PipeDetailsAssetNo=", NewLabelPipeDetail.text) 'PCN2820

    Case "SiteIDLabelChange"
        PipelineDetails.SiteID_lbl.Caption = NewLabelPipeDetail.text
        Call INI_WriteBack(MyFile, "PipeDetailsSiteID=", NewLabelPipeDetail.text)  'PCN2820

    Case "CityLabelChange"
        PipelineDetails.City_lbl.Caption = NewLabelPipeDetail.text
        Call INI_WriteBack(MyFile, "PipeDetailsCity=", NewLabelPipeDetail.text) 'PCN2820

    Case "DateLabelChange"
        PipelineDetails.Date_lbl.Caption = NewLabelPipeDetail.text
        Call INI_WriteBack(MyFile, "PipeDetailsDate=", NewLabelPipeDetail.text) 'PCN2820

    Case "TimeLabelChange"
        PipelineDetails.Time_lbl.Caption = NewLabelPipeDetail.text
        Call INI_WriteBack(MyFile, "PipeDetailsTime=", NewLabelPipeDetail.text) 'PCN2820

    Case "StNodeLabelChange"
        PipelineDetails.StartNodeNo_lbl.Caption = NewLabelPipeDetail.text
        Call INI_WriteBack(MyFile, "PipeDetailsStNode=", NewLabelPipeDetail.text) 'PCN2820

    Case "StNodeLocLabelChange"
        PipelineDetails.StNodeLoc_lbl.Caption = NewLabelPipeDetail.text
        Call INI_WriteBack(MyFile, "PipeDetailsStLoc=", NewLabelPipeDetail.text) 'PCN2820

    Case "FhNodeLabelChange"
        PipelineDetails.FinishNodeNo_lbl.Caption = NewLabelPipeDetail.text
        Call INI_WriteBack(MyFile, "PipeDetailsFhNode=", NewLabelPipeDetail.text) 'PCN2820

    Case "FhNodeLocLabelChange"
        PipelineDetails.FhNodeLoc_lbl.Caption = NewLabelPipeDetail.text
        Call INI_WriteBack(MyFile, "PipeDetailsFhLoc=", NewLabelPipeDetail.text) 'PCN2820

    Case "IntDiaLabelChange"
        PipelineDetails.InternalDiameterExpected_lbl.Caption = NewLabelPipeDetail.text
        Call INI_WriteBack(MyFile, "PipeDetailsIntDiaExp=", NewLabelPipeDetail.text) 'PCN2820

    Case "OutDiaLabelChange"
        PipelineDetails.OutsideDiameter_lbl.Caption = NewLabelPipeDetail.text
        Call INI_WriteBack(MyFile, "PipeDetailsOutDiaExp=", NewLabelPipeDetail.text) 'PCN2820

    Case "PipeLenLabelChange"
        PipelineDetails.PipeLen_lbl.Caption = NewLabelPipeDetail.text
        Call INI_WriteBack(MyFile, "PipeDetailsLength=", NewLabelPipeDetail.text) 'PCN2820

    Case "MaterialLabelChange"
        PipelineDetails.Material_lbl.Caption = NewLabelPipeDetail.text
        Call INI_WriteBack(MyFile, "PipeDetailsMaterial=", NewLabelPipeDetail.text) 'PCN2820

End Select


Exit Sub
Err_Handler:
   MsgBox Err & "-DE20:" & Error$
End Sub

Sub PipelineDetailsLabelChangeConfirmAnsi()
On Error GoTo Err_Handler

Select Case DataEntryType
    Case "AssetNoLabelChange"
        PipelineDetails.AssetNo_lbl.Caption = PipeDetailAnsi
        Call INI_WriteBack(MyFile, "PipeDetailsAssetNo=", PipeDetailAnsi) 'PCN2820
    
    Case "AssetNoTextChange"
        PipelineDetails.AssetNo.Font.Charset = LanguageCharset
        PipelineDetails.AssetNo.text = PipeDetailAnsi

    Case "SiteIDLabelChange"
        PipelineDetails.SiteID_lbl.Caption = PipeDetailAnsi
        Call INI_WriteBack(MyFile, "PipeDetailsSiteID=", PipeDetailAnsi)  'PCN2820
        
    Case "SiteIDTextChange"
        PipelineDetails.SiteID.Font.Charset = LanguageCharset
        PipelineDetails.SiteID.text = PipeDetailAnsi

    Case "CityLabelChange"
        PipelineDetails.City_lbl.Caption = PipeDetailAnsi
        Call INI_WriteBack(MyFile, "PipeDetailsCity=", PipeDetailAnsi) 'PCN2820
        
    Case "CityTextChange"
        PipelineDetails.City.Font.Charset = LanguageCharset
        PipelineDetails.City.text = PipeDetailAnsi

    Case "DateLabelChange"
        PipelineDetails.Date_lbl.Caption = PipeDetailAnsi
        Call INI_WriteBack(MyFile, "PipeDetailsDate=", PipeDetailAnsi) 'PCN2820
    
    Case "DateTextChange"
        PipelineDetails.sDate.Font.Charset = LanguageCharset
        PipelineDetails.sDate.text = PipeDetailAnsi

    Case "TimeLabelChange"
        PipelineDetails.Time_lbl.Caption = PipeDetailAnsi
        Call INI_WriteBack(MyFile, "PipeDetailsTime=", PipeDetailAnsi) 'PCN2820
    
    Case "TimeTextChange"
        PipelineDetails.sTime.Font.Charset = LanguageCharset
        PipelineDetails.sTime.text = PipeDetailAnsi

    Case "StNodeLabelChange"
        PipelineDetails.StartNodeNo_lbl.Caption = PipeDetailAnsi
        Call INI_WriteBack(MyFile, "PipeDetailsStNode=", PipeDetailAnsi) 'PCN2820
    
    Case "StNodeTextChange"
        PipelineDetails.StartNodeNo.Font.Charset = LanguageCharset
        PipelineDetails.StartNodeNo.text = PipeDetailAnsi

    Case "StNodeLocLabelChange"
        PipelineDetails.StNodeLoc_lbl.Caption = PipeDetailAnsi
        Call INI_WriteBack(MyFile, "PipeDetailsStLoc=", PipeDetailAnsi) 'PCN2820
    
    Case "StNodeLocTextChange"
        PipelineDetails.StartNodeLocation.Font.Charset = LanguageCharset
        PipelineDetails.StartNodeLocation.text = PipeDetailAnsi

    Case "FhNodeLabelChange"
        PipelineDetails.FinishNodeNo_lbl.Caption = PipeDetailAnsi
        Call INI_WriteBack(MyFile, "PipeDetailsFhNode=", PipeDetailAnsi) 'PCN2820
    
    Case "FhNodeTextChange"
        PipelineDetails.FinishNodeNo.Font.Charset = LanguageCharset
        PipelineDetails.FinishNodeNo.text = PipeDetailAnsi

    Case "FhNodeLocLabelChange"
        PipelineDetails.FhNodeLoc_lbl.Caption = PipeDetailAnsi
        Call INI_WriteBack(MyFile, "PipeDetailsFhLoc=", PipeDetailAnsi) 'PCN2820
    
    Case "FhNodeLocTextChange"
        PipelineDetails.FinishNodeLocation.Font.Charset = LanguageCharset
        PipelineDetails.FinishNodeLocation.text = PipeDetailAnsi

    Case "IntDiaLabelChange"
        PipelineDetails.InternalDiameterExpected_lbl.Caption = PipeDetailAnsi
        Call INI_WriteBack(MyFile, "PipeDetailsIntDiaExp=", PipeDetailAnsi) 'PCN2820

    Case "OutDiaLabelChange"
        PipelineDetails.OutsideDiameter_lbl.Caption = PipeDetailAnsi
        Call INI_WriteBack(MyFile, "PipeDetailsOutDiaExp=", PipeDetailAnsi) 'PCN2820

    Case "PipeLenLabelChange"
        PipelineDetails.PipeLen_lbl.Caption = PipeDetailAnsi
        Call INI_WriteBack(MyFile, "PipeDetailsLength=", PipeDetailAnsi) 'PCN2820

    Case "MaterialLabelChange"
        PipelineDetails.Material_lbl.Caption = PipeDetailAnsi
        Call INI_WriteBack(MyFile, "PipeDetailsMaterial=", PipeDetailAnsi) 'PCN2820

    Case "MaterialTextChange"
        PipelineDetails.Material.Font.Charset = LanguageCharset
        PipelineDetails.Material.text = PipeDetailAnsi
        
End Select


Exit Sub
Err_Handler:
   MsgBox Err & "-DE20:" & Error$
End Sub

Private Sub DataEntryBox_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Handler
    Select Case KeyAscii
        Case 27: Call CloseDataEntry
        Case vbKeyReturn: Call CloseDataEntry
    End Select
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-DE21:" & Error$
    End Select
End Sub

Private Sub RichDataEntryBox_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Handler
    Select Case KeyAscii
        Case 27:
            If LanguageCharset <> 0 Then
                PipeDetailAnsi = RichDataEntryBox.text
            End If
            Call CloseDataEntry
        Case vbKeyReturn
            If LanguageCharset <> 0 Then
                PipeDetailAnsi = LanguageUtil.ConvertRichToAnsi(RichDataEntryBox.TextRTF)
            End If
            Call CloseDataEntry
    End Select
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-DE21:" & Error$
    End Select
End Sub


Private Sub NewCalibration_KeyPress(KeyAscii As Integer)
Call DataEntryBox_KeyPress(KeyAscii)
End Sub

Private Sub NewDistance_KeyPress(KeyAscii As Integer)
Call DataEntryBox_KeyPress(KeyAscii)
End Sub



Private Sub NewLabelPipeDetail_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Handler

Call RichDataEntryBox_KeyPress(KeyAscii)

Exit Sub
Err_Handler:
    MsgBox Err & "-DE22:" & Error$
End Sub


Private Sub NewLimitLine_KeyPress(KeyAscii As Integer)
Call DataEntryBox_KeyPress(KeyAscii)
End Sub

Private Sub NewStartEndDistance_KeyPress(KeyAscii As Integer)
Call DataEntryBox_KeyPress(KeyAscii)
End Sub

Sub SetUpStartFinishDistances(DistPos As String)
On Error GoTo Err_Handler
        Me.Visible = False
Select Case DistPos
    Case "Start"
        Call ClearLineScreen.GotoStartMarker
        Load DataEntryForm
'        DataEntryForm.Hide
        Call DataEntryForm.SetDataEntryType("DistanceStartChange")
        DataEntryForm.Left = ClearLineScreen.VideoRecordMarkerStartAdjuster.Left
        DataEntryForm.Top = ControlsScreen.Top - (DataEntryForm.height + 600)
    Case "Finish"
        Call ClearLineScreen.GotoStopMarker
        Load DataEntryForm
'        DataEntryForm.Hide
        Call DataEntryForm.SetDataEntryType("DistanceEndChange")
        DataEntryForm.Left = ClearLineScreen.VideoRecordMarkerStopAdjuster.Left
        DataEntryForm.Top = ControlsScreen.Top - (DataEntryForm.height + 600)
End Select
DataEntryForm.Show
NewStartEndDistance.SetFocus

Exit Sub
Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & "-DE23:" & Error$
End Select
End Sub


