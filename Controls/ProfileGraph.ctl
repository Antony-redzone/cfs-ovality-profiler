VERSION 5.00
Begin VB.UserControl PVDGraphControl 
   ClientHeight    =   12255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11280
   FillStyle       =   0  'Solid
   ScaleHeight     =   12255
   ScaleWidth      =   11280
   Begin VB.Line LowerPercentageLine 
      X1              =   7920
      X2              =   8040
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line TopPercentageLine 
      X1              =   7920
      X2              =   8040
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label LowerPercentage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "2%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   120
      TabIndex        =   25
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label TopPercentage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "12%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   120
      TabIndex        =   24
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line MajourDev 
      Index           =   19
      X1              =   2280
      X2              =   2280
      Y1              =   1560
      Y2              =   1920
   End
   Begin VB.Line MajourDev 
      Index           =   18
      X1              =   4320
      X2              =   4320
      Y1              =   4080
      Y2              =   4200
   End
   Begin VB.Line MajourDev 
      Index           =   17
      X1              =   1680
      X2              =   1680
      Y1              =   1440
      Y2              =   1800
   End
   Begin VB.Line MajourDev 
      Index           =   16
      X1              =   5160
      X2              =   5280
      Y1              =   3480
      Y2              =   3720
   End
   Begin VB.Line MajourDev 
      Index           =   15
      X1              =   3240
      X2              =   3120
      Y1              =   3360
      Y2              =   3600
   End
   Begin VB.Label DistanceLbl 
      Alignment       =   2  'Center
      Caption         =   "140"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   19
      Left            =   4440
      TabIndex        =   23
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label DistanceLbl 
      Alignment       =   2  'Center
      Caption         =   "140"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   18
      Left            =   3600
      TabIndex        =   22
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label DistanceLbl 
      Alignment       =   2  'Center
      Caption         =   "140"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   17
      Left            =   5280
      TabIndex        =   21
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label DistanceLbl 
      Alignment       =   2  'Center
      Caption         =   "140"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   16
      Left            =   5160
      TabIndex        =   20
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label DistanceLbl 
      Alignment       =   2  'Center
      Caption         =   "140"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   15
      Left            =   4920
      TabIndex        =   19
      Top             =   4560
      Width           =   495
   End
   Begin VB.Line MajourDev 
      Index           =   14
      X1              =   6960
      X2              =   6960
      Y1              =   1680
      Y2              =   1800
   End
   Begin VB.Line MajourDev 
      Index           =   13
      X1              =   6480
      X2              =   6480
      Y1              =   1920
      Y2              =   2040
   End
   Begin VB.Line MajourDev 
      Index           =   12
      X1              =   5520
      X2              =   5520
      Y1              =   1680
      Y2              =   1920
   End
   Begin VB.Line MajourDev 
      Index           =   11
      X1              =   4680
      X2              =   4680
      Y1              =   1920
      Y2              =   2040
   End
   Begin VB.Line MajourDev 
      Index           =   10
      X1              =   3120
      X2              =   3120
      Y1              =   1440
      Y2              =   1680
   End
   Begin VB.Line MajourDev 
      Index           =   9
      X1              =   5000
      X2              =   5000
      Y1              =   0
      Y2              =   60
   End
   Begin VB.Line MajourDev 
      Index           =   8
      X1              =   6480
      X2              =   6480
      Y1              =   1440
      Y2              =   1560
   End
   Begin VB.Label DistanceLbl 
      Alignment       =   2  'Center
      Caption         =   "140"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   14
      Left            =   4680
      TabIndex        =   18
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label DistanceLbl 
      Alignment       =   2  'Center
      Caption         =   "140"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   13
      Left            =   4320
      TabIndex        =   17
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label DistanceLbl 
      Alignment       =   2  'Center
      Caption         =   "140"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   12
      Left            =   3840
      TabIndex        =   16
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label DistanceLbl 
      Alignment       =   2  'Center
      Caption         =   "140"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   11
      Left            =   3480
      TabIndex        =   15
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label DistanceLbl 
      Alignment       =   2  'Center
      Caption         =   "140"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   10
      Left            =   3240
      TabIndex        =   14
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label DistanceLbl 
      Alignment       =   2  'Center
      Caption         =   "140"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   9
      Left            =   3000
      TabIndex        =   13
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label DistanceLbl 
      Alignment       =   2  'Center
      Caption         =   "140"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   8
      Left            =   2880
      TabIndex        =   12
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image PrinterReportImageTwo 
      Height          =   2700
      Left            =   120
      Picture         =   "ProfileGraph.ctx":0000
      Top             =   9480
      Width           =   46080
   End
   Begin VB.Image GraphContainerTwo 
      Height          =   2295
      Left            =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   10575
   End
   Begin VB.Line LineBreak 
      X1              =   0
      X2              =   11040
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line MajourDev 
      Index           =   7
      X1              =   9600
      X2              =   9600
      Y1              =   5880
      Y2              =   5940
   End
   Begin VB.Line MajourDev 
      Index           =   6
      X1              =   8640
      X2              =   8640
      Y1              =   5880
      Y2              =   5940
   End
   Begin VB.Line MajourDev 
      Index           =   5
      X1              =   7320
      X2              =   7320
      Y1              =   5880
      Y2              =   5940
   End
   Begin VB.Line MajourDev 
      Index           =   4
      X1              =   6000
      X2              =   6000
      Y1              =   5880
      Y2              =   5940
   End
   Begin VB.Line MajourDev 
      Index           =   3
      X1              =   4680
      X2              =   4680
      Y1              =   5880
      Y2              =   5940
   End
   Begin VB.Line MajourDev 
      Index           =   2
      X1              =   3360
      X2              =   3360
      Y1              =   5880
      Y2              =   5940
   End
   Begin VB.Line MajourDev 
      Index           =   1
      X1              =   2040
      X2              =   2040
      Y1              =   6000
      Y2              =   5940
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   120
   End
   Begin VB.Line MajourDev 
      Index           =   0
      X1              =   600
      X2              =   600
      Y1              =   5880
      Y2              =   5940
   End
   Begin VB.Label DistanceLbl 
      Alignment       =   2  'Center
      Caption         =   "140"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   7
      Left            =   9360
      TabIndex        =   11
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label DistanceLbl 
      Alignment       =   2  'Center
      Caption         =   "120"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   6
      Left            =   8040
      TabIndex        =   10
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label DistanceLbl 
      Alignment       =   2  'Center
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   5
      Left            =   6720
      TabIndex        =   9
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label DistanceLbl 
      Alignment       =   2  'Center
      Caption         =   "80"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   4
      Left            =   5400
      TabIndex        =   8
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label DistanceLbl 
      Alignment       =   2  'Center
      Caption         =   "60"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   3
      Left            =   4080
      TabIndex        =   7
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label DistanceLbl 
      Alignment       =   2  'Center
      Caption         =   "40"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   2
      Left            =   2760
      TabIndex        =   6
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label DistanceLbl 
      Alignment       =   2  'Center
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   1
      Left            =   1320
      TabIndex        =   5
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label DistanceLbl 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   0
      Left            =   -240
      TabIndex        =   4
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label CommentsLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Comments:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape GraphBorder 
      Height          =   2775
      Left            =   0
      Top             =   3120
      Width           =   10455
   End
   Begin VB.Label GraphComments 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   6120
      Width           =   9855
   End
   Begin VB.Label GraphTitle 
      Caption         =   "GraphTitle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2655
   End
   Begin VB.Image PrinterReportImage 
      Height          =   2700
      Left            =   120
      Picture         =   "ProfileGraph.ctx":267F
      Top             =   6600
      Width           =   46080
   End
   Begin VB.Label GraphUnitLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "m"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10680
      TabIndex        =   0
      Top             =   5640
      Width           =   255
   End
   Begin VB.Image GraphContainer 
      Appearance      =   0  'Flat
      Height          =   2775
      Left            =   0
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   10575
   End
   Begin VB.Shape GraphBackgroundShape 
      BackColor       =   &H00DDDDA2&
      FillColor       =   &H00DDDDA2&
      FillStyle       =   0  'Solid
      Height          =   6420
      Left            =   0
      Tag             =   "Back"
      Top             =   0
      Width           =   11055
   End
End
Attribute VB_Name = "PVDGraphControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Destination As Variant
Dim RS As Single
Dim LeftPosition As Single
Dim TopPosition As Single

Private StartDistance As Single
Private EndDistance As Single
Private StartFrame As Long
Private EndFrame As Long
Private HideInfo As Boolean
Private SecondGraph As Boolean
Private GraphLength As Single
Private RulerMultiplier As Single
Private GraphUnit As String
Private DiameterUnit As String



Private Sub UserControl_Initialize()
On Error GoTo Err_Handler

HideInfo = False
SecondGraph = False
GraphLength = 150
RulerMultiplier = 10
DiameterUnit = "mm"
GraphUnit = "m"
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-GC1:" & Error$
    End Select
End Sub

Private Sub UserControl_Resize()
On Error GoTo Err_Handler
    Dim LblIndex As Integer
    Dim LblbValue As Single
    Dim ShiftX As Single
    Dim GraphWidth As Single
    
    With UserControl
        .GraphBackgroundShape.width = .width
        .GraphBackgroundShape.height = .height
        
        .GraphComments.Top = .GraphBackgroundShape.height - .GraphComments.height
        .CommentsLabel.Top = .GraphComments.Top
        .GraphContainer.width = .width - 480
        For LblIndex = 0 To 19
            .DistanceLbl(LblIndex).Caption = LblIndex * RulerMultiplier
            lblvalue = .DistanceLbl(LblIndex).Caption
            'ShiftX = (lblvalue / GraphLength * .GraphContainer.width) - (.DistanceLbl(LblIndex).width / 2)
            
            ShiftX = (.GraphContainer.width / 19 * LblIndex) - (.DistanceLbl(lbindex).width / 2)
            
            .DistanceLbl(LblIndex).Left = ShiftX
            .DistanceLbl(LblIndex).Top = .GraphComments.Top - .DistanceLbl(0).height
            .MajourDev(LblIndex).x1 = ShiftX + (.DistanceLbl(LblIndex).width / 2)
            .MajourDev(LblIndex).x2 = ShiftX + (.DistanceLbl(LblIndex).width / 2)
            .MajourDev(LblIndex).y2 = .DistanceLbl(LblIndex).Top
            .MajourDev(LblIndex).y1 = .MajourDev(LblIndex).y2 - 60
        Next LblIndex
        .GraphUnitLabel.Top = .GraphComments.Top - .DistanceLbl(0).height
        .GraphUnitLabel.Left = .GraphContainer.width + 240
        
    

        If SecondGraph Then
            .GraphContainer.height = (.MajourDev(0).y1 - .GraphTitle.height) * 2 / 3
            .GraphContainerTwo.height = (.MajourDev(0).y1 - .GraphTitle.height) * 1 / 3
            .GraphContainerTwo.Top = .GraphTitle.height
        Else
            .GraphContainer.height = (.MajourDev(0).y1 - .GraphTitle.height)
        End If
        
        .GraphContainer.Top = .MajourDev(0).y1 - GraphContainer.height + 15
    
        .GraphBorder.width = .GraphContainer.width
        .GraphBorder.height = .GraphContainer.height
        .GraphBorder.Top = .GraphContainer.Top
        .LineBreak.x1 = 0: .LineBreak.x2 = .GraphContainer.width
        .LineBreak.y1 = .height - 30: .LineBreak.y2 = .height - 30
        
        
        GraphWidth = ((EndDistance - StartDistance) / GraphLength) * UserControl.GraphContainer.width
        
        .TopPercentage.Left = GraphWidth + 75
        .LowerPercentage.Left = GraphWidth + 75
        
        .TopPercentage.Top = .GraphContainerTwo.Top - (.TopPercentage.height / 2)
        .LowerPercentage.Top = .GraphContainerTwo.Top + .GraphContainerTwo.height - (.TopPercentage.height / 2)
        
        .TopPercentageLine.y1 = .GraphContainerTwo.Top - 15
        .TopPercentageLine.y2 = .TopPercentageLine.y1
        
        .LowerPercentageLine.y1 = .GraphContainer.Top - 15
        .LowerPercentageLine.y2 = .LowerPercentageLine.y1
        
        .TopPercentageLine.x1 = 0
        .TopPercentageLine.x2 = GraphWidth + 60
        
        .LowerPercentageLine.x1 = 0
        .LowerPercentageLine.x2 = GraphWidth + 60
        
    End With
       
    

    

    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-GC2:" & Error$
    End Select
End Sub


Public Sub PrintGraph(ToWhere As Variant, ByVal RenScale As Single, _
                      ByVal Left As Single, ByVal Top As Single)
On Error GoTo Err_Handler


Dim LabelIndex As Integer

If StartFrame = EndFrame Then Exit Sub 'If they are the same dont continue, waste of time
                                       'and causes errors.
Set Destination = ToWhere
RS = RenScale

LeftPosition = Left
TopPosition = Top

If Not HideInfo Then
    Call RenderSingleLabel(UserControl.GraphTitle)
    Call RenderSingleLabel(UserControl.GraphComments)
    Call RenderSingleLabel(UserControl.CommentsLabel)

End If

Call RenderShape(UserControl.GraphBorder)
If Abs(EndDistance - StartDistance) < GraphLength Then
    Call RenderLine(UserControl.LineBreak)
End If
For LabelIndex = 0 To 19
    'UserControl.DistanceLbl(LabelIndex).Caption = (LabelIndex * RulerMultiplier) + StartDistance
    UserControl.DistanceLbl(LabelIndex).Caption = Format((GraphLength / 19 * LabelIndex) + StartDistance, "#0.0")
    If CSng(UserControl.DistanceLbl(LabelIndex).Caption) < (EndDistance + (GraphLength / 19)) Then
        Call RenderSingleLabel(UserControl.DistanceLbl(LabelIndex))
        Call RenderLine(UserControl.MajourDev(LabelIndex))
        UserControl.GraphUnitLabel.Left = UserControl.DistanceLbl(LabelIndex).Left + 480
    End If
   
   
Next LabelIndex
    Call RenderSingleLabel(UserControl.GraphUnitLabel)


'On Error GoTo Skip_DrawPVGraphsReport
'If ToWhere.name <> "Printer" Then Call DrawPVGraphsReport
'Skip_DrawPVGraphsReport:
'On Error GoTo Err_Handler



Call RenderTheGraph

If SecondGraph Then
        Call RenderSingleLabel(UserControl.TopPercentage)
        Call RenderSingleLabel(UserControl.LowerPercentage)
        Call RenderLine(UserControl.TopPercentageLine)
        Call RenderLine(UserControl.LowerPercentageLine)
End If

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-GC3:" & Error$
    End Select
End Sub

Private Sub RenderSingleLabel(Lbl As Label)
'****************************************************************************************
'Name    : RenderSingleLabel
'Created : August 9 05, PCN3593
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    :
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler

Dim XOffset As Single
Dim YOffset As Single

XOffset = LeftPosition
YOffset = TopPosition

Destination.Font = Lbl.Font
Destination.FontName = Lbl.FontName
Destination.Font.Size = (Lbl.Font.Size) * RS
Destination.Font.Italic = Lbl.Font.Italic
Destination.Font.Bold = Lbl.Font.Bold
Destination.ForeColor = Lbl.ForeColor

If Lbl.Alignment = 0 Then Destination.CurrentX = (Lbl.Left + XOffset) * RS
If Lbl.Alignment = 1 Then Destination.CurrentX = (Lbl.Left + Lbl.width - Destination.TextWidth(Lbl.Caption) + XOffset) * RS
If Lbl.Alignment = 2 Then Destination.CurrentX = (Lbl.Left + (Lbl.width / 2) - (Destination.TextWidth(Lbl.Caption) / 2) + XOffset) * RS
Destination.CurrentY = (Lbl.Top + YOffset) * RS

If Lbl.WordWrap = True And Destination.TextWidth(Lbl.Caption) > (Lbl.width - 75) Then 'PCN4389
    Call FormatTextForLabel(Lbl)
Else
    If Destination.CurrentX < Destination.width And Destination.CurrentY < Destination.height - Destination.TextHeight(Lbl.Caption) Then
        Destination.Print Lbl.Caption
    End If
End If

UserControl.GraphBorder.width = ((EndDistance - StartDistance) / GraphLength) * UserControl.GraphContainer.width

Exit Sub
Err_Handler:
MsgBox Err & "-GC4:" & Error$

End Sub

Private Sub RenderShape(DrawShape As Shape)
'****************************************************************************************
'PCN:
'Name    : RenderShape
'Created : Feb 15 2006
'Updated :
'Prg By  : Antony van Iersel
'Param   :
'Desc    : draws shape as a filled in line bax
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler

Dim CurrentStyle As Long
Dim currentwidth As Long
Dim FillColour As Long
Dim FillStyle As Long

Dim x1, x2, y1, y2 As Single

Dim XOffset As Single
Dim YOffset As Single

XOffset = LeftPosition
YOffset = TopPosition
    
    CurrentStyle = Destination.DrawStyle
    currentwidth = Destination.DrawWidth
    FillStyle = Destination.FillStyle
    FillColour = Destination.FillColor
    

If DrawShape.BorderStyle = 3 Then
    Destination.DrawStyle = vbDot
    Destination.DrawWidth = 1
End If
    Destination.DrawWidth = 1
With DrawShape
    x1 = .Left + XOffset
    y1 = .Top + YOffset
    x2 = .Left + .width + XOffset
    y2 = .Top + .height + YOffset
End With

If DrawShape.FillStyle = vbSolid Then
    Destination.FillStyle = DrawShape.FillStyle
    Destination.FillColor = DrawShape.FillColor
    Destination.Line (x1 * RS, y1 * RS)-(x2 * RS, y2 * RS), DrawShape.BorderColor, B
Else
    Destination.Line (x1 * RS, y1 * RS)-(x2 * RS, y2 * RS), DrawShape.BorderColor, B
End If

Destination.DrawStyle = CurrentStyle
Destination.DrawWidth = currentwidth
Destination.FillStyle = FillStyle
Destination.FillColor = FillColour





Exit Sub
Err_Handler:
MsgBox Err & "-GC5:" & Error$
    
End Sub

Private Sub FormatTextForLabel(ByRef Lbl As Label)
'****************************************************************************************
'Name    : RenderSingleTextBox
'Created : August 9 05, PCN3593
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    :
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
Dim TotalText As String
Dim OneLine As String
Dim Count As Integer
Dim TotalHeight As Long

TotalText = Lbl.Caption
Destination.CurrentY = (Lbl.Top) * RS
TotalHeight = 0

While TotalHeight < (Lbl.height - 50)
    Destination.CurrentX = (Lbl.Left + 10) * RS
    OneLine = ParseOneLabelWidth(Lbl.width - 20, TotalText)
    TotalHeight = TotalHeight + Destination.TextHeight(Remark)
    If Lbl.Alignment = 1 Then
        Destination.CurrentX = (Lbl.Left + Lbl.width - 10 - Destination.TextWidth(OneLine)) * RS
    End If
    If Destination.CurrentX < Destination.width And Destination.CurrentY < Destination.height - Destination.TextHeight(OneLine) Then
        If Destination.CurrentX < Destination.width And Destination.CurrentY < Destination.height - Destination.TextHeight(Lbl.Caption) Then
            Destination.Print OneLine
        End If
        
    End If
Wend

Exit Sub
Err_Handler:
MsgBox Err & "-GC6:" & Error$
End Sub

Private Function ParseOneLabelWidth(ByRef LabelWidth As Long, ByRef Remark As String) As String
'****************************************************************************************
'Name    : RenderSingleLabel
'Created : August 9 05, PCN3593
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    :
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
Dim Word As String

If Destination.TextWidth(Remark) < LabelWidth Then
    ParseOneLabelWidth = Remark
    Remark = ""
    Exit Function
End If

ParseOneLabelWidth = ""
Word = Left(Remark, InStr(Remark, " "))

While Destination.TextWidth(ParseOneLabelWidth & Word) < LabelWidth
    If Word = "" Then
        ParseOneLabelWidth = ParseOneLabelWidth & Word
        Remark = ""
        Exit Function
    End If
    ParseOneLabelWidth = ParseOneLabelWidth & Word
    Remark = Right(Remark, Len(Remark) - Len(Word))
    Word = Left(Remark, InStr(Remark, " "))
    If Word = "" And Remark <> "" Then Word = Remark
Wend

If Destination.TextWidth(ParseOneLabelWidth & Word) < LabelWidth Then ParseOneLabelWidth = ParseOneLabelWidth & Word

Exit Function
Err_Handler:
MsgBox Err & "-GC7:" & Error$
End Function

Public Sub DrawPVGraphsReport()
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
    Dim GraphWidth As Single
    Dim StoreWidth As Single
    Dim StoreHeight As Single
    

    
    SaveScreenDrawingType = ScreenDrawingType
    ScreenDrawingType = 0
    ScreenDrawingOrientation = 1


    GraphWidth = ((EndDistance - StartDistance) / GraphLength) * UserControl.GraphContainer.width
    
    StoreWidth = ImageGraphState(6).PictureImage.width
    StoreHeight = ImageGraphState(6).PictureImage.height
    
    
    
    ImageGraphState(6).PictureImage.width = GraphWidth
    UserControl.PrinterReportImageTwo.width = GraphWidth
    
    ImageGraphState(6).PictureImage.height = UserControl.GraphContainer.height * 10
    UserControl.PrinterReportImageTwo.height = UserControl.GraphContainer.height * 10
    


    Set ImageGraphState(6).PictureImage = PrecisionVisionGraph.PrinterReportImage
    Call DrawGraphImage(ImageGraphState(6), "Clear", 0, 0, 0, 0, 0, 0, 0)
    
    If StartFrame < EndFrame Then
        If EndFrame <= 1 Then EndFrame = PVDataNoOfLines
        Call DrawGraphImage(ImageGraphState(6), ImageGraphState(0).GraphType, 0, StartFrame, EndFrame, PVGraphGeneralXScale, PVGraphGeneralXOffset, PVXScaleLimitPerL, PVXScaleLimitPerR)
        Call CopyPicture(PrecisionVisionGraph.PrinterReportImage.Picture, UserControl.PrinterReportImage.Picture)
        If SecondGraph = True Then
            Call DrawGraphImage(ImageGraphState(6), "OvalityBar", 0, StartFrame, EndFrame, PVGraphGeneralXScale, PVGraphGeneralXOffset, PVXScaleLimitPerL, PVXScaleLimitPerR)
            'Call DrawGraphImage(ImageGraphState(6), "Ovality", 0, StartFrame, EndFrame, PVGraphGeneralXScale, PVGraphGeneralXOffset, PVXScaleLimitPerL, PVXScaleLimitPerR)
            Call CopyPicture(PrecisionVisionGraph.PrinterReportImage.Picture, UserControl.PrinterReportImageTwo.Picture)
        End If
        
    Else
        If StartFrame <= 1 Then StartFrame = PVDataNoOfLines
        Call DrawGraphImage(ImageGraphState(6), ImageGraphState(0).GraphType, 0, EndFrame, StartFrame, PVGraphGeneralXScale, PVGraphGeneralXOffset, PVXScaleLimitPerL, PVXScaleLimitPerR)
        Call CopyPicture(PrecisionVisionGraph.PrinterReportImage.Picture, UserControl.PrinterReportImage.Picture)
        If SecondGraph = True Then
            Call DrawGraphImage(ImageGraphState(6), "OvalityBar", 0, EndFrame, StartFrame, PVGraphGeneralXScale, PVGraphGeneralXOffset, PVXScaleLimitPerL, PVXScaleLimitPerR)
            'Call DrawGraphImage(ImageGraphState(6), "Ovality", 0, EndFrame, StartFrame, PVGraphGeneralXScale, PVGraphGeneralXOffset, PVXScaleLimitPerL, PVXScaleLimitPerR)
            Call CopyPicture(PrecisionVisionGraph.PrinterReportImage.Picture, UserControl.PrinterReportImageTwo.Picture)
        End If
    End If
    
    ScreenDrawingType = SaveScreenDrawingType
    ImageGraphState(6).PictureImage.width = StoreWidth
    ImageGraphState(6).PictureImage.height = StoreHeight
    
    
    'Call CopyPicture(PrecisionVisionGraph.PrinterReportImage.Picture, UserControl.PrinterReportImage.Picture)
'    Call CopyPicture(PrecisionVisionGraph.PrinterReportImage.Picture, UserControl.PrinterReportPicture.Picture)
    
    
Exit Sub
Err_Handler:
    MsgBox Err & "-GC8:" & Error$

End Sub

Private Sub RenderImages(Img As Image)
'****************************************************************************************
'Name    : RenderImages
'Created : August 10 05, PCN3593
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    : draws images, except for the pipe and manholes
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler

    If Img.Picture = 0 Then Exit Sub
    If Img.Tag <> "Visible" Then Exit Sub
    
    Dim XOffset As Single
    Dim YOffset As Single

    XOffset = LeftPosition
    YOffset = TopPosition

    Call Destination.PaintPicture(Img.Picture, (Img.Left + XOffset) * RS, (Img.Top + YOffset) * RS, (Img.width) * RS, (Img.height) * RS)
    If Img.BorderStyle = 1 Then Destination.Line ((Img.Left) * RS, (Img.Top) * RS)-((Img.Left + Img.width) * RS, (Img.Top + Img.height) * RS), vbBlack, B

Exit Sub
Err_Handler:
MsgBox Err & "-GC9:" & Error$
End Sub

Sub RenderTheGraph()
On Error GoTo Err_Handler

    Dim XOffset As Single
    Dim YOffset As Single
    Dim GraphWidth As Single
    Dim XOffsetTwo As Single
    Dim YOffsetTwo As Single
    
    

    XOffset = LeftPosition + UserControl.GraphContainer.Left
    YOffset = TopPosition + UserControl.GraphContainer.Top

    XOffsetTwo = LeftPosition + UserControl.GraphContainerTwo.Left
    YOffsetTwo = TopPosition + UserControl.GraphContainerTwo.Top

    
    GraphWidth = ((EndDistance - StartDistance) / GraphLength) * UserControl.GraphContainer.width
    If StartFrame < EndFrame Then
        Call Destination.PaintPicture(UserControl.PrinterReportImage.Picture, _
                                      XOffset, YOffset, _
                                      GraphWidth, _
                                      UserControl.GraphContainer.height)
        If SecondGraph = True Then
            Call Destination.PaintPicture(UserControl.PrinterReportImageTwo.Picture, _
                                      XOffsetTwo, YOffsetTwo, _
                                      GraphWidth, _
                                      UserControl.GraphContainerTwo.height)

        
        
        End If
        
    Else
        XOffset = XOffset + GraphWidth
        XOffsetTwo = XOffsetTwo + GraphWidth
        GraphWidth = GraphWidth * -1
        Call Destination.PaintPicture(UserControl.PrinterReportImage.Picture, _
                              XOffset, YOffset, _
                              GraphWidth, _
                              UserControl.GraphContainer.height)
        If SecondGraph = True Then
            Call Destination.PaintPicture(UserControl.PrinterReportImageTwo.Picture, _
                                      XOffsetTwo, YOffsetTwo, _
                                      GraphWidth, _
                                      UserControl.GraphContainerTwo.height)
        End If
    End If



Exit Sub
Err_Handler:
MsgBox Err & "-GC10:" & Error$

End Sub

Property Let SetStartDistance(ByVal Dist As Double)
On Error GoTo Err_Handler

    StartDistance = Dist
    StartFrame = PrecisionVisionGraph.GetFrameFromDistance(StartDistance)
    

Exit Property
Err_Handler:
MsgBox Err & "-GC11:" & Error$
End Property

Property Let SetEndDistance(ByVal Dist As Double)
On Error GoTo Err_Handler

    EndDistance = Dist
    EndFrame = PrecisionVisionGraph.GetFrameFromDistance(EndDistance)
    EndDistance = PVDistances(EndFrame)

Exit Property
Err_Handler:
MsgBox Err & "-GC12:" & Error$
End Property

Property Let SetCommentCaption(ByVal CommentCaption As String)
On Error GoTo Err_Handler

    UserControl.CommentsLabel.Caption = CommentCaption & ":"


Exit Property
Err_Handler:
MsgBox Err & "-GC12.5:" & Error$
End Property


Property Let SetComment(ByVal CommentString As String)
On Error GoTo Err_Handler
    Dim Char13 As Integer

    UserControl.GraphComments = Trim(CommentString)
    Do
        Char13 = InStr(UserControl.GraphComments, Chr(13))
        If Char13 = 0 Then Exit Do
        UserControl.GraphComments = Left(UserControl.GraphComments, Char13 - 1) & _
        " " & _
        Right(UserControl.GraphComments, Len(UserControl.GraphComments) - Char13)
    Loop
        
        Do
        Char13 = InStr(UserControl.GraphComments, Chr(10))
        If Char13 = 0 Then Exit Do
        UserControl.GraphComments = Left(UserControl.GraphComments, Char13 - 1) & _
        " " & _
        Right(UserControl.GraphComments, Len(UserControl.GraphComments) - Char13)
    Loop

Exit Property
Err_Handler:
MsgBox Err & "-GC13:" & Error$
End Property

Property Let SetGraphTitle(ByVal CommentString As String)
On Error GoTo Err_Handler

    UserControl.GraphTitle = CommentString
    

Exit Property
Err_Handler:
MsgBox Err & "-GC14:" & Error$
End Property


Property Let SetHideInfo(ByVal Hide As Boolean)
On Error GoTo Err_Handler

    HideInfo = Hide

Exit Property
Err_Handler:
MsgBox Err & "-GC15:" & Error$
End Property

Property Let SetSecondGraphSate(ByVal SecondGraphOn As Boolean)
On Error GoTo Err_Handler

    SecondGraph = SecondGraphOn
    RulerMultiplier = 5

    

Exit Property
Err_Handler:
MsgBox Err & "-GC16:" & Error$
End Property

Property Let SetGraphLength(ByVal TheGraphLength As Single)
On Error GoTo Err_Handler

    GraphLength = TheGraphLength
    If GraphLength = 500 Then
        RulerMultiplier = 50
        
    End If

Exit Property
Err_Handler:
MsgBox Err & "-GC17:" & Error$
End Property


Property Get GetStartFrame() As Long
On Error GoTo Err_Handler

    GetStartFrame = StartDistance

Exit Property
Err_Handler:
MsgBox Err & "-GC18:" & Error$
End Property


Property Get GetEndFrame() As Long
On Error GoTo Err_Handler

    GetEndFrame = EndFrame

Exit Property
Err_Handler:
MsgBox Err & "-GC19:" & Error$
End Property

Property Get GetStartDistance() As Double
On Error GoTo Err_Handler

    GetStartDistance = StartDistance

Exit Property
Err_Handler:
MsgBox Err & "-GC20:" & Error$
End Property


Property Get GetEndDistance() As Long
On Error GoTo Err_Handler

    GetEndDistance = EndDistance

Exit Property
Err_Handler:
MsgBox Err & "-GC21:" & Error$
End Property

Private Sub RenderLine(DrawLine As line)
'****************************************************************************************
'PCN:
'Name    : RenderLines
'Created : January 16 2006
'Updated :
'Prg By  : Antony van Iersel
'Param   :
'Desc    : draws lines, offset by possible picture box container
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler

Dim CurrentStyle As Long
Dim currentwidth As Long

Dim XOffset As Single
Dim YOffset As Single

XOffset = LeftPosition
YOffset = TopPosition

    
    CurrentStyle = Destination.DrawStyle
    currentwidth = Destination.DrawWidth

If DrawLine.BorderStyle = 3 Then
    Destination.DrawStyle = vbDot
    Destination.DrawWidth = 1
End If
Destination.DrawWidth = 1
Destination.Line ((DrawLine.x1 + XOffset) * RS, (DrawLine.y1 + YOffset) * RS)-((DrawLine.x2 + XOffset) * RS, (DrawLine.y2 + YOffset) * RS), DrawLine.BorderColor

Destination.DrawStyle = CurrentStyle
Destination.DrawWidth = currentwidth


Exit Sub
Err_Handler:
MsgBox Err & "-GC22:" & Error$

End Sub


Property Let SetGraphUnit(ByVal unit As String)
On Error GoTo Err_Handler

    GraphUnit = unit
    GraphUnitLabel.Caption = unit
    rullermultiplier = 1
    If GraphUnit = "m" Then
        DiameterUnit = "mm"
    Else
        DiamterUnit = "in"
    End If

Exit Property
Err_Handler:
MsgBox Err & "-GC23:" & Error$
End Property
