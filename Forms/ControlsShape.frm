VERSION 5.00
Begin VB.Form ControlsShape 
   BorderStyle     =   0  'None
   ClientHeight    =   1470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1470
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox PicShape 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C6C7C6&
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   3
      Left            =   3240
      Picture         =   "ControlsShape.frx":0000
      ScaleHeight     =   945
      ScaleWidth      =   1065
      TabIndex        =   6
      Top             =   0
      Width           =   1095
      Begin VB.Label ShapeName 
         Alignment       =   2  'Center
         BackColor       =   &H00C6C7C6&
         Caption         =   "Egg"
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.PictureBox PicShape 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C6C7C6&
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   2
      Left            =   2160
      Picture         =   "ControlsShape.frx":1CCA
      ScaleHeight     =   945
      ScaleWidth      =   1065
      TabIndex        =   2
      Top             =   0
      Width           =   1095
      Begin VB.Label ShapeName 
         Alignment       =   2  'Center
         BackColor       =   &H00C6C7C6&
         Caption         =   "Egg"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.PictureBox PicShape 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C6C7C6&
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   1
      Left            =   1080
      Picture         =   "ControlsShape.frx":3994
      ScaleHeight     =   945
      ScaleWidth      =   1065
      TabIndex        =   1
      Top             =   0
      Width           =   1095
      Begin VB.Label ShapeName 
         Alignment       =   2  'Center
         BackColor       =   &H00C6C7C6&
         Caption         =   "Semieliptical"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.PictureBox PicShape 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C6C7C6&
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   0
      Left            =   0
      Picture         =   "ControlsShape.frx":565E
      ScaleHeight     =   945
      ScaleWidth      =   1065
      TabIndex        =   0
      Top             =   0
      Width           =   1095
      Begin VB.Label ShapeName 
         Alignment       =   2  'Center
         BackColor       =   &H00C6C7C6&
         Caption         =   "Circle"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
   End
End
Attribute VB_Name = "ControlsShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const FormHeight As Single = 975
Const FormWidth As Single = 4335



Private Sub Form_Load()
On Error GoTo Err_Handler

    ShapeName(0).Caption = ReferenceShape(0).name
    ShapeName(1).Caption = ReferenceShape(1).name
    ShapeName(2).Caption = ReferenceShape(2).name
    ShapeName(3).Caption = ReferenceShape(3).name
    
    Me.Left = 9115
    Me.Top = 7880
    Me.height = FormHeight
    Me.width = FormWidth
    
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox error$ & " - " & Err
    End Select
End Sub

Sub Toggle()
    If Me.Visible Then
        Me.Visible = False
    Else
        Me.Visible = True
        Me.ZOrder 0
    End If


End Sub




Private Sub PicShape_Click(Index As Integer)
    DrawShapeType = Trim(ShapeName(Index).Caption)
    ControlsReference.BtnShapeType.Caption = DrawShapeType
    Set ControlsReference.BtnShapeType.Picture = PicShape(Index).Picture
    Call ScreenDrawing.DrawProfilesStartToFinish(ClearLineScreen.PVScreen, True)
    Call ControlsReference.Toggle
'    ScreenDrawing.DrawSF = PicShape(Index).width / ClearLineScreen.MainScreen.width
'    Call ScreenDrawing.DrawReferenceShape(PicShape(Index), _
'                                          Index, _
'                                          200, _
'                                          200, _
'                                          PicShape(Index).height / 10, _
'                                          vbGreen)
End Sub
