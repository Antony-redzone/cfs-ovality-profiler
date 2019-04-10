VERSION 5.00
Begin VB.Form ControlsReference 
   BorderStyle     =   0  'None
   ClientHeight    =   4815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1095
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4815
   ScaleWidth      =   1095
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer ExpandTimer 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   1200
      Top             =   120
   End
   Begin VB.CommandButton BtnShapeType 
      BackColor       =   &H00C6C7C6&
      Caption         =   "Type"
      Height          =   975
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton BtnReferenceCircle 
      BackColor       =   &H00C6C7C6&
      Caption         =   "Reference"
      Height          =   975
      Left            =   0
      Picture         =   "ControlsReference.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton BtnOutsideDia 
      BackColor       =   &H00C6C7C6&
      Caption         =   "Outside Dia"
      Height          =   975
      Left            =   0
      Picture         =   "ControlsReference.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton BtnInternalDia 
      BackColor       =   &H00C6C7C6&
      Caption         =   "Internal Dia"
      Height          =   975
      Left            =   0
      Picture         =   "ControlsReference.frx":3994
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton BtnFlipShape 
      BackColor       =   &H00C6C7C6&
      Caption         =   "Flip"
      Height          =   975
      Left            =   0
      Picture         =   "ControlsReference.frx":565E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   1095
   End
End
Attribute VB_Name = "ControlsReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const FormHeight As Single = 4815
Const FormWidth As Single = 1095
Dim Action As String


Private Sub BtnFlipShape_Click()
    Call FlipShape(GetNumShapeType(DrawShapeType))
    Call ScreenDrawing.DrawProfilesStartToFinish(ClearLineScreen.PVScreen, True)
End Sub

Private Sub BtnInternalDia_Click()
        Call Toggle
        ScreenDrawing.ShowReferenceShape = Not ScreenDrawing.ShowReferenceShape
        Call ScreenDrawing.DrawProfilesStartToFinish(ClearLineScreen.PVScreen, True)
End Sub

Private Sub BtnShapeType_Click()
    ControlsShape.Toggle
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

    Set BtnShapeType.Picture = ControlsShape.PicShape(0).Picture
    BtnShapeType.Tag = ControlsShape.ShapeName(0).Caption

    Me.Left = 7920
    Me.Top = 4040
    Me.height = FormHeight
    Me.width = FormWidth
    
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox error$ & " - " & Err
    End Select
End Sub

Public Sub Toggle()
On Error GoTo Err_Handler

    Me.height = FormHeight
    Me.width = FormWidth
    If Me.Visible Then
        Me.Visible = False
        ControlsShape.Visible = False
    Else
        Me.Visible = True
        Me.ZOrder 0
    End If
        
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox error$ & " - " & Err
    End Select
End Sub
