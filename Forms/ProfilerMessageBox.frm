VERSION 5.00
Begin VB.Form ProfilerMessageBox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Profiler Message"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   Icon            =   "ProfilerMessageBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton PMBNo 
      Caption         =   "NO"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton PMBYes 
      Caption         =   "Yes"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton ProfilerMsgBoxBtn 
      Caption         =   "OK"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label ProfilerMsgBoxLbl 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "ProfilerMessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Me.ProfilerMsgBoxLbl.Font.Charset = LanguageCharset
End Sub

Private Sub ProfilerMsgBoxBtn_Click()
On Error GoTo Err_Handler
    
    Unload Me

Exit Sub
Err_Handler:
    MsgBox Err & "-PM2:" & Error$
End Sub

Public Sub MsgBoxYesNo(ByVal Message As String)
On Error GoTo Err_Handler
    
    Me.ProfilerMsgBoxBtn.Visible = False
    Me.PMBYes.Visible = True
    Me.PMBNo.Visible = True

    Me.ProfilerMsgBoxLbl.Caption = Message
    Me.Show vbModal
    Me.ZOrder 0
    
Exit Sub
Err_Handler:
    MsgBox Err & "-PM3:" & Error$
End Sub

Private Sub PMBNo_Click()
On Error GoTo Err_Handler
    
    PMBAnswer = vbNo
    Unload Me
    
Exit Sub
Err_Handler:
    MsgBox Err & "-PM4:" & Error$
End Sub

Private Sub PMBYes_Click()
On Error GoTo Err_Handler
    
    PMBAnswer = vbYes
    Unload Me

Exit Sub
Err_Handler:
    MsgBox Err & "-PM5:" & Error$
End Sub

