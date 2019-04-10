VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Registration 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Registration"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12975
   ControlBox      =   0   'False
   Icon            =   "Registration.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   12975
   Begin VB.TextBox RegCode 
      Height          =   315
      Left            =   3960
      TabIndex        =   15
      Top             =   4725
      Width           =   6450
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Cancel"
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   975
      Left            =   150
      TabIndex        =   4
      Top             =   3720
      Width           =   2055
      Begin VB.OptionButton IDecline 
         BackColor       =   &H00C0C0C0&
         Caption         =   "I Decline"
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   540
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton Agree 
         BackColor       =   &H00C0C0C0&
         Caption         =   "I Agree"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   180
         Width           =   1695
      End
   End
   Begin VB.TextBox ProductNo 
      Height          =   315
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3780
      Width           =   4830
   End
   Begin VB.TextBox RegUserName 
      Height          =   315
      Left            =   3960
      TabIndex        =   2
      Top             =   4320
      Width           =   4830
   End
   Begin VB.CommandButton Register 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Register"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10605
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4770
      Width           =   2100
   End
   Begin VB.CommandButton EMail 
      Caption         =   "EMail Us"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8895
      TabIndex        =   0
      Top             =   3810
      Width           =   1515
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3195
      Left            =   150
      TabIndex        =   12
      Top             =   450
      Width           =   12675
      _ExtentX        =   22357
      _ExtentY        =   5636
      _Version        =   393217
      ScrollBars      =   2
      FileName        =   "C:\Cleanflow\EULA.rtf"
      TextRTF         =   $"Registration.frx":0CCA
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   12285
      Top             =   -135
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   11610
      Top             =   -135
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Reg Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   16
      Top             =   4785
      Width           =   1635
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0C0C0&
      Caption         =   "End User License Agreement"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   13
      Top             =   120
      Width           =   10155
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Product No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   3840
      Width           =   1635
   End
   Begin VB.Label UserName_lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   4380
      Width           =   1635
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "OR Contact Us at :"
      Height          =   255
      Left            =   10470
      TabIndex        =   9
      Top             =   3810
      Width           =   2355
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Phone : 64 9 479 9901"
      Height          =   225
      Left            =   10470
      TabIndex        =   8
      Top             =   4050
      Width           =   2355
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fax : 64 9 479 9904"
      Height          =   225
      Left            =   10470
      TabIndex        =   7
      Top             =   4290
      Width           =   2355
   End
End
Attribute VB_Name = "Registration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cancel_Click()
On Error GoTo Err_Handler

Unload Me
If IDecline.value = 1 Or _
    Registered = False Then 'ML310103
'   ClearLineProfilerV6.FabLock1.IsRegistered = False Then
'GL231102    If isopen("Splash4MT2") Then Unload Splash4MT2
    If IsOpen("ClearLineProfilerV6") Then Unload ClearLineProfilerV6 'PCN4171
Else
    ClearLineProfilerV6.Show 'PCN4171
    Call LoadCoreForms 'PCNGL171202
End If
Unload Registration

Exit Sub
Err_Handler:
    Select Case Err
        Case Else
            MsgBox Err & "-R1:" & Error$
            Resume Next
    End Select
End Sub

Private Sub EMail_Click()
On Error GoTo Err_Handler

Dim Address As String, SendStatus As Boolean

If Not Agree Then
'MsgBox DisplayMessage("You must Agree to the End User License Agreement Details before Registering this Product."), vbExclamation 'PCN2111
ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("You must Agree to the End User License Agreement Details before Registering this Product."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
Exit Sub
End If

Address = "Products@cbsys.co.nz"
Call RegisterProduct(Address, RegUserName, ProductNo, SendStatus)
'If SendStatus Then   'commented by ls 7/1/03
'    'Save ProductNo
'    Dim i As Integer
'    Open App.Path & "/temp.cbs" For Output As #1
'    For i = Len(ProductNo) To 1 Step -1
'        Print #1, Asc(Mid(ProductNo, i, 1)) + Len(ProductNo) - i + 1
'    Next i
'    Close #1
'End If

If SendStatus = False Then
  'MsgBox DisplayMessage("Email not Sent. Please check your EMail Settings."), vbCritical 'PCN2111
  ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Email not Sent. Please check your EMail Settings."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
Else
  'MsgBox DisplayMessage("Registration Request Sent."), vbInformation 'PCN2111
  ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Registration Request Sent."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
  EMail.Enabled = False
End If

Exit Sub
Err_Handler:
MsgBox Err & "-R2:" & Error$

End Sub



Private Sub Form_Load()

On Error GoTo Err_Handler
ConvertLanguage Me, Language 'PCN2111

Registration.height = 5790 'PCNGL300503-2
Registration.width = 13095 'PCNGL300503-2




Dim ArmClass As CArmadillo
Dim RegType As String

Set ArmClass = New CArmadillo

ProductNo = ArmClass.HardwareFingerPrint
RegUserName = ArmClass.UserName
RegCode = ArmClass.UserKey
RegType = ArmClass.ClearLineRegType

If Registered = False Then 'ML310103
'    If ClearLineProfilerV5.FabLock1.IsRegistered = False Then
    EMail.Enabled = True
    Frame4.Visible = True
    Register.Visible = True
ElseIf RegType = "Registered" Or UCase(RegType) = "SONAR" Then 'This is standard registration
    'Enable registration of the 3D module
    EMail.Enabled = True
    Frame4.Visible = True
    Register.Visible = True
    RegCode = ""
Else
    EMail.Enabled = False
    Frame4.Visible = False
    Register.Visible = False
End If
'^^^^^ ************************************************************
    
'PCN2151 --------------v
If Dir(ReadOnlyAppPath & "Language\" & EULAFilename) <> "" And Language <> "English" Then 'PCN2123
    RichTextBox1.LoadFile ReadOnlyAppPath & "Language\" & EULAFilename 'PCN2123
End If '---------------^

Exit Sub
Err_Handler:
    Select Case Err
        Case Else
            MsgBox Err & "-R3:" & Error$
            Resume Next
    End Select
End Sub

Private Sub Register_Click()

On Error GoTo Err_Handler
Dim ArmClass As CArmadillo

Set ArmClass = New CArmadillo

'ProductNo = ArmClass.HardwareFingerPrint

'ClearLineProfilerV5.FabLock1.UserName = ProductNo
'ClearLineProfilerV5.FabLock1.RegistrationCode = RegCode

If Not Agree Then
 'MsgBox DisplayMessage("You must Agree to the End User License Agreement Details before Registering this Product."), vbExclamation 'PCN2111
 ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("You must Agree to the End User License Agreement Details before Registering this Product."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
 Exit Sub
End If

'ClearLineProfilerV5.FabLock1.Register False

'If Registered = True Then 'ML310103
'MsgBox "installed = " & ArmClass.IsValidKey(RegUserName, RegCode)
If ArmClass.IsValidKey(RegUserName, RegCode) Then 'ML310103
    Call ArmClass.InstallKey(RegUserName, RegCode)
    Dim ArmRegType As String
    ArmRegType = ArmClass.ClearLineRegType
    'vvvv PCN2861 ***************************
'    Select Case ArmRegType
'        Case "RegisteredWith3D"
'            ThreeDActivated = True
'    End Select
    '^^^^ ***********************************
    'MsgBox DisplayMessage("Congratulations, your ClearLine Profiler Software is now registered. Please shut-down the application and restart to allow new settings to take effect."), vbInformation 'PCN2111
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Congratulations, your ClearLine Profiler Software is now registered. Please shut-down the application and restart to allow new settings to take effect."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Registered = True
    Frame4.Visible = False
    Register.Visible = False
    'PCN2861: Added 3D module to general registration.
    ThreeDActivated = True 'PCN2861
    EMail.Enabled = False
    If Dir(LocToSave & "temp.cbs") <> "" Then 'PCN1971
      Kill LocToSave & "temp.cbs" 'PCN1971
    End If
Else
  'MsgBox DisplayMessage("Registration Failed. Please Check your Registration Code."), vbCritical 'PCN2111
  ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Registration Failed. Please Check your Registration Code."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
End If
Unload Me 'PCN1791

Exit Sub
Err_Handler:
    Select Case Err
        Case Else
            MsgBox Err & "-R4:" & Error$
            Resume Next
    End Select
End Sub


Function SaveProductNo()

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'SaveProductNo Function  Louise Shrimpton  louiseS@cbsys.co.nz
'
'Revision history"
'   V0.0    Louise Shrimpton,    7/1/02
'
'Description:
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim i As Integer

Dim FileNo As Integer
FileNo = FreeFile 'ID5384

Open LocToSave & "temp.cbs" For Output As #FreeFile 'ID5384 #1
For i = Len(ProductNo) To 1 Step -1
    Print #FileNo, Asc(Mid(ProductNo, i, 1)) + Len(ProductNo) - i + 1
Next i
Close #FileNo 'ID5384 #1
    
'Unload Registration  'PCN1787

Exit Function
Err_Handler:
    Select Case Err
        Case Else
            MsgBox Err & "-R5:" & Error$
            Resume Next 'PCNLS170203
    End Select
End Function
