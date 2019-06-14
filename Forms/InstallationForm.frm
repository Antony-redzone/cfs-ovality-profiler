VERSION 5.00
Begin VB.Form InstallationForm 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setup"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11190
   ControlBox      =   0   'False
   Icon            =   "InstallationForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   11190
   Begin ClearLineProfiler.CBS_DropDownBox LanguageDropDown 
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   600
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   661
   End
   Begin ClearLineProfiler.CBS_DropDownBox CameraDropdown 
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   600
      Width           =   4440
      _ExtentX        =   7832
      _ExtentY        =   661
   End
   Begin ClearLineProfiler.CBS_DropDownBox UnitsDropdown 
      Height          =   375
      Left            =   9840
      TabIndex        =   6
      Top             =   600
      Width           =   960
      _ExtentX        =   2117
      _ExtentY        =   661
   End
   Begin VB.CommandButton ContinueButton 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Continue"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
   End
   Begin VB.PictureBox RulerPicture 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   9000
      Picture         =   "InstallationForm.frx":0442
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.PictureBox CameraPicture 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   3600
      Picture         =   "InstallationForm.frx":210C
      ScaleHeight     =   735
      ScaleWidth      =   375
      TabIndex        =   1
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   360
      Picture         =   "InstallationForm.frx":2A0A
      ScaleHeight     =   630
      ScaleWidth      =   585
      TabIndex        =   0
      Top             =   480
      Width           =   585
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1095
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   3015
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1095
      Left            =   3480
      TabIndex        =   8
      Top             =   240
      Width           =   5175
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1095
      Left            =   8880
      TabIndex        =   9
      Top             =   240
      Width           =   2055
   End
   Begin VB.Frame RegTypeFrame 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Registration type"
      Height          =   1215
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Width           =   4215
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.OptionButton HaspLockOption 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "xxxxxx-xxxxxx-xxxxxx-xxxxxx-xxxxxx-xxxxxx"
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   840
         Width           =   3615
      End
      Begin VB.Image HaspLockImage 
         Height          =   495
         Left            =   480
         Picture         =   "InstallationForm.frx":8C9E
         Stretch         =   -1  'True
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "InstallationForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Complete As Boolean
Private IntroTextFile As String
Private TheLanguage() As String
Private TheOther() As String
Private TheContinue() As String
Private TheSetup() As String
Private TheRegType() As String
Private TheViewer() As String
Private DropDownHeight As Single

Private Sub ContinueButton_Click()
On Error GoTo Error_handler

    Call SubmitSettings
    
Exit Sub
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-IF1:" & Error$, vbExclamation
    End Select
End Sub

Private Sub Form_Load()
On Error GoTo Error_handler
    Dim FecCameraModel As String

    Complete = False
    
    Call UnitsDropdown.AddItem("mm")
    Call UnitsDropdown.AddItem("in")
    Call CameraDropdown.AddItem("Viewer")
    
    ReDim TheLanguage(0)
    ReDim TheOther(0)
    ReDim TheContinue(0)
    ReDim TheSetup(0)
    ReDim TheRegType(0)
    ReDim TheViewer(0)
    ReDim TheFECFiles(0)

    IntroTextFile = App.Path & "\language\Introduction.txt"
    If Dir(IntroTextFile) = "" Then
        Call SubmitSettings
    End If
    Call LoadLanguageSettings
    Call LoadCameras(Me.CameraDropdown)

Exit Sub
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-IF2:" & Error$, vbExclamation
    End Select
End Sub

Private Sub SubmitSettings()
On Error GoTo Error_handler
    Dim FishEyeFile As String

    ConfigInfo.Units = UnitsDropdown.text
    Language = LanguageDropDown.text
    Call INI_WriteBack(MyFile, "MeasurementUnits=", UnitsDropdown.text)
    Call INI_WriteBack(MyFile, "Language=", LanguageDropDown.text)
    If HaspLockOption(0).value = True Then
        Call INI_WriteBack(MyFile, "HASPLock=", "true")
    Else
        Call INI_WriteBack(MyFile, "HASPLock=", "false")
    End If
        
    If CameraDropdown.ItemSelected > 0 Then
        FishEyeFile = TheFECFiles(CameraDropdown.ItemSelected - 1)
        Call FisheyeFunctions.FecLoadInformation(App.Path & "\Fec Files\", FishEyeFile)
    Else
        Call FisheyeFunctions.FECLoadDefaultSettings
    End If
    
        
    
    
    Complete = True

Exit Sub
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-IF3:" & Error$, vbExclamation
    End Select
End Sub

Private Sub LoadLanguageSettings()
On Error GoTo Error_handler

    Dim FileNo
    Dim InputString As String
    
    FileNo = FreeFile
    
    Open IntroTextFile For Input As #FileNo
        Do While Not EOF(FileNo)   ' Check for end of file.
        Line Input #FileNo, InputString ' Read line of data.
        Call ParseString(InputString)  ' Seperate the translated continue and other.
    Loop
    
    Close #FileNo
        
Exit Sub
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-IF4:" & Error$, vbExclamation
    End Select
End Sub

Private Sub ParseString(ByVal InputString As String)
On Error GoTo Error_handler

    Dim ChrPos As Integer
    Dim Lang As String
    Dim Cont As String
    Dim Oth As String
    Dim Setu As String
    Dim RegType As String
    Dim View As String
    Dim CountLang As Integer
    
    ChrPos = InStr(InputString, Chr(9))
    Lang = Left(InputString, ChrPos - 1)
    InputString = Right(InputString, Len(InputString) - ChrPos)
    
    ChrPos = InStr(InputString, Chr(9))
    Cont = Left(InputString, ChrPos - 1)
    InputString = Right(InputString, Len(InputString) - ChrPos)
    
    ChrPos = InStr(InputString, Chr(9))
    Oth = Left(InputString, ChrPos - 1)
    InputString = Right(InputString, Len(InputString) - ChrPos)
    
    ChrPos = InStr(InputString, Chr(9))
    Setu = Left(InputString, ChrPos - 1)
    InputString = Right(InputString, Len(InputString) - ChrPos)
    
    ChrPos = InStr(InputString, Chr(9))
    RegType = Left(InputString, ChrPos - 1)
    InputString = Right(InputString, Len(InputString) - ChrPos)
    
    View = InputString
    
    Call LanguageDropDown.AddItem(Lang)
  
    CountLang = UBound(TheLanguage)
    
    TheLanguage(CountLang) = Lang
    TheContinue(CountLang) = Cont
    TheOther(CountLang) = Oth
    TheSetup(CountLang) = Setu
    TheRegType(CountLang) = RegType
    TheViewer(CountLang) = View
    
    ReDim Preserve TheLanguage(CountLang + 1)
    ReDim Preserve TheOther(CountLang + 1)
    ReDim Preserve TheContinue(CountLang + 1)
    ReDim Preserve TheSetup(CountLang + 1)
    ReDim Preserve TheRegType(CountLang + 1)
    ReDim Preserve TheViewer(CountLang + 1)

Exit Sub
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-IF5:" & Error$, vbExclamation
    End Select
End Sub



Private Sub Form_LostFocus()
On Error GoTo Error_handler

    Complete = True
    
Exit Sub
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-IF6:" & Error$, vbExclamation
    End Select
End Sub

Private Sub LanguageDropDown_MouseMove()
On Error GoTo Error_handler
    Dim SelectedText As String
    Dim LangCount As Integer
    
    LangCount = UBound(TheLanguage)
    
    SelectedText = LanguageDropDown.TextHighlited
        For I = 0 To LangCount - 1
        If TheLanguage(I) = SelectedText Then
            CameraDropdown.Item(0) = TheViewer(I): FisheyeFunctions.ViewerString = TheViewer(I)
            ContinueButton.Caption = TheContinue(I)
            InstallationForm.Caption = TheSetup(I)
            RegTypeFrame.Caption = TheRegType(I)
        End If
    Next I
Exit Sub
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-IF7:" & Error$, vbExclamation
    End Select
End Sub
