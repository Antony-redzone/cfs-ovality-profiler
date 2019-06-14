VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form ControlsMain 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "ControlsMain.frx":0000
   ScaleHeight     =   9390
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ControlsDisplayImageList 
      Left            =   1800
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsMain.frx":1E792
            Key             =   "PVGraph"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsMain.frx":2046C
            Key             =   "PVGraphNotSelected"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsMain.frx":22146
            Key             =   "PipeDetails"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsMain.frx":23E20
            Key             =   "PipeDetailsNotSelected"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsMain.frx":25AFA
            Key             =   "Options"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsMain.frx":277D4
            Key             =   "OptionsNotSelected"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsMain.frx":294AE
            Key             =   "PVSettings"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsMain.frx":2B188
            Key             =   "PVSettingsNotSelected"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ControlsMainFnsImageList 
      Left            =   2880
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsMain.frx":2CE62
            Key             =   "Background"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsMain.frx":2D5DC
            Key             =   "Depress"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsMain.frx":2DD56
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlsMain.frx":2E4D0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image ControlsMainFns 
      Height          =   360
      Index           =   7
      Left            =   1965
      Picture         =   "ControlsMain.frx":2EC4A
      Tag             =   "PVD_Deployment"
      ToolTipText     =   "Create Viewer"
      Top             =   1110
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image ControlsDisplaySonar 
      Height          =   720
      Index           =   2
      Left            =   2280
      Picture         =   "ControlsMain.frx":2F3B4
      ToolTipText     =   "Options"
      Top             =   4320
      Width           =   720
   End
   Begin VB.Image ControlsDisplaySonar 
      Height          =   720
      Index           =   1
      Left            =   1440
      Picture         =   "ControlsMain.frx":3107E
      ToolTipText     =   "Pipeline Details"
      Top             =   4320
      Width           =   720
   End
   Begin VB.Image ControlsDisplaySonar 
      Height          =   720
      Index           =   0
      Left            =   600
      Picture         =   "ControlsMain.frx":32D48
      ToolTipText     =   "Sonar"
      Top             =   4320
      Width           =   720
   End
   Begin VB.Image ControlsMainFns 
      Height          =   360
      Index           =   6
      Left            =   2880
      Picture         =   "ControlsMain.frx":34A12
      Tag             =   "Sonar"
      ToolTipText     =   "Sonar"
      Top             =   1110
      Width           =   360
   End
   Begin VB.Image ControlsMainFns 
      Height          =   360
      Index           =   5
      Left            =   2400
      Picture         =   "ControlsMain.frx":3517C
      Tag             =   "LiveConnect"
      ToolTipText     =   "Connect Video"
      Top             =   1110
      Width           =   360
   End
   Begin VB.Image ControlsMainFns 
      Height          =   360
      Index           =   4
      Left            =   3840
      Picture         =   "ControlsMain.frx":358E6
      Top             =   1110
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image ControlsMainFns 
      Height          =   360
      Index           =   3
      Left            =   3360
      Picture         =   "ControlsMain.frx":36050
      Tag             =   "Information"
      ToolTipText     =   "Information and Help"
      Top             =   1110
      Width           =   360
   End
   Begin VB.Image ControlsMainFns 
      Height          =   360
      Index           =   2
      Left            =   645
      Picture         =   "ControlsMain.frx":367BA
      Tag             =   "SaveToFile"
      ToolTipText     =   "Save To File"
      Top             =   120
      Width           =   360
   End
   Begin VB.Image ControlsMainFns 
      Height          =   360
      Index           =   1
      Left            =   165
      Picture         =   "ControlsMain.frx":36F24
      Tag             =   "OpenFile"
      ToolTipText     =   "Open File"
      Top             =   120
      Width           =   360
   End
   Begin VB.Image ControlsMainFns 
      Height          =   360
      Index           =   0
      Left            =   1320
      Picture         =   "ControlsMain.frx":3768E
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image ControlsDisplay 
      Height          =   720
      Index           =   2
      Left            =   2880
      Picture         =   "ControlsMain.frx":37DF8
      Tag             =   "PVSettings"
      ToolTipText     =   "Precision Vision settings"
      Top             =   180
      Width           =   720
   End
   Begin VB.Image ControlsDisplay 
      Height          =   720
      Index           =   1
      Left            =   2040
      Picture         =   "ControlsMain.frx":39AC2
      Tag             =   "DisplayPipeDetails"
      ToolTipText     =   "Pipeline Details"
      Top             =   120
      Width           =   720
   End
   Begin VB.Image ControlsDisplay 
      Height          =   720
      Index           =   0
      Left            =   1200
      Picture         =   "ControlsMain.frx":3B78C
      Tag             =   "DisplayPVGraph"
      ToolTipText     =   "Precision Vision Graphs"
      Top             =   120
      Width           =   720
   End
   Begin VB.Image ControlsDisplay 
      Height          =   720
      Index           =   3
      Left            =   3600
      Picture         =   "ControlsMain.frx":3D456
      Tag             =   "Options"
      ToolTipText     =   "Options"
      Top             =   120
      Width           =   720
   End
   Begin VB.Image ControlViewPressed 
      Height          =   960
      Left            =   4320
      Picture         =   "ControlsMain.frx":3F120
      Top             =   75
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlViewHighlight 
      Height          =   960
      Left            =   1200
      Picture         =   "ControlsMain.frx":3FAF2
      Top             =   75
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image ControlFilePressed 
      Height          =   600
      Left            =   5040
      Picture         =   "ControlsMain.frx":4113C
      Top             =   75
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image ControlFileHighlight 
      Height          =   600
      Left            =   0
      Picture         =   "ControlsMain.frx":41815
      Top             =   75
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image ControlMainPressed 
      Height          =   600
      Left            =   5040
      Picture         =   "ControlsMain.frx":4201E
      Top             =   1015
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image ControlMainHighlight 
      Height          =   600
      Left            =   4320
      Picture         =   "ControlsMain.frx":426AD
      Top             =   1015
      Visible         =   0   'False
      Width           =   540
   End
End
Attribute VB_Name = "ControlsMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public MainViewSelected As Integer 'PCN4277


Private Sub ControlsDisplay_Click(Index As Integer)
On Error GoTo Err_Handler

MainViewSelected = Index 'PCN4277
Call ControlsDisplaySetup(ControlsDisplay(Index).Tag)
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CM1:" & Error$
End Sub

Private Sub ControlsDisplay_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

If Me.ControlViewPressed.Left <> Me.ControlsDisplay(Index).Left Then Me.ControlViewPressed.Left = Me.ControlsDisplay(Index).Left 'PCN4277 - 100
If Me.ControlViewPressed.Visible = False Then Me.ControlViewPressed.Visible = True
    

Exit Sub
Err_Handler:
    MsgBox Err & "-CM2:" & Error$
End Sub

Private Sub ControlsDisplay_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


If Me.ControlViewHighlight.Left <> Me.ControlsDisplay(Index).Left Then Me.ControlViewHighlight.Left = Me.ControlsDisplay(Index).Left 'PCN4277 - 100
If Me.ControlViewHighlight.Visible = False Then Me.ControlViewHighlight.Visible = True

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CM3:" & Error$
End Sub



Private Sub ControlsDisplay_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    If Me.ControlViewPressed.Visible = True Then Me.ControlViewPressed.Visible = False

Exit Sub
Err_Handler:
    MsgBox Err & "-CM4:" & Error$
End Sub

Private Sub ControlsMainFns_Click(Index As Integer)
On Error GoTo Err_Handler

Call ControlsAction(Me.ControlsMainFns(Index).Tag)

Exit Sub
Err_Handler:
    MsgBox Err & "-CM5:" & Error$
End Sub

Private Sub ControlsMainFns_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

'PCN4277
' Picture = Me.ControlsMainFnsImageList.ListImages("Depress").Picture

'PCN4277
If Index > 2 Then
    If Me.ControlMainPressed.Left <> Me.ControlMainHighlight.Left Then Me.ControlMainPressed.Left = Me.ControlMainHighlight.Left
    If Me.ControlMainPressed.Visible = False Then Me.ControlMainPressed.Visible = True
End If

If Index < 3 Then
    If Me.ControlFilePressed.Left <> Me.ControlFileHighlight.Left Then Me.ControlFilePressed.Left = Me.ControlFileHighlight.Left
    If Me.ControlFilePressed.Visible = False Then Me.ControlFilePressed.Visible = True
End If

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CM6:" & Error$
End Sub

Private Sub ControlsMainFns_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

'PCN4277
'If Index > 0 And Me.ControlsMainFns(0).Left <> Me.ControlsMainFns(Index).Left Then
'    Me.ControlsMainFns(0).Left = Me.ControlsMainFns(Index).Left
'End If

If Index < 3 Then
    If Me.ControlFileHighlight.Left <> Me.ControlsMainFns(Index).Left - 75 Then
        Me.ControlFileHighlight.Left = Me.ControlsMainFns(Index).Left - 75
    End If
    If Me.ControlFileHighlight.Visible = False Then Me.ControlFileHighlight.Visible = True
End If

If Index > 2 Then
    If Me.ControlMainHighlight.Left <> Me.ControlsMainFns(Index).Left - 75 Then
        Me.ControlMainHighlight.Left = Me.ControlsMainFns(Index).Left - 75
    End If
    If Me.ControlMainHighlight.Visible = False Then Me.ControlMainHighlight.Visible = True

End If


'PCN4277
'If Me.ControlsMainFns(0).Visible = False Then Me.ControlsMainFns(0).Visible = True

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CM7:" & Error$
End Sub

Private Sub ControlsMainFns_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

'If Index > 0 Then
'    Set Me.ControlsMainFns(0).Picture = Me.ControlsMainFnsImageList.ListImages("Background").Picture
'End If


If Index > 2 Then
    If Me.ControlMainPressed.Visible = True Then Me.ControlMainPressed.Visible = False
End If
If Index < 3 Then
    If Me.ControlFilePressed.Visible = True Then Me.ControlFilePressed.Visible = False
End If

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CM8:" & Error$
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
    
Me.Top = ClearLineScreen.height + ClearLineScreen.Top
Me.height = 1650
Me.Left = ControlsScreen.width
Me.width = PVPageWidth
Call ConvertLanguage(Me, Language)

'vvvv PCN3809 ********************************
If SoftwareConfiguration = "Reader" Then
    Call SetupMainForReaderConfiguration
End If
'^^^^ ****************************************

Exit Sub
Err_Handler:
    MsgBox Err & "-CM9:" & Error$
End Sub


Sub ControlsAction(ByVal Action As String)
On Error GoTo Err_Handler
   
   
Select Case Action
    Case "OpenFile": Me.ControlFileHighlight.Visible = False: Call OpenAnyFile("") 'PCN2133
    Case "SaveToFile"
         Me.ControlFileHighlight.Visible = False
        If PVRecording = True Then 'PCN2379
            'MsgBox DisplayMessage("Stop PVD recording before saving") 'PCN2762
            ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Stop PVD recording before saving"): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
            
            Exit Sub
        End If
        If Registered = False Then 'PCNML220103 'Testing ML040203
            'MsgBox DisplayMessage("Cannot save a .PVD file, please register the software to access this."), vbExclamation 'PCN2111
            ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Cannot save a .PVD file, please register the software to access this."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
        Else
            Call SaveImageAndOrData 'PCNGL110103
            Call SetDisplayMainsFns   'PCN4298
        End If
    Case "Information"
        Dim hwndHelp As Long
        If SoftwareConfiguration = "Reader" And (InStr(1, HelpFilename, "Reader_") = 0) Then 'PCN4444
            If HelpFilename <> "" And HelpFilename <> "ReaderHelpfile.chm" Then 'PCN4553 once set to reader it keeped adding reader string to filename
                HelpFilename = "Reader_" & HelpFilename 'PCN4444
            Else
                HelpFilename = "ReaderHelpfile.chm"
            End If
        End If
        
        
        If HelpFilename <> "" And Dir(ReadOnlyAppPath & "Language\" & HelpFilename) = "" Then 'Check whether that file exists actually.'PCN2167 7/8/03 by Abe
            MsgBox HelpFilename & " " & DisplayMessage("file for the language") & _
                "(" & Language & ") " & _
                DisplayMessage("does not exist. Create this file first. The default language(English) is loaded for Help file."), , "Clear Line Profiler" 'PCN2167 7/8/03 by Abe, PCN2171
        End If
        If HelpFilename <> "" And Dir(ReadOnlyAppPath & "Language\" & HelpFilename) <> "" Then  'PCN2167 7/8/03 by Abe ---------v
            hwndHelp = HtmlHelp(hwnd, ReadOnlyAppPath & "Language\" & HelpFilename, HH_DISPLAY_TOPIC, 0)
        Else '-------------------------------------------------------------------------------------------^
            hwndHelp = HtmlHelp(hwnd, App.Path & "\" & "ProfilerHelp.chm", HH_DISPLAY_TOPIC, 0) 'PCN4400
        End If

    Case "DisplayPVGraph"
        Load PrecisionVisionGraph
        PrecisionVisionGraph.Show
        PrecisionVisionGraph.ZOrder 0
    
    Case "DisplayPipeDetails"
        Load PipelineDetails
        PipelineDetails.Show
        PipelineDetails.ZOrder 0
        
    Case "Options"
        Load OptionsPage
        OptionsPage.Show
        OptionsPage.ZOrder 0
    
    Case "PVSettings"
        AutoTune.Show
        AutoTune.ZOrder 0
        Call AutoTune.SetupSelectedTask("")
        
    Case "PVD_Deployment"
        Call DeployPVD
        
    Case "LiveConnect"
        Call ClearLineScreen.LiveVideoConnect
''            'PCN2395 (22 September 2003, Ant) enable or disable live button if there is or isn't caputre device
''        Dim AnyCaptureDevices As Long
''        Call hough_anycapturedevices(AnyCaptureDevices)
''
''        If AnyCaptureDevices = 1 And VideoCaptureDevice > 0 Then          '
''        ' ClearLineScreen.ControlToolbar.Buttons.Item(1).Enabled = True      ' 'PCN4171
''        'Else                                                                '
''        ' ClearLineScreen.ControlToolbar.Buttons.Item(1).Enabled = False     '
''        End If                                                              '
''        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "Sonar"
        InvokeSonar
        'SonarConfig.Show
        'SonarConfig.ZOrder 0
        
End Select


Exit Sub
Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & "-CM10:" & Error$
End Select
End Sub

Sub InvokeSonar()
On Error GoTo Err_Handler

Dim SonarSTTFilename As String

SonarSTTFilename = App.Path & "\ProfilerSonarA.stt"
If Dir(SonarSTTFilename) <> "" Then
    Call INI_WriteBack(SonarSTTFilename, "CompanyName=", CompanyName)
    Call INI_WriteBack(SonarSTTFilename, "PhoneNo=", PhoneNo)
    Call INI_WriteBack(SonarSTTFilename, "FaxNo=", FaxNo)
    Call INI_WriteBack(SonarSTTFilename, "CompanyLogoPath=", CompanyLogoPath)
    Call INI_WriteBack(SonarSTTFilename, "MeasurementUnits=", Left(MeasurementUnits, 2))
    Call INI_WriteBack(SonarSTTFilename, "Language=", Language)

    Call INI_WriteBack(SonarSTTFilename, "PipeDetailsAssetNo=", PipelineDetails.AssetNo_lbl.Caption) 'PCN2820
    Call INI_WriteBack(SonarSTTFilename, "PipeDetailsSiteID=", PipelineDetails.SiteID_lbl.Caption)  'PCN2820
    Call INI_WriteBack(SonarSTTFilename, "PipeDetailsCity=", PipelineDetails.City_lbl.Caption) 'PCN2820
    Call INI_WriteBack(SonarSTTFilename, "PipeDetailsDate=", PipelineDetails.Date_lbl.Caption) 'PCN2820
    Call INI_WriteBack(SonarSTTFilename, "PipeDetailsTime=", PipelineDetails.Time_lbl.Caption)  'PCN2820
    Call INI_WriteBack(SonarSTTFilename, "PipeDetailsStNode=", PipelineDetails.StartNodeNo_lbl.Caption) 'PCN2820
    Call INI_WriteBack(SonarSTTFilename, "PipeDetailsStLoc=", PipelineDetails.StNodeLoc_lbl.Caption) 'PCN2820
    Call INI_WriteBack(SonarSTTFilename, "PipeDetailsFhNode=", PipelineDetails.FinishNodeNo_lbl.Caption) 'PCN2820
    Call INI_WriteBack(SonarSTTFilename, "PipeDetailsFhLoc=", PipelineDetails.FhNodeLoc_lbl.Caption) 'PCN2820
    Call INI_WriteBack(SonarSTTFilename, "PipeDetailsIntDiaExp=", PipelineDetails.InternalDiameterExpected_lbl.Caption) 'PCN2820
    Call INI_WriteBack(SonarSTTFilename, "PipeDetailsOutDiaExp=", PipelineDetails.OutsideDiameter_lbl.Caption) 'PCN2820
    Call INI_WriteBack(SonarSTTFilename, "PipeDetailsLength=", PipelineDetails.PipeLen_lbl.Caption) 'PCN2820
    Call INI_WriteBack(SonarSTTFilename, "PipeDetailsMaterial=", PipelineDetails.Material_lbl.Caption) 'PCN2820

    Call INI_WriteBack(SonarSTTFilename, "PipeDetailsValueAssetNo=", PipelineDetails.AssetNo) 'PCN2820
    Call INI_WriteBack(SonarSTTFilename, "PipeDetailsValueSiteID=", PipelineDetails.SiteID)  'PCN2820
    Call INI_WriteBack(SonarSTTFilename, "PipeDetailsValueCity=", PipelineDetails.City) 'PCN2820
    Call INI_WriteBack(SonarSTTFilename, "PipeDetailsValueDate=", PipelineDetails.sDate) 'PCN2820
    Call INI_WriteBack(SonarSTTFilename, "PipeDetailsValueTime=", PipelineDetails.sTime)  'PCN2820
    Call INI_WriteBack(SonarSTTFilename, "PipeDetailsValueStNode=", PipelineDetails.StartNodeNo) 'PCN2820
    Call INI_WriteBack(SonarSTTFilename, "PipeDetailsValueStLoc=", PipelineDetails.StartNodeLocation) 'PCN2820
    Call INI_WriteBack(SonarSTTFilename, "PipeDetailsValueFhNode=", PipelineDetails.FinishNodeNo) 'PCN2820
    Call INI_WriteBack(SonarSTTFilename, "PipeDetailsValueFhLoc=", PipelineDetails.FinishNodeLocation) 'PCN2820
    Call INI_WriteBack(SonarSTTFilename, "PipeDetailsValueIntDiaExp=", PipelineDetails.InternalDiameterExpected) 'PCN2820
    Call INI_WriteBack(SonarSTTFilename, "PipeDetailsValueOutDiaExp=", PipelineDetails.OutsideDiameter) 'PCN2820
    Call INI_WriteBack(SonarSTTFilename, "PipeDetailsValueLength=", PipelineDetails.PipeLength) 'PCN2820
    Call INI_WriteBack(SonarSTTFilename, "PipeDetailsValueMaterial=", PipelineDetails.Material) 'PCN2820
End If
    
If Dir(App.Path & "/ProfilerSonarA.exe") <> "" Then
    Call Shell(App.Path & "/ProfilerSonarA.exe -populate", vbNormalFocus)
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-CM11:" & Error$

End Sub
Sub ControlsDisplaySetup(DisplayThisPage As String)
On Error GoTo Err_Handler



Set Me.ControlsDisplay(0).Picture = Me.ControlsDisplayImageList.ListImages("PVGraphNotSelected").Picture
Set Me.ControlsDisplay(1).Picture = Me.ControlsDisplayImageList.ListImages("PipeDetailsNotSelected").Picture
Set Me.ControlsDisplay(2).Picture = Me.ControlsDisplayImageList.ListImages("PVSettingsNotSelected").Picture
Set Me.ControlsDisplay(3).Picture = Me.ControlsDisplayImageList.ListImages("OptionsNotSelected").Picture

Select Case DisplayThisPage
    Case "DisplayPVGraph"
        Set Me.ControlsDisplay(0).Picture = Me.ControlsDisplayImageList.ListImages("PVGraph").Picture
        MainViewSelected = 0 'PCN4277
    
    Case "DisplayPipeDetails"
        Set Me.ControlsDisplay(1).Picture = Me.ControlsDisplayImageList.ListImages("PipeDetails").Picture
        MainViewSelected = 1 'PCN4277
        
    Case "PVSettings"
        Set Me.ControlsDisplay(2).Picture = Me.ControlsDisplayImageList.ListImages("PVSettings").Picture
        MainViewSelected = 2 'PCN4277

    Case "Options"
        Set Me.ControlsDisplay(3).Picture = Me.ControlsDisplayImageList.ListImages("Options").Picture
        MainViewSelected = 3 'PCN4277
    
End Select
Call SetupDisplayViewHighlight 'PCN4277

Call ControlsAction(DisplayThisPage)

Call SetDisplayMainsFns   'PCN4298

    
Exit Sub
Err_Handler:
    MsgBox Err & "-CM12:" & Error$
End Sub

Sub SetDisplayMainsFns() 'PCN4298
On Error GoTo Err_Handler

'vvvv PCN4241 ****************************
'Display the appropriate Save button
If SoftwareConfiguration <> "Reader" Then
    If PVDFileName <> "" And PVDFileName <> LocToSave & DefaultPVDFileName Then 'A saved PVD is loaded
        Me.ControlsMainFns(7).Visible = True
    Else
        Me.ControlsMainFns(7).Visible = False
    End If
End If
'^^^^ ************************************

'PCN2395 (22 September 2003, Ant) enable or disable live button if there is or isn't caputre device
Dim AnyCaptureDevices As Long
Call hough_anycapturedevices(AnyCaptureDevices)

'PCN4243 Putting live video back in, it was disabled.       '
If AnyCaptureDevices = 1 And VideoCaptureDevice > 0 Then    '
    Me.ControlsMainFns(5).Enabled = True      ' 'PCN4171    '
Else                                                        '
    Me.ControlsMainFns(5).Enabled = False                   '
End If                                                      '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Exit Sub
Err_Handler:
    MsgBox Err & "-CM13:" & Error$
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


'If Me.ControlViewHighlight.Visible Then Me.ControlViewHighlight.Visible = False
'Me.ControlViewHighlight.Visible = False
If X > Me.ControlsDisplay(0).Left And X < Me.ControlsDisplay(3).Left And Y < 960 And Y > 75 Then Exit Sub

Call SetupDisplayViewHighlight 'PCN4277

'PCN4277
'If Me.ControlsMainFns(0).Visible Then Me.ControlsMainFns(0).Visible = False
If Me.ControlFileHighlight.Visible = True Then Me.ControlFileHighlight.Visible = False
If Me.ControlMainHighlight.Visible = True Then Me.ControlMainHighlight.Visible = False
    
Exit Sub
Err_Handler:
    MsgBox Err & "-CM14:" & Error$
End Sub

Sub SetupDisplayViewHighlight() ' PCN4277
On Error GoTo Err_Handler

If Me.ControlViewHighlight.Left <> Me.ControlsDisplay(MainViewSelected).Left Then Me.ControlViewHighlight.Left = Me.ControlsDisplay(MainViewSelected).Left
If Me.ControlViewHighlight.Visible = False Then Me.ControlViewHighlight.Visible = True

Exit Sub
Err_Handler:
    MsgBox Err & "-CM15:" & Error$
End Sub


Sub SetupMainForReaderConfiguration() 'PCN3809
On Error GoTo Err_Handler

'Hide all unused controls
'Hide ControlsMainFns controls
For ButtonIndex = 2 To 6
    Me.ControlsMainFns(ButtonIndex).Visible = False
Next ButtonIndex
'Hide Options button
Me.ControlsDisplay(2).Visible = False
'Hide Precision Vision Settings button
Me.ControlsDisplay(3).Visible = False

'Unhide help button
Me.ControlsMainFns(3).Visible = True 'PCN4444






Exit Sub
Err_Handler:
    MsgBox Err & "-CM16:" & Error$
End Sub

Sub DeployPVD() 'PCN4241
On Error GoTo Err_Handler

Dim Resp As Variant
Dim SourceDir As String
Dim SourceExe As String
Dim DestExe As String
Dim SourceClearLineDLL As String
Dim SourceLaserLibDLL As String
Dim SourceThreeDimDLL As String
Dim SourceCOMDLG32OCX As String 'PCN4337
Dim SourceMSCOMCTLOCX As String 'PCN4337
Dim SourceLaserTexture As String 'PCN4416
Dim SourceDeployINI As String 'PCN4431
Dim SourceMSVBVM60 As String 'PCN4423
Dim SourceReaderHelpFile As String 'PCN4445
Dim SourceLanguageTxt As String 'PCN4469

Dim SourceReaderHelpFileInternational As String

'Dim SourceReaderHelpFileGermanLanguage As String 'PCN4469
'Dim SourceReaderHelpFileSwedishLanguage As String
'Dim SourceReaderHelpFileFrenchCanadianLanguage As String

'''''''''''''''
Dim SourceMSVBVM50 As String
Dim SourceEMPGDMX As String
Dim SourceRICHTX32 As String

'''''''''''''''

Dim SourceLanguageInternationalTxt As String
'Dim SourceGermanLanguage As String
'Dim SourceSwedishLanguage As String
'Dim SourceFrenchCanadianLanguage As String
'Dim SourceJapaneseLanguage As String

Dim TargetDir As String
Dim TargetFileName As String
Dim LanguageFile As String
Dim PVDVideoFileName As String
Dim PVDVideoDir As String
Dim DeploymentFile As String

Dim ShapeFile As String

ClearLineProfilerV6.Dialog.Filter = "Precision Vision File (*.pvd)|*.pvd||"
ClearLineProfilerV6.Dialog.FileName = PVDFileName
ClearLineProfilerV6.Dialog.ShowSave

If Dir(ClearLineProfilerV6.Dialog.FileName) <> "" Then
'    'Resp = MsgBox(DisplayMessage("File already exists. Will you overwrite?"), vbYesNo)
    ProfilerMessageBox.MsgBoxYesNo (DisplayMessage("File already exists. Will you overwrite?"))
    Resp = PMBAnswer
    If Resp = vbNo Then Exit Sub
End If

'Confirm the size of the deployment and notify


'Copy all files to the deployment directory
SourceDir = App.Path & "\"

'Before first copy any INI information needed for the deploy

Call TransferINIInformation 'PCN4431

'First copy the PVD
TargetFileName = ClearLineProfilerV6.Dialog.FileName
If TargetFileName = PVDFileName Then
    'MsgBox DisplayMessage("Saving to the same file name."), vbExclamation
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Saving to the same file name."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Exit Sub
ElseIf InStr(1, TargetFileName, SourceDir) <> 0 Then
    'MsgBox DisplayMessage("Can not save to the application directory"), vbExclamation
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Can not save to the application directory"): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Exit Sub
End If

On Error Resume Next
Kill TargetFileName
FileCopy PVDFileName, TargetFileName
On Error GoTo Err_Handler
TargetFileName = Dir(ClearLineProfilerV6.Dialog.FileName)
If TargetFileName = "" Then
    'Copy failed
    Exit Sub
End If
'Determine target directory
TargetDir = Left(ClearLineProfilerV6.Dialog.FileName, InStr(1, ClearLineProfilerV6.Dialog.FileName, TargetFileName) - 1)

'Other files to copy
DestExe = "Profiler.exe"
SourceExe = "Profiler Viewer.dat"   'PCN????    'Antony 31 August 2009
SourceClearLineDLL = "clearline.dll"
SourceLaserLibDLL = "laserlib.dll"
SourceThreeDimDLL = "threedim.dll"
SourceLaserTexture = "Textures\laser.jpg" 'PCN4416
SourceDeployINI = "Deploy.ini" 'PCN4431
SourceReaderHelpFile = "Language\ReaderHelpfile.chm" 'PCN4445

SourceReaderHelpFileInternational = "Language\Reader_" & Language & "HelpFile.chm"

'SourceReaderHelpFileGermanLanguage = "Language\Reader_GermanHelpFile.chm" 'PCN4469
'SourceReaderHelpFileSwedishLanguage = "Language\Reader_SwedishHelpFile.chm"
'SourceReaderHelpFileFrenchCanadianLanguage = "Language\Reader_FrenchCanadianHelpFile.chm"

SourceLanguageTxt = "Language\Languages.txt" 'PCN4469
SourceLanguageInternationalTxt = "Language\" & Language & ".txt"

'SourceGermanLanguage = "Language\German.txt" 'PCN4469
'SourceSwedishLanguage = "Language\Swedish.txt"
'SourceFrenchCanadianLanguage = "Language\FrenchCanadian.txt"
'SourceJapaneseLanguage = "Language\Japanese.txt"

SourceMSVBVM60 = "MSVBVM60.dll" 'PCN4423 'PCN???? Antony 31 August 2009 change to local
SourceMSVBVM50 = "MSVBVM50.dll" 'Some operating system installs need this file, eg, some clean Vistand 7
SourceEMPGDMX = "EMPGDMX.AX     'Elle card demuxer if needed"
SourceCOMDLG32OCX = "COMDLG32.OCX" 'needs to copy from local
SourceMSCOMCTLOCX = "MSCOMCTL.OCX" 'needs to copy from local
SourceRICHTX32 = "RICHTX32.OCX" 'PCN6027

On Error Resume Next
Kill TargetDir & SourceExe
Kill TargetDir & SourceClearLineDLL
Kill TargetDir & SourceLaserLibDLL
Kill TargetDir & SourceThreeDimDLL
Kill TargetDir & SourceLaserTexture
Kill TargetDir & SourceDeployINI 'PCN4431
Kill TargetDir & SourceReaderHelpFile

Kill TargetDir & SourceReaderHelpFileInternational

'Kill TargetDir & SourceReaderHelpFileGermanLanguage 'PCN4469
'Kill TargetDir & SourceReaderHelpFileSwedishLanguage
'Kill TargetDir & SourceReaderHelpFileFrenchCanadianLanguage

                 
Kill TargetDir & SourceLanguageInternationalTxt


'Kill TargetDir & SourceGermanLanguage 'PCN4469
'Kill TargetDir & SourceSwedishLanguage
'Kill TargetDir & SourceFrenchCanadianLanguage
Kill TargetDir & SourceLanguageTxt 'PCN4469

' moved to local or new file 'When I mean local I mean instead getting the source from system32 make it availibe in application directory
Kill TargetDir & SourceMSVBVM60
Kill TargetDir & SourceMSVBVM50
Kill TargetDir & SourceEMPGDMX
Kill TargetDir & SourceCOMDLG32OCX
Kill TargetDir & SourceMSCOMCTLOCX
Kill targerdir & SourceRICHTX32 'PCN6026



FileCopy SourceDir & SourceExe, TargetDir & DestExe
FileCopy SourceDir & SourceClearLineDLL, TargetDir & SourceClearLineDLL
FileCopy SourceDir & SourceLaserLibDLL, TargetDir & SourceLaserLibDLL
FileCopy SourceDir & SourceThreeDimDLL, TargetDir & SourceThreeDimDLL

'''''''''''''''' the following were from system 32

FileCopy SourceDir & SourceCOMDLG32OCX, TargetDir & SourceCOMDLG32OCX
FileCopy SourceDir & SourceMSCOMCTLOCX, TargetDir & SourceMSCOMCTLOCX
FileCopy SourceDir & SourceMSVBVM60, TargetDir & SourceMSVBVM60 'PCN4423
FileCopy SourceDir & SourceMSVBVM50, TargetDir & SourceMSVBVM50 'PCN4423
FileCopy SourceDir & SourceEMPGDMX, TargetDir & SourceEMPGDMX 'PCN4423
FileCopy SourceDir & SourceRICHTX32, targerdir & SourceRICHTX32 'PCN6026

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

MkDir TargetDir & "Textures" 'PCN4416
FileCopy SourceDir & SourceLaserTexture, TargetDir & SourceLaserTexture 'PCN4416
FileCopy SourceDir & SourceDeployINI, TargetDir & SourceDeployINI 'PCN4431
MkDir TargetDir & "Language"
FileCopy SourceDir & SourceReaderHelpFile, TargetDir & SourceReaderHelpFile 'PCN4445

FileCopy SourceDir & SourceReaderHelpFileInternational, TargetDir & SourceReaderHelpFileInternational

'FileCopy SourceDir & SourceReaderHelpFileGermanLanguage, TargetDir & SourceReaderHelpFileGermanLanguage
'FileCopy SourceDir & SourceReaderHelpFileSwedishLanguage, TargetDir & SourceReaderHelpFileSwedishLanguage
'FileCopy SourceDir & SourceReaderHelpFileFrenchCanadianLanguage, TargetDir & SourceReaderHelpFileFrenchCanadianLanguage


FileCopy SourceDir & SourceLanguageInternationalTxt, TargetDir & SourceLangaugeInternationalTxt
'FileCopy SourceDir & SourceGermanLanguage, TargetDir & SourceGermanLanguage
'FileCopy SourceDir & SourceSwedishLanguage, TargetDir & SourceSwedishLanguage
'FileCopy SourceDir & SourceFrenchCanadianLanguage, TargetDir & SourceFrenchCanadianLanguage

FileCopy SourceDir & SourceLanguageTxt, TargetDir & SourceLanguageTxt


'Make shape files folder and copy all shape files - Richard Ashcroft 19/03/10
'****************************************************************************
MkDir TargetDir & "Shape Files"

ShapeFile = Dir(SourceDir & "Shape Files\*.shp")
Do While ShapeFile <> ""
    FileCopy SourceDir & "Shape Files\" & ShapeFile, TargetDir & "Shape Files\" & ShapeFile
    ShapeFile = Dir()
Loop
'****************************************************************************

'vvvv PCN4337 ******************************************
'here they were files from system 32, now manually part of install

'^^^^ **************************************************

On Error GoTo Err_Handler
'The Video file link by the PVD
'If the video already exists at this directory don't copy it
If VideoFileName <> "" Then
    PVDVideoFileName = Dir(VideoFileName)
    If PVDVideoFileName <> "" Then
        PVDVideoDir = Left(LCase(VideoFileName), InStr(1, LCase(VideoFileName), LCase(PVDVideoFileName)) - 1) 'PCN Deploy fix, Antony 12 Nov 2008
        
    
        
        
        'Check to see the target video dir is the same as the VideoDir
        If PVDVideoDir <> TargetDir Then
            'Copy video file to target dir
            FileCopy VideoFileName, TargetDir & PVDVideoFileName
        End If
    End If
End If
'Language file in a sub directory, GermanV6.txt dependant on current language setting
'Determine what language file to copy (none if in English)
LanguageFile = GetLanguageFile(Language)
'Create the language directory
On Error Resume Next
MkDir TargetDir & "Language"
If Language <> "English" And LanguageFile <> "" Then
    Kill TargetDir & "Language\" & LanguageFile 'PCN4494 this originally killed the 3dim file instead of the language file
    FileCopy SourceDir & "Language\" & LanguageFile, TargetDir & "Language\" & LanguageFile
End If
On Error GoTo Err_Handler


'Copy the CD/DVD autostart and menu files
'Copy all the files in the 'Deployment' directory 'PCN4241
DeploymentFile = Dir(SourceDir & "Deployment\*.*")
Do While DeploymentFile <> ""
    'Check to see if the file exists in the target directory
    On Error Resume Next
    FileCopy SourceDir & "Deployment\" & DeploymentFile, TargetDir & DeploymentFile
    On Error GoTo Err_Handler
    DeploymentFile = Dir()
Loop
'Now copy all files in the Autorun dir
On Error Resume Next
MkDir TargetDir & "\Autorun"
On Error GoTo Err_Handler
DeploymentFile = Dir(SourceDir & "Deployment\Autorun\*.*")
Do While DeploymentFile <> ""
    'Check to see if the file exists in the target directory
    On Error Resume Next
    FileCopy SourceDir & "Deployment\Autorun\" & DeploymentFile, TargetDir & "Autorun\" & DeploymentFile
    On Error GoTo Err_Handler
    DeploymentFile = Dir()
Loop


'MsgBox DisplayMessage("Completed")
ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Completed"): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0


Exit Sub
Err_Handler:
    Select Case Err
        Case 32755 'Cancel selected - PCN4254
            'MsgBox DisplayMessage("Creation of Viewer files cancelled."), vbInformation
            ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Creation of Viewer files cancelled."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
            Exit Sub
        Case Else
            MsgBox Err & "-CM17:" & Error$

    End Select
End Sub


'Sub DeployPVD() 'PCN4241
'On Error GoTo Err_Handler
'Dim Resp As Variant
'Dim SourceDir As String
'Dim SourceExe As String
'Dim DestExe As String
'Dim SourceClearLineDLL As String
'Dim SourceLaserLibDLL As String
'Dim SourceThreeDimDLL As String
'Dim SourceCOMDLG32OCX As String 'PCN4337
'Dim SourceMSCOMCTLOCX As String 'PCN4337
''Dim SourceSMAKEJPGOCX As String 'PCN4337
'Dim SourceLaserTexture As String 'PCN4416
'Dim SourceDeployINI As String 'PCN4431
'Dim SourceMSVBVM60 As String 'PCN4423
'Dim SourceReaderHelpFile As String 'PCN4445
'Dim SourceLanguageTxt As String 'PCN4469
'Dim SourceReaderHelpFileGermanLanguage As String 'PCN4469
'Dim SourceReaderHelpFileSwedishLanguage As String
'Dim SourceReaderHelpFileFrenchCanadianLanguage As String
'
''''''''''''''''
'Dim SourceMSVBVM50 As String
'Dim SourceEMPGDMX As String
'
''''''''''''''''
'
'
'
'Dim SourceGermanLanguage As String
'Dim SourceSwedishLanguage As String
'Dim SourceFrenchCanadianLanguage As String
'Dim SourceJapaneseLanguage As String
'
'Dim TargetDir As String
'Dim TargetFileName As String
'Dim LanguageFile As String
'Dim PVDVideoFileName As String
'Dim PVDVideoDir As String
'Dim DeploymentFile As String
'
'Dim ShapeFile As String
'
'ClearLineProfilerV6.Dialog.Filter = "Precision Vision File (*.pvd)|*.pvd||"
'ClearLineProfilerV6.Dialog.FileName = PVDFileName
'ClearLineProfilerV6.Dialog.ShowSave
'
'If Dir(ClearLineProfilerV6.Dialog.FileName) <> "" Then
'    'Resp = MsgBox(DisplayMessage("File already exists. Will you overwrite?"), vbYesNo)
'    ProfilerMessageBox.MsgBoxYesNo (DisplayMessage("File already exists. Will you overwrite?"))
'    Resp = PMBAnswer
'    If Resp = vbNo Then Exit Sub
'End If
'
''Confirm the size of the deployment and notify
'
'
''Copy all files to the deployment directory
'SourceDir = App.Path & "\"
'
''Before first copy any INI information needed for the deploy
'
'Call TransferINIInformation 'PCN4431
'
''First copy the PVD
'TargetFileName = ClearLineProfilerV6.Dialog.FileName
'If TargetFileName = PVDFileName Then
'    'MsgBox DisplayMessage("Saving to the same file name."), vbExclamation
'    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Saving to the same file name."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
'    Exit Sub
'ElseIf InStr(1, TargetFileName, SourceDir) <> 0 Then
'    'MsgBox DisplayMessage("Can not save to the application directory"), vbExclamation
'    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Can not save to the application directory"): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
'    Exit Sub
'End If
'
'On Error Resume Next
'Kill TargetFileName
'FileCopy PVDFileName, TargetFileName
'On Error GoTo Err_Handler
'TargetFileName = Dir(ClearLineProfilerV6.Dialog.FileName)
'If TargetFileName = "" Then
'    'Copy failed
'    Exit Sub
'End If
''Determine target directory
'TargetDir = Left(ClearLineProfilerV6.Dialog.FileName, InStr(1, ClearLineProfilerV6.Dialog.FileName, TargetFileName) - 1)
'
''Other files to copy
'DestExe = "Profiler.exe"
'SourceExe = "Profiler Viewer.dat"   'PCN????    'Antony 31 August 2009
'SourceClearLineDLL = "clearline.dll"
'SourceLaserLibDLL = "laserlib.dll"
'SourceThreeDimDLL = "threedim.dll"
'SourceLaserTexture = "Textures\laser.jpg" 'PCN4416
'SourceDeployINI = "Deploy.ini" 'PCN4431
'SourceReaderHelpFile = "Language\ReaderHelpfile.chm" 'PCN4445
'SourceReaderHelpFileGermanLanguage = "Language\Reader_GermanHelpFile.chm" 'PCN4469
'SourceReaderHelpFileSwedishLanguage = "Language\Reader_SwedishHelpFile.chm"
'SourceReaderHelpFileFrenchCanadianLanguage = "Language\Reader_FrenchCanadianHelpFile.chm"
'SourceLanguageTxt = "Language\Languages.txt" 'PCN4469
'SourceGermanLanguage = "Language\German.txt" 'PCN4469
'SourceSwedishLanguage = "Language\Swedish.txt"
'SourceFrenchCanadianLanguage = "Language\FrenchCanadian.txt"
'SourceJapaneseLanguage = "Language\Japanese.txt"
'
'SourceMSVBVM60 = "MSVBVM60.dll" 'PCN4423 'PCN???? Antony 31 August 2009 change to local
'SourceMSVBVM50 = "MSVBVM50.dll" 'Some operating system installs need this file, eg, some clean Vistand 7
'SourceEMPGDMX = "EMPGDMX.AX     'Elle card demuxer if needed"
'SourceCOMDLG32OCX = "COMDLG32.OCX" 'needs to copy from local
'SourceMSCOMCTLOCX = "MSCOMCTL.OCX" 'needs to copy from local
'
'On Error Resume Next
'Kill TargetDir & SourceExe
'Kill TargetDir & SourceClearLineDLL
'Kill TargetDir & SourceLaserLibDLL
'Kill TargetDir & SourceThreeDimDLL
'Kill TargetDir & SourceLaserTexture
'Kill TargetDir & SourceDeployINI 'PCN4431
'Kill TargetDir & SourceReaderHelpFile
'Kill TargetDir & SourceReaderHelpFileGermanLanguage 'PCN4469
'Kill TargetDir & SourceReaderHelpFileSwedishLanguage
'Kill TargetDir & SourceReaderHelpFileFrenchCanadianLanguage
'Kill TargetDir & SourceGermanLanguage 'PCN4469
'Kill TargetDir & SourceSwedishLanguage
'Kill TargetDir & SourceFrenchCanadianLanguage
'Kill TargetDir & SourceLanguageTxt 'PCN4469
'
'' moved to local or new file 'When I mean local I mean instead getting the source from system32 make it availibe in application directory
'Kill TargetDir & SourceMSVBVM60
'Kill TargetDir & SourceMSVBVM50
'Kill TargetDir & SourceEMPGDMX
'Kill TargetDir & SourceCOMDLG32OCX
'Kill TargetDir & SourceMSCOMCTLOCX
'
'
'FileCopy SourceDir & SourceExe, TargetDir & DestExe
'FileCopy SourceDir & SourceClearLineDLL, TargetDir & SourceClearLineDLL
'FileCopy SourceDir & SourceLaserLibDLL, TargetDir & SourceLaserLibDLL
'FileCopy SourceDir & SourceThreeDimDLL, TargetDir & SourceThreeDimDLL
'
''''''''''''''''' the following were from system 32
'
'FileCopy SourceDir & SourceCOMDLG32OCX, TargetDir & SourceCOMDLG32OCX
'FileCopy SourceDir & SourceMSCOMCTLOCX, TargetDir & SourceMSCOMCTLOCX
'FileCopy SourceDir & SourceMSVBVM60, TargetDir & SourceMSVBVM60 'PCN4423
'FileCopy SourceDir & SourceMSVBVM50, TargetDir & SourceMSVBVM50 'PCN4423
'FileCopy SourceDir & SourceEMPGDMX, TargetDir & SourceEMPGDMX 'PCN4423
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'MkDir TargetDir & "Textures" 'PCN4416
'FileCopy SourceDir & SourceLaserTexture, TargetDir & SourceLaserTexture 'PCN4416
'FileCopy SourceDir & SourceDeployINI, TargetDir & SourceDeployINI 'PCN4431
'MkDir TargetDir & "Language"
'FileCopy SourceDir & SourceReaderHelpFile, TargetDir & SourceReaderHelpFile 'PCN4445
'FileCopy SourceDir & SourceReaderHelpFileGermanLanguage, TargetDir & SourceReaderHelpFileGermanLanguage
'FileCopy SourceDir & SourceReaderHelpFileSwedishLanguage, TargetDir & SourceReaderHelpFileSwedishLanguage
'FileCopy SourceDir & SourceReaderHelpFileFrenchCanadianLanguage, TargetDir & SourceReaderHelpFileFrenchCanadianLanguage
'FileCopy SourceDir & SourceGermanLanguage, TargetDir & SourceGermanLanguage
'FileCopy SourceDir & SourceSwedishLanguage, TargetDir & SourceSwedishLanguage
'FileCopy SourceDir & SourceFrenchCanadianLanguage, TargetDir & SourceFrenchCanadianLanguage
'FileCopy SourceDir & SourceLanguageTxt, TargetDir & SourceLanguageTxt
'
'
''Make shape files folder and copy all shape files - Richard Ashcroft 19/03/10
''****************************************************************************
'MkDir TargetDir & "Shape Files"
'
'ShapeFile = Dir(SourceDir & "Shape Files\*.shp")
'Do While ShapeFile <> ""
'    FileCopy SourceDir & "Shape Files\" & ShapeFile, TargetDir & "Shape Files\" & ShapeFile
'    ShapeFile = Dir()
'Loop
''****************************************************************************
'
''vvvv PCN4337 ******************************************
''here they were files from system 32, now manually part of install
'
''^^^^ **************************************************
'
'On Error GoTo Err_Handler
''The Video file link by the PVD
''If the video already exists at this directory don't copy it
'PVDVideoFileName = Dir(VideoFileName)
'If PVDVideoFileName <> "" Then
'    PVDVideoDir = Left(LCase(VideoFileName), InStr(1, LCase(VideoFileName), LCase(PVDVideoFileName)) - 1) 'PCN Deploy fix, Antony 12 Nov 2008
'
'
'
'
'    'Check to see the target video dir is the same as the VideoDir
'    If PVDVideoDir <> TargetDir Then
'        'Copy video file to target dir
'        FileCopy VideoFileName, TargetDir & PVDVideoFileName
'    End If
'End If
'
''Language file in a sub directory, GermanV6.txt dependant on current language setting
''Determine what language file to copy (none if in English)
'LanguageFile = GetLanguageFile(Language)
''Create the language directory
'On Error Resume Next
'MkDir TargetDir & "Language"
'If Language <> "English" And LanguageFile <> "" Then
'    Kill TargetDir & "Language\" & LanguageFile 'PCN4494 this originally killed the 3dim file instead of the language file
'    FileCopy SourceDir & "Language\" & LanguageFile, TargetDir & "Language\" & LanguageFile
'End If
'On Error GoTo Err_Handler
'
'
''Copy the CD/DVD autostart and menu files
''Copy all the files in the 'Deployment' directory 'PCN4241
'DeploymentFile = Dir(SourceDir & "Deployment\*.*")
'Do While DeploymentFile <> ""
'    'Check to see if the file exists in the target directory
'    On Error Resume Next
'    FileCopy SourceDir & "Deployment\" & DeploymentFile, TargetDir & DeploymentFile
'    On Error GoTo Err_Handler
'    DeploymentFile = Dir()
'Loop
''Now copy all files in the Autorun dir
'On Error Resume Next
'MkDir TargetDir & "\Autorun"
'On Error GoTo Err_Handler
'DeploymentFile = Dir(SourceDir & "Deployment\Autorun\*.*")
'Do While DeploymentFile <> ""
'    'Check to see if the file exists in the target directory
'    On Error Resume Next
'    FileCopy SourceDir & "Deployment\Autorun\" & DeploymentFile, TargetDir & "Autorun\" & DeploymentFile
'    On Error GoTo Err_Handler
'    DeploymentFile = Dir()
'Loop
'
'
''MsgBox DisplayMessage("Completed")
'ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Completed"): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
'
'
'Exit Sub
'Err_Handler:
'    Select Case Err
'        Case 32755 'Cancel selected - PCN4254
'            'MsgBox DisplayMessage("Creation of Viewer files cancelled."), vbInformation
'            ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Creation of Viewer files cancelled."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
'            Exit Sub
'        Case Else
'            MsgBox Err & "-CM17:" & Error$
'            Resume
'    End Select
'End Sub

Sub TransferINIInformation()
On Error GoTo Err_Handler
Dim DeployINIFilename As String

Call WriteNewViewerINI 'PCN6025

DeployINIFilename = App.Path & "\deployment\deploy.ini"
If Dir(DeployINIFilename) <> "" Then
    Call INI_WriteBack(DeployINIFilename, "PipeDetailsAssetNo=", PipelineDetails.AssetNo_lbl.Caption) 'PCN2820
    Call INI_WriteBack(DeployINIFilename, "PipeDetailsSiteID=", PipelineDetails.SiteID_lbl.Caption)  'PCN2820
    Call INI_WriteBack(DeployINIFilename, "PipeDetailsCity=", PipelineDetails.City_lbl.Caption) 'PCN2820
    Call INI_WriteBack(DeployINIFilename, "PipeDetailsDate=", PipelineDetails.Date_lbl.Caption) 'PCN2820
    Call INI_WriteBack(DeployINIFilename, "PipeDetailsTime=", PipelineDetails.Time_lbl.Caption)  'PCN2820
    Call INI_WriteBack(DeployINIFilename, "PipeDetailsStNode=", PipelineDetails.StartNodeNo_lbl.Caption) 'PCN2820
    Call INI_WriteBack(DeployINIFilename, "PipeDetailsStLoc=", PipelineDetails.StNodeLoc_lbl.Caption) 'PCN2820
    Call INI_WriteBack(DeployINIFilename, "PipeDetailsFhNode=", PipelineDetails.FinishNodeNo_lbl.Caption) 'PCN2820
    Call INI_WriteBack(DeployINIFilename, "PipeDetailsFhLoc=", PipelineDetails.FhNodeLoc_lbl.Caption) 'PCN2820
    Call INI_WriteBack(DeployINIFilename, "PipeDetailsIntDiaExp=", PipelineDetails.InternalDiameterExpected_lbl.Caption) 'PCN2820
    Call INI_WriteBack(DeployINIFilename, "PipeDetailsOutDiaExp=", PipelineDetails.OutsideDiameter_lbl.Caption) 'PCN2820
    Call INI_WriteBack(DeployINIFilename, "PipeDetailsLength=", PipelineDetails.PipeLen_lbl.Caption) 'PCN2820
    Call INI_WriteBack(DeployINIFilename, "PipeDetailsMaterial=", PipelineDetails.Material_lbl.Caption) 'PCN2820
    Call INI_WriteBack(DeployINIFilename, "Language=", Language)  'PCN4469
    If OptionsPage.MedianDiameterOpt(0).value = True Then MedianFlat = True: Call INI_WriteBack(DeployINIFilename, "FlatType=", "Deflection")
    If OptionsPage.MedianDiameterOpt(1).value = True Then MedianFlat = False: Call INI_WriteBack(DeployINIFilename, "FlatType=", "Flat")
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-CM18:" & Error$
End Sub

'PCN6025 this function added 27th Jan 2011
Sub WriteNewViewerINI()
On Error GoTo Err_Handler

   Dim FileNo
   Dim DeployINIFilename As String

    DeployINIFilename = App.Path & "\deployment\deploy.ini"
    FileNo = FreeFile
    
    Open DeployINIFilename For Output As #FileNo
    
    Print #FileNo, "[Revision]"
    Print #FileNo, "INIRevision=2.0"
    Print #FileNo, "[PipeDetails]"
    Print #FileNo, "PipeDetailsAssetNo="
    Print #FileNo, "PipeDetailsSiteID="
    Print #FileNo, "PipeDetailsCity="
    Print #FileNo, "PipeDetailsDate="
    Print #FileNo, "PipeDetailsTime="
    Print #FileNo, "PipeDetailsStNode="
    Print #FileNo, "PipeDetailsStLoc="
    Print #FileNo, "PipeDetailsFhNode="
    Print #FileNo, "PipeDetailsFhLoc="
    Print #FileNo, "PipeDetailsIntDiaExp="
    Print #FileNo, "PipeDetailsOutDiaExp="
    Print #FileNo, "PipeDetailsLength="
    Print #FileNo, "PipeDetailsMaterial="
    Print #FileNo, "[Regional Options]"
    Print #FileNo, "Language="
    Print #FileNo, "[Profiler Settings]"
    Print #FileNo, "FlatType="  'PCN6025
    
    Close #FileNo
    
    Exit Sub
Err_Handler:
    MsgBox Err & "-ST37:" & Error$

End Sub


