VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form PVReport1K 
   Caption         =   "1K Report"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   12585
   Begin VB.ComboBox CmboPrinterList 
      Height          =   315
      Left            =   5160
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Text            =   "Select a Printer"
      Top             =   180
      Width           =   3135
   End
   Begin VB.PictureBox PageFramePictureBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   15480
      Left            =   0
      ScaleHeight     =   15450
      ScaleWidth      =   19065
      TabIndex        =   1
      Top             =   600
      Width           =   19095
      Begin VB.PictureBox picReportPageNth 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2055
         Index           =   0
         Left            =   12120
         ScaleHeight     =   2025
         ScaleWidth      =   5505
         TabIndex        =   21
         Tag             =   "Paper"
         Top             =   0
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.PictureBox picReportPagePg1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   16000
         Left            =   0
         ScaleHeight     =   15975
         ScaleWidth      =   11880
         TabIndex        =   2
         Tag             =   "Paper"
         Top             =   0
         Width           =   11904
         Begin ClearLineProfiler.PVDGraphControl PVGraph 
            Height          =   1515
            Index           =   0
            Left            =   480
            TabIndex        =   14
            Top             =   1560
            Width           =   10335
            _ExtentX        =   18653
            _ExtentY        =   3307
         End
         Begin VB.TextBox UserTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2363
            MousePointer    =   3  'I-Beam
            TabIndex        =   5
            Top             =   765
            Width           =   7215
         End
         Begin VB.TextBox FloatingTextDefault 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   8760
            TabIndex        =   4
            Text            =   "Default Text Setting"
            Top             =   120
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.TextBox FloatingText 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   8520
            TabIndex        =   3
            Top             =   120
            Visible         =   0   'False
            Width           =   255
         End
         Begin ClearLineProfiler.PVDGraphControl PVGraph 
            Height          =   1860
            Index           =   1
            Left            =   480
            TabIndex        =   15
            Top             =   3435
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   3281
         End
         Begin ClearLineProfiler.PVDGraphControl PVGraph 
            Height          =   1875
            Index           =   2
            Left            =   480
            TabIndex        =   16
            Top             =   5325
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   3307
         End
         Begin ClearLineProfiler.PVDGraphControl PVGraph 
            Height          =   1815
            Index           =   3
            Left            =   480
            TabIndex        =   17
            Top             =   7200
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   3201
         End
         Begin ClearLineProfiler.PVDGraphControl PVGraph 
            Height          =   2235
            Index           =   4
            Left            =   480
            TabIndex        =   18
            Top             =   9080
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   3942
         End
         Begin ClearLineProfiler.PVDGraphControl PVGraph 
            Height          =   1860
            Index           =   5
            Left            =   480
            TabIndex        =   19
            Top             =   10965
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   3281
         End
         Begin ClearLineProfiler.PVDGraphControl PVGraph 
            Height          =   1875
            Index           =   6
            Left            =   480
            TabIndex        =   20
            Top             =   12840
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   3307
         End
         Begin VB.Label lblTitle 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "1Km Project Report"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   4283
            TabIndex        =   23
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label DirectoryLbl 
            Caption         =   "Directory Label"
            Height          =   255
            Left            =   480
            TabIndex        =   22
            Top             =   1200
            Width           =   8295
            WordWrap        =   -1  'True
         End
         Begin VB.Label CompanyNameLabel 
            Caption         =   "Co"
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
            Left            =   2880
            TabIndex        =   13
            Top             =   14880
            Width           =   5055
         End
         Begin VB.Image CLPLogoImage 
            Height          =   705
            Left            =   600
            Picture         =   "PVRreport1K.frx":0000
            Stretch         =   -1  'True
            Top             =   14880
            Width           =   2010
         End
         Begin VB.Image LogoImage 
            Height          =   690
            Left            =   360
            Picture         =   "PVRreport1K.frx":0865
            Stretch         =   -1  'True
            Top             =   480
            Width           =   2010
         End
         Begin VB.Label PrintedLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "Printed:"
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
            Left            =   9120
            TabIndex        =   12
            Top             =   15120
            Width           =   855
         End
         Begin VB.Label PrintedData 
            Caption         =   "xxxxxx"
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
            Left            =   10080
            TabIndex        =   11
            Top             =   15120
            Width           =   855
         End
         Begin VB.Line FooterBreakLine 
            X1              =   480
            X2              =   11040
            Y1              =   14760
            Y2              =   14760
         End
         Begin VB.Line HeaderBreakLine 
            X1              =   480
            X2              =   11040
            Y1              =   1500
            Y2              =   1500
         End
         Begin VB.Label CleanFlowSystemWebLabel 
            Caption         =   "www.cleanflowsystems.com"
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
            Left            =   4800
            TabIndex        =   10
            Top             =   15120
            Width           =   2295
         End
         Begin VB.Label PhLabel 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ph:"
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
            Left            =   2760
            TabIndex        =   9
            Top             =   15120
            Width           =   375
         End
         Begin VB.Label PhData 
            Caption         =   "xxxxxx"
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
            Left            =   3120
            TabIndex        =   8
            Top             =   15120
            Width           =   1575
         End
         Begin VB.Label PageLabel 
            Caption         =   "Page: 1/1"
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
            Left            =   10080
            TabIndex        =   7
            Top             =   14880
            Width           =   855
         End
         Begin VB.Label CopyrightLabel 
            Caption         =   "Copyright 2009"
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
            Left            =   8040
            TabIndex        =   6
            Top             =   14880
            Width           =   1455
         End
      End
   End
   Begin MSComDlg.CommonDialog PrinterDialogBox 
      Left            =   0
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Printer Settings"
   End
   Begin VB.Image ControlsReport 
      Height          =   720
      Index           =   1
      Left            =   960
      MousePointer    =   1  'Arrow
      Picture         =   "PVRreport1K.frx":10CA
      Tag             =   "Move"
      ToolTipText     =   "Move Report or Text"
      Top             =   0
      Width           =   720
   End
   Begin VB.Image ControlsReport 
      Height          =   720
      Index           =   3
      Left            =   4320
      MousePointer    =   1  'Arrow
      Picture         =   "PVRreport1K.frx":2D94
      Tag             =   "Print"
      ToolTipText     =   "Print"
      Top             =   0
      Width           =   720
   End
   Begin VB.Image ControlHighlightDepressed 
      Height          =   585
      Left            =   11760
      Picture         =   "PVRreport1K.frx":4A5E
      Top             =   90
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Image ControlHighlight 
      Height          =   570
      Left            =   10680
      Picture         =   "PVRreport1K.frx":66A8
      Top             =   90
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Image CloseReport 
      Height          =   480
      Left            =   12840
      MousePointer    =   1  'Arrow
      Picture         =   "PVRreport1K.frx":823A
      ToolTipText     =   "Close Report"
      Top             =   150
      Width           =   480
   End
   Begin VB.Image ControlsBackPanel 
      Height          =   615
      Left            =   0
      MousePointer    =   1  'Arrow
      Picture         =   "PVRreport1K.frx":8F04
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13335
   End
End
Attribute VB_Name = "PVReport1K"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public picReportPage As PictureBox

Public NumberOfExtraPages As Integer
Public TotalNumberOfGraphs As Integer
Public DistancePerGraph As Single
Public Units As String

'Flat Graph folder select

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

'Browsing type.
Public Enum BrowseType
    BrowseForFolders = &H1
    BrowseForComputers = &H1000
    BrowseForPrinters = &H2000
    BrowseForEverything = &H4000
End Enum

'Folder Type
Public Enum FolderType
    CSIDL_BITBUCKET = 10
    CSIDL_CONTROLS = 3
    CSIDL_DESKTOP = 0
    CSIDL_DRIVES = 17
    CSIDL_FONTS = 20
    CSIDL_NETHOOD = 18
    CSIDL_NETWORK = 19
    CSIDL_PERSONAL = 5
    CSIDL_PRINTERS = 4
    CSIDL_PROGRAMS = 2
    CSIDL_RECENT = 8
    CSIDL_SENDTO = 9
    CSIDL_STARTMENU = 11
End Enum

Private Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32.dll" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, ListId As Long) As Long


Private PVDfilesFound() As String

Public Function BrowseFolders(hWndOwner As Long, sMessage As String, Browse As BrowseType, ByVal RootFolder As FolderType) As String

On Error GoTo Err_Handler

    Dim Nullpos As Integer
    Dim lpIDList As Long
    Dim Res As Long
    Dim sPath As String
    Dim BInfo As BrowseInfo
    Dim RootID As Long

    SHGetSpecialFolderLocation hWndOwner, RootFolder, RootID
    BInfo.hWndOwner = hWndOwner
    BInfo.lpszTitle = lstrcat(sMessage, "")
    BInfo.ulFlags = Browse
    If RootID <> 0 Then BInfo.pIDLRoot = RootID
    lpIDList = SHBrowseForFolder(BInfo)
    If lpIDList <> 0 Then
        sPath = String(MAX_PATH, 0)
        Res = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        Nullpos = InStr(sPath, vbNullChar)
        If Nullpos <> 0 Then
            sPath = Left(sPath, Nullpos - 1)
        End If
    End If
    BrowseFolders = sPath

Exit Function
Err_Handler:
    MsgBox Err & " - " & Error$
End Function


Sub PrintPVReport1K()
On Error GoTo Err_Handler
    Dim i As Integer
    Dim GraphIndex As Integer
    
    


    
    RenderScale = 1

    Printer.Orientation = vbPRORPortrait
    Printer.PrintQuality = vbPRPQHigh
    PrinterDialogBox.Orientation = cdlPortrait

    ScreenDrawingType = 1
    ScreenDrawingOrientation = 0
    
    Call RestoreOriginalState
    Call MarkForPrinting
    Me.PageLabel.Caption = "Page: 1/" & (1 + NumberOfExtraPages)
    Set picReportPage = picReportPagePg1
    
    Call RenderToPrinter.RenderReport(Me, Printer, 1)
    Call RenderToPrinter.RenderSingleTextBox(Me.UserTitle, Printer, False) 'PCN4277
    
    For GraphIndex = 0 To TotalNumberOfGraphs
        If (GraphIndex Mod 7 = 0) And (GraphIndex > 0) Then
            Call RestoreOriginalState
            Call MarkForPrinting
            Me.PageLabel.Caption = "Page: " & 1 + Fix(GraphIndex / 7) & "/" & (1 + NumberOfExtraPages)
            Set picReportPage = picReportPagePg1
            
            Printer.NewPage
            Call RenderToPrinter.RenderReport(Me, Printer, 1)
            Call RenderToPrinter.RenderSingleTextBox(Me.UserTitle, Printer, False) 'PCN4277
        End If
        Call PVGraph(GraphIndex).PrintGraph(Printer, 1, Me.PVGraph(GraphIndex).Left, Me.PVGraph(GraphIndex).Top)
    Next GraphIndex
        
    For i = 1 To Me.FloatingText.Count - 1 'PCN4412
        Call RenderToPrinter.RenderSingleTextBox(Me.FloatingText(i), Printer, True)
    Next i

    Call Printer.EndDoc
    ScreenDrawingType = 0
    ScreenDrawingOrientation = 0
    
    Me.RestoreOriginalState
    Call Me.InitialiseFrom
    
'    Me.picReportPagePg1.Visible = True
'    Me.picReportPagePg1.Left = 0
'    Me.picReportPagePg1.Top = 0
'
'    For I = 1 To Me.picReportPageNth.Count - 1
'        Me.picReportPageNth(I).width = Me.picReportPagePg1.width
'        Me.picReportPageNth(I).height = Me.picReportPagePg1.height
'        Me.picReportPageNth(I).Top = (Me.picReportPagePg1.height + 300) * I
'        Me.picReportPageNth(I).Left = Me.picReportPagePg1.Left
'        Me.picReportPageNth(I).Visible = True
'    Next I
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R1K1:" & Error$
    End Select
End Sub


Private Sub CloseReport_Click()
On Error GoTo Err_Handler

Unload Me
    
Exit Sub
Err_Handler:
    MsgBox Err & "-R1K2:" & Error$
End Sub

Private Sub ControlsBackPanel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

If Me.ControlHighlight.Visible Then Me.ControlHighlight.Visible = False
If Me.ControlHighlightDepressed.Visible Then Me.ControlHighlightDepressed.Visible = False
    
Exit Sub
Err_Handler:
    MsgBox Err & "-R1K3:" & Error$
End Sub

Private Sub ControlsReport_Click(Index As Integer)
On Error GoTo Err_Handler

Call ExecuteReportButton(Me, Index)
    
Exit Sub
Err_Handler:
    MsgBox Err & "-R1K4:" & Error$
End Sub

Private Sub ControlsReport_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


Me.ControlHighlightDepressed.Left = Me.ControlHighlight.Left
Me.ControlHighlightDepressed.Visible = True
Me.ControlHighlight.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-R1K5:" & Error$
End Sub

Private Sub ControlsReport_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ControlHighlight.Left = Me.ControlsReport(Index).Left - 100
If Me.ControlHighlight.Visible = False Then Me.ControlHighlight.Visible = True
    
Exit Sub
Err_Handler:
    MsgBox Err & "-R1K6:" & Error$
End Sub

Private Sub ControlsReport_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ControlHighlight.Visible = True
Me.ControlHighlightDepressed.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-R1K7:" & Error$
End Sub



Private Sub FloatingText_Change(Index As Integer)
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingText_Change(Me, Index)
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R1K8:" & Error$
    End Select
End Sub

Private Sub FloatingText_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingText_KeyPress(Me, Index, KeyAscii)
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R1K9:" & Error$
    End Select
End Sub

Private Sub FloatingText_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingText_MouseDown(Me, Index, Button, Shift, X, Y)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R1K10:" & Error$
    End Select
End Sub

Private Sub FloatingText_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingText_MouseMove(Me, Index, Button, Shift, X, Y)
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R1K11:" & Error$
    End Select
End Sub

Private Sub FloatingTextBackgroundColourMenu_Click()
On Error GoTo Err_Handler
    
    Call RenderToPrinter.FloatingTextBackgroundColourMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R1K12:" & Error$
    End Select
End Sub

Private Sub FloatingTextDefaultMenu_Click()
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingTextDefaultMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R1K13:" & Error$
    End Select
End Sub

Private Sub FloatingTextDeleteAllMenu_Click()
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingTextDeleteAllMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R1K14:" & Error$
    End Select
End Sub

Private Sub FloatingTextDeleteMenu_Click()
On Error GoTo Err_Handler
    
    Call RenderToPrinter.FloatingTextDeleteMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R1K15:" & Error$
    End Select
End Sub

Private Sub FloatingTextFontMenu_Click()
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingTextFontMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R1K16:" & Error$
    End Select
End Sub
Sub InitialiseFrom()
On Error GoTo Err_Handler
    Dim i As Integer
'    Dim PVDPath As String
'    Dim PVDName As String
'    Dim PVDExt As String
    
    Dim GraphHeight As Single
    Dim GraphWidth As Single
    
    

    
   'Set ImageGraphState(6).PictureImage = PrecisionVisionGraph.PrinterReportImage 'PCN4271
    
    Me.Left = 0
    Me.width = ClearLineProfilerV6.width - 200
    Me.Top = 0
    Me.height = ClearLineProfilerV6.height - 500
    
    i = Me.Controls.Count
    ReDim OriginalStateVisible(i)
    ReDim OriginalStateTag(i)
    ReDim OriginalStateLeft(i)
    ReDim OriginalStateTop(i)
    ReDim OriginalStateX1(i)
    ReDim OriginalStateY1(i)
    ReDim OriginalStateX2(i)
    ReDim OriginalStateY2(i)
    
    PageFramePictureBox.Left = 0
    PageFramePictureBox.width = Me.width - PageFramePictureBox.Left
    PageFramePictureBox.height = Me.height
    PageFramePictureBox.Top = 650
    Me.ControlsBackPanel.width = Me.width
    Me.CloseReport.Left = Me.width - 750
    
    Call ConvertLanguage(Me, Language) 'PCN4171
    Me.PageLabel.Caption = DisplayMessage("Page") & " 1/1"

    RenderScale = 1
    
    picReportPagePg1.width = Printer.width * RenderScale
    picReportPagePg1.height = Printer.height * RenderScale
    
    Call FooterPosition(Printer.height - 2074)
    
    
    GraphHeight = (Me.FooterBreakLine.y1 - Me.HeaderBreakLine.y1 + 30) / 7
    GraphWidth = Me.picReportPagePg1.width - (480 * 2)
    
    For i = 0 To 6
        PVGraph(i).height = GraphHeight
        PVGraph(i).Top = i * GraphHeight + Me.HeaderBreakLine.y1 + 30
        PVGraph(i).width = GraphWidth
        PVGraph(i).Left = 480
    Next i
    
    For i = 1 To Me.Controls.Count - 1
        OriginalStateVisible(i) = Me.Controls(i).Visible
        OriginalStateTag(i) = Me.Controls(i).Tag
        OriginalStateLeft(i) = Me.Controls(i).Left
        OriginalStateTop(i) = Me.Controls(i).Top
        OriginalStateX1(i) = Me.Controls(i).x1
        OriginalStateY1(i) = Me.Controls(i).y1
        OriginalStateX2(i) = Me.Controls(i).x2
        OriginalStateY2(i) = Me.Controls(i).y2
    Next i
    
    Call FillOutPrintForm
    Set PrintPreviewForm = Me
    
'    Call SplitFilePath(PVDFileName, PVDPath, PVDName, PVDExt)
'    Call FindAllPVDInDirectory(PVDPath)
    Call MarkForPrinting
    Call RenderForm
    
    'Set mouse icon for move
    PrintPreviewAction = "MoveAll"
    Call SetupReportMouseIcon(Me, 108)
    Me.PageFramePictureBox.Visible = True
Exit Sub
Err_Handler:
    Select Case Err
        Case 438, 387, 393
: Resume Next
        Case Else: MsgBox Err & "-R1K17:" & Error$
    End Select
End Sub

Sub FindAllPVDInDirectory(ByVal Path As String)
On Error GoTo Err_Handler

Dim FileName As String
Dim SplitPath As String
Dim SplitName As String
Dim SplitExt As String

Dim i As Integer


FileName = Dir(Path & "*.PVD")
While FileName <> ""
    Call SplitFilePath(FileName, SplitPath, SplitName, SplitExt)
    If LCase(SplitExt) = "pvd" Then
        i = i + 1
        ReDim Preserve PVDfilesFound(i)
        PVDfilesFound(i) = Path & FileName
    End If
    FileName = Dir
Wend


Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R1K18:" & Error$
    End Select
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
    Dim PVDPath As String
    Dim PVDName As String
    Dim PVDExt As String
    Dim GraphIndex As Integer
    Dim PVDIndex As Integer
    Dim SpanInc As Integer
    
    Me.UserTitle.Font.Charset = LanguageCharset
    Me.FloatingText(0).Font.Charset = LanguageCharset
    
    If MeasurementUnits = "mm" Then
       lblTitle.Caption = "1Km Project Report"
       SpanInc = 25

       Units = "m"
    Else
        lblTitle.Caption = "1 Mile Project Report"

        SpanInc = 125
        Units = "ft"
    End If
    
    Call ConvertLanguage(Me, Language) 'PCN4171
    Me.Caption = DisplayMessage("1K Report")
    
    If Confirm1KDialog.SpanOption(0).value = True Then DistancePerGraph = SpanInc * 1
    If Confirm1KDialog.SpanOption(1).value = True Then DistancePerGraph = SpanInc * 2
    If Confirm1KDialog.SpanOption(2).value = True Then DistancePerGraph = SpanInc * 3
    If Confirm1KDialog.SpanOption(3).value = True Then DistancePerGraph = SpanInc * 4
    If Confirm1KDialog.SpanOption(4).value = True Then DistancePerGraph = SpanInc * 5
    If Confirm1KDialog.SpanOption(5).value = True Then DistancePerGraph = SpanInc * 6
        
    
    If PVDFileName = "" Then
        PVDPath = BrowseFolders(hwnd, "Select a Folder", BrowseForFolders, CSIDL_DESKTOP)
        If PVDPath = "" Then
            Unload Me
            Exit Sub
        End If
        PVDPath = PVDPath & "\"
    Else
        Call SplitFilePath(PVDFileName, PVDPath, PVDName, PVDExt)
    End If
    
    Call ScreenDrawing.GraphSelect("Flat", 0)
    Me.CmboPrinterList.Enabled = False
                                                 
    'Me.LogoImage.Picture = LoadPicture(WindowsTempDirectory & "CBS\EmbedFile.jpg")
    Me.LogoImage.Picture = LoadPicture(WindowsTempDirectory & "EmbedFile.jpg") 'ID4601


    Call FindAllPVDInDirectory(PVDPath)
    
    For PVDIndex = 1 To UBound(PVDfilesFound)
        Call LoadFoundPVDs(PVDfilesFound(PVDIndex), GraphIndex)
    Next PVDIndex
    
    TotalNumberOfGraphs = GraphIndex - 1
    
    Call CLPProgressBar.ProgressBarPosition(1#)

    Call GetPrinterList(Me)
    Call Me.InitialiseFrom
    Me.CmboPrinterList.Enabled = True
    
    

Exit Sub
Err_Handler:
    Select Case Err
        Case 53: Resume Next 'File Not Found
        Case 9: Unload Me: Exit Sub
        Case Else: MsgBox Err & "-R1K19:" & Error$
        
    End Select
    
    
End Sub

Sub RenderForm()
On Error GoTo Err_Handler
    Dim i As Integer
    Dim CurrentNoOfExtraPages As Integer
    Dim GraphIndex As Integer
    Dim PageToRender As Integer
    
    
    NumberOfExtraPages = Fix((PVGraph.Count - 1) / 7)
    CurrentNoOfExtraPages = Me.picReportPageNth.Count
    
    For i = CurrentNoOfExtraPages To NumberOfExtraPages
        Call CreateNewPage(i)
    Next i
    
    picReportPagePg1.Cls
    picReportPagePg1.width = Printer.width * RenderScale
    picReportPagePg1.height = Printer.height * RenderScale
    Call FooterPosition((Printer.height - 2074) * RenderScale)
    
    For i = 1 To NumberOfExtraPages
        Me.picReportPageNth(i).Cls
        Me.picReportPageNth(i).width = Printer.width * RenderScale
        Me.picReportPageNth(i).height = Printer.height * RenderScale
    Next i
    
    Call RestoreOriginalState
    Call MarkForPrinting
    Me.PageLabel.Caption = "Page: 1/" & (1 + NumberOfExtraPages)
    Set picReportPage = picReportPagePg1
    Call RenderToPrinter.RenderReport(Me, picReportPage, RenderScale)




    For GraphIndex = 0 To TotalNumberOfGraphs
        If GraphIndex <= 6 Then
            Call PVGraph(GraphIndex).PrintGraph(Me.picReportPagePg1, 1, Me.PVGraph(GraphIndex).Left, Me.PVGraph(GraphIndex).Top)
        Else
            Me.PVGraph(GraphIndex).width = PVGraph(GraphIndex - 7).width
            Me.PVGraph(GraphIndex).height = PVGraph(GraphIndex - 7).height
            Me.PVGraph(GraphIndex).Top = PVGraph(GraphIndex - 7).Top
            Me.PVGraph(GraphIndex).Left = PVGraph(GraphIndex - 7).Left
            PageToRender = Fix(GraphIndex / 7)
            Call PVGraph(GraphIndex).PrintGraph(Me.picReportPageNth(PageToRender), 1, Me.PVGraph(GraphIndex).Left, Me.PVGraph(GraphIndex).Top)
        End If
    Next GraphIndex
    
    Call CLPProgressBar.ProgressBarPosition(1#)

    For i = 1 To picReportPageNth.Count - 1
        Call RestoreOriginalState
        Call MarkForPrinting
        Me.PageLabel.Caption = "Page: " & 1 + i & "/" & (1 + NumberOfExtraPages)
        Set picReportPage = picReportPageNth(i)

        Call RenderToPrinter.RenderReport(Me, picReportPage, RenderScale)
    Next i
    
    picReportPagePg1.Visible = True
    For i = 1 To (picReportPageNth.Count - 1)
        Me.picReportPageNth(i).width = Me.picReportPagePg1.width
        Me.picReportPageNth(i).height = Me.picReportPagePg1.height
        Me.picReportPageNth(i).Top = (Me.picReportPagePg1.height + 300) * i
        Me.picReportPageNth(i).Left = Me.picReportPagePg1.Left
        Me.picReportPageNth(i).Visible = True
    Next i

        
    Me.picReportPageNth(0).Visible = False

    ScreenDrawingType = 0
    ScreenDrawingOrientation = 0

    Me.UserTitle.Visible = True

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R1K20:" & Error$
    End Select

End Sub



Sub LoadFoundPVDs(ByVal FileName As String, ByRef GraphIndex As Integer)
On Error GoTo Err_Handler
    Dim StartFrameDistance As Long
    Dim EndFrameDistance As Long
    Dim PageToRender As Integer
    Dim GraphTitle As String * 40
    Dim DisplayUnits
    
    
    Dim SplitPath As String
    Dim SplitName As String
    Dim SplitExt As String
    
    If GraphIndex > 6 Then Call CreateNewGraph(GraphIndex)
    ScreenDrawingType = 0
    ScreenDrawingOrientation = 0
    PVDLoadError = False
    DebrisOn = False
    LoadVideo = False: Call OpenPVDFile(FileName): LoadVideo = True
    If PVDLoadError = True Then Exit Sub
    If Trim(ConfigInfo.DistanceProcessMethod) = "None" Then Exit Sub
    If ConfigInfo.DistanceStart <= -1000 Or ConfigInfo.DistanceFinish <= -1000 Then Exit Sub
    If PVDistances(1) <= -1000 Then Exit Sub
    
    PVGraph(GraphIndex).SetComment = PipelineInfo.Comments
    Call SplitFilePath(FileName, SplitPath, SplitName, SplitExt)
    Me.DirectoryLbl.Caption = SplitPath
    
    
    GraphTitle = Trim(SplitName) & ": "
    
    PVGraph(GraphIndex).SetGraphTitle = GraphTitle & _
                                        Trim(PipelineInfo.AssetNo) & " - " & _
                                        Round(ExpectedDiameter, 2) & MeasurementUnits & " - " & _
                                        Trim(PipelineInfo.Material) & " - " & _
                                        DisplayMessage("From") & " " & Trim(PipelineInfo.StartName) & _
                                        " " & DisplayMessage("To") & " " & Trim(PipelineInfo.FinishName)
    PVGraph(GraphIndex).SetCommentCaption = DisplayMessage("Comments")
                                        
                                      
    PVGraph(GraphIndex).SetGraphLength = DistancePerGraph
    PVGraph(GraphIndex).SetGraphUnit = Units
    
    

    
    
    StartFrameDistance = PVDistances(1)
    EndFrameDistance = PVDistances(PVDataNoOfLines)
    
    If StartFrameDistance > EndFrameDistance Then StartFrameDistance = EndFrameDistance
    
    PVGraph(GraphIndex).SetStartDistance = StartFrameDistance
    PVGraph(GraphIndex).SetEndDistance = StartFrameDistance + DistancePerGraph
    
    
    
    
    Do
        StartFrameDistance = Fix(PVGraph(GraphIndex).GetStartDistance)
        EndFrameDistance = Fix(PVGraph(GraphIndex).GetEndDistance)
        ScreenDrawingType = 2
        ScreenDrawingOrientation = 1
        Call PVGraph(GraphIndex).DrawPVGraphsReport
        ScreenDrawingType = 0
        ScreenDrawingOrientation = 0
        
        If EndFrameDistance - StartFrameDistance < DistancePerGraph Then Exit Do
        GraphIndex = GraphIndex + 1
        If GraphIndex > 6 Then Call CreateNewGraph(GraphIndex)
        PVGraph(GraphIndex).SetHideInfo = True
        PVGraph(GraphIndex).SetGraphLength = DistancePerGraph
            PVGraph(GraphIndex).SetGraphUnit = Units
        
        
            
        PVGraph(GraphIndex).SetStartDistance = PVGraph(GraphIndex - 1).GetEndDistance
        PVGraph(GraphIndex).SetEndDistance = PVGraph(GraphIndex).GetStartDistance + DistancePerGraph
        
    Loop
    GraphIndex = GraphIndex + 1
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R1K21:" & Error$
    End Select
End Sub

Sub CreateNewGraph(ByRef GraphIndex As Integer)
On Error GoTo Err_Handler

Dim PageNo As Integer

Load PVGraph(GraphIndex)
'PageNo = Me.picReportPageNth.Count

Me.PVGraph(GraphIndex).width = PVGraph(GraphIndex - 7).width
Me.PVGraph(GraphIndex).height = PVGraph(GraphIndex - 7).height
Me.PVGraph(GraphIndex).Top = PVGraph(GraphIndex - 7).Top
Me.PVGraph(GraphIndex).Left = PVGraph(GraphIndex - 7).Left

'If GraphIndex Mod 7 = 0 Then
'    Load Me.picReportPageNth(PageNo)
'    Me.picReportPageNth(PageNo).width = Me.picReportPagePg1.width
'    Me.picReportPageNth(PageNo).height = Me.picReportPagePg1.height
'    Me.picReportPageNth(PageNo).Top = (Me.picReportPagePg1.height + 300) * PageNo
'    Me.picReportPageNth(PageNo).Left = Me.picReportPagePg1.Left
'    Me.picReportPageNth(PageNo).Visible = True
'End If
    
Exit Sub
Err_Handler:
    Select Case Err
        Case 360: Exit Sub 'Allready created
        Case Else: MsgBox Err & "-R1K22:" & Error$
    End Select
End Sub

Sub CreateNewPage(ByVal PageNo As Integer)
On Error GoTo Err_Handler

    Load Me.picReportPageNth(PageNo)
    Me.picReportPageNth(PageNo).width = Me.picReportPagePg1.width
    Me.picReportPageNth(PageNo).height = Me.picReportPagePg1.height
    Me.picReportPageNth(PageNo).Top = (Me.picReportPagePg1.height + 300) * PageNo
    Me.picReportPageNth(PageNo).Left = Me.picReportPagePg1.Left
    Me.picReportPageNth(PageNo).Visible = True

Exit Sub
Err_Handler:
    Select Case Err
        Case 360: Exit Sub 'Allready created
        Case Else: MsgBox Err & "-R1K23:" & Error$
    End Select
End Sub
Private Sub Form_Resize()
On Error GoTo Err_Handler
    PageFramePictureBox.width = Me.width - PageFramePictureBox.Left
    PageFramePictureBox.height = Me.height
    PageFramePictureBox.Top = 650
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R1K24:" & Error$
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler
Dim FileSaveFail As Boolean

Call SaveToFilePipeObs(FileSaveFail)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R1K25:" & Error$
    End Select
End Sub

Private Sub PageFramePictureBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


If Me.ControlHighlight.Visible Then Me.ControlHighlight.Visible = False
If Me.ControlHighlightDepressed.Visible Then Me.ControlHighlightDepressed.Visible = False
    
Exit Sub
Err_Handler:
    MsgBox Err & "-R1K26:" & Error$
End Sub

Private Sub picReportPagePg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Call ReportPageMouseDown(Me, picReportPagePg1, Button, Shift, X, Y) 'PCN4193

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R1K27:" & Error$
    End Select

End Sub

Private Sub picReportPagePg1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    Dim i As Integer
    
    If ReportMouseDown Then
        picReportPagePg1.Left = picReportPagePg1.Left + X - ReportMouseX
        picReportPagePg1.Top = picReportPagePg1.Top + Y - ReportMouseY
        For i = 1 To picReportPageNth.Count - 1
                picReportPageNth(i).Left = picReportPageNth(i).Left + X - ReportMouseX
                picReportPageNth(i).Top = picReportPageNth(i).Top + Y - ReportMouseY
        Next i
        
        
    End If
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R1K28:" & Error$
    End Select
End Sub

Private Sub picReportPagePg1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

ReportMouseDown = False
If PrintPreviewAction = "MoveAll" Then
    Call SetupReportMouseIcon(Me, 108)
End If

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R1K29:" & Error$
    End Select
End Sub


Private Sub FillOutPrintForm()
On Error GoTo Err_Handler
    With Me
        .PrintedData.Caption = CStr(Date)
        .PhData.Caption = PhoneNo
        .LogoImage.Picture = LoadPicture(CompanyLogoPath)
        .CompanyNameLabel = CompanyName
    End With
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R1K30:" & Error$
    End Select
    
End Sub

Sub GraphSpecificSettings()
On Error GoTo Err_Handler

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R1K31:" & Error$
    End Select
End Sub

Sub MarkForPrinting()
On Error GoTo Err_Handler

Dim i As Integer
Dim ControlType As String

'Draw renderings first that are marked back
For i = 1 To Me.Controls.Count - 1
    
    With Me.Controls(i)
        If TypeName(.Container) = "PictureBox" Then
            If .Tag = "Paper" Then
                .Visible = True
            ElseIf .Tag = "Back" Then
                .Visible = False
            ElseIf .Visible Then
                .Tag = "Visible"
                .Visible = False
            Else
                .Visible = False
                .Tag = ""
            End If
        End If
No_Container:
    End With
Next i
Exit Sub
Err_Handler:
    Select Case Err
        Case 438: Resume No_Container
        Case Else: MsgBox Err & "-R1K32:" & Error$
    End Select
End Sub



Private Sub PrinterSettingsButton_Click()
On Error GoTo Err_Handler
    PrinterDialogBox.ShowPrinter
Exit Sub
Err_Handler:
    If Err = 32755 Then Exit Sub ' Cancel Printer Setting
End Sub


'===========================
'Declare the Function to select printer
'===========================
 
Function SelectPrinter(ByVal printer_name As String) As Boolean
On Error GoTo Err_Handler
    
    Dim i As Integer
 
    SelectPrinter = False
    For i = 0 To Printers.Count - 1
        ' if the specified printer is found, select it and return True
        If Printers(i).DeviceName = printer_name Then
            Set Printer = Printers(i)
            SelectPrinter = True
            Exit For
        End If
    Next i
    
Exit Function
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R1K33:" & Error$
    End Select
End Function
 

Private Sub CmboPrinterList_Click()
On Error GoTo Err_Handler
    Dim i As Integer
    
    Call SelectPrinter(CmboPrinterList.text)
    
    If Me.CmboPrinterList.Enabled = True Then
        Me.picReportPagePg1.Cls
        
        For i = 1 To Me.FloatingText.Count - 1 'We dont want the text to be rendered on the preview yet
            Me.FloatingText(i).Visible = False 'when changes printers
        Next i
        
        Me.RestoreOriginalState
        Me.InitialiseFrom
        
        For i = 1 To Me.FloatingText.Count - 1 'Even thou we dont want the text to be rendered, we still
            Me.FloatingText(i).Visible = True  'want to be able to see them after printer select changed
        Next i
    End If
    
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R1K34:" & Error$
    End Select
End Sub



Sub RestoreOriginalState()
On Error GoTo Err_Handler
    Dim i As Long

    For i = 1 To Me.Controls.Count - 1
        If Me.Controls(i).name <> "FloatingText" Then
            Me.Controls(i).Visible = OriginalStateVisible(i)
            Me.Controls(i).Tag = OriginalStateTag(i)
            Me.Controls(i).Left = OriginalStateLeft(i)
            Me.Controls(i).Top = OriginalStateTop(i)
            Me.Controls(i).x1 = OriginalStateX1(i)
            Me.Controls(i).y1 = OriginalStateY1(i)
            Me.Controls(i).x2 = OriginalStateX2(i)
            Me.Controls(i).y2 = OriginalStateY2(i)
        End If
    Next i
Exit Sub
Err_Handler:
    Select Case Err
        Case 9 'Subscript out of range
            'Addition of text will cause this error
            Exit Sub
        Case 438, 382: Resume Next
        Case Else: MsgBox Err & "-R1K35:" & Error$
    End Select

End Sub




Private Sub picReportPageNth_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    Call ReportPageMouseDown(Me, picReportPagePg1, Button, Shift, X, Y) 'PCN4193
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R1K36:" & Error$
    End Select
End Sub

Private Sub picReportPageNth_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    Call picReportPagePg1_MouseMove(Button, Shift, X, Y)
    

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R1K37:" & Error$
    End Select
End Sub

Private Sub picReportPageNth_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

ReportMouseDown = False
If PrintPreviewAction = "MoveAll" Then
    Call SetupReportMouseIcon(Me, 108)
End If

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R1K38:" & Error$
    End Select
End Sub

Private Sub UserTitle_Change()
On Error GoTo Err_Handler

Dim FileSaveFail As Boolean

UserTitleProfile = Trim(Me.UserTitle.text) 'PCN4433

Exit Sub
Err_Handler:
    Select Case Err

        Case Else: MsgBox Err & "-R1K39:" & Error$
    End Select
End Sub

Sub FooterPosition(ByVal Position As Single)
On Error GoTo Err_Handler
    FooterBreakLine.y1 = Position
    FooterBreakLine.y2 = Position
    CLPLogoImage.Top = Position + 120
    CompanyNameLabel.Top = Position + 120
    CopyrightLabel.Top = Position + 120
    PageLabel.Top = Position + 120
    PhLabel.Top = Position + 360
    PhData.Top = Position + 360
    CleanFlowSystemWebLabel.Top = Position + 360
    PrintedLabel.Top = Position + 360
    PrintedData.Top = Position + 360
    Exit Sub
Err_Handler:
    Select Case Err

        Case Else: MsgBox Err & "-R1K40:" & Error$
    End Select
    
End Sub



