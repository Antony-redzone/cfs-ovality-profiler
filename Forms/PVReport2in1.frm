VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form PVReport2in1 
   Caption         =   "1K Report"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12945
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   12945
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
      Begin VB.PictureBox picReportPagePg1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   16000
         Left            =   0
         ScaleHeight     =   15975
         ScaleWidth      =   11880
         TabIndex        =   3
         Tag             =   "Paper"
         Top             =   0
         Width           =   11904
         Begin VB.TextBox FloatingText 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   8520
            TabIndex        =   7
            Top             =   120
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox FloatingTextDefault 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   8760
            TabIndex        =   6
            Text            =   "Default Text Setting"
            Top             =   120
            Visible         =   0   'False
            Width           =   2175
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
            Left            =   2400
            MousePointer    =   3  'I-Beam
            TabIndex        =   5
            Top             =   760
            Width           =   6495
         End
         Begin ClearLineProfiler.PVDGraphControl PVGraph 
            Height          =   2475
            Index           =   0
            Left            =   480
            TabIndex        =   4
            Top             =   1560
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   4577
         End
         Begin ClearLineProfiler.PVDGraphControl PVGraph 
            Height          =   2595
            Index           =   1
            Left            =   480
            TabIndex        =   17
            Top             =   4320
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   4577
         End
         Begin ClearLineProfiler.PVDGraphControl PVGraph 
            Height          =   2595
            Index           =   2
            Left            =   480
            TabIndex        =   18
            Top             =   7080
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   4577
         End
         Begin ClearLineProfiler.PVDGraphControl PVGraph 
            Height          =   2595
            Index           =   3
            Left            =   480
            TabIndex        =   19
            Top             =   9840
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   4577
         End
         Begin ClearLineProfiler.PVDGraphControl PVGraph 
            Height          =   2595
            Index           =   4
            Left            =   480
            TabIndex        =   20
            Top             =   12600
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   4577
         End
         Begin VB.Label DirectoryLbl 
            Caption         =   "Directory Label"
            Height          =   255
            Left            =   480
            TabIndex        =   21
            Top             =   1200
            Width           =   8295
            WordWrap        =   -1  'True
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
            TabIndex        =   16
            Top             =   14880
            Width           =   1455
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
            TabIndex        =   15
            Top             =   14880
            Width           =   855
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
            TabIndex        =   14
            Top             =   15120
            Width           =   1575
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
            TabIndex        =   13
            Top             =   15120
            Width           =   375
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
            TabIndex        =   12
            Top             =   15120
            Width           =   2295
         End
         Begin VB.Line HeaderBreakLine 
            X1              =   480
            X2              =   11040
            Y1              =   1500
            Y2              =   1500
         End
         Begin VB.Line FooterBreakLine 
            X1              =   480
            X2              =   11040
            Y1              =   14760
            Y2              =   14760
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
            TabIndex        =   10
            Top             =   15120
            Width           =   855
         End
         Begin VB.Label lblTitle 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "Ovality Flat Project Report"
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
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   10815
         End
         Begin VB.Image LogoImage 
            Height          =   690
            Left            =   360
            Picture         =   "PVReport2in1.frx":0000
            Stretch         =   -1  'True
            Top             =   480
            Width           =   2010
         End
         Begin VB.Image CLPLogoImage 
            Height          =   705
            Left            =   600
            Picture         =   "PVReport2in1.frx":0865
            Stretch         =   -1  'True
            Top             =   14880
            Width           =   2010
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
            TabIndex        =   8
            Top             =   14880
            Width           =   5055
         End
      End
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
         TabIndex        =   2
         Tag             =   "Paper"
         Top             =   0
         Visible         =   0   'False
         Width           =   5535
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
      Index           =   3
      Left            =   4320
      MousePointer    =   1  'Arrow
      Picture         =   "PVReport2in1.frx":10CA
      Tag             =   "Print"
      ToolTipText     =   "Print"
      Top             =   0
      Width           =   720
   End
   Begin VB.Image CloseReport 
      Height          =   480
      Left            =   12840
      MousePointer    =   1  'Arrow
      Picture         =   "PVReport2in1.frx":2D94
      ToolTipText     =   "Close Report"
      Top             =   150
      Width           =   480
   End
   Begin VB.Image ControlHighlight 
      Height          =   570
      Left            =   10680
      Picture         =   "PVReport2in1.frx":3A5E
      Top             =   90
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Image ControlHighlightDepressed 
      Height          =   585
      Left            =   11760
      Picture         =   "PVReport2in1.frx":55F0
      Top             =   90
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Image ControlsReport 
      Height          =   720
      Index           =   4
      Left            =   2640
      MousePointer    =   1  'Arrow
      Picture         =   "PVReport2in1.frx":723A
      Tag             =   "ZoomOut"
      ToolTipText     =   "Zoom Out"
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsReport 
      Height          =   720
      Index           =   2
      Left            =   1800
      MousePointer    =   1  'Arrow
      Picture         =   "PVReport2in1.frx":8F04
      Tag             =   "ZoomIn"
      ToolTipText     =   "Zoom In"
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsReport 
      Height          =   720
      Index           =   1
      Left            =   960
      MousePointer    =   1  'Arrow
      Picture         =   "PVReport2in1.frx":ABCE
      Tag             =   "Move"
      ToolTipText     =   "Move Report or Text"
      Top             =   0
      Width           =   720
   End
   Begin VB.Image ControlsBackPanel 
      Height          =   615
      Left            =   0
      MousePointer    =   1  'Arrow
      Picture         =   "PVReport2in1.frx":C898
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13335
   End
End
Attribute VB_Name = "PVReport2in1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public picReportPage As PictureBox

Public NumberOfExtraPages As Integer
Public TotalNumberOfGraphs As Integer
Public NGPP As Integer 'Number of graphs per page
Public GraphLength As Single
Public Units As String



Private PVDfilesFound() As String

Sub PrintPVReport2in1K()
On Error GoTo Err_Handler
    Dim I As Integer
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
        If (GraphIndex Mod NGPP = 0) And (GraphIndex > 0) Then
            Call RestoreOriginalState
            
            Call MarkForPrinting
            Me.PageLabel.Caption = "Page: " & 1 + Fix(GraphIndex / NGPP) & "/" & (1 + NumberOfExtraPages)
            Set picReportPage = picReportPagePg1
            
            Printer.NewPage
            Call RenderToPrinter.RenderReport(Me, Printer, 1)
            Call RenderToPrinter.RenderSingleTextBox(Me.UserTitle, Printer, False) 'PCN4277
        End If
        
        If GraphIndex > NGPP - 1 Then
            Me.PVGraph(GraphIndex).width = PVGraph(GraphIndex - (NGPP)).width
            Me.PVGraph(GraphIndex).height = PVGraph(GraphIndex - (NGPP)).height
            Me.PVGraph(GraphIndex).Top = PVGraph(GraphIndex - (NGPP)).Top
            Me.PVGraph(GraphIndex).Left = PVGraph(GraphIndex - (NGPP)).Left
        End If
        
        Call PVGraph(GraphIndex).PrintGraph(Printer, 1, Me.PVGraph(GraphIndex).Left, Me.PVGraph(GraphIndex).Top)
    

    
    
    Next GraphIndex
        
    For I = 1 To Me.FloatingText.Count - 1 'PCN4412
        Call RenderToPrinter.RenderSingleTextBox(Me.FloatingText(I), Printer, True)
    Next I

    Call Printer.EndDoc
    ScreenDrawingType = 0
    ScreenDrawingOrientation = 0
    
    Me.RestoreOriginalState
    Call Me.InitialiseFrom
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R2in11:" & Error$
    End Select
End Sub



Private Sub CloseReport_Click()
On Error GoTo Err_Handler

Unload Me
    
Exit Sub
Err_Handler:
    MsgBox Err & "-R2in12:" & Error$
End Sub

Private Sub ControlsBackPanel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

If Me.ControlHighlight.Visible Then Me.ControlHighlight.Visible = False
If Me.ControlHighlightDepressed.Visible Then Me.ControlHighlightDepressed.Visible = False
    
Exit Sub
Err_Handler:
    MsgBox Err & "-R2in13:" & Error$
End Sub

Private Sub ControlsReport_Click(Index As Integer)
On Error GoTo Err_Handler

Call ExecuteReportButton(Me, Index)
    
Exit Sub
Err_Handler:
    MsgBox Err & "-R2in14:" & Error$
End Sub

Private Sub ControlsReport_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


Me.ControlHighlightDepressed.Left = Me.ControlHighlight.Left
Me.ControlHighlightDepressed.Visible = True
Me.ControlHighlight.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-R2in15:" & Error$
End Sub

Private Sub ControlsReport_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ControlHighlight.Left = Me.ControlsReport(Index).Left - 100
If Me.ControlHighlight.Visible = False Then Me.ControlHighlight.Visible = True
    
Exit Sub
Err_Handler:
    MsgBox Err & "-R2in16:" & Error$
End Sub

Private Sub ControlsReport_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ControlHighlight.Visible = True
Me.ControlHighlightDepressed.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-R2in17:" & Error$
End Sub



Private Sub FloatingText_Change(Index As Integer)
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingText_Change(Me, Index)
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R2in18:" & Error$
    End Select
End Sub

Private Sub FloatingText_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingText_KeyPress(Me, Index, KeyAscii)
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R2in19:" & Error$
    End Select
End Sub

Private Sub FloatingText_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingText_MouseDown(Me, Index, Button, Shift, X, Y)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R2in110:" & Error$
    End Select
End Sub

Private Sub FloatingText_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingText_MouseMove(Me, Index, Button, Shift, X, Y)
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R2in111:" & Error$
    End Select
End Sub

Private Sub FloatingTextBackgroundColourMenu_Click()
On Error GoTo Err_Handler
    
    Call RenderToPrinter.FloatingTextBackgroundColourMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R2in112:" & Error$
    End Select
End Sub

Private Sub FloatingTextDefaultMenu_Click()
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingTextDefaultMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R2in113:" & Error$
    End Select
End Sub

Private Sub FloatingTextDeleteAllMenu_Click()
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingTextDeleteAllMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R2in114:" & Error$
    End Select
End Sub

Private Sub FloatingTextDeleteMenu_Click()
On Error GoTo Err_Handler
    
    Call RenderToPrinter.FloatingTextDeleteMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R2in115:" & Error$
    End Select
End Sub

Private Sub FloatingTextFontMenu_Click()
On Error GoTo Err_Handler

    Call RenderToPrinter.FloatingTextFontMenu_Click(Me)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R2in116:" & Error$
    End Select
End Sub
Sub InitialiseFrom()
On Error GoTo Err_Handler
    Dim I As Integer
    
    Dim GraphHeight As Single
    Dim GraphWidth As Single
    
    Me.Left = 0
    Me.width = ClearLineProfilerV6.width - 200
    Me.Top = 0
    Me.height = ClearLineProfilerV6.height - 500
    
    I = Me.Controls.Count
    ReDim OriginalStateVisible(I)
    ReDim OriginalStateTag(I)
    ReDim OriginalStateLeft(I)
    ReDim OriginalStateTop(I)
    ReDim OriginalStateX1(I)
    ReDim OriginalStateY1(I)
    ReDim OriginalStateX2(I)
    ReDim OriginalStateY2(I)
    
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
    
    GraphHeight = (Me.FooterBreakLine.y1 - Me.HeaderBreakLine.y1 + 30) / NGPP
    GraphWidth = Me.picReportPagePg1.width - (480 * 2)
    
    For I = 0 To (NGPP - 1)
        PVGraph(I).height = GraphHeight
        PVGraph(I).Top = I * GraphHeight + Me.HeaderBreakLine.y1 + 30
        PVGraph(I).width = GraphWidth
        PVGraph(I).Left = 480
    Next I
    
    For I = 1 To Me.Controls.Count - 1
        OriginalStateVisible(I) = Me.Controls(I).Visible
        OriginalStateTag(I) = Me.Controls(I).Tag
        OriginalStateLeft(I) = Me.Controls(I).Left
        OriginalStateTop(I) = Me.Controls(I).Top
        OriginalStateX1(I) = Me.Controls(I).x1
        OriginalStateY1(I) = Me.Controls(I).y1
        OriginalStateX2(I) = Me.Controls(I).x2
        OriginalStateY2(I) = Me.Controls(I).y2
    Next I
    
    Call FillOutPrintForm
    Set PrintPreviewForm = Me
    
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
        Case Else: MsgBox Err & "-R2in117:" & Error$
    End Select
End Sub

Sub FindAllPVDInDirectory(ByVal Path As String)
On Error GoTo Err_Handler

Dim FileName As String
Dim SplitPath As String
Dim SplitName As String
Dim SplitExt As String

Dim I As Integer


FileName = Dir(Path & "*.PVD")
While FileName <> ""
    Call SplitFilePath(FileName, SplitPath, SplitName, SplitExt)
    If LCase(SplitExt) = "pvd" Then
        I = I + 1
        ReDim Preserve PVDfilesFound(I)
        PVDfilesFound(I) = Path & FileName
    End If
    FileName = Dir
Wend


Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R2in118:" & Error$
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
    
    If Confirm1KDialog.SpanOption(0).value = True Then GraphLength = SpanInc * 1
    If Confirm1KDialog.SpanOption(1).value = True Then GraphLength = SpanInc * 2
    If Confirm1KDialog.SpanOption(2).value = True Then GraphLength = SpanInc * 3
    If Confirm1KDialog.SpanOption(3).value = True Then GraphLength = SpanInc * 4
    If Confirm1KDialog.SpanOption(4).value = True Then GraphLength = SpanInc * 5
    If Confirm1KDialog.SpanOption(5).value = True Then GraphLength = SpanInc * 6
    
    
    NGPP = 5
 '   GraphLength = 50 'CSng(DebugForm.MetersPerGraphValue)
    
    Me.CmboPrinterList.Enabled = False
                                                 
    'Me.LogoImage.Picture = LoadPicture(WindowsTempDirectory & "CBS\EmbedFile.jpg")
    Me.LogoImage.Picture = LoadPicture(WindowsTempDirectory & "EmbedFile.jpg") 'ID4601

    Call SplitFilePath(PVDFileName, PVDPath, PVDName, PVDExt)
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
        Case Else: MsgBox Err & "-R2in119:" & Error$
        
    End Select
    
End Sub

Sub RenderForm()
On Error GoTo Err_Handler
    Dim I As Integer
    Dim CurrentNoOfExtraPages As Integer
    Dim GraphIndex As Integer
    Dim PageToRender As Integer
    
    
    NumberOfExtraPages = Fix((PVGraph.Count - 1) / NGPP)
    CurrentNoOfExtraPages = Me.picReportPageNth.Count
    
    For I = CurrentNoOfExtraPages To NumberOfExtraPages
        Call CreateNewPage(I)
    Next I
    
    picReportPagePg1.Cls
    picReportPagePg1.width = Printer.width * RenderScale
    picReportPagePg1.height = Printer.height * RenderScale
    
    For I = 1 To NumberOfExtraPages
        Me.picReportPageNth(I).Cls
        Me.picReportPageNth(I).width = Printer.width * RenderScale
        Me.picReportPageNth(I).height = Printer.height * RenderScale
    Next I
    
    Call RestoreOriginalState
    Call MarkForPrinting
    Me.PageLabel.Caption = "Page: 1/" & (1 + NumberOfExtraPages)
    Set picReportPage = picReportPagePg1
    Call RenderToPrinter.RenderReport(Me, picReportPage, RenderScale)




    For GraphIndex = 0 To TotalNumberOfGraphs
        If GraphIndex <= (NGPP - 1) Then
            Call PVGraph(GraphIndex).PrintGraph(Me.picReportPagePg1, 1, Me.PVGraph(GraphIndex).Left, Me.PVGraph(GraphIndex).Top)
        Else
            Me.PVGraph(GraphIndex).width = PVGraph(GraphIndex - (NGPP)).width
            Me.PVGraph(GraphIndex).height = PVGraph(GraphIndex - (NGPP)).height
            Me.PVGraph(GraphIndex).Top = PVGraph(GraphIndex - (NGPP)).Top
            Me.PVGraph(GraphIndex).Left = PVGraph(GraphIndex - (NGPP)).Left
            PageToRender = Fix(GraphIndex / (NGPP))
            Call PVGraph(GraphIndex).PrintGraph(Me.picReportPageNth(PageToRender), 1, Me.PVGraph(GraphIndex).Left, Me.PVGraph(GraphIndex).Top)
        
        
        End If
    Next GraphIndex
    
    Call CLPProgressBar.ProgressBarPosition(1#)

    For I = 1 To picReportPageNth.Count - 1
        Call RestoreOriginalState
        Call MarkForPrinting
        Me.PageLabel.Caption = "Page: " & 1 + I & "/" & (1 + NumberOfExtraPages)
        Set picReportPage = picReportPageNth(I)

        Call RenderToPrinter.RenderReport(Me, picReportPage, RenderScale)
    Next I
    
    picReportPagePg1.Visible = True
    For I = 1 To (picReportPageNth.Count - 1)
        Me.picReportPageNth(I).width = Me.picReportPagePg1.width
        Me.picReportPageNth(I).height = Me.picReportPagePg1.height
        Me.picReportPageNth(I).Top = (Me.picReportPagePg1.height + 300) * I
        Me.picReportPageNth(I).Left = Me.picReportPagePg1.Left
        Me.picReportPageNth(I).Visible = True
    Next I

        
    Me.picReportPageNth(0).Visible = False

    ScreenDrawingType = 0
    ScreenDrawingOrientation = 0

    Me.UserTitle.Visible = True

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R2in120:" & Error$
    End Select

End Sub



Sub LoadFoundPVDs(ByVal FileName As String, ByRef GraphIndex As Integer)
On Error GoTo Err_Handler
    Dim StartFrameDistance As Long
    Dim EndFrameDistance As Long
    Dim PageToRender As Integer
    Dim GraphTitle As String * 40
    Dim LimitString As String
    
    Dim SplitPath As String
    Dim SplitName As String
    Dim SplitExt As String
    
    If GraphIndex > (NGPP - 1) Then Call CreateNewGraph(GraphIndex)
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
    
'    If DebugForm.UseLimitTick.value = 1 Then
'        LimitString = Round(DebugForm.LimitValue, 2)
'    Else
        LimitString = Round(OvalityLimitL, 2)
'    End If
    
    PVGraph(GraphIndex).SetGraphTitle = GraphTitle & _
                                        Trim(PipelineInfo.AssetNo) & " - " & _
                                        Round(ExpectedDiameter, 2) & "mm - " & _
                                        Trim(PipelineInfo.Material) & " - " & _
                                        DisplayMessage("From") & " " & Trim(PipelineInfo.StartName) & _
                                        " " & DisplayMessage("To") & " " & Trim(PipelineInfo.FinishName) & " " & _
                                        "   (" & DisplayMessage("Red ovality limit") & ") = " & _
                                        LimitString & _
                                        "%)"

                                        
    PVGraph(GraphIndex).SetCommentCaption = DisplayMessage("Comments")
    PVGraph(GraphIndex).SetSecondGraphSate = True
    PVGraph(GraphIndex).SetGraphLength = GraphLength
    PVGraph(GraphIndex).SetGraphUnit = Units
    
                                        
                                      
    
    

    
    
    StartFrameDistance = PVDistances(1)
    EndFrameDistance = PVDistances(PVDataNoOfLines)
    
    If StartFrameDistance > EndFrameDistance Then StartFrameDistance = EndFrameDistance
    
    PVGraph(GraphIndex).SetStartDistance = StartFrameDistance
    PVGraph(GraphIndex).SetEndDistance = StartFrameDistance + GraphLength
    
    
    
    Do
        StartFrameDistance = Fix(PVGraph(GraphIndex).GetStartDistance)
        EndFrameDistance = Fix(PVGraph(GraphIndex).GetEndDistance)
        ScreenDrawingType = 2
        ScreenDrawingOrientation = 1
        Call PVGraph(GraphIndex).DrawPVGraphsReport
        ScreenDrawingType = 0
        ScreenDrawingOrientation = 0
        
        If EndFrameDistance - StartFrameDistance < GraphLength Then Exit Do
        GraphIndex = GraphIndex + 1
        If GraphIndex > (NGPP - 1) Then Call CreateNewGraph(GraphIndex)
        PVGraph(GraphIndex).SetHideInfo = True
        PVGraph(GraphIndex).SetSecondGraphSate = True
        PVGraph(GraphIndex).SetGraphLength = GraphLength
        PVGraph(GraphIndex).SetGraphUnit = Units
        
        
            
        PVGraph(GraphIndex).SetStartDistance = PVGraph(GraphIndex - 1).GetEndDistance
        PVGraph(GraphIndex).SetEndDistance = PVGraph(GraphIndex).GetStartDistance + GraphLength
    Loop
    GraphIndex = GraphIndex + 1
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R2in121:" & Error$
    End Select
End Sub

Sub CreateNewGraph(ByRef GraphIndex As Integer)
On Error GoTo Err_Handler

Dim PageNo As Integer

Load PVGraph(GraphIndex)
'PageNo = Me.picReportPageNth.Count

Me.PVGraph(GraphIndex).width = PVGraph(GraphIndex - (NGPP)).width
Me.PVGraph(GraphIndex).height = PVGraph(GraphIndex - (NGPP)).height
Me.PVGraph(GraphIndex).Top = PVGraph(GraphIndex - (NGPP)).Top
Me.PVGraph(GraphIndex).Left = PVGraph(GraphIndex - (NGPP)).Left

  
Exit Sub
Err_Handler:
    Select Case Err
        Case 360: Exit Sub 'Allready created
        Case Else: MsgBox Err & "-R2in122:" & Error$
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
        Case Else: MsgBox Err & "-R2in123:" & Error$
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
        Case Else: MsgBox Err & "-R2in124:" & Error$
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler
Dim FileSaveFail As Boolean

Call SaveToFilePipeObs(FileSaveFail)

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R2in125:" & Error$
    End Select
End Sub

Private Sub PageFramePictureBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


If Me.ControlHighlight.Visible Then Me.ControlHighlight.Visible = False
If Me.ControlHighlightDepressed.Visible Then Me.ControlHighlightDepressed.Visible = False
    
Exit Sub
Err_Handler:
    MsgBox Err & "-R2in126:" & Error$
End Sub

Private Sub picReportPagePg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Call ReportPageMouseDown(Me, picReportPagePg1, Button, Shift, X, Y) 'PCN4193

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R2in127:" & Error$
    End Select

End Sub

Private Sub picReportPagePg1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    Dim I As Integer
    
    If ReportMouseDown Then
        picReportPagePg1.Left = picReportPagePg1.Left + X - ReportMouseX
        picReportPagePg1.Top = picReportPagePg1.Top + Y - ReportMouseY
        For I = 1 To picReportPageNth.Count - 1
                picReportPageNth(I).Left = picReportPageNth(I).Left + X - ReportMouseX
                picReportPageNth(I).Top = picReportPageNth(I).Top + Y - ReportMouseY
        Next I
        
        
    End If
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R2in128:" & Error$
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
        Case Else: MsgBox Err & "-R2in129:" & Error$
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
        Case Else: MsgBox Err & "-R2in130:" & Error$
    End Select
    
End Sub

Sub GraphSpecificSettings()
On Error GoTo Err_Handler

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R2in131:" & Error$
    End Select
End Sub

Sub MarkForPrinting()
On Error GoTo Err_Handler

Dim I As Integer
Dim ControlType As String

'Draw renderings first that are marked back
For I = 1 To Me.Controls.Count - 1
    
    With Me.Controls(I)
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
Next I
Exit Sub
Err_Handler:
    Select Case Err
        Case 438: Resume No_Container
        Case Else: MsgBox Err & "-R2in132:" & Error$
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
    
    Dim I As Integer
 
    SelectPrinter = False
    For I = 0 To Printers.Count - 1
        ' if the specified printer is found, select it and return True
        If Printers(I).DeviceName = printer_name Then
            Set Printer = Printers(I)
            SelectPrinter = True
            Exit For
        End If
    Next I
    
Exit Function
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R2in133:" & Error$
    End Select
End Function
 

 

Private Sub CmboPrinterList_Click()
On Error GoTo Err_Handler
    Dim I As Integer
    
    Call SelectPrinter(CmboPrinterList.text)
    
    If Me.CmboPrinterList.Enabled = True Then
        Me.picReportPagePg1.Cls
        
        For I = 1 To Me.FloatingText.Count - 1 'We dont want the text to be rendered on the preview yet
            Me.FloatingText(I).Visible = False 'when changes printers
        Next I
        
        Me.RestoreOriginalState
        Me.InitialiseFrom
        
        For I = 1 To Me.FloatingText.Count - 1 'Even thou we dont want the text to be rendered, we still
            Me.FloatingText(I).Visible = True  'want to be able to see them after printer select changed
        Next I
    End If
    
    
Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R2in134:" & Error$
    End Select
End Sub



Sub RestoreOriginalState()
On Error GoTo Err_Handler
    Dim I As Long

    For I = 1 To Me.Controls.Count - 1
        If Me.Controls(I).name <> "FloatingText" Then
            Me.Controls(I).Visible = OriginalStateVisible(I)
            Me.Controls(I).Tag = OriginalStateTag(I)
            Me.Controls(I).Left = OriginalStateLeft(I)
            Me.Controls(I).Top = OriginalStateTop(I)
            Me.Controls(I).x1 = OriginalStateX1(I)
            Me.Controls(I).y1 = OriginalStateY1(I)
            Me.Controls(I).x2 = OriginalStateX2(I)
            Me.Controls(I).y2 = OriginalStateY2(I)
        End If
    Next I
Exit Sub
Err_Handler:
    Select Case Err
        Case 9 'Subscript out of range
            'Addition of text will cause this error
            Exit Sub
        Case 438, 382: Resume Next
        Case Else: MsgBox Err & "-R2in135:" & Error$
    End Select

End Sub




Private Sub picReportPageNth_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    
    Call ReportPageMouseDown(Me, picReportPagePg1, Button, Shift, X, Y) 'PCN4193

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R2in136:" & Error$
    End Select
End Sub

Private Sub picReportPageNth_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    Call picReportPagePg1_MouseMove(Button, Shift, X, Y)
    

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-R2in137:" & Error$
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
        Case Else: MsgBox Err & "-R2in138:" & Error$
    End Select
End Sub

Private Sub UserTitle_Change()
On Error GoTo Err_Handler

Dim FileSaveFail As Boolean

UserTitleProfile = Trim(Me.UserTitle.text) 'PCN4433

Exit Sub
Err_Handler:
    Select Case Err

        Case Else: MsgBox Err & "-R2in139:" & Error$
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





