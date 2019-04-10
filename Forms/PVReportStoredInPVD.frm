VERSION 5.00
Begin VB.Form PVReportStoredInPVD 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   10140
   ScaleWidth      =   14460
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PageFramePictureBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   14895
      Left            =   120
      ScaleHeight     =   14865
      ScaleWidth      =   17265
      TabIndex        =   3
      Top             =   720
      Width           =   17295
      Begin VB.PictureBox picReportPage 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   2535
         Index           =   0
         Left            =   120
         ScaleHeight     =   2475
         ScaleWidth      =   13515
         TabIndex        =   4
         Tag             =   "Paper"
         Top             =   120
         Width           =   13575
         Begin VB.Image StoredReportImage 
            Height          =   1815
            Index           =   0
            Left            =   360
            Stretch         =   -1  'True
            Top             =   240
            Visible         =   0   'False
            Width           =   3135
         End
      End
   End
   Begin VB.HScrollBar PageHScroll 
      Height          =   255
      Left            =   8640
      Max             =   2
      Min             =   1
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   240
      Value           =   1
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox CmboPrinterList 
      Height          =   315
      Left            =   5160
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Text            =   "Select a Printer"
      Top             =   120
      Width           =   3135
   End
   Begin VB.Image ControlsReport 
      Height          =   720
      Index           =   0
      Left            =   3480
      MousePointer    =   1  'Arrow
      Picture         =   "PVReportStoredInPVD.frx":0000
      Tag             =   "DeleteStoredReport"
      ToolTipText     =   "Delete Report"
      Top             =   -60
      Width           =   720
   End
   Begin VB.Label NoOfPagesLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Page 1 of 2"
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
      Left            =   8640
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      Top             =   30
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image CloseReport 
      Height          =   480
      Left            =   12960
      MousePointer    =   1  'Arrow
      Picture         =   "PVReportStoredInPVD.frx":1CCA
      ToolTipText     =   "Close Report"
      Top             =   90
      Width           =   480
   End
   Begin VB.Image ControlsReport 
      Height          =   720
      Index           =   4
      Left            =   2640
      MousePointer    =   1  'Arrow
      Picture         =   "PVReportStoredInPVD.frx":2994
      Tag             =   "ZoomOut"
      ToolTipText     =   "Zoom Out"
      Top             =   -60
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsReport 
      Height          =   720
      Index           =   3
      Left            =   4320
      MousePointer    =   1  'Arrow
      Picture         =   "PVReportStoredInPVD.frx":465E
      Tag             =   "Print"
      ToolTipText     =   "Print"
      Top             =   -60
      Width           =   720
   End
   Begin VB.Image ControlsReport 
      Height          =   720
      Index           =   2
      Left            =   1800
      MousePointer    =   1  'Arrow
      Picture         =   "PVReportStoredInPVD.frx":6328
      Tag             =   "ZoomIn"
      ToolTipText     =   "Zoom In"
      Top             =   -60
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ControlsReport 
      Height          =   720
      Index           =   1
      Left            =   960
      MousePointer    =   1  'Arrow
      Picture         =   "PVReportStoredInPVD.frx":7FF2
      Tag             =   "Move"
      ToolTipText     =   "Move"
      Top             =   -60
      Width           =   720
   End
   Begin VB.Image ControlHighlight 
      Height          =   570
      Left            =   10680
      Picture         =   "PVReportStoredInPVD.frx":9CBC
      Top             =   30
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Image ControlHighlightDepressed 
      Height          =   585
      Left            =   11760
      Picture         =   "PVReportStoredInPVD.frx":B84E
      Top             =   30
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Image ControlsBackPanel 
      Height          =   615
      Left            =   0
      MousePointer    =   1  'Arrow
      Picture         =   "PVReportStoredInPVD.frx":D498
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13335
   End
End
Attribute VB_Name = "PVReportStoredInPVD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ReportType As Integer
Dim PageWidth() As Single
Dim PageHeight() As Single

Dim PaintToX() As Single
Dim PaintToY() As Single
Dim PaintToWidth() As Single
Dim PaintToHeight() As Single
Dim PaintFromX() As Single
Dim PaintFromY() As Single
Dim PaintFromWidth() As Single
Dim PaintFromHeight() As Single


Dim SelectedReportIndex As Integer

Private Sub CloseReport_Click()
On Error GoTo Err_Handler

Unload Me
    
Exit Sub
Err_Handler:
    MsgBox Err & "-RPVD1:" & Error$
End Sub



Private Sub CmboPrinterList_Click()
On Error GoTo Err_Handler
    Dim NumberOfPages As Integer
    Dim I As Integer
    
    Call SelectPrinter(CmboPrinterList.text)
    Call SetupForStoredReport(SelectedReportIndex)

    
Exit Sub
Err_Handler:
    MsgBox Err & "-RPVD2:" & Error$

End Sub

Private Sub ControlsBackPanel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


If Me.ControlHighlight.Visible Then Me.ControlHighlight.Visible = False
If Me.ControlHighlightDepressed.Visible Then Me.ControlHighlightDepressed.Visible = False
    
Exit Sub
Err_Handler:
    MsgBox Err & "-RPVD3:" & Error$
End Sub

Private Sub ControlsReport_Click(Index As Integer)
On Error GoTo Err_Handler

Call ExecuteStoredReportButton(Index)
    
Exit Sub
Err_Handler:
    MsgBox Err & "-RPVD4:" & Error$
End Sub

Private Sub ControlsReport_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler


Me.ControlHighlightDepressed.Left = Me.ControlHighlight.Left
Me.ControlHighlightDepressed.Visible = True
Me.ControlHighlight.Visible = False

    
Exit Sub
Err_Handler:
    MsgBox Err & "-RPVD5:" & Error$
End Sub

Private Sub ControlsReport_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ControlHighlight.Left = Me.ControlsReport(Index).Left - 100

If Me.ControlHighlight.Visible = False Then Me.ControlHighlight.Visible = True

    
Exit Sub
Err_Handler:
    MsgBox Err & "-RPVD6:" & Error$
End Sub

Private Sub ControlsReport_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Me.ControlHighlight.Left = Me.ControlsReport(Index).Left - 100

If Me.ControlHighlight.Visible = False Then Me.ControlHighlight.Visible = True

    
Exit Sub
Err_Handler:
    MsgBox Err & "-RPVD7:" & Error$
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Dim I As Integer

    ReportType = 0 'Reset Report type
    SelectedReportIndex = 0
Call GetPrinterList(Me)


Me.Left = 0
Me.width = ClearLineProfilerV6.width - 200
Me.Top = 0
Me.height = ClearLineProfilerV6.height - 500

PageFramePictureBox.Left = 0
PageFramePictureBox.width = Me.width - PageFramePictureBox.Left
PageFramePictureBox.height = Me.height
PageFramePictureBox.Top = 650
Me.ControlsBackPanel.width = Me.width
Me.CloseReport.Left = Me.width - 750

Call ConvertLanguage(Me, Language) 'PCN4171

'Set mouse icon for move
PrintPreviewAction = "MoveAll"
Call SetupReportMouseIcon(Me, 108)
If SoftwareConfiguration = "Reader" Then Me.ControlsReport(0).Visible = False

Exit Sub
Err_Handler:
    If Err = 340 Then Resume Next

End Sub

Private Sub picReportPage_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
'    If PrintPreviewAction = "DrawText" Then
'        Call RenderToPrinter.FloatingTextAdd(Me, picReportPage(Index), Button, Shift, X, Y)
'    Else
'        ReportMouseDown = True
'    End If
'    ReportMouseX = X
'    ReportMouseY = Y

Call ReportPageMouseDown(Me, picReportPage(Index), Button, Shift, X, Y) 'PCN4193

Exit Sub
Err_Handler:
MsgBox Err & "-RPVD8:" & Error$
End Sub

Private Sub picReportPage_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Dim I As Integer


If ReportMouseDown Then
    For I = 0 To picReportPage.Count - 1
        picReportPage(I).Left = picReportPage(I).Left + X - ReportMouseX
        picReportPage(I).Top = picReportPage(I).Top + Y - ReportMouseY
    Next I
End If

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RPVD9:" & Error$
    End Select
End Sub

Private Sub picReportPage_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

ReportMouseDown = False
If PrintPreviewAction = "MoveAll" Then
    Call SetupReportMouseIcon(PVReportStoredInPVD, 108)
End If

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RPVD10:" & Error$
    End Select
End Sub

Sub SetupForStoredReport(ReportIndex As Integer)
On Error GoTo Err_Handler
Dim PageIndex As Integer

Dim ReportNumber As Integer
Dim NumberOfPages As Integer
Dim I As Integer
Dim LeftExtent As Long, RightExtent As Long, TopExtent As Long, BottomExtent As Long
Dim PageMargen As Single
Dim PageRatio As Double
Dim PasteHeightSize As Double

PageMargen = 600

If ReportIndex = 0 Then Exit Sub

ReportNumber = StoredReportArray(ReportIndex).ReportNumber
If ReportNumber = 0 Then Exit Sub

Call PageFunctions.StoredReportRetrieve(ReportNumber)

PageIndex = 1
NumberOfPages = PrecisionVisionGraph.ReportsPictureStorage.Count - 1

'Format page
ReportType = StoredReportArray(ReportIndex).ReportType
Select Case ReportType
    Case 1: Call SetPrinterPageSettings(vbPRORLandscape)  'PVReportMultiProfilex3
    Case 2: Call SetPrinterPageSettings(vbPRORPortrait) 'PVReport4in1 PCN4413
    Case 3: Call SetPrinterPageSettings(vbPRORPortrait) 'PVReportProfile
    Case 4: Call SetPrinterPageSettings(vbPRORLandscape) 'PVReportSingle
End Select


picReportPage(0).width = Printer.width * 2
picReportPage(0).height = Printer.height * 2

ReDim PaintToX(NumberOfPages)
ReDim PaintToY(NumberOfPages)
ReDim PaintToWidth(NumberOfPages)
ReDim PaintToHeight(NumberOfPages)
ReDim PaintFromX(NumberOfPages)
ReDim PaintFromY(NumberOfPages)
ReDim PaintFromWidth(NumberOfPages)
ReDim PaintFromHeight(NumberOfPages)

For I = 0 To NumberOfPages - 1

    If I > 0 Then
        Load picReportPage(I)
        Load StoredReportImage(I)
        picReportPage(I).Top = picReportPage(I - 1).Top + picReportPage(I - 1).height + 200
        picReportPage(I).Left = picReportPage(I - 1).Left
        picReportPage(I).Visible = True
        picReportPage(I).width = picReportPage(I - 1).width
        picReportPage(I).height = picReportPage(I - 1).height
    End If




    Set StoredReportImage(I).Picture = PrecisionVisionGraph.ReportsPictureStorage(I + 1)
    Call ScreenDrawing.GetPictureExtents(LeftExtent, _
                                         TopExtent, _
                                        RightExtent, _
                                        BottomExtent, _
                                        StoredReportImage(I).Picture)
    
    'StoredReportImage(I).width = picReportPage(0).width
    'StoredReportImage(I).width = picReportPage(0).width
    'StoredReportImage(I).height = picReportPage(0).height
    

    
    ReDim Preserve PageWidth(I)
    ReDim Preserve PageHeight(I)
    
    
    PageWidth(I) = picReportPage(I).width
    PageHeight(I) = picReportPage(I).height
    picReportPage(I).width = Printer.width * 2
    picReportPage(I).height = Printer.height * 2
    
    picReportPage(I).Cls
    
    PageRatio = (BottomExtent - TopExtent) / (RightExtent - LeftExtent)
    PasteHeightSize = (picReportPage(I).width - (PageMargen * 2)) * PageRatio
    
    PaintToX(I) = PageMargen
    PaintToY(I) = PageMargen
    PaintToWidth(I) = picReportPage(I).width - (PageMargen * 2)
    PaintToHeight(I) = PasteHeightSize
    PaintFromX(I) = LeftExtent * 15
    PaintFromY(I) = TopExtent * 15
    PaintFromWidth(I) = (RightExtent - LeftExtent) * 15
    PaintFromHeight(I) = (BottomExtent - TopExtent) * 15
    
    
    
    
    
    Call picReportPage(I).PaintPicture(PrecisionVisionGraph.ReportsPictureStorage(I + 1), _
                                        PaintToX(I), PaintToY(I), _
                                      PaintToWidth(I), _
                                       PaintToHeight(I), _
                                       PaintFromX(I), PaintFromY(I), _
                                       PaintFromWidth(I), PaintFromHeight(I))
                                
'       Call picReportPage(I).PaintPicture(PrecisionVisionGraph.ReportsPictureStorage(I + 1), PageMargen, PageMargen, _
'                                       picReportPage(I).width - (PageMargen * 2), _
'                                       PasteHeightSize, _
'                                       LeftExtent * 15, TopExtent * 15, _
'                                       (RightExtent - LeftExtent) * 15, (BottomExtent - TopExtent) * 15)
'

Next I

SelectedReportIndex = ReportIndex

Exit Sub
Err_Handler:
    Select Case Err
    Case 9: Exit Sub 'That one is not here, most likely been deleted PCN4561
    Case 360: Resume Next 'the page or image box is allready loaded
        Case Else: MsgBox Err & "-RPVD11:" & Error$
    End Select

End Sub

Sub SetPrinterPageSettings(ByVal Orientation)

On Error GoTo ManualOrientation

'Printer.Orientation = vbPRORLandscape
Printer.Orientation = Orientation
Printer.PrintQuality = vbPRPQHigh

Exit Sub

ManualOrientation:
On Error GoTo Err_Handler

Dim originalheight
Dim originalwidth

originalheight = Printer.height
originalwidth = Printer.width

'If printer page is allready landscape size then no need to set to landscape size
If Orientation = vbPRORLandscape Then
    If originalwidth > originalheight Then Exit Sub ' PCN4367
Else
    If originalheight > originalheight Then Exit Sub
End If
    
'Swap orientation of printer size
Printer.height = originalwidth
Printer.width = originalheight

Exit Sub
Err_Handler:
MsgBox Err & "-RPVD12:" & Error$

End Sub

Sub ExecuteStoredReportButton(Index As Integer)
On Error GoTo Err_Handler
    Dim I As Integer

Select Case Me.ControlsReport(Index).Tag
    Case "Move"
        PrintPreviewAction = "MoveAll"
        Call SetupReportMouseIcon(Me, 108)
    
    Case "ZoomIn"
        If RenderScale = 0.5 Then
            RenderScale = 0.7
        ElseIf RenderScale = 0.7 Then
            RenderScale = 0.9
        ElseIf RenderScale = 0.9 Then
            RenderScale = 1
        Else
            Exit Sub
        End If
        
    Case "ZoomOut"
        If RenderScale = 1 Then
            RenderScale = 0.9
        ElseIf RenderScale = 0.9 Then
            RenderScale = 0.7
        ElseIf RenderScale = 0.7 Then
            RenderScale = 0.5
        Else
            Exit Sub
        End If
    
    Case "Print"
        If Me.picReportPage(0).width > Me.picReportPage(0).height Then
            Printer.Orientation = vbPRORLandscape
        Else
            Printer.Orientation = vbPRORPortrait
        End If
        
        For I = 0 To Me.picReportPage.Count - 1
            If I > 0 Then Call Printer.NewPage
            'Call Printer.PaintPicture(Me.picReportPage(i).Picture, 0, 0, Me.picReportPage(i).width, Me.picReportPage(i).height)
            'Call Printer.PaintPicture(Me.picReportPage(I).Picture, 0, 0, PageWidth(I), PageHeight(I))
            'Call Printer.PaintPicture(Me.StoredReportImage(I).Picture, 0, 0, PageWidth(I), PageHeight(I))
            Call Printer.PaintPicture(Me.StoredReportImage(I).Picture, _
                                      PaintToX(I), PaintToY(I), _
                                      PaintToWidth(I), _
                                      PaintToHeight(I), _
                                      PaintFromX(I) * 2, PaintFromY(I) * 2, _
                                      PaintFromWidth(I) * 2, PaintFromHeight(I) * 2)
        Next I
        
        Call Printer.EndDoc
        
        Me.PageFramePictureBox.Visible = True
        For I = 0 To Me.picReportPage.Count - 1
           Me.picReportPage(I).Visible = True
        Next I
    Case "DeleteStoredReport"
        Call DeleteSelectedStoredReport
    
End Select

    
Exit Sub
Err_Handler:
    MsgBox Err & "-RPVD3:" & Error$

End Sub

Sub DeleteSelectedStoredReport()
On Error GoTo Err_Handler
Dim Resp As Variant
Dim MsgStr As String
Dim ReportNumber As Integer

If SelectedReportIndex = 0 Then Exit Sub

MsgStr = DisplayMessage("Delete report") & " " & StoredReportArray(SelectedReportIndex).Title & "?"
Resp = MsgBox(MsgStr, vbExclamation + vbYesNo)
If Resp = vbYes Then
    Me.PageFramePictureBox.MousePointer = 11
    ReportNumber = StoredReportArray(SelectedReportIndex).ReportNumber
    If ReportNumber = 0 Then Exit Sub
    
    Call PageFunctions.StoredReportDelete(ReportNumber)
    
    Call ControlsScreen.SetupForStoredReports
    Me.PageFramePictureBox.MousePointer = 0
    Unload Me
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-RPVD14:" & Error$
End Sub

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
        Case Else: MsgBox Err & "-RPVD15:" & Error$
    End Select
End Function
