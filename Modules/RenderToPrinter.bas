Attribute VB_Name = "RenderToPrinter"
'PCN3569.........................................
Private Const WM_NCLBUTTONDOWN As Long = &HA1&
Private Const HTCAPTION As Long = 2&
Private Declare Function ReleaseCapture Lib "user32" () As Long
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, wParam As Any, lParam As Any) As Long
'.................................................

#If Win32 Then

   Private Declare Function SetBkMode Lib "gdi32" _
   (ByVal hdc As Long, ByVal nBkMode As Long) As Long

   Private iBKMode As Long

#Else

   Private Declare Function SetBkMode Lib "GDI" (ByVal hdc As Integer _
    , ByVal nBkMode As Integer) As Integer

   Private iBKMode As Integer

#End If

   Private Const TRANSPARENT = 1
   Private Const OPAQUE = 2

Dim Destination As Variant
Public RS As Single
Public OriginalStateVisible() As Boolean
Public OriginalStateTag() As Variant
Public OriginalStateLeft() As Single
Public OriginalStateTop() As Single
Public OriginalStateX1() As Single
Public OriginalStateY1() As Single
Public OriginalStateX2() As Single
Public OriginalStateY2() As Single
Public OriginalStateWidth() As Single
Public OriginalStateHeight() As Single



Public ReportMouseX As Single
Public ReportMouseY As Single
Public ReportMouseDown As Boolean
Public PrintPreviewAction As String
Public RenderScale As Single

#If Win32 Then
 Private Declare Function SendMessageAsLong Lib "user32" _
 Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
 ByVal wParam As Long, ByVal lParam As Long) As Long
 Private Declare Function SendMessageAsString Lib "user32" _
 Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
 ByVal wParam As Long, ByVal lParam As String) As Long

 Const EM_GETLINE = 196
 Const EM_GETLINECOUNT = 186
#Else
 Private Declare Function SendMessage% Lib "user" _
 (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam As Any)

 Const EM_GETLINE = &H400 + 20
 Const EM_GETLINECOUNT = &H400 + 10
#End If
 Const MAX_CHAR_PER_LINE = 80  ' Scale this to size of text box.


'Private WrapSpace(50) As String
'Private Const EM_GETLINE = &HC4


Public Sub RenderReport(FormToPrint As Form, ToWhere As Variant, ByVal RenScale As Single)
'****************************************************************************************
'Name    : RenderReportToPrinter
'Created : August 9 05, PCN3593
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    : draws the images,text boxes and labels from the form onto the Printer object
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler
Dim i As Integer
Dim ControlType As String

Set Destination = ToWhere
RS = RenScale
If RS = 0 Then RS = 1

Destination.FontTransparent = True
'Correctly sets the background mix mode to transparent
iBKMode = SetBkMode(Printer.hdc, TRANSPARENT)

Destination.DrawWidth = 2
'Draw renderings first that are marked back
For i = 0 To FormToPrint.Controls.Count - 1
    If FormToPrint.Controls(i).Tag = "Back" Then
    
        ControlType = TypeName(FormToPrint.Controls(i))
        Select Case ControlType
            Case "TextBox"
                Destination.ForeColor = 0
                Call RenderSingleTextBox(FormToPrint.Controls(i))
            Case "Label"
                Destination.ForeColor = 0
                Call RenderSingleLabel(FormToPrint.Controls(i))
            Case "Shape"
                Call RenderShape(FormToPrint.Controls(i))
            Case "Image"
                Call RenderImages(FormToPrint.Controls(i))
            Case "PictureBox"
                Call RenderPictureBox(FormToPrint.Controls(i))
            Case "Line"
                Call RenderLine(FormToPrint.Controls(i))
        End Select
    End If
Next i

Destination.FontTransparent = True

'Now the rest of the renderings as ordered
For i = 0 To FormToPrint.Controls.Count - 1
    If FormToPrint.Controls(i).Tag = "Visible" Or FormToPrint.Controls(i).Tag = "Paper" Then
        ControlType = TypeName(FormToPrint.Controls(i))
        Select Case ControlType
            Case "TextBox"
                Call RenderSingleTextBox(FormToPrint.Controls(i))
            Case "Label"
                Call RenderSingleLabel(FormToPrint.Controls(i))
            Case "Shape"
                Call RenderShape(FormToPrint.Controls(i))
            Case "Image"
                Call RenderImages(FormToPrint.Controls(i))
            Case "PictureBox"
                Call RenderPictureBox(FormToPrint.Controls(i))
            Case "Line"
                Call RenderLine(FormToPrint.Controls(i))
        End Select
    End If
Next i
'
'Call Printer.PaintPicture(imgPipeImage.Picture, imgPipeImage.Left, imgPipeImage.Top, , , 0, 0, imgPipeImage.width, imgPipeImage.height)
'Call Printer.PaintPicture(imgUSManhole.Picture, imgUSManhole.Left, imgUSManhole.Top)
'Call Printer.PaintPicture(imgDSManhole.Picture, imgDSManhole.Left, imgDSManhole.Top)
'
'Printer.DrawWidth = 1

Exit Sub
Err_Handler:
MsgBox Err & "-RTP1:" & Error$
 
End Sub


Public Function RenderSingleTextBox(ByVal Box As TextBox, Optional AltDestination, Optional Boxing)
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

Dim XOffset As Single
Dim YOffset As Single
Dim buffer As String
Dim ndx As Long

Dim WasWidth As Single
Dim WasHeight As Single
Dim WasTop As Single
Dim WasLeft As Single
Dim WasFontSize As Single

If RS = 0 Then RS = 1

If Not IsMissing(AltDestination) Then
    Set Destination = AltDestination
Else
    Call GetOffset(XOffset, YOffset, Box)
End If



XOffset = XOffset * RS
YOffset = YOffset * RS

WasWidth = Box.width
WasHeight = Box.height
WasLeft = Box.Left
WasTop = Box.Top
WasFontSize = Box.Font.Size


Box.Left = Box.Left * RS
Box.Top = Box.Top * RS

Box.width = Box.width * RS
Box.height = Box.height * RS

Destination.Font = Box.Font
Destination.FontName = Box.FontName
Destination.Font.Size = (Box.Font.Size) * RS
Destination.Font.Italic = Box.Font.Italic
Destination.Font.Bold = Box.Font.Bold
Destination.Font.Charset = LanguageCharset

Destination.Line ((Box.Left + XOffset), (Box.Top + YOffset))-((Box.Left + Box.width + XOffset), (Box.Top + Box.height + YOffset)), vbWhite, BF


If IsMissing(Boxing) Then
    Destination.Line ((Box.Left + XOffset), (Box.Top + YOffset))-((Box.Left + Box.width + XOffset), (Box.Top + Box.height + YOffset)), vbBlack, B
Else
    If Boxing = True Then
        Destination.Line ((Box.Left + XOffset), (Box.Top + YOffset))-((Box.Left + Box.width + XOffset), (Box.Top + Box.height + YOffset)), vbBlack, B
        XOffset = XOffset + Destination.TextWidth("r")
    End If
End If
    
'Destination.CurrentX = (Box.Left + XOffset + Destination.TextWidth("r")) * RS
'
'If Box.Alignment = 0 Then Destination.CurrentX = (Box.Left + XOffset) * RS
'If Box.Alignment = 1 Then Destination.CurrentX = (Box.Left + Box.width - Destination.TextWidth(Lbl.Caption) + XOffset) * RS
'If Box.Alignment = 2 Then Destination.CurrentX = (Box.Left + (Lbl.width / 2) - (Destination.TextWidth(Lbl.Caption) / 2) + XOffset) * RS
'

'Destination.CurrentY = (Box.Top + YOffset + Round(Box.height * 0.1)) * RS
'Destination.Print Box.text

ndx& = fGetLineCount&(Box)
For n& = 1 To ndx&
   buffer = fGetLine(Box, n& - 1)
   'Destination.CurrentX = (Box.Left + XOffset + Destination.TextWidth("r")) * RS
   
   If Box.Alignment = 0 Then Destination.CurrentX = (Box.Left + XOffset)
   If Box.Alignment = 1 Then Destination.CurrentX = (Box.Left + Box.width - Destination.TextWidth(buffer) + XOffset)
   If Box.Alignment = 2 Then Destination.CurrentX = (Box.Left + (Box.width / 2) - (Destination.TextWidth(buffer) / 2) + XOffset)
   
   Destination.CurrentY = (Box.Top + YOffset + Round(Box.height * 0.1) + Destination.TextHeight("r") * (n - 1))
   If (Destination.CurrentY + Destination.TextHeight(r)) > (Box.Top + YOffset + Box.height) Then Exit For
   
    If Destination.CurrentX < Destination.width And Destination.CurrentY < Destination.height - Destination.TextHeight(buffer) Then
           Destination.Print buffer       ' ...or print to the screen.
    End If

Next n&

Box.width = WasWidth
Box.height = WasHeight
Box.Left = WasLeft
Box.Top = WasTop
Box.Font.Size = WasFontSize

Exit Function
Err_Handler:
MsgBox Err & "-RTP2:" & Error$

End Function


Private Function FormatTextForLabel(ByRef Lbl As Label)
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
Destination.CurrentY = (Lbl.Top) ' * RS
TotalHeight = 0

While TotalHeight < (Lbl.height - 50)
    Destination.CurrentX = (Lbl.Left + 10) ' * RS
    OneLine = ParseOneLabelWidth(Lbl.width - 20, TotalText)
    TotalHeight = TotalHeight + Destination.TextHeight(Remark)
    If Lbl.Alignment = 1 Then
        Destination.CurrentX = (Lbl.Left + Lbl.width - 10 - Destination.TextWidth(OneLine)) ' * RS
    End If
    If Destination.CurrentX < Destination.width And Destination.CurrentY < Destination.height - Destination.TextHeight(OneLine) Then
        If Destination.CurrentX < Destination.width And Destination.CurrentY < Destination.height - Destination.TextHeight(Lbl.Caption) Then
            Destination.Print OneLine
        End If
        
    End If
Wend

Exit Function
Err_Handler:
MsgBox Err & "-RTP3:" & Error$
End Function

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
MsgBox Err & "-RTP4:" & Error$
End Function

Public Function RenderSingleLabel(ByVal Lbl As Label, Optional AltDestination, Optional RenderScale) 'PCN4458 added AltDestination to render observation label
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

Dim WasWidth As Single
Dim WasHeight As Single
Dim WasTop As Single
Dim WasLeft As Single
Dim WasFontSize As Single

If RS = 0 Then RS = 1


'PCN4458 ''''''''''''''''''''''''''''''''''''''''
If Not IsMissing(AltDestination) Then           '
    Set Destination = AltDestination   '
Else                                            '
    Call GetOffset(XOffset, YOffset, Lbl)       '
End If                                          '
'''''''''''''''''''''''''''''''''''''''''''''''''

WasWidth = Lbl.width
WasHeight = Lbl.height
WasLeft = Lbl.Left
WasTop = Lbl.Top
WasFontSize = Lbl.Font.Size

Lbl.width = Lbl.width * RS
Lbl.height = Lbl.height * RS
Lbl.Left = Lbl.Left * RS
Lbl.Top = Lbl.Top * RS

XOffset = XOffset * RS
YOffset = YOffset * RS


Destination.Font = Lbl.Font
Destination.FontName = Lbl.FontName
Destination.Font.Size = (Lbl.Font.Size) * RS
Destination.Font.Italic = Lbl.Font.Italic
Destination.Font.Bold = Lbl.Font.Bold
'Destination.ForeColor = Lbl.ForeColor
Destination.FillStyle = vbFSTransparent
Destination.FontTransparent = True
Destination.Font.Charset = LanguageCharset


'Destination.FontName = Lbl.FontName
'Destination.Font.Size = (Lbl.Font.Size) * RS


If Lbl.Alignment = 0 Then Destination.CurrentX = (Lbl.Left + XOffset)
If Lbl.Alignment = 1 Then Destination.CurrentX = (Lbl.Left + Lbl.width) - Destination.TextWidth(Lbl.Caption) + XOffset
If Lbl.Alignment = 2 Then Destination.CurrentX = Lbl.Left + (Lbl.width / 2) - (Destination.TextWidth(Lbl.Caption) / 2) + XOffset
Destination.CurrentY = (Lbl.Top + YOffset) '* RS

If Lbl.WordWrap = True And Destination.TextWidth(Lbl.Caption) > Lbl.width - 75 Then  'PCN4389
    Call FormatTextForLabel(Lbl)
Else
    If Destination.CurrentX < Destination.width And Destination.CurrentY < Destination.height - Destination.TextHeight(Lbl.Caption) Then
        Destination.Print Lbl.Caption
    End If

End If


Lbl.width = WasWidth
Lbl.height = WasHeight
Lbl.Left = WasLeft
Lbl.Top = WasTop
Lbl.Font.Size = WasFontSize


Exit Function
Err_Handler:
MsgBox Err & "-RTP5:" & Error$

End Function

Private Function RenderImages(ByVal Img As Image)
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

If Img.Picture = 0 Then Exit Function
If Img.Tag <> "Visible" Then Exit Function

Dim XOffset As Single
Dim YOffset As Single

Call GetOffset(XOffset, YOffset, Img)

'If img.name <> "imgUSManhole" And img.name <> "imgDSManhole" And img.name <> "imgPipeImage" Then
    Call Destination.PaintPicture(Img.Picture, (Img.Left + XOffset) * RS, (Img.Top + YOffset) * RS, (Img.width) * RS, (Img.height) * RS)
'End If
    If Img.BorderStyle = 1 Then Destination.Line ((Img.Left) * RS, (Img.Top) * RS)-((Img.Left + Img.width) * RS, (Img.Top + Img.height) * RS), vbBlack, B

Exit Function
Err_Handler:
MsgBox Err & "-RTP6:" & Error$
End Function

Private Function RenderPictureBox(ByVal pict As PictureBox)
'****************************************************************************************
'PCN:
'Name    : RenderPictureBoxes
'Created : January 16 2006
'Updated :
'Prg By  : Antony van Iersel
'Param   :
'Desc    : draws images, except for the pipe and manholes
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler

'If Not pict.Visible Then Exit Function
If pict.Tag = "Paper" Then Exit Function

Destination.Line ((pict.Left) * RS, (pict.Top) * RS)-((pict.Left + pict.width) * RS, (pict.Top + pict.height) * RS), vbBlack, B

If pict.Picture = 0 Then Exit Function
If Not pict.Visible Then Exit Function
pict.Visible = False
Call Destination.PaintPicture(pict.Picture, (pict.Left * RS), (pict.Top) * RS, (pict.width) * RS, (pict.height) * RS)

Exit Function
Err_Handler:
MsgBox Err & "-RTP7:" & Error$
End Function

Private Function RenderLine(ByVal DrawLine As line)
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

Call GetOffset(XOffset, YOffset, DrawLine)
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


Exit Function
Err_Handler:
MsgBox Err & "-RTP8:" & Error$

End Function

Private Function RenderShape(ByVal DrawShape As Shape)
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

Call GetOffset(XOffset, YOffset, DrawShape)
    
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





Exit Function
Err_Handler:
MsgBox Err & "-RTP9:" & Error$
    
End Function

           
Sub GetOffset(ByRef X As Single, ByRef Y As Single, object As Variant)
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
    If TypeName(object.Container) = "PictureBox" And object.Container.Tag <> "Paper" Then
        X = object.Container.Left
        Y = object.Container.Top
    Else
        X = 0
        Y = 0
    End If
    
Exit Sub
Err_Handler:
MsgBox Err & "-RTP10:" & Error$
End Sub

Public Sub FloatingTextAdd(FormPaper As Form, PageBox As PictureBox, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error GoTo Err_Handler
Dim CurrentControl As Control
    
    
    NumberOfTextBoxes = FormPaper.FloatingText.Count
    Load FormPaper.FloatingText(NumberOfTextBoxes)
    Set FormPaper.FloatingText(NumberOfTextBoxes).Container = PageBox
    
    With FormPaper.FloatingText(NumberOfTextBoxes)
        .Left = X
        .Top = Y
        
        .Visible = True
        .ZOrder 0
        .SetFocus
        .Tag = "Visible"
    End With
    
   
    
Exit Sub
Err_Handler:
    Select Case Err
        Case 9: NumberOfTextBoxes = 0: Resume Next
        Case Else: MsgBox Err & "-RTP11:" & Error$
        

    End Select
End Sub

Public Sub FloatingText_Change(FormPaper As Form, Index As Integer)
On Error GoTo Err_Handler
    Call SetTextBoxWidthAndHeight(FormPaper, FormPaper.FloatingText(Index))
    FormPaper.FloatingText(Index).Refresh
Exit Sub
Err_Handler:
    MsgBox Err & "-RTP12:" & Error$
End Sub


Public Sub FloatingText_KeyPress(FormPaper As Form, Index As Integer, KeyAscii As Integer)
On Error GoTo Err_Handler

    If KeyAscii = 13 Then
        If FormPaper.FloatingText(Index).text = "" Then Call FloatingTextDelete(FormPaper, Index) 'PCN4193
    End If
    
Exit Sub
Err_Handler:
    MsgBox Err & "-RTP13:" & Error$
End Sub

Private Sub FloatingTextDelete(FormPaper As Form, Index As Integer)
On Error GoTo Err_Handler
    Dim NumberOfTextBoxes As Integer

    NumberOfTextBoxes = FormPaper.FloatingText.Count - 1
    FormPaper.FloatingText(Index).Left = FormPaper.FloatingText(NumberOfTextBoxes).Left
    FormPaper.FloatingText(Index).Top = FormPaper.FloatingText(NumberOfTextBoxes).Top
    FormPaper.FloatingText(Index).text = FormPaper.FloatingText(NumberOfTextBoxes).text
    Unload FormPaper.FloatingText(NumberOfTextBoxes)

Exit Sub
Err_Handler:
    MsgBox Err & "-RTP14:" & Error$
End Sub

Public Sub FloatingTextMoveAll(FormPaper As Form, X As Single, Y As Single)
On Error GoTo Err_Handler
    Dim NumberOfTextBoxes As Integer
    Dim Count As Integer

    NumberOfTextBoxes = FormPaper.FloatingText.Count - 1
    For Count = 1 To NumberOfTextBoxes
        FormPaper.FloatingText(Count).Left = FormPaper.FloatingText(Count).Left + X
        FormPaper.FloatingText(Count).Top = FormPaper.FloatingText(Count).Top + Y
    Next Count
    
Exit Sub
Err_Handler:
    MsgBox Err & "-RTP15:" & Error$

End Sub
Public Sub FloatingTextMove(FormPaper As Form, Index As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    FormPaper.FloatingText(Index).Left = FormPaper.FloatingText(Index).Left + X
    FormPaper.FloatingText(Index).Top = FormPaper.FloatingText(Index).Top + Y
    'FormPaper.FloatingText(Index).left = X
    'FormPaper.FloatingText(Index).Top = Y
    
    
Exit Sub
Err_Handler:
    MsgBox Err & "-RTP16:" & Error$

End Sub
Sub FloatingTextResetAll(FormPaper As Form)
On Error GoTo Err_Handler
    Dim NumberOfTextBoxes As Integer
    Dim Count As Integer

    NumberOfTextBoxes = FormPaper.FloatingText.Count - 1
    For Count = 1 To NumberOfTextBoxes
        FormPaper.FloatingText(Count).BackColor = FormPaper.FloatingText(0).BackColor
        FormPaper.FloatingText(Count).BorderStyle = FormPaper.FloatingText(0).BorderStyle
        FormPaper.FloatingText(Count).Font = FormPaper.FloatingText(0).Font
        FormPaper.FloatingText(Count).FontBold = FormPaper.FloatingText(0).FontBold
        FormPaper.FloatingText(Count).FontItalic = FormPaper.FloatingText(0).FontItalic
        FormPaper.FloatingText(Count).FontName = FormPaper.FloatingText(0).FontName
        FormPaper.FloatingText(Count).FontSize = FormPaper.FloatingText(0).FontSize
        FormPaper.FloatingText(Count).FontStrikethru = FormPaper.FloatingText(0).FontStrikethru
        FormPaper.FloatingText(Count).FontUnderline = FormPaper.FloatingText(0).FontUnderline
        FormPaper.FloatingText(Count).ForeColor = FormPaper.FloatingText(0).ForeColor
        Call SetTextBoxWidthAndHeight(FormPaper, FormPaper.FloatingText(Count))
        FormPaper.FloatingText(Count).Refresh
        
    Next Count
Exit Sub
Err_Handler:
    MsgBox Err & "-RTP17:" & Error$
End Sub
Sub FloatingTextResetAllToDefault(FormPaper As Form)
On Error GoTo Err_Handler
    Dim NumberOfTextBoxes As Integer
    Dim Count As Integer

    NumberOfTextBoxes = FormPaper.FloatingText.Count - 1
    For Count = 0 To NumberOfTextBoxes
        FormPaper.FloatingText(Count).BackColor = FormPaper.FloatingTextDefault.BackColor
        FormPaper.FloatingText(Count).BorderStyle = FormPaper.FloatingTextDefault.BorderStyle
        FormPaper.FloatingText(Count).Font = FormPaper.FloatingTextDefault.Font
        FormPaper.FloatingText(Count).FontBold = FormPaper.FloatingTextDefault.FontBold
        FormPaper.FloatingText(Count).FontItalic = FormPaper.FloatingTextDefault.FontItalic
        FormPaper.FloatingText(Count).FontName = FormPaper.FloatingTextDefault.FontName
        FormPaper.FloatingText(Count).FontSize = FormPaper.FloatingTextDefault.FontSize
        FormPaper.FloatingText(Count).FontStrikethru = FormPaper.FloatingTextDefault.FontStrikethru
        FormPaper.FloatingText(Count).FontUnderline = FormPaper.FloatingTextDefault.FontUnderline
        FormPaper.FloatingText(Count).ForeColor = FormPaper.FloatingTextDefault.ForeColor
        Call SetTextBoxWidthAndHeight(FormPaper, FormPaper.FloatingText(Count))
        FormPaper.FloatingText(Count).Refresh
        
    Next Count
Exit Sub
Err_Handler:
    MsgBox Err & "-RTP18:" & Error$
End Sub
Public Sub FloatingTextDeleteAll(FormPaper As Form)
On Error GoTo Err_Handler
    Dim NumberOfTextBoxes As Integer
    Dim Count As Integer

    NumberOfTextBoxes = FormPaper.FloatingText.Count - 1
    For Count = 1 To NumberOfTextBoxes
        Call FloatingTextDelete(FormPaper, 1)
    Next Count
    
Exit Sub
Err_Handler:
    MsgBox Err & "-RTP19:" & Error$
End Sub

Public Sub FloatingTextHide(FormPaper As Form)
On Error GoTo Err_Handler
    Dim Count As Integer
    Dim NumberOfTextBoxes As Integer

    NumberOfTextBoxes = FormPaper.FloatingText.Count - 1
    For Count = 1 To NumberOfTextBoxes
        FormPaper.FloatingText(Count).Visible = False
    Next Count

Exit Sub
Err_Handler:
    MsgBox Err & "-RTP20:" & Error$
End Sub

Public Sub FloatingTextShow(FormPaper As Form)
On Error GoTo Err_Handler
    
    Dim Count As Integer
    Dim NumberOfTextBoxes As Integer

    NumberOfTextBoxes = FormPaper.FloatingText.Count - 1
    For Count = 1 To NumberOfTextBoxes
        FormPaper.FloatingText(Count).Visible = True
        FormPaper.FloatingText(Count).ZOrder 0
    Next Count

Exit Sub
Err_Handler:
    MsgBox Err & "-RTP21:" & Error$
End Sub



Public Sub FloatingText_KeyUp(FormPaper As Form, Index As Integer, KeyCode As Integer, Shift As Integer)
End Sub

Public Sub FloatingText_MouseDown(FormPaper As Form, Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    If Button = vbLeftButton Then
        If PrintPreviewAction = "MoveAll" Then
            FormPaper.FloatingText(Index).MousePointer = 99
            FormPaper.FloatingText(Index).MouseIcon = LoadResPicture(123, vbResIcon) 'grab text icon
            Call ReleaseCapture
            Call SendMessage(FormPaper.FloatingText(Index).hwnd, WM_NCLBUTTONDOWN, ByVal HTCAPTION, ByVal 0&)

        End If
    ElseIf Button = vbRightButton Then
        'The following three lines disables the default edit popup menu from http://www.devx.com/vb2themax/Tip/18376
        FormPaper.FloatingText(Index).Enabled = False ' disable the textbox
        DoEvents                            ' (this DoEvents seems to be optional)
        FormPaper.FloatingText(Index).Enabled = True  ' re-enable the control, so that it doesn't appear as grayed
        FormPaper.FloatingText(0).Tag = Index
        FormPaper.PopupMenu FormPaper.FloatingTextMenu 'show your custom menu
        Call FloatingTextResetAll(FormPaper)
    
    End If
Exit Sub
Err_Handler:
    MsgBox Err & "-RTP22:" & Error$
    
End Sub

Public Sub FloatingText_MouseMove(FormPaper As Form, Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    If PrintPreviewAction = "MoveAll" Then
        FormPaper.FloatingText(Index).MousePointer = 99
        FormPaper.FloatingText(Index).MouseIcon = LoadResPicture(122, vbResIcon) 'Move holding text icon
    Else
        FormPaper.FloatingText(Index).MousePointer = vbIbeam
    End If
Exit Sub
Err_Handler:
    MsgBox Err & "-RTP23:" & Error$

End Sub

Public Sub FloatingText_MouseUp(FormPaper As Form, Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

Exit Sub
Err_Handler:
    MsgBox Err & "-RTP24:" & Error$
End Sub

Public Sub FloatingTextBackgroundColourMenu_Click(FormPaper As Form)
On Error GoTo Err_Handler
    FormPaper.FloatingTextDialog.CancelError = True
    FormPaper.FloatingTextDialog.Flags = cdlCCRGBInit
    FormPaper.FloatingTextDialog.Color = FormPaper.FloatingText(0).BackColor
    FormPaper.FloatingTextDialog.ShowColor
    
    FormPaper.FloatingText(0).BackColor = FormPaper.FloatingTextDialog.Color
    Call FloatingTextResetAll(FormPaper)

Exit Sub
Err_Handler:
    Select Case Err
        Case 32755: Exit Sub 'Cancel
        Case Else: MsgBox Err & "-RTP25:" & Error$
    End Select
End Sub

Public Sub FloatingTextDefaultMenu_Click(FormPaper As Form)
    Call FloatingTextResetAllToDefault(FormPaper)
End Sub

Public Sub FloatingTextDeleteAllMenu_Click(FormPaper As Form)
On Error GoTo Err_Handler
    Dim Index As Integer
    
    Index = FormPaper.FloatingText(0).Tag
    Call FloatingTextDeleteAll(FormPaper)
Exit Sub
Err_Handler:
    MsgBox Err & "-RTP26:" & Error$
End Sub

Public Sub FloatingTextDeleteMenu_Click(FormPaper As Form)
On Error GoTo Err_Handler
    Dim Index As Integer
    
    Index = FormPaper.FloatingText(0).Tag
    Call FloatingTextDelete(FormPaper, Index)
Exit Sub
Err_Handler:
    MsgBox Err & "-RTP27:" & Error$
End Sub

Public Sub FloatingTextFontMenu_Click(FormPaper As Form)
On Error GoTo Err_Handler
    FormPaper.FloatingTextDialog.CancelError = True

  ' Set the Flags property
    FormPaper.FloatingTextDialog.Flags = cdlCFEffects Or cdlCFBoth
    FormPaper.FloatingTextDialog.FontName = FormPaper.FloatingText(0).Font.name
    FormPaper.FloatingTextDialog.FontSize = FormPaper.FloatingText(0).Font.Size
    FormPaper.FloatingTextDialog.FontBold = FormPaper.FloatingText(0).Font.Bold
    FormPaper.FloatingTextDialog.FontItalic = FormPaper.FloatingText(0).Font.Italic
    FormPaper.FloatingTextDialog.FontUnderline = FormPaper.FloatingText(0).Font.Underline
    FormPaper.FloatingTextDialog.FontStrikethru = FormPaper.FloatingText(0).FontStrikethru
    FormPaper.FloatingTextDialog.Color = FormPaper.FloatingText(0).ForeColor

    FormPaper.FloatingTextDialog.ShowFont
    
    FormPaper.FloatingText(0).Font.name = FormPaper.FloatingTextDialog.FontName
    FormPaper.FloatingText(0).Font.Size = FormPaper.FloatingTextDialog.FontSize
    FormPaper.FloatingText(0).Font.Bold = FormPaper.FloatingTextDialog.FontBold
    FormPaper.FloatingText(0).Font.Italic = FormPaper.FloatingTextDialog.FontItalic
    FormPaper.FloatingText(0).Font.Underline = FormPaper.FloatingTextDialog.FontUnderline
    FormPaper.FloatingText(0).FontStrikethru = FormPaper.FloatingTextDialog.FontStrikethru
    FormPaper.FloatingText(0).ForeColor = FormPaper.FloatingTextDialog.Color
Exit Sub
Err_Handler:
    Select Case Err
        Case 32755: Exit Sub 'Cancel
        Case Else: MsgBox Err & "-RTP28:" & Error$
    End Select
End Sub

Sub SetTextBoxWidthAndHeight(FormPaper As Form, FloatingText As TextBox)
On Error GoTo Err_Handler
    
    Dim TextString As String
    Dim DummyPictureBox As PictureBox
    
    Set DummyPictureBox = FormPaper.Controls.Add("vb.PictureBox", "DummyPictureBox")
    DummyPictureBox.Visible = False
    
    TextString = FloatingText.text & "ww"
    
    DummyPictureBox.Font.name = FloatingText.FontName
    DummyPictureBox.Font.Size = FloatingText.FontSize
    DummyPictureBox.Font.Bold = FloatingText.FontBold
    DummyPictureBox.Font.Italic = FloatingText.FontItalic
    DummyPictureBox.Font.Underline = FloatingText.FontUnderline
    DummyPictureBox.FontStrikethru = FloatingText.FontStrikethru

    FloatingText.width = DummyPictureBox.TextWidth(TextString)
    FloatingText.height = DummyPictureBox.TextHeight(TextString)
    
    Call FormPaper.Controls.Remove("DummyPictureBox")

Exit Sub
Err_Handler:
    MsgBox Err & "-RTP29:" & Error$
End Sub



Sub ReportPageMouseDown(ReportForm As Form, ReportPage As PictureBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
Dim i As Long

    If PrintPreviewAction = "DrawText" Then
        Call FloatingTextAdd(ReportForm, ReportPage, Button, Shift, X, Y)
        i = ReportForm.Controls.Count
        ReDim Preserve OriginalStateVisible(i)
        ReDim Preserve OriginalStateTag(i)
        OriginalStateVisible(i - 1) = ReportForm.Controls(i - 1).Visible
        OriginalStateTag(i - 1) = ReportForm.Controls(i - 1).Tag

        'reset to move
        PrintPreviewAction = "MoveAll" 'PCN4193
        Call SetupReportMouseIcon(ReportForm, 108) 'PCN4193
    Else
        ReportMouseDown = True
        Call SetupReportMouseIcon(ReportForm, 109)
    End If

ReportMouseX = X
ReportMouseY = Y

Exit Sub
Err_Handler:
    Select Case Err
        Case 438
            Resume Next
        Case Else: MsgBox Err & "-RTP30:" & Error$
    End Select
End Sub

Private Function fGetLine$(TxtBx As TextBox, LineNumber As Long)
On Error GoTo Err_Handler
' This function fills the buffer with a line of text
' specified by LineNumber from the text-box control.
' The first line starts at zero.
      
      byteLo% = MAX_CHAR_PER_LINE And (255)  '[changed 5/15/92]
      byteHi% = Int(MAX_CHAR_PER_LINE / 256) '[changed 5/15/92]
      buffer$ = Chr$(byteLo%) + Chr$(byteHi%) + Space$(MAX_CHAR_PER_LINE - 2)
      #If Win32 Then
        X = SendMessageAsString(TxtBx.hwnd, EM_GETLINE, LineNumber, buffer$)
      #Else
        X = SendMessage(TxtBx.hwnd, EM_GETLINE, LineNumber, buffer$)
      #End If
      fGetLine$ = Left$(buffer$, X)
Exit Function
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RTP31:" & Error$
    End Select


End Function

      
Private Function fGetLineCount&(TxtBx As TextBox)
On Error GoTo Err_Handler
' This function will return the number of lines
' currently in the text-box control.
' Setfocus method illegal while in resize event,
' so use global flag to see if called from there
' (or use setfocus before this function call in general case).
       
    #If Win32 Then
         lCount = SendMessageAsLong(TxtBx.hwnd, EM_GETLINECOUNT, 0, 0)
    #Else
         lCount = SendMessage(TxtBx.hwnd, EM_GETLINECOUNT, 0&, 0&)
    #End If
    fGetLineCount& = lCount
Exit Function
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RTP32:" & Error$
    End Select
End Function



