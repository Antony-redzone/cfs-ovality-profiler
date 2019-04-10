VERSION 5.00
Begin VB.UserControl CBS_DropDownBox 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BackStyle       =   0  'Transparent
   ClientHeight    =   5475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6090
   ForwardFocus    =   -1  'True
   Picture         =   "CBS_DropDownBox.ctx":0000
   ScaleHeight     =   5475
   ScaleWidth      =   6090
   Begin VB.CommandButton DropDownButton 
      Height          =   495
      Left            =   5280
      Picture         =   "CBS_DropDownBox.ctx":038A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.VScrollBar ItemScrollbar 
      Height          =   1875
      LargeChange     =   540
      Left            =   5280
      SmallChange     =   270
      TabIndex        =   0
      Top             =   540
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox DropdownPictureBox 
      Height          =   3375
      Left            =   0
      ScaleHeight     =   3315
      ScaleWidth      =   5235
      TabIndex        =   1
      Top             =   480
      Width           =   5295
      Begin VB.PictureBox ItemsPictureBox 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2655
         Left            =   0
         ScaleHeight     =   2655
         ScaleWidth      =   5295
         TabIndex        =   2
         Top             =   0
         Width           =   5295
         Begin VB.Label HighliteLabel 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   0
            TabIndex        =   4
            Top             =   1680
            Width           =   5295
         End
         Begin VB.Label Items 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   5295
         End
      End
   End
   Begin VB.Label ItemSelectedlabel 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "CBS_DropDownBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ItemCount As Integer
Private ItemCurrentSelect As Integer
Private ItemCurrentHighlited As Integer
Private Collapsed As Boolean
Private DropDownHeight As Single
Private DisableResize As Boolean
Public Event MouseMove()
Public Event MouseClick()
Public Event OnSelect()

Private Sub DropDownButton_Click()
On Error GoTo Error_handler

    If Collapsed = True Then
        Call ExpandDropDownBox
    Else
        Call CollapseDropDownBox
        Call Highlite(-1)
    End If
    
Exit Sub
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-DD1:" & Error$, vbExclamation
    End Select
End Sub

Private Sub CollapseDropDownBox()
On Error GoTo Error_handler

    Dim i As Single
   
    For i = DropDownHeight To ItemSelectedlabel.height Step -270
        DisableResize = True
        height = i + ItemSelectedlabel.height
        DropdownPictureBox.height = i
        ItemScrollbar.height = i
        DoEvents
        Sleep (10)
    Next i
    DropdownPictureBox.Visible = False
    ItemScrollbar.Visible = False
    height = ItemSelectedlabel.height
    DisableResize = False
    Collapsed = True
    RaiseEvent MouseClick

Exit Sub
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-DD2:" & Error$, vbExclamation
    End Select
End Sub


Private Sub ExpandDropDownBox()
On Error GoTo Error_handler

    Dim i As Single
    DropdownPictureBox.Visible = True
    If ItemScrollbar.Enabled = True Then ItemScrollbar.Visible = True
    DisableResize = True
    For i = ItemSelectedlabel.height To DropDownHeight Step 270

        height = i + ItemSelectedlabel.height
        DropdownPictureBox.height = i
        ItemScrollbar.height = i
        DoEvents
        
        Sleep (10)
    Next i
    DropdownPictureBox.height = DropDownHeight
    height = DropDownHeight + ItemSelectedlabel.height
    ItemScrollbar.height = DropDownHeight
    
    DisableResize = False

    Collapsed = False

Exit Sub
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-DD3:" & Error$, vbExclamation
    End Select
End Sub
Private Sub Highlite(ByVal Index As Integer)
On Error GoTo Error_handler

    Dim i As Integer
    
    If Index <> -1 Then
        HighliteLabel.Caption = Items(Index).Caption
        HighliteLabel.Top = Items(Index).Top
        ItemCurrentHighlited = Index
        ItemSelectedlabel.Caption = Items(Index)
    Else
        HighliteLabel.Top = ItemsPictureBox.height
        ItemCurrentHighlited = ItemCurrentSelect
        ItemSelectedlabel.Caption = Items(ItemCurrentSelect).Caption
    End If
    RaiseEvent MouseMove
       
Exit Sub
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-DD4:" & Error$, vbExclamation

    End Select
End Sub

Public Sub AddItem(ByVal LangString As String)
On Error GoTo Error_handler

    If ItemCount = 0 Then
        Items(0).Caption = LangString
        ItemSelectedlabel.Caption = LangString
        Items(0).height = ItemSelectedlabel.height
        ItemCount = 1
        ItemsPictureBox.height = ItemSelectedlabel.height * ItemCount
        DropDownHeight = ItemSelectedlabel.height * ItemCount
        HighliteLabel.Top = ItemsPictureBox.height
        ItemCurrentSelect = 0
        Exit Sub
    End If
    
    Load Items(ItemCount)
    Items(ItemCount).Caption = LangString
    Items(ItemCount).Top = ItemSelectedlabel.height * ItemCount

    
    Items(ItemCount).Visible = True
    ItemCount = ItemCount + 1
    ItemsPictureBox.height = ItemSelectedlabel.height * ItemCount
    HighliteLabel.Top = ItemsPictureBox.height
    DropDownHeight = ItemSelectedlabel.height * ItemCount
    
    If DropDownHeight > (ItemSelectedlabel.height * 4) Then
        DropDownHeight = ItemSelectedlabel.height * 4
    End If
    
    If ItemsPictureBox.height > DropDownHeight Then
        ItemScrollbar.Max = ItemsPictureBox.height - DropDownHeight
        ItemScrollbar.Enabled = True
    End If
    
Exit Sub
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-DD5:" & Error$, vbExclamation
    End Select
End Sub

Private Sub HighliteLabel_Click()
On Error GoTo Error_handler
    
    If DropDownButton.Enabled = False Then: Exit Sub
    ItemSelectedlabel.Caption = HighliteLabel.Caption
    ItemCurrentSelect = ItemCurrentHighlited
    Call Highlite(-1)
    Call CollapseDropDownBox
    RaiseEvent OnSelect

Exit Sub
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-DD6:" & Error$, vbExclamation
    End Select
End Sub

Private Sub ItemSelectedLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Error_handler
    
    If DropDownButton.Enabled = False Then: Exit Sub
    Call Highlite(-1)
    
Exit Sub
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-DD7:" & Error$, vbExclamation
    End Select
End Sub



Private Sub UserControl_Initialize()
On Error GoTo Error_handler

    DisableResize = False
    DropDownHeight = 0
    Collapsed = True
    ItemCount = 0
    
    ItemScrollbar.height = 0
    ItemScrollbar.Enabled = False
    
Exit Sub
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-DD8:" & Error$, vbExclamation
    End Select
End Sub



Private Sub Items_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Error_handler
    
    If DropDownButton.Enabled = False Then Exit Sub
    Call Highlite(Index)

Exit Sub
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-DD9:" & Error$, vbExclamation
    End Select
End Sub



Private Sub ItemScrollbar_Change()
On Error GoTo Error_handler
    
    If DropDownButton.Enabled = False Then Exit Sub
    Call ItemScrollbar_Scroll

Exit Sub
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-DD10:" & Error$, vbExclamation
    End Select
End Sub

Private Sub ItemScrollbar_Scroll()
On Error GoTo Error_handler

     ItemsPictureBox.Top = -ItemScrollbar.value

Exit Sub
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-DD11:" & Error$, vbExclamation
    End Select
End Sub


Private Sub UserControl_Resize()
On Error GoTo Error_handler
Dim i As Integer

If DisableResize = True Then Exit Sub

ItemSelectedlabel.width = width ' - DropdownButton.width
ItemSelectedlabel.height = height

DropDownButton.Left = width - DropDownButton.width
DropDownButton.height = height

DropdownPictureBox.width = ItemSelectedlabel.width

ItemScrollbar.Left = DropDownButton.Left
ItemScrollbar.width = DropDownButton.width

ItemsPictureBox.width = ItemSelectedlabel.width
ItemsPictureBox.Top = 0

ItemScrollbar.Top = ItemSelectedlabel.height
DropdownPictureBox.Top = ItemSelectedlabel.height
DropDownHeight = ItemSelectedlabel.height * ItemCount
DropdownPictureBox.height = DropDownHeight
ItemScrollbar.height = DropDownHeight

ItemsPictureBox.height = ItemSelectedlabel.height * (ItemCount)
For i = 0 To ItemCount - 1
    Items(i).height = ItemSelectedlabel.height
    Items(i).Top = ItemSelectedlabel.height * ItemCount
Next i

HighliteLabel.width = ItemSelectedlabel.width
HighliteLabel.height = ItemSelectedlabel.height
HighliteLabel.Left = ItemSelectedlabel.Left
HighliteLabel.Top = ItemsPictureBox.height

Exit Sub
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-DD12:" & Error$, vbExclamation
    End Select
End Sub

Sub HighliteItemSelected()
On Error GoTo Error_handler

Dim i As Integer

For i = 0 To ItemCount - 1
    If ItemSelectedlabel.Caption = Items(i).Caption Then ItemCurrentSelect = i
Next i

Exit Sub
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-DD13:" & Error$, vbExclamation
    End Select
End Sub

Property Get Font() As Font
On Error GoTo Error_handler
    Dim i As Integer
    
    Set Font = ItemSelectedlabel.Font
    For i = 0 To ItemCount - 1
        Items(i).Font = ItemSelectedlabel.Font
        Items(i).Font.Bold = ItemSelectedlabel.Font.Bold
        Items(i).Font.Size = ItemSelectedlabel.Font.Size
    Next i
    HighliteLabel.Font = ItemSelectedlabel.Font
    HighliteLabel.Font.Bold = ItemSelectedlabel.Font.Bold
    HighliteLabel.Font.Size = ItemSelectedlabel.Font.Size
    
Exit Property
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-DD14:" & Error$, vbExclamation
    End Select
End Property

Property Set Font(ByVal NewFont As Font)
On Error GoTo Error_handler

    Set ItemSelectedlabel.Font = NewFont

Exit Property
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-DD15:" & Error$, vbExclamation
    End Select
End Property

Property Get text() As String
On Error GoTo Error_handler

    text = ItemSelectedlabel.Caption
    
Exit Property
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-DD16:" & Error$, vbExclamation
    End Select
End Property

Property Let text(txt As String)
On Error GoTo Error_handler

    ItemSelectedlabel.Caption = txt
    Call HighliteItemSelected
    
Exit Property
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-DD17:" & Error$, vbExclamation
    End Select
End Property

Property Get Count() As Integer
On Error GoTo Error_handler

    Count = ItemCount
    
Exit Property
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-DD18:" & Error$, vbExclamation
    End Select
End Property

Property Get ItemSelected() As Integer
On Error GoTo Error_handler

    ItemSelected = ItemCurrentSelect

Exit Property
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-DD19:" & Error$, vbExclamation
    End Select
End Property

Property Get TextHighlited() As String
On Error GoTo Error_handler

    If ItemCurrentHighlited >= 0 And ItemCurrentHighlited < ItemCount Then
        TextHighlited = Items(ItemCurrentHighlited).Caption
    End If

Exit Property
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-DD20:" & Error$, vbExclamation
    End Select
End Property

Property Let Item(ByVal Index As Integer, ByVal ItemText As String)
On Error GoTo Error_handler
    
    If Index < 0 Or Index > ItemCount Then Exit Property
    Items(Index).Caption = ItemText
    If Index = ItemCurrentSelect Then ItemSelectedlabel = ItemText

Exit Property
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-DD21:" & Error$, vbExclamation
    End Select
End Property

Property Let Enabled(ByVal TrueFalse As Boolean)
On Error GoTo Error_handler
    DropDownButton.Enabled = TrueFalse
Exit Property
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-DD22:" & Error$, vbExclamation
    End Select
End Property

Property Let SelectItem(ByVal Index As Integer)
On Error GoTo Error_handler


    If Index < 0 Or Index > ItemCount Then Exit Property
    ItemCurrentSelect = Index
    Call Highlite(-1)
    

Exit Property
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-DD23:" & Error$, vbExclamation
    End Select
End Property


Public Sub Expand()
On Error GoTo Error_handler

Call ExpandDropDownBox

Exit Sub
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-DD24:" & Error$, vbExclamation
    End Select
End Sub

Public Sub Collapse()
On Error GoTo Error_handler

Call CollapseDropDownBox

Exit Sub
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-DD25:" & Error$, vbExclamation
    End Select
End Sub

Public Sub SetCharset(ByVal LanguageCharset As Integer)
On Error GoTo Error_handler
    Dim i As Integer
    
    ItemsPictureBox.Font.Charset = LanguageCharset
    ItemSelectedlabel.Font.Charset = LanguageCharset
    DropdownPictureBox.Font.Charset = LanguageCharset
    HighliteLabel.Font.Charset = LanguageCharset
    For i = 0 To ItemCount - 1
        Items(i).Font.Charset = LanguageCharset
    Next i
    
Exit Sub
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-DD26:" & Error$, vbExclamation
    End Select
End Sub

