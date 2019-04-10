Attribute VB_Name = "ProgressIndicator"
'VBClassic   Rob Crombie   Rob@crombie.com
'Demonstrates how to use encapsulation and code re-use in VB6
'Only use VB's Intinsic controls ( no DLL Hell )
'PROGRESS BAR
'""""""""""""
'By adding this BAS file to your project, you get a progress bar.
'The sample Form in this project, does not need to be added to your project.
'In your calling Form -
'Call the Public Progress_Initialize and pass Form, PosTop and PosLeft
' The code in the BAS file will do all of the rest for you.
'   If you wish a smaller picturebox, then you may have to tweak
'   the Font sizes and message spacing in the BAS file.
'In your working routine (navigating a rs, or some other looping)
' call AdvanceProgressIndicator passing  Total As Long  and  Progress As Long
'When Total = Progress  the BAS file will remove the Progress Bar automatically.
'If you are unsure as to whether it will reach  Total = Progress
'     ( somer working loops may not pass that final value )
' then you can place this in the Sub that called your working routine -
'  Progress_Finalize


Option Explicit

Private fForm As Form
                                      
Private lngWidthContainer As Long
Private lngHeight As Long

Public Sub Progress_Initialize(passedFormReference As Form, PosTop, PosLeft)
  Set fForm = passedFormReference
  CreateControls fForm, PosTop, PosLeft
  ' Placing into BAS level variables, so not continually referring to Properties
  lngWidthContainer = fForm!fLabel.Width
  lngHeight = fForm!fLabel.Height
End Sub

Public Sub Progress_Finalize()
  On Error Resume Next
  fForm.Controls.Remove "fPicBoxBorder"
  Set fForm = Nothing
  On Error GoTo 0
End Sub

Public Sub AdvanceProgressIndicator(lngTotal As Long, lngProgress As Long)
 Dim myCurr As Currency
 Dim myLong As Long
 Dim lngPerecentAsWidth As Long
 Dim lngPercent As Long
 Static LastCalc As Long
  If lngTotal Mod 100 = 99 Then DoEvents 'Modified by Abe on 041103
  
  myCurr = (lngProgress / lngTotal) * 100
  lngPercent = Int(myCurr)
  'Only paint once for each percentage (eg 1.0 1.1 1.2 etc, only painted once)
  If lngPercent <> LastCalc Then
    LastCalc = lngPercent
    lngPerecentAsWidth = lngWidthContainer * (lngPercent / 100)
    'fForm!fShape.Move 0, 0, lngPerecentAsWidth, lngHeight
    fForm!fShape.Width = lngPerecentAsWidth
    'fForm!fShape.Refresh
'Debug.Print lngPercent
    fForm!fLabel.Caption = "RUNNING   " & lngPercent & "%"
    ' Ensure last few repaints are Refreshed immediately.
    If lngPercent > 94 Then
      fForm!fShape.Refresh
      fForm!fLabel.ForeColor = RGB(255, 0, 0)
      fForm!fLabel.Refresh
      fForm!fPicBoxBorder.BackColor = RGB(255, 0, 0)
      fForm!fPicBoxBorder.Refresh
      If lngPercent = 100 Then
        Progress_Finalize
      End If
    End If
  End If
End Sub

'                       CONTROL CREATION CODE                               '
 Private Sub CreateControls(fForm As Form, PosTop, PosLeft)
 fForm.Controls.Add "VB.PictureBox", "fPicBoxBorder", fForm
  With fForm!fPicBoxBorder
    .Top = PosTop
    .Left = PosLeft
    .Height = 515
    .Width = 4190
    .Appearance = 0   'Flat
    .BackColor = RGB(150, 150, 150)
    .Visible = True
    .ZOrder 0
    .Refresh
  End With
    
  fForm.Controls.Add "VB.PictureBox", "fPicBox", fForm!fPicBoxBorder
  With fForm!fPicBox
    .Top = 30
    .Left = 45
    .Height = fForm!fPicBoxBorder.Height - 105
    .Width = fForm!fPicBoxBorder.Width - 120
    .BackColor = RGB(255, 255, 255)   'White
    .Visible = True
    .Refresh
  End With
  
  fForm.Controls.Add "VB.Label", "fLabel", fForm!fPicBox
  With fForm!fLabel
    .Top = 0
    .Left = 0
    .Height = fForm!fPicBox.Height + 30
    .Width = fForm!fPicBox.Width
    .Appearance = 0   'Flat
    .Alignment = 2    'Center
    .BackStyle = 0    'Transparent
    '.Caption = "RUNNING    100%"
    .font = "Arial Black"
    .FontSize = 12
    .ForeColor = RGB(0, 0, 140)
    .Visible = True
    .Refresh
  End With
  
  fForm.Controls.Add "VB.Shape", "fShape", fForm!fPicBox
  With fForm!fShape
    .Top = 0
    .Left = 0
    .Height = fForm!fPicBox.Height + 30
    .BorderStyle = 0 'Transparent
    .FillStyle = 0 'Solid
    .FillColor = RGB(100, 220, 255)
    .Visible = True
    .Width = 0
    .Refresh
  End With
End Sub
 
