Attribute VB_Name = "ReportingFunctions"
 Option Explicit


Sub SetupReportMouseIcon(ReportForm As Form, MouseIconID As Long)   'PCNGL021202
On Error GoTo Err_Handler
Dim curSelect As StdPicture

Set curSelect = LoadResPicture(MouseIconID, vbResIcon)
ReportForm.MousePointer = 99
ReportForm.MouseIcon = curSelect

Exit Sub
Err_Handler:
    Select Case Err
        Case 53 'File not found
            PVReport4in1.MousePointer = 2
            Resume Next
        Case 75 'File not found
            PVReport4in1.MousePointer = 2
            Resume Next
        Case Else
            MsgBox Err & "-RF1:" & Error$
    End Select
End Sub

Sub ExecuteReportButton(ReportForm As Form, Index As Integer)
On Error GoTo Err_Handler

Select Case ReportForm.ControlsReport(Index).Tag
    Case "DrawText"
        PrintPreviewAction = "DrawText"
        ReportForm.MousePointer = 3

    Case "Move"
        PrintPreviewAction = "MoveAll"
        Call SetupReportMouseIcon(ReportForm, 108)
    
    Case "SaveReportToPVD"
        If PVDFileName = "" Then Exit Sub 'PCN4552
        Call StoreReportToPVD(ReportForm)

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
        Call ReportForm.RenderForm
        Call ReportForm.PositionReportControls
        
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
        Call ReportForm.RenderForm
        Call ReportForm.PositionReportControls
    
    Case "Print"
        Select Case ReportForm.name
            Case "PVReportMultiProfilex3"
                Call PVReportMultiProfilex3.PrintMultiProfileReport
            
            Case "PVReport4in1"
                Call PVReport4in1.Print4in1Report
            
            Case "PVReportProfile"
                Call PVReportProfile.PrintPVProfileReport
                
            Case "PVReportSingle"
                Call PVReportSingle.PrintPVSingleReport
            Case "PVReport1K"
                PVReport1K.PrintPVReport1K
            Case "PVReport2in1"
                PVReport2in1.PrintPVReport2in1K
                
                
            
        End Select
    
End Select

    
Exit Sub
Err_Handler:
    MsgBox Err & "-RF2:" & Error$
End Sub

Sub GetPrinterList(ReportForm As Form)
On Error GoTo Err_Handler

Dim dev As Printer, Index As Integer, CurrentPrinter As Integer

CurrentPrinter = -1
Index = 0

For Each dev In Printers
    ReportForm.CmboPrinterList.AddItem dev.DeviceName
    If Printer.DeviceName = dev.DeviceName Then
        CurrentPrinter = Index
    End If
    Index = Index + 1
Next

If CurrentPrinter <> -1 Then
    ReportForm.CmboPrinterList.ListIndex = CurrentPrinter
End If

Exit Sub
Err_Handler:
    Select Case Err
        Case Else:     MsgBox Err & "-RF3:" & Error$
    End Select
 
End Sub

Sub StoreReportToPVD(ReportForm As Form) 'PCN3809
On Error GoTo Err_Handler
Dim ReportNumber As Integer
Dim ReportTitle As String
Dim ReportType As Integer
Dim NoStoredReports As Integer
Dim Paper As PictureBox

Dim I, J As Integer

'PCN4279 ''''''''
If ThisFileIsReadOnly(PVDFileName) And SoftwareConfiguration <> "Reader" Then
    'MsgBox DisplayMessage("Warning this PVD is Read ONLY. Unable to store report."), vbExclamation
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Warning this PVD is Read ONLY. Unable to store report."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    Exit Sub    '
End If          '
'''''''''''''''''

ReportForm.MousePointer = 11
NoStoredReports = UBound(StoredReportArray)
ReportNumber = NextStoredReportNumber

Select Case ReportForm.name
    Case "PVReportMultiProfilex3"
        ReportType = 1
        If ReportNumber = 0 Then Exit Sub
        ReportTitle = ReportForm.UserTitle.text
        ReportForm.UserTitle.Visible = False
        
        RenderToPrinter.RenderScale = 3
        Call PVReportMultiProfilex3.RestoreOriginalState
        PVReportMultiProfilex3.InitialiseForm
        
        Call RenderToPrinter.RenderSingleTextBox(PVReportMultiProfilex3.UserTitle, ReportForm.picReportPage1, False)
        If ScreenDrawing.ImageGraphState(0).GraphType = "Ovality" Then
            Call RenderToPrinter.RenderSingleLabel(PVReportMultiProfilex3.Explination, ReportForm.picReportPage1)
        End If
        
        
        For I = 1 To PVReportMultiProfilex3.FloatingText.Count - 1
            If PVReportMultiProfilex3.FloatingText(I).Container.name = "picReportPage1" Then
                Call RenderToPrinter.RenderSingleTextBox(PVReportMultiProfilex3.FloatingText(I), ReportForm.picReportPage1, True)
            End If
        Next I
        
        
        
        Call SaveReportImageToBMPFile(ReportForm.picReportPage1)
        Call PageFunctions.StoredReportStore(ReportNumber, ReportTitle, ReportType)
        For I = 1 To ReportForm.NumberOfExtraPages
            'Call RenderToPrinter.RenderSingleTextBox(PVReportMultiProfilex3.UserTitle, ReportForm.picReportNextPage(I), False)
            
            If ScreenDrawing.ImageGraphState(0).GraphType = "Ovality" Then
                Call RenderToPrinter.RenderSingleLabel(PVReportMultiProfilex3.Explination, ReportForm.picReportNextPage(I))
            End If
        
            
            For J = 1 To PVReportMultiProfilex3.FloatingText.Count - 1
                If PVReportMultiProfilex3.FloatingText(J).Container.name = "picReportNextPage" Then
                    If PVReportMultiProfilex3.FloatingText(J).Container.Index = I Then
                        Call RenderToPrinter.RenderSingleTextBox(PVReportMultiProfilex3.FloatingText(J), ReportForm.picReportNextPage(I), True)
                    End If
                End If
            Next J
            Call SaveReportImageToBMPFile(ReportForm.picReportNextPage(I))
            Call PageFunctions.StoredReportStore(ReportNumber, ReportTitle, ReportType)
        Next I
        

    
        RenderToPrinter.RenderScale = 1
        Call PVReportMultiProfilex3.RestoreOriginalState
        PVReportMultiProfilex3.InitialiseForm
        
        
        For I = 1 To PVReportMultiProfilex3.FloatingText.Count - 1
            If PVReportMultiProfilex3.FloatingText(I).Container.name = "picReportPage1" Then
                Call RenderToPrinter.RenderSingleTextBox(PVReportMultiProfilex3.FloatingText(I), ReportForm.picReportPage1, True)
            End If
        Next I
        
        For I = 1 To ReportForm.NumberOfExtraPages
                    For J = 1 To PVReportMultiProfilex3.FloatingText.Count - 1
                If PVReportMultiProfilex3.FloatingText(J).Container.name = "picReportNextPage" Then
                    If PVReportMultiProfilex3.FloatingText(J).Container.Index = I Then
                        Call RenderToPrinter.RenderSingleTextBox(PVReportMultiProfilex3.FloatingText(J), ReportForm.picReportNextPage(I), True)
                    End If
                End If
            Next J
        Next I
        
        
        For I = 1 To PVReportMultiProfilex3.FloatingText.Count - 1
            PVReportMultiProfilex3.FloatingText(I).Visible = False 'PCN4531 instead of unload which is moved to form:unload
        Next I

    Case "PVReport4in1"
        ReportType = 2
        If ReportNumber = 0 Then Exit Sub
        ReportTitle = ReportForm.UserTitle.text
        ReportForm.UserTitle.Visible = False
        PVReport4in1.CommentsTextBox(0).Visible = False
        
        If ScreenDrawing.ImageGraphState(0).GraphType = "Ovality" Then
            Call RenderToPrinter.RenderSingleLabel(PVReport4in1.Explination(0), ReportForm.picReportPagePg1)
        End If
        
        For I = 1 To PVReport4in1.FloatingText.Count - 1
            If PVReport4in1.FloatingText(I).Container.name = "picReportPagePg1" Then
                Call RenderToPrinter.RenderSingleTextBox(PVReport4in1.FloatingText(I), ReportForm.picReportPagePg1, True)
            End If
        Next I
        
        Call RenderToPrinter.RenderSingleTextBox(PVReport4in1.UserTitle, ReportForm.picReportPagePg1, False)
        Call RenderToPrinter.RenderSingleTextBox(PVReport4in1.CommentsTextBox(0), ReportForm.picReportPagePg1, True)
        
        Call SaveReportImageToBMPFile(ReportForm.picReportPagePg1)
        Call PageFunctions.StoredReportStore(ReportNumber, ReportTitle, ReportType)
        
        
        If ReportForm.NoOfGraphPanels > 3 Then
            PVReport4in1.UserTitle.Left = PVReport4in1.lblTitle(1).Left
            PVReport4in1.UserTitle.width = PVReport4in1.lblTitle(1).width
            
            For I = 1 To PVReport4in1.FloatingText.Count - 1
                If PVReport4in1.FloatingText(I).Container.name = "picReportPagePg2" Then
                    Call RenderToPrinter.RenderSingleTextBox(PVReport4in1.FloatingText(I), ReportForm.picReportPagePg2, True)
                End If
            Next I
            
            Call RenderToPrinter.RenderSingleTextBox(PVReport4in1.UserTitle, ReportForm.picReportPagePg2, False)
            Call SaveReportImageToBMPFile(ReportForm.picReportPagePg2)
            Call PageFunctions.StoredReportStore(ReportNumber, ReportTitle, ReportType)
            
            PVReport4in1.UserTitle.Left = PVReport4in1.lblTitle(0).Left
            PVReport4in1.UserTitle.width = PVReport4in1.lblTitle(0).width
        End If
            
            
            
        For I = 1 To PVReport4in1.FloatingText.Count - 1
            PVReport4in1.FloatingText(I).Visible = False 'PCN4531 instead of unload which is moved to form:unload
        Next I
        

    Case "PVReportProfile"
        ReportType = 3
        If ReportNumber = 0 Then Exit Sub
        ReportTitle = ReportForm.UserTitle.text
        ReportForm.UserTitle.Visible = False
        PVReportProfile.CommentsTextBox.Visible = False
        
        RenderToPrinter.RenderScale = 3
        Call PVReportProfile.RestoreOriginalState
        PVReportProfile.InitialiseForm
        
        If ScreenDrawing.ImageGraphState(0).GraphType = "Ovality" Then
            Call RenderToPrinter.RenderSingleLabel(PVReportProfile.Explination, ReportForm.picReportPage)
        End If
        
        For I = 1 To PVReportProfile.FloatingText.Count - 1
            Call RenderToPrinter.RenderSingleTextBox(PVReportProfile.FloatingText(I), ReportForm.picReportPage, True)
            'Unload PVReportProfile.FloatingText(I)
        Next I
        
        Call RenderToPrinter.RenderSingleLabel(PVReportProfile.ObservationsLabel, ReportForm.picReportPage) 'PCN4458
        Call RenderToPrinter.RenderSingleTextBox(PVReportProfile.CommentsTextBox, ReportForm.picReportPage, True)
        Call RenderToPrinter.RenderSingleTextBox(PVReportProfile.UserTitle, ReportForm.picReportPage, False)
        Call SaveReportImageToBMPFile(ReportForm.picReportPage)
        Call PageFunctions.StoredReportStore(ReportNumber, ReportTitle, ReportType)
    
        RenderToPrinter.RenderScale = 1
        Call PVReportProfile.RestoreOriginalState
        PVReportProfile.InitialiseForm
        For I = 1 To PVReportProfile.FloatingText.Count - 1
            Call RenderToPrinter.RenderSingleTextBox(PVReportProfile.FloatingText(I), ReportForm.picReportPage, True)
            PVReportProfile.FloatingText(I).Visible = False 'PCN4531 instead of unload which is moved to form:unload
        Next I
    
    Case "PVReportSingle"
        ReportType = 4
        If ReportNumber = 0 Then Exit Sub
        ReportTitle = ReportForm.UserTitle.text
        ReportForm.UserTitle.Visible = False
        PVReportSingle.CommentsTextBox.Visible = False
        
        If ScreenDrawing.ImageGraphState(0).GraphType = "Ovality" Then
            Call RenderToPrinter.RenderSingleLabel(PVReportSingle.Explination, ReportForm.picReportPage)
        End If
        
        
        For I = 1 To PVReportSingle.FloatingText.Count - 1
            Call RenderToPrinter.RenderSingleTextBox(PVReportSingle.FloatingText(I), ReportForm.picReportPage, True)
            PVReportSingle.FloatingText(I).Visible = False 'PCN4531 instead of unload which is moved to form:unload
        Next I
        
        Call RenderToPrinter.RenderSingleTextBox(PVReportSingle.CommentsTextBox, ReportForm.picReportPage, True)
        Call RenderToPrinter.RenderSingleTextBox(PVReportSingle.UserTitle, ReportForm.picReportPage, False)
        Call SaveReportImageToBMPFile(ReportForm.picReportPage)
        Call PageFunctions.StoredReportStore(ReportNumber, ReportTitle, ReportType)
End Select

ReportForm.MousePointer = 0
'MsgBox DisplayMessage("Report saved to PVD"), vbInformation
ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Report saved to PVD"): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
        


Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RF4:" & Error$
    End Select






ReportForm.MousePointer = 0
End Sub

Sub SaveReportImageToBMPFile(ReportPageImage As PictureBox)
On Error GoTo Err_Handler

On Error Resume Next
'MkDir WindowsTempDirectory & "CBS" 'ID4601
Kill WindowsTempDirectory & EmbedBMPFileNameAndPath
On Error GoTo Err_Handler

SavePicture ReportPageImage.Image, WindowsTempDirectory & EmbedBMPFileNameAndPath
    
With PipelineDetails.JPGMake1
    .InputFile = WindowsTempDirectory & EmbedBMPFileNameAndPath
'    .Quality = 80
    .Quality = 100
    .OutputFile = WindowsTempDirectory & EmbedJMPFileNameAndPath
    .Go
End With

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RF5:" & Error$
    End Select
End Sub

Function NextStoredReportNumber() As Integer
On Error GoTo Err_Handler
Dim ReportIndex As Integer
Dim NoStoredReports As Integer
Dim MaxReportNumber As Integer

NoStoredReports = UBound(StoredReportArray)

For ReportIndex = 1 To NoStoredReports
    If StoredReportArray(ReportIndex).ReportNumber > MaxReportNumber Then
        MaxReportNumber = StoredReportArray(ReportIndex).ReportNumber
    End If
Next ReportIndex

NextStoredReportNumber = MaxReportNumber + 1

Exit Function
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RF6:" & Error$
    End Select
End Function

Sub GetPipeDetailsLabels(ReportForm As Form)  'PCN4171
On Error GoTo Err_Handler

With PipelineDetails
    If ReportForm.name <> "PVReport4in1" Then
        ReportForm.AssetNoLabel.Caption = .AssetNo_lbl.Caption
        ReportForm.SiteIDLabel.Caption = .SiteID_lbl.Caption
        ReportForm.CityLabel.Caption = .City_lbl.Caption
        ReportForm.DateLabel.Caption = .Date_lbl.Caption
        ReportForm.StartNodeLabel.Caption = .StartNodeNo_lbl.Caption
        ReportForm.StartLocationLabel.Caption = .StNodeLoc_lbl.Caption
        ReportForm.FinishNodeLabel.Caption = .FinishNodeNo_lbl.Caption
        ReportForm.FinishLocationLabel.Caption = .FhNodeLoc_lbl.Caption
        ReportForm.PipeDiameterLabel.Caption = .InternalDiameterExpected_lbl.Caption
        ReportForm.PipeLengthLabel.Caption = .PipeLen_lbl.Caption
        ReportForm.PipeMaterialLabel.Caption = .Material_lbl.Caption
    Else
        ReportForm.AssetNoLabel(0).Caption = .AssetNo_lbl.Caption
        ReportForm.SiteIDLabel(0).Caption = .SiteID_lbl.Caption
        ReportForm.CityLabel(0).Caption = .City_lbl.Caption
        ReportForm.DateLabel(0).Caption = .Date_lbl.Caption
        ReportForm.StartNodeLabel(0).Caption = .StartNodeNo_lbl.Caption
        ReportForm.StartLocationLabel(0).Caption = .StNodeLoc_lbl.Caption
        ReportForm.FinishNodeLabel(0).Caption = .FinishNodeNo_lbl.Caption
        ReportForm.FinishLocationLabel(0).Caption = .FhNodeLoc_lbl.Caption
        ReportForm.PipeDiameterLabel(0).Caption = .InternalDiameterExpected_lbl.Caption
        ReportForm.PipeLengthLabel(0).Caption = .PipeLen_lbl.Caption
        ReportForm.PipeMaterialLabel(0).Caption = .Material_lbl.Caption
    End If
End With

Exit Sub
Err_Handler:
    Select Case Err
        Case Else: MsgBox Err & "-RF7:" & Error$
    End Select
End Sub


