VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm ClearLineProfilerV6 
   BackColor       =   &H8000000C&
   Caption         =   "Profiler"
   ClientHeight    =   5775
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6390
   Icon            =   "ClearLineProfilerV6.frx":0000
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   2430
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "ClearLineProfilerV6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub MDIForm_Load()




On Error GoTo Err_Handler
Dim strLine As String
    'PCN4920 added diameter flat colour
    'PCN2111 -------------------------------------------------v
    'Similiar routine exists in Main() Sub routine of Starup Routine .
    'But next codes required to initiate langauge loading.
    'ClearLineProfilerV6.Caption = "Profiler CCTV Ovality V" & App.Major & "." & App.Minor & "." & App.Revision & "bj"  'Still PCN6128 + PCN6178 (inclination + design gradient) 'PCN6258 and Japanese Bug fixes
    'ClearLineProfilerV6.Caption = "Profiler CCTV Ovality V" & App.Major & "." & App.Minor & "." & App.Revision & "bl"  'Still PCN6128 + PCN6178 (inclination + design gradient) 'PCN6258 and Japanese Bug fixes
    'ClearLineProfilerV6.Caption = "Profiler CCTV Ovality V" & App.Major & "." & App.Minor & "." & App.Revision & "bm"  'Still PCN6128 + PCN6178 (inclination + design gradient) 'PCN6258 and Japanese Bug fixes
    ClearLineProfilerV6.Caption = "Profiler CCTV Ovality V" & App.Major & "." & App.Minor & "." & App.Revision & "bo"  'overflow error found in laserlib dll, on vob 352 x 480 res, memory leaks removed, spike introduction fixed, added tuning bar feed back values
    
    Dim FileNo As Integer
    FileNo = FreeFile
    
    If Language = "" Then
        Language = "English"
        
    'ClearLineProfilerV6.Caption = ",Profiler  CCTV Ovality V" & App.Major & "." & App.Minor & "." & App.Revision & "bj" 'Still PCN6128 + PCN6178 (inclination + design gradient) 'PCN6258 and Japanese Bug fixes
    'ClearLineProfilerV6.Caption = ",Profiler  CCTV Ovality V" & App.Major & "." & App.Minor & "." & App.Revision & "bl" 'Still PCN6128 + PCN6178 (inclination + design gradient) 'PCN6258 and Japanese Bug fixes
    'ClearLineProfilerV6.Caption = ",Profiler  CCTV Ovality V" & App.Major & "." & App.Minor & "." & App.Revision & "bm" 'Still PCN6128 + PCN6178 (inclination + design gradient) 'PCN6258 and Japanese Bug fixes
    ClearLineProfilerV6.Caption = ",Profiler  CCTV Ovality V" & App.Major & "." & App.Minor & "." & App.Revision & "bo"   'overflow error found in laserlib dll, on vob 352 x 480 res, memory leaks removed, spike introduction fixed, added tuning bar feed back values
        
'PCN6523 remove inf file        If Dir(WindowsTempDirectory & "CBS\Clearline.inf") <> "" Then 'PCN2123, PCN2177 'PCN6471If Dir(App.Path & "\Clearline.inf") <> "" Then 'PCN2123, PCN2177
'PCN6523 remove inf file            Open WindowsTempDirectory & "CBS\Clearline.inf" For Input As #1 'PCN2123, PCN2177 'Open App.Path & "\Clearline.inf" For Input As #1 'PCN2123, PCN2177
'PCN6523 remove inf file            Line Input #1, MyFile
'PCN6523 remove inf file            Close #1
            
            
            'MyFile = WindowsTempDirectory & "CBS\Clearline.ini" 'PCN6471 MyFile = App.Path & "\Clearline.ini"
            MyFile = WindowsTempDirectory & "Clearline.ini" 'ID4601 'PCN6471 MyFile = App.Path & "\Clearline.ini"
            If Dir(MyFile) <> "" Then
                Open MyFile For Input As #FileNo
                While Not EOF(FileNo)
                    Line Input #FileNo, strLine
                    If Left(strLine, 9) = "Language=" Then
                        Language = Right(strLine, Len(strLine) - 9)
                    End If
                Wend
                Close #FileNo
            End If
            
'PCN6523 remove inf file        End If
    End If
    
    If Language <> "English" Then
    'Similiar routine exist in ConvertLanguage function.
    'But this routine is in MIDForm_Load Sub Routine to give proper error message to the user.
    'ConvertLanguage Function is called in all forms.
    'So ConvertLanguage Function does not give the same error message.
        'Get the filename for Language
        Dim strLang As String 'to store the text filename for the Language to load.
        Dim intPos As Integer 'PCN2169
        Dim blnEnd As Boolean
        
        FileNo = FreeFile
        If Dir(ReadOnlyAppPath & "Language\Languages.txt") <> "" Then 'PCN2123
            Open ReadOnlyAppPath & "Language\Languages.txt" For Input As #FileNo 'PCN2123
            blnEnd = False
            While Not EOF(FileNo) And Not blnEnd
                Line Input #FileNo, strLang
                intPos = InStr(strLang, ",")
                If intPos > 0 Then 'PCN2172
                    If Left(strLang, intPos - 1) = Language Then
                        strLang = Trim(Right(strLang, Len(strLang) - intPos)) 'E.G.) French.txt, FrehchEULA.rtf
                        intPos = InStr(strLang, ",")
                        If intPos > 0 Then
                            EULAFilename = Trim(Right(strLang, Len(strLang) - intPos)) 'E.G.) FrenchEULA.rtf
                            strLang = Trim(Left(strLang, intPos - 1)) 'E.G.) French.txt
                        End If
                        'PCN2167 7/8/03 by Abe -------------------v
                        intPos = InStr(EULAFilename, ",")
                        If intPos > 0 Then
                            HelpFilename = Trim(Right(EULAFilename, Len(EULAFilename) - intPos)) 'E.G.) FrenchHelpFile.chm
                            EULAFilename = Trim(Left(EULAFilename, intPos - 1)) 'E.G.) FrenchEULA.rtf
                        End If '----------------------------------^
                        blnEnd = True
                    End If
                End If
            Wend
            Close #FileNo 'PCN2168 Close file before calling DisplayMessage.
            If Not blnEnd Then 'the filename for Language is not found in Languages.txt file.
                'PCN2168 DisplayMessage is not necessary if Language text file is not available.
                MsgBox "Languages.txt file does not contain necessary information to load the language" & "(" & Language & ") " & "you chose. Please edit this file first. The default Language(English) is loaded instead.", , "Clear Line Profiler"
            ElseIf Dir(ReadOnlyAppPath & "Language\" & strLang) = "" Then 'Check whether that file exists actually.'PCN2123
                MsgBox strLang & " file for the language(" & Language & ") you chose does not exist. Create this file first. The default language(English) is loaded.", , "Clear Line Profiler"
            End If
        Else
            MsgBox DisplayMessage("Languages.txt file does not exist. Please create this file first. The default language(English) is loaded."), , "Clear Line Profiler"
        End If
    End If
    ConvertLanguage Me, Language 'PCN2111
    '---------------------------------------------------------^
        
    WindowState = vbMaximized

Exit Sub
Err_Handler:
    Select Case Err
        Case Else
            MsgBox Err & "-CLP1:" & Error$
            Resume Next
    End Select
    
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer) 'PCN2397, PCN2391
On Error GoTo Err_Handler
    If IsOpen("Fisheye") Then
        Unload Fisheye
    End If
    
    Unload PVReport4in1
    Unload PVReportMultiProfilex3
    Unload PVReportProfile
    Unload PVReport4in1
    Unload PVReportSingle
    Unload ProfilerMessageBox
    
    Exit Sub
Err_Handler:
    Select Case Err
        Case Else
            MsgBox Err & "-CLP2:" & Error$
            Resume Next
    End Select
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'****************************************************************************************
'Name    : MDIForm_Unload
'Created : 7 November 2003, PCN2371
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    : The 3DFlat takes a long time to draw. If the application is closed by the user
'           during this process, the application will crash. This function is to cancel
'           the 3D Flat drawing process by the unloading of the application.
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler

''PCN3513 No longer background load (Antony, 12 may 2005)
''
''Flat3DCancel = True 'PCN2371
''BackgroundLoadCancel = True 'PCN2970
''DoEvents

Exit Sub
Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & "-CLP3:" & Error$
        Resume Next
End Select
End Sub

