Attribute VB_Name = "LanguageUtil"
Option Explicit

Public LanguageEnglishTranslated() As String
Public LanguageCharset As Integer
Public CharacterType As String

'****************************************************************************************
'USAGE:
'   To Get All Forms/Report Information,
'       1. Put next lines in the 1st line of Main() Sub Routine.
'           ReadOnlyAppPath = App.Path & "\" 'PCN2123
'           GetAllFormsInfo
'           GetAllReportsInfo
'           Exit Sub
'       2. Put Exit Sub in the 1st line of ConvertLanguage Sub Routine.
'       3. Put Exit Sub in the 1st line of DataReport_Initialize Sub Routine.
'   To convert Language,
'       0. Coding To Get All Forms/Report Information as above should be removed.
'       1. Form_Load Sub Routine of Each Form should have
'           ConvertLanguage Me, Language
'       2. DataReport_Initialize Sub Routine of Each Report should have
'           ConvertLanguageInReport Me, Language
'****************************************************************************************

Sub GetAllFormsInfo()
'****************************************************************************************
'Name    : GetAllFormsInfo
'Created : 25 July 2003, PCN2111
'Updated : 25 July 2003, PCN2111
'Prg By  : Abe Park
'Param   : (None)
'Desc    : Gets Caption and ToolTipText of controls of forms in ClearLineProfiler.
'          Result is saved in template.txt file under APP.Path.
'Usage   : Put GetAllFormsInfo & Exit Sub in Main() Sub Routine in StartUp module.
'          If ConvertLanguage Function is being used in forms,
'               Put Exit Sub in the 1st line of ConvertLanguage Function.
'****************************************************************************************
On Error GoTo Err_Handler
    Const intTotFrm As Integer = 15 'PCN2477 PCN2506 PCN2747(14->15)
    If MsgBox("Do you want to create the default language file?", vbOKCancel) = vbCancel Then Exit Sub
    'Declares an array to represent all forms in ClearLineProfiler
    Dim frmName(intTotFrm) As Form
    'Assigns each form to a frmName 0 ~ 15.
    Set frmName(0) = AutoTune
    Set frmName(1) = ClearLineProfilerV6  'PCN4171
    Set frmName(2) = ClearLineScreen

    Set frmName(3) = V6Splash
    
    Set frmName(4) = PVReport4in1
    Set frmName(5) = PVReportMultiProfilex3
    Set frmName(6) = PVReportProfile
    
    Set frmName(7) = OptionsPage 'PCN3019
    Set frmName(8) = PipelineDetails 'PCN3019
    Set frmName(9) = PrecisionVisionGraph 'PCN3019
    Set frmName(10) = Registration 'PCN3019
    Set frmName(11) = Fisheye 'PCN2477 'PCN3019
    Set frmName(12) = PVReportSingle
    Set frmName(13) = CLPProgressBar 'PCN2747 'PCN3019
    
    Set frmName(14) = ControlsMain
    Set frmName(15) = ControlsScreen
        
    'index for FOR ... NEXT Loop.
    Dim intIndex As Integer
    'Load all form invisibly.
    For intIndex = 0 To intTotFrm
        frmName(intIndex).Hide
    Next intIndex
    
    'an array to contain infomation of controls with Caption and ToolTipText
    Dim strCtrlInfo() As String 'Array to contain information about controls
    'Number of information found in a form
    Dim intCount As Integer
    Dim intTotCnt As Integer 'Total number of information found in all forms
    intTotCnt = 0
    Dim intTotCnt2 As Integer 'Total number of information found in all forms, which is not Zero-Length String
    intTotCnt2 = 0
    'index for a loop inside of a loop.
    Dim intIndex2 As Integer
    
    MkDir ReadOnlyAppPath & "Language" 'PCN2123
    MkDir ReadOnlyAppPath & "Language\Template" 'PCN2123
    
    Dim FileNo As Integer
    FileNo = FreeFile
    Open ReadOnlyAppPath & "Language\Template\Template.txt" For Output As #FileNo 'PCN2123
    For intIndex = 0 To intTotFrm
        'Get information about All Controls in frmName(intIndex).
        intCount = GetAllControlsInfo(frmName(intIndex), strCtrlInfo)
        intTotCnt = intTotCnt + intCount + 1 ' + 1 is because the form itself has caption property.
        
        'Save in template.text
        'Caption of frmName itself
        Dim strEng As String
        If frmName(intIndex).Caption <> "" Then
            intTotCnt2 = intTotCnt2 + 1
            If InStr(frmName(intIndex).Caption, ",") = 0 Then
                strEng = frmName(intIndex).Caption
            Else
                strEng = Chr(34) & frmName(intIndex).Caption & Chr(34)
            End If
            Print #FileNo, frmName(intIndex).name & Chr(9) & _
                      frmName(intIndex).name & Chr(9) & _
                      "Caption" & Chr(9) & _
                      strEng '& Chr(9) & intTotCnt2 & "(G)" & strEng 'intTotCnt2 is for TESTING purpose only!
        End If
        For intIndex2 = 0 To intCount - 1
            If strCtrlInfo(intIndex2, 2) <> "" Then
                intTotCnt2 = intTotCnt2 + 1
                If InStr(strCtrlInfo(intIndex2, 2), ",") = 0 Then
                    strEng = strCtrlInfo(intIndex2, 2)
                Else
                    strEng = Chr(34) & strCtrlInfo(intIndex2, 2) & Chr(34)
                End If
                Print #FileNo, frmName(intIndex).name & Chr(9) & _
                          strCtrlInfo(intIndex2, 0) & Chr(9) & _
                          strCtrlInfo(intIndex2, 1) & Chr(9) & _
                          strEng '& Chr(9) & intTotCnt2 & "(G)" & strEng  'intTotCnt2 is for TESTING purpose only!
            End If
        Next intIndex2
    Next intIndex
    Close #FileNo
    
    MsgBox intTotCnt2 & " lines of form information logged.", , "Clear Line Profiler"
Exit Sub
Err_Handler:
    Select Case Err
        Case 75 'Path/File access error
            Resume Next
        Case Else
            MsgBox Err & "-L1:" & Error$
    End Select
End Sub

Private Function GetAllControlsInfo(ByRef frmOne As Form, ByRef strCtrlInfo() As String) As Integer
'***********************************************************************************
'Name    : GetAllControlsInfo
'Created : 25 July 2003, PCN2111
'Updated : 29 July 2003, PCN2111
'Prg By  : Abe Park
'Param   : frmOne - One of forms in ClearLineProfiler
'          strCtrlInfo - Four values of information from each control.
'Desc    : Gets three values of information from each control in frmOne in ClearLineProfiler.
'          Result is saved in strCtrlInfo.
'Usage   : (Used by GetAllFormsInfo Sub Routine within LanguageUtil Module)
'Return  : Total number of information found
'***********************************************************************************
On Error GoTo Err_Handler
    ReDim strCtrlInfo(1000, 2) As String 'Index of controls in a form with Caption or ToolTipText Property,
                                         'Three values(0:Control Name, 1:What Property, 2:The property value) of information from each control
    Dim lngErr As Long 'Stores Err number. This value is used to decide  whether the control is an array or not.
    Dim intIndex As Integer 'Index of controls in a form with Caption or ToolTipText Property
    Dim intCount As Integer 'Counter of controls in a form with Caption or ToolTipText Property
    intCount = -1
    Dim ctlOne As Control 'controls in a form
    Dim ctlOne2  'controls in a toolbar
    Dim strCtrlName As String 'storage for the current control name.
    For Each ctlOne In frmOne.Controls
        'Check whether the control is a ToolBar
        If TypeName(ctlOne) = "Toolbar" Then
            For Each ctlOne2 In ctlOne.buttons
                'Buttons in a toolbar : [ToolbarName].Buttons.Item([ButtonKeyValue])
                strCtrlName = ctlOne.name & ".Buttons.Item(" & ctlOne2.Key & ")"
                'Get information if Caption property exists for ctlOne
                intCount = intCount + 1
                strCtrlInfo(intCount, 0) = strCtrlName        'NAME
                strCtrlInfo(intCount, 1) = "Caption"           'PROPERTY
                strCtrlInfo(intCount, 2) = ctlOne2.Caption     'ENGLISH
                
                'Get information if ToolTipText property exists for ctlOne
                intCount = intCount + 1
                strCtrlInfo(intCount, 0) = strCtrlName        'NAME
                strCtrlInfo(intCount, 1) = "ToolTipText"       'PROPERTY
                strCtrlInfo(intCount, 2) = ctlOne2.ToolTipText 'ENGLISH
            Next ctlOne2
        End If
        'Check whether the control is an array.
        lngErr = -999 'Initialize
        intIndex = ctlOne.Index 'Tries to get index of ctlOne to know whether ctlOne is an array.
        
        'Decides the name of the control.
        If lngErr = 343 Then 'ctlOne is not an array.
            strCtrlName = ctlOne.name
        Else
            strCtrlName = ctlOne.name & "(" & intIndex & ")"
        End If
        
        'Get information if Caption property exists for ctlOne
        intCount = intCount + 1
        strCtrlInfo(intCount, 0) = strCtrlName        'NAME
        strCtrlInfo(intCount, 1) = "Caption"          'PROPERTY
        strCtrlInfo(intCount, 2) = ctlOne.Caption     'ENGLISH
        
        'Get information if ToolTipText property exists for ctlOne
        intCount = intCount + 1
        strCtrlInfo(intCount, 0) = strCtrlName        'NAME
        strCtrlInfo(intCount, 1) = "ToolTipText"      'PROPERTY
        strCtrlInfo(intCount, 2) = ctlOne.ToolTipText 'ENGLISH
    Next ctlOne
    
    'Return Total number of information found
    GetAllControlsInfo = intCount + 1 ' + 1 is for 0 ~ intCount
Exit Function
Err_Handler:
    Select Case Err
        Case 343 'Object Not Array
            lngErr = Err
            Resume Next
        
        Case 438 'Object doesn't support this property
            intCount = intCount - 1
            Resume Next
            
        Case Else
            MsgBox Err & "-L2:" & Error$
    End Select
End Function


Public Sub ConvertLanguage(ByRef frmOneForm As Form, strLanguage As String)
'***********************************************************************************
'Name    : ConvertLanguage
'Created : 28 July 2003, PCN2111
'Updated : 29 July 2003, PCN2111
'          30 July 2003, When there is a line without TAB
'Prg By  : Abe Park
'Param   : frmOneForm - a form being loaded
'          strLanguage - The name of Language to use. E.G., English, French, German
'Desc    : Converts English in frmOneForm into strLanguage.
'Usage   : Put ConvertLanguage in Form_Load event of each form.
'***********************************************************************************
On Error GoTo Err_Handler
Dim ctlOne As Control 'controls in a form
Dim intTotCnt As Integer
Dim TextForTranslation As String

intTotCnt = 0
For Each ctlOne In frmOneForm.Controls
    intTotCnt = intTotCnt + 1
    With ctlOne
        Select Case TypeName(ctlOne)
            Case "Form"
                'Caption only
                TextForTranslation = .Caption
                If Len(TextForTranslation) <> 0 Then
'                    Debug.Print .name
                    .Font.Charset = LanguageCharset
                    .Caption = DisplayMessage(TextForTranslation)

'                    Debug.Print .Caption
                End If
        
            Case "TextBox"
                'Text and ToolTipText
                TextForTranslation = .text
                If Len(TextForTranslation) <> 0 Then
                    
                    .Font.Charset = LanguageCharset
                    .text = DisplayMessage(TextForTranslation)
                    
                End If
                TextForTranslation = .ToolTipText
                If Len(TextForTranslation) <> 0 Then
                    .ToolTipText = DisplayMessage(TextForTranslation)
                End If
                
            Case "Frame"
                'Caption only
                TextForTranslation = .Caption
                If Len(TextForTranslation) <> 0 Then
'                    Debug.Print .name
                    .Font.Charset = LanguageCharset
                    .Caption = DisplayMessage(TextForTranslation)

'                    Debug.Print .Caption
                End If
            
            Case "ComboBox"
                'Caption only
                TextForTranslation = .text
                If Len(TextForTranslation) <> 0 Then
'                    Debug.Print .name
                    .Font.Charset = LanguageCharset
                    .text = DisplayMessage(TextForTranslation)

'                    Debug.Print .Caption
                End If
                
            Case "CommandButton", "Label"
                'Caption and ToolTipText
                TextForTranslation = .Caption
                If Len(TextForTranslation) <> 0 Then
                    
                    .Font.Charset = LanguageCharset
                    .Caption = DisplayMessage(TextForTranslation)

                End If
                Debug.Print ctlOne.name
                TextForTranslation = .ToolTipText
                If Len(TextForTranslation) <> 0 Then
                    .ToolTipText = DisplayMessage(TextForTranslation)
                End If
                
            Case "Image", "PictureBox"
                'ToolTipText only
                TextForTranslation = .ToolTipText
                If Len(TextForTranslation) <> 0 Then
                    .ToolTipText = DisplayMessage(TextForTranslation)
                End If
                
            Case Else
        End Select
    End With
ContinueWithNextControl:
Next ctlOne


Exit Sub
Err_Handler:
    Select Case Err
        Case 383 'Active Form's property is read-only 'Test
            Resume ContinueWithNextControl
        Case 438 'Object does not support this property or method
            Resume ContinueWithNextControl
        Case 730 'Control does not exist 'PCN2193
            Resume ContinueWithNextControl
        Case 35601 'Element Not Found PCN2139
            Resume ContinueWithNextControl
        Case Else
            MsgBox Err & "-L3:" & Error$
    End Select
End Sub


Public Sub ConvertLanguageInReport(ByRef rptOneReport, strLanguage As String)
'***********************************************************************************
'Name    : ConvertLanguageInReport
'Created : 29 July 2003, PCN2111
'Updated : 29 July 2003, PCN2111
'          30 July 2003, When there is a line without TAB
'Prg By  : Abe Park
'Param   : rptOneReport - a report being loaded
'          strLanguage - The name of Language to use. E.G., English, French, German
'Desc    : Converts English in frmOneReport into strLanguage.
'Usage   : Put ConvertLanguageInReport in DataReport_Initialize event of each report.
'***********************************************************************************
On Error GoTo Err_Handler
        
    Dim strLangFile As String 'to store the text filename for the Language to load.
    Dim sectionsError  As Boolean 'PCN3297 if the sections for lanuage throw an error this is set to true
    'Get the filename for strLanguage
    strLangFile = GetLanguageFile(strLanguage)
    If strLangFile = "" Then Exit Sub
    
    Dim FileNo As Integer
    FileNo = FreeFile
    Open ReadOnlyAppPath & "Language\" & strLangFile For Input As #FileNo 'PCN2123
        
    Dim strOneLine As String 'To contain one line from Language text file.
    Dim strValue(4) As String '0:report name, 1:Component, 2:Property, 3:English, 4:Other Language
    Dim intIndex As Integer 'Used as Index value for a control array
    Dim strSection As String 'Used as Section Name in a report
    Dim strControl As String 'Used as Control Name for each section in a report
    Dim intPosition As Integer 'Used to get position of "," or "("
    Dim intPosition2 As Integer 'Used to get position of ")"
    
    'Get the 1st line which is version information.
    'Line Input #1, strOneLine
            
    While Not EOF(FileNo)
TryNextLine:
        'get the next line
        Line Input #FileNo, strOneLine
        intPosition = InStr(strOneLine, Chr(9))
        If intPosition = 0 Then 'When there is a line without TAB, Read Next line.
            If EOF(FileNo) Then
                Close #FileNo
                Exit Sub
            End If
            GoTo TryNextLine
        End If
        'Get "Report Name"
        strValue(0) = Trim(Left(strOneLine, intPosition - 1))
        'if report name is same to rptOneReport.Name, then get other values(component, property, English, Other Language) ---v
        If strValue(0) = rptOneReport.name Then
            strOneLine = Right(strOneLine, Len(strOneLine) - intPosition)
            intPosition = InStr(strOneLine, Chr(9))
            'Get "Component"
            strValue(1) = Trim(Left(strOneLine, intPosition - 1))
            strOneLine = Right(strOneLine, Len(strOneLine) - intPosition)
            intPosition = InStr(strOneLine, Chr(9))
            'Get "Property"
            strValue(2) = Trim(Left(strOneLine, intPosition - 1))
            strOneLine = Right(strOneLine, Len(strOneLine) - intPosition)
            intPosition = InStr(strOneLine, Chr(9))
            If intPosition > 0 Then
                'Get "English"
                strValue(3) = Trim(Left(strOneLine, intPosition - 1))
                'Get "Other Language"
                If Trim(Right(strOneLine, Len(strOneLine) - intPosition)) = "" Then 'If Other Language is "", Display English
                    strValue(4) = strValue(3)
                Else
                    strValue(4) = Trim(Right(strOneLine, Len(strOneLine) - intPosition))
                End If
            Else 'There is no Other Language.
                'Get "English"
                strValue(3) = strOneLine
                'Get "Other Language"
                strValue(4) = strOneLine  'Then display English.
            End If
            'Remove Quotation marks if exist -------------------------v
            If Left(strValue(4), 1) = Chr(34) Then
                strValue(4) = Right(strValue(4), Len(strValue(4)) - 1)
            End If
            If Right(strValue(4), 1) = Chr(34) Then
                strValue(4) = Left(strValue(4), Len(strValue(4)) - 1)
            End If '--------------------------------------------------^
            'Check whether the component is a control in a section.
            intPosition = InStr(strValue(1), "Sections.Item(")
            If intPosition <> 0 Then 'if the component is a control in a section
                intPosition2 = InStr(strValue(1), ")")
                strSection = Mid(strValue(1), intPosition + 14, intPosition2 - intPosition - 14) '14 is Length of ->Sections.Item(<-
                strValue(1) = Right(strValue(1), Len(strValue(1)) - intPosition2)
                intPosition = InStr(strValue(1), "Controls.Item(")
                intPosition2 = InStr(strValue(1), ")")
                strControl = Mid(strValue(1), intPosition + 14, intPosition2 - intPosition - 14) '14 is Length of ->Controls.Item(<-
                sectionsError = False 'PCN3297
                rptOneReport.Sections.Item(strSection).Controls.Item(strControl).Caption = strValue(4)
                
                'PCN3297 If above section fails try sections 1 to 5 done by resuming on error and setting
                ' sectionsError to True '''''''''
                If sectionsError Then           '
                    strSection = "Section1"     '
                    sectionsError = False       '
                    rptOneReport.Sections.Item(strSection).Controls.Item(strControl).Caption = strValue(4)
                End If                          '
                If sectionsError Then           '
                    strSection = "Section2"     '
                    sectionsError = False       '
                    rptOneReport.Sections.Item(strSection).Controls.Item(strControl).Caption = strValue(4)
                End If                          '
                If sectionsError Then           '
                    strSection = "Section3"     '
                    sectionsError = False       '
                    rptOneReport.Sections.Item(strSection).Controls.Item(strControl).Caption = strValue(4)
                End If                          '
                If sectionsError Then           '
                    strSection = "Section4"     '
                    sectionsError = False       '
                    rptOneReport.Sections.Item(strSection).Controls.Item(strControl).Caption = strValue(4) 'PCN3962 subscript outofbounds in Italian
                End If                          '
                If sectionsError Then           '
                    strSection = "Section5"     '
                    sectionsError = False       '
                    rptOneReport.Sections.Item(strSection).Controls.Item(strControl).Caption = strValue(4)
                End If                          '
                ' If all the above fail then resume with english
                '''''''''''''''''''''''''''''''''
                
            Else 'if the component is NOT a control in a section.
                If rptOneReport.name = strValue(1) Then 'This is the Caption of form itself.
                    rptOneReport.Caption = strValue(4)
                End If
            End If
        End If
    Wend
    Close #FileNo
Exit Sub
Err_Handler:
    Select Case Err
        Case 8574 'PCN3297 (7 Feb 2005, Antony), if rptOneReport.Sections.Item(strSection).... not found try the sections 1 to 5
            sectionsError = True 'by resuming. Setting this to True forces the next section to be tried
            Resume Next 'If none of the section work, just egnore and use english
        Case 9: Resume Next 'Subscript out of range, the above line caused it when in italian
        Case Else
            MsgBox Err & "-L4:" & Error$
    End Select
End Sub



Public Function DisplayMessage(strMessage As String) As String
'***********************************************************************************
'Name    : DisplayMessage
'Created :  29 July 2003, PCN2111
'Updated :  30 July 2003, 'PCN4171
'
'Prg By  : Geoff Logan
'Param   : strMessage - The English MsgBox string to be converted to the current language setting
'
'Desc    : Takes the English strMessage MsgBox string, looks for a corresponding MessageBox entry
'           in the language.txt language file and returns the result of the English translation.
'Usage   : Replace a MsgBox message string with this function. E.g. MsgBox DisplayMessage("Please calibrate first."), vbInformation
'***********************************************************************************
On Error GoTo Err_Handler
Dim LanguageIndex As Integer 'Used as Index value for a language array
Dim LanguageArraySize As Long
Dim str2Compare1 As String
Dim str2Compare2 As String
Dim str2Compare2WithOutEndString As String
Dim EndCharOfString As String

If Language = "English" Then
    DisplayMessage = strMessage
    Exit Function
End If

str2Compare2 = Trim(strMessage)
EndCharOfString = Right(str2Compare2, 1)
If EndCharOfString = "." Or EndCharOfString = "," Or EndCharOfString = "!" Or EndCharOfString = "?" Or EndCharOfString = ":" Then
    str2Compare2WithOutEndString = Trim(Left(str2Compare2, Len(str2Compare2) - 1))
Else
    str2Compare2WithOutEndString = ""
    EndCharOfString = ""
End If

LanguageArraySize = UBound(LanguageEnglishTranslated, 2)
If LanguageArraySize = 0 Then
    'There is not valid language data loaded
    Language = "English"
    DisplayMessage = strMessage
    Exit Function
End If
For LanguageIndex = 1 To LanguageArraySize
    If LanguageEnglishTranslated(1, LanguageIndex) <> "" Then
        'Exclude any character like '.', ',', '!', '?', ' ') in the end of a line.
        If Right(LanguageEnglishTranslated(0, LanguageIndex), 1) = "." Or Right(LanguageEnglishTranslated(0, LanguageIndex), 1) = "," _
            Or Right(LanguageEnglishTranslated(0, LanguageIndex), 1) = "!" Or Right(LanguageEnglishTranslated(0, LanguageIndex), 1) = "?" Then
            str2Compare1 = Trim(Left(LanguageEnglishTranslated(0, LanguageIndex), Len(LanguageEnglishTranslated(0, LanguageIndex)) - 1))
        Else
            str2Compare1 = LanguageEnglishTranslated(0, LanguageIndex)
        End If
        If UCase(str2Compare1) = UCase(str2Compare2) Then
            DisplayMessage = LanguageEnglishTranslated(1, LanguageIndex)
            Exit Function
        ElseIf UCase(str2Compare1) = UCase(str2Compare2WithOutEndString) Then  'Matching Message Found.
            If EndCharOfString = Right(LanguageEnglishTranslated(1, LanguageIndex), 1) Then
                DisplayMessage = LanguageEnglishTranslated(1, LanguageIndex)
            Else
                DisplayMessage = LanguageEnglishTranslated(1, LanguageIndex) & EndCharOfString
            End If
            Exit Function
        End If
    End If
Next LanguageIndex

DisplayMessage = strMessage

Exit Function
Err_Handler:
    Select Case Err
        Case Else
            MsgBox Err & "-L5:" & Error$
    End Select

End Function

Sub LoadLanguageFromFile(ErrorStr As String)
'***********************************************************************************
'Name    : LoadLanguageFromFile
'Created :  11 August 2006, PCN4171
'Updated :
'
'Prg By  : Geoff Logan
'Param   :
'
'Desc    : Loads the language English translation into the LanguageEnglishTranslated
'          array for use in the DisplayMessage function
'Usage   :
'***********************************************************************************
On Error GoTo Err_Handler
Dim strLangFile As String 'to store the text filename for the Language to load.
Dim strOneLine As String 'To contain one line from Language text file.
Dim strValue(1) As String '0:English, 1:Other Language
Dim intPosition As Long 'Used to get position of "," or "("
Dim FileNumber As Integer
Dim NoOfLanguageItems As Long
Dim CurrentLanguage As String
Dim FileVersion As String
Dim CurrentLineInFile As Long
Dim i As Long
Dim UnicodeByte() As Byte
Dim UnicodeString As String
Dim CharacterSetString As String

Language = "English"

ReDim LanguageEnglishTranslated(1, 0)

CurrentLanguage = GetCurrentLanguageSetting
If CurrentLanguage = "English" Then
    Exit Sub
End If

'Get the filename for strLanguage
strLangFile = GetLanguageFile(CurrentLanguage)
If strLangFile = "" Then
    ErrorStr = "Language file not found"
    Exit Sub
End If

FileNumber = FreeFile

'If CurrentLanguage <> "Japanese" Then
'For non-unicode translation



Open ReadOnlyAppPath & "Language\" & strLangFile For Input As #FileNumber

If EOF(FileNumber) Then
    Close #FileNumber
    ErrorStr = "Invalid language file"
    Exit Sub
End If


'Get Version

Line Input #FileNumber, strOneLine
intPosition = InStr(strOneLine, "Version=")
If intPosition = 0 Then
    Close #FileNumber
    ErrorStr = "Invalid language file"
    Exit Sub
End If

FileVersion = Right(strOneLine, Len(strOneLine) - 7 - intPosition)
FileVersion = Left(FileVersion, 3)
If FileVersion = "2.0" Then
    Line Input #FileNumber, strOneLine
    intPosition = InStr(strOneLine, "CharacterType=")
    If intPosition = 0 Then
        Close #FileNumber
        ErrorStr = "Invalid language file"
        Exit Sub
    End If
    CharacterType = Right(strOneLine, Len(strOneLine) - 13 - intPosition)
    CharacterType = Left(CharacterType, 7)

    Line Input #FileNumber, strOneLine
    intPosition = InStr(strOneLine, "CharacterSet=")
    If intPosition = 0 Then
        Close #FileNumber
        ErrorStr = "Invalid language file"
        Exit Sub
    End If
    CharacterSetString = Right(strOneLine, Len(strOneLine) - 12 - intPosition)
    
    
End If
    
If SafeCDbl(FileVersion) < 2 Or LCase(CharacterType) <> "unicode" Then
'If CDbl(FileVersion) < 1 Then
    While Not EOF(FileNumber)
        'Get the next line
        Line Input #FileNumber, strOneLine
        
        intPosition = InStr(strOneLine, Chr(9))
        If intPosition = 0 Then 'When there is a line without TAB, Read Next line.
            If EOF(FileNumber) Then
                Close #FileNumber
                Exit Sub
            End If
        Else
            'Get "English"
            strValue(0) = Trim(Left(strOneLine, intPosition - 1))
            
            'Get Translation
            strValue(1) = Right(strOneLine, Len(strOneLine) - intPosition)
            'Remove Quotation marks if exist -------------------------v
            If Left(strValue(0), 1) = Chr(34) Then
                strValue(0) = Right(strValue(0), Len(strValue(0)) - 1)
            End If
            If Right(strValue(0), 1) = Chr(34) Then
                strValue(0) = Left(strValue(0), Len(strValue(0)) - 1)
            End If
            If Left(strValue(1), 1) = Chr(34) Then
                strValue(1) = Right(strValue(1), Len(strValue(1)) - 1)
            End If
            If Right(strValue(1), 1) = Chr(34) Then
                strValue(1) = Left(strValue(1), Len(strValue(1)) - 1)
            End If '--------------------------------------------------^
        
            NoOfLanguageItems = UBound(LanguageEnglishTranslated, 2) + 1
            ReDim Preserve LanguageEnglishTranslated(1, NoOfLanguageItems)
            LanguageEnglishTranslated(0, NoOfLanguageItems) = strValue(0) 'English
            LanguageEnglishTranslated(1, NoOfLanguageItems) = strValue(1) 'Translation
        
        End If
    Wend
    Close #FileNumber
Else
    'For Unicode Translation
'    Open ReadOnlyAppPath & "Language\" & strLangFile For Input As #FileNumber
'
'    If EOF(FileNumber) Then
'        Close #FileNumber
'        ErrorStr = "Invalid language file"
'        Exit Sub
'    End If
'
'    'Get Version
'    Line Input #FileNumber, strOneLine
'    intPosition = InStr(strOneLine, "Version=")
'    If intPosition = 0 Then
'        Close #FileNumber
'        ErrorStr = "Invalid language file"
'        Exit Sub
'    End If
'    FileVersion = Right(strOneLine, Len(strOneLine) - 10)
    
    Close #FileNumber
    
    'Open file in Binary to get UNICODE
    Dim FileLength As Long
            
    Dim FileNo As Integer
    FileNo = FreeFile
    Open ReadOnlyAppPath & "Language\" & strLangFile For Binary As #FileNo
        FileLength = LOF(1)
        'UnicodeByte = StrConv(InputB(FileLength, #1), vbFromUnicode, 1041)
        UnicodeByte = InputB(FileLength, #FileNo)
    Close #FileNo
    
    Dim J As Long
    i = 0
    
    For J = 1 To UBound(UnicodeByte)
        If UnicodeByte(J) = 13 Then
            J = J + 2
            Do While UnicodeByte(J) <> 9
                If UnicodeByte(J) = 13 Then
                    UnicodeString = ""
                    J = J + 1
                End If
                UnicodeString = UnicodeString + Chr$(UnicodeByte(J))
                If J >= UBound(UnicodeByte) Then Exit For
                J = J + 1
            Loop
            i = i + 1
            ReDim Preserve LanguageEnglishTranslated(1, i)
            UnicodeString = StrConv(UnicodeString, vbFromUnicode, 1041)
            UnicodeString = Right(UnicodeString, Len(UnicodeString) - 1)
            LanguageEnglishTranslated(0, i) = UnicodeString
            UnicodeString = ""
        End If
        
        If UnicodeByte(J) = 9 Then
            J = J + 1
            Do While (True)
                If J >= UBound(UnicodeByte) Then Exit Do
                If (J + 1) <> UBound(UnicodeByte) Then
                    If UnicodeByte(J + 1) = 13 And UnicodeByte(J + 2) = 0 Then Exit Do
                End If
                UnicodeString = UnicodeString + ChrB(UnicodeByte(J + 1))
                'UnicodeString = ChrB(0) + ChrB(253)
                If J >= UBound(UnicodeByte) Then Exit For
                J = J + 1
            Loop
            LanguageEnglishTranslated(1, i) = UnicodeString
            UnicodeString = ""
        End If
        
        If J >= UBound(UnicodeByte) Then Exit For

    Next J
    
End If



Language = CurrentLanguage 'This is a valid language

If FileVersion = "2.0" Then
    LanguageCharset = CInt(CharacterSetString)
Else
    Select Case Language
        Case "Japanese"
            LanguageCharset = 128
        Case "Chinese"
            LanguageCharset = 136
        Case "Greek"
            LanguageCharset = 161
        Case "Turkish"
            LanguageCharset = 162
        Case "Hebrew"
            LanguageCharset = 177
        Case "Arabic"
            LanguageCharset = 178
        Case "Russian"
            LanguageCharset = 204
        Case "Thai"
            LanguageCharset = 222
        Case Else
            LanguageCharset = 0
    End Select
End If

Exit Sub
Err_Handler:
    Select Case Err
        Case Else
            MsgBox Err & "-L6:" & Error$
    End Select
End Sub


Function GetLanguageFile(ByVal strLanguage As String) As String
'***********************************************************************************
'Name    : GetLanguageFile
'Created :  30 July 2003, PCN2111
'Updated :  30 July 2003
'
'Prg By  : Abe Park
'Param   : strLanguage - The name of Language, which language file should be found.
'Return  : The name of the language text file related to strLanguage
'Desc    : First, Reads for Languages.txt. If Languages.txt file does not exist, Return False and Exit.
'          Then, Looks for a line containing strLanguage value and its language text file name.
'               If the line containing this information is not found, Return False and Exit.
'          Last, Looks for the language text file. If the language text file does not exist, Return False and exit.
'Usage   : Put
'               Dim strLangFile As String
'               strLangFile = GetLanguageFile(<Name of the Language>)
'               If strLangFile = "" Then
'                   {Return a proper value if it is a Function}
'                   Exit Sub(or Function)
'               End If
'          in the 1st line of Language Conversion Functions as below.
'               ConvertLanguage
'               ConvertLanguageInReport
'               DisplayMessage
'***********************************************************************************
On Error GoTo Err_Handler
Dim ErrCount As Integer 'PCN2193

    If strLanguage = "English" Then
        Exit Function
    End If
    
    ErrCount = 0
    Dim FileNo As Integer
    FileNo = FreeFile
    'Get the filename for strLanguage ==================================================================v
    Dim strLang As String 'to store the text filename for the Language to load.
    Dim intPos As Integer 'PCN2169
    Dim blnEnd As Boolean
    If Dir(ReadOnlyAppPath & "Language\Languages.txt") <> "" Then 'Continues only if Languages.txt file exists.'PCN2123
        ' Get the filename for strLanguage ------------------------v
        
        
        Open ReadOnlyAppPath & "Language\Languages.txt" For Input As #FileNo 'PCN2193 'PCN2123
        blnEnd = False
        While Not EOF(FileNo) And Not blnEnd  'PCN2193
            Line Input #FileNo, strLang  'PCN2193
            intPos = InStr(strLang, ",")
            If intPos > 0 Then 'PCN2172
                If Left(strLang, intPos - 1) = strLanguage Then
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
        Wend '-----------------------------------------------------^
        If Not blnEnd Then 'Exit if the text filename is not found in Languages.txt file.
            Close #FileNo 'PCN2193
            Exit Function
        End If
        Close #FileNo 'PCN2193
    Else
        Exit Function
    End If '=============================================================================================^
    
    If Dir(ReadOnlyAppPath & "Language\" & strLang) = "" Then 'If the file for strLanguage does not exist, Exit.'PCN2123
        Exit Function
    End If
    
    'Both Languages.txt and the Language text file related to strLanguage EXIST.
    GetLanguageFile = strLang 'The name of the language text file related to strLanguage

Exit Function 'PCN2193
CloseFileOnErr: 'PCN2193
    ErrCount = ErrCount + 1
    If ErrCount < 2 Then
        Close #FileNo 'PCN2193
    End If

Exit Function
Err_Handler:
    Select Case Err
        Case Else
            MsgBox Err & "-L7:" & Error$
    End Select
    GoTo CloseFileOnErr 'PCN2193
End Function

Function ConvertRichToAnsi(RichText As String) As String
    Dim UnicodeString As String
    Dim StrPos As Integer
    Dim strlength As Integer
    Dim HexStr(50) As String
    Dim StartPos As Integer
    Dim i As Integer
    
    strlength = Len(RichText)
    
    i = 0
    StrPos = 1
    StartPos = 1
    Do While StrPos < strlength + 1
        StrPos = InStr(StartPos, RichText, "'") ' + StartPos - 1
        If StrPos = 0 Then Exit Do
        HexStr(i) = Mid(RichText, (StrPos + 1), 2)
        StartPos = StrPos + 1
        i = i + 1
    Loop
    
    UnicodeString = ""
    i = 0
    
    Do While HexStr(i) <> ""
        'UnicodeString = UnicodeString & Chr$(Val("&H" & HexStr(i)))
        UnicodeString = UnicodeString + ChrB(Val("&H" & HexStr(i)))
        i = i + 1
    Loop
    UnicodeString = StrConv(UnicodeString, vbUnicode)
    
    
    ConvertRichToAnsi = UnicodeString
    
End Function

'Function ConvertUnicodeToAnsi(UNIText As String) As String
'    Dim UnicodeString As String
'    Dim StrPos As Integer
'    Dim strlength As Integer
'    Dim HexStr(50) As String
'    Dim StartPos As Integer
'    Dim i As Integer
'
'    strlength = Len(UNIText)
'
'    i = 0
'    StrPos = 0
'    Do While StrPos < strlength + 1
'        HexStr(i) = Mid(UNIText, StrPos, 2)
'        StrPos = StrPos + 2
'        i = i + 1
'    Loop
'
'    UnicodeString = ""
'    i = 0
'
'    Do While HexStr(i) <> ""
'        UnicodeString = UnicodeString + ChrB(Val("&H" & HexStr(i)))
'        i = i + 1
'    Loop
'    UnicodeString = StrConv(UnicodeString, vbUnicode)
'
'
'    ConvertUnicodeToAnsi = UnicodeString
'
'End Function
