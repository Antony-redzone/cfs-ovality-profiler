Attribute VB_Name = "Interfaces"
'PCN2133
Option Explicit

Public strVersion As String 'Version of CLPInterface.int file

Public strPVDName As String
Public strVideoName As String

'Variables for [Header Info]

Public LocalAssetNo As String
Public LocalStartNodeNo As String
Public LocalStartNodeLocation As String
Public LocalCity As String
Public LocalFinishNodeNo As String
Public LocalFinishNodeLocation As String
Public LocalsDate As String
Public LocalsTime As String
Public LocalPipeLength As String
Public LocaldblDiameter As String
Public LocalMaterial As String
Public LocalRecommendations As String
Public LocalSiteID As String
Public dblDiameter As Double

Private ConfigArray()
Private ErrorLog() As String
'----------------------------------------------------------------------^

Public Function SetupAsPerInterfaceFile(ByVal InterfaceFileName As String)
'****************************************************************************************
'Name    : SetupAsPerInterfaceFile
'Created : 31 July 2003, PCN2133
'Updated :
'Prg By  : Abe Park
'Param   : InterfaceFile - the File Path & File Name to read
'                          This File should specify Section Head Name between []
'                          Section Detail format is <Name>=. E.g., PicFileName=
'Desc    : Gets Information from InterfaceFileName and setups Clear Line Profiler.
'Usage   : This function is placed in the end of Main() sub routine.
'****************************************************************************************
On Error GoTo Err_Handler

Call GetINT_Information(InterfaceFileName)

Exit Function
Err_Handler:
    Select Case Err
        Case Else
            MsgBox Err & "-I1:" & Error$
    End Select
End Function

Private Sub GetINT_Information(InterfaceFile As String)
'****************************************************************************************
'Name    : GetINT_Information
'Created : 31 July 2003, PCN2133
'Updated :
'Prg By  : Abe Park
'Param   : InterfaceFile - the File Path & File Name to read
'                          This File should specify Section Head Name between []
'                          Section Detail format is <Name>=. E.g., PicFileName=
'Desc    : Reads InterfacFile and assigns values to proper variables.
'Usage   : Used in SetupAsPerInterfaceFile to read InterfaceFileName
'****************************************************************************************
On Error GoTo Err_Handler

' MGR 17/10/2002
' Read INI File and populate LOGO & Contractor Details

' Get INI file from current directory and load into memory

Dim ConfigLine As String
Dim Config_LineCnt As String
Dim arrayIndex As Long


Dim SectionHead As String
Dim SectionDetail As String


SectionDetail = "***"
   
Config_LineCnt = 0

Dim FileNo As Integer
FileNo = FreeFile

Open InterfaceFile For Input As #FileNo
Do While Not EOF(FileNo)
 Line Input #FileNo, ConfigLine
 Config_LineCnt = Config_LineCnt + 1
Loop
Close #FileNo


ReDim ConfigArray(Config_LineCnt)
  
arrayIndex = 1

FileNo = FreeFile
Open InterfaceFile For Input As #FileNo
Do While Not EOF(FileNo)
 Line Input #FileNo, ConfigLine
 ConfigArray(arrayIndex) = ConfigLine
 arrayIndex = arrayIndex + 1
Loop
Close #FileNo

' Run Through Entire Array and Validate Paths and Files

For arrayIndex = 1 To Config_LineCnt
    SectionHead = ConfigArray(arrayIndex): If UCase(SectionHead) = UCase("[Version Info]") Then Call GetVersionInfo(arrayIndex)
    SectionHead = ConfigArray(arrayIndex): If UCase(SectionHead) = UCase("[File Pointers]") Then Call GetFilePointers(arrayIndex)
    SectionHead = ConfigArray(arrayIndex): If UCase(SectionHead) = UCase("[Process Request]") Then Call GetProcessRequest(arrayIndex)
    SectionHead = ConfigArray(arrayIndex): If UCase(SectionHead) = UCase("[Header Info]") Then Call GetHeaderInfo(arrayIndex)
   SectionDetail = "***" ' reset to NON blank
Next arrayIndex


Exit Sub
Err_Handler:
Select Case Err
    Case 9 'Subscript out of range (end of file)
        Exit Sub
    Case 13 'Type Mismatch 'E.G., strAssetID = <string value>
        Resume Next
    Case Else
        MsgBox Err & "-I2:" & Error$
End Select
End Sub

Sub GetVersionInfo(ByRef arrayIndex As Long)
On Error GoTo Err_Handler
    Dim SectionDetail As String
    Dim Parameter As String
    Dim value As String
    
    arrayIndex = arrayIndex + 1
    SectionDetail = ConfigArray(arrayIndex)
    Do Until SectionDetail = "" Or Left(SectionDetail, 1) = "[" Or SectionDetail = "\\" 'End marker of data
        If SectionDetail <> "" Then Call GetParam(SectionDetail, Parameter, value)
        Select Case Parameter
            Case "Version="
                strVersion = value
            Case Else
        End Select
        arrayIndex = arrayIndex + 1
        SectionDetail = ConfigArray(arrayIndex)
    Loop
    Dim strTemp As String
    If strVersion <> "2.0" Then
        strTemp = DisplayMessage("CLPInterface  VERSION ERROR. Expecting ")
        strTemp = strTemp & "2.0" & DisplayMessage(", CLPInterface is currently ")
        strTemp = strTemp & Format(strVersion, "###0.0") & DisplayMessage(" - This application may not work as designed.") 'PCN2111
        MsgBox strTemp, vbCritical 'PCN2111
    End If

Exit Sub
Err_Handler:
MsgBox Err & "-I3:" & Error$

End Sub

Sub GetFilePointers(ByRef arrayIndex As Long)
On Error GoTo Err_Handler
    Dim SectionDetail As String
    Dim Parameter As String
    Dim PathName As String
    Dim FileName As String
    Dim FileExtension As String
    Dim value As String
    Dim TryFileName As String
    Dim Extension As String
    
    arrayIndex = arrayIndex + 1
    SectionDetail = ConfigArray(arrayIndex)
    Do Until SectionDetail = "" Or Left(SectionDetail, 1) = "[" Or SectionDetail = "\\" 'End marker of data
        If SectionDetail <> "" Then Call GetParam(SectionDetail, Parameter, value)
        Select Case Parameter
            Case "MediaFilePath="
                Call PageFunctions.SplitFilePath(value, PathName, FileName, Extension)
                
            Case Else
        End Select
        arrayIndex = arrayIndex + 1
        SectionDetail = ConfigArray(arrayIndex)
    Loop
    
    'Extension = UCase(Extension) 'PCD deploy error, make the ucase part of the if statement,
                                  'that way the video extension doesnt end up converted to upercase
 
    
    TryFileName = PathName & FileName & ".PVD"

    If Dir(TryFileName) <> "" Then strPVDName = TryFileName: Call LoadFile: Exit Sub
    
    
    TryFileName = PathName & FileName & "." & Extension
    If UCase(Extension) = "MPG" Or _
       UCase(Extension) = "MPA" Or _
       UCase(Extension) = "M2P" Or _
       UCase(Extension) = "MP2" Or _
       UCase(Extension) = "AVI" Or _
       UCase(Extension) = "VOB" Or _
       UCase(Extension) = "BMP" Or _
       UCase(Extension) = "JPG" Then
       strPVDName = ""
       
       If Dir(TryFileName) <> "" Then
            strVideoName = TryFileName
            Call LoadFile
            Call AddErrorCode("Cannot find PVD, loading media file only")
            Exit Sub
        End If
    End If
        
    strPVDName = ""
    strVideoName = ""
    Call AddErrorCode("Cannot find PVD or media file")
        
Exit Sub
Err_Handler:
    Select Case Err
        Case 52: Call AddErrorCode("Bad file name or directory"): Exit Sub
        Case Else: MsgBox Err & "-I4:" & Error$
    End Select
End Sub

Sub GetProcessRequest(ByRef arrayIndex As Long)
On Error GoTo Err_Handler
    Dim SectionDetail As String
    Dim Parameter As String
    Dim value As String

    arrayIndex = arrayIndex + 1
    SectionDetail = ConfigArray(arrayIndex)
    Do Until SectionDetail = "" Or Left(SectionDetail, 1) = "[" Or SectionDetail = "\\" 'End marker of data
        If SectionDetail <> "" Then Call GetParam(SectionDetail, Parameter, value)
        If InStr(1, Parameter, "Process") Then Call DoProcess(value)
        
        arrayIndex = arrayIndex + 1
        SectionDetail = ConfigArray(arrayIndex)
    Loop
Exit Sub
Err_Handler:
MsgBox Err & "-I5:" & Error$

End Sub

Sub GetHeaderInfo(ByRef arrayIndex As Long)
On Error GoTo Err_Handler
    Dim SectionDetail As String
    Dim Parameter As String
    Dim value As String
    
   
    
    
    arrayIndex = arrayIndex + 1
    SectionDetail = ConfigArray(arrayIndex)
    Do Until SectionDetail = "" Or Left(SectionDetail, 1) = "[" Or SectionDetail = "\\" 'End marker of data
        If SectionDetail <> "" Then Call GetParam(SectionDetail, Parameter, value)
        Select Case Parameter
            Case "AssetID=":        LocalAssetNo = value 'PCN2239
            Case "ProjectNumber=":
            Case "StartNode=":      LocalStartNodeNo = value
            Case "SNLocation=":     LocalStartNodeLocation = value
            Case "City=":           LocalCity = value
            Case "EndNode=":        LocalFinishNodeNo = value
            Case "ENLocation=":     LocalFinishNodeLocation = value
            Case "Date=":           LocalsDate = value
            Case "Time=":           LocalsTime = value
            Case "PipeLength=":     LocalPipeLength = value
            Case "Diameter=":       LocaldblDiameter = value
            Case "Material=":       LocalMaterial = value
            Case "Comments=":       LocalRecommendations = value
            Case "SiteID=":         LocalSiteID = value
            Case Else
        End Select
        arrayIndex = arrayIndex + 1
        If arrayIndex > UBound(ConfigArray) Then Exit Do
        SectionDetail = ConfigArray(arrayIndex)
    Loop
    Call PopulateHeaderInfo
Exit Sub
Err_Handler:
MsgBox Err & "-I6:" & Error$


End Sub

Sub DoProcess(ByVal TheProcess As String)
On Error GoTo Err_Handler
    Dim Process As String
    Dim value As String
    Dim Frame As Long
    
    Call GetProcess(TheProcess, Process, value)
    Select Case Process
        
        Case "Graph"
            If PVDFileName <> "" Then
                Select Case value
                    Case "Capacity": Call ScreenDrawing.GraphSelect("Capacity", 0)
                    Case "Ovality": Call ScreenDrawing.GraphSelect("Ovality", 0)
                    Case "Delta": Call ScreenDrawing.GraphSelect("Delta", 0)
                    Case "XYDiameter": Call ScreenDrawing.GraphSelect("XYDiameter", 0)
                    Case "MaxDiameter": Call ScreenDrawing.GraphSelect("MaxDiameter", 0)
                    Case "Flat": Call ScreenDrawing.GraphSelect("Flat", 0)
                    Case "MedianDiameter": Call ScreenDrawing.GraphSelect("MedianDiameter", 0)
                    Case "MinDiameter": Call ScreenDrawing.GraphSelect("MinDiameter", 0) 'PCN4333
'PCN6458                     Case "Inclination": Call ScreenDrawing.GraphSelect("Inclination", 0) ' PCN6128
                End Select
            End If
        Case "GotoDistance"
            If PVDFileName <> "" Then
                PVFrameNo = PrecisionVisionGraph.GetFrameFromDistance(value)
                If Frame <> -1 Then
                    Call ClearLineScreen.GotoPVProfile(PVFrameNo, True)
                    Call PrecisionVisionGraph.MoveGraph(PVFrameNo)
                Else
                    Call AddErrorCode("No distance information")
                End If
            End If
            
        Case "GotoFrame"
            If PVDFileName <> "" Then
                Call ClearLineScreen.GotoPVProfile(value, True)
                Call PrecisionVisionGraph.MoveGraph(value)
            End If
            
        Case "TakeSnapShot"
            If VideoFileName <> "" Then
               Call ControlsScreen.ExecuteControlsFixedButton(5)
            End If
        Case "OpenReport"
            If PVDFileName <> "" Then
                
                If value = "Analysis" Then Call ControlsScreen.ExecuteReportsButton(0) 'Now called single
                If value = "Single" Then Call ControlsScreen.ExecuteReportsButton(0) 'Now called single
                
                If value = "PVGraph" Then Call ControlsScreen.ExecuteReportsButton(1)
                If value = "MultiLine" Then Call ControlsScreen.ExecuteReportsButton(1)
                
                If value = "Profile" Then Call ControlsScreen.ExecuteReportsButton(2)
                
                If value = "MultiProfile" Then Call ControlsScreen.ExecuteReportsButton(3)
              
            
            End If
        Case "ShowVideo"
            If VideoFileName <> "" Then
                Call ControlsScreen.ControlsView_Click(3)
            End If
        Case "ShowPVScreen"
            If PVDFileName <> "" Then
                Call ControlsScreen.ControlsView_Click(2)
            End If
        Case "ShowSnapshot"
            If PVDFileName <> "" Then
                Call ControlsScreen.ControlsView_Click(1)
            End If
        Case "Show3D"
            If PVDFileName <> "" Then
                Call ControlsScreen.ControlsView_Click(0)
            End If
    End Select
    
Exit Sub
Err_Handler:
MsgBox Err & "-I7:" & Error$
    
End Sub

Function GetProcess(ByVal MyString, Param, value)

On Error GoTo Err_Handler

Dim Loc As Integer, X As Integer

Loc = InStr(MyString, "[")
If Loc <> 0 Then
  Param = Left(MyString, Loc - 1)
  value = Trim(Mid(MyString, Loc + 1))
  value = Left(value, Len(value) - 1)
Else
    If MyString <> "" Then Param = MyString
    value = ""
End If

Exit Function
Err_Handler:
MsgBox Err & "-I8:" & Error$

End Function

Sub AddErrorCode(ByVal ErrorStr As String)
On Error GoTo Err_Handler

Dim NumberOffErrorCodes As Integer
    
    NumberOffErrorCodes = UBound(ErrorLog)
    NumberOffErrorCodes = NumberOffErrorCodes + 1
    ReDim Preserve ErrorLog(NumberOffErrorCodes)
    ErrorLog(NumberOffErrorCodes) = ErrorStr
    
Exit Sub
Err_Handler:
    Select Case Err
        Case 9: NumberOffErrorCodes = 1: Resume Next
    Case Else
        MsgBox Err & "-I9:" & Error$
    End Select

End Sub

Sub PopulateHeaderInfo()
On Error GoTo Err_Handler
Dim bool As Boolean

            PipelineDetails.AssetNo = LocalAssetNo 'PCN2239
            PipelineDetails.StartNodeNo = LocalStartNodeNo
            PipelineDetails.StartNodeLocation = LocalStartNodeLocation
            PipelineDetails.City = LocalCity
            PipelineDetails.FinishNodeNo = LocalFinishNodeNo
            PipelineDetails.FinishNodeLocation = LocalFinishNodeLocation
            PipelineDetails.sDate = LocalsDate
            PipelineDetails.sTime = LocalsTime
            PipelineDetails.PipeLength = LocalPipeLength
            If LocaldblDiameter <> "" And SafeCDbl(LocaldblDiameter) <> 0 Then 'PCN3922 when not filled in made it 0
                PipelineDetails.InternalDiameterExpected = SafeCDbl(LocaldblDiameter)
                Call PipelineDetails.InternalDiameterExpected_Validate(bool)  'PCN2241
            End If
            PipelineDetails.Material = LocalMaterial
            PipelineDetails.GeneralComments = LocalRecommendations
            PipelineDetails.SiteID = LocalSiteID
Exit Sub
Err_Handler:
    Select Case Err
    Case Else
        MsgBox Err & "-I10:" & Error$
    End Select
End Sub

Sub LoadFile()
On Error GoTo Err_Handler

    If strPVDName <> "" Then
        Call OpenAnyFile(strPVDName)
    ElseIf strVideoName <> "" Then
        Call OpenAnyFile(strVideoName)
    End If

Exit Sub
Err_Handler:
    Select Case Err
    Case Else
        MsgBox Err & "-I11:" & Error$
    End Select
End Sub

Function ExtractInterfaceFileName(ByVal strLoadPVDFile As String) As String
On Error GoTo Err_Handler

    Dim FileName As String
    Dim StringSearch As String
    Dim X As Long
    
    X = 3
    StringSearch = ""

'Split the path the the file
    Do Until Left(StringSearch, 1) = Chr(34) Or X > Len(strLoadPVDFile)
      StringSearch = Right(strLoadPVDFile, X)
      X = X + 1
    Loop
    
    FileName = Mid(StringSearch, 2, Len(StringSearch) - 2)
    
    If Dir(FileName) <> "" Then ExtractInterfaceFileName = FileName: Exit Function
    ExtractInterfaceFileName = ""

Exit Function
Err_Handler:
    Select Case Err
    Case 52: ExtractInterfaceFileName = "": Exit Function 'Bad Filename
    Case 5: ExtractInterfaceFileName = "": Exit Function 'Mid given bad parametors
    Case Else
        MsgBox Err & "-I12:" & Error$
 
    End Select
End Function

