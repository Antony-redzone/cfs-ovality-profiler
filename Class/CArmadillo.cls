VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CArmadillo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

   '+---------------------------------------------------------------------+'
   '|                                                                     |'
   '|                     CArmadillo Wrapper Class                        |'
   '|                                                                     |'
   '| This class wraps the Armadillo registration/crack-protection shell  |'
   '| making it easy to retrive information from the shell or execute     |'
   '| commands on it.                                                     |'
   '|                                                                     |'
   '| This class is released as open source to assist others using the    |'
   '| Armadillo Shell from the Visual Basic environment and is free for   |'
   '| use in all programs, commercial and otherwise, as long as this      |'
   '| comment block remains intact.                                       |'
   '|                                                                     |'
   '| Armadillo is a product from Silicon Realms software.                |'
   '|                                                                     |'
   '+---------------------------------------------------------------------+'

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
Private Declare Function FormatMessageAPI Lib "kernel32" Alias "FormatMessageA" _
   (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, _
   ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, _
   Arguments As Long) As Long

'Put ArmAccess.DLL in your program's directory or somewhere on the path (such as the Windows directory).
'Now you can call the functions directly, and use Err.LastDllError to get any error codes if necessary

'If using one of these functions causes the certificate to expire (or detects that it's expired, due to
'date-change for instance), Armadillo simply sets the EXPIRED variable. It is up to your program to
'check for and deal with this.
'
'These functions set the last error value (retrieved with the Windows API function GetLastError()). This
'lets you determine why the call failed, if you wish to. Note that it is usually enough to check the
'return value; these codes are provided mostly for debugging purposes. The following values
'may be returned by GetLastError():
Private Declare Function ArmCheckCode Lib "ArmAccess.dll" Alias "CheckCode" (ByVal name As String, ByVal Code As String) As Byte
Private Declare Function ArmVerifyKey Lib "ArmAccess.dll" Alias "VerifyKey" (ByVal name As String, ByVal Code As String) As Byte
Private Declare Function ArmInstallKey Lib "ArmAccess.dll" Alias "InstallKey" (ByVal name As String, ByVal Code As String) As Byte
Private Declare Function ArmInstallKeyLater Lib "ArmAccess.dll" Alias "InstallKeyLater" (ByVal name As String, ByVal Code As String) As Byte
Private Declare Function ArmUninstallKey Lib "ArmAccess.dll" Alias "UninstallKey" () As Byte
Private Declare Function ArmSetDefaultKey Lib "ArmAccess.dll" Alias "SetDefaultKey" () As Byte
Private Declare Function ArmUpdateEnvironment Lib "ArmAccess.dll" Alias "UpdateEnvironment" () As Byte
Private Declare Function ArmIncrementCounter Lib "ArmAccess.dll" Alias "IncrementCounter" () As Byte
Private Declare Function ArmCopiesRunning Lib "ArmAccess.dll" Alias "CopiesRunning" () As Long
Private Declare Function ArmChangeHardwareLock Lib "ArmAccess.dll" Alias "ChangeHardwareLock" () As Byte
Private Declare Function ArmGetShellProcessID Lib "ArmAccess.dll" Alias "GetShellProcessID" () As Long

' SetWindowPos Flags
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

' SetWindowPos() hWndInsertAfter values
Private Const HWND_TOP = 0
Private Const HWND_BOTTOM = 1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Public Enum ArmadilloErrors
   ERROR_SUCCESS = 0
   ERROR_BAD_ENVIRONMENT = 10
   ERROR_INVALID_DATA = 13
   ERROR_BROKEN_PIPE = 109
   ERROR_BUSY = 170
   ERROR_ACCOUNT_RESTRICTION = 1327
   ERROR_ACCOUNT_DISABLED = 1331
   ERROR_LICENSE_QUOTA_EXCEEDED = 1395
   ERROR_ACCOUNT_EXPIRED = 1793
   ERROR_INVALID_DATATYPE = 1804
   ERROR_ACCOUNT_LOCKED_OUT = 1909
   ERROR_BAD_USERNAME = 2202
End Enum
Private m_bInIDE As Boolean


Private Function ErrCodeToString(ByVal ErrCode As Long) As String
   On Error GoTo Hell
   Dim strTemp As String
   
   Select Case ErrCode
       Case ERROR_SUCCESS: strTemp = "Operation was completed successfully."
       Case ERROR_BAD_ENVIRONMENT: strTemp = "Can't establish communications with the Armadillo shell. This usually means that you haven't protected the program."
       Case ERROR_INVALID_DATA: strTemp = "There was an error in the communications path. You might be able to get it to work if you try the call a second time. If you receive this error again, it's probably unrecoverable."
       Case ERROR_BROKEN_PIPE: strTemp = "Could not open the communications path. This is probably an unrecoverable error."
       Case ERROR_BUSY: strTemp = "Returned only by the CopiesRunning function, if the network is too busy to get a response. Should almost never be seen in the real world."
       Case ERROR_ACCOUNT_RESTRICTION: strTemp = "The certificate has 'Limit Key Validity' set, and this key is either past the date it is allowed to be installed, or Armadillo has detected a problem with the system clock."
       Case ERROR_ACCOUNT_DISABLED: strTemp = "The key is an upgrade key, and the user does not have an existing upgradable key installed."
       Case ERROR_LICENSE_QUOTA_EXCEEDED: strTemp = "The key was successfully installed, but there are now more computers using this key than it is licensed for. Please see the discussion on the Limit Copies option for more information."
       Case ERROR_ACCOUNT_EXPIRED: strTemp = "The name/key you specified was valid at one time, but is now expired. It should not be considered a usable key."
       Case ERROR_INVALID_DATATYPE: strTemp = "The certificate in use is not expire-by-uses (IncrementCounter, or the key installed is not a hardware locked key (UninstallKey."
       Case ERROR_ACCOUNT_LOCKED_OUT: strTemp = "This key is listed in the Stolen Codes Database for the program."
       Case ERROR_BAD_USERNAME: strTemp = "The name/key you specified is not valid for this program."
       Case Else
           strTemp = FormatMessage(ErrCode) '//Assume it's from the Windows API
           Debug.Assert False
   End Select
   
   ErrCodeToString = strTemp

Exit_Proc:
   Exit Function

Hell:
   ErrorIn "CArmadillo.ErrCodeToString(ErrCode)", ErrCode, EA_default
End Function

Private Function FormatMessage(ByVal ErrNum As String) As String
   Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
       Dim strBuffer As String * 512, strMsg As String
       On Error GoTo errHandler
   
   FormatMessageAPI FORMAT_MESSAGE_FROM_SYSTEM, Null, ErrNum, 0, strBuffer, 512, 0
   strMsg = strBuffer
           'Strange but necessary manipulations
   strMsg = Replace(strMsg, vbNewLine, "")
   strMsg = Replace(strMsg, Chr(0), "")
   FormatMessage = strMsg
   
   Exit Function
errHandler:
       'Nothing to do...
End Function

Private Function GetEnvironString(name As String) As String
   On Error GoTo Hell
   'Define a string for the return value
   Dim ReturnString As String

   'First, pad the return value string so that it's big enough.
   ReturnString = Space$(256)

   'Now call the GetEnvironmentVariable function. If the function indicates
   'that the variable doesn't exist, return an empty string. Otherwise,
   'return the contents of the variable.
   If GetEnvironmentVariable(name, ReturnString, 256) <> 0 Then
       GetEnvironString = Trim$(StripNulls(ReturnString))
   Else
       GetEnvironString = ""
   End If

Exit_Proc:
   Exit Function

Hell:
   ErrorIn "CArmadillo.GetEnvironString(Name)", name, EA_default
End Function

Public Function hWndLoading() As Long
   On Error GoTo Hell
   Dim strTemp As String

   'LOADINGWINDOW: If you use a custom bitmap loading window, this variable holds the
   'handle of this window, in hexadecimal format. You can use this handle to close or
   'hide the loading window when your program starts up, if you don't wish to rely on Armadillo's timer.
   
   'You could also issue a WM_CLOSE to the Window, but I thought it to be safer to leave the hWnd valid
   'for Armadillo to close the Window when the timer expires.
   
   'You should call this method when the main window of your application is fully loaded to remove the
   'loading window if you are using it as the Splash screen of your application.
   
   strTemp = GetEnvironString("LOADINGWINDOW")
   If Len(strTemp) > 0 Then
       hWndLoading = CLng("&H" & strTemp)
   End If

Exit_Proc:
   Exit Function

Hell:
   ErrorIn "CArmadillo.hWndLoading(strTemp)", strTemp, EA_default
End Function

Public Sub HideLoadingWindow()
   On Error GoTo Hell
   Dim hwnd As Long
   Dim strTemp As String

   'LOADINGWINDOW: If you use a custom bitmap loading window, this variable holds the
   'handle of this window, in hexadecimal format. You can use this handle to close or
   'hide the loading window when your program starts up, if you don't wish to rely on Armadillo's timer.
   
   'You could also issue a WM_CLOSE to the Window, but I thought it to be safer to leave the hWnd valid
   'for Armadillo to close the Window when the timer expires.
   
   'You should call this method when the main window of your application is fully loaded to remove the
   'loading window if you are using it as the Splash screen of your application.
   
   hwnd = Me.hWndLoading
   
   If hwnd <> 0 Then
       SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_HIDEWINDOW
   End If
   
'    MsgBox "window hwnd = " & hwnd


Exit_Proc:
   Exit Sub

Hell:
   ErrorIn "CArmadillo.HideLoadingWindow(hwnd)", hwnd, EA_default
End Sub

Public Sub ARMRegister()
   On Error GoTo Hell

   'PROTECTEDFILE: This variable holds the path and filename of the (protected) EXE file.
   'If your program needs to know where the original EXE file is (for instance, to locate
   'some data files that are in the same directory), it should use this variable's contents
   'instead of relying on the Windows API call GetModuleFileName(). The reason:
   'GetModuleFileName() returns the name of the actual file. In the case of an
   'Armadillo-protected program, this will be the temporary file (usually with the
   'extension .TMP0 instead of EXE), and if your program is run from a read-only
   'directory or CD, Armadillo is forced to create this file in the Windows temporary
   'directory instead of the one where the program is stored.

   Dim strTemp As String
   
   
'   If IsInIDE Then Exit Sub
   strTemp = GetEnvironString("PROTECTEDFILE")
   If Len(strTemp) > 0 Then
       Shell strTemp & " REGISTER", vbNormalFocus
   Else
        'MsgBox DisplayMessage("The registration attempt failed.  Registration shell is not present.") & vbCrLf & vbCrLf & DisplayMessage("Please contact technical support"), vbCritical
        ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("The registration attempt failed.  Registration shell is not present.") & vbCrLf & vbCrLf & DisplayMessage("Please contact technical support"): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
   End If
   
Exit_Proc:
   Exit Sub
   

Hell:
   ErrorIn "CArmadillo.Register", , EA_default
End Sub

Public Sub ShowRegInfo()
   On Error GoTo Hell

   Dim strTemp As String
   
   
   strTemp = GetEnvironString("PROTECTEDFILE")
   If Len(strTemp) > 0 Then
       Shell strTemp & " INFO", vbNormalFocus
   Else
       GoTo Hell
   End If
   
Exit_Proc:
   Exit Sub

Hell:
   ErrorIn "CArmadillo.ShowRegInfo", , EA_default
End Sub

Public Property Get HardwareFingerPrint(Optional ByVal Enhanced As Boolean = False) As String
   On Error GoTo Hell
   'ENHFINGERPRINT: Similar to the FINGERPRINT variable documented below, this
   'one holds the enhanced hardware-locking fingerprint, for certificates that use enhanced hardware locking.

   'FINGERPRINT: This variable appears only if your program includes at least one
   'hardware-locked certificate. It contains the eight-digit hexadecimal
   '"machine fingerprint" of the computer your program is currently running
   'on, which is necessary to make a hardware-locked key.
   
   If Enhanced Then
       HardwareFingerPrint = GetEnvironString("ENHFINGERPRINT")
   Else
       HardwareFingerPrint = GetEnvironString("FINGERPRINT")
   End If

Exit_Proc:
   Exit Property

Hell:
   ErrorIn "CArmadillo.HardwareFingerPrint(Enhanced)", Enhanced, EA_default
End Property

Public Property Get IsRegistered() As Boolean
   On Error GoTo Hell
   '//ASSUMES: when the 'Default' certificate is installed, the software is UNregistered.
   '//When any other certificate is installed, the software is registered.
   
#If INTERNAL_BUILD Then
   IsRegistered = True
   Exit Property
#End If
   
   IsRegistered = CBool(Len(Me.UserName) > 0)
   
Exit_Proc:
   Exit Property

Hell:
   ErrorIn "CArmadillo.IsRegistered", , EA_default
End Property
   
Public Property Get IsRegisteredEx(ByVal ExtraInfoMatch As Integer) As Boolean
   On Error GoTo Hell
   '//Returns TRUE only if the product is REGISTERED, and if the stored ExtraInfo
   '//matches the ExtraInfo passed into the method.

#If INTERNAL_BUILD Then
   IsRegisteredEx = True
   Exit Property
#End If
   
   If Me.IsRegistered Then
       IsRegisteredEx = CBool(Me.ExtraInfo = ExtraInfoMatch)
   End If

Exit_Proc:
   Exit Property

Hell:
   ErrorIn "CArmadillo.IsRegisteredEx", , EA_default
End Property


Private Function StripNulls(OriginalStr As String) As String
   On Error GoTo Hell

' Strips any trailing nulls from path names retrieved
' from the registry. This function is found in the
' following Microsoft(r) knowledge base articles:
' Q183009 "HOWTO: Enumerate Windows Using the WIN32 API"
' Q185476 "HOWTO: Search Directories to Find or List Files"
' Q190218 "HOWTO: Retrieve Settings From a Printer Driver"

 If (InStr(OriginalStr, Chr(0)) > 0) Then
   OriginalStr = Left(OriginalStr, _
       InStr(OriginalStr, Chr(0)) - 1)
 End If
 
 StripNulls = OriginalStr

Exit_Proc:
   Exit Function

Hell:
   ErrorIn "CArmadillo.StripNulls(OriginalStr)", OriginalStr, EA_default
End Function


Public Property Get UserName() As String
   On Error GoTo Hell
   Dim strTemp As String
   
   
   'USERNAME: This is the name on the user's key, or DEFAULT if they're using the default certificate.
   'Note that this string is used by some other programs; for compatibility, the user's name is also
   'stored in the ALTUSERNAME string. Note: If the registry information is somehow corrupted, and
   'you were using an older version of Armadillo before, then the user's name may be lost. In this case,
   'it will appear as the string "MISSING" (without the quotes) instead. All other information, including
   'the number of days or uses remaining, will remain.
   
   
#If INTERNAL_BUILD Then
   UserName = "Milori Insider"
   Exit Property
#End If
   
   strTemp = Trim$(GetEnvironString("ALTUSERNAME"))
   If UCase(strTemp) = "DEFAULT" Then strTemp = ""
   If IsInIDE And Len(strTemp) = 0 Then strTemp = "Program Author"
   If Len(strTemp) > 0 Then
       UserName = strTemp
   End If
   
Exit_Proc:
   Exit Property

Hell:
   ErrorIn "CArmadillo.UserName", , EA_default
End Property


Public Property Get DaysInstalled() As Long
   On Error GoTo Hell
   Dim strTemp As String
   
   
   'DAYSINSTALLED: This variable contains the number of days since the current key was
   'installed. On the day of installation, it is zero.
   
   strTemp = GetEnvironString("DAYSINSTALLED")
   If Len(strTemp) > 0 Then
       DaysInstalled = CLng(strTemp)
   Else
       DaysInstalled = 0
   End If
   
Exit_Proc:
   Exit Property

Hell:
   ErrorIn "CArmadillo.DaysInstalled", , EA_default
End Property

Public Property Get ExtraInfo() As Integer
   On Error GoTo Hell
   Dim strTemp As String
   
   'EXTRAINFO: Any "extra information" you have stored in the key is stored in this
   'variable, in decimal format. If the certificate uses the extra bits for another
   'purpose, such as storing expiration information in the key, then this variable will contain a zero.
   
   strTemp = GetEnvironString("EXTRAINFO")
   If Len(strTemp) > 0 Then
       ExtraInfo = CInt(strTemp)
   Else
       ExtraInfo = 0
   End If
   
Exit_Proc:
   Exit Property

Hell:
   ErrorIn "CArmadillo.ExtraInfo", , EA_default
End Property

Public Property Get KeyCreated() As String
   On Error GoTo Hell
   Dim strTemp As String
   
   'KEYCREATED: This contains the date, in the form "YYYY.MM.DD", that the user's key was created on.
   'It is not set for the default key.
   
   KeyCreated = GetEnvironString("KEYCREATED")
   
Exit_Proc:
   Exit Property

Hell:
   ErrorIn "CArmadillo.KeyCreated", , EA_default
End Property

Public Property Get UserKey() As String
   On Error GoTo Hell
   Dim strTemp As String
   
   'USERKEY: This is the key that the program is currently running under. It is not set for
   'the default key, and should not be trusted for expire-by-uses keys entered before version 1.73.

   UserKey = GetEnvironString("USERKEY")
   If IsInIDE And Len(UserKey) = 0 Then UserKey = "Development"
   
Exit_Proc:
   Exit Property

Hell:
   ErrorIn "CArmadillo.UserKey", , EA_default
End Property

Public Property Get DaysLeft() As Long
   On Error GoTo Hell
   Dim strTemp As String
   
   
   'DAYSLEFT: For date- and day-limited certificates, this is set to the number of valid
   'days left on the key. When this is set to 1, the key will expire at midnight.
   
   strTemp = GetEnvironString("DAYSLEFT")
   If Len(strTemp) > 0 Then
       DaysLeft = CLng(strTemp)
   Else
       DaysLeft = 0
   End If
   
Exit_Proc:
   Exit Property

Hell:
   ErrorIn "CArmadillo.DaysLeft", , EA_default
End Property

Public Property Get CopiesAllowed() As Long
   On Error GoTo Hell
   Dim strTemp As String
   
   
   'COPIESALLOWED: If your program uses the Limit Copies feature, this variable is set to the
   'number of copies permitted for the current certificate and key.
   
   strTemp = GetEnvironString("COPIESALLOWED")
   If Len(strTemp) > 0 Then
       CopiesAllowed = CLng(strTemp)
   Else
       CopiesAllowed = 0
   End If
   
Exit_Proc:
   Exit Property

Hell:
   ErrorIn "CArmadillo.CopiesAllowed", , EA_default
End Property

Public Property Get DaysInstalledBit(ByVal BitNumber As Byte) As Boolean
   On Error GoTo Hell
   Dim strTemp As String
   
   
   'DAYSINSTALLEDBITx: This is actually fifteen different variables, DAYSINSTALLEDBIT1 through DAYSINSTALLEDBIT15.
   'They contain the number of days since the corresponding "extra information" bit was last changed,
   'and can be used with Modification Keys to set up a trial period for separate modules in your program.
   
   strTemp = GetEnvironString("DAYSINSTALLEDBIT" & CStr(BitNumber))
   If Len(strTemp) > 0 Then
       DaysInstalledBit = CBool(strTemp)
   Else
       DaysInstalledBit = False
   End If
   
Exit_Proc:
   Exit Property

Hell:
   ErrorIn "CArmadillo.DaysInstalledBit(BitNumber)", BitNumber, EA_default
End Property

Public Property Get Expired() As Boolean
   On Error GoTo Hell
   Dim strTemp As String
   
   
   'EXPIRED: This variable is set only if you've chosen the "don't show expiration message"
   '(on the Expire window) and the certificate has expired. Armadillo will let the program run
   'after expiration if that box is checked, the program itself must check for this variable.
   
   strTemp = GetEnvironString("EXPIRED")
   If Len(strTemp) > 0 Then
       Expired = CBool(strTemp)
   Else
       Expired = False
   End If
   
Exit_Proc:
   Exit Property

Hell:
   ErrorIn "CArmadillo.Expired", , EA_default
End Property

Public Property Get ExpiredVersion() As String
   On Error GoTo Hell
   Dim strTemp As String
   
   
   'EXPIREVER: For the expire-by-version certificates, this tells you what version the key will expire on.
   
   strTemp = GetEnvironString("EXPIREVER")
   If Len(strTemp) > 0 Then
       ExpiredVersion = strTemp
   Else
       ExpiredVersion = ""
   End If
   
Exit_Proc:
   Exit Property

Hell:
   ErrorIn "CArmadillo.ExpiredVersion", , EA_default
End Property

Public Property Get UsesLeft() As Long
   On Error GoTo Hell
   Dim strTemp As String
   
   
   'USESLEFT: For certificates offering a limited number-of-uses, this is set to the number
   'of uses remaining. If the user is on the last use, this is set to 1.
   
   strTemp = GetEnvironString("USESLEFT")
   If Len(strTemp) > 0 Then
       UsesLeft = CLng(strTemp)
   Else
       UsesLeft = 0
   End If
   
Exit_Proc:
   Exit Property

Hell:
   ErrorIn "CArmadillo.UsesLeft", , EA_default
End Property

Public Property Get ClockBack() As Boolean
   On Error GoTo Hell
   Dim strTemp As String
   
   
   'CLOCKBACK: If you have chosen to disable Armadillo's clock-back checking
   '(by way of the Don't Report Clock-Back on the Project window), this variable
   'will be set when Armadillo detects a problem with an incorrect system clock.
   
   strTemp = GetEnvironString("CLOCKBACK")
   If Len(strTemp) > 0 Then
       ClockBack = CBool(strTemp)
   Else
       ClockBack = False
   End If
   
Exit_Proc:
   Exit Property

Hell:
   ErrorIn "CArmadillo.ClockBack", , EA_default
End Property

Public Property Get InvalidKey() As Boolean
   On Error GoTo Hell
   Dim strTemp As String
   
   
   'INVALIDKEY: This variable is only used if you have the Auto-revert on invalid key
   'option set, and the key stored on a user's system is invalid (probably because you've
   'changed an encryption template). It is only set during the first such run; after that,
   'the user will be automatically reset to the default certificate, if any. This is
   'intended to let your program gracefully warn the user of the change.
   
   strTemp = GetEnvironString("INVALIDKEY")
   If Len(strTemp) > 0 Then
       InvalidKey = CBool(strTemp)
   Else
       InvalidKey = False
   End If
   
Exit_Proc:
   Exit Property

Hell:
   ErrorIn "CArmadillo.InvalidKey", , EA_default
End Property


'Do It Yourself: Integrating Armadillo With Your Programs Using ArmAccess.DLL
'You have the option of bypassing all of Armadillo's screens and replacing them with your own if you wish
'by using the included ArmAccess.DLL file.

'What it does
'ArmAccess.DLL (which you can distribute with your programs free of charge) is a small dynamic-link
'library which hooks into the Armadillo shell around your program, allowing you two-way communication
'with it. This lets you create your own registration dialogs instead of using Armadillo's default
'"register" screen, among other things.

'The Functions
'ArmAccess.DLL exports several functions. Most of them return non-zero (a Boolean 'true' in C) on success,
'or zero if there is an error or they cannot carry out the requested function for some reason.

Public Function InstallKey(ByVal name As String, ByVal Key As String, Optional ByVal InstallOnExit As Boolean = False) As Boolean
   On Error GoTo Hell
   'CheckCode: Checks a name/key pair, and (if valid) stores it as the new key. Returns non-zero if the code is valid.
   'InstallKey: This is identical to CheckCode. You can use either name.
   'InstallKeyLater: This function, like VerifyKey (below), checks whether a key is good and returns non-zero
   '(and sets environment variables) if it is. It doesn't install the key until that copy of the program exits though.

   If InstallOnExit Then
       InstallKey = CBool(ArmInstallKeyLater(name, Key))
   Else
       InstallKey = CBool(ArmInstallKey(name, Key))
   End If

Exit_Proc:
   Exit Function

Hell:
   ErrorIn "CArmadillo.InstallKey(Name,Key,InstallOnExit)", Array(name, Key, InstallOnExit), _
        EA_default
End Function

Public Function LastErrorS() As String
   On Error GoTo Hell
   LastErrorS = ErrCodeToString(Err.LastDllError)

Exit_Proc:
   Exit Function

Hell:
   ErrorIn "CArmadillo.LastErrorS", , EA_default
End Function

Public Function UninstallKey() As String
   On Error GoTo Hell
   'UninstallKey: Securely uninstalls a hardware-locked key, and change the hardware "fingerprint"
   'so that it can't be reinstalled. It also gives you an uninstall code (in the otherwise-undocumented
   'environment variable UNINSTALLCODE), so that you can manually verify that the key has been uninstalled.
   
   '//This function uninstalls the license and returns the uninstallcode
   If CBool(ArmUninstallKey) Then
       UninstallKey = GetEnvironString("UNINSTALLCODE")
   End If

Exit_Proc:
   Exit Function

Hell:
   ErrorIn "CArmadillo.UninstallKey", , EA_default
End Function

Public Function IsValidKey(ByVal name As String, ByVal Key As String, Optional ByVal UpdateEnvironment As Boolean = False) As Boolean
   On Error GoTo Hell
   'VerifyKey: This function is similar to CheckCode/InstallKey. It checks a name/key pair,
   'returning non-zero if it's valid and could be installed. However, unlike CheckCode/InstallKey,
   'it does not install the key; the original key remains installed. After calling VerifyKey,
   'the environment variables are set as if the key were installed; to recover the environment
   'variables for the original key, call UpdateEnvironment.
   
   If CBool(ArmVerifyKey(name, Key)) Then
       IsValidKey = True
       If UpdateEnvironment Then
           '//Do nothing.  This happens automatically.
       Else
           '//Restore the environment back the way that it was, if the user didn't
           '//intend to update the environment.
           ArmUpdateEnvironment
       End If
   Else
       IsValidKey = False
   End If


Exit_Proc:
   Exit Function

Hell:
   ErrorIn "CArmadillo.IsValidKey(Name,Key,UpdateEnvironment)", Array(name, Key, UpdateEnvironment), _
        EA_default
End Function

Public Function RestoreDefaultKey() As Boolean
   On Error GoTo Hell
   'SetDefaultKey: Calling this function will force the Armadillo shell to revert to the default
   'certificate (if any) for your program.
   RestoreDefaultKey = CBool(ArmSetDefaultKey)

Exit_Proc:
   Exit Function

Hell:
   ErrorIn "CArmadillo.RestoreDefaultKey", , EA_default
End Function

Public Function UpdateEnvironment() As Boolean
   On Error GoTo Hell
   'UpdateEnvironment: You can call this function to have Armadillo update the environment strings it sets.
   'This is intended for programs that are often left running for days at a time, and lets you check
   'whether the current key is still valid.
   UpdateEnvironment = CBool(ArmUpdateEnvironment)

Exit_Proc:
   Exit Function

Hell:
   ErrorIn "CArmadillo.UpdateEnvironment", , EA_default
End Function

Public Function IncrementCounter() As Boolean
   On Error GoTo Hell
   'IncrementCounter: This function allows you to control the "usage count" of an Armadillo
   'expire-by-uses certificate. Every time it's called, it increases the usage count by one
   'and updates the environment. If you're using this function to count uses, and don't
   'want Armadillo to automatically increment the counter when the program starts, then set
   'the certificate to expire "After number of uses (counted by DLL)".
   IncrementCounter = CBool(ArmIncrementCounter)

Exit_Proc:
   Exit Function

Hell:
   ErrorIn "CArmadillo.IncrementCounter", , EA_default
End Function

Public Function CopiesRunning() As Long
   On Error GoTo Hell
   'CopiesRunning: Mainly used with network licenses, this function returns the number of copies running
   'on the network at the time it is called (including the calling copy). Returns zero if there is an error.
   'Note that this function only reports the number of copies that are using the same key as the one
   'that makes the request.
   CopiesRunning = ArmCopiesRunning

Exit_Proc:
   Exit Function

Hell:
   ErrorIn "CArmadillo.CopiesRunning", , EA_default
End Function

Public Function ChangeHardwareLock() As Boolean
   On Error GoTo Hell
   'ChangeHardwareLock: Used only with hardware-locked keys, this function randomly changes the
   'hardware "fingerprint" for a machine, just as if the TRANSFER command were used.
   ChangeHardwareLock = CBool(ArmChangeHardwareLock)

Exit_Proc:
   Exit Function

Hell:
   ErrorIn "CArmadillo.ChangeHardwareLock", , EA_default
End Function

Public Function PID() As Long
   On Error GoTo Hell
   'GetShellProcessID: Returns the Process ID of the Armadillo shell around your program,
   'for some web-update programs and other advanced usage.
   PID = ArmGetShellProcessID

Exit_Proc:
   Exit Function

Hell:
   ErrorIn "CArmadillo.PID", , EA_default
End Function


Private Function DoIDECheck() As Boolean

'this will only be called in we are running inside the IDE
m_bInIDE = True
DoIDECheck = True


End Function

Private Function IsInIDE() As Boolean


'are we running inside the VB IDE ?
Debug.Assert (DoIDECheck())
IsInIDE = m_bInIDE

End Function

Public Property Get ClearLineRegType() As String
   On Error GoTo Hell
   Dim strTemp As String
   
   
   
   strTemp = GetEnvironString("CLEARLINEREGTYPE")
   
   ClearLineRegType = strTemp
   
Exit_Proc:
   Exit Property

Hell:
   ErrorIn "CArmadillo.InvalidKey", , EA_default
End Property

Public Function SetEnvironString(name As String, value As String) As String
   On Error GoTo Hell
   
Call SetEnvironmentVariable(name, value)
   
Exit_Proc:
   Exit Function

Hell:
   ErrorIn "CArmadillo.GetEnvironString(Name)", name, EA_default
End Function

Public Property Get GetVersionNumber() As String
   On Error GoTo Hell
   Dim strTemp As String
   
   
   
   strTemp = GetEnvironString("VERSIONNUMBER")
   
   GetVersionNumber = strTemp
   
Exit_Proc:
   Exit Property

Hell:
   ErrorIn "CArmadillo.InvalidKey", , EA_default
End Property

