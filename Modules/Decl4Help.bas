Attribute VB_Name = "Decl4Help"
Public Const HH_DISPLAY_TOPIC = &H0
Public Const HH_SET_WIN_TYPE = &H4
Public Const HH_GET_WIN_TYPE = &H5
Public Const HH_GET_WIN_HANDLE = &H6
Public Const HH_DISPLAY_TEXT_POPUP = &HE    ' Display string resource ID or
                                            ' text in a pop-up window.
Public Const HH_HELP_CONTEXT = &HF          ' Display mapped numeric value in
                                            ' dwData.
Public Const HH_TP_HELP_CONTEXTMENU = &H10  ' Text pop-up help, similar to
                                            ' WinHelp's HELP_CONTEXTMENU.
Public Const HH_TP_HELP_WM_HELP = &H11      ' text pop-up help, similar to
                                            ' WinHelp's HELP_WM_HELP.
Public Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
         (ByVal hwndCaller As Long, ByVal pszFile As String, _
         ByVal uCommand As Long, ByVal dwData As Long) As Long

