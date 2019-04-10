Attribute VB_Name = "Utilities4Form"
 Option Explicit
Private Declare Function TrackMouseEvent Lib "user32" (ByRef lpEventTrack As tagTRACKMOUSEEVENT) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias _
"CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
   ByVal hwnd As Long, ByVal msg As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long

'Types
Private Type tagTRACKMOUSEEVENT
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type

'Constants
Private Const TME_HOVER = &H1
Private Const TME_LEAVE = &H2
Private Const TME_CANCEL = &H80000000
Private Const HOVER_DEFAULT = &HFFFFFFFF
Private Const WM_MOUSELEAVE = &H2A3
Private Const WM_MOUSEHOVER = &H2A1
Private Const WM_MOUSEMOVE = &H200
Private Const GWL_WNDPROC = (-4)

'Variables
Dim trackCol As Collection

Public Sub StartTrack(trak As clsTrackInfo)
On Error GoTo Err_Handler
Dim prevProc As Long

If trackCol Is Nothing Then
    Set trackCol = New Collection
End If

trak.prevProc = SetWindowLong(trak.hwnd, GWL_WNDPROC, AddressOf WindowProc)
trackCol.Add trak, CStr(trak.hwnd)

RequestTracking trak

Exit Sub
Err_Handler:
    MsgBox Err & "-U1:" & Error$
End Sub
Public Sub EndTrack(trak As clsTrackInfo)
On Error GoTo Err_Handler

If trackCol Is Nothing Then Exit Sub

Call SetWindowLong(trak.hwnd, GWL_WNDPROC, trak.prevProc)

Dim trk As tagTRACKMOUSEEVENT
trk.cbSize = 16
trk.dwFlags = TME_LEAVE Or TME_HOVER Or TME_CANCEL
trk.hwndTrack = trak.hwnd
TrackMouseEvent trk

trackCol.Remove CStr(trak.hwnd)
If trackCol.Count = 0 Then
    Set trackCol = Nothing
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-T2:" & Error$
End Sub

Private Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo Err_Handler
Dim trak As clsTrackInfo

Set trak = trackCol.Item(CStr(hwnd))

If uMsg = WM_MOUSELEAVE Then
    trak.RaiseMouseLeave
ElseIf uMsg = WM_MOUSEHOVER Then
    trak.RaiseMouseHover
ElseIf uMsg = WM_MOUSEMOVE Then
    RequestTracking trak
    WindowProc = CallWindowProc(trak.prevProc, hwnd, uMsg, wParam, lParam)
Else
    WindowProc = CallWindowProc(trak.prevProc, hwnd, uMsg, wParam, lParam)
    'Debug.Print uMsg
End If

Exit Function
Err_Handler:
'    MsgBox Err & " - " & error$
End Function

Private Sub RequestTracking(trak As clsTrackInfo)
Dim trk As tagTRACKMOUSEEVENT
trk.cbSize = 16
trk.dwFlags = TME_LEAVE Or TME_HOVER
trk.dwHoverTime = trak.HoverTime
trk.hwndTrack = trak.hwnd

TrackMouseEvent trk

Exit Sub
Err_Handler:
    MsgBox Err & "-T3:" & Error$
End Sub



Function IsOpen(FormName As String) As Boolean
On Error GoTo Err_Handler

Dim I As Integer

For I = 0 To Forms.Count - 1
    If StrComp(Forms(I).name, FormName) = 0 Then
        IsOpen = True
        Exit Function
    End If
Next
IsOpen = False

Exit Function
Err_Handler:
    MsgBox Err & "-T4:" & Error$
End Function


