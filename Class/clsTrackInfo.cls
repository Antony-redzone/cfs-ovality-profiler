VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTrackInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public hwnd As Long
Public HoverTime As Long
Public prevProc As Long

Public Event MouseHover()
Public Event MouseLeave()


Public Sub RaiseMouseHover()
RaiseEvent MouseHover
End Sub

Public Sub RaiseMouseLeave()
RaiseEvent MouseLeave
End Sub

Private Sub Class_Initialize()
HoverTime = 400
End Sub
