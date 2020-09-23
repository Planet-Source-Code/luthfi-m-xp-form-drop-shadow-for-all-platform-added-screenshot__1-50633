VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DropShadow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_hWnd As Long
Private m_MsgHandler As cMsgHandler 'Our message handler

Public Sub CreateShadow(ByVal hwnd As Long, _
  Optional ByVal Offset As Long = 8&)
  
  If IsWindow(hwnd) = 0 Then Exit Sub
  
  'Child window is not allowed to have shadow
  If (GetWindowLong(hwnd, GWL_STYLE) And WS_CHILD) = WS_CHILD Then Exit Sub
  
  Dispose
  
  Set m_MsgHandler = New cMsgHandler
  With m_MsgHandler
    .Offset = Offset
    .hwnd = hwnd
  End With
  m_hWnd = hwnd
End Sub

Private Sub Dispose()
  If m_hWnd = 0& Then Exit Sub
  Set m_MsgHandler = Nothing
  m_hWnd = 0&
End Sub

