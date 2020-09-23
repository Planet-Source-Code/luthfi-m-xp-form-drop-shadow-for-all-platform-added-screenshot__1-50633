VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMsgHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Implements IWinSub

Private m_hWndP As Long
Private m_hWndShadow As Long
Private m_ParentSub As cWinSub
Private m_ShadowSub As cWinSub
Private m_Offset As Long
Private m_DeskDC As MemoryDC
Private m_DC As MemoryDC
Private m_Gradient As Gradient

Public Property Let Offset(ByVal Value As Long)
  If m_Offset <> Value Then
    m_Offset = Value
    SynchronizeShadow WM_MOVE
  End If
End Property

Public Property Let hwnd(ByVal Value As Long)
  Dim tR As RECT
  
  Dispose
  
  m_hWndP = Value
  
  Set m_ParentSub = New cWinSub
  
  With m_ParentSub
    'we want tob notified when parent window is sized or moved
    .AddAfterMsgs WM_SIZE, WM_MOVE
    'Notify when to Show/hide shadow
    .AddBeforeMsgs WM_ACTIVATE
    .Subclass m_hWndP, Me
  End With
  
  GetWindowRect m_hWndP, tR
  
  CreateShadow tR
End Property

Private Sub Dispose()
  If m_hWndP = 0& Then Exit Sub
  
  'Destroy Shadow window
  If m_hWndShadow <> 0& Then
    Set m_ShadowSub = Nothing
    ShowWindow m_hWndShadow, SW_HIDE
    Set m_DC = Nothing
    Set m_Gradient = Nothing
    SendMessage m_hWndShadow, WM_CLOSE, 0, ByVal 0
  End If
  Set m_ParentSub = Nothing
  m_hWndP = 0&
  Set m_DeskDC = Nothing
End Sub

Private Sub CreateShadow(tR As RECT)
  Dim lStyle As Long
  
  'Create the shadow with the same size as our shadowed window
  'And position it just behind the shadowed window
  
  lStyle = WS_POPUP And Not WS_BORDER
  
  m_hWndShadow = CreateWindowEx(WS_EX_TOOLWINDOW, "Static", "", lStyle, _
    tR.Left + m_Offset, tR.Top + m_Offset, tR.Right - tR.Left, tR.Bottom - tR.Top, _
    0, 0, App.hInstance, ByVal 0)
  
  If m_hWndShadow = 0 Then Exit Sub
  
  Set m_ShadowSub = New cWinSub
  
  With m_ShadowSub
    .AddBeforeMsgs WM_PAINT, WM_ERASEBKGND
    .Subclass m_hWndShadow, Me
  End With
  'Initialize our back buffer
  Set m_DC = New MemoryDC
  Set m_Gradient = New Gradient
  
  ShowWindow m_hWndShadow, SW_SHOWNOACTIVATE
End Sub

Private Sub SynchronizeShadow(ByVal uMsg As Long)
  Dim tR As RECT
  If m_hWndP <> 0& Then
    GetWindowRect m_hWndP, tR
    
    If uMsg = WM_SIZE Then 'Recreate DC only if size has changed
      m_DC.CreateDC tR.Right - tR.Left, tR.Bottom - tR.Top
    End If
    
    'Move/Size our shadow
    SetWindowPos m_hWndShadow, 0, tR.Left + m_Offset, tR.Top + m_Offset, _
      tR.Right - tR.Left, tR.Bottom - tR.Top, SWP_NOACTIVATE Or SWP_NOZORDER
    
    'Generate WM_PAINT message
    RedrawWindow m_hWndShadow, ByVal 0, 0, RDW_INVALIDATE Or RDW_UPDATENOW
  
  End If
  
End Sub

Private Sub CaptureDesktop()
'Capture the Desktop Bitmap
  
  Dim hdc As Long
  If m_DeskDC Is Nothing Then Set m_DeskDC = New MemoryDC

'Hide our parent Window so it doesn't get captured
  SetWindowPos m_hWndP, 0, 0, 0, 0, 0, _
    SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_HIDEWINDOW Or SWP_NOSENDCHANGING
  'Allow windows to clean up the screen
  DoEvents
  
  'Capture the bitmap (hopefully it's clear now)
  hdc = GetDC(0)
  m_DeskDC.CreateDC GetDeviceCaps(hdc, HORZRES), GetDeviceCaps(hdc, VERTRES)
  BitBlt m_DeskDC.hdc, 0, 0, m_DeskDC.Width, m_DeskDC.Height, hdc, 0, 0, vbSrcCopy
  ReleaseDC 0, hdc
  
  'Capture is done, display the main window back
  SetWindowPos m_hWndP, 0, 0, 0, 0, 0, _
    SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_SHOWWINDOW Or SWP_NOSENDCHANGING

End Sub

Private Sub PaintShadow(ByVal hwnd As Long)
  Dim tPS As PAINTSTRUCT, tR As RECT
  
  On Error Resume Next
  'Fill the paintstruct buffer
  BeginPaint hwnd, tPS
  'Get the size and position of our Shadowed window
  GetWindowRect hwnd, tR
  
  With tR
    'Copy from captured bitmap
    BitBlt m_DC.hdc, 0, 0, m_DC.Width, m_DC.Height, _
      m_DeskDC.hdc, .Left, .Top, vbSrcCopy
    'Draw the shadow
    m_Gradient.DrawDropShadow m_DC.hdc, 0, 0, m_DC.Width, m_DC.Height, m_Offset
    'Copy to Shadow window
    BitBlt tPS.hdc, 0, 0, m_DC.Width, m_DC.Height, _
      m_DC.hdc, 0, 0, vbSrcCopy
  End With
  'Validate and end the paint operation
  EndPaint hwnd, tPS
End Sub

Private Sub Class_Terminate()
  Dispose
End Sub

Private Sub IWinSub_After(lReturn As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
  SynchronizeShadow uMsg 'WM_SIZE and WM_MOVE
End Sub

Private Sub IWinSub_Before(bHandled As Long, lReturn As Long, hwnd As Long, uMsg As Long, wParam As Long, lParam As Long)
  
  Select Case uMsg
    Case WM_PAINT 'Paint message for our shadow
      PaintShadow hwnd
    
    Case WM_ERASEBKGND  'Prevent window from processing this message
      lReturn = 1
      bHandled = 1
      
    Case WM_ACTIVATE
      If wParam > 0 Then 'Activate, Capture background and show shadow
        
        CaptureDesktop
        
        SetWindowPos m_hWndShadow, m_hWndP, 0, 0, 0, 0, _
          SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
      
      Else 'Deactivate, Hide shadow
        ShowWindow m_hWndShadow, SW_HIDE
      End If
      
  End Select
End Sub

