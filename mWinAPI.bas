Attribute VB_Name = "mWinAPI"
Option Explicit

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Type RGBQUAD
  rgbBlue As Byte
  rgbGreen As Byte
  rgbRed As Byte
  rgbReserved As Byte
End Type

Public Type PAINTSTRUCT
  hdc As Long
  fErase As Long
  rcPaint As RECT
  fRestore As Long
  fIncUpdate As Long
  rgbReserved(31) As Byte
End Type

Public Const BITSPIXEL As Long = 12

Public Declare Function IsWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function BeginPaint Lib "user32.dll" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Public Declare Function EndPaint Lib "user32.dll" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'Public Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Const SW_HIDE As Long = 0
Public Const SW_SHOWNOACTIVATE As Long = 4
Public Const SW_SHOW As Long = 5

Public Const SWP_HIDEWINDOW As Long = &H80
Public Const SWP_NOACTIVATE As Long = &H10
Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOSIZE As Long = &H1
Public Const SWP_NOSENDCHANGING As Long = &H400
Public Const SWP_NOZORDER As Long = &H4
Public Const SWP_SHOWWINDOW As Long = &H40

Public Const WM_CLOSE As Long = &H10
Public Const WM_SIZE As Long = &H5
Public Const WM_MOVE As Long = &H3
Public Const WM_ACTIVATE As Long = &H6
Public Const WM_PAINT As Long = &HF
Public Const WM_ERASEBKGND As Long = &H14
'Public Const WM_NCACTIVATE As Long = &H86

Public Const WS_POPUP As Long = &H80000000
Public Const WS_BORDER As Long = &H800000
Public Const WS_DLGFRAME As Long = &H400000
Public Const WS_VISIBLE As Long = &H10000000
Public Const WS_EX_TOOLWINDOW As Long = &H80&
Public Const WS_CHILD As Long = &H40000000

Public Const HORZRES As Long = 8
Public Const VERTRES As Long = 10

Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long

Public Declare Function RedrawWindow Lib "user32.dll" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Const RDW_INVALIDATE As Long = &H1
Public Const RDW_UPDATENOW As Long = &H100


Public Const GWL_STYLE As Long = -16


