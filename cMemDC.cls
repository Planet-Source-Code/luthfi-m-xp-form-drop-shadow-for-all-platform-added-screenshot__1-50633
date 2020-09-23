VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemoryDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type BITMAPINFOHEADER
  biSize As Long
  biWidth As Long
  biHeight As Long
  biPlanes As Integer
  biBitCount As Integer
  biCompression As Long
  biSizeImage As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type

Private Type BITMAPINFO256
  bmiHeader As BITMAPINFOHEADER
  bmiColors(0 To 255) As RGBQUAD
End Type

Private Const BI_RGB As Long = 0&
Private Const DIB_RGB_COLORS As Long = 0

Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function CreateDIBSection256 Lib "gdi32.dll" Alias "CreateDIBSection" _
  (ByVal hdc As Long, pBitmapInfo As BITMAPINFO256, ByVal un As Long, _
  lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" _
    (ByVal hdc As Long, _
    pBitmapInfo As BITMAPINFO256, _
    ByVal un As Long, _
    lplpVoid As Long, _
    ByVal handle As Long, _
    ByVal dw As Long) As Long
    
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFOHEADER, ByVal wUsage As Long) As Long

Private m_hDC As Long
Private m_hBmpOld As Long
Private m_hBmp As Long
Private m_lWidth As Long
Private m_lheight As Long
Private m_Created As Boolean

Private m_tBI256 As BITMAPINFO256
Private m_Is256 As Boolean
Private m_lPtr As Long

Private m_tRGB(0 To 255) As RGBQUAD

Function CreateCompatibleDIB(ByVal lWidth As Long, ByVal lHeight As Long) As Long
  pDestroy
  Dim lBitDepth As Long, lhDC As Long
  lhDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
  m_hDC = CreateCompatibleDC(lhDC)
  DeleteDC lhDC
  
  lBitDepth = GetDeviceCaps(m_hDC, BITSPIXEL)
  If lBitDepth = 8 Then
    CreateCompatibleDIB = Create256DIB(lWidth, lHeight)
    Exit Function
  End If
  GetOptimalDIBFormat m_hDC, m_tBI256.bmiHeader
  With m_tBI256.bmiHeader
    .biBitCount = 32
    .biWidth = lWidth
    .biHeight = lHeight
  End With
  m_hBmp = CreateDIBSection(m_hDC, m_tBI256, DIB_RGB_COLORS, m_lPtr, 0&, 0&)
  
  CreateCompatibleDIB = m_hBmp
  If m_hBmpOld <> 0& Then
    DeleteObject SelectObject(m_hDC, m_hBmpOld)
  End If
  m_hBmpOld = SelectObject(m_hDC, m_hBmp)
  m_Is256 = False
  m_lWidth = lWidth
  m_lheight = lHeight
  m_Created = True
End Function

Public Sub CreateDC(ByVal Width As Long, ByVal Height As Long)
  pDestroy
  pCreate Width, Height
  m_Created = True
End Sub


Public Property Get hdc() As Long
   hdc = m_hDC
End Property

Property Get BytesPerScanLine() As Long
  ' Scans must align on dword boundaries:
  BytesPerScanLine = (m_tBI256.bmiHeader.biWidth + 3&) And &HFFFFFFFC
End Property

Property Get DibPtr() As Long
  If m_Created = False Then Exit Property
  DibPtr = m_lPtr
End Property

Property Get hBitmap() As Long
  If m_Created = False Then Exit Property
  hBitmap = m_hBmp
End Property

Public Property Get Width() As Long
   Width = m_lWidth
End Property

Public Property Get Height() As Long
   Height = m_lheight
End Property

Private Sub pCreate(ByVal lW As Long, ByVal lH As Long)
Dim lhDC As Long
   pDestroy
   lhDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
   m_hDC = CreateCompatibleDC(lhDC)
   m_hBmp = CreateCompatibleBitmap(lhDC, lW, lH)
   m_hBmpOld = SelectObject(m_hDC, m_hBmp)
   If m_hBmpOld = 0 Then
      pDestroy
   Else
      m_lWidth = lW
      m_lheight = lH
   End If
   DeleteDC ByVal lhDC
   m_Created = True
End Sub

Private Sub pDestroy()
    
  If m_hBmpOld <> 0 Then
     SelectObject ByVal m_hDC, ByVal m_hBmpOld
     m_hBmpOld = 0
  End If
    
  If m_hBmp <> 0 Then
     DeleteObject ByVal m_hBmp
     m_hBmp = 0
  End If
  
  m_lWidth = 0
  m_lheight = 0
  If Not m_hDC = 0 Then
    DeleteDC ByVal m_hDC
    m_hDC = 0
  End If
  m_Created = False
End Sub

Private Sub Class_Terminate()
   pDestroy
End Sub


Private Function GetOptimalDIBFormat(ByVal hdc As Long, pbi As BITMAPINFOHEADER) As Boolean
  Dim hBm As Long
  Dim bRet As Boolean
  
  ' Create a memory bitmap that is compatible with the
  ' format of the target device.
  hBm = CreateCompatibleBitmap(hdc, 1, 1)
  If hBm Then
  ' Initialize the header.
    pbi.biSize = Len(pbi)
    
    ' First call to GetDIBits will fill in the optimal biBitCount.
    bRet = GetDIBits(hdc, hBm, 0, 1, ByVal 0&, pbi, DIB_RGB_COLORS)
    
    ' Second call to GetDIBits will get the optimal color table, o
    ' or the optimal bitfields values.
    If bRet Then bRet = GetDIBits(hdc, hBm, 0, 1, ByVal 0&, pbi, DIB_RGB_COLORS)
    
    ' Clean up.
    DeleteObject hBm
    
    GetOptimalDIBFormat = bRet
  End If
End Function

Function Create256DIB(lWidth As Long, lHeight As Long) As Long
  If m_Created = False Then Exit Function
  With m_tBI256.bmiHeader
    .biSize = LenB(m_tBI256.bmiHeader)
    .biBitCount = 8
    .biCompression = BI_RGB
    .biWidth = lWidth
    .biHeight = lHeight
    .biPlanes = 1
    .biSizeImage = .biHeight * BytesPerScanLine
  End With
  'Fill pallete
  CopyMemory VarPtr(m_tBI256.bmiColors(0)), VarPtr(m_tRGB(0)), Len(m_tRGB(0)) * 256
  m_hBmp = CreateDIBSection256(m_hDC, m_tBI256, DIB_RGB_COLORS, m_lPtr, 0&, 0&)
  Create256DIB = m_hBmp
  If m_hBmpOld <> 0& Then
    DeleteObject SelectObject(m_hDC, m_hBmpOld)
  End If
  m_hBmpOld = SelectObject(m_hDC, m_hBmp)
  m_Is256 = True
  m_lWidth = lWidth
  m_lheight = lHeight
  m_Created = True
End Function

