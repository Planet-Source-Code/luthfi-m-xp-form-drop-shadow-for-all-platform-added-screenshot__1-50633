VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Gradient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const AC_SRC_OVER = &H0
Private Const AC_SRC_ALPHA = &H1

Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type

Private Declare Function AlphaBlend Lib "msimg32.dll" ( _
  ByVal hdcDest As Long, _
  ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, _
  ByVal nWidthDest As Long, ByVal nHeightDest As Long, _
  ByVal hdcSrc As Long, _
  ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, _
  ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, _
  ByVal BLENDFUNCT As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    
Private Type SAFEARRAYBOUND
  cElements As Long
  lLbound As Long
End Type
          
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type


Private Declare Function VarPtrArray Lib "msvbvm60.dll" _
        Alias "VarPtr" (Ptr() As Any) As Long


Sub DrawDropShadow(ByVal hdc As Long, _
  ByVal lLeft As Long, ByVal lTop As Long, _
  ByVal lRight As Long, ByVal lBottom As Long, _
  ByVal lOffset As Long, _
  Optional ByVal Boxed As Boolean = True, _
  Optional ByVal Horizontal As Boolean = False, _
  Optional ByVal Alpha As Byte = 128)
  
  Dim hBr As Long, tR As RECT
  Dim cDCBack As MemoryDC
  Dim tBlend As BLENDFUNCTION, lBF As Long
  
  Set cDCBack = New MemoryDC
  With cDCBack
    .CreateCompatibleDIB lRight - lLeft, lBottom - lTop
  
    hBr = CreateSolidBrush(vbBlack)
    SetRect tR, 0, 0, .Width, .Height
    FillRect .hdc, tR, hBr
    DeleteObject hBr
  
    CreateDropShadow Boxed, Horizontal, lOffset, cDCBack
    
    With tBlend
      .BlendOp = AC_SRC_OVER
      .BlendFlags = 0
      .AlphaFormat = AC_SRC_ALPHA
      .SourceConstantAlpha = Alpha
    End With
    
    CopyMemory lBF, tBlend, 4
    
    AlphaBlend hdc, lLeft, lTop, .Width, .Height, .hdc, 0, 0, .Width, .Height, lBF
  End With
  Set cDCBack = Nothing
End Sub

Private Sub CreateDropShadow( _
  ByVal Boxed As Boolean, _
  ByVal bHorizontal As Boolean, ByVal lOffset As Long, _
  CDC As MemoryDC)
'Note Dib is reverse upsidedown
  Dim bDib() As Byte
  Dim x As Long, y As Long
  Dim lC As Long, lInitC As Long, lSize As Long
  Dim tSA As SAFEARRAY2D
  Dim lWidth As Long
  Dim lStartStep As Long, lEndStep As Long
    
  On Error Resume Next
  'This is a modified routine from www.vbAccelerator.com
  lWidth = CDC.Width * 4

  ' Get the bits in the from DIB section:
  With tSA
      .cbElements = 1
      .cDims = 2
      .Bounds(0).lLbound = 0
      .Bounds(0).cElements = CDC.Height
      .Bounds(1).lLbound = 0
      .Bounds(1).cElements = lWidth
      .pvData = CDC.DibPtr
  '        Debug.Print .pvData
  End With
  CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4
  
'  GoTo VER
  If (bHorizontal = True) Or (Boxed = True) Then
    ' horizontal:
    
    If Boxed Then
      lStartStep = lOffset
      lSize = lOffset
    Else
      lStartStep = CDC.Height - 1&
      lSize = CDC.Height
    End If
    
    
    ' at bottom:
    For x = 0 To lWidth - 1& Step 4&
      ' if x is within lSize of the start or end, then
      ' it contributes to alpha:
      
      If (x < lSize * 4) Then
         lInitC = (255 * x) \ (lSize * 4&)
      ElseIf (x >= (lWidth - lSize * 4&)) Then
         lInitC = (((lWidth - x) * 255&) \ (4& * lSize))
      Else
         lInitC = 255
      End If
       
      For y = 0 To lStartStep 'CDC.Height - 1&
         lC = (lInitC * y) \ lSize 'CDC.Height
         bDib(x + 3&, y) = lC
         bDib(x + 2&, y) = 0
         bDib(x + 1&, y) = 0
         bDib(x, y) = 0
      Next y
       
    Next x
  End If

  If (bHorizontal = False) Or (Boxed = True) Then
    ' vertical:
    If Boxed Then
      lStartStep = lWidth - (lOffset * 4&)
      lEndStep = lOffset
    Else
      lStartStep = 0&
      lEndStep = 0
    End If
    
    lSize = (lWidth - lStartStep) \ 4&
    ' at right:
    For y = lEndStep To CDC.Height - 1&
      ' if y is within lSize of the start or end, then
      ' it contributes to the alpha:
      
      'If (y < lSize) Then
      '   lInitC = (255 * y) \ lSize
      If (y >= (CDC.Height - lSize)) Then
         lInitC = (255& * (CDC.Height - y)) \ lSize
      Else
         lInitC = 255&
      End If
      
      For x = lStartStep To lWidth - 1& Step 4&
         ' the amount of alpha depends on how far we are from the left:
         lC = (lInitC * (lWidth - x)) \ (lWidth - lStartStep)
         'Debug.Print lC
         bDib(x + 3&, y) = lC
         bDib(x + 2&, y) = 0
         bDib(x + 1&, y) = 0
         bDib(x, y) = 0
      Next x
    Next y
  End If
  
  ' Clear the temporary array descriptor
  ' (This does not appear to be necessary, but
  ' for safety do it anyway)
  CopyMemory ByVal VarPtrArray(bDib), 0&, 4&
    
End Sub

