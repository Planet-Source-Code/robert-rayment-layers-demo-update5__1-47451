Attribute VB_Name = "APIMod"
' APIMOD.bas

Option Explicit

' --------------------------------------------------------------
' Used here to detect CTRL key in PicAColor
Public Declare Function GetAsyncKeyState Lib "user32" _
(ByVal vKey As Long) As Integer

Public Const VK_CONTROL = &H11
Public Const VK_SHIFT = &H10

' --------------------------------------------------------------
' Timing - for Do Loops

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
' -----------------------------------------------------------
' Windows APIs -  Function & constants to locate & make Window stay on top
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
ByVal wi As Long, ByVal ht As Long, ByVal wFlags As Long) As Long

Public Const SWP_NOSIZE = &H1
'Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOPMOST = -1
'Public Const HWND_NOTOPMOST = -2

Public Const wFlags = SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
Public Const wflags2 = SWP_SHOWWINDOW Or SWP_NOACTIVATE

'--------------------------------------------------------------
' Windows APIs - For searching list box
'Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

' NB  lParam needs to be As Long for some functions
' but As Any for Search List Box using LB_FINDSTRINGEXACT
'--------------------------------------------------
'--------------------------------------------------------------------------
' Shaping APIs

Public Declare Function CreateRoundRectRgn Lib "gdi32" _
(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal x3 As Long, ByVal y3 As Long) As Long

Public Declare Function SetWindowRgn Lib "user32" _
(ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Public Declare Function DeleteObject Lib "gdi32" _
(ByVal hObject As Long) As Long

' -----------------------------------------------------------

' For getting screen cursor position

Public Declare Function GetPixel Lib "gdi32" _
(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long


Public Type POINTAPI
  X As Long
  Y As Long
End Type
               
'Public Declare Sub SetCursorPos Lib "USER32" (ByVal IX As Long, ByVal IY As Long)
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetDC& Lib "user32" (ByVal hwnd As Long)
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

' -----------------------------------------------------------
Public Type TextFont
   fntName As String
   fntSize As Long
   fntItalic As Boolean
   fntBold As Boolean
End Type
Public TextFont As TextFont

' -----------------------------------------------------------
' API to Fill background, creating masks

Public Declare Function CreatePen Lib "gdi32" _
(ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, _
 ByVal Y As Long, ByVal crColor As Long, ByVal fuFillType As Long) As Long
Public Const FLOODFILLSURFACE = 1

'------------------------------------------------------------------------------
'Declare Function GetStretchBltMode Lib "gdi32" _
'(ByVal hdc As Long) As Long

Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long

'Public Const COLORONCOLOR = 3
Public Const HALFTONE = 4

' Use:-
'  ' oldMode = GetStretchBltMode( pic.hdc)
'
'  If SetStretchBltMode(pic.hdc, HALFTONE) = 0 Then
'     MsgBox "SetStretchBltMode error ", vbCritical, " "
'     End
'  End If
'
'  SetStretchBltMode pic.hdc, OldMode

'------------------------------------------------------------------------------
'Public Declare Function InvertRect Lib "USER32" _
'(ByVal hdc As Long, lpRect As RECT) As Long

'------------------------------------------------------------------------------
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" _
   (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function SetCursor Lib "user32" _
   (ByVal hCursor As Long) As Long

'------------------------------------------------------------------------------
'  This required instead of Screen.Height & Width for resizing
Public Declare Function GetSystemMetrics Lib "user32" _
(ByVal nIndex As Long) As Long

Public Const SM_CXSCREEN = 0  ' Screen Width
Public Const SM_CYSCREEN = 1  ' Screen Height
Public Const SM_CYCAPTION = 4 ' Height of window caption
Public Const SM_CYMENU = 15   ' Height of menu
Public Const SM_CXDLGFRAME = 7   ' Width of borders X & Y same + 1 for sizable
Public Const SM_CYSMCAPTION = 51 ' Height of small caption (Tool Windows)

'------------------------------------------------------------
' API used to check selection & clipping rectangles

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
'Public Declare Function SetRect Lib "USER32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Public Declare Function IntersectRect Lib "USER32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
'Public Declare Function IsRectEmpty Lib "USER32" (lpRect As RECT) As Long

'------------------------------------------------------------------------------
' Structures for DIBs
Public Type BITMAPINFOHEADER ' 40 bytes
   biSize As Long
   biwidth As Long
   biheight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

'Public Type RGBQUAD
'        rgbBlue As Byte
'        rgbGreen As Byte
'        rgbRed As Byte
'        rgbReserved As Byte
'End Type

Public Type BITMAPINFO
   bmiH As BITMAPINFOHEADER
'   Colors(0 To 255) As RGBQUAD
End Type
Public bmbac As BITMAPINFO
Public bmpic As BITMAPINFO

' For transferring aDrawing in an array to Form or PicBox
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, _
ByVal X As Long, ByVal Y As Long, _
ByVal DesW As Long, ByVal DesH As Long, _
ByVal SrcXOffset As Long, ByVal SrcYOffset As Long, _
ByVal PICWW As Long, ByVal PICHH As Long, _
lpBits As Any, lpBitsInfo As BITMAPINFO, _
ByVal wUsage As Long, ByVal dwRop As Long) As Long
'wUsage is one of:-
'Public Const DIB_PAL_COLORS = 1 '  uses system....
'Public Const DIB_RGB_COLORS = 0 '  uses RGBQUAD
'dwRop is vbSrcCopy


' -----------------------------------------------------------
' APIs for getting DIB bits to PicMem

Public Declare Function GetDIBits Lib "gdi32" _
(ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, _
ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Public Declare Function CreateCompatibleDC Lib "gdi32" _
(ByVal hdc As Long) As Long

Public Declare Function SelectObject Lib "gdi32" _
(ByVal hdc As Long, ByVal hObject As Long) As Long

Public Declare Function DeleteDC Lib "gdi32" _
(ByVal hdc As Long) As Long

'----------------------------------------------------------------
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, _
ByVal X As Long, ByVal Y As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, _
ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, _
ByVal dwRop As Long) As Long

'----------------------------------------------------------------
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
ByVal X As Long, ByVal Y As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'----------------------------------------------------------------
'To fill BITMAP structure
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" _
(ByVal hObject As Long, ByVal Lenbmp As Long, dimbmp As Any) As Long

Public Type BITMAP
   bmType As Long              ' Type of bitmap
   bmWidth As Long             ' Pixel width
   bmHeight As Long            ' Pixel height
   bmWidthBytes As Long        ' Byte width = 3 x Pixel width
   bmPlanes As Integer         ' Color depth of bitmap
   bmBitsPixel As Integer      ' Bits per pixel, must be 16 or 24
   bmBits As Long              ' This is the pointer to the bitmap data  !!!
End Type
Public bmp As BITMAP

'------------------------------------------------------------------------------
'Copy one array to another of same number of bytes

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
(Destination As Any, Source As Any, ByVal Length As Long)

'------------------------------------------------------------------------------

Public Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, _
lpPoint As POINTAPI, ByVal nCount As Long) As Long

'------------------------------------------------------------------------------


