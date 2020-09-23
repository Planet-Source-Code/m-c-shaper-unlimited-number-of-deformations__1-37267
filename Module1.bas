Attribute VB_Name = "Module1"
'funkciji ki kopirata in stegneta sliko
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'sledeèe konstante morajo biti tu da bitblt in stretchblt delata
Global Const SRCCOPY = &HCC0020 'Copies the source bitmap to the destination bitmap.
Global Const SRCAND = &H8800C6 'Combines pixels of the destination and source bitmap using the Boolean AND operator.
Global Const SRCPAINT = &HEE0086 'Combines pixels of the destination and source bitmap using the Boolean OR operator.
Global Const SRCERASE = &H440328  ' (DWORD) dest = source AND (NOT dest )
Global Const SRCINVERT = &H660046  ' (DWORD) dest = source XOR dest

'source picture mora imeti autoredraw = false èe hoèeš da bit in stretchblt
'prepoznata vgnezdeno image v sliki.


'deklacije set in get pixel
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

'deklaracije, ki naj bi pospešile kopiranje po pikslih v TEST command buttnu

Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long


'for filling purposes
'Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Public Const FLOODFILLBORDER = 0 ' Fill until crColor& color encountered.
Public Const FLOODFILLSURFACE = 1 ' Fill surface until crColor& color not encountered.


