Attribute VB_Name = "modRegions"
'Made for E¹
'29/09/2001
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

'Constants used by the CombineRgn() API function.
Public Const RGN_AND = 1&
Public Const RGN_OR = 2&
Public Const RGN_XOR = 3&
Public Const RGN_DIFF = 4&
Public Const RGN_COPY = 5&

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type POINTAPI
    x As Long
    y As Long
End Type


Public Function RRRegion(frm As Form, rad As Long)
Dim hRgn As Long
hRgn = CreateRoundRectRgn((frm.Width / 15) + 1, (frm.Height / 15) + 1, 0, 0, rad, rad)
SetWindowRgn frm.hwnd, hRgn, True
End Function
