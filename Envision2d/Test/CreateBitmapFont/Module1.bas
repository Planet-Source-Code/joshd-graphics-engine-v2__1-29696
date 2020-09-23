Attribute VB_Name = "Module1"
Public Enum RasterOps
    SRCCOPY = &HCC0020
    SRCAND = &H8800C6
    SRCINVERT = &H660046
    SRCPAINT = &HEE0086
    SRCERASE = &H4400328
    WHITENESS = &HFF0062
    BLACKNESS = &H42
End Enum

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As RasterOps) As Long

