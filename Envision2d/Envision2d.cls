VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Envision2d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Sine(0 To 360) As Double
Private CoSine(0 To 360) As Double
Private Const Pi = 3.14159265358979
Public Sub PaintTrans(FromHDC As Long, ToHDC As Long, xSource As Integer, ySource As Integer, xTarget As Integer, yTarget As Integer, width As Integer, height As Integer, percent As Integer)
    'Make an image transparent over the top of another - slow for all but very small images
    On Error Resume Next
    Dim x As Integer, y As Integer
    Dim col1 As Long, col2 As Long, colfinal As Long
    Dim dec1 As Double, dec2 As Double
    Dim R1 As Long, G1 As Long, B1 As Long
    Dim R2 As Long, G2 As Long, B2 As Long
    
    'The decimal percentage of each pixel being combined
    dec1 = percent * 0.01
    dec2 = 1 - dec1
    
    For x = 0 To width
        For y = 0 To height
            col1 = GetPixel(FromHDC, x + xSource, y + ySource)
            col2 = GetPixel(ToHDC, x + xTarget, y + yTarget)
            
            'Get the red, green and blue colour values
            B1 = Int(col1 / 65536)
            G1 = Int((col1 - (65536 * B1)) / 256)
            R1 = col1 - (65536 * B1 + 256 * G1)
            
            B2 = Int(col2 / 65536)
            G2 = Int((col2 - (65536 * B2)) / 256)
            R2 = col2 - (65536 * B2 + 256 * G2)
            
            'Combine the colours before placing the pixel on the hdC
            colfinal = RGB(dec1 * R1 + dec2 * R2, dec1 * G1 + dec2 * G2, dec1 * B1 + dec2 * B2)
            SetPixel ToHDC, x + xTarget, y + yTarget, colfinal
        Next y
    Next x
End Sub
Public Sub PaintKey(FromHDC As Long, ToHDC As Long, xSource As Integer, ySource As Integer, xTarget As Integer, yTarget As Integer, width As Integer, height As Integer, key As Long)
    'Copy the image to the screen but areas of the same colout as the key should be transparent
    On Error Resume Next
    Dim x As Integer, y As Integer
    Dim colour As Long
    
    For x = 0 To width
        For y = 0 To height
            colour = GetPixel(FromHDC, x + xSource, y + ySource)
            If colour <> key Then SetPixel ToHDC, x + xTarget, y + yTarget, colour
        Next y
    Next x
End Sub
Public Sub PaintMask(ImageHDC As Long, MaskHDC As Long, xSource As Integer, ySource As Integer, xTarget As Integer, yTarget As Integer, width As Integer, height As Integer, TransKey As Long)
    'Crate a black/white mask
    On Error Resume Next
    Dim x As Integer, y As Integer
    Dim colour As Long
    
    For x = 0 To width
        For y = 0 To height
            If GetPixel(ImageHDC, x + xSource, y + ySource) = TransKey Then
                SetPixel MaskHDC, x + xTarget, y + yTarget, vbWhite
            Else
                SetPixel MaskHDC, x + xTarget, y + yTarget, vbBlack
            End If
        Next y
    Next x
End Sub
Public Sub PaintRotate(FromHDC As Long, ToHDC As Long, xSource As Integer, ySource As Integer, xTarget As Integer, yTarget As Integer, width As Integer, height As Integer, angle As Long)
    'Rotate a picture around its centre (fast, but recommend "pre-rotating" all but very small images)
    'This took a bit of fidling to get it to rotate around the centre
    'It is still not right as there are "gaps" in the rotated image
    Dim x As Integer, y As Integer
    Dim halfwidth As Double, halfheight As Double
    Dim newX As Double, newY As Double
    halfwidth = width / 2
    halfheight = height / 2
    
    For x = 0 To width
        For y = 0 To width
            newY = (y - halfheight) * CoSine(360 - angle) - (x - halfwidth) * Sine(360 - angle)
            newX = (x - halfwidth) * CoSine(360 - angle) + (y - halfheight) * Sine(360 - angle)
            SetPixel ToHDC, newX + halfwidth + xTarget, newY + halfheight + yTarget, GetPixel(FromHDC, x + xSource, y + ySource)
        Next y
    Next x
End Sub
Public Function Distance(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer) As Double
    'Distance between two points
    Distance = Sqr(Abs(x2 - x1) * Abs(x2 - x1) + Abs(y2 - y1) * Abs(y2 - y1))
End Function
Public Function Intersects(x1 As Integer, y1 As Integer, width1 As Integer, height1 As Integer, x2 As Integer, y2 As Integer, width2 As Integer, height2 As Integer) As Boolean
    'Do two rectangles intersect (suitable for most games)
    Dim xIntersect As Boolean, yIntersect As Boolean
    xIntersect = (x2 >= x1 And x2 <= x1 + width1) Or (x2 + width2 >= x1 And x2 + width2 <= x1 + width1)
    yIntersect = (y2 >= y1 And y2 <= y1 + height1) Or (y2 + height2 >= x1 And y2 + height2 <= y1 + height1)
    Intersects = xIntersect And yIntersect
End Function
Public Function IntersectPixel(x1 As Integer, y1 As Integer, width1 As Integer, height1 As Integer, x2 As Integer, y2 As Integer, width2 As Integer, height2 As Integer, colourKey As Integer) As Boolean
    'If all the pixels that touch are the same as colourKey it is not a hit
    'This is much slower than the intersect function
End Function
Public Function Within(x As Integer, y As Integer, rectX As Integer, rectY As Integer, rectWidth As Integer, rectHeight As Integer) As Boolean
    'Is a point within a rectangle (such as a button on a menu)
    If x >= rectX And x <= rectX + rectWidth And y >= rectY And y <= rectY + rectHeight Then
        Within = True
    Else
        Within = False
    End If
End Function
Private Sub Class_Initialize()
    'Create a table of sin and cos values (for rotation)
    Dim i As Integer
    For i = 0 To 360
        Sine(i) = Sin(i / 180 * Pi)
    Next i
    For i = 0 To 360
        CoSine(i) = Cos(i / 180 * Pi)
    Next i
End Sub
