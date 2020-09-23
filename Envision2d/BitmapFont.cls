VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitmapFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim CharX(0 To 255) As Integer
Dim CharY(0 To 255) As Integer
Dim CharHeight(0 To 255) As Integer
Dim CharWidth(0 To 255) As Integer
Public FontHDC As Long
Public MaskHDC As Long
Public Sub LoadFile(fileSpec As String)
    'Load a file of format: ASCII Value, X, Y, Width, Height
    On Error Resume Next
    Dim i As Integer
    Open fileSpec For Input As #1
        Do While Not EOF(1)
            Input #1, i
            Input #1, CharX(i), CharY(i), CharWidth(i), CharHeight(i)
        Loop
    Close #1
End Sub
Public Sub DrawText(targetHDC As Long, text As String, x As Integer, y As Integer)
    'Simply draw the mask (white bits will be transparent)
    'and the other bit(black bits will be transparent)
    DrawCustom MaskHDC, targetHDC, text, x, y, SRCAND
    DrawCustom FontHDC, targetHDC, text, x, y, SRCPAINT
End Sub
Public Sub DrawCustom(SourceHDC As Long, targetHDC As Long, text As String, x As Integer, y As Integer, paintMethod As Long)
    Dim i As Integer, xMargin As Integer
    Dim letter As String, code As Integer
    Dim letterX As Integer, letterY As Integer
    
    letterX = x
    letterY = y
    xMargin = x
    For i = 1 To Len(text)
        letter = Mid(text, i, 1)    'Get the letter
        code = Asc(letter)
        If code = 13 Then           'Next Line
            letterX = xMargin
            letterY = letterY + CharHeight(65)
        Else                        'Copy the coresponding piece to the source
            Call BitBlt(targetHDC, letterX, letterY, CharWidth(code), CharHeight(code), SourceHDC, CharX(code), CharY(code), paintMethod)
            letterX = letterX + CharWidth(code)
        End If
    Next i
End Sub
