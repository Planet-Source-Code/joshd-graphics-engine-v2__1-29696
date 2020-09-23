VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum RasterOps
    SRCCOPY = &HCC0020
    SRCAND = &H8800C6
    SRCINVERT = &H660046
    SRCPAINT = &HEE0086
    SRCERASE = &H4400328
    WHITENESS = &HFF0062
    BLACKNESS = &H42
End Enum

Public surfaceHDC As Long
Public x As Integer
Public y As Integer
Public height As Integer
Public width As Integer
Public Sub Draw(targetHDC As Long, targetX As Integer, targetY As Integer, paintMethod As RasterOps)
    BitBlt targetHDC, targetX, targetY, width, height, surfaceHDC, x, y, paintMethod
End Sub
'THIS CLASS IS INCOMPLETE
'It will be used for each sprite
'ie. If you have a tree in your game it will
'be a sprite - and its source and coordinates
'will be here. As will its mask.
Private Sub Class_Initialize()
    surfaceHDC = 0
    x = 0
    y = 0
    height = 0
    width = 0
End Sub

