VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FontSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public fontName As String
Public size As Integer
Public colour As Long
Public bold As Boolean
Public italic As Boolean
Public underline As Boolean
Public border As Boolean
Public borderCol As Long
Public background As Boolean
Public backgroundCol As Long
Private Sub Class_Initialize()
    'Default values:
    background = False
    border = False
    backgroundCol = vbBlue
    borderCol = vbWhite
    colour = vbBlack
    italic = False
    underline = False
    bold = False
    fontName = "Arial"
    size = 12
End Sub
Public Sub WriteText(surface As Object, x As Integer, y As Integer, text As String)
    'Print the text - this function allows for background colours/border colours
    Dim width As Integer, height As Integer
    surface.FontBold = bold
    surface.FontItalic = italic
    surface.FontUnderline = underline
    surface.FontSize = size
    surface.ForeColor = colour
    width = surface.TextWidth(text)
    height = surface.TextHeight(text)
    If background Then surface.Line (x, y)-(x + width, y + height), backgroundCol, BF
    If border Then surface.Line (x, y)-(x + width, y + height), borderCol, B
    surface.CurrentX = x
    surface.CurrentY = y
    surface.Print text
End Sub
