VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TileSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public gridRows As Integer
Public gridCols As Integer
Public cellWidth As Integer
Public cellHeight As Integer
Dim TileSprite(1 To 1000, 1 To 1000) As New sprite
Dim TileBehaviour(1 To 1000, 1 To 1000) As Integer
Dim TileToken(1 To 1000, 1 To 1000) As Integer
Public Function GridToPointX(GridCol As Integer) As Double
    GridToPointX = GridCol * cellWidth
End Function
Public Function GridToPointY(GridRow As Integer) As Double
    GridToPointY = GridRow * cellHeight
End Function
Public Function PointToGridX(GridX As Integer) As Integer
    PointToGridX = GridX \ cellWidth
End Function
Public Function PointToGridY(GridY As Integer) As Integer
    PointToGridY = GridY \ cellHeight
End Function
Public Sub SetDimensions(Cols As Integer, Rows As Integer)
    If Cols > 1000 Then Cols = 1000
    If Rows > 1000 Then Rows = 1000
    
    Dim x As Integer, y As Integer
    For x = 1 To Cols
        For y = 1 To Rows
            'Do not overwrite exsting places:
            If y > gridRows And x > gridCols Then
                TileToken(x, y) = 0
                TileBehaviour(x, y) = 0
            End If
        Next y
    Next x
    gridRows = Rows
    gridCols = Cols
End Sub
Public Sub SetSprite(newSprite As sprite, col As Integer, row As Integer)
    With TileSprite(col, row)
        .height = newSprite.height
        .width = newSprite.width
        .x = newSprite.x
        .y = newSprite.y
        .surfaceHDC = newSprite.surfaceHDC
    End With
End Sub
Public Function GetSprite(col As Integer, row As Integer) As sprite
    If row > gridRows Or col > gridCols Then Exit Function
    GetSprite = TileSprite(col, row)
End Function
Public Sub DrawAll(targetHDC As Long)
    Dim x As Integer, y As Integer
    For x = 1 To gridCols
        For y = 1 To gridRows
            TileSprite(x, y).Draw targetHDC, (x - 1) * cellWidth, (y - 1) * cellHeight, SRCCOPY
        Next y
    Next x
End Sub
Private Sub Class_Initialize()
    gridRows = 0
    gridCols = 0
End Sub
