VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   533
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   361
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTiles 
      Caption         =   "Draw Tiles"
      Height          =   375
      Left            =   1800
      TabIndex        =   24
      Top             =   6240
      Width           =   3495
   End
   Begin VB.CommandButton cmdSingle 
      Caption         =   "Draw First"
      Height          =   375
      Left            =   4080
      TabIndex        =   23
      Top             =   5760
      Width           =   1215
   End
   Begin VB.PictureBox picSmileMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   480
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   22.5
      ScaleMode       =   2  'Point
      ScaleWidth      =   22.5
      TabIndex        =   22
      Top             =   5640
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picSmile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   0
      Picture         =   "Form1.frx":0B0A
      ScaleHeight     =   22.5
      ScaleMode       =   2  'Point
      ScaleWidth      =   22.5
      TabIndex        =   21
      Top             =   5640
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CommandButton cmdZOrder 
      Caption         =   "ZOrder (add 10, max 200)"
      Height          =   375
      Left            =   1800
      TabIndex        =   20
      Top             =   5760
      Width           =   2175
   End
   Begin VB.CommandButton cmdAnimStop 
      Caption         =   "Stop Animation"
      Height          =   375
      Left            =   3600
      TabIndex        =   19
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Timer tmrAnim 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   960
      Top             =   5520
   End
   Begin VB.CommandButton cmdSprite 
      Caption         =   "Draw Sprite"
      Height          =   375
      Left            =   1800
      TabIndex        =   18
      Top             =   2040
      Width           =   3495
   End
   Begin VB.CommandButton cmdAnimStart 
      Caption         =   "Start Animation"
      Height          =   375
      Left            =   1800
      TabIndex        =   17
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton cmdSShot 
      Caption         =   "ScreenShot (1)"
      Height          =   375
      Left            =   3600
      TabIndex        =   16
      Top             =   3480
      Width           =   1665
   End
   Begin VB.CommandButton cmdBitmapWrite 
      Caption         =   "Write"
      Height          =   375
      Left            =   3600
      TabIndex        =   15
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox txtBitmapWrite 
      Height          =   855
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   14
      Text            =   "Form1.frx":1614
      Top             =   3000
      Width           =   1695
   End
   Begin VB.PictureBox picBitmapWrite 
      AutoRedraw      =   -1  'True
      Height          =   1215
      Left            =   1800
      Picture         =   "Form1.frx":1635
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   229
      TabIndex        =   13
      Top             =   3960
      Width           =   3495
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1440
      Left            =   4560
      Picture         =   "Form1.frx":6574
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   583
      TabIndex        =   12
      Top             =   6240
      Visible         =   0   'False
      Width           =   8745
   End
   Begin VB.PictureBox picFont 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1410
      Left            =   4560
      Picture         =   "Form1.frx":2F6B6
      ScaleHeight     =   94
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   590
      TabIndex        =   11
      Top             =   6120
      Visible         =   0   'False
      Width           =   8850
   End
   Begin VB.CommandButton cmdMask 
      Caption         =   "CreateMask (black)"
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   1560
      Width           =   3495
   End
   Begin VB.CommandButton cmdWrite 
      Caption         =   "Write"
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox txtWrite 
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Text            =   "Score: 100"
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdPaintKey 
      Caption         =   "Key (black)"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   1080
      Width           =   3495
   End
   Begin VB.HScrollBar scrBlend 
      Height          =   255
      Left            =   2400
      Max             =   100
      TabIndex        =   5
      Top             =   600
      Width           =   2895
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   1455
   End
   Begin VB.HScrollBar scrRotate 
      Height          =   255
      Left            =   2400
      Max             =   360
      TabIndex        =   2
      Top             =   120
      Width           =   2895
   End
   Begin VB.PictureBox pic2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   120
      Picture         =   "Form1.frx":30968
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   1
      Top             =   1680
      Width           =   1500
   End
   Begin VB.PictureBox pic1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   120
      Picture         =   "Form1.frx":358A7
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label2 
      Caption         =   "Blend"
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Rotate"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Engine As New Envision2d
Dim SilverFont As New BitmapFont
Dim SSaver As New Screenshot
Dim MyAnim As New Animation
Dim MySprite As New Sprite
Dim MyElements As New ZOrder
Dim SmileSprite As New Sprite
Dim SmileSpriteMask As New Sprite
Dim MyTileSet As New TileSet
Private Sub cmdAnimStart_Click()
    tmrAnim.Enabled = True
End Sub

Private Sub cmdAnimStop_Click()
    tmrAnim.Enabled = False
End Sub

Private Sub cmdBitmapWrite_Click()
    picBitmapWrite.Cls
    SilverFont.DrawText picBitmapWrite.hDC, txtBitmapWrite.Text, 5, 5
    picBitmapWrite.Refresh
End Sub

Private Sub cmdMask_Click()
    pic2.Cls
    Engine.PaintMask pic2.hDC, pic2.hDC, 0, 0, 0, 0, 100, 100, vbBlack
    pic2.Refresh
End Sub
Private Sub cmdPaintKey_Click()
    pic2.Cls
    Engine.PaintKey pic1.hDC, pic2.hDC, 0, 0, 0, 0, 100, 100, vbBlack
    pic2.Refresh
End Sub
Private Sub cmdReset_Click()
    pic2.Cls
End Sub

Private Sub cmdSingle_Click()
    pic2.Cls
    MyElements.DrawID 1, pic2.hDC
    pic2.Refresh
End Sub

Private Sub cmdSprite_Click()
    pic2.Cls
    MySprite.Draw pic2.hDC, 10, 10, SRCCOPY
    pic2.Refresh
End Sub

Private Sub cmdSShot_Click()
    SSaver.Save picBitmapWrite
    cmdSShot.Caption = "ScreenShot (" & SSaver.screenCount & ")"
End Sub

Private Sub cmdTiles_Click()
    pic2.Cls
    MyTileSet.DrawAll pic2.hDC
    pic2.Refresh
End Sub

Private Sub cmdWrite_Click()
    Dim scoreFont As New FontSpec
    With scoreFont
        .backgroundCol = RGB(140, 0, 0)
        .background = True
        .borderCol = vbWhite
        .border = True
        .Colour = vbWhite
        .Size = 8
        .Bold = True
        .WriteText pic2, 5, 5, txtWrite.Text
    End With
    pic2.Refresh
End Sub

Private Sub cmdZOrder_Click()
    Dim i As Integer, y As Integer
    
    'Add 10 new smilies in a random position
    'We will let the zIndex be the same as y
    'this will usually be the case but not always
    For i = 1 To 10
        Randomize
        y = Rnd * 70
        Call MyElements.Add(SmileSprite, 70 * Rnd, y, y, True, SmileSpriteMask)
    Next i
    
    'Sort them all and then draw them all
    MyElements.Sort
    pic2.Cls
    MyElements.DrawAll pic2.hDC
    pic2.Refresh
    
End Sub

Private Sub Form_Load()
    Dim i As Integer, j As Integer
    'Bitmap Font
    SilverFont.FontHDC = picFont.hDC
    SilverFont.MaskHDC = picMask.hDC
    SilverFont.LoadFile App.Path & "\font.env"
    'Screen Capturer
    SSaver.Path = App.Path & "\"
    SSaver.FileName = "screen"
    'Sprite
    With MySprite
        .Height = 50
        .Width = 50
        .x = 30
        .y = 30
        .surfaceHDC = picFont.hDC
    End With
    'Create Animation
    MyAnim.masked = False
    MyAnim.x = 0
    MyAnim.y = 0
    For i = 1 To 20
        Randomize
        MyAnim.AddFrame picBitmapWrite.hDC, 50 + Rnd * 10, Rnd * 10, 40, 40
    Next i
    'Elements
    With SmileSprite
        .Height = 30
        .Width = 30
        .x = 0
        .y = 0
        .surfaceHDC = picSmile.hDC
    End With
    With SmileSpriteMask
        .Height = 30
        .Width = 30
        .x = 0
        .y = 0
        .surfaceHDC = picSmileMask.hDC
    End With
    'Tiles
    MyTileSet.SetDimensions 10, 10
    MyTileSet.cellWidth = 30
    MyTileSet.cellHeight = 30
    For i = 1 To 10
        For j = 1 To 10
            If Rnd * 2 < 1 Then
                MyTileSet.SetSprite SmileSprite, i, j
            Else
                MyTileSet.SetSprite SmileSpriteMask, i, j
            End If
        Next j
    Next i
End Sub
Private Sub scrBlend_Change()
    pic2.Cls
    Engine.PaintTrans pic1.hDC, pic2.hDC, 0, 0, 0, 0, 100, 100, scrBlend.Value
    pic2.Refresh
End Sub
Private Sub scrRotate_Change()
    pic2.Cls
    Engine.PaintRotate pic1.hDC, pic2.hDC, 20, 20, 40, 40, 50, 50, scrRotate.Value
    pic2.Refresh
End Sub
Private Sub tmrAnim_Timer()
    MyAnim.DrawNextFrame pic2.hDC
    pic2.Refresh
End Sub
