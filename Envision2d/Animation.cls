VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Animation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'THIS CLASS IS INCOMPLETE ;)
'It will contain arrays of sprites (sprites and their masks)
'and allow you to easily loop through them and produce animations
'with less code. It will prolly have its own file format too

Option Explicit
Dim sprites(1 To 100) As New Sprite
Dim mask(1 To 100) As New Sprite
Public masked As Boolean
Public frameCount As Integer
Public currentFrame As Integer
Public x As Integer
Public y As Integer
Public Sub LoadFile(fileName As String)
    ''''''''''''
End Sub
Public Sub AddSprites(newFrame As Sprite, Optional newMask As Sprite, Optional useMask As Boolean = False)
    'This procedure is not functioning
    frameCount = frameCount + 1
    With sprites(frameCount)
        .surfaceHDC = newFrame.surfaceHDC
        .y = newFrame.y
        .x = newFrame.x
        .width = newFrame.width
        .height = newFrame.height
    End With
    If useMask Then
        With mask(frameCount)
            .surfaceHDC = newMask.surfaceHDC
            .y = newMask.y
            .x = newMask.x
            .width = newMask.width
            .height = newMask.height
        End With
    End If
End Sub
Public Sub AddFrame(surfaceHDC As Long, x As Integer, y As Integer, width As Integer, height As Integer) 'Masks not supported
    frameCount = frameCount + 1
    With sprites(frameCount)
        .surfaceHDC = surfaceHDC
        .y = y
        .x = x
        .width = width
        .height = height
    End With
End Sub
Public Sub AddMask(surfaceHDC As Long, x As Integer, y As Integer, width As Integer, height As Integer) 'Masks not supported
    With mask(frameCount)
        .surfaceHDC = surfaceHDC
        .y = y
        .x = x
        .width = width
        .height = height
    End With
End Sub
Public Function NextFrame() As Sprite
    If frameCount = 0 Then NextFrame = Null
    currentFrame = currentFrame + 1
    If currentFrame > frameCount Then currentFrame = 1
    NextFrame = sprites(currentFrame)
End Function
Public Sub DrawNextFrame(targetHDC As Long)
    If frameCount = 0 Then Exit Sub
    currentFrame = currentFrame + 1
    If currentFrame > frameCount Then currentFrame = 1
    DrawFrame targetHDC, currentFrame
End Sub
Public Sub DrawFrame(targetHDC As Long, frameIndex As Integer)
    If frameCount = 0 Or frameIndex > frameCount Then Exit Sub
    If masked = True Then
        mask(frameIndex).Draw targetHDC, x, y, SRCAND
        sprites(frameIndex).Draw targetHDC, x, y, SRCPAINT
    Else
        sprites(frameIndex).Draw targetHDC, x, y, SRCCOPY
    End If
End Sub
Private Sub Class_Initialize()
    frameCount = 0
    currentFrame = 0
    masked = True
    x = 0
    y = 0
End Sub
