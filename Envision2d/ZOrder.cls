VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'THIS CLASS IS INCOMPLETE ;)
'It will store an array of ID's or sprites
'It will include a FAST sorting (not bubble sort :)
'algorithm so that objects can appear on screen in order
'eg. a tree that is behind another will be drawn behind it
'but a player character that moves around will appear
'either in front or behind it - depending on where he is
Option Explicit

Dim Elements(1 To 200) As New Element
Public elementCount As Integer
Public Sub Sort()
    If elementCount <= 1 Then Exit Sub  'If there is only 1 it is already sorted
    Dim i As Integer, j As Integer
    Dim smallest As Integer, smallID As Integer
    For i = 1 To elementCount - 1
        smallest = Elements(i).zIndex
        smallID = i
        For j = i + 1 To elementCount
            If Elements(j).zIndex < smallest Then
                smallest = Elements(j).zIndex
                smallID = j
            End If
        Next j
        Swap i, smallID
    Next i
End Sub
Public Sub Swap(index1 As Integer, index2 As Integer)
    Dim tempElement As New Element
    With tempElement
        .ID = Elements(index1).ID
        .x = Elements(index1).x
        .y = Elements(index1).y
        .zIndex = Elements(index1).zIndex
        .mask = Elements(index1).mask
        .sprite.height = Elements(index1).sprite.height
        .sprite.width = Elements(index1).sprite.width
        .sprite.x = Elements(index1).sprite.x
        .sprite.y = Elements(index1).sprite.y
        .sprite.surfaceHDC = Elements(index1).sprite.surfaceHDC
        .maskSprite.height = Elements(index1).maskSprite.height
        .maskSprite.width = Elements(index1).maskSprite.width
        .maskSprite.x = Elements(index1).maskSprite.x
        .maskSprite.y = Elements(index1).maskSprite.y
        .maskSprite.surfaceHDC = Elements(index1).maskSprite.surfaceHDC
    End With
    With Elements(index1)
        .ID = Elements(index2).ID
        .x = Elements(index2).x
        .y = Elements(index2).y
        .zIndex = Elements(index2).zIndex
        .mask = Elements(index2).mask
        .sprite.height = Elements(index2).sprite.height
        .sprite.width = Elements(index2).sprite.width
        .sprite.x = Elements(index2).sprite.x
        .sprite.y = Elements(index2).sprite.y
        .sprite.surfaceHDC = Elements(index2).sprite.surfaceHDC
        .maskSprite.height = Elements(index2).maskSprite.height
        .maskSprite.width = Elements(index2).maskSprite.width
        .maskSprite.x = Elements(index2).maskSprite.x
        .maskSprite.y = Elements(index2).maskSprite.y
        .maskSprite.surfaceHDC = Elements(index2).maskSprite.surfaceHDC
    End With
    With Elements(index2)
        .ID = tempElement.ID
        .x = tempElement.x
        .y = tempElement.y
        .zIndex = tempElement.zIndex
        .mask = tempElement.mask
        .sprite.height = tempElement.sprite.height
        .sprite.width = tempElement.sprite.width
        .sprite.x = tempElement.sprite.x
        .sprite.y = tempElement.sprite.y
        .sprite.surfaceHDC = tempElement.sprite.surfaceHDC
        .maskSprite.height = tempElement.maskSprite.height
        .maskSprite.width = tempElement.maskSprite.width
        .maskSprite.x = tempElement.maskSprite.x
        .maskSprite.y = tempElement.maskSprite.y
        .maskSprite.surfaceHDC = tempElement.maskSprite.surfaceHDC
    End With
End Sub
Public Function Add(newSprite As sprite, newX As Integer, newY As Integer, newZIndex As Integer, Optional newMask As Boolean = False, Optional newMaskSprite As sprite) As Integer
    If elementCount >= 200 Then Exit Function
    elementCount = elementCount + 1
    With Elements(elementCount)
        .ID = elementCount
        .x = newX
        .y = newY
        .zIndex = newZIndex
        .sprite.height = newSprite.height
        .sprite.width = newSprite.width
        .sprite.x = newSprite.x
        .sprite.y = newSprite.y
        .sprite.surfaceHDC = newSprite.surfaceHDC
        .mask = newMask
    End With
    If newMask Then
        With Elements(elementCount).maskSprite
            .height = newMaskSprite.height
            .width = newMaskSprite.width
            .x = newMaskSprite.x
            .y = newMaskSprite.y
            .surfaceHDC = newMaskSprite.surfaceHDC
        End With
    End If
    Add = elementCount
End Function
Public Sub SetElement(index As Integer, newElement As Element)
    Elements(index) = newElement
End Sub
Public Function GetElement(index As Integer) As Element
    GetElement = Elements(index)
End Function
Public Sub Draw(index As Integer, targetHDC As Long)
    If Elements(index).mask = True Then
        Elements(index).maskSprite.Draw targetHDC, Elements(index).x, Elements(index).y, SRCAND
        Elements(index).sprite.Draw targetHDC, Elements(index).x, Elements(index).y, SRCPAINT
    Else
        Elements(index).sprite.Draw targetHDC, Elements(index).x, Elements(index).y, SRCCOPY
    End If
End Sub
Public Sub DrawID(ID As Integer, targetHDC As Long)
    Dim i As Integer
    For i = 1 To elementCount
        If Elements(i).ID = ID Then
            Draw i, targetHDC
            Exit For
        End If
    Next i
End Sub
Public Sub DrawAll(targetHDC As Long)
    Dim i As Integer
    If elementCount < 1 Then Exit Sub
    For i = 1 To elementCount
        Draw i, targetHDC
    Next i
End Sub
'Public Sub SetY(ID As Integer, y As Integer)
'    Elements(ID).y = y
'End Sub
'Public Sub SetZIndex(ID As Integer, zIndex As Integer)
'    Elements(ID).zIndex = zIndex
'End Sub
'Public Sub SetX(ID As Integer, x As Integer)
'    Elements(ID).x = x
'End Sub
'Public Sub SetSprite(ID As Integer, sprite As sprite)
'    Elements(ID).sprite = sprite
'End Sub
'Public Function GetSprite(ID As Integer) As sprite
'    GetSprite = Elements(ID).sprite
'End Function
'Public Function GetX(ID As Integer) As Integer
'    GetX = Elements(ID).x
'End Function
'Public Function GetY(ID As Integer) As Integer
'    GetY = Elements(ID).y
'End Function
Public Function GetZIndex(index As Integer) As Integer
    GetZIndex = Elements(index).zIndex
End Function
Private Sub Class_Initialize()
    elementCount = 0
End Sub
