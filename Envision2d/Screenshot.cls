VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Screenshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public path As String
Public fileName As String
Public screenCount As Integer
Public Sub Save(pictureObject As Object)
    SavePicture pictureObject.Image, path & fileName & screenCount & ".bmp"
    screenCount = screenCount + 1
End Sub
Private Sub Class_Initialize()
    screenCount = 1
End Sub
