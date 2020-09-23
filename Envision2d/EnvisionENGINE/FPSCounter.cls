VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FPSCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim Frames As Integer
Dim FPS As Integer
Dim FrameStart As Long
Public Sub Frame()
    If GetTickCount - FrameStart >= 1000 Then    'Over a second
        FrameStart = GetTickCount
        FPS = Frames
        Frames = 0
    End If
    Frames = Frames + 1
End Sub
Public Function GetFPS() As Integer
    GetFPS = FPS
End Function
Private Sub Class_Initialize()
    Frames = 0
    FrameStart = GetTickCount
End Sub
