VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim StartTime As Long
Dim EndTime As Long
Dim Stopped As Boolean
Public Sub StartTimer()
    StartTime = GetTickCount
    Stopped = False
End Sub
Public Function StopTimer() As Long
    If Stopped Then
        StopTimer = EndTime
    Else
        Stopped = True
        EndTime = GetTickCount - StartTime
        StopTimer = EndTime
    End If
End Function
Public Function GetTime() As Long
    If Stopped Then
        GetTime = EndTime
    Else
        GetTime = GetTickCount - StartTime
    End If
End Function

