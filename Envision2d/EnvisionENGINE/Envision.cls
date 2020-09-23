VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Envision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function GetPath(Path As String) As String
    'If the path is the root directory then the path will
    'end in a slash. Otherwise the path will end in a slash
    'this function ensures it always will
    If Right(Path, 1) = "/" Or Right(Path, 1) = "\" Then
        GetPath = Path
    Else
        GetPath = Path & "\"
    End If
End Function
Public Function DLLPath() As String
    'If the path is the root directory then the path will
    'end in a slash. Otherwise the path will end in a slash
    'this function ensures it always will
    If Right(App.Path, 1) = "/" Or Right(App.Path, 1) = "\" Then
        DLLPath = App.Path
    Else
        DLLPath = App.Path & "\"
    End If
End Function
Public Function FileExists(FileName As String) As String
    'Does the filename given exist
    FileExists = (Dir$(FileName) <> "")
End Function
Public Function PathExists(Path As String) As String
    'To be implemented
    PathExists = False
End Function
Public Function ChangeRes(Width As Integer, Height As Integer) As Boolean
    'To be implemented
    ChangeRes = False
End Function
Public Function RestoreRes() As Boolean
    'To be implemented
    RestoreRes = False
End Function


