VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Get Time"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Stop"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Frame Increment"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "File"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Path"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   105
      X2              =   5025
      Y1              =   660
      Y2              =   660
   End
   Begin VB.Label lblTimer 
      Caption         =   "0"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblFPS 
      Caption         =   "0"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Engine As New Envision
Dim Counter As New FPSCounter
Dim MyTimer As New StopWatch
Private Sub Command1_Click()
    MsgBox Engine.GetPath(App.Path)
    MsgBox Engine.DLLPath
End Sub

Private Sub Command2_Click()
    MsgBox Engine.FileExists(Engine.GetPath(App.Path) & "file.txt")
End Sub

Private Sub Command3_Click()
    Counter.Frame
    lblFPS.Caption = Counter.GetFPS
End Sub

Private Sub Command4_Click()
    MyTimer.StartTimer
End Sub

Private Sub Command5_Click()
    lblTimer.Caption = MyTimer.StopTimer
End Sub

Private Sub Command6_Click()
    lblTimer.Caption = MyTimer.GetTime
End Sub

Private Sub Form_Load()

End Sub
