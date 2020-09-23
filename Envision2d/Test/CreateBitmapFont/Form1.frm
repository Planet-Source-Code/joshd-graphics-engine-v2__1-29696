VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   272
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   601
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   2280
      Width           =   975
   End
   Begin VB.PictureBox picHolder 
      Height          =   1935
      Left            =   120
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   581
      TabIndex        =   4
      Top             =   120
      Width           =   8775
      Begin VB.PictureBox picLetters 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tempus Sans ITC"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3750
         Left            =   0
         ScaleHeight     =   250
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1000
         TabIndex        =   5
         Top             =   0
         Width           =   15000
      End
   End
   Begin VB.PictureBox picOutput 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      Height          =   735
      Left            =   120
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   581
      TabIndex        =   3
      Top             =   3240
      Width           =   8775
   End
   Begin VB.CommandButton cmdWrite 
      Caption         =   "Test"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox txtText 
      Height          =   885
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2280
      Width           =   3135
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "GO"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      Height          =   855
      Left            =   5760
      TabIndex        =   7
      Top             =   2280
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CharX(0 To 155) As Integer
Dim CharY(0 To 155) As Integer
Dim CharHeight(0 To 155) As Integer
Dim CharWidth(0 To 155) As Integer
Dim maxWidth As Integer
Private Sub cmdGo_Click()
    Dim i As Integer
    maxWidth = 0
    picLetters.Cls: picLetters.Print " ";
    For i = 65 To 90
        addLetter (i)   'Upper Case
    Next i
    picLetters.Print "": picLetters.Print " ";
    For i = 97 To 122
        addLetter (i)   'Lower case
    Next i
    picLetters.Print "": picLetters.Print " ";
    For i = 32 To 64
        addLetter (i)   'Numbers/Special Characters
    Next i
    For i = 91 To 96
        addLetter (i)  'Various characters
    Next i
    picLetters.Print ""
    picLetters.Width = maxWidth
    picLetters.Height = picLetters.CurrentY
    'picLetters.Refresh
End Sub
Public Sub addLetter(code As Integer)
    CharX(code) = picLetters.CurrentX
    CharY(code) = picLetters.CurrentY
    CharWidth(code) = picLetters.TextWidth(Chr(code))
    CharHeight(code) = picLetters.TextHeight(Chr(code))
    picLetters.Print Chr(code) & " ";
    If maxWidth < picLetters.CurrentX Then maxWidth = picLetters.CurrentX
End Sub

Private Sub cmdSave_Click()
    Dim i As Integer
    Call SavePicture(picLetters.Image, App.Path & "\font.bmp")
    Open App.Path & "\font.env" For Output As #1
        For i = 65 To 90
            Write #1, i, CharX(i), CharY(i), CharWidth(i), CharHeight(i)
        Next i
        For i = 97 To 122
            Write #1, i, CharX(i), CharY(i), CharWidth(i), CharHeight(i)
        Next i
        For i = 32 To 64
            Write #1, i, CharX(i), CharY(i), CharWidth(i), CharHeight(i)
        Next i
        For i = 91 To 96
            Write #1, i, CharX(i), CharY(i), CharWidth(i), CharHeight(i)
        Next i
    Close #1
End Sub

Private Sub cmdWrite_Click()
    Dim i As Integer
    Dim x As Integer, y As Integer, xMargin As Integer
    Dim letter As String, code As Integer
    x = 5
    y = 5
    xMargin = 5
    For i = 1 To Len(txtText.Text)
        letter = Mid(txtText.Text, i, 1)
        code = Asc(letter)
        If code = 13 Then   'Next Line
            x = xMargin
            y = y + CharHeight(65)
        Else
            Call BitBlt(picOutput.hDC, x, y, CharWidth(code), CharHeight(code), picLetters.hDC, CharX(code), CharY(code), SRCCOPY)
            x = x + CharWidth(code)
        End If
    Next i
    picOutput.Refresh
End Sub

