VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   10560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "在等差数列插入一个数"
      Height          =   975
      Left            =   2520
      TabIndex        =   0
      Top             =   3480
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim a(11) As Integer
    Dim i As Integer, k As Integer, t As Integer
    t = Val(InputBox("插入"))
    For i = 0 To 9
        a(i) = i * 3 + 1
    Print a(i);
    Next i
    Print
    Print "插入" & t
    For k = 0 To 10
        If t < a(k) Then Exit For
    Next k
    For i = 9 To k Step -1
        a(i + 1) = a(i)
    Next i
    a(k) = t
    For i = 0 To 10
        Print a(i);
    Next i
    Print
End Sub
