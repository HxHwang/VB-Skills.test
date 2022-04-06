VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5925
   BeginProperty Font 
      Name            =   "ËÎÌå"
      Size            =   15.75
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   5925
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub triangle(Str As String, n As Integer)
    Dim i As Integer, j As Integer
    For i = l To n
        Print Tab(16 - i);
        For j = 1 To 2 * i - 1
            Print Str;
        Next j
        Print
    Next i
End Sub
Private Sub Form_Click()
    Dim char As String * 1, n As Integer
    char = "*"
    n = 5
    Call triangle(char, n)
    char = "+"
    triangle char, 3
End Sub

