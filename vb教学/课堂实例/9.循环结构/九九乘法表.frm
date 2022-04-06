VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   9615
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, j As Integer
Dim s As String
Private Sub Form_Click()
Print Tab(30); "九九乘法表"
Print Tab(29); "-----------"
For i = 1 To 9
 For j = 1 To i
   s = i & "*" & j & "=" & i * j
   Print Tab((j - 1) * 8 + 1); s;
 Next j
 Print
 Print
Next i
End Sub

