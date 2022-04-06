VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8040
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   15.75
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   8040
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(10) As Integer, i As Integer
Private Sub Form_Click()
 Print: Print
            '输入数组元素
 For i = 1 To 10
  a(i) = InputBox("请输入数组", "输入", 0)
  '输出数组元素
Print a(i);
 Next i
 Print: Print
            '交换数组元素
 For i = 1 To 5
  t = a(i)
  a(i) = a(11 - i)
  a(11 - i) = t
 Next i
            '输出交换后的数组元素
 For i = 1 To 10
  Print a(i);
 Next i
End Sub

