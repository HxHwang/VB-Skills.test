VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "打印"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "请输入一个整数【1-20】"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim num As Integer
Dim i As Integer, j As Integer
num = Val(Text1) ' 将Text1的文本转为数值，赋值给num 方便计算
'根据你输入的不同数字，循环次数也会不同
For i = 1 To num 'i控制的是行数
    For j = 1 To num 'j控制的是列数
        Select Case j  ' 测试表达式
            Case Is >= i '这一步很重要，慢慢体会！说不清楚... ...
                Print "1"; Space(1);
            Case Else
                Print "0"; Space(1);
        End Select
    Next j
    Print
Next i
End Sub
