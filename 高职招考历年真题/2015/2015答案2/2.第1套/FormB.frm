VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "判断正数"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "等级评价"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "显示正数"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'知识备注：Sgn(x) 求x的符号 x>0 返回1；x=0 返回0；x<0 返回-1；
Private Sub Command1_Click()
Dim number As Integer

number = Val(InputBox("请输入数据", "判断正数", "45")) '因为inputbox默认返回string哦

If number > 0 Then '想高端的话 这里可以改为 if Sgn(number)=1 then
    Print number
Else
    MsgBox "请输入正数"
End If
End Sub

Private Sub Command2_Click()
Dim number As Integer
number = Val(InputBox("请输入数据", "等级评价", "8"))
'遇到这类题，建议select case 走起
Select Case number
    Case 1 To 4
        Print "D"
    Case 5 To 10
        Print "C"
    Case 11 To 14
        Print "B"
    Case Else '以上都不满足条件，直接执行这句
        Print "A"
End Select

End Sub
