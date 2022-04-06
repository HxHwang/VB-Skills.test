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
      Caption         =   "求降序数"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'i 控制所有10位数
' cou是计数用的 如果 cou为10 则 换行显示
' ge、shi 就是 取十位数的个位和十位
Dim i As Integer, cou As Integer
Dim ge As Integer, shi As Integer
Print "所有2位的降序数有："
For i = 10 To 99
    ge = i Mod 10   '对10取余可得到个位
    shi = i \ 10 Mod 10  ' 先除以10 然后再 取余 即可得到十位
    If shi >= ge Then  '按照降序的算法 a>=b  所谓高位不低于临位 就是 十位大于个位的意思！
        Print i;   ' 分号是 每输出一个数字 就 在同一行输出 否则 会直接到下一行
        cou = cou + 1  '计数是为了 统计 换行的条件
        If cou = 10 Then ' 等于10了，那么我可以换行啦！
            cou = 0 '换行的同时，将换行的条件继续设置为0，这样就会反复换行！
            Print '因为上面的i后面有个分号，所以 一但你想要换行了 就必须输入个空的print
        End If
    End If
Next i
'因为题目有算法提示，整体来说 不难！

'小知识： /   和  \  的区别

' / 是浮点除
' \ 是整除

'例如 7/2=3.5    7\2=3
End Sub

