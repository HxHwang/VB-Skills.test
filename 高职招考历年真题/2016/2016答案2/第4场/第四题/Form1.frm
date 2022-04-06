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
   Begin VB.CommandButton Command5 
      Caption         =   "退出"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "验证密码"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "初始化"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "b"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "a"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String '用来记录a按钮和b按钮的内容
Private Sub Command1_Click()
Dim a As String
a = Command1.Caption '也可直接写 a = "a"
str = str & a '累加的效果
End Sub

Private Sub Command2_Click()
Dim b As String
b = Command2.Caption ' 也可直接写 b = "b"
str = str & b '累加的效果
End Sub

Private Sub Command3_Click()
str = "" '清空你单击的a按钮和b按钮的内容
End Sub

Private Sub Command4_Click()
'判断 你单击的a按钮和b按钮存储的str字符串 是否为'abab'
If str = "abab" Then
    Form2.Show '弹出窗体2
Else
    MsgBox "密码错误，请重新输入"
    str = ""
End If
End Sub

Private Sub Command5_Click()
End
End Sub
