VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5970
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   11100
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "结束"
      Height          =   855
      Left            =   4560
      TabIndex        =   5
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "登录"
      Height          =   735
      Left            =   1680
      TabIndex        =   4
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   3600
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   3600
      TabIndex        =   2
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "QQ密码"
      Height          =   735
      Left            =   1080
      TabIndex        =   1
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "QQ账号"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Text1.Text = "" Then MsgBox "QQ账号不能为空！", , "确认": Exit Sub

b = Len(Text2.Text)
If b < 6 Then MsgBox "QQ密码长度必须六位数以上！", , "确认"
End Sub

Private Sub Command2_Click()
End
End Sub
