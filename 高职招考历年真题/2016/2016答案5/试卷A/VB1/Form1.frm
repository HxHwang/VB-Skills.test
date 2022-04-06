VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "求平方根"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "求平方根"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim x As Single
x = Val(Text1.Text)
If x > 0 Then
    Text2.Text = Sqr(x)
Else
    Text2.Text = "不能输入负数哦！"
End If

End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Form_Load()
'清空text1和text2
Text1.Text = ""
Text2.Text = ""
End Sub
