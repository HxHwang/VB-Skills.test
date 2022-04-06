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
      Caption         =   "转换"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "输入百分制成绩"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim n As Integer
Dim res As String
n = Val(Text1.Text)

Select Case n
    Case Is >= 90
        res = "优秀"
    Case 80 To 90
        res = "良好"
    Case 70 To 80
        res = "中等"
    Case 60 To 70
        res = "及格"
    Case 0 To 60
        res = "不及格"
    Case Else
        res = "请输入0~100的整数"
End Select

' 将结果以消息框的形式，显示出来
a = MsgBox(res, , "VB3")


End Sub
