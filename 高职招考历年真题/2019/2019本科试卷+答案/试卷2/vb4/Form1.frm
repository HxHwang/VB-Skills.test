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
   Begin VB.CommandButton Command4 
      Caption         =   "退出"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "输出"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "请输入字母2"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "请输入字母1"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   735
      Left            =   1320
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String
Private Sub Command1_Click()

Dim one As String
one = InputBox("请输入字母1", "输入框")
str = str & one & Space(1)

End Sub

Private Sub Command2_Click()

Dim two As String
two = InputBox("请输入字母2", "输入框")
str = str & two & Space(1)

End Sub

Private Sub Command3_Click()
Label1.Caption = "A到D之间的字母有：" & str
End Sub

Private Sub Command4_Click()
    End
End Sub
