VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "退出"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   1560
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Text            =   "请输入！"
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Text1_Change()
Text2.Text = Text1.Text
Text2.FontSize = 18
Text3.Text = Text1.Text
Text3.FontSize = 24
End Sub
