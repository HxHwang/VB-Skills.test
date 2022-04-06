VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   ScaleHeight     =   1830
   ScaleWidth      =   4020
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Text            =   "3"
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton C1 
      Caption         =   "确定"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label L2 
      Caption         =   "允许次数"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label L1 
      Caption         =   "口令"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub C1_Click()
If Text1.Text = "123456" Then
        Text1.Text = "口令正确"
        Text1.PasswordChar = ""
    Else
        Text2.Text = Text2.Text - 1
        If Text2.Text > 0 Then
            MsgBox "第" & (3 - Text2.Text) & "次口令错误，请重新输入"
        Else
            MsgBox "3次输入错误，请退出"
            Text1.Enabled = False
        End If
    End If

End Sub

