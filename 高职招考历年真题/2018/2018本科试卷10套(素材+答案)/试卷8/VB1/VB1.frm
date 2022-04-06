VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "修改"
      Height          =   615
      Left            =   1680
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "新密码"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "旧密码"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Text1.Text = Text2.Text Or Len(Text2.Text) < 6 Then
        Print "修改失败"
    Else
        Print "修改成功"
    End If
End Sub
