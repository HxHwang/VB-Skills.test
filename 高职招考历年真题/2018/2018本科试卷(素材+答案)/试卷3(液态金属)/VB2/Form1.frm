VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox Check2 
      Caption         =   "雪碧"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "可乐"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      Caption         =   "炒面"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   1680
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "米饭"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1080
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "点餐统计"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   2400
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Option1.Value = True Then
        Text1.Text = "您点的食物是：" & Option1.Caption
    Else
        Text1.Text = "您点的食物是：" & Option2.Caption
    End If
    If Check1.Value = 1 Then
        Text1.Text = Text1.Text & "  " & Check1.Caption
    End If
    If Check2.Value = 1 Then
        Text1.Text = Text1.Text & "  " & Check2.Caption
    End If
End Sub
