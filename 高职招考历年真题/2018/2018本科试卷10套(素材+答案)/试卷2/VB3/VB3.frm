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
      Caption         =   "转换"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "输入百分制成绩"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    x = Val(Text1.Text)
    If x >= 90 Then
        MsgBox "优秀"
    ElseIf x >= 80 Then
        MsgBox "良好"
    ElseIf x >= 70 Then
        MsgBox "中等"
    ElseIf x >= 60 Then
        MsgBox "及格"
    Else
        MsgBox "不及格"
    End If
End Sub
