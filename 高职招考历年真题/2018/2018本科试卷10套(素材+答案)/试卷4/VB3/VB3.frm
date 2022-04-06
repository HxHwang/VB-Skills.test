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
      Caption         =   "函数值"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "输入x"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    x = Val(Text1.Text)
    If x <= 0 Then
        MsgBox "函数值：" & x + 3
    ElseIf x < 10 Then
        MsgBox "函数值：" & x / 2
    Else
        MsgBox "函数值：" & Sqr(x) - 3
    End If
End Sub
