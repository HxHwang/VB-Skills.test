VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "判断偶数"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "评价等级"
      Height          =   735
      Left            =   2640
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "偶数"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   1920
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim a As Integer
a = InputBox("请输入数据", "判断偶数")
If a Mod 2 = 0 Then
    Print a
End If
End Sub

Private Sub Command2_Click()
Dim a As interger
a = InputBox("请输入成绩", "评价等级")

    Select Case a
        Case 81 To 100
            Print "优秀"
        Case 60 To 80
            Print "合格"
        Case 0 To 59
            Print "不及格"
        Case Else
     MsgBox "输入错误", , "错误"
        End Select

End Sub
