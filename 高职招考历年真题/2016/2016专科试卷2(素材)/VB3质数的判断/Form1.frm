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
      Caption         =   "判断质数"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "请输入一个大于2的正整数："
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim n As Integer, i As Integer, flag As Boolean, res As String
n = Val(Text1.Text)
flag = True
For i = 2 To n - 1
    If n Mod i = 0 Then
        flag = False
        Exit For
    End If
Next i

If flag Then
    res = n & "是质数!"
Else
    res = n & "不是质数！"
End If
Label2.Caption = res

End Sub

