VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3360
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "判断质数"
      Height          =   615
      Left            =   1800
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "请输入一个大于2的正整数："
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim n As Integer
    Dim i As Integer
    Dim f As Boolean
    n = Text1.Text
    f = True '假设n是质数
    For i = 2 To n - 1
        If n Mod i = 0 Then '如果i是n的约数
            f = False '推翻假定
            Exit For  '提前退出循环
        End If
    Next i
    If f = True Then '判断是否是质数
       Label2.Caption = n & "是质数！"
    Else
        Label2.Caption = n & "不是质数！"
    End If
End Sub
