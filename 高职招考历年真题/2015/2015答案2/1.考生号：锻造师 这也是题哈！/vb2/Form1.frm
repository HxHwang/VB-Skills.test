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
      Caption         =   "缴纳水费"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'x是用水量，y是水费
Dim x As Single, y As Single
x = Val(InputBox("请输入用水量", "计算水费", "23"))
Label1.Caption = x
Select Case x
    Case Is <= 18
        y = 1.2 * x
    Case 18 To 25 '范围是18<x<=25
        y = 1.2 * 18 + 1.8 * (x - 18)
    Case Is > 25
        y = 1.2 * 18 + 1.8 * (25 - 18) + 2.4 * (x - 25)
End Select
Label2.Caption = y

End Sub
