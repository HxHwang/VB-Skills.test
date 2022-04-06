VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "判断负数"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "偶数和"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "负数"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim x As Integer
Cls
x = Val(InputBox("请输入数据", "判断负数", 0))
If x < 0 Then Print x
End Sub

Private Sub Command2_Click()
Dim m, n, k, i, sum As Integer
m = Val(Text1)
n = Val(Text2)
If m > n Then
k = -1
Else
k = 1
End If

For i = m To n Step k
If i Mod 2 = 0 Then
sum = sum + i
End If
Next
Print sum

End Sub
