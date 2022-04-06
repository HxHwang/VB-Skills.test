VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "最大值"
   ClientHeight    =   3030
   ClientLeft      =   9750
   ClientTop       =   2820
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   Begin VB.CommandButton Command2 
      Caption         =   "计算"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "最大值"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   720
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
Dim m, n As Integer

Private Sub Command1_Click()
Cls
m = Val(InputBox("请输入第一数字", "判断最小数", 0))
n = Val(InputBox("请输入第二数字", "判断最小数", 0))
If m < n Then
Print n
Label1.Caption = m
Label2.Caption = n
Else
Print m
Label1.Caption = n
Label2.Caption = m
End If
m = Val(Label1.Caption)
n = Val(Label2.Caption)
End Sub

Private Sub Command2_Click()
Dim i As Integer
Dim sum As Double

For i = m To n
If i Mod 5 = 0 Then
sum = sum + i
End If
Next
Print sum
End Sub


