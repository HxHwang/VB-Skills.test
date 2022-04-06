VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "最小值"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "计算"
      Height          =   855
      Left            =   2640
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "最小值"
      Height          =   855
      Left            =   600
      TabIndex        =   2
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   360
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
Dim a%, b%, c%
a = InputBox("请输入第一个数：", "判断最小数")
b = InputBox("请输入第二个数：", "判断最小数")
If a < b Then
Print a
Label1.Caption = a
Label2.Caption = b

Else

Print b
Label1.Caption = b
Label2.Caption = a
End If
End Sub

Private Sub Command2_Click()
Dim a%, b%, i%, s%
a = Val(Label1.Caption)
b = Val(Label2.Caption)
For i = a To b
s = s + i
Next i
Print s


End Sub
