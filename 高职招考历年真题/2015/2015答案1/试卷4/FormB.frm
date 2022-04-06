VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "最大值"
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
      Height          =   735
      Left            =   2400
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "最大值"
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   855
      Left            =   2400
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a%, b%
Private Sub Command1_Click()
a = InputBox("请输入第一数字", "判断最大数")
b = InputBox("请输入第二数字", "判断最大数")
If a > b Then
Print a
Label1.Caption = b
Label2.Caption = a
Else
Print b
Label1.Caption = a
Label2.Caption = b
End If

End Sub

Private Sub Command2_Click()
Dim i%, s%
a = Val(Label1.Caption)
b = Val(Label2.Caption)
For i = a To b
If i Mod 5 = 0 Then
s = s + i
End If
Next i
Print s
End Sub
