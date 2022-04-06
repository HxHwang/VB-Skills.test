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
   Begin VB.CommandButton Command4 
      Caption         =   "退出"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "计算"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "请输入b的值"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "请输入a的值"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sum As Integer

Private Sub Command1_Click()
Dim a As Integer
a = Val(InputBox("请输入a的值", "输入框"))
sum = sum + a
End Sub

Private Sub Command2_Click()
Dim b As Integer
b = Val(InputBox("请输入b的值", "输入框"))
sum = sum + b
End Sub

Private Sub Command3_Click()
Label1.Caption = "a-b之间的和是：" & sum
End Sub

Private Sub Command4_Click()
    End
End Sub
