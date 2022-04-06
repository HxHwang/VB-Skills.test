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
   Begin VB.CommandButton Command2 
      Caption         =   "正方形面积"
      Height          =   615
      Left            =   2520
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "符号函数"
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   2520
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   360
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
Dim x%
x = Sgn(Val(Text1.Text))



End Sub

Private Sub Command2_Click()
Dim x%
x = Val(Text1.Text)
Text2.Text = x ^ 2
End Sub
