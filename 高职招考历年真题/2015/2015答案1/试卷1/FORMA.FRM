VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   593
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2633
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "绝对值"
      Height          =   615
      Left            =   593
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "十位数"
      Height          =   615
      Left            =   2633
      TabIndex        =   0
      Top             =   1920
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

Text2.Text = Abs(Val(Text1.Text))

End Sub

Private Sub Command2_Click()

'Text2.Text = Val(Text1.Text) \ 10  '方法一
 
Text2.Text = Val(Right(Text1.Text, 2)) \ 10   '方法二 ，更优解


'Text2.Text = abs(Val(Text1.Text)) \ 10 Mod 10  '方法三，更优解
End Sub
