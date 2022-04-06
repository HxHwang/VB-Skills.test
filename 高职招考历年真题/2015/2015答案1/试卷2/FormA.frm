VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   5775
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "第5个字符"
      Height          =   495
      Left            =   3180
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "字符长度"
      Height          =   495
      Left            =   1140
      TabIndex        =   2
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   3060
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1140
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Text2.Text = Len(Text1.Text)
End Sub

Private Sub Command2_Click()
Text2.Text = Mid(Text1.Text, 5, 1)
End Sub
