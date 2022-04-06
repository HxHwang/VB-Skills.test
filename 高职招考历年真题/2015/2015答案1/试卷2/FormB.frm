VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "判断负数"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "偶数和"
      Height          =   615
      Left            =   2640
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "负数"
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim a%
a = Val(InputBox("请输入数据", "判断负数"))
If a < 0 Then
Print a
End If
End Sub

Private Sub Command2_Click()
Dim m%, n%, i%, s%

m = Val(Text1.Text)
n = Val(Text2.Text)
For i = m To n
    If i Mod 2 = 0 Then
        s = s + i
    End If
Next i
Print s
End Sub
