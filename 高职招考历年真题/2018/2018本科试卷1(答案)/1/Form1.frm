VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "请输入一个数："
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim m
m = Val(Text1.Text)
If m > 0 Then
Label2.Caption = Sqr(m)
Else
m = Abs(m)
Label2.Caption = Sqr(m)
End If

End Sub
