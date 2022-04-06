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
   Begin VB.CommandButton Command1 
      Caption         =   "求阶乘"
      Height          =   615
      Left            =   1920
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "请输入一个正整数："
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim i As Integer
    Dim s As Long
    Dim n As Integer
    s = 1
    n = Text1.Text
    For i = 1 To n
        s = s * i
    Next i
    Label2.Caption = n & "的阶乘为：" & s
End Sub
