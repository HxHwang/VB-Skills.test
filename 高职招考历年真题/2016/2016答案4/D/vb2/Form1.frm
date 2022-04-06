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
      Caption         =   "比较"
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "最大的数是："
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "请输入三个数"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim a As Single
    Dim b As Single
    Dim c As Single
    Dim max As Single
    a = Text1.Text
    b = Text2.Text
    c = Text3.Text
    max = a
    If b > max Then max = b
    If c > max Then max = c
    Label2.Caption = Label2.Caption & max
End Sub
