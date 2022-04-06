VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   3870
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "面积"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "宽"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "长"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Single, b As Single
Private Sub Command1_Click()
a = Text1.Text
b = Text2.Text
area a, b
End Sub
