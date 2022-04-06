VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "算术运算"
   ClientHeight    =   1695
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   ScaleHeight     =   1695
   ScaleWidth      =   4470
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "除"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "乘"
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "减"
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "加"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "计算结果"
      Height          =   300
      Left            =   3120
      TabIndex        =   2
      Top             =   240
      Width           =   1320
   End
   Begin VB.Label Label2 
      Caption         =   "操作数2"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "操作数1"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m As Integer, n As Integer
Private Sub Command1_Click()
m = Val(Text1.Text)
n = Val(Text2.Text)
Text3.Text = m + n
End Sub

Private Sub Command2_Click()
m = Val(Text1.Text)
n = Val(Text2.Text)
Text3.Text = m - n
End Sub

Private Sub Command3_Click()
m = Val(Text1.Text)
n = Val(Text2.Text)
Text3.Text = m * n
End Sub

Private Sub Command4_Click()
m = Val(Text1.Text)
n = Val(Text2.Text)
Text3.Text = m / n
End Sub
