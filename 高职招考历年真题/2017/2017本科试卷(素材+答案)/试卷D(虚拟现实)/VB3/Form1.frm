VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "运算"
      Height          =   855
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Width           =   3735
      Begin VB.OptionButton Option4 
         Caption         =   "整除"
         Height          =   375
         Left            =   2760
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option3 
         Caption         =   "乘"
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "减"
         Height          =   375
         Left            =   960
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "加"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "结果"
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "请输入两个整数"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m As Integer, n As Integer
Private Sub Option1_Click()
m = Val(Text1.Text)
n = Val(Text2.Text)
Label3.Caption = m + n
End Sub

Private Sub Option2_Click()
m = Val(Text1.Text)
n = Val(Text2.Text)
Label3.Caption = m - n
End Sub

Private Sub Option3_Click()
m = Val(Text1.Text)
n = Val(Text2.Text)
Label3.Caption = m * n
End Sub

Private Sub Option4_Click()
m = Val(Text1.Text)
n = Val(Text2.Text)
Label3.Caption = m \ n '注意：是整除\
End Sub
