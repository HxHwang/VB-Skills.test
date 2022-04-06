VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   11295
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "运算"
      Height          =   495
      Left            =   5040
      TabIndex        =   6
      Top             =   3600
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "第三题.frx":0000
      Left            =   840
      List            =   "第三题.frx":0010
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Left            =   2760
      TabIndex        =   4
      Top             =   2760
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "答案："
      Height          =   180
      Left            =   960
      TabIndex        =   3
      Top             =   2760
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "输入两个整数："
      Height          =   180
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   1260
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Combo1.Text = "+" Then Label3 = Val(Text1) + Val(Text2)
If Combo1.Text = "-" Then Label3 = Val(Text1) - Val(Text2)
If Combo1.Text = "*" Then Label3 = Val(Text1) * Val(Text2)
If Combo1.Text = "\" Then Label3 = Val(Text1) \ Val(Text2)
End Sub
