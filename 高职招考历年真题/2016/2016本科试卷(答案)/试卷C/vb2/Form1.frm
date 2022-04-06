VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3960
   ClientLeft      =   9285
   ClientTop       =   4170
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   4560
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   2040
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2040
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2040
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "运费"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "票价"
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   1380
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "重量"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim z, pj, yf As Single
z = Val(Text1.Text)
pj = Val(Text2.Text)
yf = 0
If z > 20 Then yf = pj * 0.015 * (z - 20)
Text3.Text = yf
End Sub

Private Sub Command2_Click()
End
End Sub

