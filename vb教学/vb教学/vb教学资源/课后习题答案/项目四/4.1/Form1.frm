VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   8625
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   4800
      TabIndex        =   6
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "显示结果"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4560
      TabIndex        =   3
      Top             =   5640
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   6120
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   4500
      TabIndex        =   1
      Top             =   4320
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   4800
      TabIndex        =   0
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   " 四   则   运   算"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   1080
      TabIndex        =   8
      Top             =   240
      Width           =   5775
   End
   Begin VB.Label Label4 
      Caption         =   "输入+-*/运算符"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "请输入第二个数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   4320
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "请输入第一个数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   2760
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Long, b As Long, m As Long
Dim x As String
Private Sub Command1_Click()
x = Text4.Text
a = Val(Text1.Text)
b = Val(Text2.Text)
Select Case x
    Case "+"
      m = a + b
    Case "-"
      m = a - b
    Case "*"
      m = a * b
    Case "/"
     m = a / b
End Select
Text3.Text = Str$(m)
End Sub

