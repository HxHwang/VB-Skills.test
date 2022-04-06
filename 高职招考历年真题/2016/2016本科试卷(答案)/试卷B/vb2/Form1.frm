VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4185
   ClientLeft      =   8775
   ClientTop       =   4125
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   4560
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   615
      Left            =   2280
      TabIndex        =   7
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "判断三角形"
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "c="
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "b="
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   1500
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "a="
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Cls
Dim a, b, c As Single
a = Val(Text1.Text)
b = Val(Text2.Text)
c = Val(Text3.Text)
If (a + b) > c And (b + c) > a And (c + a) > b Then
Print "可以构成三角形"
Else
Print "不能构成三角形"
End If
End Sub

Private Sub Command2_Click()
End
End Sub

