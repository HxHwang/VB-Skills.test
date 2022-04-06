VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "矩形面积计算"
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5550
   BeginProperty Font 
      Name            =   "楷体_GB2312"
      Size            =   15
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   5550
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   1440
      TabIndex        =   7
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   1440
      TabIndex        =   6
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1440
      TabIndex        =   5
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清空"
      Height          =   615
      Left            =   3960
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   615
      Left            =   3960
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "面积"
      Height          =   300
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "宽"
      Height          =   300
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "长"
      Height          =   300
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   315
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim area As Single
 
Sub recarea(rlen, rwid)
  area = rlen * rwid
End Sub

Private Sub Command1_Click()
    Dim a As Single, b As Single
    a = Val(Text1.Text)
    b = Val(Text2.Text)
    recarea a, b
    'Call recarea(a, b)
    Text3.Text = area
End Sub


Private Sub Command2_Click()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
End Sub

