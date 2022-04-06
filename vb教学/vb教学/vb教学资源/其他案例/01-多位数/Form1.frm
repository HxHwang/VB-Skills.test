VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   6915
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "显示"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   5160
      TabIndex        =   10
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   4380
      TabIndex        =   9
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   3600
      TabIndex        =   8
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   2820
      TabIndex        =   7
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1260
      TabIndex        =   5
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2160
      TabIndex        =   3
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "七位数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "多位数分位显示"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1845
      TabIndex        =   0
      Top             =   360
      Width           =   3285
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Dim x As Long, a As Long, b As Long, c As Long, d As Long, _
        e As Integer, f As Integer, g As Integer
    x = Val(Text1.Text)
    Text2.Text = Str$(x \ 1000000)
    a = x Mod 1000000
    Text3.Text = Str$(a \ 100000)
    b = a Mod 100000
    Text4.Text = Str$(b \ 10000)
    c = b Mod 10000
    Text5.Text = Str$(c \ 1000)
    d = c Mod 1000
    Text6.Text = Str$(d \ 100)
    e = d Mod 100
    Text7.Text = Str$(e \ 10)
    f = e Mod 10
    Text8.Text = Str$(f)
End Sub


