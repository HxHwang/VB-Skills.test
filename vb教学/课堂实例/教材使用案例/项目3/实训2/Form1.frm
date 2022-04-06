VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   BeginProperty Font 
      Name            =   "楷体_GB2312"
      Size            =   21.75
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   6870
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdEnd 
      Caption         =   "退出"
      Height          =   615
      Left            =   4200
      TabIndex        =   9
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton CmdCls 
      Caption         =   "清除"
      Height          =   615
      Left            =   2520
      TabIndex        =   8
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton CmdSub 
      Caption         =   "－"
      Height          =   615
      Left            =   1560
      TabIndex        =   7
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "＋"
      Height          =   615
      Left            =   600
      TabIndex        =   6
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox Txt2 
      Alignment       =   2  'Center
      Height          =   615
      Left            =   3000
      TabIndex        =   2
      Top             =   1230
      Width           =   1695
   End
   Begin VB.TextBox Txt1 
      Alignment       =   2  'Center
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1230
      Width           =   1695
   End
   Begin VB.Label LblResult 
      AutoSize        =   -1  'True
      Height          =   435
      Left            =   5400
      TabIndex        =   5
      Top             =   1320
      Width           =   240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "＝"
      Height          =   435
      Left            =   4800
      TabIndex        =   4
      Top             =   1320
      Width           =   465
   End
   Begin VB.Label LblSymbol 
      AutoSize        =   -1  'True
      Height          =   435
      Left            =   2400
      TabIndex        =   3
      Top             =   1320
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "加减法运算"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Double, b As Double, c As Double

Private Sub CmdAdd_Click()
    a = Val(Txt1.Text)
    b = Val(Txt2.Text)
    c = a + b
    LblSymbol.Caption = "+"
    LblResult.Caption = Str(c)
End Sub


Private Sub CmdCls_Click()
    Txt1.Text = ""
    Txt2.Text = ""
    LblSymbol.Caption = ""
    LblResult.Caption = ""
End Sub

Private Sub CmdEnd_Click()
    End
End Sub

Private Sub CmdSub_Click()
    a = Val(Txt1.Text)
    b = Val(Txt2.Text)
    c = a - b
    LblSymbol.Caption = "-"
    LblResult.Caption = Str(c)
End Sub
