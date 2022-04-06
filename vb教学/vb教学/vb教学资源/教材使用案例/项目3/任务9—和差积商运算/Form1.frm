VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5565
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   15.75
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   5565
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdAdd 
      Caption         =   "＋"
      Height          =   495
      Left            =   840
      TabIndex        =   12
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton CmdSub 
      Caption         =   "－"
      Height          =   495
      Left            =   1800
      TabIndex        =   11
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton CmdMult 
      Caption         =   "×"
      Height          =   495
      Left            =   2760
      TabIndex        =   10
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton CmdDiv 
      Caption         =   "÷"
      Height          =   495
      Left            =   3720
      TabIndex        =   9
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton CmdEnd 
      Caption         =   "退出"
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton CmdCls 
      Caption         =   "清除"
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton CmdRes 
      Caption         =   "重新"
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label LblResult 
      Height          =   315
      Left            =   3840
      TabIndex        =   5
      Top             =   1320
      Width           =   1620
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "＝"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   3360
      TabIndex        =   4
      Top             =   1320
      Width           =   330
   End
   Begin VB.Label LblNumber2 
      AutoSize        =   -1  'True
      Height          =   315
      Left            =   2520
      TabIndex        =   3
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label LblSymbol 
      AutoSize        =   -1  'True
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label LblNumber1 
      AutoSize        =   -1  'True
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label LblTitle 
      AutoSize        =   -1  'True
      Caption         =   "和差积商运算"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   2790
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer, b As Integer, c As Double

Private Sub CmdAdd_Click()
    LblSymbol.Caption = "+"
    c = a + b
    LblResult.Caption = c
End Sub

Private Sub CmdCls_Click()
    LblNumber1.Caption = ""
    LblSymbol.Caption = ""
    LblNumber2.Caption = ""
    LblResult.Caption = ""
End Sub

Private Sub CmdDiv_Click()
    LblSymbol.Caption = "÷"
    c = a / b
    LblResult.Caption = c
End Sub

Private Sub CmdEnd_Click()
    End
End Sub

Private Sub CmdMult_Click()
    LblSymbol.Caption = "×"
    c = a * b
    LblResult.Caption = c
End Sub

Private Sub CmdRes_Click()
    a = CInt(Rnd * 100)
    b = CInt(Rnd * 100)
    LblNumber1.Caption = a
    LblNumber2.Caption = b
    LblSymbol.Caption = ""
    LblResult.Caption = ""
End Sub

Private Sub CmdSub_Click()
    LblSymbol.Caption = "-"
    c = a - b
    LblResult.Caption = c
End Sub

Private Sub Form_Load()
    a = CInt(Rnd * 100)
    b = CInt(Rnd * 100)
    LblNumber1.Caption = a
    LblNumber2.Caption = b
End Sub
