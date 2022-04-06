VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "函数运算"
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   BeginProperty Font 
      Name            =   "楷体_GB2312"
      Size            =   15.75
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   6705
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdEnd 
      Caption         =   "退出"
      Height          =   495
      Left            =   4800
      TabIndex        =   8
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton CmdSQR 
      Caption         =   "SQR"
      Height          =   495
      Left            =   3720
      TabIndex        =   7
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton CmdTAN 
      Caption         =   "TAN"
      Height          =   495
      Left            =   2600
      TabIndex        =   6
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton CmdCOS 
      Caption         =   "COS"
      Height          =   495
      Left            =   1480
      TabIndex        =   5
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton CmdSIN 
      Caption         =   "SIN"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton CmdCls 
      Caption         =   "清除"
      Height          =   495
      Left            =   5520
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox TxtY 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   615
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   3615
   End
   Begin VB.TextBox TxtX 
      Alignment       =   2  'Center
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "函数运算"
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
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   1980
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Double, Y As Double
Const PI = 3.1415926

Private Sub CmdCls_Click()
    TxtX.Text = ""
    TxtY.Text = ""
End Sub

Private Sub CmdCOS_Click()
    X = Val(TxtX.Text)
    Y = Cos(X)
    TxtY.Text = Str(Y)
End Sub

Private Sub CmdEnd_Click()
    End
End Sub

Private Sub CmdSIN_Click()
    X = Val(TxtX.Text)
    Y = Sin(X)
    TxtY.Text = Str(Y)
End Sub

Private Sub CmdSQR_Click()
    X = Val(TxtX.Text)
    Y = Sqr(X)
    TxtY.Text = Str(Y)
End Sub

Private Sub CmdTAN_Click()
    X = Val(TxtX.Text)
    Y = Tan(X)
    TxtY.Text = Str(Y)
End Sub
