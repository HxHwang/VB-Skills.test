VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4080
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "确定"
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   3360
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      Caption         =   "仅允许网络身份认证用户"
      Enabled         =   0   'False
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   2520
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin VB.OptionButton Option2 
      Caption         =   "允许远程连接"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   2040
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "不允许远程连接"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "高级..."
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "允许远程协助这台计算机"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "选择一个选项，指定谁可以连接"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
