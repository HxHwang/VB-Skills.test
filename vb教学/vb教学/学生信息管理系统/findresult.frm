VERSION 5.00
Begin VB.Form findresult 
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   7005
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "选择查询条件"
      Height          =   2055
      Left            =   600
      TabIndex        =   3
      Top             =   480
      Width           =   2535
      Begin VB.OptionButton optid 
         Caption         =   "按学号"
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton optname 
         Caption         =   "按姓名"
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   1200
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "返回"
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox txtfind 
      Height          =   615
      Left            =   3600
      TabIndex        =   0
      Top             =   1200
      Width           =   2535
   End
End
Attribute VB_Name = "findresult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub
