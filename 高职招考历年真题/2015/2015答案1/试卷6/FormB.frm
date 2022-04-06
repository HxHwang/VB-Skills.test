VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "QQ登录"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "结束"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "登录"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "QQ密码"
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "QQ账号"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
If Text1.Text = "" Then
    MsgBox "QQ账号不能为空", 64, "确认"
End If
If Len(Text1.Text) < 6 Then
    MsgBox "QQ密码长度必须六位数以上", 64, "确认"
End If
End Sub

Private Sub Command2_Click()
End
End Sub
