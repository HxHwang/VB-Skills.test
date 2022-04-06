VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "List列表拒绝添加重复信息"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   4065
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CmdEnd 
      Caption         =   "退出"
      Height          =   360
      Left            =   2190
      TabIndex        =   4
      Top             =   2535
      Width           =   1425
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "添加"
      Height          =   360
      Left            =   480
      TabIndex        =   3
      Top             =   2535
      Width           =   1425
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   1020
      TabIndex        =   1
      Top             =   120
      Width           =   2580
   End
   Begin VB.ListBox List1 
      Height          =   1860
      Left            =   1020
      TabIndex        =   0
      Top             =   525
      Width           =   2580
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "请输入编号"
      Height          =   330
      Left            =   75
      TabIndex        =   2
      Top             =   180
      Width           =   990
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  List1.AddItem "a01001"
  List1.AddItem "a01002"
End Sub

Private Sub CmdAdd_Click()
  Dim Myval As Long
  For i = 0 To List1.ListCount - 1
      List1.ListIndex = i
      If List1.Text = Text1.Text Then
         MsgBox "系统不允许重复输入，请重新输入"
         Exit Sub
      End If
  Next i
  List1.AddItem Text1.Text
End Sub

Private Sub CmdEnd_Click()
  End
End Sub


