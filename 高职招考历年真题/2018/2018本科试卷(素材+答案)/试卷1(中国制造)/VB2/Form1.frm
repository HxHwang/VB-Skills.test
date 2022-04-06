VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2400
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "显示菜单信息"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   1500
      ItemData        =   "Form1.frx":0000
      Left            =   480
      List            =   "Form1.frx":0016
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "菜单列表"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If List1.ListIndex > -1 Then
        Text1.Text = "你选中的是第 " & (List1.ListIndex + 1)
        Text1.Text = Text1.Text & " 项 菜单名：" & List1.Text
    End If
End Sub
