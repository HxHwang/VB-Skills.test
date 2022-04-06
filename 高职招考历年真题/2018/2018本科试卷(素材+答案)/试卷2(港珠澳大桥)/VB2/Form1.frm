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
   Begin VB.OptionButton Option3 
      Caption         =   "经济套餐38元"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      Caption         =   "麻辣套餐56元"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "全家套餐78元"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "点餐"
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   2520
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Option1.Value = True Then
        s = Option1.Caption
    ElseIf Option2.Value Then
        s = Option2.Caption
     ElseIf Option3.Value Then
        s = Option3.Caption
    End If
    Text1.Text = "您点的套餐是：" & s & vbCrLf & "订餐时间是：" & Now
End Sub
