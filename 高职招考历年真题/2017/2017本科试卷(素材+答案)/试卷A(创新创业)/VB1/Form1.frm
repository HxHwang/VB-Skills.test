VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "时间函数"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "系统时间"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim t As Date
t = Time()
Select Case Hour(t)
    Case Is < 8
        Label1.Caption = "凌晨" & Space(1) & t
    Case Is < 12
        Label1.Caption = "上午" & Space(1) & t
    Case Is < 17
        Label1.Caption = "下午" & Space(1) & t
    Case Else
        Label1.Caption = "晚上" & Space(1) & t
End Select
End Sub

Private Sub Command2_Click()
End
End Sub
