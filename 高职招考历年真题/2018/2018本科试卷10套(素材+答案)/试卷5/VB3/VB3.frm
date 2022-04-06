VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "计算个税"
      Height          =   615
      Left            =   1680
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "输入工资"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    x = Val(Text1.Text)
    If x <= 5000 Then
        s = 0
    ElseIf x <= 10000 Then
        s = (x - 5000) * 0.05
    Else
        s = (x - 10000) * 0.1 + 250
    End If
    MsgBox "应缴税：" & s & "元"
End Sub
