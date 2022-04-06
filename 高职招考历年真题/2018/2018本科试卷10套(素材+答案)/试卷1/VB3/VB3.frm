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
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算利息"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "存款年限"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "存款金额"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    x = Val(Text1.Text)
    n = Val(Text2.Text)
    If n <= 3 Then
        s = x * 0.03 * n
    ElseIf n <= 5 Then
        s = x * 0.05 * n
    Else
        s = x * 0.07 * n
    End If
    MsgBox "利息：" & s
End Sub
