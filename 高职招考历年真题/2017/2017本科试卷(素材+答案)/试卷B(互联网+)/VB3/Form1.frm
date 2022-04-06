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
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   1440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "倒计时"
      Height          =   855
      Left            =   1920
      TabIndex        =   1
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "10"
      Height          =   855
      Left            =   1920
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As Integer
Private Sub Command1_Click()
Timer1.Enabled = True
Timer1.Interval = 1000
t = 10
End Sub

Private Sub Timer1_Timer()
t = t - 1
Label1.Caption = t
If t = 0 Then
    MsgBox "时间到！"
    Timer1.Enabled = False
    Timer1.Interval = 0
End If
End Sub
