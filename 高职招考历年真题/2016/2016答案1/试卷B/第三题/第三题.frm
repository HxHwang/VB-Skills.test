VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5085
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   10980
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2280
      Top             =   2640
   End
   Begin VB.CommandButton Command1 
      Caption         =   "倒计时"
      Height          =   495
      Left            =   4920
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "10"
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer
Private Sub Command1_Click()
n = 10
Timer1.Interval = 1000
Label1.Caption = n
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
n = n - 1
Label1 = n
If n = 0 Then
    Timer1.Enabled = False
    MsgBox ("时间到")
End If
End Sub
