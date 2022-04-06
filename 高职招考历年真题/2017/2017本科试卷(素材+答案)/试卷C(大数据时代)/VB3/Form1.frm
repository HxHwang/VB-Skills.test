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
   Begin VB.CommandButton Command2 
      Caption         =   "停止"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   1680
   End
   Begin VB.Label Label1 
      Caption         =   "考试成功！"
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   720
      Width           =   975
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
Timer1.Interval = 10
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
End Sub

Private Sub Form_Load()
Label1.Left = 0
t = 0
End Sub

Private Sub Timer1_Timer()
If t = 0 Then
    Label1.Left = Label1.Left + 10
    If Label1.Left >= Form1.Width - Label1.Width Then
        t = 1
    End If
Else
    Label1.Left = Label1.Left - 10
    If Label1.Left <= 0 Then
        t = 0
    End If
End If
End Sub
'遇到这题 必须定义一个判断来回的一个标志t
