VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12075
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   12075
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5520
      Top             =   2280
   End
   Begin VB.CommandButton Command2 
      Caption         =   "停止"
      Height          =   495
      Left            =   6720
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "考试成功！"
      Height          =   180
      Left            =   4680
      TabIndex        =   0
      Top             =   1080
      Width           =   900
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim f As Integer
Private Sub Command1_Click()
Timer1.Enabled = True
End Sub
Private Sub Command2_Click()
Timer1.Enabled = False
End Sub
Private Sub Form_Load()
Label1.Left = 0
Timer1.Enabled = False
f = 0
End Sub
Private Sub Timer1_Timer()
If Label1.Left > Form1.Width - Label1.Width Then f = 1
If Label1.Left <= 0 Then f = 0
If f = 0 Then Label1.Left = Label1.Left + 10
If f = 1 Then Label1.Left = Label1.Left - 10
End Sub
