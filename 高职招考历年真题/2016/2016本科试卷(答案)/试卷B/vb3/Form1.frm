VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   9630
   ClientTop       =   5325
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "倒计时"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1.Enabled = False
End Sub

Private Sub Command1_Click()
Timer1.Enabled = True
Timer1.Interval = 1000
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Label1.Caption - 1
If Label1.Caption = 0 Then
Timer1.Enabled = False
MsgBox "时间到", , "工程1"
End If
End Sub
