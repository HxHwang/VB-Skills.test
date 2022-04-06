VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5475
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   5475
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4680
      Top             =   600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "定时"
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   2040
      TabIndex        =   3
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "分"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "时"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "闹钟时间："
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "现在时间："
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hour, minute
Private Sub Command1_Click()
hour = Format(Text1.Text, "00")
minute = Format(Text2.Text, "00")
End Sub

Private Sub Form_Load()
Label2.Caption = Time$

End Sub

Private Sub Timer1_Timer()
Label2.Caption = Time$
If Mid(Time$, 1, 8) = hour & ":" & minute & ":00" Then
   Beep
   Print "时间到！"
End If
End Sub
