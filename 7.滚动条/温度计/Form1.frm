VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   7380
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   240
      Top             =   240
   End
   Begin VB.CommandButton Command3 
      Caption         =   "停止"
      Height          =   495
      Left            =   5760
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "降温"
      Height          =   495
      Left            =   5760
      TabIndex        =   2
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "加热"
      Height          =   495
      Left            =   5760
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.VScrollBar vbl1 
      Height          =   3975
      Left            =   4920
      Max             =   100
      TabIndex        =   0
      Top             =   240
      Value           =   100
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "100℃"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   135
      Left            =   4320
      TabIndex        =   8
      Top             =   360
      Width           =   15
   End
   Begin VB.Label Label4 
      Caption         =   "0℃"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "当前水温："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Sub Command1_Click()
i = 1
Label2.Caption = ""
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
i = 2
Label2.Caption = ""
Timer1.Enabled = True
End Sub

Private Sub Command3_Click()
Timer1.Enabled = False
End Sub

Private Sub Form_Load()
Label1.Caption = "当前水温："
Command2.Enabled = False
End Sub

Private Sub Timer1_Timer()
If i = 1 Then
vbl1.Value = vbl1.Value - 1
Command2.Enabled = True
ElseIf i = 2 Then
vbl1.Value = vbl1.Value + 1
Command1.Enabled = True
End If
Label3.Caption = Str$(-(vbl1.Value - 100)) & "℃"
If vbl1.Value = 0 Then
Label2.Caption = "水开了！"
Command1.Enabled = False
Command2.Enabled = True
i = 3
ElseIf vbl1.Value = 100 Then
Label2.Caption = "水结冰了！"
Command1.Enabled = True
Command2.Enabled = False
i = 3
End If

End Sub

