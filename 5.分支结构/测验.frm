VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7425
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   42
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   7425
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "评分"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   3120
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2880
      Width           =   5415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5280
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "出题"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   840
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   4200
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Height          =   735
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n1, n2, ok, err

Private Sub Command1_Click()
n1 = Int(10 * Rnd + 1)
n2 = Int(10 * Rnd + 1)
Label1.Caption = n1 & "+" & n2 & "="
Text1.SetFocus
Command1.Caption = "下一题"
End Sub




Private Sub Command2_Click()
Text2.Text = Text2.Text & "共有" & (ok + err) & "题" & "得" & Int(ok / (ok + err) * 100) & "分"
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
Randomize
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Text2.Text = Text2.Text & Label1.Caption & Text1.Text
If Val(Text1.Text) = n1 + n2 Then
 Text2.Text = Text2.Text & "  正确  " & vbCrLf
 ok = ok + 1
 Else
 Text2.Text = Text2.Text & "  错误  " & vbCrLf
 err = err + 1
 End If
 Text1.Text = ""
 End If
End Sub
