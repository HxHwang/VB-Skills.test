VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "分支函数计算"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "楷体_GB2312"
      Size            =   15.75
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   4335
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   720
      TabIndex        =   5
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   720
      TabIndex        =   4
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      Height          =   615
      Left            =   2520
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Y"
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "X"
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   180
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim x, y As Single
    x = Val(Text1.Text)
    If x > 0 Then
        y = 1
    ElseIf x = 0 Then
        y = 0
    Else
        y = -1
    End If
    Text2.Text = Str$(y)

End Sub

Private Sub Command2_Click()
    Text1.Text = ""
    Text2.Text = ""
End Sub
