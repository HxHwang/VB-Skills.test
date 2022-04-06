VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "闰年"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
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
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   2280
      TabIndex        =   5
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   2280
      TabIndex        =   4
      Top             =   240
      Width           =   1935
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
      Caption         =   "是否为闰年"
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   1725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "年份"
      Height          =   315
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   690
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim x As Integer
    x = Val(Text1.Text)
    If (x Mod 100) Then     '如果X不能被100整除
        If (x Mod 4 = 0) Then   '如果x能被4整除但不能被100整除
            Text2.Text = "yes"
        Else    '如果x不能被4和100整除
            Text2.Text = "no"
        End If
    ElseIf (x Mod 400 = 0) Then '如果x能被100整除，又能被400整除
        Text2.Text = "yes"
    Else
        Text2.Text = "no"
    End If


End Sub

Private Sub Command2_Click()
    Text1.Text = ""
    Text2.Text = ""
End Sub
