VERSION 5.00
Begin VB.Form frmxt1 
   Caption         =   "操作练习题"
   ClientHeight    =   3090
   ClientLeft      =   255
   ClientTop       =   8220
   ClientWidth     =   11040
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   ScaleHeight     =   3090
   ScaleWidth      =   11040
   Begin VB.CommandButton Command1 
      Caption         =   "视频讲解"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "返回"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "frmxt1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
xtjj1.Show
End Sub

Private Sub Command3_Click()
Me.Hide
form5.Show
xlapp.Visible = False '设置EXCEL对象不可见
End Sub

Private Sub Form_Load()
Label1.Caption = "体验1 " & vbCrLf & "（1）分别计算三个车间生产所有零件的合计数。" & vbCrLf & "（2）使用函数计算三个车间生产齿轮、齿轮箱、齿轮泵、轮轴、扇形齿轮各部件的总计数。"
End Sub
