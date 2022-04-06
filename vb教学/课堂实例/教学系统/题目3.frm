VERSION 5.00
Begin VB.Form frmxt3 
   Caption         =   "操作练习题"
   ClientHeight    =   3090
   ClientLeft      =   645
   ClientTop       =   7830
   ClientWidth     =   11190
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   ScaleHeight     =   3090
   ScaleWidth      =   11190
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
      Width           =   7335
   End
End
Attribute VB_Name = "frmxt3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
xtjj3.Show
End Sub

Private Sub Command3_Click()
Me.Hide
form5.Show
xlapp.Visible = False '设置EXCEL对象不可见
End Sub

Private Sub Form_Load()
Label1.Caption = "体验3" & vbCrLf & "（1）按公式：工资=基本工资+效益工资，计算每人的工资。" & vbCrLf & "（2）按公式：浮动额=工资*浮动率，计算每人的工资浮动额。" & vbCrLf & "（3） 更具'工资'和'浮动额'分别计算每人的工作总额。" & vbCrLf & "（4）计算机各工资项的平均值。"
      
End Sub
