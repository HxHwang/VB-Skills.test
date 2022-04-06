VERSION 5.00
Begin VB.Form frmxt2 
   Caption         =   "操作练习题"
   ClientHeight    =   3090
   ClientLeft      =   450
   ClientTop       =   8025
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
      Left            =   7800
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
      Left            =   7800
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
Attribute VB_Name = "frmxt2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
xtjj2.Show
End Sub

Private Sub Command3_Click()
Me.Hide
form5.Show
xlapp.Visible = False '设置EXCEL对象不可见
End Sub

Private Sub Form_Load()
Label1.Caption = "体验2" & vbCrLf & "（1）按公式：利润=售出价-进价-经营成本，计算各种产品的利润和平均利润值，结果保留两位小数。" & vbCrLf & "（2）将sheet1中的平均值以外的数据全部复制到sheet2中，并将全部数据按利润值从高到低排序。" & vbCrLf & "（3）将sheet1中的平均值以外的数据全部复制到sheet3中，在数据中筛选出林润大于100元并小于150元的数据?筛选后恢复全部数据?" & vbCrLf & "（4）将sheet1中的平均值以外的数据全部复制到sheet4中，按类别分类汇总，分别求出空调、电冰箱、洗衣机的经营成本与利润的平均值?"
       
End Sub
