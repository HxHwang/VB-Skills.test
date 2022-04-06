VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "实操练习"
   ClientHeight    =   8370
   ClientLeft      =   1830
   ClientTop       =   1425
   ClientWidth     =   10860
   LinkTopic       =   "Form5"
   Picture         =   "操作练习.frx":0000
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   5640
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   7695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "操作环境"
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9120
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9120
      Width           =   2535
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6990
      ItemData        =   "操作练习.frx":4C004
      Left            =   1680
      List            =   "操作练习.frx":4C011
      TabIndex        =   0
      Top             =   1440
      Width           =   3135
   End
End
Attribute VB_Name = "form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
Set xlapp = CreateObject("Excel.Application") '创建EXCEL对象
Select Case List1.ListIndex
       Case 0
        Set xlbook = xlapp.Workbooks.Open(App.Path & "\3.1.xls")
        xlapp.Visible = True '设置EXCEL对象可见
        Set xlsheet = xlbook.Worksheets(1)
        form5.Hide
        frmxt1.Show
       Case 1
        Set xlbook = xlapp.Workbooks.Open(App.Path & "\3.2.xls")
        xlapp.Visible = True '设置EXCEL对象可见
        Set xlsheet = xlbook.Worksheets(1)
        form5.Hide
        frmxt2.Show
       Case 2
        Set xlbook = xlapp.Workbooks.Open(App.Path & "\3.3.xls")
        xlapp.Visible = True '设置EXCEL对象可见
        Set xlsheet = xlbook.Worksheets(1)
        form5.Hide
        frmxt3.Show
End Select


End Sub

Private Sub Command3_Click()
Me.Hide
Form3.Show
End Sub

Private Sub Form_Load()
Command2.Enabled = False
Me.WindowState = 2
End Sub

Private Sub List1_Click()
Select Case List1.ListIndex
       Case 0
       Text1.Text = "体验1 " & vbCrLf & "（1）分别计算三个车间生产所有零件的合计数。" & vbCrLf & "（2）使用函数计算三个车间生产齿轮、齿轮箱、齿轮泵、轮轴、扇形齿轮各部件的总计数。"
        Command2.Enabled = True
       Case 1
        Command2.Enabled = True
         Text1.Text = "体验2" & vbCrLf & "（1）按公式：利润=售出价-进价-经营成本，计算各种产品的利润和平均利润值，结果保留两位小数。" & vbCrLf & "（2）将sheet1中的平均值以外的数据全部复制到sheet2中，并将全部数据按利润值从高到低排序。" & vbCrLf & "（3）将sheet1中的平均值以外的数据全部复制到sheet3中，在数据中筛选出林润大于100元并小于150元的数据?筛选后恢复全部数据?" & vbCrLf & "（4）将sheet1中的平均值以外的数据全部复制到sheet4中，按类别分类汇总，分别求出空调、电冰箱、洗衣机的经营成本与利润的平均值?"
       Case 2
         Text1.Text = "体验3" & vbCrLf & "（1）按公式：工资=基本工资+效益工资，计算每人的工资。" & vbCrLf & "（2）按公式：浮动额=工资*浮动率，计算每人的工资浮动额。" & vbCrLf & "（3）根据'工资'和'浮动额'分别计算每人的工作总额。" & vbCrLf & "（4）计算机各工资项的平均值。"
        Command2.Enabled = True
End Select
       
End Sub
