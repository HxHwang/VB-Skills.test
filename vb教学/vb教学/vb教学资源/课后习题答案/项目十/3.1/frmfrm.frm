VERSION 5.00
Begin VB.Form FrmEide 
   Caption         =   "学生成绩表"
   ClientHeight    =   6795
   ClientLeft      =   1110
   ClientTop       =   450
   ClientWidth     =   5340
   LinkTopic       =   "Form2"
   ScaleHeight     =   6795
   ScaleWidth      =   5340
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭(&C)"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   3840
      TabIndex        =   12
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "更新(&U)"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "刷新(&R)"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   3840
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除(&D)"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   2040
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "添加(&A)"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Connect         =   "Access"
      DatabaseName    =   "学生成绩数据库.mdb"
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "学生成绩表"
      Top             =   6300
      Width           =   5340
   End
   Begin VB.TextBox txtFields 
      DataField       =   "成绩"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   3
      Left            =   1920
      TabIndex        =   7
      Top             =   5400
      Width           =   2535
   End
   Begin VB.TextBox txtFields 
      DataField       =   "课程名"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   2
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   5
      Top             =   4280
      Width           =   2535
   End
   Begin VB.TextBox txtFields 
      DataField       =   "姓名"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   1
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   3
      Top             =   3160
      Width           =   2535
   End
   Begin VB.TextBox txtFields 
      DataField       =   "学号"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   0
      Left            =   1920
      TabIndex        =   1
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "成  绩:"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   360
      TabIndex        =   6
      Top             =   5640
      Width           =   1125
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "课程名:"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   367
      TabIndex        =   4
      Top             =   4520
      Width           =   1110
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "姓  名:"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   3400
      Width           =   1125
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "学  号:"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   2280
      Width           =   1125
   End
End
Attribute VB_Name = "FrmEide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
  Data1.Recordset.AddNew
End Sub

Private Sub cmdDelete_Click()
  '如果删除记录集的最后一条记录
  '记录或记录集中唯一的记录
  Data1.Recordset.Delete
  Data1.Recordset.MoveNext
End Sub

Private Sub cmdRefresh_Click()
  '这仅对多用户应用程序才是需要的
  Data1.Refresh
End Sub

Private Sub cmdUpdate_Click()
  Data1.UpdateRecord
  Data1.Recordset.Bookmark = Data1.Recordset.LastModified
End Sub

Private Sub cmdClose_Click()
 End
 
End Sub



Private Sub Data1_Error(DataErr As Integer, Response As Integer)
  '这就是放置错误处理代码的地方
  '如果想忽略错误，注释掉下一行代码
  '如果想捕捉错误，在这里添加错误处理代码
  MsgBox "数据错误事件命中错误：" & Error$(DataErr)
  Response = 0  '忽略错误
End Sub

Private Sub Data1_Reposition()
  Screen.MousePointer = vbDefault
  On Error Resume Next
  '这将显示当前记录位置
  '为动态集和快照
  Data1.Caption = "记录：" & (Data1.Recordset.AbsolutePosition + 1)
  '对于 Table 对象，当记录集创建后并使用下面的行时，
  '必须设置 Index 属性
  'Data1.Caption = "记录：" & (Data1.Recordset.RecordCount * (Data1.Recordset.PercentPosition * 0.01)) + 1
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)
  '这是放置验证代码的地方
  '当下面的动作发生时，调用这个事件
  Select Case Action
    Case vbDataActionMoveFirst
    Case vbDataActionMovePrevious
    Case vbDataActionMoveNext
    Case vbDataActionMoveLast
    Case vbDataActionAddNew
    Case vbDataActionUpdate
    Case vbDataActionDelete
    Case vbDataActionFind
    Case vbDataActionBookmark
    Case vbDataActionClose
  End Select
  Screen.MousePointer = vbHourglass
End Sub

