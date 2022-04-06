VERSION 5.00
Begin VB.Form frmStuff_info 
   Caption         =   "学生基本信息"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9735
   ControlBox      =   0   'False
   LinkMode        =   1  'Source
   LinkTopic       =   "FrmMain"
   ScaleHeight     =   8475
   ScaleWidth      =   9735
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   40
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   39
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "备注信息"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   20
      Top             =   7080
      Width           =   9255
      Begin VB.TextBox Remark 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Width           =   8895
      End
   End
   Begin VB.Frame workinfo 
      Caption         =   "个人工作信息"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   360
      TabIndex        =   14
      Top             =   4800
      Width           =   9255
      Begin VB.TextBox Position 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6480
         TabIndex        =   37
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox PayTime 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6480
         TabIndex        =   36
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox Dept 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   35
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox InTime 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   34
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox WorkTime 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   33
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "班级职务："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   19
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "正式上课时间："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   18
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label16 
         Caption         =   "所在班级："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   17
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "进入本校时间："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label14 
         Caption         =   "入学时间："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   15
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "学生基本信息"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   9255
      Begin VB.TextBox Email 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6000
         TabIndex        =   32
         Top             =   3480
         Width           =   2055
      End
      Begin VB.TextBox Tel 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6000
         TabIndex        =   31
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox Code 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6000
         TabIndex        =   30
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox Address 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6000
         TabIndex        =   29
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox Speciality 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6000
         TabIndex        =   28
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox Degree 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6000
         TabIndex        =   27
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox Birthday 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   26
         Top             =   3480
         Width           =   2055
      End
      Begin VB.TextBox Age 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   25
         Top             =   2880
         Width           =   2055
      End
      Begin VB.ComboBox Gender 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         TabIndex        =   23
         Text            =   "Combo1"
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox Place 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   24
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox StuffName 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   22
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox ID 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   21
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "邮政编码："
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
         Left            =   4440
         TabIndex        =   13
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Email:"
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
         Left            =   4320
         TabIndex        =   12
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "电    话："
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
         Left            =   4440
         TabIndex        =   11
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "家庭住址："
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
         Left            =   4440
         TabIndex        =   10
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "专    业："
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
         Left            =   4440
         TabIndex        =   9
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "年    级："
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
         Left            =   4440
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "籍    贯："
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
         Left            =   360
         TabIndex        =   7
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "年    龄："
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
         Left            =   360
         TabIndex        =   6
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "出生日期："
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
         Left            =   360
         TabIndex        =   5
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "性    别："
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
         Left            =   360
         TabIndex        =   4
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "学生姓名："
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
         Left            =   360
         TabIndex        =   3
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "学生编号："
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
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      Caption         =   "学生基本信息"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "frmStuff_info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub addNewRecord()
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    SQL = "select * from StuffInfo"
        Set rs = TransactSQL(SQL)
        rs.AddNew                                 '添加新记录
            rs.Fields(0) = Trim(Me.ID)
            rs.Fields(1) = Trim(Me.StuffName)
            rs.Fields(2) = Gender.Text
            rs.Fields(3) = Trim(Me.Place)
            rs.Fields(4) = Trim(Me.Age)
            rs.Fields(5) = Trim(Me.Birthday)
            rs.Fields(6) = Trim(Me.Degree)
            rs.Fields(7) = Trim(Me.Speciality)
            rs.Fields(8) = Trim(Me.Address)
            rs.Fields(9) = Trim(Me.Code)
            rs.Fields(10) = Trim(Me.Tel)
            rs.Fields(11) = Trim(Me.Email)
            rs.Fields(12) = Trim(Me.WorkTime)
            rs.Fields(13) = Trim(Me.InTime)
            rs.Fields(14) = Trim(Me.Dept)
            rs.Fields(15) = Trim(Me.PayTime)
            rs.Fields(16) = Trim(Me.Position)
            rs.Fields(17) = Trim(Me.Remark)
        rs.Update
        rs.Close
End Sub

Private Sub init()                               '初始化
        Me.StuffName = ""
        Me.Gender.ListIndex = 0
        Me.Place = ""
        Me.Age = ""
        Me.Birthday = ""
        Me.Degree = ""
        Me.Speciality = ""
        Me.Address = ""
        Me.Code = ""
        Me.Tel = ""
        Me.Email = ""
        Me.WorkTime = ""
        Me.InTime = ""
        Me.Dept = ""
        Me.PayTime = ""
        Me.Position = ""
        Me.Remark = ""
        Me.StuffName.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim SQL As String
    Dim temp As String
    Dim num As Integer
    Dim rs As New ADODB.Recordset
    If Trim(Me.StuffName) = "" Then                 '判断学生姓名是否为空
        MsgBox "请输入学生姓名！", vbOKOnly + vbExclamation, "警告！"
        Me.StuffName.SetFocus
        Exit Sub
    End If
    If Trim(Me.Age) = "" Then                       '判断年龄是否为空
        MsgBox "请输入学生年龄！", vbOKOnly + vbExclamation, "警告！"
        Me.Age.SetFocus
        Exit Sub
    End If
    If Trim(Me.Birthday) = "" Then                  '判断生日是否为空
        MsgBox "请输入学生生日！", vbOKOnly + vbExclamation, "警告！"
        Me.Birthday.SetFocus
        Exit Sub
    End If
    If Trim(Me.Dept) = "" Then                      '判断班级是否为空
        MsgBox "请输入学生所在班级！", vbOKOnly + vbExclamation, "警告！"
        Me.Dept.SetFocus
        Exit Sub
    End If
    If Trim(Me.Position) = "" Then                  '判断职务是否为空
        MsgBox "请输入学生职务！", vbOKOnly + vbExclamation, "警告！"
        Me.Position.SetFocus
    Exit Sub
    End If
    If Not IsDate(Me.Birthday) Then                 '判断生日的格式
        MsgBox "生日请按照(yyyy-mm-dd)方式输入！", vbOKOnly + vbExclamation, "警告！"
        Me.Birthday.SetFocus
        Exit Sub
        Else
        Me.Birthday = Format(Me.Birthday, "yyyy-mm-dd")
        End If
    If Not IsDate(Me.WorkTime) Then                 '判断入学时间的格式
        MsgBox "入学时间请按照(yyyy-mm-dd)方式输入！", vbOKOnly + vbExclamation, "警告！"
        Me.WorkTime.SetFocus
        Exit Sub
    Else
        Me.WorkTime = Format(Me.WorkTime, "yyyy-mm-dd")
    End If
    If Not IsDate(Me.InTime) Then                  '判断进入本校时间格式
        MsgBox "进入本校时间请按照(yyyy-mm-dd)方式输入！", vbOKOnly + vbExclamation, "警告！"
        Me.InTime.SetFocus
        Exit Sub
    Else
        Me.InTime = Format(Me.InTime, "yyyy-mm-dd")
    End If
    If Not IsDate(Me.PayTime) Then                 '正式上课时间格式
        MsgBox "正式上课时间请按照(yyyy-mm-dd)方式输入！", vbOKOnly + vbExclamation, "警告！"
        Me.PayTime.SetFocus
        Exit Sub
    Else
        Me.PayTime = Format(Me.PayTime, "yyyy-mm-dd")
    End If
    If flag = 1 Then                               '添加操作
        SQL = "select * from StuffInfo where SName='" & Trim(Me.StuffName)
        SQL = SQL & "' and SGender='" & Gender.Text & "' and SBirthday='"
        SQL = SQL & Trim(Me.Birthday) & "' and SDept='" & Trim(Me.Dept)
        SQL = SQL & "' and SPosition='" & Trim(Me.Position) & "'"
        Set rs = TransactSQL(SQL)
        If rs.EOF = False Then                     '判断是否已经存在学生记录
             MsgBox "已经存在这个学生的记录！", vbOKOnly + vbExclamation, "警告！"
             Me.StuffName.SetFocus
             Me.StuffName.SelStart = 0
             rs.Close
        Else
        Call addNewRecord
        MsgBox "记录已经成功添加！", vbOKOnly + vbExclamation, "添加结果！"
        SQL = "update PersonNum set Num= Num+1"       '计数器加1
        TransactSQL (SQL)
        SQL = "select * from PersonNum"               '学生编号初始化
        Set rs = TransactSQL(SQL)
        num = rs(0)
        num = num + 1
        temp = Right(Format(100000000 + num), 7)
        Me.ID = "P" & temp
        rs.Close
        Call init
        SQL = "select * from StuffInfo"          '显示信息列表
        frmResult.createList (SQL)
        frmResult.Show
        frmResult.ZOrder 0
        Me.ZOrder 0                              '显示窗体继续添加
        End If
    ElseIf flag = 2 Then                         '修改操作
        SQL = "update StuffInfo set SGender='" & Gender.Text & "',SPlace='"
        SQL = SQL & Trim(Me.Place) & "', SAge=" & Trim(Me.Age)
        SQL = SQL & ",SBirthday='" & Trim(Me.Birthday) & "',"
        SQL = SQL & "SDegree='" & Trim(Me.Degree) & "',"
        SQL = SQL & "SSpecial='" & Trim(Me.Speciality) & "',"
        SQL = SQL & "SAddress='" & Trim(Me.Address) & "',"
        SQL = SQL & "SCode='" & Trim(Me.Code) & "',"
        SQL = SQL & "STel='" & Trim(Me.Tel) & "',SEmail='" & Trim(Me.Email) & "',"
        SQL = SQL & "SWorkTime='" & Trim(Me.WorkTime) & "',"
        SQL = SQL & "SInTime='" & Trim(Me.InTime) & "',"
        SQL = SQL & "SDept='" & Trim(Me.Dept) & "',SPayTime='" & Trim(Me.PayTime)
        SQL = SQL & "',SPosition='" & Trim(Me.Position) & "',"
        SQL = SQL & "SRemark='" & Trim(Me.Remark) & "' where SID='" & Trim(Me.ID) & "'"
        TransactSQL (SQL)
        MsgBox "记录已经成功修改！", vbOKOnly + vbExclamation, "修改结果！"
        Unload Me
        SQL = "select * from StuffInfo"
        frmResult.createList (SQL)
        frmResult.Show
    End If
End Sub

Private Sub Form_Load()
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim num As Integer
    Dim temp As String
    With Gender                                '添加性别选项
        .AddItem "男"
        .AddItem "女"
    End With
    If flag = 1 Then                           '判断为添加信息
        Me.Caption = "添加" + Me.Caption
        Gender.ListIndex = 0
        SQL = "select * from PersonNum"
        Set rs = TransactSQL(SQL)
        num = rs(0)
        num = num + 1
        temp = Right(Format(10000000 + num), 7)
        Me.ID = "P" & temp
        rs.Close
    ElseIf flag = 2 Then                      '判断为修改信息
        Set rs = TransactSQL(gSQL)
        If rs.EOF = False Then
        With rs
            Me.ID = rs(0)
            Me.StuffName = rs(1)
            Me.Gender = rs(2)
            Me.Place = rs(3)
            Me.Age = rs(4)
            Me.Birthday = rs(5)
            Me.Degree = rs(6)
            Me.Speciality = rs(7)
            Me.Address = rs(8)
            Me.Code = rs(9)
            Me.Tel = rs(10)
            Me.Email = rs(11)
            Me.WorkTime = rs(12)
            Me.InTime = rs(13)
            Me.Dept = rs(14)
            Me.PayTime = rs(15)
            Me.Position = rs(16)
            Me.Remark = rs(17)
        End With
        rs.Close
        Me.Caption = "修改" & Me.Caption
        Me.ID.Enabled = False
        Me.StuffName.Enabled = False
        Else
            MsgBox "目前没有学生！", vbOKOnly + vbExclamation, "警告！"
        End If
    End If
End Sub

