VERSION 5.00
Begin VB.Form FrmAttendance 
   Caption         =   "学生出勤信息"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9630
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   9630
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   8
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   7
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "学生出勤信息"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   480
      TabIndex        =   12
      Top             =   2760
      Width           =   8655
      Begin VB.Frame Frame3 
         Caption         =   "出入信息"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   8295
         Begin VB.TextBox InTime 
            BeginProperty Font 
               Name            =   "楷体_GB2312"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2160
            TabIndex        =   4
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox OutTime 
            BeginProperty Font 
               Name            =   "楷体_GB2312"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6120
            TabIndex        =   6
            Top             =   600
            Width           =   1815
         End
         Begin VB.OptionButton InFlag 
            Caption         =   "上学时间："
            BeginProperty Font 
               Name            =   "楷体_GB2312"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   240
            TabIndex        =   3
            Top             =   480
            Width           =   1815
         End
         Begin VB.OptionButton OutFlag 
            Caption         =   "下学时间："
            BeginProperty Font 
               Name            =   "楷体_GB2312"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   4320
            TabIndex        =   5
            Top             =   480
            Width           =   1695
         End
      End
      Begin VB.TextBox NowDate 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   14
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "当前日期："
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   13
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "学生个人信息"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   9
      Top             =   1320
      Width           =   8655
      Begin VB.ComboBox ASID 
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         TabIndex        =   1
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox ASName 
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5880
         TabIndex        =   2
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "学生姓名："
         BeginProperty Font 
            Name            =   "楷体_GB2312"
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
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "学生编号："
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Label topic 
      Alignment       =   2  'Center
      Caption         =   "学生上下学信息"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   5655
   End
End
Attribute VB_Name = "FrmAttendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ilate As Integer                                  '迟到次数
Private iearly As Integer                                 '早退次数
Private aflag As String                                   '出入标志
Private addflag As Boolean                                '添加标志
Private firstID As String                                 '第一个学生编号

Private Sub ASID_KeyDown(KeyCode As Integer, Shift As Integer)
    TabToEnter KeyCode
End Sub

Private Sub ASID_LostFocus()
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    SQL = "select SName from StuffInfo where SID='" & Me.ASID.Text & "'"
    Set rs = TransactSQL(SQL)
    If rs.EOF = False Then
        Me.ASName = rs(0)                           '初始化学生姓名
    Else
        MsgBox "学生编号输入错误，或者没有这个学生！", vbOKOnly + vbExclamation, "警告！"
        Me.ASID = ""
        Me.ASID.SetFocus
        Me.ASID.ListIndex = 0
    End If
    rs.Close
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub CheckRecord()                           '判断是否存在记录
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    SQL = "select * from AttendanceInfo where AStuffID='" & Me.ASID.Text & "'"
    SQL = SQL & " and AFlag='" & aflag & "' and ADate=#" & Me.NowDate & "#"
        Set rs = TransactSQL(SQL)
        If rs.EOF = False Then
            MsgBox "已经存在这条记录！", vbOKOnly + vbExclamation, "警告！"
            addflag = True
        Else
            addflag = False
        End If
        rs.Close
End Sub

Private Sub in_add()                                '添加上学记录
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    SQL = "select * from AttendanceInfo"
    Set rs = TransactSQL(SQL)
    rs.AddNew
    rs.Fields(1) = Me.ASID
    rs.Fields(2) = Me.ASName
    rs.Fields(3) = Me.NowDate
    rs.Fields(4) = aflag
    rs.Fields(5) = Me.InTime
    rs.Fields(7) = ilate
    rs.Update
    rs.Close
End Sub

Private Sub out_add()                               '添加放学记录
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    SQL = "select * from AttendanceInfo"
    Set rs = TransactSQL(SQL)
    rs.AddNew
    rs.Fields(1) = Me.ASID
    rs.Fields(2) = Me.ASName
    rs.Fields(3) = Me.NowDate
    rs.Fields(4) = aflag
    rs.Fields(6) = Me.OutTime
    rs.Fields(8) = iearly
    rs.Update
    rs.Close
End Sub

Private Sub cmdOK_Click()
    Dim SQL As String
    Dim sql2 As String
    Dim rs As New ADODB.Recordset
    Dim rsTime As New ADODB.Recordset
    sql2 = "select * from AttendanceInfo order by ID desc"
    SQL = "select * from TimeSetting"
    Set rsTime = TransactSQL(SQL)
    If flag = 1 Then
    ilate = 0
    iearly = 0
    If Me.InFlag = False And Me.OutFlag = False Then
        MsgBox "请选择上下学！", vbOKOnly + vbExclamation, "警告！"
    Else
    If Me.InFlag = True Then                         '添加上学记录
        aflag = "入"
        If Me.InTime = "" Or IsDate(Me.InTime) = False Then
            MsgBox "请输入正确的时间！", vbOKOnly + vbExclamation, "警告!"
            Me.InTime = ""
            Me.InTime.SetFocus
        Else
            If DateDiff("s", Me.InTime, rsTime(0)) < 0 Then
                ilate = 1
            End If
            Call CheckRecord
            If addflag = False Then
                Call in_add
                MsgBox "已经添加上学记录！", vbOKOnly + vbExclamation, "添加结果！"
                Call init
                Me.InFlag = False
            Else
                Call init
                Me.InFlag = False
            End If
        End If
    End If
    If Me.OutFlag = True Then                        '添加放学记录
        aflag = "出"
        If Me.OutTime = "" Or IsDate(Me.OutTime) = False Then
            MsgBox "请输入正确的时间！", vbOKOnly + vbExclamation, "警告!"
            Me.OutTime = ""
            Me.OutTime.SetFocus
        Else
            If DateDiff("s", Me.OutTime, rsTime(1)) > 0 Then
                iearly = 1
            End If
            Call CheckRecord
            If addflag = False Then
                Call out_add
                MsgBox "已经添加放学记录！", vbOKOnly + vbExclamation, "添加结果！"
                Call init
                Me.OutFlag = False
            Else
                Call init
                Me.OutFlag = False
            End If
        End If
    End If
    End If
        Call frmAResult.ListTopic
        Call frmAResult.ShowData(sql2)
        frmAResult.Show
        frmAResult.ZOrder 0
        Me.ZOrder 0
    Else                                             '修改记录
        If MsgBox("确定修改编号为" & Me.ASID & "的学生信息?", vbOKCancel, "提示！") _
                                                                = vbOK Then
            If Me.InFlag = True Then
                If DateDiff("s", Me.InTime, rsTime(0)) < 0 Then
                    ilate = 1
                End If
                SQL = "update AttendanceInfo set AInTime=#" & Me.InTime & "#,"
                SQL = SQL & "ALate=" & ilate & " where ID=" & ArecordID
                TransactSQL (SQL)                     '修改上学记录
                Call frmAResult.ListTopic
                Call frmAResult.ShowData(sql2)
                frmAResult.Show
                MsgBox "信息已经修改！", vbOKOnly + vbExclamation, "修改结果！"
                Unload Me
            ElseIf Me.OutFlag = True Then
                If DateDiff("s", Me.OutTime, rsTime(1)) > 0 Then
                    iearly = 1
                End If
                SQL = "update AttendanceInfo set AOutTime=#" & Me.OutTime & "#,"
                SQL = SQL & "AEarly=" & iearly & " where ID=" & ArecordID
                TransactSQL (SQL)                     '修改放学记录
                Call frmAResult.ListTopic
                Call frmAResult.ShowData(sql2)
                frmAResult.Show
                MsgBox "信息已经修改！", vbOKOnly + vbExclamation, "修改结果！"
                Unload Me
            End If
        Else
        Unload Me
        End If
    End If
    rsTime.Close
End Sub

Private Sub Form_Load()
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    If flag = 1 Then
    SQL = "select SID from StuffInfo order by SID"
    Set rs = TransactSQL(SQL)
    If rs.EOF = False Then
        rs.MoveFirst
        firstID = rs(0)
    While Not rs.EOF
        Me.ASID.AddItem rs(0)                     '初始化学生编号
        rs.MoveNext
    Wend
        rs.Close
    Else
        MsgBox "目前没有学生！", vbOKOnly + vbExclamation, "警告！"
    End If
    Me.NowDate = Date
    Me.ASID.ListIndex = 0
    SQL = "select SName from StuffInfo where SID='" & firstID & "'"
    Set rs = TransactSQL(SQL)
    Me.ASName = rs(0)                             '初始化学生姓名
    rs.Close
    Me.OutTime = ""
    Me.InTime = ""
   ElseIf flag = 2 Then
       
        Set rs = TransactSQL(kqsql)
         'If rs.EOF = False And rs.BOF Then

        If rs.EOF = False Then
         rs.MoveFirst
        firstID = rs(0)
        With rs
            Me.ASID = rs(1)
            Me.ASName = rs(2)
            Me.NowDate = rs(3)
            If IsNull(rs(5)) = True Then
            Me.InTime = ""
            Me.OutFlag = True
            Else
            Me.InTime = rs(5)
            End If
            If IsNull(rs(6)) = True Then
            Me.OutTime = ""
            Me.InFlag = True
            Else
            Me.OutTime = rs(6)
            End If
        End With
        rs.Close
        End If
    End If
    
End Sub
Private Sub init()                                '初始化
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    SQL = "select SName from StuffInfo where SID='" & firstID & "'"
    Set rs = TransactSQL(SQL)
    Me.ASID.ListIndex = 0
    Me.ASName = rs(0)
    Me.InTime = ""
    Me.OutTime = ""
End Sub


