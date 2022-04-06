VERSION 5.00
Begin VB.Form frmAlteration 
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   8955
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   19
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   18
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "备注"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1575
      Left            =   720
      TabIndex        =   16
      Top             =   3840
      Width           =   8055
      Begin VB.TextBox ARemark 
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   7815
      End
   End
   Begin VB.TextBox AInTime 
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
      Left            =   6360
      TabIndex        =   15
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox AOutTime 
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
      Left            =   2160
      TabIndex        =   14
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox ANewPosition 
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
      Left            =   6360
      TabIndex        =   13
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox AOldPosition 
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
      Left            =   2160
      TabIndex        =   10
      Top             =   2280
      Width           =   2295
   End
   Begin VB.ComboBox ANewDept 
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
      Left            =   6360
      TabIndex        =   7
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox AOldDept 
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
      Left            =   2160
      TabIndex        =   5
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox AName 
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
      Left            =   6360
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
   Begin VB.ComboBox AID 
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
      Left            =   2160
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label8 
      Caption         =   "调入时间："
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "调出时间："
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "新职务："
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "原职务："
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "新班级名称："
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "原班级名称："
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "frmAlteration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public str1 As String                         '保存修改时的SQL语句
Public ID As Integer                              '保存记录编号
Private baddflag As Boolean

Private Sub AID_KeyDown(KeyCode As Integer, Shift As Integer)
    TabToEnter KeyCode
End Sub

Private Sub AID_LostFocus()
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    SQL = "select SName,SDept,SPosition from StuffInfo where SID='" & Me.AID.Text & "'"
    Set rs = TransactSQL(SQL)
    If rs.EOF = False Then
        Me.AName = rs(0)                           '初始化员工姓名
        Me.AOldDept = rs(1)
        Me.AOldPosition = rs(2)
   Else
        MsgBox "学生编号输入错误，或者没有这个学生！", vbOKOnly + vbExclamation, "警告！"
        Me.AID = ""
        Me.AID.SetFocus
        Me.AID.ListIndex = 0
    End If
    rs.Close
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub checkinput()
    If Me.ANewPosition = "" Then
            MsgBox "请输入新的职务！", vbOKOnly + vbExclamation, "警告！"
            Me.ANewPosition.SetFocus
        ElseIf Me.AOutTime = "" Or IsDate(Me.AOutTime) = False Then
            MsgBox "请输入正确的调出时间！", vbOKOnly + vbExclamation, "警告！"
            Me.AOutTime = ""
            Me.AOutTime.SetFocus
        ElseIf Me.AInTime = "" Or IsDate(Me.AInTime) = False Then
            MsgBox "请输入正确的调入时间！", vbOKOnly + vbExclamation, "警告！"
            Me.AInTime = ""
            Me.AInTime.SetFocus
        Else
            baddflag = True
    End If
End Sub

Private Sub cmdOK_Click()
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    baddflag = False
    Call checkinput
    If baddflag = True Then
    If flag = 1 Then                                    '添加记录
        'Call checkinput
        SQL = "select * from AlterationInfo"
        Set rs = TransactSQL(SQL)
        rs.AddNew
        rs.Fields(1) = Me.AID
        rs.Fields(2) = Me.AName
        rs.Fields(3) = Me.AOldDept
        rs.Fields(4) = Me.ANewDept
        rs.Fields(5) = Me.AOldPosition
        rs.Fields(6) = Me.ANewPosition
        rs.Fields(7) = Me.AOutTime
        rs.Fields(8) = Me.AInTime
        rs.Fields(9) = Me.ARemark
        rs.Update
        rs.Close
        SQL = "update StuffInfo set SDept='" & Me.ANewDept & "', SPosition='"
        SQL = SQL & Me.ANewPosition & "' where SID='" & Me.AID & "'"
        TransactSQL (SQL)
        MsgBox "已经添加调动信息！", vbOKOnly + vbExclamation, "添加结果！"
        SQL = "select * from AlterationInfo order by ID"
        frmAlterationResult.Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\Person.mdb"
        frmAlterationResult.Adodc1.RecordSource = SQL
        If SQL <> "" Then
            frmAlterationResult.Adodc1.Refresh
        End If
        Set frmAlterationResult.DataGrid1.DataSource = frmAlterationResult.Adodc1.Recordset
        frmAlterationResult.DataGrid1.Refresh
        frmAlterationResult.Show
        frmAlterationResult.ZOrder 0
        Call init
        Me.ZOrder 0
    
    Else                                                 '修改记录
        'Call checkinput
        SQL = "update StuffInfo set SDept='" & Me.ANewDept & "', SPosition='"
        SQL = SQL & Me.ANewPosition & "' where SID='" & Me.AID & "'"
        TransactSQL (SQL)
        SQL = "update AlterationInfo set AOldDept='" & Me.AOldDept & "',ANewDept='"
        SQL = SQL & Me.ANewDept & "',AOldPosition='" & Me.AOldPosition & "',ARemark='" & Me.ARemark
        SQL = SQL & "',ANewPosition='" & Me.ANewPosition & "',AOutTime=#" & Me.AOutTime
        
        
        SQL = SQL & "#,AInTime=#" & Me.AInTime & "# where ID=" & ID
        TransactSQL (SQL)
        MsgBox "已经修改信息！", vbOKOnly + vbExclamation, "修改结果！"
        Unload Me
        SQL = "select * from AlterationInfo order by ID"
        With frmAlterationResult.Adodc1                  '重新设置记录集
            .RecordSource = SQL
            .Refresh
        End With
        With frmAlterationResult.DataGrid1               '重新绑定记录集
            .ReBind
        End With
        frmAlterationResult.Show
        frmAlterationResult.ZOrder 0
        Unload frmAlterationResult
    frmAlterationResult.Show
    End If
    End If
End Sub

Private Sub Form_Load()
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim firstname As String
    If flag = 1 Then
        SQL = "select SID,SName,SDept,SPosition from StuffInfo order by SID"
        Set rs = TransactSQL(SQL)
        If rs.EOF = False Then
            rs.MoveFirst
            Me.AName = rs(1)
            Me.AOldDept = rs(2)
            Me.AOldPosition = rs(3)
            While Not rs.EOF
                Me.AID.AddItem rs(0)
                rs.MoveNext
            Wend
            rs.Close
            Me.AID.ListIndex = 0
        End If
        SQL = "select distinct SDept from StuffInfo"
        Set rs = TransactSQL(SQL)
        If rs.EOF = False Then
            rs.MoveFirst
            While Not rs.EOF
                Me.ANewDept.AddItem rs(0)
                rs.MoveNext
            Wend
            rs.Close
            Me.ANewDept.ListIndex = 0
        End If
        Me.AOutTime = Date
        Me.AInTime = Date
    End If
End Sub

Private Sub init()
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim firstname As String
    SQL = "select SID,SName,SDept,SPosition from StuffInfo order by SID"
    Set rs = TransactSQL(SQL)
    If rs.EOF = False Then
        rs.MoveFirst
        Me.AName = rs(1)
        Me.AOldDept = rs(2)
        Me.AOldPosition = rs(3)
        While Not rs.EOF
            Me.AID.AddItem rs(0)
            rs.MoveNext
        Wend
        rs.Close
        Me.AID.ListIndex = 0
    End If
    SQL = "select distinct SDept from StuffInfo"
    Set rs = TransactSQL(SQL)
    If rs.EOF = False Then
        rs.MoveFirst
        While Not rs.EOF
            Me.ANewDept.AddItem rs(0)
            rs.MoveNext
        Wend
        rs.Close
        Me.ANewDept.ListIndex = 0
    End If
    Me.AOutTime = Date
    Me.AInTime = Date
    Me.ANewPosition = ""
End Sub
