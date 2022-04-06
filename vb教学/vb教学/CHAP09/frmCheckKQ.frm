VERSION 5.00
Begin VB.Form frmCheckKQ 
   Caption         =   "查询学生出勤信息"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   5760
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox Timechecked 
      Caption         =   "时间"
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
      Left            =   600
      TabIndex        =   15
      Top             =   960
      Width           =   1695
   End
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
      Left            =   3480
      TabIndex        =   14
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Default         =   -1  'True
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
      Left            =   720
      TabIndex        =   13
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   5055
      Begin VB.ComboBox toMonth 
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
         Left            =   2880
         TabIndex        =   11
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox toYear 
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
         Left            =   840
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox FromMonth 
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
         Left            =   2880
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox fromYear 
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
         Left            =   840
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "月"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   12
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "年"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   10
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "到"
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
         TabIndex        =   8
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "月"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   7
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "年"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "从"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.TextBox StuffID 
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
      Left            =   2520
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.CheckBox IDchecked 
      Caption         =   "学生编号"
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
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmCheckKQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private querystring As String                       '保存查询出勤SQL语句
Private queryleave As String                        '保存查询请假SQL语句
Private queryovertime As String
Private queryerrand As String
Private fromtime As String
Private totime As String

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub setQuerystring()
    'Dim fromtime As String
    'Dim totime As String
    fromtime = Me.fromYear & "-" & Me.FromMonth & "-1"
    totime = Me.toYear & "-" & Me.toMonth & "-1"
    
    'MsgBox fromtime
    'MsgBox totime
    
    If Me.IDchecked.Value = vbChecked And Me.Timechecked.Value = vbChecked Then
        querystring = "select * from AttendanceInfo where AStuffID='" & Me.StuffID & "'"
        querystring = querystring & " and ADate between #" & fromtime & "# and #" & totime & "#"
        querystring = querystring & " order by ID"
        
        queryleave = "select * from LeaveInfo where LStuffID='" & Me.StuffID & "'"
        queryleave = queryleave & " and LFromDay between #" & fromtime & "# and #" & totime & "#"
        queryleave = queryleave & " order by LID"
        
        queryovertime = "select * from OvertimeInfo where OStuffID='" & Me.StuffID & "'"
        queryovertime = queryovertime & " and OFromDay between #" & fromtime & "# and #" & totime & "#"
        queryovertime = queryovertime & " order by OID"
        
        queryerrand = "select * from ErrandInfo where EStuffID='" & Me.StuffID & "'"
        queryerrand = queryerrand & " and EFromday between #" & fromtime & "# and #" & totime & "#"
        queryerrand = queryerrand & " order by EID"
        
    ElseIf Me.Timechecked.Value = vbChecked Then
        querystring = "select * from AttendanceInfo where ADate between #" & fromtime
        querystring = querystring & "# and #" & totime & "# order by AStuffID"
        
        queryleave = "select * from LeaveInfo where LFromDay between #" & fromtime
        queryleave = queryleave & "# and #" & totime & "# order by LStuffID"
        
        queryovertime = "select * from OvertimeInfo where OFromDay between #" & fromtime
        queryovertime = queryovertime & "# and #" & totime & "# order by OStuffID"
        
        queryerrand = "select * from ErrandInfo where EFromday between #" & fromtime
        queryerrand = queryerrand & "# and #" & totime & "# order by EStuffID"
    ElseIf Me.IDchecked.Value = vbChecked Then
        querystring = "select * from AttendanceInfo where AStuffID='" & Me.StuffID & "'"
        querystring = querystring & " order by ID"
        
        queryleave = "select * from LeaveInfo where LStuffID='" & Me.StuffID & "'"
        queryleave = queryleave & " order by LID"
        
        queryovertime = "select * from OvertimeInfo where OStuffID='" & Me.StuffID & "'"
        queryovertime = queryovertime & " order by OID"
        
        queryerrand = "select * from ErrandInfo where EStuffID='" & Me.StuffID & "'"
        queryerrand = queryerrand & " order by EID"
    Else
        querystring = "select * from AttendanceInfo order by ID"
        
        queryleave = "select * from LeaveInfo order by LID"
        
        queryovertime = "select * from OvertimeInfo order by OID"
        
        queryerrand = "select * from ErrandInfo order by EID"
    End If
End Sub
Private Sub CombineDate()
    fromtime = Me.fromYear.Text & "-" & Me.FromMonth.Text & "-1"
    fromtime = Format(Me.fromYear.Text & "-" & Me.FromMonth.Text & "-1", "yyyy-mm-dd")
    totime = Me.toYear.Text & "-" & Me.toMonth.Text & "-1"
    totime = Format(totime, "yyyy-mm-dd")
End Sub
Private Sub cmdOK_Click()
    
    If Trim(Me.StuffID) = "" And Timechecked.Value <> vbChecked Then
        MsgBox "请选择查询的条件！", vbOKOnly + vbExclamation, "警告！"
    Else
    Call CombineDate
    Call setQuerystring
    Call frmkqcheckresult.ATopic
    Call frmkqcheckresult.ShowAResult(querystring)
    Call frmkqcheckresult.LTopic
    Call frmkqcheckresult.ShowLResult(queryleave)
    Call frmkqcheckresult.OTopic
    Call frmkqcheckresult.ShowOResult(queryovertime)
    Call frmkqcheckresult.ETopic
    Call frmkqcheckresult.ShowEReslut(queryerrand)
    frmkqcheckresult.Show
    frmkqcheckresult.ZOrder 0
    Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    SQL = "select distinct ADate from AttendanceInfo"
    Set rs = TransactSQL(SQL)
    If Not rs.EOF Then
        rs.MoveFirst
        While Not rs.EOF
            If Not IsNull(rs.Fields(0)) Then            '设置年
                Me.fromYear.AddItem Left(rs(0), 4)
                Me.toYear.AddItem Left(rs(0), 4)
            End If
            rs.MoveNext
        Wend
        rs.Close
        Me.fromYear.ListIndex = 0
        Me.toYear.ListIndex = 0
    End If
    For i = 1 To 12                                     '设置月
        Me.FromMonth.AddItem i
        Me.toMonth.AddItem i
    Next i
        Me.FromMonth.ListIndex = 0
        Me.toMonth.ListIndex = 0
End Sub
