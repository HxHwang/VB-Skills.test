VERSION 5.00
Begin VB.Form frmCheckAlter 
   Caption         =   "查询调动信息"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8595
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   8595
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox StuffID 
      Height          =   300
      Left            =   2640
      TabIndex        =   15
      Top             =   840
      Width           =   2655
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
      Left            =   840
      TabIndex        =   14
      Top             =   720
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "调出时间"
      Height          =   1815
      Left            =   480
      TabIndex        =   3
      Top             =   1920
      Width           =   5055
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
         TabIndex        =   7
         Top             =   360
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
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
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
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
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
         TabIndex        =   13
         Top             =   480
         Width           =   375
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
         TabIndex        =   12
         Top             =   480
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
         TabIndex        =   11
         Top             =   480
         Width           =   495
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
         TabIndex        =   10
         Top             =   1080
         Width           =   375
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
         TabIndex        =   9
         Top             =   1080
         Width           =   375
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
         TabIndex        =   8
         Top             =   1080
         Width           =   495
      End
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
      Left            =   960
      TabIndex        =   2
      Top             =   3960
      Width           =   1575
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
      Left            =   3720
      TabIndex        =   1
      Top             =   3960
      Width           =   1575
   End
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
      Left            =   840
      TabIndex        =   0
      Top             =   1320
      Width           =   1695
   End
End
Attribute VB_Name = "frmCheckAlter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strQuery As String
Private fromtime As String                           '开始时间
Private totime As String                             '结束时间

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub setstrQuery()
    fromtime = Me.fromYear & "-" & Me.FromMonth & "-1"
    totime = Me.toYear & "-" & Me.toMonth & "-1"
    If Me.IDchecked.Value = vbChecked And Me.Timechecked.Value = vbChecked Then
        strQuery = "select * from AlterationInfo where AID='" & Me.StuffID
        strQuery = strQuery & "' and AOutTime between #" & fromtime & "# and #"
        strQuery = strQuery & totime & "#"
        'MsgBox strQuery
    ElseIf Me.IDchecked.Value = vbChecked Then
        strQuery = "select * from AlterationInfo where AID='" & Me.StuffID & "' order by ID"
    ElseIf Me.Timechecked.Value = vbChecked Then
        strQuery = "select * from AlterationInfo where AOutTime between #" & fromtime
        strQuery = strQuery & "# and #" & totime & "# order by ID"
    Else
        strQuery = "select * from AlterationInfo order by ID"
    End If
End Sub

Private Sub cmdOK_Click()

    If Trim(Me.StuffID) = "" And Timechecked.Value <> vbChecked Then
        MsgBox "请选择查询的条件！", vbOKOnly + vbExclamation, "警告！"
    Else
    Call setstrQuery
    frmAlterationResult.Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\Person.mdb"
    frmAlterationResult.Adodc1.RecordSource = strQuery
    If strQuery <> "" Then
        frmAlterationResult.Adodc1.Refresh
    End If
    Set frmAlterationResult.DataGrid1.DataSource = frmAlterationResult.Adodc1.Recordset
    frmAlterationResult.DataGrid1.Refresh
    frmAlterationResult.Show
    frmAlterationResult.ZOrder 0
    Unload Me
    End If
End Sub

Private Sub Form_Load()
 Dim i As Integer
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    SQL = "select distinct AID from AlterationInfo order by AID"
    Set rs = TransactSQL(SQL)
    If rs.EOF = False Then
        rs.MoveFirst
        While Not rs.EOF
            Me.StuffID.AddItem rs(0)
            rs.MoveNext
        Wend
        Me.StuffID.ListIndex = 0
    End If
    rs.Close
    SQL = "select distinct AOutTime from AlterationInfo"
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

