VERSION 5.00
Begin VB.Form frmCheckStuff 
   Caption         =   "查询学生基本信息"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   6375
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
      Left            =   3360
      TabIndex        =   11
      Top             =   5880
      Width           =   1695
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
      TabIndex        =   10
      Top             =   5880
      Width           =   1695
   End
   Begin VB.TextBox SName 
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
      Left            =   2640
      TabIndex        =   4
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox SID 
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
      Left            =   2640
      MaxLength       =   8
      TabIndex        =   2
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   600
      TabIndex        =   12
      Top             =   4080
      Width           =   4695
      Begin VB.ComboBox ToMonth 
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
         Left            =   2520
         TabIndex        =   9
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox ToYear 
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
         Left            =   720
         TabIndex        =   8
         Top             =   960
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
         ItemData        =   "frmCheckStuff.frx":0000
         Left            =   2520
         List            =   "frmCheckStuff.frx":0002
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox FromYear 
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
         Left            =   720
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label7 
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
         Left            =   4080
         TabIndex        =   18
         Top             =   960
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
         Left            =   4080
         TabIndex        =   17
         Top             =   480
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
         Left            =   2040
         TabIndex        =   16
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label4 
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
         Left            =   2040
         TabIndex        =   15
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label2 
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
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.CheckBox TimeCheck 
      Caption         =   "进入本校时间："
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   5
      Top             =   3360
      Width           =   3015
   End
   Begin VB.CheckBox NameCheck 
      Caption         =   "学生姓名："
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
      Left            =   600
      TabIndex        =   3
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CheckBox IDCheck 
      Caption         =   "学生编号："
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label topic 
      Alignment       =   2  'Center
      Caption         =   "选择查询条件"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   3855
   End
End
Attribute VB_Name = "frmCheckStuff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private query As String                            '保存SQL语句
Private fromdate As String                         '起始时间
Private todate As String                           '结束时间

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub CombineDate()                          '获得起始和结束时间
    fromdate = Me.FromYear.Text & "-" & Me.FromMonth.Text & "-1"
    fromdate = Format(Me.FromYear.Text & "-" & Me.FromMonth.Text & "-1", "yyyy-mm-dd")
    todate = Me.ToYear.Text & "-" & Me.ToMonth.Text & "-1"
    todate = Format(todate, "yyyy-mm-dd")
End Sub

Private Sub setSQL()                               '设置SQL语句
    If IDCheck.Value = vbChecked Then
        query = "select * from StuffInfo where SID='" & Trim(Me.SID) & "'"
    End If
    If NameCheck.Value = vbChecked Then
        query = "select * from StuffInfo where SName='" & Trim(Me.SName) & "'"
    End If
    If TimeCheck.Value = vbChecked Then
        query = "select * from StuffInfo where SInTime between #"
        query = query & fromdate & "# and  #" & todate & "#"
    End If
    If IDCheck.Value = vbChecked And NameCheck.Value = vbChecked Then
        query = "select * from StuffInfo where SID=' " & Trim(Me.SID)
        query = query & "' and SName='" & Trim(Me.SName) & "'"
    End If
    If NameCheck.Value = vbChecked And TimeCheck.Value = vbChecked Then
        query = "select * from StuffInfo where SName='" & Trim(Me.SName)
        query = query & "' and SInTime between #" & fromdate
        query = query & "# and #" & todate & "#"
    End If
End Sub

Private Sub cmdOK_Click()
    If Trim(Me.SID) = "" And Trim(Me.SName) = "" And TimeCheck.Value <> vbChecked Then
        MsgBox "请选择查询的条件！", vbOKOnly + vbExclamation, "警告！"
    Else
    Call CombineDate
    Call setSQL
    frmResult.createList (query)
    frmResult.Show
    Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim sql As String
    Dim rs As New ADODB.Recordset
    sql = "select distinct SInTime from StuffInfo"
    Set rs = TransactSQL(sql)
    If Not rs.EOF Then
        rs.MoveFirst
        While Not rs.EOF
            If Not IsNull(rs.Fields(0)) Then            '设置年
                Me.FromYear.AddItem Left(rs(0), 4)
                Me.ToYear.AddItem Left(rs(0), 4)
            End If
            rs.MoveNext
        Wend
        rs.Close
        Me.FromYear.ListIndex = 0
        Me.ToYear.ListIndex = 0
    End If
    For i = 1 To 12                                     '设置月
        Me.FromMonth.AddItem i
        Me.ToMonth.AddItem i
    Next i
        Me.FromMonth.ListIndex = 0
        Me.ToMonth.ListIndex = 0
End Sub
