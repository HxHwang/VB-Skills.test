VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmsearch 
   Caption         =   "成绩查询窗体"
   ClientHeight    =   7395
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   11835
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "选择查询的方式"
      Height          =   3135
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   9855
      Begin VB.CommandButton Command3 
         Caption         =   "返回"
         Height          =   615
         Left            =   7680
         TabIndex        =   12
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "所有"
         Height          =   615
         Left            =   7680
         TabIndex        =   11
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton Option3 
         Caption         =   "按成绩"
         Height          =   615
         Left            =   960
         TabIndex        =   10
         Top             =   2040
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         Caption         =   "按课程名"
         Height          =   615
         Left            =   960
         TabIndex        =   9
         Top             =   1320
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "按学号"
         Height          =   615
         Left            =   960
         TabIndex        =   8
         Top             =   480
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.TextBox txtscore2 
         Height          =   495
         Left            =   5640
         TabIndex        =   6
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtscore1 
         Height          =   495
         Left            =   3360
         TabIndex        =   5
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtcourse 
         Height          =   495
         Left            =   3360
         TabIndex        =   4
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox txtnumber 
         Height          =   495
         Left            =   3360
         TabIndex        =   3
         Top             =   480
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "查询"
         Height          =   615
         Left            =   7680
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "————"
         Height          =   495
         Left            =   4920
         TabIndex        =   7
         Top             =   2280
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3720
      Top             =   7200
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=学生成绩信息库.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=学生成绩信息库.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3495
      Left            =   1440
      TabIndex        =   0
      Top             =   3600
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   6165
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "学号"
         Caption         =   "学号"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "姓名"
         Caption         =   "姓名"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "性别"
         Caption         =   "性别"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "年龄"
         Caption         =   "年龄"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "课程号"
         Caption         =   "课程号"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "课程名"
         Caption         =   "课程名"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "任课教师"
         Caption         =   "任课教师"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "成绩"
         Caption         =   "成绩"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1275.024
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Option1.Value = True Then
   If txtnumber.Text = "" Then
     MsgBox "请输入要查询的学号", vbOKOnly + vbCritical, "提示"
     txtnumber.SetFocus
     Exit Sub
   Else
     Adodc1.RecordSource = "select * from 总查询表 where 学号=" & Val(txtnumber.Text)
     Adodc1.Refresh
     txtnumber.Text = ""
   End If
ElseIf Option2.Value = True Then
    If txtcourse.Text = "" Then
     MsgBox "请输入要查询的课程名", vbOKOnly + vbCritical, "提示"
     txtcourse.SetFocus
     Exit Sub
   Else
     Adodc1.RecordSource = "select * from 总查询表 where 课程名='" & Trim(txtcourse.Text) & "'"
     Adodc1.Refresh
     txtcourse.Text = ""
   End If
ElseIf Option3.Value = True Then
    If txtscore1.Text = "" Or txtscore2.Text = "" Then
     MsgBox "请输入要查询的成绩范围", vbOKOnly + vbCritical, "提示"
     txtscore1.SetFocus
     Exit Sub
   Else
     Adodc1.RecordSource = "select * from 总查询表 where 成绩 between " & Val(txtscore1.Text) & " and " & Val(txtscore2.Text)
     Adodc1.Refresh
     txtscore1.Text = ""
     txtscore2.Text = ""
   End If
End If

End Sub

Private Sub Command2_Click()
Adodc1.RecordSource = "select * from 总查询表 "
Adodc1.Refresh
End Sub

Private Sub Form_Load()
Adodc1.RecordSource = "select * from 总查询表 "
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
End Sub
