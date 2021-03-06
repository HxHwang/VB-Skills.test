VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form findresult 
   Caption         =   "学生成绩查询"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   8490
   StartUpPosition =   3  '窗口缺省
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "findresult.frx":0000
      Height          =   2625
      Left            =   600
      TabIndex        =   6
      Top             =   2640
      Width           =   7050
      _ExtentX        =   12435
      _ExtentY        =   4630
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
      ColumnCount     =   5
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
      BeginProperty Column04 
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
            Alignment       =   2
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1094.74
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "选择查询条件"
      Height          =   2055
      Left            =   600
      TabIndex        =   3
      Top             =   480
      Width           =   2535
      Begin VB.OptionButton optid 
         Caption         =   "按学号"
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton optname 
         Caption         =   "按课程名"
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   1200
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "返回"
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox txtfind 
      Height          =   615
      Left            =   3960
      TabIndex        =   0
      Top             =   840
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4320
      Top             =   5040
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=student.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=student.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"findresult.frx":0015
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
End
Attribute VB_Name = "findresult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim sqlstr As String
If Trim(txtfind.Text) = "" Then
  MsgBox "请先输入查询条件！", vbOKOnly + vbCritical, "查询"
Else
  If optid.Value = True Then
   sqlstr = "select studentinfo.学号,studentinfo.姓名,studentinfo.性别,courseinfo.课程名,result.成绩 from studentinfo,courseinfo,result where studentinfo.学号=result.学号 and courseinfo.课程编号=result.课程编号 and studentinfo.学号='" & Trim(txtfind.Text) & "'"
   Adodc1.RecordSource = sqlstr
   Adodc1.Refresh
   If Adodc1.Recordset.EOF Then
     MsgBox "查询记录为空！", vbOKOnly + vbInformation, "查询结果"
   End If
   
  Else
   sqlstr = "select studentinfo.学号,studentinfo.姓名,studentinfo.性别,courseinfo.课程名,result.成绩 from studentinfo,courseinfo,result where studentinfo.学号=result.学号 and courseinfo.课程编号=result.课程编号 and courseinfo.课程名='" & Trim(txtfind.Text) & "'"
   Adodc1.RecordSource = sqlstr
   Adodc1.Refresh
    If Adodc1.Recordset.EOF Then
     MsgBox "查询记录为空！", vbOKOnly + vbInformation, "查询结果"
   End If
  End If
End If
End Sub

Private Sub Command2_Click()
mainfrm.Show
findresult.Hide
End Sub

