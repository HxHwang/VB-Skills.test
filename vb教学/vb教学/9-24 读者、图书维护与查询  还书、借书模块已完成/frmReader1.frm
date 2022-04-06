VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmReader 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1
   ScaleMode       =   0  'User
   ScaleWidth      =   1
   Begin MSAdodcLib.Adodc adoQueryResult 
      Height          =   495
      Left            =   600
      Top             =   5280
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "adoQueryResult"
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
      Height          =   2055
      Left            =   240
      TabIndex        =   11
      Top             =   3240
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   3625
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5640
      Top             =   5280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
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
      Connect         =   ""
      OLEDBString     =   ""
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5106
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "读者记录管理"
      TabPicture(0)   =   "frmReader1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(1)"
      Tab(0).Control(1)=   "Label2(0)"
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(3)=   "Label9"
      Tab(0).Control(4)=   "Label22"
      Tab(0).Control(5)=   "Label26"
      Tab(0).Control(6)=   "txtReaderID"
      Tab(0).Control(7)=   "txtReaderName"
      Tab(0).Control(8)=   "txtReaderAddress"
      Tab(0).Control(9)=   "cmdAdd"
      Tab(0).Control(10)=   "cmdDel"
      Tab(0).Control(11)=   "cmdUpdate"
      Tab(0).Control(12)=   "cmdSave"
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "读者记录查询"
      TabPicture(1)   =   "frmReader1.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "optReaderName"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "optReaderID"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtKeyReaderName"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtKeyReaderID"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdQuery"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "dgQueryResult"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdReturn"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.CommandButton cmdReturn 
         Caption         =   "返回"
         Height          =   375
         Left            =   5640
         TabIndex        =   18
         Top             =   2280
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid dgQueryResult 
         Height          =   1695
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   2990
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查询"
         Height          =   375
         Left            =   3360
         TabIndex        =   16
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtKeyReaderID 
         Height          =   375
         Left            =   1320
         TabIndex        =   15
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtKeyReaderName 
         Height          =   375
         Left            =   1320
         TabIndex        =   14
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton optReaderID 
         Caption         =   "编号"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   1320
         Width           =   735
      End
      Begin VB.OptionButton optReaderName 
         Caption         =   "姓名"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "保存"
         Height          =   375
         Left            =   -69000
         TabIndex        =   10
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "修改"
         Height          =   375
         Left            =   -69000
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除"
         Height          =   375
         Left            =   -69000
         TabIndex        =   8
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "添加"
         Height          =   375
         Left            =   -69000
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtReaderAddress 
         Height          =   375
         Left            =   -73800
         TabIndex        =   6
         Top             =   1680
         Width           =   4095
      End
      Begin VB.TextBox txtReaderName 
         Height          =   375
         Left            =   -73800
         TabIndex        =   5
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtReaderID 
         Height          =   375
         Left            =   -73800
         TabIndex        =   4
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   -71640
         TabIndex        =   21
         Top             =   1080
         Width           =   165
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   -71640
         TabIndex        =   20
         Top             =   600
         Width           =   165
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "*号项目必须填写"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   -72480
         TabIndex        =   19
         Top             =   2280
         Width           =   2370
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "住址"
         Height          =   180
         Left            =   -74640
         TabIndex        =   3
         Top             =   1680
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "读者姓名"
         Height          =   180
         Index           =   0
         Left            =   -74640
         TabIndex        =   2
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "读者ID"
         Height          =   180
         Index           =   1
         Left            =   -74640
         TabIndex        =   1
         Top             =   720
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReaderID As String
Dim Flag As Integer  '0-未启用 1-添加 2-更新



Private Sub cmdAdd_Click()
Me.txtReaderAddress.Text = ""
Me.txtReaderID.Text = ""
Me.txtReaderName.Text = ""
Flag = 1
DongJie

End Sub

Private Sub cmdDel_Click()
Dim m As Integer
Dim sqlstr As String
ReaderID = Me.txtReaderID.Text
m = MsgBox("确定要删除此条读者记录吗？", vbOKCancel)
If m = 1 Then
 On Error GoTo erp
 sqlstr = " delete from 读者  where  读者编号='" & ReaderID & "'"
 dbcn.Execute sqlstr
 
 MsgBox "删除成功", , "删除读者档案"
 adodc1.Refresh
End If
 INIT
 Exit Sub
erp:
 MsgBox "删除失败"
 INIT
End Sub

Private Sub cmdQuery_Click()

Dim questr As String
Me.adoQueryResult.ConnectionString = CnStr
questr = "select * from 读者 where "

If Me.optReaderID.Value = True Then    '按编号查询读者
   questr = questr & "读者编号 like '%" + Me.txtKeyReaderID.Text & "%'"
   
End If
If Me.optReaderName.Value = True Then   '按读者姓名查询读者
   questr = questr & "读者姓名  like '%" + Me.txtKeyReaderName.Text & "%'"
   
End If


On Error GoTo erp

adoQueryResult.RecordSource = questr
Me.adoQueryResult.Refresh
Set Me.dgQueryResult.DataSource = adoQueryResult


If Me.adoQueryResult.Recordset.EOF Then
   MsgBox "数据库中没有符合要求的记录！", , "查询读者档案"
   Exit Sub
End If


Me.dgQueryResult.Visible = True
Me.cmdReturn.Visible = True

Exit Sub
erp:
MsgBox "查询关键字不正确，请确认查询关键字", vbExclamation, "警告"



Me.cmdReturn.Visible = True
Me.dgQueryResult.Visible = True

End Sub

Private Sub cmdReturn_Click()
INIT

End Sub

Private Sub cmdSave_Click()
Dim m As Integer
Dim sqlstr As String
If Flag = 1 Then        '添加后保存
m = MsgBox("确定要添加此条读者记录吗？", vbOKCancel)
 If m = 1 Then
   sqlstr = "insert into 读者  "
   sqlstr = sqlstr & "values('" & Me.txtReaderID & "','" & Me.txtReaderName & "','" & Me.txtReaderAddress & " ')"
   
   On Error GoTo erp

   dbcn.Execute sqlstr

   Flag = 0

   MsgBox "添加成功", , "添加读者档案"
   adodc1.Refresh
  End If
End If

If Flag = 2 Then       '修改后保存
m = MsgBox("确定要修改此条读者记录吗？", vbOKCancel)
  If m = 1 Then

   sqlstr = "update 读者  set "

   sqlstr = sqlstr & "读者编号='" & Me.txtReaderID.Text & "',读者姓名='" & Me.txtReaderName.Text & "',住址='" & Me.txtReaderAddress.Text & "'"
   
   sqlstr = sqlstr & "where 读者编号='" & ReaderID & "'"
   
   On Error GoTo erp

   dbcn.Execute sqlstr

   MsgBox "修改成功", , "修改读者档案"
   Flag = 0
   adodc1.Refresh
End If
End If
INIT
Exit Sub
erp:
MsgBox "请保证录入项目的正确性"
End Sub

Private Sub cmdUpdate_Click()
DongJie
Flag = 2
ReaderID = Me.txtReaderID.Text

End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Datagrid1.Columns(0).Text = "" Then
 Exit Sub
End If
Me.txtReaderID.Text = Datagrid1.Columns(0).Text
Me.txtReaderName.Text = Datagrid1.Columns(1).Text
Me.txtReaderAddress.Text = Datagrid1.Columns(2).Text

End Sub

Private Sub Form_Load()

INIT
End Sub
Private Sub INIT()
Me.txtReaderAddress.Enabled = False
Me.txtReaderID.Enabled = False
Me.txtReaderName.Enabled = False

Me.cmdAdd.Enabled = True
Me.cmdDel.Enabled = True
Me.cmdSave.Enabled = False
Me.cmdUpdate.Enabled = True

Me.Datagrid1.Enabled = True
Datagrid1.AllowUpdate = False
Me.Datagrid1.AllowAddNew = False
Me.Datagrid1.AllowArrows = False
Me.Datagrid1.AllowDelete = False

Me.adodc1.Visible = False

adodc1.ConnectionString = CnStr

adodc1.RecordSource = "select * from 读者 "
Set Datagrid1.DataSource = adodc1

If Datagrid1.Columns(0).Text = "" Then
 Exit Sub
End If
Me.txtReaderID.Text = Datagrid1.Columns(0).Text
Me.txtReaderName.Text = Datagrid1.Columns(1).Text
Me.txtReaderAddress.Text = Datagrid1.Columns(2).Text

Flag = 0

'读者查询初始化
Me.optReaderID.Value = False
Me.optReaderName.Value = True
Me.txtKeyReaderID.Enabled = False
Me.txtKeyReaderName.Enabled = True

Me.adoQueryResult.Visible = False
Me.dgQueryResult.Visible = False

Me.cmdQuery.Visible = True
Me.cmdReturn.Visible = False



End Sub

Private Sub DongJie()
Me.txtReaderAddress.Enabled = True
Me.txtReaderID.Enabled = True
Me.txtReaderName.Enabled = True

Me.cmdAdd.Enabled = False
Me.cmdDel.Enabled = False
Me.cmdSave.Enabled = True
Me.cmdUpdate.Enabled = False

Me.Datagrid1.Enabled = False
Datagrid1.AllowUpdate = False
Me.Datagrid1.AllowAddNew = False
Me.Datagrid1.AllowArrows = False
Me.Datagrid1.AllowDelete = False

Me.adodc1.Visible = False

End Sub

Private Sub EnterQuerying(txt1 As TextBox, txt2 As TextBox)
txt1.Enabled = True
txt2.Enabled = False  '单选钮旁边的文本框之间的状态转化

End Sub



Private Sub optReaderID_Click()
If Me.optReaderID.Value = True Then
  EnterQuerying Me.txtKeyReaderID, Me.txtKeyReaderName
End If
End Sub

Private Sub optReaderName_Click()
If Me.optReaderName.Value = True Then
  EnterQuerying Me.txtKeyReaderName, Me.txtKeyReaderID
End If
End Sub

