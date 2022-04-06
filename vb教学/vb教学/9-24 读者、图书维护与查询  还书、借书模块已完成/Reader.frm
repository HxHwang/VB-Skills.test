VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBook 
   Caption         =   "图书管理"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8130
   ScaleMode       =   0  'User
   ScaleWidth      =   10485
   Begin MSAdodcLib.Adodc adoQueryResult 
      Height          =   375
      Left            =   840
      Top             =   7800
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   570
      Left            =   7200
      Top             =   7800
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1005
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3135
      Left            =   720
      TabIndex        =   1
      Top             =   4440
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   5530
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   6588
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "图书记录管理"
      TabPicture(0)   =   "Reader.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label10"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label11"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label12"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label13"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label14"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label15"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label16"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdAdd"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdDel"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdUpdate"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdSave"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtBookID"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtBookISBN"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtBookName"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtBookAuthor"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtBookPrice"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtBookBz"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtBookPrint"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "DTPicker1"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).ControlCount=   27
      TabCaption(1)   =   "图书记录查询"
      TabPicture(1)   =   "Reader.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dgdQueryResult"
      Tab(1).Control(1)=   "cmdReturn"
      Tab(1).Control(2)=   "cmdQuery"
      Tab(1).Control(3)=   "txtKeyBookPrint"
      Tab(1).Control(4)=   "txtKeyBookName"
      Tab(1).Control(5)=   "txtKeyAuthor"
      Tab(1).Control(6)=   "txtKeyISBN"
      Tab(1).Control(7)=   "optKeyBookPrint"
      Tab(1).Control(8)=   "optKeyBookName"
      Tab(1).Control(9)=   "optKeyAuthor"
      Tab(1).Control(10)=   "optKeyISBN"
      Tab(1).ControlCount=   11
      Begin MSDataGridLib.DataGrid dgdQueryResult 
         Height          =   2415
         Left            =   -74760
         TabIndex        =   40
         Top             =   480
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   4260
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
      Begin VB.CommandButton cmdReturn 
         Caption         =   "返回"
         Height          =   495
         Left            =   -67080
         TabIndex        =   39
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查询"
         Height          =   495
         Left            =   -70680
         TabIndex        =   38
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtKeyBookPrint 
         Height          =   375
         Left            =   -73680
         TabIndex        =   37
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox txtKeyBookName 
         Height          =   375
         Left            =   -73680
         TabIndex        =   36
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtKeyAuthor 
         Height          =   375
         Left            =   -73680
         TabIndex        =   35
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox txtKeyISBN 
         Height          =   375
         Left            =   -73680
         TabIndex        =   34
         Top             =   1680
         Width           =   2655
      End
      Begin VB.OptionButton optKeyBookPrint 
         Caption         =   "出版社"
         Height          =   495
         Left            =   -74640
         TabIndex        =   33
         Top             =   2160
         Width           =   1455
      End
      Begin VB.OptionButton optKeyBookName 
         Caption         =   "书名"
         Height          =   375
         Left            =   -74640
         TabIndex        =   32
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton optKeyAuthor 
         Caption         =   "作者"
         Height          =   495
         Left            =   -74640
         TabIndex        =   31
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton optKeyISBN 
         Caption         =   "ISBN"
         Height          =   495
         Left            =   -74640
         TabIndex        =   30
         Top             =   1680
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   5880
         TabIndex        =   29
         Top             =   1800
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   25624577
         CurrentDate     =   39712
      End
      Begin VB.TextBox txtBookPrint 
         Height          =   375
         Left            =   1200
         TabIndex        =   28
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox txtBookBz 
         Height          =   855
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   2760
         Width           =   5895
      End
      Begin VB.TextBox txtBookPrice 
         Height          =   375
         Left            =   1200
         TabIndex        =   18
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox txtBookAuthor 
         Height          =   375
         Left            =   5880
         TabIndex        =   17
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtBookName 
         Height          =   375
         Left            =   1200
         TabIndex        =   16
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtBookISBN 
         Height          =   375
         Left            =   5880
         TabIndex        =   15
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtBookID 
         Height          =   375
         Left            =   1200
         TabIndex        =   14
         Top             =   600
         Width           =   2895
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "保存"
         Height          =   495
         Left            =   8160
         TabIndex        =   13
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "修改"
         Height          =   495
         Left            =   8160
         TabIndex        =   12
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除"
         Height          =   495
         Left            =   8160
         TabIndex        =   11
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "添加"
         Height          =   495
         Left            =   8160
         TabIndex        =   10
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   7800
         TabIndex        =   27
         Top             =   1800
         Width           =   165
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   7800
         TabIndex        =   26
         Top             =   1200
         Width           =   165
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   7800
         TabIndex        =   25
         Top             =   600
         Width           =   165
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   4200
         TabIndex        =   24
         Top             =   2280
         Width           =   165
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   4200
         TabIndex        =   23
         Top             =   1680
         Width           =   165
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   4200
         TabIndex        =   22
         Top             =   1200
         Width           =   165
      End
      Begin VB.Label Label10 
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
         Left            =   4200
         TabIndex        =   21
         Top             =   600
         Width           =   165
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "价格"
         Height          =   180
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "备注"
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   2760
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "购买日期"
         Height          =   180
         Left            =   4920
         TabIndex        =   7
         Top             =   1800
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "出版社"
         Height          =   180
         Left            =   240
         TabIndex        =   6
         Top             =   2280
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "ISBN"
         Height          =   180
         Left            =   4920
         TabIndex        =   5
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "作者"
         Height          =   180
         Left            =   4920
         TabIndex        =   4
         Top             =   1200
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "书名"
         Height          =   180
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "图书编号"
         Height          =   180
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   720
      End
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
      Left            =   4320
      TabIndex        =   20
      Top             =   3960
      Width           =   2370
   End
End
Attribute VB_Name = "frmBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Flag As Integer  '0-未启用 1-添加 2-更新
Dim BookID As String


Private Sub cmdAdd_Click()  '添加按钮
Flag = 1
Me.txtBookID.Text = ""
Me.txtBookName.Text = ""
Me.txtBookAuthor.Text = ""
Me.txtBookPrint.Text = ""
Me.txtBookISBN.Text = ""
Me.txtBookBz.Text = ""
Me.txtBookPrice.Text = ""
DongJie
End Sub

Private Sub cmdDel_Click() '删除
BookID = Me.txtBookID.Text
m = MsgBox("确定要删除此条图书记录吗？", vbOKCancel)
If m = 1 Then
 On Error GoTo erp
 sqlstr = " delete from 图书  where  图书编号='" & BookID & "'"
 DBCn.Execute sqlstr
 
 MsgBox "删除成功", , "删除图书档案"
 Adodc1.Refresh
End If
 INIT
 Exit Sub
erp:
 MsgBox "删除失败"
End Sub

Private Sub cmdQuery_Click()  '查询按钮
Dim questr As String
Me.adoQueryResult.ConnectionString = CnStr
questr = "select * from 图书 where "

If Me.optKeyAuthor.Value = True Then    '按作者查询图书
   questr = questr & "作者 like '%" + Me.txtKeyAuthor.Text & "%'"
   
End If
If Me.optKeyBookName.Value = True Then   '按书名查询图书
   questr = questr & "书名  like '%" + Me.txtKeyBookName.Text & "%'"
   
End If
If Me.optKeyISBN.Value = True Then   '按ISBN查询图书
   questr = questr & "isbn  like '%" + Me.txtKeyISBN.Text & "%'"
   
End If
If Me.optKeyBookPrint.Value = True Then   '按出版社查询图书
   questr = questr & "出版社  like  '%" + Me.txtKeyBookPrint.Text & "%'"
   
End If

On Error GoTo erp

adoQueryResult.RecordSource = questr
Me.adoQueryResult.Refresh
Set Me.dgdQueryResult.DataSource = adoQueryResult


If Me.adoQueryResult.Recordset.EOF Then
   MsgBox "数据库中没有符合要求的记录！", , "查询图书档案"
   Exit Sub
End If


Me.dgdQueryResult.Visible = True
Me.cmdReturn.Visible = True

Exit Sub
erp:
MsgBox "查询关键字不正确，请确认查询关键字", vbExclamation, "警告"


End Sub

Private Sub cmdReturn_Click()  '返回按钮
Me.cmdReturn.Visible = False
Me.dgdQueryResult.Visible = False
INIT
End Sub

Private Sub cmdSave_Click() '保存
If Flag = 1 Then        '添加后保存
m = MsgBox("确定要添加此条图书记录吗？", vbOKCancel)
 If m = 1 Then
   sqlstr = "insert into 图书  "
   sqlstr = sqlstr & "values('" & Me.txtBookID & "','" & Me.txtBookName & "','" & Me.txtBookAuthor & "','" & Me.txtBookPrint & "','" & Me.txtBookISBN & "','" & Me.DTPicker1.Value & "','" & Me.txtBookBz & "'," & Me.txtBookPrice & " )"
   
   On Error GoTo erp

   DBCn.Execute sqlstr

   Flag = 0

   MsgBox "添加成功", , "添加图书档案"
   Adodc1.Refresh
  End If
INIT
End If

If Flag = 2 Then       '修改后保存
m = MsgBox("确定要修改此条图书记录吗？", vbOKCancel)
  If m = 1 Then

   sqlstr = "update 图书  set "

   sqlstr = sqlstr & "图书编号='" & Me.txtBookID & "',书名='" & Me.txtBookName & "',作者='" & Me.txtBookAuthor & "',出版社='" & Me.txtBookPrint & "',ISBN='" & Me.txtBookISBN & "',购买日期='" & Me.DTPicker1.Value & "',备注='" & Me.txtBookBz & "',价格='" & Me.txtBookPrice & "'"
   
   sqlstr = sqlstr & "where 图书编号='" & BookID & "'"
   
   On Error GoTo erp

   DBCn.Execute sqlstr

   MsgBox "修改成功", , "修改图书档案"
   Flag = 0
   Adodc1.Refresh
End If
INIT
End If
Exit Sub
erp:
MsgBox "请保证录入项目的正确性"



End Sub

Private Sub cmdUpdate_Click() '修改
Flag = 2
DongJie
BookID = Me.txtBookID
End Sub



Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If DataGrid1.Columns(0).Text = "" Then
 Exit Sub
End If
Me.txtBookID.Text = DataGrid1.Columns(0).Text
Me.txtBookName.Text = DataGrid1.Columns(1).Text
Me.txtBookAuthor.Text = DataGrid1.Columns(2).Text
Me.txtBookPrint.Text = DataGrid1.Columns(3).Text
Me.txtBookISBN.Text = DataGrid1.Columns(4).Text
Me.DTPicker1.Value = DataGrid1.Columns(5).Text
Me.txtBookBz.Text = DataGrid1.Columns(6).Text
Me.txtBookPrice.Text = DataGrid1.Columns(7).Text
End Sub

Private Sub Form_Load()

INIT
End Sub
Private Sub INIT()
Flag = 0

Me.cmdAdd.Enabled = True
Me.cmdDel.Enabled = True
Me.cmdUpdate.Enabled = True
Me.cmdSave.Enabled = False
  
Me.txtBookAuthor.Locked = True
Me.txtBookBz.Locked = True

Me.DTPicker1.Enabled = False

Me.txtBookID.Locked = True
Me.txtBookISBN.Locked = True
Me.txtBookName.Locked = True
Me.txtBookPrice.Locked = True
Me.txtBookPrint.Locked = True

DataGrid1.Enabled = True


DataGrid1.AllowUpdate = False
Me.DataGrid1.AllowAddNew = False
Me.DataGrid1.AllowArrows = False
Me.DataGrid1.AllowDelete = False

    
Adodc1.ConnectionString = CnStr
Adodc1.RecordSource = "select * from 图书 "

Set DataGrid1.DataSource = Adodc1

If DataGrid1.Columns(0).Text = "" Then
 Exit Sub
End If

If DataGrid1.Columns(0).Text = "" Then
 Exit Sub
End If


Me.txtBookID.Text = DataGrid1.Columns(0).Text
Me.txtBookName.Text = DataGrid1.Columns(1).Text
Me.txtBookAuthor.Text = DataGrid1.Columns(2).Text
Me.txtBookPrint.Text = DataGrid1.Columns(3).Text
Me.txtBookISBN.Text = DataGrid1.Columns(4).Text
Me.DTPicker1.Value = DataGrid1.Columns(5).Text
Me.txtBookBz.Text = DataGrid1.Columns(6).Text
Me.txtBookPrice.Text = DataGrid1.Columns(7).Text


'查询选项卡
Me.txtKeyAuthor.Visible = True
Me.txtKeyBookName.Visible = True
Me.txtKeyBookPrint.Visible = True
Me.txtKeyISBN.Visible = True


Me.txtKeyAuthor.Enabled = False
Me.txtKeyBookName.Enabled = True
Me.txtKeyBookPrint.Enabled = False
Me.txtKeyISBN.Enabled = False

Me.optKeyAuthor.Value = False
Me.optKeyBookName.Value = True
Me.optKeyBookPrint.Value = False
Me.optKeyISBN.Value = False

Me.cmdReturn.Visible = False
Me.cmdQuery.Visible = True

Me.dgdQueryResult.Visible = False

Me.adoQueryResult.Visible = False

End Sub

Private Sub DongJie()
Me.cmdAdd.Enabled = False
Me.cmdDel.Enabled = False
Me.cmdUpdate.Enabled = False
Me.cmdSave.Enabled = True

Me.txtBookAuthor.Locked = False
Me.txtBookBz.Locked = False

Me.DTPicker1.Enabled = True


Me.txtBookID.Locked = False
Me.txtBookISBN.Locked = False
Me.txtBookName.Locked = False
Me.txtBookPrice.Locked = False
Me.txtBookPrint.Locked = False


DataGrid1.Enabled = False
End Sub



Private Sub EnterQuerying(txt1 As TextBox, txt2 As TextBox, txt3 As TextBox, txt4 As TextBox)
txt1.Enabled = True
txt2.Enabled = False  '单选钮旁边的文本框之间的状态转化
txt3.Enabled = False
txt4.Enabled = False
End Sub

Private Sub optKeyAuthor_Click()
If Me.optKeyAuthor.Value = True Then
   EnterQuerying Me.txtKeyAuthor, Me.txtKeyBookName, Me.txtKeyBookPrint, Me.txtKeyISBN
End If
  
End Sub

Private Sub optKeyBookName_Click()
If Me.optKeyBookName.Value = True Then
   EnterQuerying Me.txtKeyBookName, Me.txtKeyBookPrint, Me.txtKeyISBN, Me.txtKeyAuthor
End If
End Sub

Private Sub optKeyBookPrint_Click()
If Me.optKeyBookPrint.Value = True Then
   EnterQuerying Me.txtKeyBookPrint, Me.txtKeyISBN, Me.txtKeyAuthor, Me.txtKeyBookName
End If
End Sub

Private Sub optKeyISBN_Click()
If Me.optKeyISBN.Value = True Then
   EnterQuerying Me.txtKeyISBN, Me.txtKeyAuthor, Me.txtKeyBookName, Me.txtKeyBookPrint
End If
   
End Sub
