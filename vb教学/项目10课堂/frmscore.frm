VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "Msadodc.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form scorefrm 
   Caption         =   "添加和修改学生成绩"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13905
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   13905
   StartUpPosition =   3  '窗口缺省
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   495
      Left            =   7920
      Top             =   6720
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      RecordSource    =   "课程信息表"
      Caption         =   "Adodc2"
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
   Begin VB.Frame Frame1 
      Caption         =   "学生成绩"
      Height          =   5775
      Left            =   6480
      TabIndex        =   1
      Top             =   960
      Width           =   6135
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   1800
         TabIndex        =   14
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "返回"
         Height          =   495
         Left            =   4440
         TabIndex        =   12
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "修改"
         Height          =   495
         Left            =   4440
         TabIndex        =   11
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "添加"
         Height          =   495
         Left            =   4440
         TabIndex        =   10
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "删除"
         Height          =   495
         Left            =   4440
         TabIndex        =   9
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox txtnumber 
         Height          =   495
         Left            =   1800
         TabIndex        =   4
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtname 
         Height          =   495
         Left            =   1800
         TabIndex        =   3
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox txtscore 
         Height          =   495
         Left            =   1800
         TabIndex        =   2
         Top             =   4320
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "（单击选择要添加的课程名）"
         Height          =   255
         Left            =   1680
         TabIndex        =   15
         Top             =   2640
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "学号："
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "课程名："
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "课程号："
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "成绩："
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   4440
         Width           =   855
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmscore.frx":0000
      Height          =   5775
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   10186
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
      Left            =   1080
      Top             =   6840
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      RecordSource    =   "学生成绩表"
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
   Begin VB.Label Label5 
      Caption         =   "添加和修改学生成绩"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   13
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "scorefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer
Dim i As Integer
Dim tianjia As Boolean


Private Sub Combo1_Click()
Adodc2.Refresh
Adodc2.Recordset.Find "课程名='" & Combo1.Text & "'"
txtname.Text = Str(Adodc2.Recordset.Fields(0))
End Sub

Private Sub Command1_Click()

Main.Show
infofrm.Hide
End Sub

Private Sub Command2_Click()
If Command2.Caption = "添加" Then
  Command2.Caption = "保存"
  txtnumber.Enabled = True
  txtname.Enabled = True
  Combo1.Enabled = True
  txtscore.Enabled = True
  txtnumber.Text = ""
  txtname.Text = ""
  Combo1.Text = ""
  txtscore.Text = ""
  tianjia = True
  Label6.Visible = True
Else
  Adodc1.Refresh
  For i = 0 To Adodc1.Recordset.RecordCount - 1
     If Adodc1.Recordset.Fields(0) = Val(txtnumber.Text) Then
        If Adodc1.Recordset.Fields(1) = Val(txtname.Text) Then
           MsgBox "该学生该门课程已有成绩！", vbOKOnly + vbCritical, "提示"
           Exit Sub
        End If
     End If
     Adodc1.Recordset.MoveNext
   Next i
  Adodc1.Recordset.AddNew
  Adodc1.Recordset.Fields(0) = Trim(txtnumber.Text)
  Adodc1.Recordset.Fields(1) = Trim(txtname.Text)
  Adodc1.Recordset.Fields(2) = Trim(txtscore.Text)
  Adodc1.Recordset.Update
  MsgBox "添加成功！", vbOKOnly + vbInformation, "添加"
  DataGrid1.Refresh
  Command2.Caption = "添加"
  txtnumber.Enabled = False
  txtname.Enabled = False
  Combo1.Enabled = False
  txtscore.Enabled = False
  tianjia = False
  Label6.Visible = False
End If
End Sub

Private Sub Command3_Click()
If Command3.Caption = "修改" Then
  Command3.Caption = "保存"
  txtnumber.Enabled = True
  txtname.Enabled = True
  Combo1.Enabled = True
  txtscore.Enabled = True
Else
  Adodc1.Recordset.Fields(0) = Trim(txtnumber.Text)
  Adodc1.Recordset.Fields(1) = Trim(txtname.Text)
  
  Adodc1.Recordset.Fields(2) = Trim(txtscore.Text)
  MsgBox "修改成功！", vbOKOnly + vbInformation, "添加"
  DataGrid1.Refresh
  Command3.Caption = "修改"
  txtnumber.Enabled = False
  txtname.Enabled = False
  Combo1.Enabled = False
  txtscore.Enabled = False
End If
End Sub

Private Sub Command4_Click()
On Error GoTo aaa
Dim answer As Integer
answer = MsgBox("确认要删除该学生信息吗？", vbOKCancel + vbQuestion, "删除")
If answer = vbOK Then
  Adodc1.Recordset.Delete
  Adodc1.Recordset.MoveNext
  Exit Sub
Else
  Exit Sub
End If
aaa:
Adodc1.Recordset.MoveLast
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If tianjia = False Then
  txtnumber.Text = DataGrid1.Columns(0).Text
  txtname.Text = DataGrid1.Columns(1).Text
  txtscore.Text = DataGrid1.Columns(2).Text
  Adodc2.Refresh
  Adodc2.Recordset.Find "课程号=" & Val(txtname.Text)
  Combo1.Text = Adodc2.Recordset.Fields(1)
End If
End Sub

Private Sub Form_Load()
tianjia = False
txtnumber.Enabled = False
txtname.Enabled = False
Combo1.Enabled = False
txtscore.Enabled = False
Adodc2.Refresh
For i = 0 To Adodc2.Recordset.RecordCount - 1
  Combo1.AddItem Adodc2.Recordset.Fields(1)
  Adodc2.Recordset.MoveNext
Next i
End Sub
