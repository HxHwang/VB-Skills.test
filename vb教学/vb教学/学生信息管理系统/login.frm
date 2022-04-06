VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form login 
   Caption         =   "登录界面"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3960
      Top             =   2880
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\学生信息管理系统\student.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\学生信息管理系统\student.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "userinfo"
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
   Begin VB.CommandButton cmdexit 
      Caption         =   "退出"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "登录"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txtpwd 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox txtname 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "密码："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "用户名："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
Dim ans As String
ans = MsgBox("确定退出系统？", vbYesNo + vbInformation, "退出")
If ans = vbYes Then
   End
End If
End Sub

Private Sub cmdok_Click()
Adodc1.Refresh
If Trim(txtname.Text) = "" Then
  MsgBox "用户名不能为空！", vbOKOnly + vbInformation, "登录"
Else
  Adodc1.Recordset.Find "username='" & Trim(txtname.Text) & "'"
  If Adodc1.Recordset.EOF Then
    MsgBox "无此用户！", , "登录"
  Else
    If Adodc1.Recordset.Fields("userpwd") = Trim(txtpwd.Text) Then
      mainfrm.Show
      login.Hide
    Else
      MsgBox "密码不正确，请重试！", , "登录"
    End If
  End If
End If
    
'Dim idinfo As Recordset
'Dim sqlstr As String
'DBEngine.DefaultType = dbUseJet
'Set coursedb = DBEngine.OpenDatabase("student.mdb", False, False)
'sqlstr = "select username,userpwd from userinfo where username='" & txtname & "'"
'Set idinfo = coursedb.OpenRecordset(sqlstr, dbOpenSnapshot, dbReadOnly)
'If (idinfo.RecordCount = 0) Then
  'MsgBox "无此用户！", , "登录"
'Else
  'If (idinfo.Fields("userpwd").Value = txtpwd) Then
    'introfrm.Hide
    'Unload login
    'mainfrm.Show
  'Else
    'MsgBox "无效的密码，请重试！", , "登录"
  'End If
'End If

    

'If Trim(txtname.Text) = "admin" And Trim(txtpsw.Text) = "123" Then
  'introfrm.Hide
  'login.Hide
  'mainfrm.Show
'Else
  'MsgBox "请输入正确的用户名和密码！", vbOKOnly + vbCritical, "登录"
  'txtname.Text = ""
  'txtpsw.Text = ""
'End If
End Sub

