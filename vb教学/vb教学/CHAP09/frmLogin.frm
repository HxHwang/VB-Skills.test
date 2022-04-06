VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   Caption         =   "用户登录"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3780
   ScaleMode       =   0  'User
   ScaleWidth      =   5400
   StartUpPosition =   2  '屏幕中心
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2400
      Top             =   3120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Connect         =   $"frmLogin.frx":0000
      OLEDBString     =   $"frmLogin.frx":00A6
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "UserInfo"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
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
      Left            =   3360
      TabIndex        =   6
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
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
      Left            =   480
      TabIndex        =   5
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox PassWord 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   4
      Text            =   "1111"
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox UserName 
      Height          =   480
      Left            =   2280
      TabIndex        =   3
      Text            =   "Admin"
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "用户密码"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "用户名称"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "学生档案管理管理系统"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pwdCount As Integer

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdOK_Click()
    Dim SQL As String
    Dim rs As ADODB.Recordset
    If Trim(UserName.Text = "") Then
        MsgBox "没有输入用户名称，请重新输入！", vbOKOnly + vbExclamation, "警告"
        UserName.SetFocus
    Else                                                '查询用户
        SQL = "select * from UserInfo where UserID='" & UserName.Text & "'"
    Set rs = TransactSQL(SQL)
        If iflag = 1 Then
            If rs.EOF = True Then
                MsgBox "没有这个用户，请重新输入！", vbOKOnly + vbExclamation, "警告"
                UserName.SetFocus
            Else
                If Trim(rs.Fields(1)) = Trim(PassWord.Text) Then
                    rs.Close
                    Me.Hide
                    gUserName = Trim(UserName.Text)         '保存用户名称
                    FrmMain.Show
                    Unload Me
                Else
                    MsgBox "密码不正确，请重新输入！", vbOKOnly + vbExclamation, "警告"
                    PassWord.SetFocus
                    PassWord.Text = ""
                End If
            End If
        Else
            Unload Me
        End If
    End If
    pwdCount = pwdCount + 1                             '判断输入次数
    If pwdCount = 3 Then
        Unload Me
        Exit Sub
    End If
End Sub



Private Sub Form_Load()
 
    pwdCount = 0
    UserName = ""
End Sub

Private Sub PassWord_KeyDown(KeyCode As Integer, Shift As Integer)
  TabToEnter KeyCode
End Sub

Private Sub UserName_KeyDown(KeyCode As Integer, Shift As Integer)
TabToEnter KeyCode
End Sub
