VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form adduser 
   Caption         =   "添加新用户"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   6180
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3840
      Top             =   3000
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=student.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=student.mdb;Persist Security Info=False"
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
   Begin VB.TextBox Text3 
      DataSource      =   "Adodc1"
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc1"
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "确认密码："
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "用户密码："
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label label1 
      Caption         =   "新用户名："
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "adduser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Trim(text1.Text) = "" Then
 MsgBox "用户名不能为空！", vbOKOnly + vbExclamation, "提示"
Else
 Adodc1.Recordset.Find "username='" & Trim(text1.Text) & "'"
 If Adodc1.Recordset.EOF = False Then
   MsgBox "该用户名已被使用，请重输！", vbOKOnly + vbExclamation, "提示"
   text1.SetFocus
   text1.Text = ""
   Text2.Text = ""
   text3.Text = ""
 Else
   If Trim(Text2.Text) <> Trim(text3.Text) Then
     MsgBox "两次密码不一致，请重输！", vbOKOnly + vbExclamation, "提示"
     Text2.SetFocus
     Text2.Text = ""
     text3.Text = ""
   Else
     Adodc1.Recordset.AddNew
     Adodc1.Recordset.Fields(0) = Trim(text1.Text)
     Adodc1.Recordset.Fields(1) = Trim(Text2.Text)
     Adodc1.Recordset.Update
     MsgBox "恭喜你，注册成功！", vbOKOnly + vbExclamation, "注册"
     text1.Text = ""
     Text2.Text = ""
     text3.Text = ""
   End If
 End If
End If
End Sub

Private Sub Command2_Click()
mainfrm.Show
adduser.Hide
End Sub
