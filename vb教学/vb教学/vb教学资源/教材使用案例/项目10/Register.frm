VERSION 5.00
Begin VB.Form Register 
   Caption         =   "注册为用户"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5055
   LinkTopic       =   "Form2"
   ScaleHeight     =   4560
   ScaleWidth      =   5055
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame FrameReg 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4815
      Begin VB.TextBox txtUserName 
         DataField       =   "用户名"
         DataSource      =   "Data1"
         Height          =   270
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "学生成绩信息库.mdb"
         DefaultCursorType=   0  '缺省游标
         DefaultType     =   2  '使用 ODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "帐号管理"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "确定"
         Height          =   495
         Left            =   2760
         TabIndex        =   17
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox txtUnit 
         DataField       =   "班级"
         DataSource      =   "Data1"
         Height          =   270
         Left            =   1200
         TabIndex        =   16
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtAge 
         DataField       =   "年龄"
         DataSource      =   "Data1"
         Height          =   270
         Left            =   1200
         TabIndex        =   15
         Top             =   2280
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtRealname 
         DataField       =   "真实姓名"
         DataSource      =   "Data1"
         Height          =   270
         Left            =   1200
         TabIndex        =   14
         Top             =   1800
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.TextBox txtPassword 
         DataField       =   "密码"
         DataSource      =   "Data1"
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.TextBox txtPwAgain 
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   1200
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1320
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.CommandButton cmdReg 
         Caption         =   "注册"
         Default         =   -1  'True
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "用户名："
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   375
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "密码："
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   855
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "密码确认："
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1335
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "真实姓名："
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1815
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "年龄："
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "班级："
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "*"
         Height          =   255
         Index           =   0
         Left            =   4320
         TabIndex        =   5
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "*"
         Height          =   255
         Index           =   1
         Left            =   4320
         TabIndex        =   4
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "*"
         Height          =   255
         Index           =   2
         Left            =   4320
         TabIndex        =   3
         Top             =   1320
         Width           =   255
      End
   End
   Begin VB.Label Label1 
      Caption         =   "欢迎注册成为新用户！(打*号为必添项）"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "Register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Click()
On Error GoTo AddErr
    If txtUserName.Text = "" Then
        MsgBox "请填写用户名！", vbOKOnly + vbInformation, "注意"
        txtUserName.SetFocus
        Exit Sub
    ElseIf txtPassword.Text = "" Then
        MsgBox "请填写密码！", vbOKOnly + vbInformation, "注意"
        txtPassword.SetFocus
        Exit Sub
    ElseIf txtPassword.Text <> txtPwAgain.Text Then
        MsgBox "两次密码输入不一致！", vbOKOnly + vbInformation, "注意"
        txtPassword.SetFocus
        Exit Sub
    End If
    
    'Data1.Recordset.Update
    Data1.UpdateRecord
    Unload Register
    If Reg = 1 Then
        Login.Show
    ElseIf Reg = 2 Then
        Main.Show
    End If

    Exit Sub
   
AddErr:
    MsgBox "注册错误！已有该用户名！", vbOKOnly + vbInformation, "错误"
    
    
End Sub

Private Sub cmdReg_Click()
    Data1.Recordset.AddNew
    
    txtUserName.Visible = True
    txtPassword.Visible = True
    txtPwAgain.Visible = True
    txtRealname.Visible = True
    txtAge.Visible = True
    txtUnit.Visible = True
    cmdReg.Enabled = False
End Sub

