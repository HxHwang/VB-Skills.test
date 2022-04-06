VERSION 5.00
Begin VB.Form Login 
   Caption         =   "登陆学生成绩查询系统"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3795
   LinkTopic       =   "Form2"
   ScaleHeight     =   3720
   ScaleWidth      =   3795
   StartUpPosition =   3  '窗口缺省
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\fjlg\Desktop\项目10课堂\学生成绩信息库.mdb"
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "帐号管理"
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "注册新用户"
      Height          =   1095
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "登录"
      Default         =   -1  'True
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.TextBox Text2 
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   1080
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   1
         Text            =   "administrator"
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "密码："
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "用户名："
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NameFind As String
Dim PwdFind As String

Private Sub cmdCancel_Click()
    Unload Login
End Sub

Private Sub cmdOK_Click()
    
If Text1.Text = "" Then
     MsgBox "请填写用户名！", vbOKOnly + vbInformation, "注意"
     Text1.SetFocus
     Exit Sub
End If
If Text2.Text = "" Then
     MsgBox "请填写密码！", vbOKOnly + vbInformation, "注意"
     Text2.SetFocus
     Exit Sub
End If
Data1.Recordset.FindFirst "用户名='" & Trim(Text1.Text) & " '"
NameFind = Data1.Recordset.Bookmark          '记录找到的用户名的书签
If Data1.Recordset.NoMatch = False Then          '找到用户名
    Data1.Recordset.FindFirst "密码='" & Trim(Text2.Text) & "'"
    PwdFind = Data1.Recordset.Bookmark
    If Data1.Recordset.NoMatch = False And NameFind = PwdFind Then
      If Trim(Text1.Text) = "administrator" Then
         Admin = True
      End If
      Main.Show
      Unload Me
    Else
      MsgBox "密码不正确！请重试……", vbOKOnly + vbInformation, "错误"
      Text2.Text = ""
    End If
Else
   MsgBox "无此用户！请先注册……", vbOKOnly + vbInformation, "错误"
        
End If

    
End Sub

Private Sub cmdRegister_Click()
    Unload Login
    Reg = 1
    Register.Show
End Sub



 
