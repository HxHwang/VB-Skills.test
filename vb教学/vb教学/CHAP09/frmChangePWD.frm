VERSION 5.00
Begin VB.Form frmChangePWD 
   Caption         =   "修改密码"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5805
   LinkTopic       =   "FrmMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   5805
   StartUpPosition =   2  '屏幕中心
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
      Left            =   3720
      TabIndex        =   5
      Top             =   2760
      Width           =   1695
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
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox ConfirmPWD 
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2880
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox NewPWD 
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2880
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "请确认新密码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "请输入新密码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "frmChangePWD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()



Dim SQL As String
    Dim rs As ADODB.Recordset

    'If Trim(OldPWD.Text) = UserPWD Then                            '判断是否输入旧密码
    'MsgBox "请输入旧密码！", vbOKOnly + vbExclamation, "警告"
        'OldPWD.SetFocus
        'Exit Sub
    'Else
        If Trim(NewPWD.Text) = "" Then                        '判断是否输入新密码
            MsgBox "请输入新密码！", vbOKOnly + vbExclamation, "警告"
            NewPWD.SetFocus
            Exit Sub
        ElseIf Trim(NewPWD.Text) <> Trim(ConfirmPWD.Text) Then '判断两次密码是否相同
            MsgBox "两次密码不同！", vbOKOnly + vbExclamation, "警告"
            NewPWD.Text = ""
            ConfirmPWD.Text = ""
            NewPWD.SetFocus
        Else
             ' If Trim(OldPWD.Text) = UserPWD Then                                                    '修改密码
            SQL = "update UserInfo set UserPWD = '" & NewPWD & "'where UserID='"
            SQL = SQL & gUserName & "'"
            TransactSQL (SQL)
            MsgBox "密码已经修改！", vbOKOnly + vbExclamation, "修改结果"
            Unload Me
           
    End If

End Sub

Private Sub Form_Load()
    'OldPWD.Text = ""
    NewPWD.Text = ""
    ConfirmPWD.Text = ""
End Sub
