VERSION 5.00
Begin VB.Form frmAddUser 
   Caption         =   "����û�"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5580
   ControlBox      =   0   'False
   LinkTopic       =   "FrmMain"
   ScaleHeight     =   4095
   ScaleWidth      =   5580
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   7
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox confirmPWD 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2400
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2280
      Width           =   2535
   End
   Begin VB.TextBox PassWord 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2400
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox UserName 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      MaxLength       =   8
      TabIndex        =   3
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "  ȷ������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "  �û�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "���û�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim sql As String
    Dim rs As ADODB.Recordset
    If Trim(UserName.Text) = "" Then                              '�ж��û������Ƿ�Ϊ��
        MsgBox "�������û�����!", vbOKOnly + vbExclamation, "����"
        Exit Sub
        UserName.SetFocus
    Else
        sql = "select * from UserInfo where UserID='" & UserName & "'"
        Set rs = TransactSQL(sql)
        If rs.EOF = False Then                                    '�ж��Ƿ��Ѿ������û�
            MsgBox "����û��Ѿ����ڣ������������û����ƣ�", vbOKOnly + vbExclamation, "����"
            UserName.SetFocus
            UserName.Text = ""
            PassWord.Text = ""
            ConfirmPWD.Text = ""
            Exit Sub
        Else
            If Trim(PassWord.Text) <> Trim(ConfirmPWD.Text) Then  '�ж����������Ƿ���ͬ
                MsgBox "������������벻һ�£��������������룡", vbOKOnly + vbExclamation, "����"
                PassWord.Text = ""
                ConfirmPWD.Text = ""
                PassWord.SetFocus
                Exit Sub
            ElseIf Trim(PassWord.Text) = "" Then                  '�ж������Ƿ�Ϊ��
                MsgBox "���벻��Ϊ�գ�", vbOKOnly + vbExclamation, "����"
                PassWord.Text = ""
                ConfirmPWD = ""
                PassWord.SetFocus
            Else                                                 '����û�
                sql = "insert into UserInfo (UserID,UserPWD) values('" & UserName
                sql = sql & "','" & PassWord & "')"
                TransactSQL (sql)
                MsgBox "��ӳɹ���", vbOKOnly + vbExclamation, "��ӽ��"
                                                                 '�������ó�ʼ��Ϊ��
                UserName.Text = ""
                PassWord.Text = ""
                ConfirmPWD.Text = ""
                UserName.SetFocus
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    UserName.Text = ""
    PassWord.Text = ""
    ConfirmPWD.Text = ""
End Sub
