VERSION 5.00
Begin VB.Form frmChangePWD 
   Caption         =   "�޸�����"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5805
   LinkTopic       =   "FrmMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   5805
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
      Left            =   3720
      TabIndex        =   5
      Top             =   2760
      Width           =   1695
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
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox ConfirmPWD 
      BeginProperty Font 
         Name            =   "����_GB2312"
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
         Name            =   "����_GB2312"
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
      Caption         =   "��ȷ��������"
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
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "������������"
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

    'If Trim(OldPWD.Text) = UserPWD Then                            '�ж��Ƿ����������
    'MsgBox "����������룡", vbOKOnly + vbExclamation, "����"
        'OldPWD.SetFocus
        'Exit Sub
    'Else
        If Trim(NewPWD.Text) = "" Then                        '�ж��Ƿ�����������
            MsgBox "�����������룡", vbOKOnly + vbExclamation, "����"
            NewPWD.SetFocus
            Exit Sub
        ElseIf Trim(NewPWD.Text) <> Trim(ConfirmPWD.Text) Then '�ж����������Ƿ���ͬ
            MsgBox "�������벻ͬ��", vbOKOnly + vbExclamation, "����"
            NewPWD.Text = ""
            ConfirmPWD.Text = ""
            NewPWD.SetFocus
        Else
             ' If Trim(OldPWD.Text) = UserPWD Then                                                    '�޸�����
            SQL = "update UserInfo set UserPWD = '" & NewPWD & "'where UserID='"
            SQL = SQL & gUserName & "'"
            TransactSQL (SQL)
            MsgBox "�����Ѿ��޸ģ�", vbOKOnly + vbExclamation, "�޸Ľ��"
            Unload Me
           
    End If

End Sub

Private Sub Form_Load()
    'OldPWD.Text = ""
    NewPWD.Text = ""
    ConfirmPWD.Text = ""
End Sub
