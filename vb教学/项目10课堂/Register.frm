VERSION 5.00
Begin VB.Form Register 
   Caption         =   "ע��Ϊ�û�"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5055
   LinkTopic       =   "Form2"
   ScaleHeight     =   4560
   ScaleWidth      =   5055
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame FrameReg 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4815
      Begin VB.TextBox txtUserName 
         DataField       =   "�û���"
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
         DatabaseName    =   "ѧ���ɼ���Ϣ��.mdb"
         DefaultCursorType=   0  'ȱʡ�α�
         DefaultType     =   2  'ʹ�� ODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "�ʺŹ���"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "ȷ��"
         Height          =   495
         Left            =   2760
         TabIndex        =   17
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox txtUnit 
         DataField       =   "�༶"
         DataSource      =   "Data1"
         Height          =   270
         Left            =   1200
         TabIndex        =   16
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtAge 
         DataField       =   "����"
         DataSource      =   "Data1"
         Height          =   270
         Left            =   1200
         TabIndex        =   15
         Top             =   2280
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtRealname 
         DataField       =   "��ʵ����"
         DataSource      =   "Data1"
         Height          =   270
         Left            =   1200
         TabIndex        =   14
         Top             =   1800
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.TextBox txtPassword 
         DataField       =   "����"
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
         Caption         =   "ע��"
         Default         =   -1  'True
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "�û�����"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   375
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "���룺"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   855
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "����ȷ�ϣ�"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1335
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "��ʵ������"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1815
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "���䣺"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "�༶��"
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
      Caption         =   "��ӭע���Ϊ���û���(��*��Ϊ�����"
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
On Error GoTo aaa
If txtUserName.Text = "" Then
End If
If txtPassword.Text <> txtPwAgain.Text Then
End If
'Data1.Recordset.Update
Data1.UpdateRecord
Data1.Recordset.MoveLast
Exit Sub

aaa:
MsgBox ""
End Sub

Private Sub cmdReg_Click()
Cmd.Enabled = True
cmdReg.Enabled = False
txtUserName.Visible = True
txtPassword.Visible = True
txtPwAgain.Visible = True
Data1.Recordset.AddNew
txtUserName.SetFocus
End Sub

Private Sub Form_Load()
Cmd.Enabled = False
End Sub
