VERSION 5.00
Begin VB.Form Login 
   Caption         =   "��½ѧ���ɼ���ѯϵͳ"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3795
   LinkTopic       =   "Form2"
   ScaleHeight     =   3720
   ScaleWidth      =   3795
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\fjlg\Desktop\��Ŀ10����\ѧ���ɼ���Ϣ��.mdb"
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "�ʺŹ���"
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��"
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "ע�����û�"
      Height          =   1095
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "��¼"
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
         Caption         =   "���룺"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "�û�����"
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
     MsgBox "����д�û�����", vbOKOnly + vbInformation, "ע��"
     Text1.SetFocus
     Exit Sub
End If
If Text2.Text = "" Then
     MsgBox "����д���룡", vbOKOnly + vbInformation, "ע��"
     Text2.SetFocus
     Exit Sub
End If
Data1.Recordset.FindFirst "�û���='" & Trim(Text1.Text) & " '"
NameFind = Data1.Recordset.Bookmark          '��¼�ҵ����û�������ǩ
If Data1.Recordset.NoMatch = False Then          '�ҵ��û���
    Data1.Recordset.FindFirst "����='" & Trim(Text2.Text) & "'"
    PwdFind = Data1.Recordset.Bookmark
    If Data1.Recordset.NoMatch = False And NameFind = PwdFind Then
      If Trim(Text1.Text) = "administrator" Then
         Admin = True
      End If
      Main.Show
      Unload Me
    Else
      MsgBox "���벻��ȷ�������ԡ���", vbOKOnly + vbInformation, "����"
      Text2.Text = ""
    End If
Else
   MsgBox "�޴��û�������ע�ᡭ��", vbOKOnly + vbInformation, "����"
        
End If

    
End Sub

Private Sub cmdRegister_Click()
    Unload Login
    Reg = 1
    Register.Show
End Sub



 
