VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form editfrm 
   Caption         =   "�޸�ѧ����Ϣ"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   8280
   StartUpPosition =   3  '����ȱʡ
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2880
      Top             =   4920
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
      RecordSource    =   "studentinfo"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Caption         =   "����"
      Height          =   495
      Left            =   7080
      TabIndex        =   17
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "����"
      Height          =   495
      Left            =   7080
      TabIndex        =   16
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "ɾ��"
      Height          =   495
      Left            =   7080
      TabIndex        =   15
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "�޸�"
      Height          =   495
      Left            =   7080
      TabIndex        =   14
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "ѧ����Ϣ"
      Height          =   4455
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   6495
      Begin VB.ComboBox Comboxb 
         DataField       =   "�Ա�"
         DataSource      =   "Adodc1"
         Height          =   300
         ItemData        =   "editfrm.frx":0000
         Left            =   1200
         List            =   "editfrm.frx":000A
         TabIndex        =   18
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txthome 
         DataField       =   "��Դ��"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   4560
         TabIndex        =   13
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox txtbj 
         DataField       =   "�༶"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   4560
         TabIndex        =   12
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox txtnj 
         DataField       =   "�꼶"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   4560
         TabIndex        =   11
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtage 
         DataField       =   "��������"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1200
         TabIndex        =   10
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox txtname 
         DataField       =   "����"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1200
         TabIndex        =   9
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtid 
         DataField       =   "ѧ��"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "�������ڣ�"
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "�Ա�"
         Height          =   495
         Left            =   480
         TabIndex        =   6
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "������"
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "ѧ�ţ�"
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "��Դ��:"
         Height          =   615
         Left            =   3840
         TabIndex        =   3
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "�༶:"
         Height          =   495
         Left            =   3840
         TabIndex        =   2
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "�꼶��"
         Height          =   495
         Left            =   3840
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "editfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ChangeEnabled(choose As Boolean)
Dim i As Integer
txtid.Enabled = choose
txtname.Enabled = choose
Comboxb.Enabled = choose
txtage.Enabled = choose
txtnj.Enabled = choose
txtbj.Enabled = choose
txthome.Enabled = choose

End Sub



Private Sub cmddelete_Click()
Adodc1.Recordset.Delete
Adodc1.Recordset.MoveNext
End Sub

Private Sub cmdedit_Click()
If cmdedit.Caption = "�޸�" Then
  cmdedit.Caption = "����"
  Call ChangeEnabled(True)
Else
  Adodc1.Recordset.Fields("ѧ��") = txtid
  Adodc1.Recordset.Fields("����") = txtname
  Adodc1.Recordset.Fields("�Ա�") = Comboxb
  Adodc1.Recordset.Fields("��������") = txtage
  Adodc1.Recordset.Fields("�꼶") = txtnj
  Adodc1.Recordset.Fields("�༶") = txtbj
  Adodc1.Recordset.Fields("��Դ��") = txthome
  Adodc1.Recordset.Update
  MsgBox "��ѧ����Ϣ���޸ģ�", vbOKOnly + vbInformation, "��ʾ"
  Call ChangeEnabled(False)
End If

End Sub

Private Sub cmdexit_Click()
editfrm.Hide
mainfrm.Show
End Sub

Private Sub cmdfind_Click()
findfrm.Show
editfrm.Hide
End Sub


Private Sub Form_Load()
Call ChangeEnabled(False)
End Sub
