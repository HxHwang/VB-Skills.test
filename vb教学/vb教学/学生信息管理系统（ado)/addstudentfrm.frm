VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form addstudentfrm 
   Caption         =   "���ѧ����Ϣ"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   8280
   StartUpPosition =   3  '����ȱʡ
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   4920
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   582
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
      Left            =   5520
      TabIndex        =   15
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "���"
      Height          =   495
      Left            =   2760
      TabIndex        =   14
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "ѧ����Ϣ"
      Height          =   4455
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   6495
      Begin VB.ComboBox Comboxb 
         DataSource      =   "Adodc1"
         Height          =   300
         ItemData        =   "addstudentfrm.frx":0000
         Left            =   1200
         List            =   "addstudentfrm.frx":000A
         TabIndex        =   16
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txthome 
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   4560
         TabIndex        =   13
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox txtbj 
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   4560
         TabIndex        =   12
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox txtnj 
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   4560
         TabIndex        =   11
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtage 
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1200
         TabIndex        =   10
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox txtname 
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1200
         TabIndex        =   9
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtid 
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
Attribute VB_Name = "addstudentfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdadd_Click()
If Trim(txtid.Text) = "" Then
  MsgBox "ѧ�Ų���Ϊ�գ������䣡", vbOKOnly + vbInformation, "��ʾ"
  Exit Sub
  txtid.SetFocus
End If
If Trim(txtname.Text) = "" Then
  MsgBox "��������Ϊ�գ������䣡", vbOKOnly + vbInformation, "��ʾ"
  Exit Sub
  txtname.SetFocus
End If
If Len(Trim(txtid.Text)) <> 5 Or IsNumeric(Trim(txtid.Text)) = False Then
  MsgBox "ѧ����������������5λ����", vbOKOnly + vbInformation, "��ʾ"
  Exit Sub
  txtid.SetFocus
End If
Adodc1.Recordset.Find "ѧ��='" & Trim(txtid.Text) & "'"
If Adodc1.Recordset.EOF = False Then
    MsgBox "ѧ���ظ�������������", vbOKOnly + vbInformation, "��ʾ"
Else
    Adodc1.Recordset.AddNew
    Adodc1.Recordset.Fields("ѧ��") = txtid.Text
    Adodc1.Recordset.Fields("����") = txtname.Text
    Adodc1.Recordset.Fields("�Ա�") = Comboxb.Text
    Adodc1.Recordset.Fields("��������") = txtage.Text
    Adodc1.Recordset.Fields("�꼶") = txtnj.Text
    Adodc1.Recordset.Fields("�༶") = txtbj.Text
    Adodc1.Recordset.Fields("��Դ��") = txthome.Text
    Adodc1.Recordset.Update
    Adodc1.Refresh
    MsgBox "��ӳɹ���", vbOKOnly + vbInformation, "��Ӽ�¼"
    infofrm.Show
    addstudentfrm.Hide
End If
End Sub
Private Sub cmdexit_Click()
mainfrm.Show
addstudentfrm.Hide

End Sub

