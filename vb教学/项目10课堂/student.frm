VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form infofrm 
   Caption         =   "���ѧ����Ϣ"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   10725
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame1 
      Caption         =   "ѧ����Ϣ"
      Height          =   3495
      Left            =   240
      TabIndex        =   1
      Top             =   4200
      Width           =   9855
      Begin VB.CommandButton Command1 
         Caption         =   "����"
         Height          =   495
         Left            =   6960
         TabIndex        =   13
         Top             =   2640
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "�޸�"
         Height          =   495
         Left            =   2640
         TabIndex        =   12
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "���"
         Height          =   495
         Left            =   600
         TabIndex        =   11
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "ɾ��"
         Height          =   495
         Left            =   4680
         TabIndex        =   10
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox txtnumber 
         Height          =   615
         Left            =   2400
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtname 
         Height          =   615
         Left            =   6720
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtsex 
         Height          =   615
         Left            =   2400
         TabIndex        =   3
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtage 
         Height          =   615
         Left            =   6720
         TabIndex        =   2
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "ѧ�ţ�"
         Height          =   375
         Left            =   1080
         TabIndex        =   9
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "�Ա�"
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "������"
         Height          =   375
         Left            =   5640
         TabIndex        =   7
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "���䣺"
         Height          =   375
         Left            =   5640
         TabIndex        =   6
         Top             =   1560
         Width           =   855
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "student.frx":0000
      Height          =   2895
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   5106
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2640
      Top             =   3480
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ѧ���ɼ���Ϣ��.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ѧ���ɼ���Ϣ��.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from ѧ����Ϣ�� order by ѧ�� asc"
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
End
Attribute VB_Name = "infofrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer
Private Sub Command1_Click()

Main.Show
infofrm.Hide
End Sub

Private Sub Command2_Click()
If Command2.Caption = "���" Then
  Command2.Caption = "����"
  txtnumber.Enabled = True
  txtname.Enabled = True
  txtsex.Enabled = True
  txtage.Enabled = True
  txtnumber.Text = ""
  txtname.Text = ""
  txtsex.Text = ""
  txtage.Text = ""
Else
  Adodc1.Recordset.AddNew
  Adodc1.Recordset.Fields(0) = Trim(txtnumber.Text)
  Adodc1.Recordset.Fields(1) = Trim(txtname.Text)
  Adodc1.Recordset.Fields(2) = Trim(txtsex.Text)
  Adodc1.Recordset.Fields(3) = Trim(txtage.Text)
  MsgBox "��ӳɹ���", vbOKOnly + vbInformation, "���"
  DataGrid1.Refresh
  Command2.Caption = "���"
  txtnumber.Enabled = False
  txtname.Enabled = False
  txtsex.Enabled = False
  txtage.Enabled = False
End If
End Sub

Private Sub Command3_Click()
If Command3.Caption = "�޸�" Then
  Command3.Caption = "����"
  txtnumber.Enabled = True
  txtname.Enabled = True
  txtsex.Enabled = True
  txtage.Enabled = True
Else
  Adodc1.Recordset.Fields(0) = Trim(txtnumber.Text)
  Adodc1.Recordset.Fields(1) = Trim(txtname.Text)
  Adodc1.Recordset.Fields(2) = Trim(txtsex.Text)
  Adodc1.Recordset.Fields(3) = Trim(txtage.Text)
  MsgBox "�޸ĳɹ���", vbOKOnly + vbInformation, "���"
  DataGrid1.Refresh
  Command3.Caption = "�޸�"
  txtnumber.Enabled = False
  txtname.Enabled = False
  txtsex.Enabled = False
  txtage.Enabled = False
End If
End Sub

Private Sub Command4_Click()
On Error GoTo aaa
Dim answer As Integer
answer = MsgBox("ȷ��Ҫɾ����ѧ����Ϣ��", vbOKCancel + vbQuestion, "ɾ��")
If answer = vbOK Then
  Adodc1.Recordset.Delete
  Adodc1.Recordset.MoveNext
  Exit Sub
Else
  Exit Sub
End If
aaa:
Adodc1.Recordset.MoveLast
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

txtnumber.Text = DataGrid1.Columns(0).Text
txtname.Text = DataGrid1.Columns(1).Text
txtsex.Text = DataGrid1.Columns(2).Text
txtage.Text = DataGrid1.Columns(3).Text
End Sub

Private Sub Form_Load()
txtnumber.Enabled = False
txtname.Enabled = False
txtsex.Enabled = False
txtage.Enabled = False

End Sub
