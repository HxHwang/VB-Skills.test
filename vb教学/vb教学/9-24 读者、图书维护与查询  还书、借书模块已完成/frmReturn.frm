VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmReturn 
   Caption         =   "���鴦��"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8220
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   4155
   ScaleWidth      =   8220
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "��ʾ������Ϣ"
      Height          =   495
      Left            =   3480
      TabIndex        =   19
      Top             =   120
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1800
      Top             =   4320
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   735
      Left            =   240
      TabIndex        =   18
      Top             =   4080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
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
   Begin VB.CommandButton cmdReturn 
      Caption         =   "����"
      Height          =   495
      Left            =   5040
      TabIndex        =   17
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtBookID 
      Height          =   375
      Left            =   1440
      TabIndex        =   16
      Top             =   120
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "������Ϣ"
      Height          =   3135
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   7695
      Begin VB.TextBox txtReaderAddress 
         Height          =   375
         Left            =   4800
         TabIndex        =   14
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txtReaderID 
         Height          =   375
         Left            =   4800
         TabIndex        =   13
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtReaderName 
         Height          =   375
         Left            =   4800
         TabIndex        =   12
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtBookPrint 
         Height          =   375
         Left            =   1200
         TabIndex        =   11
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txtBookAuthor 
         Height          =   375
         Left            =   1200
         TabIndex        =   10
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtBookName 
         Height          =   375
         Left            =   1200
         TabIndex        =   9
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtLendDate 
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblFanJin 
         BackColor       =   &H0000FFFF&
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   4800
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblDisplay 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   3720
         TabIndex        =   20
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "����סַ"
         Height          =   180
         Left            =   3720
         TabIndex        =   8
         Top             =   2400
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "���߱��"
         Height          =   180
         Left            =   3720
         TabIndex        =   7
         Top             =   1800
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Left            =   3720
         TabIndex        =   6
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Left            =   120
         TabIndex        =   3
         Top             =   2400
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   240
         TabIndex        =   2
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   240
         TabIndex        =   1
         Top             =   1080
         Width           =   360
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ͼ����"
      Height          =   180
      Left            =   480
      TabIndex        =   15
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub INIT()
Me.txtBookID.Locked = False
Me.txtBookAuthor.Enabled = False
Me.txtBookName.Enabled = False
Me.txtBookPrint.Enabled = False
Me.txtLendDate.Enabled = False
Me.txtReaderAddress.Enabled = False
Me.txtReaderID.Enabled = False
Me.txtReaderName.Enabled = False

Me.cmdReturn.Enabled = True

Me.DataGrid1.Visible = False
Me.Adodc1.Visible = False

Me.lblDisplay.Visible = False
Me.lblDisplay.Caption = ""
Me.lblFanJin.Visible = False
Me.lblFanJin.Caption = ""

Me.txtBookAuthor.Text = ""
Me.txtBookID.Text = ""
Me.txtBookName.Text = ""
Me.txtBookPrint.Text = ""
Me.txtLendDate.Text = ""
Me.txtReaderAddress.Text = ""
Me.txtReaderID.Text = ""
Me.txtReaderName.Text = ""

Me.cmdDisplay.Enabled = True
Me.cmdReturn.Enabled = False


Adodc1.ConnectionString = CnStr  '����ADODC�����ݿ�

End Sub


Private Sub cmdDisplay_Click()
Dim questr As String
Me.cmdReturn.Enabled = True
Me.cmdDisplay.Enabled = False

'��ʾͼ����Ϣ
questr = "select ����,������,����   from ͼ�� where ͼ����='" & Me.txtBookID.Text & "'"
Adodc1.RecordSource = questr
Set DataGrid1.DataSource = Adodc1

If Me.Adodc1.Recordset.EOF Then
   MsgBox "û�д�ͼ��"
   INIT
   Exit Sub
End If
Me.txtBookAuthor.Text = Me.DataGrid1.Columns(2).Text
Me.txtBookName.Text = Me.DataGrid1.Columns(0).Text
Me.txtBookPrint.Text = Me.DataGrid1.Columns(1).Text
Adodc1.Refresh
Me.DataGrid1.Refresh
'��ʾ�������ڡ����߱��
questr = "select  ��������,���߱��  from ͼ����� where   ͼ����='" & Me.txtBookID.Text & "'"
Adodc1.RecordSource = questr
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
Me.DataGrid1.Refresh
If Me.DataGrid1.Columns(0) = "" Then
   MsgBox "û�д�ͼ��"
   INIT
   Exit Sub
End If
Me.txtLendDate.Text = ""
Me.txtReaderID.Text = ""
Me.txtLendDate.Text = Me.DataGrid1.Columns(0).Text
Me.txtReaderID.Text = Me.DataGrid1.Columns(1).Text
'��ʾ������
questr = "select  ��������,סַ  from ���� where   ���߱��='" & Me.txtReaderID.Text & "'"
Adodc1.RecordSource = questr
Adodc1.Refresh

Set DataGrid1.DataSource = Adodc1
Me.DataGrid1.Refresh
If Me.DataGrid1.Columns(0) = "" Then
   MsgBox "û�д�ͼ��"
   INIT
   Exit Sub
End If
Me.txtReaderName.Text = Me.DataGrid1.Columns(0).Text
Me.txtReaderAddress.Text = Me.DataGrid1.Columns(1).Text
'�ж��Ƿ��ѹ���
Dim x As Date
Dim y As Date
Dim z As Double

x = Format(Now - 21, "yyyy - mm - dd")
y = Format(Me.txtLendDate, "yyyy - mm - dd")
If x > y Then
   Me.lblDisplay.Visible = True
   Me.lblDisplay.Caption = "�ѳ���, ����:"
   Me.lblDisplay.AutoSize = True
   
   '���㷣��
   
   z = x - y
   Me.lblFanJin.Visible = True
   Me.lblFanJin.Caption = z * Val(FaJin)
   End If
End Sub

Private Sub cmdReturn_Click()
Dim questr As String

'������� ��������
If Me.lblDisplay.Caption <> "" Then
  questr = "insert into ����  values('" & Me.txtReaderID.Text & "','" & Me.txtBookID.Text & "'," & Me.lblFanJin.Caption & ")"
    If DataManage(questr, dbcn, Adodc1) = 0 Then 'ʧ��
      MsgBox ("ʧ��")
      Exit Sub
    End If
End If
'����
questr = "delete from ͼ����� where  ͼ���� = '" & Me.txtBookID.Text & "'"
dbcn.Execute questr

INIT



End Sub

Private Sub Form_Load()
INIT
End Sub


