VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmLend 
   Caption         =   "����"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6630
   ScaleWidth      =   8910
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   7680
      Top             =   480
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
      Caption         =   "Adodc2"
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
      Height          =   495
      Left            =   6840
      TabIndex        =   23
      Top             =   240
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
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
   Begin VB.CommandButton cmdLend 
      Caption         =   "����"
      Height          =   495
      Left            =   4320
      TabIndex        =   13
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "��ʾ��Ϣ"
      Height          =   495
      Left            =   3000
      TabIndex        =   12
      Top             =   360
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "ͼ����Ϣ"
      Height          =   2175
      Left            =   120
      TabIndex        =   11
      Top             =   4320
      Width           =   8775
      Begin VB.TextBox txtBookISBN 
         Height          =   375
         Left            =   5520
         TabIndex        =   21
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtBookPrint 
         Height          =   375
         Left            =   5520
         TabIndex        =   20
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox txtBookAuthor 
         Height          =   375
         Left            =   840
         TabIndex        =   19
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox txtBookName 
         Height          =   375
         Left            =   840
         TabIndex        =   18
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label lblBookInfo 
         BackColor       =   &H0000FFFF&
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
         Height          =   495
         Left            =   3360
         TabIndex        =   22
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "ISBN"
         Height          =   180
         Left            =   4800
         TabIndex        =   17
         Top             =   1200
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Left            =   4800
         TabIndex        =   16
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   360
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   600
      Top             =   6600
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
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
   Begin VB.Frame frame1 
      Caption         =   "������Ϣ"
      Height          =   3015
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   8775
      Begin VB.Frame Frame2 
         Caption         =   "�ѽ�ͼ��"
         Height          =   2055
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   8415
         Begin MSDataGridLib.DataGrid DGLendedBook 
            Height          =   1695
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   2990
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
      End
      Begin VB.TextBox txtReaderAddrass 
         Height          =   375
         Left            =   3840
         TabIndex        =   8
         Top             =   240
         Width           =   3495
      End
      Begin VB.TextBox txtReaderName 
         Height          =   375
         Left            =   840
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "סַ"
         Height          =   255
         Left            =   3240
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "����"
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.TextBox txtBookID 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtReaderID 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ͼ����"
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "���߱��"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmLend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDisplay_Click()
Me.cmdDisplay.Enabled = False
Me.cmdLend.Enabled = True
Dim str As String
Dim QueryInt As Integer
'��ѯ����
str = "select ��������,סַ from ���� where ���߱��='" & Me.txtReaderID.Text & "'"
QueryInt = DataQuery(str, Adodc1, Me.DGLendedBook)
If QueryInt = 0 Then
 MsgBox "����ʧ��"
 INIT
 Exit Sub
End If
If QueryInt = 1 Then
   Me.txtReaderAddrass.Text = Me.DGLendedBook.Columns(1).Text
   Me.txtReaderName.Text = Me.DGLendedBook.Columns(0).Text
End If

'��ѯͼ��
str = "select  ����,����,������,ISBN from ͼ�� where ͼ���� =  '" & Me.txtBookID.Text & "'"

QueryInt = DataQuery(str, Adodc1, Me.DGLendedBook)
If QueryInt = 0 Or QueryInt = 2 Then
 MsgBox "����ʧ��"
 INIT
 Exit Sub
End If
Me.txtBookAuthor.Text = Me.DGLendedBook.Columns(1).Text
Me.txtBookISBN.Text = Me.DGLendedBook.Columns(3).Text
Me.txtBookName.Text = Me.DGLendedBook.Columns(0).Text
Me.txtBookPrint.Text = Me.DGLendedBook.Columns(2).Text

'��ѯ��ͼ���Ƿ񱻽��
str = "select * from ͼ����� where ͼ����='" & Me.txtBookID.Text & "'"
QueryInt = DataQuery(str, Adodc1, Me.DGLendedBook)
If QueryInt = 0 Then
  MsgBox "����ʧ��"
  INIT
  Exit Sub
End If
If QueryInt = 1 Then
  Me.lblBookInfo.Visible = True
  Me.lblBookInfo.Caption = "��ͼ���ѱ����"
  Me.lblBookInfo.AutoSize = True
End If


'��ѯ�����ѽ�ͼ��
str = "select ͼ�����.ͼ����,���� ,����,������ ,�������� from ͼ����� ,ͼ��  where ͼ�����.ͼ����=ͼ��.ͼ���� and  ���߱��='" & Me.txtReaderID.Text & "'"
QueryInt = DataQuery(str, Adodc1, Me.DGLendedBook)
If QueryInt = 0 Then
 MsgBox "����ʧ��"
 Exit Sub
End If
If Me.Adodc1.Recordset.RecordCount > MaxBook Then  '�������ֵ
   MsgBox "�װ��Ķ��ߣ������ͼ���ѳ��������Ŀ���뾡�컹�飬�ټ������飬лл"
   INIT
   Exit Sub
End If


'��ͼ�鱻���
If Me.lblBookInfo.Caption <> "" Then
   Me.cmdDisplay.Enabled = True
   Me.cmdLend.Enabled = False
   Me.txtBookID.Text = ""
   Me.txtReaderID.Text = ""
End If

End Sub

Private Sub cmdLend_Click()
Me.cmdDisplay.Enabled = True
Me.cmdLend.Enabled = False
Dim str As String
Dim x As Date
x = Format(Now, "yyyy-mm-dd")

str = "insert into ͼ����� values('" & Me.txtBookID.Text & "','" & Me.txtReaderID.Text & "'," & x & ")"
Dim opt As Integer
opt = DataManage(str, dbcn, Adodc1)
If opt = 0 Then
   MsgBox "����ʧ��"
   INIT
   Exit Sub
End If
MsgBox "�����ɹ�"
   
INIT
End Sub

Private Sub Form_Load()
INIT
End Sub
Private Sub INIT()
Me.txtBookAuthor.Enabled = False
Me.txtBookID.Enabled = True
Me.txtBookISBN.Enabled = False
Me.txtBookName.Enabled = False
Me.txtBookPrint.Enabled = False
Me.txtReaderAddrass.Enabled = False
Me.txtReaderID.Enabled = True
Me.txtReaderName.Enabled = False

Me.txtReaderAddrass.Text = ""
Me.txtReaderName.Text = ""
Me.txtBookAuthor.Text = ""
Me.txtBookID.Text = ""
Me.txtBookISBN.Text = ""
Me.txtBookName.Text = ""
Me.txtBookPrint.Text = ""




Me.DGLendedBook.AllowAddNew = False
Me.DGLendedBook.AllowDelete = False
Me.DGLendedBook.AllowUpdate = False

Me.lblBookInfo.Visible = False
Me.lblBookInfo.Caption = ""

Me.cmdDisplay.Enabled = True
Me.cmdLend.Enabled = False

Me.Adodc1.Visible = False

Me.Adodc1.ConnectionString = CnStr

Me.Adodc2.Visible = False
Me.Adodc2.ConnectionString = CnStr
Me.DataGrid1.Visible = False



End Sub

