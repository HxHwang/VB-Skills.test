VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "ͼ�����ϵͳ"
   ClientHeight    =   8235
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10740
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '����ȱʡ
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Align           =   1  'Align Top
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   975
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   1720
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
      Align           =   1  'Align Top
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   1720
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
   Begin VB.Menu Reader 
      Caption         =   "���ߵ���"
      Begin VB.Menu ReaderManage 
         Caption         =   "���ߵ�������"
      End
      Begin VB.Menu ReaderQuery 
         Caption         =   "���ߵ�����ѯ"
      End
   End
   Begin VB.Menu Book 
      Caption         =   "ͼ�鵵��"
      Begin VB.Menu BookManage 
         Caption         =   "ͼ�鵵������"
      End
      Begin VB.Menu BookQuery 
         Caption         =   "ͼ�鵵����ѯ"
      End
   End
   Begin VB.Menu Operation 
      Caption         =   "ҵ�����"
      Begin VB.Menu Lend 
         Caption         =   "���鴦��"
      End
      Begin VB.Menu Return 
         Caption         =   "���鴦��"
      End
   End
   Begin VB.Menu System 
      Caption         =   "ϵͳ"
      Begin VB.Menu Quit 
         Caption         =   "�˳�"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub BookManage_Click()
frmBook.Show
    frmBook.SSTab1.Tab = 0
    frmBook.SetFocus
End Sub

Private Sub BookQuery_Click()
 frmBook.Show
    frmBook.SSTab1.Tab = 1
    frmBook.SetFocus
End Sub

Private Sub Lend_Click()
    frmLend.Show
   
    frmLend.SetFocus
End Sub

Private Sub MDIForm_Load()
OpenData
'ȡ��ϵͳ����
On Error GoTo erp:

Me.DataGrid1.Visible = False
Me.Adodc1.Visible = False

Adodc1.ConnectionString = CnStr
Adodc1.RecordSource = "select * from ���� "

Set DataGrid1.DataSource = Adodc1

Day = Me.DataGrid1.Columns(0).Text
FaJin = Me.DataGrid1.Columns(1).Text
MaxBook = Me.DataGrid1.Columns(2).Text
Exit Sub
erp:
 MsgBox ("��ȡ����ʧ��")
 
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
CloseData
End Sub

Private Sub Quit_Click()
End

End Sub

Private Sub ReaderManage_Click()
    frmReader.Show
    frmReader.SSTab1.Tab = 0
    frmReader.SetFocus
End Sub

Private Sub ReaderQuery_Click()
    frmReader.Show
    frmReader.SSTab1.Tab = 1
    frmReader.SetFocus
End Sub

Private Sub Return_Click()
    frmReturn.Show
  
    frmReturn.SetFocus
End Sub
