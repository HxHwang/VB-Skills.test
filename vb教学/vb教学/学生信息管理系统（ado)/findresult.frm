VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form findresult 
   Caption         =   "ѧ���ɼ���ѯ"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   8490
   StartUpPosition =   3  '����ȱʡ
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "findresult.frx":0000
      Height          =   2625
      Left            =   600
      TabIndex        =   6
      Top             =   2640
      Width           =   7050
      _ExtentX        =   12435
      _ExtentY        =   4630
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "ѧ��"
         Caption         =   "ѧ��"
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
         DataField       =   "����"
         Caption         =   "����"
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
      BeginProperty Column02 
         DataField       =   "�Ա�"
         Caption         =   "�Ա�"
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
      BeginProperty Column03 
         DataField       =   "�γ���"
         Caption         =   "�γ���"
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
      BeginProperty Column04 
         DataField       =   "�ɼ�"
         Caption         =   "�ɼ�"
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
            Alignment       =   2
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1094.74
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "ѡ���ѯ����"
      Height          =   2055
      Left            =   600
      TabIndex        =   3
      Top             =   480
      Width           =   2535
      Begin VB.OptionButton optid 
         Caption         =   "��ѧ��"
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton optname 
         Caption         =   "���γ���"
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   1200
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox txtfind 
      Height          =   615
      Left            =   3960
      TabIndex        =   0
      Top             =   840
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4320
      Top             =   5040
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=student.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=student.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"findresult.frx":0015
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
Attribute VB_Name = "findresult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim sqlstr As String
If Trim(txtfind.Text) = "" Then
  MsgBox "���������ѯ������", vbOKOnly + vbCritical, "��ѯ"
Else
  If optid.Value = True Then
   sqlstr = "select studentinfo.ѧ��,studentinfo.����,studentinfo.�Ա�,courseinfo.�γ���,result.�ɼ� from studentinfo,courseinfo,result where studentinfo.ѧ��=result.ѧ�� and courseinfo.�γ̱��=result.�γ̱�� and studentinfo.ѧ��='" & Trim(txtfind.Text) & "'"
   Adodc1.RecordSource = sqlstr
   Adodc1.Refresh
   If Adodc1.Recordset.EOF Then
     MsgBox "��ѯ��¼Ϊ�գ�", vbOKOnly + vbInformation, "��ѯ���"
   End If
   
  Else
   sqlstr = "select studentinfo.ѧ��,studentinfo.����,studentinfo.�Ա�,courseinfo.�γ���,result.�ɼ� from studentinfo,courseinfo,result where studentinfo.ѧ��=result.ѧ�� and courseinfo.�γ̱��=result.�γ̱�� and courseinfo.�γ���='" & Trim(txtfind.Text) & "'"
   Adodc1.RecordSource = sqlstr
   Adodc1.Refresh
    If Adodc1.Recordset.EOF Then
     MsgBox "��ѯ��¼Ϊ�գ�", vbOKOnly + vbInformation, "��ѯ���"
   End If
  End If
End If
End Sub

Private Sub Command2_Click()
mainfrm.Show
findresult.Hide
End Sub

