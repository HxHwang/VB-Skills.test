VERSION 5.00
Begin VB.Form Search 
   Caption         =   "ѧ���ɼ���ѯ"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   6705
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox TextFind 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   26
      Top             =   5760
      Width           =   1575
   End
   Begin VB.ComboBox CobFind 
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "Search.frx":0000
      Left            =   4800
      List            =   "Search.frx":0002
      TabIndex        =   25
      Text            =   "ѧ��"
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton CmdFindPrevious 
      Caption         =   "<<"
      Height          =   375
      Left            =   5640
      TabIndex        =   24
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton CmdFindNext 
      Caption         =   ">>"
      Height          =   375
      Left            =   4800
      TabIndex        =   23
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton CmdBack 
      Caption         =   "�� ��"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   22
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton CmdFind 
      Caption         =   "�� ѯ"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   21
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton CmdLast 
      Caption         =   "���һ��"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   20
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton CmdPrevious 
      Caption         =   "��һ��"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   19
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   "��һ��"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   18
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton CmdFirst 
      Caption         =   "��һ��"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   17
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      DataField       =   "�ɼ�"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   16
      Top             =   6120
      Width           =   2415
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "�ον�ʦ"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   15
      Top             =   5382
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      DataField       =   "�γ���"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   14
      Top             =   4645
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      DataField       =   "�γ̺�"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   13
      Top             =   3908
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      DataField       =   "����"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   12
      Top             =   3171
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      DataField       =   "�Ա�"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   11
      Top             =   2434
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataField       =   "����"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   10
      Top             =   1697
      Width           =   2415
   End
   Begin VB.TextBox TxtID 
      Alignment       =   2  'Center
      DataField       =   "ѧ��"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   9
      Top             =   960
      Width           =   2415
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "ѧ���ɼ���Ϣ��.mdb"
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "�ܲ�ѯ��"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   8
      Top             =   6240
      Width           =   1065
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "�ον�ʦ"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   195
      TabIndex        =   7
      Top             =   5490
      Width           =   1395
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "�γ���"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   6
      Top             =   4755
      Width           =   1065
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "�γ̺�"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   5
      Top             =   4005
      Width           =   1065
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   4
      Top             =   3270
      Width           =   1065
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   1065
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   1785
      Width           =   1065
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "ѧ  ��"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Top             =   1035
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ѧ���ɼ���ѯ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msg
Private Sub CmdFirst_Click() '����һ������ť����
Data1.Recordset.MoveFirst
CmdFirst.Enabled = False
CmdPrevious.Enabled = False
If CmdNext.Enabled = False Then CmdNext.Enabled = True
If CmdLast.Enabled = False Then CmdLast.Enabled = True

End Sub

Private Sub CmdNext_Click() '����һ������ť����
Data1.Recordset.MoveNext
If Data1.Recordset.EOF Then
 Data1.Recordset.MoveLast
  MsgBox "�Ѿ��ﵽ���һ����¼��", 65, "��ʾ"
  CmdNext.Enabled = False
  CmdLast.Enabled = False
Else
  If CmdFirst.Enabled = False Then CmdFirst.Enabled = True
  If CmdPrevious.Enabled = False Then CmdPrevious.Enabled = True
End If

End Sub

Private Sub CmdPrevious_Click() '����һ������ť����
Data1.Recordset.MovePrevious
 If Data1.Recordset.BOF Then
  Data1.Recordset.MoveFirst
  MsgBox "�Ѿ��ﵽ��һ����¼��", 65, "��ʾ"
  CmdFirst.Enabled = False
  CmdPrevious.Enabled = False
Else
  If CmdNext.Enabled = False Then CmdNext.Enabled = True
  If CmdLast.Enabled = False Then CmdLast.Enabled = True

 End If
End Sub

Private Sub CmdLast_Click() '�����һ������ť����
Data1.Recordset.MoveLast
CmdNext.Enabled = False
CmdLast.Enabled = False
If CmdFirst.Enabled = False Then CmdFirst.Enabled = True
If CmdPrevious.Enabled = False Then CmdPrevious.Enabled = True
  
End Sub

Private Sub CmdFind_Click() '����ѯ����ť����
    If TextFind.Text = "" Then
        MsgBox "�������ѯ���ݣ�", 48, "��ʾ"
        Exit Sub
    End If
    If CobFind.Text = "����" Then
        msg = "����=" & "'" & TextFind.Text & "'"
        Data1.Recordset.FindFirst "����=" & "'" & TextFind.Text & "'"
    ElseIf CobFind.Text = "ѧ��" Then
        msg = "ѧ�� Like" & "'" & TextFind.Text & "'"
        Data1.Recordset.FindFirst "ѧ�� Like" & "'" & TextFind.Text & "'"
    ElseIf CobFind.Text = "�γ���" Then
        msg = "�γ���=" & "'" & TextFind.Text & "'"
        Data1.Recordset.FindFirst "�γ���=" & "'" & TextFind.Text & "'"
    ElseIf CobFind.Text = "�γ̺�" Then
        msg = "�γ̺� Like" & "'" & TextFind.Text & "'"
        Data1.Recordset.FindFirst "�γ̺� Like" & "'" & TextFind.Text & "'"
    End If
        If Data1.Recordset.NoMatch Then
        MsgBox "��¼�����ڣ�", 64, "��ʾ"
    End If

End Sub

Private Sub CmdFindNext_Click() '��>>�����ҷ�����������һ����¼
Data1.Recordset.FindNext msg
End Sub

Private Sub CmdFindPrevious_Click() '��<<�����ҷ�����������һ����¼
Data1.Recordset.FindPrevious msg
End Sub

Private Sub CmdBack_Click() '���رա���ť����
    Unload Search
    Main.Show
End Sub

Private Sub Form_Load()
CmdFirst.Enabled = False
CmdPrevious.Enabled = False
End Sub
