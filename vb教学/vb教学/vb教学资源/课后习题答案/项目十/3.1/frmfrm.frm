VERSION 5.00
Begin VB.Form FrmEide 
   Caption         =   "ѧ���ɼ���"
   ClientHeight    =   6795
   ClientLeft      =   1110
   ClientTop       =   450
   ClientWidth     =   5340
   LinkTopic       =   "Form2"
   ScaleHeight     =   6795
   ScaleWidth      =   5340
   Begin VB.CommandButton cmdClose 
      Caption         =   "�ر�(&C)"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   3840
      TabIndex        =   12
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "����(&U)"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "ˢ��(&R)"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   3840
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ɾ��(&D)"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   2040
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "���(&A)"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Connect         =   "Access"
      DatabaseName    =   "ѧ���ɼ����ݿ�.mdb"
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ѧ���ɼ���"
      Top             =   6300
      Width           =   5340
   End
   Begin VB.TextBox txtFields 
      DataField       =   "�ɼ�"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   3
      Left            =   1920
      TabIndex        =   7
      Top             =   5400
      Width           =   2535
   End
   Begin VB.TextBox txtFields 
      DataField       =   "�γ���"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   2
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   5
      Top             =   4280
      Width           =   2535
   End
   Begin VB.TextBox txtFields 
      DataField       =   "����"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   1
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   3
      Top             =   3160
      Width           =   2535
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ѧ��"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   0
      Left            =   1920
      TabIndex        =   1
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "��  ��:"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   360
      TabIndex        =   6
      Top             =   5640
      Width           =   1125
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "�γ���:"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   367
      TabIndex        =   4
      Top             =   4520
      Width           =   1110
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "��  ��:"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   3400
      Width           =   1125
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "ѧ  ��:"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   2280
      Width           =   1125
   End
End
Attribute VB_Name = "FrmEide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
  Data1.Recordset.AddNew
End Sub

Private Sub cmdDelete_Click()
  '���ɾ����¼�������һ����¼
  '��¼���¼����Ψһ�ļ�¼
  Data1.Recordset.Delete
  Data1.Recordset.MoveNext
End Sub

Private Sub cmdRefresh_Click()
  '����Զ��û�Ӧ�ó��������Ҫ��
  Data1.Refresh
End Sub

Private Sub cmdUpdate_Click()
  Data1.UpdateRecord
  Data1.Recordset.Bookmark = Data1.Recordset.LastModified
End Sub

Private Sub cmdClose_Click()
 End
 
End Sub



Private Sub Data1_Error(DataErr As Integer, Response As Integer)
  '����Ƿ��ô��������ĵط�
  '�������Դ���ע�͵���һ�д���
  '����벶׽������������Ӵ��������
  MsgBox "���ݴ����¼����д���" & Error$(DataErr)
  Response = 0  '���Դ���
End Sub

Private Sub Data1_Reposition()
  Screen.MousePointer = vbDefault
  On Error Resume Next
  '�⽫��ʾ��ǰ��¼λ��
  'Ϊ��̬���Ϳ���
  Data1.Caption = "��¼��" & (Data1.Recordset.AbsolutePosition + 1)
  '���� Table ���󣬵���¼��������ʹ���������ʱ��
  '�������� Index ����
  'Data1.Caption = "��¼��" & (Data1.Recordset.RecordCount * (Data1.Recordset.PercentPosition * 0.01)) + 1
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)
  '���Ƿ�����֤����ĵط�
  '������Ķ�������ʱ����������¼�
  Select Case Action
    Case vbDataActionMoveFirst
    Case vbDataActionMovePrevious
    Case vbDataActionMoveNext
    Case vbDataActionMoveLast
    Case vbDataActionAddNew
    Case vbDataActionUpdate
    Case vbDataActionDelete
    Case vbDataActionFind
    Case vbDataActionBookmark
    Case vbDataActionClose
  End Select
  Screen.MousePointer = vbHourglass
End Sub

