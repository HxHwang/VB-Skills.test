VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "ʵ����ϰ"
   ClientHeight    =   8370
   ClientLeft      =   1830
   ClientTop       =   1425
   ClientWidth     =   10860
   LinkTopic       =   "Form5"
   Picture         =   "������ϰ.frx":0000
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   5640
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   7695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9120
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9120
      Width           =   2535
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6990
      ItemData        =   "������ϰ.frx":4C004
      Left            =   1680
      List            =   "������ϰ.frx":4C011
      TabIndex        =   0
      Top             =   1440
      Width           =   3135
   End
End
Attribute VB_Name = "form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
Set xlapp = CreateObject("Excel.Application") '����EXCEL����
Select Case List1.ListIndex
       Case 0
        Set xlbook = xlapp.Workbooks.Open(App.Path & "\3.1.xls")
        xlapp.Visible = True '����EXCEL����ɼ�
        Set xlsheet = xlbook.Worksheets(1)
        form5.Hide
        frmxt1.Show
       Case 1
        Set xlbook = xlapp.Workbooks.Open(App.Path & "\3.2.xls")
        xlapp.Visible = True '����EXCEL����ɼ�
        Set xlsheet = xlbook.Worksheets(1)
        form5.Hide
        frmxt2.Show
       Case 2
        Set xlbook = xlapp.Workbooks.Open(App.Path & "\3.3.xls")
        xlapp.Visible = True '����EXCEL����ɼ�
        Set xlsheet = xlbook.Worksheets(1)
        form5.Hide
        frmxt3.Show
End Select


End Sub

Private Sub Command3_Click()
Me.Hide
Form3.Show
End Sub

Private Sub Form_Load()
Command2.Enabled = False
Me.WindowState = 2
End Sub

Private Sub List1_Click()
Select Case List1.ListIndex
       Case 0
       Text1.Text = "����1 " & vbCrLf & "��1���ֱ������������������������ĺϼ�����" & vbCrLf & "��2��ʹ�ú����������������������֡������䡢���ֱá����ᡢ���γ��ָ��������ܼ�����"
        Command2.Enabled = True
       Case 1
        Command2.Enabled = True
         Text1.Text = "����2" & vbCrLf & "��1������ʽ������=�۳���-����-��Ӫ�ɱ���������ֲ�Ʒ�������ƽ������ֵ�����������λС����" & vbCrLf & "��2����sheet1�е�ƽ��ֵ���������ȫ�����Ƶ�sheet2�У�����ȫ�����ݰ�����ֵ�Ӹߵ�������" & vbCrLf & "��3����sheet1�е�ƽ��ֵ���������ȫ�����Ƶ�sheet3�У���������ɸѡ���������100Ԫ��С��150Ԫ������?ɸѡ��ָ�ȫ������?" & vbCrLf & "��4����sheet1�е�ƽ��ֵ���������ȫ�����Ƶ�sheet4�У�����������ܣ��ֱ�����յ�������䡢ϴ�»��ľ�Ӫ�ɱ��������ƽ��ֵ?"
       Case 2
         Text1.Text = "����3" & vbCrLf & "��1������ʽ������=��������+Ч�湤�ʣ�����ÿ�˵Ĺ��ʡ�" & vbCrLf & "��2������ʽ��������=����*�����ʣ�����ÿ�˵Ĺ��ʸ����" & vbCrLf & "��3������'����'��'������'�ֱ����ÿ�˵Ĺ����ܶ" & vbCrLf & "��4����������������ƽ��ֵ��"
        Command2.Enabled = True
End Select
       
End Sub
