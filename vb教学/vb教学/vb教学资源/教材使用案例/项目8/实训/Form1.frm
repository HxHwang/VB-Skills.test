VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "��̬�����ļ������ļ�����������"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   1965
   ScaleWidth      =   4950
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "�����ļ�����������"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3840
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4935
      Begin VB.Label Label1 
         Caption         =   "�����뱣���ļ���λ�ü��ļ���"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   2895
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type stu
  stname As String * 10
  num As String
  age As Integer
  addr As String
  End Type
Private Sub Command1_Click()
   CommonDialog1.Filter = "txt(*.txt)|*.txt|doc(*.doc)|*.doc"     '�����ļ�����
   CommonDialog1.ShowSave                                         '����Ի���
   Text1.Text = CommonDialog1.FileName
End Sub

Private Sub Command2_Click()
    If Text1.Text = "" Then                    '�ļ�������Ϊ��
        MsgBox "�ļ�������Ϊ��"
    Else
        Open Text1.Text For Output As #1    '�ڶ�Ӧ��λ���½��ļ�
        MsgBox "�����ļ��ɹ����밴����ʾ����ѧ����Ϣ��"
        Static stud() As stu                          '���徲̬����
        n = InputBox("������ѧ������")             '�������뺯����������
        ReDim stud(n) As stu
        For i = 1 To n
            stud(i).stname = InputBox("����������:")
            stud(i).num = InputBox("�������꼶:")
            stud(i).age = InputBox("����������:")
            stud(i).addr = InputBox("�������ַ:")
            Write #1, stud(i).stname, stud(i).num, stud(i).age, stud(i).addr
        Next i
        Close #1
        MsgBox "������ϣ�"
        End
    End If
End Sub

Private Sub Form_DblClick()
 End
End Sub

