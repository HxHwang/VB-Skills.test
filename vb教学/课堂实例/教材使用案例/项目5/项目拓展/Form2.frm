VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "ѧ��������Ϣ"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4485
   LinkTopic       =   "Form2"
   ScaleHeight     =   1965
   ScaleWidth      =   4485
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton OKButton 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   1440
      Width           =   1815
   End
   Begin VB.ComboBox cmbYear 
      Height          =   300
      Left            =   3240
      TabIndex        =   8
      Text            =   "Combo2"
      Top             =   840
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "�Ա�"
      Height          =   615
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   1815
      Begin VB.OptionButton optFemale 
         Caption         =   "Ů"
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton optMale 
         Caption         =   "��"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.ComboBox cmbDepart 
      Height          =   300
      Left            =   720
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Text            =   "������"
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "��ѧʱ��"
      Height          =   180
      Left            =   2400
      TabIndex        =   7
      Top             =   960
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "רҵ"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   360
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
'Ϊ"רҵ"��Ͽ������Ŀ��������Ĭ����
cmbDepart.AddItem "�����"
cmbDepart.AddItem "���"
cmbDepart.AddItem "�г�Ӫ��"
cmbDepart.AddItem "����"
cmbDepart.ListIndex = 0
'Ϊ"��ѧʱ��"��Ͽ������Ŀ��������Ĭ����
cmbYear.AddItem "2001��9��"
cmbYear.AddItem "2002��9��"
cmbYear.AddItem "2003��9��"
cmbYear.ListIndex = 2
'����Ĭ���Ա�
optMale.Value = True
End Sub

Private Sub OKButton_Click()
'����һ�����ڴ洢�Ա���ַ���
Dim man As String
'������ѡ���Ա𣬽��Ա𸳸���������ַ���
If optMale Then
man = "��"
Else
man = "Ů"
End If
'���������ѧ��������Ϣ��ʾ����������
Form1.Print txtName.Text + "   " + man + "    " + cmbYear.Text + _
"    " + cmbDepart.Text
'���ضԻ���
Form2.Hide
End Sub


