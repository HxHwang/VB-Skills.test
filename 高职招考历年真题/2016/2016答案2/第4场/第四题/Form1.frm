VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command5 
      Caption         =   "�˳�"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "��֤����"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��ʼ��"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "b"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "a"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String '������¼a��ť��b��ť������
Private Sub Command1_Click()
Dim a As String
a = Command1.Caption 'Ҳ��ֱ��д a = "a"
str = str & a '�ۼӵ�Ч��
End Sub

Private Sub Command2_Click()
Dim b As String
b = Command2.Caption ' Ҳ��ֱ��д b = "b"
str = str & b '�ۼӵ�Ч��
End Sub

Private Sub Command3_Click()
str = "" '����㵥����a��ť��b��ť������
End Sub

Private Sub Command4_Click()
'�ж� �㵥����a��ť��b��ť�洢��str�ַ��� �Ƿ�Ϊ'abab'
If str = "abab" Then
    Form2.Show '��������2
Else
    MsgBox "�����������������"
    str = ""
End If
End Sub

Private Sub Command5_Click()
End
End Sub
