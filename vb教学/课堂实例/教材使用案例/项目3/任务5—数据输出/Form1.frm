VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�������"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   BeginProperty Font 
      Name            =   "����"
      Size            =   15.75
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   6870
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "���"
      Height          =   615
      Left            =   5040
      TabIndex        =   0
      Top             =   5040
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    '�������
    Dim i As Integer, j As Integer      '�������α���
    Dim S As String                     '����䳤�ַ���
    Dim S1 As String * 10, S2 As String * 5, S3 As String * 1       '���嶨���ַ���
    
    i = 2:    j = -5
    Print "�����ֵ����:"   '����ַ�������
    Print "i="; i           '�����ֵ����
    Print "j="; j
    Print "i+j="; i + j     '���������ʽ��ֵ
    Print                   '���һ������
    
    S = "abcde"
    Print "ʹ�÷ֺ�����䳤�ַ������ݣ�"
    Print S; "ABCDE"        'ʹ�÷ֺ�����䳤�ַ����������ַ�������
    Print
    Print "ʹ�ö�������䳤�ַ������ݣ�"
    Print S, "ABCDE"         'ʹ�ö�������䳤�ַ����������ַ�������
    Print
    
    S1 = "xyz": S2 = "xyz": S3 = "xyz"
    Print "ʹ�÷ֺ���������ַ�������S1,S2,S3��"
    Print S1; S2;           'β���ӷֺű�ʾ��һ���������������
    Print S3
End Sub


