VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   3990
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdInput 
      Caption         =   "����ѧ����Ϣ"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   1920
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdInput_Click()
'��ʾ"ѧ��������Ϣ"�Ի���
Form2.Show 0
End Sub

Private Sub Form_Load()
'�ڴ�������ʾ"����    �Ա�  ��ѧʱ��      רҵ"
Form1.Print "����    �Ա�  ��ѧʱ��      רҵ"
End Sub

