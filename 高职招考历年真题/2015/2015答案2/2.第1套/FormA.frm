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
   Begin VB.CommandButton Command2 
      Caption         =   "ʮλ��"
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����ֵ"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim jue As Integer
jue = Text1.Text
Text2.Text = Abs(jue) 'ʹ��abs(n) ����ʹnתΪ����ֵ
End Sub

Private Sub Command2_Click()
Dim n As Integer
Dim shi As Integer
n = Val(Text1.Text)
shi = n \ 10 '�мǣ�/��\������  /���Ǹ���� \�������� ѧ����ô��
Text2.Text = shi
End Sub
