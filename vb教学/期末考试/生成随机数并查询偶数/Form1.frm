VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   6825
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "���������"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      TabIndex        =   1
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��ѯż��"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      TabIndex        =   0
      Top             =   5160
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(0 To 20) As Integer
Dim j, m, n As Integer



Private Sub Command1_Click()
Command2.Enabled = True
FontSize = 24 '�����ı���ʾ����
Print  '�ڴ����ϴ�������
Print "������20�������Ϊ��"
For j = 1 To 20
 a(j) = CInt(Rnd * 100) '����100���ڵ������
 Print a(j);  '�ڴ�������ʾû�������
 If j Mod 5 = 0 Then  '������5�������Ϊһ��
  Print
 End If
Next j
End Sub



Private Sub Command2_Click()
n = 0
Print "����ż��Ϊ��"
For j = 1 To 20
m = Sushu(a(j)) '���ú����ж�a(j)���������Ƿ�Ϊ����
If m = 0 Then  '�ڴ����ϴ�ӡ����
Print a(j);
n = n + 1
End If
If n Mod 5 = 0 And n <> 0 Then  '����ÿ5������Ϊһ��
n = 0
Print
End If
Next j

End Sub
