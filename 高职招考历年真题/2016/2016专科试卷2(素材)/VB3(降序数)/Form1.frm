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
   Begin VB.CommandButton Command1 
      Caption         =   "������"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'i ��������10λ��
' cou�Ǽ����õ� ��� couΪ10 �� ������ʾ
' ge��shi ���� ȡʮλ���ĸ�λ��ʮλ
Dim i As Integer, cou As Integer
Dim ge As Integer, shi As Integer
Print "����2λ�Ľ������У�"
For i = 10 To 99
    ge = i Mod 10   '��10ȡ��ɵõ���λ
    shi = i \ 10 Mod 10  ' �ȳ���10 Ȼ���� ȡ�� ���ɵõ�ʮλ
    If shi >= ge Then  '���ս�����㷨 a>=b  ��ν��λ��������λ ���� ʮλ���ڸ�λ����˼��
        Print i;   ' �ֺ��� ÿ���һ������ �� ��ͬһ����� ���� ��ֱ�ӵ���һ��
        cou = cou + 1  '������Ϊ�� ͳ�� ���е�����
        If cou = 10 Then ' ����10�ˣ���ô�ҿ��Ի�������
            cou = 0 '���е�ͬʱ�������е�������������Ϊ0�������ͻᷴ�����У�
            Print '��Ϊ�����i�����и��ֺţ����� һ������Ҫ������ �ͱ���������յ�print
        End If
    End If
Next i
'��Ϊ��Ŀ���㷨��ʾ��������˵ ���ѣ�

'С֪ʶ�� /   ��  \  ������

' / �Ǹ����
' \ ������

'���� 7/2=3.5    7\2=3
End Sub

