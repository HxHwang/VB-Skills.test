VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�ж�����"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "�ȼ�����"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʾ����"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'֪ʶ��ע��Sgn(x) ��x�ķ��� x>0 ����1��x=0 ����0��x<0 ����-1��
Private Sub Command1_Click()
Dim number As Integer

number = Val(InputBox("����������", "�ж�����", "45")) '��ΪinputboxĬ�Ϸ���stringŶ

If number > 0 Then '��߶˵Ļ� ������Ը�Ϊ if Sgn(number)=1 then
    Print number
Else
    MsgBox "����������"
End If
End Sub

Private Sub Command2_Click()
Dim number As Integer
number = Val(InputBox("����������", "�ȼ�����", "8"))
'���������⣬����select case ����
Select Case number
    Case 1 To 4
        Print "D"
    Case 5 To 10
        Print "C"
    Case 11 To 14
        Print "B"
    Case Else '���϶�������������ֱ��ִ�����
        Print "A"
End Select

End Sub
