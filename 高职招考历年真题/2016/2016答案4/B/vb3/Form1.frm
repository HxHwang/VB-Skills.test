VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "������"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "������"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim i As Integer
    Dim t As Integer
    Dim a As Integer
    Dim b As Integer
    Print "����2λ�Ľ������У�"
    For i = 10 To 99  '��������λ������ѭ���ж�
        a = i \ 10 '���ʮλ�ϵ���
        b = i Mod 10 '�����λ�ϵ���
        If a >= b Then
            Print i;
            t = t + 1
            If t = 10 Then
                Print  '����
                t = 0 '����������Ϊ0
            End If
        End If
    Next i
End Sub
