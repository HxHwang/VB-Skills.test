VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ˮ�ɻ���"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "��ʾˮ�ɻ���"
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   2280
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim a As Integer
    Dim b As Integer
    Dim c As Integer
    Dim n As Integer
    Dim i As Integer
    Print "ˮ�ɻ����У�"
    For i = 100 To 999 'ѭ���ж���λ��
        a = i \ 100 '��λ��
        b = i \ 10 Mod 10 'ʮλ��
        c = i Mod 10 '��λ��
        n = a * a * a + b * b * b + c * c * c
        If n = i Then
            Print i
        End If
    Next i
End Sub
