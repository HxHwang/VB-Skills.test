VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3360
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   4560
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "�ж�����"
      Height          =   615
      Left            =   1800
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "������һ������2����������"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim n As Integer
    Dim i As Integer
    Dim f As Boolean
    n = Text1.Text
    f = True '����n������
    For i = 2 To n - 1
        If n Mod i = 0 Then '���i��n��Լ��
            f = False '�Ʒ��ٶ�
            Exit For  '��ǰ�˳�ѭ��
        End If
    Next i
    If f = True Then '�ж��Ƿ�������
       Label2.Caption = n & "��������"
    Else
        Label2.Caption = n & "����������"
    End If
End Sub
