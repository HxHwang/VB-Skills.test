VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   5730
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "��ӡ"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "������һ��������1-20��"
      Height          =   495
      Left            =   2640
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
Private Sub Command1_Click()
Cls ' ÿ�ε����������һ������
Dim i As Integer, j As Integer
Dim n As Integer
n = Val(Text1.Text)

If n > 20 Or n <= 0 Then
    MsgBox "������1-20��Χ�ڵ�����"
    Exit Sub ' �˳�sub����
End If


For i = 1 To n Step 1
    For j = 1 To n Step 1
        ' ���Խ����ж�
        If i = j Then
            Print 0;
        Else
            Print 1;
        End If
    Next j
    Print ' ����
Next i


End Sub
