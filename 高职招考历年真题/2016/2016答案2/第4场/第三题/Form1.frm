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
      Caption         =   "��ӡ"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "������һ��������1-20��"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim num As Integer
Dim i As Integer, j As Integer
num = Val(Text1) ' ��Text1���ı�תΪ��ֵ����ֵ��num �������
'����������Ĳ�ͬ���֣�ѭ������Ҳ�᲻ͬ
For i = 1 To num 'i���Ƶ�������
    For j = 1 To num 'j���Ƶ�������
        Select Case j  ' ���Ա��ʽ
            Case Is >= i '��һ������Ҫ��������ᣡ˵�����... ...
                Print "1"; Space(1);
            Case Else
                Print "0"; Space(1);
        End Select
    Next j
    Print
Next i
End Sub
