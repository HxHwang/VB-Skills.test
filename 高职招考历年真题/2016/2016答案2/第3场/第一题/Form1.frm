VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   4815
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��(&K)"
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "�ַ�(&C)"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "��(&Y)"
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   2400
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label1 
      Height          =   1095
      Left            =   840
      TabIndex        =   1
      Top             =   960
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'�������ʽ��Int(Rnd*(n-m+1)+m)
'֪ʶ�㣺String(n,�ַ���) �����ַ����ĵ�һ���ַ�n��
'        Chr(ASCII��) ������Ӧ����ĸ������
'        ASC(��ĸ������) ����ASCII��
Text1 = Int(Rnd * 26 + 65)
If Option1 = True Then
   Label1 = String(Val(Text1), "��")
Else
    Label1 = Chr(Text1)
End If
End Sub
