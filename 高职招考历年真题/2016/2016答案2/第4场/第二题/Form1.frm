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
      Caption         =   "���Լ��"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
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
Dim x As Integer, y As Integer
Dim i As Integer, max As Integer
x = Val(Text1)
y = Val(Text2)
For i = 1 To x
    If x Mod i = 0 And y Mod i = 0 Then
        'Print "��Լ����" & i
        If i > max Then
            max = i
        End If
    End If
Next i
Label1 = "���Լ����" & max
End Sub
'����� ��Ҫ���,���Լ����һ�����ܱ�һ��������
'���� 4 ���Ա�1��2��4����  2 ���Ա� 1��2����
'    ��ô2��4�Ĺ�Լ������ 1��2  ���ľ���2 ����������Լ��
'��������һ�����ܱ���Щ������

Private Sub Form_Load()
Text1 = ""
Text2 = ""
Label1 = ""
End Sub
