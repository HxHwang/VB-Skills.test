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
      Caption         =   "�ж�"
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "��ѧ�ɼ�"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "���ĳɼ�"
      Height          =   495
      Left            =   720
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

Dim a As Integer, b As Integer

a = Val(Text1.Text)
b = Val(Text2.Text)

' �жϷ�Χ�Ƿ�����
If a < 0 Or a > 100 Or b < 0 Or b > 100 Then
    MsgBox "������0~100��Χ�ڵ�����"
    Exit Sub
End If

If a >= 90 And b >= 90 Then
    res = "��õ��ѧ��"
ElseIf a = 100 Or b = 100 Then
    res = "��õ��ѧ��"
Else
    res = "û�н�ѧ��"
End If

MsgBox res




End Sub
