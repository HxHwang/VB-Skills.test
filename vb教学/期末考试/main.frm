VERSION 5.00
Begin VB.Form main 
   Caption         =   "��ĩ����"
   ClientHeight    =   6510
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9210
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   9210
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Menu Mnukaoti 
      Caption         =   "����"
      Begin VB.Menu Mnunum1 
         Caption         =   "��һ��"
      End
      Begin VB.Menu Mnunum2 
         Caption         =   "�ڶ���"
      End
      Begin VB.Menu Mnunum3 
         Caption         =   "������"
      End
      Begin VB.Menu Mnunum4 
         Caption         =   "������"
      End
      Begin VB.Menu Mnunum5 
         Caption         =   "������"
      End
   End
   Begin VB.Menu Mnuexit 
      Caption         =   "�˳�"
      Begin VB.Menu Mnutuichu 
         Caption         =   "�˳�����"
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Mnunum1_Click()
Form1.Show
End Sub

Private Sub Mnunum2_Click()
Form2.Show
End Sub

Private Sub Mnunum3_Click()
Form3.Show
End Sub

Private Sub Mnunum4_Click()
Form4.Show
End Sub

Private Sub Mnunum5_Click()
Form6.Show
End Sub

Private Sub Mnutuichu_Click()
Dim answer As Integer
answer = MsgBox("ȷ��Ҫ�˳���ǰ������", vbOKCancel + vbQuestion, "�˳�")
If answer = vbOK Then
   End
Else
  Exit Sub
End If
End Sub
