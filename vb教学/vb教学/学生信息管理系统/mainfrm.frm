VERSION 5.00
Begin VB.Form mainfrm 
   BackColor       =   &H00E0E0E0&
   Caption         =   "ѧ����Ϣ����ϵͳ"
   ClientHeight    =   5310
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   8550
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Menu filemenu 
      Caption         =   "�ļ�"
      Begin VB.Menu addusermenu 
         Caption         =   "����û�"
      End
      Begin VB.Menu exitmenu 
         Caption         =   "�˳�"
      End
   End
   Begin VB.Menu xuejimenu 
      Caption         =   "ѧ������"
      Begin VB.Menu personinfomenu 
         Caption         =   "�鿴ѧ����Ϣ"
      End
      Begin VB.Menu addinfomenu 
         Caption         =   "���ѧ����Ϣ"
      End
      Begin VB.Menu editinfomenu 
         Caption         =   "�޸�ѧ����Ϣ"
      End
      Begin VB.Menu findinfomenu 
         Caption         =   "����ѧ����Ϣ"
      End
   End
   Begin VB.Menu chengjimenu 
      Caption         =   "�ɼ�����"
      Begin VB.Menu addscoremenu 
         Caption         =   "��ӳɼ�"
      End
      Begin VB.Menu editmenu 
         Caption         =   "�޸ĳɼ�"
      End
      Begin VB.Menu findscoremenu 
         Caption         =   "��ѯ�ɼ�"
      End
   End
End
Attribute VB_Name = "mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub addinfomenu_Click()
addstudentfrm.Show
mainfrm.Hide
End Sub


Private Sub addusermenu_Click()
adduser.Show
mainfrm.Hide
End Sub

Private Sub editinfomenu_Click()
editfrm.Show
mainfrm.Hide
End Sub

Private Sub exitmenu_Click()
Dim ans As String
ans = MsgBox("ȷ���˳�ϵͳ��", vbYesNo + vbInformation, "�˳�")
If ans = vbYes Then
   End
End If
End Sub

Private Sub findinfomenu_Click()
findfrm.Show
mainfrm.Hide
End Sub

Private Sub personinfomenu_Click()
infofrm.Show
mainfrm.Hide

End Sub
