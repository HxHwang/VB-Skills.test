VERSION 5.00
Begin VB.MDIForm FrmMain 
   BackColor       =   &H8000000C&
   Caption         =   "ѧ����������ϵͳ"
   ClientHeight    =   3555
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6405
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin VB.Menu System 
      Caption         =   "ϵͳ"
      Index           =   1
      Begin VB.Menu Add_User 
         Caption         =   "����û�"
      End
      Begin VB.Menu Change_PWD 
         Caption         =   "�޸�����"
      End
      Begin VB.Menu System_EXIT 
         Caption         =   "�˳�"
      End
   End
   Begin VB.Menu Stuff_Info 
      Caption         =   "ѧ��������Ϣ"
      Index           =   2
      Begin VB.Menu Add_Stuff 
         Caption         =   "���ѧ����Ϣ"
      End
      Begin VB.Menu Change_Stuff 
         Caption         =   "�޸�ѧ����Ϣ"
      End
      Begin VB.Menu Check_Stuff 
         Caption         =   "��ѯѧ����Ϣ"
      End
      Begin VB.Menu Del_Stuff 
         Caption         =   "ɾ��ѧ����Ϣ"
      End
   End
   Begin VB.Menu Stuff_Checkin 
      Caption         =   "ѧ��������Ϣ"
      Index           =   3
      Begin VB.Menu Add_Checkin 
         Caption         =   "��ӳ�����Ϣ"
         Begin VB.Menu AddAttendance 
            Caption         =   "�������ѧ��Ϣ"
         End
         Begin VB.Menu AddOtherKQ 
            Caption         =   "�������������Ϣ"
         End
      End
      Begin VB.Menu Change_Checkin 
         Caption         =   "�޸ĳ�����Ϣ"
         Begin VB.Menu ChangeAttendance 
            Caption         =   "�޸�����ѧ��Ϣ"
         End
         Begin VB.Menu ChangeOtherKQ 
            Caption         =   "�޸�����������Ϣ"
         End
      End
      Begin VB.Menu Check_Checkin 
         Caption         =   "��ѯ������Ϣ"
      End
      Begin VB.Menu Del_Checkin 
         Caption         =   "ɾ��������Ϣ"
         Begin VB.Menu delInOut 
            Caption         =   "ɾ������ѧ��Ϣ"
         End
         Begin VB.Menu delOtherKQ 
            Caption         =   "ɾ������������Ϣ"
         End
      End
      Begin VB.Menu SetTime 
         Caption         =   "��������ѧʱ��"
      End
   End
   Begin VB.Menu Stuff_Alteration 
      Caption         =   "ѧ��������Ϣ"
      Index           =   4
      Begin VB.Menu Add_Alter 
         Caption         =   "��ӵ�����Ϣ"
      End
      Begin VB.Menu Chage_Alter 
         Caption         =   "�޸ĵ�����Ϣ"
      End
      Begin VB.Menu Check_Alter 
         Caption         =   "��ѯ������Ϣ"
      End
      Begin VB.Menu Del_Alter 
         Caption         =   "ɾ��������Ϣ"
      End
   End
   Begin VB.Menu System_Help 
      Caption         =   "����"
      Index           =   6
      Begin VB.Menu About 
         Caption         =   "����"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private SQL As String

Private Sub About_Click()                         '���ڴ���
    frmAbout.Show
    frmAbout.ZOrder 0
End Sub

Private Sub Add_Alter_Click()                     '��ӵ�����Ϣ
    flag = 1
    frmAlteration.Caption = "���ѧ��������Ϣ"
    frmAlteration.Show
    frmAlteration.ZOrder 0
End Sub

Private Sub Add_Stuff_Click()                      '���Ա����Ϣ
    flag = 1
    frmStuff_info.Show
    frmStuff_info.ZOrder 0
End Sub

Private Sub Add_User_Click()                       '����û�
    Dim fAdd As New frmAddUser
    fAdd.Show
    fAdd.ZOrder 0
End Sub

Private Sub AddAttendance_Click()                  '�������ѧ��Ϣ
    flag = 1
    FrmAttendance.Show
    FrmAttendance.ZOrder 0
End Sub

Private Sub AddOtherKQ_Click()                     '�������������Ϣ
    flag = 1
    frmOtherKQ.Show
    frmOtherKQ.ZOrder 0
End Sub

Private Sub Chage_Alter_Click()                     '�޸ĵ�����Ϣ
    frmAlterationResult.Show
    frmAlterationResult.ZOrder 0
End Sub

Private Sub Change_PWD_Click()                      '�޸�����
    Dim fChangePWD As New frmChangePWD
    fChangePWD.Show
End Sub

Private Sub Change_Stuff_Click()                    '�޸�Ա����Ϣ
    frmCheckStuff.topic = "ѡ���޸�����"
    frmCheckStuff.Caption = "�޸�ѧ��������Ϣ"
    SQL = "select * from StuffInfo order by SID"
    frmResult.createList (SQL)
    frmResult.Show
    frmResult.ZOrder 0
    frmCheckStuff.Show
    frmCheckStuff.ZOrder 0
End Sub

Private Sub ChangeAttendance_Click()                 '�޸�����ѧ��Ϣ
    frmAResult.Show
    'frmAResult.ZOrder 0
End Sub

Private Sub changeOtherKQ_Click()                    '�޸�����������Ϣ
    frmOKQResult.Show
    frmOKQResult.ZOrder 0
End Sub

Private Sub Check_Alter_Click()                       '��ѯ������Ϣ
    frmCheckAlter.Show
    frmCheckAlter.ZOrder 0
End Sub

Private Sub Check_Checkin_Click()                    '��ѯ����������Ϣ
    frmCheckKQ.Show
    frmCheckKQ.ZOrder 0
End Sub

Private Sub Check_Stuff_Click()                       '��ѯԱ����Ϣ
    SQL = "select * from StuffInfo"
    frmResult.createList (SQL)
    frmResult.Show
    frmCheckStuff.Show
    frmResult.ZOrder 1
    frmCheckStuff.ZOrder 0
End Sub

Private Sub Del_Alter_Click()                          'ɾ��������Ϣ
    frmAlterationResult.Show
    frmAlterationResult.ZOrder 0
End Sub

Private Sub Del_Stuff_Click()                          'ɾ��Ա����Ϣ
    frmCheckStuff.topic = "ѡ��ɾ������"
    frmCheckStuff.Caption = "ɾ��ѧ��������Ϣ"
    SQL = "select * from StuffInfo"
    frmResult.createList (SQL)
    frmResult.Show
    frmCheckStuff.Show
    frmResult.ZOrder 1
    frmCheckStuff.ZOrder 0
End Sub

Private Sub delInOut_Click()                            'ɾ������ѧ��Ϣ
    Dim SQL As String
    SQL = "select * from AttendanceInfo order by ID desc"
    Call frmAResult.ListTopic
    Call frmAResult.ShowData(SQL)
    frmAResult.Show
    frmAResult.ZOrder 0
End Sub

Private Sub delOtherKQ_Click()                            'ɾ������������Ϣ
    frmOKQResult.Show
    frmOKQResult.ZOrder 0
End Sub

Private Sub SetTime_Click()                               '��������ѧʱ��
    frmSetTime.Show
    frmSetTime.ZOrder 0
End Sub

Private Sub System_EXIT_Click()
    Unload Me
    Exit Sub
End Sub
