VERSION 5.00
Begin VB.Form popmenu 
   Caption         =   "�˵�"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   Visible         =   0   'False
   Begin VB.Menu popmenu 
      Caption         =   "popmenu"
      Visible         =   0   'False
      Begin VB.Menu add 
         Caption         =   "���ѧ��������Ϣ"
      End
      Begin VB.Menu Change 
         Caption         =   "�޸�ѧ��������Ϣ"
      End
      Begin VB.Menu Check 
         Caption         =   "��ѯѧ��������Ϣ"
      End
      Begin VB.Menu Del 
         Caption         =   "ɾ��ѧ��������Ϣ"
      End
   End
   Begin VB.Menu popmenu1 
      Caption         =   "popmenu1"
      Visible         =   0   'False
      Begin VB.Menu addInOut 
         Caption         =   "�������ѧ��Ϣ"
      End
      Begin VB.Menu checkKQ1 
         Caption         =   "��ѯ������Ϣ"
      End
      Begin VB.Menu delKQ1 
         Caption         =   "ɾ������ѧ��Ϣ"
      End
   End
   Begin VB.Menu popmenu2 
      Caption         =   "popmenu2"
      Visible         =   0   'False
      Begin VB.Menu addOtherKQ 
         Caption         =   "�������������Ϣ"
      End
      Begin VB.Menu checkKQ2 
         Caption         =   "��ѯ������Ϣ"
      End
      Begin VB.Menu delKQ2 
         Caption         =   "ɾ������������Ϣ"
      End
   End
   Begin VB.Menu popmenu3 
      Caption         =   "popmenu3"
      Visible         =   0   'False
      Begin VB.Menu addAlteration 
         Caption         =   "��ӵ�����Ϣ"
      End
      Begin VB.Menu ChangeAlter 
         Caption         =   "�޸ĵ�����Ϣ"
      End
      Begin VB.Menu Checkalter 
         Caption         =   "��ѯ������Ϣ"
      End
      Begin VB.Menu DelAlter 
         Caption         =   "ɾ��������Ϣ"
      End
   End
End
Attribute VB_Name = "popmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public str1 As String

Private Sub add_Click()                             '���ѧ����Ϣ
    flag = 1
    frmStuff_info.Show vbModal
    frmStuff_info.ZOrder 0
    frmResult.ZOrder 1
End Sub

Private Sub addAlteration_Click()                   '��ӵ�����Ϣ
    flag = 1
    frmAlteration.Show
    frmAlteration.ZOrder 0
End Sub

Private Sub addInOut_Click()                        '�������ѧ��Ϣ
    flag = 1
    FrmAttendance.Show
    FrmAttendance.ZOrder 0
End Sub

Private Sub AddOtherKQ_Click()                       '�������������Ϣ
    flag = 1
    frmOtherKQ.Show
    frmOtherKQ.ZOrder 0
End Sub

Private Sub change_Click()  '�޸�ѧ����Ϣ
    flag = 2
    If frmResult.rsGrid.Rows > 1 Then
    
        gSQL = "select * from StuffInfo where SID='" & Trim(frmResult.rsGrid.TextMatrix( _
                frmResult.rsGrid.Row, 0)) & "'"
        frmStuff_info.Show
        frmStuff_info.ZOrder 0
    Else
     MsgBox "Ŀǰû��ѧ����Ϣ,�������ѧ����Ϣ��", vbOKOnly + vbExclamation, "����!"
     flag = 1
     frmStuff_info.Show
    End If
End Sub

Private Sub aa_Click()                           '�޸�����������Ϣ
    flag = 2
    If frmOKQResult.LRecordList.Rows > 1 Then
        kqsql2 = "select * from StuffInfo where SID='" & Trim(frmOKQResult.LRecordList.TextMatrix( _
                frmOKQResult.LRecordList.Row, 0)) & "'"
        frmOtherKQ.Show
        frmOtherKQ.ZOrder 0
    Else
     MsgBox "Ŀǰû��ѧ����Ϣ,�������ѧ����Ϣ��", vbOKOnly + vbExclamation, "����!"
     flag = 1
     frmOtherKQ.Show
    End If
End Sub


Private Sub tt_Click()                           '�޸�����ѧ��Ϣ
    flag = 2
    'FrmAttendance.Caption = "�޸�ѧ������ѧ��Ϣ"
    If frmAResult.recordlist.Rows > 1 Then
        kqsql = "select * from AttendanceInfo where ID='" & Trim(frmAResult.recordlist.TextMatrix( _
                frmAResult.recordlist.Row, 0)) & "'"
        FrmAttendance.Show
        FrmAttendance.ZOrder 0
    Else
     MsgBox "Ŀǰû������ѧ��Ϣ,���������Ϣ��", vbOKOnly + vbExclamation, "����!"
     flag = 1
     FrmAttendance.Show
    End If
End Sub
Private Sub ChangeAlter_Click()                      '�޸ĵ�����Ϣ
    Dim rs As New ADODB.Recordset
    flag = 2
    frmAlteration.Caption = "�޸�ѧ��������Ϣ"
    If frmAlterationResult.DataGrid1.Row < 0 Then
        MsgBox "Ŀǰû�м�¼��", vbOKOnly + vbExclamation, "��ʾ��"
        flag = 1
        frmAlteration.Show
        frmAlteration.ZOrder 0
    Else
       str1 = "select * from AlterationInfo where ID=" & Trim( _
            frmAlterationResult.DataGrid1.Columns(0))
        frmAlteration.ID = Trim(frmAlterationResult.DataGrid1.Columns(0))
        Set rs = TransactSQL(str1)
        If rs.EOF = False Then
        With rs
            frmAlteration.AID = rs(1)
            frmAlteration.AName = rs(2)
            frmAlteration.AOldDept = rs(3)
            frmAlteration.ANewDept = rs(4)
            frmAlteration.AOldPosition = rs(5)
            frmAlteration.ANewPosition = rs(6)
            frmAlteration.AOutTime = rs(7)
            frmAlteration.AInTime = rs(8)
            frmAlteration.ARemark = rs(9)
        End With
            rs.Close
        End If
        frmAlteration.Show
        frmAlteration.ZOrder 0
    End If
End Sub

Private Sub check_Click()                            '��ѯѧ����Ϣ
    frmCheckStuff.Show
End Sub

Private Sub checkKQ1_Click()                         '��ѯ������Ϣ
    frmCheckKQ.Show
    frmCheckKQ.ZOrder 0
End Sub

Private Sub checkKQ2_Click()                         '��ѯ������Ϣ
    frmCheckKQ.Show
    frmCheckKQ.ZOrder 0
End Sub

Private Sub del_Click()                              'ɾ��ѧ����Ϣ
    Dim SQL As String
    If frmResult.rsGrid.Rows = 1 Then
        MsgBox "Ŀǰû��ѧ����Ϣ,�������ѧ����Ϣ��", vbOKOnly + vbExclamation, "����!"
        flag = 1
        frmStuff_info.Show
        frmStuff_info.ZOrder 0
    Else
        SQL = "delete from StuffInfo where SID='" & Trim(frmResult.rsGrid.TextMatrix( _
                frmResult.rsGrid.Row, 0)) & "'"
        If MsgBox("���Ҫɾ��������¼ô��", vbOKCancel + vbExclamation, "��ʾ��") = vbOK _
        Then
        TransactSQL (SQL)
        MsgBox "ѧ����Ϣ��¼�Ѿ�ɾ����", vbOKOnly + vbExclamation, "����!"
        Unload Me
        SQL = "select * from StuffInfo"
       
        frmResult.createList (SQL)
        Unload frmResult
        frmResult.Show
        End If
    End If
End Sub

Private Sub delKQ1_Click()                           'ɾ������ѧ��Ϣ
     Dim SQL As String
    If frmAResult.recordlist.Rows = 1 Then
        MsgBox "Ŀǰû������ѧ��Ϣ��", vbOKOnly + vbExclamation, "���棡"
        flag = 1
        FrmAttendance.Show
        FrmAttendance.ZOrder 0
    Else
            SQL = "delete from AttendanceInfo where ID=" & Trim(frmAResult.recordlist.TextMatrix( _
                frmAResult.recordlist.Row, 0))
        If MsgBox("���Ҫɾ��������¼ô��", vbOKCancel + vbExclamation, "��ʾ��") = vbOK _
        Then
        TransactSQL (SQL)
        MsgBox "��¼�Ѿ�ɾ����", vbOKOnly + vbExclamation, "����!"
        Unload Me
        SQL = "select * from AttendanceInfo"
        'frmAResult.ListTopic
        frmAResult.ShowData (SQL)
        
        Unload frmAResult
        frmAResult.Show
         Else
         Unload frmAResult
         End If
    End If
    
End Sub

Private Sub DelAlter_Click()                         'ɾ��������Ϣ
    Dim SQL As String
    If frmAlterationResult.DataGrid1.Row < 0 Then
        MsgBox "Ŀǰû�м�¼��", vbOKOnly + vbExclamation, "��ʾ��"
        flag = 1
        frmAlteration.Show
        frmAlteration.ZOrder 0
    Else
        SQL = "delete from AlterationInfo where ID=" & Trim( _
            frmAlterationResult.DataGrid1.Columns(0).CellText( _
            frmAlterationResult.DataGrid1.Bookmark))
        If MsgBox("���Ҫɾ��������¼ô��", vbOKCancel) = vbOK Then
            TransactSQL (SQL)
            MsgBox "��¼�Ѿ�ɾ����", vbOKOnly + vbExclamation, "��ʾ��"
            SQL = "select * from AlterationInfo order by ID"
            frmAlterationResult.Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\Person.mdb"
            frmAlterationResult.Adodc1.RecordSource = SQL
            If SQL <> "" Then
                frmAlterationResult.Adodc1.Refresh
            End If
            Set frmAlterationResult.DataGrid1.DataSource = frmAlterationResult.Adodc1.Recordset
            frmAlterationResult.DataGrid1.Refresh
            frmAlterationResult.Show
            frmAlterationResult.ZOrder 0
        End If
    End If
End Sub

Private Sub delKQ2_Click()                            'ɾ������������Ϣ
    Dim SQL As String
    Select Case frmOKQResult.SSTab.Caption
    Case "ѧ�������Ϣ�б�"
        If frmOKQResult.LRecordList.Rows = 1 Then
            MsgBox "Ŀǰû�������Ϣ��", vbOKOnly + vbExclamation, "���棡"
            flag = 1
            frmOtherKQ.Show
            frmOtherKQ.ZOrder 0
        Else
            SQL = "delete from LeaveInfo where LID="
            SQL = SQL & Trim(frmOKQResult.LRecordList.TextMatrix( _
                                    frmOKQResult.LRecordList.Row, 0))
            If MsgBox("���Ҫɾ��������¼ô��", vbOKCancel + vbExclamation, "��ʾ��") = vbOK _
            Then
                TransactSQL (SQL)
                MsgBox "��¼�Ѿ�ɾ����", vbOKOnly + vbExclamation, "����!"
                Unload Me
                SQL = "select * from LeaveInfo"
                Call frmOKQResult.LeaveTopic
                Call frmOKQResult.ShowLRecord(SQL)
                Unload frmOKQResult
                frmOKQResult.Show
                 frmOKQResult.SSTab.Caption = "ѧ�������Ϣ�б�"
            End If
          End If
    Case "ѧ��������Ϣ�б�"
        If frmOKQResult.ORecordList.Rows = 1 Then
            MsgBox "Ŀǰû�в�����Ϣ��", vbOKOnly + vbExclamation, "���棡"
            flag = 1
            frmOtherKQ.Show
            frmOtherKQ.ZOrder 0
        Else
            SQL = "delete from OvertimeInfo where OID="
            SQL = SQL & Trim(frmOKQResult.ORecordList.TextMatrix( _
                                    frmOKQResult.ORecordList.Row, 0))
            If MsgBox("���Ҫɾ��������¼ô��", vbOKCancel + vbExclamation, "��ʾ��") = vbOK _
            Then
                TransactSQL (SQL)
                MsgBox "��¼�Ѿ�ɾ����", vbOKOnly + vbExclamation, "����!"
                Unload Me
                SQL = "select * from OvertimeInfo"
                Call frmOKQResult.OverTimeTopic
                Call frmOKQResult.ShowORecord(SQL)
                 Unload frmOKQResult
                frmOKQResult.Show
                 frmOKQResult.SSTab.Caption = "ѧ��������Ϣ�б�"
            End If
        End If
    Case "ѧ��������Ϣ�б�"
        If frmOKQResult.ERecordList.Rows = 1 Then
            MsgBox "Ŀǰû�п�����Ϣ��", vbOKOnly + vbExclamation, "���棡"
            flag = 1
            frmOtherKQ.Show
            frmOtherKQ.ZOrder 0
        Else
            SQL = "delete from ErrandInfo where EID="
            SQL = SQL & Trim(frmOKQResult.ERecordList.TextMatrix( _
                                    frmOKQResult.ERecordList.Row, 0))
            If MsgBox("���Ҫɾ��������¼ô��", vbOKCancel + vbExclamation, "��ʾ��") = vbOK _
            Then
                TransactSQL (SQL)
                MsgBox "��¼�Ѿ�ɾ����", vbOKOnly + vbExclamation, "����!"
                Unload Me
                SQL = "select * from ErrandInfo"
                Call frmOKQResult.ErrandTopic
                Call frmOKQResult.ShowERecord(SQL)
                 Unload frmOKQResult
                'frmOKQResult.Show
                 frmOKQResult.SSTab.Caption = "ѧ��������Ϣ��"
            End If
        End If
    End Select
End Sub



