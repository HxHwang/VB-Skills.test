Attribute VB_Name = "Module1"
Public gUserName As String                        '�����û�����
Public flag As Integer                            '��Ӻ��޸ĵı�־
Public gSQL As String                             '����SQL���
Public kqsql As String                            '�����ѯ���ڽ��SQL���
Public kqsql2 As String                           '�����ѯ�������ڽ��SQL���
Public ArecordID As Integer                       '��������ѧ��¼���
Public LrecordID As Integer                       '������ټ�¼���
Public OrecordID As Integer                       '���油�μ�¼���
Public ErecordID As Integer                       '������μ�¼���
Public iflag As Integer                           '���ݿ��Ƿ�򿪱�־
Public conn As New ADODB.Connection



Public Function TransactSQL(ByVal sql As String) As ADODB.Recordset
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim strConnection As String
Dim strArray() As String
Set con = New ADODB.Connection                  '��������
Set rs = New ADODB.Recordset                    '������¼��
On Error GoTo TransactSQL_Error
    strConnection = "Provider=Microsoft.jet.oledb.4.0;Data Source=" & App.Path & "\Person.mdb"
    strArray = Split(sql)
    con.Open strConnection                      '������
    If StrComp(UCase$(strArray(0)), "select", vbTextCompare) = 0 Then
        rs.Open Trim$(sql), con, adOpenKeyset, adLockOptimistic
        Set TransactSQL = rs                   '���ؼ�¼��
        iflag = 1
    Else
        con.Execute sql                        'ִ������
        iflag = 1
    End If
TransactSQL_Exit:
    Set rs = Nothing
    Set con = Nothing
    Exit Function
TransactSQL_Error:
    MsgBox "��ѯ����" & Err.Description
    iflag = 2
    Resume TransactSQL_Exit
End Function

Public Sub TabToEnter(Key As Integer)
    If Key = 13 Then                            '�ж��Ƿ�Ϊ�س���
    SendKeys "{TAB}"                            'ת��ΪTab��
    End If
End Sub

Sub main()
    Dim fLogin As New frmLogin
    fLogin.Show vbModual                        '��ʾ����
End Sub



