Attribute VB_Name = "Module1"
Option Explicit

Public dbcn As New ADODB.Connection
Public CnStr As String
Public Day As String
Public FaJin As String
Public MaxBook As String


Public Sub OpenData()
   '�����������ݿ��ַ���
    CnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
    CnStr = CnStr & App.Path & "\Library.mdb;Persist Security Info=False"
    '�������ݿ�
    dbcn.Open CnStr
    
End Sub

Public Sub CloseData()
dbcn.Close

End Sub

'������������  ���ݿ� ���ӡ�ɾ�����޸� ����
'�������� sqlstr ��������
'       dbcn  �������ݿ�
'       adodc ADO����
'����ֵ   0 ����ʧ��
'         1 �����ɹ�
Public Function DataManage(sqlstr As String, dbcn As ADODB.Connection, Adodc1 As Adodc) As Integer
On Error GoTo erp:
dbcn.Execute sqlstr
Adodc1.Refresh
DataManage = 1
Exit Function
erp:
DataManage = 0
End Function
'���ܣ���ѯ����
'������sqlstr ��ѯ����
'      adodc1 ADO����
'      Datagrid1  װ��ѯ�����
'����ֵ   0 ����ʧ��
'         1 �����ɹ�
'         2 δ�鵽����
Public Function DataQuery(sqlstr As String, Adodc1 As Adodc, DataGrid1 As DataGrid) As Integer
'On Error GoTo erp:

Adodc1.RecordSource = sqlstr
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh

If Adodc1.Recordset.EOF Then
   'MsgBox "���ݿ���û�з���Ҫ��ļ�¼��"
   DataQuery = 2
   Exit Function
End If
DataQuery = 1
Exit Function
erp:
DataQuery = 0

End Function



