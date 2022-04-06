Attribute VB_Name = "Module1"
Option Explicit

Public dbcn As New ADODB.Connection
Public CnStr As String
Public Day As String
Public FaJin As String
Public MaxBook As String


Public Sub OpenData()
   '设置连接数据库字符串
    CnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
    CnStr = CnStr & App.Path & "\Library.mdb;Persist Security Info=False"
    '连接数据库
    dbcn.Open CnStr
    
End Sub

Public Sub CloseData()
dbcn.Close

End Sub

'函数功能描述  数据库 增加、删除、修改 操作
'参数描述 sqlstr 操作描述
'       dbcn  连接数据库
'       adodc ADO数据
'返回值   0 操作失败
'         1 操作成功
Public Function DataManage(sqlstr As String, dbcn As ADODB.Connection, Adodc1 As Adodc) As Integer
On Error GoTo erp:
dbcn.Execute sqlstr
Adodc1.Refresh
DataManage = 1
Exit Function
erp:
DataManage = 0
End Function
'功能：查询操作
'参数：sqlstr 查询描述
'      adodc1 ADO数据
'      Datagrid1  装查询结果的
'返回值   0 操作失败
'         1 操作成功
'         2 未查到数据
Public Function DataQuery(sqlstr As String, Adodc1 As Adodc, DataGrid1 As DataGrid) As Integer
'On Error GoTo erp:

Adodc1.RecordSource = sqlstr
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh

If Adodc1.Recordset.EOF Then
   'MsgBox "数据库中没有符合要求的记录！"
   DataQuery = 2
   Exit Function
End If
DataQuery = 1
Exit Function
erp:
DataQuery = 0

End Function



