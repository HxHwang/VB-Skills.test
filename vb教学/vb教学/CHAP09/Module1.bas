Attribute VB_Name = "Module1"
Public gUserName As String                        '保存用户名称
Public flag As Integer                            '添加和修改的标志
Public gSQL As String                             '保存SQL语句
Public kqsql As String                            '保存查询考勤结果SQL语句
Public kqsql2 As String                           '保存查询其他考勤结果SQL语句
Public ArecordID As Integer                       '保存上下学记录编号
Public LrecordID As Integer                       '保存请假记录编号
Public OrecordID As Integer                       '保存补课记录编号
Public ErecordID As Integer                       '保存旷课记录编号
Public iflag As Integer                           '数据库是否打开标志
Public conn As New ADODB.Connection



Public Function TransactSQL(ByVal sql As String) As ADODB.Recordset
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim strConnection As String
Dim strArray() As String
Set con = New ADODB.Connection                  '创建连接
Set rs = New ADODB.Recordset                    '创建记录集
On Error GoTo TransactSQL_Error
    strConnection = "Provider=Microsoft.jet.oledb.4.0;Data Source=" & App.Path & "\Person.mdb"
    strArray = Split(sql)
    con.Open strConnection                      '打开连接
    If StrComp(UCase$(strArray(0)), "select", vbTextCompare) = 0 Then
        rs.Open Trim$(sql), con, adOpenKeyset, adLockOptimistic
        Set TransactSQL = rs                   '返回记录集
        iflag = 1
    Else
        con.Execute sql                        '执行命令
        iflag = 1
    End If
TransactSQL_Exit:
    Set rs = Nothing
    Set con = Nothing
    Exit Function
TransactSQL_Error:
    MsgBox "查询错误：" & Err.Description
    iflag = 2
    Resume TransactSQL_Exit
End Function

Public Sub TabToEnter(Key As Integer)
    If Key = 13 Then                            '判断是否为回车键
    SendKeys "{TAB}"                            '转换为Tab键
    End If
End Sub

Sub main()
    Dim fLogin As New frmLogin
    fLogin.Show vbModual                        '显示窗体
End Sub



