<!--#include file ="conn.asp"--> 
<%
userid=request.form("userid")
pwd=request.form("pwd")
exec="select * from res where(�ʺ�='"&userid&"' and ����='"&pwd&"')"
set rs=server.createobject("adodb.recordset")
rs.open exec,conn
if not rs.eof then
rs.Close
conn.Close
session("userid")=userid
session("checked")="yes"
response.Redirect "welcome.asp"
else
response.Write("�û������������")
end if
%>