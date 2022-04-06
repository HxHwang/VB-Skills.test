<!--#include file ="conn.asp"--> 
<%
userid=request.form("userid")
pwd=request.form("pwd")
exec="select * from res where(ÕÊºÅ='"&userid&"' and ÃÜÂë='"&pwd&"')"
set rs=server.createobject("adodb.recordset")
rs.open exec,conn
if not rs.eof then
rs.Close
conn.Close
session("userid")=userid
session("checked")="yes"
response.Redirect "welcome.asp"
else
response.Write("ÓÃ»§Ãû»òÃÜÂë´íÎó£¡")
end if
%>