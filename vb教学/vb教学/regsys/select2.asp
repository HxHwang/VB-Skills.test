<% if request.Form("Options")<>"" then%>
<%
if not request.ServerVariables("REMOTE_ADDR")=request.Cookies("IPAddress") then
response.Cookies("IPAddress")=request.ServerVariables("REMOTE_ADDR")
%>
<!--#include file ="conn.asp"--> 
<%
selected=request.Form("Options")
set rs=server.createobject("adodb.recordset")
exec="update resh set select"&selected&"=select"&selected&"+1 where id=2"
rs.open exec,conn,3,3
set rs=nothing
conn.close
set conn=nothing
response.Redirect("view2.asp")
else
response.Write("ͶƱʧ����ʾ�����ղ���Ͷ��Ʊ��лл����֧�֣�")
end if
else
response.Write("ͶƱʧ����ʾ��������ѡ���ˣ�")
end if
%>
