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
response.Write("投票失败提示：您刚才已投了票，谢谢您的支持！")
end if
else
response.Write("投票失败提示：您忘记选择了！")
end if
%>
