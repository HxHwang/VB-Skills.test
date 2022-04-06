<!--#include file ="conn.asp"--> 
<%
userid=trim(request.Form("userid"))
pwd=trim(request.Form("pwd"))
repwd=trim(request.Form("repwd"))
username=trim(request.Form("username"))
email=trim(request.Form("email"))

exec="select * from res where 帐号='" & userid & "'" 
set rs=server.CreateObject("adodb.recordset") 
rs.open exec,conn,1,1 
if rs.eof and rs.bof then
  exec="insert into res(帐号,密码,姓名,email)values('"+userid+"','"+pwd+"','"+username+"','"+email+"')" 
  conn.execute exec 
  rs.close
  set rs=nothing
  conn.close
  set conn=nothing
  response.Write("祝贺你注册成功！按<a href=login.htm>这里</a>进入登录页")
else
rs.close
set rs=nothing
conn.close
set conn=nothing
showMsg("此用户名已被注册！")
end if
%>
<% sub showMsg(msg)%>
<body>
<center>
<h3><%=msg%></h3>
</center>
<form>
<p align="center"><input type="button" value="返回重新填写" onClick="history.back();"></p>
</form>
</body>
<% end sub%>