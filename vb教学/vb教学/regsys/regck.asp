<!--#include file ="conn.asp"--> 
<%
userid=trim(request.Form("userid"))
pwd=trim(request.Form("pwd"))
repwd=trim(request.Form("repwd"))
username=trim(request.Form("username"))
email=trim(request.Form("email"))

exec="select * from res where �ʺ�='" & userid & "'" 
set rs=server.CreateObject("adodb.recordset") 
rs.open exec,conn,1,1 
if rs.eof and rs.bof then
  exec="insert into res(�ʺ�,����,����,email)values('"+userid+"','"+pwd+"','"+username+"','"+email+"')" 
  conn.execute exec 
  rs.close
  set rs=nothing
  conn.close
  set conn=nothing
  response.Write("ף����ע��ɹ�����<a href=login.htm>����</a>�����¼ҳ")
else
rs.close
set rs=nothing
conn.close
set conn=nothing
showMsg("���û����ѱ�ע�ᣡ")
end if
%>
<% sub showMsg(msg)%>
<body>
<center>
<h3><%=msg%></h3>
</center>
<form>
<p align="center"><input type="button" value="����������д" onClick="history.back();"></p>
</form>
</body>
<% end sub%>