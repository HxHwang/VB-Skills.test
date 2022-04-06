<!--#include file ="conn.asp"--> 
<% set rs=server.CreateObject("Adodb.Recordset")
sword=request.Form("word")
sword="%"&sword&"%"
exec= "select * from list where word like '"&sword&"'"
rs.open exec,conn,1,1
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>

<body>

<div align="center">
  <p>统计查询共有<%=rs.recordcount%>条记录</p>
  <table width="200" border="1" cellspacing="0" cellpadding="0">
    <tr>
      <td>id</td>
      <td>主题</td>
      <td>查看</td>
    </tr>
   <% while not rs.eof%>
    <tr>
      <td><%=rs("id")%></td>
      <td><%=rs("title")%></td>
      <td><a href=<%=rs("url")%>>GO</a></td>
    </tr>
	<%rs.movenext%>
	<% wend%>
  </table>
  <%rs.close
  conn.close%>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
</div>
</body>
</html>
