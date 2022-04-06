<!--#include file ="conn.asp"--> 
<%set rs=server.createobject("adodb.recordset")
exec="select * from resh where id=2"
rs.open exec,conn,1,1

total=rs("select1")+rs("select2")+rs("select3")+rs("select4")+rs("select5")+rs("select6")+rs("select7")
if total>0 then
ps1=rs("select1")/total
ps2=rs("select2")/total
ps3=rs("select3")/total

ws1=600*ps1
ws2=600*ps2
ws3=600*ps3

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>
<body>
<div align="center">
  <p>谢谢您的参与，已经有<font color="#FF0000"><%=total%></font>人参加了调查，下面是当前的调查结果</p>
  <table width="466" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="117">◇ 西藏</td>
      <td width="68" height="40"><%=rs("select1")%></td>
      <td width="189"><table width=<%=ws1%> height="20" border="0" cellpadding="0" cellspacing="0" bgcolor="#FF0000">
        <tr>
          <td>&nbsp;</td>
        </tr>
      </table></td>
      <td width="92"><%=FormatPercent(ps1)%></td>
    </tr>
    <tr>
      <td>◇ 桂林</td>
      <td height="40"><%=rs("select2")%></td>
      <td><table width=<%=ws2%> height="20" border="0" cellpadding="0" cellspacing="0" bgcolor="#00FF00">
        <tr>
          <td>&nbsp;</td>
        </tr>
      </table></td>
      <td><%=FormatPercent(ps2)%></td>
    </tr>
    <tr>
      <td>◇ 海南</td>
      <td height="40"><%=rs("select3")%></td>
      <td><table width=<%=ws3%> height="20" border="0" cellpadding="0" cellspacing="0" bgcolor="#0000FF">
        <tr>
          <td>&nbsp;</td>
        </tr>
      </table></td>
      <td><%=FormatPercent(ps3)%></td>
    </tr>
  </table>
  <p><a href="javascript:window.close()">关闭窗口</a></p>
  <% else 
  response.Write("还没有人参与调查！")
  end if%>
</div>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing%>
</body>
</html>
