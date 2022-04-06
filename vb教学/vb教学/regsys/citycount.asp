<%@LANGUAGE="VBSCRIPT"%>
<%
'从文本文件中读出城市得票数
function ReadCt(cfile)
Set ofs = Server.CreateObject("Scripting.FileSystemObject")
set ofsFile=ofs.OpenTextFile(Server.MapPath(cfile),1,true)
if not atendofstream then
ReadCt=Clng(ofsFile.readline)
else
ReadCt=0
end if
ofsFile.close
set ofs=nothing
end function
%>
<%
xz1=ReadCt("xz.txt")
gl1=ReadCt("gl.txt")
hn1=ReadCt("hn.txt")

'计算总票数，并分别计算各城市得票的百分比
total=xz1+gl1+hn1
pxz=xz1/total
pgl=gl1/total
phn=hn1/total

'根据各城市得票数的百分比计算横条图的宽度
wxz=600*pxz
wgl=600*pgl
whn=600*phn
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>查看投票结果</title>
</head>

<body>
<div align="center">
  <table width="778" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td height="82" colspan="3" align="center">目前共计收到选票<font color="#FF0000"><%=total%></font>张，其中：</td>
    </tr>
    <tr>
      <td width="227" rowspan="2" align="center"><a href=city.asp?vote=xz><img src="img/200510120025.jpg" width="200" height="100" border="0"></a></td>
      <td width="330" height="50" align="left"><p>西藏得票：<%=xz1%>票</p>        </td>
      <td width="221" align="left">&nbsp;</td>
    </tr>
    <tr>
      <td align="left"><table width=<%=wxz%> height="20" border="0" cellpadding="0" cellspacing="0" bgcolor="#00FF00">
        <tr>
          <td>&nbsp;</td>
        </tr>
      </table></td>
      <td width="221" align="left"><%=FormatPercent(pxz)%></td>
    </tr>
    <tr>
      <td height="21" align="center">西藏</td>
      <td align="left">&nbsp;</td>
      <td align="left">&nbsp;</td>
    </tr>
    <tr>
      <td height="21" rowspan="2" align="center"><a href=city.asp?vote=gl><img src="img/200510120029.jpg" width="200" height="100" border="0" /></a></td>
      <td height="50" align="left"><p>挂林得票：<%=gl1%>票</p>        </td>
      <td align="left">&nbsp;</td>
    </tr>
    <tr>
      <td align="left"><table width=<%=wgl%> height="20" border="0" cellpadding="0" cellspacing="0" bgcolor="#FF0000">
        <tr>
          <td>&nbsp;</td>
        </tr>
      </table></td>
      <td align="left"><%=FormatPercent(pgl)%></td>
    </tr>
    <tr>
      <td height="21" align="center">桂林</td>
      <td align="left">&nbsp;</td>
      <td align="left">&nbsp;</td>
    </tr>
    <tr>
      <td height="21" rowspan="2" align="center"><a href=city.asp?vote=hn><img src="img/200510120033.jpg" width="200" height="100" border="0" /></a></td>
      <td height="50" align="left"><p>海南得票：<%=hn1%>票</p>
        </td>
      <td align="left">&nbsp;</td>
    </tr>
    <tr>
      <td align="left"><table width=<%=whn%> height="20" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFF00">
        <tr>
          <td>&nbsp;</td>
        </tr>
      </table></td>
      <td align="left"><%=FormatPercent(phn)%></td>
    </tr>
    <tr>
      <td height="21" align="center">海南</td>
      <td align="center">&nbsp;</td>
      <td align="center">&nbsp;</td>
    </tr>
  </table>
</div>
</body>
</html>
