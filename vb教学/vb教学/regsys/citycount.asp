<%@LANGUAGE="VBSCRIPT"%>
<%
'���ı��ļ��ж������е�Ʊ��
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

'������Ʊ�������ֱ��������е�Ʊ�İٷֱ�
total=xz1+gl1+hn1
pxz=xz1/total
pgl=gl1/total
phn=hn1/total

'���ݸ����е�Ʊ���İٷֱȼ������ͼ�Ŀ��
wxz=600*pxz
wgl=600*pgl
whn=600*phn
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�鿴ͶƱ���</title>
</head>

<body>
<div align="center">
  <table width="778" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td height="82" colspan="3" align="center">Ŀǰ�����յ�ѡƱ<font color="#FF0000"><%=total%></font>�ţ����У�</td>
    </tr>
    <tr>
      <td width="227" rowspan="2" align="center"><a href=city.asp?vote=xz><img src="img/200510120025.jpg" width="200" height="100" border="0"></a></td>
      <td width="330" height="50" align="left"><p>���ص�Ʊ��<%=xz1%>Ʊ</p>        </td>
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
      <td height="21" align="center">����</td>
      <td align="left">&nbsp;</td>
      <td align="left">&nbsp;</td>
    </tr>
    <tr>
      <td height="21" rowspan="2" align="center"><a href=city.asp?vote=gl><img src="img/200510120029.jpg" width="200" height="100" border="0" /></a></td>
      <td height="50" align="left"><p>���ֵ�Ʊ��<%=gl1%>Ʊ</p>        </td>
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
      <td height="21" align="center">����</td>
      <td align="left">&nbsp;</td>
      <td align="left">&nbsp;</td>
    </tr>
    <tr>
      <td height="21" rowspan="2" align="center"><a href=city.asp?vote=hn><img src="img/200510120033.jpg" width="200" height="100" border="0" /></a></td>
      <td height="50" align="left"><p>���ϵ�Ʊ��<%=hn1%>Ʊ</p>
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
      <td height="21" align="center">����</td>
      <td align="center">&nbsp;</td>
      <td align="center">&nbsp;</td>
    </tr>
  </table>
</div>
</body>
</html>
