<%@LANGUAGE="VBSCRIPT"%>
<%
'��ÿͻ���IP
function getip() 
 getip=Request.ServerVariables("REMOTE_ADDR") 
end function
%>
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
sub writeCt(cfile,ct)
Set ofs = Server.CreateObject("Scripting.FileSystemObject")
set ofsFile=ofs.OpenTextFile(Server.MapPath(cfile),2,true)
ofsFile.writeline(ct)
ofsfile.close
set ofs=nothing
end sub
%>
<%
application.Lock()
vote=request("vote")
xz1=ReadCt("xz.txt")
gl1=ReadCt("gl.txt")
hn1=ReadCt("hn.txt")
select case vote
 case "xz"
  xz1=xz1+1
  writeCt "xz.txt",xz1
 case "gl" 
  gl1=gl1+1
  writeCt "gl.txt",gl1
 case "hn" 
  hn1=hn1+1
  writeCt "hn.txt",hn1
end select
application.UnLock()
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�������� ����ͶƱ</title>
</head>
<body>
<div align="center">
  <p>�������� ����ͶƱ </p>
  <table width="778" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td align="center"><a href=city.asp?vote=xz><img src="img/200510120025.jpg" width="200" height="100" border="0"></a></td>
      <td align="center"><a href=city.asp?vote=gl><img src="img/200510120029.jpg" width="200" height="100" border="0"></a></td>
      <td align="center"><a href=city.asp?vote=hn><img src="img/200510120033.jpg" width="200" height="100" border="0"></a></td>
    </tr>
    <tr>
      <td height="21" align="center">���ص�Ʊ����<%=xz1%></td>
      <td align="center">���ֵ�Ʊ����<%=gl1%></td>
      <td align="center">���ϵ�Ʊ����<%=hn1%></td>
    </tr>
    <tr>
      <td height="40" colspan="3" align="center">���������һ������ϲ���ĳ���ͼƬ���ó��оͿɻ���㱦���һƱ��</td>
    </tr>
    <tr>
      <td height="40" colspan="3" align="center"><a href="citycount.asp">�鿴ͶƱ���</a></td>
    </tr>
  </table>
  <p>&nbsp;</p>
</div>
</body>
</html>
