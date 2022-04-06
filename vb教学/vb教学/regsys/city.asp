<%@LANGUAGE="VBSCRIPT"%>
<%
'获得客户端IP
function getip() 
 getip=Request.ServerVariables("REMOTE_ADDR") 
end function
%>
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
<title>魅力城市 网上投票</title>
</head>
<body>
<div align="center">
  <p>魅力城市 网上投票 </p>
  <table width="778" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td align="center"><a href=city.asp?vote=xz><img src="img/200510120025.jpg" width="200" height="100" border="0"></a></td>
      <td align="center"><a href=city.asp?vote=gl><img src="img/200510120029.jpg" width="200" height="100" border="0"></a></td>
      <td align="center"><a href=city.asp?vote=hn><img src="img/200510120033.jpg" width="200" height="100" border="0"></a></td>
    </tr>
    <tr>
      <td height="21" align="center">西藏得票数：<%=xz1%></td>
      <td align="center">挂林得票数：<%=gl1%></td>
      <td align="center">海南得票数：<%=hn1%></td>
    </tr>
    <tr>
      <td height="40" colspan="3" align="center">请用鼠标点击一下你所喜欢的城市图片，该城市就可获得你宝贵得一票！</td>
    </tr>
    <tr>
      <td height="40" colspan="3" align="center"><a href="citycount.asp">查看投票结果</a></td>
    </tr>
  </table>
  <p>&nbsp;</p>
</div>
</body>
</html>
