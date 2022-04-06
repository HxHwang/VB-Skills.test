<%
set conn=server.CreateObject("adodb.connection") 
conn.open "driver={microsoft access driver (*.mdb)};dbq="&server.MapPath("res.mdb") 
%>
