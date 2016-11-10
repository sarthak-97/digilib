<% 
dim con,rs,cs
set con=server.createobject("adodb.connection")
cs="provider=microsoft.jet.oledb.4.0 ;data source=G:\GDG\web\accs\db1.mdb"
con.open cs
response.write("connection successful")
response.write(" ")
set rs=server.createobject("ADODB.recordset")
rs.open "s", con
response.write("the login table contains ")
response.write("<br>")
while not rs.eof
response.write(rs("name") & " " &  rs("pass") & "<br>")
rs.movenext
wend
%>