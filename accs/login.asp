
<%
dim n,p
n=request.form("t1")
n=trim(n)
p=request.form("t2")
p=trim(p)

dim cn,cs,rs
set cn=server.createobject("adodb.connection")
cs="provider=microsoft.jet.oledb.4.0.;data source=c:\inetpub\wwwroot\web\accs\db1.mdb"
cn.open cs
set rs=server.createobject("adodb.recordset")
rs.open "s",cn
found=0
do while not rs.eof
if rs("name")=n and rs("pass")=p then
found=1
exit do 
end if 
rs.movenext
loop
if found=0 then
response.redirect("signin1.html")


end if 

%>



 