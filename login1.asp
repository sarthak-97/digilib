<%
dim n,p
n=request.form("t1")
n=trim(n)
q=request.form("t1")
q=trim(q)
a=ucase(q)
p=request.form("t2")
p=trim(p)
if len(n)<>0 then
dim cn,cs,rs
set cn=server.createobject("adodb.connection")
cs="provider=microsoft.jet.oledb.4.0.;data source=G:\GDG\web\accs\db1.mdb"
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
elseif found=1 then
response.redirect("book.asp")
end if 
else
response.redirect("signin2.html")
end if
%>
