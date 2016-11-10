
<%
dim n,p
n=request.form("t1")
n=trim(n)
a=lcase(n)

dim cn,cs,rs
set cn=server.createobject("adodb.connection")
cs="provider=microsoft.jet.oledb.4.0.;data source=G:\GDG\web\accs\db1.mdb"
cn.open cs
set rs=server.createobject("adodb.recordset")
rs.open "books",cn
found=0
do while not rs.eof
if rs("Name")=a then
found=1
exit do 
end if 
rs.movenext
loop
if found=0 then
response.redirect("book2.asp")
elseif found=1 then
response.redirect("download.asp")
end if 
%>

