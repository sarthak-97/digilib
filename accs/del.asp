<%
dim n,p,found
n=request.form("t1")
n=trim(n)
found=0
dim cn,cs,rs
set cn=server.createobject("adodb.connection")
cn.open "provider=microsoft.jet.oledb.4.0;data source=G:\GDG\web\accs\db1.mdb"
set rs=server.createobject("adodb.recordset")
rs.open "s",cn,2,2
do until rs.eof
if rs("name")=n then 
found=1
rs.delete
response.write("record deleted")
exit do
end if
rs.movenext
loop
if found=0 then 
response.write("record not found")
end if

%>
 