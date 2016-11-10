<%
dim n,p
n=request.form("t1")
n=trim(n)
p=request.form("t2")
p=trim(p)
q=request.form("t4")
q=trim(q)
k=request.form("t5")
k=trim(k)

if (len(n)=0) or (len(p)=0) or (len(k)=0) then
response.redirect("addrec3.html")
else
dim cn,cs,rs
set cn=server.createobject("adodb.connection")
cn.open "provider=microsoft.jet.oledb.4.0;data source=G:\GDG\web\accs\db1.mdb"
set rs=server.createobject("adodb.recordset")
rs.open "s",cn, 2,2
if p=k then
rs.addnew
rs("name")=n
rs("pass")=p
rs("age")=q
rs.update
response.redirect("register.html")
else 
response.redirect("addrec2.html")
end if
end if
%>
 