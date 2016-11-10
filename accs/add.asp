<%
dim n,p
n=request.form("t1")
n=trim(n)
p=request.form("t2")
p=trim(p)
q=request.form("t4")
q=trim(q)
dim cn,cs,rs
set cn=server.createobject("adodb.connection")
cn.open "provider=microsoft.jet.oledb.4.0;data source=c:\inetpub\wwwroot\web\accs\db1.mdb"
set rs=server.createobject("adodb.recordset")
rs.open "s",cn, 2,2
rs.addnew
rs("name")=n
rs("pass")=p
rs("age")=q
rs.update
response.write("record added")

%>
 