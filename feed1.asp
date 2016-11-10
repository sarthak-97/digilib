<HTML>
<HEAD><TITLE>ACCESS RECORD</TITLE>
</HEAD>
<%
    n=request.form("t1")
n=trim(n)
p=request.form("t2")
p=trim(p)
q=request.form("g1")
q=trim(q)
k=request.form("t3")
k=trim(k)
    if (len(n)=0) or (len(p)=0) or (len(k)=0) or (len(q)=0)then
response.redirect("feed3.asp")
else
Dim Cn ,Scn , Rs 
set Cn = Server.CreateObject ( "ADODB.Connection" )
Scn = "Provider=microsoft.Jet.OLEDB.4.0;Data Source=G:\GDG\web\accs\db1.mdb"
Cn.Open Scn
Set Rs = Server.CreateObject ( "ADODB.RecordSet" )
	Rs.open "feed",Cn, 2,2
	
                   Rs.Addnew
 		 Rs("name")=request.form("t1")
 		Rs("email")=request.form("t2")
 		Rs("rating")=request.form("g1")
 		Rs("sugg")=request.form("t3")
 		
                    RS.UPDATE      
                 RS.close
response.redirect("feed2.asp")
end if
      %>
  </html>