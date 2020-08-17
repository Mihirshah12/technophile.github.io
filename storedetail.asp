<%@language="vbscript"%>
<%
	dim a,b,c,d
	a=Request.form("name")
	b=Request.form("age")
	c=Request.form("user")
	d=Request.form("pass")
	e=Request.form("ph")

	dim conn,res
	set conn=Server.CreateObject("ADODB.Connection")
	conn.provider= "Microsoft.Jet.OLEDB.4.0"
	conn.open "C:\\inetpub\wwwroot\Web\website\UserInfo.mdb"
	set res=Server.CreateObject("ADODB.Recordset")
	res.open "Table1",conn,0,3,2
	res.MoveFirst
	x=0
	Do while not res.EOF
		
	If res("Uname")=c then
		x=1
		res.MoveNext
	End If
Loop

	If x=1 then
		Response.write "Username Exists"
	else
	

	res.AddNew
	res("Name")=a
	res("Age")=b
	res("Uname")=c
	res("Password")=d
	res("Phone")=e
	res.Update

	response.redirect "Main.html"
end If


%>
