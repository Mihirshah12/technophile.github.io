<%@Language= "vbscript"%>
<%
		dim conn,res
	set conn=Server.CreateObject("ADODB.Connection")
	conn.Provider= "Microsoft.Jet.OLEDB.4.0"
	conn.Open "C:\\inetpub\wwwroot\Web\website\DB.mdb"
	set res=Server.CreateObject("ADODB.Recordset")
	res.Open "Table1",conn,0,3,2

	a=Request.form("uname") 
	b=request.form("pass")
	Session("user")=a
	'sql="UPDATE Table1 SET Visits='" & res("Visits")+1 &"
	res.MoveFirst
	Do while not res.EOF
		If res("User")=a and res("Pass")=b then
			Response.Redirect "progress.asp"
		End if
		res.MoveNext
	Loop

	Response.Write "Invalid Username Or Password"
%>