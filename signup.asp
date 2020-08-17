<%@ language="vbscript" %>
<%
	dim con,tsql,res
	set con = Server.createObject("ADODB.Connection")
	con.provider="Microsoft.Jet.OLEDB.4.0"
	con.open("C:\\inetpub\wwwroot\Web\website\DB.mdb")

	user=Request.Form("user")
	pass=Request.Form("pass")
	phone=Request.Form("phone")
	email=Request.Form("email")


	Session("user") = Request.Form("user")

	set res = Server.createObject("ADODB.Recordset")
	res.Open "Table1",con,0,3,2

	res.AddNew
	res("User") = user
	res("Pass") = pass
	res("Phone") = phone
	res("Email") = email
	res("Visits") = 1
	res.Update
	res.MoveNext
	Response.Redirect("progress.asp")
%>