<%@ language="vbscript" %>
<html>

<body onload="move()">
	<progress id="progressbar" max="100" value="1" style="width: 80%; position: absolute; top: 300px; left: 60px"></progress>
	</br>
	
	<% 
		dim con,tsql,res
		set con = Server.createObject("ADODB.Connection")
		con.provider="Microsoft.Jet.OLEDB.4.0"
		con.open("C:\\inetpub\wwwroot\Web\website\DB.mdb")
		set res = Server.createObject("ADODB.Recordset")
		res.Open "Table1",con,0,3,2

		Response.Write("<p style='position: absolute; top: 166px; left: 593px; font-size: 22px;'>Hello "&Session("user")& "</p>   <br><br>") 

		Do while not res.EOF
			if res("User") = Session("user") then
				Response.Write("<p style='position: absolute; top: 210px; left: 580px; font-size: 22px;'>Your Visits: "&res("visits")& "</p><br>") 
			End if
			res.MoveNext
		Loop
	%>

	<p id="demo" style="position: absolute; top: 320px; left: 625px; font-size: 22px;"></p>
</body>

<script type="text/javascript">
	var progress = document.getElementById("progressbar");

	function move() 
	{
		var a = setInterval(function() {
    	progress.value = progress.value + 1;

    	document.getElementById("demo").innerHTML=progress.value

    	if (progress.value == 100) 
    	{
    		setTimeout(function(){
			window.location = "Main.html";
		}, 0)

       	}
  		}, 25);
	}
</script>
</html>