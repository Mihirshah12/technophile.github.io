<%@language="vbscript"%>
<%
dim obj
Set obj=Server.CreateObject("MSWC.AdRotator")
%>
<html>
	<body>
		please support my advertisement
	    	<%=obj.GetAdvertisement("schedulefile.txt")%>
		<p>welcome</p>
		<%=obj.GetAdvertisement("schedulefile.txt")%>
	</body>
</html>
