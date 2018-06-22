<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<link rel="stylesheet" type="text/css" href="mainStyleSheet.css">
		<title>Logisticorp Driver Home Page</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<%
		DriverID=Request.Cookies("DriverInfo")("DriverID")
		If DriverID="" then
			Response.Redirect("DriverLogin.asp")
		End if
		%>
	</head>
	<body>
	<table cellspacing="0" cellpadding="0" border="0">
		<tr>
			<td>
				Driver #<%=DriverID%> Home Page (<%=Request.Cookies("Phone")("BT_ID")%>))
			</td>
		</tr>
		<tr><td>&nbsp;</td></tr>
		<tr><td><a href="orderentry/DriverIfabPhoneEmulator.asp?DriverID=<%=DriverID%>">Change Order Status</a></td></tr>
		<tr><td><a href="Link2.asp">Current Truck Load</a></td></tr>
		<tr><td><a href="DriverLogin.asp">Log in as New Driver</a></td></tr>
	</table>
	</body>
</html>
