<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
	<!--
	<bgsound src="file://\windows\ringer.wav" loop="-1">
	-->
	<!--
	<EMBED src="file://\windows\ringer.wav" width="144" height="60" autostart="true" loop="false" hidden="true">
	-->
		<meta http-equiv="refresh" content="120" />
		<link rel="stylesheet" type="text/css" href="mainStyleSheet.css">
		<!-- #include file="../v9web/include/ifabsettings.inc" -->
		<!-- #include file="driverinfo.inc" -->	
		<title>Logisticorp Driver Home Page</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<%
		'-----------------------------------------------------------------
		'Response.Write "unitid="&UnitID&"<BR>"
		'Response.Write "vehicleid="&vehicleid&"<BR>"
		'Response.Write "driverid="&driverid&"<BR>"
		
				Set oRs = Server.CreateObject("ADODB.Command")
				with oRs
				oRs.ActiveConnection = DATABASE	
				oRs.CommandText = "exec fc_mdt.dbo.wc_count_opn_jobs"
				oRs.CommandType = adCmdStoredProc
				
' Define stored procedure params and append tocommand.

params.Append oRs.CreateParameter("@RETURN_VALUE", adInteger,adParamReturnValue, 0)
params.Append oRs.CreateParameter("&vehicleid&", adChar, adParamInput, 0)
				End with
			oRs.Execute
				Response.Write "oRs.CommandText="&oRs.CommandText&"<BR>"
				NumberOfOrders = oRs("@RETURN_VALUE")
		'-----------------------------------------------------------------
		Response.Write "NumberOfOrders="& NumberOfOrders &"<br>"
		%>
	</head>
	<body>
	<table cellspacing="0" cellpadding="0" border="0" width="300">
		<tr>
			<td class="mainpagetextboldcenter">
				Welcome&nbsp;&nbsp;<%=FirstName%>&nbsp;&nbsp;<%=LastName%>
				<br>
				Vehicle: <%=VehicleName%>
			</td>
		</tr>
		<tr><td>&nbsp;</td></tr>
		<form method="post" action="DriverAcknowledge.asp">
			<tr><td><input type="submit" value="Acknowledge Orders (<%=NumberOfOrders%>)"></td></tr>
		</form>	
		<tr><td>&nbsp;</td></tr>
		<form method="post" action="DriverIfabPhoneEmulator.asp" ID="Form1">
			<tr><td><input type="submit" value="Drop Off/Pick Up"></td></tr>
			<input type="hidden" name="UserID" value="<%=UserID%>">
		</form>
		<tr><td>&nbsp;</td></tr>
		<form method="post" action="DriverTruckLoad.asp" ID="Form2">	
			<tr><td><input type="submit" value="Current Routing"></td></tr>
		</form>	
		<tr><td>&nbsp;</td></tr>
		<%If UnitID="303551" or UnitID="303552" or UnitID="303553" or UnitID="303554" Then%>
		<form method="post" action="DriverHandOff.asp" ID="Form5">	
			<tr><td><input type="submit" value="Handoff a Job" ID="Submit1" NAME="Submit1"></td></tr>
		</form>	
		<tr><td>&nbsp;</td></tr>
		<%End if%>		
		<form method="post" action="DriverVehicle.asp" ID="Form3">
			<tr><td><input type="submit" value="Change Vehicle"></td></tr>
		</form>
		<tr><td>&nbsp;</td></tr>	
		<form method="post" action="DriverLogin.asp" ID="Form4">
			<tr><td><input type="submit" value="Log Out"></td></tr>
		</form>
	</table>
	</body>
</html>
