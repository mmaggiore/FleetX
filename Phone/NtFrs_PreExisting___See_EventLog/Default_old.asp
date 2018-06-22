<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<meta http-equiv="refresh" content="120" />
		<link rel="stylesheet" type="text/css" href="mainStyleSheet.css">
		<!-- #include file="../v9web/include/ifabsettings.inc" -->
		<!-- #include file="driverinfo.inc" -->	
		<title>Logisticorp Driver Home Page</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<%
		'-----------------------------------------------------------------
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				
				SQL = "SELECT Count(Fl_SF_ID) as NumberOfOrders FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
				SQL = SQL&" WHERE (Fl_dr_ID='"&VehicleID&"')"
				SQL = SQL&" AND (Fh_Status='OPN')"
				'SQL = SQL&" ORDER BY Fh_Priority, fh_status desc, fh_id"
				'Response.write "<br><font color='blue'>SQL="&SQL&"<BR></font>"
				oRs.Open SQL, DATABASE, 1, 3
					If not oRs.eof then
						NumberOfOrders = oRs("NumberOfOrders")
						'Response.Write "NumberOfOrders="&NumberOfOrders&"<BR>"
					End if
					oRs.Close		
		'-----------------------------------------------------------------
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
