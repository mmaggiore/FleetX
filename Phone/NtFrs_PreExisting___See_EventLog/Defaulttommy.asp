<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
	<!--	
	<bgsound src="file://\sounds\alert1.wav" >
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
		Response.Cookies("Phone")("PageStatus")=""
		Response.Cookies("Phone")("AliasCode")=""
		Response.Cookies("Phone")("FakeSubmit")=""
		'-----------------------------------------------------------------
			'Response.Write "VehicleID="&VehicleID&"<BR>"
				yyy=0
				if isnumeric(VehicleID) then
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				'Response.Write "vehicleID="&vehicleID&"<BR>"
				SQL = "SELECT fh_priority FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
				SQL = SQL&" WHERE (Fl_dr_ID='"&cInt(VehicleID)&"')"
				'''''If VehicleID<>124 AND VehicleID<>313 then
					SQL = SQL&" AND (((Fh_Status='OPN') AND (fl_sf_id<>'55')) OR"
					'''''else
					SQL = SQL&" ((Fh_Status='ARV') AND (fl_secacc is NULL))) "
				'''''End if
				'SQL = SQL&" ORDER BY Fh_Priority, fh_status desc, fh_id"
				'Response.write "<br><font color='blue'>SQL="&SQL&"<BR></font>"
				oRs.Open SQL, DATABASE, 1, 3
					Do while not oRs.eof 
						NewJob="y"
						yyy=yyy+1
						thepriority = oRs("fh_priority")
						If thepriority="P0" then
							PZero="y"
						End if
						'Response.Write "NumberOfOrders="&NumberOfOrders&"<BR>"

					oRs.movenext
					Loop
					oRs.Close
					Set oRs=Nothing	
					end if
						NumberOfOrders=yyy
						'Response.Write "NewJob="&NewJob&"<BR>"
						'Response.Write "VehicleID="&VehicleID&"<BR>"
						If trim(VehicleID)="xxx" then
							response.Write "GOT HERE!!!!<BR>"
							%>
							<bgsound src="file://\sounds\lswarn.wav">
							<%
						End if
						If NewJob="y" and PZero<>"y" then
							If trim(vehicleID)<>"1" and trim(vehicleID)<>"2" then
								%>
								<bgsound src="file://\sounds\lswarn.wav">
								<%else%>
								<bgsound src="sounds/newjob.wav">
								<%
							End if
						End if
						If PZero="y" then
							If trim(vehicleID)<>"1" and trim(vehicleID)<>"2" then
								%>
								<bgsound src="file://\sounds\lswarn.wav">
								<%else%>
								<bgsound src="sounds/PZero.wav">
								<%
							End if
						End if											
					BillToID=Request.Cookies("Phone")("sBT_ID")	
				if isnumeric(VehicleID) and BillToID=48 then
				'response.Write "got here<br>!"
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				'Response.Write "vehicleID="&vehicleID&"<BR>"
				SQL = "SELECT Count(Fl_SF_ID) as NumberOfPaper FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
				SQL = SQL&" WHERE (Fl_dr_ID='"&cInt(VehicleID)&"')"
				If VehicleID<>124 then
					SQL = SQL&" AND (Fh_Status='ACC')"
				End if
				'SQL = SQL&" ORDER BY Fh_Priority, fh_status desc, fh_id"
				'Response.write "<br><font color='blue'>SQL="&SQL&"<BR></font>"
				oRs.Open SQL, DATABASE, 1, 3
					If not oRs.eof then
						NumberOfPaper = oRs("NumberOfPaper")
					End if
					oRs.Close
					Set oRs=Nothing	
					end if						
		'-----------------------------------------------------------------
		'Response.Write "VehicleID="&VehicleID&"<br>"
		%>
	</head>
	<body>
	<table cellspacing="0" cellpadding="0" border="0" width="300">
		<tr>
			<td class="mainpagetextboldcenter">
				<font color="blue">Last update: <%=Time()%></font>
			</td>
		</tr>		
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
		<%If UnitID="1" or UnitID="2" Then%>
		<form method="post" action="DriverPOB.asp" ID="Form6">	
			<tr><td><input type="submit" value="Paper On Board  (<%=NumberOfPaper%>)" ID="Submit2" NAME="Submit1"></td></tr>
		</form>	
		<tr><td>&nbsp;</td></tr>
		<form method="post" action="DriverExceptions.asp" ID="Form7">	
			<tr><td><input type="submit" value="Exceptions" ID="Submit3" NAME="Submit1"></td></tr>
		</form>	
		<tr><td>&nbsp;</td></tr>		
		<%End if%>		
		<form method="post" action="DriverIfabPhoneEmulator.asp" ID="Form1">
			<tr><td><input type="submit" value="Drop Off/Pick Up"></td></tr>
			<input type="hidden" name="UserID" value="<%=UserID%>">
		</form>
		<tr><td>&nbsp;</td></tr>
		<form method="post" action="DriverTruckLoad.asp" ID="Form2">	
			<tr><td><input type="submit" value="Current Routing"></td></tr>
		</form>	
		<tr><td>&nbsp;</td></tr>
		<%
		'Response.Write "UnitID="&UnitID&"<BR>"
		If UnitID="303551" or UnitID="303552" or UnitID="303553" or UnitID="303554" or UnitID="srv" or UnitId="ofb" Then%>
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
