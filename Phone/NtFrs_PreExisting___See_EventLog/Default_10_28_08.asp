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
''''''''''''''''''''''''''''''''''''''''''
SecureYes = Request.ServerVariables ("HTTPS")
'If SecureYes="off" then
If SecureYes="on" then
	''''''''''''''''''''''''''''''''''''''''''''
	something=Request.ServerVariables("HTTP_HOST") 
	Something=left(Something,4) 
	If lcase(something)="test" then 
		Response.redirect("http://test.logisticorp.us/phone/default.asp")
		'Response.Write "GOT HERE!!!" 
		MPMSendEmail="n" 
		else
		Response.redirect("http://www.logisticorp.us/phone/default.asp")
	End if 
	'Response.Write "Something="&Something&"<BR>"
	'Response.Write "MPMSendEmail="&MPMSendEmail&"<BR>"

	'''''''''''''''''''''''''''''''''''''''''''
		
End if
''''''''''''''''''''''''''''''''''''''''''		
		Response.Cookies("Phone")("PageStatus")=""
		Response.Cookies("Phone")("AliasCode")=""
		Response.Cookies("Phone")("FakeSubmit")=""
		mark=request.QueryString("mark")
		if mark="y" then
			'response.write "helloooo!"
			'response.write "VehicleID="&VehicleID&"<BR>"
		End if
		'-----------------------------------------------------------------
			'Response.Write "VehicleID="&VehicleID&"<BR>"
				Tommy=request.QueryString("Tommy")
				yyy=0
				if isnumeric(VehicleID) then
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				'Response.Write "vehicleID="&vehicleID&"<BR>"
				SQL = "SELECT fh_priority, fh_user5 FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
				SQL = SQL&" WHERE (Fl_dr_ID='"&cInt(VehicleID)&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL))"
				'''''If VehicleID<>124 AND VehicleID<>313 then
					SQL = SQL&" AND (((Fh_Status='OPN')) OR"
					'''''else
					'''''SQL = SQL&" ((Fh_Status='ARV') AND (fl_secacc is NULL) AND (fl_st_id<>'TOPPAN'))) "
					SQL = SQL&" ((Fh_Status='ARV') AND (fl_secacc is NULL))) "
				'''''End if
				'SQL = SQL&" ORDER BY Fh_Priority, fh_status desc, fh_id"
				'If mark="y" then
					'response.write "<br><font color='blue'>SQL="&SQL&"<BR></font>"
				'End if
				
				oRs.Open SQL, DATABASE, 1, 3
					Do while not oRs.eof 
						NewJob="y"
						yyy=yyy+1
						thepriority = oRs("fh_priority")
						MaterialType = oRs("FH_User5")
						If thepriority="P0" or MaterialType="Secure Waf" then
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
							'response.Write "GOT HERE!!!!<BR>"
							%>
							<bgsound src="file://\sounds\lswarn.wav">
							<%
						End if
						'response.Write "vehicleID="&vehicleID&"<br>"
						If NewJob="y" and PZero<>"y" then
							If trim(vehicleID)<>"1" and trim(vehicleID)<>"2" and trim(vehicleID)<>"145" then
								'Response.Write "vehicleID="&VehicleID&".<BR>"
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
				SQL = SQL&" WHERE (Fl_dr_ID='"&cInt(VehicleID)&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL))"
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
		<%If UnitID="1" or UnitID="2" or UnitID="3" or UnitID="4" or UnitID="5" or UnitID="6" or UnitID="7" Then%>
		<form method="post" action="DriverPOB.asp" ID="Form6">	
			<tr><td><input type="submit" value="Paper On Board  (<%=NumberOfPaper%>)" ID="Submit2" NAME="Submit1"></td></tr>
		</form>	
		<tr><td>&nbsp;</td></tr>
		<form method="post" action="DriverExceptions.asp" ID="Form7">	
			<tr><td><input type="submit" value="Exceptions" ID="Submit3" NAME="Submit1"></td></tr>
		</form>	
		<tr><td>&nbsp;</td></tr>
		<form method="post" action="DriverInterimScan.asp" ID="Form11">	
			<tr><td><input type="submit" value="Interim Shipments" ID="Submit7" NAME="Submit1"></td></tr>
		</form>	
		<tr><td>&nbsp;</td></tr>				
		<%End if
		'Response.Write "UnitID="&UnitID&"<BR>"
		If UnitID="AIMS1" or UnitID="AIMS2" Then
		
		
		
			if isnumeric(VehicleID) and BillToID=75 then
				'response.Write "got here<br>!"
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				'Response.Write "vehicleID="&vehicleID&"<BR>"
				SQL = "SELECT Count(Distinct(Fl_FH_ID)) as BOLNeeded FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN fcrefs ON fclegs.fl_fh_id = fcrefs.rf_fh_id"
				SQL = SQL&" WHERE (Fl_dr_ID='"&cInt(VehicleID)&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL))"
				If VehicleID<>124 then
					SQL = SQL&" AND (Fh_Status='ACC') AND (fcrefs.rf_box = '')"
				End if
				'SQL = SQL&" ORDER BY Fh_Priority, fh_status desc, fh_id"
				'Response.write "<br><font color='blue'>SQL="&SQL&"<BR></font>"
				oRs.Open SQL, DATABASE, 1, 3
					If not oRs.eof then
						BOLNeeded = oRs("BOLNeeded")
					End if
					oRs.Close
					Set oRs=Nothing	
			end if
					%>						
		
		
		
		<form method="post" action="DriverBOL.asp" ID="Form8">	
			<tr><td><input type="submit" value="Create a BOL (<%=BOLNeeded%>)" ID="Submit4" NAME="Submit1"></td></tr>
		</form>	
		<tr><td>&nbsp;</td></tr>
		<%End if
		If BillToID=75 then
		%>
		<form method="post" action="DriverAIMSLocations.asp" ID="Form9">
			<tr><td><input type="submit" value="Drop Off/Pick Up" ID="Submit5" NAME="Submit5"></td></tr>
			<input type="hidden" name="UserID" value="<%=UserID%>" ID="Hidden1">
		</form>		
		<%
		Else
		%>				
		<form method="post" action="DriverIfabPhoneEmulator.asp" ID="Form1">
			<tr><td><input type="submit" value="Drop Off/Pick Up"></td></tr>
			<input type="hidden" name="UserID" value="<%=UserID%>">
		</form>
		<%
		End if
		%>
		<tr><td>&nbsp;</td></tr>
		<form method="post" action="DriverTruckLoad.asp" ID="Form2">	
			<tr><td><input type="submit" value="Current Routing"></td></tr>
		</form>	
		<tr><td>&nbsp;</td></tr>
		<%
		'Response.Write "BillToID="&BillToID&"<BR>"
		'Response.Write "UnitID="&UnitID&"<BR>"
		
		If BillToID="75" then%>
			<form method="post" action="DriverExceptions.asp" ID="Form10">	
				<tr><td><input type="submit" value="Exceptions" ID="Submit6" NAME="Submit1"></td></tr>
			</form>	
			<tr><td>&nbsp;</td></tr>		
			<%
		End if
		'Response.Write "UnitID="&UnitID&"<BR>"
		If UnitID="303551" or UnitID="303552" or UnitID="303553" or UnitID="303554" or UnitID="srv" or UnitId="ofb" or UnitID="SHERMAN" or UnitID="OCV" or lcase(UnitID)="srv" Then%>
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
