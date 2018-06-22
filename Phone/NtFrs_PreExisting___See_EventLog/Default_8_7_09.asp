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
		<META HTTP-EQUIV="PRAGMA" CONTENT="NO-CACHE">
		<meta http-equiv="refresh" content="120" />

<script language="JavaScript">
<!--

var sURL = unescape(window.location.pathname);

function doLoad()
{
    // the timeout value should be the same as in the "refresh" meta-tag
    setTimeout( "refresh()", 130*1000 );
}

function refresh()
{
    //  This version of the refresh function will cause a new
    //  entry in the visitor's history.  It is provided for
    //  those browsers that only support JavaScript 1.0.
    //
    window.location.href = sURL;
}

//-->
</script>

<script language="JavaScript1.1">
<!--
function refresh()
{
    //  This version does NOT cause an entry in the browser's
    //  page view history.  Most browsers will always retrieve
    //  the document from the web-server whether it is already
    //  in the browsers page-cache or not.
    //  
    window.location.replace( sURL );
}
//-->
</script>

<script language="JavaScript1.2">
<!--
function refresh()
{
    //  This version of the refresh function will be invoked
    //  for browsers that support JavaScript version 1.2
    //
    
    //  The argument to the location.reload function determines
    //  if the browser should retrieve the document from the
    //  web-server.  In our example all we need to do is cause
    //  the JavaScript block in the document body to be
    //  re-evaluated.  If we needed to pull the document from
    //  the web-server again (such as where the document contents
    //  change dynamically) we would pass the argument as 'true'.
    //  
    window.location.reload( false );
}
//-->
</script>
		
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
			response.write "helloooo!"
			response.write "VehicleID="&VehicleID&"<BR>"
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
				SQL = "SELECT fh_priority, fh_user5, fl_rt_type FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
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
						'fl_rt_type = trim(oRs("fl_rt_type"))
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
					'response.Write "BillToID="&BillToID&"<BR>"
					If not isnumeric(BillToID) then
						Response.redirect("driverlogin.asp")
					End if
				if isnumeric(VehicleID) and (BillToID=48 or BillToID=80) then
				'response.Write "got here<br>!"
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				'Response.Write "vehicleID="&vehicleID&"<BR>"
				SQL = "SELECT Count(Fl_SF_ID) as NumberOfPaper FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
				If trim(vehicleID)="199" then
					SQL = SQL&" WHERE (Fl_dr_ID='"&cInt(VehicleID)&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL))"
					SQL = SQL&" AND (Fh_Status='ONB') and (fl_rt_type<>'out')"
					else
					SQL = SQL&" WHERE (Fl_dr_ID='"&cInt(VehicleID)&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL))"
				End if
				If VehicleID<>124 and VehicleID<>199 then
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
<body onload="doLoad()">


	
	<table cellspacing="0" cellpadding="0" border="0" width="300">
		<%If trim(vehicleID)="312" or trim(vehicleID)="313" or trim(vehicleID)="212" then%>
			<tr>
				<td class="mainpagetextboldcenter">
					<font color="black">Last update: <%=Time()%></font>
				</td>
			</tr>
			<tr><td class="mainpagetextboldcenter"><font color="blue">
			<%else%>
			<tr>
				<td class="mainpagetextboldcenter">
					<font color="blue">Last update: <%=Time()%></font>
				</td>
			</tr>
			<tr><td class="mainpagetextboldcenter"><font color="blue">			
			
			<%
		End if	
		
		
		
		
		'response.Write "vehicleid="&vehicleID&"<BR>"
		If trim(vehicleID)="312" or trim(vehicleID)="313" or trim(vehicleID)="212" then
			''''''''''''FIND P0'S'''''''''''''''''''''''''''''''''''''''''''''''''''
			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.CursorLocation = 3
			oRs.CursorType = 3
			oRs.ActiveConnection = DATABASE	
			'Response.Write "vehicleID="&vehicleID&"<BR>"
			SQL = "SELECT OrdersInVehicle from Mark_Wafer3_P0"
			oRs.Open SQL, DATABASE, 1, 3
				If not oRs.eof then
					OrdersInVehicle = oRs("OrdersInVehicle")
					If OrdersInVehicle>0 then
						Wafer3P0="yes"
					End if
				End if
			oRs.Close
			Set oRs=Nothing
			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.CursorLocation = 3
			oRs.CursorType = 3
			oRs.ActiveConnection = DATABASE	
			'Response.Write "vehicleID="&vehicleID&"<BR>"
			SQL = "SELECT OrdersInVehicle from Mark_Wafer2_P0"
			oRs.Open SQL, DATABASE, 1, 3
				If not oRs.eof then
					OrdersInVehicle = oRs("OrdersInVehicle")
					If OrdersInVehicle>0 then
						Wafer2P0="yes"
					End if
				End if
			oRs.Close
			Set oRs=Nothing
			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.CursorLocation = 3
			oRs.CursorType = 3
			oRs.ActiveConnection = DATABASE	
			'Response.Write "vehicleID="&vehicleID&"<BR>"
			SQL = "SELECT OrdersInVehicle from Mark_Wafer1_P0"
			oRs.Open SQL, DATABASE, 1, 3
				If not oRs.eof then
					OrdersInVehicle = oRs("OrdersInVehicle")
					If OrdersInVehicle>0 then
						Wafer1P0="yes"
					End if
				End if
			oRs.Close
			Set oRs=Nothing	
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			''''''''''''FIND LATE ORDERS
			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.CursorLocation = 3
			oRs.CursorType = 3
			oRs.ActiveConnection = DATABASE	
			'Response.Write "vehicleID="&vehicleID&"<BR>"
			SQL = "SELECT OrdersInVehicle from Mark_Wafer3_Late"
			oRs.Open SQL, DATABASE, 1, 3
				If not oRs.eof then
					OrdersInVehicle = oRs("OrdersInVehicle")
					If OrdersInVehicle>0 then
						AlertMessage=AlertMessage&"Wafer 3 has a late order."
					End if
				End if
			oRs.Close
			Set oRs=Nothing
			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.CursorLocation = 3
			oRs.CursorType = 3
			oRs.ActiveConnection = DATABASE	
			'Response.Write "vehicleID="&vehicleID&"<BR>"
			SQL = "SELECT OrdersInVehicle from Mark_Wafer2_Late"
			oRs.Open SQL, DATABASE, 1, 3
				If not oRs.eof then
					OrdersInVehicle = oRs("OrdersInVehicle")
					If OrdersInVehicle>0 then
						AlertMessage=AlertMessage&"Wafer 2 has a late order."
					End if
				End if
			oRs.Close
			Set oRs=Nothing
			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.CursorLocation = 3
			oRs.CursorType = 3
			oRs.ActiveConnection = DATABASE	
			'Response.Write "vehicleID="&vehicleID&"<BR>"
			SQL = "SELECT OrdersInVehicle from Mark_Wafer1_Late"
			oRs.Open SQL, DATABASE, 1, 3
				If not oRs.eof then
					OrdersInVehicle = oRs("OrdersInVehicle")
					If OrdersInVehicle>0 then
						AlertMessage=AlertMessage&"Wafer 1 has a late order."
					End if
				End if
			oRs.Close
			Set oRs=Nothing				
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''		
		
		
		
		
		
			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.CursorLocation = 3
			oRs.CursorType = 3
			oRs.ActiveConnection = DATABASE	
			'Response.Write "vehicleID="&vehicleID&"<BR>"
			SQL = "SELECT OrdersInVehicle from Mark_Wafer1_Orders"
			oRs.Open SQL, DATABASE, 1, 3
				If not oRs.eof then
					OrdersInVehicle = oRs("OrdersInVehicle")
					If Wafer1P0="yes" then
						Response.Write "<font color='red'>"
						else
						Response.Write "<font color='blue'>"
					End if
					Response.Write "Wafer1:"&OrdersInVehicle&"</font>&nbsp;/&nbsp;"
					'If OrdersInVehicle>0 then
						'AlertMessage=AlertMessage&"Wafer 1 has a late order."
					'End if
				End if
			oRs.Close
			Set oRs=Nothing	
			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.CursorLocation = 3
			oRs.CursorType = 3
			oRs.ActiveConnection = DATABASE	
			'Response.Write "vehicleID="&vehicleID&"<BR>"
			SQL = "SELECT OrdersInVehicle from Mark_Wafer2_Orders"
			oRs.Open SQL, DATABASE, 1, 3
				If not oRs.eof then
					OrdersInVehicle = oRs("OrdersInVehicle")
					If Wafer2P0="yes" then
						Response.Write "<font color='red'>"
						else
						Response.Write "<font color='blue'>"
					End if					
					Response.Write "Wafer2:"&OrdersInVehicle&"</font>&nbsp;/&nbsp;"
					'If OrdersInVehicle>0 then
						'AlertMessage=AlertMessage&"Wafer 2 has a late order."
					'End if
				End if
			oRs.Close
			Set oRs=Nothing	
			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.CursorLocation = 3
			oRs.CursorType = 3
			oRs.ActiveConnection = DATABASE	
			'Response.Write "vehicleID="&vehicleID&"<BR>"
			SQL = "SELECT OrdersInVehicle from Mark_Wafer3_Orders"
			oRs.Open SQL, DATABASE, 1, 3
				If not oRs.eof then
					OrdersInVehicle = oRs("OrdersInVehicle")
					If Wafer3P0="yes" then
						Response.Write "<font color='red'>"
						else
						Response.Write "<font color='blue'>"
					End if					
					Response.Write "Wafer3:"&OrdersInVehicle&"</font>"

				End if
			oRs.Close
			Set oRs=Nothing
		End if
		'Response.Write "AlertMessage="&AlertMessage&"****<BR>"
		If trim(alertmessage)>"" then
		%>
			<SCRIPT type="text/javascript" language="JavaScript">	
				var AlertMessage="<%=alertmessage%>";
				alert(AlertMessage);
			</script>		
		<%
		end if							
		%>
		</font>
		</td></tr>				
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
		<%
		If UnitID="1" or UnitID="2" or UnitID="3" or UnitID="4" or UnitID="5" or UnitID="6" or UnitID="7" or UnitID="199" or UnitID="198" Then
		'response.Write "billtoid="&billtoid&"<BR>"
		If BillToID="80" then
		%>
		<form method="post" action="DriverAIMSPOBLocations.asp" ID="Form12">
		<%
		else
		%>
		<form method="post" action="DriverPOB.asp" ID="Form6">
		<%
		end if
		%>	
			<tr><td><input type="submit" value="Paper On Board  (<%=NumberOfPaper%>)" ID="Submit2" NAME="Submit1"></td></tr>
		</form>	
		<tr><td>&nbsp;</td></tr>
		<form method="post" action="DriverExceptions.asp" ID="Form7">	
			<tr><td><input type="submit" value="Exceptions" ID="Submit3" NAME="Submit1"></td></tr>
		</form>	
		<tr><td>&nbsp;</td></tr>
		<%
		End if
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
		'response.Write "BillToID="&BillToID&"<BR>"
		If BillToID=75 or BillToID=80 then
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
		If UnitID="303551" or UnitID="303552" or UnitID="303553" or UnitID="303554" or UnitID="srv" or UnitId="ofb" or UnitID="SHERMAN" or UnitID="OCV" or lcase(UnitID)="srv" or lcase(UnitID)="1" or lcase(UnitID)="2" or lcase(UnitID)="3" or lcase(UnitID)="4" or lcase(UnitID)="5" or lcase(UnitID)="6" or lcase(UnitID)="7" Then%>
		<form method="post" action="DriverHandOff.asp" ID="Form5">	
			<tr><td><input type="submit" value="Handoff a Job" ID="Submit1" NAME="Submit1"></td></tr>
		</form>	
		<tr><td>&nbsp;</td></tr>
		<%End if%>	
		<%if UnitID="1" or UnitID="2" then%>		
			<form method="post" action="DriverInterimScan.asp" ID="Form11">	
				<tr><td><input type="submit" value="Interim Shipments" ID="Submit7" NAME="Submit1"></td></tr>
			</form>	
			<tr><td>&nbsp;</td></tr>				
		<%End if%>			
		<form method="post" action="DriverVehicle.asp" ID="Form3">
			<tr><td><input type="submit" value="Change Vehicle"></td></tr>
		</form>
		<tr><td>&nbsp;</td></tr>
		<form method="post" action="DriverPhoneList.asp" ID="Form13">
			<tr><td><input type="submit" value="Phone List" ID="Submit8" NAME="Submit8"></td></tr>
		</form>
		<tr><td>&nbsp;</td></tr>			
		<form method="post" action="DriverLogin.asp" ID="Form4">
			<tr><td><input type="submit" value="Log Out"></td></tr>
		</form>
	</table>
	</body>
</html>
