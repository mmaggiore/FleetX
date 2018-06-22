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
''''''''''''SETS VEHICLE JOB COUNTS ETC. TO ZERO''''''''''''
NumJobs303551=0
NumJobs303552=0
NumJobs303553=0
NumJobsSRVRFAB=0
XP303551=0
XP303552=0
XP303553=0
XPSRVRFAB=0
LATE303551=0
LATE303552=0
LATE303553=0
LATESRVRFAB=0
'If SecureYes="off" then
If SecureYes="on" then
	''''''''''''''''''''''''''''''''''''''''''''
	''''''''THIS NEEDS TO BE CHANGED WHEN MOVED TO PRODUCTION!''''''''
	something=Request.ServerVariables("HTTP_HOST") 
	Something=left(Something,4) 
	If lcase(something)="test" then 
		'''Response.redirect("http://test.logisticorp.us/h5redo/phone/defaultc.asp")
		'Response.Write "GOT HERE!!!" 
		MPMSendEmail="n" 
		else
		'''Response.redirect("http://www.logisticorp.us/h5redo/phone/defaultc.asp")
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
		'-----FIND THE JOBS THAT NEED TO BE ACKNOWLEDGED------------------------------------------------------------
			'Response.Write "VehicleID="&VehicleID&"<BR>"
				Tommy=request.QueryString("Tommy")
				yyy=0
				if isnumeric(VehicleID) then
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				'Response.Write "vehicleID="&vehicleID&"<BR>"
				SQL = "SELECT fh_priority, fh_user5, fl_un_id FROM VW_Driver_FindOpenJobs"
				SQL = SQL&" WHERE (Fl_DR_ID='"&VehicleID&"')"
				'response.write "<br><font color='blue'>XXXSQL="&SQL&"<BR></font>"
				'End if
				
				oRs.Open SQL, DATABASE, 1, 3
					Do while not oRs.eof 
						NewJob="y"
						yyy=yyy+1
						thepriority = oRs("fh_priority")
						MaterialType = oRs("FH_User5")
						'fl_rt_type = trim(oRs("fl_rt_type"))
						If thepriority="P0" or MaterialType="Secure Waf" or MaterialType="ITAR" then
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
		''''''''''''NEW VIEW TO SHOW THE OPEN ORDERS/PO's/LATES FOR THE OTHER VEHICLES''''''''''
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				'Response.Write "vehicleID="&vehicleID&"<BR>"
				SQL = "SELECT fh_priority, fh_user5, fl_un_id, fl_st_rta, fl_dr_id FROM VW_Driver_FindAllOpenJobsForAllWaferTrucks"
				SQL = SQL&" WHERE (Fl_DR_ID IN ('312','313','212', '113'))"
				'response.write "<br><font color='Red'>XXXSQL="&SQL&"<BR></font>"
				'End if

				oRs.Open SQL, DATABASE, 1, 3
					Do while not oRs.eof 
						NewJob="y"
						X66=X66+1
				       ' Response.Write "X66="&X66&"<BR>"
						AllPriority = oRs("fh_priority")
						AllMaterialType = oRs("FH_User5")
						AllVehicle = Trim(oRs("fl_dr_id"))
						AllDueTime = oRs("fl_st_rta")
						'Response.Write "****AllVehicle="&AllVehicle&"<BR>"
						Select Case AllVehicle
						    Case "312"
						        NumJobs303551=NumJobs303551+1
						        If AllPriority="P0" or AllPriority="P1" or AllPriority="XP" or AllMaterialType="Secure Waf" or AllMaterialType="ITAR" then
						            XP303551=XP303551+1
						        End if
						        If AllDueTime<Now() then
						            Late303551=Late303551+1
						        End if
						        'Response.Write "Got here 1<br>"
						        'Response.Write "NumJobs303551="&NumJobs303551&"<BR>"
						        'Response.Write "XP303551="&XP303551&"<BR>"
						        'Response.Write "Late303551="&Late303551&"<BR>"							        
						    Case "313"
						        NumJobs303552=NumJobs303552+1
						        If AllPriority="P0" or AllPriority="P1" or AllPriority="XP" or AllMaterialType="Secure Waf" or AllMaterialType="ITAR" then
						            XP303552=XP303552+1
						        End if	
						        If AllDueTime<Now() then
						            Late303552=Late303552+1
						        End if	
						        'Response.Write "Got here 2<br>"
						        'Response.Write "NumJobs303552="&NumJobs303552&"<BR>"
						        'Response.Write "XP303552="&XP303552&"<BR>"
						        'Response.Write "Late303552="&Late303552&"<BR>"					        					        
						    Case "212"
						        NumJobs303553=NumJobs303553+1
						        If AllPriority="P0" or AllPriority="P1" or AllPriority="XP" or AllMaterialType="Secure Waf" or AllMaterialType="ITAR" then
						            XP303553=XP303553+1
						        End if
						        If AllDueTime<Now() then
						            Late303553=Late303553+1
						        End if
						        'Response.Write "Got here 3<br>"	
						        'Response.Write "NumJobs303553="&NumJobs303553&"<BR>"
						        'Response.Write "XP303553="&XP303553&"<BR>"
						        'Response.Write "Late303553="&Late303553&"<BR>"							        					        						        
						    Case "113"
						        
						        NumJobsSRVRFAB=NumJobsSRVRFAB+1
						        If AllPriority="P0" or AllPriority="P1" or AllPriority="XP" or AllMaterialType="Secure Waf" or AllMaterialType="ITAR" then
						            XPSRVRFAB=XPSRVRFAB+1
						        End if
						        If AllDueTime<Now() then
						            LateSRVRFAB=LateSRVRFAB+1
						        End if	
						        
						        'Response.Write "Got here 4<br>"
						        'Response.Write "NumJobsSRVRFAB="&NumJobsSRVRFAB&"<BR>"
						        'Response.Write "XPSRVRFAB="&XPSRVRFAB&"<BR>"
						        'Response.Write "LateSRVRFAB="&LateSRVRFAB&"<BR>"							        					        						        
						End Select
						
						'fl_rt_type = trim(oRs("fl_rt_type"))
						'Response.Write "NumberOfOrders="&NumberOfOrders&"<BR>"

					oRs.movenext
					Loop
					oRs.Close
					Set oRs=Nothing	
						
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
		

		'Response.Write "AlertMessage="&AlertMessage&"****<BR>"
		If trim(alertmessage)>"" then
			Response.Write "<font color='red'><br>*******************<BR>"
			Response.Write alertmessage&"<BR>"
			Response.Write "*******************<BR></font>"
			If   NewJob<>"y" and PZero<>"y" then
				If trim(vehicleID)<>"1" and trim(vehicleID)<>"2" then
					%>
					<bgsound src="file://\sounds\cshk-effect.wav">
					<%else%>
					<bgsound src="sounds/PZero.wav">
					<%
				End if
				else
				
			End if
			%>		
			<!--
				<SCRIPT type="text/javascript" language="JavaScript">	
					var AlertMessage="<%=alertmessage%>";
					alert(AlertMessage);
				</script>		
			-->
			
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
		<%If (UnitID="303551" OR UnitID="303552" Or UnitID="303553" Or UnitID="303554" Or UnitID="srvrfab") then 
		If NumJobs303551>10 then JobFont303551="red" else JobFont303551="Green" end if 
		If NumJobs303552>10 then JobFont303552="red" else JobFont303552="Green" end if 
		If NumJobs303553>10 then JobFont303553="red" else JobFont303553="Green" end if 
		If NumJobsSRVRFAB>10 then JobFontSRVRFAB="red" else JobFontSRVRFAB="Green" end if 
		If XP303551>0 then XPFont303551="red" else XPFont303551="Green" end if 
		If XP303552>0 then XPFont303552="red" else XPFont303552="Green" end if 
		If XP303553>0 then XPFont303553="red" else XPFont303553="Green" end if 
		If XPSRVRFAB>0 then XPFontSRVRFAB="red" else XPFontSRVRFAB="Green" end if 
		If Late303551>0 then LateFont303551="red" else LateFont303551="Green" end if 
		If Late303552>0 then LateFont303552="red" else LateFont303552="Green" end if 
		If Late303553>0 then LateFont303553="red" else LateFont303553="Green" end if 
		If LateSRVRFAB>0 then LateFontSRVRFAB="red" else LateFontSRVRFAB="Green" end if 		
		'Response.Write "unitID="&UnitID&"<BR>"
		%>
		    <tr><td>&nbsp;</td></tr>
		    <%'''If UnitID<>"303551" then %>
		        <tr><td class="mainpagetextbold">W1: <font color="<%=JobFont303551%>">Jobs=<%=NumJobs303551%></font> / <font color="<%=XPFont303551%>">P0=<%=XP303551%></font> / <font color="<%=LateFont303551%>">Late=<%=Late303551%></font></td></tr>
		    <%'''End if%>
		    <%'''If UnitID<>"303552" then %>
		        <tr><td class="mainpagetextbold">W2: <font color="<%=JobFont303552%>">Jobs=<%=NumJobs303552%></font> / <font color="<%=XPFont303552%>">P0=<%=XP303552%></font> / <font color="<%=LateFont303552%>">Late=<%=Late303552%></font></td></tr>
		    <%'''End if%>
		    <%'''If UnitID<>"303553" then %>
		        <tr><td class="mainpagetextbold">W3: <font color="<%=JobFont303553%>">Jobs=<%=NumJobs303553%></font> / <font color="<%=XPFont303553%>">P0=<%=XP303553%></font> / <font color="<%=LateFont303553%>">Late=<%=Late303553%></font></td></tr>
		    <%'''End if%>
		    <%If ucase(UnitID)="SRVRFAB" then %>
		        <tr><td class="mainpagetextbold">RFAB: <font color="<%=JobFontSRVRFAB%>">Jobs=<%=NumJobsSRVRFAB%></font> / <font color="<%=XPFontSRVRFAB%>">P0=<%=XPSRVRFAB%></font> / <font color="<%=LateFontSRVRFAB%>">Late=<%=LateSRVRFAB%></font></td></tr>
		    <%End if%>
		<%end if %>
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
		<%If trim(vehicleID)="312" or trim(vehicleID)="313" or trim(vehicleID)="212" or trim(vehicleID)="314" then
	'''''''''''''FOR LUNCH/BREAK CODE...IF THEY LEAVE THE PAGE TO CIRCUMVENT THE BREAK PAGE''''''''''
			Set oConn = Server.CreateObject("ADODB.Connection")
			oConn.ConnectionTimeout = 100
			oConn.Provider = "MSDASQL"
			oConn.Open DATABASE
				l_cSQL = "UPDATE breaktable SET EndTime = '"& Now() &"' WHERE Userid = '" & UserID & "' and EndTime is NULL"
				oConn.Execute(l_cSQL)
			Set oConn=Nothing
			'Response.Write "l_cSQL="&l_cSQL&"<BR>"	
		%>
		<tr><td>&nbsp;</td></tr>			
		<form method="post" action="DriverBreak.asp?a=l&b=<%=DateAdd("n",30,now())%>&c=y" ID="Form14">
			<tr><td><input type="submit" value="Take Lunch"></td></tr>
		</form>
		<tr><td>&nbsp;</td></tr>			
		<form method="post" action="DriverBreak.asp?a=b&b=<%=DateAdd("n",15,now())%>&c=y" ID="Form15">
			<tr><td><input type="submit" value="Take a Break"></td></tr>
		</form>				
		<%end if%>
		<tr><td>&nbsp;</td></tr>			
		<form method="post" action="DriverLogin.asp" ID="Form4">
			<tr><td><input type="submit" value="Log Out"></td></tr>
		</form>
	</table>

	</body>
</html>
