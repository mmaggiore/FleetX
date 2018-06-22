<%@ Language=VBScript %>
<!-- 10/08/2005 Span-Renuka included this screen for CCF 2023 -->
<!-- #include file="../v9web/include/checkstring.inc" -->
<!-- #include file="../v9web/include/custom.inc" -->
<!-- #include file="../v9web/include/ifabsettings.inc" -->
<!-- #include file="driverinfo.inc" -->	
<HTML>
	<HEAD>
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
		<link rel="stylesheet" type="text/css" href="mainStyleSheet.css">
		<%
		DriverID=Request.Form("DriverID")
		%>
	</HEAD>
	<BODY leftMargin=0 TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<table width="300" cellpadding="0" cellspacing="0" border="0" align="left" ID="Table2">
			<tr><td align="center" colspan="9"><a href="default.asp" class="mainpagelink"><form method="post" action="default.asp" ID="Form8"><input type="submit" value="Return to Menu" ID="Submit3" NAME="Submit3"></form></td></tr>
			<tr>
				<td align="center" colspan="13" class="purpleseparator"><b>CURRENT STATUS OF <%=uCase(VehicleName)%></b></td>
			</tr>						
			<tr>
				<td align="center" class="purpleseparator" colspan="13"><b>ORDERS IN VEHICLE</b></td>
			</tr>		
			<tr>
				<!--td align="center" colspan="2">&nbsp;</td-->						
				<td align="center" nowrap><b>Jobs</b></td>
				<td width="5">&nbsp;</td>
				<td align="center" nowrap><b>Due In</b></td>
				<td width="5">&nbsp;</td>
				<td align="center" nowrap><b>To</b></td>
			</tr>
			<%
			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.CursorLocation = 3
			oRs.CursorType = 3
			oRs.ActiveConnection = DATABASE	
			SQL = "SELECT Fl_SF_ID, Fh_ID, Fl_ST_ID, fl_st_rta, fh_bt_id, FH_Status, Fh_Priority FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
			SQL = SQL&" WHERE (Fl_dr_ID='"&VehicleID&"') AND (fh_ship_dt>'"&now()-7&"')"
			SQL = SQL&" AND ((fh_status='ONB'))"
			SQL = SQL&" ORDER BY fh_priority, fl_st_rta"
			'Response.write "SQL="&SQL&"<BR>"
			oRs.Open SQL, DATABASE, 1, 3
			If not oRs.EOF then
					CloseTable="y"
					ELSE
					Response.Write "<tr><td colspan='13' align='center'>There are currently no orders in the vehicle.</td></tr><tr><td>&nbsp;</td></tr>"
			End if
			Do while not oRs.eof
				X=X+1
				FromLocation = oRs("Fl_SF_ID")
				JobNumber = oRs("Fh_ID")
				ToLocation = oRs("Fl_ST_ID")
				'Response.Write "ToLocation="&ToLocation&"<BR>"
				'Response.Write "JobNumber="&JobNumber&"<BR>"
				JobStatus = oRs("fh_status")
				Priority = oRs("fh_priority")
				If FColor2="" and Priority="P1" then
					FColor2="purple"
				End if
				If Priority="P0" then
					FColor2="red"
				End if				
				fh_bt_id=Trim(cStr(oRs("fh_bt_id")))
				DueTime=oRs("fl_st_rta")
				TimeTillDue=DateDiff("n",now(),DueTime)	
				If TimeTillDue<0 then
					DisplayTimeTillDue="LATE"
					Else
					HoursTillDue=Int(TimeTillDue/60)
					MinutesTillDue=TimeTillDue-(HoursTillDue*60)
					DisplayTimeTillDue=HoursTillDue&"h "&MinutesTilldue&"m"
				End if
				If TempToLocation<>ToLocation then
					X=1
					'Response.Write "*********************GOT HERE!<BR>"
					DisplayToLocation=ToLocation
					DisplayDisplayTimeTillDue=DisplayTimeTillDue
					'Response.Write "JobNumber="&JobNumber&"<BR>"
					'Response.Write "DisplayTimeTillDue="&DisplayTimeTillDue&"<BR>"
					'REsponse.Write "***********************<BR>"
					Set Recordset1 = Server.CreateObject("ADODB.Recordset")
					Recordset1.ActiveConnection = DATABASE
					Recordset1.Source = "SELECT count(fh_id) as NumberOfJobs FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id WHERE (Fl_dr_ID='"&VehicleID&"') AND (fh_ship_dt>'"&now()-7&"') AND ((fh_status='ONB'))"
					Recordset1.CursorType = 0
					Recordset1.CursorLocation = 2
					Recordset1.LockType = 1
					Recordset1.Open()
					Recordset1_numRows = 0
					if NOT Recordset1.EOF then
						NumberOfJobs=Recordset1("NumberOfJobs")
					End if
					Recordset1.Close()
					Set Recordset1 = Nothing
					showhr2=showhr2+1	
					If showhr2>1 then
						Response.Write "<tr><td colspan='7'><hr></td></tr>"					
					End if										
					%>
					<form method="post" action="whatever.asp">
					<tr>
						<!--td align="center" colspan="2">&nbsp;</td-->						
						<td align="center" nowrap><font color="<%=FColor2%>"><%=NumberOfJobs%></font></td>
						<td width="5">&nbsp;</td>
						<td align="center" nowrap><font color="<%=FColor2%>"><%=DisplayDisplayTimeTillDue%></font></td>
						<td width="5">&nbsp;</td>
						<td align="center" nowrap><font color="<%=FColor2%>"><%=DisplayToLocation%></font></td>
						<td width="5">&nbsp;</td>
						<td align="center" nowrap><input type="submit" value="details"></td>
					</tr>
					</form>					
					<%
				End if
				TempToLocation=ToLocation
			oRs.Movenext
			Loop
			oRs.Close
			'Response.Write "X="&X&"<BR>"											
			%>

			<tr><td>&nbsp;</td></tr>
			<tr>
				<td align="center" class="purpleseparator" colspan="13"><b>ORDERS TO BE PICKED UP</b></td>
			</tr>


			<tr>
				<!--td colspan="2">&nbsp;</td-->
				<td align="center" nowrap><b>Jobs</b></td>
				<td width="5">&nbsp;</td>
				<td align="center" nowrap><b>Due in</b></td>
				<td width="5">&nbsp;</td>
				<td align="center" nowrap><b>From/To</b></td>
				<td width="5">&nbsp;</td>
				<td align="center" nowrap>
				<%
				if trim(fh_bt_id)<>"26" then
				%>
					<b>Lots</b>
				<%
				End if
				%>
				</td>
			</tr>
			<%
			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.CursorLocation = 3
			oRs.CursorType = 3
			oRs.ActiveConnection = DATABASE	
			SQL = "SELECT Fl_SF_ID, Fh_ID, Fl_ST_ID, fl_st_rta, fh_bt_id, FH_Status, Fh_Priority FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
			SQL = SQL&" WHERE (Fl_dr_ID='"&VehicleID&"') AND (fh_ship_dt>'"&now()-7&"')"
			SQL = SQL&" AND ((fh_status='ACC'))"
			SQL = SQL&" ORDER BY fh_priority, fl_st_rta"
			'Response.write "SQL="&SQL&"<BR>"
			oRs.Open SQL, DATABASE, 1, 3
			If not oRs.EOF then
					CloseTable="y"
					ELSE
					Response.Write "<tr><td colspan='13' align='center'>There are currently no orders in the vehicle.</td></tr><tr><td>&nbsp;</td></tr>"
			End if
			Do while not oRs.eof
				XX=XX+1
				FromLocation = oRs("Fl_SF_ID")
				JobNumber = oRs("Fh_ID")
				ToLocation = oRs("Fl_ST_ID")
				'Response.Write "ToLocation="&ToLocation&"<BR>"
				'Response.Write "JobNumber="&JobNumber&"<BR>"
				JobStatus = oRs("fh_status")
				Priority = oRs("fh_priority")
				If FColor="" and Priority="P1" then
					FColor="purple"
				End if
				If Priority="P0" then
					FColor="red"
				End if
				fh_bt_id=Trim(cStr(oRs("fh_bt_id")))
				DueTime=oRs("fl_st_rta")
				TimeTillDue=DateDiff("n",now(),DueTime)	
				If TimeTillDue<0 then
					DisplayTimeTillDue="LATE"
					Else
					HoursTillDue=Int(TimeTillDue/60)
					MinutesTillDue=TimeTillDue-(HoursTillDue*60)
					DisplayTimeTillDue=HoursTillDue&"h "&MinutesTilldue&"m"
				End if
				If TempToLocation<>ToLocation OR TempFromLocation<>FromLocation then
					XX=1
					'Response.Write "*********************GOT HERE!<BR>"
					DisplayToLocation=ToLocation
					DisplayFromLocation=FromLocation
					DisplayDisplayTimeTillDue=DisplayTimeTillDue
					'Response.Write "JobNumber="&JobNumber&"<BR>"
					'Response.Write "DisplayTimeTillDue="&DisplayTimeTillDue&"<BR>"
					'REsponse.Write "***********************<BR>"
					Set Recordset1 = Server.CreateObject("ADODB.Recordset")
					Recordset1.ActiveConnection = DATABASE
					Recordset1.Source = "SELECT count(fh_id) as NumberOfJobs FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id WHERE (Fl_dr_ID='"&VehicleID&"') AND (fh_ship_dt>'"&now()-7&"') AND (fh_status='ACC') AND (Fl_ST_ID='"&ToLocation&"') AND (Fl_SF_ID='"&FromLocation&"')"
					Recordset1.CursorType = 0
					Recordset1.CursorLocation = 2
					Recordset1.LockType = 1
					Recordset1.Open()
					Recordset1_numRows = 0
					if NOT Recordset1.EOF then
						NumberOfJobs=Recordset1("NumberOfJobs")
					End if
					Recordset1.Close()
					Set Recordset1 = Nothing	
					showhr=showhr+1	
					If showhr>1 then
						Response.Write "<tr><td colspan='7'><hr></td></tr>"					
					End if			
					%>
					<form method="post" action="whatever.asp">
					<tr>
						<!--td align="center" colspan="2">&nbsp;</td-->						
						<td align="center" nowrap valign="top"><font color="<%=FColor%>"><%=NumberOfJobs%></font></td>
						<td width="5">&nbsp;</td>
						<td align="center" nowrap valign="top"><font color="<%=FColor%>"><%=DisplayDisplayTimeTillDue%></font></td>
						<td width="5">&nbsp;</td>
						<td align="center" nowrap valign="top"><font color="<%=FColor%>"><%=DisplayFromLocation%><br><%=DisplayToLocation%></font></td>
						<td width="5">&nbsp;</td>
						<td align="center" nowrap><input type="submit" value="details" ID="Submit1" NAME="Submit1"></td>					
					</tr>
					</form>
									
					<%
				End if
				TempToLocation=ToLocation
				TempFromLocation=FromLocation
			oRs.Movenext
			Loop
			oRs.Close
			'Response.Write "X="&X&"<BR>"											
			%>

			<tr><td>&nbsp;</td></tr>
			<tr><td>&nbsp;</td></tr>
			</table>				
	</BODY>
</HTML>
