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
		Dim ListOfFrom(200)
		Dim ListOfToM(200)
		Dim ListOfTo(200)
		DriverID=Request.Form("DriverID")
		BillToID=Request.Cookies("Phone")("sBT_ID")	
		mark=Request.QueryString("Mark")
		JobNumber=Request.Form("JobNumber")
		BOL=Request.Form("BOL")
		If BOL>"" then
			Set oConn = Server.CreateObject("ADODB.Connection")
			oConn.ConnectionTimeout = 100
			oConn.Provider = "MSDASQL"
			oConn.Open DATABASE
				l_cSQL = "UPDATE fcrefs SET rf_box = '"& BOL &"' WHERE rf_fh_id = '" & JobNumber & "'"
				''''response.Write "l_cSQL="&l_cSQL&"<br>"
				'oConn.Execute "PHONE_CHANGE_STATUS '" & txtJobNumber & "', '4', 'ACC', '','','',''" 
				oConn.Execute(l_cSQL)
			Set oConn=Nothing		
		End if
		%>
	</HEAD>
	<BODY leftMargin=0 TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<table width="300" cellpadding="0" cellspacing="0" border="0" align="left" ID="Table2">
			<tr><td align="center" colspan="9"><a href="default.asp" class="mainpagelink"><form method="post" action="default.asp" ID="Form8"><input type="submit" value="Return to Menu" ID="Submit3" NAME="Submit3"></form></td></tr>
			<tr>
				<td align="center" colspan="13" class="purpleseparator"><b>CURRENT STATUS OF <%=uCase(VehicleName)%></b></td>
			</tr>						
			<tr>
				<td align="center" class="purpleseparator" colspan="13"><b>ORDERS REQUIRING BOL #</b></td>
			</tr>


			<tr>
				<!--td colspan="2">&nbsp;</td-->
				<td align="center" nowrap><b>Jobs</b></td>
				<td width="5">&nbsp;</td>
				<td align="center" nowrap><b>Due in</b></td>
				<td width="5">&nbsp;</td>
				<td align="center" nowrap><b>From/To</b></td>
				<td width="5">&nbsp;</td>
			</tr>
			<%
			Showhr=0
			DontShow=""
			Showdetails=""
			YYY=0
			Z=0
			XX=0
			TempToLocation=""
			TempFromLocation=""			
			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.CursorLocation = 3
			oRs.CursorType = 3
			oRs.ActiveConnection = DATABASE	
			SQL = "SELECT Distinct(Fl_SF_ID), Fh_ID, Fl_ST_ID, fl_st_rta, fl_firstdrop, fh_bt_id, FH_Status, Fh_Priority, fh_user5 FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN fcrefs ON fclegs.fl_fh_id = fcrefs.rf_fh_id"
			SQL = SQL&" WHERE (Fl_dr_ID='"&trim(VehicleID)&"') AND (fh_ship_dt>'"&now()-30&"')"
			SQL = SQL&" AND ((fh_status='ACC')"
			SQL = SQL&" AND (rf_box='')"
			SQL = SQL&" )"

			
			
			SQL = SQL&" ORDER BY fl_st_rta, fh_priority, fl_sf_id"
			If mark="y" then
				response.write "to be picked up SQL="&SQL&"<BR>"
			end if
			'response.write "to be picked up SQL="&SQL&"<BR>"
			oRs.Open SQL, DATABASE, 1, 3
			If not oRs.EOF then
					CloseTable="y"
					ELSE
					If WereP0s<>"y" then
						Response.Write "<tr><td colspan='13' align='center'>There are currently no orders without BOLs.</td></tr><tr><td>&nbsp;</td></tr>"
						
					End if
			End if
			Do while not oRs.eof
				XX=XX+1
				FromLocation = oRs("Fl_SF_ID")
				JobNumber = oRs("Fh_ID")
				ToLocation = oRs("Fl_ST_ID")
				fl_firstdrop = oRs("Fl_firstdrop")
				'Response.Write "ToLocation="&ToLocation&"<BR>"
				'Response.Write "JobNumber="&JobNumber&"<BR>"
				JobStatus = oRs("fh_status")
				Priority = oRs("fh_priority")
				If FColor="" and Priority="P1" then
					FColor="purple"
					else
					If Priority="P0" then
						FColor="red"
						else 
						FColor="black"
					End if
				End if
				fh_bt_id=Trim(cStr(oRs("fh_bt_id")))
				MaterialType = oRs("fh_user5")
				DueTime=oRs("fl_st_rta")
				If trim(FromLocation)="55" or trim(FromLocation)="72" then
					'Response.Write "Got here???<BR>"
					If Priority="P0" then
						DueTime=DateAdd("n", 45, Fl_firstdrop)
						else
						DueTime=DateAdd("n", 90, Fl_firstdrop)
					End if					
				End if				
				TimeTillDue=DateDiff("n",now(),DueTime)	
				If TimeTillDue<0 then
					DisplayTimeTillDue="LATE"
					Else
					HoursTillDue=Int(TimeTillDue/60)
					MinutesTillDue=TimeTillDue-(HoursTillDue*60)
					DisplayTimeTillDue=HoursTillDue&"h "&MinutesTilldue&"m"
				End if
					DisplayToLocation=ToLocation
					DisplayFromLocation=FromLocation
					DisplayDisplayTimeTillDue=DisplayTimeTillDue
					DisplayMaterialSymbol=MaterialSymbol

				
					showhr=showhr+1	
					If showhr>1 OR (showhr=1 AND WereP0s="y") then
						Response.Write "<tr><td colspan='7'><hr></td></tr>"					
					End if
						%>
						<form method="post" action="DriverBOL.asp">
						<tr>
							<!--td align="center" colspan="2">&nbsp;</td-->						
							<td align="center" nowrap valign="top"><font color="<%=FColor%>"><%=MaterialSymbol%><a href="DriverTracking.asp?JobNumber=<%=JobNumber%>&fh_bt_id=75"><%=JobNumber%></a><%=MaterialSymbol%></font></td>
							<td width="5">&nbsp;</td>
							<td align="center" nowrap valign="top"><font color="<%=FColor%>"><%=DisplayDisplayTimeTillDue%></font></td>
							<td width="5">&nbsp;</td>
							<td align="center" nowrap valign="top"><font color="<%=FColor%>"><%=DisplayFromLocation%>/<%=DisplayToLocation%></font></td>
						</tr>
						<tr>
							<td colspan="4"><input type="text" name="BOL"></td>
							<td nowrap valign="top">
							<input type="hidden" name="JobNumber" value="<%=JobNumber%>">
							<input type="submit" value="submit BOL" ID="Submit1" NAME="Submit1">
							</td>					
						</tr>
						</form>
										
						<%
						MaterialSymbol=""
					DontShow="n"
			oRs.Movenext
			Loop
			oRs.Close
			WereP0s=""
			%>

			<tr><td>&nbsp;</td></tr>
			<tr><td>&nbsp;</td></tr>
			</table>				
	</BODY>
</HTML>
