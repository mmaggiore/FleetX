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
		AcknowledgeIt=Request.Form("AcknowledgeIt")
		DriverID=Request.Form("DriverID")
		LocationCode=Request.Form("LocationCode")
		Submit=Request.Form("Submit")
		PageStatus=Request.Form("PageStatus")
		PageStatus="loggedin"
		txtJobNumber=Request.Form("txtJobNumber")
		If Submit="submit" then
			'If LocationCode="" then
			'	ErrorMessage="You must provide your location code"
			'End if				
			If DriverID="" then
				ErrorMessage="You must provide your driver id"
			End if
		If ErrorMessage="" then PageStatus="loggedin" End if
		End if
		%>
	</HEAD>
	<BODY leftMargin=0 TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<%
		Select Case PageStatus
			Case "loggedin"
'-------------------STARTS THE DROP OFF				
				X=0
				Y=0
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
				%>
					<table width="300" cellpadding="0" cellspacing="0" border="0" align="left" ID="Table1">
						<tr><td align="center" colspan="9"><a href="default.asp" class="mainpagelink"><form method="post" action="default.asp"><input type="submit" value="Return to Menu"></form></td></tr>
						<tr>
							<td align="center" colspan="13" class="purpleseparator"><b>CURRENT STATUS OF <%=uCase(VehicleName)%></b></td>
						</tr>						
						<tr>
							<td align="center" class="purpleseparator" colspan="13"><b>ORDERS IN VEHICLE</b></td>
						</tr>
				<%
				If not oRs.EOF then
						'm_Logit "OrdersToBeDroppedOff " & DriverID, oConn
						'm_Logit "OrdersToBeDroppedOff " & LocationCode, oConn

						CloseTable="y"
						ELSE
						Response.Write "<tr><td colspan='13' align='center'>There are currently no orders in the vehicle.</td></tr><tr><td>&nbsp;</td></tr>"
				End if
				Do while not oRs.eof
					FromLocation = oRs("Fl_SF_ID")
					JobNumber = oRs("Fh_ID")
					ToLocation = oRs("Fl_ST_ID")
					JobStatus = oRs("fh_status")
					Priority = oRs("fh_priority")
					fh_bt_id=Trim(cStr(oRs("fh_bt_id")))
					DueTime=oRs("fl_st_rta")
					TimeTillDue=DateDiff("n",now(),DueTime)
					'Response.Write "TimeTillDue="&TimeTillDue&"<BR>"
					xxx=xxx+1
					if xxx=1 then
					%>

						<tr>
							<!--td align="center" colspan="2">&nbsp;</td-->						
							<td align="center" nowrap><b>Job #</b></td>
							<td width="5">&nbsp;</td>
							<td align="center" nowrap><b>Due In</b></td>
							<td width="5">&nbsp;</td>
							<td align="center" nowrap><b>To</b></td>
							<td width="5">&nbsp;</td>
							<td align="center" nowrap>
							<%if trim(fh_bt_id)<>"26" then%>
								<b>Lots</b>
							<%end if%>
							</td>
							</tr>
						<%
					End if					
					'rESPONSE.Write "timetilldue="&timetilldue&"<BR>"
					If TimeTillDue<0 then
						DisplayTimeTillDue="LATE"
						Else
						HoursTillDue=Int(TimeTillDue/60)
						'rESPONSE.Write "HoursTillDue="&HoursTillDue&"<BR>"
						MinutesTillDue=TimeTillDue-(HoursTillDue*60)
						'rESPONSE.Write "MinutesTillDue="&MinutesTillDue&"<BR>"
						DisplayTimeTillDue=HoursTillDue&"h "&MinutesTilldue&"m"
					End if					
					if trim(fh_bt_id)<>"26" then
						'Response.Write "Got here????<BR>"
						Set Recordset1 = Server.CreateObject("ADODB.Recordset")
						Recordset1.ActiveConnection = DATABASE
						Recordset1.Source = "SELECT count(rf_ref) as NumberOfLots FROM FCREFS WHERE (rf_fh_id='"&JobNumber&"')"
						Recordset1.CursorType = 0
						Recordset1.CursorLocation = 2
						Recordset1.LockType = 1
						Recordset1.Open()
						Recordset1_numRows = 0
						if NOT Recordset1.EOF then
							NumberOfLots=Recordset1("NumberOfLots")
							If NumberOfLots>1 then WordLots="Lots" end if
							If NumberOfLots=1 then WordLots="Lot" end if
							If NumberOfLots=0 then WordLots="" end if
							Else
							ErrorMessage="Incorrect driver ID or password"
						End if
						Recordset1.Close()
						Set Recordset1 = Nothing					
					End if
					
					
					If Priority="P0" then 
						PriorityColor="red"
						ButtonClass="ButtonRed"
						Else
						ButtonClass="Button1"
						PriorityColor="black"
					End if					
					Select Case JobStatus
						Case "OPN"
							JobStatus="Open"
							ButtonText="ACK"
						Case "ACC"
							JobStatus="ACK"
							ButtonText="ONB"
						Case "ONB"
							JobStatus="ONB"
							ButtonText="CLS"
					End Select
					'FromLocation = oRs("Fl_SF_ID")
					'If JobNumber<>TempJobNumber then
					'Response.Write TempX&"</font></td></tr>"
						If X>0 then
							Response.Write "<tr><td colspan='13' bgcolor='white'><img src='images/pixel.gif' height='2' width='1' border='0'></td></tr>"
							Response.Write "<tr><td colspan='13' bgcolor='gray'><img src='images/pixel.gif' height='1' width='1' border='0'></td></tr>"
							Response.Write "<tr><td colspan='13' bgcolor='white'><img src='images/pixel.gif' height='2' width='1' border='0'></td></tr>"							
							X=0
						End if
						Y=Y+1
						If Priority="P0" then
							ButtonClass="ButtonRed"
							else
							ButtonClass="Button1"
						End if
						Select Case JobStatus
							Case "ACK","ONB"
								%>
								<form method="post" action="getjobdetails.asp" ID="Form3">
								<!--td valign="top"><input type="submit" value="<%=ButtonText%>" name="submit" class="<%=ButtonClass%>" ID="Submit4"></td>
								<td width="20">&nbsp;</td-->
								<input type="hidden" name="txtcaller" value="<%=DriverID%>" ID="Hidden6">
								<input type="hidden" name="txtstation" value="<%=ToLocation%>" ID="Hidden7">
								<input type="hidden" name="txtjobnumber" value="<%=jobnumber%>" ID="Hidden8">
								<input type="hidden" name="DriverID" value="<%=DriverID%>" ID="Hidden28">
								<input type="hidden" name="LocationCode" value="<%=ToLocation%>" ID="Hidden29">
								<input type="hidden" name="jobnumber" value="<%=jobnumber%>" ID="Hidden30">									
								<!--
								<input type="hidden" name="" value="<%=x%>" ID="Hidden3">
								<input type="hidden" name="" value="<%=x%>" ID="Hidden4">
								-->
								</form>
								<%
							Case "Open"
								%>
								<form method="post" action="DriverTruckLoad.asp" ID="Form5">
								<!--td valign="top"><input type="submit" value="<%=ButtonText%>" name="submit" class="<%=ButtonClass%>" ID="Submit6"></td>
								<td width="20">&nbsp;</td-->
								<input type="hidden" name="txtcaller" value="<%=DriverID%>" ID="Hidden18">
								<input type="hidden" name="txtstation" value="<%=ToLocation%>" ID="Hidden19">
								<input type="hidden" name="txtjobnumber" value="<%=jobnumber%>" ID="Hidden20">
								<input type="hidden" name="DriverID" value="<%=DriverID%>" ID="Hidden21">
								<input type="hidden" name="LocationCode" value="<%=LocationCode%>" ID="Hidden22">
								<input type="hidden" name="AcknowledgeIt" value="y" ID="Hidden23">
								<input type="hidden" name="PageStatus" value="loggedin" ID="Hidden24">
								</form>
								<%								
							Case Else
								%>
								<td valign="top">&nbsp;</td>
								<td width="20">&nbsp;</td>								
								<%
						End Select
						'Response.Write "<td>&nbsp;</td>"
						%>
						<form method="post" action="DriverTracking.asp" ID="Form7">
						<%						
						Response.Write "<td valign='top' nowrap><font color='"&PriorityColor&"'>"&Y&")" %>
						
						<input type="hidden" name="JobNumber" value="<%=JobNumber%>">
						<input type="hidden" name="fh_bt_id" value="<%=fh_bt_id%>" ID="Hidden14">
						<input type="submit" name="submit" value="<%=Right(JobNumber,5)%>">
						
						</font></td>
						</form>
						<%
						Response.Write "<td>&nbsp;</td>"
						Response.Write "<td valign='top' nowrap><font color='"&PriorityColor&"'>"&DisplayTimeTillDue&"</font></td>"
						Response.Write "<td>&nbsp;</td>"
						Response.Write "<td valign='top' nowrap><font color='"&PriorityColor&"'>"&ToLocation&"</font></td>"
						Response.Write "<td>&nbsp;</td>"
						Response.Write "<td valign='top' nowrap align='center'><font color='"&PriorityColor&"'>"&NumberOfLots&"</td></tr><tr><td>&nbsp;</td></tr>"
					'End if
					x=x+1
					'Response.Write X&") "&LotNumber&"<br>"
					'TempJobNumber=JobNumber
					'TempX=X
					'Response.Write "----------------------------<br>"
					oRs.Movenext
				Loop
				Response.Write "</font>"
				oRs.Close	
				'If CloseTable="y" then
					'Response.Write TempX
					'Response.Write "</font></td>"					
					'
					'</tr><!--/table-->
					'
					'CloseTable="n"
				'End if
'-------------------STARTS THE PICK UP				
				X=0
				Y=0
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				
				SQL = "SELECT Fl_SF_ID, Fh_ID, fh_bt_id, Fl_ST_ID, fl_st_rta, FH_Status, Fh_Priority FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
				SQL = SQL&" WHERE (Fl_dr_ID='"&VehicleID&"') AND (fh_ship_dt>'"&now()-7&"')"
				'SQL = SQL&" AND ((fh_status='OPN') or (fh_status='ACC'))"
				SQL = SQL&" AND (fh_status='ACC')"
				SQL = SQL&" ORDER BY fh_priority, fl_st_rta"
				oRs.Open SQL, DATABASE, 1, 3
				'REsponse.Write "SQL="&SQL&"<BR>"
				%>
					<!--table width="700" cellpadding="0" cellspacing="0" border="1" align="center" ID="Table2"-->
						<tr><td>&nbsp;</td></tr>
						<tr>
							<td align="center" class="purpleseparator" colspan="13"><b>ORDERS TO BE PICKED UP</b></td>
						</tr>
				<%
				If not oRs.EOF then
						'm_Logit "OrdersToBePickedUp " & DriverID, oConn
						'm_Logit "OrdersToBePickedUp " & LocationCode, oConn

						CloseTable="y"
						ELSE
						Response.Write "<tr><td colspan='13' align='center'>There are currently no orders waiting to be picked up.</td></tr><tr><td>&nbsp;</td></tr>"
				End if
				Do while not oRs.eof
					FromLocation = trim(oRs("Fl_SF_ID"))
					JobNumber = trim(oRs("Fh_ID"))
					ToLocation = trim(oRs("Fl_ST_ID"))
					JobStatus = trim(oRs("fh_status"))
					Priority = trim(oRs("fh_priority"))
					fh_bt_id=Trim(cStr(oRs("fh_bt_id")))
					DueTime=oRs("fl_st_rta")
					TimeTillDue=DateDiff("n",now(),DueTime)	
					'rESPONSE.Write "TimeTillDue="&TimeTillDue&"<BR>"				
					If TimeTillDue<0 then
						DisplayTimeTillDue="LATE"
						Else
						HoursTillDue=Int(TimeTillDue/60)
						'rESPONSE.Write "HoursTillDue="&HoursTillDue&"<BR>"
						MinutesTillDue=TimeTillDue-(HoursTillDue*60)
						'rESPONSE.Write "MinutesTillDue="&MinutesTillDue&"<BR>"
						DisplayTimeTillDue=HoursTillDue&"h "&MinutesTilldue&"m"
					End if	
					yyy=yyy+1
					if yyy=1 then
					%>

						<tr>
							<!--td colspan="2">&nbsp;</td-->
							<td align="center" nowrap><b>Job #</b></td>
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
					End if	
					'Response.Write "fh_bt_id="&fh_bt_id&"*****<BR>"								
					if trim(fh_bt_id)<>"26" then
					
						Set Recordset1 = Server.CreateObject("ADODB.Recordset")
						Recordset1.ActiveConnection = DATABASE
						Recordset1.Source = "SELECT count(rf_ref) as NumberOfLots FROM FCREFS WHERE (rf_fh_id='"&JobNumber&"')"
						Recordset1.CursorType = 0
						Recordset1.CursorLocation = 2
						Recordset1.LockType = 1
						Recordset1.Open()
						Recordset1_numRows = 0
						if NOT Recordset1.EOF then
							NumberOfLots=Recordset1("NumberOfLots")
							If NumberOfLots>1 then WordLots="Lots" end if
							If NumberOfLots=1 then WordLots="Lot" end if
							If NumberOfLots=0 then WordLots="" end if
							Else
							ErrorMessage="Incorrect driver ID or password"
						End if
						Recordset1.Close()
						Set Recordset1 = Nothing					
					End if
					
					
					If Priority="P0" then 
						PriorityColor="red"
						ButtonClass="ButtonRed"
						Else
						ButtonClass="Button1"
						PriorityColor="black"
					End if					
					Select Case JobStatus
						Case "OPN"
							JobStatus="Open"
							ButtonText="ACK"
						Case "ACC"
							JobStatus="ACK"
							ButtonText="ONB"
						Case "ONB"
							JobStatus="ONB"
							ButtonText="CLS"
					End Select
					'FromLocation = oRs("Fl_SF_ID")
					'If JobNumber<>TempJobNumber then
						'If TempJobNumber>"" then
						'	Response.Write TempX&"</font></td></tr>"
						'End if
						If X>0 then
							Response.Write "<tr><td colspan='13' bgcolor='white'><img src='images/pixel.gif' height='2' width='1' border='0'></td></tr>"
							Response.Write "<tr><td colspan='13' bgcolor='gray'><img src='images/pixel.gif' height='1' width='1' border='0'></td></tr>"
							Response.Write "<tr><td colspan='13' bgcolor='white'><img src='images/pixel.gif' height='2' width='1' border='0'></td></tr>"
							X=0
						End if
						Y=Y+1
						If Priority="P0" then
							ButtonClass="ButtonRed"
							else
							ButtonClass="Button1"
						End if
						Select Case JobStatus
							Case "ACK","ONB"
								%>
								<form method="post" action="getjobdetails.asp" ID="Form2">
								<!--td valign="top"><input type="submit" value="<%=ButtonText%>" name="submit" class="<%=ButtonClass%>" ID="Submit2"></td>
								<td width="20">&nbsp;</td-->
								<input type="hidden" name="txtcaller" value="<%=DriverID%>" ID="Hidden3">
								<input type="hidden" name="txtstation" value="<%=FromLocation%>" ID="Hidden4">
								<input type="hidden" name="txtjobnumber" value="<%=jobnumber%>" ID="Hidden5">
								<input type="hidden" name="DriverID" value="<%=DriverID%>" ID="Hidden25">
								<input type="hidden" name="LocationCode" value="<%=FromLocation%>" ID="Hidden26">
								<input type="hidden" name="jobnumber" value="<%=jobnumber%>" ID="Hidden27">								
								<!--
								<input type="hidden" name="" value="<%=x%>" ID="Hidden3">
								<input type="hidden" name="" value="<%=x%>" ID="Hidden4">
								-->
								</form>
								<%
							Case "Open"
								%>
								<form method="post" action="DriverTruckLoad.asp" ID="Form4">
								<!--td valign="top"><input type="submit" value="<%=ButtonText%>" name="submit" class="<%=ButtonClass%>" ID="Submit5"></td>
								<td width="20">&nbsp;</td-->
								<input type="hidden" name="txtcaller" value="<%=DriverID%>" ID="Hidden9">
								<input type="hidden" name="txtstation" value="<%=ToLocation%>" ID="Hidden10">
								<input type="hidden" name="txtjobnumber" value="<%=jobnumber%>" ID="Hidden11">
								<input type="hidden" name="DriverID" value="<%=DriverID%>" ID="Hidden12">
								<input type="hidden" name="LocationCode" value="<%=LocationCode%>" ID="Hidden13">
								<input type="hidden" name="AcknowledgeIt" value="y" ID="Hidden1">
								<input type="hidden" name="PageStatus" value="loggedin" ID="Hidden2">
								</form>
								<%								
							Case Else
								%>
								<td valign="top">&nbsp;</td>
								<td width="20">&nbsp;</td>								
								<%
						End Select
						'Response.Write "<td>&nbsp;</td>"
						%>
						<form method="post" action="DriverTracking.asp" ID="Form1">
						<%						
						Response.Write "<td valign='top' nowrap><font color='"&PriorityColor&"'>"&Y&")"%>
						
						<input type="hidden" name="JobNumber" value="<%=JobNumber%>" ID="Hidden15">
						<input type="hidden" name="fh_bt_id" value="<%=fh_bt_id%>" ID="Hidden16">
						<input type="submit" name="submit" value="<%=Right(JobNumber,5)%>" ID="Submit2">
											 
						 </font></td>
						 </form>	
						 <%						
						Response.Write "<td>&nbsp;</td>"
						Response.Write "<td valign='top' nowrap><font color='"&PriorityColor&"'>"&DisplayTimeTillDue&"</font></td>"
						Response.Write "<td>&nbsp;</td>"
						Response.Write "<td valign='top' nowrap><font color='"&PriorityColor&"'>"&FromLocation&"<br>"&ToLocation&"</font></td>"
						Response.Write "<td>&nbsp;</td>"
						Response.Write "<td valign='top' nowrap align='center'><font color='"&PriorityColor&"'>"&NumberOfLots&"</td></tr>"
						
					'End if
					x=x+1
					'Response.Write X&") "&LotNumber&"<br>"
					TempJobNumber=JobNumber
					TempX=X
					'Response.Write "----------------------------<br>"
					oRs.Movenext
				Loop
				Response.Write "</font>"
				oRs.Close	
				'If CloseTable="y" then
					'Response.Write TempX
					'Response.Write "</font></td>"						
				'	
				'	</tr><!--/table-->
				'	
				'	CloseTable="n"
				'End if			
			Case else
			%>
			<FORM ACTION="DriverTruckLoad.asp" method="post" name="thisForm" ID="Form6">
				<TABLE WIDTH="300" align="left" cellpadding="0" cellspacing="5" ID="Table3">
					<TR> 
						<td width="50%"> 
							<div class="purpleseparator"> 
								<table border="0" cellpadding="2" cellspacing="0" ID="Table4">
									<%if errormessage>"" then%>
										<tr>
											<td class='generalcontenthighlight'colspan='2' align="center"><font color="red"><br><b><%=ErrorMessage%></b><br><br></font></td>
										</tr>
									<%end if%>									
									<tr>
										<td class='generalcontenthighlight' width='25%'></td>
										<td width='75%' class='generalcontent'>
											<input type="submit" name="submit" value="submit" ID="Submit1">									
										</td>
									</tr>
									
									<tr> 
										<td class="subheader" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
									</tr>
								</table>
							</div>
						</td>
						<!--Dummy section-->
						<td align=left width="50%" height="100%"> 
						</TD>
					</TR>
					

				</TABLE>
			</FORM>
			
			<%
			End Select
			%>
			<tr><td>&nbsp;</td></tr>
			</table>				
	</BODY>
</HTML>