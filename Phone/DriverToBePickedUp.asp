<%@ Language=VBScript %>
<!-- 10/08/2005 Span-Renuka included this screen for CCF 2023 -->
<!-- #include file="../DedicatedFleets/include/checkstring.inc" -->
<!-- #include file="../DedicatedFleets/include/custom.inc" -->
<!-- #include file="../DedicatedFleets/include/ifabsettings.inc" -->
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
		
		
	
		'FUNCTION m_Logit (strDetails, oConnection)
		'	'	---- Log Errors ----
'
'
		'		Set oConnection = Server.CreateObject("ADODB.Connection")
		'		oConnection.ConnectionTimeout = 100
		'		oConnection.Provider = "MSDASQL"
		'		oConnection.Open DATABASE	
		'		l_cSQL ="INSERT INTO applog (ap_date, ap_log, ap_user, ap_bt_id) " & _
		'			"VALUES ('" & DATE() & " " & TIME() & "', " & _
		'			"'" & strDetails & "','DriverTruckLoad.asp','')"
		'			'response.Write "l_cSQL="&l_cSQL&"<BR>"
		'		Set oRs666 = oConnection.Execute(l_cSQL)
		'		Set oConnection=Nothing
		'	
		'END FUNCTION		
		
		%>
	</HEAD>
	<BODY leftMargin=0 TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<%
		Select Case PageStatus
			Case "loggedin"
				If AcknowledgeIt="y" then
					'Response.write "GOT HERE!"
					Set oConn = Server.CreateObject("ADODB.Connection")
					oConn.ConnectionTimeout = 100
					oConn.Provider = "MSDASQL"
					oConn.Open DATABASE
					' 7/26/04 KK: Added canceljob functionality to be able to update the status if cancel button is pressed.
					' 11/30/04 DEC: Changed from CAN/98 TO DEL/99 to be consistent with dispatchOffice			
						oConn.Execute "PHONE_CHANGE_STATUS '" & txtJobNumber & "', '4', 'ACC'" 
						'''''l_cSQL = "UPDATE fcfgthd SET fh_status = 'ACC', fh_statcode = 4 WHERE fh_id = '" & txtJobNumber & "'"
						'Response.write "l_cSQL="&l_cSQL&"<BR>"
						
						'''''oConn.Execute(l_cSQL)
						'm_logit "AcknowledgesOnFCFGTHD " & txtJobNumber, oConn
					Set oConn=Nothing
					'''''Set oConn = Server.CreateObject("ADODB.Connection")
					'''''oConn.ConnectionTimeout = 100
					'''''oConn.Provider = "MSDASQL"
					'''''oConn.Open DATABASE
					' 7/26/04 KK: Added canceljob functionality to be able to update the status if cancel button is pressed.
					' 11/30/04 DEC: Changed from CAN/98 TO DEL/99 to be consistent with dispatchOffice			
						'''''l_cSQL = "UPDATE fclegs SET fl_t_acc='"&now()&"' WHERE fl_fh_id = '" & txtJobNumber & "'"
						'Response.write "l_cSQL="&l_cSQL&"<BR>"
						
						'''''oConn.Execute(l_cSQL)
						'm_logit "AcknowledgesOnFCLEGS " & txtJobNumber, oConn
					'''''Set oConn=Nothing					
				End if
'-------------------STARTS THE DROP OFF				
				X=0
				Y=0
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				
				SQL = "SELECT Fl_SF_ID, Fh_ID, Fl_ST_ID, FH_Status, Fh_Priority, RF_Ref FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN fcrefs ON fclegs.fl_fh_id = fcrefs.rf_fh_id"
				SQL = SQL&" WHERE (Fl_dr_ID='"&VehicleID&"') AND (fh_ship_dt>'"&now()-30&"')"
				SQL = SQL&" AND ((fh_status='ONB'))"
				SQL = SQL&" ORDER BY fh_priority, fh_id"
				'Response.write "SQL="&SQL&"<BR>"
				oRs.Open SQL, DATABASE, 1, 3
				%>
					<table width="300" cellpadding="0" cellspacing="0" border="0" align="left" ID="Table1">
						<tr>
							<td align="center" colspan="13"><b>CURRENT STATUS OF <%=uCase(VehicleName)%></b></td>
						</tr>						
						<tr>
							<td align="center" class="purpleseparator" colspan="13"><b>ORDERS IN VEHICLE</b></td>
						</tr>
				<%
				If not oRs.EOF then
						'm_logit "OrdersToBeDroppedOff " & DriverID, oConn
						'm_logit "OrdersToBeDroppedOff " & LocationCode, oConn
					%>

						<tr>
							<!--td align="center" colspan="2">&nbsp;</td-->						
							<td align="center" nowrap><b>Job #</b></td>
							<td width="20">&nbsp;</td>
							<td align="center" nowrap><b>&nbsp;</b></td>
							<td width="20">&nbsp;</td>
							<td align="center" nowrap><b>To</b></td>
							<td width="20">&nbsp;</td>
							<td align="center" nowrap><b>Lots</b></td>
							</tr>
						<%
						CloseTable="y"
						ELSE
						Response.Write "<tr><td colspan='13' align='center'>There are currently no orders to drop off at this location.</td></tr><tr><td>&nbsp;</td></tr>"
				End if
				Do while not oRs.eof
					FromLocation = oRs("Fl_SF_ID")
					JobNumber = oRs("Fh_ID")
					ToLocation = oRs("Fl_ST_ID")
					JobStatus = oRs("fh_status")
					Priority = oRs("fh_priority")
					If Priority="P0" then 
						PriorityColor="red"
						ButtonClass="ButtonRed"
						Else
						ButtonClass="Button1"
						PriorityColor="black"
					End if					
					LotNumber = oRs("rf_ref")
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
					If JobNumber<>TempJobNumber then
					Response.Write TempX&"</font></td></tr>"
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
						Response.Write "<td valign='top' nowrap><font color='"&PriorityColor&"'>"&Y&") "&JobNumber&"</font></td>"
						Response.Write "<td>&nbsp;</td>"
						Response.Write "<td valign='top' nowrap><font color='"&PriorityColor&"'>"&Priority&"</font></td>"
						Response.Write "<td>&nbsp;</td>"
						Response.Write "<td valign='top' nowrap><font color='"&PriorityColor&"'>"&ToLocation&"</font></td>"
						Response.Write "<td>&nbsp;</td>"
						Response.Write "<td valign='top' nowrap align='center'><font color='"&PriorityColor&"'>"
						
					End if
					x=x+1
					'Response.Write X&") "&LotNumber&"<br>"
					TempJobNumber=JobNumber
					TempX=X
					'Response.Write "----------------------------<br>"
					oRs.Movenext
				Loop
				Response.Write "</font>"
				oRs.Close	
				If CloseTable="y" then
					Response.Write TempX
					Response.Write "</font></td>"					
					%>
					</tr><!--/table-->
					<%
					CloseTable="n"
				End if
'-------------------STARTS THE PICK UP				
				X=0
				Y=0
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				
				SQL = "SELECT Fl_SF_ID, Fh_ID, Fl_ST_ID, FH_Status, Fh_Priority, RF_Ref FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN fcrefs ON fclegs.fl_fh_id = fcrefs.rf_fh_id"
				SQL = SQL&" WHERE (Fl_dr_ID='"&VehicleID&"') AND (fh_ship_dt>'"&now()-30&"')"
				'SQL = SQL&" AND ((fh_status='OPN') or (fh_status='ACC'))"
				SQL = SQL&" AND (fh_status='ACC')"
				SQL = SQL&" ORDER BY fh_priority, fh_id"
				oRs.Open SQL, DATABASE, 1, 3
				%>
					<!--table width="700" cellpadding="0" cellspacing="0" border="1" align="center" ID="Table2"-->
						<tr><td>&nbsp;</td></tr>
						<tr>
							<td align="center" class="purpleseparator" colspan="13"><b>ORDERS TO BE PICKED UP</b></td>
						</tr>
				<%
				If not oRs.EOF then
						'm_logit "OrdersToBePickedUp " & DriverID, oConn
						'm_logit "OrdersToBePickedUp " & LocationCode, oConn
					%>

						<tr>
							<!--td colspan="2">&nbsp;</td-->
							<td align="center" nowrap><b>Job #</b></td>
							<td width="20">&nbsp;</td>
							<td align="center" nowrap><b>&nbsp;</b></td>
							<td width="20">&nbsp;</td>
							<td align="center" nowrap><b>From</b></td>
							<td width="20">&nbsp;</td>
							<td align="center" nowrap><b>Lots</b></td>
							</tr>
						<%
						CloseTable="y"
						ELSE
						Response.Write "<tr><td colspan='13' align='center'>There are currently no orders to pick up at this location.</td></tr><tr><td>&nbsp;</td></tr>"
				End if
				Do while not oRs.eof
					FromLocation = trim(oRs("Fl_SF_ID"))
					JobNumber = trim(oRs("Fh_ID"))
					ToLocation = trim(oRs("Fl_ST_ID"))
					JobStatus = trim(oRs("fh_status"))
					Priority = trim(oRs("fh_priority"))
					If Priority="P0" then 
						PriorityColor="red"
						ButtonClass="ButtonRed"
						Else
						ButtonClass="Button1"
						PriorityColor="black"
					End if					
					LotNumber = oRs("rf_ref")
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
					If JobNumber<>TempJobNumber then
						Response.Write TempX&"</font></td></tr>"
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
						Response.Write "<td valign='top' nowrap><font color='"&PriorityColor&"'>"&Y&") "&JobNumber&"</font></td>"
						Response.Write "<td>&nbsp;</td>"
						Response.Write "<td valign='top' nowrap><font color='"&PriorityColor&"'>"&Priority&"</font></td>"
						Response.Write "<td>&nbsp;</td>"
						Response.Write "<td valign='top' nowrap><font color='"&PriorityColor&"'>"&FromLocation&"</font></td>"
						Response.Write "<td>&nbsp;</td>"
						Response.Write "<td valign='top' nowrap align='center'><font color='"&PriorityColor&"'>"
						
					End if
					x=x+1
					'Response.Write X&") "&LotNumber&"<br>"
					TempJobNumber=JobNumber
					TempX=X
					'Response.Write "----------------------------<br>"
					oRs.Movenext
				Loop
				Response.Write "</font>"
				oRs.Close	
				If CloseTable="y" then
					Response.Write TempX
					Response.Write "</font></td>"						
					%>
					</tr><!--/table-->
					<%
					CloseTable="n"
				End if			
			Case else
			'LocationAlias=Request.Cookies("Location_Logisticorp")("LocationAlias")
			'If LocationAlias="" then
			'	Response.Redirect("ifabpmonepage.asp")
			'End if 
			Set oConn = Server.CreateObject("ADODB.Connection")
			oConn.ConnectionTimeout = 100
			oConn.Provider = "MSDASQL"
			oConn.Open DATABASE
				l_cSQL = "select st_id, st_addr1 from fcshipto  " &_
						 "WHERE st_alias = '" & TRIM(LocationAlias)&"'" 
				SET oRs = oConn.Execute(l_cSql)

				'Response.Write "l_cSQL="&l_cSQL&"<BR>"
				'Response.Write "st_id="&st_id&"<BR>"
				'Response.Write "st_addr1="&st_addr1&"<BR>"
				IF not oRs.EOF then	
						XYZ=XYZ+1
						st_addr1=oRs("st_addr1")
						LocationCode=oRs("st_id")
						'm_logit "SETCOOKIE " & LocationAlias, oConn
				End if
			Set oConn=Nothing				
			%>
			<FORM ACTION="DriverTruckLoad.asp" method="post" name="thisForm" ID="Form6">
				<TABLE WIDTH="300" align="left" cellpadding="0" cellspacing="5" ID="Table3">
					<TR> 
						<td width="50%"> 
							<div class="purpleseparator"> 
								<table border="0" cellpadding="2" cellspacing="0" ID="Table4">
									<tr> 
										<td class="subheader" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
									</tr>
									<tr> 
										<td class="subheader" colspan="2">Log in to phone</td>
									</tr>
									<tr>									
										<td class='generalcontenthighlight' width='25%'>Driver ID :</td>
										<td width='75%' class='generalcontent'>
											<input maxlength='25' size='25' name='DriverID' id='DriverID' class='inputgeneral' value='<%=DriverID%>'>
										</td>
									</tr>
									<tr>
										<td class='generalcontenthighlight' width='25%'>Station Location:</td>
										<td width='75%' class='generalcontent'>
											<%
											if XYZ>1 then%>
												<%=LocationCode%>
												<input type="hidden" name="LocationCode" value="<%=LocationCode%>" ID="Hidden34">
												<%else%>
											
												<select name="LocationCode" ID="Select1">
												<%
													Set oConn = Server.CreateObject("ADODB.Connection")
													oConn.ConnectionTimeout = 100
													oConn.Provider = "MSDASQL"
													oConn.Open DATABASE
														l_cSQL = "select st_id, st_addr1 from fcshipto  " &_
																"WHERE st_alias = '" & TRIM(LocationAlias)&"'" 
														SET oRs = oConn.Execute(l_cSql)

																'Response.Write "l_cSQL="&l_cSQL&"<BR>"
																'Response.Write "st_id="&st_id&"<BR>"
																'Response.Write "st_addr1="&st_addr1&"<BR>"
																Do While not oRs.EOF
																st_addr1=oRs("st_addr1")
																st_id=oRs("st_id")								
															%>
															<option value="<%=st_id%>" <%if lstPickup=st_id then response.Write " selected" end if%>><%=st_id%></option>
															<%
														oRs.movenext
														LOOP
													Set oConn=Nothing									
													%>
												</select>											
											<%	
											end if
											%>
											<!--input maxlength='20' name='LocationCode' id='txtstation' class='input' size='25' value='<%=LocationCode%>'-->
										</td>
									</tr>
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
			<tr><td align="center" colspan="9"><a href="default.asp" class="mainpagelink">Return Home 66</a></td></tr><tr><td>&nbsp;</td></tr>
			</table>				
	</BODY>
</HTML>
