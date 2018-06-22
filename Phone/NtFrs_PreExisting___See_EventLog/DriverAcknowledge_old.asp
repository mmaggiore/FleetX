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
		
		
	
		FUNCTION m_Logit (strDetails, oConnection)
			'	---- Log Errors ----


				Set oConnection = Server.CreateObject("ADODB.Connection")
				oConnection.ConnectionTimeout = 100
				oConnection.Provider = "MSDASQL"
				oConnection.Open DATABASE	
				l_cSQL ="INSERT INTO applog (ap_date, ap_log, ap_user, ap_bt_id) " & _
					"VALUES ('" & DATE() & " " & TIME() & "', " & _
					"'" & strDetails & "','DriverifabPhoneEm','')"
					'response.Write "l_cSQL="&l_cSQL&"<BR>"
				Set oRs666 = oConnection.Execute(l_cSQL)
				Set oConnection=Nothing
			
		END FUNCTION		
		
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
						l_cSQL = "UPDATE fcfgthd SET fh_status = 'ACC', fh_statcode = 4 WHERE fh_id = '" & txtJobNumber & "'"
						'Response.write "l_cSQL="&l_cSQL&"<BR>"
						
						oConn.Execute(l_cSQL)
						m_logit "AcknowledgesOnFCFGTHD " & txtJobNumber, oConn
					Set oConn=Nothing
					Set oConn = Server.CreateObject("ADODB.Connection")
					oConn.ConnectionTimeout = 100
					oConn.Provider = "MSDASQL"
					oConn.Open DATABASE
					' 7/26/04 KK: Added canceljob functionality to be able to update the status if cancel button is pressed.
					' 11/30/04 DEC: Changed from CAN/98 TO DEL/99 to be consistent with dispatchOffice			
						l_cSQL = "UPDATE fclegs SET fl_t_acc='"&now()&"' WHERE fl_fh_id = '" & txtJobNumber & "'"
						'Response.write "l_cSQL="&l_cSQL&"<BR>"
						
						oConn.Execute(l_cSQL)
						m_logit "AcknowledgesOnFCLEGS " & txtJobNumber, oConn
					Set oConn=Nothing					
				End if
'-------------------STARTS THE OTHER ORDERS IN THE PHONE		
				X=0
				Y=0
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				
				SQL = "SELECT Fl_SF_ID, Fh_ID, fh_ship_dt, Fl_ST_ID, FH_Status, Fh_Priority, RF_Ref FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN fcrefs ON fclegs.fl_fh_id = fcrefs.rf_fh_id"
				SQL = SQL&" WHERE (Fl_dr_ID='"&VehicleID&"')"
				SQL = SQL&" AND (Fh_Status='OPN')"
				SQL = SQL&" ORDER BY Fh_Priority, fh_status desc, fh_id"
				'Response.write "<br><font color='blue'>SQL="&SQL&"<BR></font>"
				oRs.Open SQL, DATABASE, 1, 3
				%>
					<table cellpadding="0" width="300" cellspacing="0" bordercolor="red" border="0" align="left" ID="Table5">
						<tr><td>&nbsp;</td></tr>
						<tr>
							<td align="center" class="purpleseparator" colspan="3"><b>NEW ORDERS</b></td>
						</tr>
						<%
						If not oRs.EOF then
							m_logit "OtherOrdersInPhone " & DriverID, oConn
							m_logit "OtherOrdersInPhone " & LocationCode, oConn
							'Response.Write "GOT HERE<BR>"
							CloseTable="y"
							ELSE
							Response.Write "<tr><td colspan='3' align='center'>There are currently no unacknowledged orders.</td></tr><tr><td>&nbsp;</td></tr><tr><td align='center'><a href='default.asp' class='mainpagelink'>Return Home</a></td></tr>"
						End if
						Do while not oRs.eof
						FromLocation = oRs("Fl_SF_ID")
						JobNumber = oRs("Fh_ID")
						ToLocation = oRs("Fl_ST_ID")
						JobStatus = oRs("fh_status")
						Priority = oRs("fh_priority")
						ShipTime = oRs("fh_ship_dt")
						TimeSincePlaced=DateDiff("n",shiptime,now())
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
								ButtonText="Ack"
							Case "ACC"
								JobStatus="Acknowledged"
								ButtonText="On Board"
							Case "ONB"
								JobStatus="On Board"
								ButtonText="Close"
						End Select
						'FromLocation = oRs("Fl_SF_ID")
						If JobNumber<>TempJobNumber then
							If X>0 then
								Response.Write "<font color='"&PreviousPriorityColor&"'>"&X&"</font><br>"
								Response.Write "<tr><td colspan='3' bgcolor='white'><img src='images/pixel.gif' height='2' width='1' border='0'></td></tr>"
								Response.Write "<tr><td colspan='3' bgcolor='gray'><img src='images/pixel.gif' height='1' width='1' border='0'></td></tr>"
								Response.Write "<tr><td colspan='3' bgcolor='white'><img src='images/pixel.gif' height='2' width='1' border='0'></td></tr>"							
								PreviousPriorityColor=PriorityColor
								X=0
							End if
							Y=Y+1
									%>
									<form method="post" action="DriverAcknowledge.asp" ID="Form1">
									<tr><td valign="top" width="40"><input type="submit" value="<%=ButtonText%>" name="submit" class="<%=ButtonClass%>" ID="Submit3"></td>
									<input type="hidden" name="txtcaller" value="<%=DriverID%>" ID="Hidden14">
									<input type="hidden" name="txtstation" value="<%=ToLocation%>" ID="Hidden15">
									<input type="hidden" name="txtjobnumber" value="<%=jobnumber%>" ID="Hidden16">
									<input type="hidden" name="DriverID" value="<%=DriverID%>" ID="Hidden17">
									<input type="hidden" name="LocationCode" value="<%=LocationCode%>" ID="Hidden31">
									<input type="hidden" name="AcknowledgeIt" value="y" ID="Hidden32">
									<input type="hidden" name="PageStatus" value="loggedin" ID="Hidden33">								
									<!--
									<input type="hidden" name="" value="<%=x%>" ID="Hidden3">
									<input type="hidden" name="" value="<%=x%>" ID="Hidden4">
									-->
									
									<%
							'Response.Write "<td>&nbsp;</td>"						
							Response.Write "<td class='mainpagetextboldright'>Job #:&nbsp;&nbsp;</td><td valign='top' nowrap><font color='"&PriorityColor&"'>"&JobNumber&"</font></td></tr>"
							Response.Write "<tr><td>&nbsp;</td><td class='mainpagetextboldright'>Ordered:&nbsp;&nbsp;</td><td valign='top' nowrap><font color='"&PriorityColor&"'>"&TimeSincePlaced&" minutes ago</font></td></tr>"
							Response.Write "<tr><td>&nbsp;</td><td class='mainpagetextboldright'>Priority:&nbsp;&nbsp;</td><td valign='top' nowrap><font color='"&PriorityColor&"'>"&Priority&"</font></td></tr>"
							Response.Write "<tr><td>&nbsp;</td><td class='mainpagetextboldright'>Lane:&nbsp;&nbsp;</td><td valign='top' nowrap><font color='"&PriorityColor&"'>"&FromLocation&" to "&ToLocation&"</font></td></tr>"
							Response.Write "</form>"
							Response.Write "<tr><td>&nbsp;</td><td valign='top' class='mainpagetextboldright'>Lots:&nbsp;&nbsp;</td><td valign='top' nowrap>"
						End if
						x=x+1
						'Response.Write "<font color='"&PriorityColor&"'>"&X&") "&LotNumber&"</font><br>"
						PreviousPriorityColor=PriorityColor
						TempJobNumber=JobNumber
						oRs.Movenext
						Loop
						
						oRs.Close
						Response.Write X	
						'Response.Write "</form>GOT HERE!!!!<BR>"
						If CloseTable="y" then
							%>
							</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td align="center" colspan="3"><a href="default.asp" class="mainpagelink">Return Home</a></td></tr><tr><td>&nbsp;</td></tr></table>					
							<%
							CloseTable="n"
						End if					
					Case else
						LocationAlias=Request.Cookies("Location_Logisticorp")("LocationAlias")
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
							m_logit "SETCOOKIE " & LocationAlias, oConn
						End if
						Set oConn=Nothing				
						%>
						
						
						
						
						
						<FORM ACTION="DriverAcknowledge.asp" method="post" name="thisForm" ID="Form6">
						<TABLE WIDTH="100%" cellpadding="0" cellspacing="5" ID="Table3">
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
	</BODY>
</HTML>
