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
	</HEAD>
	<BODY leftMargin=0 TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<%
'-------------------STARTS THE PICK UP
				
				'Response.Write "VehicleID="&VehicleID&"<BR>"
				'Response.Write "VehicleName="&VehicleName&"<BR>"
				'Response.Write "UnitID="&UnitID&"<BR>"
			If Request.Form("page") = "" Then
				intPage = 1	
				Else
				intPage = Request.Form("page")
			End If				
				HandOffUnitID=Request.Form("HandOffUnitId")
				PageStatus=Request.Form("PageStatus")
				JobNumber=Request.Form("JobNumber")
				'Response.Write "PageStatus="&PageStatus&"<BR>"
				'Response.Write "JobNumber="&JobNumber&"<BR>"
				If PageStatus="PerformHandOff" then
				If JobNumber>"" then
				'''''''''''''''''''''''''''''''
				''''''''''''''EMAILS WARNING
				Body = "Greetings,<br><br>"   & _
				"FYI:  A driver has 'handed off' a job.<br><br>"& _
				"driver: "&FirstName&" "&LastName&"<br>" & _
				"job number: "&JobNumber&"<br><br>"& _
				"job taken from: "& HandOffUnitID &"<br><br>"& _
				"job given to: "&VehicleID&"<br><br>"& _
				"date/time: "&now()&"<br><br>"& _
				"Thank you,<br><br>" & _
				"Mark Maggiore<br>"  & _
				"LogistiCorp Web Developer<br>"  & _
				"mark.maggiore@logisticorp.us<br>"  & _ 
				"(214) 956-0400 xt. 212<br><br>"
				'Recipient = "mark.maggiore@logisticorp.us"
				Set objMail = CreateObject("CDONTS.Newmail")
				objMail.From = "system.monitor@logisticorp.us"
				objMail.To = "mark.maggiore@logisticorp.us"
				'objMail.cc = "x0031708@ti.com"
				objMail.Subject = "Vehicle Hand Off"
				objMail.MailFormat = cdoMailFormatMIME
				objMail.BodyFormat = cdoBodyFormatHTML
				objMail.Body = Body
				objMail.Send
				Set objMail = Nothing		
				'''''''''''''''''''''''''''''
					Set oConn = Server.CreateObject("ADODB.Connection")
					oConn.ConnectionTimeout = 100
					oConn.Provider = "MSDASQL"
					oConn.Open DATABASE
						'l_cSQL = "UPDATE FCREFS SET ref_status = '"&ORDERSTATUS&"' WHERE (rf_fh_id = '"&TheJobNumber&"') AND (rf_ref = '" &TheBarCode& "')"
						'Response.write "UPDATE FCREFS="&l_cSQL&"<BR>"
						'oConn.Execute(l_cSQL)
						oConn.Execute "PHONE_CHANGE_STATUS '" & JobNumber & "', '4', 'ACC', '','','',''" 
						'''''m_cSQL = "UPDATE FCFGTHD SET fh_status = 'ACC', fh_statcode=4 WHERE fh_id = '" & JobNumber&"'"
						'response.write "UPDATE FCFGTHD="&m_cSQL&"<BR>"
						'''''oConn.Execute(m_cSQL)
						n_cSQL = "UPDATE FCLEGS SET fl_dr_id = '"&VehicleID&"' WHERE fl_fh_id = '" & JobNumber&"'"
						'response.write "UPDATE FCLEGS="&n_cSQL&"<BR>"
						oConn.Execute(n_cSQL)
					Set oConn=Nothing
					ErrorMessage="<font color='green'><b>Job #"&JobNumber&" is now yours</b></font>"
					Else
					'Response.Write "hello?"
					ErrorMessage="<font color='red'><b>You did not select a job to take</b></font>"
					
				End if	
				'''''''''''''''''''''''''''''''
				'PageStatus="ShowOrders"
				PageStatus="ShowOrders"
				End if
				
								
				Select Case PageStatus
				Case "ShowOrders"
				X=0
				Y=0
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				
				SQL = "SELECT un_dr_id FROM fcunits"
				SQL = SQL&" WHERE (Un_ID='"& HandOffUnitID &"')"
				'SQL = SQL&" AND ((fh_status='ACC') OR (fh_status='OPN'))"
				'SQL = SQL&" ORDER BY fh_id"
				oRs.Open SQL, DATABASE, 1, 3
				If not oRs.eof then
					HandOffUnitID = trim(oRs("UN_DR_ID"))
				End if
				'oRs=Nothing
				oRs.Close
				'REsponse.Write "HandOffUnitID="& HandOffUnitID &"<BR>"				
				
				
				
				
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				
				SQL = "SELECT Fl_SF_ID, Fh_ID, fh_bt_id, Fl_ST_ID, fl_st_rta, FH_Status, Fh_Priority FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
				SQL = SQL&" WHERE (Fl_dr_ID='"& HandOffUnitID &"') AND (fh_ship_dt>'"&now()-30&"')"
				SQL = SQL&" AND ((fh_status='ACC') OR (fh_status='OPN') OR (fh_status='ARV') OR (fh_status='AC2'))"
				SQL = SQL&" ORDER BY fh_id"
				oRs.Open SQL, DATABASE, 1, 3
				'REsponse.Write "SQL="&SQL&"<BR>"
				%>
					<table cellpadding="0" cellspacing="0" border="0" align="left" bordercolor="red" ID="Table2">
						<tr><td align="center" colspan="9"><form method="post" action="default.asp" ID="Form5"><input type="submit" value="Return to Menu" ID="Submit7" NAME="Submit7"></form></td></tr>
						<tr>
							<td align="center" class="purpleseparator" colspan="13"><b>AVAILABLE HAND OFFS</b></td>
						</tr>
					<form method="post" action="DriverHandOff.asp" ID="Form6">
						<input type="hidden" name="HandOffUnitID" value="<%=HandOffUnitID%>" ID="Hidden9">
						<input type="hidden" name="fh_bt_id" value="<%=fh_bt_id%>" ID="Hidden17">
						<input type="hidden" name="PageStatus" value="PerformHandOff" ID="Hidden8">	
						<input type="hidden" name="page" value="<%=intPage%>" ID="Hidden12">				
				<%
				If not oRs.EOF then
						'm_Logit "OrdersToBePickedUp " & DriverID, oConn
						'm_Logit "OrdersToBePickedUp " & LocationCode, oConn

						CloseTable="y"
						ELSE
						Response.Write "<tr><td colspan='13' align='center'>None currently exist</td></tr><tr><td>&nbsp;</td></tr>"
				End if
				'''''''''''''''''''''''''''''''''''''''''
				'Do while not oRs.eof
				oRS.PageSize = 6
				oRS.CacheSize = oRS.PageSize
				intPageCount = oRS.PageCount
				intRecordCount = oRS.RecordCount
				If (oRS.EOF) then
					'Response.Redirect("default.asp")
					Sendback2="y"
					'Response.Write "Got here #3<br>"
				End if
				If NOT (oRS.BOF AND oRS.EOF) Then

				If CInt(intPage) > CInt(intPageCount) Then intPage = intPageCount
					If CInt(intPage) <= 0 Then intPage = 1
						If intRecordCount > 0 Then
							oRS.AbsolutePage = intPage
							intStart = oRS.AbsolutePosition
							If CInt(intPage) = CInt(intPageCount) Then
								intFinish = intRecordCount
							Else
								intFinish = intStart + (oRS.PageSize - 1)
							End if
						End If
					If intRecordCount > 0 Then
						For intRecord = 1 to oRS.PageSize					
				''''''''''''''''''''''''''''''''''''''''''
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
					'yyy=yyy+1
					'if yyy=1 then
					%>
						<!--	
						<tr>
							<td align="center" nowrap><b>&nbsp;</b></td>
							<td nowrap><b>Job #</b></td>
							<td nowrap><b>Due in</b></td>
							<td nowrap><b>From/To</b></td>
							<td nowrap>
							<%
							if trim(fh_bt_id)<>"26" then
							%>
								<b>Lots</b>
							<%
							End if
							%>
							</td>
						</tr>
						-->
						<%
					'Response.Write "fh_bt_id="&fh_bt_id&"*****<BR>"								
					'if trim(fh_bt_id)<>"26" then
					
					'	Set Recordset1 = Server.CreateObject("ADODB.Recordset")
					'	Recordset1.ActiveConnection = DATABASE
					'	Recordset1.Source = "SELECT count(rf_ref) as NumberOfLots FROM FCREFS WHERE (rf_fh_id='"&JobNumber&"')"
					'	Recordset1.CursorType = 0
					'	Recordset1.CursorLocation = 2
					'	Recordset1.LockType = 1
					'	Recordset1.Open()
					'	Recordset1_numRows = 0
					'	if NOT Recordset1.EOF then
					'		NumberOfLots=Recordset1("NumberOfLots")
					'		If NumberOfLots>1 then WordLots="Lots" end if
					'		If NumberOfLots=1 then WordLots="Lot" end if
					'		If NumberOfLots=0 then WordLots="" end if
					'		Else
					'		ErrorMessage="Incorrect driver ID or password"
					'	End if
					'	Recordset1.Close()
					'	Set Recordset1 = Nothing					
					'End if
					
					
					If Priority="P0" then 
						PriorityColor="red"
						ButtonClass="ButtonRed"
						Else
						ButtonClass="Button1"
						PriorityColor="black"
					End if					
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
						%>
						<TR>
						<td valign="top" nowrap><font color="<%=PriorityColor%>">&nbsp;&nbsp;</font>
							<input type="radio" name="JobNumber" value="<%=JobNumber%>">
						</td>
						 <%	
''''''''''''''''''''''''''''''''
					DisplayToLocation=ToLocation
					DisplayFromLocation=FromLocation
					If Trim(ToLocation)="55" then
						DisplayToLocation="SB-HUB"
					End if
					If Trim(FromLocation)="55" then
						DisplayFromLocation="SB-HUB"
					End if	
''''''''''''''''''''''''''''''''						 
						 
						 					
						Response.Write "<td><a href='DriverTracking.asp?JobNumber="&JobNumber&"&fh_bt_id="&fh_bt_id&"'><font color='"&PriorityColor&"'>"&Right(JobNumber,5)&"</font></a> / "
						Response.Write "<font color='"&PriorityColor&"'>"&DisplayTimeTillDue&" / </font>"
						Response.Write "<font color='"&PriorityColor&"'>"&DisplayFromLocation&"-"&DisplayToLocation&"</font>"
						Response.Write "</td></tr>"
						
					'End if
					x=x+1
					'Response.Write X&") "&LotNumber&"<br>"
					TempJobNumber=JobNumber
					TempX=X
					oRS.MoveNext
					If colorchanger = 1 Then
						colorchanger = 0
						color1 = "class=headerwhite"
						color2 = "class=header"
					Else
						colorchanger = 1
						color1 = "class=header"
						color2 = "class=headerwhite"	
					End If
					If oRS.EOF Then Exit for
						Next
						End if
						End if	
				'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
				'If CloseTable="y" then
					'Response.Write TempX
					'Response.Write "</font></td>"						
				'	
				'	</tr><!--/table-->
				'	
				'	CloseTable="n"
				'End if	
				%>
					<!------------------------------------------------------------->
					 <tr><td align="center" colspan="2"><%=ErrorMessage%></td></tr>
					 <%If CloseTable="y" then%>
						<tr><td align="center" colspan="2"><input type="submit" value="Take Selected Job"></td></tr>
					 <%End if%>
					 </form>
					 <%If CloseTable="y" then%>					
										<tr>
											<td colspan="6">
											<table ID="Table6" width="300" align="center">
							<td width="50%" align="left" valign="top"><font face="Verdana, arial" size="1">
								<%If cInt(intPage) > 1 Then%>
									<form method="post" ID="Form10">
									<input type="submit" name="submit" value="<<Previous" ID="Submit8">
									<input type="hidden" name="pagestatus" value="<%=pagestatus%>" ID="Hidden10">
									<input type="hidden" name="HandOffUnitID" value="<%=HandOffUnitID%>" ID="Hidden11">
									<input type="hidden" name="LocationCode" value="<%=LocationCode%>" ID="Hidden22">
									<input type="hidden" name="truckstatus" value="<%=truckstatus%>" ID="Hidden23">						
									<input type="hidden" name="page" value="<%=intPage - 1%>" ID="Hidden24"></form>
									</form>
									<!--
									<a href="?orderby=<%=orderBy%>&page=<%=intPage - 1%>&SearchFor=<%=SearchFor%>&SearchWhere=<%=SearchWhere%>&SearchVariable=<%=SearchVariable%>"><< <b>Prev</b></a>
									-->
									<%
									else
									Response.write "&nbsp;"
								End If%>
								</font>
							</td>
							<td width="50%" align="right" valign="top"><font face="Verdana, arial" size="1" >
								<%If cInt(intPage) < cInt(intPageCount) Then%>
									<form method="post" ID="Form11">
									<input type="submit" name="submit" value="Next>>" ID="Submit9">
									<input type="hidden" name="pagestatus" value="<%=pagestatus%>" ID="Hidden32">
									<input type="hidden" name="HandOffUnitID" value="<%=HandOffUnitID%>" ID="Hidden34">
									<input type="hidden" name="LocationCode" value="<%=LocationCode%>" ID="Hidden35">
									<input type="hidden" name="truckstatus" value="<%=truckstatus%>" ID="Hidden36">
									<input type="hidden" name="page" value="<%=intPage + 1%>" ID="Hidden37"></form>
									</form>
									<!--
									<a href="?orderby=<%=orderBy%>&page=<%=intPage + 1%>&SearchFor=<%=SearchFor%>&SearchWhere=<%=SearchWhere%>&SearchVariable=<%=SearchVariable%>"><b>Next</b> >></a>
									-->
									<%
									else
									Response.write "&nbsp;"
								End If%>
								</font>
							</td>			</table>				
											</td>
										</tr>
						<%End if%>						
					<!------------------------------------------------------------->				
				<%		
			Case else
			%>
				<TABLE WIDTH="300" border="0" align="left" cellpadding="0" cellspacing="0" ID="Table3">
					<form method="post" action="default.asp" ID="Form4"><tr><td align="center"><input type="submit" value="Return to Menu" ID="Submit6" NAME="Submit6"></td></tr></form>
					<TR> 
						<td> 
								<table border="0" bordercolor="blue" width="300" cellpadding="0" cellspacing="0" ID="Table4">
										<tr>
											<td class='mainpagetextboldcenter' colspan='2' align="center"><b>Job Hand Off<br>(Select a vehicle to take a handoff from)<br></b><br></td>
										</tr>								
									<%if errormessage>"" then%>
										<tr>
											<td class='generalcontenthighlight'colspan='2' align="center"><font color="red"><br><b><%=ErrorMessage%></b><br><br></font></td>
										</tr>
									<%end if
									UnitID=trim(UnitID)
									'Response.Write "unitid="&unitid&"***<BR>"
									Select Case UnitID
										Case "303551", "303552", "303553", "303554"
											%>	
											
											<%if UnitID<>"303551" Then%>
											<FORM ACTION="DriverHandOff.asp" method="post" name="thisForm" ID="Form8">								
											<tr>
												<td class='generalcontent'>
													<input type="submit" name="submit" value="Wafer1" ID="Submit4">									
												</td>
											</tr>
											<input type="hidden" name="HandOffUnitID" value="303551">
											<input type="hidden" name="PageStatus" value="ShowOrders" ID="Hidden1">
											</FORM>
											<tr> 
												<td class="subheader" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
											</tr>									
											<tr> 
												<td class="subheader" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
											</tr>
											<tr><td>&nbsp;</td></tr>									
											<%End if%>

											
											<%if UnitID<>"303552" Then%>
											<FORM ACTION="DriverHandOff.asp" method="post" name="thisForm" ID="Form1">								
											<tr>
												<td class='generalcontent'>
													<input type="submit" name="submit" value="Wafer2" ID="Submit1">									
												</td>
											</tr>
											<input type="hidden" name="HandOffUnitID" value="303552" ID="Hidden2">
											<input type="hidden" name="PageStatus" value="ShowOrders" ID="Hidden3">
											</FORM>
											<tr> 
												<td class="subheader" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
											</tr>									
											<tr> 
												<td class="subheader" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
											</tr>
											<tr><td>&nbsp;</td></tr>										
											<%End if%>

											
											<%if UnitID<>"303553" Then%>
											<FORM ACTION="DriverHandOff.asp" method="post" name="thisForm" ID="Form2">								
											<tr>
												<td class='generalcontent'>
													<input type="submit" name="submit" value="Wafer3" ID="Submit2">									
												</td>
											</tr>
											<input type="hidden" name="HandOffUnitID" value="303553" ID="Hidden4">
											<input type="hidden" name="PageStatus" value="ShowOrders" ID="Hidden5">
											</FORM>
											<tr> 
												<td class="subheader" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
											</tr>									
											<tr> 
												<td class="subheader" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
											</tr>
											<tr><td>&nbsp;</td></tr>										
											<%End if%>

											
											<%if UnitID<>"303554" Then%>
											<FORM ACTION="DriverHandOff.asp" method="post" name="thisForm" ID="Form3">								
											<tr>
												<td class='generalcontent'>
													<input type="submit" name="submit" value="Wafer4" ID="Submit5">									
												</td>
											</tr>
											<input type="hidden" name="HandOffUnitID" value="303554" ID="Hidden6">
											<input type="hidden" name="PageStatus" value="ShowOrders" ID="Hidden7">
											</FORM>
											<tr> 
												<td class="subheader" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
											</tr>	
											<tr> 
												<td class="subheader" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
											</tr>
											<tr><td>&nbsp;</td></tr>																		
											<%End if
										Case "ofb"
											if UnitID<>"srv" Then%>
											<FORM ACTION="DriverHandOff.asp" method="post" name="thisForm" ID="Form7">								
											<tr>
												<td class='generalcontent'>
													<input type="submit" name="submit" value="Stockroom Van" ID="Submit3">									
												</td>
											</tr>
											<input type="hidden" name="HandOffUnitID" value="srv" ID="Hidden13">
											<input type="hidden" name="PageStatus" value="ShowOrders" ID="Hidden14">
											</FORM>
											<tr> 
												<td class="subheader" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
											</tr>									
											<tr> 
												<td class="subheader" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
											</tr>
											<tr><td>&nbsp;</td></tr>										
											<%End if%>
										
										<%
										'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
										Case "srv", "SHERMAN", "OCV"
											if UnitID<>"srv" Then%>
											<FORM ACTION="DriverHandOff.asp" method="post" name="thisForm" ID="Form12">								
											<tr>
												<td class='generalcontent'>
													<input type="submit" name="submit" value="Stockroom Van" ID="Submit11">									
												</td>
											</tr>
											<input type="hidden" name="HandOffUnitID" value="srv" ID="Hidden18">
											<input type="hidden" name="PageStatus" value="ShowOrders" ID="Hidden19">
											</FORM>
											<tr> 
												<td class="subheader" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
											</tr>									
											<tr> 
												<td class="subheader" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
											</tr>
											<tr><td>&nbsp;</td></tr>										
											<%End if
											if UnitID<>"SHERMAN" Then%>
											<FORM ACTION="DriverHandOff.asp" method="post" name="thisForm" ID="Form13">								
											<tr>
												<td class='generalcontent'>
													<input type="submit" name="submit" value="Sherman Vehicle" ID="Submit12">									
												</td>
											</tr>
											<input type="hidden" name="HandOffUnitID" value="SHERMAN" ID="Hidden20">
											<input type="hidden" name="PageStatus" value="ShowOrders" ID="Hidden21">
											</FORM>
											<tr> 
												<td class="subheader" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
											</tr>									
											<tr> 
												<td class="subheader" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
											</tr>
											<tr><td>&nbsp;</td></tr>										
											<%End if
											if UnitID<>"OCV" Then%>
											<FORM ACTION="DriverHandOff.asp" method="post" name="thisForm" ID="Form14">								
											<tr>
												<td class='generalcontent'>
													<input type="submit" name="submit" value="On Call Vehicle" ID="Submit13">									
												</td>
											</tr>
											<input type="hidden" name="HandOffUnitID" value="OCV" ID="Hidden25">
											<input type="hidden" name="PageStatus" value="ShowOrders" ID="Hidden26">
											</FORM>
											<tr> 
												<td class="subheader" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
											</tr>									
											<tr> 
												<td class="subheader" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
											</tr>
											<tr><td>&nbsp;</td></tr>										
											<%End if													
											'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
											%>
											
											<%if UnitID<>"ofb" Then%>
											<FORM ACTION="DriverHandOff.asp" method="post" name="thisForm" ID="Form9">								
											<tr>
												<td class='generalcontent'>
													<input type="submit" name="submit" value="Overflow Bobtail" ID="Submit10">									
												</td>
											</tr>
											<input type="hidden" name="HandOffUnitID" value="ofb" ID="Hidden15">
											<input type="hidden" name="PageStatus" value="ShowOrders" ID="Hidden16">
											</FORM>
											<tr> 
												<td class="subheader" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
											</tr>	
											<tr> 
												<td class="subheader" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
											</tr>
											<tr><td>&nbsp;</td></tr>																		
											<%End if
									End select
										
									%>
																											
								</table>
						</td>
					</TR>
					

				</TABLE>
			<%
			End Select
			%>
			<tr><td>&nbsp;</td></tr>
			</table>				
	</BODY>
</HTML>
