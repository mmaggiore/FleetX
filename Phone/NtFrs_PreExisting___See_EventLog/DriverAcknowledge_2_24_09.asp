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
		FORMJOBSTATUS=TRIM(Request.Form("FORMJOBSTATUS"))
		AcknowledgeIt=Request.Form("AcknowledgeIt")
		DriverID=Request.Form("DriverID")
		LocationCode=Request.Form("LocationCode")
		Submit=Request.Form("Submit")
		PageStatus=Request.Form("PageStatus")
		PageStatus="loggedin"
		txtJobNumber=Request.Form("txtJobNumber")
		If Submit="submit" then
			If DriverID="" then
				ErrorMessage="You must provide your driver id"
			End if
		If ErrorMessage="" then PageStatus="loggedin" End if
		End if
		'Response.Write "userid="&userID&"<br>"
		'Response.Write "vehicleid="&vehicleID&"<br>"
		'Response.Write "unitid="&unitid&"<br>"
		'Response.Write "driverid="&driverid&"<br>"
		%>
	</HEAD>
	<BODY leftMargin=0 TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<%
		
		Select Case PageStatus
			Case "loggedin"
				If AcknowledgeIt="y" then
''''''''''''''''''''ERROR HANDLING TO PREVENT TIME FLIPPING'''''''''''''''''''''''''''''''
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				SQL = "SELECT Fh_ID FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
				SQL = SQL&" WHERE (Fl_dr_ID='"&VehicleID&"') AND (Fh_ID='"&txtJobNumber&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL))"
				''''If VehicleID=124 then
					'''''SQL = SQL&" AND ((((Fh_Status='ARV') AND (fl_st_id<>'TOPPAN')) AND (Fl_SecAcc is NULL)) "
					SQL = SQL&" AND (((Fh_Status='ARV') AND (Fl_SecAcc is NULL)) "
					'''''else
					SQL = SQL&" OR (Fh_Status='OPN')) "
				'''''End if
				SQL = SQL&" ORDER BY Fh_Priority, fh_status desc, fh_id"
				
				'response.write "XXXXXXXXSQL="&SQL&"<BR>"
				'''''''''''''''''''''''''
				oRs.Open SQL, DATABASE, 1, 3
						If not oRs.EOF then
							OKToChange="y"
							ELSE
							OKToChange="n"
						End if	
				oRs.Close
				Set oRs=Nothing						
''''''''''''''''''''END ERROR HANDLING'''''''''''''''''''''''''''''''				
				
					'If VehicleID<>124 then
					If FORMJOBSTATUS<>"ARV" AND OKToChange="y" then
						Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open DATABASE
								''''l_cSQL = "UPDATE fcfgthd SET fh_status = 'ACC', fh_statcode = 4 WHERE fh_id = '" & txtJobNumber & "'"
							'response.Write "txtJobNumber="&txtJobNumber&"<br>"
							oConn.Execute "PHONE_CHANGE_STATUS '" & txtJobNumber & "', '4', 'ACC', '','','',''" 
							'response.write "UPDATE 1<BR>"
							'oConn.Execute(l_cSQL)
						Set oConn=Nothing
					End if						
					Set oConn = Server.CreateObject("ADODB.Connection")
					oConn.ConnectionTimeout = 100
					oConn.Provider = "MSDASQL"
					oConn.Open DATABASE
						'If VehicleID=124 then
						If FORMJOBSTATUS="ARV" AND OKToChange="y" then
							oConn.Execute "PHONE_CHANGE_STATUS '" & txtJobNumber & "', '55', 'AC2','','','',''" 
							'response.write "UPDATE 2<BR>"
							'else
							'l_cSQL = "UPDATE fclegs SET fl_t_acc='"&now()&"' WHERE fl_fh_id = '" & txtJobNumber & "'"
						End if
						'oConn.Execute(l_cSQL)
					Set oConn=Nothing					
				End if
'------------------------------ACKNOWLEDGES ALL
			If AcknowledgeIt="all" then
				X=0
				Y=0
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				
				SQL = "SELECT Fl_SF_ID, Fh_ID, fh_ship_dt, Fl_ST_ID, FH_Status, Fh_Priority, Fh_User5 FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
				SQL = SQL&" WHERE (Fl_dr_ID='"&VehicleID&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL))"
				SQL = SQL&" AND (((Fh_Status='OPN'))"
				''If VehicleID=124 or VehicleID=313 then
					SQL = SQL&" OR (Fh_Status='ARV')) "
				'''End if
				'Response.Write "ZZZZZZZZZZZZZZZzSQL="&SQL&"<BR>"
				oRs.Open SQL, DATABASE, 1, 3
						If not oRs.EOF then
							ELSE
						End if
						Do while not oRs.eof
						FromLocation = oRs("Fl_SF_ID")
						JobNumber = oRs("Fh_ID")
						ToLocation = oRs("Fl_ST_ID")
						MaterialType = oRs("Fh_User5")
						JobStatus = oRs("fh_status")
						FORMJOBSTATUS=Trim(JobStatus)
						Priority = oRs("fh_priority")
						ShipTime = oRs("fh_ship_dt")
						TimeSincePlaced=DateDiff("n",shiptime,now())
						'Response.Write "JobNumber="&JobNumber&"<BR>"
						'response.write "MaterialType="&MaterialType&"***<BR>"
						If priority="P0" or priority="P1" or MaterialType="Secure Waf" or MaterialType="secret" then
							DontRedirect="y"
						End if
						If AcknowledgeIt="all" and Priority<>"P0" and Priority<>"P1" and MaterialType<>"Secure Waf" and MaterialType<>"secret" then
							If FORMJOBSTATUS<>"ARV" then
								Set oConn = Server.CreateObject("ADODB.Connection")
								oConn.ConnectionTimeout = 100
								oConn.Provider = "MSDASQL"
								oConn.Open DATABASE
										oConn.Execute "PHONE_CHANGE_STATUS '" & JobNumber & "', '4', 'ACC', '','','',''" 
										'response.write "UPDATE 3<BR>"
										'''''l_cSQL = "UPDATE fcfgthd SET fh_status = 'ACC', fh_statcode = 4 WHERE fh_id = '" & JobNumber & "'"
									'''''oConn.Execute(l_cSQL)
								oConn.close
								Set oConn=Nothing
							End if
							Set oConn = Server.CreateObject("ADODB.Connection")
							oConn.ConnectionTimeout = 100
							oConn.Provider = "MSDASQL"
							oConn.Open DATABASE
								If FORMJOBSTATUS="ARV" then
									oConn.Execute "PHONE_CHANGE_STATUS '" & JobNumber & "', '55', 'AC2', '','','',''" 
									'response.write "UPDATE 4<BR>"
									'''''else
									'''''l_cSQL = "UPDATE fclegs SET fl_t_acc='"&now()&"' WHERE fl_fh_id = '" & JobNumber & "'"
								End if
								'''''oConn.Execute(l_cSQL)
							oConn.close
							Set oConn=Nothing					
						End if							
							Y=Y+1
						PreviousPriorityColor=PriorityColor
						TempJobNumber=JobNumber
						oRs.Movenext
						Loop
						oRs.Close
						Set oRs=Nothing	
						If DontRedirect<>"y" then			
							Response.Redirect("default.asp")
							'response.write "got here as well!<BR>"
						End if
				End if
'-------------------STARTS THE OTHER ORDERS IN THE PHONE
%>
					<table cellpadding="0" width="300" cellspacing="0" bordercolor="red" border="0" align="left" ID="Table5">
						<tr><td align="center" colspan="3"><form method="post" action="default.asp" ID="Form2"><input type="submit" value="Return to Menu" ID="Submit2" NAME="Submit2"></form></td></tr>
						<tr>
							<td align="center" class="purpleseparator" colspan="3"><b>NEW ORDERS</b></td>
						</tr>
						<form method="post" action="DriverAcknowledge.asp" ID="Form333">
						<tr><td valign="top" colspan="3" align="center"><input type="submit" value="Acknowledge ALL" name="submit" class="<%=ButtonGrey%>" ID="Submit4"></td></tr>
						<input type="hidden" name="txtcaller" value="<%=DriverID%>" ID="Hidden1">
						<input type="hidden" name="txtstation" value="<%=ToLocation%>" ID="Hidden2">
						<input type="hidden" name="txtjobnumber" value="<%=jobnumber%>" ID="Hidden3">
						<input type="hidden" name="DriverID" value="<%=DriverID%>" ID="Hidden4">
						<input type="hidden" name="LocationCode" value="<%=LocationCode%>" ID="Hidden5">
						<input type="hidden" name="AcknowledgeIt" value="all" ID="Hidden6">
						<input type="hidden" name="PageStatus" value="loggedin" ID="Hidden7">
						</form>	
<%

				X=0
				Y=0
				If Request.Form("page") = "" Then
					intPage = 1	
					Else
					intPage = Request.Form("page")
				End If				
				
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				SQL = "SELECT Fl_SF_ID, Fh_ID, Fh_User5, fh_ship_dt, fh_bt_id, fh_user5, Fl_ST_ID, Fl_St_Rta, Fl_FirstDrop, FH_Status, Fh_Priority FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
				SQL = SQL&" WHERE (Fl_dr_ID='"&VehicleID&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL))"
				''''If VehicleID=124 then
					'''''SQL = SQL&" AND ((((Fh_Status='ARV') AND (fl_st_id<>'TOPPAN')) AND (Fl_SecAcc is NULL)) "
					SQL = SQL&" AND (((Fh_Status='ARV') AND (Fl_SecAcc is NULL)) "
					'''''else
					SQL = SQL&" OR (Fh_Status='OPN')) "
				'''''End if
				SQL = SQL&" ORDER BY Fh_Priority, fh_status desc, fh_id"
				
				'response.write "XXXXXXXXSQL="&SQL&"<BR>"
				'''''''''''''''''''''''''
				oRs.Open SQL, DATABASE, 1, 3
				
				
				
				
'RS.Open SQL, INTRANET, 1, 3
oRS.PageSize = 6
oRS.CacheSize = oRS.PageSize
intPageCount = oRS.PageCount
intRecordCount = oRS.RecordCount
If (oRS.EOF) then
	'response.write "SQL="&SQL&"<BR>"
	'response.write "got here!"
	Response.Redirect("default.asp")
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
				''''''''''''''''''''''''''

						'''''''''''''''''''''''''''''''
	'					If not oRs.EOF then
	'						'CloseTable="y"
	'						ELSE
	'						Response.Redirect("default.asp")
	'						'Response.Write "<tr><td colspan='3' align='center'>There are currently no unacknowledged orders.</td></tr><tr><td>&nbsp;</td></tr>"
	'					End if
	'					Do while not oRs.eof
						'''''''''''''''''''''''''''''''
						FromLocation = oRs("Fl_SF_ID")
						JobNumber = oRs("Fh_ID")
						MaterialType = oRs("Fh_User5")
						BillToID=Trim(cStr(oRs("Fh_bt_id")))
						ToLocation = oRs("Fl_ST_ID")
						JobStatus = oRs("fh_status")
						FORMJOBSTATUS=JobStatus
						'response.write "FORMJOBSTATUS="&FORMJOBSTATUS&"<BR>"
						Priority = oRs("fh_priority")
						ShipTime = oRs("fh_ship_dt")
						'MaterialType = oRs("fh_user5")
						'Response.Write "materialtype="&MaterialType&"<BR>"
						If MaterialType="300 mm Waf" or MaterialType="Foup/Fosby" then
							MaterialSymbol="*"
							else
							MaterialSymbol=""							
						End if			
						DueTime=oRs("fl_st_rta")
						Fl_FirstDrop=oRs("Fl_FirstDrop")
						'Response.Write "fromlocation="&fromlocation&"<br>"
						If trim(fromLocation)="xx55" or trim(fromLocation)="72" then
							'Response.Write "GOT HERE????<BR>"
							If Priority="P0" then
								DueTime=DateAdd("n", 45, Fl_firstdrop)
								else
								DueTime=DateAdd("n", 120, Fl_firstdrop)
							End if
						End if						
						TimeSincePlaced=DateDiff("n",shiptime,now())
						
						TimeTillDue=DateDiff("n",now(),DueTime)
						'Response.Write "TimeTillDue="&TimeTillDue&"<Br>"
						If TimeTillDue<0 then
							DisplayTimeTillDue="LATE"
							Else
							HoursTillDue=TimeTillDue/60
							HoursTillDue=Int(HoursTillDue)
							MinutesTillDue=TimeTillDue-(HoursTillDue*60)
							DisplayTimeTillDue=trim(HoursTillDue&"h "&MinutesTilldue&"m")
							'Response.Write "DisplayTimeTillDue="&DisplayTimeTillDue&"<BR>"
						End if
						If AcknowledgeIt="all" and Priority<>"P0" and Priority<>"P1" and MaterialType<>"Secure Waf" and MaterialType<>"secret" then
							If FORMJOBSTATUS<>"ARV" then
								Set oConn = Server.CreateObject("ADODB.Connection")
								oConn.ConnectionTimeout = 100
								oConn.Provider = "MSDASQL"
								oConn.Open DATABASE
								oConn.Execute "PHONE_CHANGE_STATUS '" & JobNumber & "', '4', 'ACC', '','','',''" 
										'''''l_cSQL = "UPDATE fcfgthd SET fh_status = 'ACC', fh_statcode = 4 WHERE fh_id = '" & JobNumber & "'"
									'''''oConn.Execute(l_cSQL)
								Set oConn=Nothing
							End if
							Set oConn = Server.CreateObject("ADODB.Connection")
							oConn.ConnectionTimeout = 100
							oConn.Provider = "MSDASQL"
							oConn.Open DATABASE
								If FORMJOBSTATUS="ARV" then
									oConn.Execute "PHONE_CHANGE_STATUS '" & JobNumber & "', '55', 'AC2', '','','',''" 
									'''''else
									'''''l_cSQL = "UPDATE fclegs SET fl_t_acc='"&now()&"' WHERE fl_fh_id = '" & JobNumber & "'"
								End if
								'''''oConn.Execute(l_cSQL)
							Set oConn=Nothing					
						End if							
						If Priority="P0" then 
							PriorityColor="red"
							ButtonClass="ButtonRed"
							Else
							ButtonClass="Button1"
							PriorityColor="black"
						End if
						If MaterialType="Secure Waf" or MaterialType="secret" then
							PriorityColor="Orange"
						End if
						
						Select Case JobStatus
							Case "OPN", "ARV"
								JobStatus="Open"
								ButtonText="Ack"
							Case "ACC"
								JobStatus="Acknowledged"
								ButtonText="On Board"
							Case "ONB"
								JobStatus="On Board"
								ButtonText="Close"
						End Select
							Y=Y+1
									%>
									<form method="post" action="DriverAcknowledge.asp" ID="Form1">
									<tr><td valign="top" width="20"><input type="submit" value="<%=ButtonText%>" name="submit" class="<%=ButtonGrey%>" ID="Submit3"></td>
									<input type="hidden" name="page" value="<%=intPage%>" ID="Hidden9">
									<input type="hidden" name="txtcaller" value="<%=DriverID%>" ID="Hidden14">
									<input type="hidden" name="txtstation" value="<%=ToLocation%>" ID="Hidden15">
									<input type="hidden" name="txtjobnumber" value="<%=jobnumber%>" ID="Hidden16">
									<input type="hidden" name="DriverID" value="<%=DriverID%>" ID="Hidden17">
									<input type="hidden" name="LocationCode" value="<%=LocationCode%>" ID="Hidden31">
									<input type="hidden" name="AcknowledgeIt" value="y" ID="Hidden32">
									<input type="hidden" name="PageStatus" value="loggedin" ID="Hidden33">	
									<input type="hidden" name="FORMJOBSTATUS" value="<%=FORMJOBSTATUS%>" ID="Hidden10">								
									<%
									If BillToID<>"26" then
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
											If NumberOfLots>1 then WordLots="Items"&MaterialSymbol&"/" end if
											If NumberOfLots=1 then WordLots="Item"&MaterialSymbol&"/" end if
											If NumberOfLots=0 then WordLots="" end if
											Else
											ErrorMessage="Incorrect driver ID or password"
										End if
										NumberOfLots="/"&MaterialSymbol&NumberOfLots
										Recordset1.Close()
										Set Recordset1 = Nothing
										Priority="/"&Priority
										Else
										Priority=""
									End if	
							''''''''''''''''''''''''''''''''
							'Response.Write "VehicleID="&VehicleID&"<BR>"
							'Response.Write "ToLocation="&ToLocation&"<BR>"
							'Response.Write "FromLocation="&FromLocation&"<BR>"
							'Response.Write "JobStatus="&JobStatus&"<BR>"
							
							DisplayToLocation=trim(ToLocation)
							DisplayFromLocation=trim(FromLocation)
							If Trim(ToLocation)="55" or Trim(ToLocation)="CPGP" or Trim(ToLocation)="72" or Trim(ToLocation)="TOPPAN" or Trim(ToLocation)="PHO" then
								'Response.Write "GOT HERE!<BR>"
								DisplayFromLocation="SB-HUB"
								If (trim(ToLocation)="TOPPAN" or trim(ToLocation)="CPGP" or trim(ToLocation)="PHO") and lcase(JobStatus)="open" and (trim(VehicleID)<>"611" and trim(VehicleID)<>"612" and trim(VehicleID)<>"613" and trim(VehicleID)<>"112" and trim(VehicleID)<>"123") then
									DisplayFromLocation=trim(FromLocation)
									DisplayToLocation="SB-HUB"
								End if
							End if

							'response.Write "VehicleID="&trim(VehicleID)&"***<br>"
							If Trim(FromLocation)="55" or Trim(FromLocation)="72" or Trim(FromLocation)="TOPPAN" or Trim(FromLocation)="PHO" then
								DisplayFromLocation="SB-HUB"
								If trim(VehicleID)="611" or trim(VehicleID)="612" or trim(VehicleID)="613" or trim(VehicleID)="112" or trim(VehicleID)="123" then
									'response.Write "GOT HERE!!!<BR>"
									DisplayFromLocation=FromLocation
									If trim(DisplayFromLocation)="55" then DisplayFromLocation="CPGP" End if
									DisplayToLocation="SB-HUB"
								End if
							End if
							
							If trim(VehicleID)="123" and trim(ToLocation)="TISHERMA" then
								'REsponse.Write "GOT HERE!<BR>"
								DisplayToLocation="SB-HUB"
							End if
							If trim(VehicleID)="613" and trim(ToLocation)="TISHERMA" then
								'REsponse.Write "GOT HERE!<BR>"
								DisplayFromLocation="SB-HUB"
							End if							
							
							Response.Write "<td valign='top' nowrap>&nbsp;&nbsp;<font color='"&PriorityColor&"'><a href='DriverTracking.asp?JobNumber="&JobNumber&"&fh_bt_id="&fh_bt_id&"'><font color='"&PriorityColor&"'>"&Right(JobNumber,5)&"</font></a>/"&Priority&NumberOfLots&" "&WordLots&"</font>"
							Response.Write "<font color='"&PriorityColor&"'>"&DisplayTimeTillDue&"&nbsp;</font>/"							
							Response.Write "<font color='"&PriorityColor&"'>"&DisplayFromLocation&"-"&DisplayToLocation&"</font></td></tr>"
							Response.Write "</form>"
							Response.Write "<tr><td colspan='3'><hr width='100%'></td></tr>"
						PreviousPriorityColor=PriorityColor
						TempJobNumber=JobNumber
						
						
						
						
						
						'''''''''''''''''''''''''''''''''
'						oRs.Movenext
'						Loop
'						oRs.Close
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
	'End if
						''''''''''''''''''''''''''''''''''
						If CloseTable="y" then
							%>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr></table>					
							<%
							CloseTable="n"
						End if					
					Case else
						Response.Write "Error #1465<br>"
				End Select
			%>
							<tr>
								<td colspan="2">
								<table ID="Table1" width="300" align="center">
				<td width="50%" align="left" valign="top"><font face="Verdana, arial" size="1">
					<%If cInt(intPage) > 1 Then%>
						<form method="post" ID="Form3">
						<input type="submit" name="submit" value="<<Previous" ID="Submit1">
						<input type="hidden" name="page" value="<%=intPage - 1%>" ID="Hidden8"></form>
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
						<form method="post" ID="Form4">
						<input type="submit" name="submit" value="Next>>" ID="Submit5">
						<input type="hidden" name="page" value="<%=intPage + 1%>" ID="Hidden11"></form>
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
							</tr></table>		
			
			
	</BODY>
</HTML>
