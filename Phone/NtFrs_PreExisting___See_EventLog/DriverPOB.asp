<%@ Language=VBScript %>
<!-- 10/08/2005 Span-Renuka included this screen for CCF 2023 -->
<!-- #include file="../v9web/include/checkstring.inc" -->
<!-- #include file="../v9web/include/custom.inc" -->
<!-- #include file="../v9web/include/ifabsettings.inc" -->
<!-- #include file="driverinfo.inc" -->	
<HTML>
	<HEAD>
		<link rel="stylesheet" type="text/css" href="mainStyleSheet.css">
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
		<script type="text/javascript">
		function formSubmit()
		{
		document.getElementById("thisForm").submit()
		}
		</script>		
		<%
		Dim MultiLocationCode(100)
		Dim MultiBillToID(100)
				
		If Request.Form("page") = "" Then
			intPage = 1	
			Else
			intPage = Request.Form("page")
		End If	
		If Request.Form("page2") = "" Then
			'Response.Write "got here1<br>"
			intPage2 = 1	
			Else
			'Response.Write "got here2<br>"
			intPage2 = Request.Form("page2")
		End If			
		AcknowledgeIt=Request.Form("AcknowledgeIt")
		AliasCode=Request.Form("AliasCode")
		If AliasCode="" then
			AliasCode=Request.QueryString("AliasCode")
		End if
		'If trim(AliasCode)="" then AliasCode="666" end if
		LocationCode=Request.Form("LocationCode")
		FakeSubmit=Request.Form("FakeSubmit")
		If FakeSubmit="" then
			FakeSubmit=Request.QueryString("FakeSubmit")
		End if	
		If FakeSubmit>"" then
			If trim(AliasCode)="" then
				Response.Redirect("default.asp")
				''Response.write "Got here #1<br>"
			End if		
		end if	
		PageStatus=Request.Form("PageStatus")
		If PageStatus="" then
			PageStatus=Request.QueryString("PageStatus")
		End if
		txtJobNumber=Request.Form("txtJobNumber")
		''''''''''''''''''''''''''''''''''''''
		'I had to duplicate this to get rid of the overuse of alias code
		'If I ever have a free few hours, I'll rework it.
		'Love, Mark
		'---------------------------------------------------------------
		If AliasCode>"" then	
			Set Recordset1 = Server.CreateObject("ADODB.Recordset")
			Recordset1.ActiveConnection = DATABASE
			Recordset1.Source = "SELECT st_id, sb_bt_id FROM fcshipto INNER JOIN fcshipbt ON fcshipto.st_id = fcshipbt.sb_st_id WHERE (st_alias='"&AliasCode&"')"
			'Response.write "Recordset1.Source="&Recordset1.Source&"<BR>"
			Recordset1.CursorType = 0
			Recordset1.CursorLocation = 2
			Recordset1.LockType = 1
			Recordset1.Open()
			Recordset1_numRows = 0
			If Recordset1.eof then
				ErrorMessage="That is not a valid location"
			End if			
			
			Do While NOT Recordset1.EOF 
				'SetArrivalTime="y"
				'm=m+1
				LocationCode=trim(Recordset1("st_id"))
				'MultiLocationCode(m)=trim(Recordset1("st_id"))
				'MultiBillToID(m)=Trim(cStr(Recordset1("sb_bt_id")))
				'BillToID=Trim(cStr(Recordset1("sb_bt_id")))
				Recordset1.Movenext
				''Response.write "m="&m&"<br>"
				'Response.write "LocationCode="&LocationCode&"<br>"
				''Response.write "MultiLocationCode(m)="&MultiLocationCode(m)&"<br>"
				''Response.write "MultiBillToID(m)="&MultiBillToID(m)&"<br>"
				Loop
					'Response.Write "</font>"
			Recordset1.Close()
			Set Recordset1 = Nothing
		End if
			'--------------------------------------------------------------		
		'Response.Cookies("FleetXPhone")("sBT_ID")
		'suid=session("suid")
		'response.Write "aliascode="&aliascode&"<BR>"
		'response.Write "billtoID="&BillToID&"<BR>"
		'response.Write "sBT_ID="&sBT_ID&"<BR>"
		If (Locationcode="KWEO" or Locationcode="kweo")  then		
			Set Recordset1 = Server.CreateObject("ADODB.Recordset")
			Recordset1.ActiveConnection = DATABASE
			Recordset1.Source = "SELECT st_pkey, st_id, sb_bt_id FROM fcshipto INNER JOIN fcshipbt ON fcshipto.st_id = fcshipbt.sb_st_id WHERE (st_alias='"&AliasCode&"')"
			'response.write "Recordset1.Source="&Recordset1.Source&"<BR>"
			Recordset1.CursorType = 0
			Recordset1.CursorLocation = 2
			Recordset1.LockType = 1
			Recordset1.Open()
			Recordset1_numRows = 0
			If Recordset1.eof then
				ErrorMessage="That is not a valid location"
			End if			
			If not Recordset1.EOF then 
				SetArrivalTime="y"
				m=m+1
				LocationCode=cStr(trim(Recordset1("st_id")))
				LocationCodePkey=cStr(trim(Recordset1("st_pkey")))
				'Response.Write "LocationCode="&LocationCode&"***<BR>"
				'Response.Write "Now="&now()&"<BR>"
				'Response.Write "BillToID="&BillToID&"<BR>"
				
				
				
				
'-------------------STARTS THE PAPERWORK PICK UP	
				If SetArrivalTime="y" then
						'-----ADDED TO SERVE AS A CLOCK IN FUNCTION IN THE EVENT THAT NO ORDERS EXIST------
						Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open DATABASE
							'response.write "PHONE_ATPAPERWORK_ORDERS '" & PaperOrder & "', '" & LocationCodePkey &"'<br>" 
							oConn.Execute "PHONE_ATPAPERWORK_ORDERS 'ARRIVED', '" & LocationCodePkey &"'" 
						oConn.Close
						Set oConn=Nothing	
						'-----------END NEW FUNCTION				
						'Response.Write "*************GOT HERE!!!"
						Set oRs = Server.CreateObject("ADODB.Recordset")
						oRs.CursorLocation = 3
						oRs.CursorType = 3
						oRs.ActiveConnection = DATABASE	
						''''''CHECKS TO SEE IF ITS OKAY TO CHANGE THE ORDER STATUS
						SQL = "SELECT fl_fh_id FROM fcwttime RIGHT OUTER JOIN fclegs ON fcwttime.wt_fh_id = fclegs.fl_fh_id"
						SQL = SQL&" WHERE (fl_t_int = '1/1/1900') AND (Fl_dr_ID='"&trim(VehicleID)&"') AND (wt_start Is NULL) AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) "
						oRs.Open SQL, DATABASE, 1, 3
						'response.write "sql="&sql&"<BR>"
						If oRs.eof then
							'response.write "none!!!! END OF FILE!!!!<br>"
						End if
						Do while not oRs.EOF
							'Response.Write "Got Here! insert into paper<br>"
							PaperOrder=oRs("fl_fh_id")
							Set oConn = Server.CreateObject("ADODB.Connection")
							oConn.ConnectionTimeout = 100
							oConn.Provider = "MSDASQL"
							oConn.Open DATABASE
								'response.write "PHONE_ATPAPERWORK_ORDERS '" & PaperOrder & "', '" & LocationCodePkey &"'<br>" 
								oConn.Execute "PHONE_ATPAPERWORK_ORDERS '" & PaperOrder & "', '" & LocationCodePkey &"'" 
							oConn.Close
							Set oConn=Nothing
							''''''''''''''''''''''''''''''''''''''''
						oRs.movenext
						Loop
						oRs.Close
						Set oRs=Nothing	
						
						Pagestatus="viewpaper"				
				End if				
				
				
				
				
				
				
				
				
			End if
			Recordset1.Close()
			Set Recordset1 = Nothing		
			else
			If AliasCode>"" and  sBT_ID<>"80" then
				ERRORMESSAGE="You may ONLY scan in at the KWE OFFICE location."
				else
				If trim(ErrorMessage)="" then
					PageStatus="viewpaper"
				End if
			End if
		End if
		''''''''''''''''''''''''''''''''
		
		If FakeSubmit="" then
			'Response.Write "<font color='red'>IF I GOT HERE, THERE'S A PROBLEMO!<BR></font>"
		If ErrorMessage="" then PageStatus="loggedin" End if
		End if
	
		'Response.Write "<font color='red'>pagestatus="&pagestatus&"<BR></font>"
		%>
	</HEAD>
	<%if pagestatus>"" and pagestatus<>"loggedin" then
		%>
		<BODY leftMargin=0 TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<%
		else%>
		<BODY leftMargin=0 TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" OnLoad=document.thisForm.AliasCode.focus()>
	<%end if%>	
	<table width="300" cellpadding="0" cellspacing="0" border="0" ID="Table1">
	<tr><td align="center" colspan="9"><a href="default.asp" class="mainpagelink"><form method="post" action="default.asp" ID="Form8"><input type="submit" value="Return to Menu" ID="Submit1" NAME="Submit3"></form></td></tr>
	
		<%
		'Response.Write "PageStatus="&PageStatus&"<BR>"
		'Response.Write "AcknowledgeIt="&AcknowledgeIt&"<BR>"
		'Response.Write "BillToID="&BillToID&"<BR>"
		Select Case PageStatus
			Case "loggedin"
				%>
				<TR> 
					<td> 
						<div class="purpleseparator"> 
							<table border="0" cellpadding="2" cellspacing="0" ID="Table5" width="100%" bordercolor="blue">
								<tr> 
									<td class="mainpagetextboldright" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
								</tr>
								<tr>
									<td class='mainpagetextboldcenter' colspan="2" nowrap align="center">SCAN in location code</td>
								</tr>
								<form method="post" name="thisForm" id="thisForm">
								<tr>
									<td colspan='2' class='generalcontent' align="center">
										<input maxlength="20" name="AliasCode" id="AliasCode" type="password" size="15">
										<input maxlength='25' size='25' name='VehicleID' id="Hidden2" value='<%=VehicleID%>' type="hidden">
										<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden6">
									</td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan="2" align="center"><input size="8" maxlength="30" type="text" value="<%=Time()%>" name="bogus" ID="bogus" onFocus="formSubmit()" readonly="readonly" class="InvisibleTextFieldPurple"></td></tr>				
								</form>
								<%if errormessage>"" then%>
									<tr>
										<td class='generalcontenthighlight'colspan='2' align="center"><font color="red"><br><b><%=ErrorMessage%></b><br><br></font></td>
									</tr>
								<%end if%>
								<!--									
								<tr>
									<td colspan="2" align="center" class='generalcontent'>
										<input type="submit" name="submit" value="submit" ID="Submit1">									
									</td>
								</tr>
								-->
								
								
								
								<tr> 
									<td class="subheader" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
								</tr>
							</table>
						</div>
					</td>
					<!--Dummy section-->
				</TR>
				<tr><td align="center" colspan="4">&nbsp;</td></tr>		
				<%
			Case "viewpaper"
'-------------------STARTS THE PAPER WORK ON BOARD			
				X=0
				Y=0
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				if trim(vehicleID)<>"199" then
					SQL = "SELECT Fl_SF_ID, Fh_ID, Fl_ST_ID, fl_sf_comment, FH_Status, Fh_Priority FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
					SQL = SQL&" WHERE (fh_ship_dt>'"&now()-30&"') AND  (Fl_dr_ID='"&VehicleID&"')  AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL))"
					SQL = SQL&" AND ((fh_status='ACC'))"
					SQL = SQL&" ORDER BY fh_priority, fh_id"
					else
					SQL = "SELECT Fl_SF_ID, Fh_ID, Fl_ST_ID, fl_sf_comment, FH_Status, Fh_Priority FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
					SQL = SQL&" WHERE (fh_ship_dt>'"&now()-30&"') AND  (Fl_dr_ID='"&VehicleID&"')  AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL))"
					SQL = SQL&" AND ((fh_status='ONB')) AND (fl_sf_id='"& LocationCode &"') and (fl_rt_type<>'out')"
					SQL = SQL&" ORDER BY fh_priority, fh_id"
				End if					
				'Response.Write "SQL="&SQL&"<BR>"
				oRs.Open SQL, DATABASE, 1, 3
				%>
							<tr>
							<td align="center" class="purpleseparator" colspan="3"><b>PAPERWORK TO PICK UP</b></td>
						</tr>
				<%
				If not oRs.EOF then
					%>
						<tr>
							<td>&nbsp;</td>
							<td align="left" nowrap><b>&nbsp;&nbsp;&nbsp;&nbsp;Job #</b></td>
							<!--td align="center" nowrap><b>Details</b></td-->
							<td align="center" nowrap>
							<%If BillToID<>"26" then%>
							<b>Lots</b>
							<%End if%>
							</td>
							</tr>
						<%
						CloseTable="y"
						ELSE
						Response.Write "<tr><td colspan='4' align='center'>No paperwork to pick up here.</td></tr><tr><td>&nbsp;</td></tr>"
				End if
				'''''''''''''''''''''''''''''''''''''''''''''''''''''
				'Do while not oRs.eof
				oRS.PageSize = 4
				oRS.CacheSize = oRS.PageSize
				intPageCount2 = oRS.PageCount
				intRecordCount2 = oRS.RecordCount
				If (oRS.EOF) then
					'Response.Redirect("default.asp")
					Sendback2="y"
					'Response.Write "Got here #3<br>"
				End if
				If NOT (oRS.BOF AND oRS.EOF) Then

				If CInt(intPage2) > CInt(intPageCount2) Then intPage2 = intPageCount2
					If CInt(intPage2) <= 0 Then intPage2 = 1
						If intRecordCount2 > 0 Then
							oRS.AbsolutePage = intPage2
							intStart = oRS.AbsolutePosition
							If CInt(intPage2) = CInt(intPageCount2) Then
								intFinish = intRecordCount
							Else
								intFinish = intStart + (oRS.PageSize - 1)
							End if
						End If
					If intRecordCount2 > 0 Then
						For intRecord2 = 1 to oRS.PageSize				
				'''''''''''''''''''''''''''''''''''''''''''''''''''''
					FromLocation = trim(oRs("Fl_SF_ID"))
					JobNumber = trim(oRs("Fh_ID"))
					ToLocation = trim(oRs("Fl_ST_ID"))
					fl_sf_comment = trim(oRs("fl_sf_comment"))
					JobStatus = trim(oRs("fh_status"))
					Priority = trim(oRs("fh_priority"))
					If Priority="P0" then 
						PriorityColor="red"
						ButtonClass="ButtonRed"
						Else
						ButtonClass="Button1"
						PriorityColor="black"
					End if					
					'LotNumber = oRs("rf_ref")
					Select Case JobStatus
						Case "OPN"
							JobStatus="Open"
							ButtonText="Acknowledge"
						Case "ACC"
							JobStatus="Acknowledged"
							ButtonText="POB"
						Case "ONB"
							JobStatus="ONB"
							ButtonText="POB"
					End Select
					If JobNumber<>TempJobNumber then
						If X>0 then
							'IF trim(BillToID)<>"26" Then
							'	Response.Write TempX
							'	else
							'	Response.Write "&nbsp;"
							'End if
							''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''							
							if trim(fh_bt_id)<>"26" then
							
								Set Recordset1 = Server.CreateObject("ADODB.Recordset")
								Recordset1.ActiveConnection = DATABASE
								Recordset1.Source = "SELECT count(rf_ref) as NumberOfLots FROM FCREFS WHERE (rf_fh_id='"&TempJobNumber&"')"
								Recordset1.CursorType = 0
								Recordset1.CursorLocation = 2
								Recordset1.LockType = 1
								Recordset1.Open()
								Recordset1_numRows = 0
								if NOT Recordset1.EOF then
									NumberOfLots=Recordset1("NumberOfLots")
									Response.Write NumberOfLots
									If NumberOfLots>1 then WordLots="Lots" end if
									If NumberOfLots=1 then WordLots="Lot" end if
									If NumberOfLots=0 then WordLots="" end if
									Else
									Response.Write "&nbsp;"
									ErrorMessage="Incorrect driver ID or password"
								End if
								Recordset1.Close()
								Set Recordset1 = Nothing					
							End if							
							'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
							Response.Write "</font></td></tr>"
							Response.Write "<tr><td colspan='4' bgcolor='white'><img src='images/pixel.gif' height='2' width='1' border='0'></td></tr>"
							Response.Write "<tr><td colspan='4' bgcolor='gray'><img src='images/pixel.gif' height='1' width='1' border='0'></td></tr>"
							Response.Write "<tr><td colspan='4' bgcolor='white'><img src='images/pixel.gif' height='2' width='1' border='0'></td></tr>"
							X=0
						End if
						Y=Y+1
						If Priority="P0" then
							ButtonClass="ButtonRed"
							else
							ButtonClass="Button1"
						End if
						'Response.Write "JobStatus="&JobStatus&"<BR>"
						Select Case JobStatus
							Case "Acknowledged","ONB"
								''''If show=0 then
									'Response.Write "Got here!"
									'Response.Write "ButtonText="&ButtonText&"<BR>"
									''''show=show+1
									BillToID=Request.Cookies("FleetXPhone")("sBT_ID")
									'Response.Write "BillToID="&BillToID&"<BR>"	
									'If BillToID<>"26" then
									
									Select Case BillToID
										Case "26"
										%>
										<form method="post" action="DriverCloseWafer.asp" ID="Form4">
										<%
										Case "48"
										'Response.Write "GOT HERE!!!<BR>"
										%>
										<form method="post" action="DriverPOBChange.asp" ID="Form3">
										<%										
										Case Else
										%>
										<form method="post" action="DriverPOBChange.asp" ID="Form1">
										<!--form method="post" action="DriverClose.asp" ID="Form2"-->
										<%
									End Select
									'End if
									%>
									<td valign="top"><input type="submit" value="<%=ButtonText%>" name="submit" class="<%=ButtonClass%>" ID="Submit2"></td>
									<input type="hidden" name="txtcaller" value="<%=trim(VehicleID)%>" ID="Hidden3">
									<input type="hidden" name="txtstation" value="<%=trim(FromLocation)%>" ID="Hidden4">
									<input type="hidden" name="txtjobnumber" value="<%=trim(jobnumber)%>" ID="Hidden5">
									<input type="hidden" name="VehicleID" value="<%=trim(VehicleID)%>" ID="Hidden25">
									<input type="hidden" name="LocationCode" value="<%=Trim(LocationCode)%>" ID="Hidden26">
									<input type="hidden" name="jobnumber" value="<%=Trim(jobnumber)%>" ID="Hidden27">	
									<input type="hidden" name="AliasCode" value="<%=Trim(AliasCode)%>" ID="Hidden31">
									<input type="hidden" name="BillToID" value="<%=Trim(BillToID)%>" ID="Hidden1">
									<input type="hidden" name="PageStatus" value="POB" ID="Hidden14">								
									</form>
									<%
									''''else
									''''Response.Write "<td>&nbsp;</td>"
								''''End if
							Case Else
								%>
								<td valign="top">&nbsp;</td>
								<%
						End Select
						Response.Write "<td valign='top' nowrap><font color='"&PriorityColor&"'>"&Y&") <a href='DriverTracking.asp?JobNumber="&JobNumber&"&fh_bt_id="&fh_bt_id&"'><font color='"&PriorityColor&"'>"&JobNumber&"</font></a></font></td>"
						Response.Write "<td valign='top' nowrap align='center'><font color='"&PriorityColor&"'>"
						
					End if
					x=x+1
					TempJobNumber=JobNumber
					TempX=X
					''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
					'oRs.Movenext
				'Loop
				Response.Write "</font>"
				'oRs.Close
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
				''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				If CloseTable="y" then
					If BillToID<>"26" then
						Response.Write TempX
					End if
				Response.Write "</font></td>"				
					%>
					</tr>
					<%
					if trim(fl_sf_comment)>"" then
						Response.Write "<tr><td colspan='3'>***"&fl_sf_comment&"<BR></td></tr>"
					End if
					%>
					
					<tr><td>&nbsp;</td></tr><!--/table-->
					<%
					CloseTable="n"
				End if			
	
				If CloseTable="y" then
					%>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					</table>
					<%
					CloseTable="n"
				End if	
				%>

					<!------------------------------------------------------------->
										<tr>
											<td colspan="6">
											<table ID="Table6" width="300" align="center">
							<td width="50%" align="left" valign="top"><font face="Verdana, arial" size="1">
								<%If cInt(intPage2) > 1 Then%>
									<form method="post" ID="Form10">
									<input type="submit" name="submit" value="<<Previous" ID="Submit6">
									<input type="hidden" name="fakesubmit" value="Next>>" ID="Hidden7">
									<input type="hidden" name="pagestatus" value="<%=pagestatus%>" ID="Hidden9">
									<input type="hidden" name="VehicleID" value="<%=VehicleID%>" ID="Hidden10">
									<input type="hidden" name="LocationCode" value="<%=LocationCode%>" ID="Hidden22">
									<input type="hidden" name="truckstatus" value="<%=truckstatus%>" ID="Hidden23">	
									<input type="hidden" name="AliasCode" value="<%=AliasCode%>" ID="Hidden42">							
									<input type="hidden" name="page2" value="<%=intPage2 - 1%>" ID="Hidden24"></form>
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
								<%If cInt(intPage2) < cInt(intPageCount2) Then%>
									<form method="post" ID="Form11">
									<input type="submit" name="submit" value="Next>>" ID="Submit7">
									<input type="hidden" name="fakesubmit" value="Next>>" ID="Submit3">
									<input type="hidden" name="pagestatus" value="<%=pagestatus%>" ID="Hidden32">
									<input type="hidden" name="VehicleID" value="<%=VehicleID%>" ID="Hidden34">
									<input type="hidden" name="LocationCode" value="<%=LocationCode%>" ID="Hidden35">
									<input type="hidden" name="truckstatus" value="<%=truckstatus%>" ID="Hidden36">
									<input type="hidden" name="AliasCode" value="<%=AliasCode%>" ID="Hidden43">		
									<input type="hidden" name="page2" value="<%=intPage2 + 1%>" ID="Hidden37"></form>
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
					<!------------------------------------------------------------->
			<%
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

				'Response.Write "****l_cSQL="&l_cSQL&"<BR>"
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
				<!--TABLE WIDTH="300" cellpadding="0" cellspacing="5" ID="Table2" align="left" border="0" bordercolor="red">
					<tr><td align="center" colspan="3"><form method="post" action="default.asp" ID="Form1"><input type="submit" value="Return to Menuxxx" ID="Submit3" NAME="Submit1"></form></td></tr>			
				</table>
				<br clear="all"-->
			<FORM ACTION="DriverPOB.asp" method="post" name="thisForm" ID="Form6">
				
				<TABLE WIDTH="300" cellpadding="0" cellspacing="5" ID="Table3" align="left" border="0" bordercolor="red">
					<TR> 
						<td> 
							<div class="purpleseparator"> 
								<table border="0" cellpadding="2" cellspacing="0" ID="Table4" width="100%" bordercolor="blue">
									<tr> 
										<td class="mainpagetextboldright" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
									</tr>
									<tr>
										<td class='mainpagetextboldcenter' colspan="2" nowrap align="center">SCAN in location code</td>
									</tr>
									<tr>
										<td colspan='2' class='generalcontent' align="center">
											<input maxlength="20" name="AliasCode" id="txtstation" type="password" size="15">
											<input maxlength='25' size='25' name='VehicleID' id='VehicleID' value='<%=VehicleID%>' type="hidden">
											<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden16">
										</td>
									</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan="2" align="center"><input size="8" maxlength="30" type="text" value="<%=Time()%>" name="bogus" ID="Text1" onFocus="formSubmit()" readonly="readonly" class="InvisibleTextFieldPurple"></td></tr>				
									
									<%if errormessage>"" then%>
										<tr>
											<td class='generalcontenthighlight'colspan='2' align="center"><font color="red"><br><b><%=ErrorMessage%></b><br><br></font></td>
										</tr>
									<%end if%>
									<!--									
									<tr>
										<td colspan="2" align="center" class='generalcontent'>
											<input type="submit" name="submit" value="submit" ID="Submit1">									
										</td>
									</tr>
									-->
									
									
									
									<tr> 
										<td class="subheader" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
									</tr>
								</table>
							</div>
						</td>
						<!--Dummy section-->
					</TR>
					<tr><td align="center" colspan="4">&nbsp;</td></tr>					
				</TABLE>
			</FORM>
			
			<%
			End Select
			If sendback1="y" and sendback2="y" then
				'Response.Write "Got here sendback!<br>"
				'Response.Redirect("default.asp")
			End if
			'Response.Write "AliasCode="&AliasCode&"<BR>"
			%>
			
	</BODY>
</HTML>
