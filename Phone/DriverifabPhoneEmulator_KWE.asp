<%@ Language=VBScript %>
<!-- 10/08/2005 Span-Renuka included this screen for CCF 2023 -->
<!-- #include file="../v9web/include/checkstring.inc" -->
<!-- #include file="../v9web/include/custom.inc" -->
<!-- #include file="../v9web/include/ifabsettings.inc" -->
<!-- #include file="driverinfo.inc" -->	
<HTML>
	<HEAD>
		<!--meta http-equiv="refresh" content="100"-->
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
		OtherBillToID=Request.Cookies("Phone")("sBT_ID")	
		fh_bt_id=Request.Cookies("Phone")("sBT_ID")
		'Response.Write "OtherBillToID="&OtherBillToID&"<BR>"
		Hub2=Request.Form("Hub2")		
		Hub3=Request.Form("Hub3")	
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
		If AliasCode>"" then Response.Cookies("Phone")("AliasCode")=AliasCode end if
		If aliasCode="" then aliasCode=Request.Cookies("Phone")("AliasCode") end if
		'Response.write "***AliasCode="&AliasCode&"<BR>"
		'If AliasCode>"" then
		''Response.write "ALIASCODE=XX"&AliasCode&"XX<br>"
		'End if
		'If trim(AliasCode)="" then AliasCode="666" end if
		LocationCode=Request.Form("LocationCode")
		FakeSubmit=Request.Form("FakeSubmit")
		If FakeSubmit="" then
			FakeSubmit=Request.QueryString("FakeSubmit")
		End if
		If FakeSubmit>"" then
			Response.Cookies("Phone")("FakeSubmit")=FakeSubmit
		End if
		If FakeSubmit="" then
			FakeSubmit=Request.Cookies("Phone")("FakeSubmit")
		End if
		'Response.write "FakeSubmit="&FakeSubmit&"<BR>"		
		PageStatus=Request.Form("PageStatus")

		txtJobNumber=Request.Form("txtJobNumber")
		If FakeSubmit="fakesubmit" then
		If trim(AliasCode)="" then
			'Response.Redirect("default.asp")
			'Response.write "Got here #1<br>"
		End if
			'Response.Write "DATABASE="&DATABASE&"<BR>"
			'Response.Write "billtoID="&fh_bt_id&"<BR>"
			Set Recordset1 = Server.CreateObject("ADODB.Recordset")
			Recordset1.ActiveConnection = DATABASE
			Recordset1.Source = "SELECT st_id, sb_bt_id FROM fcshipto INNER JOIN fcshipbt ON fcshipto.st_id = fcshipbt.sb_st_id WHERE (st_alias='"&AliasCode&"')"
			'response.write "Recordset1.Source="&Recordset1.Source&"<BR>"
			Recordset1.CursorType = 0
			Recordset1.CursorLocation = 2
			Recordset1.LockType = 1
			Recordset1.Open()
			Recordset1_numRows = 0
			If Recordset1.eof then
				ErrorMessage="That is not a valid location"
			End if			
			
			Do While NOT Recordset1.EOF 
			'Response.write "BillToID="&BillToID&"<br>"
			'Response.write "suid="& session("suid") &"<br>"
				SetArrivalTime="y"
				m=m+1
				LocationCode=Recordset1("st_id")
				'REsponse.Write "xxxxxxxLocationCode="&LocationCode&"<BR>"
				MultiLocationCode(m)=trim(Recordset1("st_id"))
				MultiBillToID(m)=Trim(cStr(Recordset1("sb_bt_id")))
				BillToID=Trim(cStr(Recordset1("sb_bt_id")))
				If OtherBillToID="80" then
					BillToID="80"
				End if
				'Response.write "BillToID="&BillToID&"<br>"
				'Response.write "BillToID="&BillToID&"<br>"
				Recordset1.Movenext
				''Response.write "m="&m&"<br>"
				'Response.write "LocationCode="&LocationCode&"<br>"
				''Response.write "MultiLocationCode(m)="&MultiLocationCode(m)&"<br>"
				'Response.write "MultiBillToID(m)="&MultiBillToID(m)&"<br>"
				'Response.write "BillToID="&BillToID&"<br>"
				Loop
					Response.Write "</font>"
			Recordset1.Close()
			Set Recordset1 = Nothing		
		
		
		''Response.write "GOT HERE 1<br>"
		'If UCASE(AliasCode)="EBHUB" or UCASE(AliasCode)="13601" or UCASE(AliasCode)="K13536" then
			'If UCASE(AliasCode)="EBHUB" then
			AliasCode=UCASE(ALIASCODE)
			LocationCode=Trim(UCASE(LOCATIONCODE))
			'Response.Write "AliasCode=*"&AliasCode&"*<BR>"
			'Response.Write "LocationCode=*"&LocationCode&"*<BR>"
			
			Select Case LocationCode
				Case "EBHUB"
				BillToID="26"
				'LocationCode="EBHUB"
				Hub="y"
				'Response.Write "Got here #5...<BR>"
				Case "13601", "13536"
				'Response.write "Got here 2<br>"
				BillToID="48"
				'LocationCode=UCASE(AliasCode)
				Hub2="y"
				'Response.Write "Got here #4...<BR>"
				Case "SBRT"
					BillToID="38"
					Hub3="y"
					'Response.Write "Got here #3...<BR>"
			End Select
				DisplayLocationCode=LocationCode
				If Trim(LocationCode)="SBRT" then DisplayLocationCode="SB-HUB" end if	
				If Trim(LocationCode)="80" then DisplayLocationCode="LSP Warehouse" end if			
			'Response.Write "Got here #6...<BR>"
			'End if			
		'Else
		'End if
			'REsponse.Write "Database="&Database&"<BR>"

		'End if		
		If ErrorMessage="" then PageStatus="loggedin" End if
		End if
		If PageStatus>"" then 
			Response.Cookies("Phone")("PageStatus")=PageStatus 
			'Response.write "sets the cookie?<br>"		
		end if
		If PageStatus="" then 
			PageStatus=Request.Cookies("Phone")("PageStatus") 
			'Response.write "takes it from the cookie?<br>"
		end if
		'Response.write "PageStatus="&PageStatus&"<BR>"	
		
		%>
	</HEAD>
	<%if pagestatus>"" then%>
		<BODY leftMargin=0 TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<%else%>
		<BODY leftMargin=0 TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" OnLoad=document.thisForm.AliasCode.focus()>
	<%end if%>	
		<%
		'Response.Write "PageStatus="&PageStatus&"<BR>"
		'Response.Write "AcknowledgeIt="&AcknowledgeIt&"<BR>"
		'response.write "LocationCode="&LocationCode&"<BR>"
		''''''''''Response.write "PageStatus="&PageStatus&"<BR>"
		'Response.write "hub="&hub&"<BR>"
		'Response.write "hub1="&hub1&"<BR>"
		'Response.write "hub2="&hub2&"<BR>"
		'Response.write "hub3="&hub3&"<BR>"
		
		Select Case PageStatus
			Case "loggedin"
'-------------------STARTS THE DROP OFF				
				
				X=0
				Y=0
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				
				SQL = "SELECT Fl_SF_ID, Fh_ID, fh_User5, Fl_ST_ID, fl_sf_comment, FH_Status, Fh_Priority FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
				SQL = SQL&" WHERE (fh_ship_dt>'"&now()-30&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND (Fl_dr_ID='"&VehicleID&"') and ("
				If HUB2<>"y" and HUB3<>"y" then
					'''''Response.Write "got here 4<br>"
					SQL = SQL&"(Fl_st_ID='"&LocationCode&"')"
					'''''Response.Write "***SQL="&SQL&"<BR>"
				End if
				If trim(DisplayLocationCode)="D1" then
					SQL = SQL&" OR (Fl_st_ID='D7')"
					SQL = SQL&" OR (Fl_st_ID='P1')"
				End if
				If HUB="y" then
				'Response.Write "got here 3<br>"
					SQL = SQL&" OR (Fl_st_ID='D6W3')"
					SQL = SQL&" OR (Fl_st_ID='D6N2')"
					SQL = SQL&" OR (Fl_st_ID='D6N1')"
					SQL = SQL&" OR (Fl_st_ID='DM4M')"
					SQL = SQL&" OR (Fl_st_ID='DM5M')"
					SQL = SQL&" OR (Fl_st_ID='DPI2')"
					SQL = SQL&" OR (Fl_st_ID='DPI3')"
					SQL = SQL&" OR (Fl_st_ID='ESTK')"
					SQL = SQL&" OR (Fl_st_ID='DM5Q')"
					SQL = SQL&" OR (Fl_st_ID='DM6Q')"
				End if
				If HUB2="y" then
					SQL = SQL&" (Fl_st_ID='xxx')"
				End if	
				If HUB3="y" and VehicleID<>"113" then
					''''Response.Write "got here 1<br>"
					SQL = SQL&" (((Fl_st_ID='CPGP') or (Fl_sf_ID='55') or (FL_ST_ID='TOPPAN') or (FL_SF_ID='TOPPAN') or (FL_SF_ID='PHO') or (FL_ST_ID='PHO') or (FL_ST_ID='TISHERMA') or (FL_ST_ID='TISHERMANRET') or (FL_SF_ID='TISHERMANRET')) AND (Fh_Status<>'DPV') )"
				End if								
				For mm=1 to m
				'''''Response.Write "got here 2<br>"
					'response.write "got hereXXXYYYZZZ???<br>"
					SQL = SQL&" OR (Fl_st_ID='"&MultiLocationCode(mm)&"')"
					'SQL666 = SQL666&" OR (Fl_st_ID='"&MultiLocationCode(mm)&"')"
					'Response.Write "SQL666="&sql666&"<BR>"
				next
				''''Response.Write "HUB="&HUB&"<BR>"
				''''Response.Write "VehicleID="&VEhicleID&"<BR>"
				'If VehicleID=124 then
				'	Response.Write "got here 1<br>"
				'	If HUB="y" then
				'		Response.Write "got here 2<br>"
				'		SQL = SQL&" ) AND ((fh_status='DPV')"
				'		else
				'		SQL = SQL&" ) AND ((fh_status='ONB')"
				'	End if
				'	else
				'	SQL = SQL&" ) AND ((fh_status='ONB')"
				'End if
				If (vehicleID=124 and HUB="y") or (vehicleID=113 and HUB3="y") then
					SQL = SQL&") AND (fh_status='ONB')"
					else
					if trim(vehicleID)="199" then
						SQL = SQL&") AND (((((fh_status='ONB') AND  (fl_rt_type='out')) OR (fh_status='PUO'))))"
						else
						'SQL = SQL&") AND (((fh_status='DPV') and (FL_ST_ID<>'TOPPAN')) OR (fh_status='ONB'))"
						SQL = SQL&") AND (((fh_status='DPV')) OR (fh_status='ONB'))"
					End if
				End if
				'End if
				'SQL = SQL&" ORDER BY fh_priority, fh_id"
				
				'response.write "DROP OFF SQL="&SQL&"<BR>"
				'response.write "HUB="&HUB&"<BR>"
				'response.write "HUB2="&HUB2&"<BR>"
				'response.write "HUB3="&HUB3&"<BR>"
				
				oRs.Open SQL, DATABASE, 1, 3
				If trim(DisplayLocationCode)="55" then DisplayLocationCode="CPGP" end if
				If trim(DisplayLocationCode)="48" then DisplayLocationCode="KWEO" end if
				'If trim(DisplayLocationCode)="80" then DisplayLocationCode="LSP Warehouse" end if
				%>
					<table width="300" cellpadding="0" cellspacing="0" border="0" bordercolor="green" align="left" ID="Table1">
						<tr><td align="center" colspan="3"><form method="post" action="default.asp" ID="Form7"><input type="submit" value="Return to Menu" ID="Submit1" NAME="Submit1"></form></td></tr>
						<tr>
							<td class="mainpagetextboldcenter" colspan="3" align="center">
								<font color="blue">Last update: <%=Time()%></font>
							</td>
						</tr>						
						<tr>
							<td align="center" class="purpleseparator" colspan="3"><b>DROP OFFS AT <%=uCase(DisplayLocationCode)%></b></td>
						</tr>
				<%
				If not oRs.EOF then
					%>
						<tr>
							<td align="center">&nbsp;</td>						
							<td align="left" nowrap><b>&nbsp;&nbsp;&nbsp;&nbsp;Job #</b></td>
							<!--td align="center" nowrap><b>Details</b></td-->
							<td align="center" nowrap>
							<%If BillToID<>"26" then%>
								<%If BillToID="80" then%>
									<b>Items</b>
									<%else%>
									<b>Lots</b>
									<%
								End if
							End if%>
							</td>
							</tr>
						<%
						CloseTable="y"
						ELSE
						Response.Write "<tr><td colspan='4' align='center'>No orders to drop off here.</td></tr><tr><td>&nbsp;</td></tr>"
				End if
				'''''''''''''''''''''''''''''''''''''''''''''''''''''''
				
				'Do while not oRs.eof
				oRS.PageSize = 4
				oRS.CacheSize = oRS.PageSize
				intPageCount = oRS.PageCount
				intRecordCount = oRS.RecordCount
				If (oRS.EOF) then
					''Response.Redirect("default.asp")
					sendback1="y"
					'Response.Write "Got here #2<br>"
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
				
				
				'''''''''''''''''''''''''''''''''''''''''''''''''''''''
					FromLocation = oRs("Fl_SF_ID")
					JobNumber = oRs("Fh_ID")
					MaterialType = oRs("Fh_User5")
					ToLocation = oRs("Fl_ST_ID")
					fl_sf_comment = oRs("fl_sf_comment")
					JobStatus = oRs("fh_status")
					'response.Write "JobNumber="&JobNumber&"<BR>"
					'response.Write "JobStatus="&JobStatus&"<BR>"
					Priority = oRs("fh_priority")
					If Priority="P0" then 
						PriorityColor="red"
						ButtonClass="ButtonRed"
						Else
						ButtonClass="Button1"
						PriorityColor="black"
					End if
					If MaterialType="Secure Waf" OR MaterialType="secret" OR MaterialType="ITAR" then
						PriorityColor="Orange"
					End if					
					'LotNumber = oRs("rf_ref")
					Select Case JobStatus
						Case "OPN"
							JobStatus="Open"
							ButtonText="Acknowledge"
						Case "ACC"
							JobStatus="Acknowledged"
							ButtonText="ONB"
						Case "ONB", "DPV"
							JobStatus="ONB"
							ButtonText="CLS"
						Case "PUO"
							JobStatus="Paper on Board"
							ButtonText="CLS"							
					End Select
					'FromLocation = oRs("Fl_SF_ID")
					If JobNumber<>TempJobNumber then
						If X>0 or X=0 then
							'If BillToID<>"26" then
							'	Response.Write TempX
							'End if
							''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''							
							if trim(fh_bt_id)<>"26" then
							
								Set Recordset1 = Server.CreateObject("ADODB.Recordset")
								Recordset1.ActiveConnection = DATABASE
								SQL_111="SELECT count(rf_ref) as NumberOfLots FROM FCREFS WHERE (rf_fh_id='"&JobNumber&"')"
								Recordset1.Source = SQL_111
								'Response.write "SQL_111="&SQL_111&"<BR>"
								Recordset1.CursorType = 0
								Recordset1.CursorLocation = 2
								Recordset1.LockType = 1
								Recordset1.Open()
								Recordset1_numRows = 0
								if NOT Recordset1.EOF then
									NumberOfLots=Recordset1("NumberOfLots")
									'Response.Write NumberOfLots
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
							If trim(fl_sf_comment)>"" then
								Response.Write "<tr><td colspan='3'>***"&fl_sf_comment&"</td></tr>"
							end if							
							Response.Write "<tr><td colspan='4' bgcolor='white'><img src='images/pixel.gif' height='2' width='1' border='0'></td></tr>"							
							X=0
							
						End if
						Y=Y+1
						If Priority="P0" then
							ButtonClass="ButtonRed"
							else
							ButtonClass="Button1"
						End if
						'Response.Write "got here #1<BR>"
						Select Case JobStatus
							Case "Acknowledged","ONB", "Paper on Board"
								''''If show2=0 then
									'Response.Write "Got here!"
									'Response.Write "cookie?(sBT_ID)="& request.cookies("sBT_ID") &"<BR>"
									''''show2=show2+1
					
									If (Request.Cookies("Phone")("sBT_ID")<>"26" AND Request.Cookies("Phone")("sBT_ID")<>"80"  AND Request.Cookies("Phone")("sBT_ID")<>"48" AND Request.Cookies("Phone")("sBT_ID")<>"75" AND MaterialType<>"xxxSecure Waf") OR (Request.Cookies("Phone")("sBT_ID")="26" AND trim(FromLocation)="PHO") then
										'Response.Write "cookie?(sBT_ID)="&Request.Cookies("Phone")("sBT_ID")&"<BR>"
										'Response.Write "BillToID="&BillToID&"<BR>"
										'Response.Write "materialtype="&materialtype&"***<BR>"
										'Response.Write "Driver Close Wafer<br>"
										%>
										<form method="post" action="DriverCloseWafer.asp" ID="Form3">

										<%
										Else
										'Response.Write "cookie?(sBT_ID)="&Request.Cookies("Phone")("sBT_ID")&"<BR>"
										'Response.Write "fromlocation="&fromlocation&"<BR>"
										'Response.Write "tolocation="&tolocation&"***<BR>"
										'Response.Write "Driver Close<br>"
										'response.Write "GOT HERE!!!!!<BR>"
										%>

										<form method="post" action="DriverClose.asp" ID="Form5">
										<%
									End if
									%>
									<tr><td valign="top"><input type="submit" value="<%=ButtonText%>" name="submit" class="<%=ButtonClass%>" ID="Submit4"></td>
									<input type="hidden" name="txtcaller" value="<%=trim(VehicleID)%>" ID="Hidden6">
									<input type="hidden" name="txtstation" value="<%=trim(ToLocation)%>" ID="Hidden7">
									<input type="hidden" name="txtjobnumber" value="<%=trim(jobnumber)%>" ID="Hidden8">
									<input type="hidden" name="VehicleID" value="<%=trim(VehicleID)%>" ID="Hidden28">
									<input type="hidden" name="LocationCode" value="<%=trim(LocationCode)%>" ID="Hidden29">
									<input type="hidden" name="jobnumber" value="<%=trim(jobnumber)%>" ID="Hidden30">	
									<input type="hidden" name="PageStatus" value="CLS" ID="Hidden15">
									<input type="hidden" name="BillToID" value="<%=Request.Cookies("Phone")("sBT_ID")%>" ID="Hidden2">
									<input type="hidden" name="AliasCode" value="<%=trim(AliasCode)%>" ID="Hidden33">
									<input type="hidden" name="Hub2" value="<%=Hub2%>" ID="Hidden49">
									<input type="hidden" name="Hub" value="<%=Hub%>" ID="Hidden50">																		
									<!--
									<input type="hidden" name="" value="<%=x%>" ID="Hidden3">
									<input type="hidden" name="" value="<%=x%>" ID="Hidden4">
									-->
									</form>
									<%
									''''ELSE
									''''Response.Write "<td>&nbsp;</td>"
								''''End if
							Case Else
								%>
								<tr><td valign="top">&nbsp;</td>
								<%
						End Select
						'Response.Write "<td>|</td>"						
						Response.Write "<td valign='top' nowrap><font color='"&PriorityColor&"'>"&Y&") <a href='DriverTracking.asp?JobNumber="&JobNumber&"&fh_bt_id="&fh_bt_id&"'><font color='"&PriorityColor&"'>"&JobNumber&"</font></a></font></td>"
						'Response.Write "<td valign='top' nowrap align='center'><font color='"&PriorityColor&"'>"&Priority&"</font>-"
						'Response.Write "<font color='"&PriorityColor&"'>"&FromLocation&"</font></td>"
						Response.Write "<td valign='top' nowrap align='center'><font color='"&PriorityColor&"'>"
'''''''''''''''''''''''

''''''''''''''''''''''''''''''						
					End if
					x=x+1
					'Response.Write X&") "&LotNumber&"<br>"
					TempJobNumber=JobNumber
					TempX=X
					If NumberOfLots>=1 then
						TempX=NumberOfLots
					end if					
					'Response.Write "----------------------------<br>"
					''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
					'oRs.Movenext
					'Loop
				
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
					
				'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
				If CloseTable="y" then
					If BillToID<>"26" then
						Response.Write TempX
					End if
				Response.Write "</font></td>"
					%>
					</tr><!--/table-->
					<%
					'fl_sf_comment="here is a comment!!!!"
					'''If trim(fl_sf_comment)>"" then
					'''	Response.Write "<tr><td colspan='3'>***"&fl_sf_comment&"</td></tr>"
					'''end if
					%>
					<!------------------------------------------------------------->
										<tr>
											<td colspan="6">
											<table ID="Table6" width="300" align="center">
							<td width="50%" align="left" valign="top"><font face="Verdana, arial" size="1">
								<%If cInt(intPage) > 1 Then%>
									<form method="post" ID="Form8">
									<input type="submit" name="submit" value="<<Previous" ID="Submit5">
									<input type="hidden" name="pagestatus" value="<%=pagestatus%>" ID="Hidden11">
									<input type="hidden" name="VehicleID" value="<%=VehicleID%>" ID="Hidden20">
									<input type="hidden" name="LocationCode" value="<%=LocationCode%>" ID="Hidden21">
									<input type="hidden" name="truckstatus" value="<%=truckstatus%>" ID="Hidden12">						
									<input type="hidden" name="page" value="<%=intPage - 1%>" ID="Hidden13">
									<input type="hidden" name="AliasCode" value="<%=Trim(AliasCode)%>" ID="Hidden40">	
									<input type="hidden" name="Hub2" value="<%=Hub2%>" ID="Hidden44">
									<input type="hidden" name="Hub" value="<%=Hub%>" ID="Hidden45">		
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
									<form method="post" ID="Form9">
									
									
									<input type="hidden" name="txtcaller" value="<%=trim(VehicleID)%>" ID="Hidden48">
									<input type="hidden" name="txtstation" value="<%=trim(ToLocation)%>" ID="Hidden51">
									<!--input type="hidden" name="txtjobnumber" value="<%=trim(jobnumber)%>" ID="Hidden52"-->
									<!--input type="hidden" name="jobnumber" value="<%=trim(jobnumber)%>" ID="Hidden55"-->	
									<!--input type="hidden" name="PageStatus" value="CLS" ID="Hidden56"-->
									<input type="hidden" name="BillToID" value="<%=BillToID%>" ID="Hidden57">
										
									
									
									
									<input type="hidden" name="AliasCode" value="<%=AliasCode%>" ID="Hidden41">		
									<input type="hidden" name="Hub2" value="<%=Hub2%>" ID="Hidden46">
									<input type="hidden" name="Hub" value="<%=Hub%>" ID="Hidden47">									
									<input type="hidden" name="pagestatus" value="<%=pagestatus%>" ID="Hidden17">
									<input type="hidden" name="VehicleID" value="<%=VehicleID%>" ID="Hidden38">
									<input type="hidden" name="LocationCode" value="<%=LocationCode%>" ID="Hidden39">
									<input type="hidden" name="truckstatus" value="<%=truckstatus%>" ID="Hidden18">
									<input type="hidden" name="page" value="<%=intPage + 1%>" ID="Hidden19">										
									<input type="submit" name="submit" value="Next>>" ID="Submit8">
								
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
					CloseTable="n"
				End if
'-------------------STARTS THE PICK UP	
				'Response.Write "BillToID="&BillToID&"<BR>"
				'Response.Write "SetArrivalTime="&SetArrivalTime&"<BR>"	
				If SetArrivalTime="y" and (BillToID="48" or BillToID="80") then
						'Response.Write "GOT HERE!!!"
						Set oRs = Server.CreateObject("ADODB.Recordset")
						oRs.CursorLocation = 3
						oRs.CursorType = 3
						oRs.ActiveConnection = DATABASE	
						''''''CHECKS TO SEE IF ITS OKAY TO CHANGE THE ORDER STATUS
						SQL = "SELECT fl_fh_id FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id "
						'SQL = SQL&" WHERE (Fl_sf_ID='"&trim(LocationCode)&"') AND (fl_t_atp = '1/1/1900') AND (Fl_dr_ID='"&trim(VehicleID)&"') AND (fl_t_int > '1/1/1900') AND (fh_ready<='"& NOW() &"')"
						Select Case BillToID
							Case "48"
								SQL = SQL&" WHERE (Fl_sf_ID='"&trim(LocationCode)&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND (fl_t_atp = '1/1/1900') AND (Fl_dr_ID='"&trim(VehicleID)&"') AND (fl_t_int > '1/1/1900')"
							Case else
								if trim(vehicleID)="198" then
									SQL = SQL&" WHERE (Fl_sf_ID='"&trim(LocationCode)&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND (fl_t_atp = '1/1/1900') AND (Fl_dr_ID='"&trim(VehicleID)&"') AND (fl_t_int > '1/1/1900')"
									else
									SQL = SQL&" WHERE (Fl_sf_ID='"&trim(LocationCode)&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND (fl_t_atp = '1/1/1900') AND (Fl_dr_ID='"&trim(VehicleID)&"')"
								End if
						End Select
						'SQL = SQL&" AND (Ref_Status='o')"
						'SQL = SQL&" ORDER BY Fh_Priority, fh_status desc, fh_id"
						'Response.write "<br><font color='blue'>Change the status?="&SQL&"<BR></font>"
						'Response.Write "SQL="&SQL&"<BR>"
						oRs.Open SQL, DATABASE, 1, 3
						If oRs.eof then
							'Response.Write "none!!!!<br>"
						End if
						Do while not oRs.EOF
							AtAirlineOrder=oRs("fl_fh_id")
							'Response.Write "AtAirlineOrder="&AtAirlineOrder&"<BR>"
							''''''''''''''''''''''''''''''''''''''''
							'''''n_cSQL = "UPDATE FCLEGS SET fl_t_atd = '"&now()&"' WHERE fl_fh_id = '" & TheJobNumber&"'"
							'Response.write "UPDATE FCLEGS="&n_cSQL&"<BR>"
							'''''oConn.Execute(n_cSQL)							
							'''''m_cSQL = "UPDATE FCFGTHD SET fh_status = 'CLS', fh_statcode=9 WHERE fh_id = '" & TheJobNumber&"'"
							'Response.write "UPDATE FCFGTHD="&m_cSQL&"<BR>"
							'''''oConn.Execute(m_cSQL)
							'Response.Write "Database="&database&"<BR>"
							Set oConn = Server.CreateObject("ADODB.Connection")
							oConn.ConnectionTimeout = 100
							oConn.Provider = "MSDASQL"
							oConn.Open DATABASE
								oConn.Execute "PHONE_ATAIRLINE_ORDERS '" & AtAirlineOrder & "'" 
							oConn.Close
							Set oConn=Nothing
							''''''''''''''''''''''''''''''''''''''''
						oRs.movenext
						Loop
						oRs.Close
						Set oRs=Nothing					
				
				
				
				
				
				
				
				
					''Set oConn = Server.CreateObject("ADODB.Connection")
					''oConn.ConnectionTimeout = 100
					''oConn.Provider = "MSDASQL"
					''oConn.Open DATABASE	
					'Response.Write "Database="&Database&"<BR>"			
					''l_cSQL = "UPDATE FCLEGS SET fl_t_und = '"&now()&"' WHERE (Fl_sf_ID='"&trim(LocationCode)&"') AND (fl_t_und = '1/1/1900') AND (Fl_dr_ID='"&trim(VehicleID)&"') AND (fl_t_int > '1/1/1900')"
					''Response.write "UPDATE Wafers="&l_cSQL&"<BR>"
					''oConn.Execute(l_cSQL)	
					''oConn.Close
				End if					
				X=0
				Y=0
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				
				SQL = "SELECT Fl_SF_ID, Fh_ID, fh_User5, Fl_ST_ID, fl_sf_comment, FH_Status, Fh_Priority FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
				SQL = SQL&" WHERE (fh_ship_dt>'"&now()-30&"') AND ((FL_Leg_Status='c')OR (FL_Leg_Status is NULL)) AND  (Fl_dr_ID='"&VehicleID&"') "
				If VehicleID=124 or vehicleID=113 then
					SQL = SQL&" and ((Fl_sf_ID='zzz') "
					else
					SQL = SQL&" and ((Fl_sf_ID='"&LocationCode&"') "
				End if
				If HUB="y" and (VehicleID=124 or vehicleID=113) then
					SQL = SQL&" OR (Fl_st_ID='D6W3')"
					SQL = SQL&" OR (Fl_st_ID='D6N2')"
					SQL = SQL&" OR (Fl_st_ID='D6N1')"
					SQL = SQL&" OR (Fl_st_ID='DM4M')"
					SQL = SQL&" OR (Fl_st_ID='DM5M')"
					SQL = SQL&" OR (Fl_st_ID='DPI2')"
					SQL = SQL&" OR (Fl_st_ID='DPI3')"
					SQL = SQL&" OR (Fl_st_ID='ESTK')"
					SQL = SQL&" OR (Fl_st_ID='DM5Q')"
					SQL = SQL&" OR (Fl_st_ID='DM6Q')"
				End if
				'Response.Write "Hub3="&Hub3&"*<br>"
				'Response.Write "vehicleID="&vehicleID&"*<br>"
				
				If HUB3="y" and (trim(vehicleID)<>"611" and trim(vehicleID)<>"612" and trim(vehicleID)<>"613" and trim(vehicleID)<>"112" and trim(vehicleID)<>"123") then
					SQL = SQL&" OR (Fl_sf_ID='CPGP') OR (Fl_sf_ID='55') OR (Fl_sf_ID='72') OR (FL_sf_ID='TOPPAN') OR (FL_st_ID='TOPPAN') OR (FL_st_ID='PHO') OR (FL_sf_ID='PHO') OR (FL_sf_ID='TISHERMANRET')"
				End if	
				If HUB3="y" and (trim(vehicleID)="611" or trim(vehicleID)="612" or trim(vehicleID)="613" or trim(vehicleID)="112" or trim(vehicleID)="123") then
					SQL = SQL&" OR (Fl_st_ID='CPGP') OR (Fl_st_ID='55') OR (Fl_st_ID='72') OR (FL_st_ID='TOPPAN') OR (FL_st_ID='PHO') OR (FL_st_ID='TISHERMA') OR (FL_st_ID='HFABRET') OR (FL_st_ID='TISHERMANRET')"
				End if	
				
								
				For nn=1 to m
					'Response.Write "got here<br>"
					SQL = SQL&" OR (Fl_sf_ID='"&MultiLocationCode(nn)&"')"
					'SQL666 = SQL666&" OR (Fl_st_ID='"&MultiLocationCode(nn)&"')"
					'Response.Write "SQL666="&sql666&"<BR>"
				next
				'''If VehicleID<>124 then				
					SQL = SQL&" ) "	
				'''End if
				If OtherBillToID="48" or trim(vehicleID)="198" then
					SQL = SQL&" AND ((fh_status='PUO') or (fh_status='AC2'))"
					Else	
					SQL = SQL&" AND ((fh_status='ACC')"
					'''''If VehicleID=124 then
						SQL = SQL&" OR ((fh_status='AC2') AND (fl_secacc is not null))"
						'If fl_st_id="TISHERMANRET" then
						'	SQL = SQL&" AND (ref_status='H')"
						'End if
					'''''End if					
					SQL = SQL&") "
				End if
				SQL = SQL&" ORDER BY fh_priority, fh_id"
				
				'response.write "<br><br>PICK UP SQL="&SQL&"<BR>"
				
				oRs.Open SQL, DATABASE, 1, 3
				DisplayLocationCode=LocationCode
				If Trim(LocationCode)="SBRT" then DisplayLocationCode="SB-HUB" end if
				If Trim(LocationCode)="55" then DisplayLocationCode="CPGP" end if
				If trim(LocationCode)="48" then DisplayLocationCode="KWEO" end if
				If trim(LocationCode)="80" then DisplayLocationCode="LSP Warehouse" end if
				%>
					<!--table width="700" cellpadding="0" cellspacing="0" border="0" align="center" ID="Table2"-->
						<tr>
							<td align="center" class="purpleseparator" colspan="3"><b>PICK UPS AT <%=uCase(DisplayLocationCode)%></b></td>
						</tr>
				<%
				If not oRs.EOF then
					%>
						<tr>
							<td>&nbsp;</td>
							<td align="left" nowrap><b>&nbsp;&nbsp;&nbsp;&nbsp;Job #</b></td>
							<!--td align="center" nowrap><b>Details</b></td-->
							<td align="center" nowrap>
							<%If BillToID<>"26" and BillToID<>"80" then%>
							<b>Lots</b>
							<%else%>
							<b>Refs</b>
							<%End if%>
							</td>
							</tr>
						<%
						CloseTable="y"
						ELSE
						Response.Write "<tr><td colspan='4' align='center'>No orders to pick up here.</td></tr><tr><td>&nbsp;</td></tr>"
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
					MaterialType = oRs("Fh_user5")
					ToLocation = trim(oRs("Fl_ST_ID"))
					fl_sf_comment=trim(oRs("fl_sf_comment"))
					JobStatus = trim(oRs("fh_status"))
					Priority = trim(oRs("fh_priority"))
					If Priority="P0" then 
						PriorityColor="red"
						ButtonClass="ButtonRed"
						Else
						ButtonClass="Button1"
						PriorityColor="black"
					End if
					If MaterialType="Secure Waf" OR MaterialType="secret" OR MaterialType="ITAR" then
						PriorityColor="Orange"
					End if					
					'LotNumber = oRs("rf_ref")
					Select Case JobStatus
						Case "OPN"
							JobStatus="Open"
							ButtonText="Acknowledge"
						Case "ACC", "PUO", "ARV", "AC2"
							JobStatus="Acknowledged"
							ButtonText="ONB"
						Case "ONB"
							JobStatus="ONB"
							ButtonText="CLS"
					End Select
					If JobNumber<>TempJobNumber then
						'Response.write "X="&X&"<BR>"
						If X>=0 then
							'IF trim(BillToID)<>"26" Then
							'	Response.Write TempX
							'	else
							'	Response.Write "&nbsp;"
							'End if
							''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''							
							if trim(fh_bt_id)<>"26" then
								'Response.Write "I GOT HERE!!!!<BR>"
								'Response.Write "fh_bt_id="&fh_bt_id&"<BR>"
								Set Recordset1 = Server.CreateObject("ADODB.Recordset")
								SQL_99="SELECT count(rf_ref) as NumberOfLots FROM FCREFS WHERE (rf_fh_id='"&JobNumber&"') AND ((ref_status<>'X') OR (ref_status is NULL))"
								'Response.Write "SQL_99="&SQL_99&"<BR>"
								Recordset1.ActiveConnection = DATABASE
								Recordset1.Source = SQL_99
								Recordset1.CursorType = 0
								Recordset1.CursorLocation = 2
								Recordset1.LockType = 1
								Recordset1.Open()
								Recordset1_numRows = 0
								if NOT Recordset1.EOF then
									NumberOfLots=Recordset1("NumberOfLots")
									'Response.Write "XXXNumberofLots="&NumberOfLots&"<BR>"
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
						'Response.Write "Got here 2!<BR>"
						Select Case JobStatus
							Case "Acknowledged","ONB", "ARV", "AC2"
								''''If show=0 then
									'Response.Write "Got here!"
									'Response.Write "ButtonText="&ButtonText&"<BR>"
									''''show=show+1
									'Response.Write "BillToID=***"&BillToID&"****<BR>"	
									'Response.Write "cookie?(sBT_ID)=***"&cookie?("sBT_ID")&"****<BR>"	
									If Request.Cookies("Phone")("sBT_ID")<>"26" AND Request.Cookies("Phone")("sBT_ID")<>"48" AND Request.Cookies("Phone")("sBT_ID")<>"75" AND Request.Cookies("Phone")("sBT_ID")<>"80"   then
										%>
										<form method="post" action="DriverCloseWafer.asp" ID="Form4">
										<%
										'Response.Write "got here!<br>"
										Else
										'Response.write "Didnt get here!<br>"
										%>
										<form method="post" action="DriverClose.asp" ID="Form2">
										<%
									End if
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
									<input type="hidden" name="PageStatus" value="ONB" ID="Hidden14">								
									</form>
									<%
									'response.Write "FromLocation="&FromLocation&"<BR>"
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
				Response.Write NumberOfLots&"</font>"
						numberoflots=0
				'oRs.Close
				
						If trim(fl_sf_comment)>"" then
						Response.Write "<tr><td colspan='3'>***"&fl_sf_comment&"</td></tr>"
					end if			
				
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
						'Response.Write TempX
						'Response.Write NumberOfLots
						'numberoflots=0
					End if
					Response.Write "</font></td>"				
					%>
					</tr>
					<%
					'If trim(fl_sf_comment)>"" then
					'	Response.Write "<tr><td colspan='3'>***"&fl_sf_comment&"</td></tr>"
					'End if
					%>
					<tr><td>&nbsp;</td></tr><!--/table-->
					<%
					CloseTable="n"
					'else
					'Response.Write NumberOfLots&"XXX"
					'numberoflots=0	
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
									<input type="hidden" name="txtcaller" value="<%=trim(VehicleID)%>" ID="Hidden52">
									<input type="hidden" name="txtstation" value="<%=trim(ToLocation)%>" ID="Hidden53">
									<!--input type="hidden" name="txtjobnumber" value="<%=trim(jobnumber)%>" ID="Hidden52"-->
									<!--input type="hidden" name="jobnumber" value="<%=trim(jobnumber)%>" ID="Hidden55"-->	
									<!--input type="hidden" name="PageStatus" value="CLS" ID="Hidden56"-->
									<input type="hidden" name="BillToID" value="<%=BillToID%>" ID="Hidden54">
										
									
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
						'Response.Write "GOT HERE!!!!!<BR>"
						'm_logit "SETCOOKIE " & LocationAlias, oConn
				End if
			Set oConn=Nothing				
			%>
				<TABLE WIDTH="300" cellpadding="0" cellspacing="5" ID="Table2" align="left" border="0" bordercolor="red">
					<tr><td align="center" colspan="3"><form method="post" action="default.asp" ID="Form1"><input type="submit" value="Return to Menu" ID="Submit3" NAME="Submit1"></form></td></tr>			
				</table>
				<br clear="all">
			<FORM ACTION="DriverifabPhoneEmulator_KWE.asp" method="post" name="thisForm" ID="thisForm">
				
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
