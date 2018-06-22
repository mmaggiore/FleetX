<%@ Language=VBScript %>
<!-- #include file="FleetX.inc" -->
<!-- #include file="driverinfo.inc" -->	
<HTML>
	<HEAD>
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
		<link rel="stylesheet" type="text/css" href="mainStyleSheet.css">
		<%
		BillToID=Request.Cookies("Phone")("sBT_ID")	
		AcknowledgeIt=Request.Form("AcknowledgeIt")
		DriverID=Request.Form("DriverID")
		LocationCode=Request.Form("LocationCode")
		Submit=Request.Form("Submit")
		PageStatus=Request.Form("PageStatus")
		PageStatus="loggedin"
		If Request.Form("page") = "" Then
			intPage = 1	
			Else
			intPage = Request.Form("page")
		End If
		txtJobNumber=Request.Form("txtJobNumber")
		TruckStatus=Request.Form("TruckStatus")
		'Response.Write "TruckStatus="&TruckStatus&"<BR>"
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
    <!-- #include file="LogoSection.asp" -->
		<%
		Select Case PageStatus
			Case "loggedin"
'-------------------STARTS THE DROP OFF	
				If TruckStatus="dropoff" then			
				X=0
				Y=0
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				
				SQL = "SELECT Fl_SF_ID, Fh_ID, fh_User5, Fl_ST_ID, fl_st_name, fl_st_building, fl_st_addr1, fl_st_addr2, fl_st_city, fl_st_rta, fl_firstdrop, fl_sf_comment, fh_bt_id, FH_Status, Fh_Priority FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
				SQL = SQL&" WHERE (Fl_un_ID='"&VehicleID&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND (fh_ship_dt>'"&now()-30&"')"
				If trim(vehicleID)<>"199" then
					SQL = SQL&" AND ((fh_status='ONB')"
					else
					SQL = SQL&" AND (((fh_status='ONB') AND  (fl_rt_type='out')) OR (fh_status='PUO')"
				End if	
				'''''If VehicleID=124 then
					'SQL = SQL&" OR ((fh_status='DPV') AND (fl_st_id<>'CPGP'))"
					SQL = SQL&" OR ((fh_status='DPV'))"
				'''''End if
				SQL = SQL&") ORDER BY fh_priority desc, fl_st_rta"
				'Response.write "SQL="&SQL&"<BR>"
				oRs.Open SQL, DATABASE, 1, 3
				%>
					<table width="300" cellpadding="0" cellspacing="0" border="0" align="left" ID="Table1">
                    <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>
                    <form method="post" action="DriverTruckLoad.asp" ID="Form1">
						<tr><td align="center" colspan="4"><a href="default.asp" class="mainpagelink"><input type="submit" value="Return to Routing" id="gobutton" name="Submit1"></td></tr>
                        </form>
                        <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>
                        <tr>
		                    <td class="FleetXRedSection" colspan="4" align="center">
			                    ORDERS IN <%=uCase(VehicleName)%>
		                    </td>
	                    </tr>
                        <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>
				<%
				''''''If not oRs.EOF then
						'm_Logit "OrdersToBeDroppedOff " & DriverID, oConn
						'm_Logit "OrdersToBeDroppedOff " & LocationCode, oConn

						''''''CloseTable="y"
						''''''ELSE
						''''''Response.Write "<tr><td colspan='7' align='center'>There are currently no orders in the vehicle.</td></tr><tr><td>&nbsp;</td></tr>"
				''''''End if
				''''''Do while not oRs.eof
'''''''''''''''''''''''''''''''''''''''''''''''
oRS.PageSize = 6
oRS.CacheSize = oRS.PageSize
intPageCount = oRS.PageCount
intRecordCount = oRS.RecordCount
If (oRS.EOF) then
	'response.write "REDIRECT!<BR>"
	'Response.Redirect("default.asp")
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
						
					FromLocation = oRs("Fl_SF_ID")
					JobNumber = oRs("Fh_ID")
					MaterialType = oRs("Fh_User5")
					ToLocation = oRs("Fl_ST_ID")


                    fl_st_name = oRs("fl_st_name")
                    fl_st_building = oRs("fl_st_building")
                    fl_st_addr1 = oRs("fl_st_addr1")
                    fl_st_addr2 = oRs("fl_st_addr2")
                    fl_st_city = oRs("fl_st_city")




					fl_firstdrop = oRs("Fl_Firstdrop")
					fl_sf_comment = oRs("fl_sf_comment")
					JobStatus = oRs("fh_status")
					Priority = oRs("fh_priority")
					fh_bt_id=Trim(cStr(oRs("fh_bt_id")))
					DueTime=oRs("fl_st_rta")
					If MaterialType="300 mm Waf" or MaterialType="Foup/Fosby" then
						MaterialSymbol="*"
						else
						MaterialSymbol=""							
					End if
					If trim(FromLocation)="72" then
						'Response.Write "Got here???<BR>"
						If Priority="P0" then
							DueTime=DateAdd("n", 45, Fl_firstdrop)
							else
							DueTime=DateAdd("n", 120, Fl_firstdrop)
						End if					
					End if						
					TimeTillDue=DateDiff("n",now(),DueTime)
					'Response.Write "TimeTillDue="&TimeTillDue&"<BR>"
          
         
					xxx=xxx+1
					if xxx=1 then
					%>

						<tr>
							<!--td align="center" colspan="2">&nbsp;</td-->						
							<td align="center" nowrap><b>Job #</b></td>
							<td align="center" nowrap><b>Due In</b></td>
							<td align="center" nowrap><b>To</b></td>
							<td align="center" nowrap>
							<%if trim(fh_bt_id)<>"26" and trim(fh_bt_id)<>"80" then%>
								<b>Lots</b>
								<%else%>
								<b>Refs</b>
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
						'Response.Write "SELECT count(rf_ref) as NumberOfLots FROM FCREFS WHERE (rf_fh_id='"&JobNumber&"') and (ref_status<>'X') "
						'Response.Write "Got here????<BR>"
						Set Recordset1 = Server.CreateObject("ADODB.Recordset")
						Recordset1.ActiveConnection = DATABASE
						Recordset1.Source = "SELECT count(rf_ref) as NumberOfLots FROM FCREFS WHERE (rf_fh_id='"&JobNumber&"') AND ((ref_status<>'X') OR (ref_status is NULL)) "
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
					
					
					If Priority="P0" or Priority="XP" then 
						PriorityColor="red"
						ButtonClass="ButtonRed"
                        else
                        If Priority="P1" then
						    ButtonClass="Button1"
						    PriorityColor="blue"
						    Else
						    ButtonClass="Button1"
						    PriorityColor="black"
					End if	
                    End if
					If MaterialType="Secure Waf" OR MaterialType="secret" OR MaterialType="ITAR" then
						PriorityColor="Orange"
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
							Response.Write "<tr><td colspan='4' bgcolor='white'><img src='images/pixel.gif' height='2' width='1' border='0'></td></tr>"
							Response.Write "<tr><td colspan='4' bgcolor='gray'><img src='images/pixel.gif' height='1' width='1' border='0'></td></tr>"
							Response.Write "<tr><td colspan='4' bgcolor='white'><img src='images/pixel.gif' height='2' width='1' border='0'></td></tr>"							
							X=0
						End if
						Y=Y+1
						If Priority="P0" or Priority="XP" then
							ButtonClass="ButtonRed"
                            else
                            If Priority="P1" then
                                ButtonClass="ButtonBlue"
							    else
							    ButtonClass="Button1"
						End if
                        End if
						Select Case JobStatus
							Case "ACK","ONB","DPV", "PUO"
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
						
						<input type="hidden" name="JobNumber" value="<%=JobNumber%>" ID="Hidden1">
						<input type="hidden" name="fh_bt_id" value="<%=fh_bt_id%>" ID="Hidden14">
						<input type="submit" name="submit" value="<%=Right(JobNumber,5)%>" id="gobutton2">
						
						</font></td>
						</form>
						<%
						''''''''''''''''''''''''''''''''
											'Response.Write "%%%%FromLocation="&FromLocation&"<BR>"
											DisplayToLocation=trim(ToLocation)
											DisplayFromLocation=trim(FromLocation)
											
											'''REMOVED ON 2/21/11 If trim(VehicleID)="613" and trim(JobStatus)="ONB" AND (trim(FromLocation)="PHO" or trim(FromLocation)="CPGP" or trim(FromLocation)="TOPPAN")	then
											'''REMOVED ON 2/21/11 	DisplayToLocation="SB-HUB"
											'''REMOVED ON 2/21/11 	'Response.Write "got here!!!<br>"
											'''REMOVED ON 2/21/11 End if											
											'response.write "VehicleID="&VehicleID&"<BR>"
											'''REMOVED ON 2/21/11 If (Trim(ToLocation)="TOPPAN" or Trim(ToLocation)="CPGP" or Trim(ToLocation)="72" or Trim(ToLocation)="PHO" or Trim(ToLocation)="TISHERMANRET") and (trim(VehicleID)<>"611" AND trim(VehicleID)<>"612" AND trim(VehicleID)<>"613" AND trim(VehicleID)<>"112" AND trim(VehicleID)<>"123") then
											'''REMOVED ON 2/21/11 	DisplayToLocation="SB-HUB"
											'''REMOVED ON 2/21/11 	'response.write "Got here 1<BR>"
											'''REMOVED ON 2/21/11 End if
											'''REMOVED ON 2/21/11 If (Trim(FromLocation)="55" and (VehicleID<>"611" OR VehicleID<>"612" OR VehicleID<>"613" OR VehicleID<>"112" OR VehicleID<>"123")) or Trim(FromLocation)="72" then
											'''REMOVED ON 2/21/11 	DisplayFromLocation="SB-HUB"
											'''REMOVED ON 2/21/11 	'response.write "Got here 2<BR>"
											'''REMOVED ON 2/21/11 End if	
											'''REMOVED ON 2/21/11 If ((Trim(FromLocation)="55" OR Trim(FromLocation)="TOPPAN" OR Trim(FromLocation)="PHO" OR Trim(FromLocation)="HFABRET") AND (Trim(VehicleID)="611" or trim(VehicleID)="612" or trim(VehicleID)="613" or trim(VehicleID)="112" or trim(VehicleID)="123")) AND (trim(FromLocation)<>"TISHERMANRET" AND trim(ToLocation)<>"PHO") AND (trim(ToLocation)<>"TISHERMANRET" AND trim(FromLocation)<>"PHO") then 
											'''REMOVED ON 2/21/11 	DisplayToLocation="SB-HUB" 
											'''REMOVED ON 2/21/11 	ToName="South Building HUB"
											'''REMOVED ON 2/21/11 	'response.write "Got here 3<BR>"
											'''REMOVED ON 2/21/11 End if
											'''REMOVED ON 2/21/11 If trim(VehicleID)="613" and trim(ToLocation)="TISHERMA" or Trim(ToLocation)="TISHERMANRET" then
												'REsponse.Write "GOT HERE!<BR>"
											'''REMOVED ON 2/21/11 	DisplayFromLocation="SB-HUB"
											'''REMOVED ON 2/21/11 End if											
										''''''''''''''''''''''''''''''''''
											'Response.Write "vehicleID="&vehicleID&"<BR>"
											'Response.Write "ToLocation="&ToLocation&"<BR>"
												'''REMOVED ON 2/21/11 If trim(VehicleID)="123" and trim(ToLocation)="TISHERMA" then
												'''REMOVED ON 2/21/11 	'REsponse.Write "GOT HERE!<BR>"
												'''REMOVED ON 2/21/11 	DisplayToLocation="SB-HUB"
												'''REMOVED ON 2/21/11 End if	
											If trim(ToLocation)="80" then
												DisplayToLocation="LSP Warehouse"
											End if
											If trim(FromLocation)="80" then
												DisplayFromLocation="LSP Warehouse"
											End if	
											Select Case DisplayToLocation
												Case "D7"
													DisplayToLocation="D1"
												Case "P1"
													DisplayToLocation="D1"
											End Select														
										''''''''''''''''''''''''''''''''''											
						''''''''''''''''''''''''''''''''						
						Response.Write "<td valign='top' nowrap><font color='"&PriorityColor&"'>"&DisplayTimeTillDue&"</font></td>"
						Response.Write "<td valign='top' nowrap><font color='"&PriorityColor&"'>"%>

                             <%If trim(fl_st_name)>"" then response.write fl_st_name&"<BR>" end if%><%If trim(fl_st_building)>"" then response.write fl_st_Building&"<BR>" end if%><%If trim(fl_st_addr1)>"" then response.write fl_st_addr1&"<BR>" End if%><%If trim(fl_st_addr2)>"" then response.write fl_st_addr2&"<BR>"%><%If trim(fl_st_city)>"" then response.write fl_st_city end if%>
                          <%
                          Response.write "</font></td>"  


						Response.Write "<td valign='top' nowrap align='center'><font color='"&PriorityColor&"'>"&MaterialSymbol&NumberOfLots&MaterialSymbol&"</td></tr><tr><td>&nbsp;</td></tr>"
					If trim(fl_sf_comment)>"" then
						Response.Write "<tr><td colspan='3'>***"&fl_sf_comment&"</td></tr>"
					End if
					'End if
					x=x+1
					'Response.Write X&") "&LotNumber&"<br>"
					'TempJobNumber=JobNumber
					'TempX=X
					'Response.Write "----------------------------<br>"
					'''''''oRs.Movenext
				'''''''Loop
				Response.Write "</font>"
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
				
				
				''''''oRs.Close	
				'If CloseTable="y" then
					'Response.Write TempX
					'Response.Write "</font></td>"					
					'
					'</tr><!--/table-->
					'
					'CloseTable="n"
				'End if
				End if
'-------------------STARTS THE PICK UP	
				If TruckStatus="pickup" then			
				X=0
				Y=0
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				
				SQL = "SELECT Fl_SF_ID, Fh_ID, fh_User5, fh_bt_id, Fl_ST_ID, fl_sf_name, fl_sf_building, fl_sf_addr1, fl_sf_addr2, fl_sf_city, fl_st_name, fl_st_building, fl_st_addr1, fl_st_addr2, fl_st_city, fl_st_rta, fl_firstdrop, fl_sf_comment, FH_Status, Fh_Priority, fh_user5 FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
				SQL = SQL&" WHERE (Fl_un_ID='"&VehicleID&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND (fh_ship_dt>'"&now()-30&"')"
				'SQL = SQL&" AND ((fh_status='OPN') or (fh_status='ACC'))"
				If BillToID="48" or trim(vehicleID)="198" then
					SQL = SQL&" AND ((fh_status='PUO') or (fh_status='AC2'))"
					Else
					SQL = SQL&" AND ((fh_status='ACC')"
					'''''If VehicleID=124 then
						SQL = SQL&" OR (fh_status='ARV') OR (fh_status='AC2')"
					'''''End if	
					SQL = SQL&" )"				
				End if
				SQL = SQL&" ORDER BY fh_priority desc, fl_st_rta"
				oRs.Open SQL, DATABASE, 1, 3
				'REsponse.Write "SQL="&SQL&"<BR>"
				%>
					<!--table width="700" cellpadding="0" cellspacing="0" border="1" align="center" ID="Table2"-->
					<table width="300" cellpadding="0" cellspacing="0" border="0" align="left" ID="Table2">
                        <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>
						<tr><td align="center" colspan="4"><a href="default.asp" class="mainpagelink"><form method="post" action="DriverTruckLoad.asp" ID="Form9"><input type="submit" value="Return to Routing" id="gobutton" name="Submit1" /></form></td></tr>

                        <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>
                        <tr>
		                    <td class="FleetXRedSection" colspan="7" align="center">
			                    <%=uCase(VehicleName)%> STATUS
		                    </td>
	                    </tr>
                         <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>
                        <tr>
		                    <td class="FleetXRedSection" colspan="7" align="center">
			                    ORDERS TO BE PICKED UP
		                    </td>
	                    </tr>
                        <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>
				<%
				'If not oRs.EOF then
						'm_Logit "OrdersToBePickedUp " & DriverID, oConn
						'm_Logit "OrdersToBePickedUp " & LocationCode, oConn

				'		CloseTable="y"
				'		ELSE
				'		Response.Write "<tr><td colspan='7' align='center'>There are currently no orders waiting to be picked up.</td></tr><tr><td>&nbsp;</td></tr>"
				'End if
				'Do while not oRs.eof
''''''''''''''''''''''''''''''''''''''''''''''				
oRS.PageSize = 6
oRS.CacheSize = oRS.PageSize
intPageCount = oRS.PageCount
intRecordCount = oRS.RecordCount
If (oRS.EOF) then
	'response.write "REDIRECT#2<BR>"
	'Response.Redirect("default.asp")
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
				
				
				
				
					FromLocation = trim(oRs("Fl_SF_ID"))
					JobNumber = trim(oRs("Fh_ID"))
					MaterialType = oRs("FH_User5")
					ToLocation = trim(oRs("Fl_ST_ID"))

                    fl_sf_name = trim(oRs("fl_sf_name"))
                    fl_sf_building = trim(oRs("fl_sf_building"))
                    fl_sf_addr1 = trim(oRs("fl_sf_addr1"))
                    fl_sf_addr2 = trim(oRs("fl_sf_addr2"))
                    fl_sf_city = trim(oRs("fl_sf_city"))
                    fl_st_name = trim(oRs("fl_st_name"))
                    fl_st_building = trim(oRs("fl_st_building"))
                    fl_st_addr1 = trim(oRs("fl_st_addr1"))
                    fl_st_addr2 = trim(oRs("fl_st_addr2"))
                    fl_st_city = trim(oRs("fl_st_city"))
                    
                              




					fl_firstdrop = oRs("Fl_firstdrop")
					fl_sf_comment = oRs("fl_sf_comment")
					JobStatus = trim(oRs("fh_status"))
					Priority = trim(oRs("fh_priority"))
					fh_bt_id=Trim(cStr(oRs("fh_bt_id")))
          
          
						MaterialType = oRs("fh_user5")
						'Response.Write "materialType="&materialtype&"<BR>"
						If MaterialType="300 mm Waf" or MaterialType="Foup/Fosby" then
							MaterialSymbol="*"
							else
							MaterialSymbol=""							
						End if
						If MaterialType="Secure Waf"  OR MaterialType="secret" OR MaterialType="ITAR" then
							MaterialSymbol="!"
						End if					
					DueTime=oRs("fl_st_rta")
					If trim(FromLocation)="72" then
						'Response.Write "Got here???<BR>"
						If Priority="P0" then
							DueTime=DateAdd("n", 45, Fl_firstdrop)
							else
							DueTime=DateAdd("n", 120, Fl_firstdrop)
						End if					
					End if					
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
							<td align="center" nowrap><b>Due in</b></td>
							<td align="center" nowrap><b>From/To</b></td>
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
						''''Response.Write "SELECT count(rf_ref) as NumberOfLots FROM FCREFS WHERE (rf_fh_id='"&JobNumber&"') and (ref_status<>'X') "
						Set Recordset1 = Server.CreateObject("ADODB.Recordset")
						Recordset1.ActiveConnection = DATABASE
						Recordset1.Source = "SELECT count(rf_ref) as NumberOfLots FROM FCREFS WHERE (rf_fh_id='"&JobNumber&"')  AND ((ref_status<>'X') OR (ref_status is NULL)) "
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
					
					
					If Priority="P0" or Priority="XP" then 
						PriorityColor="red"
						ButtonClass="ButtonRed"
                        Else
                        If Priority="P1" then
						PriorityColor="blue"
						ButtonClass="ButtonBlue"
						Else
						ButtonClass="Button1"
						PriorityColor="black"
					End if
                    End if	
					If MaterialType="Secure Waf"  OR MaterialType="secret" OR MaterialType="ITAR" then
						PriorityColor="Orange"
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
							Response.Write "<tr><td colspan='4' bgcolor='white'><img src='images/pixel.gif' height='2' width='1' border='0'></td></tr>"
							Response.Write "<tr><td colspan='4' bgcolor='gray'><img src='images/pixel.gif' height='1' width='1' border='0'></td></tr>"
							Response.Write "<tr><td colspan='4' bgcolor='white'><img src='images/pixel.gif' height='2' width='1' border='0'></td></tr>"
							X=0
						End if
						Y=Y+1
						If Priority="P0" or Priority="XP" then
							ButtonClass="ButtonRed"
                            else
                            If Priority="P1" then
                                ButtonClass="ButtonBlue"
							    else
							    ButtonClass="Button1"
                        End if
						End if
						Select Case JobStatus
							Case "ACK","ONB","PUO", "ARV", "AC2"
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
								<input type="hidden" name="AcknowledgeIt" value="y" ID="Hidden2">
								<input type="hidden" name="PageStatus" value="loggedin" ID="Hidden15">
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
						<form method="post" action="DriverTracking.asp" ID="Form6">
						<%						
						Response.Write "<td valign='top' nowrap><font color='"&PriorityColor&"'>"&Y&")"%>
						
						<input type="hidden" name="JobNumber" value="<%=JobNumber%>" ID="Hidden16">
						<input type="hidden" name="fh_bt_id" value="<%=fh_bt_id%>" ID="Hidden17">
						<input type="submit" name="submit" value="<%=Right(JobNumber,5)%>" ID="gobutton2">
											 
						 </font></td>
						 </form>	
						 <%	
''''''''''''''''''''''''''''''''
					DisplayToLocation=ToLocation
					DisplayFromLocation=FromLocation
					'Response.Write "VehicleID="&trim(VehicleID)&"***<BR>"
					'''REMOVED ON 2/21/11 If (Trim(ToLocation)="CPGP" or  Trim(ToLocation)="TOPPAN" or  Trim(ToLocation)="PHO") and (Trim(VehicleID)<>"611" AND trim(VehicleID)<>"612" AND trim(VehicleID)<>"613" AND trim(VehicleID)<>"112" AND trim(VehicleID)<>"123") then
					'''REMOVED ON 2/21/11 	DisplayToLocation="SB-HUB"
					'''REMOVED ON 2/21/11 	'response.write "Got here 1xxx<BR>"
					'''REMOVED ON 2/21/11 End if
					'''REMOVED ON 2/21/11 If ((Trim(FromLocation)="55" or Trim(FromLocation)="72" or Trim(FromLocation)="TOPPAN" or Trim(FromLocation)="PHO") AND (Trim(VehicleID)<>"611" AND trim(VehicleID)<>"612" AND trim(VehicleID)<>"613" AND trim(VehicleID)<>"112" AND trim(VehicleID)<>"123")) then
					'''REMOVED ON 2/21/11 	DisplayFromLocation="SB-HUB"
					'''REMOVED ON 2/21/11 	'response.write "Got here 2xxx<BR>"
					'''REMOVED ON 2/21/11 End if	
					'''REMOVED ON 2/21/11 If ((Trim(ToLocation)="CPGP" or Trim(FromLocation)="72" or Trim(ToLocation)="TOPPAN" or Trim(ToLocation)="PHO") AND (Trim(VehicleID)="611" OR trim(VehicleID)="612" OR trim(VehicleID)="613" OR trim(VehicleID)="112" OR trim(VehicleID)="123")) then
					'''REMOVED ON 2/21/11 	DisplayFromLocation="SB-HUB"
					'''REMOVED ON 2/21/11 	'response.write "Got here 3xxx<BR>"
					'''REMOVED ON 2/21/11 End if	
					'''REMOVED ON 2/21/11 If ((Trim(FROMLocation)="CPGP" or Trim(FromLocation)="72" or Trim(FROMLocation)="TOPPAN" or Trim(FROMLocation)="PHO" or Trim(FromLocation)="HFABRET") AND (Trim(VehicleID)="611" OR trim(VehicleID)="612" OR trim(VehicleID)="613" OR trim(VehicleID)="112" OR trim(VehicleID)="123")) then
					'''REMOVED ON 2/21/11 			If (Trim(VehicleID)<>"613" and Trim(FromLocation)<>"HFABRET") AND (Trim(VehicleID)<>"612" and Trim(ToLocation)<>"HFABRET") then
					'''REMOVED ON 2/21/11 				DisplayToLocation="SB-HUB"
					'''REMOVED ON 2/21/11 				'response.write "Got here 4xxx<BR>"
					'''REMOVED ON 2/21/11 				'response.write "VehicleID="&VehicleID&"<BR>"
					'''REMOVED ON 2/21/11 			End if
					'''REMOVED ON 2/21/11 	if Trim(FromLocation)="HFABRET" and trim(VehicleID)="612" then
					'''REMOVED ON 2/21/11 		'response.write "Got here 5xxx<BR>"
					'''REMOVED ON 2/21/11 		DisplayFromLocation=FromLocation
					'''REMOVED ON 2/21/11 	End if
					'''REMOVED ON 2/21/11 	'Response.Write "Got here 1<BR>"
					'''REMOVED ON 2/21/11 End if	
					''''''''''''''''''''''''''''''''''
						'Response.Write "vehicleID="&vehicleID&"<BR>"
						'Response.Write "ToLocation="&ToLocation&"<BR>"
							'''REMOVED ON 2/21/11 If trim(VehicleID)="123" and trim(ToLocation)="TISHERMA" then
							'''REMOVED ON 2/21/11 	'REsponse.Write "GOT HERE!<BR>"
							'''REMOVED ON 2/21/11 	DisplayToLocation="SB-HUB"
							'''REMOVED ON 2/21/11 End if
							'''REMOVED ON 2/21/11 If trim(VehicleID)="613" and ((trim(ToLocation)="TISHERMA" or Trim(ToLocation)="TISHERMANRET") AND Trim(FromLocation)<>"CSSF") then
							'''REMOVED ON 2/21/11 	'REsponse.Write "GOT HERE!<BR>"
							'''REMOVED ON 2/21/11 	DisplayFromLocation="SB-HUB"
							'''REMOVED ON 2/21/11 End if												
					''''''''''''''''''''''''''''''''''					
					If DisplayFromLocation="55" then DisplayFromLocation="CPGP" end if
					If Trim(ToLocation)="80" then
						DisplayToLocation="LSP Warehouse"
						'response.Write "Got here 8<BR>"
					End if	
					If Trim(FromLocation)="80" then
						DisplayFromLocation="LSP Warehouse"
						'response.Write "Got here 8<BR>"
					End if			
            'response.write "731 JobStatus=" & JobStatus & "<br>"
            if (JobStatus = "ARV" or JobStatus = "AC2") AND VehicleID=912780 then
              DisplayFromLocation = "SRHUB"
              fl_sf_name = "SRHUB"
             end if           
''''''''''''''''''''''''''''''''						 					
						Response.Write "<td valign='top' nowrap><font color='"&PriorityColor&"'>"&DisplayTimeTillDue&"</font></td>"
						Response.Write "<td valign='top' nowrap><font color='"&PriorityColor&"'>"%>
                             <%If trim(fl_sf_name)>"" then response.write fl_sf_name&"<BR>" end if%><%If trim(fl_sf_building)>"" then response.write fl_sf_Building&"<BR>" end if%><%If trim(fl_sf_addr1)>"" then response.write fl_sf_addr1&"<BR>" End if%><%If trim(fl_sf_addr2)>"" then response.write fl_sf_addr2&"<BR>"%><%If trim(fl_sf_city)>"" then response.write fl_sf_city&"<BR>-------------------------------------<br>"%>
                             <%If trim(fl_st_name)>"" then response.write fl_st_name&"<BR>" end if%><%If trim(fl_st_building)>"" then response.write fl_st_Building&"<BR>" end if%><%If trim(fl_st_addr1)>"" then response.write fl_st_addr1&"<BR>" End if%><%If trim(fl_st_addr2)>"" then response.write fl_st_addr2&"<BR>"%><%If trim(fl_st_city)>"" then response.write fl_st_city&" "%>
                                                   
                        
                        <%
                        '&DisplayFromLocation&"<br>"&DisplayToLocation&"</font></td>"
                        Response.write "</font></td>"
						Response.Write "<td valign='top' nowrap align='center'><font color='"&PriorityColor&"'>"&MaterialSymbol&NumberOfLots&MaterialSymbol&"</td></tr>"
						
                        
   										Set Recordset1 = Server.CreateObject("ADODB.Recordset")
										Recordset1.ActiveConnection = DATABASE
										Recordset1.Source = "SELECT NumberOfPieces, rf_box FROM FCREFS WHERE (rf_fh_id='"&JobNumber&"') and ((ref_status<>'X') or (ref_status is NULL))"
										Recordset1.CursorType = 0
										Recordset1.CursorLocation = 2
										Recordset1.LockType = 1
										Recordset1.Open()
										Recordset1_numRows = 0
										if NOT Recordset1.EOF then
                                            NumberOfPieces=Recordset1("NumberOfPieces")
                                            rf_box=Recordset1("rf_box")
                                            If trim(NumberOfPieces)>"" then
                                                Response.write "<tr><td colspan='3'><b>"&numberofpieces&" "&rf_box&"</b></td></tr>"
                                            End if
											Else
											ErrorMessage="Incorrect driver ID or password"
										End if
										Recordset1.Close()
										Set Recordset1 = Nothing                          
                        
                        
                        
                        If trim(fl_sf_comment)>"" then
							Response.Write "<tr><td colspan='3'>***"&fl_sf_comment&"</td></TR>"
						End if
					'End if
					x=x+1
					'Response.Write X&") "&LotNumber&"<br>"
					TempJobNumber=JobNumber
					TempX=X
					'Response.Write "----------------------------<br>"
					''''''oRs.Movenext
				'''''''Loop
				Response.Write "</font>"
				
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
				
				
				
				''''''oRs.Close	
				'If CloseTable="y" then
					'Response.Write TempX
					'Response.Write "</font></td>"						
				'	
				'	</tr><!--/table-->
				'	
				'	CloseTable="n"
				'End if
				End if
				%>
							<tr>
								<td colspan="4">
								<table ID="Table5" width="300" align="center">
				<td width="50%" align="left" valign="top"><font face="Verdana, arial" size="1">
					<%If cInt(intPage) > 1 Then%>
						<form method="post" ID="Form10">
						<input type="submit" name="submit" value="<<Previous" ID="Submit6">
						<input type="hidden" name="pagestatus" value="<%=pagestatus%>" ID="Hidden35">
						<input type="hidden" name="truckstatus" value="<%=truckstatus%>" ID="Hidden36">						
						<input type="hidden" name="page" value="<%=intPage - 1%>" ID="Hidden31"></form>
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
						<input type="submit" name="submit" value="Next>>" ID="Submit7">
						<input type="hidden" name="pagestatus" value="<%=pagestatus%>" ID="Hidden33">
						<input type="hidden" name="truckstatus" value="<%=truckstatus%>" ID="Hidden34">
						<input type="hidden" name="page" value="<%=intPage + 1%>" ID="Hidden32"></form>
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
				<%			
			Case else
			%>
			<FORM ACTION="DriverTruckLoad.asp" method="post" name="thisForm" ID="Form8">
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
											<input type="submit" name="submit" value="submit" ID="Submit4">									
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