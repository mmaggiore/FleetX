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
			'Response.Write "XXXXDisplayP0="&DisplayP0&"<BR>"
			'Response.Write "VehicleID="&VehicleID&"<BR>"
			If DisplayP0<>"xyx" then
			''''''''''''''''''''''''START OF IN VEHICLE P0 ONLY'''''''''''''''''''''''''''''''''''''''''''
			'response.write "GOT HERE 1<BR>"
			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.CursorLocation = 3
			oRs.CursorType = 3
			oRs.ActiveConnection = DATABASE	
			SQL = "SELECT distinct(Fl_ST_ID), Fh_ID, fh_user5, Fl_SF_ID, fl_st_rta, fl_firstdrop, convert(varchar(150), fl_sf_comment) as fl_sf_comment, fh_bt_id, FH_Status, Fh_Priority FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
			SQL = SQL&" WHERE (Fl_dr_ID='"&trim(VehicleID)&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND ((fh_user5='Secure Waf') OR (fh_user5='secret') OR (Fh_Priority='P0'))  AND (fh_ship_dt>'"&now()-30&"')"
			'''''If VehicleID=124 then
				'SQL = SQL&" AND ((((fh_status='DPV') AND (fl_st_id<>'CPGP')))"
				SQL = SQL&" AND ((((fh_status='DPV')))"
				'''''else
				If trim(vehicleID)="199" then
					SQL = SQL&" OR ((fh_status='ONB'))) and (fl_rt_type<>'out')"
					else
					SQL = SQL&" OR ((fh_status='PUO') or (fh_status='ONB')))"
				End if
			'''''End if			
			SQL = SQL&" ORDER BY fh_user5, fl_st_rta, fl_st_id"
			if mark="y" then
				response.write "in vehicle SQL="&SQL&"<BR>"
			end if
			''''''''''''''''''''''''''''''''''''''''''''
			'response.write "***in vehicle SQL="&SQL&"<BR><br>"
			''''''''''''''''''''''''''''''''''''''''''''
			oRs.Open SQL, DATABASE, 1, 3
			If not oRs.EOF then
					CloseTable="y"
					
					ELSE
					'Response.Write "<tr><td colspan='13' align='center'>There are currently no orders in the vehicle.</td></tr><tr><td>&nbsp;</td></tr>"
			End if
			Do while not oRs.eof
				WereP0s="y"
				X=X+1
				MaterialType = oRs("fh_user5")
				'Response.Write "zzzMaterialType="&MaterialType&"<BR>"
				FromLocation = oRs("Fl_SF_ID")
				JobNumber = oRs("Fh_ID")
				ToLocation = oRs("Fl_ST_ID")
				Fl_firstdrop = oRs("Fl_firstdrop")
				fl_sf_comment = oRs("fl_sf_comment")
				'Response.Write "fl_sf_comment="&fl_sf_comment&"<BR>"
				'Response.Write "JobNumber="&JobNumber&"<BR>"
				JobStatus = trim(oRs("fh_status"))
				'Response.Write "JobStatus="&JobStatus&"<BR>"
				Priority = oRs("fh_priority")
				If FColor2="" and Priority="P1" then
					FColor2="purple"
					else
					If Priority="P0" then
						FColor2="red"
						else 
						FColor2="black"
					End if
				End if
				If MaterialType="Secure Waf" or MaterialType="secret" then
					FColor2="Orange"
				End if
				fh_bt_id=Trim(cStr(oRs("fh_bt_id")))
				DueTime=oRs("fl_st_rta")
				'Response.Write "from location="&fromlocation&"<br>"
				If trim(FromLocation)="55xx" or trim(FromLocation)="72" then
					'Response.Write "Got here<BR>"
					If Priority="P0" then
						DueTime=DateAdd("n", 45, Fl_firstdrop)
						else
						DueTime=DateAdd("n", 120, Fl_firstdrop)
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
				If TempToLocation<>ToLocation then
					X=1
					for MMM=1 to M
						'Response.Write "Got here/FromLocation="&FromLocation&"/XFROM="&FromLocation"
						If ToLocation=ListOfToM(MMM) then
							DontShowM="y"
							'Response.Write "Dont Show!<BR>"
						End if
					Next
					M=M+1
					'ListOfTo(M)=ToLocation
					ListOfToM(M)=ToLocation						
					'Response.Write "*********************GOT HERE!<BR>"
					DisplayToLocation=trim(ToLocation)
					
					If trim(VehicleID)="613" and trim(JobStatus)="ONB" AND (trim(FromLocation)="PHO" or trim(FromLocation)="CPGP" or trim(FromLocation)="TOPPAN")	then
						DisplayToLocation="SB-HUB"
						'Response.Write "got here!!!<br>"
					End if					
						
					If Trim(ToLocation)="55" or Trim(ToLocation)="CPGP" or Trim(ToLocation)="72" or Trim(ToLocation)="TOPPAN" or Trim(ToLocation)="PHO" then
						DisplayToLocation="SB-HUB"
						'response.write "Got here 1<br>"
					End if
					
					If Trim(ToLocation)="TISHERMANRET" and VehicleID="212" then
						DisplayToLocation="SB-HUB"
					End if
					If Trim(FromLocation)="55" or Trim(FromLocation)="CPGP" or Trim(FromLocation)="72" or (Trim(FromLocation)="TOPPAN" or (Trim(FromLocation)="PHO") or (Trim(FromLocation)="HFABRET")) then
						DisplayFromLocation="SB-HUB"
						'response.write "Got here 2<br>"
						'If 
					End if
					'Response.Write "FromLocation="&FromLocation&"<BR>"
					If (FromLocation="55" OR FromLocation="TOPPAN" OR FromLocation="PHO" OR FromLocation="TISHERMANRET") AND (trim(vehicleID)="611" or trim(VehicleID)="612" or trim(VehicleID)="613" or trim(VehicleID)="112" or trim(VehicleID)="123") then
						DisplayToLocation="SB-HUB"
						'response.write "Got here 3<br>"
					End if
					'Response.Write "FromLocation="&FromLocation&"<BR>"	
					If trim(FromLocation)="80" then
						DisplayFromLocation="LSP Warehouse"
					End if
					If trim(ToLocation)="80" then
						DisplayToLocation="LSP Warehouse"
					End if					
					''''''''''''''''''''''''''''''''''
						'Response.Write "vehicleID="&vehicleID&"<BR>"
						'Response.Write "ToLocation="&ToLocation&"<BR>"
							If trim(VehicleID)="123" and (trim(ToLocation)="TISHERMA" OR trim(ToLocation)="TISHERMANRET") then
								'REsponse.Write "GOT HERE!<BR>"
								DisplayToLocation="SB-HUB"
							End if	
					Select Case DisplayToLocation
						Case "D7"
							DisplayToLocation="D1"
						Case "P1"
							DisplayToLocation="D1"
					End Select
											
					''''''''''''''''''''''''''''''''''
					DisplayDisplayTimeTillDue=DisplayTimeTillDue
					'Response.Write "JobNumber="&JobNumber&"<BR>"
					'Response.Write "DisplayTimeTillDue="&DisplayTimeTillDue&"<BR>"
					'REsponse.Write "***********************<BR>"
					Set Recordset1 = Server.CreateObject("ADODB.Recordset")
					Recordset1.ActiveConnection = DATABASE
					Recordset1.Source = "SELECT count(fh_id) as NumberOfJobs FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id WHERE ((Fh_Priority='P0') or (fh_user5='Secure Waf') or (fh_user5='secret')) AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND (Fl_dr_ID='"&VehicleID&"') AND (fh_ship_dt>'"&now()-30&"') AND "
					'''''If VehicleID=124 then
						Recordset1.Source = Recordset1.Source&" ((fh_status='DPV') "
						'''''else
						Recordset1.Source = Recordset1.Source&" OR (fh_status='ONB')) "
					'''''End if
					Recordset1.Source = Recordset1.Source&"  AND (Fl_ST_ID='"&ToLocation&"')"
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
					If DontShowM<>"y" then
						If showhr2>1 then
							Response.Write "<tr><td colspan='7'><hr></td></tr>"					
						End if										
						%>
						<form method="post" action="DriverInTruck.asp" ID="Form2">
						<tr>
							<!--td align="center" colspan="2">&nbsp;</td-->						
							<td align="center" nowrap><font color="<%=FColor2%>"><%=NumberOfJobs%>lll</font></td>
							<td width="5">&nbsp;</td>
							<td align="center" nowrap><font color="<%=FColor2%>"><%=DisplayDisplayTimeTillDue%></font></td>
							<td width="5">&nbsp;</td>
							<td align="center" nowrap><font color="<%=FColor2%>"><%=DisplayToLocation%></font></td>
							<td width="5">&nbsp;</td>
							<td align="center" nowrap>
							<%if showdetails2<>"no" then
								ShowButton1="n"
								%>
								<input type="submit" value="details" ID="Submit4" NAME="Submit4">
								<input type="hidden" name="truckstatus" value="dropoff" ID="Hidden3">
							<%showdetails2="no"
							End if%>
							</td>
						</tr>
						</form>					
						<%
						if trim(fl_sf_comment)>"" then
							Response.write "<tr><td colspan='7'>***"&fl_sf_comment&"</td></tr>"
						end if						
					End if
					DontShowM="n"
				End if
				TempToLocation=ToLocation
			oRs.Movenext
			Loop
			oRs.Close
			''''''''''''''''''''''''''''END OF IN VEHICLE P0'''''''''''''''''''''''''''''''''''''''''
			End if
			'response.write "in vehicle SQL="&SQL&"<BR>"
			''''''''''''''''''''''''START OF IN VEHICLE'''''''''''''''''''''''''''''''''''''''''''
			Showhr2=0
			DontShowM=""
			Showdetails2=""
			MMM=0
			M=0
			X=0
			TempToLocation=""
			'Response.Write "DID I GET HERE?!?<BR>"
			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.CursorLocation = 3
			oRs.CursorType = 3
			oRs.ActiveConnection = DATABASE	
			SQL = "SELECT distinct(Fl_ST_ID), Fh_ID, Fl_SF_ID, fl_st_rta, fl_firstdrop, convert(varchar, fl_sf_comment) as fl_sf_comment, fh_bt_id, FH_Status, Fh_Priority FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
			SQL = SQL&" WHERE (Fl_dr_ID='"&trim(VehicleID)&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND (Fh_Priority<>'P0') AND (fh_ship_dt>'"&now()-30&"')"
			'''''If VehicleID=124 then
				'SQL = SQL&" AND ((((fh_status='DPV') AND (fl_st_id<>'CPGP')))"
				SQL = SQL&" AND ((((fh_status='DPV')))"
				'''''else
				If trim(vehicleID)<>"199" then
					SQL = SQL&" OR ((fh_status='ONB')))"
					else
					SQL = SQL&" OR ((fh_status='ONB') AND  (fl_rt_type='out')) OR ((fh_status='PUO')))"
				End if
			'''''End if			
			SQL = SQL&" ORDER BY fl_st_rta, fh_priority, fl_st_id"
			if mark="y" then
				response.write "in vehicle SQL="&SQL&"<BR>"
			end if
			'''''''''''''''''''''''''''''''''''''''''
			'response.write "*****in vehicle SQL="&SQL&"<BR><BR>"
			'''''''''''''''''''''''''''''''''''''''''
			'response.Write "VehicleID="&vehicleid&"<br>"
			oRs.Open SQL, DATABASE, 1, 3
			If not oRs.EOF then
					CloseTable="y"
					ELSE
					If WereP0s<>"y" then
						Response.Write "<tr><td colspan='13' align='center'>There are currently no orders in the vehicle.</td></tr><tr><td>&nbsp;</td></tr>"
						
					End if
			End if
			Do while not oRs.eof
				'response.write "do while not?<BR>"
				X=X+1
				FromLocation = oRs("Fl_SF_ID")
				JobNumber = oRs("Fh_ID")
				ToLocation = oRs("Fl_ST_ID")
				Fl_firstdrop = oRs("Fl_firstdrop")
				fl_sf_comment = oRs("fl_sf_comment")
				'Response.Write "ToLocation="&ToLocation&"<BR>"
				'Response.Write "JobNumber="&JobNumber&"<BR>"
				JobStatus = trim(oRs("fh_status"))
				'Response.Write "JobStatus="&JobStatus&"<BR>"
				Priority = oRs("fh_priority")
				If FColor2="" and Priority="P1" then
					FColor2="purple"
					else
					If Priority="P0" then
						FColor2="red"
						else 
						FColor2="black"
					End if
				End if

				fh_bt_id=Trim(cStr(oRs("fh_bt_id")))
				DueTime=oRs("fl_st_rta")
				'Response.Write "from location="&fromlocation&"<br>"
				If trim(FromLocation)="72" then
					'Response.Write "Got here<BR>"
					If Priority="P0" then
						DueTime=DateAdd("n", 45, Fl_firstdrop)
						else
						DueTime=DateAdd("n", 120, Fl_firstdrop)
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
				If TempToLocation<>ToLocation then
					X=1
					for MMM=1 to M
						'Response.Write "Got here/FromLocation="&FromLocation&"/XFROM="&FromLocation"
						If ToLocation=ListOfToM(MMM) then
							DontShowM="y"
							'Response.Write "Dont Show!<BR>"
						End if
					Next
					M=M+1
					'ListOfTo(M)=ToLocation
					ListOfToM(M)=ToLocation	
					DisplayToLocation=trim(ToLocation)
					
					
					'Response.Write "Jobstatus="& JobStatus &"<BR>"	
					'Response.Write "FromLocation="& FromLocation &"<BR>"
					'Response.Write "VehicleID="& VehicleID &"<BR>"	
					If trim(VehicleID)="613" and trim(JobStatus)="ONB" AND (trim(FromLocation)="PHO" or trim(FromLocation)="CPGP" or trim(FromLocation)="TOPPAN")	then
						DisplayToLocation="SB-HUB"
						'Response.Write "got here!!!<br>"
					End if				
					'Response.Write "*********************GOT HERE!<BR>"
					
					
					If (Trim(ToLocation)="TOPPAN" or Trim(ToLocation)="CPGP" or Trim(ToLocation)="PHO" or Trim(ToLocation)="72" or Trim(ToLocation)="TISHERMANRET") and (trim(VehicleID)<>"611" AND trim(VehicleID)<>"612" AND trim(VehicleID)<>"613" AND trim(VehicleID)<>"112" AND trim(VehicleID)<>"123") then
						DisplayToLocation="SB-HUB"
						'response.write "Got here 11<br>"
					End if					
					''''If Trim(ToLocation)="55" or Trim(ToLocation)="CPGP" or Trim(ToLocation)="72" or (Trim(ToLocation)="TOPPAN" AND jobstatus<>"DPV") then
					''''	DisplayToLocation="SB-HUB"
					''''End if
					If Trim(FromLocation)="55" or Trim(FromLocation)="CPGP" or Trim(FromLocation)="72" or Trim(FromLocation)="TOPPAN" or Trim(FromLocation)="PHO" or (Trim(FromLocation)="HFABRET") then
						DisplayFromLocation="SB-HUB"
						'response.write "Got here 22<br>"
						If JobStatus="ONB" or JobStatus="DPV" and (trim(vehicleID)="611" or trim(vehicleID)="612" or trim(vehicleID)="613" or trim(vehicleID)="112" or trim(vehicleID)="123") then
							If trim(vehicleID)<>"613" AND trim(ToLocation)<>"HFABRET" THEN
								'Response.Write "FromLocation="&FromLocation&"***<BR>"
								'Response.Write "ToLocation="&ToLocation&"***<BR>"
								If (trim(FromLocation)<>"TISHERMANRET" AND trim(ToLocation)<>"PHO") AND (trim(ToLocation)<>"TISHERMANRET" AND trim(FromLocation)<>"PHO") then
									DisplayToLocation="SB-HUB"
								End if
								'response.write "Got here 33<br>"
							End if
						End if
					End if
					If trim(FromLocation)="80" then
						DisplayFromLocation="LSP Warehouse"
					End if
					If trim(ToLocation)="80" then
						DisplayToLocation="LSP Warehouse"
					End if					
					''''''''''''''''''''''''''''''''''
						'Response.Write "vehicleID="&vehicleID&"***<BR>"
						'Response.Write "FromLocation="&FromLocation&"***<BR>"
							If (trim(VehicleID)="123" and (trim(ToLocation)="TISHERMA" OR trim(ToLocation)="TISHERMANRET")) OR (trim(VehicleID)="613" and trim(FromLocation)="TISHERMANRET" AND trim(ToLocation)<>"PHO") then
								'REsponse.Write "GOT HERE!<BR>"
								DisplayToLocation="SB-HUB"
							End if					
					''''''''''''''''''''''''''''''''''					
					DisplayDisplayTimeTillDue=DisplayTimeTillDue
					'Response.Write "JobNumber="&JobNumber&"<BR>"
					'Response.Write "DisplayTimeTillDue="&DisplayTimeTillDue&"<BR>"
					'REsponse.Write "***********************<BR>"
					Set Recordset1 = Server.CreateObject("ADODB.Recordset")
					Recordset1.ActiveConnection = DATABASE
					Recordset1.Source = "SELECT count(fh_id) as NumberOfJobs FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id WHERE (Fl_dr_ID='"&VehicleID&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND (Fh_Priority<>'P0') AND (fh_ship_dt>'"&now()-30&"') AND "
					'''''If VehicleID=124 then
						Recordset1.Source = Recordset1.Source&" ((fh_status='DPV') "
						'''''else
									If trim(vehicleID)<>"199" then
										Recordset1.Source = Recordset1.Source&" OR (fh_status='ONB')) "
										else
										Recordset1.Source = Recordset1.Source&" OR (((fh_status='ONB') AND  (fl_rt_type='out')) OR (fh_status='PUO'))) "
									End if	
						
					'''''End if
					Recordset1.Source = Recordset1.Source&"  AND (Fl_ST_ID='"&ToLocation&"')"
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
					Select Case DisplayToLocation
						Case "D7"
							DisplayToLocation="D1"
						Case "P1"
							DisplayToLocation="D1"
					End Select						
					If DontShowM<>"y" then
						If showhr2>1 OR (showhr2=1 AND WereP0s="y") then
							Response.Write "<tr><td colspan='7'><hr></td></tr>"					
						End if										
						%>
						<form method="post" action="DriverInTruck.asp">
						<tr>
							<!--td align="center" colspan="2">&nbsp;</td-->						
							<td align="center" nowrap><font color="<%=FColor2%>"><%=NumberOfJobs%></font></td>
							<td width="5">&nbsp;</td>
							<td align="center" nowrap><font color="<%=FColor2%>"><%=DisplayDisplayTimeTillDue%></font></td>
							<td width="5">&nbsp;</td>
							<td align="center" nowrap><font color="<%=FColor2%>"><%=DisplayToLocation%></font></td>
							<td width="5">&nbsp;</td>
							<td align="center" nowrap>
							<%if showdetails2<>"no" AND showbutton1<>"n" then%>
								<input type="submit" value="details">
								<input type="hidden" name="truckstatus" value="dropoff" ID="Hidden1">
							<%showdetails2="no"
							End if%>
							</td>
						</tr>
						</form>					
						<%
						if trim(fl_sf_comment)>"" then
							Response.write "<tr><td colspan='7'>***"&fl_sf_comment&"</td></tr>"
						end if	
					End if
					DontShowM="n"
				End if
				
				TempToLocation=ToLocation
			oRs.Movenext
			Loop
			oRs.Close
			WereP0s=""
			''''''''''''''''''''''''''''END OF IN VEHICLE'''''''''''''''''''''''''''''''''''''''''
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
				<!--
				<td align="center" nowrap>
				<%
				'Response.Write "fh_bt_id="&fh_bt_id&"<BR>"
				'if trim(fh_bt_id)<>"26" then
				%>
					<b>Lots</b>
				<%
				'End if
				%>
				</td>
				-->
			</tr>
			<%
					'Response.Write "vehicleID="&vehicleID&"******<BR>"
					'Response.Write "FromLocation="&FromLocation&"******<BR>"			
			If DisplayP0<>"xyx" then
			'''''''''''''''''''''''''START OF TO BE PICKED UP PO''''''''''''''''''''''''''
			'Response.Write "GOT HERE 2<BR>"
			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.CursorLocation = 3
			oRs.CursorType = 3
			oRs.ActiveConnection = DATABASE	
			SQL = "SELECT distinct(Fl_SF_ID), Fh_ID, fh_user5, Fl_ST_ID, fl_st_rta, fl_firstdrop, convert(varchar, fl_sf_comment) as fl_sf_comment, fh_bt_id, FH_Status, Fh_Priority, fh_user5 FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
			SQL = SQL&" WHERE (Fl_dr_ID='"&trim(VehicleID)&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND (Fh_Priority='P0') AND (fh_ship_dt>'"&now()-30&"')"
			If BillToID="48" or trim(vehicleID)="198" then
				SQL = SQL&" AND ((fh_status='PUO'))"
				Else
				SQL = SQL&" AND ((fh_status='ACC')"
				'''''If VehicleID=124 then
					SQL = SQL&" OR (((fh_status='ARV') OR (fh_status='AC2')) AND (fl_secacc>'1/1/1900')) "
					'response.write "HELLO?<BR>"
				'''''End if
				SQL = SQL&" )"
			End if
			
			
			SQL = SQL&" ORDER BY fh_user5, fl_st_rta, fh_priority, fl_sf_id"
			If mark="y" then
				response.write "to be picked up SQL="&SQL&"<BR>"
			end if
			'''''''''''''''''''''''''''''''''''''''''''''''''''''
			'response.write "********to be picked up SQL="&SQL&"<BR><BR>"
			'''''''''''''''''''''''''''''''''''''''''''''''''''''
			oRs.Open SQL, DATABASE, 1, 3
			If not oRs.EOF then
					CloseTable="y"
					WereP0s="y"
					ELSE
					'Response.Write "<tr><td colspan='13' align='center'>There are currently no orders to be picked up.</td></tr><tr><td>&nbsp;</td></tr>"
			End if
			Do while not oRs.eof
				XX=XX+1
				MaterialType = oRs("fh_user5")
				FromLocation = oRs("Fl_SF_ID")
				JobNumber = oRs("Fh_ID")
				ToLocation = oRs("Fl_ST_ID")
				fl_firstdrop = oRs("Fl_firstdrop")
				fl_sf_comment = oRs("fl_sf_comment")
				'Response.Write "ToLocation="&ToLocation&"<BR>"
				'Response.Write "JobNumber="&JobNumber&"<BR>"
				JobStatus = trim(oRs("fh_status"))
				Priority = oRs("fh_priority")
				'Response.Write "MaterialType1="&MaterialType&"<BR>"
				If FColor="" and Priority="P1" then
					FColor="purple"
					else
					If Priority="P0" then
						FColor="red"
						else 
						FColor="black"
					End if
				End if
				If MaterialType="Secure Waf" OR MaterialType="secret" then
					FColor="Orange"
				End if
				fh_bt_id=Trim(cStr(oRs("fh_bt_id")))
				'MaterialType = oRs("fh_user5")
				'response.Write "ID="&jobnumber&"///"&MaterialType&"<BR>"
				''If MaterialType="300 mm Waf" then
				''	MaterialSymbol="*"
					'Response.Write "GOT HERE<BR>"
					'else
					'MaterialSymbol=""
				''End if	
				''Response.Write "materialSymbol="&materialsymbol&"<BR>"			
				DueTime=oRs("fl_st_rta")
				If trim(FromLocation)="55xx" or trim(FromLocation)="72" then
					'Response.Write "Got here<BR>"
					If Priority="P0" then
						DueTime=DateAdd("n", 45, Fl_firstdrop)
						else
						DueTime=DateAdd("n", 120, Fl_firstdrop)
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
				'Response.Write "*********************GOT HERE!<BR>"
				If TempToLocation<>ToLocation OR TempFromLocation<>FromLocation then
					XX=1
					for YYY=1 to Z
						'Response.Write "Got here/FromLocation="&FromLocation&"/XFROM="&FromLocation"
						If FromLocation=ListOfFrom(YYY) and ToLocation=ListOfTo(YYY) then
							DontShow="y"
							'Response.Write "Dont Show!<BR>"
						End if
					Next
					Z=Z+1
					ListOfFrom(Z)=FromLocation
					ListOfTo(Z)=ToLocation				
					'Response.Write "*********************GOT HERE!<BR>"
					DisplayToLocation=ToLocation
					DisplayFromLocation=FromLocation
					'Response.Write "vehicleID="&vehicleID&"******<BR>"
					'Response.Write "FromLocation="&FromLocation&"******<BR>"						
					If trim(VehicleID)="212" and trim(FromLocation)="TISHERMANRET" then
						'REsponse.Write "GOT HERE!<BR>"
						DisplayFromLocation="SB-HUB"
					End if					
					If (Trim(ToLocation)="55" or Trim(ToLocation)="72" or Trim(ToLocation)="TOPPAN" or Trim(ToLocation)="PHO" or Trim(ToLocation)="TISHERMANRET") AND (trim(VehicleID)<>"611" AND trim(VehicleID)<>"612" AND trim(VehicleID)<>"613" AND trim(VehicleID)<>"112" AND trim(VehicleID)<>"123") then
						DisplayToLocation="SB-HUB"
						'response.write "Got here 111<br>"
						'response.Write "Got here 6<BR>"
					End if
					If (Trim(ToLocation)="55" or Trim(ToLocation)="72" or Trim(ToLocation)="TOPPAN" or Trim(ToLocation)="PHO") AND (trim(VehicleID)="611" or trim(VehicleID)="612" or trim(VehicleID)="613" or trim(VehicleID)="112" or trim(VehicleID)="123") then
						DisplayFromLocation="SB-HUB1"
						'response.write "Got here 222<br>"
						'response.Write "Got here 8<BR>"
					End if	
					'Response.Write "FromLocation="&FromLocation&"<BR>"				
					If Trim(FromLocation)="55" or Trim(FromLocation)="72" or Trim(FromLocation)="TOPPAN" or Trim(FromLocation)="PHO" or (Trim(FromLocation)="HFABRET") or (Trim(FromLocation)="TISHERMANRET") then
						DisplayFromLocation="SB-HUB"
						'response.write "Got here 333<br>"
								If trim(VehicleID)="611" or trim(VehicleID)="612" or trim(VehicleID)="613" or trim(VehicleID)="112" or trim(VehicleID)="123" then
									'response.Write "GOT HERE5!!!<BR>"
									DisplayFromLocation=FromLocation
									DisplayToLocation="SB-HUB"
									'response.write "Got here 444<br>"
								End if						
					End if	
					If Trim(ToLocation)="80" then
						DisplayToLocation="LSP Warehouse"
						'response.Write "Got here 8<BR>"
					End if	
					If Trim(FromLocation)="80" then
						DisplayFromLocation="LSP Warehouse"
						'response.Write "Got here 8<BR>"
					End if
																
					''''''''''''''''''''''''''''''''''
						'Response.Write "vehicleID="&vehicleID&"<BR>"
						'Response.Write "ToLocation="&ToLocation&"<BR>"
							If trim(VehicleID)="123" and (trim(ToLocation)="TISHERMA" OR trim(ToLocation)="TISHERMANRET") then
								'REsponse.Write "GOT HERE!<BR>"
								DisplayToLocation="SB-HUB"
							End if					
					''''''''''''''''''''''''''''''''''					
					DisplayDisplayTimeTillDue=DisplayTimeTillDue
					DisplayMaterialSymbol=MaterialSymbol
					'Response.Write "JobNumber="&JobNumber&"<BR>"
					'Response.Write "DisplayTimeTillDue="&DisplayTimeTillDue&"<BR>"
					'REsponse.Write "***********************<BR>"
					Set Recordset1 = Server.CreateObject("ADODB.Recordset")
					Recordset1.ActiveConnection = DATABASE
					Recordset1.Source = "SELECT count(fh_id) as NumberOfJobs FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id WHERE (Fl_dr_ID='"&VehicleID&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND (Fh_Priority='P0') AND (fh_ship_dt>'"&now()-30&"') "
					If BillToID="48" or trim(vehicleID)="198" then
						Recordset1.Source =Recordset1.Source& " AND ((fh_status='PUO') or (fh_status='AC2')) "
						else
						Recordset1.Source =Recordset1.Source& " AND ((fh_status='ACC') "
						'''''If VehicleID=124 then
							Recordset1.Source =Recordset1.Source& " OR (((fh_status='ARV') OR (fh_status='AC2')) AND (fl_secacc>'1/1/1900')) "
						'''''End if
						Recordset1.Source =Recordset1.Source& ")"
					End if
					Recordset1.Source =Recordset1.Source& " AND (Fl_ST_ID='"&ToLocation&"') AND (Fl_SF_ID='"&FromLocation&"')"
					
					'response.write "I GOT TO THIS PART!!!!<BR>"
					'response.write "Recordset1.Source="&Recordset1.Source&"<BR>"
					
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
'''''''''''''''''''''''''''''''''''''''					
					Set Recordset1 = Server.CreateObject("ADODB.Recordset")
					Recordset1.ActiveConnection = DATABASE
					Recordset1.Source = "SELECT count(fh_id) as Any300 FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id WHERE (Fl_dr_ID='"&VehicleID&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND (Fh_Priority='P0') AND ((fh_user5='300 mm Waf') OR (fh_user5='Foup/Fosby')) AND (fh_ship_dt>'"&now()-30&"') "
					If BillToID="48" or trim(vehicleID)="198" then
						Recordset1.Source =Recordset1.Source& " AND ((fh_status='PUO') OR (fh_status='AC2')) "
						else
						Recordset1.Source =Recordset1.Source& " AND ((fh_status='ACC') "
						'''''If VehicleID=124 then
							Recordset1.Source =Recordset1.Source& " OR (((fh_status='ARV') OR (fh_status='AC2')) AND (fl_secacc>'1/1/1900')) "
						'''''End if
						Recordset1.Source =Recordset1.Source& ")"
					End if
					Recordset1.Source =Recordset1.Source& " AND (Fl_ST_ID='"&ToLocation&"') AND (Fl_SF_ID='"&FromLocation&"')"
					
					
					'response.write "Recordset1.Source="&Recordset1.Source&"<BR>"
					
					
					Recordset1.CursorType = 0
					Recordset1.CursorLocation = 2
					Recordset1.LockType = 1
					Recordset1.Open()
					Recordset1_numRows = 0
					if NOT Recordset1.EOF then
						Any300=Recordset1("Any300")
					End if
					Recordset1.Close()
					Set Recordset1 = Nothing
					If Any300>0 then
						MaterialSymbol="*"
						else
						MaterialSymbol=""
					end if
					If MaterialType="Secure Waf" or MaterialType="secret" then
						MaterialSymbol="!"
					End if					
'''''''''''''''''''''''''''''''''''''''	
					Select Case DisplayToLocation
						Case "D7"
							DisplayToLocation="D1"
						Case "P1"
							DisplayToLocation="D1"
					End Select				
					showhr=showhr+1	
					If DontShow<>"y" then
					If showhr>1 then
						Response.Write "<tr><td colspan='7'><hr></td></tr>"					
					End if
						%>
						<form method="post" action="DriverInTruck.asp" ID="Form1">
						<tr>
							<!--td align="center" colspan="2">&nbsp;</td-->						
							<td align="center" nowrap valign="top"><font color="<%=FColor%>"><%=MaterialSymbol%><%=NumberOfJobs%><%=MaterialSymbol%></font></td>
							<td width="5">&nbsp;</td>
							<td align="center" nowrap valign="top"><font color="<%=FColor%>"><%=DisplayDisplayTimeTillDue%></font></td>
							<td width="5">&nbsp;</td>
							<td align="center" nowrap valign="top"><font color="<%=FColor%>"><%=DisplayFromLocation%><br><%=DisplayToLocation%></font></td>
							<td width="5">&nbsp;</td>
							<td align="center" nowrap valign="top">
							<%if showdetails<>"no" then
								ShowButton2="n"
								%>
								<input type="submit" value="details" ID="Submit2" NAME="Submit1">
								<input type="hidden" name="truckstatus" value="pickup" ID="Hidden2">
								<%
								showdetails="no"
							end if%>
							</td>					
						</tr>
						</form>
										
						<%
						if trim(fl_sf_comment)>"" then
							Response.write "<tr><td colspan='7'>***"&fl_sf_comment&"</td></tr>"
						end if						
						MaterialSymbol=""
					End if
					DontShow="n"
				End if
				TempToLocation=ToLocation
				TempFromLocation=FromLocation
			oRs.Movenext
			Loop
			oRs.Close
			'''''''''''''''''''''''''''''''''''''''END OF TO BE PICKED UP PO''''''''''''''''''''''''
			End if			
			'''''''''''''''''''''''''START OF TO BE PICKED UP''''''''''''''''''''''''''
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
			SQL = "SELECT distinct(Fl_SF_ID), Fh_ID, fh_user5, Fl_ST_ID, fl_st_rta, fl_firstdrop, convert(varchar(150), fl_sf_comment) as fl_sf_comment, fh_bt_id, FH_Status, Fh_Priority, fh_user5 FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
			SQL = SQL&" WHERE (Fl_dr_ID='"&trim(VehicleID)&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND (Fh_Priority<>'P0') AND (fh_ship_dt>'"&now()-30&"')"
			If BillToID="48" or trim(vehicleID)="198" then
				SQL = SQL&" AND ((fh_status='PUO') or (fh_status='AC2'))"
				Else
				SQL = SQL&" AND ((fh_status='ACC')"
				'''''If VehicleID=124 then
					SQL = SQL&" OR (((fh_status='ARV') OR (fh_status='AC2')) AND (fl_secacc>'1/1/1900')) "
					'response.write "HELLO?<BR>"
				'''''End if
				SQL = SQL&" )"
			End if
			
			
			SQL = SQL&" ORDER BY fl_st_rta, fh_priority, fl_sf_id"
			If mark="y" then
				response.write "to be picked up SQL="&SQL&"<BR>"
			end if
			'''''''''''''''''''''''''''''''''''''''''''
			'response.write "*****to be picked up SQL="&SQL&"<BR><BR>"
			'''''''''''''''''''''''''''''''''''''''''''
			oRs.Open SQL, DATABASE, 1, 3
			If not oRs.EOF then
					CloseTable="y"
					ELSE
					If WereP0s<>"y" then
						Response.Write "<tr><td colspan='13' align='center'>There are currently no orders to be picked up.</td></tr><tr><td>&nbsp;</td></tr>"
						
					End if
			End if
			Do while not oRs.eof
				XX=XX+1
				FromLocation = oRs("Fl_SF_ID")
				JobNumber = oRs("Fh_ID")
				MaterialType = oRs("fh_user5")
				ToLocation = oRs("Fl_ST_ID")
				fl_firstdrop = oRs("Fl_firstdrop")
				fl_sf_comment = trim(oRs("fl_sf_comment"))
				'Response.Write "fl_sf_comment="&fl_sf_comment&"<BR>"
				fh_status=oRs("Fh_Status")
				'Response.Write "ToLocation="&ToLocation&"<BR>"
				'Response.Write "JobNumber="&JobNumber&"<BR>"
				JobStatus = trim(oRs("fh_status"))
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
				If MaterialType="Secure Waf" or MaterialType="secret" then
					FColor="Orange"
				End if
				fh_bt_id=Trim(cStr(oRs("fh_bt_id")))
				
				'response.Write "ID="&jobnumber&"///"&MaterialType&"<BR>"
				''If MaterialType="300 mm Waf" then
				''	MaterialSymbol="*"
				''	Response.Write "GOT HERE<BR>"
					'else
					'MaterialSymbol=""
				''End if	
				''Response.Write "materialSymbol="&materialsymbol&"<BR>"			
				DueTime=oRs("fl_st_rta")
				If trim(FromLocation)="72" then
					'Response.Write "Got here<BR>"
					If Priority="P0" then
						DueTime=DateAdd("n", 45, Fl_firstdrop)
						else
						DueTime=DateAdd("n", 120, Fl_firstdrop)
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
				If TempToLocation<>ToLocation OR TempFromLocation<>FromLocation then
					XX=1
					for YYY=1 to Z
						'Response.Write "Got here/FromLocation="&FromLocation&"/XFROM="&FromLocation"
						If FromLocation=ListOfFrom(YYY) and ToLocation=ListOfTo(YYY) then
							DontShow="y"
							'Response.Write "Dont Show!<BR>"
						End if
					Next
					Z=Z+1
					ListOfFrom(Z)=FromLocation
					ListOfTo(Z)=ToLocation				
					'Response.Write "*********************GOT HERE!<BR>"
					DisplayToLocation=ToLocation
					DisplayFromLocation=FromLocation
					'Response.Write "vehicleID="&vehicleID&"<BR>"
					'Response.Write "FromLocation="&FromLocation&"<BR>"						
					If trim(VehicleID)="212" and trim(FromLocation)="TISHERMANRET" then
						'REsponse.Write "GOT HERE!<BR>"
						DisplayFromLocation="SB-HUB"
					End if					
					If (Trim(ToLocation)="55" or Trim(ToLocation)="72" or Trim(ToLocation)="TOPPAN" or Trim(ToLocation)="PHO" or Trim(ToLocation)="TISHERMANRET") AND (trim(VehicleID)<>"611" AND trim(VehicleID)<>"612" AND trim(VehicleID)<>"613" AND trim(VehicleID)<>"112" AND trim(VehicleID)<>"123") then
						DisplayToLocation="SB-HUB"
						'response.write "Got here 1111<br>"
						'response.Write "Got here 6<BR>"
					End if
					If (Trim(ToLocation)="55" or Trim(ToLocation)="72" or ((Trim(ToLocation)="TOPPAN" or Trim(ToLocation)="PHO") AND Trim(FromLocation)<>"TISHERMANRET")) AND (trim(VehicleID)="611" or trim(VehicleID)="612" or trim(VehicleID)="613" or trim(VehicleID)="112" or trim(VehicleID)="123") then
						DisplayFromLocation="SB-HUB"
						'response.write "Got here 2222<br>"
						'response.Write "Got here 8<BR>"
					End if	
					'Response.Write "FromLocation="&FromLocation&"<BR>"
					'Response.Write "VehicleID="&VehicleID&"<BR>"				
					If Trim(FromLocation)="55" or Trim(FromLocation)="72" or Trim(FromLocation)="TOPPAN" or Trim(FromLocation)="PHO" or (Trim(FromLocation)="HFABRET") or (Trim(FromLocation)="TISHERMANRET" AND Trim(ToLocation)<>"PHO" AND Trim(ToLocation)<>"TOPPAN" AND Trim(ToLocation)<>"CPGP") then
						'response.write "Got here 3333<br>"
						DisplayFromLocation="SB-HUB"
								If (trim(VehicleID)="611" or trim(VehicleID)="612" or trim(VehicleID)="613" or trim(VehicleID)="112" or trim(VehicleID)="123") and (Trim(ToLocation)<>"PHO" and (Trim(ToLocation)<>"TOPPAN") and (Trim(ToLocation)<>"CPGP")) then
									'response.write "Got here 4444<br>"
									'response.Write "GOT HERE5!!!<BR>"
									'Response.Write "FromLocation="&FromLocation&"<BR>"
									If trim(ToLocation)<>"HFABRET" then
										DisplayFromLocation=FromLocation
										DisplayToLocation="SB-HUB"
									End if
								End if						
					End if
					If Trim(VehicleID)="613" AND trim(ToLocation)="TISHERMANRET" AND (Trim(FromLocation)="TOPPAN" or Trim(FromLocation)="CPGP") then
						DisplayToLocation=ToLocation
					End if
					If Trim(ToLocation)="TISHERMANRET" and trim(FromLocation)="PHO" then
						DisplaytoLocation=ToLocation
					end if
					If (Trim(FromLocation)="TISHERMANRET" AND Trim(ToLocation)="PHO" OR Trim(ToLocation)="TOPPAN" OR Trim(ToLocation)="CPGP") and trim(VehicleID)="611" then
						'response.write "Got here 3333<br>"
						DisplayFromLocation="SB-HUB"
					End if					
					If Trim(FromLocation)="TISHERMANRET" and (Trim(ToLocation)="TOPPAN" Or Trim(ToLocation)="CPGP") and (trim(VehicleID)="612" or trim(VehicleID)="613") then
						DisplayFromLocation=FromLocation
						DisplayToLocation="SB-HUB"
					End if
					If Trim(ToLocation)="80" then
						DisplayToLocation="LSP Warehouse"
						'response.Write "Got here 8<BR>"
					End if	
					If Trim(FromLocation)="80" then
						DisplayFromLocation="LSP Warehouse"
						'response.Write "Got here 8<BR>"
					End if						
					''''''''''''''''''''''''''''''''''
						'Response.Write "vehicleID="&vehicleID&"<BR>"
						'Response.Write "ToLocation="&ToLocation&"<BR>"
							If trim(VehicleID)="123" and (trim(ToLocation)="TISHERMA" OR trim(ToLocation)="TISHERMANRET") then
								'REsponse.Write "GOT HERE!<BR>"
								DisplayToLocation="SB-HUB"
							End if					
					''''''''''''''''''''''''''''''''''					
					DisplayDisplayTimeTillDue=DisplayTimeTillDue
					DisplayMaterialSymbol=MaterialSymbol
					'Response.Write "JobNumber="&JobNumber&"<BR>"
					'Response.Write "DisplayTimeTillDue="&DisplayTimeTillDue&"<BR>"
					'REsponse.Write "***********************<BR>"
					Set Recordset1 = Server.CreateObject("ADODB.Recordset")
					Recordset1.ActiveConnection = DATABASE
					Recordset1.Source = "SELECT count(fh_id) as NumberOfJobs FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id WHERE (Fl_dr_ID='"&VehicleID&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND (Fh_Priority<>'P0') AND (fh_ship_dt>'"&now()-30&"') "
					If BillToID="48" or trim(vehicleID)="198" then
						Recordset1.Source =Recordset1.Source& " AND ((fh_status='PUO') or (fh_status='AC2')) "
						else
						Recordset1.Source =Recordset1.Source& " AND ((fh_status='ACC') "
						'''''If VehicleID=124 then
							Recordset1.Source =Recordset1.Source& " OR (((fh_status='ARV') OR (fh_status='AC2')) AND (fl_secacc>'1/1/1900')) "
						'''''End if
						Recordset1.Source =Recordset1.Source& ")"
					End if
					Recordset1.Source =Recordset1.Source& " AND (Fl_ST_ID='"&ToLocation&"') AND (Fl_SF_ID='"&FromLocation&"')"
					'Response.Write "Recordset1.Source="&Recordset1.Source&"<BR>"
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
'''''''''''''''''''''''''''''''''''''''					
					Set Recordset1 = Server.CreateObject("ADODB.Recordset")
					Recordset1.ActiveConnection = DATABASE
					Recordset1.Source = "SELECT count(fh_id) as Any300 FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id WHERE (Fl_dr_ID='"&VehicleID&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND (Fh_Priority<>'P0') AND ((fh_user5='300 mm Waf') OR (fh_user5='Foup/Fosby')) AND (fh_ship_dt>'"&now()-30&"') "
					If BillToID="48" or trim(vehicleID)="198" then
						Recordset1.Source =Recordset1.Source& " AND (fh_status='PUO') "
						else
						Recordset1.Source =Recordset1.Source& " AND ((fh_status='ACC') "
						'''''If VehicleID=124 then
							Recordset1.Source =Recordset1.Source& " OR (((fh_status='ARV') OR (fh_status='AC2')) AND (fl_secacc>'1/1/1900')) "
						'''''End if
						Recordset1.Source =Recordset1.Source& ")"
					End if
					Recordset1.Source =Recordset1.Source& " AND (Fl_ST_ID='"&ToLocation&"') AND (Fl_SF_ID='"&FromLocation&"')"
					'Response.Write "Recordset1.Source="&Recordset1.Source&"<BR>"
					Recordset1.CursorType = 0
					Recordset1.CursorLocation = 2
					Recordset1.LockType = 1
					Recordset1.Open()
					Recordset1_numRows = 0
					if NOT Recordset1.EOF then
						Any300=Recordset1("Any300")
					End if
					Recordset1.Close()
					Set Recordset1 = Nothing
					If Any300>0 then
						MaterialSymbol="*"
						else
						MaterialSymbol=""
					end if	
					If MaterialType="Secure Waf" or MaterialType="secret" then
						MaterialSymbol="!"
					End if										
'''''''''''''''''''''''''''''''''''''''					
					showhr=showhr+1	
					If DontShow<>"y" then
					'response.write "showhr="&showhr&"<BR>"
					'response.write "fh_status="&fh_status&"***<BR>"
					If showhr>1 OR (showhr=1 AND WereP0s="y") then
						Response.Write "<tr><td colspan='7'><hr></td></tr>"					
					End if
					If trim(DisplayFromLocation)="55" then DisplayFromLocation="CPGP" end if
					If trim(DisplayToLocation)="55" then DisplayToLocation="CPGP" end if
					If (Trim(ToLocation)="55" or Trim(ToLocation)="72" or Trim(ToLocation)="TOPPAN" or Trim(ToLocation)="PHO") AND fh_status="AC2" AND (trim(VehicleID)<>"611" and trim(VehicleID)<>"612" and trim(VehicleID)<>"613" and trim(VehicleID)<>"112" and trim(VehicleID)<>"123") then
						DisplayToLocation="SB-HUB"
						'Response.Write "got here 1xxx<BR>"
					End if
					If (Trim(ToLocation)="CPGP" AND trim(FromLocation)<>"TISHERMANRET") and (trim(VehicleID)="611" or trim(VehicleID)="612" or trim(VehicleID)="613" or trim(VehicleID)="112" or trim(VehicleID)="123") then
						DisplayFromLocation="SB-HUB"
						'Response.Write "got here 2<BR>"
					End if
					If (Trim(FromLocation)="55" or Trim(FromLocation)="72" or Trim(FromLocation)="TOPPAN" or Trim(FromLocation)="PHO")  and (trim(VehicleID)<>"611" AND trim(VehicleID)<>"612" AND trim(VehicleID)<>"613" AND trim(VehicleID)<>"112" AND trim(VehicleID)<>"123") then
						DisplayFromLocation="SB-HUB"
						'Response.Write "got here 3<BR>"
					End if	
					If trim(VehicleID)="613" and (trim(ToLocation)="TISHERMA" or trim(ToLocation)="TISHERMANRET") AND trim(FromLocation)<>"PHO" then
						'REsponse.Write "GOT HERE!<BR>"
						DisplayFromLocation="SB-HUB"
					End if
					Select Case DisplayToLocation
						Case "D7"
							DisplayToLocation="D1"
						Case "P1"
							DisplayToLocation="D1"
					End Select					
					'If Trim(ToLocation)="CPGP" and (trim(VehicleID)="611" or trim(VehicleID)="612") then
					'	DisplayFromLocation="SB-HUB"
					'End if					
						%>
						<form method="post" action="DriverInTruck.asp">
						<tr>
							<!--td align="center" colspan="2">&nbsp;</td-->						
							<td align="center" nowrap valign="top"><font color="<%=FColor%>"><%=MaterialSymbol%><%=NumberOfJobs%><%=MaterialSymbol%></font></td>
							<td width="5">&nbsp;</td>
							<td align="center" nowrap valign="top"><font color="<%=FColor%>"><%=DisplayDisplayTimeTillDue%></font></td>
							<td width="5">&nbsp;</td>
							<td align="center" nowrap valign="top"><font color="<%=FColor%>"><%=DisplayFromLocation%><br><%=DisplayToLocation%></font></td>
							<td width="5">&nbsp;</td>
							<td align="center" nowrap valign="top">
							<%if showdetails<>"no" AND ShowButton2<>"n" then%>
							<input type="submit" value="details" ID="Submit1" NAME="Submit1">
							<input type="hidden" name="truckstatus" value="pickup">
							<%
							showdetails="no"
							end if%>
							</td>					
						</tr>
						</form>
										
						<%
						'fl_sf_comment="lkjsf sdlfkj sdflkj sdflkj sdflkjsd sdflkjlk sdfsdsdf sdfsdfsd sdfffdf sdrewr wersdfs sdgjkhe kjhwerkjh"
						if trim(fl_sf_comment)>"" then
							Response.write "<tr><td colspan='7'>***"&trim(fl_sf_comment)&"</td></tr>"
						end if
						MaterialSymbol=""
					End if
					DontShow="n"
				End if
				TempToLocation=ToLocation
				TempFromLocation=FromLocation
			oRs.Movenext
			Loop
			oRs.Close
			WereP0s=""
			'''''''''''''''''''''''''''''''''''''''END OF TO BE PICKED UP''''''''''''''''''''''''
			'Response.Write "X="&X&"<BR>"											
			%>

			<tr><td>&nbsp;</td></tr>
			<tr><td>&nbsp;</td></tr>
			</table>				
	</BODY>
</HTML>
