<%@ Language=VBScript %>
<!-- #include file="FleetX.inc" -->
<!-- #include file="driverinfo.inc" -->	
<HTML>
	<HEAD>
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
		<link rel="stylesheet" type="text/css" href="mainStyleSheet.css">
		<%
		Dim ListOfFrom(200)
		Dim ListOfToM(200)
		Dim ListOfTo(200)
		BillToID=Request.Cookies("FleetXPhone")("sBT_ID")	
		mark=Request.QueryString("Mark")

		%>
	</HEAD>
	<BODY leftMargin=0 TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
    <!-- #include file="LogoSection.asp" -->
        
		<table width="300" cellpadding="0" cellspacing="0" border="0" align="left" ID="Table2">
            <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>
            <form method="post" action="default.asp" ID="Form8">
			<tr><td align="center" colspan="9"><a href="default.asp" class="mainpagelink"><input type="submit" value="Return to Menu" id="gobutton" name="Submit3" /></td></tr>
            </form> 
                        <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>
                        <tr>
		                    <td class="FleetXRedSection" colspan="7" align="center">
			                    <%=uCase(VehicleName)%> Bill of Lading
		                    </td>
	                    </tr>
                        <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>			
			<tr>
				<td align="center" class="purpleseparator" colspan="6"><b></b></td>
			</tr>


			<tr>
				<!--td colspan="2">&nbsp;</td-->
				<td align="center" nowrap><b>Jobs</b></td>
				<td width="5">&nbsp;</td>
				<td align="center" nowrap><b>ONB</b></td>
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
      
      'SQL = "SELECT DISTINCT fclegs.fl_st_id, fcfgthd.fh_id, fcfgthd.fh_user5, fclegs.fl_sf_id, fclegs.fl_st_rta, fclegs.fl_firstdrop, CONVERT(varchar(150), fclegs.fl_sf_comment) AS fl_sf_comment, fcfgthd.fh_bt_id, fcfgthd.fh_status, fcfgthd.fh_priority, FCJobExceptions.ExceptionID, DriverExceptionList.ExceptionDescription FROM DriverExceptionList INNER JOIN FCJobExceptions ON DriverExceptionList.ExceptionID = FCJobExceptions.ExceptionID RIGHT OUTER JOIN fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id ON FCJobExceptions.fh_id = fclegs.fl_fh_id"
      'SQL = SQL & " WHERE (Fl_un_ID=" & VehicleID & ") AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND (Fh_Priority<>'P0') AND (Fh_Priority<>'XP') AND (fh_ship_dt>@CurrentDateTime)"
      'SQL = SQL & " AND ((fh_status<>'CLS') AND fh_status<>'CAN'))"
      
			'SQL = "SELECT Distinct(Fl_SF_ID), Fh_ID, Fl_ST_ID, fl_st_rta, fl_firstdrop, fh_bt_id, FH_Status, Fh_Priority, fh_user5 FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN fcrefs ON fclegs.fl_fh_id = fcrefs.rf_fh_id"
			'SQL = SQL&" WHERE (Fl_dr_ID='"&trim(VehicleID)&"') AND (fh_ship_dt>'"&now()-30&"')"
			'SQL = SQL&" AND ((fh_status='ACC')"
			'SQL = SQL&" AND (rf_box='')"
			'SQL = SQL&" )"
      
      SQL = "SELECT DISTINCT fclegs.fl_un_id, fclegs.fl_t_atp, fcrefs.MaterialDescription, fcrefs.NumberofPieces, fcrefs.rf_box, fcrefs.weight, fcrefs.DimHeight, fcrefs.DimWidth, fcrefs.DimLength, fcfgthd.fh_id, fcfgthd.fh_user5, fclegs.fl_sf_id, fclegs.fl_st_id, fclegs.fl_st_rta, fclegs.fl_firstdrop, CONVERT(varchar(150), fclegs.fl_sf_comment) AS fl_sf_comment, fcfgthd.fh_bt_id, fcfgthd.fh_status, fcfgthd.fh_priority  "
      SQL = SQL & " FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN fcrefs ON fcfgthd.fh_id = fcrefs.rf_fh_id"
      SQL = SQL & " WHERE (fclegs.fl_un_id = '"&trim(VehicleID)&"' ) AND (fclegs.fl_leg_status = 'c' OR fclegs.fl_leg_status IS NULL) AND (fcfgthd.fh_status <> 'CLS') AND (fcfgthd.fh_status <> 'CAN')"
	
			
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
						Response.Write "<tr><td colspan='13' align='center'>There are currently no open orders</td></tr><tr><td>&nbsp;</td></tr>"
						
					End if
			End if
			Do while not oRs.eof
				XX=XX+1
				FromLocation = oRs("Fl_SF_ID")
				JobNumber = oRs("Fh_ID")
        ONB = oRs("fl_t_atp")
        MatDesc = oRs("MaterialDescription")
        NbrPieces = oRs("NumberofPieces")
        box = oRs("rf_box")
        dweight = oRs("weight")
        dlength = oRs("dimLength")
        dwidth = oRs("dimWidth")
        dheight = oRs("dimHeight")
        if len(trim(dweight)) < 1 or isNull(dweight) then
          dweight = 0
        end if
        if len(trim(dlength)) < 1 or isNull(dlength) then
          dlength = 0
        end if
        if len(trim(dwidth)) < 1 or isNull(dwidth) then
          dwidth = 0
        end if
        if len(trim(dheight)) < 1 or isNull(dheight) then
          dheight = 0
        end if
        dimsize = dweight & " lbs/" & dwidth & "w " & dheight & "h " & dlength & "l"
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

				
						%>

						<tr>
							<td align="center" nowrap valign="top"><font color="<%=FColor%>"><%=MaterialSymbol%><a href="DriverTracking.asp?JobNumber=<%=JobNumber%>&fh_bt_id=75"><%=JobNumber%></a><%=MaterialSymbol%></font></td>
							<td width="5">&nbsp;</td>
							<td align="center" nowrap valign="top"><font color="<%=FColor%>"><%=ONB%></font></td>
							<td width="5">&nbsp;</td>
							<td align="center" nowrap valign="top"><font color="<%=FColor%>"><%=DisplayFromLocation%>/<%=DisplayToLocation%></font></td>
						</tr>
										
						<tr>
							<td align="center" nowrap valign="top"><font color="<%=FColor%>"><%=dimsize%></font></td>
							<td width="5">&nbsp;</td>
							<td align="center" nowrap valign="top"><font color="<%=FColor%>"><%=MatDesc%></font></td>
							<td width="5">&nbsp;</td>
							<td align="center" nowrap valign="top"><font color="<%=FColor%>">PCS: <%=NbrPieces%>&nbsp;<%=box%></font></td>
						</tr>
            
            <tr><td colspan=6><hr></td></tr>
										
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
