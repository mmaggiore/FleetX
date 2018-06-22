<%@ LANGUAGE="VBSCRIPT"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<SCRIPT Language="Javascript" SRC="Script/Calendar1-902.js"></SCRIPT> 
	<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<TITLE><% Response.Write(D_TITLEBAR) %></TITLE>
	<!-- added the include style.css-->
<link rel="stylesheet" type="text/css" href="mainStyleSheet.css">
<!-- #include file="../v9web/include/checkstring.inc" -->
<!-- #include file="../v9web/include/custom.inc" -->
<!-- #include file="../v9web/include/ifabsettings.inc" -->
<!-- #include file="driverinfo.inc" -->	
</head>
<% 
Response.buffer = True 
changedirectory="../marketing/"
PageNameText="Track and Trace"
Submit=Request.Form("Submit")
If Submit="" then
	Submit=Request.QueryString("Submit")
End if
Submit="xx"
DateSent=Request.Form("DateSent")
'Response.write "DateSent="&DateSent&"<BR>"
'If DateSent="" then
	'DateSent=Date()
	'else
If Submit="" or DateSent="" then
	DateSent=Date()
End if
If DateSent>"" then
	DayAfter=cDate(DateSent)+1
End if
	'Response.write "DateSent="&DateSent&"<BR>"
	'Response.write "DayAfter="&DayAfter&"<BR>"
'End if
DocumentNumber=Request.Form("DocumentNumber")
If DocumentNumber="" then
	DocumentNumber=Request.QueryString("DocumentNumber")
End if
DocumentNumber=Replace(DocumentNumber,"""","")
DocumentNumber=Replace(DocumentNumber,"'","")
LotNumber=Request.Form("LotNumber")
If LotNumber="" then
	LotNumber=Request.QueryString("LotNumber")
End if
LotNumber=Replace(LotNumber,"""","")
LotNumber=Replace(LotNumber,"'","")
SortBy=Request.Form("SortBy")
ToLocation=Request.Form("ToLocation")
FromLocation=Request.Form("FromLocation")
JobNumber=Request.Form("JobNumber")
If JobNumber="" then
	JobNumber=Request.QueryString("JobNumber")
End if
ReferenceNumber=Request.Form("ReferenceNumber")
JobNumber=Replace(JobNumber,"""","")
JobNumber=Replace(JobNumber,"'","")
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				
				SQL = "SELECT Fl_SF_ID, Fh_ID, fh_User5, Fl_ST_ID, fh_bt_id, FH_Status, Fh_Priority FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
				SQL = SQL&" WHERE (Fh_ID='"&JobNumber&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL))"
				'Response.write "SQL="&SQL&"<BR>"
				oRs.Open SQL, DATABASE, 1, 3
				If not oRs.eof then
					suid=oRs("fh_bt_id")
				End if
				oRs.Close
				
'suid=Request.QueryString("fh_bt_id")
ReferenceNumber=Replace(ReferenceNumber,"""","")
ReferenceNumber=Replace(ReferenceNumber,"'","")
Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
	RSEVENTS2.CursorLocation = 3
	RSEVENTS2.CursorType = 3
	RSEVENTS2.ActiveConnection = Database
	l_csql = "SELECT bt_fhu5req FROM fcbillto WHERE (bt_id='"&suid&"')"
	'response.write("Query:" & l_cSQL)
	RSEVENTS2.Open l_cSQL, Database, 1, 3
	If not RSEVENTS2.EOF then
		UsesLots=RSEVENTS2("bt_fhu5req")
		Else
		ErrorMessage="You must log out and log back in."	
	End if
	RSEVENTS2.close
Set RSEVENTS2 = Nothing
'Response.write "SortBy="&SortBy&"*<BR>"
%>

<BODY leftMargin=0 TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
	<TABLE WIDTH="300" border="0" bordercolor="red" cellpadding="0" cellspacing="5" align="left" ID="Table1">
		<tr><td align="center"><input type="button" value="Back" name="ClickBack" onclick=(history.back()) ID="Button1"></td></tr>
		<tr><td align="left">
	
		<%
		If Submit>"" then
		
			Dim colorset, i, numcolors
			'/--- This is your array of colors to use. -------------\ 
			colorset = split("#D2D9FC,White",",")
			numcolors = ubound(colorset)+1

		
		
			Server.ScriptTimeout = 1000
			optJobSel=Request.Querystring("optJobSel")
			optJobSel=Replace(optJobSel,"""","")
			optJobSel=Replace(optJobSel,"'","")
			If ReferenceNumber>"" then optJobSel="ByRef" end if
			If JobNumber>"" then optJobSel="ByJob" end if
			'Response.write "******optJobSel="&optJobSel&"<BR>"
			Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				RSEVENTS2.CursorLocation = 3
				RSEVENTS2.CursorType = 3
				RSEVENTS2.ActiveConnection = Database
				'Response.write "GOT HERE #2!<BR>"
				'Response.write "Database="&Database&"<BR>"
				l_csql = " SELECT DISTINCT(fcfgthd.fh_id), fcfgthd.fh_bt_id, fclegs.fl_sf_rta, fclegs.fl_st_rta, fclegs.fl_firstdrop, fcfgthd.fh_status, fcfgthd.fh_ship_dt, fcfgthd.fh_user5, fclegs.fl_sf_id, "
                l_csql = l_csql&" fclegs.fl_sf_name, fclegs.fl_st_id, fclegs.fl_st_name, fclegs.fl_t_atp, fclegs.fl_t_atd, fclegs.fl_pod, fcfgthd.fh_custpo, fcfgthd.fh_priority, fcrefs.rf_ref, "
                l_csql = l_csql&" fclegs.fl_sf_addr1, fclegs.fl_sf_addr2, fclegs.fl_sf_city, fclegs.fl_sf_state, fclegs.fl_sf_country, fclegs.fl_sf_zip, fclegs.fl_sf_clname, "
                l_csql = l_csql&" fclegs.fl_sf_cfname, fclegs.fl_sf_phone, fclegs.fl_st_clname, fclegs.fl_st_cfname, fclegs.fl_st_phone, fclegs.fl_st_addr1, fclegs.fl_st_addr2, "
                l_csql = l_csql&" fclegs.fl_st_city, fclegs.fl_st_state, fclegs.fl_st_zip, fclegs.fl_st_country, convert(varchar(150), fclegs.Fl_SF_Comment) as fl_sf_comment "
				'If USESLOTS=TRUE then
					l_csql = l_csql&" FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN fcrefs ON fclegs.fl_fh_id = fcrefs.rf_fh_id "	
					'else
					'l_csql = l_csql&" FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id "	
				'End if				
				l_csql = l_csql&" WHERE (fh_bt_id>'') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND ((ref_status<>'X') OR (ref_status is NULL)) "
						If trim(ReferenceNumber)>"" then
							l_csql = L_csql& "AND (rf_ref='"&ReferenceNumber&"') "
						End if
						If trim(DocumentNumber)>"" then
							l_csql = L_csql& "AND (fh_custpo LIKE '%"&DocumentNumber&"') "
						End if
						If trim(LotNumber)>"" then
							l_csql = L_csql& "AND (rf_ref LIKE '%"&LotNumber&"') "
						End if
						If trim(JobNumber)>"" then
							l_csql = L_csql& "AND (fh_id LIKE '%"&JobNumber&"') "
						End if												
						If DateSent>"" and (JobNumber="" and DocumentNumber="" and LotNumber="") then
							If UsesLots=FALSE then
								l_csql = L_csql& "AND (fl_sf_rta>='"&DateSent&"') AND (fl_sf_rta<'"&DayAfter&"') "
								else
								l_csql = L_csql& "AND (fh_ship_dt>='"&DateSent&"') AND (fh_ship_dt<'"&DayAfter&"') "
							End if
						End if
						If ToLocation>"" then
							l_csql = L_csql& "AND (fl_st_id='"&ToLocation&"') "
						End if
						If FromLocation>"" then
							l_csql = L_csql& "AND (fl_sf_id='"&FromLocation&"') "
						End if
						If USESLOTS=TRUE then
							SortBy="rf_ref"
						End if
						If SortBy>"" then
							l_csql = L_csql& " ORDER BY "&Sortby
						End if

					
			'Response.write("Query3:" & l_cSQL)
			''''''''''''''''''''''''''''''''''''''''
				RSEVENTS2.Open l_cSQL, Database, 1, 3
				If RSEVENTS2.eof then
						ErrorMessage="No jobs were found that match your criteria."	
				End if				
				Do while not RSEVENTS2.EOF 
					fh_id=RSEVENTS2("fh_id")
					BillToID=trim(RSEVENTS2("fh_bt_id"))
					fl_sf_rta=RSEVENTS2("fl_sf_rta")
					fl_st_rta=RSEVENTS2("fl_st_rta")
					fh_status=RSEVENTS2("fh_status")
					fh_ship_dt=RSEVENTS2("fh_ship_dt")
					MaterialType = RSEVENTS2("fh_user5")
					If MaterialType="300 mm Waf" or MaterialType="Foup/Fosby" then
						MaterialSymbol="*"
						SecureWarning=""
						else
						MaterialSymbol=""
						SecureWarning=""						
					End if	
					If MaterialType="Secure Waf" then
						MaterialSymbol="!"
						SecureWarning=" - (SECURE WAFER!)"
					End if				
					fl_sf_id=RSEVENTS2("fl_sf_id")
					fl_sf_addr1=RSEVENTS2("fl_sf_addr1")
					fl_sf_addr2=RSEVENTS2("fl_sf_addr2")
					fl_sf_city=RSEVENTS2("fl_sf_city")
					fl_sf_state=RSEVENTS2("fl_sf_state")
					fl_sf_country=RSEVENTS2("fl_sf_country")
					fl_sf_zip=RSEVENTS2("fl_sf_zip")
					fl_sf_clname=RSEVENTS2("fl_sf_clname")
					fl_sf_cfname=RSEVENTS2("fl_sf_cfname")
					fl_sf_phone=RSEVENTS2("fl_sf_phone")
					
 					

					fl_st_addr1=RSEVENTS2("fl_st_addr1")
					fl_st_addr2=RSEVENTS2("fl_st_addr2")
					fl_st_city=RSEVENTS2("fl_st_city")
					fl_st_state=RSEVENTS2("fl_st_state")
					fl_st_country=RSEVENTS2("fl_st_country")
					fl_st_zip=RSEVENTS2("fl_st_zip")
					fl_st_clname=RSEVENTS2("fl_st_clname")
					fl_st_cfname=RSEVENTS2("fl_st_cfname")
					fl_st_phone=RSEVENTS2("fl_st_phone")					
					
					
					fl_st_id=RSEVENTS2("fl_st_id")					
					fl_sf_name=RSEVENTS2("fl_sf_name")
					fl_st_name=RSEVENTS2("fl_st_name")
					fl_t_atp=RSEVENTS2("fl_t_atp")
					fl_t_atd=RSEVENTS2("fl_t_atd")
					fl_firstdrop=RSEVENTS2("fl_firstdrop")
					fl_pod=RSEVENTS2("fl_pod")
					fh_custpo=RSEVENTS2("fh_custpo")
					fh_priority=RSEVENTS2("fh_priority")
					fl_sf_rta=RSEVENTS2("fl_sf_rta")
					Fl_SF_Comment=RSEVENTS2("Fl_SF_Comment")
					If USESLOTS=TRUE then
						rf_ref=RSEVENTS2("rf_ref")
					End if
			Select Case fh_status
				Case "CLS"
					Display_fh_status="Closed"
				Case "OPN"
					Display_fh_status="Open"
				Case "ACC"
					Display_fh_status="Accepted"
				Case "ONB"
					Display_fh_status="On Board"
				Case "ATD"
					Display_fh_status="At Destination"
				Case "CAN"
					Display_fh_status="Cancelled"
				Case "DEL"
					Display_fh_status="Deleted"	
				Case Else
					Display_fh_status=fh_status																			
			End Select
			if fh_ship_dt="1/1/1900" then fh_ship_dt="&nbsp;" end if
			if fl_t_atp="1/1/1900" then fl_t_atp="&nbsp;" end if
			if fl_t_atd="1/1/1900" then fl_t_atd="&nbsp;" end if
			'If ErrorMessage="" then
			%>
		<%
		'''''Response.Write "BillToID="&BillToID&"<BR>"
		Select Case BillToID
			Case "75"
				titleword="PO Number(s)"
			Case "80"
				titleword="HAWB Number(s)"				
			Case else
				titleword="Lot Number(s)"
		End select
		If x=0 then
		%>
			<table cellpadding="3" cellspacing="0" width="300" border="1" align="left" ID="Table5">
				<tr><td class="mainpagetextboldcenter" nowrap colspan="2">Job #<%=JobNumber%></td></tr>
				<tr>
				<%If UsesLots=FALSE then
						ColspanNumber="7"
				%>
					<td  class="mainpagetextboldcenter" nowrap colspan="2">
						Document Number(s)
					</td>				
					<%else
						ColspanNumber="8"
					%>
					<td  class="mainpagetextboldcenter" nowrap colspan="2">
						<%=titleword%>
					</td>				
				<%End if%>
																				
				</tr>				
	<%end if%>		
			
			
			
			
			
				<tr>
					<%
					X=X+1
					If UsesLots=FALSE then
					
					%>
						<td colspan="2" class="MainPageTextSmaller" nowrap valign="top" bgcolor="<%=colorset(i mod numcolors)%>">
							<%=fh_custpo%>
						</td>				
						<%else
						%>
						<td colspan="2" class="MainPageTextSmaller" nowrap valign="top" bgcolor="<%=colorset(i mod numcolors)%>">
							<%=X%>)&nbsp;<%=MaterialSymbol%><%=trim(rf_ref)%><%=MaterialSymbol%> <%=SecureWarning%>
						</td>				
					<%End if%>
				</tr>
				
				
				<%
				i=i+1
				RSEVENTS2.movenext
				LOOP
				RSEVENTS2.close
			Set RSEVENTS2 = Nothing
			If trim(Fl_SF_Comment)>"" then
				Response.Write "<tr><td colspan='2'>***"&trim(Fl_SF_Comment)&"</td></tr>"
			End if
			
			If BillToID="75" or BillToID="80" then
				Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
					RSEVENTS2.CursorLocation = 3
					RSEVENTS2.CursorType = 3
					RSEVENTS2.ActiveConnection = Database
					l_csql = "SELECT pieces, piecetype, skids, fc_description, Weight, DimWeight, Dimensions FROM FCRefs_Details WHERE (fh_id='"&fh_id&"')"
					'response.write("Query:" & l_cSQL)
					RSEVENTS2.Open l_cSQL, Database, 1, 3
					Do while not RSEVENTS2.EOF 
						pieces=RSEVENTS2("pieces")
						piecetype=RSEVENTS2("piecetype")
						skids=RSEVENTS2("skids")
						fc_description=RSEVENTS2("fc_description")
						RealWeight=RSEVENTS2("Weight")
						DimWeight=RSEVENTS2("DimWeight")
						Dimensions=RSEVENTS2("Dimensions")
				%>
					<tr>
						<td colspan="2" class="MainPageTextSmaller" nowrap valign="top" bgcolor="<%=colorset(i mod numcolors)%>">
							Description:&nbsp;&nbsp;<%=fc_description%><br>
							<%=PieceType%>:&nbsp;&nbsp;&nbsp;&nbsp;<%=Pieces%>&nbsp;&nbsp;&nbsp;&nbsp;Skids:&nbsp;&nbsp;&nbsp;&nbsp;<%=Skids%><br>
							Weight:&nbsp;&nbsp;&nbsp;&nbsp;<%=RealWeight%>&nbsp;&nbsp;&nbsp;&nbsp;Dim Weight:&nbsp;&nbsp;&nbsp;&nbsp;<%=DimWeight%><br>
							Dimensions:&nbsp;&nbsp;&nbsp;&nbsp;<%=Dimensions%>
						</td>					
					</tr>			
				<%
					RSEVENTS2.Movenext
					LOOP
					RSEVENTS2.close
				Set RSEVENTS2 = Nothing				
			End if
						
			
			Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				RSEVENTS2.CursorLocation = 3
				RSEVENTS2.CursorType = 3
				RSEVENTS2.ActiveConnection = Database
				l_csql = "SELECT st_name FROM fcshipto WHERE (st_id='"&fl_sf_id&"')"
				'response.write("Query:" & l_cSQL)
				RSEVENTS2.Open l_cSQL, Database, 1, 3
				If not RSEVENTS2.EOF then
					FromName=RSEVENTS2("st_name")
				End if
				RSEVENTS2.close
			Set RSEVENTS2 = Nothing	
			
			Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				RSEVENTS2.CursorLocation = 3
				RSEVENTS2.CursorType = 3
				RSEVENTS2.ActiveConnection = Database
				l_csql = "SELECT st_name FROM fcshipto WHERE (st_id='"&fl_st_id&"')"
				'response.write("Query:" & l_cSQL)
				RSEVENTS2.Open l_cSQL, Database, 1, 3
				If not RSEVENTS2.EOF then
					ToName=RSEVENTS2("st_name")
				End if
				RSEVENTS2.close
			Set RSEVENTS2 = Nothing						
			'Response.Write "i="&i&"<BR>"
			'If i>0 then
			'	If i>1 then
			'		PluralResults="s"
			'	End if
			'	Response.Write "<tr><td align='left' class='miniheader' colspan='"&ColspanNumber&"'>"&i&" Result"&PluralResults&"</td></tr>"
			'End if
			'Response.Write "ColspanNumber="&ColspanNumber&"<BR>"
					DisplayToLocation=fl_st_id
					DisplayFromLocation=fl_sf_id
					If Trim(fl_st_id)="80" then
						DisplayToLocation="LSP Warehouse"
					End if	
					If Trim(fl_sf_id)="80" then
						DisplayFromLocation="LSP Warehouse"
					End if										
					If (Trim(fl_st_id)="55" AND (Trim(VehicleID)<>"611" AND trim(VehicleID)<>"612" AND trim(VehicleID)<>"613" AND trim(VehicleID)<>"112" AND trim(VehicleID)<>"123")) or Trim(fl_st_id)="72" then
						DisplayToLocation="SB-HUB"
					End if
					'If ((Trim(fl_sf_id)="55" OR Trim(fl_sf_id)="TOPPAN") AND (Trim(VehicleID)<>"611" AND trim(VehicleID)<>"612")) or Trim(fl_st_id)="72" then
					'	DisplayFromLocation="SB-HUB"
					'End if					
					If ((Trim(fl_sf_id)="55" OR Trim(fl_sf_id)="TOPPAN" OR Trim(fl_sf_id)="PHO") AND (Trim(VehicleID)="611" or trim(VehicleID)="612" or trim(VehicleID)="613" or trim(VehicleID)="112" or trim(VehicleID)="123")) then 
						DisplayToLocation="SB-HUB" 
						ToName="South Building HUB"
					End if
					If ((Trim(fl_sf_id)="55" OR Trim(fl_sf_id)="TOPPAN" OR Trim(fl_sf_id)="PHO") AND (Trim(VehicleID)<>"611" AND trim(VehicleID)<>"612" AND trim(VehicleID)<>"613" AND trim(VehicleID)<>"112" AND trim(VehicleID)<>"123")) then 
						DisplayFromLocation="SB-HUB" 
						FromName="South Building HUB"
					End if					
					If Trim(fl_sf_id)="72" then
						DisplayFromLocation="SB-HUB"
							If Fh_Priority="P0" then
								fl_st_rta=DateAdd("n", 45, Fl_firstdrop)
								else
								fl_st_rta=DateAdd("n", 120, Fl_firstdrop)
							End if						
					End if
					''''''''''''''''''''''''''''''''''
						'Response.Write "vehicleID="&vehicleID&"<BR>"
						'Response.Write "fl_st_id="&fl_st_id&"<BR>"
							If trim(VehicleID)="123" and trim(fl_st_id)="TISHERMA" then
								'REsponse.Write "GOT HERE!<BR>"
								DisplayToLocation="SB-HUB"
							End if
							If trim(VehicleID)="613" and trim(fl_st_id)="TISHERMA" then
								'REsponse.Write "GOT HERE!<BR>"
								DisplayFromLocation="SB-HUB"
							End if												
					''''''''''''''''''''''''''''''''''								
					If trim(DisplayFromLocation)="55" then DisplayFromLocation="CPGP" end if
	%>
			<tr><td class="mainpagetextbold" width="50%" valign="top">
			From:<br>
					<%=DisplayFromLocation%><br>
					<%=FromName%><br>
					<%=fl_sf_addr1%><br>
					<%If trim(fl_sf_addr2)>"" then response.write fl_sf_addr2 & "<br>"%>
					<%=fl_sf_city%>, <%=fl_sf_state%> <%=fl_sf_zip%><br>
					<%if trim(fl_sf_cfname)>"" or trim(fl_sf_clname)>"" then%>
					<%=fl_sf_cfname%> <%=fl_sf_clname%><br>
					<%end if%>
					<%if trim(fl_sf_phone)>"" then%>
					<%=fl_sf_phone%><br>
					<%end if%>	
					<%
					StreetAddressFrom=Replace(Trim(fl_sf_addr1), " ", "+")
					LookitupFrom=trim(StreetAddressFrom)&"+"&trim(fl_sf_city)&"+"&trim(fl_sf_state)&"+"&trim(fl_sf_zip)
					%>
					<!--
					<a href="http://maps.yahoo.com/print?ard=1&v3=0&mvt=m&tp=1&q1=<%=LookitupFrom%>/" target="_blank">Yahoo Map</a>
					<a href="http://maps.google.com/maps?near=<%=LookitupFrom%>/" target="_blank">Google Map</a>		
					-->
			</td>
			<td class="mainpagetextbold" width="50%" valign="top">To: 
					<br>
					<%=DisplayToLocation%><br>
					<%=ToName%><br>
					<%=fl_st_addr1%><br>
					<%If trim(fl_st_addr2)>"" then response.write fl_st_addr2 & "<br>"%>
					<%=fl_st_city%>, <%=fl_st_state%> <%=fl_st_zip%><br>
					<%if trim(fl_st_cfname)>"" or trim(fl_st_clname)>"" then%>
					<%=fl_st_cfname%> <%=fl_st_clname%><br>
					<%end if%>
					<%if trim(fl_st_phone)>"" then%>
					<%=fl_st_phone%><br>
					<%end if%>	
					<%
					StreetAddressTo=Replace(trim(fl_st_addr1), " ", "+")
					LookitupTo=trim(StreetAddressTo)&"+"&trim(fl_st_city)&"+"&trim(fl_st_state)&"+"&trim(fl_st_zip)
					%>
					<!--
					<a href="http://maps.yahoo.com/print?ard=1&v3=0&mvt=m&tp=1&q1=<%=LookitupTo%>/" target="_blank">Yahoo Map</a>						
					-->	
			</td></tr>
			<tr><td class="mainpagetextboldcenter" colspan="2">Due by: <%=fl_st_rta%></td></tr>
			<tr><td align="center" class="miniheader" colspan="<%=ColspanNumber%>"><%=ErrorMessage%></td></tr>
			</table>	
			<tr><td>&nbsp;</td></tr>
			<tr><td>&nbsp;</td></tr>
			<tr><td>&nbsp;</td></tr>			

	</Table>
<%end if%>
	</td></tr>
</body>
</html>