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
LLL=0
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
				
				SQL = "SELECT Fl_SF_ID, Fh_ID, Fl_ST_ID, fh_bt_id, FH_Status, Fh_Priority FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
				SQL = SQL&" WHERE (Fh_ID='"&JobNumber&"')"
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
		%>
			<table cellpadding="3" cellspacing="0" width="300" border="1" align="left" ID="Table5">
				<tr><td class="mainpagetextboldcenter" nowrap colspan="2">Locations</td></tr>
		<%
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
			
			
			
			
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				RSEVENTS2.CursorLocation = 3
				RSEVENTS2.CursorType = 3
				RSEVENTS2.ActiveConnection = Database
				'Response.write "GOT HERE #2!<BR>"
				'Response.write "Database="&Database&"<BR>"
				L_CSQL = "SELECT DISTINCT(fcshipto.st_name), fcshipto.st_addr1, fcshipto.st_addr2, fcshipto.st_city, fcshipto.st_state, fcshipto.st_zip, fcshipto.st_clname, fcshipto.st_cfname, fcshipto.st_cphone, fcshipto.st_id, fcshipto.st_alias, fcshipto.st_name FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN fcshipto ON fclegs.fl_sf_id = fcshipto.st_id"
				L_CSQL = L_CSQL&" WHERE (Fl_dr_ID='"&VehicleID&"') AND (fh_ship_dt>'"&now()-30&"')"
				L_CSQL = L_CSQL&" AND ((fh_status='ACC')"
				'''''If VehicleID=124 then
					L_CSQL = L_CSQL&" OR (fh_status='OPN')"
				If trim(VehicleID)="198" then
					L_CSQL = L_CSQL&" OR (fh_status='PUO')"				
				End if
				'''''End if
				L_CSQL = L_CSQL&") ORDER BY fcshipto.st_name"

					
			'Response.write("XXXX:" & l_cSQL & "<br>")
			''''''''''''''''''''''''''''''''''''''''
				RSEVENTS2.Open l_cSQL, Database, 1, 3
				Do while not RSEVENTS2.EOF 
				'response.Write "GOT HERE!<BR>"
					fl_sf_id=RSEVENTS2("st_id")
					fl_sf_addr1=RSEVENTS2("st_addr1")
					fl_sf_addr2=RSEVENTS2("st_addr2")
					fl_sf_city=RSEVENTS2("st_city")
					fl_sf_state=RSEVENTS2("st_state")
					fl_sf_zip=RSEVENTS2("st_zip")
					fl_sf_clname=RSEVENTS2("st_clname")
					fl_sf_cfname=RSEVENTS2("st_cfname")
					fl_sf_phone=RSEVENTS2("st_cphone")
					fl_sf_alias=RSEVENTS2("st_alias")	
					fl_sf_Name=RSEVENTS2("st_Name")	
					'Response.Write "fl_sf_Name="&fl_sf_Name&"<BR>"
					'Response.Write "fl_sf_alias="&fl_sf_alias&"<BR>"	
					'Response.Write "database="&database&"<BR>"			
					If trim(fl_sf_id)="80" then
						fl_sf_name="LSP Warehouse"
					End if
			%>
				<tr>
						<td colspan="2" class="MainPageTextSmaller" nowrap valign="top" bgcolor="<%=colorset(i mod numcolors)%>">
							<FORM ACTION="DriverifabPhoneEmulator.asp" method="post" name="thisForm" ID="Form1">
											<input name="AliasCode" id="Hidden1" type="hidden" value="<%=fl_sf_Alias%>">
											<input name='VehicleID' id="Hidden2" value='<%=VehicleID%>' type="hidden">
											<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden3">
											<input type="hidden" name="RecordTheTime" value="y" ID="Hidden4">
											<input type="submit" value="<%=fl_sf_Name%>" name="submit" ID="Submit1">
											<%
											LLL=LLL+1
											Response.Write "<BR>" & fl_sf_addr1 & "<BR>"
											If trim(fl_sf_addr2)>"" then 
												Response.Write fl_sf_addr2 & "<BR>"
											End if
											Response.Write fl_sf_city & ", " & fl_sf_state & " " & fl_sf_zip & "<BR>"
											If trim(fl_sf_cfname)>"" or trim(fl_sf_clname)>"" then
												Response.Write fl_sf_cfname & " " & fl_sf_clname & "<br>"
											End if
											If trim(fl_sf_phone)>"" then
												Response.Write fl_sf_phone & "<br>"
											End if

											%>								
							</form>								
						</td>				

				</tr>
				
				<%
				i=i+1
				RSEVENTS2.movenext
				LOOP
				RSEVENTS2.close
			Set RSEVENTS2 = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''				
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''			
			Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				RSEVENTS2.CursorLocation = 3
				RSEVENTS2.CursorType = 3
				RSEVENTS2.ActiveConnection = Database
				'Response.write "GOT HERE #2!<BR>"
				'Response.write "Database="&Database&"<BR>"
				L_CSQL = "SELECT DISTINCT(fcshipto.st_name), fcshipto.st_addr1, fcshipto.st_addr2, fcshipto.st_city, fcshipto.st_state, fcshipto.st_zip, fcshipto.st_clname, fcshipto.st_cfname, fcshipto.st_cphone, fcshipto.st_id, fcshipto.st_alias, fcshipto.st_name FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN fcshipto ON fclegs.fl_st_id = fcshipto.st_id"
				L_CSQL = L_CSQL&" WHERE (Fl_dr_ID='"&VehicleID&"') AND (fh_ship_dt>'"&now()-30&"')"
				If trim(vehicleID)<>"199" then
					L_CSQL = L_CSQL&" AND ((fh_status='ONB')"
					else
					L_CSQL = L_CSQL&" AND ((((fh_status='ONB') AND  (fl_rt_type='out')) OR (fh_status='PUO'))"
				End if
				'''''If VehicleID=124 then
					L_CSQL = L_CSQL&" OR (fh_status='DPV')"
				'''''End if
				L_CSQL = L_CSQL&") ORDER BY fcshipto.st_name"

					
			'Response.write("******:" & l_cSQL)
			''''''''''''''''''''''''''''''''''''''''
				RSEVENTS2.Open l_cSQL, Database, 1, 3
				Do while not RSEVENTS2.EOF 
					'Response.Write "GOT HERE!!!!<BR>"
					fl_st_id=RSEVENTS2("st_id")
					fl_st_addr1=RSEVENTS2("st_addr1")
					fl_st_addr2=RSEVENTS2("st_addr2")
					fl_st_city=RSEVENTS2("st_city")
					fl_st_state=RSEVENTS2("st_state")
					fl_st_zip=RSEVENTS2("st_zip")
					fl_st_clname=RSEVENTS2("st_clname")
					fl_st_cfname=RSEVENTS2("st_cfname")
					fl_st_phone=RSEVENTS2("st_cphone")	
					fl_st_Alias=RSEVENTS2("st_alias")
					fl_st_Name=RSEVENTS2("st_name")	
					If trim(fl_st_id)="80" then
						fl_st_name="LSP Warehouse"
					End if				

			%>
				<tr>
						<td colspan="2" class="MainPageTextSmaller" nowrap valign="top" bgcolor="<%=colorset(i mod numcolors)%>">
							<FORM ACTION="DriverifabPhoneEmulator.asp" method="post" name="thisForm" ID="Form6">
											<input name="AliasCode" id="txtstation" type="hidden" value="<%=fl_st_Alias%>">
											<input name='VehicleID' id='VehicleID' value='<%=VehicleID%>' type="hidden">
											<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden16">
											<input type="submit" value="<%=fl_st_Name%>" name="submit">
											<%
											'Response.Write "aliascode="&aliascode&"<BR>"
											'Response.Write "fl_st_Alias="&fl_st_Alias&"<BR>"
											LLL=LLL+1
											Response.Write "<BR>" & fl_st_addr1 & "<BR>"
											If trim(fl_st_addr2)>"" then 
												Response.Write fl_st_addr2 & "<BR>"
											End if
											Response.Write fl_st_city & ", " & fl_st_state & " " & fl_st_zip & "<BR>"
											If trim(fl_st_cfname)>"" or trim(fl_st_clname)>"" then
												Response.Write fl_st_cfname & " " & fl_st_clname & "<br>"
											End if
											If trim(fl_st_phone)>"" then
												Response.Write fl_st_phone & "<br>"
											End if

											%>				
							</form>							
						</td>				

				</tr>
				
				<%
				i=i+1
				RSEVENTS2.movenext
				LOOP
				RSEVENTS2.close
			Set RSEVENTS2 = Nothing
			'Response.Write "LLL="&LLL&"<BR>"
			If LLL=0 then
			%>
			<tr><td>&nbsp;</td></tr>
			<tr><td>There are currently no drop offs or pick ups</td></tr>
			<tr><td>&nbsp;</td></tr>
			<%
			End if
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''		
	%>
			</table>	
			<tr><td>&nbsp;</td></tr>
			<tr><td>&nbsp;</td></tr>
			<tr><td>&nbsp;</td></tr>			

	</Table>
<%end if%>
	</td></tr>
</body>
</html>