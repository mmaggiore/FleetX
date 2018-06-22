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
'Response.Write "VehicleID="&VehicleID&"<BR>"
'If VehicleID=14 then Response.Write "GOT HERE!!!<br>" END IF
AliasCode=Request.Form("AliasCode")
'response.write "AliasCode="&AliasCode&"<BR>"
LocationCode=Request.Form("LocationCode")
'response.write "LocationCode="&LocationCode&"<BR>"
BarCode=Request.Form("BarCode")
'response.Write "VehicleID="&VehicleID&"<BR>"
'''''''''Below is the code for ABBOTT
If VehicleID=666 and BarCode>"" then
'rESPONSE.Write "got here!!!!<br>"
	Barcode=Right(Barcode,10)
	Barcode=Left(Barcode,9)
	Barcode=0&Barcode
End if
''''''''End ABBOTT code
Response.buffer = True 
'changedirectory="../marketing/"
'PageNameText="Track and Trace"
Submit=Request.Form("Submit")
'Response.Write "Submit="&Submit&"<BR>"
'Response.Write "Barcode="&Barcode&"<BR>"
If Submit="" and Barcode="" then
	response.Redirect("DriverifabPhoneEmulator.asp")
End if
PageStatus=Request.Form("PageStatus")
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
				Response.write "SQL="&SQL&"<BR>"
				oRs.Open SQL, DATABASE, 1, 3
				If not oRs.eof then
					suid=oRs("fh_bt_id")
				End if
				oRs.Close
				
'suid=Request.QueryString("fh_bt_id")
ReferenceNumber=Replace(ReferenceNumber,"""","")
ReferenceNumber=Replace(ReferenceNumber,"'","")





IF BarCode>"" THEN
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.CursorLocation = 3
	oRs.CursorType = 3
	oRs.ActiveConnection = DATABASE	

	SQL = "SELECT rf_pkey FROM FCREFS"
	SQL = SQL&" WHERE (Rf_Fh_ID='"&JobNumber&"') and (rf_ref='"&BarCode&"')"
	'Response.write "SQL="&SQL&"<BR>"
	oRs.Open SQL, DATABASE, 1, 3
	'If oRs.eof then
	'	Response.Write "THIS IS DONE!!!<BR>"
	'	Response.Write "PageStatus="&PageStatus&"<BR>"
	'	Response.Write "JobNumber="&JobNumber&"<BR>"
	'End if
	If not oRs.eof then
		RefLineNumber=oRs("rf_pkey")
		'response.write "RefLineNumber="&RefLineNumber&"<BR>"
		'response.write "PageStatus="&PageStatus&"<BR>"
		If PageStatus="ONB" then
			ORDERSTATUS="o"
			else
			ORDERSTATUS="c"
		End if
		
		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.ConnectionTimeout = 100
		oConn.Provider = "MSDASQL"
		oConn.Open DATABASE
		' 7/26/04 KK: Added canceljob functionality to be able to update the status if cancel button is pressed.
		' 11/30/04 DEC: Changed from CAN/98 TO DEL/99 to be consistent with dispatchOffice			
			l_cSQL = "UPDATE FCREFS SET ref_status = '"&ORDERSTATUS&"' WHERE rf_pkey = '" & RefLineNumber & "'"
			'Response.write "l_cSQL="&l_cSQL&"<BR>"
			oConn.Execute(l_cSQL)
			'm_logit "AcknowledgesOnFCFGTHD " & txtJobNumber, oConn
		Set oConn=Nothing		
		
		
		
		else 
		ErrorMessage="That is not a valid item number"
		'REsponse.Write "ErrorMessage="&ErrorMessage&"<BR>"
	End if
	oRs.Close
	
	
	
	Set oRs5 = Server.CreateObject("ADODB.Recordset")
	oRs5.CursorLocation = 3
	oRs5.CursorType = 3
	oRs5.ActiveConnection = DATABASE	
	SQL5 = "SELECT rf_pkey FROM FCREFS"
	SQL5 = SQL5&" WHERE (Rf_Fh_ID='"&JobNumber&"') "
	If PageStatus="ONB" then
		SQL5 = SQL5&" and (ref_status is NULL)"
		else
		SQL5 = SQL5&" and (ref_status<>'c')"
	End if	
	'response.write "SQL="&SQL&"<BR>"
	oRs5.Open SQL5, DATABASE, 1, 3	
	If oRs5.EOF then
		If PageStatus="ONB" then
			Set oConn = Server.CreateObject("ADODB.Connection")
			oConn.ConnectionTimeout = 100
			oConn.Provider = "MSDASQL"
			oConn.Open DATABASE
				l_cSQL = "UPDATE FCFGTHD SET fh_status = 'ONB', fh_statcode=5 WHERE fh_id = '" & JobNumber & "'"
				oConn.Execute(l_cSQL)
			Set oConn=Nothing	
			Set oConn = Server.CreateObject("ADODB.Connection")
			oConn.ConnectionTimeout = 100
			oConn.Provider = "MSDASQL"
			oConn.Open DATABASE
				l_cSQL = "UPDATE FCLEGS SET fl_t_atp = '"&now()&"' WHERE fl_fh_id = '" & JobNumber & "'"
				oConn.Execute(l_cSQL)
			Set oConn=Nothing				
			Session("TempJobNumber")=JobNumber
			%>
			<form method="post" action="DriverifabPhoneEmulator.asp" name="form666" ID="Form2">
				<input type="hidden" name="LocationCode" value="<%=LocationCode%>" ID="Hidden1">
				<input type="hidden" name="DriverID" value="<%=DriverID%>" ID="Hidden2">
				<input type="hidden" name="AliasCode" value="<%=AliasCode%>" ID="Hidden17">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden8">
				<input type="hidden" name="PageStatus" value="loggedin" ID="Hidden6">
			</form>
			<SCRIPT LANGUAGE='JavaScript'> 
			document.forms.form666.submit(); 
			</SCRIPT> 			
			<%	
		End if
		If PageStatus="CLS" then
			Set oConn = Server.CreateObject("ADODB.Connection")
			oConn.ConnectionTimeout = 100
			oConn.Provider = "MSDASQL"
			oConn.Open DATABASE
				l_cSQL = "UPDATE FCFGTHD SET fh_status = 'CLS', fh_statcode=9 WHERE fh_id = '" & JobNumber & "'"
				oConn.Execute(l_cSQL)
			Set oConn=Nothing	
			Set oConn = Server.CreateObject("ADODB.Connection")
			oConn.ConnectionTimeout = 100
			oConn.Provider = "MSDASQL"
			oConn.Open DATABASE
				l_cSQL = "UPDATE FCLEGS SET fl_t_atd = '"&now()&"' WHERE fl_fh_id = '" & JobNumber & "'"
				oConn.Execute(l_cSQL)
			Set oConn=Nothing				
			Session("TempJobNumber")=JobNumber
			%>
			<form method="post" action="DriverifabPhoneEmulator.asp" name="form666" ID="Form3">
				<input type="hidden" name="LocationCode" value="<%=LocationCode%>" ID="Hidden9">
				<input type="hidden" name="DriverID" value="<%=DriverID%>" ID="Hidden10">
				<input type="hidden" name="AliasCode" value="<%=AliasCode%>" ID="Hidden11">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden12">
				<input type="hidden" name="PageStatus" value="loggedin" ID="Hidden13">
			</form>
			<SCRIPT LANGUAGE='JavaScript'> 
			document.forms.form666.submit(); 
			</SCRIPT> 			
			<%		
		End if
	End if	
	oRs5.Close
	Set oRs5=NOTHING
	
	
	
	
	
END IF

'''''''''''''''REMOVED WHOLE USES LOTS THING'''''''''''''''''''''''
'Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
'	RSEVENTS2.CursorLocation = 3
'	RSEVENTS2.CursorType = 3
'	RSEVENTS2.ActiveConnection = Database
'	l_csql = "SELECT bt_fhu5req FROM fcbillto WHERE (bt_id='"&suid&"')"
	'response.write("Query:" & l_cSQL)
'	RSEVENTS2.Open l_cSQL, Database, 1, 3
'	If not RSEVENTS2.EOF then
'		UsesLots=RSEVENTS2("bt_fhu5req")
'		Else
'		ErrorMessage="You must log out and log back in."	
'	End if
'	RSEVENTS2.close
'Set RSEVENTS2 = Nothing
'Response.write "SortBy="&SortBy&"*<BR>"
%>
<body leftMargin=0 TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" OnLoad=document.Form1.BarCode.focus()>
	<TABLE WIDTH="300" border="0" bordercolor="red" cellpadding="0" cellspacing="5" align="left" ID="Table1">
		<tr><td align="center"><input type="button" value="Back" name="ClickBack" onclick=(history.back()) ID="Button1"></td></tr>
		<tr><td align="left">
		<%
		If Submit>"" then
		%>
			<table cellpadding="3" cellspacing="0" width="300" border="1" align="left" ID="Table5">
				<tr>
					<%
						ColspanNumber="8"
					%>
					<td  class="mainpagetextboldcenter" nowrap>
						Job #<%=JobNumber%>
					</td>				
																				
				</tr>
				<form name="Form1" id="Form1" method="post">
					<tr><td class="mainpagetextboldcenter" nowrap>Scan Barcode: <input type="text" name="BarCode" onBlur="form.submit()" ID="Text1"></td></tr>		
					<input type="hidden" name="Scanned" value="y" ID="Hidden3">
					<input type="hidden" name="PageStatus" value="<%=PageStatus%>" ID="Hidden4">
					<input type="hidden" name="txtcaller" value="<%=VehicleID%>" ID="Hidden5">
					<input type="hidden" name="txtstation" value="<%=FromLocation%>" ID="Hidden7">
					<input type="hidden" name="txtjobnumber" value="<%=jobnumber%>" ID="Hidden14">
					<input type="hidden" name="VehicleID" value="<%=VehicleID%>" ID="Hidden25">
					<input type="hidden" name="LocationCode" value="<%=FromLocation%>" ID="Hidden26">
					<input type="hidden" name="jobnumber" value="<%=jobnumber%>" ID="Hidden27">	
					<input type="hidden" name="AliasCode" value="<%=AliasCode%>" ID="Hidden15">				
				</form>
				<%If ErrorMessage>"" Then%>
				<tr><td class="ErrorMessageBoldCenter"><%=ErrorMessage%></td></tr>
				<%End if%>
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
			Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				RSEVENTS2.CursorLocation = 3
				RSEVENTS2.CursorType = 3
				RSEVENTS2.ActiveConnection = Database
				'Response.write "GOT HERE #2!<BR>"
				'Response.write "optJobSel="&optJobSel&"<BR>"
				'If USESLOTS=TRUE then
					l_csql = "SELECT Distinct(fcfgthd.fh_id), fclegs.fl_sf_rta, fcfgthd.fh_status, fcfgthd.fh_ship_dt, fclegs.fl_sf_id, fclegs.fl_sf_name, fclegs.fl_st_id, fclegs.fl_st_name, fclegs.fl_t_atp, fclegs.fl_t_atd, fclegs.fl_pod, fcfgthd.fh_custpo, fcfgthd.fh_priority, fcrefs.rf_ref FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN fcrefs ON fclegs.fl_fh_id = fcrefs.rf_fh_id "	
					'else
					'l_csql = "SELECT Distinct(fcfgthd.fh_id), fclegs.fl_sf_rta, fcfgthd.fh_status, fcfgthd.fh_ship_dt, fclegs.fl_sf_id, fclegs.fl_sf_name, fclegs.fl_st_id, fclegs.fl_st_name, fclegs.fl_t_atp, fclegs.fl_t_atd, fclegs.fl_pod, fcfgthd.fh_custpo, fcfgthd.fh_priority FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id "	
				'End if				
				l_csql = l_csql&" WHERE (fh_bt_id>'') "
						If PageStatus="ONB" then
							l_csql = L_csql& "AND (ref_status is NULL) "
						end if
						If PageStatus="CLS" then
							l_csql = L_csql& "AND (ref_status='o') "
						End if
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
							'If UsesLots=FALSE then
								'l_csql = L_csql& "AND (fl_sf_rta>='"&DateSent&"') AND (fl_sf_rta<'"&DayAfter&"') "
								'else
								l_csql = L_csql& "AND (fh_ship_dt>='"&DateSent&"') AND (fh_ship_dt<'"&DayAfter&"') "
							'End if
						End if
						If ToLocation>"" then
							l_csql = L_csql& "AND (fl_st_id='"&ToLocation&"') "
						End if
						If FromLocation>"" then
							l_csql = L_csql& "AND (fl_sf_id='"&FromLocation&"') "
						End if
						'If USESLOTS=TRUE then
							SortBy="rf_ref"
						'End if
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
					fl_sf_rta=RSEVENTS2("fl_sf_rta")
					fh_status=RSEVENTS2("fh_status")
					fh_ship_dt=RSEVENTS2("fh_ship_dt")
					fl_sf_id=RSEVENTS2("fl_sf_id")
					fl_st_id=RSEVENTS2("fl_st_id")					
					fl_sf_name=RSEVENTS2("fl_sf_name")
					fl_st_name=RSEVENTS2("fl_st_name")
					fl_t_atp=RSEVENTS2("fl_t_atp")
					fl_t_atd=RSEVENTS2("fl_t_atd")
					fl_pod=RSEVENTS2("fl_pod")
					fh_custpo=RSEVENTS2("fh_custpo")
					fh_priority=RSEVENTS2("fh_priority")
					fl_sf_rta=RSEVENTS2("fl_sf_rta")
					'If USESLOTS=TRUE then
						rf_ref=RSEVENTS2("rf_ref")
					'End if
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

				
				<tr>
					<%
					X=X+1
					%>
						<td class="MainPageTextSmaller" nowrap valign="top" bgcolor="<%=colorset(i mod numcolors)%>">
							<%=X%>)&nbsp;<%=rf_ref%>
						</td>				
				</tr>

<%
				i=i+1
				RSEVENTS2.movenext
				LOOP
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
	%>
			<!--
			<tr><td align="center" class="miniheader" colspan="<%=ColspanNumber%>"><%=ErrorMessage%></td></tr>
			-->
			</table>	
			<tr><td>&nbsp;</td></tr>
			<tr><td>&nbsp;</td></tr>
			<tr><td>&nbsp;</td></tr>			

	</Table>
<%end if%>
	</td></tr>
</body>
</html>
