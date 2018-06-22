<%@ LANGUAGE="VBSCRIPT"%>
<%
Response.buffer = True
%>
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

Dim FormBarCode(12)
Dim AllegedBarCode(12)
Dim FormJobNumber(12)
FormBarCode(1)=Request.Form("FormBarCode1")
FormBarCode(2)=Request.Form("FormBarCode2")
FormBarCode(3)=Request.Form("FormBarCode3")
FormBarCode(4)=Request.Form("FormBarCode4")
FormBarCode(5)=Request.Form("FormBarCode5")
FormBarCode(6)=Request.Form("FormBarCode6")
FormBarCode(7)=Request.Form("FormBarCode7")
FormBarCode(8)=Request.Form("FormBarCode8")
FormBarCode(9)=Request.Form("FormBarCode9")
FormBarCode(10)=Request.Form("FormBarCode10")
FormBarCode(11)=Request.Form("FormBarCode11")
FormBarCode(12)=Request.Form("FormBarCode12")

AllegedBarCode(1)=Request.Form("AllegedBarCode(1)")
AllegedBarCode(2)=Request.Form("AllegedBarCode(2)")
AllegedBarCode(3)=Request.Form("AllegedBarCode(3)")
AllegedBarCode(4)=Request.Form("AllegedBarCode(4)")
AllegedBarCode(5)=Request.Form("AllegedBarCode(5)")
AllegedBarCode(6)=Request.Form("AllegedBarCode(6)")
AllegedBarCode(7)=Request.Form("AllegedBarCode(7)")
AllegedBarCode(8)=Request.Form("AllegedBarCode(8)")
AllegedBarCode(9)=Request.Form("AllegedBarCode(9)")
AllegedBarCode(10)=Request.Form("AllegedBarCode(10)")
AllegedBarCode(11)=Request.Form("AllegedBarCode(11)")
AllegedBarCode(12)=Request.Form("AllegedBarCode(12)")

FormJobNumber(1)=Request.Form("FormJobNumber(1)")
FormJobNumber(2)=Request.Form("FormJobNumber(2)")
FormJobNumber(3)=Request.Form("FormJobNumber(3)")
FormJobNumber(4)=Request.Form("FormJobNumber(4)")
FormJobNumber(5)=Request.Form("FormJobNumber(5)")
FormJobNumber(6)=Request.Form("FormJobNumber(6)")
FormJobNumber(7)=Request.Form("FormJobNumber(7)")
FormJobNumber(8)=Request.Form("FormJobNumber(8)")
FormJobNumber(9)=Request.Form("FormJobNumber(9)")
FormJobNumber(10)=Request.Form("FormJobNumber(10)")
FormJobNumber(11)=Request.Form("FormJobNumber(11)")
FormJobNumber(12)=Request.Form("FormJobNumber(12)")

AliasCode=Request.Form("AliasCode")
LocationCode=Request.Form("LocationCode")
If UCASE(LocationCode)="EBHUB" then
	'Response.Write "I GOT HERE!"
	Hub="y"
End if
BarCode=Request.Form("BarCode")
BillToID=Request.Form("BillToID")
'Response.Write "LocationCode="&LocationCode&"<BR>"
If BillToID>"" then
	Suid=BillToID
End if

													'''''''''Below is the code for ABBOTT
													If VehicleID=666 and BarCode>"" then
													'rESPONSE.Write "got here!!!!<br>"
														Barcode=Right(Barcode,10)
														Barcode=Left(Barcode,9)
														Barcode=0&Barcode
													End if
													''''''''End ABBOTT code
 
Submit=Request.Form("Submit")
If Submit="" and Barcode="" then
	response.Redirect("DriverifabPhoneEmulator.asp")
End if
PageStatus=Request.Form("PageStatus")
If Submit="" then
	Submit=Request.QueryString("Submit")
End if
'Submit="xx"
DateSent=Request.Form("DateSent")
If Submit="" or DateSent="" then
	DateSent=Date()
End if
If DateSent>"" then
	DayAfter=cDate(DateSent)+1
End if
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
ReferenceNumber=Replace(ReferenceNumber,"""","")
ReferenceNumber=Replace(ReferenceNumber,"'","")





IF Submit="submit" THEN
	'Response.Write "GOT HERE AFTER SUBMIT<BR>"

		If PageStatus="ONB" then
			ORDERSTATUS="o"
			else
			ORDERSTATUS="c"
			if Hub="y" then
				ORDERSTATUS="H"
			End if
		End if
		

		If PageStatus="ONB" then
			For q=1 to 12
				If (trim(AllegedBarCode(q))=trim(FormBarCode(q))) AND trim(FormBarCode(q))>""  then
				TheJobNumber=trim(FormJobNumber(q))
				'TheAllegedBarCode=trim(AllegedBarCode(q))
				TheBarCode=trim(FormBarCode(q))				
					Set oConn = Server.CreateObject("ADODB.Connection")
					oConn.ConnectionTimeout = 100
					oConn.Provider = "MSDASQL"
					oConn.Open DATABASE
						l_cSQL = "UPDATE FCREFS SET ref_status = '"&ORDERSTATUS&"' WHERE (rf_fh_id = '"&TheJobNumber&"') AND (rf_ref = '" &TheBarCode& "')"
						'Response.write "UPDATE FCREFS="&l_cSQL&"<BR>"
						oConn.Execute(l_cSQL)
						m_cSQL = "UPDATE FCFGTHD SET fh_status = 'ONB', fh_statcode=5 WHERE fh_id = '" & TheJobNumber&"'"
						'response.write "UPDATE FCFGTHD="&m_cSQL&"<BR>"
						oConn.Execute(m_cSQL)
						n_cSQL = "UPDATE FCLEGS SET fl_t_atp = '"&now()&"' WHERE fl_fh_id = '" & TheJobNumber&"'"
						'response.write "UPDATE FCLEGS="&n_cSQL&"<BR>"
						oConn.Execute(n_cSQL)
					Set oConn=Nothing	
				End if			
				''''Session("TempJobNumber")=JobNumber
			Next
	
		End if
		If PageStatus="CLS" then
			Set oConn = Server.CreateObject("ADODB.Connection")
			oConn.ConnectionTimeout = 100
			oConn.Provider = "MSDASQL"
			oConn.Open DATABASE		
			For qq=1 to 12
					If (trim(AllegedBarCode(qq))=trim(FormBarCode(qq))) AND trim(FormBarCode(qq))>""  then
					TheJobNumber=trim(FormJobNumber(qq))
					TheBarCode=trim(FormBarCode(qq))				

						l_cSQL = "UPDATE FCREFS SET ref_status = '"&ORDERSTATUS&"' WHERE (rf_fh_id = '" & TheJobNumber& "') AND (rf_ref = '" & TheBarCode& "')"
						oConn.Execute(l_cSQL)
						l_cSQL = "UPDATE FCFGTHD SET fh_status = 'CLS', fh_statcode=9 WHERE fh_id = '" & TheJobNumber& "'"
						oConn.Execute(l_cSQL)
						l_cSQL = "UPDATE FCLEGS SET fl_t_atd = '"&now()&"' WHERE fl_fh_id = '" & TheJobNumber& "'"
						oConn.Execute(l_cSQL)
								
				End if
			Next
		End if
		'oConn.close
		Set oConn=Nothing	
	
END IF


%>
<body leftMargin=0 TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" OnLoad=document.Form1.FormBarCode1.focus()>
	<TABLE WIDTH="300" border="0" bordercolor="red" cellpadding="0" cellspacing="5" align="left" ID="Table1">
		<tr><td align="center" colspan="3"><form method="post" action="DriverIfabPhoneEmulator.asp" ID="Form7"><input type="hidden" name="Aliascode" value="<%=LocationCode%>"><input type="submit" value="Return to Drop Off/Pick Up" ID="Submit1" NAME="Submit1"></form></td></tr>
		<tr><td align="left">
		<%
		If Submit>"" then
		'Response.Write "GOT HERE!<BR>"
		%>
			<table cellpadding="3" cellspacing="0" width="300" border="1" align="left" ID="Table5">
				<tr>
					<%
						ColspanNumber="8"
					%>
					<td  class="mainpagetextboldcenter" nowrap colspan="2">
						<%=PageStatus%>
					</td>				
																				
				</tr>

				<%If ErrorMessage>"" Then%>
				<tr><td class="ErrorMessageBoldCenter"><%=ErrorMessage%></td></tr>
				<%End if%>
				<form name="Form1" id="Form1" method="post">
					<input type="hidden" name="Scanned" value="y" ID="Hidden3">
					<input type="hidden" name="PageStatus" value="<%=PageStatus%>" ID="Hidden4">
					<input type="hidden" name="txtcaller" value="<%=VehicleID%>" ID="Hidden5">
					<input type="hidden" name="txtstation" value="<%=FromLocation%>" ID="Hidden7">
					<input type="hidden" name="txtjobnumber" value="<%=jobnumber%>" ID="Hidden14">
					<input type="hidden" name="VehicleID" value="<%=VehicleID%>" ID="Hidden25">
					<!--input type="hidden" name="LocationCode" value="<%=FromLocation%>" ID="Hidden26"-->
					<input type="hidden" name="jobnumber" value="<%=jobnumber%>" ID="Hidden27">	
					<input type="hidden" name="AliasCode" value="<%=AliasCode%>" ID="Hidden15">				
			
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
					l_csql = "SELECT Top 12 fcfgthd.fh_id, fclegs.fl_sf_rta, fcfgthd.fh_status, fcfgthd.fh_ship_dt, fclegs.fl_sf_id, fclegs.fl_sf_name, fclegs.fl_st_id, fclegs.fl_st_name, fclegs.fl_t_atp, fclegs.fl_t_atd, fclegs.fl_pod, fcfgthd.fh_custpo, fcfgthd.fh_priority, fcrefs.rf_ref FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN fcrefs ON fclegs.fl_fh_id = fcrefs.rf_fh_id "	
					'else
					'l_csql = "SELECT Distinct(fcfgthd.fh_id), fclegs.fl_sf_rta, fcfgthd.fh_status, fcfgthd.fh_ship_dt, fclegs.fl_sf_id, fclegs.fl_sf_name, fclegs.fl_st_id, fclegs.fl_st_name, fclegs.fl_t_atp, fclegs.fl_t_atd, fclegs.fl_pod, fcfgthd.fh_custpo, fcfgthd.fh_priority FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id "	
				'End if				
				l_csql = l_csql&" WHERE (Fl_dr_ID='"&VehicleID&"') AND fh_ship_dt>'"&now()-30&"'"
						
						
						If PageStatus="ONB" then
							l_csql = L_csql& "AND (fh_status='ACC') "
							l_csql = L_csql& "AND (ref_status is NULL) "
							l_csql = L_csql& "AND ((fl_sf_id='"&LocationCode&"') "
							If HUB="y" then
								l_csql = l_csql&" OR (Fl_sf_ID='D6W3')"
								l_csql = l_csql&" OR (Fl_sf_ID='D6N2')"
								l_csql = l_csql&" OR (Fl_sf_ID='D6N1')"
								l_csql = l_csql&" OR (Fl_sf_ID='DM4M')"
								l_csql = l_csql&" OR (Fl_sf_ID='DM5M')"
								l_csql = l_csql&" OR (Fl_sf_ID='DPI2')"
								l_csql = l_csql&" OR (Fl_sf_ID='DPI3')"
								l_csql = l_csql&" OR (Fl_sf_ID='ESTK')"
							End if	
							l_csql = l_csql& " ) "						
						end if
						If PageStatus="CLS" then
							l_csql = L_csql& "AND (fh_status='ONB') "
							l_csql = L_csql& "AND (ref_status='o') "
							l_csql = L_csql& "AND ((fl_st_id='"&LocationCode&"') "
							If HUB="y" then
								l_csql = l_csql&" OR (Fl_st_ID='D6W3')"
								l_csql = l_csql&" OR (Fl_st_ID='D6N2')"
								l_csql = l_csql&" OR (Fl_st_ID='D6N1')"
								l_csql = l_csql&" OR (Fl_st_ID='DM4M')"
								l_csql = l_csql&" OR (Fl_st_ID='DM5M')"
								l_csql = l_csql&" OR (Fl_st_ID='DPI2')"
								l_csql = l_csql&" OR (Fl_st_ID='DPI3')"
								l_csql = l_csql&" OR (Fl_st_ID='ESTK')"
							End if	
							l_csql = l_csql& " ) "								
						End if
						'Response.Write "pagestatus="&PageStatus&"<BR>"

						'If USESLOTS=TRUE then
							'SortBy="rf_ref"
							SortBy="fh_priority, fh_id"
						'End if
						If SortBy>"" then
							l_csql = L_csql& " ORDER BY "&Sortby
						End if

					'Response.write("Query3:" & l_cSQL)
			''''''''''''''''''''''''''''''''''''''''
				RSEVENTS2.Open l_cSQL, Database, 1, 3
				If RSEVENTS2.eof then
						Response.Redirect("DriverIfabPhoneEmulator.asp?AliasCode="&AliasCode&"&FakeSubmit=fakesubmit")
						'Response.Write "IM HERE!<BR>"
						ErrorMessage="No jobs were found that match your criteria."	
				End if
				'If not RSEVENTS2.EOF THEN				
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
					X=X+1
					%>
				<tr>
					<td class="mainpagetextboldcenter" nowrap><input type="text" name="FormBarCode<%=X%>" ID="Text2" size="3"></td>	
					<td class="MainPageTextSmaller" nowrap valign="top" bgcolor="<%=colorset(i mod numcolors)%>">
						<b><%=fh_id%></b><br><%=X%>)&nbsp;<%=rf_ref%>
					</td>				
				</tr>
				<input type="hidden" name="FormJobNumber(<%=X%>)" value="<%=fh_id%>">	
				<input type="hidden" name="AllegedBarCode(<%=X%>)" value="<%=rf_ref%>" ID="Hidden1">				

<%
				i=i+1
				'END IF
				RSEVENTS2.movenext
				LOOP
				RSEVENTS2.close
			Set RSEVENTS2 = Nothing

	%>
			<input type="hidden" name="LocationCode" value="<%=LocationCode%>">
			<input type="hidden" name="BillToID" value="<%=BillToID%>">
			<tr>
				<td colspan="2">
					<input type="submit" name="submit" value="submit">
				</td>
			</tr>
			</form>	
			</table>	
			<tr><td>&nbsp;</td></tr>
			<tr><td>&nbsp;</td></tr>
			<tr><td>&nbsp;</td></tr>			

	</Table>
<%end if%>
	</td></tr>
</body>
</html>
