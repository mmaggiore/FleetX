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
	response.Redirect("DriverPOB.asp")
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
	'Response.write "GOT HERE AFTER SUBMIT<BR>"

		If PageStatus="POB" then
			ORDERSTATUS="p"
			else
			ORDERSTATUS=""
			if Hub="y" then
				ORDERSTATUS="H"
			End if
		End if
		

		If PageStatus="POB" then
			For q=1 to 12
				'Response.Write "formBarCode="&FormBarCode(q)&"<BR>"
				If trim(FormBarCode(q))>""  then
					Set oRs = Server.CreateObject("ADODB.Recordset")
					oRs.CursorLocation = 3
					oRs.CursorType = 3
					oRs.ActiveConnection = DATABASE	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''					
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
					SQL = "SELECT fcfgthd.fh_id, fclegs.fl_sf_rta, fcfgthd.fh_status, fcfgthd.fh_ship_dt, fclegs.fl_sf_id, fclegs.fl_sf_name, fclegs.fl_st_id, fclegs.fl_st_name, fclegs.fl_t_atp, fclegs.fl_t_atd, fclegs.fl_pod, fcfgthd.fh_custpo, fcfgthd.fh_priority, fcrefs.rf_ref FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN fcrefs ON fclegs.fl_fh_id = fcrefs.rf_fh_id "	
					SQL = SQL&" WHERE (Fl_dr_ID='"&VehicleID&"') AND (fh_ship_dt>'"&now()-30&"') AND (rf_ref='"&FormBarCode(q)&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL))"
					If PageStatus="POB" then
						If trim(vehicleID)<>"199" then
							SQL = SQL& "AND (fh_status='ACC') "
							SQL = SQL& "AND (ref_status is NULL) "
							else
							SQL = SQL& "AND (fh_status='ONB') "
							SQL = SQL& "AND (ref_status='o') "							
						End if
						'SQL = SQL& "AND ((fl_sf_id='"&LocationCode&"') "
						'If HUB="y" then
						'	SQL = SQL&" OR (Fl_sf_ID='D6W3')"
						'	SQL = SQL&" OR (Fl_sf_ID='D6N2')"
						'	SQL = SQL&" OR (Fl_sf_ID='D6N1')"
						'	SQL = SQL&" OR (Fl_sf_ID='DM4M')"
						'	SQL = SQL&" OR (Fl_sf_ID='DM5M')"
						'	SQL = SQL&" OR (Fl_sf_ID='DPI2')"
						'	SQL = SQL&" OR (Fl_sf_ID='DPI3')"
						'	SQL = SQL&" OR (Fl_sf_ID='ESTK')"
						'End if	
						'SQL = SQL& " ) "						
					end if
					'SortBy="fh_priority, fh_id"
					'If SortBy>"" then
					'	SQL = SQL& " ORDER BY "&Sortby
					'End if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''					
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''					
					'Response.write "<br><font color='blue'>First Select="&SQL&"<BR></font>"
					oRs.Open SQL, DATABASE, 1, 3
					If oRs.EOF then
						ErrorMessage=ErrorMessage&" "&FormBarCode(q)&" is not correct.<br>"
					End if
					If not oRs.eof then
						TheJobNumber = oRs("fh_id")
						TheBarCode = FormBarCode(q)
						'Response.write "TheJobNumber=***"&TheJobNumber&"***<BR>"
						'Response.write "TheBarCode=***"&TheBarCode&"***<BR>"
						''Response.write "GOT HERE!<BR>"
						'TheJobNumber=trim(FormJobNumber(q))
						'TheAllegedBarCode=trim(AllegedBarCode(q))
						'TheBarCode=trim(FormBarCode(q))				
						'If TheJobNumber>"" then
						Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open DATABASE
						''''UPDATES THE WAFER
						l_cSQL = "UPDATE FCREFS SET ref_status = '"&ORDERSTATUS&"' WHERE (rf_fh_id = '"&TheJobNumber&"') AND (rf_ref = '" &TheBarCode& "')"
						'Response.write "UPDATE Wafers="&l_cSQL&"<BR>"
						oConn.Execute(l_cSQL)
						''''''''''''''''''''''''''''''''''''''''
						Set oRs = Server.CreateObject("ADODB.Recordset")
						oRs.CursorLocation = 3
						oRs.CursorType = 3
						oRs.ActiveConnection = DATABASE	
						''''''CHECKS TO SEE IF ITS OKAY TO CHANGE THE ORDER STATUS
						SQL = "SELECT rf_fh_id FROM fcrefs"
						SQL = SQL&" WHERE (rf_fh_id='"&TheJobNumber&"')"
						SQL = SQL&" AND (Ref_Status IS NULL)"
						'SQL = SQL&" ORDER BY Fh_Priority, fh_status desc, fh_id"
						'Response.write "<br><font color='blue'>Change the status?="&SQL&"<BR></font>"
						oRs.Open SQL, DATABASE, 1, 3
						If oRs.EOF then
							''''''''''''''''''''''''''''''''''''''''
							oConn.Execute "PHONE_CHANGE_STATUS '" & TheJobNumber & "', '13', 'PUO', '', '', '', ''" 
							'''''m_cSQL = "UPDATE FCFGTHD SET fh_status = 'PUO', fh_statcode=13 WHERE fh_id = '" & TheJobNumber&"'"
							'Response.write "UPDATE FCFGTHD="&m_cSQL&"<BR>"
							'''''oConn.Execute(m_cSQL)
							'''''n_cSQL = "UPDATE FCLEGS SET fl_t_int = '"&now()&"' WHERE fl_fh_id = '" & TheJobNumber&"'"
							'Response.write "UPDATE FCLEGS="&n_cSQL&"<BR>"
							'''''oConn.Execute(n_cSQL)
							o_cSQL = "UPDATE FCWTTIME SET wt_stop = '"&now()&"' WHERE wt_fh_id = '" & TheJobNumber&"'"
							'Response.write "UPDATE FCLEGS="&n_cSQL&"<BR>"
							oConn.Execute(o_cSQL)							
							''''''''''''''''''''''''''''''''''''''''
						End if
						oRs.Close
						Set oRs=Nothing							
						''''''''''''''''''''''''''''''''''''''''
						oConn.Close	
						Set oConn=Nothing	
						'End if
						TheJobNumber=""
						TheBarCode=""	
					End if
					''''cookie?("TempJobNumber")=JobNumber
				End if
			Next
			'End if
			'oRs.Close
			'Set oRs=Nothing		
		End if
			
			
			
			
			
			
			
			
		
		If PageStatus="CLS" then
			For q=1 to 12
				'Response.Write "formBarCode="&FormBarCode(q)&"<BR>"
				If trim(FormBarCode(q))>""  then
					Set oRs = Server.CreateObject("ADODB.Recordset")
					oRs.CursorLocation = 3
					oRs.CursorType = 3
					oRs.ActiveConnection = DATABASE	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''					
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
					SQL = "SELECT fcfgthd.fh_id, fclegs.fl_sf_rta, fcfgthd.fh_status, fcfgthd.fh_ship_dt, fclegs.fl_sf_id, fclegs.fl_sf_name, fclegs.fl_st_id, fclegs.fl_st_name, fclegs.fl_t_atp, fclegs.fl_t_atd, fclegs.fl_pod, fcfgthd.fh_custpo, fcfgthd.fh_priority, fcrefs.rf_ref FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN fcrefs ON fclegs.fl_fh_id = fcrefs.rf_fh_id "	
					SQL = SQL&" WHERE (Fl_dr_ID='"&VehicleID&"') AND (fh_ship_dt>'"&now()-30&"') AND (rf_ref='"&FormBarCode(q)&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL))"
					If PageStatus="ONB" then
						SQL = SQL& "AND (fh_status='ACC') "
						SQL = SQL& "AND (ref_status is NULL) "
						SQL = SQL& "AND ((fl_sf_id='"&LocationCode&"') "
						If HUB="y" then
							SQL = SQL&" OR (Fl_sf_ID='D6W3')"
							SQL = SQL&" OR (Fl_sf_ID='D6N2')"
							SQL = SQL&" OR (Fl_sf_ID='D6N1')"
							SQL = SQL&" OR (Fl_sf_ID='DM4M')"
							SQL = SQL&" OR (Fl_sf_ID='DM5M')"
							SQL = SQL&" OR (Fl_sf_ID='DPI2')"
							SQL = SQL&" OR (Fl_sf_ID='DPI3')"
							SQL = SQL&" OR (Fl_sf_ID='ESTK')"
						End if	
						SQL = SQL& " ) "						
					end if
					If PageStatus="CLS" then
						SQL = SQL& "AND (fh_status='ONB') "
						SQL = SQL& "AND (ref_status='o') "
						SQL = SQL& "AND ((fl_st_id='"&LocationCode&"') "
						If HUB="y" then
							SQL = SQL&" OR (Fl_st_ID='D6W3')"
							SQL = SQL&" OR (Fl_st_ID='D6N2')"
							SQL = SQL&" OR (Fl_st_ID='D6N1')"
							SQL = SQL&" OR (Fl_st_ID='DM4M')"
							SQL = SQL&" OR (Fl_st_ID='DM5M')"
							SQL = SQL&" OR (Fl_st_ID='DPI2')"
							SQL = SQL&" OR (Fl_st_ID='DPI3')"
							SQL = SQL&" OR (Fl_st_ID='ESTK')"
						End if	
						SQL = SQL& " ) "								
					End if
					'SortBy="fh_priority, fh_id"
					'If SortBy>"" then
					'	SQL = SQL& " ORDER BY "&Sortby
					'End if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''					
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''					
					'Response.write "<br><font color='blue'>First Select="&SQL&"<BR></font>"
					oRs.Open SQL, DATABASE, 1, 3
					If oRs.EOF then
						ErrorMessage=ErrorMessage&" "&FormBarCode(q)&" is not correct.<br>"
					End if
					If not oRs.eof then
						TheJobNumber = oRs("fh_id")
						TheBarCode = FormBarCode(q)
						'Response.write "TheJobNumber=***"&TheJobNumber&"***<BR>"
						'Response.write "TheBarCode=***"&TheBarCode&"***<BR>"
						''Response.write "GOT HERE!<BR>"
						'TheJobNumber=trim(FormJobNumber(q))
						'TheAllegedBarCode=trim(AllegedBarCode(q))
						'TheBarCode=trim(FormBarCode(q))				
						'If TheJobNumber>"" then
						Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open DATABASE
						''''UPDATES THE WAFER
						l_cSQL = "UPDATE FCREFS SET ref_status = '"&ORDERSTATUS&"' WHERE (rf_fh_id = '"&TheJobNumber&"') AND (rf_ref = '" &TheBarCode& "')"
						'Response.write "UPDATE Wafers="&l_cSQL&"<BR>"
						oConn.Execute(l_cSQL)
						''''''''''''''''''''''''''''''''''''''''
						Set oRs = Server.CreateObject("ADODB.Recordset")
						oRs.CursorLocation = 3
						oRs.CursorType = 3
						oRs.ActiveConnection = DATABASE	
						''''''CHECKS TO SEE IF ITS OKAY TO CHANGE THE ORDER STATUS
						SQL = "SELECT rf_fh_id FROM fcrefs"
						SQL = SQL&" WHERE (rf_fh_id='"&TheJobNumber&"')"
						SQL = SQL&" AND (Ref_Status='o')"
						'SQL = SQL&" ORDER BY Fh_Priority, fh_status desc, fh_id"
						'Response.write "<br><font color='blue'>Change the status?="&SQL&"<BR></font>"
						oRs.Open SQL, DATABASE, 1, 3
						If oRs.EOF then
							''''''''''''''''''''''''''''''''''''''''
							oConn.Execute "PHONE_CHANGE_STATUS '" & TheJobNumber & "', '9', 'CLS', '', '', '', ''" 
							'''''m_cSQL = "UPDATE FCFGTHD SET fh_status = 'CLS', fh_statcode=9 WHERE fh_id = '" & TheJobNumber&"'"
							'Response.write "UPDATE FCFGTHD="&m_cSQL&"<BR>"
							'''''oConn.Execute(m_cSQL)
							'''''n_cSQL = "UPDATE FCLEGS SET fl_t_atd = '"&now()&"' WHERE fl_fh_id = '" & TheJobNumber&"'"
							'Response.write "UPDATE FCLEGS="&n_cSQL&"<BR>"
							'''''oConn.Execute(n_cSQL)
							''''''''''''''''''''''''''''''''''''''''
						End if
						oRs.Close
						Set oRs=Nothing							
						''''''''''''''''''''''''''''''''''''''''
						oConn.Close	
						Set oConn=Nothing	
						'End if
						TheJobNumber=""
						TheBarCode=""	
					End if
					''''cookie?("TempJobNumber")=JobNumber
				End if
			Next
			'End if
			'oRs.Close
			'Set oRs=Nothing	

		End if
		
		
		
		
		
		
		
		'oConn.close
		Set oConn=Nothing	
	
END IF


%>
<body leftMargin=0 TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" OnLoad=document.Form1.FormBarCode1.focus()>
	<TABLE WIDTH="300" border="0" bordercolor="red" cellpadding="0" cellspacing="5" align="left" ID="Table1">
		<tr><td align="center" colspan="3"><form method="post" action="DriverPOB.asp" ID="Form7"><input type="hidden" name="Aliascode" value="<%=LocationCode%>" ID="Hidden1"><input type="submit" value="Return to Drop Off/Pick Up" ID="Submit1" NAME="Submit1"></form></td></tr>
		<tr><td align="left">
		<%
		If Submit>"" then
		''Response.write "GOT HERE!<BR>"
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
				<tr><td class="ErrorMessageBoldCenter" colspan="2"><%=ErrorMessage%></td></tr>
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
			''Response.write "******optJobSel="&optJobSel&"<BR>"
'			Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
'				RSEVENTS2.CursorLocation = 3
'				RSEVENTS2.CursorType = 3
'				RSEVENTS2.ActiveConnection = Database
				''Response.write "GOT HERE #2!<BR>"
				''Response.write "optJobSel="&optJobSel&"<BR>"
				'If USESLOTS=TRUE then
					l_csql = "SELECT fcfgthd.fh_id, fclegs.fl_sf_rta, fcfgthd.fh_status, fcfgthd.fh_ship_dt, fclegs.fl_sf_id, fclegs.fl_sf_name, fclegs.fl_st_id, fclegs.fl_st_name, fclegs.fl_t_atp, fclegs.fl_t_atd, fclegs.fl_pod, fcfgthd.fh_custpo, fcfgthd.fh_priority, fcrefs.rf_ref FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN fcrefs ON fclegs.fl_fh_id = fcrefs.rf_fh_id "	
					'else
					'l_csql = "SELECT Distinct(fcfgthd.fh_id), fclegs.fl_sf_rta, fcfgthd.fh_status, fcfgthd.fh_ship_dt, fclegs.fl_sf_id, fclegs.fl_sf_name, fclegs.fl_st_id, fclegs.fl_st_name, fclegs.fl_t_atp, fclegs.fl_t_atd, fclegs.fl_pod, fcfgthd.fh_custpo, fcfgthd.fh_priority FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id "	
				'End if				
				l_csql = l_csql&" WHERE (Fl_dr_ID='"&VehicleID&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND fh_ship_dt>'"&now()-30&"'"
				''Response.write "l_csql="&l_csql&"<BR>"		
						
						If PageStatus="POB" then
							If trim(vehicleID)<>"199" then
								l_csql = L_csql& " AND (fh_status='ACC') "
								l_csql = L_csql& "AND (ref_status is NULL) "
								else
								l_csql = L_csql& " AND (fh_status='ONB') "
								l_csql = L_csql& "AND (ref_status='o') "
								l_csql = L_csql& "AND (fl_sf_id='"&LocationCode&"') "
							End if
							
							'l_csql = L_csql& "AND ((fl_sf_id='"&LocationCode&"') "
							'If HUB="y" then
							'	l_csql = l_csql&" OR (Fl_sf_ID='D6W3')"
							'	l_csql = l_csql&" OR (Fl_sf_ID='D6N2')"
							'	l_csql = l_csql&" OR (Fl_sf_ID='D6N1')"
							'	l_csql = l_csql&" OR (Fl_sf_ID='DM4M')"
							'	l_csql = l_csql&" OR (Fl_sf_ID='DM5M')"
							'	l_csql = l_csql&" OR (Fl_sf_ID='DPI2')"
							'	l_csql = l_csql&" OR (Fl_sf_ID='DPI3')"
							'	l_csql = l_csql&" OR (Fl_sf_ID='ESTK')"
							'End if	
							'l_csql = l_csql& " ) "						
						end if
						If PageStatus="CLS" then
							l_csql = L_csql& "AND (fh_status='ONB') "
							l_csql = L_csql& "AND (ref_status='o') "
							'l_csql = L_csql& "AND ((fl_st_id='"&LocationCode&"') "
							'If HUB="y" then
							'	l_csql = l_csql&" OR (Fl_st_ID='D6W3')"
							'	l_csql = l_csql&" OR (Fl_st_ID='D6N2')"
							'	l_csql = l_csql&" OR (Fl_st_ID='D6N1')"
							'	l_csql = l_csql&" OR (Fl_st_ID='DM4M')"
							'	l_csql = l_csql&" OR (Fl_st_ID='DM5M')"
							'	l_csql = l_csql&" OR (Fl_st_ID='DPI2')"
							'	l_csql = l_csql&" OR (Fl_st_ID='DPI3')"
							'	l_csql = l_csql&" OR (Fl_st_ID='ESTK')"
							'End if	
							'l_csql = l_csql& " ) "								
						End if
						''Response.write "pagestatus="&PageStatus&"<BR>"

						'If USESLOTS=TRUE then
							'SortBy="rf_ref"
							SortBy="fh_priority, fh_id, rf_ref"
						'End if
						If SortBy>"" then
							l_csql = L_csql& " ORDER BY "&Sortby
						End if

					'Response.write("Query3XXX:" & l_cSQL)
			''''''''''''''''''''''''''''''''''''''''
'				RSEVENTS2.Open l_cSQL, Database, 1, 3
'				If RSEVENTS2.eof then
'						Response.Redirect("DriverPOB.asp?PageStatus=loggedin&AliasCode="&AliasCode&"&FakeSubmit=fakesubmit")
'						''Response.write "IM HERE!<BR>"
'						ErrorMessage="No jobs were found that match your criteria."	
'				End if
'				'If not RSEVENTS2.EOF THEN				
'				Do while not RSEVENTS2.EOF 
				'''''''''''''''''''''''''''''''''''''''''''''''''''''
				'Do while not RSEVENTS2.eof
				Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				RSEVENTS2.CursorLocation = 3
				RSEVENTS2.CursorType = 3
				RSEVENTS2.ActiveConnection = DATABASE					
				RSEVENTS2.Open L_cSQL, DATABASE, 1, 3
				%>
							<tr>
							<td align="center" class="purpleseparator" colspan="3"><b>PAPERWORK TO PICK UP</b></td>
						</tr>
				<%
				If not RSEVENTS2.EOF then
						ELSE
						Response.Redirect("DriverPOB.asp?PageStatus=loggedin&AliasCode="&AliasCode&"&FakeSubmit=fakesubmit")
						''Response.write "IM HERE!<BR>"
						ErrorMessage="No jobs were found that match your criteria."	
				End if				
				
				
				
				
				RSEVENTS2.PageSize = 8
				RSEVENTS2.CacheSize = RSEVENTS2.PageSize
				intPageCount2 = RSEVENTS2.PageCount
				intRecordCount2 = RSEVENTS2.RecordCount
				If (RSEVENTS2.EOF) then
					'Response.Redirect("default.asp")
					Sendback2="y"
					'Response.Write "Got here #3<br>"
				End if
				If NOT (RSEVENTS2.BOF AND RSEVENTS2.EOF) Then

				If CInt(intPage2) > CInt(intPageCount2) Then intPage2 = intPageCount2
					If CInt(intPage2) <= 0 Then intPage2 = 1
						If intRecordCount2 > 0 Then
							RSEVENTS2.AbsolutePage = intPage2
							intStart = RSEVENTS2.AbsolutePosition
							If CInt(intPage2) = CInt(intPageCount2) Then
								intFinish = intRecordCount
							Else
								intFinish = intStart + (RSEVENTS2.PageSize - 1)
							End if
						End If
					If intRecordCount2 > 0 Then
						For intRecord2 = 1 to RSEVENTS2.PageSize				
				'''''''''''''''''''''''''''''''''''''''''''''''''''''



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
						<!--<b><%=fh_id%></b><br>--><%=X%>)&nbsp;<%=rf_ref%>
					</td>				
				</tr>
				<!--
				<input type="hidden" name="FormJobNumber(<%=X%>)" value="<%=fh_id%>" ID="Hidden2">	
				<input type="hidden" name="AllegedBarCode(<%=X%>)" value="<%=rf_ref%>" ID="Hidden6">				
				-->
<%
					''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
					'RSEVENTS2.Movenext
				'Loop
				Response.Write "</font>"
				'RSEVENTS2.Close
				RSEVENTS2.MoveNext
				If colorchanger = 1 Then
					colorchanger = 0
					color1 = "class=headerwhite"
					color2 = "class=header"
				Else
					colorchanger = 1
					color1 = "class=header"
					color2 = "class=headerwhite"	
				End If
				If RSEVENTS2.EOF Then Exit for
					Next
					End if
					End if					
				''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'				i=i+1
'				'END IF
'				RSEVENTS2.movenext
'				LOOP
'				RSEVENTS2.close
'			Set RSEVENTS2 = Nothing

	%>
			<input type="hidden" name="LocationCode" value="<%=LocationCode%>" ID="Hidden8">
			<input type="hidden" name="BillToID" value="<%=BillToID%>" ID="Hidden9">
			
			<tr>
				<td colspan="2">
					<input type="submit" name="submit" value="submit" ID="Submit2">
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
