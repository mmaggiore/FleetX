<%@ LANGUAGE="VBSCRIPT"%>
<%Response.buffer = true%>
<!-- #include file="../fleetexpress.inc" -->
<%
'response.Write "hello???<BR>"
BillToID=Session("Suid")
If BillToID="" then
	BillToID=Request.QueryString("BillToID")
End if
Session("sBT_ID")=BillToID
whatevah=Session("sBT_ID")
BillToName=trim(Session("sUsername"))
'Response.write "WHATEVAH!="&whatevah&"***<BR>"
'Response.write "BillToID="&BillToID&"***<BR>"
%>
<html>
<head>
<link rel="stylesheet" href="../themes/classic.css" id="theme_base">
<link rel="stylesheet" href="../themes/classic.date.css" id="theme_date">
<link rel="stylesheet" href="../themes/classic.time.css" id="theme_time">
<% 
i=0
ii=0

changedirectory="../marketing/"
PageNameText="Track and Trace"
Submit=Request.Form("Submit")
'ResetButton=Request.Form("ResetButton")
'Response.Write "XXXXXSubmit="&Submit&"<BR>"
If Submit="" then
	Submit=Request.QueryString("Submit")
End if
'Response.Write "YYYYYYSubmit="&Submit&"<BR>"
If Submit>"" then Submit2=Submit end if
If ResetButton<>"clear search" then
	DateSentFrom=Request.Form("DateSentFrom")
	DateSentTo=Request.Form("DateSentTo")
	DocumentNumber=Request.Form("DocumentNumber")
	If DocumentNumber="" then
		DocumentNumber=Request.QueryString("DocumentNumber")
	End if
	LotNumber=Request.Form("LotNumber")
	If LotNumber="" then
		LotNumber=Request.QueryString("LotNumber")
	End if
	SortBy=Request.Form("SortBy")
	ToLocation=Request.Form("ToLocation")
	FromLocation=Request.Form("FromLocation")
	JobNumber=Request.Form("JobNumber")
    DocumentNumber=Request.form("DocumentNumber")
	If JobNumber="" then
		JobNumber=Request.QueryString("JobNumber")
	End if
	ReferenceNumber=Request.Form("ReferenceNumber")
	Priority=Request.Form("Priority")
	JobStatus=Request.Form("JobStatus")
End if
'Response.Write "53 tracking - jobnumber = " & jobnumber & "<br>"
'Response.Write "Priority="&Priority&"<BR>"
'Response.Write "BillToID="&BillToID&"<BR>"
Select Case BillToID
	Case 48 'KWE
		LotWord="HAWB Number"
	Case 36 'WAFER
		LotWord="Lot Number"
	Case 38, 72 'RETICLES
		LotWord="Reticle Number"
	Case 13, 14, 25 'ABBOTT ROSS
		LotWord="BOL Number"		
	Case 26 'RETICLES
		LotWord="Document Number"
	Case 75 'TI-AIMS
		LotWord="PO Number"	
	Case 76 'TOPAN
		LotWord="Reticle Number"
		'response.Write "Got here 1<BR>"				
	Case else
		LotWord="Order Number"
		'response.Write "Got here 2Based on previous delivery<BR>"			
End Select
'If DateSent="" then
	'DateSent=Date()
	'else
If Submit="" or DateSentFrom="" or DateSentTo="" then
	DateSentFrom=Date()-7
	DateSentTo=Date()
End if
If DateSentTo>"" then
	SQLDateSentTo=cDate(DateSentTo)+1
End if
	''Response.write "DateSent="&DateSent&"<BR>"
	''Response.write "DayAfter="&DayAfter&"<BR>"
'End if


DocumentNumber=Replace(DocumentNumber,"""","")
DocumentNumber=Replace(DocumentNumber,"'","")

LotNumber=Replace(LotNumber,"""","")
LotNumber=Replace(LotNumber,"'","")
JobNumber=Replace(JobNumber,"""","")
JobNumber=Replace(JobNumber,"'","")
'''suid=Session("suid")
ReferenceNumber=Replace(ReferenceNumber,"""","")
ReferenceNumber=Replace(ReferenceNumber,"'","")

USESLOTS = FALSE
'Response.write "115 tracking USESLOTS="&USESLOTS&"<BR>"
''Response.write "SortBy="&SortBy&"*<BR>"
%>
<SCRIPT Language="Javascript" SRC="Script/Calendar1-902.js"></SCRIPT> 
<link rel="stylesheet" type="text/css" href="../css/Style.css">
<LINK REL="stylesheet" TYPE="text/css" HREF="Script/Calendar.css"> 

 <%
    ColorSelect=Request.form("ColorSelect")
    ColorSelect=ColorSelect+1
    If ColorSelect>4 then ColorSelect=1 End if
    ColorSelect=3
    Select Case ColorSelect
        Case 1
            HeaderBorderColor="#cc1126"
            BorderColor="#cc1126"
            LinkClass="FleetExpressRed"
        Case 2
             HeaderBorderColor="#216194"
            BorderColor="#216194"
            LinkClass="FleetExpressBlue"
        Case 3 
            'HeaderBorderColor="#B7B8B8" 
            HeaderBorderColor="#d71e26"  
            BorderColor="#d71e26"
            LinkClass="FleetXRed"
        Case else 
            HeaderBorderColor="black"  
            BorderColor="black"
            LinkClass="FleetExpressBlack"
    End Select
    HighlightedField="RequestorName"
    CurrentDateTime=Now()
    PageTitle="TRACKING"

%>
<title>FleetX - <%=PageTitle %></title>
</head>

<body onload="document.FindUser.requestorName.focus();">
	<table align="center" border="0" bordercolor="black" cellpadding="0" cellspacing="0" ID="Table1" width="100%" height="100%">
        <tr><td align="left" height="10"><img src="../images/pixel.gif" height="10" width="1" /><img src="../images/pixel.gif" height="1" width="1" /></td></tr>
        <tr>
            <td align="left" height="50"><img src="../images/FleetX_Small.jpg" height="50" width="168" /></td>
            <td align="right" valign="bottom"><!-- #include file="../topnavbar.asp" --></td>
        </tr>
        <tr><td align="left" colspan="2" height="1"><img src="../images/pixel.gif" /></td></tr>	
    <tr><td colspan="2">
<!-- <form action="/Intranet/FleetX/NewUser.asp" method="post" name="FindUser"> -->
<table border="0" bordercolor="green" Cellspacing="0" Cellpadding="0" align="left" width="100%" height="100%">
 <tr><td align="left" colspan="2" bgcolor="<%=HeaderBorderColor%>"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>

        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetXRedSection" colspan="2"><%=PageTitle%></td></tr>

        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
        <!--tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressBodyWhite" colspan="2">In order to reach the Fleet Express Order page, please correctly type in the green verification code in the supplied text box and click "Submit."</td></tr-->
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
<tr><td>
<table  border="0" bordercolor="blue" valign="top" align="center" class="MainPageText" width="100%" >
	<tr height="40">
		<td width="650" valign="top">&nbsp;</td>
	</tr>

    <tr><td width="100%" align="center" valign="top">
    
	 <TABLE WIDTH="100%" valign="top" border="0" bordercolor="red" cellpadding="0" cellspacing="5" ID="Table1">
        <tr><td> 
      
		<table cellpadding="0" valign="top" cellspacing="0" border="0" bordercolor="brown" align="center" ID="Table2">  
			<tr><td class="MainPageText">Provide as much, or as little search criteria as needed</td></tr>
			<tr><td>&nbsp;</td></tr>
			<tr>
				<td>
					<table border="0" bordercolor="red" align="center" ID="Table3" class="MainPageText">
					<form method="post" action="Tracking.asp" name="thisForm" ID="Form1">
						
				<!--		<tr>
							<td class='subheader' align="right">
							<%=LotWord%>:
							</td>
							<td>
								<input type="text" name="LotNumber" value="<%=LotNumber%>" size="20" ID="Text4">
							</td>					
						</tr>   -->
						<tr>
							<td class="MainPageText" align="right">
							Order Number:
							</td>
							<td>
								<input type="text" name="JobNumber" value="<%=JobNumber%>" size="20" ID="Text2">
							</td>					
						</tr>
                        <%
                        'response.write "BillToID="&BillToID&"<BR>" 
                        'response.write "CustomerID="&CustomerID&"<BR>"
                         'response.write "RequestorCompany="&RequestorCompany&"<BR>"
                        %>
 						<tr>
							<td class="MainPageText" align="right">
							Document Number:
							</td>
							<td>
								<input type="text" name="DocumentNumber" value="<%=DocumentNumber%>" size="20" ID="Text4">
							</td>					
						</tr>                       												
						<tr>
							<td class="MainPageText" align="right" valign="top">
							Date Sent Range:
							</td>
                        </tr>
									<tr>
										<td class="MainPageText" align="right">From:</td>
										<td>
											<input type="text" name="DateSentFrom" value="<%=DateSentFrom%>" size="8" ID="date_1">
                                            <!--
											<a href="javascript:show_calendar('thisForm.DateSentFrom');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;"><img src="../images/cal.gif" width="16" height="15" border="0" name="calendar" alt="Calendar" title="Calendar" align="ABSMIDDLE"></a>						
										    -->
                                        </td>
									</tr>
									<tr>
										<td class="MainPageText" align="right">To:</td>
										<td>
											<input type="text" name="DateSentTo" value="<%=DateSentTo%>" size="8" ID="date_2">
                                            <!--
											<a href="javascript:show_calendar('thisForm.DateSentTo');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;"><img src="../images/cal.gif" width="16" height="15" border="0" name="calendar" alt="Calendar" title="Calendar" align="ABSMIDDLE"></a>						
										    -->
                                        </td>
									</tr>
						
						<tr>
							<td class="MainPageText" align="right">
								Priority:
							</td>
							<%'response.Write "ToLocation="&ToLocation&"***<BR>"%>
							<td>
								<select name="Priority" ID="Select4">
									<option value="" <%if Priority="" then response.Write " Selected " end if%>>All Priorities</option>
									<%
									Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
										RSEVENTS2.CursorLocation = 3
										RSEVENTS2.CursorType = 3
										RSEVENTS2.ActiveConnection = Database
										l_csql = "SELECT DISTINCT PriorityDescription AS PriorityDescription, PriorityMinutes AS PriorityTime, Priority AS PriorityAbbreviation FROM priorities WHERE (Priority_BT_ID = '"&BillToID&"') order by Priority"
										'response.write("Query:" & l_cSQL)
										RSEVENTS2.Open l_cSQL, Database, 1, 3
										Do while not RSEVENTS2.EOF
											PriorityDescription=RSEVENTS2("PriorityDescription")
											PriorityAbbreviation=RSEVENTS2("PriorityAbbreviation")
										%>
											<option value="<%=PriorityAbbreviation%>" <%if PriorityAbbreviation=Priority then response.Write "selected" end if%>><%=PriorityDescription%></option>
										<%	
										RSEVENTS2.movenext
										LOOP
										RSEVENTS2.close
									Set RSEVENTS2 = Nothing							
									%>
								</select>
								<%
								'Response.Write "l_csql="&l_csql&"<BR>"
								%>
							</td>
						</tr>						
						
						<tr>
							<td class="MainPageText" align="right">
								Order Status:
							</td>
							<%'response.Write "ToLocation="&ToLocation&"***<BR>"%>
							<td>
								<select name="JobStatus" ID="Select5">
									<option value="" <%if JobStatus="" then response.Write " Selected " end if%>>All Order Statuses</option>
									<option value="98" <%if JobStatus="98" then response.Write "selected" end if%>>Cancelled</option>
									<option value="9" <%if JobStatus="9" then response.Write "selected" end if%>>Closed</option>
									<option value="OPEN" <%if JobStatus="OPEN" then response.Write "selected" end if%>>Open</option>
								</select>
								<%
								'Response.Write "l_csql="&l_csql&"<BR>"
								%>
							</td>
						</tr>							
						
						<tr>
							<td class="MainPageText" align="right">
								Origination:
							</td>
							<td>
								<select name="FromLocation" ID="Select2">
									<option value="" <%if FromLocation="" then response.Write " Selected " end if%>>All Locations</option>
									<%
									Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
										RSEVENTS2.CursorLocation = 3
										RSEVENTS2.CursorType = 3
										RSEVENTS2.ActiveConnection = Database
										l_csql = "SELECT fcshipbt.sb_bt_id, fcshipbt.sb_pkey, fcshipto.st_name, fcshipbt.sb_st_id FROM fcshipto INNER JOIN fcshipbt ON fcshipto.st_id = fcshipbt.sb_st_id WHERE (sb_bt_id='"&BillToID&"')"
										If BilltoID=36 then
										    l_csql=l_csql&" OR sb_st_id='TISHERMA' "
										End if										
										If BillToID=48 then
											l_csql=l_csql&" AND (St_Priapt='DFW')"
										End if
										If BillToID=38 then
											l_csql=l_csql&" AND (st_id<>'CPGP')"
										End if
										If BillToID=76 then
											l_csql=l_csql&" OR (st_id='TOPPAN')"
										End if												
										l_csql=l_csql&" ORDER BY st_name"										
										'response.write("Query2:" & l_cSQL)
										
										RSEVENTS2.Open l_cSQL, Database, 1, 3
										Do while not RSEVENTS2.EOF
										sb_st_id=RSEVENTS2("sb_st_id")
										st_name=RSEVENTS2("st_name")
										%>
										
											<option value="<%=sb_st_id%>" <%if sb_st_id=FromLocation then response.Write " Selected " end if%>><%=st_name%></option>
										
										<%	
										RSEVENTS2.movenext
										LOOP
										RSEVENTS2.close
									Set RSEVENTS2 = Nothing							
									%>
								</select>
								<%
								'Response.Write "l_csql="&l_csql&"<BR>"
								%>
							</td>
						</tr>						
						<tr>
							<td class="MainPageText" align="right">
								Destination:
							</td>
							<%'response.Write "ToLocation="&ToLocation&"***<BR>"%>
							<td>
								<select name="ToLocation" ID="Select1">
									<option value="" <%if ToLocation="" then response.Write " Selected " end if%>>All Locations</option>
									<%
									Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
										RSEVENTS2.CursorLocation = 3
										RSEVENTS2.CursorType = 3
										RSEVENTS2.ActiveConnection = Database
										l_csql = "SELECT fcshipbt.sb_bt_id, fcshipbt.sb_pkey, fcshipto.st_name, fcshipbt.sb_st_id FROM fcshipto INNER JOIN fcshipbt ON fcshipto.st_id = fcshipbt.sb_st_id WHERE (sb_bt_id='"&BillToID&"')"
										If BilltoID=36 then
										    l_csql=l_csql&" OR sb_st_id='TISHERMA' "
										End if
										If BillToID=48 then
											l_csql=l_csql&" AND (St_Name<>'KWE')"
										End if
										If BillToID=38 then
											l_csql=l_csql&" AND (st_id<>'55')"
										End if
										If BillToID=76 then
											l_csql=l_csql&" OR (st_id='TOPPAN')"
										End if										
										l_csql=l_csql&" ORDER BY st_name"																				
										'response.write("Query:" & l_cSQL)
										
										RSEVENTS2.Open l_cSQL, Database, 1, 3
										Do while not RSEVENTS2.EOF
										sb_st_id=RSEVENTS2("sb_st_id")
										st_name=RSEVENTS2("st_name")
										%>
										
											<option value="<%=sb_st_id%>" <%if sb_st_id=ToLocation then response.Write " Selected " end if%>><%=st_name%></option>
										
										<%	
										RSEVENTS2.movenext
										LOOP
										RSEVENTS2.close
									Set RSEVENTS2 = Nothing							
									%>
								</select>
								<%
								'Response.Write "l_csql="&l_csql&"<BR>"
								%>
							</td>
						</tr>

						<tr>
							<td class="MainPageText" align="right">
								Sort By:
							</td>
							<td>
								<select name="SortBy" ID="Select3">
								<%If BillToID=26 then%>
									<option value="fh_custpo asc" <%if SortBy="fh_custpo asc" then response.Write " Selected " end if%>><%=LotWord%> (Ascending)</option>
									<option value="fh_custpo desc" <%if SortBy="fh_custpo desc" then response.Write " Selected " end if%>><%=LotWord%> (Descending)</option>
									<option value="fl_sf_rta asc" <%if SortBy="" or SortBy="fl_sf_rta asc" or SortBy="" then response.Write " Selected " end if%>>SAP Order Time (Ascending)</option>									
									<option value="fl_sf_rta desc" <%if SortBy="fl_sf_rta desc" then response.Write " Selected " end if%>>SAP Order Time (Descending)</option>									
									
									<%
									else
									%>
									<option value="rf_ref asc" <%if SortBy="rf_ref asc" then response.Write " Selected " end if%>><%=LotWord%>  (Ascending)</option>
									<option value="rf_ref desc" <%if SortBy="rf_ref desc" then response.Write " Selected " end if%>><%=LotWord%> (Descending)</option>
									<option value="fh_ship_dt asc" <%if SortBy="fh_ship_dt asc" or SortBy="" then response.Write " Selected " end if%>>Booked Time (Ascending)</option>								
									<option value="fh_ship_dt desc" <%if SortBy="" or SortBy="fh_ship_dt desc" then response.Write " Selected " end if%>>Booked Time (Descending)</option>								

								<%end if%>
								</select>
							</td>
						</tr>
						<tr><td>&nbsp;</td></tr>						
						<tr><td><img src="../images/pixel.gif" height="1" width="1" border="0"></td></tr>
						<tr><td align="center" colspan="2"><input type="submit" name="submit" value="search" id="gobutton"></td></tr>	
					</form>
			<tr><td>&nbsp;</td></tr>
			<tr><td>&nbsp;</td></tr>
			<tr><td>&nbsp;</td></tr>
		</table>
				</td>
			</tr>
		</table>
		<%                                                                    
		'Response.Write "got here #1<BR>"
		'Response.Write "ZZZsubmit2="&submit2&"<BR>"
		If Submit2>"" then
		'Response.Write "got here 2<BR>"
		%>

           <tr><td>
			<table cellpadding="3" cellspacing="0" border="0" align="center" ID="Table5">
				<tr>
				<%If UsesLots=FALSE then
						ColspanNumber="7"
				%>
					<td class="SubHeader" nowrap>
						Order Number
					</td>				
					<%else
						ColspanNumber="8"
					%>
					<td class="SubHeader" nowrap>
						<%=LotWord%> 
					</td>				
				<%End if%>
					<!--
					<td class="MainPageTextBoldCentered" nowrap>
						Pickup
					</td>	
					<td class="MainPageTextBoldCentered" nowrap>
						Dropoff
					</td>
					-->
					<td class="SubHeader" nowrap>
						From
					</td>
					<td class="SubHeader" nowrap>
						To
					</td>										
					<td class="SubHeader" nowrap>
						Status
					</td>
					<%If UsesLots=TRUE then%>
					<td class="SubHeader" nowrap>
						Priority
					</td>
					<%End if%>					
					<td class="SubHeader" nowrap>
						Entered
					</td>	
					<td class="SubHeader" nowrap>
						Picked Up
					</td>				
					<td class="SubHeader" nowrap>
						Delivered
					</td>																						
				</tr>		
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
				'Response.write "Database="&Database&"<BR>"
				''''If USESLOTS=TRUE then
				'Response.write "DateSentFrom="&DateSentFrom&"<BR>"
				'Response.write "DateSentTo="&DateSentTo&"<BR>"
				NumberofDays=datediff("d",DateSentFrom, DateSentTo)
				'Response.write "NumberofDays="&NumberofDays&"<BR>"
					l_csql = "SELECT "
					If NumberofDays>0 then
						l_csql = l_csql&" Top 300 "	
					End if		
					'l_csql = l_csql&" Report_Data.fh_id, Report_Data.fh_status, Report_Data.fh_ship_dt, Report_Data.fl_sf_id, Report_Data.fl_sf_name, Report_Data.fl_st_id,Report_Data.fl_st_name, Report_Data.fl_t_acc, Report_Data.fl_t_atp, Report_Data.fl_t_atd, Report_Data.fh_custpo, Report_Data.fh_priority, Report_Refs.RF_REF, Report_Refs.RF_STATUS  FROM Report_Data INNER JOIN Report_Refs ON Report_Data.fh_id = Report_Refs.RF_FH_ID "	
					'Response.write "504 tracking l_csql="&l_csql&"<BR><br>"
                    'l_csql = l_csql&" WHERE fl_st_id=fl_finaldestination "
					''''else
          
					l_csql = "SELECT Distinct(fcfgthd.fh_id), fclegs.fl_sf_rta, fcfgthd.fh_status, fcfgthd.fh_ship_dt, fclegs.fl_sf_id, fclegs.fl_sf_name, fclegs.fl_st_id, fclegs.fl_st_name, fclegs.fl_t_atp, fclegs.fl_t_acc, fclegs.fl_t_atd, fclegs.fl_pod, fcfgthd.fh_priority FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id "	
				l_csql = L_csql& " WHERE (fh_id>'') "
        ''''End if
				'response.write "BillToName="&BillToName&"***<BR>"
				'Select Case BillToName
					'Case "comps"
						'l_csql = l_csql&" WHERE (((fl_st_id='CPGP') OR (fl_sf_id='55')) "
					'Case "Toppan"
						'l_csql = l_csql&" WHERE (((fl_st_id='TOPPAN') OR (fl_sf_id='TOPPAN')) "
					'Case "tiret"
						'l_csql = l_csql&" WHERE (((fh_bt_id='"&BillToID&"') OR ((fh_bt_id<>'26') AND (fh_bt_id<>'36'))) "
					'Case else
						'l_csql = l_csql&" WHERE ((fh_bt_id='"&BillToID&"') "
				'End Select

						'''If trim(ReferenceNumber)>"" then
						'''	l_csql = L_csql& "AND (rf_ref='"&ReferenceNumber&"') "
						'''End if
						'''If trim(DocumentNumber)>"" then
						'''	l_csql = L_csql& "AND (fh_custpo LIKE '%"&DocumentNumber&"') "
						'''End if
						'If trim(LotNumber)>"" then
						'	l_csql = L_csql& "AND (rf_ref LIKE '%"&LotNumber&"%') "
						'End if
						'If trim(JobNumber)>"" then
						'	l_csql = L_csql& "AND (fh_id LIKE '%"&JobNumber&"') "
						'End if	
						If trim(Priority)>"" then
						    If Priority="XP" then
						        l_csql = L_csql& "AND ((fh_priority = 'P1') OR (fh_priority = 'P0') OR (fh_priority = 'XP') ) "
						        else
							    l_csql = L_csql& "AND (fh_priority = '"&Priority&"') "
							End if
						End if							
						If trim(JobStatus)>"" then
							Select Case JobStatus
								Case "9"
									l_csql = L_csql& "AND (fh_status = 'CLS') "
								Case "98"
									l_csql = L_csql& "AND (fh_status = 'CAN') "
								Case "OPEN"
									l_csql = L_csql& "AND ((fh_status <> 'CLS') AND (fh_status <> 'CAN'))"
							End Select
							
						End if
                        
                        If trim(DocumentNumber)>"" then
                            l_csql = L_csql& "AND (fh_custpo = '"&DocumentNumber&"') "
                        End if							
																	
						'If DateSentTo>"" And DateSentFrom>"" and trim(LotNumber)="" and trim(JobNumber)="" then
                        If DateSentTo>"" And DateSentFrom>"" then
							If BillToID=26 then
								l_csql = L_csql& "AND (fl_sf_rta>='"&DateSentFrom&"') AND (fl_sf_rta<'"&SQLDateSentTo&"') "
								else
								l_csql = L_csql& "AND (fh_ship_dt>='"&DateSentFrom&"') AND (fh_ship_dt<'"&SQLDateSentTo&"') "
							End if
						End if
						If ToLocation>"" then
							l_csql = L_csql& "AND (fl_st_id='"&ToLocation&"') "
						End if
						If FromLocation>"" then
							l_csql = L_csql& "AND (fl_sf_id='"&FromLocation&"') "
						End if
							'l_csql = L_csql& ") "
						If trim(LotNumber)>"" and trim(JobNumber)="" then
							l_csql = L_csql& "AND (rf_ref LIKE '%"&trim(LotNumber)&"%') "
						End if
						If trim(JobNumber)>"" AND trim(LotNumber)="" then
							l_csql = L_csql& "AND (fh_id LIKE '%"&trim(JobNumber)&"') "
						End if	
						If trim(JobNumber)>"" AND trim(LotNumber)>"" then
							l_csql = L_csql& "AND ((rf_ref LIKE '%"&trim(LotNumber)&"%') AND (fh_id LIKE '%"&trim(JobNumber)&"')) "
						End if							
						GenericSortBy="fh_ship_dt desc"
						If SortBy>"" then
							l_csql = L_csql& " ORDER BY "&Sortby
							else
							l_csql = L_csql& " ORDER BY "&GenericSortby
						End if

					
			'response.write(" 583 tracking Query3:" & l_cSQL & "<br>")
			''''''''''''''''''''''''''''''''''''''''
				RSEVENTS2.Open l_cSQL, Database, 1, 3
				If RSEVENTS2.eof then
						ErrorMessage="No jobs were found that match your criteria.<br><br>Please check your criteria and try again."	
				End if				
				Do while not RSEVENTS2.EOF 
					fh_id=RSEVENTS2("fh_id")
					'fl_sf_rta=RSEVENTS2("fl_sf_rta")
					fh_status=RSEVENTS2("fh_status")
					fh_ship_dt=RSEVENTS2("fh_ship_dt")
					fl_t_acc=trim(RSEVENTS2("fl_t_acc"))
					fl_sf_id=trim(RSEVENTS2("fl_sf_id"))
					fl_st_id=RSEVENTS2("fl_st_id")					
					fl_sf_name=RSEVENTS2("fl_sf_name")
					fl_st_name=RSEVENTS2("fl_st_name")
					fl_t_atp=RSEVENTS2("fl_t_atp")
					'Response.Write "i="&i&"<BR>"
					'Response.Write "fl_t_atp="&fl_t_atp&"<BR>"
					'''''If ii=0 then
					    FirstONB=fl_t_atp
					''''''End if  
					fl_t_atd=RSEVENTS2("fl_t_atd")
					'fl_pod=RSEVENTS2("fl_pod")
					'fh_custpo=RSEVENTS2("fh_custpo")
					fh_priority=RSEVENTS2("fh_priority")
					'fl_sf_rta=RSEVENTS2("fl_sf_rta")
					'fl_finalDestination=RSEVENTS2("fl_finalDestination")
					If USESLOTS=TRUE then
						rf_ref=RSEVENTS2("rf_ref")
						'PODDateTime=RSEVENTS2("PODDateTime")
					End if
                    rf_status=RSEVENTS2("fh_status")
			Select Case fl_sf_id
				CASE "55"
					Fl_sf_id="CPGP"
				CASE "72"
					Fl_sf_id="CRI"					
			End Select
			Select Case fh_priority
				Case "WF", "CS"
					Displayfh_Priority="Standard"
				Case "XP"
					Displayfh_Priority="Expedited"					
				Case "AS"
					Displayfh_Priority="Next Day"
				Case "A0"
					Displayfh_Priority="Hot Shot"
				Case "A1"
					Displayfh_Priority="Same Day"															
				Case ELSE
					DisplayFH_Priority=FH_Priority
			End Select
			Select Case fh_status
				Case "RAP"
					Display_fh_status="Booked"			
				Case "CLS"
					Display_fh_status="Closed"
				Case "OPN"
					Display_fh_status="Open"
				Case "ACC"
					Display_fh_status="Accepted"
				Case "PUO"
					Display_fh_status="POB"					
				Case "ONB"
					Display_fh_status="On Board"
				Case "ATD"
					Display_fh_status="At Destination"
				Case "CAN"
					Display_fh_status="Cancelled"
				Case "DEL"
					Display_fh_status="Deleted"	
				Case "ARV"
					Display_fh_status="Arrived At HUB"
				Case "AC2"
					Display_fh_status="Arrived At HUB*"	
				Case "DPV"
					Display_fh_status="Departed HUB"											
				Case Else
					Display_fh_status=fh_status																			
			End Select
            'Response.write "fh_status="&fh_status&"<BR>"
            If rf_status="X" then Display_FH_Status="Cancelled" end if
			if fh_ship_dt="1/1/1900" then fh_ship_dt="&nbsp;" end if
			if FirstONB="1/1/1900" then FirstONB="&nbsp;" end if
			if fl_t_atd="1/1/1900" then fl_t_atd="&nbsp;" end if
			if fl_t_acc="1/1/1900" then fl_t_acc="&nbsp;" end if
			'If ErrorMessage="" then
			'Response.Write "UsesLots="&UsesLots&"<BR>"
			'Response.Write "("&fh_id&")fl_finaldestination="&fl_finaldestination&"****<BR>"
			'If (trim(fl_st_id)=trim(fl_finaldestination)) Or (isnull(fl_finaldestination)) then
			%>
				<tr>

 						        <td class="MainPageTextSmaller" nowrap valign="top" bgcolor="<%=colorset(i mod numcolors)%>">
							        <a href="../reporting/OrderDetails.asp?inputlotnumber=<%=trim(rf_ref)%>&inputjobnumber=<%=fh_id%>"><%=fh_id%></a>
						        </td>				
              
          
					<!--				
					<td class="MainPageTextSmaller" valign="top">
						<%=fl_sf_name%>
					</td>	
					<td class="MainPageTextSmaller" valign="top">
						<%=fl_st_name%>
					</td>
					-->
					<td class="MainPageTextSmaller" nowrap valign="top" bgcolor="<%=colorset(i mod numcolors)%>">
						<%=fl_sf_name%>
						<%If (trim(fl_st_id)<>trim(fl_pod) AND trim(fl_pod)>"") then response.Write "<br><font color='red'>DISCREPANCY</font>" end if%>
					</td>
					<td class="MainPageTextSmaller" nowrap valign="top" bgcolor="<%=colorset(i mod numcolors)%>">
						<%=fl_st_name%>
					</td>					
					<td class="MainPageTextSmaller" nowrap valign="top" bgcolor="<%=colorset(i mod numcolors)%>">
						<%=Display_fh_status%>
					</td>
					<%If UsesLots=TRUE then%>
					<td class="MainPageTextSmaller" nowrap valign="top" bgcolor="<%=colorset(i mod numcolors)%>">
						<%=Displayfh_priority%>
					</td>					
					<%end if%>
					<td class="MainPageTextSmaller" nowrap valign="top" bgcolor="<%=colorset(i mod numcolors)%>">
						<%=fh_ship_dt%>
					</td>	
					<td class="MainPageTextSmaller" nowrap valign="top" bgcolor="<%=colorset(i mod numcolors)%>">
						<%
						'Response.Write "Fl_sf_id="&Fl_sf_id&"<BR>"
						If fl_sf_id="CPGP" then
							if trim(fh_custpo)>"" then
								If firstONB<>"&nbsp;" then
									%>
									<a href="http://www.quickonline.com/cgi-bin/WebObjects/BOLSearch?bolNumber=<%=fh_custpo%>" target="_blank"><%=firstONB%></a>
									<%
									'Response.Write fl_t_atp&" ("&fh_custpo&")"
									else
									%>
									<a href="http://www.quickonline.com/cgi-bin/WebObjects/BOLSearch?bolNumber=<%=fh_custpo%>" target="_blank"><%=fl_t_acc%></a>
									<%									
									'Response.Write fl_t_acc&" ("&fh_custpo&")"
								End if
								else
								Response.Write FirstONB
							End if
							else
							Response.Write FirstONB
						End if
						%>
					</td>				
					<td class="MainPageTextSmaller" nowrap valign="top" bgcolor="<%=colorset(i mod numcolors)%>">
						<%

						'Response.Write "BILLTOID="&BILLTOID&"<BR>"

						If trim(BILLTOID)="48" then
						
							Set RSEVENTS22 = Server.CreateObject("ADODB.Recordset")
								RSEVENTS22.CursorLocation = 3
								RSEVENTS22.CursorType = 3
								'response.Write "Liberty="&Liberty&"<BR>"
								RSEVENTS22.ActiveConnection = LIBERTY
								l_csql = "SELECT * FROM F_HAWB_DATA WHERE (SZF1='"& rf_ref &"')"
								'Response.write("Query:" & l_cSQL)
								RSEVENTS22.Open l_cSQL, LIBERTY, 1, 3
								If not RSEVENTS22.EOF then
									ULID=RSEVENTS22("ULID")
									HexULID=Hex(ULID)
									else
									ULID=""
								End if
								RSEVENTS22.close
							Set RSEVENTS22 = Nothing
						End if



						
						If trim(ULID)>"" then
							%>
							<!--
							<a href="../KWEPODS/<%=trim(rf_ref)%>.pdf" target="_blank"><%=fl_t_atd%></a>
							-->
							<a href="http://document.logisticorp.us:8080/LibertyIMS::/User=WebUser;pwd=Internet42;sys=LogistiCorp/Cmd%3DGetRawDocument%3BFolder%3D%2321%3BDoc%3D<%=HexULID%>%3Bformat%3DLIC/" target="_blank"><%=fl_t_atd%></a>&nbsp;
							<%
							else
							If isdate(PODDateTime) then
								%>
								<a href="../KWEPODS/<%=trim(rf_ref)%>.pdf" target="_blank"><%=fl_t_atd%></a>
								<%
								else
							%>
							<%=fl_t_atd%>
							<%
							End if
						End if
						If trim(fl_pod)>"" then response.Write " to "&fl_pod end if
						%>						
					</td>																						
				</tr>

<%
				i=i+1
				'END IF
				ii=ii+1
				RSEVENTS2.movenext
				LOOP
				RSEVENTS2.close
			Set RSEVENTS2 = Nothing
			'Response.Write "i="&i&"<BR>"
					'Response.Write "BillToID="&BillToID&"<BR>"
					'Response.Write "fh_custpo="&fh_custpo&"<BR>"
					'Response.Write "rf_ref="&rf_ref&"<BR>"
					'Response.Write "fh_id="&fh_id&"<BR>"			
			If i>0 then
				If i>1 then
					PluralResults="s"
					else
					'Response.Write "BillToID="&BillToID&"<BR>"
					'Response.Write "fh_custpo="&fh_custpo&"<BR>"
					'Response.Write "rf_ref="&rf_ref&"<BR>"
					'Response.Write "fh_id="&fh_id&"<BR>"
					Select Case BillToID
						Case 26
							'Response.Write "Redirect to SR?<BR>"
							Response.Redirect("../reporting/jobanalysis.asp?inputdocumentnumber="&fh_custpo)
						Case 36
							'Response.Write "Redirect to Wafer?<BR>"
							Response.Redirect("../reporting/OrderDetails.asp?inputlotnumber="&trim(rf_ref)&"&inputjobnumber="&fh_id)
						Case 38
							'Response.Write "Redirect to Reticle?<BR>"
							'Response.Redirect("../reporting/jobanalysis.asp?inputlotnumber="&fh_custpo&"&inputjobnumber="&fh_id)
							Response.Redirect("../reporting/OrderDetails.asp?inputjobnumber="&fh_id)																								
						Case 48, 13, 14, 25
							'Response.Write "Redirect to KWE?<BR>"
							Response.Redirect("../reporting/jobanalysis.asp?inputlotnumber="&rf_ref&"&inputjobnumber="&fh_id)
					End Select
					
				End if
				Response.Write "<tr><td align='left' class='miniheader' colspan='"&ColspanNumber&"'>"&i&" Result"&PluralResults
				If (i=20 and NumberofDays>0) or i=300 then
					Response.Write " - The maximum page display is 300 results.  There may be more results, please narrow your search criteria."
				end if
				Response.Write "</td></tr>"
			End if
			'Response.Write "ColspanNumber="&ColspanNumber&"<BR>"
	%>
			<tr><td align="center" class="miniheader" colspan="<%=ColspanNumber%>"><%=ErrorMessage%></td></tr>
			</table>	  
<%end if%>    
    
			
			
    
    
    </td></tr>


</table>

</td></tr>

</table>
</td></tr>
<tr><td colspan="<%=ColspanNumber%>">&nbsp;</td></tr>
<tr ><td height="100%">&nbsp;</td></tr>
<!-- </form>  -->

<tr><td>
<table width="100%" cellpadding=0 cellspacing=0>
<tr>
    <td height="100" class="FleetXGreySection" colspan="2">
        <!-- #include file="../BottomSection.asp" -->
    </td>
</tr>
<tr><td height="15" class="FleetXRedSectionSmall" colspan="2" align="center"><%=CopywriteNotice %></td></tr>
</table>
</td></tr></table>
<script src="../jquery-2.1.0.min.js"></script> 
<script src="../pickadate.js"></script> 
<script type="text/javascript">
    // PICKADATE FORMATTING
    $('#date_1').pickadate({
        format: 'mm/dd/yyyy', 	// Friendly format displayed to user
        formatSubmit: 'mm/dd/yyyy', // Actual format used by application
        hiddenName: false			// Allows two different formats
    });
    $('#date_2').pickadate({
        format: 'mm/dd/yyyy', 	// Friendly format displayed to user
        formatSubmit: 'mm/dd/yyyy', // Actual format used by application
        hiddenName: false			// Allows two different formats
    });

    $('#time_1').pickatime({
        format: 'h:i A', 		// Displayed and application format
        interval: 10, 			// Interval between values (in minutes)
        min: '12:00 AM', 			// Starting value
        max: '11:59 PM'				// Ending value
    });

</script>
</body>
</html>
