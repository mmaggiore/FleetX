<%@ Language=VBScript CodePage=65001 %>
<%  
	l_cSel = Request.Form("STSEL")
	
	' This is the DATE the user entered
	TheCurrentTime=Now()
	Submit=Request.Form("Submit")
	SQLDriverID=Request.Form("DriverID")
	NumberOfHours=request.Form("NumberOfHours")
	OrderType=Request.Form("OrderType")
    'Session("NumberOfHours")=""
	If NumberOfHours="" and Session("NumberOfHours")="" then 
		NumberOfHours=96
		Session("NumberOfHours")=96
		OrderType="open"
		Session("OrderType")=OrderType
    'rESPONSE.WRITE "LINE 16 NUMBER OF HOURS="&nUMBEROFHOURS&"<br>"
	End if

    'rESPONSE.WRITE "LINE 17 NUMBER OF HOURS="&nUMBEROFHOURS&"<br>"
	If Submit>"" then
		Session("NumberOfHours")=NumberOfHours
		Session("SQLDriverID")=SQLDriverID
		Session("OrderType")=OrderType
	End if
	OrderType=Session("OrderType")
	SQLDriverID=Session("SQLDriverID")	
	NumberOfHours=Session("NumberOfHours")
	%>
<html>
<head>

<!-- #include file="../fleetexpress.inc" -->
<link rel="stylesheet" type="text/css" href="../css/Style.css">
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
    PageTitle="REPORTS"

%>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<meta http-equiv="refresh" content="90" />
<title>FleetX - <%=PageTitle %></title>
</head>

<body onload="document.FindUser.requestorName.focus();">
	<table align="center" border="0" bordercolor="black" cellpadding="0" cellspacing="0" ID="Table1" height="100%" width="100%">
        <tr><td align="left"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>
        <tr>
            <td align="left"><img src="../images/FleetX_Small.jpg" height="50" width="168" /></td>
            <td align="right" valign="bottom"><!-- #include file="../topnavbar.asp" --></td>
        </tr>
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" /></td></tr>	
    <tr><td colspan="2">
<!-- <form action="NewUser.asp" method="post" name="FindUser">   -->
<table border="0" bordercolor="green" Cellspacing="0" Cellpadding="0" align="left" width="100%">
 <tr><td align="left" colspan="2" bgcolor="<%=HeaderBorderColor%>"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>

        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetXRedSection" colspan="2"><%=PageTitle%></td></tr>

        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
        <!--tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressBodyWhite" colspan="2">In order to reach the Fleet Express Order page, please correctly type in the green verification code in the supplied text box and click "Submit."</td></tr-->
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
<tr><td>
<table  border="0" bordercolor="blue" align="center" class="MainPageText" width="100%">
	<tr height="20">
		<td width="650">&nbsp;</td>
	</tr>



    <tr><td align=center width="100%"><!-- main page stuff goes here! -->
    
    	<CENTER>
	<%
	'Response.Write("<H2><FONT COLOR='red'>This page is currently being worked on.<br>  It may become unavailable for brief periods.<br>  Thanks,<br>  Mark</FONT></H2>")
	Response.Write("<H2><FONT COLOR=#d71e26>LIVE SHIPMENT MONITOR OF LAST "&NumberOfHours&" HOURS FROM "&TheCurrentTime&"</FONT></H2>")
	%>

	<form action="FleetXLiveMonitor.asp" method="POST" ID="Form1">
	<table border="0" bordercolor="blue" cellpadding="0" cellspacing="0" ID="Table1">
	<tr>
	<td>
	<table border="0" bordercolor="red" width="400" ID="Table2">
	<form name="GetJobParms" action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="POST" ID="Form2">
			<%
			suid=Trim(Session("suid"))
			'l_cSQL="SELECT  distinct(fclegs.fl_dr_id) FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id WHERE (fh_bt_id='"&suid&"') ORDER BY fl_dr_id"
			'Response.Write "l_cSQL="&l_cSQL&"<BR>"           
            %>
			<tr>
				<td class="generalcontent">
                <!--input type="submit" value="Submit/Refresh" name="submit" class="buttongrey" ID="Submit1"-->
                <input id="gobutton" name="submit" type="submit" value="Submit/Refresh" /></td>
				<td class="generalcontent">&nbsp;</td>				
				<td class="generalcontent">
					<select name="OrderType" ID="Select2">
						<option value="open" <%If OrderType="open" then response.Write " selected" end if%>>Open Orders</option>
						<option value="all" <%If OrderType="all" then response.Write " selected" end if%>>All Orders</option>
					</select>           
				</td>
				<%If morethanonestockroomtruck="yes" then%>			
				<td class="generalcontent">&nbsp;</td>
				<td class="generalcontent">
					<select name="DriverID" ID="Select3">
					<option value="">All Vehicles</option>
					<%    
			            
				Set oRs21 = Server.CreateObject("ADODB.Recordset")
				oRs21.CursorLocation = 3
				oRs21.CursorType = 3
				oRs21.ActiveConnection = DATABASE	

					'On Error Resume Next                                                  
					Err.Clear
					

					oRs21.Open l_cSQL, DATABASE, 1, 3
					
					If Err.Number <> 0 Then                                               
					Response.Write "Error Executing the query.  Error:" & Err.Description
					Else
						IF NOT oRs21.EOF THEN
							oRs21.MoveFirst

							DO WHILE NOT oRs21.EOF
							ShowDriverID=trim(oRs21("fl_dr_id"))
							If ShowDriverID>"" and Left(ShowDriverID,1)>"." AND  ShowDriverID<>"111" AND ShowDriverID<>"EFRANCO" then			

							%>          
							<option value="<%=ShowDriverID%>" <%if ShowDriverID=SQLDriverID then response.Write " selected" end if%>>Vehicle #<%=ShowDriverID%></option>
							<%				
							
							End if
							oRs21.MoveNext
							LOOP
						End if
					End If
					oRs21.close
					Set oRs21=Nothing
					%> 
					</select>           
				</td>
				<%end if%>
				<td class="generalcontent">&nbsp;</td>				
				<td class="generalcontent" width="70%">
					<select name="NumberOfHours" ID="Select1">

						<option value="3" <%If NumberOfHours=3 then response.Write " selected" end if%>>3 Hours</option>
						<option value="6" <%If NumberOfHours=6 then response.Write " selected" end if%>>6 Hours</option>
						<option value="12" <%If NumberOfHours=12 then response.Write " selected" end if%>>12 Hours</option>
						<option value="18" <%If NumberOfHours=18 then response.Write " selected" end if%>>18 Hours</option>
						<option value="24" <%If NumberOfHours=24 then response.Write " selected" end if%>>24 Hours</option>
						<option value="48" <%If NumberOfHours=48 then response.Write " selected" end if%>>48 Hours</option>
						<option value="72" <%If NumberOfHours=72 then response.Write " selected" end if%>>72 Hours</option>
						<option value="96" <%If NumberOfHours=96 then response.Write " selected" end if%>>96 Hours</option>
					</select>           
				</td>
			</tr>
            <tr><td>&nbsp;</td></tr>
			<tr>
				<td colspan="7">
					<table cellpadding="0" cellspacing="0" border="0" bordercolor="brown" ID="Table3">
						<tr>
							<!--
							<td bgcolor="#3120FF" class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;</td><td nowrap class="generalcontent">&nbsp;Pick</td>
							<td class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
							-->	
							<!------------------------------------------------------------------------>	
							<!------------------------------------------------------------------------>	
							<!------------------------------------------------------------------------>	
							<!------------------------------------------------------------------------>
							<td bgcolor="#94FAA2" class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;</td><td nowrap class="generalcontent">&nbsp;Dispatched</td>
							<td class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>                            					
							<td bgcolor="#FBB5B5" class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;</td><td nowrap class="generalcontent">&nbsp;Acknowledged</td>
							<td class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
						<!--	<td bgcolor="#C1C1F9" class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;</td><td nowrap class="generalcontent">&nbsp;POB</td>
							<td class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
							<td bgcolor="#C8F9C1" class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;</td><td nowrap class="generalcontent" nowrap>&nbsp;Arrived at Airline</td>
							<td class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>	-->										
							<td bgcolor="#F8DAA1" class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;</td><td nowrap class="generalcontent">&nbsp;On Board</td>
							<td class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
							<td bgcolor="#A1F7F8" class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;</td><td nowrap class="generalcontent">&nbsp;Delivered</td>
                            <!--
							<td class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
							<td nowrap class="generalcontent" height="5"><img src="../images/bluedot.gif" height="10" width="10" border="0"></td><td nowrap class="generalcontent">&nbsp;POD Scanned</td>
							-->
                            <!--
							<td bgcolor="red" class="generalcontent"><img src="../images/legendstripes.gif" height="14" width="17"></td><td nowrap class="generalcontent">&nbsp;Time Since Last Status Change</td>						
							-->
						</tr>
					</table>
				</td>				
				
				
			</tr> 
			<!--END--> 
            </table>
      </td>
    </TR>
    <tr>
		<td>
		
		
		
			<table border="1" cellpadding="0" cellspacing="0" ID="Table4" bordercolor="black">
			<tr>
				<td class="FleetXRedSectionLM"  nowrap="nowrap" >&nbsp;&nbsp;ORDER Number&nbsp;&nbsp;</td>
                <td class="FleetXRedSectionLM"  nowrap="nowrap" >&nbsp;&nbsp;Vehicle&nbsp;&nbsp;</td>
                <td class="FleetXRedSectionLM"  nowrap="nowrap" >&nbsp;&nbsp;Driver&nbsp;&nbsp;</td>
				<td class="FleetXRedSectionLM"  nowrap="nowrap" width="100">&nbsp;&nbsp;Dispatched&nbsp;&nbsp;</td>
				<td class="FleetXRedSectionLM"  nowrap="nowrap" width="100">&nbsp;&nbsp;Acknowledged&nbsp;&nbsp;</td>
		<!--		<td class="FleetXRedSectionLM" nowrap width="100">&nbsp;&nbsp;Paper on Board&nbsp;&nbsp;</td>
				<td class="FleetXRedSectionLM" nowrap width="100">&nbsp;&nbsp;Arrived at Airline&nbsp;&nbsp;</td>   -->
				<td class="FleetXRedSectionLM"  nowrap="nowrap" width="100">&nbsp;&nbsp;On Board&nbsp;&nbsp;</td>
				<td class="FleetXRedSectionLM"  nowrap="nowrap" width="100">&nbsp;&nbsp;Delivered&nbsp;&nbsp;</td>
				<td class="FleetXRedSectionLM"  nowrap="nowrap" width="100">&nbsp;&nbsp;Exceptions&nbsp;&nbsp;</td>
		<!--		<td class="FleetXRedSectionLM" nowrap width="100">&nbsp;&nbsp;EDI Sent&nbsp;&nbsp;</td>  -->
				<td class="FleetXRedSectionLM" nowrap="nowrap">&nbsp;&nbsp;Time Remaining&nbsp;&nbsp;</td>
			</tr>
			<% 
					l_cSQL="SELECT * from LiveMonitor "
					'''l_cSQL="SELECT fcfgthd.fh_bt_id AS Customer, fcfgthd.fh_id AS JobNo, "
					'''l_cSQL=l_cSQL&"fcfgthd.fh_ship_dt AS BookTime, fclegs.fl_PKey AS fl_Pkey, fclegs.fl_sf_rta AS StartTime, "
					'''l_cSQL=l_cSQL&"fclegs.fl_dr_id AS DriverID, fclegs.fl_t_disp AS DispatchTime, "
					'''l_cSQL=l_cSQL&"fclegs.fl_t_acc AS AcknowledgeTime, fclegs.fl_t_int AS POBTime, "
					'''l_cSQL=l_cSQL&"fclegs.fl_t_und AS AtAirlineTime, fclegs.fl_t_atp AS OnBoardTime, "
					'''l_cSQL=l_cSQL&"fclegs.fl_t_atd AS DeliveryTime, fcfgthd.fh_priority AS Priority, "
					'''l_cSQL=l_cSQL&"fcfgthd.fh_status AS Status, fclegs.fl_sf_id AS FromLocationID, "
					'''l_cSQL=l_cSQL&"fclegs.fl_sf_name AS FromLocation, fclegs.fl_st_id AS ToLocationID, "
					'''l_cSQL=l_cSQL&"fclegs.fl_st_name AS ToLocation, fclegs.fl_job_closed AS fl_Job_Closed, fcrefs.rf_ref AS HAWB, fcrefs.EDI_DateTime AS EDI_DateTime "
					'l_cSQL=l_cSQL&"FCJobExceptions.ExceptionID AS ExceptionID "
					'''l_cSQL=l_cSQL&"FROM fcfgthd INNER JOIN "
                    '''l_cSQL=l_cSQL&"fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN "
                    '''l_cSQL=l_cSQL&"fcrefs ON fclegs.fl_fh_id = fcrefs.rf_fh_id "
                    'LEFT OUTER JOIN "
                    'l_cSQL=l_cSQL&"FCJobExceptions ON fcrefs.rf_fh_id = FCJobExceptions.fh_id "			
					'l_cSQL="SELECT fcfgthd.fh_bt_id AS Customer, fcfgthd.fh_id AS JobNo, "
					'l_cSQL=l_cSQL&"fcfgthd.fh_ship_dt AS BookTime, fclegs.fl_sf_rta AS StartTime, "
					'l_cSQL=l_cSQL&"fclegs.fl_dr_id AS DriverID, fclegs.fl_t_disp AS DispatchTime, "
					'l_cSQL=l_cSQL&"fclegs.fl_t_acc AS AcknowledgeTime, fclegs.fl_t_int AS POBTime, "
					'l_cSQL=l_cSQL&"fclegs.fl_t_und AS AtAirlineTime, fclegs.fl_t_atp AS OnBoardTime, "
					'l_cSQL=l_cSQL&"fclegs.fl_t_atd AS DeliveryTime, fcfgthd.fh_priority AS Priority, "
					'l_cSQL=l_cSQL&"fcprior.fp_rtd_om AS PriorityTime, fcfgthd.fh_status AS Status, "
					'l_cSQL=l_cSQL&"fclegs.fl_sf_ID AS FromLocationID, fclegs.fl_sf_name AS FromLocation, "
					'l_cSQL=l_cSQL&"fclegs.fl_st_ID AS ToLocationID, fclegs.fl_st_name AS ToLocation "
					'l_cSQL=l_cSQL&"FROM fcfgthd INNER JOIN "
					'l_cSQL=l_cSQL&"fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN "
					'l_cSQL=l_cSQL&"fcprior ON fcfgthd.fh_priority = fcprior.fp_id "
					'l_cSQL=l_cSQL&"WHERE ((fcfgthd.fh_bt_id = '"&suid&"') AND (fcfgthd.fh_ship_dt >= '"&(TheCurrentTime-7)&"') AND (fclegs.fl_dr_id <> '111') AND (fclegs.fl_dr_id <> 'EFRANCO') AND"
					l_cSQL=l_cSQL&"WHERE ("
					If SQLDriverID>"" then
						l_cSQL=l_cSQL&" (DriverID = '"&SQLDriverID&"') AND "
					End if
                    'Response.write "line 272 ordertype="&OrderType&"<BR>"
					If trim(lcase(OrderType))="open" then
						l_cSQL=l_cSQL&" (status <> 'CLS') AND "
					End if					
					If NumberOfHours=3 then
						l_cSQL=l_cSQL&"(BookTime >= '"&(TheCurrentTime-.125)&"') "
					End if
					If NumberOfHours=6 then
						l_cSQL=l_cSQL&"(BookTime >= '"&(TheCurrentTime-.25)&"') "
					End if										
					If NumberOfHours=12 then
						l_cSQL=l_cSQL&"(BookTime >= '"&(TheCurrentTime-.5)&"') "
					End if
					If NumberOfHours=18 then
						l_cSQL=l_cSQL&"(BookTime >= '"&(TheCurrentTime-.75)&"') "
					End if	
					If NumberOfHours=24 then
						l_cSQL=l_cSQL&"(BookTime >= '"&(TheCurrentTime-1)&"') "
					End if	
					If NumberOfHours=48 then
						l_cSQL=l_cSQL&"(BookTime >= '"&(TheCurrentTime-2)&"') "
					End if
					If NumberOfHours=72 then
						l_cSQL=l_cSQL&"(BookTime >= '"&(TheCurrentTime-3)&"') "
					End if
					If NumberOfHours=96 then
						l_cSQL=l_cSQL&"(BookTime >= '"&(TheCurrentTime-4)&"') "
					End if	
					l_cSQL=l_cSQL&" OR ((DeliveryTime='1/1/1900') AND (fl_job_closed='1/1/1900') )) "																					
					l_cSQL=l_cSQL&" AND ((status <> 'CAN') AND (status <> 'DEL') "
					'l_cSQL=l_cSQL&" AND (Customer = '"&suid&"') AND (DriverID <> '111') AND (DriverID <> 'EFRANCO') "
					'l_cSQL=l_cSQL&" AND (DriverID <> '111') AND (DriverID <> 'EFRANCO') "
					If SQLDriverID>"" then
						l_cSQL=l_cSQL&" AND (DriverID = '"&SQLDriverID&"') "
					End if						
					l_cSQL=l_cSQL&") "
					'If OrderType="open" then
					'	l_cSQL=l_cSQL&"OR (fcfgthd.fh_status <> 'CLS') AND (fcfgthd.fh_status <> 'CAN') AND (fcfgthd.fh_status <> 'DEL') "
					'	else
						'l_cSQL=l_cSQL&"OR (fcfgthd.fh_status <> 'CLS') "
					'End if
					l_cSQL=l_cSQL&"Order By JobNo Desc, HAWB, fl_pkey asc "
  					
  					'Response.Write "304 vehkwe l_cSQL="&l_cSQL&"<BR>"
  					
				Set oRs22 = Server.CreateObject("ADODB.Recordset")
				oRs22.CursorLocation = 3
				oRs22.CursorType = 3
				oRs22.ActiveConnection = DATABASE	
					Err.Clear
					oRs22.Open l_cSQL, DATABASE, 1, 3
					If Err.Number <> 0 Then                                               
					Response.Write "Error Executing the query.  Error:" & Err.Description
					Else
						IF NOT oRs22.EOF THEN
							oRs22.MoveFirst
							DO WHILE NOT oRs22.EOF
							JobNo=trim(oRs22("JobNo"))
							BookTime=trim(oRs22("BookTime"))
                            DispatchTime = trim(oRs22("DispatchTime"))
                            'Response.Write "Line 336 BookTime="&BookTime&"***DispatchTime="&DispatchTime&"/jobno="&jobno&"<BR>"
							fl_PKey=trim(oRs22("fl_PKey"))
							StartTime=trim(oRs22("StartTime"))
							AcknowledgeTime=trim(oRs22("AcknowledgeTime"))
							If HAWB="590012279104" then
								'Response.Write "***fl_pkey="&fl_pkey&"<BR>"
								'Response.Write "***booktime="&BookTime&"<BR>"
								'Response.Write "***AcknowledgeTime="&AcknowledgeTime&"<BR>"
							End if							
							POBTime=trim(oRs22("POBTime"))
							AtAirlineTime=trim(oRs22("AtAirlineTime"))
							OnBoardTime=trim(oRs22("OnBoardTime"))
							DeliveryTime=trim(oRs22("DeliveryTime"))
                            
              'response.write "333 vehkwe deliverytime=" & DeliveryTime & "<br>"
							Priority=trim(oRs22("Priority"))
							Fl_Job_Closed=oRs22("Fl_Job_Closed")
							'DeliveryTime=Fl_Job_Closed
							Status=ucase(trim(oRs22("Status")))
							HAWB=trim(oRs22("HAWB"))
							EDI_DateTime=trim(oRs22("EDI_DateTime"))
                            'HAWBNUMBER=trim(oRs22("HAWBNUMBER"))
							If isdate(EDI_DateTime) then
								EDI_DateTimeMonth=Month(EDI_DateTime)
								EDI_DateTimeDay=Day(EDI_DateTime)
								EDI_DateTimeTime=FormatDateTime(EDI_DateTime,4)
								DisplayEDI_DateTime=EDI_DateTimeMonth & "/" & EDI_DateTimeDay & " " & EDI_DateTimeTime
								else
								DisplayEDI_DateTime="&nbsp;"
							End if

							PODDateTime=trim(oRs22("PODDateTime"))
							'ExceptionID=trim(oRs22("ExceptionID"))
							DueTime=trim(oRs22("DueTime"))
							VehicleID=trim(oRs22("VehicleID"))
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                     	Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
                    		RSEVENTS.CursorLocation = 3
                    		RSEVENTS.CursorType = 3
                    		RSEVENTS.ActiveConnection = DATABASE
                   		  SQL = "SELECT AvailableVehicles.VehicleName AS VehicleName, lcintranet.dbo.Intranet_Users.FirstName AS DriverFirstName, lcintranet.dbo.Intranet_Users.LastName AS DriverLastName FROM AvailableVehicles INNER JOIN lcintranet.dbo.Intranet_Users ON AvailableVehicles.DriverID = lcintranet.dbo.Intranet_Users.UserID WHERE VehicleID = '" & VehicleID & "' and AvailableStatus = 'c'"
                    		'Response.Write "632 SQL="&SQL&"<BR>"
                    		RSEVENTS.Open SQL, DATABASE, 1, 3
                            If NOT RSEVENTS.EOF then
                                VehicleName = RSEVENTS("VehicleName")
                                DriverFirstName = RSEVENTS("DriverFirstName")
                                DriverLastName = RSEVENTS("DriverLastName")
                            End if
                        'response.write "636 couriertype=" & couriertype & "<br>"
                        RSEVENTS.close
                    	  Set RSEVENTS = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''








							If jobNo<>tempjobno then
							
							
							If trim(Priority)<>"P0" then
								LinkClass="tableinnercontent"
								else
								LinkClass="redinnercontent"
							End if
							If Status="DEL" or Status="CLS" or Status="CAN" then
								ShowDots="y"
								ELSE
								ShowDots="n"
							End if
							If Status="DEL" then Status="DELETED" END IF
							If Status="CAN" then Status="CANCELLED" END IF
							'BookTimeEllapsed=DateDiff("n",StartTime,BookTime)
							'PickTimeEllapsed=DateDiff("n",StartTime,BookTime)
							'If abs(PickTimeEllapsed)>99999 then PickTimeEllapsed=0 End if
							If HAWB="590012279104" then
								'Response.Write "booktime="&BookTime&"<BR>"
								'Response.Write "AcknowledgeTime="&AcknowledgeTime&"<BR>"
							End if
              
                            'response.write "378 dispatchtime=" & dispatchtime & "<br>"
                            If dispatchtime="1/1/1900" then dispatchtime=now() end if
							DispatchTimeEllapsed=DateDiff("n",BookTime,DispatchTime)
							If DispatchTimeEllapsed=0 then DispatchTimeEllapsed=1 end if
							'If JobNo="00101848" then
								'response.write "BookTime="&BookTime&"<BR>"
								'response.write "AcknowledgeTime="&AcknowledgeTime&"<BR>"
								'response.write "AcknowledgeTimeEllapsed="&AcknowledgeTimeEllapsed&"<BR>"
							'End if							
							If abs(DispatchTimeEllapsed)>99999 then DispatchTimeEllapsed=0 End if	
							TotalTimeEllapsed=DispatchTimeEllapsed
							If DispatchTimeEllapsed>59 then
								temphours=fix(DispatchTimeEllapsed/60)
								tempminutes=DispatchTimeEllapsed-(temphours*60)
								DispatchTimeEllapsed=Temphours&" hrs "&TempMinutes&" mins"
								else
								DispatchTimeEllapsed=DispatchTimeEllapsed&" mins"
							End if						
							''''''''''''''''''''''''''''''''''''''''''''
                            '''''If showacknowledge="y" then

                                'response.write "showacknowledge="&showacknowledge&"<BR>"
                                '''If AcknowledgeTime="1/1/1900" and showacknowledge="y" then AcknowledgeTime=Now() end if
                                If AcknowledgeTime="1/1/1900" then 
                                    AcknowledgeTime=Now() 
                                    ShowAcknowledgeColor="n"
                                    showOnBoard="n"
                                    else
                                    ShowAcknowledgeColor="y"
                                    showOnBoard="y"
                                end if
                                'response.write "AcknowledgeTime="&AcknowledgeTime&"<BR>"
                                'response.write "DispatchTime="&DispatchTime&"<BR>"
							    AcknowledgeTimeEllapsed=DateDiff("n",DispatchTime,AcknowledgeTime)
                                'response.write "Line 419 AcknowledgeTimeEllapsed="&AcknowledgeTimeEllapsed&"<BR>"
							    If AcknowledgeTimeEllapsed=0 and AcknowledgeTime>"1/1/1900" and DispatchTime>"1/1/1900" then AcknowledgeTimeEllapsed=1 end if
							    'If AcknowledgeTimeEllapsed=0 then AcknowledgeTimeEllapsed=1 end if
							    'If JobNo="00101848" then
								    'response.write "BookTime="&BookTime&"<BR>"
								    'response.write "AcknowledgeTime="&AcknowledgeTime&"<BR>"
								    'response.write "AcknowledgeTimeEllapsed="&AcknowledgeTimeEllapsed&"<BR>"
							    'End if							
							    If abs(AcknowledgeTimeEllapsed)>99999 then AcknowledgeTimeEllapsed=0 End if	
							    TotalTimeEllapsed=AcknowledgeTimeEllapsed
							    If AcknowledgeTimeEllapsed>59 then
								    temphours=fix(AcknowledgeTimeEllapsed/60)
								    tempminutes=AcknowledgeTimeEllapsed-(temphours*60)
								    AcknowledgeTimeEllapsed=Temphours&" hrs "&TempMinutes&" mins"
								    else
								    AcknowledgeTimeEllapsed=AcknowledgeTimeEllapsed&" mins"
							    End if	
                           ''''' End if					
							''''''''''''''''''''''''''''''''''''''''''''
                            'Response.write "Line #445 showOnBoard="&showOnBoard&"<BR>"
							If showOnBoard="y" then
                                'Response.write "Line #446 OnBoardTime="&OnBoardTime&"<BR>"
                                If OnBoardTime="1/1/1900" then 
                                    OnboardTime=Now() 
                                    ShowOnboardColor="n"
                                    else
                                    ShowOnBoardColor="y"
                                end if
							    OnBoardTimeEllapsed=DateDiff("n",AcknowledgeTime,OnBoardTime)
								    'response.write "OnBoardTime="&OnBoardTime&"<BR>"
								    'response.write "AcknowledgeTime="&AcknowledgeTime&"<BR>"
								    'response.write "OnBoardTimeEllapsed="&OnBoardTimeEllapsed&"<BR>"
							    If OnBoardTimeEllapsed=0 and AcknowledgeTime>"1/1/1900" and OnBoardTime>"1/1/1900" then OnBoardTimeEllapsed=1 end if
							    If JobNo="00077112x" then
								    response.write "OnBoardTimeEllapsed="&OnBoardTimeEllapsed&"<BR>"
							    End if							
							    If abs(OnBoardTimeEllapsed)>99999 then OnBoardTimeEllapsed=0 End if	
							    TotalTimeEllapsed=TotalTimeEllapsed+OnBoardTimeEllapsed
							    If OnBoardTimeEllapsed>59 then
								    temphours=fix(OnBoardTimeEllapsed/60)
								    tempminutes=OnBoardTimeEllapsed-(temphours*60)
								    OnBoardTimeEllapsed=Temphours&" hrs "&TempMinutes&" mins"
								    else
								    OnBoardTimeEllapsed=OnBoardTimeEllapsed&" mins"
							    End if		
                            End if	
                            'Response.write "line 499 status="&Status&"<BR>"					
							If status="CLS" or status="ONB" then
							    DeliveryTimeEllapsed=DateDiff("n",OnBoardTime,DeliveryTime)
								    'response.write "445 DeliveryTimeEllapsed="&DeliveryTimeEllapsed&"<BR>"
								    'response.write "DeliveryTime="&DeliveryTime&"<BR>"
								    'response.write "DeliveryTimeEllapsed="&DeliveryTimeEllapsed&"<BR>"
							    If DeliveryTimeEllapsed=0 and OnBoardTime>"1/1/1900" and DeliveryTime="1/1/1900" then DeliveryTimeEllapsed=1 end if
							    If JobNo="00077112x" then
								    response.write "OnBoardTime="&OnBoardTime&"<BR>"
								    response.write "DeliveryTime="&DeliveryTime&"<BR>"
								
								    response.write "DeliveryTimeEllapsed="&DeliveryTimeEllapsed&"<BR>"
							    End if							
							    If abs(DeliveryTimeEllapsed)>99999 then DeliveryTimeEllapsed=0 End if
							    TotalTimeEllapsed=TotalTimeEllapsed+DeliveryTimeEllapsed
							    If DeliveryTimeEllapsed>59 then
								    temphours=fix(DeliveryTimeEllapsed/60)
								    tempminutes=DeliveryTimeEllapsed-(temphours*60)
								    DeliveryTimeEllapsed=Temphours&" hrs "&TempMinutes&" mins"
								    else
								    DeliveryTimeEllapsed=DeliveryTimeEllapsed&" mins"
							    End if		
                            End if						
							
							CurrentOrderTime=DateDiff("n",BookTime,TheCurrentTime)	
							OrderClosed="n"	
							ShowDisplay="y"
							
	'If whatever="xxx" then						
							If DispatchTimeEllapsed="mins" then dispatchedTimeEllapsed="" end if
								'if DispatchTimeEllapsed<>"0 mins" then 
                                if DispatchTimeEllapsed<>""  then 
                                    'Response.write "line 478 status="&status&"/"&jobno&"/"&DispatchTime&"<BR>"
									IF (status="CLS" or status="DEL" or status="CAN" or status="ONB" or status="ACC" or status="OPN") then
                                        dispatchbgcolor="#94faa2"
                                        showacknowledge="y"
                                        'Response.write "Line 493 GOT HERE!!!!<BR>"
                                    End if
									DisplayDispatchTimeEllapsed=DispatchTimeEllapsed
									else
									dispatchbgcolor="white"
									DisplayDispatchTimeEllapsed="&nbsp;"
                                    showacknowledge="n"
								 end if
                    If showacknowledge="y" then
                                'REsponse.write "line 490 jobno="&jobno&"<BR>"
                               ' REsponse.write "line 491 AcknowledgeTimeEllapsed="&AcknowledgeTimeEllapsed&"<BR>"
								If trim(AcknowledgeTimeEllapsed)="mins" then AcknowledgeTimeEllapsed="" end if
                                'if AcknowledgeTimeEllapsed<>"0 mins" then 
                                '''''if AcknowledgeTimeEllapsed<>""  then 
                                If ShowAcknowledgeColor="y" then
									ackbgcolor="#FBB5B5"
									DisplayAcknowledgeTimeEllapsed=AcknowledgeTimeEllapsed
                                    showonboard="y"
									else
									ackbgcolor="white"
									DisplayAcknowledgeTimeEllapsed="&nbsp;"
                                    showonboard="n"
								 end if
                                 ''''''''''''''''''''''''''''''''''''''''
								If DisplayAcknowledgeTimeEllapsed="&nbsp;" then
									TimeNow=now()
									AcknowledgeTimeEllapsed=DateDiff("n",dispatchTime,TimeNow)
									TotalTimeEllapsed=TotalTimeEllapsed+AcknowledgeTimeEllapsed
									If AcknowledgeTimeEllapsed=0 then AcknowledgeTimeEllapsed=1 end if
									If JobNo="00077112x" then
										response.write "AcknowledgeTimeEllapsed="&AcknowledgeTimeEllapsed&"<BR>"
									End if							
									If abs(AcknowledgeTimeEllapsed)>99999 then AcknowledgeTimeEllapsed=0 End if	
									If AcknowledgeTimeEllapsed>59 then
										temphours=fix(AcknowledgeTimeEllapsed/60)
										tempminutes=AcknowledgeTimeEllapsed-(temphours*60)
										AcknowledgeTimeEllapsed=Temphours&" hrs "&TempMinutes&" mins"
										else
										AcknowledgeTimeEllapsed=AcknowledgeTimeEllapsed&" mins"
									End if
									DisplayAcknowledgeTimeEllapsed=AcknowledgeTimeEllapsed								
								End if
                                 ''''''''''''''''''''''''''''''''''''''''
                    End if
                    If showonboard="y" then
                                'Response.write "OnboardTimeEllapsed="&Onboardtimeellapsed&"<BR>"
                                'If trim(OnBoardTimeEllapsed)="mins" then response.write "HELLO!!!!<BR>" end if
                                If trim(OnBoardTimeEllapsed)="mins" then OnBoardTimeEllapsed="" end if
								'if OnBoardTimeEllapsed<>"0 mins" then 
                                '''''if OnBoardTimeEllapsed<>"" then 
                                If ShowOnboardColor="y" then
									onbbgcolor="#F8DAA1" 
                                    showdelivered="y"
									DisplayOnBoardTimeEllapsed=OnBoardTimeEllapsed
									else
									onbbgcolor="white" 
                                    showdelivered="n"
									DisplayOnBoardTimeEllapsed="&nbsp;"
									If DisplayAtAirlineTimeEllapsed<>"&nbsp;" and ShowDisplay<>"n" then
										DisplayOnBoardTimeEllapsed=DateDiff("n",AcknowledgeTime,now())
										TotalTimeEllapsed=TotalTimeEllapsed+DisplayOnBoardTimeEllapsed
										If DisplayOnBoardTimeEllapsed>59 then
											temphours=fix(DisplayOnBoardTimeEllapsed/60)
											tempminutes=DisplayOnBoardTimeEllapsed-(temphours*60)
											DisplayOnBoardTimeEllapsed=Temphours&" hrs "&TempMinutes&" mins"
											else
											DisplayOnBoardTimeEllapsed=DisplayOnBoardTimeEllapsed&" mins"
										End if										
										ShowDisplay="n"
									End if										
								end if
                        End if
                        If showdelivered="y" then
                                'REsponse.write "GOT HERE!  Line 607<BR>"
                                'Response.write "Line 608 DeliveryTimeEllapsed="&DeliveryTimeEllapsed&"<BR>"
                                If trim(DeliveryTimeEllapsed)="mins" then DeliveryTimeEllapsed="" end if
								'if DeliveryTimeEllapsed<>"0 mins" then 
                                if Status="CLS" then 
									dropbgcolor="#A1F7F8" 
									DisplayDeliveryTimeEllapsed=DeliveryTimeEllapsed
									else
									dropbgcolor="white" 
									DisplayDeliveryTimeEllapsed="&nbsp;"
									If DisplayOnBoardTimeEllapsed<>"&nbsp;" and ShowDisplay<>"n" then
										DisplayDeliveryTimeEllapsed=DateDiff("n",OnBoardTime,now())
										TotalTimeEllapsed=TotalTimeEllapsed+DisplayDeliveryTimeEllapsed
										If DisplayDeliveryTimeEllapsed>59 then
											temphours=fix(DisplayDeliveryTimeEllapsed/60)
											tempminutes=DisplayDeliveryTimeEllapsed-(temphours*60)
											DisplayDeliveryTimeEllapsed=Temphours&" hrs "&TempMinutes&" mins"
											else
											DisplayDeliveryTimeEllapsed=DisplayDeliveryTimeEllapsed&" mins"
										End if										
										ShowDisplay="n"
									End if										
								End if
                        End if
							'Response.Write "**********************<BR>"	
							'Response.Write "JobNo="&JobNo&"<BR>"
							'Response.Write "TempJobNo="&TempJobNo&"<BR>"
							'Response.Write "FL_PKey="&FL_PKey&"<BR>"
							'Response.Write "OnBoardTime="&OnBoardTime&"<BR>"
							'Response.Write "**********************<BR>"
        'End if
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''							
							
							%>
							<tr><td valign="middle" align="center"><a href="OrderDetails.asp?inputjobnumber=<%=JobNo%>" target="_blank" class="FleetXRedMain"><%=JobNo%></a> &nbsp; </td>
							<%
							If status="CANCELLED" or status="DELETED" then
								%>
								<td class="subheader">&nbsp;&nbsp;<%=status%></td>
								<td class="generalcontent" align="center" nowrap>&nbsp;</td>
								<td class="generalcontent" align="center" nowrap>&nbsp;</td>
								<td class="generalcontent" align="center" nowrap>&nbsp;</td>
								<td class="generalcontent" align="center" nowrap>&nbsp;</td>
							</tr>
								<%
								MinuteColor="white"
								else
								%>
                                <td height="30" valign="middle" class="subheader" align="center" nowrap="nowrap">&nbsp;<%=VehicleName%>&nbsp;</td>
                                <td height="30" valign="middle" class="subheader" align="center" nowrap="nowrap">&nbsp;<%=DriverFirstName%>&nbsp;&nbsp;<%=DriverLastName %>&nbsp;</td>

								<td height="30" valign="middle" class="subheader" align="center" nowrap bgcolor="<%=dispatchbgcolor%>"><%=DisplayDispatchTimeEllapsed%></td>
								<td height="30" valign="middle" class="subheader" align="center" nowrap bgcolor="<%=ackbgcolor%>"><%=DisplayAcknowledgeTimeEllapsed%></td>
						<!--		<td class="subheader" align="center" nowrap bgcolor="<%=pobbgcolor%>"><%=DisplayPOBTimeEllapsed%></td>
								<td class="subheader" align="center" nowrap bgcolor="<%=atairlinebgcolor%>"><%=DisplayAtAirlineTimeEllapsed%></td>   -->
								<td class="subheader" align="center" nowrap bgcolor="<%=onbbgcolor%>"><%=DisplayOnBoardTimeEllapsed%></td>
								<td class="subheader" align="center" nowrap bgcolor="<%=dropbgcolor%>"><%=DisplayDeliveryTimeEllapsed%></td>
								<td class="subheader" align="center" nowrap>&nbsp;
									<%
									L_cSQL="SELECT ExceptionID FROM FCJobExceptions WHERE (Ref_num = '"& HAWB &"') AND (fh_id = '"& JobNo &"')"
									Set oRs21 = Server.CreateObject("ADODB.Recordset")
									oRs21.CursorLocation = 3
									oRs21.CursorType = 3
									oRs21.ActiveConnection = DATABASE	
										'On Error Resume Next 
										'Response.Write "l_cSQL="&l_cSQL&"<BR>"                                                 
										Err.Clear
										oRs21.Open l_cSQL, DATABASE, 1, 3
										If Err.Number <> 0 Then                                               
										Response.Write "Error Executing the query.  Error:" & Err.Description
										Else
											IF NOT oRs21.EOF THEN
												oRs21.MoveFirst
												DO WHILE NOT oRs21.EOF
												If rrr>0 then response.Write ", " end if
												rrr=rrr+1
												ExceptionID=trim(oRs21("ExceptionID"))
												Response.Write ExceptionID&" "
												oRs21.MoveNext
												LOOP
											End if
										End If
										oRs21.close
										Set oRs21=Nothing
										rrr=0
                                        ''''''''''CHANGED CODE BELOW WHEN THEY DECIDED THEY ONLY WANTED A COUNTDOWN CLOCK INSTEAD OF A TIME ELLAPSED COUNTER! 9/3/15
                                        'Response.write "Line 678 DeliveryTime="&DeliveryTime&"<BR>"
                                        If Status="CLS" then
                                            TotalTimeEllapsed=DateDiff("n", DeliveryTime, DueTime)
                                            Else
                                            TotalTimeEllapsed=DateDiff("n", Now(), DueTime)
                                        
                                        End if
                                        If TotalTimeEllapsed<0 then
                                            DisplayTotalTimeEllapsed="<font color='#d71e26'>LATE</font>"
                                            ELSE

										        DisplayTotalTimeEllapsed=TotalTimeEllapsed	
										        If DisplayTotalTimeEllapsed>59 then
											        temphours=fix(DisplayTotalTimeEllapsed/60)
											        tempminutes=DisplayTotalTimeEllapsed-(temphours*60)
											        DisplayTotalTimeEllapsed=Temphours&" hrs "&TempMinutes&" mins"
											        else
											        DisplayTotalTimeEllapsed=DisplayTotalTimeEllapsed&" mins"
										        End if
										        If trim(EDI_DateTime)="" or isnull(EDI_DateTime) then
											        EDI_DateTime="&nbsp;"
										        End if	

                                            End if
                                        																		
									%>
								</td>
						<!--		<td align="center" class="subheader"><%=DisplayEDI_DateTime%></td> -->
								<td align="center" class="subheader"><%=DisplayTotalTimeEllapsed%></td>
							</tr>								
								<%
								End if							
								%>
							</tr>
							<%
							LLL=LLL+1
							If LLL=17 then
								%>
								<tr>
									<td class="FleetXRedSectionLM"  nowrap="nowrap" >&nbsp;&nbsp;ORDER Number&nbsp;&nbsp;</td>
                                    <td class="FleetXRedSectionLM"  nowrap="nowrap" >&nbsp;&nbsp;Vehicle&nbsp;&nbsp;</td>
                                    <td class="FleetXRedSectionLM"  nowrap="nowrap" >&nbsp;&nbsp;Driver&nbsp;&nbsp;</td>
									<td class="FleetXRedSectionLM"  nowrap="nowrap" width="100">&nbsp;&nbsp;Dispatched&nbsp;&nbsp;</td>
									<td class="FleetXRedSectionLM"  nowrap="nowrap" width="100">&nbsp;&nbsp;Acknowledged&nbsp;&nbsp;</td>
								<!--	<td class="FleetXRedSectionLM" nowrap width="100">&nbsp;&nbsp;Paper on Board&nbsp;&nbsp;</td>
									<td class="FleetXRedSectionLM" nowrap width="100">&nbsp;&nbsp;Arrived at Airline&nbsp;&nbsp;</td>  -->
									<td class="FleetXRedSectionLM"  nowrap="nowrap" width="100">&nbsp;&nbsp;On Board&nbsp;&nbsp;</td>
									<td class="FleetXRedSectionLM"  nowrap="nowrap" width="100">&nbsp;&nbsp;Delivered&nbsp;&nbsp;</td>
									<td class="FleetXRedSectionLM"  nowrap="nowrap" width="100">&nbsp;&nbsp;Exceptions&nbsp;&nbsp;</td>
								<!--	<td class="FleetXRedSectionLM" nowrap width="100">&nbsp;&nbsp;EDI Sent&nbsp;&nbsp;</td>  -->
									<td class="FleetXRedSectionLM"  nowrap="nowrap">&nbsp;&nbsp;Time Remaining&nbsp;&nbsp;</td>
								</tr>							
								<%
								LLL=0
							END IF
							End if
							
							tempjobno=jobno
							'Response.Write "Hello 1<BR>"
							
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''							
							oRs22.MoveNext
							LOOP
							'''Response.Write "Hello 2<BR>"
						End if
					End If
					oRs22.close
					Set oRs22 = Nothing						
					%>
			</table>
		</td>
	</tr>
	<!----------------------------------->
	<tr><td><img src="images/pixel.gif" height="5" width="10"></td></tr>
	<tr>
		<td colspan="7">
			<table cellpadding="0" cellspacing="0" border="0" bordercolor="brown" ID="Table5">
				<tr>
					<td bgcolor="#94FAA2" class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;</td><td nowrap class="generalcontent">&nbsp;Dispatched</td>
					<td class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td> 
					<td bgcolor="#FBB5B5" class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;</td><td nowrap class="generalcontent">&nbsp;Acknowledged</td>
					<td class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<!-- <td bgcolor="#C1C1F9" class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;</td><td nowrap class="generalcontent">&nbsp;POB</td>
					<td class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<td bgcolor="#C8F9C1" class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;</td><td nowrap class="generalcontent" nowrap>&nbsp;Arrived at Airline</td>
					<td class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>	-->										
					<td bgcolor="#F8DAA1" class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;</td><td nowrap class="generalcontent">&nbsp;On Board</td>
					<td class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<td bgcolor="#A1F7F8" class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;</td><td nowrap class="generalcontent">&nbsp;Delivered</td>
					<!--
                    <td class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<td nowrap class="generalcontent" height="5"><img src="../images/bluedot.gif" height="10" width="10" border="0"></td><td nowrap class="generalcontent">&nbsp;POD Scanned</td>
					-->
                    <!--
					<td bgcolor="red" class="generalcontent"><img src="../images/legendstripes.gif" height="14" width="17"></td><td nowrap class="generalcontent">&nbsp;Time Since Last Status Change</td>						
					-->
				</tr>
			</table>
		</td>				
	</tr> 	
	<!----------------------------------->
	</table>
	</form>
	</CENTER>
    
    
    </td></tr>



 
	<tr Height="50">
		<td>&nbsp;</td>
	</tr>

  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
      &nbsp;
    </td>
	<td width="5">&nbsp;</td>
    <td width="136"> 
     &nbsp;
    </td>
	<td width="5">&nbsp;</td>
    <td width="725"> 
      &nbsp;
    </td>
  </tr>
</table>
</td></tr>
<%
if ErrorMessage>"" then%>
<tr><td>
<table width="100%" border="0" bordercolor="Yellow" cellspacing="0" cellpadding="0" align="center" class="MainPageText">
	<tr><td>&nbsp;</td></tr>  
	<tr>
    <td align="center" class="Errormessage"><%=ErrorMessage%></td>
  </tr>
	<tr><td>&nbsp;</td></tr>
</table>
</td></tr>
<%end if%>
</table>
<!-- </form>  -->
<tr><td Height="90%">&nbsp;</td></tr>
<tr>
    <td height="100" class="FleetXGreySection" colspan="2">
        <!-- #include file="../BottomSection.asp" -->
    </td>
</tr>
<tr><td height="15" class="FleetXRedSectionSmall" colspan="2" align="center"><%=CopywriteNotice %></td></tr>
</td></tr></table>


</body>
</html>

