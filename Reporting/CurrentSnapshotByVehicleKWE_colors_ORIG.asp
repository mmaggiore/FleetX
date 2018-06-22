<%@ Language=VBScript CodePage=65001 %>
<%  
	l_cSel = Request.Form("STSEL")
	
	' This is the DATE the user entered
	TheCurrentTime=Now()
	Submit=Request.Form("Submit")
	SQLDriverID=Request.Form("DriverID")
	NumberOfHours=request.Form("NumberOfHours")
	OrderType=Request.Form("OrderType")
	If NumberOfHours="" and Session("NumberOfHours")="" then 
		NumberOfHours=48
		Session("NumberOfHours")=48
		OrderType="all"
		Session("OrderType")="all"	
	End if
	If Submit>"" then
		Session("NumberOfHours")=NumberOfHours
		Session("SQLDriverID")=SQLDriverID
		Session("OrderType")=OrderType
	End if
	OrderType=Session("OrderType")
	SQLDriverID=Session("SQLDriverID")	
	NumberOfHours=Session("NumberOfHours")
	%>
	<html><head>
	<link rel="stylesheet" type="text/css" href="../css/style.css">
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<meta http-equiv="refresh" content="90" />
	<TITLE><% Response.Write(D_TITLEBAR) %></TITLE>
	</head>
	<CENTER>
	<%
	'Response.Write("<H2><FONT COLOR='red'>This page is currently being worked on.<br>  It may become unavailable for brief periods.<br>  Thanks,<br>  Mark</FONT></H2>")
	Response.Write("<H2><FONT COLOR=#612099>Vehicle Monitor of Last "&NumberOfHours&" Hours from "&TheCurrentTime&"</FONT></H2>")
	%>
	<!-- #include file="../include/settings.inc" -->

	<form action="CurrentSnapshotbyvehicleKWE_colors.asp" method="POST" ID="Form1">
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
				<td class="generalcontent"><input type="submit" value="Submit/Refresh" name="submit" class="buttongrey" ID="Submit1"></td>
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
							<td bgcolor="#FBB5B5" class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;</td><td nowrap class="generalcontent">&nbsp;Acknowledge</td>
							<td class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
							<td bgcolor="#C1C1F9" class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;</td><td nowrap class="generalcontent">&nbsp;POB</td>
							<td class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
							<td bgcolor="#C8F9C1" class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;</td><td nowrap class="generalcontent" nowrap>&nbsp;Arrived at Airline</td>
							<td class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>											
							<td bgcolor="#F8DAA1" class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;</td><td nowrap class="generalcontent">&nbsp;On Board</td>
							<td class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
							<td bgcolor="#A1F7F8" class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;</td><td nowrap class="generalcontent">&nbsp;Delivery</td>
							<td class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
							<td nowrap class="generalcontent" height="5"><img src="../images/bluedot.gif" height="10" width="10" border="0"></td><td nowrap class="generalcontent">&nbsp;POD Scanned</td>
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
				<td class="subheader" nowrap bgcolor="#E3E3DF">&nbsp;&nbsp;HAWB Number&nbsp;&nbsp;</td>
				<td class="subheader" align="center" nowrap width="100" bgcolor="#E3E3DF">&nbsp;&nbsp;Acknowledged&nbsp;&nbsp;</td>
				<td class="subheader" align="center" nowrap width="100" bgcolor="#E3E3DF">&nbsp;&nbsp;Paper on Board&nbsp;&nbsp;</td>
				<td class="subheader" align="center" nowrap width="100" bgcolor="#E3E3DF">&nbsp;&nbsp;Arrived at Airline&nbsp;&nbsp;</td>
				<td class="subheader" align="center" nowrap width="100" bgcolor="#E3E3DF">&nbsp;&nbsp;On Board&nbsp;&nbsp;</td>
				<td class="subheader" align="center" nowrap width="100" bgcolor="#E3E3DF">&nbsp;&nbsp;Delivered&nbsp;&nbsp;</td>
				<td class="subheader" align="center" nowrap width="100" bgcolor="#E3E3DF">&nbsp;&nbsp;Exceptions&nbsp;&nbsp;</td>
				<td class="subheader" align="center" nowrap width="100" bgcolor="#E3E3DF">&nbsp;&nbsp;EDI Sent&nbsp;&nbsp;</td>
				<td class="subheader" align="center" nowrap width="100" bgcolor="#E3E3DF">&nbsp;&nbsp;Elapsed Time&nbsp;&nbsp;</td>
			</tr>
			<% 
					l_cSQL="SELECT * from Mark_KWE_LiveMonitor "
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
					If OrderType="open" then
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
					l_cSQL=l_cSQL&" AND (Customer = '"&suid&"') AND (DriverID <> '111') AND (DriverID <> 'EFRANCO') "
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
  					
  					'Response.Write "l_cSQL="&l_cSQL&"<BR>"
  					
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
							Priority=trim(oRs22("Priority"))
							Fl_Job_Closed=oRs22("Fl_Job_Closed")
							DeliveryTime=Fl_Job_Closed
							Status=ucase(trim(oRs22("Status")))
							HAWB=trim(oRs22("HAWB"))
							EDI_DateTime=trim(oRs22("EDI_DateTime"))
                            HAWBNUMBER=trim(oRs22("HAWBNUMBER"))
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
							AcknowledgeTimeEllapsed=DateDiff("n",BookTime,AcknowledgeTime)
							If AcknowledgeTimeEllapsed=0 then AcknowledgeTimeEllapsed=1 end if
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
							''''''''''''''''''''''''''''''''''''''''''''
							
							POBTimeEllapsed=DateDiff("n",AcknowledgeTime,POBTime)
							If POBTimeEllapsed=0 and AcknowledgeTime>"1/1/1900" and POBTime>"1/1/1900" then POBTimeEllapsed=1 end if
							If JobNo="00077112x" then
								response.write "POBTimeEllapsed="&POBTimeEllapsed&"<BR>"
							End if
							If abs(POBTimeEllapsed)>99999 then POBTimeEllapsed=0 End if	
							TotalTimeEllapsed=TotalTimeEllapsed+POBTimeEllapsed
							If POBTimeEllapsed>59 then
								temphours=fix(POBTimeEllapsed/60)
								tempminutes=POBTimeEllapsed-(temphours*60)
								POBTimeEllapsed=Temphours&" hrs "&TempMinutes&" mins"
								else
								POBTimeEllapsed=POBTimeEllapsed&" mins"
							End if								
							
							AtAirlineTimeEllapsed=DateDiff("n",POBTime,AtAirlineTime)
							If AtAirlineTimeEllapsed=0 and POBTime>"1/1/1900" and AtAirlineTime>"1/1/1900" then AtAirlineTimeEllapsed=1 end if
							If JobNo="00077112x" then
								response.write "AtAirlineTimeEllapsed="&AtAirlineTimeEllapsed&"<BR>"
							End if
							If abs(AtAirlineTimeEllapsed)>99999 then AtAirlineTimeEllapsed=0 End if	
							TotalTimeEllapsed=TotalTimeEllapsed+AtAirlineTimeEllapsed
							If AtAirlineTimeEllapsed>59 then
								temphours=fix(AtAirlineTimeEllapsed/60)
								tempminutes=AtAirlineTimeEllapsed-(temphours*60)
								AtAirlineTimeEllapsed=Temphours&" hrs "&TempMinutes&" mins"
								else
								AtAirlineTimeEllapsed=AtAirlineTimeEllapsed&" mins"
							End if								
																				
							'''''''''''''''''''''''''''''''''''''''''''''

							OnBoardTimeEllapsed=DateDiff("n",AtAirlineTime,OnBoardTime)
							If OnBoardTimeEllapsed=0 and AtAirlineTime>"1/1/1900" and OnBoardTime>"1/1/1900" then OnBoardTimeEllapsed=1 end if
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
							
							DeliveryTimeEllapsed=DateDiff("n",OnBoardTime,DeliveryTime)
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
							
							CurrentOrderTime=DateDiff("n",BookTime,TheCurrentTime)	
							OrderClosed="n"	
							ShowDisplay="y"
							
							
							
								if AcknowledgeTimeEllapsed<>"0 mins" then 
									ackbgcolor="#FBB5B5"
									DisplayAcknowledgeTimeEllapsed=AcknowledgeTimeEllapsed
									else
									ackbgcolor="white"
									DisplayAcknowledgeTimeEllapsed="&nbsp;"
								 end if
								if POBTimeEllapsed<>"0 mins" then 
									POBbgcolor="#C1C1F9"
									DisplayPOBTimeEllapsed=POBTimeEllapsed
									else
									POBbgcolor="white"
									DisplayPOBTimeEllapsed="&nbsp;"
									If DisplayAcknowledgeTimeEllapsed<>"&nbsp;" then
										DisplayPOBTimeEllapsed=DateDiff("n",AcknowledgeTime,now())
										TotalTimeEllapsed=TotalTimeEllapsed+DisplayPOBTimeEllapsed
										If DisplayPOBTimeEllapsed>59 then
											temphours=fix(DisplayPOBTimeEllapsed/60)
											tempminutes=DisplayPOBTimeEllapsed-(temphours*60)
											DisplayPOBTimeEllapsed=Temphours&" hrs "&TempMinutes&" mins"
											else
											DisplayPOBTimeEllapsed=DisplayPOBTimeEllapsed&" mins"
										End if										
										ShowDisplay="n"
									End if
								end if
								If DisplayAcknowledgeTimeEllapsed="&nbsp;" then
									TimeNow=now()
									AcknowledgeTimeEllapsed=DateDiff("n",BookTime,TimeNow)
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
								if AtAirlineTimeEllapsed<>"0 mins" then 
									atairlinebgcolor="#C8F9C1" 
									DisplayAtAirlineTimeEllapsed=AtAirlineTimeEllapsed
									else
									atairlinebgcolor="white" 
									DisplayAtAirlineTimeEllapsed="&nbsp;"
									If DisplayPOBTimeEllapsed<>"&nbsp;" and ShowDisplay<>"n" then
										DisplayAtAirlineTimeEllapsed=DateDiff("n",POBTime,now())
										TotalTimeEllapsed=TotalTimeEllapsed+DisplayAtAirlineTimeEllapsed
										If DisplayAtAirlineTimeEllapsed>59 then
											temphours=fix(DisplayAtAirlineTimeEllapsed/60)
											tempminutes=DisplayAtAirlineTimeEllapsed-(temphours*60)
											DisplayAtAirlineTimeEllapsed=Temphours&" hrs "&TempMinutes&" mins"
											else
											DisplayAtAirlineTimeEllapsed=DisplayAtAirlineTimeEllapsed&" mins"
										End if										
										ShowDisplay="n"
									End if									
								end if
								if OnBoardTimeEllapsed<>"0 mins" then 
									onbbgcolor="#F8DAA1" 
									DisplayOnBoardTimeEllapsed=OnBoardTimeEllapsed
									else
									onbbgcolor="white" 
									DisplayOnBoardTimeEllapsed="&nbsp;"
									If DisplayAtAirlineTimeEllapsed<>"&nbsp;" and ShowDisplay<>"n" then
										DisplayOnBoardTimeEllapsed=DateDiff("n",AtAirlineTime,now())
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
								if DeliveryTimeEllapsed<>"0 mins" then 
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
							'Response.Write "**********************<BR>"	
							'Response.Write "JobNo="&JobNo&"<BR>"
							'Response.Write "TempJobNo="&TempJobNo&"<BR>"
							'Response.Write "FL_PKey="&FL_PKey&"<BR>"
							'Response.Write "OnBoardTime="&OnBoardTime&"<BR>"
							'Response.Write "**********************<BR>"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''							
							
							%>
							<tr><td valign="middle"><a href="jobanalysis.asp?inputjobnumber=<%=JobNo%>&sbt_id=48" target="_blank" class="<%=LinkClass%>"><%=HAWB%></a> &nbsp; <%if Trim(HAWBNUMBER)>"" then%><img src="../images/bluedot.gif" height="10" width="10" border="0"><%end if%></td>
							<%
							If status="CANCELLED" or status="DELETED" then
								%>
								<td class="subheader">&nbsp;&nbsp;<%=status%></td>
								<td class="generalcontent" align="center" nowrap>&nbsp;</td>
								<td class="generalcontent" align="center" nowrap>&nbsp;</td>
								<td class="generalcontent" align="center" nowrap>&nbsp;</td>
								<td class="generalcontent" align="center" nowrap>&nbsp;</td>
								<td class="generalcontent" align="center" nowrap>&nbsp;</td>
								<td class="generalcontent" align="center" nowrap>&nbsp;</td>
							</tr>
								<%
								MinuteColor="white"
								else
								%>
								<td height="30" valign="middle" class="subheader" align="center" nowrap bgcolor="<%=ackbgcolor%>"><%=DisplayAcknowledgeTimeEllapsed%></td>
								<td class="subheader" align="center" nowrap bgcolor="<%=pobbgcolor%>"><%=DisplayPOBTimeEllapsed%></td>
								<td class="subheader" align="center" nowrap bgcolor="<%=atairlinebgcolor%>"><%=DisplayAtAirlineTimeEllapsed%></td>
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
									%>
								</td>
								<td align="center" class="subheader"><%=DisplayEDI_DateTime%></td>
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
									<td class="subheader" nowrap bgcolor="#E3E3DF">&nbsp;&nbsp;HAWB Number&nbsp;&nbsp;</td>
									<td class="subheader" align="center" nowrap width="100" bgcolor="#E3E3DF">&nbsp;&nbsp;Acknowledged&nbsp;&nbsp;</td>
									<td class="subheader" align="center" nowrap width="100" bgcolor="#E3E3DF">&nbsp;&nbsp;Paper on Board&nbsp;&nbsp;</td>
									<td class="subheader" align="center" nowrap width="100" bgcolor="#E3E3DF">&nbsp;&nbsp;Arrived at Airline&nbsp;&nbsp;</td>
									<td class="subheader" align="center" nowrap width="100" bgcolor="#E3E3DF">&nbsp;&nbsp;On Board&nbsp;&nbsp;</td>
									<td class="subheader" align="center" nowrap width="100" bgcolor="#E3E3DF">&nbsp;&nbsp;Delivered&nbsp;&nbsp;</td>
									<td class="subheader" align="center" nowrap width="100" bgcolor="#E3E3DF">&nbsp;&nbsp;Exceptions&nbsp;&nbsp;</td>
									<td class="subheader" align="center" nowrap width="100" bgcolor="#E3E3DF">&nbsp;&nbsp;EDI Sent&nbsp;&nbsp;</td>
									<td class="subheader" align="center" nowrap width="100" bgcolor="#E3E3DF">&nbsp;&nbsp;Elapsed Time&nbsp;&nbsp;</td>
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
					<td bgcolor="#FBB5B5" class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;</td><td nowrap class="generalcontent">&nbsp;Acknowledge</td>
					<td class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<td bgcolor="#C1C1F9" class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;</td><td nowrap class="generalcontent">&nbsp;POB</td>
					<td class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<td bgcolor="#C8F9C1" class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;</td><td nowrap class="generalcontent" nowrap>&nbsp;Arrived at Airline</td>
					<td class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>											
					<td bgcolor="#F8DAA1" class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;</td><td nowrap class="generalcontent">&nbsp;On Board</td>
					<td class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<td bgcolor="#A1F7F8" class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;</td><td nowrap class="generalcontent">&nbsp;Delivery</td>
					<td class="generalcontent">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<td nowrap class="generalcontent" height="5"><img src="../images/bluedot.gif" height="10" width="10" border="0"></td><td nowrap class="generalcontent">&nbsp;POD Scanned</td>
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
