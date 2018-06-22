<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Fleet Express Vehicle Monitor</title>
    <meta http-equiv="refresh" content="120" />
    <link rel="stylesheet" type="text/css" href="../css/Style.css">
   <%
   '''''''''''HARDCODED STUFF
   'sBT_ID="84"
   'Session("sBT_ID")=sBT_ID
   sBT_ID=84
   'If trim(sBT_ID)="" then
   '     sBT_ID=Request.QueryString("bid")
   Session("sBT_ID")=sBT_ID
   UserID=Session("UserID")
   'Response.write "userid="&userid&"<BR>"
   Dim DisplayReservations(24)
   Dim DisplayTime(25,25)
   Dim DisplayVehicleName(25)
    'End if
   ''''''''''''''''''''''''''
    %>
    <!-- #include file="../include/ifabsettings.inc" -->
    <!-- #include file="../include/checkstring.inc" -->
<script src="datetimepicker_css.js">
    //Date Time Picker script- by TengYong Ng of http://www.rainforestnet.com
    //Script featured on JavaScript Kit (http://www.javascriptkit.com)
    //For this script, visit http://www.javascriptkit.com 
</script>
    <%
  
    TableWidth="460"
    RequestedDate=Request.form("RequestedDate")
    If Trim(RequestedDate)="" then
        RequestedDate=Date()
    End if
    'Response.write "RequestedDate="&RequestedDate&"<BR>"
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
            HeaderBorderColor="#41924B"  
            BorderColor="#41924B"
            LinkClass="FleetExpressGreen"
        Case else 
            HeaderBorderColor="black"  
            BorderColor="black"
            LinkClass="FleetExpressBlack"
    End Select
    HighlightedField="RequestorName"
    CurrentDateTime=Now()
     %>
</head>
<BODY leftMargin="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
    <table border="0" cellpadding="0" cellspacing="0" align="center">
        <tr><td align="left"><img src="images/pixel.gif" height="10" width="1" /></td></tr>
        <tr>
            <td align="left"><img src="images/logo_FleetExpress_space.gif" height="87" width="100" /></td>
            <td align="right" valign="bottom"><a href="mailto:mark.maggiore@logisticorp.us" class="<%=LinkClass%>">Click here to report a problem with this page</a></td>
        </tr>
        <tr><td align="left" colspan="2"><img src="images/pixel.gif" /></td></tr>
        <tr><td colspan="2">
             <table border="0" width="100%" bordercolor="<%=HeaderBorderColor%>" cellpadding="0" cellspacing="0" align="center">

                <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2">Live Vehicle Monitor - <%=RequestedDate%></td></tr>
                <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2"><img src="images/pixel.gif" height="5" width="1" /></td></tr>
                <!--tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressBodyWhite" colspan="2">Please complete all areas below</td></tr-->
                <form method="post" name="formabc1" id="formabc1">
                <tr><td><table border="0" cellpadding="0" cellspacing="0"><tr><td>
                   &nbsp;Date: <input type="text" name="RequestedDate" id="RequestedDate" value="<%=RequestedDate%>" size="10" /><img src="images2/cal.gif" onclick="javascript:NewCssCal ('RequestedDate','mmddyyyy','arrow')" />   
                &nbsp;&nbsp;<input type="submit" value="submit" /></td>
                <td width="50">&nbsp;</td><td bgcolor="green">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td>&nbsp;&nbsp;Pick Up</td>
                <td width="50">&nbsp;</td><td bgcolor="red">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td>&nbsp;&nbsp;Delivery</td></tr></table></td></tr>                
                </form>
                <tr><td align="left" colspan="2"><img src="images/pixel.gif" height="10" width="1" /></td></tr>
                 <tr>
                    <td align="center" colspan="2">
                        <table border="1" bordercolor="<%=HeaderBorderColor%>" cellpadding="0" cellspacing="0" align="center"  class="VEHICLEMONITOR">

                            <tr>
                   
       <%
 

'''''''''''''''''''''''''''''''MOVED IT TO HERE, OBVIOUSLY!
				                l_cSQL2 = "Select VehicleName from  FleetVehicles where ((vehiclestatus='c') and (LogisticorpOwned='y') and (IsRaytheon='y' or IsFleetExpress='y')) and (VehicleName<>'Austin Van') order by VehicleName" 
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                'Response.write "INTRANET="&INTRANET&"<BR>"
                                'Response.write "l_cSQ2L="&l_cSQL2&"<BR>"
				                Set oRs2 = Server.CreateObject("ADODB.Recordset")
				                oRs2.CursorLocation = 3
				                oRs2.CursorType = 3
				                oRs2.ActiveConnection = INTRANET	
					                Err.Clear
					                oRs2.Open l_cSQL2, INTRANET, 1, 3
					                If Err.Number <> 0 Then                                               
					                Response.Write "Error Executing the query.  Error:" & Err.Description
					                Else
						                DO WHILE NOT oRs2.EOF
                                            VehicleName=trim(oRs2("VehicleName"))

                                            ZYX=ZYX+1
                                            DisplayVehicleName(ZYX)=VehicleName
                                            'Response.write "vehiclename="&vehiclename&"<BR>"



            '''''''''''''''Stored Procedure''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				                l_cSQL2 = "EXEC Mark_LiveVehicleMonitor " & _
					                "@FromDate = '" & RequestedDate & "', " & _ 
                                    "@VehicleName = '" & VehicleName & "' " 
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                'Response.write "INTRANET="&INTRANET&"<BR>"
                                'Response.write "l_cSQ2L="&l_cSQL2&"<BR>"
				                Set oRs23 = Server.CreateObject("ADODB.Recordset")
				                oRs23.CursorLocation = 3
				                oRs23.CursorType = 3
				                oRs23.ActiveConnection = INTRANET	
					                Err.Clear
					                oRs23.Open l_cSQL2, INTRANET, 1, 3
					                If Err.Number <> 0 Then                                               
					                Response.Write "Error Executing the query.  Error:" & Err.Description
					                Else
						                DO WHILE NOT oRs23.EOF 
                                        x=x+1
                                        If x=1 and x=2 then
                                            %>
                                                 <td bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" nowrap align="left" valign="top">
                                                    <table border="0" cellpadding="0" cellspacing="0">
                                                        <tr>
                                                            <td>&nbsp;</td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" nowrap>12:00 AM</td>
                                                        </tr>                                                       
                                                        <tr>
                                                            <td align="right" nowrap>1:00 AM</td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" nowrap>2:00 AM</td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" nowrap>3:00 AM</td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" nowrap>4:00 AM</td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" nowrap>5:00 AM</td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" nowrap>6:00 AM</td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" nowrap>7:00 AM</td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" nowrap>8:00 AM</td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" nowrap>9:00 AM</td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" nowrap>10:00 AM</td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" nowrap>11:00 AM</td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" nowrap>12:00 PM</td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" nowrap>1:00 PM</td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" nowrap>2:00 PM</td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" nowrap>3:00 PM</td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" nowrap>4:00 PM</td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" nowrap>5:00 PM</td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" nowrap>6:00 PM</td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" nowrap>7:00 PM</td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" nowrap>8:00 PM</td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" nowrap>9:00 PM</td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" nowrap>10:00 PM</td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" nowrap>11:00 PM</td>
                                                        </tr>

                                                    </table>
                                                    </td>
                                            <%                                      
                                        End if                         														
							                'if User3="yes" then
								            'Response.Write "JobNo="&JobNo&"<BR>"
								            '''Response.Write "Got here!<br>"
                                            'BookTime=oRs23("fh_ship_dt")

								            JobNo=trim(oRs23("fh_id"))
                                            PickDrop=trim(oRs23("PickDrop"))
                                            Courier=trim(oRs23("Courier"))
                                            Location=trim(oRs23("Location"))
    							            ArrivalTime=trim(oRs23("ArrivalTime"))
								            Pieces=trim(oRs23("Pieces"))
                                            POD=trim(oRs23("POD"))
                                            DisplayBTID=trim(oRs23("BTID"))
                                            If PickDrop="p" then
                                                LinkClass="FleetLiveMonitorGreen"
                                                else
                                                LinkClass="FleetLiveMonitorRed"
                                            End if


                                            Select Case DisplayBTID
                                                Case "84"
                                                    DisplayBTID="R"
                                                    DisplayLink="RaytheonOrderDispatch.asp?SearchJobNumber="&jobno&"&PageStatus=disp&findjob=y&BTID=84"
                                                    
                                                Case "86"
                                                    DisplayBTID="F"
                                                    DisplayLink="FleetExpressOrderDispatch.asp?SearchJobNumber="&jobno&"&PageStatus=disp&findjob=y&BTID=86"
                                            End Select
                                            DeliveryPeriod=trim(oRs23("DeliveryPeriod"))
                                            'If hour(ArrivalTime)=xyz then
                                            '    Response.write "GOT HERE!!!<BR>"
                                                LLL=LLL+1
                                                '''If LLL>1 and hour(ArrivalTime)<>LLL then
                                                        DisplayTime(ZYX, hour(ArrivalTime))=DisplayTime(ZYX, hour(ArrivalTime))&"<a href='"&DisplayLink&"' target='_blank' class='"&LinkClass&"'>"&DisplayBTID&JobNo&"("&DeliveryPeriod&" hrs)</a><BR>&nbsp;"
                                                    '''else
                                                        '''DisplayTime(ZYX, hour(ArrivalTime))="<a href='"&DisplayLink&"' target='_blank' class='"&LinkClass&"'>"&DisplayTime(ZYX, hour(ArrivalTime))&DisplayBTID&JobNo&"("&DeliveryPeriod&" hrs)yyy</a>"
                                                     'DisplayTime(ZYX, hour(ArrivalTime))=DisplayTime(ZYX, hour(ArrivalTime))&DisplayBTID&JobNo&"("&DeliveryPeriod&" hrs)"
                                                '''End if
                                                'Response.write "********************<br>"
                                                'Response.write "DisplayTime("&ZYX&","&hour(ArrivalTime)&")="&DisplayTime(ZYX, hour(ArrivalTime))&"<BR>"
                                                'Response.write "********************<br>"
								        oRs23.movenext
								        LOOP
                                    End if
					                oRs23.close
					                Set oRs23 = Nothing	
                   
                   ''''''''''''PUT A TABLE HERE????                

                                     
                'For rrr=0 to 23
                '    DisplayTime(rrr)=""
                'next

				oRs2.movenext
				LOOP
            End if
			oRs2.close
			Set oRs2 = Nothing
            %>
            <tr>
            <%
            DisplayVehicleName(0)="&nbsp;"
            For var1 = 0 to ZYX
                %>
                    <td nowrap bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" ><%=DisplayVehicleName(var1) %></td>
                <%
            Next

            %>
                <td nowrap bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" >&nbsp;</td>
            </tr> 
            <%
            For var3=0 to 24
                Select Case Var3
                    case 0
                    TimeDef="12:00 AM"
                    case 1
                    TimeDef="1:00 AM"
                    case 2
                    TimeDef="2:00 AM"
                    case 3
                    TimeDef="3:00 AM"
                    case 4
                    TimeDef="4:00 AM"
                    case 5
                    TimeDef="5:00 AM"
                    case 6
                    TimeDef="6:00 AM"
                    case 7
                    TimeDef="7:00 AM"
                    case 8
                    TimeDef="8:00 AM"
                    case 9
                    TimeDef="9:00 AM"
                    case 10
                    TimeDef="10:00 AM"
                    case 11
                    TimeDef="11:00 AM"
                    case 12
                    TimeDef="12:00 PM"
                    case 13
                    TimeDef="1:00 PM"
                    case 14
                    TimeDef="2:00 PM"
                    case 15
                    TimeDef="3:00 PM"
                    case 16
                    TimeDef="4:00 PM"
                    case 17
                    TimeDef="5:00 PM"
                    case 18
                    TimeDef="6:00 PM"
                    case 19
                    TimeDef="7:00 PM"
                    case 20
                    TimeDef="8:00 PM"
                    case 21
                    TimeDef="9:00 PM"
                    case 22
                    TimeDef="10:00 PM"
                    case 23
                    TimeDef="11:00 PM"
                End Select
                %>
                <tr><td nowrap bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" ><%=TimeDef%></td>
                <%
                For var2=1 to ZYX
                    %>
                   <td align="left" width="135" valign="top" nowrap>&nbsp;<%=DisplayTime(var2, var3) %>&nbsp;</td>
                    <%
                Next
                %>
                    <td nowrap bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" ><%=TimeDef%></td>
                </tr>
                <%
            Next      	


'''''''''''''''''''MOVED STORED PROCEDURE FROM HERE TO INCLUDE IT IN THE LOOP!     
        %>
                            </tr>
                        </table>
                      </td>
                    </tr>
                 <tr><td align="left"><img src="images/pixel.gif" height="10" width="1" /></td></tr>
                 <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2">Fleet Express Transportation Call Center 972-499-3415</td></tr>
                 
            </table>
        </td></tr>
        <tr><td align="left"><img src="images/pixel.gif" height="30" width="1" /></td></tr>
    </table>

</body>
</html>
