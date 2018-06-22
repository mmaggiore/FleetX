<html>
<head>
    <meta http-equiv="refresh" content="60" />

<!-- #include file="../fleetexpress.inc" -->
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
<script language="javascript" type="text/javascript" src="datetimepicker.js">
    //Date Time Picker script- by TengYong Ng of http://www.rainforestnet.com
</script>
<%
    TableWidth="460"
    RequestedDate=valid8(Request.form("RequestedDate"))
    If Trim(RequestedDate)="" then
        RequestedDate=Date()
    End if
    'Response.write "RequestedDate="&RequestedDate&"<BR>"
    ColorSelect=valid8(Request.form("ColorSelect"))
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
    PageTitle="LIVE VEHICLE MONITOR"

%>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<meta http-equiv="refresh" content="60" />

<title>FleetX - <%=PageTitle %></title>
</head>

<body leftMargin="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
	<table align="center" border="0" bordercolor="black" cellpadding="0" cellspacing="0" ID="Table1" height="100%" width="100%">
        <tr><td align="left"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>
        <tr>
            <td align="left"><img src="../images/FleetX_Small.jpg" height="50" width="168" /></td>
            <td align="right" valign="bottom"><!-- #include file="../topnavbar.asp" --></td>
        </tr>
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" /></td></tr>	
    <tr><td colspan="2">

<table border="0" bordercolor="green" Cellspacing="0" Cellpadding="0" align="left" width="100%">
 <tr><td align="left" colspan="2" bgcolor="<%=HeaderBorderColor%>"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>

        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetXRedSection" colspan="2"><%=PageTitle%></td></tr>

        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
        <!--tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressBodyWhite" colspan="2">In order to reach the Fleet Express Order page, please correctly type in the green verification code in the supplied text box and click "Submit."</td></tr-->
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
<tr><td>
<table  border="0" bordercolor="blue" align="center" class="MainPageText" width="100%">
	<tr height="40">
		<td width="650">&nbsp;</td>
	</tr>
    <tr><td><!-- main page stuff goes here! -->
 <%     HeaderBorderColor = "green" %>
   
             <table border="0" width="100%" bordercolor="<%=HeaderBorderColor%>" cellpadding="4" cellspacing="0" align="center">

                <!-- <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2">Live Vehicle Monitor - <%=RequestedDate%></td></tr>  -->
               <!-- <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>   -->
                <!--tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressBodyWhite" colspan="2">Please complete all areas below</td></tr-->
               <!-- <form method="post" name="formabc1" id="formabc1">  -->
                <tr><td><table border="0" cellpadding="0" cellspacing="0" align=center><tr><td>
                <!--
                   &nbsp;Date: <input type="text" name="RequestedDate" id="RequestedDate" value="<%=RequestedDate%>" size="10" /><a href="javascript:NewCal('RequestedDate','MMddyyyy',true,12)"><img src="../images/cal.gif" width="16" height="16" border="0" alt="Pick a date"></a>    
                &nbsp;&nbsp;<input type="submit" value="submit" />
                 -->
                </td>
                <td width="50">&nbsp;</td><td bgcolor="black">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td>&nbsp;&nbsp;New</td>
                <td width="50">&nbsp;</td><td bgcolor="green">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td>&nbsp;&nbsp;Dispatched</td>
                <td width="50">&nbsp;</td><td bgcolor="blue">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td>&nbsp;&nbsp;Acknowledged</td>
                <td width="50">&nbsp;</td><td bgcolor="orange">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td>&nbsp;&nbsp;In Transit</td>                
                <td width="50">&nbsp;</td><td bgcolor="red">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td>&nbsp;&nbsp;Late</td></tr></table></td></tr>                
               <!-- </form>   -->
                <tr><td align="left" colspan="2"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>
                 <tr>
                    <td align="center" colspan="2">
                        <table border="0" cellpadding="0" cellspacing="0" align="center" width="85%">
       <%
 				                'l_cSQL2 = "Select VehicleName, DriverID FROM AvailableVehicles WHERE AvailableStatus='c' ORDER BY VehicleName"
                        
                        l_cSQL2 =  "select fh_id, fh_status, fh_carr_id, fl_st_rta, fl_un_id from fcfgthd, fclegs where fclegs.fl_fh_id = fcfgthd.fh_id and fh_status<>'CAN' and fh_status<>'CLS' ORDER BY fl_un_id"

				                Set oRs23 = Server.CreateObject("ADODB.Recordset")
				                oRs23.CursorLocation = 3
				                oRs23.CursorType = 3
				                oRs23.ActiveConnection = DATABASE	
					                Err.Clear
					                oRs23.Open l_cSQL2, DATABASE, 1, 3
					                If Err.Number <> 0 Then                                               
					                Response.Write "Error Executing the query.  Error:" & Err.Description
					                Else
                           if oRs23.eof then
                                  response.write "<tr><td bgcolor='#e8e5e5' width=240 height=50>No Jobs Found</td></tr>"
                            else 
                            oldvehicle = 0
                            
						                DO WHILE NOT oRs23.EOF
                                
                                VehicleID=trim(oRs23("fl_un_id"))
                                if not isnumeric(vehicleid) then
                                  vehicleID = 0
                                end if 
                                if oldvehicle <> VehicleID or oldvehicle = 0 then
                                  if oldvehicle <> 0 then
                                     %></tr></table></td></tr><%
                                  end if
                                      %>
                                        <tr><td colspan=2 bgcolor=black><img src="../images/pixel.gif" height="1" width="100%" /></td></tr>
                                      <%
                                      
                                      l_cSQL233 = "Select VehicleName, DriverID FROM AvailableVehicles WHERE VehicleID='" & VehicleID & "' and AvailableStatus='c'"
                				                'response.write "144 " & l_cSQL233 & "<br>"
                                        Set oRs233 = Server.CreateObject("ADODB.Recordset")
                				                oRs233.CursorLocation = 3
                				                oRs233.CursorType = 3
                				                oRs233.ActiveConnection = DATABASE	
                					                Err.Clear
                					                oRs233.Open l_cSQL233, DATABASE, 1, 3
                					                'response.write "151 err=" & Err.Number & "<br>"
                                          If not oRs233.eof Then                                               
                                                VehicleName = oRs233("VehicleName")
                                                Set oConnA = Server.CreateObject("ADODB.Connection")
                              									oConnA.ConnectionTimeout = 100
                              									oConnA.Provider = "MSDASQL"
                              									oConnA.Open INTRANET
                              										iuSQL = "Select FirstName, LastName FROM intranet_users WHERE UserID = " & trim(oRs233("DriverID"))
                              										'response.write "157 sql=" & iuSQL & "<br>"
                                                  SET oRsa2 = oConnA.Execute(iuSql)
                                                    DriverName=oRsa2("FirstName") & " " & oRsa2("LastName")
                                                    if len(trim(DriverName)) < 1 then
                                                      DriverName = "N/A"
                                                    end if
                             									    oRsa2.close
                                                  Set oConnA=Nothing
                                         Else
                                            VehicleName = VehicleID
                                            DriverName = "(No Driver)"
                                         End If 
                                       
                                  %>
                                           <tr><td nowrap align="left" valign="top">
                                               <%=VehicleName%><br><%=DriverName%>
                                          </td>
                                    <td align="left" valign="top" height=50>
                                     <%
                                          x=0
                                          v=0
                                          response.write "<table border=0 cellpadding=0 cellspacing=0><tr>"
                                          bcolor="white"
                                          oldvehicle = VehicleID
                                    %>
                                    <% End If %>
                                  
                                    <%
                                 if bcolor="white" then
                                  bcolor="#e8e5e5"
                                 else
                                  bcolor="white"
                                 end if
                                 x=x+1
                                 v=v+1
                                    if oRs23("fh_status") = "ONB" then
                                    LinkClass="FleetLiveMonitorOrange"
                                 elseif oRs23("fh_status") = "ACC" then
                                    LinkClass="FleetLiveMonitorBlue"
                                 elseif oRs23("fh_status") = "OPN" then
                                    LinkClass="FleetLiveMonitorGreen"
                                 elseif len(trim(oRs23("fh_status"))) < 1 or ISNULL(oRs23("fh_status")) or oRs23("fh_status") = "NEW" then
                                    LinkClass="FleetLiveMonitorBlack"
                                 end if

                                 if v>5 then
                                  response.write "</td></tr><tr><td bgcolor=" & bcolor & " width=240 height=50 nowrap>"
                                  v=1
                                 elseif x>1 then
                                  response.write "</td><td bgcolor=" & bcolor & " width=240 height=50 nowrap>"
                                 else
                                  response.write "<td bgcolor=" & bcolor & " width=240 height=50 nowrap>"
                                 end if

                                        duedate = oRs23("fl_st_rta")
                                        'response.write "164 duedate=" & duedate &"<br>"
                                        duediff = DateDiff("n",Now(),duedate)
                                        if duediff < 0 then
                                          'duediff = "<font color='red'><b>LATE</b></font>"
                                          duediff = " "
                                          LinkClass="FleetLiveMonitorRed"
                                        else
                                          'duediff = duediff & " mins"
                                          duediff = datediffCNV(Now(),duedate)
                                        end if
                                %>
                                <a class="<%=LinkClass%>" href="FleetXOrderDispatch.asp?SearchJobNumber=<%=oRs23("fh_id")%>&PageStatus=disp&findjob=y"><%=oRs23("fh_id")%></a> <%=duediff%>
                                
                                <%
        								    oRs23.movenext
        								    LOOP
                              %>                                  
                            </td></tr></table></td></tr>
                            <tr><td colspan=2 bgcolor=black><img src="../images/pixel.gif" height="1" width="100%" /></td></tr>
                          <%  
                          End if
					                ENd If
                          oRs23.close
					                Set oRs23 = Nothing	
        %>
                        </table>
                      </td>
                    </tr>
                 <tr><td align="left"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>
                 
            </table>    
  
    
    </td></tr>



 
	<tr Height="50">
		<td>&nbsp;</td>
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
