<html>
<head>

<!-- #include file="fleetexpress.inc" -->
<link rel="stylesheet" type="text/css" href="css/Style.css">
<%
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

    ''''''''CHECKS TO SEE IF USER IS ADMIN
    Set Recordset1 = Server.CreateObject("ADODB.Recordset")
    'Response.write "Database="&Database&"<br>"
    Recordset1.ActiveConnection = Database
    SQL="SELECT * FROM PreExistingRequestor WHERE (RequestorID='"&Request.cookies("FleetXCookie")("UserID")&"') AND (RequestorStatus='c')"
    Recordset1.Source = SQL
    'response.write "SQL="& SQL &"<BR>"
    Recordset1.CursorType = 0
    Recordset1.CursorLocation = 2
    Recordset1.LockType = 1
    Recordset1.Open()
    Recordset1_numRows = 0
    'response.write "<font color='red'>SQL1="&Recordset1.Source&"<BR></font>"
	    if NOT Recordset1.EOF then
            Supervisor=Recordset1("RequestorType")
	    End if
Recordset1.Close()
Set Recordset1 = Nothing   


    HighlightedField="RequestorName"
    CurrentDateTime=Now()
    If Supervisor="A" then
        PageTitle="ADMIN HOME PAGE"
        else
        PageTitle="HOME PAGE"
    End if

	UserID=Request.cookies("FleetXCookie")("UserID")
	UserName=Request.cookies("FleetXCookie")("UserName")
	'Response.write "XXXUserID="&UserID&"XXX<BR>"

%>
<title>FleetX - <%=PageTitle %></title>
</head>

<body onload="document.FindUser.requestorName.focus();">
	<table align="center" border="0" bordercolor="black" cellpadding="0" cellspacing="0" ID="Table1" height="100%" width="100%">
        <tr><td align="left"><img src="images/pixel.gif" height="10" width="1" /></td></tr>
        <tr>
            <td align="left"><img src="images/FleetX_Small.jpg" height="50" width="168" /></td>
            <td align="right" valign="bottom"><!-- #include file="topnavbar.asp" --></td>
        </tr>
        <tr><td align="left" colspan="2"><img src="images/pixel.gif" /></td></tr>	
    <tr><td colspan="2">
<table border="0" bordercolor="green" Cellspacing="0" Cellpadding="0" align="left" width="100%">
 <tr><td align="left" colspan="2" bgcolor="<%=HeaderBorderColor%>"><img src="images/pixel.gif" height="5" width="1" /></td></tr>

        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetXRedSection" colspan="2"><%=PageTitle%></td></tr>

        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2"><img src="images/pixel.gif" height="5" width="1" /></td></tr>
        <!--tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressBodyWhite" colspan="2">In order to reach the Fleet Express Order page, please correctly type in the green verification code in the supplied text box and click "Submit."</td></tr-->
        <tr><td align="left" colspan="2"><img src="images/pixel.gif" height="5" width="1" /></td></tr>
<tr><td>
 	<%
	If now()<cdate("8/31/2016 1:30:00 PM") then
    'response.write "DisplayUserName="&DisplayUserName&"<BR>"
   ' If trim(WelcomeName)="HFAB QC (Houston)" then
	    %>

	    <table cellspacing="0" cellpadding="2"  border="1" bordercolor="red" align="center" ID="Table2">
        <tr><td class="mainpagetextbold">
	    <font color="red"><br /><CENTER>SYSTEM MESSAGE - 8/29/2016 </CENTER><br />***On Wednesday, August 31st between 12:00 PM and 1:30 PM, we will be performing maintenance on one of our servers.  During that time you may experience a brief period when you will be unable to access the FleetX website.***
        <br /><br />
        Please make best efforts to schedule shipment requests either before or after this time frame.
        <br /><br />
    </font>
        </td></tr>
        </table><br />
        <%End if %> 
<table  border="0" cellpadding="10" cellspacing="0" bordercolor="blue" align="center" class="MainPageText" width="100%">
	<tr height="40">
		<td width="50%">WELCOME <%=Ucase(UserName) %></td>
        <td width="50%" align="right" nowrap><%=Now() %></td>
	</tr>
    <tr><td colspan="2">Welcome to the new FleetX website.  If you have any questions/issues, please contact:  <a href="mailto:mark.maggiore@logisticorp.us" class="FleetXRedMain">Mark Maggiore</a></td></tr>
	<%
    If now()<cdate("12/04/2016 1:30:00 PM") then
	    %>
        <tr><td colspan="2">
	        <table cellspacing="0" cellpadding="2"  border="1" bordercolor="red" align="center" ID="Table3">
            <tr><td class="mainpagetextbold">
	        <font color="red"><br /><CENTER>SYSTEM MESSAGE - 11/04/2016 </CENTER><br />***11/4/2016 - New feature - You now have the ability to cancel your own orders.  Simply click on the cancel button below and find your order.***<br /><br />
            </font>
            </td></tr>
            </table>
        </td></tr>
        <%End if %>


    <tr><td colspan=2 width="100%"> <!-- main page stuff goes here! -->
         <%
            if Supervisor = "A" then 
                Set oConnA = Server.CreateObject("ADODB.Connection")
                oConnA.ConnectionTimeout = 100
                oConnA.Provider = "MSDASQL"
                oConnA.Open DATABASE
                iuSQL = "Select TOP 1 * FROM AutoDispatchStatus ORDER BY adid DESC"
                'response.write "1702 sql=" & iuSQL & "<br>"
                SET oRsa2 = oConnA.Execute(iuSql)
                if oRsa2.eof then
                  autodispatch = "OFF"
                  adbutton = "ON"
                  adtime = Now()
                else
                  adstatus = oRsa2("status")
                  adtimeon = oRsa2("dateon")
                  adtimeoff = oRsa2("dateoff")
                  if adstatus = "c" then
                    autodispatch = "ON"
                    adbutton = "OFF"
                    adtime = adtimeon
                  else
                    autodispatch = "OFF"
                    adbutton = "ON"
                    adtime = adtimeoff
                  end if
                end if
                oRsa2.close
                Set oConnA=Nothing

                thisUser = Request.cookies("FleetXCookie")("UserID")
                '''''Select Case thisUser
                	'''''Case "1", "220", "227", "146"
                    %><div style="float: right">
                     <table border=1 width=100 cellspacing=1 cellpadding=5 align=right>
                     <tr><td align="center">AutoDispatch is <b><%=autodispatch%></b><br>As of <%=adtime%><br>
                     <br><form method="post" action="Admin/AutoDispatchSwitch.asp"><input type="submit" id="gobutton" value="TURN AUTODISPATCH <%=adbutton%>" /></form></td></tr>
                     </table>
                   <!--  &nbsp;<br><img src="images/pixel.gif" height=5><font color="white">hhhh</font><br>&nbsp;&nbsp;<form method="post" action="Admin/AutoDispatchScheduler.asp"><input type="submit" id="gobutton" value="Auto-Dispatch Scheduler" />      -->
                    </div> <%
                	'''''Case Else
                '''''End Select

            End If 
         %>
       
            <table cellspacing="5" border="0">
                <tr><td><form method="post" action="ReprintWaybill.asp"><input type="submit" id="gobutton" value="Reprint Waybill" /></form></td></tr>
                <tr><td><form method="post" action="UpdateUserInfo.asp"><input type="submit" id="gobutton" value="Update User Info" /></form></td></tr>
                <%if billtoid<>"91" then %>
                    
                    <tr><td><form method="post" action="OrderEntry/FleetXAddressBook.asp"><input type="submit" id="gobutton" value="View/Edit Address Book" /></form></td></tr>
                    <tr><td><form method="post" action="orderentry/CancelPageFreight.asp"><input type="submit" id="gobutton" value="Cancel Courier/Freight Order" /></form></td></tr>
                    <tr><td><form method="get" action="images/FleetXTrainingDocumentationV2.pdf" target="_blank"><input type="submit" id="gobutton" value="Training Documentation" /></form></td></tr>
                    <%
                End if
                'Response.write "billtoid="&billtoid&"<BR>"



        If BillToID="91" or Supervisor = "A" then
            %>
            <tr><td><form method="post" action="orderentry/CancelPage.asp"><input type="submit" id="gobutton" value="Cancel Stockroom Order" /></form></td></tr>
            <%
        End if

        if Supervisor = "A" then 
        If userID="1" then
            %>
            <tr><td><form method="post" action="Admin/AliasCodes.asp"><input type="submit" id="gobutton" value="Alias Codes" /></form></td></tr>
            <%
            End if
         %>
            <!--h2><font color="black">ADMIN MENU:</font></h2-->
                <tr><td><form method="post" action="Admin/RateMaint.asp"><input type="submit" id="gobutton" value="Rate Charge Maintenance" /></form></td></tr>
                <tr><td><form method="post" action="Admin/FuelChargeMaint.asp"><input type="submit" id="gobutton" value="Fuel Charge Maintenance" /></form></td></tr>
                <tr><td><form method="post" action="Admin/AccessorialMaint.asp"><input type="submit" id="gobutton" value="Accessorial Maintenance" /></form></td></tr>
                <tr><td><form method="post" action="Admin/AccessorialCharges.asp"><input type="submit" id="gobutton" value="Accessorial Charges" /></form></td></tr>
                <tr><td><form method="post" action="Admin/NewUserApproval.asp"><input type="submit" id="gobutton" value="New User Approval" /></form></td></tr>

                <tr><td><form method="post" action="Admin/UpdateCarillonInfo.asp"><input type="submit" id="gobutton" value="Update Carillon ID Numbers" /></form></td></tr>
                <tr><td><form method="post" action="Admin/TIBilling.asp"><input type="submit" id="gobutton" value="TI Billing" /></form></td></tr>

                <tr><td><form method="post" action="Admin/UserList.asp"><input type="submit" id="gobutton" value="View/Edit Users" /></form></td></tr>
                <tr><td><form method="post" action="Admin/CompanyList.asp"><input type="submit" id="gobutton" value="View/Edit Companies" /></form></td></tr>
                <tr><td><form method="post" action="OrderEntry/FleetXOrderDispatch.asp"><input type="submit" id="gobutton" value="Dispatch Orders" /></form></td></tr>
                <tr><td><form method="post" action="OrderEntry/FleetXOrderEdit.asp"><input type="submit" id="gobutton" value="Edit Orders" /></form></td></tr>
                <tr><td><form method="post" action="Admin/JobManagement.asp"><input type="submit" id="gobutton" value="Job Management" /></form></td></tr>
                <tr><td><form method="post" action="OrderEntry/FleetXOrderReDispatch.asp"><input type="submit" id="gobutton" value="Re-Dispatch Orders" /></form></td></tr>
                <tr><td><form method="post" action="OrderEntry/LogoutDrivers.asp"><input type="submit" id="gobutton" value="Logout Drivers" /></form></td></tr>
                <tr><td><form target="blank" method="post" action="OrderEntry/FleetXVehicleMonitor.asp"><input type="submit" id="gobutton" value="Live Vehicle Monitor" /></form></td></tr>
                <tr><td><form target="_blank" method="post" action="reporting/FleetXLiveMonitor.asp"><input type="submit" id="gobutton" value="LIVE SHIPMENT MONITOR" /></form></td></tr>

         <%
        end if

 

 %>
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
</form>
<tr><td Height="90%">&nbsp;</td></tr>
<tr>
    <td height="100" class="FleetXGreySection" colspan="2">
        <!-- #include file="BottomSection.asp" -->
    </td>
</tr>
<tr><td height="15" class="FleetXRedSectionSmall" colspan="2" align="center"><%=CopywriteNotice %></td></tr>
</td></tr></table>


</body>
</html>
