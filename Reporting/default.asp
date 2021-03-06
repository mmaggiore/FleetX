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

   		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.ConnectionTimeout = 100
		oConn.Provider = "MSDASQL"
		oConn.Open DATABASE
			l_cSQL = "Select RequestorType FROM PreExistingRequestor WHERE  RequestorID='"& UserID &"'"
            'Response.Write "l_cSql="&l_cSql&"<BR>"
			SET oRs = oConn.Execute(l_cSql)
					if NOT oRs.EOF then
                        RequestorType=oRs("RequestorType")
                    End if								
		Set oConn=Nothing
        Set oRS=Nothing

        'Response.write "RequestorType="&RequestorType&"<BR>"
%>
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

<table border="0" bordercolor="green" Cellspacing="0" Cellpadding="0" align="left" width="100%">
 <tr><td align="left" colspan="2" bgcolor="<%=HeaderBorderColor%>"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>

        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetXRedSection" colspan="2"><%=PageTitle%></td></tr>

        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
        <!--tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressBodyWhite" colspan="2">In order to reach the Fleet Express Order page, please correctly type in the green verification code in the supplied text box and click "Submit."</td></tr-->
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
<tr><td>
<table  border="0" bordercolor="blue" align="left" class="MainPageText" width="100%">
	<tr height="40">
		<td width="650">&nbsp;</td>
	</tr>


    <form method="post" action="FleetXLiveMonitor.asp">
    <tr><td align="left" width="100%">
       <input type="submit" id="gobutton" value="LIVE SHIPMENT MONITOR" />
    </td></tr>
    </form>

    <form method="post" action="FleetXLiveMonitor_Chain.asp">
    <tr><td align="left" colspan="2" height="5"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
    <tr><td align="left" width="100%">
       <input type="submit" id="gobutton" value="CHAIN OF CUSTODY LIVE SHIPMENT MONITOR" />
    </td></tr>
    </form>

    <form method="post" action="FleetX_SR_Metrics.asp">
    <tr><td align="left" colspan="2" height="5"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
    <tr><td align="left" width="100%"><input type="submit" id="gobutton" value="Metrics - SR" /></td></tr>
    </form>
    <%If trim(RequestorType)="A" then %>

    <form method="post" action="FleetX_Unbilled.asp">
        <tr><td align="left" colspan="2" height="5"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
        <tr><td align="left" width="100%"><input type="submit" id="gobutton" value="ADMIN - UNBILLED - COURIER/FREIGHT" /></td></tr>
    </form>

    <form method="post" action="FleetX_All_Metrics_Admin.asp">
        <tr><td align="left" colspan="2" height="5"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
        <tr><td align="left" width="100%"><input type="submit" id="gobutton" value="ADMIN - BILLING - ALL" /></td></tr>
    </form>
    <!--
    <form method="post" action="FleetX_Freight_Metrics_Admin.asp">
        <tr><td align="left" colspan="2" height="5"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
        <tr><td align="left" width="100%"><input type="submit" id="gobutton" value="ADMIN - BILLING - FREIGHT" /></td></tr>
    </form>
    -->
    <form method="post" action="FleetX_HotShots_Freight_Metrics_Admin.asp">
        <tr><td align="left" colspan="2" height="5"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
        <tr><td align="left" width="100%"><input type="submit" id="gobutton" value="ADMIN - BILLING - FREIGHT" /></td></tr>
    </form>
    <form method="post" action="FleetX_SRandCourier_Metrics_Admin.asp">
        <tr><td align="left" colspan="2" height="5"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
        <tr><td align="left" width="100%"><input type="submit" id="gobutton" value="ADMIN - BILLING - SR AND COURIER" /></td></tr>
    </form>

    <form method="post" action="FleetX_SR_Metrics_Admin.asp">
        <tr><td align="left" colspan="2" height="5"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
        <tr><td align="left" width="100%"><input type="submit" id="gobutton" value="ADMIN - BILLING - SR" /></td></tr>
    </form>

    <form method="post" action="FleetX_Standing_Metrics_Admin.asp">
        <tr><td align="left" colspan="2" height="5"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
        <tr><td align="left" width="100%"><input type="submit" id="gobutton" value="ADMIN - BILLING - STANDING ORDERS" /></td></tr>
    </form>

    <form method="post" action="FleetX_Metrics_ByDriver.asp">
        <tr><td align="left" colspan="2" height="5"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
        <tr><td align="left" width="100%"><input type="submit" id="gobutton" value="METRICS BY DRIVER" /></td></tr>
    </form>
    <form method="post" action="FleetX_Courier_Histogram.asp">
        <tr><td align="left" colspan="2" height="5"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
        <tr><td align="left" width="100%"><input type="submit" id="gobutton" value="HISTOGRAM - COURIER" /></td></tr>
    </form>
    <%End if %>

 
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

