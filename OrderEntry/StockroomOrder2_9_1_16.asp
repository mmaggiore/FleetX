<html>
<head>

<!-- #include file="../fleetexpress.inc" -->
<link rel="stylesheet" type="text/css" href="../css/Style.css">
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
    HighlightedField="RequestorName"
    CurrentDateTime=Now()
    PageTitle="ORDER PAGE - DEFINE SHIPMENT"



l_cJobNum=valid8(Request.Querystring("l_cJobNum"))
If trim(l_cJobNum)="" then
    l_cJobNum=valid8(Request.Form("l_cJobNum"))
End if
ShipmentType=valid8(lcase(Request.form("ShipmentType")))
'Response.write "Line 41 ShipmentType="&ShipmentType&"<BR>"
If trim(ShipmentType)>"" then
    Select Case ShipmentType
        Case "light package"
            VehicleType="Van"
        Case "heavy freight"
            VehicleType="Bobtail"
    End Select
    'Response.write "Line 48 VehicleType="&VehicleType&"<BR>"
	    Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
	    RSEVENTS.Open "FCFGTHD", Database, 2, 2
	    RSEVENTS.Find "FH_ID='"& l_cJobNum &"'"
		    RSEVENTS("fh_status") = "RAP"
            RSEVENTS("fh_statcode") = "2"
            RSEVENTS("fh_user4") = VehicleType
	    RSEVENTS.update
	    RSEVENTS.close
	    set RSEVENTS = nothing
    Response.Redirect("../include/fnlrecap.asp?l_cJobNum=" & l_cJobNum & " ")

End if

'Response.write "l_cJobNum="&l_cJobNum&"<BR>"
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
<table  border="0" bordercolor="blue" align="center" class="MainPageText" width="100%">
	<tr height="40">
		<td width="650">&nbsp;</td>
	</tr>



    <tr><td align="center">
   <form method="post" action="StockroomOrder2.asp">
        <table>
            <tr>
                <td colspan="3">
                Which most closely describes your shipment?
                </td>
            </tr>
            <tr><td>&nbsp;</td></tr>
             <tr><td>&nbsp;</td></tr>
            <tr>
                <td>
                <img src="../images/lightfreight.gif" height="205" width="300" />
                </td>
                <td width="80">&nbsp;</td>
                <td>
                <img src="../images/heavyfreight.jpg" height="192" width="263" />
                </td>
            </tr>
             <tr><td>&nbsp;</td></tr>
            <tr>
                <td align="center">
                <input id="gobutton" name="ShipmentType" type="submit" value="Light Package" />
                </td>
                <td>&nbsp;</td>
                <td align="center">
                <input id="gobutton" name="ShipmentType" type="submit" value="Heavy Freight" />
                </td>
            </tr>
            <input type="hidden" name="l_cJobNum" value="<%=l_cJobNum %>" />
        </table>
    </form>
    
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
