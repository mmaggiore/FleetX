<%@ Language=VBScript %>
<!-- #include file="FleetX.inc" -->
<!-- #include file="driverinfo.inc" -->	
<HTML>
	<HEAD>
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
		<link rel="stylesheet" type="text/css" href="mainStyleSheet.css">
<%
		FORMJOBSTATUS=TRIM(Request.Form("FORMJOBSTATUS"))
		AcknowledgeIt=Request.Form("AcknowledgeIt")
		DriverID=Request.Form("DriverID")
		LocationCode=Request.Form("LocationCode")
		Submit=Request.Form("Submit")
		PageStatus=Request.Form("PageStatus")
		PageStatus="loggedin"
		txtJobNumber=Request.Form("txtJobNumber")
		If Submit="submit" then
			If DriverID="" then
				ErrorMessage="You must provide your driver id"
			End if
		If ErrorMessage="" then PageStatus="loggedin" End if
		End if
		'Response.Write "userid="&userID&"<br>"
		'Response.Write "vehicleid="&vehicleID&"<br>"
		'Response.Write "unitid="&unitid&"<br>"
		'Response.Write "driverid="&driverid&"<br>"
		%>
	</HEAD>
	<BODY leftMargin=0 TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
    <!-- #include file="LogoSection.asp" -->
		<%
		

'-------------------STARTS THE OTHER ORDERS IN THE PHONE
%>
                <form method="post" action="default.asp" ID="Form2">
					<table cellpadding="0" width="300" cellspacing="0" bordercolor="red" border="0" align="left" ID="Table5">
                    <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>
						
                        <tr><td align="center" colspan="3"><input type="submit" value="Return to Menu" id="gobutton" NAME="Submit2"></td></tr>
        <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>            
        <tr>
		    <td class="FleetXRedSection" colspan="2" align="center">
			    CONTACT NUMBERS
		    </td>
	    </tr>
        <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>	
						<tr><td class="mainpagetext">Mark Maggiore<BR>Phone Issues</td></tr>
						<tr><td class="mainpagetext">817-591-2956</td></tr>
						<tr><td class="mainpagetext"><hr></td></tr>
						<tr><td class="mainpagetext">David Mercer<BR>Wafers/On Campus Reticles</td></tr>
						<tr><td class="mainpagetext">214-882-1292</td></tr>
						<tr><td><hr></td></tr>
						<tr><td class="mainpagetext">Keith Chitwood<BR>Wafers/Reticles/Stockroom</td></tr>
						<tr><td class="mainpagetext">214-882-5423</td></tr>
						<tr><td><hr></td></tr>
						<tr><td class="mainpagetext"><B>KWE after hours contacts:</b></td></tr>
						<tr><td class="mainpagetext">Call contacts in order.<br>If no answer, leave a message.<br>After 5 minutes, call next person on the list</td></tr>
						<tr><td class="mainpagetext">THOY-469-387-6730<br>Mary Busch-469-438-2807</td></tr>
						<tr><td><hr></td></tr>
						<tr><td class="mainpagetext">Photronics Shipping</td></tr>
						<tr><td class="mainpagetext">972-889-6222</td></tr>	
						<tr><td><hr></td></tr>
						<tr><td class="mainpagetext">Toppan Shipping</td></tr>
						<tr><td class="mainpagetext">512-310-6154</td></tr>																	
						<tr><td><hr></td></tr>												
						</table>				
	        </form>
			
			
	</BODY>
</HTML>
