<%@ Language=VBScript %>
<!-- 10/08/2005 Span-Renuka included this screen for CCF 2023 -->
<!-- #include file="../v9web/include/checkstring.inc" -->
<!-- #include file="../v9web/include/custom.inc" -->
<!-- #include file="../v9web/include/ifabsettings.inc" -->
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
		<%
		

'-------------------STARTS THE OTHER ORDERS IN THE PHONE
%>
					<table cellpadding="0" width="300" cellspacing="0" bordercolor="red" border="0" align="left" ID="Table5">
						<tr><td align="center" colspan="3"><form method="post" action="default.asp" ID="Form2"><input type="submit" value="Return to Menu" ID="Submit2" NAME="Submit2"></form></td></tr>
						<tr>
							<td align="center" class="purpleseparator" colspan="3"><b>Contact Numbers</b></td>
						</tr>
						<tr><td>Mark Maggiore<BR>Phone Issues</td></tr>
						<tr><td>214-956-0400 xt. 212</td></tr>
						<tr><td><hr></td></tr>
						<tr><td>David Mercer<BR>Wafers/On Campus Reticles</td></tr>
						<tr><td>214-882-1292</td></tr>
						<tr><td><hr></td></tr>
						<tr><td>Keith Chitwood<BR>KWE/Stockroom/Off Campus Reticles</td></tr>
						<tr><td>214-882-5423</td></tr>
						<tr><td><hr></td></tr>
						<tr><td><B>KWE after hours contacts:</b></td></tr>
						<tr><td>Call contacts in order.<br>If no answer, leave a message.<br>After 5 minutes, call next person on the list</td></tr>
						<tr><td>THOY-469-387-6730<br>YADI-972-510-4946</td></tr>
						<tr><td><hr></td></tr>
						<tr><td>Photronics Shipping</td></tr>
						<tr><td>972-889-6222</td></tr>	
						<tr><td><hr></td></tr>
						<tr><td>Toppan Shipping</td></tr>
						<tr><td>512-310-6154</td></tr>																	
						<tr><td><hr></td></tr>												
						</table>				
	
			
			
	</BODY>
</HTML>
