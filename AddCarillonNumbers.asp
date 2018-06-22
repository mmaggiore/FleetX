<html>
<head>
<title>Fleet Express - Retrieve Password</title>
<!-- #include file="fleetexpress.inc" -->
<link rel="stylesheet" type="text/css" href="css/Style.css">

</head>

<body leftMargin="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" bgcolor="#FFFFFF" text="#000000" onload="document.FindUser.requiredemail.focus();">
	<table align="center" border="0" bordercolor="red" cellpadding="0" cellspacing="0" ID="Table1">
<tr><td align="center">
<table border="0" bordercolor="brown" cellspacing="0" cellpadding="0" align="center">
  <tr><td>&nbsp;</td></tr>
	<tr><td align="center" class="MainPageText">
<b>Pre-Existing Locations</b>
</td></tr>
<tr><td>&nbsp;</td></tr>
<%
Whatever=Request.Form("Whatever")
							Set oConn = Server.CreateObject("ADODB.Connection")
							oConn.ConnectionTimeout = 100
							oConn.Provider = "MSDASQL"
							oConn.Open DATABASE
								l_cSQL2 = "select CompanyID, CarillonID from Sheet1$  " 
										'if trim(displayusername)="comps" or trim(displayusername)="Compugraphics"  then 
										'l_cSQL2 = l_cSQL2 & "  AND st_id<>'CPGP'" 
										'end if
								'Response.Write "l_cSQL2="&l_cSQL2&"<BR>"
								SET oRs = oConn.Execute(l_cSql2)

								Do While not oRs.EOF
                                CompanyID=oRs("CompanyID")
                                CarillonID=oRs("CarillonID")
                                Response.write "CompanyID="&CompanyID&"<BR>"
                                Response.write "CarillonID="&CarillonID&"<BR>"

If whatever="whatever" then
                         Set oConn2 = Server.CreateObject("ADODB.Connection")
						oConn2.ConnectionTimeout = 100
						oConn2.Provider = "MSDASQL"
						oConn2.Open DATABASE
						' 7/26/04 KK: Added canceljob functionality to be able to update the status if cancel button is pressed.
						' 11/30/04 DEC: Changed from CAN/98 TO DEL/99 to be consistent with dispatchOffice			
							l_cSQL = "UPDATE PreExistingCompanies SET CarillonID = '"&CarillonID&"' WHERE CompanyID = '" & CompanyID & "'"
							Response.write "l_cSQL="&l_cSQL&"<BR>"
							oConn2.Execute(l_cSQL)
						Set oConn2=Nothing
End if





								oRs.movenext
								LOOP
							Set oConn=Nothing

                            Response.write "DONE!!!!!<BR>"
 %>

</table>
</td></tr>
</table>

</body>
</html>
