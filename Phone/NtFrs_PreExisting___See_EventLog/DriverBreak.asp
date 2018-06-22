<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
	<!--	
	<bgsound src="file://\sounds\alert1.wav" >
	<bgsound src="file://\windows\ringer.wav" loop="-1">
	-->
	<!--
	<EMBED src="file://\windows\ringer.wav" width="144" height="60" autostart="true" loop="false" hidden="true">
	-->
		<META HTTP-EQUIV="PRAGMA" CONTENT="NO-CACHE">
		<meta http-equiv="refresh" content="60" />

<script language="JavaScript">
<!--

var sURL = unescape(window.location.pathname);

function doLoad()
{
    // the timeout value should be the same as in the "refresh" meta-tag
    setTimeout( "refresh()", 65*1000 );
}

function refresh()
{
    //  This version of the refresh function will cause a new
    //  entry in the visitor's history.  It is provided for
    //  those browsers that only support JavaScript 1.0.
    //
    window.location.href = sURL;
}

//-->
</script>

<script language="JavaScript1.1">
<!--
function refresh()
{
    //  This version does NOT cause an entry in the browser's
    //  page view history.  Most browsers will always retrieve
    //  the document from the web-server whether it is already
    //  in the browsers page-cache or not.
    //  
    window.location.replace( sURL );
}
//-->
</script>

<script language="JavaScript1.2">
<!--
function refresh()
{
    //  This version of the refresh function will be invoked
    //  for browsers that support JavaScript version 1.2
    //
    
    //  The argument to the location.reload function determines
    //  if the browser should retrieve the document from the
    //  web-server.  In our example all we need to do is cause
    //  the JavaScript block in the document body to be
    //  re-evaluated.  If we needed to pull the document from
    //  the web-server again (such as where the document contents
    //  change dynamically) we would pass the argument as 'true'.
    //  
    window.location.reload( false );
}
//-->
</script>
		
		<link rel="stylesheet" type="text/css" href="mainStyleSheet.css">
		<!-- #include file="../v9web/include/ifabsettings.inc" -->
		<!-- #include file="driverinfo.inc" -->	
		<title>Logisticorp Driver Break Page</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<%
''''''''''''''''''''''''''''''''''''''''''
SecureYes = Request.ServerVariables ("HTTPS")
'If SecureYes="off" then
If SecureYes="on" then
	''''''''''''''''''''''''''''''''''''''''''''
	something=Request.ServerVariables("HTTP_HOST") 
	Something=left(Something,4) 
	If lcase(something)="test" then 
		Response.redirect("http://test.logisticorp.us/phone/default.asp")
		'Response.Write "GOT HERE!!!" 
		MPMSendEmail="n" 
		else
		Response.redirect("http://www.logisticorp.us/phone/default.asp")
	End if 
	'Response.Write "Something="&Something&"<BR>"
	'Response.Write "MPMSendEmail="&MPMSendEmail&"<BR>"

	'''''''''''''''''''''''''''''''''''''''''''
		
End if
''''''''''''''''''''''''''''''''''''''''''		
		Response.Cookies("Phone")("PageStatus")=""
		Response.Cookies("Phone")("AliasCode")=""
		Response.Cookies("Phone")("FakeSubmit")=""
		mark=request.QueryString("mark")
		BreakType=Request.Querystring("a")
		BreakTime=Request.Querystring("b")
		InsertRecord=request.QueryString("c")
		RedirectVAR=Request.Form("RedirectVAR")
		If RedirectVAR="" then
		    RedirectVAR=Request.QueryString("RedirectVAR")
		End if
		If RedirectVAR="y" then
		    ''''''''''''''''''''''''''''''''''''''''''''''
		    ''''''''''''ADD DATABASE STUFF HERE!!!!!!!
			Set oConn = Server.CreateObject("ADODB.Connection")
			oConn.ConnectionTimeout = 100
			oConn.Provider = "MSDASQL"
			oConn.Open DATABASE
				l_cSQL = "UPDATE breaktable SET EndTime = '"& Now() &"' WHERE Userid = '" & UserID & "' and EndTime is NULL"
				oConn.Execute(l_cSQL)
			Set oConn=Nothing
			'Response.Write "l_cSQL="& l_cSQL &"<BR>"
			Response.Redirect("Default.asp")	
			'Response.write "Redirect#3<BR>"	    
		    ''''''''''''''''''''''''''''''''''''''''''''''
		End if
		if mark="y" then
			response.write "helloooo!"
			response.write "VehicleID="&VehicleID&"<BR>"
		End if
		Select Case BreakType
		    Case "l"
		        TimeLeftMessage="<b>Time left on lunch:</b>"
		    Case "b"
		        TimeLeftMessage="<b>Time left on break:</b>"
		End Select
		'Response.Write "BreakType="&BreakType&"<br>"
		'Response.Write "BreakTime="&BreakTime&"<br>"
		'Response.Write "now="&now()&"<br>"
		'-----------------------------------------------------------------
					
		'-----------------------------------------------------------------
		If InsertRecord="y" then
		Response.Write "GOT HERE!!!!<BR>"
		    'Response.Write "VehicleID="&VehicleID&"<br>"
			Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				RSEVENTS2.Open "BreakTable", DATABASE, 2, 2
				RSEVENTS2.addnew
				RSEVENTS2("UserID")=UserID
				RSEVENTS2("BreakType")=BreakType
				RSEVENTS2("StartTime")=Now()
				'RSEVENTS2("EndTime")=Now()
				RSEVENTS2("Status")="c"								
				RSEVENTS2.update
				RSEVENTS2.close			
			set RSEVENTS2 = nothing	
			Response.Redirect("DriverBreak.asp?a="&BreakType&"&b="&BreakTime)
			'Response.write "Redirect#2<BR>"
		End if	
        TimeLeft=datediff("n",BreakTime,now())
        DisplayTimeLeft=Abs(TimeLeft)
        'Response.Write "<BR>"&TimeLeft
        If Timeleft>=0 then 
            'Response.Write "DONE!!!!"
            Response.redirect("DriverBreak.asp?RedirectVar=y")
           'Response.write "Redirect#1<BR>"
        End if
		
		%>
	</head>
<body onload="doLoad()">
	<table cellspacing="0" cellpadding="0" border="0" width="300">
			<tr>
				<td class="mainpagetextboldcenter">
					<font color="black">Last update: <%=Time()%></font>
				</td>
			</tr>	
	    <tr><td align="center"><%=TimeLeftMessage%></td></tr>
        <tr><td class="HugeCountdownBlue"><%=DisplayTimeLeft%></td></tr>
       <tr><td align="center"><b>Minutes</b></td></tr>
       <tr><td>&nbsp;</td></tr>
       <tr>
            <td align="center">
                <form method="post" action="DriverBreak.asp">
                    <input type="submit" value="Leave break early">
                    <input type="hidden" name="redirectVar" value="y">
                </form>
            </td>
       </tr>
	</table>
	</body>
</html>
