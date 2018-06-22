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
		<meta http-equiv="refresh" content="120" />




		
		<link rel="stylesheet" type="text/css" href="mainStyleSheet.css">
		<!-- #include file="../DedicatedFleets/include/ifabsettings.inc" -->
		<!-- #include file="driverinfo.inc" -->	
		<title>Logisticorp Driver Home Page</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<%
''''''''''''''''''''''''''''''''''''''''''
SomeVariable=Request.ServerVariables ("HTTP_USER_AGENT")
'Response.Write "SomeVariable="& SomeVariable &"<BR>"
If SomeVariable="Motorola_ES405B/20 Mozilla/4.0 (compatible; MSIE 6.0; Windows CE; IEMobile 8.12; MSIEMobile 6.5)" then
    Response.Write "<br><br>YES<BR>"
    Else
    Response.Write "<br><br>NO<BR>"
End if
''''''''''''SETS VEHICLE JOB COUNTS ETC. TO ZERO''''''''''''
%>
</head>
<body>
<form method="post" action="default.asp">
<input type="submit" value="click me" /></form>
</body>
</html>
