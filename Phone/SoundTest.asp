<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
	<%
	soundmethod=Request.QueryString("SoundMethod")
	%>
	<%
	if SoundMethod="1" then%>
	<bgsound src="file://\windows\ringer.wav" loop="-1">
	<%
	Response.Write "<font color='red'>Should have played file://\windows\ringer.wav using the bgsound method<br></font>"
	end if
	If SoundMethod="2" then%>
	<EMBED src="file://\windows\ringer.wav" width="144" height="60" autostart="true" loop="false" hidden="true">
	<%
	Response.Write "<font color='red'>Should have played file://\windows\ringer.wav using the embed method<br></font>"
	end if%>
		<title>Symbol Sound Test</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
	</head>
	<body>
	<br><br>
	This page tries to play a sound located on the mc9097 using two different HTML methods on an .asp page:
	<br><br>
	(To see the specifics of the code, select a choice below, then right click and view source of the page after it
	has displayed)<br><br>
	<a href="SoundTest.asp?SoundMethod=1">bgsound method</a><br><br>
	<a href="SoundTest.asp?SoundMethod=2">embed method</a><br><br>
	</body>
</html>
