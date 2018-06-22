<%@ Language=VBScript %>
<!-- 10/08/2005 Span-Renuka included this screen for CCF 2023 -->
<!-- #include file="../DedicatedFleets/include/checkstring.inc" -->
<!-- #include file="../DedicatedFleets/include/custom.inc" -->
<!-- #include file="../DedicatedFleets/include/ifabsettings.inc" -->
<!-- #include file="driverinfo.inc" -->	
<HTML>
	<HEAD>
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
		<link rel="stylesheet" type="text/css" href="mainStyleSheet.css">
		<%
		Dim ListOfFrom(200)
		Dim ListOfToM(200)
		Dim ListOfTo(200)
		DriverID=Request.Form("DriverID")
		BillToID=Request.Cookies("Phone")("sBT_ID")	
		mark=Request.QueryString("Mark")
		%>
	<SCRIPT type="text/javascript" language="JavaScript">
	//Span-Renuka Fucntion to Open the Pop up windows
	function EvalSound(soundobj) {
	var thissound=document.getElementByID(soundobj);
	thissound.Play();
	}
	</script>		
	</HEAD>
	<embed src="success.wav" autostart=false width=0 height=0 id="sound1" enablejavascript="true">
	<BODY leftMargin=0 TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">

		<form>
		<input type="button" value="Play Sound" onclick="EvalSound('sound1')">
		</form>		
	</BODY>
</HTML>
