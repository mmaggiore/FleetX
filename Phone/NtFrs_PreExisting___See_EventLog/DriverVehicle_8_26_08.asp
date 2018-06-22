<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<%
'Response.Write "test...<BR>"
sbt_id=Request.Form("sbt_id")
if sbt_id>"" then
	Response.Cookies("Phone")("sBT_ID")=trim(sBT_ID)
	else 
	Response.Cookies("Phone")("sBT_ID")=""
end if
%>
<link rel="stylesheet" type="text/css" href="mainStyleSheet.css">
<!-- #include file="../v9web/include/ifabsettings.inc" -->
<!-- #include file="driverinfo.inc" -->	
<title>LogistiCorp Driver Vehicle Page</title>
<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<%

VehicleSet=Request.Cookies("Phone")("VehicleSet")
'Response.Write "Database="&Database&"<BR>"
FakeSubmit=Request.Form("FakeSubmit")
If Fakesubmit>"" then
	
	'REsponse.Write "Got here too!"
	
	VehicleID=Request.Form("VehicleID")
End if
'response.write "Database="&Database&"<BR>"
'Response.Write "vehicleID="&VehicleID&"<BR>"
If VehicleID>"" and Fakesubmit>"" then
	Set Recordset1 = Server.CreateObject("ADODB.Recordset")
	Recordset1.ActiveConnection = Database
	Recordset1.Source = "SELECT * FROM FCUNITS WHERE (UN_ID='"&VehicleID&"')"
	Recordset1.CursorType = 0
	Recordset1.CursorLocation = 2
	Recordset1.LockType = 1
	Recordset1.Open()
	Recordset1_numRows = 0
	'response.write "*****Recordset1.Source="&Recordset1.Source&"<BR>"
		If NOT Recordset1.EOF then 
			'response.write "GOT HERE!<BR>"
			VehicleName=Recordset1("UN_DESC")
			PenchantVehicleID=Recordset1("un_dr_id")
			''''''''''''''''''''''''''''
			'response.write "penchantVehicleID="&penchantVehicleID&"<BR>"
			'response.write "VehicleID="&VehicleID&"<BR>"
			'response.write "VehicleName="&VehicleName&"<BR>"
			'''''''''''''''''''''''''''''
			Response.Cookies("Phone")("UnitID")=VehicleID
			Response.Cookies("Phone")("VehicleID")=PenchantVehicleID
			Response.Cookies("Phone")("VehicleName")=VehicleName
			Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				RSEVENTS2.Open "DriverLog", Intranet, 2, 2
				RSEVENTS2.addnew	
				RSEVENTS2("DriverID")=UserID		
				RSEVENTS2("VehicleID") = VehicleID
				'RSEVENTS2("LogInOut") = "o"
				RSEVENTS2("LogTime")=Now()		
				RSEVENTS2("LogStatus") = "c"
				RSEVENTS2.update
				RSEVENTS2.close			
			set RSEVENTS2 = nothing	
			'''Recordset1.Close()
			'''Set Recordset1 = Nothing			
			Response.Redirect("Default.asp")
			ELSE
			ErrorMessage="That is not a valid vehicle ID"			
		End if
	Recordset1.Close()
	Set Recordset1 = Nothing

End if
%>
</head>
<body OnLoad=document.Form1.VehicleID.focus()>
<table cellspacing="0" cellpadding="0" width="300" border="0" bordercolor="black" ID="Table1">
	<!--
	<tr><td align="center" colspan="9"><a href="default.asp" class="mainpagelink">Return Home</a></td></tr>
	-->
	<tr>
		<td class="mainpagetextboldcenter" colspan="2" align="center">
			Driver Vehicle Page
		</td>
	</tr>
	<tr><td>&nbsp;</td></tr>
	<tr><td align="left" colspan="2">

	<%If VehicleSet="" or isnull(VehicleSet) then
	''response.write "GOT HERE!"
	%>
		<form method="post" ID="Form1" name="Form1">
		<div class="purpleseparator"> 
			<table border="0" cellpadding="2" cellspacing="0" ID="Table4" width="100%" bordercolor="blue">
				<tr> 
					<td class="mainpagetextboldright" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
				</tr>
				<tr>
					<td class='mainpagetextboldcenter' colspan="2" nowrap align="center">SCAN in vehicle code</td>
				</tr>
				<tr>
					<td colspan='2' class='generalcontent' align="center">
						<input type="password" maxlength='25' size='25' name='VehicleID' id='VehicleID' onBlur="form.submit()">
						<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden16">
					</td>
				</tr>
				<%if errormessage>"" then%>
					<tr>
						<td class='generalcontenthighlight'colspan='2' align="center"><font color="red"><br><b><%=ErrorMessage%></b><br><br></font></td>
					</tr>
				<%end if%>
				<!--									
				<tr>
					<td colspan="2" align="center" class='generalcontent'>
						<input type="submit" name="submit" value="submit" ID="Submit1">									
					</td>
				</tr>
				-->
				<tr> 
					<td class="subheader" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
				</tr>
			</table>
		</div>
		</form>
			<%else
			If VehicleSet="48" or VehicleSet="mm" or VehicleSet="sr48" or VehicleSet="ksmo" or VehicleSet="ktop" then
			%>
			<form method="post" ID="Form7" name="Form2">
				<input type="hidden" name="VehicleID" value="1" ID="Hidden10">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden11">
				<input type="hidden" name="sbt_id" value="48" ID="Hidden12">	
				<input type="submit" name="submit" value="KWE Van" ID="Submit5">
			</form>	
			<form method="post" ID="Form8" name="Form2">
				<input type="hidden" name="VehicleID" value="2" ID="Hidden19">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden20">
				<input type="hidden" name="sbt_id" value="48" ID="Hidden21">	
				<input type="submit" name="submit" value="KWE PDC Van" ID="Submit6">
			</form>							
			<%
			End if				
			If VehicleSet="w" or VehicleSet="mm" then
			%>
			<form method="post" ID="Form2" name="Form2">
				<input type="hidden" name="sbt_id" value="36" ID="Hidden13">	
				<input type="hidden" name="VehicleID" value="303551">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden1">	
				<input type="submit" name="submit" value="Wafer 1">
			</form>	
			<form method="post" ID="Form3" name="Form2">
				<input type="hidden" name="sbt_id" value="36" ID="Hidden14">
				<input type="hidden" name="VehicleID" value="303552" ID="Hidden2">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden3">	
				<input type="submit" name="submit" value="Wafer 2" ID="Submit1">
			</form>					
			<form method="post" ID="Form4" name="Form2">
				<input type="hidden" name="sbt_id" value="36" ID="Hidden15">
				<input type="hidden" name="VehicleID" value="303553" ID="Hidden4">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden5">	
				<input type="submit" name="submit" value="Wafer 3" ID="Submit2">
			</form>
			<form method="post" ID="Form5" name="Form2">
				<input type="hidden" name="sbt_id" value="36" ID="Hidden17">
				<input type="hidden" name="VehicleID" value="303554" ID="Hidden6">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden7">	
				<input type="submit" name="submit" value="Wafer 4" ID="Submit3">
			</form>								
			<%
			End if
			If VehicleSet="srmh" or VehicleSet="sr" or VehicleSet="mm" or VehicleSet="sr48" or VehicleSet="ksmo" then
			%>
			<form method="post" ID="Form6" name="Form2">
				<input type="hidden" name="sbt_id" value="26" ID="Hidden18">
				<input type="hidden" name="VehicleID" value="srv" ID="Hidden8">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden9">	
				<input type="submit" name="submit" value="Stockroom Van" ID="Submit4">
			</form>	
			<form method="post" ID="Form10" name="Form2">
				<input type="hidden" name="sbt_id" value="26" ID="Hidden25">
				<input type="hidden" name="VehicleID" value="ofb" ID="Hidden26">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden27">	
				<input type="submit" name="submit" value="Overflow Bobtail" ID="Submit8">
			</form>				
			<%
			End if
			If VehicleSet="srmh" or VehicleSet="mm" or VehicleSet="ksmo" then
			%>			
			<form method="post" ID="Form9" name="Form2">
				<input type="hidden" name="sbt_id" value="26" ID="Hidden22">
				<input type="hidden" name="VehicleID" value="srb" ID="Hidden23">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden24">	
				<input type="submit" name="submit" value="SR-Material Handler" ID="Submit7">
			</form>							
			<%
			End if
			If VehicleSet="mm" or VehicleSet="ross" then
			%>			
			<form method="post" ID="Form12" name="Form2">
				<input type="hidden" name="sbt_id" value="14" ID="Hidden31">
				<input type="hidden" name="VehicleID" value="ABROSS" ID="Hidden32">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden33">	
				<input type="submit" name="submit" value="ABBROSS1" ID="Submit10">
			</form>	
			<form method="post" ID="Form13" name="Form2">
				<input type="hidden" name="sbt_id" value="14" ID="Hidden34">
				<input type="hidden" name="VehicleID" value="ABROSS2" ID="Hidden35">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden36">	
				<input type="submit" name="submit" value="ABBROSS2" ID="Submit11">
			</form>										
			<%
			End if	
			If VehicleSet="mm" or VehicleSet="aims" then
			%>			
			<form method="post" ID="Form11" name="Form2">
				<input type="hidden" name="sbt_id" value="75" ID="Hidden28">
				<input type="hidden" name="VehicleID" value="AIMS1" ID="Hidden29">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden30">	
				<input type="submit" name="submit" value="AIMS1" ID="Submit9">
			</form>	
			<form method="post" ID="Form14" name="Form2">
				<input type="hidden" name="sbt_id" value="75" ID="Hidden37">
				<input type="hidden" name="VehicleID" value="AIMS2" ID="Hidden38">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden39">	
				<input type="submit" name="submit" value="AIMS2" ID="Submit12">
			</form>										
			<%
			End if
			If VehicleSet="mm" or VehicleSet="top" or VehicleSet="ktop" then
			%>			
			<form method="post" ID="Form15" name="Form2">
				<input type="hidden" name="sbt_id" value="38" ID="Hidden40">
				<input type="hidden" name="VehicleID" value="AUSTIN" ID="Hidden41">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden42">	
				<input type="submit" name="submit" value="AUSTIN BOBTAIL" ID="Submit13">
			</form>	
			<form method="post" ID="Form16" name="Form2">
				<input type="hidden" name="sbt_id" value="38" ID="Hidden43">
				<input type="hidden" name="VehicleID" value="HOUSTON" ID="Hidden44">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden45">	
				<input type="submit" name="submit" value="HOUSTON BOBTAIL" ID="Submit14">
			</form>										
			<%
			End if									
		End if%>
	</td></tr>

	<!--
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="Submit" ID="Submit1"></td></tr>
	-->
</table>



</body>
</html>

