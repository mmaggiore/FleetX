<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<%
'Response.Write "test...<BR>"
sbt_id=Request.Form("sbt_id")
if sbt_id>"" then
	Response.Cookies("FleetXPhone")("sBT_ID")=trim(sBT_ID)
	else 
	Response.Cookies("FleetXPhone")("sBT_ID")=""
end if
%>
<link rel="stylesheet" type="text/css" href="mainStyleSheet.css">
<!-- #include file="../v9web/include/ifabsettings.inc" -->
<!-- #include file="driverinfo.inc" -->	
<title>LogistiCorp Driver Vehicle Page</title>
<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<%

VehicleSet=Request.Cookies("FleetXPhone")("VehicleSet")
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
			Response.Cookies("FleetXPhone")("UnitID")=VehicleID
			Response.Cookies("FleetXPhone")("VehicleID")=PenchantVehicleID
			Response.Cookies("FleetXPhone")("VehicleName")=VehicleName
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
			Set Recordset16 = Server.CreateObject("ADODB.Recordset")
			Recordset16.ActiveConnection = Database
			Recordset16.Source = "SELECT * FROM FCUNITS WHERE (UN_ID='"&VehicleID&"')"
			Recordset16.CursorType = 0
			Recordset16.CursorLocation = 2
			Recordset16.LockType = 1
			Recordset16.Open()
			Recordset16_numRows = 0
			'response.write "*****Recordset1.Source="&Recordset1.Source&"<BR>"
				If NOT Recordset16.EOF then 
				
				End if
			Recordset16.Close()
			Set Recordset16 = Nothing										
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
			<form method="post" ID="Form20" name="Form2">
				<input type="hidden" name="VehicleID" value="3" ID="Hidden56">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden57">
				<input type="hidden" name="sbt_id" value="48" ID="Hidden58">	
				<input type="submit" name="submit" value="KWE Spring Creek" ID="Submit19">
			</form>	
			<form method="post" ID="Form21" name="Form2">
				<input type="hidden" name="VehicleID" value="4" ID="Hidden59">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden60">
				<input type="hidden" name="sbt_id" value="48" ID="Hidden61">	
				<input type="submit" name="submit" value="KWE Stafford Bobtail" ID="Submit20">
			</form>	
			<form method="post" ID="Form22" name="Form2">
				<input type="hidden" name="VehicleID" value="5" ID="Hidden62">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden63">
				<input type="hidden" name="sbt_id" value="48" ID="Hidden64">	
				<input type="submit" name="submit" value="KWE Stafford TT" ID="Submit21">
			</form>	
			<form method="post" ID="Form23" name="Form2">
				<input type="hidden" name="VehicleID" value="6" ID="Hidden65">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden66">
				<input type="hidden" name="sbt_id" value="48" ID="Hidden67">	
				<input type="submit" name="submit" value="KWE Sherman" ID="Submit22">
			</form>	
			<form method="post" ID="Form24" name="Form2">
				<input type="hidden" name="VehicleID" value="7" ID="Hidden68">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden69">
				<input type="hidden" name="sbt_id" value="48" ID="Hidden70">	
				<input type="submit" name="submit" value="KWE Alliance" ID="Submit23">
			</form>	
			<form method="post" ID="Form259" name="Form2">
				<input type="hidden" name="VehicleID" value="8" ID="Hidden71">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden72">
				<input type="hidden" name="sbt_id" value="48" ID="Hidden73">	
				<input type="submit" name="submit" value="KWE Campus TT" ID="Submit24">
			</form>																									
			<%
			End if				
			If VehicleSet="w" or VehicleSet="mm" then
			%>
			<form method="post" ID="Form2" name="Form2">
				<input type="hidden" name="sbt_id" value="36" ID="Hidden13">	
				<input type="hidden" name="VehicleID" value="303551" ID="Hidden1">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden2">	
				<input type="submit" name="submit" value="Wafer 1" ID="Submit1">
			</form>	
			<form method="post" ID="Form3" name="Form2">
				<input type="hidden" name="sbt_id" value="36" ID="Hidden14">
				<input type="hidden" name="VehicleID" value="303552" ID="Hidden3">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden4">	
				<input type="submit" name="submit" value="Wafer 2" ID="Submit2">
			</form>					
			<form method="post" ID="Form4" name="Form2">
				<input type="hidden" name="sbt_id" value="36" ID="Hidden15">
				<input type="hidden" name="VehicleID" value="303553" ID="Hidden5">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden6">	
				<input type="submit" name="submit" value="Wafer 3" ID="Submit3">
			</form>
			<form method="post" ID="Form5" name="Form2">
				<input type="hidden" name="sbt_id" value="36" ID="Hidden17">
				<input type="hidden" name="VehicleID" value="303554" ID="Hidden7">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden8">	
				<input type="submit" name="submit" value="Wafer 4" ID="Submit4">
			</form>	
			
			<form method="post" ID="Form28" name="Form2">
				<input type="hidden" name="sbt_id" value="36" ID="Hidden83">
				<input type="hidden" name="VehicleID" value="SCMH" ID="Hidden84">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden85">	
				<input type="submit" name="submit" value="SCB-Material Handler" ID="Submit28">
			</form>				
			
			<form method="post" ID="Form18" name="Form2">
				<input type="hidden" name="sbt_id" value="36" ID="Hidden50">
				<input type="hidden" name="VehicleID" value="OCV" ID="Hidden51">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden52">	
				<input type="submit" name="submit" value="On Call Vehicle" ID="Submit17">
			</form>											
			<%
			End if
			If VehicleSet="srmh" or VehicleSet="sr" or VehicleSet="mm" or VehicleSet="sr48" or VehicleSet="ksmo" then
			%>
			<form method="post" ID="Form6" name="Form2">
				<input type="hidden" name="sbt_id" value="26" ID="Hidden18">
				<input type="hidden" name="VehicleID" value="srv" ID="Hidden9">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden22">	
				<input type="submit" name="submit" value="Stockroom Van" ID="Submit7">
			</form>	
			<form method="post" ID="Form10" name="Form2">
				<input type="hidden" name="sbt_id" value="26" ID="Hidden25">
				<input type="hidden" name="VehicleID" value="ofb" ID="Hidden26">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden27">	
				<input type="submit" name="submit" value="Overflow Bobtail" ID="Submit8">
			</form>
			<form method="post" ID="Form19" name="Form2">
				<input type="hidden" name="sbt_id" value="36" ID="Hidden53">
				<input type="hidden" name="VehicleID" value="OCV" ID="Hidden54">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden55">	
				<input type="submit" name="submit" value="On Call Vehicle" ID="Submit18">
			</form>								
			<%
			End if
			If VehicleSet="sr" or VehicleSet="mm" or VehicleSet="w" then
			%>
			<form method="post" ID="Form27" name="Form2">
				<input type="hidden" name="sbt_id" value="26" ID="Hidden80">
				<input type="hidden" name="VehicleID" value="srvrfab" ID="Hidden81">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden82">	
				<input type="submit" name="submit" value="RFAB Van" ID="Submit27">
			</form>				
			<%
			End if
			If VehicleSet="srmh" or VehicleSet="mm" or VehicleSet="ksmo" then
			%>			
			<form method="post" ID="Form9" name="Form2">
				<input type="hidden" name="sbt_id" value="26" ID="Hidden23">
				<input type="hidden" name="VehicleID" value="srb" ID="Hidden24">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden28">	
				<input type="submit" name="submit" value="SR-Material Handler" ID="Submit9">
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
				<input type="hidden" name="sbt_id" value="75" ID="Hidden29">
				<input type="hidden" name="VehicleID" value="AIMS1" ID="Hidden30">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden37">	
				<input type="submit" name="submit" value="AIMS1" ID="Submit12">
			</form>	
			<form method="post" ID="Form14" name="Form2">
				<input type="hidden" name="sbt_id" value="75" ID="Hidden38">
				<input type="hidden" name="VehicleID" value="AIMS2" ID="Hidden39">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden40">	
				<input type="submit" name="submit" value="AIMS2" ID="Submit13">
			</form>										
			<%
			End if
			If VehicleSet="mm" or VehicleSet="top" or VehicleSet="ktop" then
			%>			
			<form method="post" ID="Form15" name="Form2">
				<input type="hidden" name="sbt_id" value="38" ID="Hidden41">
				<input type="hidden" name="VehicleID" value="AUSTIN" ID="Hidden42">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden43">	
				<input type="submit" name="submit" value="AUSTIN BOBTAIL" ID="Submit14">
			</form>	
			<form method="post" ID="Form16" name="Form2">
				<input type="hidden" name="sbt_id" value="38" ID="Hidden44">
				<input type="hidden" name="VehicleID" value="HOUSTON" ID="Hidden45">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden46">	
				<input type="submit" name="submit" value="HOUSTON BOBTAIL" ID="Submit15">
			</form>										
			<%
			End if	
			If VehicleSet="mm" or VehicleSet="sher" then
			%>			
			<form method="post" ID="Form17" name="Form2">
				<input type="hidden" name="sbt_id" value="45" ID="Hidden47">
				<input type="hidden" name="VehicleID" value="SHERMAN" ID="Hidden48">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden49">	
				<input type="submit" name="submit" value="SHERMAN BOBTAIL" ID="Submit16">
			</form>	
			<%
			End if	
			If VehicleSet="mm" or VehicleSet="demo" then
			%>			
			<form method="post" ID="Form25" name="Form2">
				<input type="hidden" name="sbt_id" value="80" ID="Hidden74">
				<input type="hidden" name="VehicleID" value="199" ID="Hidden75">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden76">	
				<input type="submit" name="submit" value="DEMO VEHICLE" ID="Submit25">
			</form>	
			<form method="post" ID="Form26" name="Form2">
				<input type="hidden" name="sbt_id" value="80" ID="Hidden77">
				<input type="hidden" name="VehicleID" value="198" ID="Hidden78">
				<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden79">	
				<input type="submit" name="submit" value="DEMO VEHICLE #2" ID="Submit26">
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

