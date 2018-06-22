<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<%
''''This makes page avoid the driver log in check...
LoginCheck="n"

  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"

'Response.Write "test...<BR>"
sbt_id=Request.Form("sbt_id")
if sbt_id>"" then
	Response.Cookies("FleetXPhone")("sBT_ID")=trim(sBT_ID)
	else 
	Response.Cookies("FleetXPhone")("sBT_ID")=""
end if
%>
<!-- #include file="FleetX.inc" -->
<!-- #include file="driverinfo.inc" -->	
<title>LogistiCorp Driver Vehicle Page</title>
<script type="text/javascript">
    function formSubmit() {
        document.getElementById("Form1").submit()
    }
</script>

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
Logout=Request.form("Logout")
''''If logout="y" then
'Response.write "got here Line 41<BR>"
TakenVehicle=Request.form("TakenVehicle")
If trim(TakenVehicle)>"" then
    VehicleID=TakenVehicle
    FakeSubmit="y"
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.ConnectionTimeout = 100
	oConn.Provider = "MSDASQL"
	oConn.Open DATABASE
  Set oConni = Server.CreateObject("ADODB.Connection")
  oConni.ConnectionTimeout = 100
  oConni.Provider = "MSDASQL"
  oConni.Open INTRANET

        'l_cSQL = "UPDATE AvailableVehicles SET logouttime=getdate(), availablestatus = 'x' WHERE VehicleID = '" & TakenVehicle & "' and AvailableStatus='c'"

       'check is current user is either Mark or Betty don't deactivate other drivers from the vehicle
       if UserID <> 1 and UserID <> 508 then
          l_cSQL = "UPDATE AvailableVehicles SET logouttime=getdate(), availablestatus = 'x' "
          l_cSQL = l_cSQL & "WHERE VehicleID = '" & TakenVehicle & "' and AvailableStatus='c' AND DriverID<> '1' AND DriverID <> '508'"
	        'Response.write "55 l_cSQL="&l_cSQL&"<BR>"
          oConn.Execute(l_cSQL)
        end if
        
	oConn.close
	Set oConn=Nothing
	oConni.close
	Set oConni=Nothing
End if
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.ConnectionTimeout = 100
	oConn.Provider = "MSDASQL"
	oConn.Open DATABASE
        l_cSQL = "UPDATE AvailableVehicles SET logouttime=getdate(), availablestatus = 'x' WHERE DriverID = '" & UserID & "' and AvailableStatus='c'"
	    'Response.write "65 l_cSQL="&l_cSQL&"<BR>"
    oConn.Execute(l_cSQL)
	oConn.close
	Set oConn=Nothing
    Response.Cookies("FleetXPhone")("VehicleID")=""
    VehicleName=""
''''End if

'response.write "Database="&Database&"<BR>"
'Response.Write "vehicleID="&VehicleID&"<BR>"
If VehicleID>"" and Fakesubmit>"" then
    'Response.write "Got here line 37!<BR>"
    'REsponse.write "SELECT lcintranet.dbo.Intranet_Users.FirstName AS Expr1, lcintranet.dbo.Intranet_Users.LastName AS Expr2 FROM AvailableVehicles INNER JOIN lcintranet.dbo.Intranet_Users ON AvailableVehicles.DriverID = lcintranet.dbo.Intranet_Users.UserID  WHERE (VehicleID='"&VehicleID&"') and (AvailableStatus='c')"
	Set Recordset123 = Server.CreateObject("ADODB.Recordset")
	Recordset123.ActiveConnection = Database
	if UserID <> "1" and UserID <>"508" then
    Recordset123.Source = "SELECT lcintranet.dbo.Intranet_Users.FirstName AS Expr1, lcintranet.dbo.Intranet_Users.LastName AS Expr2 FROM AvailableVehicles INNER JOIN lcintranet.dbo.Intranet_Users ON AvailableVehicles.DriverID = lcintranet.dbo.Intranet_Users.UserID  WHERE (VehicleID='"&trim(VehicleID)&"') and (AvailableStatus='c') and AvailableVehicles.DriverID <> '1' and AvailableVehicles.DriverID <>'508'"
  else
    Recordset123.Source = "SELECT lcintranet.dbo.Intranet_Users.FirstName AS Expr1, lcintranet.dbo.Intranet_Users.LastName AS Expr2 FROM AvailableVehicles INNER JOIN lcintranet.dbo.Intranet_Users ON AvailableVehicles.DriverID = lcintranet.dbo.Intranet_Users.UserID  WHERE (VehicleID='"&trim(VehicleID)&"') and (AvailableStatus='c') "
	end if
  Recordset123.CursorType = 0
	Recordset123.CursorLocation = 2
	Recordset123.LockType = 1
	Recordset123.Open()
	Recordset123_numRows = 0
    'response.write "*****Database="&Database&"<BR>"
	'response.write "*****Recordset123.Source="&Recordset123.Source&"<BR>"
		If NOT Recordset123.EOF AND UserID <> "1" AND UserID <> "508" then 
            'Response.write "XXXXGot here line 49!XXXX<BR>"

            Expr1=Recordset123("Expr1")
            Expr2=Recordset123("Expr2")
            ErrorMessage="Sorry, that vehicle is currently checked out by "&Expr1&" "&Expr2&".<br><br>To log "&expr1&" out and log yourself into vehicle #"&vehicleID&"... <form method='post' action='DriverVehicle.asp'><Input type='hidden' name='takenvehicle' value='"& VehicleID &"'><Input type='submit' id='gobutton' value='click here'></form>"

            Else
            'Response.write "Oh noooooooooo!  I got here!<BR>"

	Set Recordset1 = Server.CreateObject("ADODB.Recordset")
	Recordset1.ActiveConnection = Database
	''''Recordset1.Source = "SELECT top 1 un_desc, un_id, un_model, availablestatus FROM fcunits LEFT OUTER JOIN AvailableVehicles ON fcunits.un_id = AvailableVehicles.VehicleID  WHERE (UN_ID='"&VehicleID&"') ORDER BY AvailableVehicles.Available_ID DESC"
    Recordset1.Source = "SELECT top 1 un_desc, un_id, un_model FROM fcunits WHERE (UN_ID='"&VehicleID&"') and (UnitStatus='c')"
	Recordset1.CursorType = 0
	Recordset1.CursorLocation = 2
	Recordset1.LockType = 1
	Recordset1.Open()
	Recordset1_numRows = 0
	'response.write "*****Recordset1.Source="&Recordset1.Source&"<BR>"
		If NOT Recordset1.EOF then 
			'response.write "GOT HERE!<BR>"
			VehicleName=Recordset1("UN_DESC")
			PenchantVehicleID=Recordset1("un_id")
            VehicleID=Recordset1("un_id")
            VehicleType=Recordset1("un_model")
            ''''AvailableStatus=Recordset1("AvailableStatus")
			''''''''''''''''''''''''''''
			'response.write "penchantVehicleID="&penchantVehicleID&"<BR>"
			'response.write "VehicleID="&VehicleID&"<BR>"
            'response.write "PenchantVehicleID="&PenchantVehicleID&"<BR>"
			'response.write "VehicleName="&VehicleName&"<BR>"
			'''''''''''''''''''''''''''''
			Response.Cookies("FleetXPhone")("UnitID")=VehicleID
			Response.Cookies("FleetXPhone")("VehicleID")=PenchantVehicleID
			Response.Cookies("FleetXPhone")("VehicleName")=VehicleName
            Response.Cookies("FleetXPhone")("VehicleType")=VehicleType
			'Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
			'	RSEVENTS2.Open "DriverLog", Intranet, 2, 2
			'	RSEVENTS2.addnew	
			'	RSEVENTS2("DriverID")=UserID		
			'	RSEVENTS2("VehicleID") = VehicleID
				'RSEVENTS2("LogInOut") = "o"
			'	RSEVENTS2("LogTime")=Now()		
			'	RSEVENTS2("LogStatus") = "c"
			'	RSEVENTS2.update
			'	RSEVENTS2.close			
			'set RSEVENTS2 = nothing	
''''''''''New Login/Logout Feature
            'Response.write "SAVED INTO THE NEW VEHICLE!!!!"
			Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				RSEVENTS2.Open "AvailableVehicles", Database, 2, 2
				RSEVENTS2.addnew	
				RSEVENTS2("VehicleID") = VehicleID
				RSEVENTS2("VehicleName") = VehicleName
                RSEVENTS2("VehicleType") = VehicleType
                RSEVENTS2("DriverID")=UserID	
				RSEVENTS2("LogInTime")=Now()		
				RSEVENTS2("AvailableStatus") = "c"
				RSEVENTS2.update
				RSEVENTS2.close			
			set RSEVENTS2 = nothing	
			'''Recordset1.Close()
			'''Set Recordset1 = Nothing	
            'REsponse.write "DriverVehicle Line 105<br>"	
            
            	
			' UNCOMMENT AFTER TESTING 
      Response.Redirect("Default.asp")
      
      
			ELSE
			ErrorMessage="Not a valid vehicle code"
			''''Set Recordset16 = Server.CreateObject("ADODB.Recordset")
			''''Recordset16.ActiveConnection = Database
			''''Recordset16.Source = "SELECT * FROM FCUNITS WHERE (UN_ID='"&VehicleID&"')"
			''''Recordset16.CursorType = 0
			''''Recordset16.CursorLocation = 2
			''''Recordset16.LockType = 1
			''''Recordset16.Open()
			''''Recordset16_numRows = 0
			'response.write "*****Recordset1.Source="&Recordset1.Source&"<BR>"
				''''If NOT Recordset16.EOF then 
				
				''''End if
			''''Recordset16.Close()
			''''Set Recordset16 = Nothing										
		End if
	Recordset1.Close()
	Set Recordset1 = Nothing
    End if
	Recordset123.Close()
	Set Recordset123 = Nothing


End if
'Response.write "UserID="&UserID&"<BR>"
'Response.write "LIne 110 UserID="&UserID&"<BR>"


%>
<script type="text/javascript">
    function setFocusToTextBox() {
        document.getElementById("VehicleID").focus();
    }
</script>
</head>
<body onload="document.Form1.VehicleID.focus()">
<!-- #include file="LogoSection.asp" -->
<table cellspacing="0" cellpadding="0" width="300" border="0" bordercolor="black" ID="Table1">
	<!--
	<tr><td align="center" colspan="9"><a href="default.asp" class="mainpagelink">Return Home</a></td></tr>
	-->
	<tr>
		<td class="FleetXRedSection" colspan="2" align="center">
			Driver Vehicle Page
		</td>
	</tr>
	<tr><td>&nbsp;</td></tr>
	<tr><td align="left" colspan="2">

	<%
    VehicleSet=""
    If VehicleSet="" or isnull(VehicleSet) then
	''response.write "GOT HERE!"
	%>
		<form method="post" id="Form1" name="Form1" action="DriverVehicle.asp">
			<table border="0" cellpadding="2" cellspacing="0" ID="Table4" width="100%" bordercolor="blue">
				<tr> 
					<td class="mainpagetextboldright" colspan="2"><img src="images/pixel.gif" height="2"></td>
				</tr>
				<tr>
					<td class='mainpagetextcenter' colspan="2" nowrap="nowrap" align="center">SCAN in vehicle code</td>
				</tr>
                <tr><td>&nbsp;</td></tr>
				<tr>
					<td colspan='2' class='generalcontent' align="center">
						<input type="password" maxlength='25' size='25' name='VehicleID' id='VehicleID' />
						<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden16" />
					</td>
				</tr>
                <tr><td>&nbsp;</td></tr>
                </form>
	            <tr><td colspan="2" align="center"><input size="8" maxlength="30" type="text" value="<%=Time()%>" name="bogus" id="bogus" onfocus="formSubmit()" readonly="readonly" class="InvisibleTextFieldWhite" /></td></tr>
				<%if errormessage>"" then%>
					<tr>
						<td class='errormessage'colspan='2' align="center"><br><%=ErrorMessage%><br><br></td>
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

		
			<%else
			

														
		End if%>
	</td></tr>

	<!--
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="Submit" ID="Submit1"></td></tr>
	-->
</table>



</body>
</html>

