<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<link rel="stylesheet" type="text/css" href="mainStyleSheet.css">
<!-- #include file="../v9web/include/ifabsettings.inc" -->
<title>LogistiCorp Driver Log In Page</title>
<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<script type="text/javascript">
function formSubmit()
{
document.getElementById("Form1").submit()
}
</script>
<%


''''''''''''''''''''''''''''''''''''''''''
SecureYes = Request.ServerVariables ("HTTPS")
'If SecureYes="off" then
If SecureYes="on" then
	''''''''''''''''''''''''''''''''''''''''''''
	something=Request.ServerVariables("HTTP_HOST") 
	Something=left(Something,4) 
	If lcase(something)="test" then 
		Response.redirect("https://test.logisticorp.us/phone/DriverLogin.asp")
		'Response.Write "GOT HERE!!!" 
		MPMSendEmail="n" 
		'else
		'Response.redirect("http://www.logisticorp.us/phone/DriverLogin.asp")
	End if 
	'Response.Write "Something="&Something&"<BR>"
	'Response.Write "MPMSendEmail="&MPMSendEmail&"<BR>"

	'''''''''''''''''''''''''''''''''''''''''''
		
End if
''''''''''''''''''''''''''''''''''''''''''
FakeSubmit=Request.Form("FakeSubmit")
'Response.Write "FakeSubmit="&FakeSubmit&"<BR>"
If FakeSubmit="" then
	Session.Abandon
End if
UserName=Request.Form("UserName")
Username=Replace(Username,"'","")
Username=Replace(Username,"""","")

Password=Request.Form("Password")
Password=Replace(Password,"'","")
Password=Replace(Password,"""","")
If UserName>"" and Password>"" then
Set Recordset1 = Server.CreateObject("ADODB.Recordset")
'Response.Write "Intranet="&Intranet&"***<BR>"
SQL777="SELECT * FROM INTRANET_USERS WHERE (USERNAME='"&UserName&"') AND (PASSWORD='"&Password&"') AND (Status='c') AND ((Rights='u') OR (Rights='a') OR (Rights='g'))"
'Response.Write "SQL777="&SQL777&"***<BR>"
Recordset1.ActiveConnection = Intranet
Recordset1.Source = SQL777
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()
Recordset1_numRows = 0
	if NOT Recordset1.EOF then
		UserID=Recordset1("UserID")
		Rights=Recordset1("Rights")
		FirstName=Recordset1("FirstName")
		LastName=Recordset1("LastName")
		DriverEmail=Recordset1("Email")
		VehicleSet=Recordset1("DriverVehicle")
		'Response.write "LOGGED IN!"
		LoggedIn="yes"
		Response.Cookies("FleetXPhone")("VehicleSet")=VehicleSet
		Response.Cookies("FleetXPhone")("DriverEmail")= DriverEmail
		Response.Cookies("FleetXPhone")("DriverUserID") = UserID
		Response.Cookies("FleetXPhone")("DriverFirstName") = FirstName
		Response.Cookies("FleetXPhone")("DriverLastName") = LastName
		Recordset1.Close()
		Set Recordset1 = Nothing		
		Response.Redirect("DriverMessage.asp")
		
		'Response.write "UserID="&UserID&"<BR>"
		'Response.write "Rights="&Rights&"<BR>"
		'Response.write "FirstName="&FirstName&"<BR>"
		'Response.write "LastName="&LastName&"<BR>"
		'Response.write "DriverEmail="&DriverEmail&"<BR>"
		Else
		ErrorMessage="Incorrect driver ID or password"
		End if
		Recordset1.Close()
		Set Recordset1 = Nothing
End if

If FakeSubmit>"" and (UserName="" or Password="") then
	ErrorMessage="Both a driver ID and password are required"
End if

%>
</head>
<body OnLoad=document.Form1.UserName.focus()>
<table cellspacing="0" cellpadding="0" width="300" border="0" bordercolor="black" >
	<form method="post" ID="Form1" name="Form1">
	<tr>
		<td class="mainpagetextboldcenter" colspan="2" align="center">
			Driver Log In Page
		</td>
	</tr>
	
	<tr><td>&nbsp;</td></tr>
	<tr><td class="mainpagetextboldright">Driver ID:&nbsp;&nbsp;</td><td><input size="15" maxlength="30" type="Password" name="UserName"></td></tr>
	<tr><td>&nbsp;</td></tr>
	<tr><td class="mainpagetextboldright">Password:&nbsp;&nbsp;</td><td><input size="15" maxlength="10" type="Password" name="Password" ID="Text1"></td></tr>
	<tr><td>&nbsp;</td></tr>
	<tr><td colspan="2" align="center"><input size="8" maxlength="30" type="text" value="<%=Time()%>" name="bogus" ID="bogus" onFocus="formSubmit()" readonly="readonly" class="InvisibleTextFieldWhite"></td></tr>
	
	<!--
	<tr><td class="mainpagetextboldright">Password:&nbsp;&nbsp;</td><td><input size="15" maxlength="10" type="Password" name="Password" ID="Password1" onBlur="formSubmit()"></td></tr>
	-->
	<tr><td>&nbsp;</td></tr>
	<input type="hidden" name="FakeSubmit" value="Submit">
	<!--
	<tr><td colspan="2" align="center"><table cellpadding="0" cellspacing="0" border="0"><tr><td align="center"><input type="submit" name="submit" value="Submit"></td></form><form method="post" ID="Form2" name="Form2"><td align="center"><input type="submit" name="reset" value="Reset"></td></form></tr></table></td></tr>
	-->
	
	<tr><td colspan="2" align="center" class="ErrorMessageBold" nowrap><%=ErrorMessage%></td></tr>
	<tr><td>&nbsp;</td></tr>
	<tr><td>&nbsp;</td></tr>
	<tr><td>&nbsp;</td></tr>
	<tr><td>&nbsp;</td></tr>
	<tr><td>&nbsp;</td></tr>
	<tr><td>&nbsp;</td></tr>

</table>



</body>
</html>
