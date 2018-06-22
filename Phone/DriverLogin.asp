<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<%
''''This makes page avoid the driver log in check...
LoginCheck="n"
 %>

<title>LogistiCorp Driver Log In Page</title>
<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta name="viewport" content="width=300">
<!-- #include file="fleetX.inc" -->
<script type="text/javascript">
function formSubmit()
{
document.getElementById("Form1").submit()
}
</script>
<%
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"

		SomeVariable=Request.ServerVariables ("HTTP_USER_AGENT")
        'Response.Write "xxxSomeVariable="& SomeVariable &"<BR>"
        If left(SomeVariable,15)="Motorola_ES405B" then
            'Response.Write "YES<BR>"
            %>
            <META NAME="MobileOptimized" CONTENT="0">
            <%
            Else
            'Response.Write "NO<BR>"
        End if
'showtop=request.QueryString("ShowTop")
'If showtop="y" then
'    session("showtop")="y"
'End if
session("showtop")=""
markx=request.querystring("x")
If trim(markx)="1" then
    Response.redirect("driverlogin.asp")
End if
''''''''''''''''''''''''''''''''''''''''''
SecureYes = Request.ServerVariables ("HTTPS")
'If SecureYes="off" then
If SecureYes="on" then
	''''''''''''''''''''''''''''''''''''''''''''
	something=Request.ServerVariables("HTTP_HOST") 
	Something=left(Something,4) 
	'''If lcase(something)="test" then 
	'''	Response.redirect("http://test.logisticorp.us/phone/DriverLogin.asp")
		'Response.Write "GOT HERE!!!" 
	'''	MPMSendEmail="n" 
	'''	else
	'''	Response.redirect("http://www.logisticorp.us/phone/DriverLogin.asp")
	'''End if 
	'Response.Write "Something="&Something&"<BR>"
	'Response.Write "MPMSendEmail="&MPMSendEmail&"<BR>"

	'''''''''''''''''''''''''''''''''''''''''''
		
End if
''''''''''''''''''''''''''''''''''''''''''
ComputerID=Request.Cookies("LegalComputer")("ComputerID")
PhoneUser=Request.ServerVariables("HTTP_UA_OS") 
If trim(ComputerID)="" and Trim(PhoneUser)<>"Windows CE (Pocket PC) - Version 5.1" and Trim(PhoneUser)<>"Windows CE (Pocket PC) - Version 5.2"  then

				Body = "The users OS is<BR><BR>"& PhoneUser &"<br><br>"
				'Recipient=FirstName&" "&LastName
				SentToEmail="mark.maggiore@logisticorp.us"
				'Email="KWETI.Mailbox@am.kwe.com"
				'Email="mark@maggiore.net"
				'''Set objMail = CreateObject("CDONTS.Newmail")
				'''objMail.From = "System.Notification@logisticorp.us"
				'''objMail.To = SentToEmail
				'objMail.cc = "4692269939@tmomail.net;htmlmale@yahoo.com"
				'objMail.cc = "4692269939@tmomail.net;pcurrin@ti.com"
				'''objMail.Subject = "User OS"
				'''objMail.MailFormat = cdoMailFormatMIME
				'''objMail.BodyFormat = cdoBodyFormatHTML
				'''objMail.Body = Body
				'''objMail.Send
				'''Set objMail = Nothing

End if







FakeSubmit=Request.Form("FakeSubmit")
'Response.Write "FakeSubmit="&FakeSubmit&"<BR>"
If FakeSubmit="12345" then
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
        Response.Cookies("FleetXPhone")("UserID") = UserID
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
		'Recordset1.Close()
		'Set Recordset1 = Nothing
End if

If FakeSubmit>"" and (UserName="" or Password="") then
	ErrorMessage="Both a driver ID and password are required"
End if

%>
</head>
<body onload="document.Form1.UserName.focus()">
<%If showtop="y" then %>
<!-- #include file="../../dedicatedfleets/nav/ifabnavbar.inc" -->
<%end if %>

<!-- #include file="LogoSection.asp" -->	


<form method="post" id="Form1" name="Form1" action="DriverLogin.asp">
<table cellspacing="0" cellpadding="0" width="300" border="0" bordercolor="black" >
	
	<tr>
		<td class="FleetXRedSection" colspan="2" align="center">
			Driver Log In Page
		</td>
	</tr>
	
	<tr><td>&nbsp;</td></tr>
	<tr><td class="MainPageText">Driver ID:&nbsp;&nbsp;</td><td><input size="15" maxlength="30" type="text" name="UserName" id="text1c" /></td></tr>
	<tr><td>&nbsp;</td></tr>
	<tr><td class="MainPageText">Password:&nbsp;&nbsp;</td><td><input size="15" maxlength="15" type="password" name="Password" id="text1" /></td></tr>
	<tr><td>&nbsp;</td></tr>
	<tr><td colspan="2" align="center"><input size="8" maxlength="30" type="text" value="<%=Time()%>" name="bogus" id="bogus" onfocus="formSubmit()" readonly="readonly" class="InvisibleTextFieldWhite" /></td></tr>
	
	<!--
	<tr><td class="mainpagetextboldright">Password:&nbsp;&nbsp;</td><td><input size="15" maxlength="10" type="Password" name="Password" ID="Password1" onBlur="formSubmit()"></td></tr>
	-->
	<tr><td>&nbsp;</td></tr>
	<input type="hidden" name="FakeSubmit" value="Submit">
	<!--
	<tr><td colspan="2" align="center"><table cellpadding="0" cellspacing="0" border="0"><tr><td align="center"><input type="submit" name="submit" value="Submit"></td></form><form method="post" ID="Form2" name="Form2"><td align="center"><input type="submit" name="reset" value="Reset"></td></form></tr></table></td></tr>
	-->
	
	<tr><td colspan="2" align="center" class="ErrorMessageBold" nowrap><%=ErrorMessage%></td></tr>
	<tr><td colspan="2">&nbsp;</td></tr>
	<tr><td>&nbsp;</td></tr>
	<tr><td>&nbsp;</td></tr>
	<tr><td>&nbsp;</td></tr>
	<tr><td>&nbsp;</td></tr>
	<tr><td>&nbsp;</td></tr>

</table>
<%
		Response.Cookies("FleetXPhone")("DriverEmail")=""
		Response.Cookies("FleetXPhone")("DriverUserID")=""
        Response.Cookies("FleetXPhone")("UserID")=""
		Response.Cookies("FleetXPhone")("DriverFirstName")=""
		Response.Cookies("FleetXPhone")("DriverLastName")=""
		Response.Cookies("FleetXPhone")("Rights")=""
		Response.Cookies("FleetXPhone")("VehicleID")=""
		Response.Cookies("FleetXPhone")("VehicleName")=""
		Response.Cookies("FleetXPhone")("UnitID")=""
		Response.Cookies("FleetXPhone")("VehicleID")=""


            'Response.write "VEhicleID="&VehicleID&"<BR>"
			'Response.write "VehicleName="&VehicleName&"<BR>"
			'Response.write "VehicleType="&VehicleType&"<BR>"
			'Response.write "UserID="&UserID&"<BR>"


%>

</form>
</body>
</html>
