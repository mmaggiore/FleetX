<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<link rel="stylesheet" type="text/css" href="mainStyleSheet.css">
<!-- #include file="../DedicatedFleets/include/ifabsettings.inc" -->
<title>LogistiCorp Driver Log In Page</title>
<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<%
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

' -----------------DATABASE INFORMATION REMOVED---------------

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
	<tr><td class="mainpagetextboldright">Password:&nbsp;&nbsp;</td><td><input size="15" maxlength="10" type="Password" name="Password" ID="Text1" onBlur="form.submit()"></td></tr>
	<tr><td>&nbsp;</td></tr>
	<input type="hidden" name="FakeSubmit" value="Submit">
	<!--
	<tr><td colspan="2" align="center"><table cellpadding="0" cellspacing="0" border="0"><tr><td align="center"><input type="submit" name="submit" value="Submit"></td></form><form method="post" ID="Form2" name="Form2"><td align="center"><input type="submit" name="reset" value="Reset"></td></form></tr></table></td></tr>
	-->
	<tr><td colspan="2" align="center" class="ErrorMessageBold" nowrap><%=ErrorMessage%></td></tr>
	
</table>



</body>
</html>
