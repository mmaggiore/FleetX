<html>
<head>
<title>FleetX - Retrieve Password</title>
<!-- #include file="settings.inc" -->
<link rel="stylesheet" type="text/css" href="mainStyleSheet.css">
<%
Message="Simply submit your e-mail address and your login name and password will be emailed to you"
Email=request.form("RequiredEmail")
PageStatus=request.form("PageStatus")


'Response.Write "Intranet="&Intranet&"<BR>"
If lcase(PageStatus)="find" then
	Set Recordset1 = Server.CreateObject("ADODB.Recordset")
	
	Recordset1.ActiveConnection = Intranet
	Recordset1.Source = "SELECT * FROM Intranet_Users WHERE (email='"&email&"') and (Status='c')"
	Recordset1.CursorType = 0
	Recordset1.CursorLocation = 2	
	Recordset1.LockType = 1
	Recordset1.Open()
	Recordset1_numRows = 0
	if NOT Recordset1.EOF then
		LastName=Recordset1("LastName")
		FirstName=Recordset1("FirstName")
		Email=Recordset1("Email")	
		Password=Recordset1("Password")
		UserName=Recordset1("Username")
		PageStatus="mail"
		else
		ErrorMessage="That email address is not in our system."
	End if
	Recordset1.Close()
	Set Recordset1 = Nothing	
	If lcase(PageStatus)="mail" then
		Body = "ATTN:&nbsp;&nbsp;"&FirstName&" "&LastName &"<br><br>"   & _
		"Below are your user name and password for the "& TaskMasterCompanyName &" Intranet Website.<br><br>"   & _
		"user name: "&UserName&"<br>"  & _
		"password: "&Password&"<br><br>"   & _
		"The address is: "& TaskMasterURL &" <br><br>"   & _ 			
		"If you'd like to personalize your password, you may do so on the home page, after you have logged in for the first time.<br><br>"   & _
		"If you have any questions, please do not hesitate to contact me.<br><br>"   & _
		"Thank you,<br><br>"   & _
		"Mark Maggiore<br>"  & _
		"LogistiCorp Web Developer<br>"  & _
		"mark.maggiore@LogistiCorp.us<br>"  & _ 
		"214/956-0400 xt 212<br><br>"
		Recipient=FirstName&" "&LastName


		Set objMail = CreateObject("CDONTS.Newmail")
		objMail.From = "system.monitor@logisticorp.us"
		objMail.To = Email
		objMail.Subject = "Username/Password Information"
		objMail.MailFormat = cdoMailFormatMIME
		objMail.BodyFormat = cdoBodyFormatHTML
		objMail.Body = Body
		objMail.Send
		Set objMail = Nothing	


		'if not Mailer.SendMail then
		  	'ErrorMessage = "Please try again later as the Email server is experiencing difficulties"
			'else
  			ErrorMessage = "An email has been sent to "&email&".<br>You should be recieving your information shortly."
		'end if	
	end if
End if
%>
</head>
<body bgcolor="#FFFFFF" text="#000000" onload="document.FindUser.requiredemail.focus();" topmargin="0" leftmargin="0">
	<table width="<%=TaskmasterLogoWidth%>" border="0" cellpadding="0" cellspacing="0" ID="Table1">
	<tr>
		<td><a href="default2.asp"><img src="images/Main_Banner_Top_Intranet.jpg" width="<%=TaskMasterLogoWidth%>" height="<%=TaskMasterLogoHeight%>" border="0" alt="<%=TaskmasterCompanyName%> Main Logo"></a></td>
	</tr>
	<tr>
		<td>
			<table border="1" cellpadding="0" cellspacing="0" width="<%=TaskMasterLogoWidth%>" ID="Table2">
				<tr>
					<td bgcolor="#E6E5E5" border="1"><img src="images/pixel.gif" height="10" width="1" border="0"></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr><td>
<form action="RequestUserInfo.asp" method="post" name="FindUser">
<table Width="750" Cellspacing="0" Cellpadding="0" align="left">
<tr><td>&nbsp;</td></tr>
<tr><td>
<table width="432" border="0" align="center" class="MainPageText" background="images/login.jpg">
	<tr height="40">
		<td width="150">&nbsp;</td>
	</tr>
  <tr Height="30"> 
    <td colspan="5" valign="middle" align="left" class="MainPageText"> 
      	Submit your email address below and your username and password will be 
		emailed to you.	
	
    </td>

  </tr>
<tr><td>&nbsp;</td></tr>

  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
      Email Address:
    </td>
	<td width="5">&nbsp;</td>
    <td width="136"> 
     <INPUT TYPE="text" NAME="requiredemail" value="<%=email%>" size="30">
    </td>
	<input type="hidden" name="pagestatus" value="find">
	<td width="5">&nbsp;</td>
    <td width="211"> 
      <INPUT TYPE="submit" name="ButtonValue" VALUE="Submit">
    </td>
  </tr>
	<tr Height="50">
		<td>&nbsp;</td>
	</tr>
</table>
</td></tr>
<%if ErrorMessage>"" then%>
<tr><td>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" class="MainPageText">
	<tr><td>&nbsp;</td></tr>  
	<tr>
    <td align="center" class="Errormessage"><%=ErrorMessage%></td>
  </tr>
	<tr><td>&nbsp;</td></tr>
</table>
</td></tr>
<%end if%>
<tr><td align="center">
<table border="0" cellspacing="0" cellpadding="0" align="center">
  <TR>  
	<td align="left" valign="center"><span class="MainPageText"><br>
		To request a login ID and password  Email: <a href="mailto:<%=TaskMasterWebMasterEmail%>" class="MainPageLink"><%=TaskMasterWebMasterEmail%></a>
	</td>
  </tr> 
  <tr><td>&nbsp;</td></tr>
	<tr><td align="center" class="MainPageText">
		<%
		'If pagestatus="mail" then
		%>
		<a href="login.asp" class="MainPageLink">Click here</a> to return to the login page.
		<%
		'end if
		%>
&nbsp;
</td></tr>
</table>
</td></tr>
</table>
</form>
</td></tr></table>

</body>
</html>
