<html>
<head>
<title>FleetX - New User</title>
<!-- #include file="fleetexpress.inc" -->
<link rel="stylesheet" type="text/css" href="css/Style.css">
<%
    ColorSelect=valid8(Request.form("ColorSelect"))
    ColorSelect=ColorSelect+1
    If ColorSelect>4 then ColorSelect=1 End if
    ColorSelect=3
    Select Case ColorSelect
        Case 1
            HeaderBorderColor="#cc1126"
            BorderColor="#cc1126"
            LinkClass="FleetExpressRed"
        Case 2
             HeaderBorderColor="#216194"
            BorderColor="#216194"
            LinkClass="FleetExpressBlue"
        Case 3 
            'HeaderBorderColor="#B7B8B8" 
            HeaderBorderColor="#d71e26"  
            BorderColor="#d71e26"
            LinkClass="FleetXRed"
        Case else 
            HeaderBorderColor="black"  
            BorderColor="black"
            LinkClass="FleetExpressBlack"
    End Select
    HighlightedField="RequestorName"
    CurrentDateTime=Now()
   '''''''''''HARDCODED STUFF
   sBT_ID=86
Message="Simply submit your e-mail address and your login name and password will be emailed to you"
Email=valid8(request.form("RequiredEmail"))
PageStatus=valid8(request.form("PageStatus"))
RequestorName=valid8(request.form("RequestorName"))
RequestorCompany=valid8(request.form("RequestorCompany"))
RequestorAddress=valid8(request.form("RequestorAddress"))
RequestorCity=valid8(request.form("RequestorCity"))
RequestorZipCode=valid8(request.form("RequestorZipCode"))
RequestorCostCenter=valid8(request.form("RequestorCostCenter"))
RequestorPhone=valid8(request.form("RequestorPhone"))
RequestorEmail=valid8(request.form("RequestorEmail"))
RetypeRequestorEmail=valid8(request.form("RetypeRequestorEmail"))
RequestorPassword=valid8(request.form("RequestorPassword"))
varCaptcha=valid8(Request.Form("varCaptcha"))
CaptchaSubmit=valid8(Request.Form("CaptchaSubmit"))

'Response.Write "Intranet="&Intranet&"<BR>"
If lcase(PageStatus)="find" then
    RequestorName=Replace(RequestorName, """", "`")
    RequestorName=Replace(RequestorName, "'", "`")
    RequestorCompany=Replace(RequestorCompany, """", "`")
    RequestorCompany=Replace(RequestorCompany, "'", "`")
    RequestorAddress=Replace(RequestorAddress, """", "`")
    RequestorAddress=Replace(RequestorAddress, "'", "`")
    RequestorCity=Replace(RequestorCity, """", "`")
    RequestorCity=Replace(RequestorCity, "'", "`")
    RequestorZipCode=Replace(RequestorZipCode, """", "`")
    RequestorZipCode=Replace(RequestorZipCode, "'", "`")
    RequestorCostCenter=Replace(RequestorCostCenter, """", "`")
    RequestorCostCenter=Replace(RequestorCostCenter, "'", "`")
    RequestorPhone=Replace(RequestorPhone, """", "")
    RequestorPhone=Replace(RequestorPhone, "'", "")
    RequestorEmail=Replace(RequestorEmail, """", "")
    RequestorEmail=Replace(RequestorEmail, "'", "")
    RequestorPassword=Replace(RequestorPassword, """", "")
    RequestorPassword=Replace(RequestorPassword, "'", "")
If CaptchaSubmit<>varCaptcha then
    ErrorMessage="You did not supply the correct verification code"
End if
If trim(RequestorPassword)="" then
    ErrorMessage="You must enter a Password"
End if
If trim(RetypeRequestorEmail)<>trim(RequestorEmail) then
    ErrorMessage="Your Email Address and Re-type Email Address do not match"
End if
If trim(RetypeRequestorEmail)="" then
    ErrorMessage="You must enter a Retype Email Address"
End if
If trim(RequestorEmail)="" then
    ErrorMessage="You must enter an Email Address"
End if
If trim(RequestorPHone)="" then
    ErrorMessage="You must enter a Phone Number"
End if
If trim(RequestorZipCode)="" then
    ErrorMessage="You must enter a Zip Code"
End if
If trim(RequestorCity)="" then
    ErrorMessage="You must enter a City"
End if
If trim(RequestorAddress)="" then
    ErrorMessage="You must enter an Address"
End if
If trim(RequestorCompany)="" then
    ErrorMessage="You must enter a Company"
End if
If trim(RequestorName)="" then
    ErrorMessage="You must enter a Name"
End if

If trim(ErrorMessage)="" then
	Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
		RSEVENTS2.Open "PreExistingRequestor", DATABASE, 2, 2
		RSEVENTS2.addnew
		RSEVENTS2("RequestorName")=RequestorName
        RSEVENTS2("RequestorCompany")=RequestorCompany
        RSEVENTS2("RequestorAddress")=RequestorAddress
        RSEVENTS2("RequestorCity")=RequestorCity
        RSEVENTS2("RequestorState")="TX"
        RSEVENTS2("RequestorZipCode")=RequestorZipCode
        RSEVENTS2("CostCenter")=RequestorCostCenter
        RSEVENTS2("RequestorPhone")=RequestorPhone
        RSEVENTS2("RequestorEmail")=RequestorEmail
        RSEVENTS2("RequestorPassword")=RequestorPassword
        'response.write "112 requestoremail=" & RequestorEmail & "<br>"
        If lcase(right(trim(RequestorEmail),6))="ti.com" then
            RSEVENTS2("RequestorStatus")="c"
            SendEmail="n"
            RSEVENTS2("bt_id")=92
        else
            RSEVENTS2("RequestorStatus")="n"
		    End if
        'response.write "119 SendEmail=" & SendEmail & "<br>"
        RSEVENTS2.update
		RSEVENTS2.close			
	set RSEVENTS2 = nothing 
 'Dim iMsg, iConf, Flds
    If SendEmail="n" then
 		Body = "ATTN:&nbsp;&nbsp;"&RequestorName &"<br><br>"   & _
		"Thank you for registering for the FleetX site.<br><br>"   & _
		"Your login information is:<br>" & _
        "Username: "&RequestorEmail&"<br>" & _	
        "Password: "&RequestorPassword&"<br><BR>" & _		
		"If you have any questions, please do not hesitate to contact me.<br><br>"   & _
		"Thank you,<br><br>"   & _
		"Mark Maggiore<br>"  & _
		"LogistiCorp Web Developer<br>"  & _
		"mark.maggiore@LogistiCorp.us<br>"  & _ 
		"817-591-2956<br><br>"
		Recipient=LastName

     'response.write "137 auto-approved<br>"
		'Set objMail = CreateObject("CDONTS.Newmail")
		'objMail.From = "FleetX@LogisticorpGroup.com"
		'objMail.To = RequestorEmail
		''objMail.To = "bettywalker@wiseweblady.com"
		'objMail.Subject = "Thank you for registering with FleetX"
		'objMail.MailFormat = cdoMailFormatMIME
		'objMail.BodyFormat = cdoBodyFormatHTML
		'objMail.Body = Body
		'objMail.Send
		'Set objMail = Nothing  

Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")
Set Flds = iConf.Fields
With Flds
	.Item("http://schemas.microsoft.com/cdo/configuration/sEndusing")				= AWS_SendUsingPort
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")				= AWS_SMTPServer
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl")				= AWS_SMTPUseSSL
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport")			= AWS_SMTPServerPort
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")		= AWS_SMTPAuthenticate
	.Item("http://schemas.microsoft.com/cdo/configuration/sendusername")			= AWS_SendUserName
	.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword")			= AWS_SendPassword
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout")	= AWS_SMTPConnectionTimeout
	.Update
End With
Set iMsg.Configuration = iConf

	iMsg.To = RequestorEmail
	iMsg.BCC = "Mark.Maggiore@Logisticorp.us"
	SentMail="y"
With iMsg
	Set .Configuration = iConf
	.From ="System.Notification@logisticorp.us"
	.Subject = "Thank you for registering with FleetX"
	.HTMLBody = Body
	.Send
End With   
        else	
	'If lcase(PageStatus)="mail" then
		Body = "ATTN:&nbsp;&nbsp;"&RequestorName &"<br><br>"   & _
		"Thank you for registering for the FleetX site.<br><br>"   & _
		"Your request will be reviewed, and you should hear back from us within one business day.<br><BR>" & _		
		"If you have any questions, please do not hesitate to contact me.<br><br>"   & _
		"Thank you,<br><br>"   & _
		"Mark Maggiore<br>"  & _
		"LogistiCorp Web Developer<br>"  & _
		"mark.maggiore@LogistiCorp.us<br>"  & _ 
		"817-591-2956<br><br>"
		Recipient=LastName

		'Set objMail = CreateObject("CDONTS.Newmail")
		'objMail.From = "FleetX@LogisticorpGroup.com"
        ''objMail.From = "mark.maggiore@Logisticorp.us"
		'objMail.To = RequestorEmail
        'objMail.cc = "linda.holt@logisticorp.us"
		'objMail.Subject = "Thank you for registering with FleetX"
		'objMail.MailFormat = cdoMailFormatMIME
		'objMail.BodyFormat = cdoBodyFormatHTML
		'objMail.Body = Body
		'objMail.Send
		'Set objMail = Nothing

Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")
Set Flds = iConf.Fields
With Flds
	.Item("http://schemas.microsoft.com/cdo/configuration/sEndusing")				= AWS_SendUsingPort
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")				= AWS_SMTPServer
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl")				= AWS_SMTPUseSSL
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport")			= AWS_SMTPServerPort
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")		= AWS_SMTPAuthenticate
	.Item("http://schemas.microsoft.com/cdo/configuration/sendusername")			= AWS_SendUserName
	.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword")			= AWS_SendPassword
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout")	= AWS_SMTPConnectionTimeout
	.Update
End With
Set iMsg.Configuration = iConf

	iMsg.To = RequestorEmail
	iMsg.BCC = "Mark.Maggiore@Logisticorp.us;linda.holt@logisticorp.us"
	SentMail="y"
With iMsg
	Set .Configuration = iConf
	.From = "System.Notification@logisticorp.us"
	.Subject = "Thank you for registering with FleetX"
	.HTMLBody = Body
	.Send
End With 
        
 		Body = "ATTN:&nbsp;&nbsp;"&RequestorName &" has registered to become a FleetX user.<br><br>"   & _
		"Below is their information:<br><br>"   & _
		"User: "&RequestorName&"<br>"  & _
        "Company: "&RequestorCompany&"<br>"  & _
        "Address: "&RequestorAddress&"<br>"  & _
        "City: "&RequestorCity&"<br>"  & _
        "State: TX<br>"  & _
        "Zip Code: "&RequestorZipCode&"<br>"  & _
        "Cost Center: "&RequestorCostCenter&"<br>"  & _
        "Phone Number: "&RequestorPhone&"<br>"  & _
		"Email Address: "&RequestorEmail&"<br><br>"   & _

        "To review/approve this user, click here: <a href='"&WhichSite&"'/home.asp'>FleetX Site</a><br><br>" &_

		"If you have any questions, please do not hesitate to contact me.<br><br>"   & _
		"Thank you,<br><br>"   & _
		"Mark Maggiore<br>"  & _
		"LogistiCorp Web Developer<br>"  & _
		"mark.maggiore@LogistiCorp.us<br>"  & _ 
		"214/956-0650 xt 212<br><br>"
		Recipient=LastName


		'Set objMail = CreateObject("CDONTS.Newmail")
		'objMail.From = "FleetX@LogisticorpGroup.com"
		'objMail.To = "Mark.Maggiore@LogistiCorp.us"
		'objMail.Subject = "New FleetX User"
		'objMail.MailFormat = cdoMailFormatMIME
		'objMail.BodyFormat = cdoBodyFormatHTML
		'objMail.Body = Body
		'objMail.Send
		'Set objMail = Nothing 
Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")
Set Flds = iConf.Fields
With Flds
	.Item("http://schemas.microsoft.com/cdo/configuration/sEndusing")				= AWS_SendUsingPort
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")				= AWS_SMTPServer
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl")				= AWS_SMTPUseSSL
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport")			= AWS_SMTPServerPort
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")		= AWS_SMTPAuthenticate
	.Item("http://schemas.microsoft.com/cdo/configuration/sendusername")			= AWS_SendUserName
	.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword")			= AWS_SendPassword
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout")	= AWS_SMTPConnectionTimeout
	.Update
End With
Set iMsg.Configuration = iConf

	iMsg.To = "Mark.Maggiore@Logisticorp.us"
	SentMail="y"
With iMsg
	Set .Configuration = iConf
	.From ="System.Notification@logisticorp.us"
	.Subject =  "New FleetX User"
	.HTMLBody = Body
	.Send
End With
        'REsponse.write "requestoremail="&requestoremail&"<BR>"
        'REsponse.write "Body="&Body&"<BR>"      	
    End if

		'if not Mailer.SendMail then
		  	'ErrorMessage = "Please try again later as the Email server is experiencing difficulties"
			'else
  			ThankYouMessage = UCASE("Thank you for your interest in FleetX.<br><br>A verification email has been sent to "&requestoremail&".<br><br>You should be recieving your information shortly.<br><br>To return to the login page, <a href='login.asp' class='fleetxredmain'>click here</a>.")
            Dontshow="y"
		'end if	
	end if
End if
%>
</head>

<body onload="document.FindUser.requestorName.focus();">
	<table align="center" border="0" bordercolor="black" cellpadding="0" cellspacing="0" ID="Table1" width="100%" height="100%">
        <tr><td align="left"><img src="images/pixel.gif" height="10" width="1" /></td></tr>
        <tr>
            <td align="left"><a href="home.asp"><img src="images/FleetX_Small.jpg" height="50" width="168" /></a></td>
            <td align="right" valign="bottom">&nbsp;</td>
        </tr>
        <tr><td align="left" colspan="2"><img src="images/pixel.gif" /></td></tr>	
    <tr><td colspan="2">
<form action="NewUser.asp" method="post" name="FindUser">
<table border="0" bordercolor="green" Cellspacing="0" Cellpadding="0" align="left" width="100%">
 <tr><td align="left" colspan="2" bgcolor="<%=HeaderBorderColor%>"><img src="images/pixel.gif" height="5" width="1" /></td></tr>
        <%If trim(UserId)>"" then%>
        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2">Expedited Transportation Request</td></tr>
        <%else %>
        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetXRedSection" colspan="2">NEW USER APPLICATION</td></tr>
        <%end if %>
        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2"><img src="images/pixel.gif" height="5" width="1" /></td></tr>
        <!--tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressBodyWhite" colspan="2">In order to reach the Fleet Express Order page, please correctly type in the green verification code in the supplied text box and click "Submit."</td></tr-->
        <tr><td align="left" colspan="2"><img src="images/pixel.gif" height="5" width="1" /></td></tr>
<tr><td>
<table  border="0" bordercolor="blue" align="center" class="MainPageText">
	<tr height="40">
		<td width="650">&nbsp;</td>
	</tr>
    <%if trim(dontshow)<>"y" then %>
  <tr Height="30"> 
    <td colspan="5" valign="middle" align="center" class="MainPageText"> 
      	COMPLETE THIS FORM AND SUBMIT TO BECOME A NEW USER
	
    </td>

  </tr>
<tr><td>&nbsp;</td></tr>
<%if trim(ErrorMessage)>"" then%>
<tr><td colspan="5" align="center">
<table width="100%" border="0" bordercolor="Yellow" cellspacing="0" cellpadding="0" align="center" class="MainPageText">
	<tr><td>&nbsp;</td></tr>  
	<tr>
    <td align="center" class="Errormessage"><%=ErrorMessage%></td>
  </tr>
	<tr><td>&nbsp;</td></tr>
</table>
</td></tr>
<%end if%>

  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
      NAME:
    </td>
	<td width="5">&nbsp;</td>
    <td width="136"> 
     <INPUT TYPE="text" NAME="requestorName" value="<%=requestorName%>" maxlength="50" size="30">
    </td>
	<td width="5">&nbsp;</td>
    <td width="650"> 
      &nbsp;
    </td>
  </tr>
  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
      COMPANY:
    </td>
	<td width="5">&nbsp;</td>
    <td width="136"> 
     <INPUT TYPE="text" NAME="requestorCompany" value="<%=requestorCompany%>" maxlength="50" size="30">
    </td>
	<td width="5">&nbsp;</td>
    <td width="650"> 
      &nbsp;
    </td>
  </tr>
  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
      ADDRESS:
    </td>
	<td width="5">&nbsp;</td>
    <td width="136"> 
     <INPUT TYPE="text" NAME="requestorAddress" value="<%=requestorAddress%>" maxlength="50" size="30">
    </td>
	<td width="5">&nbsp;</td>
    <td width="650"> 
      &nbsp;
    </td>
  </tr>
  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
        CITY:
    </td>
	<td width="5">&nbsp;</td>
    <td width="136"> 
     <INPUT TYPE="text" NAME="requestorCity" value="<%=requestorCity%>" maxlength="50" size="30">
    </td>
	<td width="5">&nbsp;</td>
    <td width="650"> 
      &nbsp;
    </td>
  </tr>
  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
        STATE:
    </td>
	<td width="5">&nbsp;</td>
    <td width="136"> 
     <INPUT TYPE="hidden" NAME="requestorState" value="TX">TX
    </td>
	<td width="5">&nbsp;</td>
    <td width="650"> 
      &nbsp;
    </td>
  </tr>
  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
        ZIP CODE:
    </td>
	<td width="5">&nbsp;</td>
    <td width="136"> 
     <INPUT TYPE="text" NAME="requestorZipCode" value="<%=requestorZipCode%>" maxlength="12" size="30">
    </td>
	<td width="5">&nbsp;</td>
    <td width="650"> 
      &nbsp;
    </td>
  </tr>
  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
        COST CENTER:
    </td>
	<td width="5">&nbsp;</td>
    <td width="136"> 
     <INPUT TYPE="text" NAME="RequestorCostCenter" value="<%=requestorCostCenter%>" maxlength="12" size="30">
    </td>
	<td width="5">&nbsp;</td>
    <td width="650"> 
      &nbsp;
    </td>
  </tr>
  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
        PHONE NUMBER:
    </td>
	<td width="5">&nbsp;</td>
    <td width="136"> 
     <INPUT TYPE="text" NAME="requestorPhone" value="<%=requestorPhone%>" maxlength="20" size="30">
    </td>
	<td width="5">&nbsp;</td>
    <td width="650"> 
      &nbsp;
    </td>
  </tr>
  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
        EMAIL ADDRESS:
    </td>
	<td width="5">&nbsp;</td>
    <td width="136"> 
     <INPUT TYPE="text" NAME="requestorEmail" value="<%=requestorEmail%>" maxlength="100" size="30">
    </td>
	<td width="5">&nbsp;</td>
    <td width="650"> 
      &nbsp;
    </td>
  </tr>
  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
        RE-TYPE EMAIL ADDRESS:
    </td>
	<td width="5">&nbsp;</td>
    <td width="136"> 
     <INPUT TYPE="text" NAME="retyperequestorEmail" value="<%=retyperequestorEmail%>" maxlength="100" size="30">
    </td>
	<td width="5">&nbsp;</td>
    <td width="650"> 
      &nbsp;
    </td>
  </tr>


  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
      PASSWORD:
    </td>
	<td width="5">&nbsp;</td>
    <td width="136"> 
     <INPUT TYPE="text" NAME="requestorpassword" value="<%=requestorpassword%>" maxlength="15" size="30">
    </td>
	<input type="hidden" name="pagestatus" value="find">
	<td width="5">&nbsp;</td>
    <td width="650"> 
      &nbsp;
    </td>
  </tr>


 <%
 
    'varCAPTCHA=date()&time()
Randomize()
varCAPTCHA1=Int(100 * Rnd())   
Randomize()
varCAPTCHA2=Int(100 * Rnd()) 
Randomize()
varCAPTCHA3=Int(100 * Rnd()) 
Randomize()
varCAPTCHA4=Int(100 * Rnd())
Randomize()
varCAPTCHA5=Int(100 * Rnd())   
 
    'varCAPTCHA=now()-666
    

    
    CAPTCHA1=mid(varCAPTCHA1,1,1)
    CAPTCHA2=mid(varCAPTCHA2,1,1)
    CAPTCHA3=mid(varCAPTCHA3,1,1)
    CAPTCHA4=mid(varCAPTCHA4,1,1)
    CAPTCHA5=mid(varCAPTCHA5,1,1)
    
    varCAPTCHA=CAPTCHA1&CAPTCHA2&CAPTCHA3&CAPTCHA4&CAPTCHA5
    
   'response.Write "xxxvarCAPTCHA="&varCAPTCHA&"<BR>" 
    
  %>
  <tr><td>&nbsp;</td></tr>
  <tr>
    <td colspan="5" align="center">
        <img src="images/captcha/<%=CAPTCHA1%>.gif" height="37" width="37" border="0" />
        <img src="images/captcha/<%=CAPTCHA2%>.gif" height="37" width="37" border="0" />
        <img src="images/captcha/<%=CAPTCHA3%>.gif" height="37" width="37" border="0" />
        <img src="images/captcha/<%=CAPTCHA4%>.gif" height="37" width="37" border="0" />
        <img src="images/captcha/<%=CAPTCHA5%>.gif" height="37" width="37" border="0" />
    </td>
  </tr>

         <input type="hidden" name="varCAPTCHA"
            value="<% = varCAPTCHA %>" />

      <tr>
	       <td NOWRAP valign="middle" align="right" class="MainPageText"> 
          VERIFICATION CODE:
            </td>
	        <td width="5">&nbsp;</td>
            <td width="136">
          <input name="CAPTCHASubmit">
            </td>
      </tr>
      <tr><td>&nbsp;</td></tr>  



  <tr><td>&nbsp;</td></tr>
	<tr Height="50">
		<td align="center" colspan="5"><INPUT TYPE="submit" id="gobutton" name="ButtonValue" VALUE="Submit"></td>
	</tr>
    <%end if %>
  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
      &nbsp;
    </td>
	<td width="5">&nbsp;</td>
    <td width="136"> 
     &nbsp;
    </td>
	<td width="5">&nbsp;</td>
    <td width="725"> 
      &nbsp;
    </td>
  </tr>
</table>
</td></tr>
<%if trim(ErrorMessage)>"" then%>
<tr><td>
<table width="100%" border="0" bordercolor="Yellow" cellspacing="0" cellpadding="0" align="center" class="MainPageText">
	<tr><td>&nbsp;</td></tr>  
	<tr>
    <td align="center" class="Errormessage"><%=ErrorMessage%></td>
  </tr>
	<tr><td>&nbsp;</td></tr>
</table>
</td></tr>
<%end if%>
<%if trim(ThankYouMessage)>"" then%>
<tr><td>
<table width="100%" border="0" bordercolor="Yellow" cellspacing="0" cellpadding="0" align="center" class="MainPageText">
	<tr><td>&nbsp;</td></tr>  
	<tr>
    <td align="center"><%=ThankYouMessage%></td>
  </tr>
	<tr><td>&nbsp;</td></tr>
</table>
</td></tr>
<%end if%>
</table>
</form>
<tr><td Height="90%">&nbsp;</td></tr>
<tr>
    <td height="100" class="FleetXGreySection" colspan="2">
        <!-- #include file="BottomSection.asp" -->
    </td>
</tr>
<tr><td height="15" class="FleetXRedSectionSmall" colspan="2" align="center"><%=CopywriteNotice %></td></tr>
</td></tr></table>


</body>
</html>
