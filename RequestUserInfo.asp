<html>
<head>

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
    PageTitle="RETRIEVE PASSWORD"

%>
<title>FleetX - <%=PageTitle %></title>
<%
Message="Simply submit your e-mail address and your login name and password will be emailed to you"
Email=valid8(request.form("RequiredEmail"))
PageStatus=valid8(request.form("PageStatus"))
varCaptcha=valid8(Request.Form("varCaptcha"))
CaptchaSubmit=valid8(Request.Form("CaptchaSubmit"))

'Response.Write "Intranet="&Intranet&"<BR>"
    If CaptchaSubmit<>varCaptcha then
        ErrorMessage="You did not supply the correct verification code"
    End if



If lcase(PageStatus)="find" and trim(ErrorMessage)="" then
      Set oConn = Server.CreateObject("ADODB.Connection")
      oConn.ConnectionTimeout = 100
      oConn.Provider = "MSDASQL"
      oConn.Open DATABASE
	SQL = "SELECT * FROM PreExistingRequestor WHERE (RequestorEmail='"&email&"') and (RequestorStatus='c')"
      SET oRs = oConn.Execute(SQL)
	if NOT oRs.EOF then
		RequestorName=oRs("RequestorName")
		Email=oRs("RequestorEmail")	
		Password=oRs("RequestorPassword")
		PageStatus="mail"
		else
		ErrorMessage="That email address is not in our system."
	End if
        Set oConn=Nothing
        Set oRS=Nothing
	If lcase(PageStatus)="mail" then
		Body = "ATTN:&nbsp;&nbsp;"&RequestorName &"<br><br>"   & _
		"Below are your user name and password for the FleetX website.<br><br>"   & _
		"user name: "&Email&"<br>"  & _
		"password: "&Password&"<br><br>"   & _
		"Thank you,<br><br>"   & _
		"Mark Maggiore<br>"  & _
		"LogistiCorp Web Developer<br>"  & _
		"mark.maggiore@LogistiCorp.us<br>"  & _ 
		"817-591-2956<br><br>"
		Recipient=RequestorName


		'Set objMail = CreateObject("CDONTS.Newmail")
		'objMail.From = "FleetX@LogisticorpGroup.com"
		varTo = Email
		varSubject = "Username/Password Information"
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

	                        iMsg.To = varTo
                            iMsg.CC = varCC
	                        iMsg.BCC = "Mark.Maggiore@Logisticorp.us"
	                        SentMail="y"
                        With iMsg
	                        Set .Configuration = iConf
	                        .From ="System.Notification@logisticorp.us"
	                        .Subject = varSubject
	                        .HTMLBody = Body
	                        .Send
                        End With

		'if not Mailer.SendMail then
		  	'ErrorMessage = "Please try again later as the Email server is experiencing difficulties"
			'else
  			ErrorMessage = "An email has been sent to "&email&".<br>You should be receiving your information shortly."
		'end if	
	end if
End if
%>
</head>

<body onload="document.FindUser.requestorName.focus();document.FindUser.requiredemail.focus();">
	<table align="center" border="0" bordercolor="black" cellpadding="0" cellspacing="0" ID="Table1" height="100%" width="100%">
        <tr><td align="left"><img src="images/pixel.gif" height="10" width="1" /></td></tr>
        <tr>
            <td align="left"><img src="images/FleetX_Small.jpg" height="50" width="168" /></td>
            <td align="right" valign="bottom"></td>
        </tr>
        <tr><td align="left" colspan="2"><img src="images/pixel.gif" /></td></tr>	
    <tr><td colspan="2">
<!-- <form action="NewUser.asp" method="post" name="FindUser">    -->
<table border="0" bordercolor="green" Cellspacing="0" Cellpadding="0" align="left" width="100%">
 <tr><td align="left" colspan="2" bgcolor="<%=HeaderBorderColor%>"><img src="images/pixel.gif" height="5" width="1" /></td></tr>

        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetXRedSection" colspan="2"><%=PageTitle%></td></tr>

        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2"><img src="images/pixel.gif" height="5" width="1" /></td></tr>
        <!--tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressBodyWhite" colspan="2">In order to reach the Fleet Express Order page, please correctly type in the green verification code in the supplied text box and click "Submit."</td></tr-->
        <tr><td align="left" colspan="2"><img src="images/pixel.gif" height="5" width="1" /></td></tr>
<tr><td>
<table  border="0" bordercolor="blue" align="center" class="MainPageText" width="100%">
	<tr height="40">
		<td width="650">&nbsp;</td>
	</tr>



    <tr><td align="center" colspan="2"><!-- main page stuff goes here! -->
    
        <table width="<%=TaskmasterLogoWidth%>" border="0" cellpadding="0" cellspacing="0" ID="Table1">

      	<tr><td>
      <form action="RequestUserInfo.asp" method="post" name="FindUser">
      <table Width="750" Cellspacing="0" Cellpadding="0" align="left" border=0>
      <tr><td>&nbsp;</td></tr>
      <tr><td>
      <table width="432" border="0" align="center" class="MainPageText">
      	<tr height="40">
      		<td width="150">&nbsp;</td>
      	</tr>
        <tr Height="30"> 
          <td colspan="5" valign="middle" align="left" class="MainPageText"> 
            	SUBMIT YOUR EMAIL ADDRESS BELOW AND YOUR USERNAME AND PASSWORD WILL BE 
      		EMAILED TO YOU.	
      	
          </td>
      
        </tr>
      <tr><td>&nbsp;</td></tr>
      
        <tr Height="30"> 
          <td NOWRAP valign="middle" align="right" class="MainPageText"> 
            EMAIL ADDRESS:
          </td>
      	<td width="5">&nbsp;</td>
          <td width="136"> 
           <INPUT TYPE="text" NAME="requiredemail" value="<%=email%>" size="30">
          </td>
      	<input type="hidden" name="pagestatus" value="find">
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
    <td colspan="3" align="center">
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

      <tr>
          <td align="center" colspan="3"> 
            <INPUT id="gobutton" name="gobutton" TYPE="submit" name="ButtonValue" VALUE="SUBMIT">
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
        <tr><td>&nbsp;</td></tr>
      	<tr><td align="center" class="MainPageText">
      		<%
      		'If pagestatus="mail" then
      		%>
      		<a href="login.asp"  class="FleetXRedMain">CLICK HERE</a> TO RETURN TO THE LOGIN PAGE.
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
    
    
    
    
    
    
    </td></tr>



 
	<tr Height="50">
		<td>&nbsp;</td>
	</tr>

</table>
</td></tr>

</table>
<!-- </form> -->
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
