<html>
<head>

<!-- #include file="fleetexpress.inc" -->
<link rel="stylesheet" type="text/css" href="css/Style.css">
<%
'response.write "xxx="&lcase(left(Request.ServerVariables ("URL"), 4))
SecureYes = Request.ServerVariables ("HTTPS")
If SecureYes="off" and sitename<>"TEST" then
	''''''''''''''''''''''''''''''''''''''''''''
	something=Request.ServerVariables("HTTP_HOST") 
	Something=left(Something,4) 

	'''''''''''''''''''''''''''''''''''''''''''
	Response.redirect(Whichsite&"/login.asp")	
End if

varCaptcha=valid8(Request.Form("varCaptcha"))
CaptchaSubmit=valid8(Request.Form("CaptchaSubmit"))
txtUserName=valid8(Request.form("txtUserName"))
txtPassword=valid8(Request.form("txtPassword"))
GoButton=valid8(Request.form("GoButton"))
'response.write "GoButton="&GoButton&"<BR>"
Submit=valid8(Request.Form("Submit"))
Logout=valid8(request.querystring("Logout"))
If trim(Logout)="y" then



Response.Buffer=True
Dim objCookie
'loop through cookie collection
'Response.Write "Deleting cookies...<BR>"
For Each objCookie In Request.Cookies
    'delete the cookie by setting its expiration Date
    'to a Date In the past
    Response.Cookies(objCookie).Expires = "September 7, 1998"
Next
'Response.Write "Done."



    'Response.write "I DONE GOT HERE!<BR>"
    Response.Cookies ("FleetXCookie")("UserID") = ""
    Response.Cookies("FleetXCookie").expires = dateadd("n",+240,now())
    Response.Cookies ("FleetXCookie")("UserName") = ""
    Response.Cookies("FleetXCookie").expires = dateadd("n",+240,now())
    Response.Cookies ("FleetXCookie")("BillToID") = ""
    Response.Cookies("FleetXCookie").expires = dateadd("n",+240,now())
End if

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
    PageTitle="LOG IN"

If trim(GoButton)="SUBMIT" then




    'response.write "HEY>>>I GOT HERE!<BR>"
    'Response.Write "Intranet="&Intranet&"<BR>"
    '--------------------------------------------
    '''''''''''''''''''''''''''''''''''''''''''''''
    If CaptchaSubmit<>varCaptcha then
        ErrorMessage="You did not supply the correct verification code"
    End if
    If trim(txtPassword)="" then
        ErrorMessage="You must enter your password"
    End if
    If trim(txtUserName)="" then
        ErrorMessage="You must enter your user name"
    End if
    If trim(CaptchaSubmit)="" and trim(txtPassword)="" and trim(txtUserName)="" then
        ErrorMessage="You must enter username, password, and the verification code"
    End if

  
    If trim(ErrorMessage)="" then
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''

    'response.write "XXXXXXXXX got here 2 the LOGIN PART!<br>"
    Set Recordset1 = Server.CreateObject("ADODB.Recordset")
    'Response.write "Intranet="&Intranet&"<br>"
    Recordset1.ActiveConnection = Database

    Recordset1.Source = "SELECT RequestorID, RequestorName, RequestorCompany, bt_id FROM PreExistingRequestor WHERE (RequestorEmail='"&txtUserName&"') AND (REQUESTORPASSWORD='"&txtPassword&"') AND (RequestorStatus='c')"
    'response.write "Recordset1.Source="& Recordset1.Source &"<BR>"
    Recordset1.CursorType = 0
    Recordset1.CursorLocation = 2
    Recordset1.LockType = 1
    Recordset1.Open()
    Recordset1_numRows = 0
    'response.write "<font color='red'>SQL1="&Recordset1.Source&"<BR></font>"
	    if NOT Recordset1.EOF then
		    UserID=Recordset1("RequestorID")
            UserName=Recordset1("RequestorName")
            RequestorCompany=Recordset1("RequestorCompany")
            BillToID=Recordset1("bt_id")
            Response.Cookies ("FleetXCookie")("UserID") = UserID
            Response.Cookies("FleetXCookie").expires = dateadd("n",+120,now())
            Session("UserID")=UserID
            Response.Cookies ("FleetXCookie")("UserName") = UserName
            Response.Cookies("FleetXCookie").expires = dateadd("n",+120,now())
            session("UserName")=Username
            'Response.write "***UserID="&UserID&"***<BR>"
            'Response.write "***RequestorCompany="&RequestorCompany&"***<BR>"
            Response.Cookies ("FleetXCookie")("RequestorCompany") = RequestorCompany
            Response.Cookies("FleetXCookie").expires = dateadd("n",+120,now())
            Session("RequestorCompany")=RequestorCompany
            Response.write "LIne 124 - BillToID="&BillToID&"<br>"
            Response.Cookies ("FleetXCookie")("BillToID") = BillToID
            Response.Cookies("FleetXCookie").expires = dateadd("n",+120,now())
            Session("BillToID")=BillToID
            'Response.write "UserID="&UserID&"<BR>"
	        'UserID=Request.cookies("FleetXCookie")("UserID")
	        'UserName=Request.cookies("FleetXCookie")("UserName")
            'Response.write "***UserID="&UserID&"***<BR>"
             Response.redirect("home.asp")
            Else
            ErrorMessage="That username/password combination does not exist"
        End if
   	Recordset1.Close()
	Set Recordset1 = Nothing

    End if
End if



%>
<title>FleetX - <%=PageTitle %></title>
</head>

<body onload="document.FindUser.Text1.focus();">
	<table align="center" border="0" bordercolor="black" cellpadding="0" cellspacing="0" ID="Table1" height="100%" width="100%">
        <tr><td align="left"><img src="images/pixel.gif" height="10" width="1" /></td></tr>
        <tr>
            <td align="left"><img src="images/FleetX_Small.jpg" height="50" width="168" /></td>
            <td align="right" valign="bottom"></td>
        </tr>
        <tr><td align="left" colspan="2"><img src="images/pixel.gif" /></td></tr>	
    <tr><td colspan="2">
<form action="Login.asp" method="post" name="FindUser">
<table border="0" bordercolor="green" Cellspacing="0" Cellpadding="0" align="left" width="100%" class="MainPageText" >
 <tr><td align="left" colspan="2" bgcolor="<%=HeaderBorderColor%>"><img src="images/pixel.gif" height="5" width="1" /></td></tr>

        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetXRedSection" colspan="2"><%=PageTitle%></td></tr>

        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2"><img src="images/pixel.gif" height="5" width="1" /></td></tr>
        <!--tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressBodyWhite" colspan="2">In order to reach the Fleet Express Order page, please correctly type in the green verification code in the supplied text box and click "Submit."</td></tr-->
        <tr><td align="left" colspan="2"><img src="images/pixel.gif" height="5" width="1" /></td></tr>
        <tr><td>NEW USER?  <A href="NewUser.asp" class="FleetXRedMain">CLICK HERE</a> TO REGISTER</td></tr>
        <tr><td>&nbsp;</td></tr>
        <tr><td><a href="images/FleetXTrainingDocumentationV2.pdf" class="FleetXRedMain" target="_blank">CLICK HERE</a> TO VIEW TRAINING DOCUMENTATION.</td></tr>
        <tr><td>&nbsp;</td></tr>
        <tr><td>TRACK AN ORDER ONLY?<a href="tracking/GenericTracking.asp" class="FleetXRedMain"> &nbsp;&nbsp;CLICK HERE</a></td></tr>       
<tr><td>
<table  border="0" bordercolor="blue" align="center" class="MainPageText">
	<tr height="40">
		<td width="650">&nbsp;</td>
	</tr>

    <tr><td>&nbsp;</td></tr>
    <tr><td>&nbsp;</td></tr>
    <tr><td align="center" class="errormessage">&nbsp;<%=ErrorMessage %>&nbsp;</td></tr>
    <tr><td>&nbsp;</td></tr>
    <tr><td>&nbsp;</td></tr>
    <tr><td colspan="5" align="center">
        
        <table cellpadding="0" cellspacing="0" border="0" bordercolor="red">

            <tr Height="30">
                <td NOWRAP valign="bottom" align="right" class="MainPageText">
                    USERNAME:
                </td>
	            <td width="5">&nbsp;</td>
                <td width="136"> 
                  <input type="text" name="txtUserName" class="MainPageTextPlain" value="<%=txtUserName%>" ID="Text1">
                </td>
            </tr>
            <tr Height="30"> 
                <td NOWRAP valign="middle" align="right" class="MainPageText"> 
                      PASSWORD:
                </td>
                <td width="5">&nbsp;</td>
                <td width="136"> 
                      <input type="password" name="txtPassword" class="MainPageText" value="<%=txtPassword%>" ID="Password1">
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
      <tr><td>&nbsp;</td></tr>    
      <tr>
        <td align="center" colspan="3">
            <input id="gobutton" name="gobutton" type="submit" value="SUBMIT" />
        </td>
      </tr>
  </form>


<%
  ' Delete the captchas object.
  Set captchas = Nothing
%>

        </table>
    </td></tr>



 
	<tr Height="50">
		<td>&nbsp;</td>
	</tr>


</table>
</td></tr>
<%
if ErrorMessage>"" then%>
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
</table>

<tr><td Height="90%">&nbsp;</td></tr>
 <tr><td>FORGOTTEN YOUR USERNAME/PASSWORD?  <A href="RequestUserInfo.asp" class="FleetXRedMain">CLICK HERE</a></td></tr>
<tr>
    <td height="100" class="FleetXGreySection" colspan="2">
        <!-- #include file="BottomSection.asp" -->
    </td>
</tr>
<tr><td height="15" class="FleetXRedSectionSmall" colspan="2" align="center"><%=CopywriteNotice %></td></tr>
</td></tr></table>


</body>
</html>
