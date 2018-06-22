<html>
<head>

<!-- #include file="../fleetexpress.inc" -->
<link rel="stylesheet" type="text/css" href="../css/Style.css">
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
    PageTitle="Edit User Information"


    PhoneBookID=Session("PhoneBookID")
    If trim(PHoneBookID)="" then
        response.Redirect("FreightOrder.asp")
    End if
    Submit=Request.Form("Submit")
    PageStatus=valid8(Request.Form("PageStatus"))
    'Response.write "PageStatus="&PageStatus&"<BR>"
If trim(PageStatus)<>"find" then
		    Set oConn = Server.CreateObject("ADODB.Connection")
		    oConn.ConnectionTimeout = 100
		    oConn.Provider = "MSDASQL"
		    oConn.Open DATABASE
			    l_cSQL = "SELECT * FROM PreExistingRequestor WHERE (RequestorID='"& PhoneBookID &"') and  (RequestorStatus <> 'x')"
			    'Response.write "l_cSQL="&l_cSQL&"<BR>"
                SET oRs = oConn.Execute(l_cSql)
					    if oRs.EOF then
                            Email=valid8(request.form("RequiredEmail"))
                            PageStatus=valid8(request.form("PageStatus"))
                            RequestorName=valid8(request.form("RequestorName"))
                            RequestorCompany=valid8(request.form("RequestorCompany"))
                            RequestorAddress=valid8(request.form("RequestorAddress"))
                            RequestorCity=valid8(request.form("RequestorCity"))
                            RequestorZipCode=valid8(request.form("RequestorZipCode"))
                            RequestorPhone=valid8(request.form("RequestorPhone"))
                            RequestorEmail=valid8(request.form("RequestorEmail"))
                            RetypeRequestorEmail=valid8(request.form("RetypeRequestorEmail"))
                            RequestorPassword=valid8(request.form("RequestorPassword"))
                            ErrorMessage="That username/password combination is not valid."
                            else

                            'PageStatus=oRs("PageStatus")
                            RequestorName=oRs("RequestorName")
                            RequestorCompany=oRs("RequestorCompany")
                            RequestorAddress=oRs("RequestorAddress")
                            RequestorCity=oRs("RequestorCity")
                            RequestorZipCode=oRs("RequestorZipCode")
                            RequestorPhone=oRs("RequestorPhone")
                            RequestorEmail=oRs("RequestorEmail")
                            RetypeRequestorEmail=RequestorEmail
                            RequestorPassword=oRs("RequestorPassword")
 
                        End if								
		    Set oConn=Nothing
        else

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
            HeaderBorderColor="#41924B"  
            BorderColor="#41924B"
            LinkClass="FleetExpressGreen"
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


PageStatus=valid8(Request.Form("PageStatus"))
'Response.Write "Intranet="&Intranet&"<BR>"
If lcase(PageStatus)="find" then
                            Email=valid8(request.form("RequiredEmail"))
                            PageStatus=valid8(request.form("PageStatus"))
                            RequestorName=valid8(request.form("RequestorName"))
                            RequestorCompany=valid8(request.form("RequestorCompany"))
                            RequestorAddress=valid8(request.form("RequestorAddress"))
                            RequestorCity=valid8(request.form("RequestorCity"))
                            RequestorZipCode=valid8(request.form("RequestorZipCode"))
                            RequestorPhone=valid8(request.form("RequestorPhone"))
                            RequestorEmail=valid8(request.form("RequestorEmail"))
                            RetypeRequestorEmail=valid8(request.form("RetypeRequestorEmail"))
                            RequestorPassword=valid8(request.form("RequestorPassword"))



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

    RequestorPhone=Replace(RequestorPhone, """", "")
    RequestorPhone=Replace(RequestorPhone, "'", "")
    RequestorEmail=Replace(RequestorEmail, """", "")
    RequestorEmail=Replace(RequestorEmail, "'", "")
    RequestorPassword=Replace(RequestorPassword, """", "")
    RequestorPassword=Replace(RequestorPassword, "'", "")

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
'Response.write "ErrorMessage="&ErrorMessage&"<BR>"
If trim(ErrorMessage)="" then


    'Response.write "PhoneBookID="&PhoneBookID&"<BR>"
	Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
		RSEVENTS2.Open "PreExistingRequestor", DATABASE, 2, 2
		RSEVENTS2.Find "RequestorID='" & PhoneBookID & "'"
		RSEVENTS2("RequestorName")=RequestorName
        RSEVENTS2("RequestorCompany")=RequestorCompany
        RSEVENTS2("RequestorAddress")=RequestorAddress
        RSEVENTS2("RequestorCity")=RequestorCity
        RSEVENTS2("RequestorState")="TX"
        RSEVENTS2("RequestorZipCode")=RequestorZipCode
        RSEVENTS2("RequestorPhone")=RequestorPhone
        RSEVENTS2("RequestorEmail")=RequestorEmail
        RSEVENTS2("RequestorPassword")=RequestorPassword
        'RSEVENTS2("RequestorStatus")="c"
		RSEVENTS2.update
		RSEVENTS2.close			
	set RSEVENTS2 = nothing 	
	'If lcase(PageStatus)="mail" then
		Body = "ATTN:&nbsp;&nbsp;"&RequestorName &"<br><br>"   & _
		"Below are your user name and password for the FleetX Website.<br><br>"   & _
		"user name: "&RequestorEmail&"<br>"  & _
		"password: "&RequestorPassword&"<br><br>"   & _
		"The address is: https://www.FleetXDFW.com <br><br>"   & _ 			
		"If you have any questions, please do not hesitate to contact me.<br><br>"   & _
		"Thank you,<br><br>"   & _
		"Mark Maggiore<br>"  & _
		"LogistiCorp Web Developer<br>"  & _
		"mark.maggiore@LogistiCorp.us<br>"  & _ 
		"214/956-0650 xt 212<br><br>"
		Recipient=LastName


		'Set objMail = CreateObject("CDONTS.Newmail")
		'objMail.From = "system.monitor@logisticorp.us"
		'objMail.To = RequestorEmail
		'objMail.Subject = "Congratulations new Fleet Express User"
		'objMail.MailFormat = cdoMailFormatMIME
		'objMail.BodyFormat = cdoBodyFormatHTML
		'objMail.Body = Body
		'objMail.Send
		'Set objMail = Nothing	
'End if

		'if not Mailer.SendMail then
		  	'ErrorMessage = "Please try again later as the Email server is experiencing difficulties"
			'else
  			ErrorMessage = "Your information has been successfully updated."
            'Dontshow="y"
		'end if	
	end if
End if

%>
<title>FleetX - <%=PageTitle %></title>
</head>

<body onload="document.FindUser.requestorName.focus();">
	<table align="center" border="0" bordercolor="black" cellpadding="0" cellspacing="0" ID="Table1" height="100%" width="100%">
        <tr><td align="left"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>
        <tr>
            <td align="left"><img src="../images/FleetX_Small.jpg" height="50" width="168" /></td>
            <td align="right" valign="bottom"><!-- #include file="../topnavbar.asp" --></td>
        </tr>
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" /></td></tr>	
    <tr><td colspan="2">
<form action="EditUserInfo.asp" method="post" name="FindUser">
<table border="0" bordercolor="green" Cellspacing="0" Cellpadding="0" align="left" width="100%">
 <tr><td align="left" colspan="2" class="FleetXRedSection"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>

        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetXRedSection" colspan="2"><%=PageTitle%></td></tr>

        <tr><td align="center" class="FleetXRedSection" class="FleetExpressHeaderWhite" colspan="2"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
        <!--tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressBodyWhite" colspan="2">In order to reach the Fleet Express Order page, please correctly type in the green verification code in the supplied text box and click "Submit."</td></tr-->
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
<tr><td>
<table  border="0" bordercolor="blue" align="center" class="MainPageText" width="100%">
	<tr height="40">
		<td width="650">&nbsp;</td>
	</tr>
    <tr><td>
    
    
    
    
    
    
    
    

<form action="EditUserInfo.asp" method="post" name="FindUser">

<table  border="0" bordercolor="black" align="center" class="MainPageText">
	<tr height="40">
		<td width="650">&nbsp;</td>
	</tr>
    <%if trim(dontshow)<>"y" then %>
  <tr Height="30"> 
    <td colspan="5" valign="middle" align="center" class="MainPageText"> 
      	Update your information and click submit.
	
    </td>

  </tr>
<tr><td>&nbsp;</td></tr>


  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
      Name:
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
      Company:
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
      Address:
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
        City:
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
        State:
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
        Zip Code:
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
        Phone Number:
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
        Email Address:
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
        Re-type Email Address:
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
      Password:
    </td>
	<td width="5">&nbsp;</td>
    <td width="136"> 
     <INPUT TYPE="text" NAME="requestorpassword" value="<%=requestorpassword%>" maxlength="15" size="30">
    </td>
	<input type="hidden" name="pagestatus" value="find">
	<td width="5">&nbsp;</td>
    <td width="650"> 
      <input id="gobutton" type="submit" name="ButtonValue" value="Submit">
    </td>
  </tr>
	<tr Height="50">
		<td>&nbsp;</td>
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
<%if ErrorMessage>"" then%>
<tr><td>
<table width="100%" border="0" bordercolor="blue" cellspacing="0" cellpadding="0" align="center" class="MainPageText">
	<tr><td>&nbsp;</td></tr>  
	<tr>
    <td align="center" class="Errormessage"><%=ErrorMessage%></td>
  </tr>
	<tr><td>&nbsp;</td></tr>
</table>
</td></tr>
<%end if%>
<tr><td align="center">
<table border="0" bordercolor="brown" cellspacing="0" cellpadding="0" align="center">
  <tr><td>&nbsp;</td></tr>
	<tr><td align="center" class="MainPageText">
To return to the Fleet Express Order Page, <a href="FreightOrder.asp?loggedin=y&varA=123"  class="FleetXRedMain">Click Here</a>.
</td></tr>
</table>
</form>

   
    
    
    
    
    
    
    
    
    
   </td></tr>
	<tr Height="50">
		<td>&nbsp;</td>
	</tr>


</table>
</td></tr>

</table>
</form>
<tr><td Height="90%">&nbsp;</td></tr>
<tr>
    <td height="100" class="FleetXGreySection" colspan="2">
        <!-- #include file="../BottomSection.asp" -->
    </td>
</tr>
<tr><td height="15" class="FleetXRedSectionSmall" colspan="2" align="center"><%=CopywriteNotice %></td></tr>
</td></tr></table>


</body>
</html>
