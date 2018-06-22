<html>
<head>

<!-- #include file="../fleetexpress.inc" -->
<link rel="stylesheet" type="text/css" href="../css/Style.css">
<%
ID = trim(request.form("ID"))
ID = valid8(ID)
if len(ID) < 1 then
  ID = valid8(trim(request.querystring("ID")))
end if

ErrorMessage =""

Edit = valid8(trim(request.form("Edit")))

if Edit = "Y" then
      Set oConn = Server.CreateObject("ADODB.Connection")
      oConn.ConnectionTimeout = 100
      oConn.Provider = "MSDASQL"
      oConn.Open DATABASE
      NewStatus = valid8(request.form("actionStatus"))
      SQL="UPDATE PreExistingRequestor set CostCenter = '" & valid8(request.form("CostCenter")) & "', requestortype = '" & valid8(request.form("UserType")) & "', bt_id =" & valid8(request.form("btCust")) & ", RequestorStatus = '" & newStatus & "' WHERE requestorID=" & ID
      'response.write "22 sql=" & SQL & "<br>"
      SET oRs = oConn.Execute(SQL)
      SQL = "SELECT * from PreExistingRequestor WHERE requestorID = " & ID
      SET oRs = oConn.Execute(SQL)
      
      if NewStatus = "c" then
        'approved
        		Body = "ATTN:&nbsp;&nbsp;"& oRs("RequestorName") &"<br><br>"   & _
        		"Welcome to the FLEETX website!<br><br>" & _
            "Below are your user name and password.<br><br>"   & _
        		"user name: "&oRs("RequestorEmail")&"<br>"  & _
        		"password: "&oRs("RequestorPassword")&"<br><br>"   & _
                "To log in, click here: <a href='"& whichsite &"'/home.asp'>FleetX Site</a><br><br>" &_
        		"Thank you,<br><br>"   & _
        		"Mark Maggiore<br>"  & _
        		"LogistiCorp Web Developer<br>"  & _
        		"mark.maggiore@LogistiCorp.us<br>"  & _ 
        		"817-591-2956<br><br>"
        		Recipient=oRs("RequestorName")
        
        		'Set objMail = CreateObject("CDONTS.Newmail")
        		'objMail.From = "FleetX@LogisticorpGroup.com"
        		varTo = oRs("RequestorEmail")
        		varSubject = "Your FleetX account has been approved"
        		'objMail.MailFormat = cdoMailFormatMIME
        		'objMail.BodyFormat = cdoBodyFormatHTML
        		'objMail.Body = Body
        		'objMail.Send
        		'Set objMail = Nothing	
    '''''''''''''''''''''''''''''''''''''''''''''''''''
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
      elseif NewStatus = "d" then
        'disapproved
        		Body = "ATTN:&nbsp;&nbsp;"& oRs("RequestorName") &"<br><br>"   & _
        		"We were not able to approve your user request at this time for the FLEETX website.<br><br>"   & _
        		"Regards,<br><br>"   & _
        		"Mark Maggiore<br>"  & _
        		"LogistiCorp Web Developer<br>"  & _
        		"mark.maggiore@LogistiCorp.us<br>"  & _ 
        		"817-591-2956<br><br>"
        		Recipient=oRs("RequestorName")
        
        		'Set objMail = CreateObject("CDONTS.Newmail")
        		'objMail.From = "FleetX@LogisticorpGroup.com"
        		varTo = oRs("RequestorEmail")
        		varSubject = "Your FLEETX User Request"
        		'objMail.MailFormat = cdoMailFormatMIME
        		'objMail.BodyFormat = cdoBodyFormatHTML
        		'objMail.Body = Body
        		'objMail.Send
        		'Set objMail = Nothing
     '''''''''''''''''''''''''''''''''''''''''''''''''''
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
      end if
      response.redirect "NewUserApproval.asp"
end if

''''''''DISPLAY NEW USER
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.ConnectionTimeout = 100
oConn.Provider = "MSDASQL"
oConn.Open DATABASE
SQL="SELECT * from PreExistingRequestor WHERE requestorID=" & ID
SET oRs = oConn.Execute(SQL)
if oRs.EOF then
  ErrorMessage = "User not found" 
end if 

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
    PageTitle="NEW USER REVIEW/APPROVAL"

%>
<title>FleetX - <%=PageTitle %></title>
</head>

<body onload="document.NewUserReview.RequestorName.focus();">
	<table align="center" border="0" bordercolor="black" cellpadding="0" cellspacing="0" ID="Table1" height="100%" width="100%">
        <tr><td align="left"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>
        <tr>
            <td align="left"><img src="../images/FleetX_Small.jpg" height="50" width="168" /></td>
            <td align="right" valign="bottom"><!-- #include file="../topnavbar.asp" --></td>
        </tr>
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" /></td></tr>	
    <tr><td colspan="2">
<form action="NewUserReview.asp" method="post" name="NewUserReview">
<table border="0" bordercolor="green" Cellspacing="0" Cellpadding="0" align="left" width="100%">
 <tr><td align="left" colspan="2" bgcolor="<%=HeaderBorderColor%>"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>

        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetXRedSection" colspan="2"><%=PageTitle%></td></tr>

        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
        <!--tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressBodyWhite" colspan="2">In order to reach the Fleet Express Order page, please correctly type in the green verification code in the supplied text box and click "Submit."</td></tr-->
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
<tr><td>
<table  border="0" bordercolor="blue" align="center" class="MainPageText" width="100%">
	<tr height="40">
		<td width="650">&nbsp;</td>
	</tr>



    <tr><td align=center width="100%"><!-- main page stuff goes here! -->
    
    
 <%
%>
<table align="center" cellspacing=3 cellpadding=3>
<% if len(ErrorMessage) < 1 then %>
  <% if len(ErrorUpdate) > 0 then %>
    <tr><td colspan=6>
    <table width="100%" border="0" bordercolor="Yellow" cellspacing="0" cellpadding="0" align="center" class="MainPageText">
    	<tr><td>&nbsp;</td></tr>  
    	<tr>
        <td align="center" class="Errormessage"><%=ErrorUpdate%></td>
      </tr>
    	<tr><td>&nbsp;</td></tr>
    </table>
    </td></tr>
  <% end if %>
  <tr><td align="right"><b>Name</b></td><td><%=oRs("RequestorName")%></td></tr>
  <tr><td align="right"><b>Submitted Company</b></td><td><%=oRs("RequestorCompany")%></td></tr>
  <tr><td align="right"><b>Company</b></td><td>
  <select name="btCust">
  <%
  SQL = "SELECT * FROM fcbillto WHERE bt_status = 'c' order by bt_desc asc"
  SET oRsN = oConn.Execute(SQL)
  Do Until oRsN.EOF
        selectd = ""
        BillTo = oRs("bt_id")
        if NOT isNULL(BillTo) and len(BillTo) > 1 then
          if cint(BillTo) = cint(oRsN("bt_id")) then 
            selectd = " selected" 
          end if
        end if
        %><option value=<%=cint(oRsN("bt_id"))%> <%=selectd%>><%=trim(oRsN("bt_desc"))%></option><%
  oRsN.MoveNext
  Loop
  %>
  </select>
  &nbsp;&nbsp;&nbsp;&nbsp;<a href="CompanyList.asp" class="FleetXRedMain">Add a Company</a>
  </td></tr>
  <tr><td align="right"><b>Cost Center</b></td><td><input type="text" maxlength=15 name="CostCenter" value="<%=oRs("CostCenter")%>"></td></tr>
  <tr><td align="right"><b>Address</b></td><td><%=oRs("RequestorAddress")%></td></tr>
  <tr><td align="right"><b>City</b></td><td><%=oRs("RequestorCity")%></td></tr>
  <tr><td align="right"><b>State</b></td><td><%=oRs("RequestorState")%></td></tr>
  <tr><td align="right"><b>Zip code</b></td><td><%=oRs("RequestorZipcode")%></td></tr>
  <tr><td align="right"><b>Phone</b></td><td><%=oRs("RequestorPhone")%></td></tr>
  <tr><td align="right"><b>Email</b></td><td><%=oRs("RequestorEmail")%></td></tr>
  <tr><td align="right"><b>User Type</b></td><td>
  <% utype = trim(oRs("RequestorType")) %>
  <input type="radio" name="UserType" value="" <% if (ISNULL(utype) or len(utype) < 1) then response.write " checked" end if%>> Standard User&nbsp;&nbsp;
  <input type="radio" name="UserType" value="A" <% if trim(oRs("RequestorType")) = "A" then response.write " checked" end if%>> Admin<br>
  </td></tr>

  <tr><td align="right"><b>Action</b></td><td>
  <select name="actionStatus">
    <option value="c">APPROVE</option>
    <option value="d">DISAPPROVE</option>
  </select>
  </td></tr>
  <tr><td> </td><td><INPUT id="gobutton" name="gobutton" TYPE="submit" name="ButtonValue" VALUE="SUBMIT"></td></tr> 
  <input type="hidden" name="ID" value="<%=id%>">
  <input type="hidden" name="Edit" value="Y">
<% else %>
  <tr><td colspan=6>
  <table width="100%" border="0" bordercolor="Yellow" cellspacing="0" cellpadding="0" align="center" class="MainPageText">
  	<tr><td>&nbsp;</td></tr>  
  	<tr>
      <td align="center" class="Errormessage"><%=ErrorMessage%></td>
    </tr>
  	<tr><td>&nbsp;</td></tr>
  </table>
  </td></tr>
<% end if %>


  
<tr><td colspan=6>&nbsp;<br><br><a href="NewUserApproval.asp" class="FleetXRedMain">CLICK HERE</a> to Return to New Users List Page</td></tr>


</td></tr>
</table>

   
    </td></tr>



 
	<tr Height="50">
		<td>&nbsp;</td>
	</tr>

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

