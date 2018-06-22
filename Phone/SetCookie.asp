<%@ Language=VBScript %>
<HTML>
	<HEAD>
		<!--meta http-equiv="refresh" content="100"-->
<!-- #include file="FleetX.inc" -->
<!-- #include file="driverinfo.inc" -->
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
        <%


ComputerID=Request.Cookies("LegalComputer")("ComputerID")
PhoneUser=Request.ServerVariables("HTTP_UA_OS") 




If Trim(PhoneUser)<>"Windows CE (Pocket PC) - Version 5.1" and Trim(PhoneUser)<>"Windows CE (Pocket PC) - Version 5.2"  then

				Body = "The users OS is<BR><BR>"& PhoneUser &"<br><br>"
                Body = Body & "USER INFORMATION:<BR>"  
                Body = Body & "UserName: "&  UserID &"<br><br>"   
                Body = Body & "FirstName: "&  FirstName &"<br><br>" 
                Body = Body & "LastName: "&  LastName &"<br><br>" 
                Body = Body & "COOKIE INFORMATION:<BR>"
                Body = Body & "Location: "&  Location &"<br><br>"  
				Body = Body & "Love,<br><br>"  
				Body = Body & "Mark<br>"
				'Recipient=FirstName&" "&LastName
				SentToEmail="mark.maggiore@logisticorp.us"
				'Email="KWETI.Mailbox@am.kwe.com"
				'Email="mark@maggiore.net"
				'Set objMail = CreateObject("CDONTS.Newmail")
				'objMail.From = "System.Notification@logisticorp.us"
				varTo = SentToEmail
				'objMail.cc = "4692269939@tmomail.net;htmlmale@yahoo.com"
				'objMail.cc = "4692269939@tmomail.net;pcurrin@ti.com"
				varSubject = "User OS-FROM SET COOKIE PAGE"
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



End if




        'PagePassword="mercer"
DATABASE="DATABASE=fc_mdt;DSN=SQLConnect;UID=sa;Password=cadre;"
INTRANET="DATABASE=lcintranet;DSN=Intranet;UID=sa;Password=cadre;"

fakesubmit=request.form("fakesubmit")
deletecookie=trim(request.querystring("removecookie"))
location=request.form("location")
password=request.form("password")
adminpassword=request.form("adminpassword")

submit=request.form("submit")
validated=request.form("validated")
validated="y"
UserName=Request.Form("UserName")
Username=Replace(Username,"'","")
Username=Replace(Username,"""","")
'''''''''''''''''''''''''''''''''''''''''
Password=Request.Form("Password")
Password=Replace(Password,"'","")
Password=Replace(Password,"""","")

If FakeSubmit>"" then
    IF trim(UserName)=""  then
	    ErrorMessage="You must scan your username"
    End if
    IF trim(Password)="" then
	    ErrorMessage="You must scan your password"
    End if
    If trim(Location)="" then
        ErrorMessage="You must select a location"
    End if
    'IF trim(AdminPassword)<>"chili" and trim(AdminPassword)<>"emergency" then
	'    ErrorMessage="You must enter the correct admin password"
    'End if
End if

'Response.write "userID="&userID&"<BR>"
'Response.write "ComputerID="&ComputerID&"<BR>"

If fakesubmit>"" and Trim(ErrorMessage)="" then
Set Recordset1 = Server.CreateObject("ADODB.Recordset")
'response.write "Intranet="&Intranet&"***<BR>"
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
		FirstName=Recordset1("FirstName")
		LastName=Recordset1("LastName")
		LoggedIn="yes"
		'Recordset1.Close()
		'Set Recordset1 = Nothing
            'response.write "location="&location&"<BR>"
                 Response.Cookies("LegalComputer").Expires = Date() + 3500
                Response.Cookies("LegalComputer")("ComputerID")=Location
                '''''''''''''SENDS EMAIL ALERT!
				    Body = "Someone just set a location identification cookie!:<br><br>"
                    Body = Body & "USER INFORMATION:<BR>"  
                    Body = Body & "UserName: "&  UserName &"<br><br>"   
                    Body = Body & "COOKIE INFORMATION:<BR>"
                    Body = Body & "Location: "&  Location &"<br><br>"  
				    Body = Body & "Love,<br><br>"  
				    Body = Body & "Mark<br>"
                    'response.write "body="&body&"<BR>"  
				    'Recipient=FirstName&" "&LastName
			        SentToEmail="mark.maggiore@logisticorp.us"
				    'Email="KWETI.Mailbox@am.kwe.com"
				    'Email="mark@maggiore.net"
				    'Set objMail = CreateObject("CDONTS.Newmail")
				    'objMail.From = "System.Notification@logisticorp.us"
				    varTo = SentToEmail
				    'objMail.cc = RequestorEmailAddress
				    'objMail.cc = "4692269939@tmomail.net;pcurrin@ti.com"
				    varSubject = "New Cookie Just Set"
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

		'Response.Redirect("DriverMessage.asp")
		
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


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If trim(Fakesubmit)="" then
            DeleteCookie="y"
        End if
        If DeleteCookie="y" then
            Response.Cookies("LegalComputer").Expires = Date() + 3500
            Response.Cookies("LegalComputer")("ComputerID")=""
        End if
        ComputerID=Request.Cookies("LegalComputer")("ComputerID")
        'Response.write "computerid="&computerid&"<BR>"
         %>
    </HEAD>
    <body>
      <!-- #include file="LogoSection.asp" -->

		<table width="300" cellpadding="0" cellspacing="0" border="0" align="left" ID="Table2">
            <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>
			<tr><td align="center" ><a href="default.asp" class="mainpagelink"><input type="submit" value="Return to Menu" id="gobutton" name="Submit3" /></td></tr>
                        <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>
                        <tr>
		                    <td class="FleetXRedSection" colspan="6" align="center">
			                    Set Location Cookie
		                    </td>
	                    </tr>
                        <tr>
                            <td>
                               <%
		                        If trim(ComputerID)="" and Trim(PhoneUser)<>"Windows CE (Pocket PC) - Version 5.1" and Trim(PhoneUser)<>"Windows CE (Pocket PC) - Version 5.2"  then
			                        RequireCookie="y"
			                        Response.write "<font color='red'>2/7/11-To use this computer<br>to drop off/pick up orders, select the<br>dropzone and scan in your username<BR>and password.</font>"
		                        End if
                                %>                           
                            </td>
                        </tr>
                        <tr><td>&nbsp;</td></tr>
                        <tr><td>
        <%If validated="y" then %>
        <% If trim(ComputerID)="" then %>
        <form method="post" action="SetCookie.asp">
         Computer's Location:   <select name="Location">
         <option value="">Select Location</option>
         <%'if trim(userid)="1" or trim(userid)="60"  then %>
         <%if trim(userid)="1" or trim(userid)="508" or trim(UserID)="82" then %>
          <option value="14141414">God Mode</option>
        <%
            End if
        

			Set Recordset1 = Server.CreateObject("ADODB.Recordset")
			Recordset1.ActiveConnection = DATABASE
			Recordset1.Source = "SELECT st_pkey, st_id, sb_bt_id FROM fcshipto INNER JOIN fcshipbt ON fcshipto.st_id = fcshipbt.sb_st_id WHERE SB_BT_ID IN('26','36', '38') AND st_id<>'55' order by sb_bt_id, st_id"
			'response.write "Recordset1.Source="&Recordset1.Source&"<BR>"
			Recordset1.CursorType = 0
			Recordset1.CursorLocation = 2
			Recordset1.LockType = 1
			Recordset1.Open()
			Recordset1_numRows = 0
			If Recordset1.eof then
				ErrorMessage="That is not a valid location"
			End if			
			Do While NOT Recordset1.EOF 
				LocationCode=Recordset1("st_id")
				BillToID=Trim(cStr(Recordset1("sb_bt_id")))
                st_pkey=Trim(cStr(Recordset1("st_pkey")))
                If BillToID="26" then ShowBillToID="Stockroom" end if
                If BillToID="36" then ShowBillToID="Wafer" end if
                If BillToID="38" then ShowBillToID="Reticle" end if
                If BillToID="9999" then ShowBillToID="Hummer" end if
                %>
                <option value="<%=st_pkey %>" <%if trim(Location)=trim(st_pkey) then response.write " selected "%>><%=LocationCode %> (<%=ShowBillToID %>)</option>
                <%
				Recordset1.Movenext
				Loop
			Recordset1.Close()
			Set Recordset1 = Nothing
        %>
        </select>
        <br /><br />
        <b>Scan from driver badge:</b><br />
        Username: <input type="password" name="UserName" />
        <br /><br />
        Password: <input type="password" name="password" />
        <br /><br />
        <!--
        Admin Password: <input type="adminpassword" name="adminpassword" />
        <br /><br />
        -->
        <%If trim(errormessage)>"" then %>
        <Font color="red"><b><%=ErrorMessage%></b></Font><br /><br />
        <%end if%>
        <input type="hidden" name="validated" value="y" ID="Hidden124">
        <input type="submit" value="submit" name="fakesubmit" id="gobutton" />
        <br /><br />
        </form>
      
        
        <%
        End if
        'Response.write "ComputerID="&ComputerID&"<br /><br />"
			If trim(ComputerID)>"" then
                Set Recordset1 = Server.CreateObject("ADODB.Recordset")
			    Recordset1.ActiveConnection = DATABASE
			    Recordset1.Source = "SELECT st_id, sb_bt_id FROM fcshipto INNER JOIN fcshipbt ON fcshipto.st_id = fcshipbt.sb_st_id WHERE SB_BT_ID IN('26', '36', '38', '9999') AND st_id<>'55' and st_pkey='" & ComputerID &"'"
			    'response.write "Recordset1.Source="&Recordset1.Source&"<BR>"
			    Recordset1.CursorType = 0
			    Recordset1.CursorLocation = 2
			    Recordset1.LockType = 1
			    Recordset1.Open()
			    Recordset1_numRows = 0
			    If Recordset1.eof then
				    ErrorMessage="That is not a valid location"
			    End if			
			    If NOT Recordset1.EOF then 
				    ComputerLocationCode=Recordset1("st_id")
				    ComputerBillToID=Trim(cStr(Recordset1("sb_bt_id")))
                    If ComputerBillToID="26" then ShowComputerBillToID="Stockroom" end if
                    If ComputerBillToID="36" then ShowComputerBillToID="Wafer" end if
                    If ComputerBillToID="38" then ShowComputerBillToID="Reticle" end if
                    If ComputerBillToID="9999" then ShowComputerBillToID="Hummer" end if
                    %>
                    This computer is:  <%=ComputerLocationCode %> (<%=ShowComputerBillToID %>)
				    <%
                End if
			    Recordset1.Close()
			    Set Recordset1 = Nothing
                If trim(ComputerID)=14141414 then
                    ComputerLocationCode="God Mode"
                End if
            Response.write "This computer is currently set as "& ComputerLocationCode &".<BR><br>"
            %>
            <!--
            <form method="post">
                <br /><br />To remove this computer's cookie 
                <input type="hidden" name="validated" value="y" ID="Hidden1">
                <input type="hidden" name="removecookie" value="y" ID="Hidden2">
                <input type="submit" name="submit" value="click here" />
            </form>
            -->
            <%
            else
            'Response.write "This is not currently an authorized computer.<BR><BR>"
            End if
            else
                'Response.write "<br><br>You have reached this page without going through proper channels...Please go away."
            End if
            %>
            <form method="post" action="default.asp">
            <input type="submit" name="submit" value="Return to emulator" id="gobutton" />
            </form>
</td></tr>
</table>
    </body>