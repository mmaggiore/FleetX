<%@ Language=VBScript %>
<HTML>
	<HEAD>
		<!--meta http-equiv="refresh" content="100"-->
		<link rel="stylesheet" type="text/css" href="mainStyleSheet.css">
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
        <%
        'PagePassword="mercer"
DATABASE="DATABASE=fc_mdt;DSN=SQLConnect;UID=sa;Password=cadre;"
INTRANET="DATABASE=lcintranet;DSN=Intranet;UID=sa;Password=cadre;"

fakesubmit=request.form("fakesubmit")
deletecookie=trim(request.querystring("removecookie"))
location=request.form("location")
password=request.form("password")
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
End if



If fakesubmit>"" and Trim(ErrorMessage)="" then
Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Response.Write "Intranet="&Intranet&"***<BR>"
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
            Response.write "location="&location&"<BR>"
                 Response.Cookies("LegalComputer").Expires = Date() + 3500
                Response.Cookies("LegalComputer")("ComputerID")=Location
                '''''''''''''SENDS EMAIL ALERT!
				    Body = "Someone just set a location identification cookie!:<br><br>"   
                    Body = Body & "COOKIE INFORMATION:<BR>"
                    Body = Body & "Location: "&  Location &"<br><br>"  
				    Body = Body & "Love,<br><br>"  
				    Body = Body & "Mark<br>"
                    Response.write "body="&body&"<BR>"  
				    'Recipient=FirstName&" "&LastName
			        SentToEmail="mark.maggiore@logisticorp.us"
				    'Email="KWETI.Mailbox@am.kwe.com"
				    'Email="mark@maggiore.net"
				    Set objMail = CreateObject("CDONTS.Newmail")
				    objMail.From = "System.Notification@logisticorp.us"
				    objMail.To = SentToEmail
				    'objMail.cc = RequestorEmailAddress
				    'objMail.cc = "4692269939@tmomail.net;pcurrin@ti.com"
				    objMail.Subject = "New Cookie Just Set"
				    objMail.MailFormat = cdoMailFormatMIME
				    objMail.BodyFormat = cdoBodyFormatHTML
				    objMail.Body = Body
				    objMail.Send
				    Set objMail = Nothing       		
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


        If DeleteCookie="y" then
            Response.Cookies("LegalComputer").Expires = Date() + 3500
            Response.Cookies("LegalComputer")("ComputerID")=""
        End if
        ComputerID=Request.Cookies("LegalComputer")("ComputerID")
        'Response.write "computerid="&computerid&"<BR>"
         %>
    </HEAD>
    <body>
        <%If validated="y" then %>
        <br /><b><font size="4" color="blue">
        Set this computer's cookie to<br />work with phone emulator
        <br /><br /></font></b>
        <% If trim(ComputerID)="" then %>
        <form method="post" action="SetCookietest.asp">
         Computer's Location:   <select name="Location">
         <option value="">Select Location</option>
        <%
        

			Set Recordset1 = Server.CreateObject("ADODB.Recordset")
			Recordset1.ActiveConnection = DATABASE
			Recordset1.Source = "SELECT st_pkey, st_id, sb_bt_id FROM fcshipto INNER JOIN fcshipbt ON fcshipto.st_id = fcshipbt.sb_st_id WHERE SB_BT_ID IN('36', '38', '9999') AND st_id<>'55' order by sb_bt_id, st_id"
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
        <%If trim(errormessage)>"" then %>
        <Font color="red"><b><%=ErrorMessage%></b></Font><br /><br />
        <%end if%>
        <input type="hidden" name="validated" value="y" ID="Hidden124">
        <input type="submit" value="submit" name="fakesubmit"
        <br /><br />
        </form>
      
        
        <%
        End if
        'Response.write "ComputerID="&ComputerID&"<br /><br />"
			If trim(ComputerID)>"" then
                Set Recordset1 = Server.CreateObject("ADODB.Recordset")
			    Recordset1.ActiveConnection = DATABASE
			    Recordset1.Source = "SELECT st_id, sb_bt_id FROM fcshipto INNER JOIN fcshipbt ON fcshipto.st_id = fcshipbt.sb_st_id WHERE SB_BT_ID IN('36', '38', '9999') AND st_id<>'55' and st_pkey='" & ComputerID &"'"
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
                    If ComputerBillToID="36" then ShowComputerBillToID="Wafer" end if
                    If ComputerBillToID="38" then ShowComputerBillToID="Reticle" end if
                    If ComputerBillToID="9999" then ShowComputerBillToID="Hummer" end if
                    %>
                    This computer is:  <%=ComputerLocationCode %> (<%=ShowComputerBillToID %>)
				    <%
                End if
			    Recordset1.Close()
			    Set Recordset1 = Nothing
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
            <input type="submit" name="submit" value="Return to phone emulator" />
            </form>
    </body>