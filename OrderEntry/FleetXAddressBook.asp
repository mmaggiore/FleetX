<html>
<head>

<!-- #include file="../fleetexpress.inc" -->
<link rel="stylesheet" type="text/css" href="../css/Style.css">
<%

    sortby=request.Form("SortBy")
    If trim(SortBy)="" then
        SortBy="CompanyName"
    End if
    spacersize=5
    Var1=valid8(Request.Querystring("Var1"))
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
    PageTitle="ADDRESS BOOK"



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
    PhoneBookID=UserID
    PreExistingRequestor=Request.Cookies("MyCookie")("PreExistingRequestor")
    If trim(PreExistingRequestor)>"" then
        PhoneBookID=PreExistingRequestor
    End if
    If trim(PHoneBookID)="" then
        response.Redirect("FreightOrder.asp?Var1="&Var1)
    End if
Message="Simply submit your e-mail address and your login name and password will be emailed to you"
Email=valid8(request.form("RequiredEmail"))
PageStatus=valid8(request.form("PageStatus"))
RequestorName=valid8(request.form("RequestorName"))
RequestorCompany=valid8(request.form("RequestorCompany"))
RequestorBuilding=valid8(request.form("RequestorBuilding"))
RequestorAddress=valid8(request.form("RequestorAddress"))
RequestorSuite=valid8(request.form("RequestorSuite"))
RequestorCity=valid8(request.form("RequestorCity"))
RequestorZipCode=valid8(request.form("RequestorZipCode"))
RequestorPhone=valid8(request.form("RequestorPhone"))
RequestorEmail=valid8(request.form("RequestorEmail"))
RetypeRequestorEmail=valid8(request.form("RetypeRequestorEmail"))
RequestorPassword=valid8(request.form("RequestorPassword"))
ContactName=valid8(request.Form("ContactName"))
UpdateCompanyID=valid8(Request.Form("UpdateCompanyID"))
UpdateStatus=valid8(Request.Form("UpdateStatus"))
CompanyCostCenter=valid8(Request.Form("CompanyCostCenter"))
Submit1=valid8(Request.Form("Submit1"))

If trim(lcase(Submit1))="edit" then
    Response.redirect("FleetXEditaddressBook.asp?loggedin=y&varA=123&var1="&Var1&"&varb="&updatecompanyid)
End if




If UpdateStatus="y" then
	    Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
	    RSEVENTS.Open "PreExistingCompanies", Database, 2, 2
	    RSEVENTS.Find "CompanyID='"& UpdateCompanyID &"'"
		    RSEVENTS("CompanyStatus") = "x"
	    RSEVENTS.update
	    RSEVENTS.close
	    set RSEVENTS = nothing
        ErrorMessage="You have successfully deleted a location from your Address Book."
End if

'Response.Write "Intranet="&Intranet&"<BR>"
If lcase(PageStatus)="find" then
    RequestorName=Replace(RequestorName, """", "`")
    RequestorName=Replace(RequestorName, "'", "`")
    RequestorCompany=Replace(RequestorCompany, """", "`")
    RequestorCompany=Replace(RequestorCompany, "'", "`")
    RequestorBuilding=Replace(RequestorBuilding, """", "`")
    RequestorBuilding=Replace(RequestorBuilding, "'", "`")
    RequestorAddress=Replace(RequestorAddress, """", "`")
    RequestorAddress=Replace(RequestorAddress, "'", "`")
    RequestorSuite=Replace(RequestorSuite, """", "`")
    RequestorSuite=Replace(RequestorSuite, "'", "`")
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
    ContactName=Replace(ContactName, """", "`")
    ContactName=Replace(ContactName, "'", "`")
'If trim(RequestorPassword)="" then
'    ErrorMessage="You must enter a Password"
'End if
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
If trim(ContactName)="" then
    ErrorMessage="You must enter a Contact Name"
End if
If trim(CompanyCostCenter)>"" then
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.ConnectionTimeout = 100
	oConn.Provider = "MSDASQL"
	oConn.Open DATABASE
		l_cSQL = "Select * FROM TICostCenters WHERE costcenterstatus='c' and CostCenterNumber='"& CompanyCostCenter &"'"
		'Response.write "CostCenter="&CostCenter&"<BR>"
        SET oRs = oConn.Execute(l_cSql)
				if oRs.EOF then
                    'OrderStatus="1"
                    ErrorMessage="You did not provide a valid Cost Center number.<br><br>The 'Cost Center' number consists of 'C1' and your three digit division code and your five digit cost center number (ex. C112312345).<BR><BR>You are not required to supply a valid cost center here, so you can leave it blank.  However, you are required to have a valid cost center to place an order."
                    'REsponse.write "Line 781 - Got here!<BR>"
                    else
                End if								
	Set oConn=Nothing
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
'If trim(RequestorName)="" then
'    ErrorMessage="You must enter a Name"
'End if

If trim(ErrorMessage)="" then
    '''''''CREATE UNIQUE ID''''''''''''''''''''''
    'varUnique=Now()
    'Response.write "varUnique="&varUnique&"<BR>"
    'varUnique=Replace(varUnique, "AM","1")
    'Response.write "varUnique="&varUnique&"<BR>"
    'varUnique=Replace(varUnique, "PM","2")
    'Response.write "varUnique="&varUnique&"<BR>"
    'varUnique=Replace(varUnique, "/","")
    'Response.write "varUnique="&varUnique&"<BR>"
    'varUnique=Replace(varUnique, " ","")
    'Response.write "varUnique="&varUnique&"<BR>"
    'varUnique=Replace(varUnique, ":","")
    'Response.write "varUnique="&varUnique&"<BR>"
    'varUnique=phonebookid&varUnique
    'Response.write "varUnique="&varUnique&"<BR>"
    ''''''''''''''''''''''''''''''''''''''''''

TempRequestorCompany=lcase(RequestorCompany)
TempRequestorBuilding=lcase(RequestorBuilding)
varXYZ=InStr(TempRequestorCompany, "pack n` ship")
If varXYZ>0 then
    courierok="y"
End if
varXYZ=InStr(TempRequestorBuilding, "pack n` ship")
If varXYZ>0 then
    courierok="y"
End if

varXYZ=InStr(TempRequestorCompany, "DFW Test")
If varXYZ>0 then
    courierok="y"
End if
varXYZ=InStr(TempRequestorBuilding, "DFW Test")
If varXYZ>0 then
    courierok="y"
End if

varXYZ=InStr(TempRequestorCompany, "chip target")
If varXYZ>0 then
    courierok="y"
End if
varXYZ=InStr(TempRequestorBuilding, "chip target")
If varXYZ>0 then
    courierok="y"
End if
varXYZ=InStr(TempRequestorCompany, "priority packaging")
If varXYZ>0 then
    courierok="y"
End if
varXYZ=InStr(TempRequestorBuilding, "priority packaging")
If varXYZ>0 then
    courierok="y"
End if

varXYZ=InStr(TempRequestorCompany, "priority package")
If varXYZ>0 then
    courierok="y"
End if
varXYZ=InStr(TempRequestorBuilding, "priority package")
If varXYZ>0 then
    courierok="y"
End if

varXYZ=InStr(TempRequestorCompany, "priority lab")
If varXYZ>0 then
    courierok="y"
End if
varXYZ=InStr(TempRequestorBuilding, "priority lab")
If varXYZ>0 then
    courierok="y"
End if
varXYZ=InStr(TempRequestorCompany, "priority labs")
If varXYZ>0 then
    courierok="y"
End if
varXYZ=InStr(TempRequestorBuilding, "priority labs")
If varXYZ>0 then
    courierok="y"
End if

varXYZ=InStr(TempRequestorCompany, "vlsip")
If varXYZ>0 then
    courierok="y"
End if
varXYZ=InStr(TempRequestorBuilding, "vlsip")
If varXYZ>0 then
    courierok="y"
End if
varXYZ=InStr(TempRequestorCompany, "priority lab")
If varXYZ>0 then
    courierok="y"
End if
varXYZ=InStr(TempRequestorBuilding, "priority lab")
If varXYZ>0 then
    courierok="y"
End if
varXYZ=InStr(TempRequestorCompany, "priority labs")
If varXYZ>0 then
    courierok="y"
End if
varXYZ=InStr(TempRequestorBuilding, "priority labs")
If varXYZ>0 then
    courierok="y"
End if
If left(RequestorZipCode, 5)="75243" or courierok="y" then
    IsCourier="y" 
    else
    IsCourier="n"
End if
	Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
		RSEVENTS2.Open "PreExistingCompanies", DATABASE, 2, 2
		RSEVENTS2.addnew
        RSEVENTS2("CompanyOwner")=PhoneBookID
        RSEVENTS2("CompanyName")=RequestorCompany
        RSEVENTS2("CompanyBuilding")=RequestorBuilding
        RSEVENTS2("CompanyAddress")=RequestorAddress
        RSEVENTS2("CompanySuite")=RequestorSuite
        RSEVENTS2("CompanyCity")=RequestorCity
        RSEVENTS2("CompanyState")="TX"
        RSEVENTS2("CompanyZip")=RequestorZipCode
        RSEVENTS2("ContactName")=ContactName
        RSEVENTS2("CompanyPhone")=RequestorPhone
        RSEVENTS2("CompanyEmail")=RequestorEmail
        RSEVENTS2("CompanyCostCenter")=CompanyCostCenter
        RSEVENTS2("IsCourier")=IsCourier
        'RSEVENTS2("st_alias")=varUnique
        RSEVENTS2("CompanyStatus")="c"
		RSEVENTS2.update
		RSEVENTS2.close			
	set RSEVENTS2 = nothing 




	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.ConnectionTimeout = 100
	oConn.Provider = "MSDASQL"
	oConn.Open DATABASE
		l_cSQL = "Select CarillonID FROM CarillonIDs WHERE BillingStatus='c' and CompanyAddress='"& RequestorAddress &"'"
		'Response.write "CostCenter="&CostCenter&"<BR>"
        SET oRs = oConn.Execute(l_cSql)
				if oRs.EOF then
                    else
                    CarillonID=oRs("CarillonID")
                End if								
	Set oConn=Nothing


    If trim(CarillonID)>"" then
        Set oConn = Server.CreateObject("ADODB.Connection")
		    oConn.ConnectionTimeout = 100
		    oConn.Provider = "MSDASQL"
		    oConn.Open DATABASE
		    ''''UPDATES THE WAFER
		    l_cSQL = "UPDATE PreExistingCompanies SET CarillonID = '"&CarillonID&"', WhoSetCarillonID='"&UserID&"' WHERE (CompanyAddress = '"&RequestorAddress&"') AND (companystatus = 'c') And (CarillonID<1 or CarillonID is NULL)"
		    oConn.Execute(l_cSQL)
            oConn.close
        Set oConn=Nothing
        else
			'Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
			'	RSEVENTS2.Open "CarillonIDs", DATABASE, 2, 2
			'	RSEVENTS2.addnew
			'	RSEVENTS2("CarillonID")=""
			'	RSEVENTS2("CompanyAddress")=CompanyAddress
			'	RSEVENTS2("BillingStatus")="c"
			'	RSEVENTS2.update
			'	RSEVENTS2.close			
			'set RSEVENTS2 = nothing
            PersonsName="Theresa"
            PersonsEmail="anup.sharma@logisticorp.us"	
		    Body = "ATTN:&nbsp;&nbsp;"&PersonsName &"<br><br>"   & _
		    "A new address without a corresponding Carillon number has just been entered into our system..<br><br>"   & _
		    "Address: "&RequestorAddress&"<br><br>"  & _
		    "Please go to the 'Update Carillon Info' page and enter the correct number: "&WhichSite&"/Admin/UpdateCarillonInfo.asp <br><br>"   & _ 			
		    "If you have any questions, please do not hesitate to contact me.<br><br>"   & _
		    "Thank you,<br><br>"   & _
		    "Mark Maggiore<br>"  & _
		    "LogistiCorp Web Developer<br>"  & _
		    "mark.maggiore@LogistiCorp.us<br>"  & _ 
		    "214/956-0650 xt 212<br><br>"
		    Recipient=LastName


		    'Set objMail = CreateObject("CDONTS.Newmail")
		    'objMail.From = "system.monitor@logisticorp.us"
            varTo = PersonsEmail
		    varcc = "mark.maggiore@logisticorp.us"
		    varSubject = "New Address without Carillon ID"
		   ' objMail.MailFormat = cdoMailFormatMIME
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
    End if

                        'Response.write "l_cSQL="&l_cSQL&"<BR>"





	''Set Recordset1 = Server.CreateObject("ADODB.Recordset")
	''SQL_99="SELECT CompanyID FROM PreExistingCompanies WHERE (CompanyOwner='"&PhoneBookID&"') AND (CompanyName='"&RequestorCompany&"') AND (CompanyBuilding='"&RequestorBuilding&"') AND (CompanyAddress='"&RequestorAddress&"') AND (CompanySuite='"&RequestorSuite&"') AND (CompanyCity='"&RequestorCity&"') AND (CompanyZip='"&RequestorZipCode&"') AND (ContactName='"&ContactName&"') AND (CompanyPhone='"&RequestorPhone&"') AND (CompanyEmail='"&RequestorEmail&"') AND (CompanyStatus='c')"                                
	'Response.Write "SQL_99="&SQL_99&"<BR>"
	''Recordset1.ActiveConnection = DATABASE
	''Recordset1.Source = SQL_99
	''Recordset1.CursorType = 0
	''Recordset1.CursorLocation = 2
	''Recordset1.LockType = 1
	''Recordset1.Open()
	''Recordset1_numRows = 0
	''if NOT Recordset1.EOF then
		''CompanyID=Recordset1("CompanyID")
        'Response.write "CompanyID="&CompanyID&"<BR>"
	''End if
	''Recordset1.Close()
	''Set Recordset1 = Nothing
    
 	''Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
	''RSEVENTS.Open "PreExistingCompanies", Database, 2, 2
	''RSEVENTS.Find "CompanyID='"& CompanyID &"'"
		''RSEVENTS("st_id") = CompanyID
        ''RSEVENTS("st_alias") = CompanyID
	''RSEVENTS.update
	''RSEVENTS.close
	''set RSEVENTS = nothing   	


    


    
    	
	'If lcase(PageStatus)="mail" then
		Body = "ATTN:&nbsp;&nbsp;"&RequestorName &"<br><br>"   & _
		"Below are your user name and password for the Fleet Express Website.<br><br>"   & _
		"user name: "&RequestorEmail&"<br>"  & _
		"password: "&RequestorPassword&"<br><br>"   & _
		"The address is: "& WhichSite &" <br><br>"   & _ 			
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
  			ErrorMessage = "Congratulations!  You have successfully added a location to your address book.</a>"
            'Response.write "XXXRequestorZipCode="&RequestorZipCode&"<BR>"
            If iscourier<>"y" then
                ErrorMessage=ErrorMessage & "<br><br>FYI...This location is not eligible for courier delivery rates.<br>The negotiated lower rates are only for TI North and South Campus deliveries.<br>If this location is within that area, please notify us by <a href='mailto:mark.maggiore@logisticorp.us' class='FleetXRedMain'><b>clicking here</b></a>."
            End if
            RequestorName=""
            RequestorCompany=""
            RequestorBuilding=""
            RequestorAddress=""
            RequestorSuite=""
            RequestorCity=""
            RequestorZipCode=""
            RequestorPhone=""
            RequestorEmail=""
            RetypeRequestorEmail=""
            RequestorPassword=""
            CompanyCostCenter=""
            ContactName=""
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
<form action="FleetXAddressBook.asp?var1=<%=var1%>" method="post" name="FindUser">
<table border="0" bordercolor="green" Cellspacing="0" Cellpadding="0" align="left" width="100%">
 <tr><td align="left" colspan="2" class="FleetXRedSection"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>

        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetXRedSection" colspan="2"><%=PageTitle%></td></tr>

        <tr><td align="center" class="FleetXRedSection" colspan="2"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
        <!--tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressBodyWhite" colspan="2">In order to reach the Fleet Express Order page, please correctly type in the green verification code in the supplied text box and click "Submit."</td></tr-->
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
<tr><td>
<table  border="0" bordercolor="blue" align="center" class="MainPageText" width="100%">
	<tr height="40">
		<td width="650">&nbsp;</td>
	</tr>
    <tr><td align="center">



<form action="FleetXAddressBook.asp?varA=123&var1=<%=var1%>" method="post" name="FindUser">

<table  border="0" bordercolor="black" align="center" class="MainPageText">
    <%if trim(dontshow)<>"y" then %>
  <tr Height="30"> 
    <td colspan="3" valign="top" align="center" class="MainPageText"> 
      	<b>Complete this form and submit to add a location to your address book</b>
	
    </td>

  </tr>
  <tr><td colspan="3">To return to the Fleet Express Order Page, <a href="FreightOrder.asp?loggedin=y&varA=123&var1=<%=Var1%>" class="FleetXRedMain">Click Here</a>.</TD></tr>
<tr><td>&nbsp;</td></tr>


  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
      Company Name:
    </td>
	<td width="5">&nbsp;</td>
    <td width="136"> 
     <INPUT TYPE="text" NAME="requestorCompany" value="<%=requestorCompany%>" maxlength="50" size="30">
    </td>

  </tr>
  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
      Company Building:
    </td>
	<td width="5">&nbsp;</td>
    <td width="136"> 
     <INPUT TYPE="text" NAME="requestorBuilding" value="<%=requestorBuilding%>" maxlength="50" size="30">
    </td>

  </tr>
  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
      Company Address:
    </td>
	<td width="5">&nbsp;</td>
    <td width="136"> 
     <INPUT TYPE="text" NAME="requestorAddress" value="<%=requestorAddress%>" maxlength="50" size="30">
    </td>

  </tr>
  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
      Company Floor/Suite/Dock *:<br />
      <b>*If shipping freight, you MUST include a dock number</b>
    </td>
	<td width="5">&nbsp;</td>
    <td width="136"> 
     <INPUT TYPE="text" NAME="requestorSuite" value="<%=requestorSuite%>" maxlength="50" size="30">
    </td>

  </tr>
  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
        Company City:
    </td>
	<td width="5">&nbsp;</td>
    <td width="136"> 
     <INPUT TYPE="text" NAME="requestorCity" value="<%=requestorCity%>" maxlength="50" size="30">
    </td>

  </tr>
  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
        Company State:
    </td>
	<td width="5">&nbsp;</td>
    <td width="136"> 
     <INPUT TYPE="hidden" NAME="requestorState" value="TX">TX
    </td>

  </tr>
  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
        Company Zip Code:
    </td>
	<td width="5">&nbsp;</td>
    <td width="136"> 
     <INPUT TYPE="text" NAME="requestorZipCode" value="<%=requestorZipCode%>" maxlength="12" size="30">
    </td>

  </tr>
  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
        Contact Name:
    </td>
	<td width="5">&nbsp;</td>
    <td width="136"> 
     <INPUT TYPE="text" NAME="ContactName" value="<%=ContactName%>" maxlength="100" size="30">
    </td>

  </tr>
  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
        Cost Center (Optional):
    </td>
	<td width="5">&nbsp;</td>
    <td width="136"> 
     <INPUT TYPE="text" NAME="CompanyCostCenter" value="<%=CompanyCostCenter%>" maxlength="100" size="30">
    </td>

  </tr>
  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
        Contact Phone Number:
    </td>
	<td width="5">&nbsp;</td>
    <td width="136"> 
     <INPUT TYPE="text" NAME="requestorPhone" value="<%=requestorPhone%>" maxlength="20" size="30">
    </td>

  </tr>
  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
        Contact Email Address:
    </td>
	<td width="5">&nbsp;</td>
    <td width="136"> 
     <INPUT TYPE="text" NAME="requestorEmail" value="<%=requestorEmail%>" maxlength="100" size="30">
    </td>

  </tr>
  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
        Re-type Contact Email Address:
    </td>
	<td width="5">&nbsp;</td>
    <td width="136"> 
     <INPUT TYPE="text" NAME="retyperequestorEmail" value="<%=retyperequestorEmail%>" maxlength="100" size="30">
    </td>

  </tr>


  <tr Height="30"> 
    <td NOWRAP valign="middle" align="center" class="MainPageText" colspan="3"> 
	<input type="hidden" name="pagestatus" value="find">
    <input type="hidden" name="PhoneBookID" value="<%=PhoneBookID%>" />
    <input type="hidden" name="sortby" value="<%=sortby%>" />
      <input id="gobutton" type="submit" name="ButtonValue" value="Submit">
    </td>
  </tr>
  </form>
	<tr Height="1">
		<td>&nbsp;</td>
	</tr>
    <%end if %>
  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
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
   <form method="post">
	<tr><td align="left" class="MainPageText" colspan="4">
<b>Pre-Existing Locations</b> 
</td>
<td colspan="10" class="MainPageText"><b>Sort by:</b>  
                                <select name="sortby" ID="Select3"  onchange="form.submit()">
								<option value="CompanyName" <%if trim(sortby)="" then response.Write " selected" end if%>>Company Name</option>
								<option value="CompanyBuilding" <%if trim(sortby)="CompanyBuilding" then response.Write " selected" end if%>>Company Building</option>
                                <option value="CompanyAddress" <%if trim(sortby)="CompanyAddress" then response.Write " selected" end if%>>Company Address</option>
                                <option value="CompanySuite" <%if trim(sortby)="CompanySuite" then response.Write " selected" end if%>>Company Floor/Suite/Dock</option>
                                <option value="CompanyCity" <%if trim(sortby)="CompanyCity" then response.Write " selected" end if%>>Company City</option>
                                <option value="CompanyZip" <%if trim(sortby)="CompanyZip" then response.Write " selected" end if%>>Company Zip</option>
                                <option value="ContactName" <%if trim(sortby)="ContactName" then response.Write " selected" end if%>>Contact Name</option>
                                <option value="CompanyCostCenter" <%if trim(sortby)="CompanyCostCenter" then response.Write " selected" end if%>>Cost Center</option>
                                <option value="CompanyPhone" <%if trim(sortby)="CompanyPhone" then response.Write " selected" end if%>>Company Phone</option>
                                <option value="CompanyEmail" <%if trim(sortby)="CompanyEmail" then response.Write " selected" end if%>>Company Email</option>

								</select> 
</td>
</tr>
    </form>
<tr><td>&nbsp;</td></tr>
<%
							Set oConn = Server.CreateObject("ADODB.Connection")
							oConn.ConnectionTimeout = 100
							oConn.Provider = "MSDASQL"
							oConn.Open DATABASE
								l_cSQL2 = "select * from PreExistingCompanies  " &_
										"WHERE CompanyStatus = 'c' AND CompanyOwner = '" & TRIM(PHoneBookID)&"' order by "& SortBy &" "  
										'if trim(displayusername)="comps" or trim(displayusername)="Compugraphics"  then 
										'l_cSQL2 = l_cSQL2 & "  AND st_id<>'CPGP'" 
										'end if
                                'Response.Write "userid="&userid&"<BR>"
								'Response.Write "l_cSQL2="&l_cSQL2&"<BR>"
								SET oRs = oConn.Execute(l_cSql2)

								Do While not oRs.EOF
                                DisplayCompanyID=oRs("CompanyID")
                                DisplayCompanyName=oRs("CompanyName")
                                DisplayCompanyBuilding=oRs("CompanyBuilding")
                                DisplayCompanyAddress=oRs("CompanyAddress")
                                DisplayCompanySuite=oRs("CompanySuite")
                                DisplayCompanyCity=oRs("CompanyCity")
                                DisplayCompanyState=oRs("CompanyState")
                                DisplayCompanyZip=oRs("CompanyZip")
                                DisplayCompanyCostCenter=oRs("CompanyCostCenter")
                                DisplayContactName=oRs("ContactName")
                                DisplayCompanyPhone=oRs("CompanyPhone")
                                DisplayCompanyEmail=oRs("CompanyEmail")
                                xyz=xyz+1

%>

 <form method="post">
 <%If xyz=1 then %>
    <tr>
    <td width="<%=spacersize%>"><img="image/pixel.gif" height="1" width="1" /></td>
    <td width="<%=spacersize%>"><img="image/pixel.gif" height="1" width="1" /></td>
    <td width="<%=spacersize%>"><img="image/pixel.gif" height="1" width="1" /></td>
    <td width="<%=spacersize%>"><img="image/pixel.gif" height="1" width="1" /></td>
    <td nowrap=nowrap><b>Company Name</b></td>
    <td width="<%=spacersize%>"><img="image/pixel.gif" height="1" width="1" /></td>
    <td nowrap=nowrap><b>Building</b></td>
    <td width="<%=spacersize%>"><img="image/pixel.gif" height="1" width="1" /></td>
    <td nowrap=nowrap><b>Address</b></td>
    <td width="<%=spacersize%>"><img="image/pixel.gif" height="1" width="1" /></td>
    <td nowrap=nowrap><b>Floor/Suite/Dock</b></td>
    <td width="<%=spacersize%>"><img="image/pixel.gif" height="1" width="1" /></td>
    <td nowrap=nowrap><b>City</b></td>
    <td width="<%=spacersize%>"><img="image/pixel.gif" height="1" width="1" /></td>
    <td nowrap=nowrap><b>Zip Code</b></td>
    <td width="<%=spacersize%>"><img="image/pixel.gif" height="1" width="1" /></td>
    <td nowrap=nowrap><b>Contact Name</b></td>
    <td width="<%=spacersize%>"><img="image/pixel.gif" height="1" width="1" /></td>
    <td nowrap=nowrap><b>Cost Center</b></td>
    <td width="<%=spacersize%>"><img="image/pixel.gif" height="1" width="1" /></td>
    <td nowrap=nowrap><b>Phone</b></td>
    <td width="<%=spacersize%>"><img="image/pixel.gif" height="1" width="1" /></td>
    <td nowrap=nowrap><b>email address</b></td></tr>
    <td width="<%=spacersize%>"><img="image/pixel.gif" height="1" width="1" /></td>
 <%End if %>
<tr>
    <td align="center"><input id="gobutton" type="submit" name="submit1" value="Edit" /></td>
    <td width="5"><img="image/pixel.gif" height="1" width="1" /></td>
    <td align="center"><input id="gobutton" type="submit" value="Delete" /></td>
    <td width="<%=spacersize%>"><img="image/pixel.gif" height="1" width="1" /></td>
    <td class="MainPageTextAddressBook"><%=left(DisplayCompanyName, 30) %></td>
    <td width="<%=spacersize%>"><img="image/pixel.gif" height="1" width="1" /></td>
    <td class="MainPageTextAddressBook"><%=DisplayCompanyBuilding %></td>
    <td width="<%=spacersize%>"><img="image/pixel.gif" height="1" width="1" /></td>
    <td class="MainPageTextAddressBook"><%=DisplayCompanyAddress %></td>
    <td width="<%=spacersize%>"><img="image/pixel.gif" height="1" width="1" /></td>
    <td class="MainPageTextAddressBook"><%=DisplayCompanySuite %></td>
    <td width="<%=spacersize%>"><img="image/pixel.gif" height="1" width="1" /></td>
    <td class="MainPageTextAddressBook"><%=DisplayCompanyCity %></td>
    <td width="<%=spacersize%>"><img="image/pixel.gif" height="1" width="1" /></td>
    <td class="MainPageTextAddressBook"><%=DisplayCompanyZip %></td>
    <td width="<%=spacersize%>"><img="image/pixel.gif" height="1" width="1" /></td>
    <td class="MainPageTextAddressBook"><%=DisplayContactName %></td>
    <td width="<%=spacersize%>"><img="image/pixel.gif" height="1" width="1" /></td>
    <td class="MainPageTextAddressBook"><%=DisplayCompanyCostCenter %></td>
    <td width="<%=spacersize%>"><img="image/pixel.gif" height="1" width="1" /></td>
    <td class="MainPageTextAddressBook"><%=DisplayCompanyPhone %></td>
    <td width="<%=spacersize%>"><img="image/pixel.gif" height="1" width="1" /></td>
    <td class="MainPageTextAddressBook"><%=DisplayCompanyEmail %></td>
    <td width="<%=spacersize%>"><img="image/pixel.gif" height="1" width="1" /></td>
</tr>
<tr><td colspan="11"><hr /></td></tr>
<input type="hidden" name="UpdateCompanyID" value="<%=DisplaycompanyID %>" />
<input type="hidden" name="updatestatus" value="y" />
<input type="hidden" name="PhoneBookID" value="<%=PhoneBookID %>" />
<input type="hidden" name="sortby" value="<%=sortby%>" />
</form>
<%

								oRs.movenext
								LOOP
							Set oConn=Nothing
 %>
</table>

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
