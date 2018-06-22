<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <%
    '''''''''''HARDCODED STUFF
   sBT_ID="88"
   Session("sBT_ID")=sBT_ID 
    fleetexpresscourier="y"
    RequestorState="TX"     
     %>
    <title>Fleet Express Courier Order Page</title>
    <link rel="stylesheet" type="text/css" href="../MainStyleSheet.css">
    <!-- #include file="../settings.inc" -->
    <!-- #include file="../../dedicatedfleets/include/ifabsettings.inc" -->
    <!-- #include file="../../dedicatedfleets/include/checkstring.inc" -->
   <%
   TempPreExistingOrigination=Request.Form("TempPreExistingOrigination")
   TempPreExistingDestination=Request.Form("TempPreExistingDestination")

    ColorSelect=Request.form("ColorSelect")
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

   ''''''''''''''''''''''''''
   submitbutton=Request.Form("submitbutton")
   'Response.write "submitbutton="&submitbutton&"<BR>"
   FXCourieruserid=Request.form("FXCourierUserID")

''''''''CHECKS TO SEE IF A SUPERVISOR
Set Recordset1 = Server.CreateObject("ADODB.Recordset")
'Response.write "Database="&Database&"<br>"
Recordset1.ActiveConnection = Database
SQL="SELECT * FROM UserList WHERE (USERID='"&FXCourierUserID&"') AND (UserStatus='c')"
 Recordset1.Source = SQL
'response.write "SQL="& SQL &"<BR>"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()
Recordset1_numRows = 0
'response.write "<font color='red'>SQL1="&Recordset1.Source&"<BR></font>"
	if NOT Recordset1.EOF then
        Supervisor=Recordset1("Supervisor")
        DispUserFirstName=Recordset1("UserFirstName")
        else
        Supervisor="n"
        PreExistingRequestor=FXCourierUserID
	End if
Recordset1.Close()
Set Recordset1 = Nothing
'response.write   "Supervisor="&Supervisor&"<BR>"
'response.write   "DispUserFirstName="&DispUserFirstName&"<BR>"
'response.write   "FXCourieruserid="&FXCourieruserid&"<BR>"



   CelCarrier=Request.Form("CelCarrier")
   TempRequestorPhoneCarrier=CelCarrier
   If trim(UserID)="" and xxx="xxx" then
        'REsponse.write "GOT HERE2!!!<br>"
        UserID=session("FXCourierUserID")
    End if
   LogInVerified=Request.form("LogInVerified")
   'response.write "FXCourierUserID="&FXCourierUserID&"<BR>"
    MarkTemp=Request.Form("MarkTemp")
    OrderSubmitted=Request.Form("OrderSubmitted")



    ''''''''REMOVED THIS FOR CAPTCHA....MIGHT NEED TO PUT IT BACK!
    'If trim(UserID)="" and trim(MarkTemp)=""  then
    '    Response.redirect("http://www.logisticorp.us/intranet")
    '    else
    '    MarkTemp="yes"
    'End if
    CaptchaSubmit=Request.form("CaptchaSubmit")
    varCaptcha=Request.form("varCaptcha")
    If CaptchaSubmit<>varCaptcha then
        ErrorMessage="You did not supply the correct verification code"
    End if
    'Response.write "CaptchaSubmit="&CaptchaSubmit&"<BR>"
    'Response.write "varCaptcha="&varCaptcha&"<BR>"
    'Response.write "UserID="&UserID&"<BR>"
    If  trim(FXCourierUserID)>"" then
    %>

<script language="javascript" type="text/javascript" src="../datetimepicker.js">

    //Date Time Picker script- by TengYong Ng of http://www.rainforestnet.com
    //Script featured on JavaScript Kit (http://www.javascriptkit.com)
    //For this script, visit http://www.javascriptkit.com
    function onPageLoad() {
        if (document.OrderForm1.RequestorName.value.length == 0) {
            document.OrderForm1.RequestorName.focus();
        }
    }
</script>

<script type="text/javascript">
    function AnyCost() {
        if (document.OrderForm1.Priority.value == "Time Critical") {
            alert("Warning:  If you select 'Time Critical' as the service level, then surcharges will be incurred and charged to the requestor's cost center or the provided charge number \n  \n Charging approval is indicated by the selection of this option");
        }
    }
</script>


    <%
    Today=Now()
    TodayHour=Hour(Now())
    If TodayHour>=0 and TodayHour<=7 then
        DeliveryType="t"
        else
        DeliveryType="d"
    End if

    ''''''''CHECKS TO SEE IF A SUPERVISOR
    Set Recordset1 = Server.CreateObject("ADODB.Recordset")
    'Response.write "Database="&Database&"<br>"
    Recordset1.ActiveConnection = Database

    Recordset1.Source = "SELECT * FROM UserList WHERE (USERID='"&FXCourierUserID&"') AND (UserStatus='c')"
    'response.write "Recordset1.Source="& Recordset1.Source &"<BR>"
    Recordset1.CursorType = 0
    Recordset1.CursorLocation = 2
    Recordset1.LockType = 1
    Recordset1.Open()
    Recordset1_numRows = 0
    'response.write "<font color='red'>SQL1="&Recordset1.Source&"<BR></font>"
	    if NOT Recordset1.EOF then
            Supervisor=Recordset1("Supervisor")
            else
            Supervisor="n"
	    End if
    Recordset1.Close()
    Set Recordset1 = Nothing


    'Response.write "Supervisor="&Supervisor&"<BR>"


    'Response.write "today="&today&"<BR>"
    'Response.write "todayhour="&todayhour&"<BR>"
    'Response.write "DeliveryType="&DeliveryType&"<BR>"

    timesthrough=Request.form("timesthrough")
    
    TableWidth="460"
    Internal=Request.QueryString("Internal")
    
    PreExistingRequestor=Request.Form("PreExistingRequestor")
    PreExistingOrigination=Request.Form("PreExistingOrigination")
    PreExistingDestination=Request.Form("PreExistingDestination")
    'Response.write "xxxPreExistingOrigination="&PreExistingOrigination&"<BR>"
    'Response.write "PreExistingRequestor="&PreExistingRequestor&"<BR>"
    'Response.write "PreExistingDestination="&PreExistingDestination&"<BR>"
    If trim(PreExistingOrigination)="" OR trim(PreExistingDestination)="" OR trim(PreExistingRequestor)="" then
        'PageStatus="x"
        PageStatus=Request.form("PageStatus")
        else
        PageStatus=Request.form("PageStatus")
    End if
    'Response.write "***PageStatus="&PageStatus&"<BR>"
    RequestorFirstName=Request.form("RequestorFirstName")
    RequestorLastName=Request.form("RequestorLastName")
    RequestorPhoneNumber=Request.form("RequestorPhoneNumber")
    RequestorEmailAddress=Request.form("RequestorEmailAddress")

    RequestorAddress=Request.form("RequestorAddress")
    RequestorCity=Request.form("RequestorCity")
    RequestorZipCode=Request.form("RequestorZipCode")
    RequestorSTate=Request.form("RequestorSTate")
    RequestorPhoneCarrier=Request.form("CelCarrier")
    RequestorName=Request.form("RequestorName")
    RequestorPassword=trim(Request.form("RequestorPassword"))
    RequestorDeliveryNotifications=Request.form("RequestorDeliveryNotifications")
    DeliveryNotifications=Request.form("DeliveryNotifications")
    TempRequestorDeliveryNotifications=DeliveryNotifications
    'REsponse.write "RequestorDeliveryNotifications="&RequestorDeliveryNotifications&"<BR>"
    'REsponse.write "DeliveryNotifications="&DeliveryNotifications&"<BR>"
    'REsponse.write "TempRequestorDeliveryNotifications="&TempRequestorDeliveryNotifications&"<BR>"

    'PONumber=Request.form("PONumber")
    'CostCenterNumber=Request.form("CostCenterNumber")
    Pieces=Request.form("Pieces")
    NumberOfPallets=Request.form("NumberOfPallets")
    DimWeight=Request.form("DimWeight")
    DimLength=Request.form("DimLength")
    DimWidth=Request.form("DimWidth")
    DimHeight=Request.form("DimHeight")
    DimValue=Request.form("DimValue")
    IsHazmat=Request.form("IsHazmat")
    OriginationCompany=Request.form("OriginationCompany")
    OriginationAddress=Request.form("OriginationAddress")
    OriginationBuilding=Request.form("OriginationBuilding")
    OriginationSuite=Request.form("OriginationSuite")
    OriginationCity=Request.form("OriginationCity")
    OriginationState=Request.form("OriginationState")
    'Response.write "***OriginationState="&OriginationState&"<BR>"
    OriginationZipCode=Request.form("OriginationZipCode")
    OriginationContactName=Request.form("OriginationContactName")
    OriginationPhoneNumber=Request.form("OriginationPhoneNumber")
    OriginationEmail=Request.form("OriginationEmail")
    DestinationCompany=Request.form("DestinationCompany")
    DestinationAddress=Request.form("DestinationAddress")
    DestinationBuilding=Request.form("DestinationBuilding")
    DestinationSuite=Request.form("DestinationSuite")
    DestinationCity=Request.form("DestinationCity")
    DestinationState=Request.form("DestinationState")
    DestinationZipCode=Request.form("DestinationZipCode")
    DestinationContactName=Request.form("DestinationContactName")
    DestinationPhoneNumber=Request.form("DestinationPhoneNumber")
    DestinationEmail=Request.form("DestinationEmail")

    POorNWA=Request.form("POorNWA")
    GenericNumber=Request.form("GenericNumber")
    BasicCharge=Request.Form("BasicCharge")
    PartNumber=trim(Request.Form("PartNumber"))
    PartNumber=Replace(Partnumber,"*","")
    PartNumber=Replace(Partnumber," ","")
    PartNumber=Replace(Partnumber,"""","")
    PartNumber=Replace(Partnumber,"'","")
    Select Case POorNWA
        Case "P/O #"
            PONumber=GenericNumber
        Case "Cost Center #"
            CostCenterNumber=GenericNumber
    End Select
    Comments=Request.form("Comments")
    RequestorName=Replace(RequestorName, """", "`")
    RequestorName=Replace(RequestorName, "'", "`")
    RequestorPhoneNumber=Replace(RequestorPhoneNumber, """", "")
    RequestorPhoneNumber=Replace(RequestorPhoneNumber, "'", "")
    RequestorEmailAddress=Replace(RequestorEmailAddress, """", "")
    RequestorEmailAddress=Replace(RequestorEmailAddress, "'", "")
    Pieces=Replace(Pieces, """", "")
    Pieces=Replace(Pieces, "'", "")
    rf_box=Request.Form("rf_box")
    DimWeight=Replace(DimWeight, """", "")
    DimWeight=Replace(DimWeight, "'", "")
    DimLength=Replace(DimLength, """", "")
    DimLength=Replace(DimLength, "'", "")
    DimWidth=Replace(DimWidth, """", "")
    DimWidth=Replace(DimWidth, "'", "")
    DimHeight=Replace(DimHeight, """", "")
    DimHeight=Replace(DimHeight, "'", "")
    OriginationCompany=Replace(OriginationCompany, """", "`")
    OriginationCompany=Replace(OriginationCompany, "'", "`")
    OriginationAddress=Replace(OriginationAddress, """", "`")
    OriginationAddress=Replace(OriginationAddress, "'", "`")
    OriginationSuite=Replace(OriginationSuite, """", "`")
    OriginationSuite=Replace(OriginationSuite, "'", "`")
    OriginationCity=Replace(OriginationCity, """", "`")
    OriginationCity=Replace(OriginationCity, "'", "`")
    OriginationZipCode=Replace(OriginationZipCode, """", "")
    OriginationZipCode=Replace(OriginationZipCode, "'", "")
    OriginationContactName=Replace(OriginationContactName, """", "`")
    OriginationContactName=Replace(OriginationContactName, "'", "`")
    OriginationPhoneNumber=Replace(OriginationPhoneNumber, """", "")
    OriginationPhoneNumber=Replace(OriginationPhoneNumber, "'", "")
    OriginationEmail=Replace(OriginationEmail, """", "")
    OriginationEmail=Replace(OriginationEmail, "'", "")



    DestinationCompany=Replace(DestinationCompany, """", "`")
    DestinationCompany=Replace(DestinationCompany, "'", "`")
    DestinationAddress=Replace(DestinationAddress, """", "`")
    DestinationAddress=Replace(DestinationAddress, "'", "`")
    DestinationSuite=Replace(DestinationSuite, """", "`")
    DestinationSuite=Replace(DestinationSuite, "'", "`")
    DestinationCity=Replace(DestinationCity, """", "`")
    DestinationCity=Replace(DestinationCity, "'", "`")
    DestinationZipCode=Replace(DestinationZipCode, """", "")
    DestinationZipCode=Replace(DestinationZipCode, "'", "")
    DestinationContactName=Replace(DestinationContactName, """", "`")
    DestinationContactName=Replace(DestinationContactName, "'", "`")
    DestinationPhoneNumber=Replace(DestinationPhoneNumber, """", "")
    DestinationPhoneNumber=Replace(DestinationPhoneNumber, "'", "")
    DestinationEmail=Replace(DestinationEmail, """", "")
    DestinationEmail=Replace(DestinationEmail, "'", "")


    Comments=Replace(Comments, """", "`")
    Comments=Replace(Comments, "'", "`")
    Refrigerate=Request.form("Refrigerate")
    Priority=Request.form("Priority")
    PickUpDateTime=Request.form("PickUpDateTime")
    If trim(PickUpDateTime)="" then
        PickUpDateTime=now()+.0034722
    End if
    DeliveryDateTime=Request.form("DeliveryDateTime")
     If trim(DeliveryDateTime)="" then
        DeliveryDateTime=DateAdd("n",180,PickUpDateTime)
    End if   


    ''''''''PRE-POPULATES ITEMS FOR LOGISTICORP PEEPS
    If Internal="y" and TRIM(timesthrough)="" then
        'Response.write "Got here!<br>"
        'RequestorEmailAddress="FleetX@LogisticorpGroup.com"
        RequestorEmailAddress="FleetXDFW@logisticorp.us"
        'RequestorEmailAddress="mark.maggiore@logisticorp.us"
        OriginationContactName="Dispatch"
        OriginationPhoneNumber="972-499-3415"
        'OriginationEmail="FleetX@LogisticorpGroup.com"
        OriginationEmail="FleetXDFW@logisticorp.us"
        DestinationContactName="Dispatch"
        DestinationPhoneNumber="972-499-3415"
        'DestinationEmail="FleetX@LogisticorpGroup.com"
        DestinationEmail="FleetXDFW@logisticorp.us"
    End if
















If submitbutton="Submit Order" then
        'If trim(DestinationEmail)="" then
        '    ErrorMessage="You must provide the Destination's Email"
       ' End if
       'If not isdate(DeliveryDateTime) then DeliveryDateTime=now() end if
        'If not isdate(PickUPDateTime) then PickUPDateTime=now() end if
        If cdate(PickUpDateTime)<now() then PickUpDateTime=Now() end if
       ' Response.write "PickUpDateTime="&PickUpDateTime&"<BR>"
       ' Response.write "CurrentDateTime="&CurrentDateTime&"<BR>"
        If trim(PreExistingOrigination)=trim(PreExistingDestination) and trim(PreExistingOrigination)>"" and trim(PreExistingDestination)>"" then
            ErrorMessage="Your Origination and Destination cannot be the same"
        End if
        If trim(RequestorFirstName)="" then
            ErrorMessage="You must provide the Requestor's First Name"
        End if
        If trim(RequestorLastName)="" then
            ErrorMessage="You must provide the Requestor's Last Name"
        End if
        If trim(RequestorAddress)="" then
            ErrorMessage="You must provide the Requestor's Address"
        End if
        If trim(RequestorCity)="" then
            ErrorMessage="You must provide the Requestor's City"
        End if
        If trim(RequestorZipCode)="" then
            ErrorMessage="You must provide the Requestor's Zip Code"
        End if
        If trim(RequestorPhoneNumber)="" then
            ErrorMessage="You must provide the Requestor's Phone Number"
        End if
        If trim(celcarrier)="" then
            ErrorMessage="You must provide the Requestor's Phone Carrier Service"
        End if


        If trim(RequestorEmailAddress)="" then
            ErrorMessage="You must provide the Requestor's Email Address"
        End if
        'If trim(UserName)="" then
        '    ErrorMessage="You must provide the Requestor's User Name"
        'End if
        If trim(RequestorPassword)="" then
            ErrorMessage="You must provide the Requestor's Password"
        End if
        If trim(DeliveryNotifications)="" then
            ErrorMessage="You must provide the Requestor's Delivery Notifications Preference"
        End if

        If trim(PartNumber)="" then
            ErrorMessage="You must provide the Control/Part Number"
        End if

        If trim(CostCenterNumber)="" AND trim(PONumber)="" then
            ErrorMessage="You must provide the Cost Center Number or P/O Number"
        End if

        If trim(Comments)="" then
            ErrorMessage="You have not provided any Special Instructions for this delivery.  If there are no special instructions, please type in N/A"
        End if
        If trim(Pieces)="" then
            ErrorMessage="You must provide the Number of Pieces"
        End if
        If trim(DimWeight)="" then
            ErrorMessage="You must provide the Commodity's Weight"
        End if
        'If trim(NumberOfPallets)="" and isPalletized="y" then
        '    ErrorMessage="You must provide the Number of Pallets"
        'End if
        If trim(DimLength)="" then
            ErrorMessage="You must provide the Commodity's Length"
        End if
        If trim(DimWidth)="" then
            ErrorMessage="You must provide the Commodity's Width"
        End if
        CubicTotal=int(Dimlength)+Int(DimWidth)+Int(DimHeight)
        If CubicTotal>50 then
            ErrorMessage="The total dimension (L + W + H) of your item cannot exceed 50 inches.<br>Your item is "& CubicTotal &" inches."
        End if
        '''If trim(OriginationCompany)="" then
        '''    ErrorMessage="You must provide the Origination's Company"
        '''End if
        '''If trim(OriginationAddress)="" then
        '''    ErrorMessage="You must provide the Origination's Address"
        '''End if
        '''If trim(OriginationCity)="" then
        '''    ErrorMessage="You must provide the Origination's City"
        '''End if
        '''If trim(OriginationZipCode)="" then
        '''    ErrorMessage="You must provide the Origination's Zip Code"
        '''End if
        If trim(OriginationBuilding)="" then
            ErrorMessage="You must provide the Origination's Building"
        End if
        If trim(OriginationSuite)="" then
            ErrorMessage="You must provide the Origination's Floor/Cube Number"
        End if
        If trim(OriginationContactName)="" then
            ErrorMessage="You must provide the Origination's Contact Name"
        End if
        If trim(OriginationPhoneNumber)="" then
            ErrorMessage="You must provide the Origination's Phone Number"
        End if
        If NOT isdate(PickUpDateTime) then
            ErrorMessage="You must provide a valid ready date/time"
        End if
        'If trim(OriginationEmail)="" then
        '    ErrorMessage="You must provide the Origination's Email"
        'End if
        If isdate(PickUpDateTime)  and cdate(CurrentDateTime)>cdate(PickUpDateTime) then
            ErrorMessage="The ready time cannot be before the current date/time"
        End if
        '''If trim(DestinationCompany)="" then
        '''    ErrorMessage="You must provide the Destination's Company"
        '''End if
        '''If trim(DestinationAddress)="" then
        '''    ErrorMessage="You must provide the Destination's Address"
        '''End if
        '''If trim(DestinationCity)="" then
        '''    ErrorMessage="You must provide the Destination's City"
        '''End if
        '''If trim(DestinationZipCode)="" then
        '''    ErrorMessage="You must provide the Destination's Zip Code"
        '''End if
        If trim(DestinationBuilding)="" then
            ErrorMessage="You must provide the Destination's Building"
        End if
        If trim(DestinationSuite)="" then
            ErrorMessage="You must provide the Destination's Floor/Cube Number"
        End if

        If trim(DestinationContactName)="" then
            ErrorMessage="You must provide the Destination's Contact Name"
        End if
        If trim(DestinationPhoneNumber)="" then
        ErrorMessage="You must provide the Destination's Phone Number"
        End if
        If DateDiff("n", PickUpDateTime, DeliveryDateTime)<179 then
            ErrorMessage="There must be a minimum of 3 hours difference between the pick up and delivery time"
        End if
        'Response.write "DeliveryDateTime="&DeliveryDateTime&"<BR>"
        If NOT isdate(DeliveryDateTime) then
            ErrorMessage="You must provide a valid destination date/time"
        End if
        'Response.write "PickUpDateTime="&PickUpDateTime&"<BR>"
        'Response.write "DeliveryDateTime="&DeliveryDateTime&"<BR>"
        'Response.write DateDiff("n", PickUpDateTime, DeliveryDateTime)
        'Response.write DateDiff("n", PickUpDateTime, DeliveryDateTime)
        If isdate(DeliveryDateTime) and isdate(PickUPDateTime) and cdate(PickUPDateTime)>=cdate(DeliveryDateTime) then
            ErrorMessage="The delivery date/time cannot be before the ready date/time"
        End if
        If isdate(DeliveryDateTime)  and cdate(CurrentDateTime)>=cdate(DeliveryDateTime) then
            ErrorMessage="The delivery date/time cannot be before the current date/time"
        End if
End if

'Response.write "ErrorMessage="&ErrorMessage&"****<BR>"
'Response.write "submitbutton="&submitbutton&"****<BR>"
'Response.write "PreexistingRequestor="&PreexistingRequestor&"****<BR>"
If trim(ErrorMessage)="" and submitbutton="Submit Order" and trim(PreexistingRequestor)="" then
            loginname=lcase(left(RequestorFirstName,1)&Requestorlastname)
	        LoginLength=Len(LoginName)
	        LoginLengthPlusOne=LoginLength+1
	        status="c"
	        Set RS2 = Server.CreateObject("ADODB.Recordset")
		        RS2.CursorLocation = 3
		        RS2.CursorType = 3
                'Response.write "Database="&Database&"<BR>"
		        RS2.ActiveConnection = Database
		        SQL = "SELECT * FROM UserList where (UserName='"&LoginName&"') or ((Left(UserName,"&LoginLength&") = '"&LoginName&"') and (right(UserName,1)<'a')) ORDER BY UserName desc"
		        'Response.write "Database="&Database&"<BR>"
                'Response.write "SQL="&SQL&"<BR>"		        
                RS2.Open SQL, Database, 1, 3
		        
                'Response.write "SQL="&SQL&"<BR>"
                If not RS2.EOF then
			        PreUserName=RS2("UserName")
			        PreFXCourierUserID=RS2("UserID")
			        TempLoginName=RS2("UserName")
			        TempLoginLength=Len(TempLoginName)
			        If TempLoginName=LogInName then
				        LogInName=LogInName&"02"
				        else
				        AddToLogin=Right(TempLoginName,(TempLoginLength-LoginLength))
				        AddToLogin=AddToLogin+1
				        If AddToLogin<10 then
					        LoginName=LoginName&"0"&AddToLogin
					        else
					        LoginName=LoginName&AddToLogin
				        end if
			        end if
		        end if
		        RS2.close
	        Set RS2 = Nothing
	        PreFXCourierUserID=PreFXCourierUserID+1		
	        Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
		        RSEVENTS.Open "UserList", Database, 2, 2
		        RSEVENTS.addnew	
		        'response.Write "UserName="&UserName&"<BR>"	
		        'response.Write "LoginName="&LoginName&"<BR>"				
		        'RSEVENTS("UserName") = 	UserName
		        RSEVENTS("UserFirstName") = RequestorFirstName
		        RSEVENTS("UserLastName") = Requestorlastname
                RSEVENTS("UserAddress") = RequestorAddress
                RSEVENTS("UserCity") = RequestorCity
                RSEVENTS("UserState") = RequestorState
                RSEVENTS("UserZipCode") = RequestorZipCode
                RSEVENTS("UserPhoneNumber") = RequestorPhoneNumber
                RSEVENTS("UserPhoneCarrier") = CelCarrier
                RSEVENTS("UserEmailAddress") = RequestorEmailAddress
                RSEVENTS("UserName") = LoginName
		        RSEVENTS("Password") = RequestorPassword
		        RSEVENTS("DeliveryNotifications") = DeliveryNotifications
		        RSEVENTS("Userstatus") = "c"
		        RSEVENTS("DateCreated") = Date()
		        RSEVENTS.update
		        RSEVENTS.close			
	        set RSEVENTS = nothing
	        Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
		        RSEVENTS.CursorLocation = 3
		        RSEVENTS.CursorType = 3
		        RSEVENTS.ActiveConnection = Database
		        SQL = "SELECT UserID FROM UserList ORDER BY UserID desc"
		        RSEVENTS.Open SQL, Database, 1, 3
		        PreFXCourierUserID=RSEVENTS("UserID")
		        RSEVENTS.close
	        Set RSEVENTS = Nothing	
	        Body = "Greetings,<br><br>"   & _
			        "Below are your user name and password for the Fleet Express Courier Website.<br><br>"& _
			        "user name: "&LoginName & "<br>" & _
			        "password: "& RequestorPassword &"<br><br>"& _
			        "The address is: http://www.FleetXDFW.com/index.html <br><br>"& _ 	
                    "Next time, avoid the phone.  You can log in and place your order online. <br><br>"& _ 			
			        "If you have any questions, please do not hesitate to contact me.<br><br>"& _
			        "Thank you,<br><br>" & _
			        "Fleet Express Courier<br>"  & _
			        "FleetX@LogisticorpGroup.com<br>"  & _ 
			        "(214) 882-0620<br><br>"
	        'Recipient = "mark.maggiore@logisticorp.us"
		        Set objMail = CreateObject("CDONTS.Newmail")
		        objMail.From = "FleetX@LogisticorpGroup.com"
		        objMail.To = RequestorEmailAddress
                objMail.cc = "mark.maggiore@logisticorp.us"
		        objMail.Subject = "Welcome Fleet Express Courier New User"
		        objMail.MailFormat = cdoMailFormatMIME
		        objMail.BodyFormat = cdoBodyFormatHTML
		        objMail.Body = Body
		        objMail.Send
		        Set objMail = Nothing

	        Body = "Greetings,<br><br>"   & _
			        "A new user has registered on the Fleet Express Courier Website.  Below is the basic info:<br><br>"& _
			        "user name: "& LoginName & "<br>" & _
			        "password: *******<br><br>"& _
                    "first name: " &RequestorFirstName & "<br>" & _
                    "last name: "& Requestorlastname & "<br>" & _
                    "address: "& RequestorAddress & "<br>" & _
                    "city: "& RequestorCity & "<br>" & _
                    "zip code: "& RequestorZipCode & "<br>" & _
                    "email address: "& RequestorEmailAddress & "<br>" & _
                    "phone number: "& RequestorPhoneNumber & "<br>" & _
    		        "If you have any questions, please do not hesitate to contact me.<br><br>"& _
			        "Thank you,<br><br>" & _
			        "Fleet Express Courier<br>"  & _
			        "FleetX@LogisticorpGroup.com<br>"  & _ 
			        "(214) 882-0620<br><br>"
	        'Recipient = "mark.maggiore@logisticorp.us"
		        Set objMail = CreateObject("CDONTS.Newmail")
		        objMail.From = "FleetX@LogisticorpGroup.com"
		        objMail.To = "FleetX@LogisticorpGroup.com"
                objMail.To = "FleetXDFW@logisticorp.us"
                objMail.cc = "mark.maggiore@logisticorp.us"
		        objMail.Subject = "New Fleet Express Courier User"
		        objMail.MailFormat = cdoMailFormatMIME
		        objMail.BodyFormat = cdoBodyFormatHTML
		        objMail.Body = Body
		        objMail.Send
		        Set objMail = Nothing

	            Set RS2 = Server.CreateObject("ADODB.Recordset")
		            RS2.CursorLocation = 3
		            RS2.CursorType = 3
                    'Response.write "Database="&Database&"<BR>"
		            RS2.ActiveConnection = Database
		            SQL = "SELECT * FROM UserList where (UserName='"&LoginName&"') and UserStatus='c' ORDER BY UserName desc"
		            'Response.write "Database="&Database&"<BR>"
                    'Response.write "SQL="&SQL&"<BR>"		        
                    RS2.Open SQL, Database, 1, 3
		        
                    'Response.write "SQL="&SQL&"<BR>"
                    If not RS2.EOF then
			            PreExistingRequestor=RS2("UserID")
		            end if
		            RS2.close
	            Set RS2 = Nothing




        End if





 
        'If trim(ErrorMessage)="" and submitbutton="Submit Order" and trim(PreexistingRequestor)>"" and trim(OriginationCompany)>"" and trim(DestinationCompany)>""  then
        'If trim(ErrorMessage)="" and submitbutton="Submit Order" and trim(PreexistingRequestor)>"" and trim(PreExistingOrigination)>"" and trim(PreExistingDestination)>""  then
        If trim(ErrorMessage)="" and submitbutton="Submit Order" and trim(PreexistingRequestor)>"" then
              
        'response.write "GOT HERE!!!<BR>"
		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.ConnectionTimeout = 100
		oConn.Provider = "MSDASQL"
		oConn.Open DATABASE
			l_cSQL = "Select * FROM PreExistingCompanies WHERE companystatus='c' and CompanyEmail='"& OriginationEmail &"'  and CompanyPhone='"& OriginationPhoneNumber &"'  and ContactName='"& OriginationContactName &"'  and CompanyZip='"& OriginationZipCode &"'  and CompanyCity='"& OriginationCity &"' and  CompanyAddress='"& Originationaddress &"'  and CompanyName='"& OriginationCompany &"' and CompanyBuilding='"& OriginationBuilding &"' and CompanySuite='"& OriginationSuite &"'"
            If Supervisor<>"y" or isNULL(Supervisor) then
                l_cSQL = l_cSQL & " AND CompanyOwner='"& PreExistingRequestor &"'"
                else
                'l_cSQL = l_cSQL & " AND CompanyOwner is NULL "
                l_cSQL = l_cSQL & " AND CompanyOwner='"& PreExistingRequestor &"'"
            End if
            'response.write "OriginationXXXXXl_cSql="&l_cSql&"<BR>"
			SET oRs = oConn.Execute(l_cSql)
					if oRs.EOF then
                        'Response.write "ADD a NEW ORIGINATION!!!!<BR>"
''''''''''''''''''''''''''''''''ADDED THIS
                        strSQL = "INSERT INTO PreExistingCompanies(CompanyName, CompanyAddress, CompanyBuilding, CompanySuite, CompanyCity, CompanyState, CompanyZip, ContactName, CompanyPhone, CompanyEmail, CompanyOwner, CompanyStatus) VALUES('"& OriginationCompany &"', '"& OriginationAddress &"', '"& OriginationBuilding &"', '"& OriginationSuite &"', '"& OriginationCity &"', '"& OriginationState &"', '"& OriginationZipCode &"', '"& OriginationContactName &"', '"& OriginationPhoneNumber &"', '"& OriginationEmail &"','"& PreExistingRequestor &"', 'c')"
                        'response.write "strSQL="&strSQL&"<BR>"
                         strConnection = DATABASE
 
                        Set commInsert = Server.CreateObject("ADODB.Connection")
                         commInsert.Open strConnection
                         Set rsNewID = commInsert.Execute(strSQL)
                         'sf_id = rsNewID("NewID") 
                         'Response.write "got here<br>"
                        commInsert.Close()
                         Set commInsert = Nothing
                         'rsNewID.Close()
                         Set rsNewID = Nothing
                 ''''''''''''FINDS NEW COMPANY ID
 		                Set oConn2 = Server.CreateObject("ADODB.Connection")
		                oConn2.ConnectionTimeout = 100
		                oConn2.Provider = "MSDASQL"
		                oConn2.Open DATABASE
			                l_cSQL = "Select * FROM PreExistingCompanies WHERE companystatus='c' and CompanyEmail='"& OriginationEmail &"'  and CompanyPhone='"& OriginationPhoneNumber &"'  and ContactName='"& OriginationContactName &"'  and CompanyZip='"& OriginationZipCode &"'  and CompanyCity='"& OriginationCity &"' and  CompanyAddress='"& Originationaddress &"'  and CompanyName='"& OriginationCompany &"' and CompanyBuilding='"& OriginationBuilding &"' and CompanySuite='"& OriginationSuite &"'"
                            If Supervisor<>"y" or isNULL(Supervisor) then
                                l_cSQL = l_cSQL & " AND CompanyOwner='"& PreExistingRequestor &"'"
                                else
                                'l_cSQL = l_cSQL & " AND CompanyOwner is NULL "
                                l_cSQL = l_cSQL & " AND CompanyOwner='"& PreExistingRequestor &"'"
                            End if
                            'response.write "OriginationXXXXXl_cSql="&l_cSql&"<BR>"
			                SET oRs2 = oConn2.Execute(l_cSql)
					                if not oRs2.EOF then
                                    sf_id=oRs2("CompanyID")
                                    end if
                            SET oRs2 = Nothing
                        Set oConn2 = Nothing







                       else
                       sf_id=PreexistingOrigination
                    End if								
		Set oConn=Nothing

		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.ConnectionTimeout = 100
		oConn.Provider = "MSDASQL"
		oConn.Open DATABASE
			l_cSQL = "Select * FROM PreExistingCompanies WHERE companystatus='c' and CompanyEmail='"& DestinationEmail &"'  and CompanyPhone='"& DestinationPhoneNumber &"'  and ContactName='"& DestinationContactName &"'  and CompanyZip='"& DestinationZipCode &"'  and CompanyCity='"& DestinationCity &"' and  CompanyAddress='"& Destinationaddress &"'  and CompanyName='"& DestinationCompany &"' and CompanyBuilding='"& DestinationBuilding &"' and CompanySuite='"& DestinationSuite &"'"
            If Supervisor<>"y" or isNULL(Supervisor) then
                l_cSQL = l_cSQL & " AND CompanyOwner='"& PreExistingRequestor &"'"
                else
                '''l_cSQL = l_cSQL & " AND CompanyOwner is NULL "
                l_cSQL = l_cSQL & " AND CompanyOwner='"& PreExistingRequestor &"'"
            End if
            'Response.Write "DestinationXXXXXl_cSql="&l_cSql&"<BR>"
			SET oRs = oConn.Execute(l_cSql)
					if oRs.EOF then
                        'Response.write "ADD a NEW DESTINATION!!!!<BR>"
''''''''''''''''''''''''''''''''ADDED THIS
                        strSQL = "INSERT INTO PreExistingCompanies(CompanyName, CompanyAddress, CompanyBuilding, CompanySuite, CompanyCity, CompanyState, CompanyZip, ContactName, CompanyPhone, CompanyEmail, CompanyOwner, CompanyStatus) VALUES('"& DestinationCompany &"', '"& DestinationAddress &"', '"& DestinationBuilding &"', '"& DestinationSuite &"', '"& DestinationCity &"', '"& DestinationState &"', '"& DestinationZipCode &"', '"& DestinationContactName &"', '"& DestinationPhoneNumber &"', '"& DestinationEmail &"','"& PreExistingRequestor &"', 'c')"

                         strConnection = DATABASE
 
                        Set commInsert = Server.CreateObject("ADODB.Connection")
                         commInsert.Open strConnection
                         Set rsNewID = commInsert.Execute(strSQL)
                         'st_id = rsNewID("CompanyID") 
                        commInsert.Close()
                         Set commInsert = Nothing
                         'rsNewID.Close()
                         Set rsNewID = Nothing
                  ''''''''''''FINDS NEW COMPANY ID
 		                Set oConn2 = Server.CreateObject("ADODB.Connection")
		                oConn2.ConnectionTimeout = 100
		                oConn2.Provider = "MSDASQL"
		                oConn2.Open DATABASE
			                l_cSQL = "Select * FROM PreExistingCompanies WHERE companystatus='c' and CompanyEmail='"& DestinationEmail &"'  and CompanyPhone='"& DestinationPhoneNumber &"'  and ContactName='"& DestinationContactName &"'  and CompanyZip='"& DestinationZipCode &"'  and CompanyCity='"& DestinationCity &"' and  CompanyAddress='"& Destinationaddress &"'  and CompanyName='"& DestinationCompany &"' and CompanyBuilding='"& DestinationBuilding &"' and CompanySuite='"& DestinationSuite &"'"
                            If Supervisor<>"y" or isNULL(Supervisor) then
                                l_cSQL = l_cSQL & " AND CompanyOwner='"& PreExistingRequestor &"'"
                                else
                                'l_cSQL = l_cSQL & " AND CompanyOwner is NULL "
                                l_cSQL = l_cSQL & " AND CompanyOwner='"& PreExistingRequestor &"'"
                            End if
                            'response.write "DestinationXXXXXl_cSql="&l_cSql&"<BR>"
			                SET oRs2 = oConn2.Execute(l_cSql)
					                if not oRs2.EOF then
                                    st_id=oRs2("CompanyID")
                                    end if
                            SET oRs2 = Nothing
                        Set oConn2 = Nothing
''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''REMOVED THIS
			           ' Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				       '     RSEVENTS2.Open "PreExistingCompanies", DATABASE, 2, 2
				       '     RSEVENTS2.addnew
				       '     RSEVENTS2("CompanyName")=OriginationCompany
                       '     RSEVENTS2("CompanyAddress")=OriginationAddress
                       '     RSEVENTS2("CompanyBuilding")=OriginationBuilding
                       '     RSEVENTS2("CompanySuite")=OriginationSuite
                       '     RSEVENTS2("CompanyCity")=OriginationCity
                       '     RSEVENTS2("CompanyState")=OriginationState
                       '     RSEVENTS2("CompanyZip")=OriginationZipCode
                       '     RSEVENTS2("ContactName")=OriginationContactName
                       '     RSEVENTS2("CompanyPhone")=OriginationPhoneNumber
                       '     RSEVENTS2("CompanyEmail")=OriginationEmail
                            'If Supervisor<>"y" or isNULL(Supervisor) then
                       '     RSEVENTS2("CompanyOwner")=PreExistingRequestor
                            'End if
                       '     RSEVENTS2("CompanyStatus")="c"
				       '     RSEVENTS2.update
				       '     RSEVENTS2.close			
			           ' set RSEVENTS2 = nothing 
                       else
                       st_id=PreexistingDestination
                    End if	


   ''''''''ERROR HANDLING''''''''''
    'Response.write "GOT HERE!!!<BR>"
    'Response.write "PageStatus="&PageStatus&"***<BR>"
    If PageStatus="submit" and trim(OrderSubmitted)="yes" then
     'Response.write "GOT HERE two!!!<BR>"

''''''''''''''MOVED ERROR MESSAGES FROM HERE!










        ''sf_id=""
        Response.write "errormessage="&errormessage&"***<BR>"
        Response.write "sf_id="&sf_id&"***<BR>"
        Response.write "st_id="&st_id&"***<BR>"
        If trim(errormessage)="" and trim(sf_id)>"" and trim(st_id)>"" then
        'Response.write "GOT HERE!!!!!!!<BR>"
        'response.write "PreExistingRequestor="&PreExistingRequestor&"<BR>"
        'response.write "RequestorFirstName="&RequestorFirstName&"<BR>"
        'response.write "Requestorlastname="&Requestorlastname&"<BR>"
        'response.write "RequestorAddress="&RequestorAddress&"<BR>"
        'response.write "RequestorCity="&RequestorCity&"<BR>"
        'response.write "RequestorState="&RequestorState&"<BR>"
        'response.write "RequestorZipCode="&RequestorZipCode&"<BR>"
        'response.write "RequestorPhoneNumber="&RequestorPhoneNumber&"<BR>"
        'response.write "CelCarrier="&CelCarrier&"<BR>"
        'response.write "RequestorEmailAddress="&RequestorEmailAddress&"<BR>"
        'response.write "RequestorPassword="&RequestorPassword&"<BR>"
        'response.write "RequestorDeliveryNotification="&RequestorDeliveryNotification&"<BR>"

        If trim(PreExistingRequestor)>"" then
		    Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
		    RSEVENTS.Open "UserList", Database, 2, 2
		    RSEVENTS.Find "UserID='" & PreExistingRequestor & "'"
			    RSEVENTS("UserFirstName")=RequestorFirstName
			    RSEVENTS("Userlastname")=Requestorlastname
			    RSEVENTS("UserAddress")=RequestorAddress
			    RSEVENTS("UserCity")=RequestorCity
			    'RSEVENTS("UserState")=RequestorState
			    RSEVENTS("UserZipCode")=RequestorZipCode
			    RSEVENTS("UserPhoneNumber")=RequestorPhoneNumber
			    RSEVENTS("UserPhoneCarrier")=CelCarrier
			    RSEVENTS("UserEmailAddress")=RequestorEmailAddress
			    RSEVENTS("Password")=RequestorPassword
			    RSEVENTS("DeliveryNotifications")=DeliveryNotifications
		    RSEVENTS.update
		    RSEVENTS.close
		    set RSEVENTS = nothing

        End if













            'Response.write "Database="&Database&"<BR>"
		    'If trim(st_id)="DNP" or trim(st_id)="CPGPSCOT" then
            Set oConn = Server.CreateObject("ADODB.Connection")
		        oConn.ConnectionTimeout = 100
		        oConn.Provider = "MSDASQL"
		        oConn.Open DATABASE
		        ''''GETSNEWJOBNUMBER
		        'Response.Write "GOT HERE....<BR>"
		        l_cSql = "EXEC pr_GetJobNum"
		        Set oRs = oConn.Execute(l_cSql)
		        newjobnum = oRs.Fields("fh_id")			       
                'Response.write "xxxnewjobnum="&newjobnum&"<BR>"
                'Response.write "xxxpriority="&priority&"<BR>"

                '''''''''''''''''''''              
                oConn.Close
            Set oConn=Nothing 
            'Response.write "costcenterNumber="&costcenterNumber&"<BR>"
            'Response.write "PoNumber="&PoNumber&"<BR>"

           
			Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				RSEVENTS2.Open "fcfgthd", DATABASE, 2, 2
				RSEVENTS2.addnew
				RSEVENTS2("fh_ID")=NewJobNum
                'RSEVENTS2("fh_Status")="SCD"
                RSEVENTS2("fh_Status")="OPN"
                RSEVENTS2("fh_dispatcher")="AUTO"                
                RSEVENTS2("fh_ship_dt")=now()
                RSEVENTS2("fh_ready")=PickUpDateTime
                RSEVENTS2("Fh_Priority")=Priority
                RSEVENTS2("fh_lastchg")=now()
                RSEVENTS2("fh_bt_ID")=Trim(sBT_ID)
                RSEVENTS2("fh_co_id")=Trim(RequestorFirstName)&" "&Trim(RequestorLastName)
                RSEVENTS2("fh_co_phone")=Trim(RequestorPhoneNumber)
                RSEVENTS2("fh_co_email")=Trim(RequestoremailAddress)
                RSEVENTS2("fh_co_costcenter")=Trim(costcenterNumber)
                RSEVENTS2("fh_custpo")=Trim(PoNumber)
                RSEVENTS2("fh_statcode")="2"
                RSEVENTS2("DeliveryType")=DeliveryType
                RSEVENTS2("BasicCharge")=BasicCharge
                RSEVENTS2("fh_user1")=FXCourieruserid
				RSEVENTS2.update
				RSEVENTS2.close			
			set RSEVENTS2 = nothing  	       
           
            Set oConn=Nothing 
            'Response.write "xxxDatabase="&database&"<BR>"
			'''''''''''MOST BASIC AUTO ROUTING EVER!
            TheTime=time()
            'Response.write "TheTime="&TheTime&"<BR>"
            'tempVar=cdate("7:30:00 AM")
            TheWeekDay=WeekDay(PickupDateTime) 

            'If thetime>cdate("6:00:00 AM") and TheTime<cdate("11:59:59 PM") and TheWeekDay>1 and TheWeekDay<7 then


            If cInt(TheWeekday)=7 or cint(TheWeekday)=1 or (cint(TheWeekday)=2 AND thetime<cdate("6:00:00 AM")) then
                DriverID="803"
                UnitID="OnCall"
                else
                DriverID="801"
                UnitID="Dedicated1"
            End if
            '''''''''''''''''''''END''''''''''''''''
            Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				RSEVENTS2.Open "fclegs", DATABASE, 2, 2
				RSEVENTS2.addnew
				RSEVENTS2("fl_fh_ID")=Trim(NewJobNum)
                RSEVENTS2("fl_sf_id")=Trim(sf_id)
                RSEVENTS2("fl_sf_name")=Trim(OriginationCompany)
                RSEVENTS2("fl_sf_clname")=Trim(OriginationContactName)
                RSEVENTS2("fl_sf_phone")=Trim(OriginationPhoneNumber)
                RSEVENTS2("fl_sf_email")=Trim(OriginationEmail)
                RSEVENTS2("fl_sf_addr1")=Trim(OriginationAddress)
                RSEVENTS2("fl_sf_Building")=Trim(OriginationBuilding)
                RSEVENTS2("fl_sf_addr2")=Trim(OriginationSuite)
                RSEVENTS2("fl_sf_city")=Trim(OriginationCity)
                RSEVENTS2("fl_sf_state")=Trim(OriginationState)
                RSEVENTS2("fl_sf_country")="US"
                RSEVENTS2("fl_sf_zip")=Trim(OriginationZipCode)
                RSEVENTS2("fl_st_id")=Trim(st_id)
                RSEVENTS2("fl_st_name")=Trim(DestinationCompany)
                RSEVENTS2("fl_st_clname")=Trim(DestinationContactName)
                RSEVENTS2("fl_st_phone")=Trim(DestinationPhoneNumber)
                RSEVENTS2("fl_st_email")=Trim(DestinationEmail)
                RSEVENTS2("fl_st_addr1")=Trim(DestinationAddress)
                RSEVENTS2("fl_st_Building")=Trim(DestinationBuilding)
                RSEVENTS2("fl_st_addr2")=Trim(DestinationSuite)
                RSEVENTS2("fl_st_city")=Trim(DestinationCity)
                RSEVENTS2("fl_st_state")=Trim(DestinationState)
                RSEVENTS2("fl_st_country")="US"
                RSEVENTS2("fl_st_zip")=Trim(DestinationZipCode)
                RSEVENTS2("fl_un_id")=Trim(UnitID)
                RSEVENTS2("fl_dr_id")=Trim(DriverID)
                RSEVENTS2("fl_sf_comment")=Trim(Comments)
                RSEVENTS2("fl_t_disp")=now()                
                RSEVENTS2("fl_st_rta")=Trim(DeliveryDateTime)
                RSEVENTS2("fl_leg_status")="c"
				RSEVENTS2.update
				RSEVENTS2.close			
			set RSEVENTS2 = nothing 

			Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				RSEVENTS2.Open "fcrefs", DATABASE, 2, 2
				RSEVENTS2.addnew
				RSEVENTS2("rf_fh_ID")=Trim(NewJobNum)
                'If trim(Partnumber)>"" then
                RSEVENTS2("PartNumber")=Trim(PartNumber)
                    'else
                RSEVENTS2("rf_ref")=Trim(NewJobNum)
                'End if
                RSEVENTS2("ref_status")=NULL
                RSEVENTS2("rf_box")=trim(rf_box)
                RSEVENTS2("NumberOfPieces")=Trim(Pieces)
                RSEVENTS2("NumberOfPallets")=Trim(NumberOfPallets)
                RSEVENTS2("Weight")=Trim(DimWeight)
                RSEVENTS2("DimLength")=Trim(DimLength)
                RSEVENTS2("DimWidth")=Trim(DimWidth)
                RSEVENTS2("DimHeight")=Trim(DimHeight)
                RSEVENTS2("MeasurementType")=Trim(MeasurementType)
                RSEVENTS2("Hazmat")=Trim(isHazmat)
                RSEVENTS2("Refrigerate")=Trim(Refrigerate)


				RSEVENTS2.update
				RSEVENTS2.close			
			set RSEVENTS2 = nothing 

		        
               'Response.Write "YAY!!!! You're DONE!!!!<BR>"
		
         
            'Response.write "newjobnum="&newjobnum&"<BR>"
  ''''''''''''''''LETTER to Requestor'''''''''''''''''''
   				    Body = "Your shipment request (#"& newjobnum &") has been successfully placed online:<br><br>"  
                    Body = Body & "Your shipment will be picked up sometime between "& pickupdatetime &" and "& DeliveryDateTime &" <br><br>"  
  			        Body = Body & "PLEASE PRINT OUT THIS NOTICE AND ATTACH IT TO YOUR SHIPMENT! <br><br>"
                    Body = Body & "If a barcode is not visible below, then please <a href='http://www.logisticorp.us/intranet/courier/orderentry/FleetExpressCourierOrderConfirmation.asp?jid="&newjobnum&"'>click here</a> for a printable waybill. <br><br>"
                    BarCodeText=newjobnum
			        'BarCodeText="1234/567-89"
			        'Response.Write "BarCodeText="&BarCodeText&"<BR>"
			        If BarCodeText>"" then
				        'Response.write BarCodeText&"<br><IMG SRC=""../images/barcodes/asterisk.gif"" WIDTH=""17"" HEIGHT=""40"">"
                        TheBarCode="<IMG SRC='http://www.logisticorp.us/intranet/courier/images/barcodes/asterisk.gif' WIDTH='17' HEIGHT='40'>"
				        For x = 1 to Len(Trim(BarCodeText))
					        DisplayBarCode=mid(BarCodeText,x,1)
					        If DisplayBarCode="/" then
						        'Response.write "<IMG SRC=""../images/barcodes/!slash.gif"" WIDTH=""17"" HEIGHT=""40"">"
						        TheBarCode=TheBarCode&"<IMG SRC='http://www.logisticorp.us/intranet/courier/images/barcodes/!slash.gif' WIDTH='17' HEIGHT='40'>"
                                else
						        'Response.Write "<IMG SRC=""../images/barcodes/" & DisplayBarCode & _
								'        ".gif"" WIDTH=""17"" HEIGHT=""40"">"
                                TheBarCode=TheBarCode&"<IMG SRC='http://www.logisticorp.us/intranet/courier/images/barcodes/" & DisplayBarCode & ".gif' WIDTH='17' HEIGHT='40'>"
					        End if
				        Next

				        'Code 39 barcodes require an asterisk as the start and stop characters
				        'Response.write "<IMG SRC=""../images/barcodes/asterisk.gif"" WIDTH=""17"" HEIGHT=""40"">"
                        TheBarCode=TheBarCode&"<IMG SRC='http://www.logisticorp.us/intranet/courier/images/barcodes/asterisk.gif' WIDTH='17' HEIGHT='40'>"
			        End if                   
                    Body = Body & TheBarCode&"<BR><BR>"
                    Body = Body & "REQUESTOR INFORMATION:<BR>"
                    Body = Body & "Name: "&  RequestorFirstName &" "&RequestorLastName&"<br>"  
                    Body = Body & "Phone Number: "&  RequestorPhoneNumber &"<br>"  
                    Body = Body & "Email Address: "&  RequestorEmailAddress &"<br>"   
                    Body = Body & "PO Number: "&  PONumber &"<br>"  
                    Body = Body & "Cost Center Number: "&  CostCenterNumber &"<br><br>" 
                    Body = Body & "COMMODITY INFORMATION:<BR>" 
                    Body = Body & "PartNumber: "&  PartNumber &"<br>" 
                    Body = Body & "Pieces: "&  Pieces &" "& rf_box &"<br>" 
                    'Body = Body & "Palletized: "&  IsPalletized &"<br>"   
                    'Body = Body & "Number Of Pallets: "&  NumberOfPallets &"<br>"  
                    Body = Body & "Weight: "&  DimWeight &"LBS<br>"
                    Body = Body & " Dimensions: "&  DimLength &" X "&  DimWidth &" X "&  DimHeight &" inches<br>"       
                    Body = Body & "ORIGINATION:<BR>"   
                    Body = Body & "Company: "&  OriginationCompany &"<br>"   
                    Body = Body & "Address: "&  OriginationAddress &"<br>"   
                    Body = Body & "Building: "&  OriginationBuilding &"<br>"   
                    Body = Body & "Floor/Cube: "&  OriginationSuite &"<br>"   
                    Body = Body & "City: "&  OriginationCity &"<br>"   
                    Body = Body & "State: "&  OriginationState &"<br>"  
                    Body = Body & "Zip Code: "&  OriginationZipCode &"<br>"   
                    Body = Body & "Contact Name: "&  OriginationContactName &"<br>"   
                    Body = Body & "Phone Number: "&  OriginationPhoneNumber &"<br>"   
                    Body = Body & "Email: "&  OriginationEmail &"<br>" 
                    Body = Body & "Ready Date/Time: "&  PickUpDateTime &"<br><br>" 
                    Body = Body & "DESTINATION:<BR>"  
                    Body = Body & "Company: "&  DestinationCompany &" (POD REQUIRED)<br>"  
                    Body = Body & "Address: "&  DestinationAddress &"<br>"  
                    Body = Body & "Building: "&  DestinationBuilding &"<br>" 
                    Body = Body & "Floor/Cube: "&  DestinationSuite &"<br>" 
                    Body = Body & "City: "&  DestinationCity &"<br>"   
                    Body = Body & "State: "&  DestinationState &"<br>"  
                    Body = Body & "Zip Code: "&  DestinationZipCode &"<br>"  
                    Body = Body & "Contact Name: "&  DestinationContactName &"<br>"  
                    Body = Body & "Phone Number: "&  DestinationPhoneNumber &"<br>"  
                    Body = Body & "Email: "&  DestinationEmail &"<br>"   
                    Body = Body & "Delivery Date/Time: "&  DeliveryDateTime &"<br><br>" 
                    If trim(Comments)>"" then
                        Body = Body & "SPECIAL INSTRUCTIONS:<br>" 
                        Body = Body & ""&  comments &"<br><br>" 
                    End if
                    'Body = Body & "Once the order has been reviewed, you will recieve notification whether it has been accepted or refused.<br><br>" 


				    Body = Body & "Thank you,<br><br>"  
				    Body = Body & "Fleet Express Courier<br>"  
				    Body = Body &  "FleetX@LogisticorpGroup.com<br>"  
				    Body = Body & "(214) 882-0620<br><br>"
				    'Recipient=FirstName&" "&LastName
			        SentToEmail=RequestorEmailAddress
				    'Email="KWETI.Mailbox@am.kwe.com"
				    'Email="mark@maggiore.net"
				    Set objMail = CreateObject("CDONTS.Newmail")
				    objMail.From = "FleetX@LogisticorpGroup.com"
				    objMail.To = SentToEmail
				    objMail.cc = "mark.maggiore@logistiCorp.us"
				    'objMail.cc = "4692269939@tmomail.net;pcurrin@ti.com"
				    objMail.Subject = "Thank you for your Fleet Express Courier shipment request"
				    objMail.MailFormat = cdoMailFormatMIME
				    objMail.BodyFormat = cdoBodyFormatHTML
				    objMail.Body = Body
                    If trim(lcase(SentToEmail))<>"fleetexpress@logisticorp.us" then
				        objMail.Send
                    End if
				    Set objMail = Nothing         
            
            		
				    Body = "There has been a new Fleet Express Courier shipment request (#"& newjobnum &") placed online:<br><br>"   

  			        
                    BarCodeText=newjobnum
			        'BarCodeText="1234/567-89"
			        'Response.Write "BarCodeText="&BarCodeText&"<BR>"
			        If BarCodeText>"" then
				        'Response.write BarCodeText&"<br><IMG SRC=""../images/barcodes/asterisk.gif"" WIDTH=""17"" HEIGHT=""40"">"
                        TheBarCode="<IMG SRC='http://www.logisticorp.us/intranet/courier/images/barcodes/asterisk.gif' WIDTH='17' HEIGHT='40'>"
				        For x = 1 to Len(Trim(BarCodeText))
					        DisplayBarCode=mid(BarCodeText,x,1)
					        If DisplayBarCode="/" then
						        'Response.write "<IMG SRC=""../images/barcodes/!slash.gif"" WIDTH=""17"" HEIGHT=""40"">"
						        TheBarCode=TheBarCode&"<IMG SRC='http://www.logisticorp.us/intranet/courier/images/barcodes/!slash.gif' WIDTH='17' HEIGHT='40'>"
                                else
						        'Response.Write "<IMG SRC=""../images/barcodes/" & DisplayBarCode & _
								'        ".gif"" WIDTH=""17"" HEIGHT=""40"">"
                                TheBarCode=TheBarCode&"<IMG SRC='http://www.logisticorp.us/intranet/courier/images/barcodes/" & DisplayBarCode & ".gif' WIDTH='17' HEIGHT='40'>"
					        End if
				        Next

				        'Code 39 barcodes require an asterisk as the start and stop characters
				        'Response.write "<IMG SRC=""../images/barcodes/asterisk.gif"" WIDTH=""17"" HEIGHT=""40"">"
                        TheBarCode=TheBarCode&"<IMG SRC='http://www.logisticorp.us/intranet/courier/images/barcodes/asterisk.gif' WIDTH='17' HEIGHT='40'>"
			        End if


                    Body = Body & "REQUESTOR INFORMATION:<BR>"
                    Body = Body & "Name: "&  RequestorFirstName &" "& RequestorLastName &"<br>"  
                    Body = Body & "Phone Number: "&  RequestorPhoneNumber &"<br>"  
                    Body = Body & "Email Address: "&  RequestorEmailAddress &"<br>"   
                    Body = Body & "PO Number: "&  PONumber &"<br>"  
                    Body = Body & "Cost Center Number: "&  CostCenterNumber &"<br><br>" 
                    Body = Body & "COMMODITY INFORMATION:<BR>"
                    Body = Body & "Part Number: "&  PartNumber &"<br>" 
                    Body = Body & "Pieces: "&  Pieces &" "& rf_box &"<br>" 
                    'Body = Body & "Palletized: "&  IsPalletized &"<br>"   
                    'Body = Body & "Number Of Pallets: "&  NumberOfPallets &"<br>"  
                    Body = Body & "Weight: "&  DimWeight &"LBS<br>"
                    Body = Body & " Dimensions: "&  DimLength &" X "&  DimWidth &" X "&  DimHeight &" inches<br>"       
  
                    
                    Body = Body & "Hazmat: "&  IsHazmat &"<br>"
                    Body = Body & "Refrigerate: "&  Refrigerate &"<br><br>"
                    Body = Body & "ORIGINATION:<BR>"   
                    Body = Body & "Company: "&  OriginationCompany &"<br>"   
                    Body = Body & "Address: "&  OriginationAddress &"<br>"  
                    Body = Body & "Building: "&  OriginationBuilding &"<br>" 
                     Body = Body & "Floor/Cube: "&  OriginationSuite &"<br>" 
                    Body = Body & "City: "&  OriginationCity &"<br>"   
                    Body = Body & "State: "&  OriginationState &"<br>"  
                    Body = Body & "Zip Code: "&  OriginationZipCode &"<br>"   
                    Body = Body & "Contact Name: "&  OriginationContactName &"<br>"   
                    Body = Body & "Phone Number: "&  OriginationPhoneNumber &"<br>"   
                    Body = Body & "Email: "&  OriginationEmail &"<br>" 
                    Body = Body & "Ready Date/Time: "&  PickUpDateTime &"<br><br>" 
                    Body = Body & "DESTINATION:<BR>"  
                    Body = Body & "Company: "&  DestinationCompany &" (POD REQUIRED)<br>"  
                    Body = Body & "Address: "&  DestinationAddress &"<br>"  
                    Body = Body & "Building: "&  DestinationBuilding &"<br>" 
                    Body = Body & "Floor/Cube: "&  DestinationSuite &"<br>" 
                    Body = Body & "City: "&  DestinationCity &"<br>"   
                    Body = Body & "State: "&  DestinationState &"<br>"  
                    Body = Body & "Zip Code: "&  DestinationZipCode &"<br>"  
                    Body = Body & "Contact Name: "&  DestinationContactName &"<br>"  
                    Body = Body & "Phone Number: "&  DestinationPhoneNumber &"<br>"  
                    Body = Body & "Email: "&  DestinationEmail &"<br>"   
                    Body = Body & "Delivery Date/Time: "&  DeliveryDateTime &"<br><br>"
                     If trim(Comments)>"" then
                        Body = Body & "SPECIAL INSTRUCTIONS:<br>" 
                        Body = Body & ""&  comments &"<br><br>" 
                    End if                    
                    'Body = Body & "<a href='http://www.logisticorp.us/intranet/dedicatedfleets/orderentry/FleetExpressOrderConfirmation.asp?bid=86&pid=disp&jid="& newjobnum &"'>To Route or Cancel this request, click here</a><br><br>" 
				    Body = Body & "Thank you,<br><br>"  
				    Body = Body & "Fleet Express Courier<br>"  
				    Body = Body &  "FleetX@LogisticorpGroup.com<br>"  
				    Body = Body & "(214) 882-0620<br><br>"
				    'Recipient=FirstName&" "&LastName
			        'SentToEmail="mark.maggiore@logisticorp.us;FleetX@LogisticorpGroup.com"
                    SentToEmail="mark.maggiore@logisticorp.us;FleetXDFW@logisticorp.us"
				    'Email="KWETI.Mailbox@am.kwe.com"
				    'Email="mark@maggiore.net"
				    Set objMail = CreateObject("CDONTS.Newmail")
				    objMail.From = "FleetX@LogisticorpGroup.com"
				    objMail.To = SentToEmail
				    'objMail.cc = RequestorEmailAddress
				    'objMail.cc = "4692269939@tmomail.net;pcurrin@ti.com"
				    objMail.Subject = "New Fleet Express Courier Shipment Request"
				    objMail.MailFormat = cdoMailFormatMIME
				    objMail.BodyFormat = cdoBodyFormatHTML
				    objMail.Body = Body
				    objMail.Send
				    Set objMail = Nothing
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    If (int(DimLength)>42 or int(DimWidth)>48 or int(DimHeight)>48) AND xyz="removethisfornow" then
				        Body = "There has been an oversized Fleet Express Courier shipment request (#"& newjobnum &") placed online:<br><br>"   

  			        
                    BarCodeText=newjobnum
			        'BarCodeText="1234/567-89"
			        'Response.Write "BarCodeText="&BarCodeText&"<BR>"
			        If BarCodeText>"" then
				        'Response.write BarCodeText&"<br><IMG SRC=""../images/barcodes/asterisk.gif"" WIDTH=""17"" HEIGHT=""40"">"
                        TheBarCode="<IMG SRC='http://www.logisticorp.us/intranet/courier/images/barcodes/asterisk.gif' WIDTH='17' HEIGHT='40'>"
				        For x = 1 to Len(Trim(BarCodeText))
					        DisplayBarCode=mid(BarCodeText,x,1)
					        If DisplayBarCode="/" then
						        'Response.write "<IMG SRC=""../images/barcodes/!slash.gif"" WIDTH=""17"" HEIGHT=""40"">"
						        TheBarCode=TheBarCode&"<IMG SRC='http://www.logisticorp.us/intranet/courier/images/barcodes/!slash.gif' WIDTH='17' HEIGHT='40'>"
                                else
						        'Response.Write "<IMG SRC=""../images/barcodes/" & DisplayBarCode & _
								'        ".gif"" WIDTH=""17"" HEIGHT=""40"">"
                                TheBarCode=TheBarCode&"<IMG SRC='http://www.logisticorp.us/intranet/courier/images/barcodes/" & DisplayBarCode & ".gif' WIDTH='17' HEIGHT='40'>"
					        End if
				        Next

				        'Code 39 barcodes require an asterisk as the start and stop characters
				        'Response.write "<IMG SRC=""../images/barcodes/asterisk.gif"" WIDTH=""17"" HEIGHT=""40"">"
                        TheBarCode=TheBarCode&"<IMG SRC='http://www.logisticorp.us/intranet/courier/images/barcodes/asterisk.gif' WIDTH='17' HEIGHT='40'>"
			        End if 
                        
                        Body = Body & "REQUESTOR INFORMATION:<BR>"
                        Body = Body & "Name: "&  RequestorFirstName &" "& RequestorLastName &"<br>"  
                        Body = Body & "Phone Number: "&  RequestorPhoneNumber &"<br>"  
                        Body = Body & "Email Address: "&  RequestorEmailAddress &"<br>"   
                        Body = Body & "PO Number: "&  PONumber &"<br>"  
                        Body = Body & "Cost Center Number: "&  CostCenterNumber &"<br><br>" 
                        Body = Body & "COMMODITY INFORMATION:<BR>" 
                        Body = Body & "Part Number: "&  PartNumber &"<br>" 
                        Body = Body & "Pieces: "&  Pieces &" "& rf_box &"<br>" 
                        'Body = Body & "Palletized: "&  IsPalletized &"<br>"   
                        'Body = Body & "Number Of Pallets: "&  NumberOfPallets &"<br>"  
                        Body = Body & "Weight: "&  DimWeight &"LBS<br>"
                        Body = Body & " Dimensions: "&  DimLength &" X "&  DimWidth &" X "&  DimHeight &" inches<br>"       
  
                    
                        Body = Body & "Hazmat: "&  IsHazmat &"<br>"
                        Body = Body & "Refrigerate: "&  Refrigerate &"<br><br>"
                        Body = Body & "ORIGINATION:<BR>"   
                        Body = Body & "Company: "&  OriginationCompany &"<br>"   
                        Body = Body & "Address: "&  OriginationAddress &"<br>" 
                        Body = Body & "Building: "&  OriginationBuilding &"<br>" 
                        Body = Body & "Floor/Cube: "&  OriginationSuite &"<br>"                           
                        Body = Body & "City: "&  OriginationCity &"<br>"   
                        Body = Body & "State: "&  OriginationState &"<br>"  
                        Body = Body & "Zip Code: "&  OriginationZipCode &"<br>"   
                        Body = Body & "Contact Name: "&  OriginationContactName &"<br>"   
                        Body = Body & "Phone Number: "&  OriginationPhoneNumber &"<br>"   
                        Body = Body & "Email: "&  OriginationEmail &"<br>" 
                        Body = Body & "Ready Date/Time: "&  PickUpDateTime &"<br><br>" 
                        Body = Body & "DESTINATION:<BR>"  
                        Body = Body & "Company: "&  DestinationCompany &" (POD REQUIRED)<br>"  
                        Body = Body & "Address: "&  DestinationAddress &"<br>" 
                        Body = Body & "Building: "&  DestinationBuilding &"<br>" 
                        Body = Body & "Floor/Cube: "&  DestinationSuite &"<br>"                         
                        Body = Body & "City: "&  DestinationCity &"<br>"   
                        Body = Body & "State: "&  DestinationState &"<br>"  
                        Body = Body & "Zip Code: "&  DestinationZipCode &"<br>"  
                        Body = Body & "Contact Name: "&  DestinationContactName &"<br>"  
                        Body = Body & "Phone Number: "&  DestinationPhoneNumber &"<br>"  
                        Body = Body & "Email: "&  DestinationEmail &"<br>"   
                        Body = Body & "Delivery Date/Time: "&  DeliveryDateTime &"<br><br>"
                         If trim(Comments)>"" then
                            Body = Body & "SPECIAL INSTRUCTIONS:<br>" 
                            Body = Body & ""&  comments &"<br><br>" 
                        End if
                        Body = Body & "Thank you,<br><br>"  
				        Body = Body & "Fleet Express Courier<br>"  
				        Body = Body &  "FleetX@LogisticorpGroup.com<br>"  
				        Body = Body & "(214) 882-0620<br><br>"
				        'Recipient=FirstName&" "&LastName
			            'SentToEmail="mark.maggiore@logisticorp.us;FleetX@LogisticorpGroup.com"
                        SentToEmail="mark.maggiore@logisticorp.us;FleetXDFW@logisticorp.us"
				        'Email="KWETI.Mailbox@am.kwe.com"
				        'Email="mark@maggiore.net"
				        Set objMail = CreateObject("CDONTS.Newmail")
				        objMail.From = "FleetX@LogisticorpGroup.com"
				        objMail.To = SentToEmail
				        'objMail.cc = RequestorEmailAddress
				        'objMail.cc = "4692269939@tmomail.net;pcurrin@ti.com"
				        objMail.Subject = "Oversized Fleet Express Courier Shipment Request"
				        objMail.MailFormat = cdoMailFormatMIME
				        objMail.BodyFormat = cdoBodyFormatHTML
				        objMail.Body = Body
				        objMail.Send
				        Set objMail = Nothing
                    End if
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                    session("FXCourierUserID")=FXCourierUserID


                    Response.Redirect("FleetExpressCourierOrderConfirmation.asp?x=1&y=1&bid=86&pid=view&jid="& newjobnum &"&Internal="&Internal&"&VarA="&FXCourieruserid&"&VarB="&Supervisor)	
         else
        ErrorMessage=ErrorMessage & "<br>You've encountered an unexpected error.  Please try again.  If you continue to have this problem, contact Mark Maggiore at 214-956-0400 xt. 212 for assistance."		  
          
          
           End if
          	
        End if

    End if
    ''''''''END ERROR HANDLING''''''

     %>
</head>
<BODY leftMargin="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" OnLoad=document.OrderForm1.<%=HighlightedField%>.focus()>
<%
TheTime=time()
'Response.write "TheTime="&TheTime&"<BR>"
tempVar=cdate("7:30:00 AM")
TheWeekDay=WeekDay(Now()) 
'Response.write "TheWeekDay="&TheWeekDay&"<BR>"
'Response.write "tempVar="&tempVar&"<BR>"
'''If thetime>cdate("7:30:00 AM") and TheTime<cdate("7:00:00 PM") and TheWeekDay>1 and TheWeekDay<7 then
'If thetime>cdate("10:40:00 AM") and TheTime<cdate("7:00:00 PM") then
%>
<table border="0" cellpadding="0" cellspacing="0" align="center" bgcolor="<%=bgcolor%>" width="770">
<tr>
    <td><img src="../images/pixel.gif" width="30" height="1" /></td>
    <td>
<table cellpadding="0" cellspacing="0" border="0" width="100%">
<tr>
    <td class="MainPageTextCenterLargeBlack" colspan="3">Fleet Express Courier Order Page</td>
</tr>
<tr><td>&nbsp;</td></tr>
<tr><td>&nbsp;</td></tr>
</table>
<table cellpadding="2" cellspacing="5" border="0" width="100%" class="BlueTable">
<tr>
    <td class="FleetExpressTextBlueBold" colspan="3"><b>SYSTEM MESSAGE:  8/27/2012</b><br /><BR />Please include the <b>floor number</b> on all origins/destinations, where applicable.<BR /><BR />Failure to do so may result in late pick-ups/deliveries.</td>
</tr>
</table>
<!--table cellpadding="0" cellspacing="0" border="1">
<tr>
    <td colspan="3"><font color="red">ATTENTION:  Fleet Express Courier Service is currently only available Monday-Friday
    from 7:30 AM to 7:00 PM.  if you have an order needing delivery outside of those times, please call our drivers direct at 214-882-0620.</font></td>
</tr>
</table-->

<form method="post" name="OrderForm1" action="FleetExpressCourierOrder.asp?Internal=<%=Internal%>">
    <table border="0" cellpadding="0" cellspacing="0" align="center">
        
        <tr><td align="left"><img src="../images/pixel.gif" height="15" width="1" /></td></tr>
        <tr><td colspan="4">Fields marked <font color="red">*</font> are required.</td></tr>

         <%If Errormessage>"" then%>
         <tr><td>&nbsp;</td></tr>
         <tr>
            <td align="center" colspan="4">
                        
                <table cellpadding="2" cellspacing="2" border="1" bordercolor="red">
                    <tr>
                        <td class="errormessage">
                            <%
                            Response.write " * * * ERROR:  "&ErrorMessage& " * * * "
                             %>
                        </td>
                    </tr>
                </table>
                
           </td>
        </tr> 
        <tr><td>&nbsp;</td></tr>
        <%End if%>


        <tr><td colspan="4" bgcolor="black"><img src="../images/pixel.gif" height="1" width="1" /></td></tr>
        <tr><td><img src="../images/pixel.gif" height="2" width="1" /></td></tr>
        <tr>
            <%somenumber=13
            If Supervisor<>"y" or isNULL(Supervisor) then
            SomeNumber=SomeNumber-1
            End if
            'Response.write "somenumber="&SomeNumber&"<BR>"
             %>
            <td valign="top" rowspan="<%=SomeNumber %>" class="OrderHeader">REQUESTOR INFORMATION<img src="../images/pixel.gif" height="1" width="15" /></td>
            

            <%
            'response.write   "PreExistingRequestor="&PreExistingRequestor&"<BR>"
            'response.write   "Supervisor="&Supervisor&"<BR>"
            'response.write   "RequestorFirstName="&RequestorFirstName&"<BR>"
            If Trim(Supervisor)="y" then
                'response.write   "XXX GOT HERE #4<BR>"
                'If Supervisor="y" and  trim(RequestorFirstName)="" then
                If Supervisor="y" then 
                'response.write   "XXX GOT HERE #5 !!!!<BR>"
                %>
                
                <td class="FleetExpressTextBlackBold" align="left" nowrap><font color="red">&nbsp;&nbsp;</font>Pre-Existing Requestor</td>
                <td width="10"><img src="../images/pixel.gif" /></td>
                <td align="left">
								<%
                                'response.write  "THIS ONE! PreExistingRequestor="&PreExistingRequestor&"<BR>"
                                 %>
                                <select name="PreExistingRequestor" ID="Select2"  onChange="form.submit()" class="textbox" >
								<option value="" <%if trim(PreExistingRequestor)="" then response.Write " selected" end if%>>Select from this list or fill in all fields below</option>
                                <%

									Set oConn = Server.CreateObject("ADODB.Connection")
									oConn.ConnectionTimeout = 100
									oConn.Provider = "MSDASQL"
									oConn.Open DATABASE
										l_cSQL = "Select * FROM UserList WHERE userstatus='c' and supervisor is NULL ORDER BY UserFirstName, UserLastName"
										'response.write "l_cSQL="&l_cSQL&"<BR>"
                                        SET oRs = oConn.Execute(l_cSql)
												Do While not oRs.EOF
                                                supertempRequestorUserID=trim(oRs("UserID"))
												supertempRequestorFirstName=trim(oRs("UserFirstName"))
							                    supertempRequestorLastName=trim(oRs("UserLastName"))
                                                SupertempRequestorName=SupertempRequestorFirstName&" "&SupertempRequestorLastName
                                                If trim(Preexistingrequestor)>"" and trim(Preexistingrequestor)=trim(SuperTempRequestorUserID) then
                                                'If trim(Preexistingrequestor)>"" then
                                                    'response.write  "Got here #666 XXX<BR>"
                                                    tempRequestorUserID=trim(oRs("UserID"))
												    tempRequestorFirstName=trim(oRs("UserFirstName"))
							                        tempRequestorLastName=trim(oRs("UserLastName"))                                                   
                                                    tempRequestorAddress=trim(oRs("UserAddress"))
                                                    tempRequestorCity=trim(oRs("UserCity"))
												    tempRequestorState=trim(oRs("UserState"))
							                        tempRequestorZipCode=trim(oRs("UserZipCode"))
                                                    tempRequestorPhoneNumber=trim(oRs("UserPhoneNumber"))
                                                    tempRequestorPhoneCarrier=trim(oRs("UserPhoneCarrier"))
												    tempRequestorEmailAddress=trim(oRs("UserEmailAddress"))
							                        tempRequestorUsername=trim(oRs("UserName"))
                                                    tempRequestorPassword=trim(oRs("Password"))
                                                    tempRequestorDeliveryNotifications=trim(oRs("DeliveryNotifications"))
												    tempRequestorSupervisor=trim(oRs("Supervisor"))
							                        tempRequestorUserStatus=trim(oRs("UserStatus"))
                                                    tempRequestorDateCreated=trim(oRs("DateCreated"))
                                                    else
                                                    If trim(tempRequestorDateCreated)="" then
                                                    ''''''''NEW Requestor, so base info on form only!
                                                    TempRequestorUserID=Request.form("PreExistingRequestor")
                                                    PreExistingRequestor=Request.form("PreExistingRequestor")
                                                    TempRequestorFirstName=Request.form("RequestorFirstName")
                                                    TempRequestorLastName=Request.form("RequestorLastName")
                                                    TempRequestorPhoneNumber=Request.form("RequestorPhoneNumber")
                                                    TempRequestorEmailAddress=Request.form("RequestorEmailAddress")

                                                    TempRequestorAddress=Request.form("RequestorAddress")
                                                    TempRequestorCity=Request.form("RequestorCity")
                                                    TempRequestorZipCode=Request.form("RequestorZipCode")
                                                    TempRequestorSTate=Request.form("RequestorSTate")
                                                    TempRequestorPhoneCarrier=Request.form("celcarrier")
                                                    TempRequestorName=Request.form("RequestorName")
                                                    TempRequestorPassword=Request.form("RequestorPassword")
                                                    TempRequestorDeliveryNotifications=Request.form("DeliveryNotifications")
                                                    'response.write  "test GOT HERE #16969<BR>"
                                                    End if
												    'tempRequestorFirstName=""
							                        'tempRequestorLastName=""
                                                End if

                                                'response.write  "test GOT HERE #1<BR>"
                                                'If trim(PreExistingRequestor)=trim(tempRequestorID) and trim(PreExistingRequestor)>"" then
												'    RequestorName=TempRequestorName
												'    RequestorPhoneNumber=TempRequestorPhone
                                                '    RequestorEmailAddress=TempRequestorEmail
                                                'End if								
											%>
											<option value="<%=SuperTempRequestorUserID%>" <%if trim(Preexistingrequestor)=trim(SuperTempRequestorUserID) then response.write "selected" End if%>><%=SuperTempRequestorName%></option>
											<%
										oRs.movenext
										LOOP
									Set oConn=Nothing									
									%>
								</select> 
                                <%'response.write  "Here's the var I'm looking for:"&tempRequestorFirstName&"<BR>" %>
           </td>
       </tr>
       <tr><td><img src="../images/pixel.gif" height="3" width="1" /><%'response.write "l_cSQL="&l_cSQL&"<BR>"%></td></tr>
        <tr>
        <%
        else 
        'response.write   "XXX GOT HERE1!!!<BR>"
									'If trim(RequestorFirstName)="" then
                                        Set oConn = Server.CreateObject("ADODB.Connection")
									    oConn.ConnectionTimeout = 100
									    oConn.Provider = "MSDASQL"
									    oConn.Open DATABASE
										    l_cSQL = "Select * FROM UserList WHERE UserID='"& FXCourierUserID &"' and Userstatus='c'"
										    'response.write "l_cSQL="&l_cSQL&"<BR>"
                                            SET oRs = oConn.Execute(l_cSql)
												    If not oRs.EOF then
                                                        'response.write  "GOT HERE!!!!<BR>"
                                                        'tempRequestorID=trim(oRs("RequestorID"))
                                                        PreExistingRequestor=trim(oRs("UserID"))
                                                        tempRequestorUserID=trim(oRs("UserID"))
												        tempRequestorFirstName=trim(oRs("UserFirstName"))
							                            tempRequestorLastName=trim(oRs("UserLastName"))
                                                        tempRequestorAddress=trim(oRs("UserAddress"))
                                                        tempRequestorCity=trim(oRs("UserCity"))
												        tempRequestorState=trim(oRs("UserState"))
							                            tempRequestorZipCode=trim(oRs("UserZipCode"))
                                                        tempRequestorPhoneNumber=trim(oRs("UserPhoneNumber"))
                                                        tempRequestorPhoneCarrier=trim(oRs("UserPhoneCarrier"))
												        tempRequestorEmailAddress=trim(oRs("UserEmailAddress"))
							                            tempRequestorUsername=trim(oRs("UserName"))
                                                        tempRequestorPassword=trim(oRs("Password"))
                                                        tempRequestorDeliveryNotifications=trim(oRs("DeliveryNotifications"))
                                                        'Response.write "test Got Here #2<BR>"
												        tempRequestorSupervisor=trim(oRs("Supervisor"))
							                            tempRequestorUserStatus=trim(oRs("UserStatus"))
                                                        tempRequestorDateCreated=trim(oRs("DateCreated"))
                                                        formattedphonenumber=Replace(tempRequestorPhone," ","")
                                                        formattedphonenumber=Replace(formattedphonenumber,"-","")
                                                        formattedphonenumber=Replace(formattedphonenumber,".","")
                                                        formattedphonenumber=Replace(formattedphonenumber,"(","")
                                                        formattedphonenumber=Replace(formattedphonenumber,")","")
                                                        Select Case DeliveryNotifications
                                                            Case "email"
                                                                TempRequestorEmail=TempRequestorEmail
                                                            Case "text"
                                                                TempRequestorEmail=FormattedPhoneNumber&"@"&tempUserPhoneCarrier
                                                            Case "both"
                                                                TempRequestorEmail=TempRequestorEmail&";"&FormattedPhoneNumber&"@"&tempUserPhoneCarrier
                                                        End Select

                                                        RequestorName=TempRequestorName
												        RequestorPhoneNumber=TempRequestorPhone
                                                        RequestorEmailAddress=TempRequestorEmail
                                                    End if
                                                'End if
							

									Set oConn=Nothing
                                    %>
                                    <input type="hidden" name="PreExistingRequestor" value="<%=PreExistingRequestor %>" />
                                    <%
                        End if
                        else
                        If trim(PreExistingRequestor)="" then
        'response.write   "XXX GOT HERE89898989898989!!!<BR>"
									'If trim(RequestorFirstName)="" then
                                        Set oConn = Server.CreateObject("ADODB.Connection")
									    oConn.ConnectionTimeout = 100
									    oConn.Provider = "MSDASQL"
									    oConn.Open DATABASE
										    l_cSQL = "Select * FROM UserList WHERE UserID='"& FXCourierUserID &"' and Userstatus='c'"
										    'response.write "l_cSQL="&l_cSQL&"<BR>"
                                            SET oRs = oConn.Execute(l_cSql)
												    If not oRs.EOF then
                                                        'response.write  "GOT HERE!!!!<BR>"
                                                        'tempRequestorID=trim(oRs("RequestorID"))
                                                        PreExistingRequestor=trim(oRs("UserID"))
                                                        tempRequestorUserID=trim(oRs("UserID"))
												        tempRequestorFirstName=trim(oRs("UserFirstName"))
							                            tempRequestorLastName=trim(oRs("UserLastName"))
                                                        tempRequestorAddress=trim(oRs("UserAddress"))
                                                        tempRequestorCity=trim(oRs("UserCity"))
												        tempRequestorState=trim(oRs("UserState"))
							                            tempRequestorZipCode=trim(oRs("UserZipCode"))
                                                        tempRequestorPhoneNumber=trim(oRs("UserPhoneNumber"))
                                                        tempRequestorPhoneCarrier=trim(oRs("UserPhoneCarrier"))
												        tempRequestorEmailAddress=trim(oRs("UserEmailAddress"))
							                            tempRequestorUsername=trim(oRs("UserName"))
                                                        tempRequestorPassword=trim(oRs("Password"))
                                                        tempRequestorDeliveryNotifications=trim(oRs("DeliveryNotifications"))
                                                        'Response.write "test Got Here #2<BR>"
												        tempRequestorSupervisor=trim(oRs("Supervisor"))
							                            tempRequestorUserStatus=trim(oRs("UserStatus"))
                                                        tempRequestorDateCreated=trim(oRs("DateCreated"))
                                                        formattedphonenumber=Replace(tempRequestorPhone," ","")
                                                        formattedphonenumber=Replace(formattedphonenumber,"-","")
                                                        formattedphonenumber=Replace(formattedphonenumber,".","")
                                                        formattedphonenumber=Replace(formattedphonenumber,"(","")
                                                        formattedphonenumber=Replace(formattedphonenumber,")","")
                                                        Select Case DeliveryNotifications
                                                            Case "email"
                                                                TempRequestorEmail=TempRequestorEmail
                                                            Case "text"
                                                                TempRequestorEmail=FormattedPhoneNumber&"@"&tempUserPhoneCarrier
                                                            Case "both"
                                                                TempRequestorEmail=TempRequestorEmail&";"&FormattedPhoneNumber&"@"&tempUserPhoneCarrier
                                                        End Select

                                                        RequestorName=TempRequestorName
												        RequestorPhoneNumber=TempRequestorPhone
                                                        RequestorEmailAddress=TempRequestorEmail
                                                    End if
                                                'End if
							

									Set oConn=Nothing
                                    %>
                                    <input type="hidden" name="PreExistingRequestor" value="<%=PreExistingRequestor %>" />
                                    <%
                            else
                            'response.write  "yo, yo, yo...I got here!<BR>"
                            ''''''''Already selected a Requestor, so base info on form only!
                            TempRequestorUserID=Request.form("PreExistingRequestor")
                            PreExistingRequestor=Request.form("PreExistingRequestor")
                            TempRequestorFirstName=Request.form("RequestorFirstName")
                            TempRequestorLastName=Request.form("RequestorLastName")
                            TempRequestorPhoneNumber=Request.form("RequestorPhoneNumber")
                            TempRequestorEmailAddress=Request.form("RequestorEmailAddress")

                            TempRequestorAddress=Request.form("RequestorAddress")
                            TempRequestorCity=Request.form("RequestorCity")
                            TempRequestorZipCode=Request.form("RequestorZipCode")
                            TempRequestorSTate=Request.form("RequestorSTate")
                            TempRequestorPhoneCarrier=Request.form("celcarrier")
                            TempRequestorName=Request.form("RequestorName")
                            TempRequestorPassword=Request.form("RequestorPassword")
                            TempRequestorDeliveryNotifications=Request.form("DeliveryNotifications")
                            'response.write  "test XXX GOT HERE #3<BR>"
                                        %>
                                        <input type="hidden" name="PreExistingRequestor" value="<%=PreExistingRequestor %>" />
                                        <%
                        End if

        End if 
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''WHOLE NEW SECTION'''''''USER INFO, INSTEAD OF REQUESTOR INFO!!!!
        %>
        <td class="FleetExpressTextBlackBold" align="left">
            <font color="red">*</font>First Name:</td>
            <td width="10"><img src="../images/pixel.gif" /></td>
            <td align="left" class="FleetExpressTextBlack"><input type="text" name="RequestorFirstName" value="<%=TempRequestorFirstName%>" class="textbox" />
         </td>
    </tr>
    <tr>
         <td class="FleetExpressTextBlackBold" align="left">           
            <font color="red">*</font>Last Name:</td>
            <td width="10"><img src="../images/pixel.gif" /></td>
            <td align="left" class="FleetExpressTextBlack"><input type="text" name="RequestorLastName" value="<%=TempRequestorLastName%>" class="textbox" />
          </td>
    </tr>
    <tr>       
         <td class="FleetExpressTextBlackBold" align="left">           
            <font color="red">*</font>Password:</td>
            <td width="10"><img src="../images/pixel.gif" /></td>
            <td align="left" class="FleetExpressTextBlack"><input type="Password" name="RequestorPassword" value="<%=TempRequestorPassword %>" class="textbox" />
          </td>
    </tr>
    <tr>       
         <td class="FleetExpressTextBlackBold" align="left">           
            <font color="red">*</font>Mailing Address:</td>
            <td width="10"><img src="../images/pixel.gif" /></td>
            <td align="left" class="FleetExpressTextBlack"><input type="text" name="RequestorAddress" value="<%=TempRequestorAddress%>" class="textbox" />
         </td>
    </tr>
    <tr>        
        <td class="FleetExpressTextBlackBold" align="left">           
            <font color="red">*</font>City:</td>
            <td width="10"><img src="../images/pixel.gif" /></td>
            <td align="left" class="FleetExpressTextBlack"><input type="text" name="RequestorCity" value="<%=TempRequestorCity%>" class="textbox" />
            <input type="hidden" name="State" value="TX" />
         </td>
    </tr>
    <tr>         
         <td class="FleetExpressTextBlackBold" align="left">           
            <font color="red">*</font>Zip Code:</td>
            <td width="10"><img src="../images/pixel.gif" /></td>
            <td align="left" class="FleetExpressTextBlack"><input type="text" name="RequestorZipCode" value="<%=TempRequestorZipCode%>" class="textbox" />
         </td>
    </tr>
    <tr>         
         <td class="FleetExpressTextBlackBold" align="left">           
            <font color="red">*</font>Email Address:</td>
            <td width="10"><img src="../images/pixel.gif" /></td>
            <td align="left" class="FleetExpressTextBlack"><input type="text" name="RequestorEmailAddress" value="<%=TempRequestorEmailAddress%>" class="textbox" />
         </td>
    </tr>
    <!--
    <tr>         
         <td class="FleetExpressTextBlackBold" align="left" nowrap>           
            Re-type Email Address:</td>
            <td width="10"><img src="../images/pixel.gif" /></td>
            <td align="left" class="FleetExpressTextBlack"><input type="text" name="RetypeEmailAddress" value="<%=TempRequestorRetypeEmailAddress%>" class="textbox" />
         </td>
    </tr>
    -->
    <tr>         
         <td class="FleetExpressTextBlackBold" align="left">    
            <font color="red">*</font>Cell Phone Number:</td>
            <td width="10"><img src="../images/pixel.gif" /></td>
            <td align="left" class="FleetExpressTextBlack"><input type="text" name="RequestorPhoneNumber" value="<%=TempRequestorPhoneNumber%>" class="textbox" />
         </td>
    </tr>
    <tr>        
         <td class="FleetExpressTextBlackBold" align="left">  
            <font color="red">*</font>Cell Phone Carrier: </td>
            <td width="10"><img src="../images/pixel.gif" /></td>
            <td align="left" class="FleetExpressTextBlack">
            <select name="celcarrier" class="textbox">
                <option value="">Select your carrier</option>
                <option value="txt.att.net" <%If trim(TempRequestorPhoneCarrier)="txt.att.net" then response.write "Selected" end if %>>AT&T</option>
                <option value="messaging.nextel.com" <%If trim(TempRequestorPhoneCarrier)="messaging.nextel.com" then response.write "Selected" end if %>>Nextel</option>
                <option value="messaging.sprintpcs.com" <%If trim(TempRequestorPhoneCarrier)="messaging.sprintpcs.com" then response.write "Selected" end if %>>Sprint</option>
                <option value="tmomail.net" <%If trim(TempRequestorPhoneCarrier)="tmomail.net" then response.write "Selected" end if %>>T-Mobile</option>
                <option value="email.uscc.net" <%If trim(TempRequestorPhoneCarrier)="email.uscc.net" then response.write "Selected" end if %>>US Cellular</option>
                <option value="vtext.com" <%If trim(TempRequestorPhoneCarrier)="vtext.com" then response.write "Selected" end if %>>Verizon</option>
                <option value="vmobl.com" <%If trim(TempRequestorPhoneCarrier)="vmobl.com" then response.write "Selected" end if %>>Virgin</option>
            </select>
        </td>
    </tr>
    <tr>        
         <td class="FleetExpressTextBlackBold" align="left"> 
            <font color="red">*</font>Notify Me By: </td>
            <td width="10"><img src="../images/pixel.gif" /></td>
            <td align="left" class="FleetExpressTextBlack">
            <select name="DeliveryNotifications" class="textbox">
                <option value="email" <%If trim(TempRequestorDeliveryNotifications)="email" then response.write "Selected" end if %>>Email</option>
                <option value="text" <%If trim(TempRequestorDeliveryNotifications)="text" then response.write "Selected" end if %>>Text Message</option>
                <option value="both" <%If trim(TempRequestorDeliveryNotifications)="both" then response.write "Selected" end if %>>Email and Text Message</option>
            </select>
        </td>
    </tr>

        <tr>




        
        <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
        <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
        <tr><td colspan="4" bgcolor="black"><img src="../images/pixel.gif" height="1" width="1" /></td></tr>
        <tr><td><img src="../images/pixel.gif" height="2" width="1" /></td></tr>
                 <tr>
                 <td valign="top" rowspan="11"  class="OrderHeader">COMMODITY INFORMATION<img src="../images/pixel.gif" height="1" width="15" /></td>
                        <td class="FleetExpressTextBlackBold" align="left"><font color="red">*</font>Control/Part Number</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                         <td align="left">
                            <input type="text" class="textbox"  value="<%=PartNumber%>" name="PartNumber" maxlength="20" />
                        </td>
                    </tr>
                    <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left"><font color="red">*</font>Number of Pieces</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td nowrap class="FleetExpressTextBlackBold">
                             <select name="pieces" class="textbox" >
                                <option value="1" <%If trim(pieces)="1" then Response.write "selected" end if%>>1</option>
                                <option value="2" <%If trim(pieces)="2" then Response.write "selected" end if%>>2</option>
                                <option value="3" <%If trim(pieces)="3" then Response.write "selected" end if%>>3</option>
                                <option value="4" <%If trim(pieces)="4" then Response.write "selected" end if%>>4</option>
                                <option value="5" <%If trim(pieces)="5" then Response.write "selected" end if%>>5</option>
                                <!--
                                <option value="6" <%If trim(pieces)="6" then Response.write "selected" end if%>>6</option>
                                <option value="7" <%If trim(pieces)="7" then Response.write "selected" end if%>>7</option>
                                <option value="8" <%If trim(pieces)="8" then Response.write "selected" end if%>>8</option>
                                <option value="9" <%If trim(pieces)="9" then Response.write "selected" end if%>>9</option>
                                <option value="10" <%If trim(pieces)="10" then Response.write "selected" end if%>>10</option>
                                <option value="11" <%If trim(pieces)="11" then Response.write "selected" end if%>>11</option>
                                <option value="12" <%If trim(pieces)="12" then Response.write "selected" end if%>>12</option>
                                -->
                            </select>                           
                            &nbsp;&nbsp;
                            <select name="rf_box" class="textbox" >
                                <option value="Boxes"<%If trim(rf_box)="Boxes" then Response.write "selected" end if%>>Boxes</option>
                                <option value="Envelopes"<%If trim(rf_box)="Envelopes" then Response.write "selected" end if%>>Envelopes</option>
                            </select>
                        </td>
                    </tr>
                    <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                    <%'Response.write "DimWeight="&DimWeight&"<BR>" %>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left"><font color="red">*</font>Weight</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack">
                            <select name="DimWeight" class="textbox" >
                                <option value="1" <%If trim(dimweight)="1" then Response.write "selected" end if%>>1</option>
                                <option value="2" <%If trim(dimweight)="2" then Response.write "selected" end if%>>2</option>
                                <option value="3" <%If trim(dimweight)="3" then Response.write "selected" end if%>>3</option>
                                <option value="4" <%If trim(dimweight)="4" then Response.write "selected" end if%>>4</option>
                                <option value="5" <%If trim(dimweight)="5" then Response.write "selected" end if%>>5</option>
                                <option value="6" <%If trim(dimweight)="6" then Response.write "selected" end if%>>6</option>
                                <option value="7" <%If trim(dimweight)="7" then Response.write "selected" end if%>>7</option>
                                <option value="8" <%If trim(dimweight)="8" then Response.write "selected" end if%>>8</option>
                                <option value="9" <%If trim(dimweight)="9" then Response.write "selected" end if%>>9</option>
                                <option value="10" <%If trim(dimweight)="10" then Response.write "selected" end if%>>10</option>
                                <option value="11" <%If trim(dimweight)="11" then Response.write "selected" end if%>>11</option>
                                <option value="12" <%If trim(dimweight)="12" then Response.write "selected" end if%>>12</option>
                                <option value="13" <%If trim(dimweight)="13" then Response.write "selected" end if%>>13</option>
                                <option value="14" <%If trim(dimweight)="14" then Response.write "selected" end if%>>14</option>
                                <option value="15" <%If trim(dimweight)="15" then Response.write "selected" end if%>>15</option>
                                <option value="16" <%If trim(dimweight)="16" then Response.write "selected" end if%>>13</option>
                                <option value="17" <%If trim(dimweight)="17" then Response.write "selected" end if%>>17</option>
                                <option value="18" <%If trim(dimweight)="18" then Response.write "selected" end if%>>18</option>
                                <option value="19" <%If trim(dimweight)="19" then Response.write "selected" end if%>>19</option>
                                <option value="20" <%If trim(dimweight)="20" then Response.write "selected" end if%>>20</option>
                                <option value="21" <%If trim(dimweight)="21" then Response.write "selected" end if%>>21</option>
                                <option value="22" <%If trim(dimweight)="22" then Response.write "selected" end if%>>22</option>
                                <option value="23" <%If trim(dimweight)="23" then Response.write "selected" end if%>>23</option>
                                <option value="24" <%If trim(dimweight)="24" then Response.write "selected" end if%>>24</option>
                                <option value="25" <%If trim(dimweight)="25" then Response.write "selected" end if%>>25 (Maximum)</option>
                            </select>
                        Pound(s)</td>
                    </tr>
                    <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left"><font color="red">*</font>Dimensions (50" Max)</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td class="FleetExpressTextBlack" align="left" nowrap>
                            L:&nbsp;
                            <select name="DimLength" class="textbox" >
                                <option value="1" <%If trim(DimLength)="1" then Response.write "selected" end if%>>1</option>
                                <option value="2" <%If trim(DimLength)="2" then Response.write "selected" end if%>>2</option>
                                <option value="3" <%If trim(DimLength)="3" then Response.write "selected" end if%>>3</option>
                                <option value="4" <%If trim(DimLength)="4" then Response.write "selected" end if%>>4</option>
                                <option value="5" <%If trim(DimLength)="5" then Response.write "selected" end if%>>5</option>
                                <option value="6" <%If trim(DimLength)="6" then Response.write "selected" end if%>>6</option>
                                <option value="7" <%If trim(DimLength)="7" then Response.write "selected" end if%>>7</option>
                                <option value="8" <%If trim(DimLength)="8" then Response.write "selected" end if%>>8</option>
                                <option value="9" <%If trim(DimLength)="9" then Response.write "selected" end if%>>9</option>
                                <option value="10" <%If trim(DimLength)="10" then Response.write "selected" end if%>>10</option>
                                <option value="11" <%If trim(DimLength)="11" then Response.write "selected" end if%>>11</option>
                                <option value="12" <%If trim(DimLength)="12" then Response.write "selected" end if%>>12</option>
                                <option value="13" <%If trim(DimLength)="13" then Response.write "selected" end if%>>13</option>
                                <option value="14" <%If trim(DimLength)="14" then Response.write "selected" end if%>>14</option>
                                <option value="15" <%If trim(DimLength)="15" then Response.write "selected" end if%>>15</option>
                                <option value="16" <%If trim(DimLength)="16" then Response.write "selected" end if%>>16</option>
                                <option value="17" <%If trim(DimLength)="17" then Response.write "selected" end if%>>17</option>
                                <option value="18" <%If trim(DimLength)="18" then Response.write "selected" end if%>>18</option>
                                <option value="19" <%If trim(DimLength)="19" then Response.write "selected" end if%>>19</option>
                                <option value="20" <%If trim(DimLength)="20" then Response.write "selected" end if%>>20</option>
                                <option value="21" <%If trim(DimLength)="21" then Response.write "selected" end if%>>21</option>
                                <option value="22" <%If trim(DimLength)="22" then Response.write "selected" end if%>>22</option>
                                <option value="23" <%If trim(DimLength)="23" then Response.write "selected" end if%>>23</option>
                                <option value="24" <%If trim(DimLength)="24" then Response.write "selected" end if%>>24</option>
                                <option value="25" <%If trim(DimLength)="25" then Response.write "selected" end if%>>25</option>
                                <option value="26" <%If trim(DimLength)="26" then Response.write "selected" end if%>>26</option>
                                <option value="27 <%If trim(DimLength)="27" then Response.write "selected" end if%>>27</option>
                                <option value="28" <%If trim(DimLength)="28" then Response.write "selected" end if%>>28</option>
                                <option value="29" <%If trim(DimLength)="29" then Response.write "selected" end if%>>29</option>
                                <option value="30" <%If trim(DimLength)="30" then Response.write "selected" end if%>>30</option>
                                <option value="31" <%If trim(DimLength)="31" then Response.write "selected" end if%>>31</option>
                                <option value="32" <%If trim(DimLength)="32" then Response.write "selected" end if%>>32</option>
                                <option value="33" <%If trim(DimLength)="33" then Response.write "selected" end if%>>33</option>
                                <option value="34" <%If trim(DimLength)="34" then Response.write "selected" end if%>>34</option>
                                <option value="35" <%If trim(DimLength)="35" then Response.write "selected" end if%>>35</option>
                                <option value="36" <%If trim(DimLength)="36" then Response.write "selected" end if%>>36</option>
                                <option value="37" <%If trim(DimLength)="37" then Response.write "selected" end if%>>37</option>
                                <option value="38" <%If trim(DimLength)="38" then Response.write "selected" end if%>>38</option>
                                <option value="39" <%If trim(DimLength)="39" then Response.write "selected" end if%>>39</option>
                                <option value="40" <%If trim(DimLength)="40" then Response.write "selected" end if%>>40</option>
                                <option value="41" <%If trim(DimLength)="42" then Response.write "selected" end if%>>41</option>
                                <option value="42" <%If trim(DimLength)="42" then Response.write "selected" end if%>>42</option>
                                <option value="43" <%If trim(DimLength)="43" then Response.write "selected" end if%>>43</option>
                                <option value="44" <%If trim(DimLength)="44" then Response.write "selected" end if%>>44</option>
                                <option value="45" <%If trim(DimLength)="45" then Response.write "selected" end if%>>45</option>
                                <option value="46" <%If trim(DimLength)="46" then Response.write "selected" end if%>>46</option>
                                <option value="47" <%If trim(DimLength)="47" then Response.write "selected" end if%>>47</option>
                                <option value="48" <%If trim(DimLength)="48" then Response.write "selected" end if%>>48</option>
                            </select>
                            W:&nbsp;
                             <select name="DimWidth" class="textbox" >
                                <option value="1" <%If trim(DimWidth)="1" then Response.write "selected" end if%>>1</option>
                                <option value="2" <%If trim(DimWidth)="2" then Response.write "selected" end if%>>2</option>
                                <option value="3" <%If trim(DimWidth)="3" then Response.write "selected" end if%>>3</option>
                                <option value="4" <%If trim(DimWidth)="4" then Response.write "selected" end if%>>4</option>
                                <option value="5" <%If trim(DimWidth)="5" then Response.write "selected" end if%>>5</option>
                                <option value="6" <%If trim(DimWidth)="6" then Response.write "selected" end if%>>6</option>
                                <option value="7" <%If trim(DimWidth)="7" then Response.write "selected" end if%>>7</option>
                                <option value="8" <%If trim(DimWidth)="8" then Response.write "selected" end if%>>8</option>
                                <option value="9" <%If trim(DimWidth)="9" then Response.write "selected" end if%>>9</option>
                                <option value="10" <%If trim(DimWidth)="10" then Response.write "selected" end if%>>10</option>
                                <option value="11" <%If trim(DimWidth)="11" then Response.write "selected" end if%>>11</option>
                                <option value="12" <%If trim(DimWidth)="12" then Response.write "selected" end if%>>12</option>
                                <option value="13" <%If trim(DimWidth)="13" then Response.write "selected" end if%>>13</option>
                                <option value="14" <%If trim(DimWidth)="14" then Response.write "selected" end if%>>14</option>
                                <option value="15" <%If trim(DimWidth)="15" then Response.write "selected" end if%>>15</option>
                                <option value="16" <%If trim(DimWidth)="16" then Response.write "selected" end if%>>16</option>
                                <option value="17" <%If trim(DimWidth)="17" then Response.write "selected" end if%>>17</option>
                                <option value="18" <%If trim(DimWidth)="18" then Response.write "selected" end if%>>18</option>
                                <option value="19" <%If trim(DimWidth)="19" then Response.write "selected" end if%>>19</option>
                                <option value="20" <%If trim(DimWidth)="20" then Response.write "selected" end if%>>20</option>
                                <option value="21" <%If trim(DimWidth)="21" then Response.write "selected" end if%>>21</option>
                                <option value="22" <%If trim(DimWidth)="22" then Response.write "selected" end if%>>22</option>
                                <option value="23" <%If trim(DimWidth)="23" then Response.write "selected" end if%>>23</option>
                                <option value="24" <%If trim(DimWidth)="24" then Response.write "selected" end if%>>24</option>
                                <option value="25" <%If trim(DimWidth)="25" then Response.write "selected" end if%>>25</option>
                                <option value="26" <%If trim(DimWidth)="26" then Response.write "selected" end if%>>26</option>
                                <option value="27 <%If trim(DimWidth)="27" then Response.write "selected" end if%>>27</option>
                                <option value="28" <%If trim(DimWidth)="28" then Response.write "selected" end if%>>28</option>
                                <option value="29" <%If trim(DimWidth)="29" then Response.write "selected" end if%>>29</option>
                                <option value="30" <%If trim(DimWidth)="30" then Response.write "selected" end if%>>30</option>
                                <option value="31" <%If trim(DimWidth)="31" then Response.write "selected" end if%>>31</option>
                                <option value="32" <%If trim(DimWidth)="32" then Response.write "selected" end if%>>32</option>
                                <option value="33" <%If trim(DimWidth)="33" then Response.write "selected" end if%>>33</option>
                                <option value="34" <%If trim(DimWidth)="34" then Response.write "selected" end if%>>34</option>
                                <option value="35" <%If trim(DimWidth)="35" then Response.write "selected" end if%>>35</option>
                                <option value="36" <%If trim(DimWidth)="36" then Response.write "selected" end if%>>36</option>
                                <option value="37" <%If trim(DimWidth)="37" then Response.write "selected" end if%>>37</option>
                                <option value="38" <%If trim(DimWidth)="38" then Response.write "selected" end if%>>38</option>
                                <option value="39" <%If trim(DimWidth)="39" then Response.write "selected" end if%>>39</option>
                                <option value="40" <%If trim(DimWidth)="40" then Response.write "selected" end if%>>40</option>
                                <option value="41" <%If trim(DimWidth)="42" then Response.write "selected" end if%>>41</option>
                                <option value="42" <%If trim(DimWidth)="42" then Response.write "selected" end if%>>42</option>
                                <option value="43" <%If trim(DimWidth)="43" then Response.write "selected" end if%>>43</option>
                                <option value="44" <%If trim(DimWidth)="44" then Response.write "selected" end if%>>44</option>
                                <option value="45" <%If trim(DimWidth)="45" then Response.write "selected" end if%>>45</option>
                                <option value="46" <%If trim(DimWidth)="46" then Response.write "selected" end if%>>46</option>
                                <option value="47" <%If trim(DimWidth)="47" then Response.write "selected" end if%>>47</option>
                                <option value="48" <%If trim(DimWidth)="48" then Response.write "selected" end if%>>48</option>
                            </select>                           
                           H:&nbsp;
                            <select name="DimHeight" class="textbox" >
                                <option value="1" <%If trim(DimHeight)="1" then Response.write "selected" end if%>>1</option>
                                <option value="2" <%If trim(DimHeight)="2" then Response.write "selected" end if%>>2</option>
                                <option value="3" <%If trim(DimHeight)="3" then Response.write "selected" end if%>>3</option>
                                <option value="4" <%If trim(DimHeight)="4" then Response.write "selected" end if%>>4</option>
                                <option value="5" <%If trim(DimHeight)="5" then Response.write "selected" end if%>>5</option>
                                <option value="6" <%If trim(DimHeight)="6" then Response.write "selected" end if%>>6</option>
                                <option value="7" <%If trim(DimHeight)="7" then Response.write "selected" end if%>>7</option>
                                <option value="8" <%If trim(DimHeight)="8" then Response.write "selected" end if%>>8</option>
                                <option value="9" <%If trim(DimHeight)="9" then Response.write "selected" end if%>>9</option>
                                <option value="10" <%If trim(DimHeight)="10" then Response.write "selected" end if%>>10</option>
                                <option value="11" <%If trim(DimHeight)="11" then Response.write "selected" end if%>>11</option>
                                <option value="12" <%If trim(DimHeight)="12" then Response.write "selected" end if%>>12</option>
                                <option value="13" <%If trim(DimHeight)="13" then Response.write "selected" end if%>>13</option>
                                <option value="14" <%If trim(DimHeight)="14" then Response.write "selected" end if%>>14</option>
                                <option value="15" <%If trim(DimHeight)="15" then Response.write "selected" end if%>>15</option>
                                <option value="16" <%If trim(DimHeight)="16" then Response.write "selected" end if%>>16</option>
                                <option value="17" <%If trim(DimHeight)="17" then Response.write "selected" end if%>>17</option>
                                <option value="18" <%If trim(DimHeight)="18" then Response.write "selected" end if%>>18</option>
                                <option value="19" <%If trim(DimHeight)="19" then Response.write "selected" end if%>>19</option>
                                <option value="20" <%If trim(DimHeight)="20" then Response.write "selected" end if%>>20</option>
                                <option value="21" <%If trim(DimHeight)="21" then Response.write "selected" end if%>>21</option>
                                <option value="22" <%If trim(DimHeight)="22" then Response.write "selected" end if%>>22</option>
                                <option value="23" <%If trim(DimHeight)="23" then Response.write "selected" end if%>>23</option>
                                <option value="24" <%If trim(DimHeight)="24" then Response.write "selected" end if%>>24</option>
                                <option value="25" <%If trim(DimHeight)="25" then Response.write "selected" end if%>>25</option>
                                <option value="26" <%If trim(DimHeight)="26" then Response.write "selected" end if%>>26</option>
                                <option value="27 <%If trim(DimHeight)="27" then Response.write "selected" end if%>>27</option>
                                <option value="28" <%If trim(DimHeight)="28" then Response.write "selected" end if%>>28</option>
                                <option value="29" <%If trim(DimHeight)="29" then Response.write "selected" end if%>>29</option>
                                <option value="30" <%If trim(DimHeight)="30" then Response.write "selected" end if%>>30</option>
                                <option value="31" <%If trim(DimHeight)="31" then Response.write "selected" end if%>>31</option>
                                <option value="32" <%If trim(DimHeight)="32" then Response.write "selected" end if%>>32</option>
                                <option value="33" <%If trim(DimHeight)="33" then Response.write "selected" end if%>>33</option>
                                <option value="34" <%If trim(DimHeight)="34" then Response.write "selected" end if%>>34</option>
                                <option value="35" <%If trim(DimHeight)="35" then Response.write "selected" end if%>>35</option>
                                <option value="36" <%If trim(DimHeight)="36" then Response.write "selected" end if%>>36</option>
                                <option value="37" <%If trim(DimHeight)="37" then Response.write "selected" end if%>>37</option>
                                <option value="38" <%If trim(DimHeight)="38" then Response.write "selected" end if%>>38</option>
                                <option value="39" <%If trim(DimHeight)="39" then Response.write "selected" end if%>>39</option>
                                <option value="40" <%If trim(DimHeight)="40" then Response.write "selected" end if%>>40</option>
                                <option value="41" <%If trim(DimHeight)="42" then Response.write "selected" end if%>>41</option>
                                <option value="42" <%If trim(DimHeight)="42" then Response.write "selected" end if%>>42</option>
                                <option value="43" <%If trim(DimHeight)="43" then Response.write "selected" end if%>>43</option>
                                <option value="44" <%If trim(DimHeight)="44" then Response.write "selected" end if%>>44</option>
                                <option value="45" <%If trim(DimHeight)="45" then Response.write "selected" end if%>>45</option>
                                <option value="46" <%If trim(DimHeight)="46" then Response.write "selected" end if%>>46</option>
                                <option value="47" <%If trim(DimHeight)="47" then Response.write "selected" end if%>>47</option>
                                <option value="48" <%If trim(DimHeight)="48" then Response.write "selected" end if%>>48</option>
                            </select>
                            &nbsp;Inches
                        </td>
                    </tr>
                    <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left"><font color="red">*</font>
                                <select name="POorNWA" class="textbox" >
                                    <option value="Cost Center #"<%If trim(POorNWA)="Cost Center #" then Response.write "selected" end if%>>Cost Center #</option>
                                    <option value="P/O #"<%If trim(POorNWA)="P/O #" then Response.write "selected" end if%>>P/O #</option>
                                </select>
                        </td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left"><input type="text" class="textbox"  maxlength="20" name="GenericNumber" value="<%=GenericNumber%>"></td>
                    </tr>
                    <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left" valign="top"><font color="red">*</font>Special Instructions<br />&nbsp;&nbsp;(ex. Hand cart required)</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left"><textarea name="comments" rows="2" cols="30" class="textbox" ><%=Comments%></textarea></td>
                    </tr>
                    <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                    <!--
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Hazmat</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left"  class="FleetExpressTextBlackBold">
                             <select name="IsHazmat">
                                <option value="n" <%If trim(IsHazmat)="n" then Response.write "selected" end if%>>No</option>
                                <option value="y" <%If trim(IsHazmat)="y" then Response.write "selected" end if%>>Yes</option>
                            </select>                     
                        </td>
                    </tr>
                    -->
                    <!--
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Refrigerate</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left"  class="FleetExpressTextBlackBold">
                             <select name="Refrigerate">
                                <option value="n" <%If trim(Refrigerate)="n" then Response.write "selected" end if%>>No</option>
                                <option value="y" <%If trim(Refrigerate)="y" then Response.write "selected" end if%>>Yes</option>
                            </select>                      
                        </td>
                    </tr>
                    -->
                    <!--
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Service Level</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left"  class="FleetExpressTextBlackBold">
                             <select name="Priority" onchange="AnyCost();">
                                <option value="Next Day" <%If trim(Priority)="Next Day" then Response.write "selected" end if%>>Next Day</option>
                                <option value="Same Day" <%If trim(Priority)="Same Day" then Response.write "selected" end if%>>Same Day</option>
                                <option value="Time Critical" <%If trim(Priority)="Time Critical" then Response.write "selected" end if%>>Time Critical</option>
                            </select>                      
                        </td>
                    </tr>
                    -->
            </td>
        </tr>

        <%
	    'Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
	    '	RSEVENTS.CursorLocation = 3
	    '	RSEVENTS.CursorType = 3
	    '	RSEVENTS.ActiveConnection = DATABASE
	    '	SQL = "SELECT PriorityMinutes FROM Priorities where (priority = '"&priority&"') AND (Priority_BT_ID='"&BillToID&"') AND (PriorityOrigination='"&st_id&"') AND (PriorityDestination='"&Destination&"')"
	    '	'Response.Write "SQL="&SQL&"<BR>"
	    '	RSEVENTS.Open SQL, DATABASE, 1, 3
	    '	If Not RSEVENTS.EOF then
	    '	    'Response.Write "GOT HERE 1!<br>"
	    '       PriorityTime=RSEVENTS("PriorityMinutes")
	    '     Response.Write "PriorityTime1111111="&PriorityTime&"<BR>"
	    '   End if
	    '	RSEVENTS.close
	    'Set RSEVENTS = Nothing  
        'Response.write "PreExistingRequestor="&PreExistingRequestor&"<BR>"      
         %>
                <input type="Hidden" value="standard" name="Priority" />

                <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                <tr><td colspan="4" bgcolor="black"><img src="../images/pixel.gif" height="1" width="1" /></td></tr>
                <tr><td><img src="../images/pixel.gif" height="2" width="1" /></td></tr>
                <tr>
                    <td valign="top" rowspan="19"  class="OrderHeader">ORIGINATION INFORMATION<img src="../images/pixel.gif" height="1" width="15" /></td>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap><font color="red">&nbsp;&nbsp;</font>Pre-Existing Location</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left">
								<%
                                'response.write  "PreExistingOrigination="&PreExistingOrigination&"<BR>"
                                'response.write  "CompanyID="&CompanyID&"<BR>"
                                'response.write  "TempPreExistingOrigination="&TempPreExistingOrigination&"<BR>"
                                 %>
                                <select name="PreExistingOrigination" ID="Select3"  onChange="form.submit()" class="textbox" >
								<option value="" <%if trim(PreExistingOrigination)="" then response.Write " selected" end if%>>Select from this list or fill in all fields below</option>
                                <%
                                    
									Set oConn = Server.CreateObject("ADODB.Connection")
									oConn.ConnectionTimeout = 100
									oConn.Provider = "MSDASQL"
									oConn.Open DATABASE
										l_cSQL = "Select * FROM PreExistingCompanies WHERE companystatus='c'"
                                         'If Supervisor<>"y" or isNULL(Supervisor) then
                                            l_cSQL = l_cSQL & " AND CompanyOwner='"& PreExistingRequestor &"'"
                                            if PreExistingRequestor="90xxx" then
                                                l_cSQL = l_cSQL & " AND CompanyID = '4705'"
                                            End if
                                            'else
                                            'l_cSQL = l_cSQL & " AND CompanyOwner is NULL"
                                        'End if                                       
                                        l_cSQL = l_cSQL & " ORDER BY CompanyName"
                                        'Response.write "l_cSQL="&l_cSQL&"<BR>"
                                       
										SET oRs = oConn.Execute(l_cSql)
												Do While not oRs.EOF
                                                DisplayCompanyID=oRs("CompanyID")
                                                tempCompanySuite=oRs("CompanySuite")
                                                    If trim(tempCompanySuite)>"" then
                                                        DisplayOriginationSuite="/"&tempCompanySuite
                                                        else
                                                        DisplayOriginationSuite=""
                                                    End if
                                                    DisplayCompanyName=oRs("CompanyName")
                                                If Trim(DisplayCompanyID)=Trim(PreExistingOrigination) then
                                                CompanyID=oRs("CompanyID")
												CompanyName=oRs("CompanyName")
												CompanyAddress=oRs("CompanyAddress")
                                                CompanyBuilding=oRs("CompanyBuilding")
                                                CompanySuite=oRs("CompanySuite")
                                                    'If trim(CompanySuite)>"" then
                                                    '    DisplayOriginationSuite="/"&CompanySuite
                                                    '    else
                                                    '    DisplayOriginationSuite=""
                                                    'End if
												CompanyCity=oRs("CompanyCity")
                                                CompanyState=oRs("CompanyState")
                                                CompanyZip=oRs("CompanyZip")
                                                ContactName=oRs("ContactName")
                                                CompanyPhone=oRs("CompanyPhone")
                                                CompanyEmail=oRs("CompanyEmail")
                                                End if


                                                If trim(PreExistingOrigination)<>trim(TempPreExistingOrigination) then

                                                'If trim(PreExistingOrigination)="" then
												   'response.write "GOT HERE!!!!<BR>"
                                                    OriginationCompany=CompanyName
												    OriginationAddress=CompanyAddress
                                                    OriginationBuilding=CompanyBuilding
                                                    OriginationSuite=CompanySuite
                                                    'If trim(OriginationSuite)>"" then
                                                    '    DisplayOriginationSuite="/"&CompanySuite
                                                    '    else
                                                    '    DisplayOriginationSuite=""
                                                    'End if
												    OriginationCity=CompanyCity
                                                    OriginationState=CompanyState
                                                    OriginationZipCode=CompanyZip
                                                    tempOriginationContactName=ContactName
                                                    tempOriginationPhoneNumber=CompanyPhone
                                                    tempOriginationEmail=CompanyEmail
                                                    If trim(tempOriginationContactName)>"" then OriginationContactName=tempOriginationContactName end if
                                                    If trim(tempOriginationPhoneNumber)>"" then OriginationPhoneNumber=tempOriginationPhoneNumber end if
                                                    If trim(tempOriginationEmail)>"" then OriginationEmail=tempOriginationEmail end if
                                                End if
                                                If Trim(DisplayCompanyName)="" then DisplayOriginationSuite=Replace(DisplayOriginationSuite,"/","") end if									
											%>
											<option value="<%=DisplayCompanyID%>" <%if trim(PreExistingOrigination)=trim(DisplayCompanyID) then response.Write " selected" end if%>><%=OriginationBuilding%><%=DisplayOriginationSuite%></option>
											<%
										oRs.movenext
										LOOP
									Set oConn=Nothing
                                    
                                   									
									%>
								</select> 
                    </td>
                </tr>
                <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap><font color="red">&nbsp;&nbsp;</font>Company Name</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left"><input type="text" class="textbox"  name="OriginationCompany" value="<%=OriginationCompany%>" size="45" maxlength="40" /></td>
                </tr>
                <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left"><font color="red">&nbsp;&nbsp;</font>Address</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left"><input type="text" class="textbox"  name="OriginationAddress" value="<%=OriginationAddress%>" size="45" maxlength="40" /></td>
                </tr>
                <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left"><font color="red">*</font>Building</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left"><input type="text" class="textbox"  name="OriginationBuilding" value="<%=OriginationBuilding%>" size="45" maxlength="40" /></td>
                </tr>
                <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left"><font color="red">*</font>Floor/Cube Number</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left"><input type="text" class="textbox"  name="OriginationSuite" value="<%=OriginationSuite%>" size="45" maxlength="40" /></td>
                </tr>
                <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left"><font color="red">&nbsp;&nbsp;</font>City/State/Zip</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td class="FleetExpressTextBlackBold">
                       
                        <input type="text" class="textbox"  name="OriginationCity" value="<%=OriginationCity%>" size="20" maxlength="30" />
                         <!--
                        <select name="OriginationState">
	                        <option value="AL" <%If trim(OriginationState)="AL" then Response.write "selected" end if%>>AL</option>
	                        <option value="AK" <%If trim(OriginationState)="AK" then Response.write "selected" end if%>>AK</option>
	                        <option value="AZ <%If trim(OriginationState)="AZ" then Response.write "selected" end if%>">AZ</option>
	                        <option value="AR <%If trim(OriginationState)="AR" then Response.write "selected" end if%>">AR</option>
	                        <option value="CA" <%If trim(OriginationState)="CA" then Response.write "selected" end if%>>CA</option>
	                        <option value="CO" <%If trim(OriginationState)="CO" then Response.write "selected" end if%>>CO</option>
	                        <option value="CT" <%If trim(OriginationState)="CT" then Response.write "selected" end if%>>CT</option>
	                        <option value="DE" <%If trim(OriginationState)="DE" then Response.write "selected" end if%>>DE</option>
	                        <option value="DC" <%If trim(OriginationState)="DC" then Response.write "selected" end if%>>DC</option>
	                        <option value="FL" <%If trim(OriginationState)="FL" then Response.write "selected" end if%>>FL</option>
	                        <option value="GA" <%If trim(OriginationState)="GA" then Response.write "selected" end if%>>GA</option>
	                        <option value="HI" <%If trim(OriginationState)="HI" then Response.write "selected" end if%>>HI</option>
	                        <option value="ID" <%If trim(OriginationState)="ID" then Response.write "selected" end if%>>ID</option>
	                        <option value="IL" <%If trim(OriginationState)="IL" then Response.write "selected" end if%>>IL</option>
	                        <option value="IN" <%If trim(OriginationState)="IN" then Response.write "selected" end if%>>IN</option>
	                        <option value="IA" <%If trim(OriginationState)="IA" then Response.write "selected" end if%>>IA</option>
	                        <option value="KS" <%If trim(OriginationState)="KS" then Response.write "selected" end if%>>KS</option>
	                        <option value="KY" <%If trim(OriginationState)="KY" then Response.write "selected" end if%>>KY</option>
	                        <option value="LA" <%If trim(OriginationState)="LA" then Response.write "selected" end if%>>LA</option>
	                        <option value="ME" <%If trim(OriginationState)="ME" then Response.write "selected" end if%>>ME</option>
	                        <option value="MD" <%If trim(OriginationState)="MD" then Response.write "selected" end if%>>MD</option>
	                        <option value="MA" <%If trim(OriginationState)="MA" then Response.write "selected" end if%>>MA</option>
	                        <option value="MI" <%If trim(OriginationState)="MI" then Response.write "selected" end if%>>MI</option>
	                        <option value="MN" <%If trim(OriginationState)="MN" then Response.write "selected" end if%>>MN</option>
	                        <option value="MS" <%If trim(OriginationState)="MS" then Response.write "selected" end if%>>MS</option>
	                        <option value="MO" <%If trim(OriginationState)="MO" then Response.write "selected" end if%>>MO</option>
	                        <option value="MT" <%If trim(OriginationState)="MT" then Response.write "selected" end if%>>MT</option>
	                        <option value="NE" <%If trim(OriginationState)="NE" then Response.write "selected" end if%>>NE</option>
	                        <option value="NV" <%If trim(OriginationState)="NV" then Response.write "selected" end if%>>NV</option>
	                        <option value="NH" <%If trim(OriginationState)="NH" then Response.write "selected" end if%>>NH</option>
	                        <option value="NJ" <%If trim(OriginationState)="NJ" then Response.write "selected" end if%>>NJ</option>
	                        <option value="NM" <%If trim(OriginationState)="NM" then Response.write "selected" end if%>>NM</option>
	                        <option value="NY" <%If trim(OriginationState)="NY" then Response.write "selected" end if%>>NY</option>
	                        <option value="NC" <%If trim(OriginationState)="NC" then Response.write "selected" end if%>>NC</option>
	                        <option value="ND" <%If trim(OriginationState)="ND" then Response.write "selected" end if%>>ND</option>
	                        <option value="OH" <%If trim(OriginationState)="OH" then Response.write "selected" end if%>>OH</option>
	                        <option value="OK" <%If trim(OriginationState)="OK" then Response.write "selected" end if%>>OK</option>
	                        <option value="OR" <%If trim(OriginationState)="OR" then Response.write "selected" end if%>>OR</option>
	                        <option value="PA" <%If trim(OriginationState)="PA" then Response.write "selected" end if%>>PA</option>
	                        <option value="RI" <%If trim(OriginationState)="RI" then Response.write "selected" end if%>>RI</option>
	                        <option value="SC" <%If trim(OriginationState)="SC" then Response.write "selected" end if%>>SC</option>
	                        <option value="SD" <%If trim(OriginationState)="SD" then Response.write "selected" end if%>>SD</option>
	                        <option value="TN" <%If trim(OriginationState)="TN" then Response.write "selected" end if%>>TN</option>
	                        <option value="TX" <%If trim(OriginationState)="TX" or trim(OriginationState)="" then Response.write "selected" end if%>>TX</option>
	                        <option value="UT" <%If trim(OriginationState)="UT" then Response.write "selected" end if%>>UT</option>
	                        <option value="VT" <%If trim(OriginationState)="VT" then Response.write "selected" end if%>>VT</option>
	                        <option value="VA" <%If trim(OriginationState)="VA" then Response.write "selected" end if%>>VA</option>
	                        <option value="WA" <%If trim(OriginationState)="WA" then Response.write "selected" end if%>>WA</option>
	                        <option value="WV" <%If trim(OriginationState)="WV" then Response.write "selected" end if%>>WV</option>
	                        <option value="WI" <%If trim(OriginationState)="WI" then Response.write "selected" end if%>>WI</option>
	                        <option value="WY" <%If trim(OriginationState)="WY" then Response.write "selected" end if%>>WY</option>
                        </select>
                        -->&nbsp;TX&nbsp;&nbsp;
                        <input type="hidden" name="OriginationState" value="TX">
                        <input type="text" class="textbox"  name="OriginationZipCode" value="<%=OriginationZipCode%>" size="11" maxlength="10" />

                    </td>
                </tr>
                <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>

                <tr>
                    <td class="FleetExpressTextBlackBold" align="left"><font color="red">*</font>Contact Name</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left"><input type="text" class="textbox"  name="OriginationContactName" value="<%=OriginationContactName%>" size="45" maxlength="25" /></td>
                </tr>
                <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left"><font color="red">*</font>Phone Number</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left"><input type="text" class="textbox"  name="OriginationPhoneNumber" value="<%=OriginationPhoneNumber%>" size="45" maxlength="20" /></td>
                </tr> 
                <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap><font color="red">&nbsp;&nbsp;</font>Email Address</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left"><input type="text" class="textbox"  name="OriginationEmail" value="<%=OriginationEmail%>" size="45" maxlength="100" /></td>
                </tr>
                 <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left"><font color="red">*</font>Ready Date/Time</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left"><input type="text" class="textbox"  id="PickupDateTime" name="PickupDateTime" value="<%=PickUpDateTime%>" size="30" maxlength="30" />
                    <a href="javascript:NewCal('PickupDateTime','MMddyyyy',true,12)"><img src="../images/cal.gif" width="16" height="16" border="0" alt="Pick a date"></a>                 
                    </td>
                </tr>
                <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                <tr><td colspan="4" bgcolor="black"><img src="../images/pixel.gif" height="1" width="1" /></td></tr>
                <tr><td><img src="../images/pixel.gif" height="2" width="1" /></td></tr>
                <tr>
                    <td valign="top" rowspan="19"  class="OrderHeader">DESTINATION INFORMATION<img src="../images/pixel.gif" height="1" width="15" />
                                  
                    </td>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap><font color="red">&nbsp;&nbsp;</font>Pre-Existing Location</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left">
								<%
                                'REsponse.write "PreExistingDestination="&PreExistingDestination&"<BR>"
                                'REsponse.write "CompanyID="&CompanyID&"<BR>"
                                'response.write "TempPreExistingDestination="&TempPreExistingDestination&"<BR>"
                                 %>
                                <select name="PreExistingDestination" ID="Select1"  onChange="form.submit()" class="textbox" >
								<option value="" <%if trim(PreExistingDestination)="" then response.Write " selected" end if%>>Select from this list or fill in all fields below</option>
                                <%

									Set oConn = Server.CreateObject("ADODB.Connection")
									oConn.ConnectionTimeout = 100
									oConn.Provider = "MSDASQL"
									oConn.Open DATABASE
										l_cSQL = "Select * FROM PreExistingCompanies WHERE companystatus='c'"
                                         'If Supervisor<>"y" or isNULL(Supervisor) then
                                            l_cSQL = l_cSQL & " AND CompanyOwner='"& PreExistingRequestor &"'"
                                            'else
                                            'l_cSQL = l_cSQL & " AND CompanyOwner is NULL"
                                        'End if                                       
                                        l_cSQL = l_cSQL & " ORDER BY CompanyName"
										SET oRs = oConn.Execute(l_cSql)
												Do While not oRs.EOF
                                                DisplaybCompanyID=oRs("CompanyID")
												DisplaybCompanyName=oRs("CompanyName")                                                
                                                DisplaybCompanySuite=oRs("CompanySuite")
                                                    If trim(DisplaybCompanySuite)>"" then
                                                        DisplayDestinationSuite="/"&DisplaybCompanySuite
                                                        else
                                                        DisplayDestinationSuite=""
                                                    End if 
                                                                                                   
                                                If Trim(DisplaybCompanyID)=Trim(PreExistingDestination) then
                                                    bCompanyID=oRs("CompanyID")
												    bCompanyName=oRs("CompanyName")
												    bCompanyAddress=oRs("CompanyAddress")
                                                    bCompanyBuilding=oRs("CompanyBuilding")
                                                    DisplayDestinationBuilding=bCompanyBuilding
                                                    bCompanySuite=oRs("CompanySuite")
                                                        'If trim(bCompanySuite)>"" then
                                                        '    DisplayDestinationSuite="/"&bCompanySuite
                                                        '    else
                                                         '   DisplayDestinationSuite=""
                                                        'End if
												    bCompanyCity=oRs("CompanyCity")
                                                    bCompanyState=oRs("CompanyState")
                                                    bCompanyZip=oRs("CompanyZip")
                                                    bContactName=oRs("ContactName")
                                                    bCompanyPhone=oRs("CompanyPhone")
                                                    bCompanyEmail=oRs("CompanyEmail")
                                                End if
                                                If trim(PreExistingDestination)<>trim(TempPreExistingDestination) then
												    DestinationCompany=bCompanyName
												    DestinationAddress=bCompanyAddress
                                                    DestinationBuilding=bCompanyBuilding
                                                    DisplayDestinationBuilding=DestinationBuilding
                                                    DestinationSuite=bCompanySuite
                                                    'If trim(DestinationSuite)>"" then
                                                    '    DisplayDestinationSuite="/"&bCompanySuite
                                                     '   else
                                                    '    DisplayDestinationSuite=""
                                                    'End if
												    DestinationCity=bCompanyCity
                                                    DestinationState=bCompanyState
                                                    DestinationZipCode=bCompanyZip
                                                    tempDestinationContactName=bContactName
                                                    tempDestinationPhoneNumber=bCompanyPhone
                                                    tempDestinationEmail=bCompanyEmail
                                                    If trim(tempDestinationContactName)>"" then DestinationContactName=tempDestinationContactName end if
                                                    If trim(tempDestinationPhoneNumber)>"" then DestinationPhoneNumber=tempDestinationPhoneNumber end if
                                                    If trim(tempDestinationEmail)>"" then DestinationEmail=tempDestinationEmail end if
                                                End if	
                                                If Trim(DisplaybCompanyName)="" then DisplayDestinationSuite=Replace(DisplayDestinationSuite,"/","") end if							
											%>
											<option value="<%=DisplaybCompanyID%>" <%if trim(PreExistingDestination)=trim(DisplaybCompanyID) then response.Write " selected" end if%>><%=DisplayDestinationBuilding%><%=DisplayDestinationSuite %></option>
											<%
										oRs.movenext
										LOOP
									Set oConn=Nothing									
									%>
								</select> 
                    </td>
                </tr>
                <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left" nowrap><font color="red">&nbsp;&nbsp;</font>Company Name</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left"><input type="text" class="textbox"  name="DESTINATIONCompany" value="<%=DESTINATIONCompany%>" size="45" maxlength="40" /></td>
                    </tr>
                    <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left"><font color="red">&nbsp;&nbsp;</font>Address</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left"><input type="text" class="textbox"  name="DESTINATIONAddress" value="<%=DESTINATIONAddress%>" size="45" maxlength="40" /></td>
                    </tr>
                    <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left"><font color="red">*</font>Building</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left"><input type="text" class="textbox"  name="DESTINATIONBuilding" value="<%=DESTINATIONBuilding%>" size="45" maxlength="40" /></td>
                    </tr>
                    <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left"><font color="red">*</font>Floor/Cube Number</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left"><input type="text" class="textbox"  name="DESTINATIONSuite" value="<%=DESTINATIONSuite%>" size="45" maxlength="40" /></td>
                    </tr>
                    <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left"><font color="red">&nbsp;&nbsp;</font>City/State/Zip</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td class="FleetExpressTextBlackBold">
                            <input type="text" class="textbox"  name="DESTINATIONCity" value="<%=DESTINATIONCity%>" size="20" maxlength="30" />
                            <!--
                            <select name="DESTINATIONState">
	                        <option value="AL" <%If trim(DestinationState)="AL" then Response.write "selected" end if%>>AL</option>
	                        <option value="AK" <%If trim(DestinationState)="AK" then Response.write "selected" end if%>>AK</option>
	                        <option value="AZ <%If trim(DestinationState)="AZ" then Response.write "selected" end if%>">AZ</option>
	                        <option value="AR <%If trim(DestinationState)="AR" then Response.write "selected" end if%>">AR</option>
	                        <option value="CA" <%If trim(DestinationState)="CA" then Response.write "selected" end if%>>CA</option>
	                        <option value="CO" <%If trim(DestinationState)="CO" then Response.write "selected" end if%>>CO</option>
	                        <option value="CT" <%If trim(DestinationState)="CT" then Response.write "selected" end if%>>CT</option>
	                        <option value="DE" <%If trim(DestinationState)="DE" then Response.write "selected" end if%>>DE</option>
	                        <option value="DC" <%If trim(DestinationState)="DC" then Response.write "selected" end if%>>DC</option>
	                        <option value="FL" <%If trim(DestinationState)="FL" then Response.write "selected" end if%>>FL</option>
	                        <option value="GA" <%If trim(DestinationState)="GA" then Response.write "selected" end if%>>GA</option>
	                        <option value="HI" <%If trim(DestinationState)="HI" then Response.write "selected" end if%>>HI</option>
	                        <option value="ID" <%If trim(DestinationState)="ID" then Response.write "selected" end if%>>ID</option>
	                        <option value="IL" <%If trim(DestinationState)="IL" then Response.write "selected" end if%>>IL</option>
	                        <option value="IN" <%If trim(DestinationState)="IN" then Response.write "selected" end if%>>IN</option>
	                        <option value="IA" <%If trim(DestinationState)="IA" then Response.write "selected" end if%>>IA</option>
	                        <option value="KS" <%If trim(DestinationState)="KS" then Response.write "selected" end if%>>KS</option>
	                        <option value="KY" <%If trim(DestinationState)="KY" then Response.write "selected" end if%>>KY</option>
	                        <option value="LA" <%If trim(DestinationState)="LA" then Response.write "selected" end if%>>LA</option>
	                        <option value="ME" <%If trim(DestinationState)="ME" then Response.write "selected" end if%>>ME</option>
	                        <option value="MD" <%If trim(DestinationState)="MD" then Response.write "selected" end if%>>MD</option>
	                        <option value="MA" <%If trim(DestinationState)="MA" then Response.write "selected" end if%>>MA</option>
	                        <option value="MI" <%If trim(DestinationState)="MI" then Response.write "selected" end if%>>MI</option>
	                        <option value="MN" <%If trim(DestinationState)="MN" then Response.write "selected" end if%>>MN</option>
	                        <option value="MS" <%If trim(DestinationState)="MS" then Response.write "selected" end if%>>MS</option>
	                        <option value="MO" <%If trim(DestinationState)="MO" then Response.write "selected" end if%>>MO</option>
	                        <option value="MT" <%If trim(DestinationState)="MT" then Response.write "selected" end if%>>MT</option>
	                        <option value="NE" <%If trim(DestinationState)="NE" then Response.write "selected" end if%>>NE</option>
	                        <option value="NV" <%If trim(DestinationState)="NV" then Response.write "selected" end if%>>NV</option>
	                        <option value="NH" <%If trim(DestinationState)="NH" then Response.write "selected" end if%>>NH</option>
	                        <option value="NJ" <%If trim(DestinationState)="NJ" then Response.write "selected" end if%>>NJ</option>
	                        <option value="NM" <%If trim(DestinationState)="NM" then Response.write "selected" end if%>>NM</option>
	                        <option value="NY" <%If trim(DestinationState)="NY" then Response.write "selected" end if%>>NY</option>
	                        <option value="NC" <%If trim(DestinationState)="NC" then Response.write "selected" end if%>>NC</option>
	                        <option value="ND" <%If trim(DestinationState)="ND" then Response.write "selected" end if%>>ND</option>
	                        <option value="OH" <%If trim(DestinationState)="OH" then Response.write "selected" end if%>>OH</option>
	                        <option value="OK" <%If trim(DestinationState)="OK" then Response.write "selected" end if%>>OK</option>
	                        <option value="OR" <%If trim(DestinationState)="OR" then Response.write "selected" end if%>>OR</option>
	                        <option value="PA" <%If trim(DestinationState)="PA" then Response.write "selected" end if%>>PA</option>
	                        <option value="RI" <%If trim(DestinationState)="RI" then Response.write "selected" end if%>>RI</option>
	                        <option value="SC" <%If trim(DestinationState)="SC" then Response.write "selected" end if%>>SC</option>
	                        <option value="SD" <%If trim(DestinationState)="SD" then Response.write "selected" end if%>>SD</option>
	                        <option value="TN" <%If trim(DestinationState)="TN" then Response.write "selected" end if%>>TN</option>
	                        <option value="TX" <%If trim(DestinationState)="TX" or trim(DestinationState)="" then Response.write "selected" end if%>>TX</option>
	                        <option value="UT" <%If trim(DestinationState)="UT" then Response.write "selected" end if%>>UT</option>
	                        <option value="VT" <%If trim(DestinationState)="VT" then Response.write "selected" end if%>>VT</option>
	                        <option value="VA" <%If trim(DestinationState)="VA" then Response.write "selected" end if%>>VA</option>
	                        <option value="WA" <%If trim(DestinationState)="WA" then Response.write "selected" end if%>>WA</option>
	                        <option value="WV" <%If trim(DestinationState)="WV" then Response.write "selected" end if%>>WV</option>
	                        <option value="WI" <%If trim(DestinationState)="WI" then Response.write "selected" end if%>>WI</option>
	                        <option value="WY" <%If trim(DestinationState)="WY" then Response.write "selected" end if%>>WY</option>
                            </select>
                        -->&nbsp;TX&nbsp;&nbsp;
                        <input type="hidden" name="DestinationState" value="TX">
                            <input type="text" class="textbox"  name="DESTINATIONZipCode" value="<%=DESTINATIONZipCode%>" size="11" maxlength="10" />

                        </td>
                    </tr>

                    <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left"><font color="red">*</font>Contact Name</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left"><input type="text" class="textbox"  name="DESTINATIONContactName" value="<%=DESTINATIONContactName%>" size="45" maxlength="25" /></td>
                    </tr>
                    <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left"><font color="red">*</font>Phone Number</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left"><input type="text" class="textbox"  name="DESTINATIONPhoneNumber" value="<%=DESTINATIONPhoneNumber%>" size="45" maxlength="20" /></td>
                    </tr>
                    <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr> 
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left" nowrap><font color="red">&nbsp;&nbsp;</font>Email Address</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left"><input type="text" class="textbox"  name="DESTINATIONEmail" value="<%=DESTINATIONEmail%>" size="45" maxlength="100" /></td>
                    </tr>
                    <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr> 
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left"><font color="red">*</font>Delivery Date/Time</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left"><input type="text" class="textbox"  name="DeliveryDateTime" id="DeliveryDateTime" value="<%=DeliveryDateTime%>" size="30" maxlength="30" />
                     <a href="javascript:NewCal('DeliveryDateTime','MMddyyyy',true,12)"><img src="../images/cal.gif" width="16" height="16" border="0" alt="Pick a date"></a>
                    </td>
                </tr>

<%
									Set oConn = Server.CreateObject("ADODB.Connection")
									oConn.ConnectionTimeout = 100
									oConn.Provider = "MSDASQL"
									oConn.Open DATABASE
										l_cSQL = "Select * FROM RateList WHERE Ratestatus='c'"
										SET oRs = oConn.Execute(l_cSql)
												If not oRs.EOF then
                                                    RateCharge=oRs("RateCharge")
												    Surcharge=oRs("Surcharge")
                                                End if								
									    Set oConn=Nothing
                                        EstimatedCharge=Round(ccur(RateCharge)+(ccur(RateCharge)*cDBL(Surcharge)), 2)
                                        'Response.write "cDBL(Surcharge)="&cDBL(Surcharge)&"<BR>"	
 %>
        <!--
        <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
        <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
        <tr><td colspan="4" bgcolor="black"><img src="../images/pixel.gif" height="1" width="1" /></td></tr>
        <tr><td><img src="../images/pixel.gif" height="2" width="1" /></td></tr>
                 <tr>
                 <td valign="top" rowspan="2"  class="OrderHeader">CHARGES<img src="../images/pixel.gif" height="1" width="15" /></td>
                        <td class="FleetExpressTextBlackBold" align="left">Estimated Costs</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td nowrap class="FleetExpressTextBlackBold">
                            $<%=EstimatedCharge %>
                        </td>
                    </tr>
                    <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
        -->
                    <input type="hidden" value="<%=EstimatedCharge %>" name="BasicCharge" />



         
        <input type="hidden" value="1" name="Timesthrough" />
         <%If Errormessage>"" then%>
         <tr><td>&nbsp;</td></tr>
         <tr>
            <td align="center" colspan="4">
                        
                <table cellpadding="2" cellspacing="2" border="1" bordercolor="red">
                    <tr>
                        <td class="errormessage">
                            <%
                            Response.write " * * * ERROR:  "&ErrorMessage& " * * * "
                             %>
                        </td>
                    </tr>
                </table>
                
           </td>
        </tr> 
        <tr><td>&nbsp;</td></tr>
        <%End if 
        TempPreExistingOrigination=trim(PreExistingOrigination)
        TempPreExistingDestination=trim(PreExistingDestination)
        %>
        <tr><td>&nbsp;</td></tr>
        <tr><td>&nbsp;</td></tr>
         <tr><td align="center" colspan="5">
         <!--input type="image" src="../images/submitorder.gif" alt="submit order" /><br /-->
         <input type="submit" name="submitbutton" id="submitbutton" value="Submit Order" /><br clear="all" />
         <img src="../images/pixel.gif" width="1" height="1000" />
         </td></tr> 
         <input type="hidden" name="TempPreExistingOrigination" value="<%=TempPreExistingOrigination %>" />
          <input type="hidden" name="TempPreExistingDestination" value="<%=TempPreExistingDestination %>" />
        <input type="hidden" name="ColorSelect" value="<%=ColorSelect %>" />
        <input type="hidden" name="MarkTemp" value="<%=MarkTemp %>" />
        <input type="hidden" name="CaptchaSubmit" value="<%=CaptchaSubmit %>" />
        <input type="hidden" name="varCaptcha" value="<%=varCaptcha %>" />
        <input type="hidden" name="FXCourierUserID" value="<%=FXCourierUserID %>" />
        <input type="hidden" name="LogInVerified" value="<%=LogInVerified %>" />
        <input type="hidden" name="ordersubmitted" value="yes" />
        <input type="hidden" name="pagestatus" value="submit" />
        </form>
        
    </table>
</td>
<td><img src="../images/pixel.gif" width="30" height="1" /></td>
</tr>
</table>
<%

else
%>
<table align="center">
    <tr><td>&nbsp;</td></tr>
    <tr><td>&nbsp;</td></tr>
    <tr><td>&nbsp;</td></tr>
    <tr><td>&nbsp;</td></tr>
    <tr><td>&nbsp;</td></tr>
    <tr>
        <td>
        <font color="white" size="5">
            Fleet Express Courier Service only accepts online orders between 7:30 AM and 7:00 PM Monday-Friday.<br /><br />
            During non-business hours, please call in your order to the drivers directly at 214-882-0620.
        </font>
        </td>
    </tr>
</table>
<%
'REsponse.write "WHOOPITY DOO!!!<BR>"
'Response.write "PageStatus="&PageStatus&"<BR>"
End if  'For the WHOLE only taking orders during certain times THANG!

''''else
''''Response.redirect("../home.asp")
'end if
''''end if
%>
</body>
</html>
