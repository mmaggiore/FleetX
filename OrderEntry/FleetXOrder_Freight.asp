<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include file="../fleetexpress.inc" -->
<script type="text/javascript">

    var _gaq = _gaq || [];
    _gaq.push(['_setAccount', 'UA-37615940-5']);
    _gaq.push(['_trackPageview']);

    (function () {
        var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
        ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
        var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
    })();
 
</script>

    <title>Fleet Express Transportation Request</title>
    <link rel="stylesheet" type="text/css" href="../css/Style.css">
   <%
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
   '''''''''''HARDCODED STUFF
   sBT_ID=86
    UserID=Request.Form("UserID")
    LoggedIn=Request.QueryString("LoggedIn")
    If trim(LoggedIn)="y" then
        UserID=Session("PhoneBookID")
    End if
    If trim(UserID)>"" then
        Session("PhoneBookID")=UserID
    End if
    'If trim(userID)="" then
    '    UserID=Session("UserID")
    'End if
    'If trim(UserID)>"" then
    '    Session("UserID")=UserID
    'End if
'Response.write "***UserID="&UserID&"<BR>"
   ''''''''''''''''''''''''''


   CurrentHour=Hour(Now())
   CurrentMinute=Minute(Now())
   'Response.write "CurrentHour="&CurrentHour&"<BR>"
   'Response.write "CurrentMinute="&CurrentMinute&"<BR>"
   '''userid=session("UserID")
   OrigEmail=Request.Form("OrigEmail")
   DestEmail=Request.Form("DestEmail")
   OtherEmail=Request.Form("OtherEmail")
   AdditionalEmail=Request.Form("AdditionalEmail")
   XSquare=Request.form("XSquare")
   'Response.write "XSquare="&XSquare&"<BR>"
    MarkTemp=Request.Form("MarkTemp")
    loginattempt=Request.Form("LoginAttempt")
    UserName=Request.Form("UserName")
    Password=Request.Form("Password")
    ''''''''REMOVED THIS FOR CAPTCHA....MIGHT NEED TO PUT IT BACK!
    'If trim(UserID)="" and trim(MarkTemp)=""  then
    '    Response.redirect("http://www.logisticorp.us/intranet")
    '    else
    '    MarkTemp="yes"
    'End if
    'Response.write "loginattempt="&loginattempt&"<BR>"
    If loginattempt="y" then
        CaptchaSubmit=Request.form("CaptchaSubmit")
        varCaptcha=Request.form("varCaptcha")
        If CaptchaSubmit<>varCaptcha then
            ErrorMessage="You did not supply the correct verification code"
        End if
        If trim(Username)="" then
            ErrorMessage="You did not supply a username"
        End if
        If trim(Password)="" then
            ErrorMessage="You did not supply a password"
        End if
        If trim(ErrorMessage)="" then
            'Response.write "Database="&Database&"<BR>"
		    Set oConn = Server.CreateObject("ADODB.Connection")
		    oConn.ConnectionTimeout = 100
		    oConn.Provider = "MSDASQL"
		    oConn.Open DATABASE
			    l_cSQL = "SELECT * FROM PreExistingRequestor WHERE (RequestorEmail='"& Username &"') and (RequestorPassword='"& Password &"') and  (RequestorStatus <> 'x')"
			    SET oRs = oConn.Execute(l_cSql)
					    if oRs.EOF then
                            ErrorMessage="That username/password combination is not valid."
                            else
                            UserID=oRs("RequestorID")
                            Session("PhoneBookID")=UserID
                        End if								
		    Set oConn=Nothing
        End if
    End if
    'Response.write "CaptchaSubmit="&CaptchaSubmit&"<BR>"
    'Response.write "varCaptcha="&varCaptcha&"<BR>"
    'Response.write "UserID="&UserID&"<BR>"
    'If (CaptchaSubmit=varCaptcha and trim(CaptchaSubmit)>"" and trim(varCaptcha)>"") or trim(UserID)>"" then
    If trim(UserID)>"" then
    %>

    <!-- #include file="../include/checkstring.inc" -->
<script language="javascript" type="text/javascript" src="datetimepicker.js">

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


    timesthrough=Request.form("timesthrough")
    
    TableWidth="460"
    Internal=Request.QueryString("Internal")
    
    PreExistingRequestor=Request.Form("PreExistingRequestor")
    PreExistingOrigination=Request.Form("PreExistingOrigination")
    PreExistingDestination=Request.Form("PreExistingDestination")
    'Response.write "xxxPreExistingOrigination="&PreExistingOrigination&"<BR>"
    'Response.write "PreExistingRequestor="&PreExistingRequestor&"<BR>"
    'Response.write "PreExistingDestination="&PreExistingDestination&"<BR>"
    'If trim(PreExistingOrigination)>"" OR trim(PreExistingDestination)>"" OR trim(PreExistingRequestor)>"" then
    If (trim(PreExistingOrigination)="" OR trim(PreExistingDestination)="") then
        PageStatus="x"
        else
        PageStatus=Request.form("PageStatus")
    End if
    'Response.write "***PageStatus="&PageStatus&"<BR>"
    RequestorName=Request.form("RequestorName")
    RequestorPhoneNumber=Request.form("RequestorPhoneNumber")
    RequestorEmailAddress=Request.form("RequestorEmailAddress")
    'PONumber=Request.form("PONumber")
    'CostCenterNumber=Request.form("CostCenterNumber")
    Pieces=Request.form("Pieces")
    NumberOfPallets=Request.form("NumberOfPallets")
    DimWeight=Request.form("DimWeight")
    DimLength=Request.form("DimLength")
    DimWidth=Request.form("DimWidth")
    DimHeight=Request.form("DimHeight")

    IsPalletized=Request.form("IsPalletized")
    DimValue=Request.form("DimValue")
    IsHazmat=Request.form("IsHazmat")
    OriginationID=Request.Form("OriginationID")
    'Response.write "OriginationID="&OriginationID&"<BR>"
    OriginationCompany=Request.form("OriginationCompany")
    OriginationAddress=Request.form("OriginationAddress")
    OriginationCity=Request.form("OriginationCity")
    OriginationState=Request.form("OriginationState")
    'Response.write "***OriginationState="&OriginationState&"<BR>"
    OriginationZipCode=Request.form("OriginationZipCode")
    OriginationContactName=Request.form("OriginationContactName")
    OriginationPhoneNumber=Request.form("OriginationPhoneNumber")
    OriginationEmail=Request.form("OriginationEmail")
    DestinationID=Request.Form("DestinationID")
    'Response.write "DestinationID="&DestinationID&"<BR>"
    DestinationCompany=Request.form("DestinationCompany")
    DestinationAddress=Request.form("DestinationAddress")
    DestinationCity=Request.form("DestinationCity")
    DestinationState=Request.form("DestinationState")
    DestinationZipCode=Request.form("DestinationZipCode")
    DestinationContactName=Request.form("DestinationContactName")
    DestinationPhoneNumber=Request.form("DestinationPhoneNumber")
    DestinationEmail=Request.form("DestinationEmail")

    POorNWA=Request.form("POorNWA")
    GenericNumber=Request.form("GenericNumber")
   OriginationNotifications=Request.Form("OriginationNotifications")
   DestinationNotifications=Request.Form("DestinationNotifications")
   If trim(OriginationNotifications)="y" then
        SendTo=OriginationEmail
   End if
   If trim(DestinationNotifications)="y" then
        SendTo=DestinationEmail
   End if
   If trim(OriginationNotifications)="y" and trim(DestinationNotifications)="y" then
        SendTo=OriginationEmail&";"&DestinationEmail
   End if


    Select Case POorNWA
        Case "TI P/O #"
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
    'If trim(PickUpDateTime)="" then
    '    PickUpDateTime=now()
   ' End if
    DeliveryDateTime=Request.form("DeliveryDateTime")
    If trim(XSquare)="y" then
        If trim(DimWeight)="" then DimWeight="50" end if
        If trim(DimLength)="" then DimLength="10" end if
        If trim(DimWidth)="" then DimWidth="10" end if
        If trim(DimHeight)="" then DimHeight="10" end if
        If trim(Comments)="" then Comments="NON-STOP" end if
        'If trim(GenericNumber)="" then GenericNumber="XXX" end if
        PickUpDateTime=Now()
        DeliveryDateTime=Now()+.1666664
    End if    


    ''''''''PRE-POPULATES ITEMS FOR LOGISTICORP PEEPS
    If Internal="y" and TRIM(timesthrough)="" then
        'Response.write "Got here!<br>"
        RequestorEmailAddress="FleetExpress@logisticorp.us"
        'RequestorEmailAddress="mark.maggiore@logisticorp.us"
        OriginationContactName="Dispatch"
        OriginationPhoneNumber="972-499-3415"
        OriginationEmail="FleetExpress@logisticorp.us"
        'OriginationEmail="mark.maggiore@logisticorp.us"
        DestinationContactName="Dispatch"
        DestinationPhoneNumber="972-499-3415"
        DestinationEmail="FleetExpress@logisticorp.us"
        'DestinationEmail="mark.maggiore@logisticorp.us"
    End if
''''''''''''''SAVES ORDER INFO WHEN GOING TO OTHER PAGES''''''''''''''''''
Submit=Request.Form("Submit")
ButtonSubmit=Request.Form("ButtonSubmit")
varA=Request.QueryString("varA")
If trim(varA)="123" then
    PreExistingRequestor=Session("PreExistingRequestor")
    RequestorName=Session("RequestorName")
    RequestorPhoneNumber=Session("RequestorPhoneNumber")
    RequestorEmailAddress=Session("RequestorEmailAddress")
    POorNWA=Session("POorNWA")
    GenericNumber=Session("GenericNumber")
    comments=Session("comments")
    pieces=Session("pieces")
    rf_box=Session("rf_box")
    IsPalletized=Session("IsPalletized")
    DimWeight=Session("DimWeight")
    DimLength=Session("DimLength")
    DimWidth=Session("DimWidth")
    DimHeight=Session("DimHeight")
    IsHazmat=Session("IsHazmat")
    Refrigerate=Session("Refrigerate")
    XSquare=Session("XSquare")
    Priority=Session("Priority")
    PreExistingOrigination=Session("PreExistingOrigination")
    OriginationID=Session("OriginationID")
    OriginationNotifications=Session("OriginationNotifications")
    PreExistingDestination=Session("PreExistingDestination")
    DestinationID=Session("DestinationID")
    DestinationNotifications=Session("DestinationNotifications")
    PickupDateTime=Session("PickupDateTime")
    DeliveryDateTime=Session("DeliveryDateTime")
    ColorSelect=Session("ColorSelect")
    MarkTemp=Session("MarkTemp")
    CaptchaSubmit=Session("CaptchaSubmit")
    varCaptcha=Session("varCaptcha")
    UserID=Session("UserID")
    pagestatus=Session("pagestatus")
End if
If lcase(trim(ButtonSubmit))="edit requestor information" OR lcase(trim(ButtonSubmit))="add/edit locations in your address book"  then
    Session("PreExistingRequestor")=PreExistingRequestor
    Session("RequestorName")=RequestorName
    Session("RequestorPhoneNumber")=RequestorPhoneNumber
    Session("RequestorEmailAddress")=RequestorEmailAddress
    Session("POorNWA")=POorNWA
    Session("GenericNumber")=GenericNumber
    Session("comments")=comments
    Session("pieces")=pieces
    Session("rf_box")=rf_box
    Session("IsPalletized")=IsPalletized
    Session("DimWeight")=DimWeight
    Session("DimLength")=DimLength
    Session("DimWidth")=DimWidth
    Session("DimHeight")=DimHeight
    Session("IsHazmat")=IsHazmat
    Session("Refrigerate")=Refrigerate
    Session("XSquare")=XSquare
    Session("Priority")=Priority
    Session("PreExistingOrigination")=PreExistingOrigination
    Session("OriginationID")=OriginationID
    Session("OriginationNotifications")=OriginationNotifications
    Session("PreExistingDestination")=PreExistingDestination
    Session("DestinationID")=DestinationID
    Session("DestinationNotifications")=DestinationNotifications
    Session("PickupDateTime")=PickupDateTime
    Session("DeliveryDateTime")=DeliveryDateTime
    Session("ColorSelect")=ColorSelect
    Session("MarkTemp")=MarkTemp
    Session("CaptchaSubmit")=CaptchaSubmit
    Session("varCaptcha")=varCaptcha
    Session("UserID")=UserID
    Session("pagestatus")=pagestatus

If lcase(trim(ButtonSubmit))="edit requestor information" then
        Response.redirect("FleetExpressUserEdit.asp")
        else
        Response.redirect("FleetExpressAddressBook.asp")
    End IF
End if
''''''''''''''END







    ''''''''ERROR HANDLING''''''''''
    'Response.write "GOT HERE!!!<BR>"
   ' Response.write "PageStatus="&PageStatus&"***<BR>"
    If PageStatus="submit" then
    'If Trim(OrigEmail)="y" then
   '     AllNotifications=OriginationEmail
   ' End if
    'If Trim(DestEmail)="y" then
    '    AllNotifications=AllNotifications&";"&DestinationEmail
    'End if
   ' If Trim(OtherEmail)="y" then
    '    AllNotifications=AllNotifications&";"&AdditionalEmail
    'End if
    'somevar=len(AllNotifications)
    'If left(AllNotifications,1)=";" then 
    '    AllNotifications=Right(AllNotifications,(somevar-1))
    'End if
    'Response.write "AllNotifications="&AllNotifications&"<BR>"
     'Response.write "GOT HERE two!!!<BR>"

        'If trim(DestinationEmail)="" then
        '    ErrorMessage="You must provide the Destination's Email"
       ' End if
       If not isdate(DeliveryDateTime) then DeliveryDateTime=now() end if
        If not isdate(PickUPDateTime) then PickUPDateTime=now() end if
        If cdate(PickUpDateTime)<now() then PickUpDateTime=Now() end if
       
       If Priority="Four Hours" then
           'Response.write "PickUpDateTime="&PickUpDateTime&"<BR>"
           'Response.write "CurrentDateTime="&CurrentDateTime&"<BR>"
           TimePeriod=DateDiff("n", PickUPDateTime, DeliveryDateTime)
           'Response.write "TimePeriod="&TimePeriod&"<BR>"
           If TimePeriod<=239 then
                'Response.write "Got heya!<br>"
                ErrorMessage="The delivery date/time cannot be less than four hours after the ready date/time"
           End if
       End if
       If Priority="2 Hours" then
           'Response.write "PickUpDateTime="&PickUpDateTime&"<BR>"
           'Response.write "CurrentDateTime="&CurrentDateTime&"<BR>"
           TimePeriod=DateDiff("n", PickUPDateTime, DeliveryDateTime)
           'Response.write "TimePeriod="&TimePeriod&"<BR>"
           If TimePeriod<=119 then
                'Response.write "Got heya!<br>"
                ErrorMessage="The delivery date/time cannot be less than two hours after the ready date/time"
           End if
       End if 
       If Priority="4 Hours" then
           'Response.write "PickUpDateTime="&PickUpDateTime&"<BR>"
           'Response.write "CurrentDateTime="&CurrentDateTime&"<BR>"
           TimePeriod=DateDiff("n", PickUPDateTime, DeliveryDateTime)
           'Response.write "TimePeriod="&TimePeriod&"<BR>"
           If TimePeriod<=239 then
                'Response.write "Got heya!<br>"
                ErrorMessage="The delivery date/time cannot be less than four hours after the ready date/time"
           End if
       End if
       
   'Response.write "CurrentHour="&CurrentHour&"<BR>"
   'Response.write "DayOfWeek="&Weekday(now())&"<BR>"
   'Response.write "DayOfWeek="&WeekdayName(Weekday(now()))&"<BR>"
        If trim(XSquare)="y" then
            If (CurrentHour=6 and currentMinute<50) or CurrentHour<6 or CurrentHour>11 or WeekdayName(Weekday(now()))="Saturday" or WeekdayName(Weekday(now()))="Sunday" then
                   ErrorMessage="Due to dock hours, X Square Probe Card orders can only be placed Monday-Friday between 7:00 AM and 12:00 PM"
            End if
        End if 

        If trim(OtherEmail)="y" AND (inStr(AdditionalEmail,"@") = 0 OR inStr(AdditionalEmail,".") = 0) THEN
           ErrorMessage="The 'Additional' email that provided '"& AdditionalEmail &"' is not a valid email address."
        End if
        If trim(OtherEmail)="y" and (trim(AdditionalEmail)="" or trim(AdditionalEmail)="Additional") then
            ErrorMessage="You checked the 'Additional' email notification box, you must enter an additional email address."
        end if
        If isdate(DeliveryDateTime)  and cdate(CurrentDateTime)>=cdate(DeliveryDateTime) then
            ErrorMessage="The delivery date/time cannot be before the current date/time"
        End if
        If isdate(DeliveryDateTime) and isdate(PickUPDateTime) and cdate(PickUPDateTime)>=cdate(DeliveryDateTime) then
            ErrorMessage="The delivery date/time cannot be before the ready date/time"
        End if
        'Response.write "DeliveryDateTime="&DeliveryDateTime&"<BR>"
        If NOT isdate(DeliveryDateTime) then
            ErrorMessage="You must provide a valid destination date/time"
        End if
        If trim(DestinationCompany)=trim(OriginationCompany) AND trim(DestinationAddress)=Trim(OriginationAddress) then
        ErrorMessage="Your Origination and Destination cannot be the same"
        End if
        If trim(DestinationPhoneNumber)="" then
        ErrorMessage="You must provide the Destination's Phone Number"
        End if
        If trim(DestinationContactName)="" then
            ErrorMessage="You must provide the Destination's Contact Name"
        End if
        If trim(DestinationZipCode)="" then
            ErrorMessage="You must provide the Destination's Zip Code"
        End if
        If trim(DestinationCity)="" then
            ErrorMessage="You must provide the Destination's City"
        End if
        If trim(DestinationAddress)="" then
            ErrorMessage="You must provide the Destination's Address"
        End if
        If trim(DestinationCompany)="" then
            ErrorMessage="You must provide the Destination's Company"
        End if
        'If trim(OriginationEmail)="" then
        '    ErrorMessage="You must provide the Origination's Email"
        'End if
        If isdate(PickUpDateTime)  and cdate(CurrentDateTime)>cdate(PickUpDateTime) then
            ErrorMessage="The ready time cannot be before the current date/time"
        End if
        If NOT isdate(PickUpDateTime) then
            ErrorMessage="You must provide a valid ready date/time"
        End if
        If trim(OriginationPhoneNumber)="" then
            ErrorMessage="You must provide the Origination's Phone Number"
        End if
        If trim(OriginationContactName)="" then
            ErrorMessage="You must provide the Origination's Contact Name"
        End if
        If trim(OriginationZipCode)="" then
            ErrorMessage="You must provide the Origination's Zip Code"
        End if
        If trim(OriginationCity)="" then
            ErrorMessage="You must provide the Origination's City"
        End if
        If trim(OriginationAddress)="" then
            ErrorMessage="You must provide the Origination's Address"
        End if
        If trim(OriginationCompany)="" then
            ErrorMessage="You must provide the Origination's Company"
        End if
        If trim(DimHeight)="" then
            ErrorMessage="You must provide the Commodity's Height"
        End if
        If trim(DimWidth)="" then
            ErrorMessage="You must provide the Commodity's Width"
        End if
        If trim(DimLength)="" then
            ErrorMessage="You must provide the Commodity's Length"
        End if
        If trim(DimWeight)="" then
            ErrorMessage="You must provide the Commodity's Weight"
        End if
        'If trim(NumberOfPallets)="" and isPalletized="y" then
        '    ErrorMessage="You must provide the Number of Pallets"
        'End if
        If trim(Pieces)="" then
            ErrorMessage="You must provide the Number of Pieces"
        End if
        If trim(Comments)="" then
            ErrorMessage="You have not provided any Special Instructions for this delivery.<br>If there are no special instructions, please type in N/A"
        End if
        If trim(CostCenterNumber)="" AND trim(PONumber)="" then
            ErrorMessage="You must provide the Cost Center Number or P/O Number"
        End if
        'If trim(PONumber)="" then
        '    ErrorMessage="You must provide the P/O Number"
        'End if
        If trim(RequestorEmailAddress)="" then
            ErrorMessage="You must provide the Requestor Email Address"
        End if
        If trim(RequestorPhoneNumber)="" then
            ErrorMessage="You must provide the Requestor Phone Number"
        End if
        If trim(RequestorName)="" then
            ErrorMessage="You must provide the Requestor Name"
        End if
        'Response.write "ErrorMessage="&ErrorMessage&"<BR>"
        If trim(ErrorMessage)="" then

		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.ConnectionTimeout = 100
		oConn.Provider = "MSDASQL"
		oConn.Open DATABASE
			l_cSQL = "Select * FROM PreExistingCompanies WHERE companystatus='c' and CompanyName='"& OriginationCompany &"' and CompanyAddress='"& OriginationAddress &"'"
			SET oRs = oConn.Execute(l_cSql)
					if oRs.EOF then
                        'Response.write "ADD a NEW ORIGINATION!!!!<BR>"
			            Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				            RSEVENTS2.Open "PreExistingCompanies", DATABASE, 2, 2
				            RSEVENTS2.addnew
				            RSEVENTS2("CompanyName")=OriginationCompany
                            RSEVENTS2("CompanyAddress")=OriginationAddress
                            RSEVENTS2("CompanyCity")=OriginationCity
                            RSEVENTS2("CompanyState")=OriginationState
                            RSEVENTS2("CompanyZip")=OriginationZipCode
                            RSEVENTS2("CompanyStatus")="c"
				            RSEVENTS2.update
				            RSEVENTS2.close			
			            set RSEVENTS2 = nothing 
                    End if								
		Set oConn=Nothing

		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.ConnectionTimeout = 100
		oConn.Provider = "MSDASQL"
		oConn.Open DATABASE
			l_cSQL = "Select * FROM PreExistingRequestor WHERE requestorstatus='c' and RequestorName='"& RequestorName &"'"
			SET oRs = oConn.Execute(l_cSql)
					if oRs.EOF then
                        'Response.write "ADD a NEW ORIGINATION!!!!<BR>"
			            Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				            RSEVENTS2.Open "PreExistingRequestor", DATABASE, 2, 2
				            RSEVENTS2.addnew
				            RSEVENTS2("RequestorName")=RequestorName
                            RSEVENTS2("RequestorPhone")=RequestorPhoneNumber
                            RSEVENTS2("RequestorEmail")=RequestorEmailAddress
                            RSEVENTS2("RequestorStatus")="c"
				            RSEVENTS2.update
				            RSEVENTS2.close			
			            set RSEVENTS2 = nothing 
                    End if								
		Set oConn=Nothing


		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.ConnectionTimeout = 100
		oConn.Provider = "MSDASQL"
		oConn.Open DATABASE
			l_cSQL = "Select * FROM PreExistingCompanies WHERE companystatus='c' and CompanyName='"& DestinationCompany &"' and CompanyAddress='"& DestinationAddress &"'"
			SET oRs = oConn.Execute(l_cSql)
					if oRs.EOF then
                    'Response.write "ADD a NEW DESTINATION!!!!<BR>"
			            Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				            RSEVENTS2.Open "PreExistingCompanies", DATABASE, 2, 2
				            RSEVENTS2.addnew
				            RSEVENTS2("CompanyName")=DestinationCompany
                            RSEVENTS2("CompanyAddress")=DestinationAddress
                            RSEVENTS2("CompanyCity")=DestinationCity
                            RSEVENTS2("CompanyState")=DestinationState
                            RSEVENTS2("CompanyZip")=DestinationZipCode
                            RSEVENTS2("CompanyStatus")="c"
				            RSEVENTS2.update
				            RSEVENTS2.close			
			            set RSEVENTS2 = nothing 
                    End if								
		Set oConn=Nothing



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
            'Response.write "sBT_ID="&sBT_ID&"<BR>"
           
			Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				RSEVENTS2.Open "fcfgthd", DATABASE, 2, 2
				RSEVENTS2.addnew
				RSEVENTS2("fh_ID")=NewJobNum
                RSEVENTS2("fh_Status")="SCD"
                RSEVENTS2("fh_ship_dt")=now()
                RSEVENTS2("fh_ready")=PickUpDateTime
                RSEVENTS2("Fh_Priority")=Priority
                RSEVENTS2("fh_lastchg")=now()
                RSEVENTS2("fh_bt_ID")=sBT_ID
                RSEVENTS2("fh_co_id")=Trim(RequestorName)
                RSEVENTS2("fh_co_phone")=Trim(RequestorPhoneNumber)
                RSEVENTS2("fh_co_email")=Trim(SendTo)
                RSEVENTS2("fh_co_costcenter")=Trim(costcenterNumber)
                RSEVENTS2("fh_custpo")=Trim(PoNumber)
                RSEVENTS2("fh_statcode")="2"
				RSEVENTS2.update
				RSEVENTS2.close			
			set RSEVENTS2 = nothing  	       
           
            Set oConn=Nothing 


            If trim(OriginationID)="" then
		        Set oConn = Server.CreateObject("ADODB.Connection")
		        oConn.ConnectionTimeout = 100
		        oConn.Provider = "MSDASQL"
		        oConn.Open DATABASE
			        l_cSQL = "Select * FROM PreExistingCompanies WHERE companystatus='c' and CompanyName='"& OriginationCompany &"' and CompanyAddress='"& OriginationAddress &"'"
			        SET oRs = oConn.Execute(l_cSql)
					        if oRs.EOF then
                                else
                                OriginationID=oRs("CompanyID")
 
                            End if								
		        Set oConn=Nothing
            End if
            If trim(DestinationID)="" then
		        Set oConn = Server.CreateObject("ADODB.Connection")
		        oConn.ConnectionTimeout = 100
		        oConn.Provider = "MSDASQL"
		        oConn.Open DATABASE
			        l_cSQL = "Select * FROM PreExistingCompanies WHERE companystatus='c' and CompanyName='"& DestinationCompany &"' and CompanyAddress='"& DestinationAddress &"'"
			        SET oRs = oConn.Execute(l_cSql)
					        if oRs.EOF then
                                else
                                DestinationID=oRs("CompanyID")
 
                            End if								
		        Set oConn=Nothing
            End if




            'Response.write "xxxDatabase="&database&"<BR>"
			Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				RSEVENTS2.Open "fclegs", DATABASE, 2, 2
				RSEVENTS2.addnew
				RSEVENTS2("fl_fh_ID")=Trim(NewJobNum)
                RSEVENTS2("fl_sf_ID")=Trim(OriginationID)
                RSEVENTS2("fl_sf_name")=Trim(OriginationCompany)
                RSEVENTS2("fl_sf_clname")=Trim(OriginationContactName)
                RSEVENTS2("fl_sf_phone")=Trim(OriginationPhoneNumber)
                RSEVENTS2("fl_sf_email")=Trim(OriginationEmail)
                RSEVENTS2("fl_sf_addr1")=Trim(OriginationAddress)
                RSEVENTS2("fl_sf_city")=Trim(OriginationCity)
                RSEVENTS2("fl_sf_state")=Trim(OriginationState)
                RSEVENTS2("fl_sf_country")="US"
                RSEVENTS2("fl_sf_zip")=Trim(OriginationZipCode)
                
                RSEVENTS2("fl_st_ID")=Trim(DestinationID)
                RSEVENTS2("fl_st_name")=Trim(DestinationCompany)
                RSEVENTS2("fl_st_clname")=Trim(DestinationContactName)
                RSEVENTS2("fl_st_phone")=Trim(DestinationPhoneNumber)
                RSEVENTS2("fl_st_email")=Trim(DestinationEmail)
                RSEVENTS2("fl_st_addr1")=Trim(DestinationAddress)
                RSEVENTS2("fl_st_city")=Trim(DestinationCity)
                RSEVENTS2("fl_st_state")=Trim(DestinationState)
                RSEVENTS2("fl_st_country")="US"
                RSEVENTS2("fl_st_zip")=Trim(DestinationZipCode)
                RSEVENTS2("fl_sf_comment")=Trim(Comments)
                RSEVENTS2("fl_st_rta")=Trim(DeliveryDateTime)
               
                RSEVENTS2("fl_leg_status")="c"
				RSEVENTS2.update
				RSEVENTS2.close			
			set RSEVENTS2 = nothing 

			Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				RSEVENTS2.Open "fcrefs", DATABASE, 2, 2
				RSEVENTS2.addnew
				RSEVENTS2("rf_fh_ID")=Trim(NewJobNum)
                RSEVENTS2("rf_ref")=Trim(NewJobNum)
                RSEVENTS2("rf_box")=trim(rf_box)
                RSEVENTS2("NumberOfPieces")=Trim(Pieces)
                RSEVENTS2("IsPalletized")=Trim(IsPalletized)
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
                    Body = Body & "***Should you need to cancel this order, please call 972-499-3415***<BR><BR>"

                    Body = Body & "REQUESTOR INFORMATION:<BR>"
                    Body = Body & "Name: "&  RequestorName &"<br>"  
                    Body = Body & "Phone Number: "&  RequestorPhoneNumber &"<br>"  
                    Body = Body & "Email Address: "&  RequestorEmailAddress &"<br>"   
                    Body = Body & "PO Number: "&  PONumber &"<br>"  
                    Body = Body & "Cost Center Number: "&  CostCenterNumber &"<br><br>" 
                    Body = Body & "COMMODITY INFORMATION:<BR>" 
                    Body = Body & "Pieces: "&  Pieces &" "& rf_box &"<br>" 
                    Body = Body & "Palletized: "&  IsPalletized &"<br>"   
                    'Body = Body & "Number Of Pallets: "&  NumberOfPallets &"<br>"  
                    Body = Body & "Weight: "&  DimWeight &"LBS<br>"
                    Body = Body & " Dimensions: "&  DimLength &" X "&  DimWidth &" X "&  DimHeight &" inches<br>"       
  
                    
                    'Body = Body & "Hazmat: "&  IsHazmat &"<br>"
                    'Body = Body & "Refrigerate: "&  Refrigerate &"<br><br>"
                    Body = Body & "ORIGINATION:<BR>"   
                    Body = Body & "Company: "&  OriginationCompany &"<br>"   
                    Body = Body & "Address: "&  OriginationAddress &"<br>"   
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
				    Body = Body & "FleetX Services<br>"  
				    Body = Body &  "FleetX@LogisticorpGroup.com<br>"  
				    Body = Body & "972/499-3415<br><br>"
				    'Recipient=FirstName&" "&LastName
			        SentToEmail=RequestorEmailAddress
				    'Email="KWETI.Mailbox@am.kwe.com"
				    'Email="mark@maggiore.net"
				    Set objMail = CreateObject("CDONTS.Newmail")
				    objMail.From = "FleetX@LogisticorpGroup.com"
				    objMail.To = SentToEmail
				    objMail.cc = "mark.maggiore@logistiCorp.us"
				    'objMail.cc = "4692269939@tmomail.net;pcurrin@ti.com"
				    objMail.Subject = "Thank you for your FleetX shipment request"
				    objMail.MailFormat = cdoMailFormatMIME
				    objMail.BodyFormat = cdoBodyFormatHTML
				    objMail.Body = Body
                    If trim(lcase(SentToEmail))<>"fleetexpress@logisticorp.us" then
				        objMail.Send
                    End if
				    Set objMail = Nothing         
            
            		
				    Body = "There has been a new Fleet Express shipment request (#"& newjobnum &") placed online:<br><br>"   

                    Body = Body & "REQUESTOR INFORMATION:<BR>"
                    Body = Body & "Name: "&  RequestorName &"<br>"  
                    Body = Body & "Phone Number: "&  RequestorPhoneNumber &"<br>"  
                    Body = Body & "Email Address: "&  RequestorEmailAddress &"<br>"   
                    Body = Body & "PO Number: "&  PONumber &"<br>"  
                    Body = Body & "Cost Center Number: "&  CostCenterNumber &"<br><br>" 
                    Body = Body & "COMMODITY INFORMATION:<BR>" 
                    Body = Body & "Pieces: "&  Pieces &" "& rf_box &"<br>" 
                    Body = Body & "Palletized: "&  IsPalletized &"<br>"   
                    'Body = Body & "Number Of Pallets: "&  NumberOfPallets &"<br>"  
                    Body = Body & "Weight: "&  DimWeight &"LBS<br>"
                    Body = Body & " Dimensions: "&  DimLength &" X "&  DimWidth &" X "&  DimHeight &" inches<br>"       
  
                    
                    'Body = Body & "Hazmat: "&  IsHazmat &"<br>"
                    'Body = Body & "Refrigerate: "&  Refrigerate &"<br><br>"
                    Body = Body & "ORIGINATION:<BR>"   
                    Body = Body & "Company: "&  OriginationCompany &"<br>"   
                    Body = Body & "Address: "&  OriginationAddress &"<br>"   
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
				    Body = Body & "FleetX Services<br>"  
				    Body = Body &  "FleetX@LogisticorpGroup.com<br>"  
				    Body = Body & "972/499-3415<br><br>"
				    'Recipient=FirstName&" "&LastName
			        SentToEmail="mark.maggiore@logisticorp.us;FleetExpress@LogistiCorp.us"
                    'SentToEmail="mark.maggiore@logisticorp.us"
				    'Email="KWETI.Mailbox@am.kwe.com"
				    'Email="mark@maggiore.net"
				    Set objMail = CreateObject("CDONTS.Newmail")
				    objMail.From = "FleetX@LogisticorpGroup.com"
				    objMail.To = SentToEmail
				    'objMail.cc = RequestorEmailAddress
				    'objMail.cc = "4692269939@tmomail.net;pcurrin@ti.com"
				    'objMail.Subject = "New Fleet Express Shipment Request"
				    If trim(Priority)="Time Critical" then
                        objMail.Subject = "* TIME CRITICAL * New FleetX Shipment Request"
                        objMail.Importance = 2 'High
                        else
                        objMail.Subject = "New FleetX Shipment Request"
                    End if
				    objMail.MailFormat = cdoMailFormatMIME
				    objMail.BodyFormat = cdoBodyFormatHTML
				    objMail.Body = Body
				    objMail.Send
				    Set objMail = Nothing
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    If (int(DimLength)>42 or int(DimWidth)>48 or int(DimHeight)>48) AND xyz="removethisfornow" then
				        Body = "There has been an oversized Fleet Express shipment request (#"& newjobnum &") placed online:<br><br>"   

                        Body = Body & "REQUESTOR INFORMATION:<BR>"
                        Body = Body & "Name: "&  RequestorName &"<br>"  
                        Body = Body & "Phone Number: "&  RequestorPhoneNumber &"<br>"  
                        Body = Body & "Email Address: "&  RequestorEmailAddress &"<br>"   
                        Body = Body & "PO Number: "&  PONumber &"<br>"  
                        Body = Body & "Cost Center Number: "&  CostCenterNumber &"<br><br>" 
                        Body = Body & "COMMODITY INFORMATION:<BR>" 
                        Body = Body & "Pieces: "&  Pieces &" "& rf_box &"<br>" 
                        Body = Body & "Palletized: "&  IsPalletized &"<br>"   
                        'Body = Body & "Number Of Pallets: "&  NumberOfPallets &"<br>"  
                        Body = Body & "Weight: "&  DimWeight &"LBS<br>"
                        Body = Body & " Dimensions: "&  DimLength &" X "&  DimWidth &" X "&  DimHeight &" inches<br>"       
  
                    
                        'Body = Body & "Hazmat: "&  IsHazmat &"<br>"
                        'Body = Body & "Refrigerate: "&  Refrigerate &"<br><br>"
                        Body = Body & "ORIGINATION:<BR>"   
                        Body = Body & "Company: "&  OriginationCompany &"<br>"   
                        Body = Body & "Address: "&  OriginationAddress &"<br>"   
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
				        Body = Body & "FleetX Services<br>"  
				        Body = Body &  "FleetX@LogisticorpGroup.com<br>"  
				        Body = Body & "972/499-3415<br><br>"
				        'Recipient=FirstName&" "&LastName
			            SentToEmail="mark.maggiore@logisticorp.us;FleetExpress@LogistiCorp.us"
                        'SentToEmail="mark.maggiore@logisticorp.us"
				        'Email="KWETI.Mailbox@am.kwe.com"
				        'Email="mark@maggiore.net"
				        Set objMail = CreateObject("CDONTS.Newmail")
				        objMail.From = "FleetX@LogisticorpGroup.com"
				        objMail.To = SentToEmail
				        'objMail.cc = RequestorEmailAddress
				        'objMail.cc = "4692269939@tmomail.net;pcurrin@ti.com"
				        'objMail.Subject = "Oversized Fleet Express Shipment Request"
				        If trim(Priority)="Time Critical" then
                            objMail.Subject = "* TIME CRITICAL * Oversized FleetX Shipment Request"
                            objMail.Importance = 2 'High
                            else
                            objMail.Subject = "Oversized FleetX Shipment Request"
                        End if
				        objMail.MailFormat = cdoMailFormatMIME
				        objMail.BodyFormat = cdoBodyFormatHTML
				        objMail.Body = Body
				        objMail.Send
				        Set objMail = Nothing
                    End if
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


                   Response.Redirect("FleetExpressOrderConfirmation.asp?x=1&y=1&bid=86&pid=view&jid="& newjobnum &"&Internal="&Internal&"&XSquare="&XSquare)	
		    'End if	
        End if
    End if
    ''''''''END ERROR HANDLING''''''

     %>
</head>
<BODY leftMargin="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" OnLoad=document.OrderForm1.<%=HighlightedField%>.focus()>
<%
'''''''''Form
%>
<form method="post" name="OrderForm1" action="FleetExpressOrder.asp?Internal=<%=Internal%>">
    <table border="0" cellpadding="0" cellspacing="0" align="center">
        <tr><td align="left"><img src="images/pixel.gif" height="10" width="1" /></td></tr>
        <tr>
            <td align="left"><img src="images/logo_FleetExpress_space.gif" height="87" width="100" /></td>
            <td align="right" valign="bottom">
            <a href="FleetExpressTraining.pdf" target="_blank" class="<%=LinkClass%>">Click here to download training documentation</a><br />
            <a href="mailto:mark.maggiore@logisticorp.us" class="<%=LinkClass%>">Click here to report a problem with this page</a>
            </td>
        </tr>
        <tr><td align="left" colspan="2"><img src="images/pixel.gif" /></td></tr>
        <%If trim(UserId)>"" then%>
        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2">Expedited Transportation Request</td></tr>
        <%else %>
        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2">Fleet Express Login Page</td></tr>
        <%end if %>
        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2"><img src="images/pixel.gif" height="5" width="1" /></td></tr>
        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressBodyWhite" colspan="2">Please complete all areas below</td></tr>
        <tr><td align="left" colspan="2"><img src="images/pixel.gif" height="10" width="1" /></td></tr>
                  <%If trim(XSquare)="y" then %>
                <tr><td>&nbsp;</td></tr>
                <tr><td colspan="3"  class="FleetExpressTextBlackBold" align="center"><font color="blue">***XSquare Probe Card orders may only be placed between 7:00 AM and 12:00 PM Monday-Friday.***</font></td></tr>
               <tr><td>&nbsp;</td></tr>
                <%end if %>        
        <tr>
            <td align="center" colspan="2">
            <table border="0" bordercolor="<%=BorderColor%>" cellpadding="0" cellspacing="0">
            <tr> <td valign="top"><table border="1" bordercolor="<%=BorderColor%>" cellpadding="0" cellspacing="0" width="<%=tablewidth%>">
                <tr> <td valign="top"><table cellpadding="3" cellspacing="0" width="100%">
                <tr>
                    <td colspan="3" align="center" bgcolor="<%=BorderColor%>" class="FleetExpressBodyWhiteBold">
                        REQUESTOR INFORMATION
                    </td>
                </tr>
                <tr><td align="left" colspan="3"><input type="submit" name="buttonsubmit" value="Edit Requestor Information" /></td></tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap>Pre-Existing Requestor</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <%If Internal="y" then  %>
                    <td align="left">
								<%
                                'Response.write "PreExistingRequestor="&PreExistingRequestor&"<BR>"
                                'response.write "Database="&database&"<BR>"
                                 %>
                                <select name="PreExistingRequestor" ID="Select2"  onChange="form.submit()">
								<option value="" <%if trim(PreExistingRequestor)="" then response.Write " selected" end if%>>Select from this list or fill in all fields below</option>
                                <%
                                    
									Set oConn = Server.CreateObject("ADODB.Connection")
									oConn.ConnectionTimeout = 100
									oConn.Provider = "MSDASQL"
									oConn.Open DATABASE
                                        If trim(XSquare)="y" then
										    'l_cSQL = "Select * FROM PreExistingRequestor WHERE requestorstatus='c' and requestorID IN('1') ORDER BY RequestorName"
                                            l_cSQL = "Select * FROM PreExistingRequestor WHERE requestorstatus='c' and requestorID>'199' ORDER BY RequestorName"
                                            else
                                            l_cSQL = "Select * FROM PreExistingRequestor WHERE requestorstatus='c' ORDER BY RequestorName"
                                        End if
										SET oRs = oConn.Execute(l_cSql)
												Do While not oRs.EOF
                                                tempRequestorID=trim(oRs("RequestorID"))
												tempRequestorName=trim(oRs("RequestorName"))
							                    tempRequestorPhone=trim(oRs("RequestorPhone"))
                                                tempRequestorEmail=trim(oRs("RequestorEmail"))

                                                If trim(PreExistingRequestor)=trim(tempRequestorID) and trim(PreExistingRequestor)>"" then
												    RequestorName=TempRequestorName
												    RequestorPhoneNumber=TempRequestorPhone
                                                    RequestorEmailAddress=TempRequestorEmail
                                                End if								
											%>
											<option value="<%=TempRequestorID%>" <%If trim(TempRequestorID)=trim(UserID) then response.write "selected" end if %>><%=TempRequestorName%></option>
											<%
										oRs.movenext
										LOOP
									Set oConn=Nothing									
									%>
								</select> 
                    </td>
                    <%else 
                    	Set oConn = Server.CreateObject("ADODB.Connection")
							oConn.ConnectionTimeout = 100
							oConn.Provider = "MSDASQL"
							oConn.Open DATABASE
                                'If trim(XSquare)="y" then
									'''l_cSQL = "Select * FROM PreExistingRequestor WHERE requestorstatus='c' and requestorID IN('1') ORDER BY RequestorName"
                                   ' l_cSQL = "Select * FROM PreExistingRequestor WHERE requestorstatus='c' and requestorID>'199' ORDER BY RequestorName"
                                    'else
                                    l_cSQL = "Select * FROM PreExistingRequestor WHERE (requestorID='"& UserID &"') AND requestorstatus='c' ORDER BY RequestorName"
                                'End if
								SET oRs = oConn.Execute(l_cSql)
										If not oRs.EOF then
                                        tempRequestorID=trim(oRs("RequestorID"))
										tempRequestorName=trim(oRs("RequestorName"))
							            tempRequestorPhone=trim(oRs("RequestorPhone"))
                                        tempRequestorEmail=trim(oRs("RequestorEmail"))

                                        'If trim(PreExistingRequestor)=trim(tempRequestorID) and trim(PreExistingRequestor)>"" then
											RequestorName=TempRequestorName
											RequestorPhoneNumber=TempRequestorPhone
                                            RequestorEmailAddress=TempRequestorEmail
                                        'End if
                                    Response. write "<td align='left'>"&TempRequestorName&"</td>"
								End if
							Set oConn=Nothing
                end if %>
                </tr>
                <%
                'Response.write "RequestorName="&RequestorName&"*<BR>"
                'Response.write "GenericNumber="&GenericNumber&"*<BR>"
                If trim(RequestorName)="Jake Weber" and trim(GenericNumber)="" then
                    'Response.write "got here!!!<BR>"
                    GenericNumber="2883"
                End if               
                 %>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">Requestor Name</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left"><input type="hidden" name="RequestorName" value="<%=RequestorName%>" /><%=RequestorName%></td>
                </tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">Phone Number</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left"><input type="hidden" name="RequestorPhoneNumber" value="<%=RequestorPhoneNumber%>"/><%=RequestorPhoneNumber%></td>
                </tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">Email Address</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left"><input type="hidden" name="RequestorEmailAddress" value="<%=RequestorEmailAddress%>"/><%=RequestorEmailAddress%></td>
                </tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">
                            <select name="POorNWA">
                                <option value="Cost Center #"<%If trim(POorNWA)="Cost Center #" then Response.write "selected" end if%>>Cost Center #</option>
                                <option value="TI P/O #"<%If trim(POorNWA)="TI P/O #" then Response.write "selected" end if%>>TI P/O #</option>
                            </select>
                    </td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left"><textarea name="GenericNumber" rows="2" cols="30"><%=GenericNumber%></textarea></td>
                </tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" valign="top">Special Instructions</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left"><textarea name="comments" rows="2" cols="30"><%=Comments%></textarea></td>
                </tr>
                <tr><td><img src="images/pixel.gif" width="1" height="3" /></td></tr>
                </table>
                 </td></tr></table></td>
                 <td align="left"><img src="images/pixel.gif" height="1" width="25" /></td>
                 <td valign="top">
                    <table border="1" bordercolor="<%=BorderColor%>" cellpadding="0" cellspacing="0" width="<%=tablewidth%>"> 
                    <tr><td align="left"> <table cellpadding="3" cellspacing="0" width="100%">               <tr>
                        <td colspan="3" align="center" bgcolor="<%=BorderColor%>" class="FleetExpressBodyWhiteBold">
                            COMMODITY INFORMATION
                        </td>
                    </tr>
                    <!--
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Pieces</td>
                        <td width="10"><img src="images/pixel.gif" /></td>
                        <td nowrap class="FleetExpressTextBlack">
                            Number of Pieces:&nbsp;&nbsp;<input type="text" name="Pieces" value="<%=Pieces%>" size="3" maxlength="4" />
                        </td>
                    </tr>
                    -->
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Number of Pieces</td>
                        <td width="10"><img src="images/pixel.gif" /></td>
                        <td nowrap class="FleetExpressTextBlackBold">
                            <%If trim(XSquare)="y" then %>
                                <select name="pieces">
                                    <option value="1"<%If trim(pieces)="1" or trim(pieces)="" then Response.write "selected" end if%>>1</option>
                                    <option value="2"<%If trim(pieces)="2" then Response.write "selected" end if%>>2</option>
                                </select>
                            <%else %>
                                <input type="text" value="<%=pieces%>" name="pieces" size="3" maxlength="4" />
                            <%end if %>
                            &nbsp;&nbsp;
                            <%If trim(XSquare)="y" then %>
                            <input type="hidden" name="rf_box" value="X Square Probe Card(s)">
                            X Square Probe Card(s)
                                    
                            <%else %>
                            <select name="rf_box">
                                <option value="Boxes"<%If trim(rf_box)="Boxes" then Response.write "selected" end if%>>Boxes</option>
                                <option value="Crates"<%If trim(rf_box)="Crates" then Response.write "selected" end if%>>Crates</option>
                                <option value="Envelopes"<%If trim(rf_box)="Envelopes" then Response.write "selected" end if%>>Envelopes</option>
                                <option value="Skids"<%If trim(rf_box)="Skids" then Response.write "selected" end if%>>Skids</option>
                            </select>
                            <%end if %>
                        </td>
                    </tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Palletization</td>
                        <td width="10"><img src="images/pixel.gif" /></td>
                        <td nowrap class="FleetExpressTextBlackBold">
                            <select name="IsPalletized">
                                <option value="y"<%If trim(IsPalletized)="y" then Response.write "selected" end if%>>Palletized</option>
                                <option value="n"<%If trim(IsPalletized)="n" then Response.write "selected" end if%>>Not Palletized</option>
                                <option value="Trailer Only Move"<%If trim(IsPalletized)="Trailer Only Move" then Response.write "selected" end if%>>Trailer Only Move</option>
                            </select>
                        </td>
                    </tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Weight</td>
                        <td width="10"><img src="images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><input type="text" name="DimWeight" value="<%=DimWeight%>" size="6" maxlength="5" /> Pounds</td>
                    </tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Dimensions</td>
                        <td width="10"><img src="images/pixel.gif" /></td>
                        <td class="FleetExpressTextBlack" align="left" nowrap>
                            L:&nbsp;&nbsp;<input type="text" name="DimLength" value="<%=DimLength%>" size="5"  maxlength="4"/> 
                            W:&nbsp;&nbsp;<input type="text" name="DimWidth" value="<%=DimWidth%>" size="5" maxlength="4" /> 
                            H:&nbsp;&nbsp;<input type="text" name="DimHeight" value="<%=DimHeight%>" size="5"  maxlength="4"/> 
                            &nbsp;&nbsp;Inches
                        </td>
                    </tr>

                    <!--
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Hazmat</td>
                        <td width="10"><img src="images/pixel.gif" /></td>
                        <td align="left"  class="FleetExpressTextBlackBold">
                             <select name="IsHazmat">
                                <option value="n" <%If trim(IsHazmat)="n" then Response.write "selected" end if%>>No</option>
                                <option value="y" <%If trim(IsHazmat)="y" then Response.write "selected" end if%>>Yes</option>
                            </select>                     
                        </td>
                    </tr>
                    -->
                    <input type="hidden" name="IsHazmat" value="n" />
                    <!--
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Refrigerate</td>
                        <td width="10"><img src="images/pixel.gif" /></td>
                        <td align="left"  class="FleetExpressTextBlackBold">
                             <select name="Refrigerate">
                                <option value="n" <%If trim(Refrigerate)="n" then Response.write "selected" end if%>>No</option>
                                <option value="y" <%If trim(Refrigerate)="y" then Response.write "selected" end if%>>Yes</option>
                            </select>                      
                        </td>
                    </tr>
                    -->
                    <input type="hidden" name="Refrigerate" value="n" />
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Service Level</td>
                        <td width="10"><img src="images/pixel.gif" /></td>
                        <td align="left"  class="FleetExpressTextBlackBold">
                        <%if trim(XSquare)="y" then %>
                              <select name="Priority" onchange="AnyCost();">
                                <option value="Four Hours" <%If trim(Priority)="Four Hours" then Response.write "selected" end if%>>Four Hours</option>
                            </select>  
                       <%else%>                           
                             <select name="Priority" onchange="AnyCost();">
                                <option value="Next Day" <%If trim(Priority)="Next Day" then Response.write "selected" end if%>>Next Day</option>
                                <!--option value="Same Day" <%If trim(Priority)="Same Day" then Response.write "selected" end if%>>Same Day</option-->
                                <option value="2 Hour" <%If trim(Priority)="2 Hour" then Response.write "selected" end if%>>2 Hour</option>
                                <option value="4 Hour" <%If trim(Priority)="4 Hour" then Response.write "selected" end if%>>4 Hour</option>
                                <option value="Time Critical" <%If trim(Priority)="Time Critical" then Response.write "selected" end if%>>Time Critical</option>
                            </select>
                       <%end if %>                      
                        </td>
                    </tr>
                    <tr><td><img src="images/pixel.gif" height="83" width="1" /></td></tr>
                     </table>
                     </td></tr></table>                 
                 </td></tr> 
                                                                                                                           
            </table>
         
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
         %>

        <tr><td align="left"><img src="images/pixel.gif" height="10" width="1" /></td></tr>
        <tr>
            <td align="center"  colspan="2">
            <table border="0" bordercolor="<%=BorderColor%>" cellpadding="0" cellspacing="0">
            <tr><td align="left"><table border="1" bordercolor="<%=BorderColor%>" cellpadding="0" cellspacing="0" width="<%=tablewidth%>">
                <tr> <td valign="top"><table cellpadding="3" cellspacing="0" width="100%">
                <tr>
                    <td colspan="3" align="center" bgcolor="<%=BorderColor%>" class="FleetExpressBodyWhiteBold">
                        ORIGINATION 
                    </td>
                </tr>
                <tr><td align="left" colspan="3"><input type="submit" name="buttonsubmit" value="Add/Edit Locations in your Address Book" /></td></tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap>Pre-Existing Company</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left">
								<%
                                'Response.write "PreExistingOrigination="&PreExistingOrigination&"<BR>"
                                 %>
                                <select name="PreExistingOrigination" ID="Select3"  onChange="form.submit()">
								<option value="" <%if trim(PreExistingOrigination)="" then response.Write " selected" end if%>>Select your origination</option>
                                <%

									Set oConn = Server.CreateObject("ADODB.Connection")
									oConn.ConnectionTimeout = 100
									oConn.Provider = "MSDASQL"
									oConn.Open DATABASE
                                        If trim(XSquare)="y" then
                                            l_cSQL = "Select * FROM PreExistingCompanies WHERE companystatus='c' and (CompanyName='EBT Probe Card Ship Room' or CompanyName='SC Building Probe Card Shop') ORDER BY CompanyName"
                                        else
										    l_cSQL = "Select * FROM PreExistingCompanies WHERE companystatus='c' and CompanyOwner='"& trim(UserID) & "' ORDER BY CompanyName"
                                        End if
										SET oRs = oConn.Execute(l_cSql)
												Do While not oRs.EOF
                                                CompanyID=oRs("CompanyID")
												CompanyName=oRs("CompanyName")
												CompanyAddress=oRs("CompanyAddress")
												CompanyCity=oRs("CompanyCity")
                                                CompanyState=oRs("CompanyState")
                                                CompanyZip=oRs("CompanyZip")
                                                ContactName=oRs("ContactName")
                                                CompanyPhone=oRs("CompanyPhone")
                                                CompanyEmail=oRs("CompanyEmail")
                                                If trim(PreExistingOrigination)=trim(CompanyID) and trim(PreExistingOrigination)>"" then
												    OriginationID=CompanyID
                                                    
                                                    OriginationCompany=CompanyName
												    OriginationAddress=CompanyAddress
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
											%>
											<option value="<%=CompanyID%>" <%if trim(PreExistingOrigination)=trim(CompanyID) then response.Write " selected" end if%>><%=CompanyName%></option>
											<%
										oRs.movenext
										LOOP
									Set oConn=Nothing									
									%>
								</select> 
                    </td>
                </tr>
                <%
                'Response.write "XXXXOriginationID="&OriginationID&"<BR>"
                 %>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap>Company Name</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left"><%=OriginationCompany %><input type="hidden" name="OriginationCompany" value="<%=OriginationCompany%>" size="45" maxlength="40" /></td>
                </tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">Address</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left"><%=OriginationAddress%><input type="hidden" name="OriginationAddress" value="<%=OriginationAddress%>" size="45" maxlength="40" /></td>
                </tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">City/State/Zip</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td class="FleetExpressTextBlack"><%=OriginationCity%>
                       
                        <input type="hidden" name="OriginationCity" value="<%=OriginationCity%>" size="20" maxlength="30" />
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
                        <%
                        'Response.write "OriginationID="&OriginationID&"<BR>" 
                        %>
                        <input type="hidden" name="OriginationID" value="<%=OriginationID%>" />
                        <input type="hidden" name="OriginationState" value="TX"><%=OriginationZipCode%>
                        <input type="Hidden" name="OriginationZipCode" value="<%=OriginationZipCode%>" size="11" maxlength="10" />

                    </td>
                </tr>


                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">Contact Name</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left"><input type="hidden" name="OriginationContactName" value="<%=OriginationContactName%>" size="45" maxlength="25" /><%=OriginationContactName%></td>
                </tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">Phone Number</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left"><%=OriginationPhoneNumber%><input type="hidden" name="OriginationPhoneNumber" value="<%=OriginationPhoneNumber%>" size="45" maxlength="20" /></td>
                </tr> 
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">Email Address</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left"><%=OriginationEmail%><input type="hidden" name="OriginationEmail" value="<%=OriginationEmail%>" size="45" maxlength="100" /></td>
                </tr> 
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">Receive Notifications</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left"><input type="radio" name="OriginationNotifications" value="y" <%If trim(OriginationNotifications)="" or trim(OriginationNotifications)="y" then Response.write "Checked" end if %>/>Yes&nbsp;&nbsp;&nbsp;&nbsp;<input type="radio" name="OriginationNotifications" value="n" <%If trim(OriginationNotifications)="n" then Response.write "Checked" end if %>/>No</td>
                </tr>
                <%
                If trim(XSquare)="y" then

                else 
                %>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">Ready Date/Time</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left"><input type="text" id="PickupDateTime" name="PickupDateTime" value="<%=PickUpDateTime%>" size="30" maxlength="30" />
                    <a href="javascript:NewCal('PickupDateTime','MMddyyyy',true,12)"><img src="images/cal.gif" width="16" height="16" border="0" alt="Pick a date"></a>                 
                    </td>
                </tr>
                <%end if %>
                 </table>
                 </td></tr></table></td>
                 <td align="left"><img src="images/pixel.gif" height="1" width="25" /></td>
                 <td align="left">
                    <table border="1" bordercolor="<%=BorderColor%>" cellpadding="0" cellspacing="0" width="<%=tablewidth%>"> 
                    <tr> <td valign="top"> <table cellpadding="3" cellspacing="0" width="100%">               <tr>
                        <td colspan="3" align="center" bgcolor="<%=BorderColor%>" class="FleetExpressBodyWhiteBold">
                            DESTINATION
                        </td>
                    </tr>
                    <tr><td align="left" colspan="3"><input type="submit" name="buttonsubmit" value="Add/Edit Locations in your Address Book" /></td></tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap>Pre-Existing Company</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left">
								<%
                                'Response.write "PreExistingDestination="&PreExistingDestination&"<BR>"
                                 %>
                                <select name="PreExistingDestination" ID="Select1"  onChange="form.submit()">
								<option value="" <%if trim(PreExistingDestination)="" then response.Write " selected" end if%>>Select your destination</option>
                                <%

									Set oConn = Server.CreateObject("ADODB.Connection")
									oConn.ConnectionTimeout = 100
									oConn.Provider = "MSDASQL"
									oConn.Open DATABASE
                                        If trim(XSquare)="y" then
                                            l_cSQL = "Select * FROM PreExistingCompanies WHERE companystatus='c' and (CompanyName='EBT Probe Card Ship Room' or CompanyName='SC Building Probe Card Shop') ORDER BY CompanyName"
                                            else
										    l_cSQL = "Select * FROM PreExistingCompanies WHERE companystatus='c' and CompanyOwner='"& trim(UserID) & "' ORDER BY CompanyName"
                                        End if
										SET oRs = oConn.Execute(l_cSql)
												Do While not oRs.EOF
                                                bCompanyID=oRs("CompanyID")
												bCompanyName=oRs("CompanyName")
												bCompanyAddress=oRs("CompanyAddress")
												bCompanyCity=oRs("CompanyCity")
                                                bCompanyState=oRs("CompanyState")
                                                bCompanyZip=oRs("CompanyZip")
                                                bContactName=oRs("ContactName")
                                                bCompanyPhone=oRs("CompanyPhone")
                                                bCompanyEmail=oRs("CompanyEmail")
                                                If trim(PreExistingDestination)=trim(bCompanyID) and trim(PreExistingDestination)>"" then
												    DestinationID=bCompanyID
                                                    DestinationCompany=bCompanyName
												    DestinationAddress=bCompanyAddress
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
											%>
											<option value="<%=bCompanyID%>" <%if trim(PreExistingDestination)=trim(bCompanyID) then response.Write " selected" end if%>><%=bCompanyName%></option>
											<%
										oRs.movenext
										LOOP
									Set oConn=Nothing									
									%>
								</select> 
                    </td>
                </tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left" nowrap>Company Name</td>
                        <td width="10"><img src="images/pixel.gif" /></td>
                        <td align="left"><%=DESTINATIONCompany%><input type="hidden" name="DESTINATIONCompany" value="<%=DESTINATIONCompany%>" size="45" maxlength="40" /></td>
                    </tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Address</td>
                        <td width="10"><img src="images/pixel.gif" /></td>
                        <td align="left"><%=DESTINATIONAddress%><input type="hidden" name="DESTINATIONAddress" value="<%=DESTINATIONAddress%>" size="45" maxlength="40" /></td>
                    </tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">City/State/Zip</td>
                        <td width="10"><img src="images/pixel.gif" /></td>
                        <td class="FleetExpressTextBlack">
                            <%=DESTINATIONCity%><input type="hidden" name="DESTINATIONCity" value="<%=DESTINATIONCity%>" size="20" maxlength="30" />
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
                        <%
                        'Response.write "DestinationID="&DestinationID&"<BR>" 
                        %>
                        <input type="hidden" name="DestinationID" value="<%=DestinationID%>" />
                        <input type="hidden" name="DestinationState" value="TX">
                            <%=DESTINATIONZipCode%><input type="hidden" name="DESTINATIONZipCode" value="<%=DESTINATIONZipCode%>" size="11" maxlength="10" />

                        </td>
                    </tr>


                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Contact Name</td>
                        <td width="10"><img src="images/pixel.gif" /></td>
                        <td align="left"><%=DESTINATIONContactName%><input type="hidden" name="DESTINATIONContactName" value="<%=DESTINATIONContactName%>" size="45" maxlength="25" /></td>
                    </tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Phone Number</td>
                        <td width="10"><img src="images/pixel.gif" /></td>
                        <td align="left"><%=DESTINATIONPhoneNumber%><input type="hidden" name="DESTINATIONPhoneNumber" value="<%=DESTINATIONPhoneNumber%>" size="45" maxlength="20" /></td>
                    </tr> 
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Email Address</td>
                        <td width="10"><img src="images/pixel.gif" /></td>
                        <td align="left"><%=DESTINATIONEmail%><input type="hidden" name="DESTINATIONEmail" value="<%=DESTINATIONEmail%>" size="45" maxlength="100" /></td>
                    </tr>
                 <tr>
                    <td class="FleetExpressTextBlackBold" align="left">Receive Notifications</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left"><input type="radio" name="DestinationNotifications" value="y" <%If trim(DestinationNotifications)="" or trim(DestinationNotifications)="y" then Response.write "Checked" end if %>/>Yes&nbsp;&nbsp;&nbsp;&nbsp;<input type="radio" name="DestinationNotifications" value="n" <%If trim(DestinationNotifications)="n" then Response.write "Checked" end if %>/>No</td>
                </tr>                    
                <%If trim(XSquare)="y" then

                else %>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">Delivery Date/Time</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left">
                    <input type="text" name="DeliveryDateTime" id="DeliveryDateTime" value="<%=DeliveryDateTime%>" size="30" maxlength="30" />
                     <a href="javascript:NewCal('DeliveryDateTime','MMddyyyy',true,12)"><img src="images/cal.gif" width="16" height="16" border="0" alt="Pick a date"></a>
                    </td>
                </tr>
                <%end if %>
                     </table>
                     </td></tr></table>                 
                 </td></tr>
                                                                                                           
            </table>
         
            </td>
        </tr>
        
        
         <tr><td align="left"><img src="images/pixel.gif" height="10" width="1" /></td></tr>
         <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2">Fleet Express Transportation Call Center 972-499-3415</td></tr>
         <tr><td align="left"><img src="images/pixel.gif" height="10" width="1" /></td></tr>
        <input type="hidden" value="1" name="Timesthrough" />
          <tr>
            <td align="center" colspan="2">
                <%If Errormessage>"" then%>              
                <table cellpadding="2" cellspacing="2" border="1" bordercolor="red">
                    <tr>
                        <td class="errormessage">
                            <%
                            Response.write " * * * ERROR:  "&ErrorMessage& " * * * "
                             %>
                        </td>
                    </tr>
                </table>
                <%End if %>
           </td>
           </tr>        
         <tr>
            <td align="left" valign="top">
            &nbsp;
           </td>
           <td align="right">
                  <input type="image" src="images/submit_fleetexpress2.gif" alt="submit order" /> 
            </td></tr>   
        <input type="hidden" name="ColorSelect" value="<%=ColorSelect %>" />
        <input type="hidden" name="MarkTemp" value="<%=MarkTemp %>" />
        <input type="hidden" name="CaptchaSubmit" value="<%=CaptchaSubmit %>" />
        <input type="hidden" name="varCaptcha" value="<%=varCaptcha %>" />
        <input type="hidden" name="XSquare" value="<%=XSquare %>" />
        <input type="hidden" name="UserID" value="<%=UserID %>" />
        
        <input type="hidden" name="pagestatus" value="submit" />
    </table>
</form>
<%
else
%>
<BODY leftMargin="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" OnLoad=document.OrderForm1.<%=HighlightedField%>.focus()>
<%
'''''''''Form
%>
<form method="post" name="OrderForm1" action="FleetExpressOrder.asp?Internal=<%=Internal%>">
    <table border="0" cellpadding="0" cellspacing="0" align="center" width="80%">
        <tr><td align="left"><img src="images/pixel.gif" height="10" width="1" /></td></tr>
        <tr>
            <td align="left"><img src="images/logo_FleetExpress_space.gif" height="87" width="100" /></td>
            <td align="right" valign="bottom">
            <a href="FleetExpressTraining.pdf" target="_blank" class="<%=LinkClass%>">Click here to download training documentation</a><br />
            <a href="mailto:mark.maggiore@logisticorp.us" class="<%=LinkClass%>">Click here to report a problem with this page</a>
            </td>
        </tr>
        <tr><td align="left" colspan="2" bgcolor="<%=HeaderBorderColor%>"><img src="images/pixel.gif" height="10" width="1" /></td></tr>
        <%If trim(UserId)>"" then%>
        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2">Expedited Transportation Request</td></tr>
        <%else %>
        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2">Fleet Express Login Page</td></tr>
        <%end if %>
        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2"><img src="images/pixel.gif" height="5" width="1" /></td></tr>
        <!--tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressBodyWhite" colspan="2">In order to reach the Fleet Express Order page, please correctly type in the green verification code in the supplied text box and click "Submit."</td></tr-->
        <tr><td align="left" colspan="2"><img src="images/pixel.gif" height="10" width="1" /></td></tr>
        <tr>
<%
'Response.write "HeaderBorderColor="& HeaderBorderColor &"<BR>"
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
  <tr><td colspan="2"><b>Log in below.  If you've forgotten your username/password, <a href="FleetExpressGetPassword.asp">click here</a>.  If you're a new user and need to create an account,
  then <a href="FleetExpressUserInfo.asp">click here</a></b></td></tr>
  <tr><td>&nbsp;</td></tr>
  <tr><td colspan="2"><b>***There have been several recent enhancements to this application, so <a href="FleetExpressTraining.pdf" target="_blank">click here</a>  to download training documentation.***</b></td></tr>
  <tr><td>&nbsp;</td></tr>
<tr><td>&nbsp;</td></tr>
<tr><td>&nbsp;</td></tr>
        <tr><td colspan="2" align="center"><table border="0" cellpadding="0" cellspacing="0">
        <tr>

            <td NOWRAP valign="middle" align="left" class="MainPageText"> 
          User Name:</td><td>
          <input type="text" value="<%=UserName %>" name="UserName">
            </td>
      </tr>
      <tr><td>&nbsp;</td></tr>
        <tr>

            <td NOWRAP valign="middle" align="left" class="MainPageText"> 
          Password:</td><td>
          <input type="Password" value="<%=Password %>" name="Password">
            </td>
      </tr>
    
   <tr><td>&nbsp;</td></tr>
  <tr>
    <td colspan="2" align="center">
        <img src="../../images/captcha/<%=CAPTCHA1%>.gif" height="37" width="37" border="0" />
        <img src="../../images/captcha/<%=CAPTCHA2%>.gif" height="37" width="37" border="0" />
        <img src="../../images/captcha/<%=CAPTCHA3%>.gif" height="37" width="37" border="0" />
        <img src="../../images/captcha/<%=CAPTCHA4%>.gif" height="37" width="37" border="0" />
        <img src="../../images/captcha/<%=CAPTCHA5%>.gif" height="37" width="37" border="0" />
    </td>
  </tr>

         <input type="hidden" name="varCAPTCHA" value="<% = varCAPTCHA %>" />

      <tr>

            <td NOWRAP valign="middle" align="center" class="MainPageText"> 
          Verification Code:</td><td>
          <input name="CAPTCHASubmit">
            </td>
      </tr>
      <tr><td>&nbsp;</td></tr>      
      <tr>
        <td align="center" colspan="2">
           <input name="submit" type="submit" value="Submit" />
           <input name="loginattempt" value="y" type="hidden" />
        </td>
      </tr>
      <input type="hidden" name="XSquare" value="<%=XSquare %>" />
  </form>


<%
  ' Delete the captchas object.
  Set captchas = Nothing
%>
	<tr Height="30">
		<td>&nbsp;</td>
	</tr>
</table></td></tr>
</table>
<%
'Response.write "ErrorMessage="&ErrorMessage&"<BR>"
if ErrorMessage>"" then%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" class="MainPageText" ID="Table5">
	<tr>
    <td align="center" class="errormessage"><%=ErrorMessage%></td>
  </tr>
	<tr><td>&nbsp;</td></tr>
</table>
<%
end if
end if
'Response.write "PageStatus="&PageStatus&"<BR>"
%>
</body>
</html>
