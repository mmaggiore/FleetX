<%@ LANGUAGE="VBSCRIPT"%>
<!-- #include file="../fleetexpress.inc" -->
<html>
<head>

<link rel="stylesheet" type="text/css" href="../css/Style.css">
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

<%
'Response.write "Hello 1<BR>"
    DimWeight=0
    DimHeight=0
    DimWidth=0
    DimLength=0
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
    PageTitle="Freight/Courier Order Page"



'Response.write "UserID="&UserID&"<BR>"
'Response.write "sBT_ID="&sBT_ID&"<BR>"

   '''''''''''HARDCODED STUFF
   sBT_ID=BillToID
    'UserID=Request.Form("UserID")
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
   CurrentHour=Hour(Now())
   CurrentMinute=Minute(Now())

      ''''''DATA COOKIE CAPBOB''''''''''''''''''''
    CAPBOB=trim(Request.Form("CAPBOB"))
    If trim(CAPBOB)>"" then
        Response.Cookies ("MyCookie")("CAPBOB")=CAPBOB
    End if
    CAPBOB=Request.Cookies("MyCookie")("CAPBOB")
   'Response.write "CAPBOB="&CAPBOB&"<br>"
    '''''''''''''''''''''''''
      ''''''DATA COOKIE IsStandingOrder''''''''''''''''''''
    IsStandingOrder=Request.Form("IsStandingOrder")
    If trim(IsStandingOrder)>"" then
        Response.Cookies ("MyCookie")("IsStandingOrder")=IsStandingOrder
    End if
    IsStandingOrder=Request.Cookies("MyCookie")("IsStandingOrder")
   'Response.write "ShipmentType="&ShipmentType&"<br>"
    '''''''''''''''''''''''''
      ''''''DATA COOKIE ShipmentType''''''''''''''''''''
    PONumber=Request.Form("PONumber")
    If trim(PONumber)>"" then
        Response.Cookies ("MyCookie")("PONumber")=PONumber
    End if
    PONumber=Request.Cookies("MyCookie")("PONumber")
   'Response.write "ShipmentType="&ShipmentType&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE ShipmentType''''''''''''''''''''
    RequestorCompany=Request.Form("RequestorCompany")
    If trim(RequestorCompany)>"" then
        Response.Cookies ("FleetXCookie")("RequestorCompany")=RequestorCompany
    End if
    RequestorCompany=Request.Cookies("FleetXCookie")("RequestorCompany")
    'Response.write "RequestorCompany="&RequestorCompany&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE ShipmentType''''''''''''''''''''
    PickUpDate=Request.Form("PickUpDate")
    If trim(PickUpDate)>"" then
        Response.Cookies ("MyCookie")("PickUpDate")=PickUpDate
    End if
    PickUpDate=Request.Cookies("MyCookie")("PickUpDate")
    'Response.write "PickUpDate="&PickUpDate&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE ShipmentType''''''''''''''''''''
    PickUpTime=Request.Form("PickUpTime")
    If trim(PickUpTime)>"" then
        Response.Cookies ("MyCookie")("PickUpTime")=PickUpTime
    End if
    PickUpTime=Request.Cookies("MyCookie")("PickUpTime")
    'Response.write "PickUpTime="&PickUpTime&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE PickUpDateTime''''''''''''''''''''
    'PickUpDateTime=Request.Form("PickUpDateTime")
    'If trim(PickUpDateTime)>"" then
   '     Response.Cookies ("MyCookie")("PickUpDateTime")=PickUpDateTime
    'End if
    'PickUpDateTime=Request.Cookies("MyCookie")("PickUpDateTime")
    'Response.write "PickUpDateTime="&PickUpDateTime&"<br>"
    '''''''''''''''''''''''''
    PickUpDateTime=PickUpDate&" "&PickUpTime
    'REsponse.write "PickUpDateTime="&PickUpDateTime&"<BR>"

   ''''''DATA COOKIE ShipmentType''''''''''''''''''''
    ShipmentType=Request.Form("ShipmentType")
    If trim(ShipmentType)>"" then
        Response.Cookies ("MyCookie")("ShipmentType")=ShipmentType
    End if
    ShipmentType=Request.Cookies("MyCookie")("ShipmentType")
   'Response.write "ShipmentType="&ShipmentType&"<br>"
    '''''''''''''''''''''''''
    'Select Case ShipmentType
    '    Case "Light Package"
    '        BillToID="92"
    '    Case "Heavy Freight"
    '        BillToID="93"
    'End Select
    ''''''DATA COOKIE OrderStatus''''''''''''''''''''
    OrderStatus=Request.Form("OrderStatus")
    FuelChargeDollars=Request.Form("FuelChargeDollars")

    If orderstatus="C" then
        'REsponse.write "Got here C!<BR>"
        '''''''''''''''''WIPE OUT ALL THE COOKIES!!!'''''''''''''''''
        Response.Cookies ("MyCookie")("CAPBOB")=""
        Response.Cookies ("MyCookie")("AddSkids")=""
		Response.Cookies ("MyCookie")("AddLargeSkids")=""
        Response.Cookies ("MyCookie")("AddSkidsCost")=""
		Response.Cookies ("MyCookie")("AddLargeSkidsCost")=""
        Response.Cookies ("MyCookie")("FuelCharge")=""
        Response.Cookies ("MyCookie")("RtDescr")=""
        Response.Cookies ("MyCookie")("RateCharge")=""
        Response.Cookies ("MyCookie")("rtBillCode")=""    
        Response.Cookies ("MyCookie")("IsStandingOrder")=""
        Response.Cookies ("MyCookie")("MaterialDescription")=""
        Response.Cookies ("MyCookie")("PickUpDate")=""
        Response.Cookies ("MyCookie")("PickUpTime")=""
        Response.Cookies ("MyCookie")("ShipmentType")=""
        Response.Cookies ("MyCookie")("OrigEmail")=""
        Response.Cookies ("MyCookie")("DestEmail")=""
        Response.Cookies ("MyCookie")("OtherEmail")=""
        Response.Cookies ("MyCookie")("AdditionalEmail")=""
        Response.Cookies ("MyCookie")("XSquare")=""
        Response.Cookies ("MyCookie")("MarkTemp")=""
        Response.Cookies ("MyCookie")("Username")=""
        Response.Cookies ("MyCookie")("Password")=""
        Response.Cookies ("MyCookie")("ShipmentType")=""
        Response.Cookies ("MyCookie")("TimesThrough")=""
        Response.Cookies ("MyCookie")("PreExistingRequestor")=""
        Response.Cookies ("MyCookie")("PreExistingOrigination")=""
        Response.Cookies ("MyCookie")("PreExistingDestination")=""
        Response.Cookies ("MyCookie")("RequestorName")=""
        Response.Cookies ("MyCookie")("RequestorPhoneNumber")=""
        Response.Cookies ("MyCookie")("RequestorEmailAddress")=""
        Response.Cookies ("MyCookie")("Pieces")=""
        Response.Cookies ("MyCookie")("NumberOfPallets")=""
        Response.Cookies ("MyCookie")("DimWeight")=""
        Response.Cookies ("MyCookie")("DimLength")=""
        Response.Cookies ("MyCookie")("DimWidth")=""
        Response.Cookies ("MyCookie")("DimHeight")=""
        Response.Cookies ("MyCookie")("IsPalletized")=""
        Response.Cookies ("MyCookie")("IsStacked")=""
        Response.Cookies ("MyCookie")("DimValue")=""
        Response.Cookies ("MyCookie")("IsHazmat")=""
        Response.Cookies ("MyCookie")("OriginationID")=""
        Response.Cookies ("MyCookie")("OriginationCompany")=""
        Response.Cookies ("MyCookie")("OriginationBuilding")=""
        Response.Cookies ("MyCookie")("OriginationAddress")=""
        Response.Cookies ("MyCookie")("OriginationSuite")=""
        Response.Cookies ("MyCookie")("OriginationCity")=""
        Response.Cookies ("MyCookie")("OriginationState")=""
        Response.Cookies ("MyCookie")("OriginationZipCode")=""
        Response.Cookies ("MyCookie")("OriginationIsCourier")=""
        Response.Cookies ("MyCookie")("OriginationAliasCode")=""
        Response.Cookies ("MyCookie")("OriginationContactName")=""
        Response.Cookies ("MyCookie")("OriginationPhoneNumber")=""
        Response.Cookies ("MyCookie")("OriginationEmail")=""
        Response.Cookies ("MyCookie")("DestinationID")=""
        Response.Cookies ("MyCookie")("DestinationCompany")=""
        Response.Cookies ("MyCookie")("DestinationBuilding")=""
        Response.Cookies ("MyCookie")("DestinationAddress")=""
        Response.Cookies ("MyCookie")("DestinationSuite")=""
        Response.Cookies ("MyCookie")("DestinationCity")=""
        Response.Cookies ("MyCookie")("DestinationState")=""
        Response.Cookies ("MyCookie")("DestinationZipCode")=""
        Response.Cookies ("MyCookie")("DestinationIsCourier")=""
        Response.Cookies ("MyCookie")("DestinationAliasCode")=""
        Response.Cookies ("MyCookie")("DestinationContactName")=""
        Response.Cookies ("MyCookie")("DestinationPhoneNumber")=""
        Response.Cookies ("MyCookie")("DestinationEmail")=""
        Response.Cookies ("MyCookie")("OriginationIsCourier")=""
        Response.Cookies ("MyCookie")("POorNWA")=""
        Response.Cookies ("MyCookie")("GenericNumber")=""
        Response.Cookies ("MyCookie")("bGenericNumber")=""
        Response.Cookies ("MyCookie")("PONumber")=""
        Response.Cookies ("MyCookie")("OriginationNotifications")=""
        Response.Cookies ("MyCookie")("RequestorNotifications")=""
        Response.Cookies ("MyCookie")("OriginationWaybill")=""
        Response.Cookies ("MyCookie")("DestinationNotifications")=""
        Response.Cookies ("MyCookie")("Comments")=""
        Response.Cookies ("MyCookie")("rf_box")=""
        Response.Cookies ("MyCookie")("Refrigerate")=""
        Response.Cookies ("MyCookie")("Priority")=""
        Response.Cookies ("MyCookie")("DeliveryDateTime")=""
        Response.Cookies ("MyCookie")("Submit")=""
        Response.Cookies ("MyCookie")("VarA")=""
For Each cookie in Response.Cookies
    'REsponse.write "got here!!!<BR>"
    Response.Cookies("MyCookie").Expires = DateAdd("d",-1,now())
Next

'''''''''''''''''END - WIPE OUT ALL THE COOKIES!!!'''''''''''''''''
    End if


    var1b=Request.Form("var1b")
    If lcase(Var1b)="<<<back" then
        OrderStatus="1b"
    End if
    var2=Request.Form("var2")
    If lcase(Var2)="<<<back" then
        OrderStatus="2"
    End if
    var3=Request.Form("var3")
    If lcase(Var3)="<<<back" then
        OrderStatus="3"
    End if
    Var1=Request.QueryString("Var1")
    If trim(Var1)>"" then
        OrderStatus=Var1
    End if
    'If trim(OrderStatus)>"" then
    '    Response.Cookies ("MyCookie")("OrderStatus")=OrderStatus
    'End if
    'OrderStatus=Request.Cookies("MyCookie")("OrderStatus")
    'Response.write "OrderStatus="&OrderStatus&"<br>"
    'Response.Cookies ("MyCookie")("OrderStatus")=1
    '''''''''''''''''''''''''

   ''''''DATA COOKIE OrigEmail''''''''''''''''''''
    MaterialDescription=Request.Form("MaterialDescription")
    If trim(MaterialDescription)>"" then
        Response.Cookies ("MyCookie")("MaterialDescription")=MaterialDescription
    End if
    MaterialDescription=Request.Cookies("MyCookie")("MaterialDescription")
    'Response.write "OrigEmail="&OrigEmail&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE AddSkids''''''''''''''''''''
    AddSkids=Request.Form("AddSkids")
    If trim(AddSkids)>"" then
        Response.Cookies ("MyCookie")("AddSkids")=AddSkids
    End if
    AddSkids=Request.Cookies("MyCookie")("AddSkids")
    'Response.write "OrigEmail="&OrigEmail&"<br>"
   ''''''DATA COOKIE AddSkidsCost''''''''''''''''''''
    AddSkidsCost=Request.Form("AddSkidsCost")
    If trim(AddSkidsCost)>"" then
        Response.Cookies ("MyCookie")("AddSkidsCost")=AddSkidsCost
    End if
    AddSkidsCost=Request.Cookies("MyCookie")("AddSkidsCost")
    'Response.write "OrigEmail="&OrigEmail&"<br>"
   ''''''DATA COOKIE AddLargeSkids''''''''''''''''''''
    AddLargeSkids=Request.Form("AddLargeSkids")
    If trim(AddLargeSkids)>"" then
        Response.Cookies ("MyCookie")("AddLargeSkids")=AddLargeSkids
    End if
    AddLargeSkids=Request.Cookies("MyCookie")("AddLargeSkids")
    'Response.write "OrigEmail="&OrigEmail&"<br>"
   ''''''DATA COOKIE AddLargeSkidsCost''''''''''''''''''''
    AddLargeSkidsCost=Request.Form("AddLargeSkidsCost")
    If trim(AddLargeSkidsCost)>"" then
        Response.Cookies ("MyCookie")("AddLargeSkidsCost")=AddLargeSkidsCost
    End if
    AddLargeSkidsCost=Request.Cookies("MyCookie")("AddLargeSkidsCost")
    'Response.write "OrigEmail="&OrigEmail&"<br>"

   ''''''DATA COOKIE AddFuelCharge''''''''''''''''''''
    FuelCharge=Request.Form("FuelCharge")
    If trim(FuelCharge)>"" then
        Response.Cookies ("MyCookie")("FuelCharge")=FuelCharge
    End if
    FuelCharge=Request.Cookies("MyCookie")("FuelCharge")
    'Response.write "OrigEmail="&OrigEmail&"<br>"
   ''''''DATA COOKIE AddRTDescr''''''''''''''''''''
    RtDescr=Request.Form("RtDescr")
    If trim(RtDescr)>"" then
        Response.Cookies ("MyCookie")("RtDescr")=RtDescr
    End if
    RtDescr=Request.Cookies("MyCookie")("RtDescr")
    'Response.write "OrigEmail="&OrigEmail&"<br>"


    
   ''''''DATA COOKIE rtBillCode''''''''''''''''''''
    rtBillCode=Request.Form("rtBillCode")
    If trim(rtBillCode)>"" then
        Response.Cookies ("MyCookie")("rtBillCode")=rtBillCode
    End if
    rtBillCode=Request.Cookies("MyCookie")("rtBillCode")
    'Response.write "OrigEmail="&OrigEmail&"<br>"
   ''''''DATA COOKIE RateCharge''''''''''''''''''''
    RateCharge=Request.Form("RateCharge")
    If trim(RateCharge)>"" then
        Response.Cookies ("MyCookie")("RateCharge")=RateCharge
    End if
    RateCharge=Request.Cookies("MyCookie")("RateCharge")
    'Response.write "OrigEmail="&OrigEmail&"<br>"



   ''''''DATA COOKIE OrigEmail''''''''''''''''''''
      ''''''DATA COOKIE PriorityDescription''''''''''''''''''''
    PriorityDescription=Request.Form("PriorityDescription")
    If trim(PriorityDescription)>"" then
        Response.Cookies ("MyCookie")("PriorityDescription")=PriorityDescription
    End if
    PriorityDescription=Request.Cookies("MyCookie")("PriorityDescription")
    'Response.write "OrigEmail="&OrigEmail&"<br>"
   ''''''DATA COOKIE PriorityCost''''''''''''''''''''
    PriorityCost=Request.Form("PriorityCost")
    If trim(PriorityCost)>"" then
        Response.Cookies ("MyCookie")("PriorityCost")=PriorityCost
    End if
    PriorityCost=Request.Cookies("MyCookie")("PriorityCost")
    'Response.write "OrigEmail="&OrigEmail&"<br>"



   ''''''DATA COOKIE OrigEmail''''''''''''''''''''




    OrigEmail=Request.Form("OrigEmail")
    If trim(OrigEmail)>"" then
        Response.Cookies ("MyCookie")("OrigEmail")=OrigEmail
    End if
    OrigEmail=Request.Cookies("MyCookie")("OrigEmail")
    'Response.write "OrigEmail="&OrigEmail&"<br>"
    '''''''''''''''''''''''''
    If trim(OrderStatus)="" then OrderStatus=1 End if
   ''''''DATA COOKIE DestEmail''''''''''''''''''''
    DestEmail=Request.Form("DestEmail")
    If trim(DestEmail)>"" then
        Response.Cookies ("MyCookie")("DestEmail")=DestEmail
    End if
    DestEmail=Request.Cookies("MyCookie")("DestEmail")
    'Response.write "DestEmail="&DestEmail&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE OtherEmail''''''''''''''''''''
    OtherEmail=Request.Form("OtherEmail")
    If trim(OtherEmail)>"" then
        Response.Cookies ("MyCookie")("OtherEmail")=OtherEmail
    End if
    OtherEmail=Request.Cookies("MyCookie")("OtherEmail")
    'Response.write "OtherEmail="&OtherEmail&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE AdditionalEmail''''''''''''''''''''
    AdditionalEmail=Request.Form("AdditionalEmail")
    If trim(AdditionalEmail)>"" then
        Response.Cookies ("MyCookie")("AdditionalEmail")=AdditionalEmail
    End if
    AdditionalEmail=Request.Cookies("MyCookie")("AdditionalEmail")
    'Response.write "AdditionalEmail="&AdditionalEmail&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE XSquare''''''''''''''''''''
    XSquare=Request.Form("XSquare")
    If trim(XSquare)>"" then
        Response.Cookies ("MyCookie")("XSquare")=XSquare
    End if
    XSquare=Request.Cookies("MyCookie")("XSquare")
    'Response.write "XSquare="&XSquare&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE MarkTemp''''''''''''''''''''
    MarkTemp=Request.Form("MarkTemp")
    If trim(MarkTemp)>"" then
        Response.Cookies ("MyCookie")("MarkTemp")=MarkTemp
    End if
    MarkTemp=Request.Cookies("MyCookie")("MarkTemp")
    'Response.write "MarkTemp="&MarkTemp&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE Username''''''''''''''''''''
    Username=Request.Form("Username")
    If trim(Username)>"" then
        Response.Cookies ("MyCookie")("Username")=Username
    End if
    Username=Request.Cookies("MyCookie")("Username")
    'Response.write "Username="&Username&"<br>"
    '''''''''''''''''''''''''
       ''''''DATA COOKIE Password''''''''''''''''''''
    Password=Request.Form("Password")
    If trim(Password)>"" then
        Response.Cookies ("MyCookie")("Password")=Password
    End if
    Password=Request.Cookies("MyCookie")("Password")
    'Response.write "Password="&Password&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE ShipmentType''''''''''''''''''''
    ShipmentType=Request.Form("ShipmentType")
    If trim(ShipmentType)>"" then
        Response.Cookies ("MyCookie")("ShipmentType")=ShipmentType
    End if
    ShipmentType=Request.Cookies("MyCookie")("ShipmentType")
    'Response.write "ShipmentType="&ShipmentType&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE TimesThrough''''''''''''''''''''
    TimesThrough=Request.Form("TimesThrough")
    If trim(TimesThrough)>"" then
        Response.Cookies ("MyCookie")("TimesThrough")=TimesThrough
    End if
    TimesThrough=Request.Cookies("MyCookie")("TimesThrough")
    'Response.write "TimesThrough="&TimesThrough&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE Internal''''''''''''''''''''
    Internal=Request.Form("Internal")
    If trim(Internal)>"" then
        Response.Cookies ("MyCookie")("Internal")=Internal
    End if
    Internal=Request.Cookies("MyCookie")("Internal")
    'Response.write "Internal="&Internal&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE PreExistingRequestor''''''''''''''''''''
    PreExistingRequestor=Request.Form("PreExistingRequestor")
    If trim(PreExistingRequestor)>"" then
        Response.Cookies ("MyCookie")("PreExistingRequestor")=PreExistingRequestor
    End if
    PreExistingRequestor=Request.Cookies("MyCookie")("PreExistingRequestor")
    'Response.write "PreExistingRequestor="&PreExistingRequestor&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE PreExistingOrigination''''''''''''''''''''
    PreExistingOrigination=Request.Form("PreExistingOrigination")
    If trim(PreExistingOrigination)>"" then
        Response.Cookies ("MyCookie")("PreExistingOrigination")=PreExistingOrigination
    End if
    PreExistingOrigination=Request.Cookies("MyCookie")("PreExistingOrigination")
    'Response.write "PreExistingOrigination="&PreExistingOrigination&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE PreExistingDestination''''''''''''''''''''
    PreExistingDestination=Request.Form("PreExistingDestination")
    If trim(PreExistingDestination)>"" then
        Response.Cookies ("MyCookie")("PreExistingDestination")=PreExistingDestination
    End if
    PreExistingDestination=Request.Cookies("MyCookie")("PreExistingDestination")
    'Response.write "PreExistingDestination="&PreExistingDestination&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE PageStatus''''''''''''''''''''
    PageStatus=Request.Form("PageStatus")
    'If trim(PageStatus)>"" then
    '    Response.Cookies ("MyCookie")("PageStatus")=PageStatus
    'End if
    'PageStatus=Request.Cookies("MyCookie")("PageStatus")
    'Response.write "PageStatus="&PageStatus&"<br>"
    '''''''''''''''''''''''''
        TableWidth="460"
        If (trim(PreExistingOrigination)="" OR trim(PreExistingDestination)="") then
            PageStatus="x"
        End if
   ''''''DATA COOKIE RequestorName''''''''''''''''''''
    RequestorName=Request.Form("RequestorName")
    If trim(RequestorName)>"" then
        Response.Cookies ("MyCookie")("RequestorName")=RequestorName
    End if
    RequestorName=Request.Cookies("MyCookie")("RequestorName")
    'Response.write "RequestorName="&RequestorName&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE RequestorPhoneNumber''''''''''''''''''''
    RequestorPhoneNumber=Request.Form("RequestorPhoneNumber")
    If trim(RequestorPhoneNumber)>"" then
        Response.Cookies ("MyCookie")("RequestorPhoneNumber")=RequestorPhoneNumber
    End if
    RequestorPhoneNumber=Request.Cookies("MyCookie")("RequestorPhoneNumber")
    'Response.write "RequestorPhoneNumber="&RequestorPhoneNumber&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE RequestorEmailAddress''''''''''''''''''''
    RequestorEmailAddress=Request.Form("RequestorEmailAddress")
    If trim(RequestorEmailAddress)>"" then
        Response.Cookies ("MyCookie")("RequestorEmailAddress")=RequestorEmailAddress
    End if
    RequestorEmailAddress=Request.Cookies("MyCookie")("RequestorEmailAddress")
    'Response.write "RequestorEmailAddress="&RequestorEmailAddress&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE Pieces''''''''''''''''''''
    Pieces=Request.Form("Pieces")
    If trim(Pieces)>"" then
        Response.Cookies ("MyCookie")("Pieces")=Pieces
    End if
    Pieces=Request.Cookies("MyCookie")("Pieces")
    'Response.write "Pieces="&Pieces&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE NumberOfPallets''''''''''''''''''''
    NumberOfPallets=Request.Form("NumberOfPallets")
    If trim(NumberOfPallets)>"" then
        Response.Cookies ("MyCookie")("NumberOfPallets")=NumberOfPallets
    End if
    NumberOfPallets=Request.Cookies("MyCookie")("NumberOfPallets")
    'Response.write "NumberOfPallets="&NumberOfPallets&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE DimWeight''''''''''''''''''''
    DimWeight=Request.Form("DimWeight")
    If trim(DimWeight)>"" then
        Response.Cookies ("MyCookie")("DimWeight")=DimWeight
    End if
    DimWeight=Request.Cookies("MyCookie")("DimWeight")
    'Response.write "DimWeight="&DimWeight&"<br>"
    'Response.write "OrderStatus="&OrderStatus&"<br>"
    'Response.write "ShipmentType="&ShipmentType&"<br>"
    If trim(dimweight)>"" then
        If trim(Orderstatus)="3" and trim(ShipmentType)="Light Package" then
            If ((Int(DimWeight)/int(Pieces)>25) or Int(DimWeight)>100) then
                'REsponse.write "GOT HERE! Line 331<BR>"
                OrderStatus="2"
                ShipmentType="Heavy Freight"
                Response.Cookies ("MyCookie")("ShipmentType")="Heavy Freight"
                ErrorMessage="Your shipment is too heavy to be sent as a light package.  We have converted your order to heavy freight.<BR>Please complete section above to continue."
            End if
        End if

    End if

    '''''''''''''''''''''''''
   ''''''DATA COOKIE DimLength''''''''''''''''''''
    DimLength=Request.Form("DimLength")
    If trim(DimLength)>"" then
        Response.Cookies ("MyCookie")("DimLength")=DimLength
    End if
    DimLength=Request.Cookies("MyCookie")("DimLength")
    'Response.write "DimLength="&DimLength&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE DimWidth''''''''''''''''''''
    DimWidth=Request.Form("DimWidth")
    If trim(DimWidth)>"" then
        Response.Cookies ("MyCookie")("DimWidth")=DimWidth
    End if
    DimWidth=Request.Cookies("MyCookie")("DimWidth")
    'Response.write "DimWidth="&DimWidth&"<br>"
    '''''''''''''''''''''''''
       ''''''DATA COOKIE DimHeight''''''''''''''''''''
    DimHeight=Request.Form("DimHeight")
    If trim(DimHeight)>"" then
        Response.Cookies ("MyCookie")("DimHeight")=DimHeight
    End if
    DimHeight=Request.Cookies("MyCookie")("DimHeight")
    'Response.write "DimHeight="&DimHeight&"<br>"
	'Response.write "ShipmentType="&ShipmentType&"<br>"
    '''''''''''''''''''''''''
If Trim(DimHeight)>"" and Trim(DimLength)>"" and trim(DimWidth)>"" then
    If trim(Orderstatus)="3" and trim(ShipmentType)="Light Package" then
        DimCubed=DimHeight*DimWidth*DimLength
        If DimCubed>=4096 and ((Int(DimWeight)/int(Pieces)>25) or Int(DimWeight)>100)then
            'REsponse.write "GOT HERE! Line 331<BR>"
            OrderStatus="2"
            ShipmentType="Heavy Freight"
            Response.Cookies ("MyCookie")("ShipmentType")="Heavy Freight"
            ErrorMessage="Your shipment is too large to be sent as a light package.  We have converted your order to heavy freight.<BR>Please complete section above to continue."
        End if
    End if
    If trim(Orderstatus)="3" and trim(ShipmentType)="Heavy Freight" then
        If Int(DimHeight)>87 then
            'REsponse.write "GOT HERE! Line 331<BR>"
            OrderStatus="2"
            'ShipmentType="Heavy Freight"
            'Response.Cookies ("MyCookie")("ShipmentType")="Heavy Freight"
            ErrorMessage="Your shipment is too large for us to handle.  We are unable to accept shipments greater than 87 inches tall."
        End if
    End if
End if
   ''''''DATA COOKIE IsPalletized''''''''''''''''''''
    IsPalletized=Request.Form("IsPalletized")
    If trim(IsPalletized)>"" then
        Response.Cookies ("MyCookie")("IsPalletized")=IsPalletized
    End if
    IsPalletized=Request.Cookies("MyCookie")("IsPalletized")
    'Response.write "IsPalletized="&IsPalletized&"<br>"
    '''''''''''''''''''''''''

If trim(Orderstatus)="3" and trim(ShipmentType)="Heavy Freight" then
    If trim(IsPalletized)="" then
        'REsponse.write "GOT HERE! Line 331<BR>"
        OrderStatus="2"
        ShipmentType="Heavy Freight"
        Response.Cookies ("MyCookie")("ShipmentType")="Heavy Freight"
        ErrorMessage="You must indicate whether your shipment is palletized or not."
    End if
End if


   ''''''DATA COOKIE Isstacked''''''''''''''''''''
    Isstacked=Request.Form("Isstacked")
    If trim(Isstacked)>"" then
        Response.Cookies ("MyCookie")("Isstacked")=Isstacked
    End if
    Isstacked=Request.Cookies("MyCookie")("Isstacked")
    'Response.write "Isstacked="&Isstacked&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE DimValue''''''''''''''''''''
    DimValue=Request.Form("DimValue")
    If trim(DimValue)>"" then
        Response.Cookies ("MyCookie")("DimValue")=DimValue
    End if
    DimValue=Request.Cookies("MyCookie")("DimValue")
    'Response.write "DimValue="&DimValue&"<br>"
    '''''''''''''''''''''''''
       ''''''DATA COOKIE IsHazmat''''''''''''''''''''
    IsHazmat=Request.Form("IsHazmat")
    If trim(IsHazmat)>"" then
        Response.Cookies ("MyCookie")("IsHazmat")=IsHazmat
    End if
    IsHazmat=Request.Cookies("MyCookie")("IsHazmat")
    'Response.write "IsHazmat="&IsHazmat&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE OriginationID''''''''''''''''''''
    OriginationID=Request.Form("OriginationID")
    If trim(OriginationID)>"" then
        Response.Cookies ("MyCookie")("OriginationID")=OriginationID
    End if
    OriginationID=Request.Cookies("MyCookie")("OriginationID")
    'Response.write "OriginationID="&OriginationID&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE OriginationCompany''''''''''''''''''''
    OriginationCompany=Request.Form("OriginationCompany")
    If trim(OriginationCompany)>"" then
        Response.Cookies ("MyCookie")("OriginationCompany")=OriginationCompany
    End if
    OriginationCompany=Request.Cookies("MyCookie")("OriginationCompany")
    'Response.write "OriginationCompany="&OriginationCompany&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE OriginationBuilding''''''''''''''''''''
    OriginationBuilding=Request.Form("OriginationBuilding")
    If trim(OriginationBuilding)>"" then
        Response.Cookies ("MyCookie")("OriginationBuilding")=OriginationBuilding
    End if
    OriginationBuilding=Request.Cookies("MyCookie")("OriginationBuilding")
    'Response.write "OriginationBuilding="&OriginationBuilding&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE OriginationAddress''''''''''''''''''''
    OriginationAddress=Request.Form("OriginationAddress")
    If trim(OriginationAddress)>"" then
        Response.Cookies ("MyCookie")("OriginationAddress")=OriginationAddress
    End if
    OriginationAddress=Request.Cookies("MyCookie")("OriginationAddress")
    'Response.write "OriginationAddress="&OriginationAddress&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE OriginationSuite''''''''''''''''''''
    OriginationSuite=Request.Form("OriginationSuite")
    If trim(OriginationSuite)>"" then
        Response.Cookies ("MyCookie")("OriginationSuite")=OriginationSuite
    End if
    OriginationSuite=Request.Cookies("MyCookie")("OriginationSuite")
    'Response.write "OriginationSuite="&OriginationSuite&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE OriginationCity''''''''''''''''''''
    OriginationCity=Request.Form("OriginationCity")
    If trim(OriginationCity)>"" then
        Response.Cookies ("MyCookie")("OriginationCity")=OriginationCity
    End if
    OriginationCity=Request.Cookies("MyCookie")("OriginationCity")
    'Response.write "OriginationCity="&OriginationCity&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE OriginationState''''''''''''''''''''
    OriginationState=Request.Form("OriginationState")
    If trim(OriginationState)>"" then
        Response.Cookies ("MyCookie")("OriginationState")=OriginationState
    End if
    OriginationState=Request.Cookies("MyCookie")("OriginationState")
    'Response.write "OriginationState="&OriginationState&"<br>"
    '''''''''''''''''''''''''
    ''''''DATA COOKIE OriginationZipCode''''''''''''''''''''
    OriginationZipCode=Request.Form("OriginationZipCode")
    OriginationIsCourier=Request.Form("OriginationIsCourier")
    If trim(OriginationIsCourier)>"" then
        Response.Cookies ("MyCookie")("OriginationIsCourier")=OriginationIsCourier
    End if
    If trim(OriginationZipCode)>"" then
        Response.Cookies ("MyCookie")("OriginationZipCode")=OriginationZipCode
    End if
    OriginationZipCode=Request.Cookies("MyCookie")("OriginationZipCode")
    OriginationIsCourier=Request.Cookies("MyCookie")("OriginationIsCourier")
    If OriginationIsCourier<>"y" and trim(OriginationIsCourier)>"" and trim(ShipmentType)="Light Package"  then
    'Response.write "OriginationZipCode="&Left(trim(OriginationZipCode),5)&"<BR>"
    'Response.write "ShipmentType="&ShipmentType&"<BR>"
        'REsponse.write "GOT HERE! Line 331<BR>"
        OrderStatus="2"
        ShipmentType="Heavy Freight"
        Response.Cookies ("MyCookie")("ShipmentType")="Heavy Freight"
        OriginationZipCode=""
        OriginationIsCourier=""
        Response.Cookies ("MyCookie")("OriginationZipCode")=OriginationZipCode
        Response.Cookies ("MyCookie")("OriginationIsCourier")=OriginationIsCourier
        ErrorMessage="We do not provide light package service from your origination.  We have converted your order to heavy freight.<BR>Please complete section above to continue."
    End if
    'Response.write "OriginationZipCode="&OriginationZipCode&"<br>"
    '''''''''''''''''''''''''
    ''''''DATA COOKIE OriginationAliasCode''''''''''''''''''''
    OriginationAliasCode=Request.Form("OriginationAliasCode")
    If trim(OriginationAliasCode)>"" then
        Response.Cookies ("MyCookie")("OriginationAliasCode")=OriginationAliasCode
    End if
    OriginationAliasCode=Request.Cookies("MyCookie")("OriginationAliasCode")
    'Response.write "OriginationAliasCode="&OriginationAliasCode&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE OriginationContactName''''''''''''''''''''
    OriginationContactName=Request.Form("OriginationContactName")
    If trim(OriginationContactName)>"" then
        Response.Cookies ("MyCookie")("OriginationContactName")=OriginationContactName
    End if
    OriginationContactName=Request.Cookies("MyCookie")("OriginationContactName")
    'Response.write "OriginationContactName="&OriginationContactName&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE OriginationPhoneNumber''''''''''''''''''''
    OriginationPhoneNumber=Request.Form("OriginationPhoneNumber")
    If trim(OriginationPhoneNumber)>"" then
        Response.Cookies ("MyCookie")("OriginationPhoneNumber")=OriginationPhoneNumber
    End if
    OriginationPhoneNumber=Request.Cookies("MyCookie")("OriginationPhoneNumber")
    'Response.write "OriginationPhoneNumber="&OriginationPhoneNumber&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE OriginationEmail''''''''''''''''''''
    OriginationEmail=Request.Form("OriginationEmail")
    If trim(OriginationEmail)>"" then
        Response.Cookies ("MyCookie")("OriginationEmail")=OriginationEmail
    End if
    OriginationEmail=Request.Cookies("MyCookie")("OriginationEmail")
    'Response.write "OriginationEmail="&OriginationEmail&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE DestinationID''''''''''''''''''''
    DestinationID=Request.Form("DestinationID")
    If trim(DestinationID)>"" then
        Response.Cookies ("MyCookie")("DestinationID")=DestinationID
    End if
    DestinationID=Request.Cookies("MyCookie")("DestinationID")
    'Response.write "DestinationID="&DestinationID&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE DestinationCompany''''''''''''''''''''
    DestinationCompany=Request.Form("DestinationCompany")
    If trim(DestinationCompany)>"" then
        Response.Cookies ("MyCookie")("DestinationCompany")=DestinationCompany
    End if
    DestinationCompany=Request.Cookies("MyCookie")("DestinationCompany")
    'Response.write "DestinationCompany="&DestinationCompany&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE DestinationBuilding''''''''''''''''''''
    DestinationBuilding=Request.Form("DestinationBuilding")
    If trim(DestinationBuilding)>"" then
        Response.Cookies ("MyCookie")("DestinationBuilding")=DestinationBuilding
    End if
    DestinationBuilding=Request.Cookies("MyCookie")("DestinationBuilding")
    'Response.write "DestinationBuilding="&DestinationBuilding&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE DestinationAddress''''''''''''''''''''
    DestinationAddress=Request.Form("DestinationAddress")
    If trim(DestinationAddress)>"" then
        Response.Cookies ("MyCookie")("DestinationAddress")=DestinationAddress
    End if
    DestinationAddress=Request.Cookies("MyCookie")("DestinationAddress")
    'Response.write "DestinationAddress="&DestinationAddress&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE DestinationSuite''''''''''''''''''''
    DestinationSuite=Request.Form("DestinationSuite")
    If trim(DestinationSuite)>"" then
        Response.Cookies ("MyCookie")("DestinationSuite")=DestinationSuite
    End if
    DestinationSuite=Request.Cookies("MyCookie")("DestinationSuite")
    'Response.write "DestinationSuite="&DestinationSuite&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE DestinationCity''''''''''''''''''''
    DestinationCity=Request.Form("DestinationCity")
    If trim(DestinationCity)>"" then
        Response.Cookies ("MyCookie")("DestinationCity")=DestinationCity
    End if
    DestinationCity=Request.Cookies("MyCookie")("DestinationCity")
    'Response.write "DestinationCity="&DestinationCity&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE DestinationState''''''''''''''''''''
    DestinationState=Request.Form("DestinationState")
    If trim(DestinationState)>"" then
        Response.Cookies ("MyCookie")("DestinationState")=DestinationState
    End if
    DestinationState=Request.Cookies("MyCookie")("DestinationState")
    'Response.write "DestinationState="&DestinationState&"<br>"
    '''''''''''''''''''''''''
    ''''''DATA COOKIE DestinationZipCode''''''''''''''''''''
    DestinationZipCode=Request.Form("DestinationZipCode")
    If trim(DestinationZipCode)>"" then
        Response.Cookies ("MyCookie")("DestinationZipCode")=DestinationZipCode
    End if
    DestinationZipCode=Request.Cookies("MyCookie")("DestinationZipCode")
    DestinationIsCourier=Request.Form("DestinationIsCourier")
    If trim(DestinationIsCourier)>"" then
        Response.Cookies ("MyCookie")("DestinationIsCourier")=DestinationIsCourier
    End if
    DestinationZipCode=Request.Cookies("MyCookie")("DestinationZipCode")
    If DestinationIsCourier<>"y" and trim(DestinationIsCourier)>"" and trim(ShipmentType)="Light Package" then
        'REsponse.write "GOT HERE! Line 331<BR>"
        OrderStatus="2"
        ShipmentType="Heavy Freight"
        Response.Cookies ("MyCookie")("ShipmentType")="Heavy Freight"
        ErrorMessage="We do not provide light package service to your destination.  We have converted your order to heavy freight.<BR>Please complete section above to continue."
        DestinationZipCode=""
        Response.Cookies ("MyCookie")("DestinationZipCode")=""
        DestinationIsCourier=""
        Response.Cookies ("MyCookie")("DestinationIsCourier")=""
    End if
    'Response.write "DestinationZipCode="&DestinationZipCode&"<br>"
    '''''''''''''''''''''''''
    ''''''DATA COOKIE DestinationAliasCode''''''''''''''''''''
    DestinationAliasCode=Request.Form("DestinationAliasCode")
    If trim(DestinationAliasCode)>"" then
        Response.Cookies ("MyCookie")("DestinationAliasCode")=DestinationAliasCode
    End if
    DestinationAliasCode=Request.Cookies("MyCookie")("DestinationAliasCode")
    'Response.write "DestinationAliasCode="&DestinationAliasCode&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE DestinationContactName''''''''''''''''''''
    DestinationContactName=Request.Form("DestinationContactName")
    If trim(DestinationContactName)>"" then
        Response.Cookies ("MyCookie")("DestinationContactName")=DestinationContactName
    End if
    DestinationContactName=Request.Cookies("MyCookie")("DestinationContactName")
    'Response.write "DestinationContactName="&DestinationContactName&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE DestinationPhoneNumber''''''''''''''''''''
    DestinationPhoneNumber=Request.Form("DestinationPhoneNumber")
    If trim(DestinationPhoneNumber)>"" then
        Response.Cookies ("MyCookie")("DestinationPhoneNumber")=DestinationPhoneNumber
    End if
    DestinationPhoneNumber=Request.Cookies("MyCookie")("DestinationPhoneNumber")
    'Response.write "DestinationPhoneNumber="&DestinationPhoneNumber&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE DestinationEmail''''''''''''''''''''
    DestinationEmail=Request.Form("DestinationEmail")
    If trim(DestinationEmail)>"" then
        Response.Cookies ("MyCookie")("DestinationEmail")=DestinationEmail
    End if
    DestinationEmail=Request.Cookies("MyCookie")("DestinationEmail")
    'Response.write "DestinationEmail="&DestinationEmail&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE POorNWA''''''''''''''''''''
    POorNWA=Request.Form("POorNWA")
    If trim(POorNWA)>"" then
        Response.Cookies ("MyCookie")("POorNWA")=POorNWA
    End if
    POorNWA=Request.Cookies("MyCookie")("POorNWA")
    'Response.write "POorNWA="&POorNWA&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE GenericNumber''''''''''''''''''''

    GenericNumber=Request.Form("GenericNumber")
    bGenericNumber=Request.Form("bGenericNumber")
    If trim(GenericNumber)>"" then
        Response.Cookies ("MyCookie")("GenericNumber")=GenericNumber
        ELSE
        If trim(bGenericNumber)>"" then
            Response.Cookies ("MyCookie")("GenericNumber")=bGenericNumber
        End if
    End if
    GenericNumber=Request.Cookies("MyCookie")("GenericNumber")
    'Response.write "XXXGenericNumber="&GenericNumber&"XXX<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE GenericNumber''''''''''''''''''''
    bGenericNumber=Request.Form("bGenericNumber")
    If trim(bGenericNumber)>"" then
        Response.Cookies ("MyCookie")("bGenericNumber")=bGenericNumber
    End if
    bGenericNumber=Request.Cookies("MyCookie")("bGenericNumber")
    'Response.write "GenericNumber="&GenericNumber&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE OriginationNotifications''''''''''''''''''''
    OriginationNotifications=Request.Form("OriginationNotifications")
    If trim(OriginationNotifications)>"" then
        Response.Cookies ("MyCookie")("OriginationNotifications")=OriginationNotifications
    End if
    OriginationNotifications=Request.Cookies("MyCookie")("OriginationNotifications")
    'Response.write "OriginationNotifications="&OriginationNotifications&"<br>"
    '''''''''''''''''''''''''
       ''''''DATA COOKIE RequestorNotifications''''''''''''''''''''
    RequestorNotifications=Request.Form("RequestorNotifications")
    If trim(RequestorNotifications)>"" then
        Response.Cookies ("MyCookie")("RequestorNotifications")=RequestorNotifications
    End if
    RequestorNotifications=Request.Cookies("MyCookie")("RequestorNotifications")
    'Response.write "RequestorNotifications="&RequestorNotifications&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE OriginationNotifications''''''''''''''''''''
    OriginationWaybill=Request.Form("OriginationWaybill")
    If trim(OriginationWaybill)>"" then
        Response.Cookies ("MyCookie")("OriginationWaybill")=OriginationWaybill
    End if
    OriginationWaybill=Request.Cookies("MyCookie")("OriginationWaybill")
    'Response.write "OriginationNotifications="&OriginationNotifications&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE DestinationNotifications''''''''''''''''''''
    DestinationNotifications=Request.Form("DestinationNotifications")
    If trim(DestinationNotifications)>"" then
        Response.Cookies ("MyCookie")("DestinationNotifications")=DestinationNotifications
    End if
    DestinationNotifications=Request.Cookies("MyCookie")("DestinationNotifications")
    'Response.write "DestinationNotifications="&DestinationNotifications&"<br>"
    '''''''''''''''''''''''''
   If trim(OriginationNotifications)="y" then
        SendTo=OriginationEmail
   End if

   If trim(OriginationWaybill)="y" then
        SendWaybillTo=OriginationEmail
   End if

   If trim(DestinationNotifications)="y" then
        SendTo=DestinationEmail
   End if
   If trim(OriginationNotifications)="y" and trim(DestinationNotifications)="y" then
        SendTo=OriginationEmail&";"&DestinationEmail
   End if
   If trim(RequestorNotifications)="y" then
        If trim(SendTo)>"" then
            SendTo=RequestorEmailAddress&";"&SendTo
        else
            SendTo=RequestorEmailAddress
        End if
   End if
   If trim(GenericNumber)="" and trim(bGenericNumber)>"" then 
        GenericNumber=bGenericNumber 
   End if
    Select Case POorNWA
        Case "P/O #"
            PONumber=GenericNumber
        Case "Cost Center #"
            CostCenterNumber=GenericNumber
    End Select
   ''''''DATA COOKIE Comments''''''''''''''''''''
    Comments=Request.Form("Comments")
    If trim(Comments)>"" then
        Response.Cookies ("MyCookie")("Comments")=Comments
    End if
    Comments=Request.Cookies("MyCookie")("Comments")
    'Response.write "Comments="&Comments&"<br>"
    '''''''''''''''''''''''''
    RequestorName=Replace(RequestorName, """", "`")
    RequestorName=Replace(RequestorName, "'", "`")
    RequestorPhoneNumber=Replace(RequestorPhoneNumber, """", "")
    RequestorPhoneNumber=Replace(RequestorPhoneNumber, "'", "")
    RequestorEmailAddress=Replace(RequestorEmailAddress, """", "")
    RequestorEmailAddress=Replace(RequestorEmailAddress, "'", "")
    Pieces=Replace(Pieces, """", "")
    Pieces=Replace(Pieces, "'", "")
   ''''''DATA COOKIE rf_box''''''''''''''''''''
    rf_box=Request.Form("rf_box")
    If trim(rf_box)>"" then
        Response.Cookies ("MyCookie")("rf_box")=rf_box
    End if
    rf_box=Request.Cookies("MyCookie")("rf_box")
    'Response.write "rf_box="&rf_box&"<br>"
    '''''''''''''''''''''''''  
    ButtonSubmit=Request.form("ButtonSubmit")
    'Response.write "buttonsubmit="&buttonsubmit&"<br>" 
    'Response.write "var1b="&var1b&"<br>" 
    If trim(lcase(rf_box))="envelopes" and ((trim(lcase(buttonsubmit))<>"next >>>") and (trim(lcase(var1b))<>"<<<back"))then
        DontShowDim="y"
        'OrderStatus=2
    End if
    'rf_box=Request.Form("rf_box")
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
    OriginationIsCourier=Replace(OriginationIsCourier, """", "")
    OriginationIsCourier=Replace(OriginationIsCourier, "'", "")
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
    DestinationIsCourier=Replace(DestinationIsCourier, """", "")
    DestinationIsCourier=Replace(DestinationIsCourier, "'", "")
    DestinationContactName=Replace(DestinationContactName, """", "`")
    DestinationContactName=Replace(DestinationContactName, "'", "`")
    DestinationPhoneNumber=Replace(DestinationPhoneNumber, """", "")
    DestinationPhoneNumber=Replace(DestinationPhoneNumber, "'", "")
    DestinationEmail=Replace(DestinationEmail, """", "")
    DestinationEmail=Replace(DestinationEmail, "'", "")


    Comments=Replace(Comments, """", "`")
    Comments=Replace(Comments, "'", "`")
   ''''''DATA COOKIE Refrigerate''''''''''''''''''''
    Refrigerate=Request.Form("Refrigerate")
    If trim(Refrigerate)>"" then
        Response.Cookies ("MyCookie")("Refrigerate")=Refrigerate
    End if
    Refrigerate=Request.Cookies("MyCookie")("Refrigerate")
    'Response.write "Refrigerate="&Refrigerate&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE Priority''''''''''''''''''''
    Priority=Request.Form("Priority")
    If trim(Priority)>"" then
        Response.Cookies ("MyCookie")("Priority")=Priority
    End if
    Priority=Request.Cookies("MyCookie")("Priority")
    'Response.write "XXXPriority="&Priority&"<br>"
    '''''''''''''''''''''''''

   ''''''DATA COOKIE DeliveryDateTime''''''''''''''''''''
    DeliveryDateTime=Request.Form("DeliveryDateTime")
    If trim(DeliveryDateTime)>"" then
        Response.Cookies ("MyCookie")("DeliveryDateTime")=DeliveryDateTime
    End if
    DeliveryDateTime=Request.Cookies("MyCookie")("DeliveryDateTime")
    'Response.write "DeliveryDateTime="&DeliveryDateTime&"<br>"
    '''''''''''''''''''''''''
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
        RequestorEmailAddress="FleetX@LogisticorpGroup.com"
        'RequestorEmailAddress="mark.maggiore@logisticorp.us"
        OriginationContactName="Dispatch"
        OriginationPhoneNumber="817-458-4594"
        OriginationEmail="FleetX@LogisticorpGroup.com"
        'OriginationEmail="mark.maggiore@logisticorp.us"
        DestinationContactName="Dispatch"
        DestinationPhoneNumber="817-458-4594"
        DestinationEmail="FleetX@LogisticorpGroup.com"
        'DestinationEmail="mark.maggiore@logisticorp.us"
    End if
''''''''''''''SAVES ORDER INFO WHEN GOING TO OTHER PAGES''''''''''''''''''
   ''''''DATA COOKIE Submit''''''''''''''''''''
    Submit=Request.Form("Submit")
    If trim(Submit)>"" then
        Response.Cookies ("MyCookie")("Submit")=Submit
    End if
    Submit=Request.Cookies("MyCookie")("Submit")
    'Response.write "Submit="&Submit&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE ButtonSubmit''''''''''''''''''''
    ButtonSubmit=Request.Form("ButtonSubmit")
    'If trim(ButtonSubmit)>"" then
    '    Response.Cookies ("MyCookie")("ButtonSubmit")=ButtonSubmit
    'End if
    'ButtonSubmit=Request.Cookies("MyCookie")("ButtonSubmit")
    'Response.write "ButtonSubmit="&ButtonSubmit&"<br>"
    '''''''''''''''''''''''''
   ''''''DATA COOKIE VarA''''''''''''''''''''
    VarA=Request.Form("VarA")
    If trim(VarA)>"" then
        Response.Cookies ("MyCookie")("VarA")=VarA
    End if
    VarA=Request.Cookies("MyCookie")("VarA")
    'Response.write "VarA="&VarA&"<br>"
    '''''''''''''''''''''''''
If lcase(trim(ButtonSubmit))="edit" OR lcase(trim(ButtonSubmit))="add/edit locations in your address book"  then
    If lcase(trim(ButtonSubmit))="edit" then
        Response.redirect("EditUserInfo.asp")
        else
        Response.redirect("FleetXAddressBook.asp?var1="&OrderStatus)
    End IF
End if
If lcase(trim(ButtonSubmit))="submit order" then
    pagestatus="submit"
End if

    Select Case ShipmentType
        Case "Light Package"
            LegacyBillToID="92"
        Case "Heavy Freight"
            LegacyBillToID="93"

    End select
    If IsStandingOrder="y" then
        LegacyBillToID="93"
    End if
        

''''''''''''''END
'Response.write "OrderStatus="&OrderStatus&"<BR>"
    If Trim(OrderStatus)<>"1" then
        'REsponse.write "Line 770 Got here!<BR>"

       'If not isdate(DeliveryDateTime) then DeliveryDateTime=now() end if
        If not isdate(PickUPDateTime) then PickUPDateTime=now() end if
        If cdate(PickUpDateTime)<now() then PickUpDateTime=Now() end if
       
       
   'Response.write "CurrentHour="&CurrentHour&"<BR>"
   'Response.write "DayOfWeek="&Weekday(now())&"<BR>"
   'Response.write "DayOfWeek="&WeekdayName(Weekday(now()))&"<BR>"
        If trim(XSquare)="y" then
            If (CurrentHour=6 and currentMinute<50) or CurrentHour<6 or CurrentHour>11 or WeekdayName(Weekday(now()))="Saturday" or WeekdayName(Weekday(now()))="Sunday" then
                   ErrorMessage="Due to dock hours, X Square Probe Card orders can only be placed Monday-Friday between 7:00 AM and 12:00 PM"
            End if
        End if 

        'If trim(OtherEmail)="y" AND (inStr(AdditionalEmail,"@") = 0 OR inStr(AdditionalEmail,".") = 0) THEN
        '   ErrorMessage="The 'Additional' email that provided '"& AdditionalEmail &"' is not a valid email address."
        'End if
        'If trim(OtherEmail)="y" and (trim(AdditionalEmail)="" or trim(AdditionalEmail)="Additional") then
        '    ErrorMessage="You checked the 'Additional' email notification box, you must enter an additional email address."
        'end if
        'If isdate(DeliveryDateTime)  and cdate(CurrentDateTime)>=cdate(DeliveryDateTime) then
        '    ErrorMessage="The delivery date/time cannot be before the current date/time"
        'End if
        'If isdate(DeliveryDateTime) and isdate(PickUPDateTime) and cdate(PickUPDateTime)>=cdate(DeliveryDateTime) then
        '    ErrorMessage="The delivery date/time cannot be before the ready date/time"
        'End if
        'Response.write "DeliveryDateTime="&DeliveryDateTime&"<BR>"
        'If NOT isdate(DeliveryDateTime) then
        '    ErrorMessage="You must provide a valid destination date/time"
        'End if
        
        'Response.write "rtBillCode="&rtBillCode&"<BR>"

        If trim(DestinationCompany)=trim(OriginationCompany) AND trim(DestinationAddress)=Trim(OriginationAddress) and trim(OriginationContactName)=trim(DestinationContactName) and OrderSTatus="5" then
        ErrorMessage="Your Origination and Destination cannot be the same"
        OrderSTatus="4"
        End if
        'If trim(DestinationPhoneNumber)="" then
        'ErrorMessage="You must provide the Destination's Phone Number"
        'End if
        'If trim(DestinationContactName)="" then
        '    ErrorMessage="You must provide the Destination's Contact Name"
        'End if
        'If trim(DestinationZipCode)="" then
        '    ErrorMessage="You must provide the Destination's Zip Code"
        'End if
        'If trim(DestinationCity)="" then
        '    ErrorMessage="You must provide the Destination's City"
        'End if
        'If trim(DestinationAddress)="" then
        '    ErrorMessage="You must provide the Destination's Address"
        'End if
        'If trim(DestinationCompany)="" then
        '    ErrorMessage="You must provide the Destination's Company"
        'End if
        'If trim(OriginationEmail)="" then
        '    ErrorMessage="You must provide the Origination's Email"
        'End if
        'Response.write "PickUpDateTime="&PickUpDateTime&"<BR>"
        'Response.write "CurrentDateTime="&CurrentDateTime&"<BR>"
        'Response.write "OrderSTatus="&OrderSTatus&"<BR>"
        If isdate(PickUpDateTime)  and cdate(CurrentDateTime)>cdate(PickUpDateTime) and OrderSTatus="4" then
            ErrorMessage="The ready time cannot be before the current date/time"
            OrderSTatus="3"
        End if
        If NOT isdate(PickUpDateTime)  and OrderSTatus="4" then
            ErrorMessage="You must provide a valid ready date/time"
            OrderStatus="3"
        End if
        If trim(OriginationID)=""  and OrderSTatus="4" then
            ErrorMessage="You must select an origination"
            OrderStatus="3"
        End if

        If trim(DestinationID)=""  and OrderSTatus="5" then
            ErrorMessage="You must select a destination"
            OrderStatus="4"
        End if
        'If trim(OriginationPhoneNumber)="" then
        '    ErrorMessage="You must provide the Origination's Phone Number"
        'End if
        'If trim(OriginationContactName)="" then
        '    ErrorMessage="You must provide the Origination's Contact Name"
        'End if
        'If trim(OriginationZipCode)="" then
        '    ErrorMessage="You must provide the Origination's Zip Code"
        'End if
        'If trim(OriginationCity)="" then
       '     ErrorMessage="You must provide the Origination's City"
       ' End if
       ' If trim(OriginationAddress)="" then
       '     ErrorMessage="You must provide the Origination's Address"
        'End if
        'If trim(OriginationCompany)="" then
        '    ErrorMessage="You must provide the Origination's Company"
        'End if
        If trim(rf_box)<>"Envelopes" then
            'If trim(DimHeight)=""  and OrderSTatus="3" and lcase(ShipmentType)="light package" then
			If trim(DimHeight)=""  and OrderSTatus="3"  then
                ErrorMessage="You must provide the Commodity's Height"
                OrderSTatus="2"
            End if
            'If trim(DimWidth)=""  and OrderSTatus="3" and lcase(ShipmentType)="light package" then
			If trim(DimWidth)=""  and OrderSTatus="3" then
                ErrorMessage="You must provide the Commodity's Width"
                OrderSTatus="2"
            End if
            'If trim(DimLength)=""  and OrderSTatus="3" and lcase(ShipmentType)="light package" then
			If trim(DimLength)=""  and OrderSTatus="3" then
                ErrorMessage="You must provide the Commodity's Length"
                OrderSTatus="2"
            End if
        End if

        If trim(DimWeight)=""  and OrderSTatus="3" and trim(lcase(rf_box))<>"envelopes" then
            ErrorMessage="You must provide the Commodity's Weight"
            OrderSTatus="2"
        End if
        'If trim(NumberOfPallets)="" and isPalletized="y" then
        '    ErrorMessage="You must provide the Number of Pallets"
        'End if
        If trim(Pieces)=""  and OrderSTatus="3" then
            ErrorMessage="You must provide the Number of Pieces"
            OrderSTatus="2"
        End if
        'If trim(Comments)="" then
        '    ErrorMessage="You have not provided any Special Instructions for this delivery.<br>If there are no special instructions, please type in N/A"
        'End if
        If trim(CostCenterNumber)="" AND trim(PONumber)="" and OrderSTatus="6" then
            ErrorMessage="You must provide the Cost Center Number or PO Number"
            OrderStatus="5"
            buttonSubmit=""
        End if
        'If trim(PONumber)="" then
        '    ErrorMessage="You must provide the P/O Number"
        'End if
        'If trim(RequestorEmailAddress)="" then
        '    ErrorMessage="You must provide the Requestor Email Address"
        'End if
        'If trim(RequestorPhoneNumber)="" then
        '    ErrorMessage="You must provide the Requestor Phone Number"
        'End if
        'If trim(RequestorName)="" then
        '    ErrorMessage="You must provide the Requestor Name"
        'End if
        'Response.write "ErrorMessage="&ErrorMessage&"<BR>"

        VehicleType="Van"
        IF IsPalletized="y" then
            VehicleType="Bobtail"
        End if

        If trim(DimWeight)>"" then

            If ((Int(DimWeight)/int(Pieces)>25) or Int(DimWeight)>100) then
                VehicleType="Bobtail"
            End if

            If DimWeight>10000 then
                VehicleType="Tractor"
            End if
        End if
        If trim(Pieces)>"" then
            If pieces>12 and trim(rf_box)="Skids" and trim(isstacked)="n" then
                VehicleType="Tractor"
            End if
            If pieces>20 and trim(rf_box)="Skids" and trim(ispalletized)="y" and trim(isstacked)="y" then
                VehicleType="Tractor"
            End if
            If pieces>23 and trim(rf_box)="Skids" and trim(isstacked)="n" then
                ErrorMessage="That shipment exceeds our vehicle capability.  You will need to break it up into orders with no more than 23 unstacked skids."
            End if
            If pieces>46 and trim(rf_box)="Skids" and trim(ispalletized)="y" and trim(isstacked)="y" then
                ErrorMessage="That shipment exceeds our vehicle capability.  You will need to break it up into orders with no more than 46 stacked skids."
            End if
			''''''''''''''Large''''''''''''''''''''''''
            If pieces>12 and trim(rf_box)="Large Skids" and trim(isstacked)="n" then
                VehicleType="Tractor"
            End if
            If pieces>20 and trim(rf_box)="Large Skids" and trim(ispalletized)="y" and trim(isstacked)="y" then
                VehicleType="Tractor"
            End if
            If pieces>23 and trim(rf_box)="Large Skids" and trim(isstacked)="n" then
                ErrorMessage="That shipment exceeds our vehicle capability.  You will need to break it up into orders with no more than 23 unstacked large skids."
            End if
            If pieces>46 and trim(rf_box)="Large Skids" and trim(ispalletized)="y" and trim(isstacked)="y" then
                ErrorMessage="That shipment exceeds our vehicle capability.  You will need to break it up into orders with no more than 46 stacked large skids."
            End if
			'''''''''''''''''''''''''''''''''''''''''''''''''''
        End if
        If OrderSTatus="6" and trim(PONumber)="" then
                'REsponse.write "Line 772 Got here!<BR>"
		        Set oConn = Server.CreateObject("ADODB.Connection")
		        oConn.ConnectionTimeout = 100
		        oConn.Provider = "MSDASQL"
		        oConn.Open DATABASE
			        l_cSQL = "Select * FROM TICostCenters WHERE costcenterstatus='c' and CostCenterNumber='"& CostCenterNumber &"'"
			        'Response.write "CostCenter="&CostCenter&"<BR>"
                    SET oRs = oConn.Execute(l_cSql)
					        if oRs.EOF then
                                OrderStatus="5"
                                ErrorMessage="You did not provide a valid Cost Center number."
                                buttonSubmit=""
                                'REsponse.write "Line 781 - Got here!<BR>"
                                else
                            End if								
		        Set oConn=Nothing


        End if

        If trim(ErrorMessage)="" and buttonsubmit="Submit Order" then





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
            'Response.write "VehicleType="&VehicleType&"<BR>"
            'Response.write "DimWeight="&DimWeight&"<BR>"
            'Response.write "pieces="&pieces&"<BR>"
            'Response.write "rf_box="&rf_box&"<BR>"
            'Response.write "ispalletized="&ispalletized&"<BR>"
            'Response.write "isstacked="&isstacked&"<BR>"
            '''''''''Special code so Irene Foreman will recieve ALL CLAB notices''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Response.write "DestinationSuite="&DestinationSuite&"<BR>"
            'Response.write "DestinationBuilding="&DestinationBuilding&"<BR>"
            Vara1=(InStr(DestinationSuite,"clab"))
            Vara2=(InStr(DestinationBuilding,"clab"))
            'Response.write "Vara1="&Vara1&"<BR>"
            'Response.write "vara2="&vara2&"<BR>"
            If Vara1>0 or vara2>0 then
                'Response.write "GOT HERE!!!!<BR>"
                    SendTo=SendTo+";i-foreman@ti.com"
            End if
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            If trim(PONumber)>"" then costcenterNumber=trim(PONumber) end if


           
			Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				RSEVENTS2.Open "fcfgthd", DATABASE, 2, 2
				RSEVENTS2.addnew
				RSEVENTS2("fh_ID")=NewJobNum
                RSEVENTS2("fh_Status")="SCD"
                RSEVENTS2("fh_ship_dt")=now()
                RSEVENTS2("fh_ready")=PickUpDateTime
                RSEVENTS2("Fh_Priority")=Priority
                RSEVENTS2("fh_lastchg")=now()
                RSEVENTS2("fh_bt_ID")=LegacyBillToID
                RSEVENTS2("fh_user_id")=Trim(UserID)
                RSEVENTS2("fh_co_id")=Trim(RequestorName)
                RSEVENTS2("fh_RequestorEmail")=Trim(RequestorEmailAddress)
                RSEVENTS2("fh_co_phone")=Trim(RequestorPhoneNumber)
                RSEVENTS2("fh_co_email")=Trim(SendTo)
                RSEVENTS2("fh_co_costcenter")=Trim(costcenterNumber)
                RSEVENTS2("fh_custpo")=Trim(PoNumber)
                RSEVENTS2("fh_statcode")="2"
                RSEVENTS2("fh_user4")=VehicleType
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


            'OriginationContactName=Replace(OriginationContactName, "/"," ")
            'Response.write "xxxOriginationContactName="&OriginationContactName&"XXX<BR>"
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
                RSEVENTS2("fl_sf_building")=Trim(OriginationBuilding)
                RSEVENTS2("fl_sf_addr1")=Trim(OriginationAddress)
                RSEVENTS2("fl_sf_addr2")=Trim(OriginationSuite)
                RSEVENTS2("fl_sf_city")=Trim(OriginationCity)
                RSEVENTS2("fl_sf_state")=Trim(OriginationState)
                RSEVENTS2("fl_sf_country")="US"
                RSEVENTS2("fl_sf_zip")=Trim(OriginationZipCode)
                RSEVENTS2("fl_sf_alias")=Trim(OriginationAliasCode)
                
                RSEVENTS2("fl_st_ID")=Trim(DestinationID)
                RSEVENTS2("fl_st_name")=Trim(DestinationCompany)
                RSEVENTS2("fl_st_clname")=Trim(DestinationContactName)
                RSEVENTS2("fl_st_phone")=Trim(DestinationPhoneNumber)
                RSEVENTS2("fl_st_email")=Trim(DestinationEmail)
                RSEVENTS2("fl_st_Building")=Trim(DestinationBuilding)
                RSEVENTS2("fl_st_addr1")=Trim(DestinationAddress)
                RSEVENTS2("fl_st_addr2")=Trim(DestinationSuite)
                RSEVENTS2("fl_st_city")=Trim(DestinationCity)
                RSEVENTS2("fl_st_state")=Trim(DestinationState)
                RSEVENTS2("fl_st_country")="US"
                RSEVENTS2("fl_st_zip")=Trim(DestinationZipCode)
                RSEVENTS2("fl_st_alias")=Trim(DestinationAliasCode)
                RSEVENTS2("fl_sf_comment")=Trim(Comments)
                '''''DeliveryDateTime=Now()+.10
                Response.write "DeliveryDateTime="&DeliveryDateTime&"<BR>"
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
                RSEVENTS2("MaterialDescription")=Trim(MaterialDescription)
                RSEVENTS2("rf_box")=trim(rf_box)
                RSEVENTS2("NumberOfPieces")=Trim(Pieces)
                RSEVENTS2("IsPalletized")=Trim(IsPalletized)
                RSEVENTS2("IsStacked")=Trim(IsStacked)
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
                If trim(PriorityCost)="" then PriorityCost=0 end if
                If PriorityCost>0 and CAPBOB<>"y" then
				    Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
					    RSEVENTS2.Open "JobCharges", DATABASE, 2, 2
					    RSEVENTS2.addnew
					    RSEVENTS2("fh_id")=newjobnum
					    RSEVENTS2("billtoid")=LegacyBillToID
					    RSEVENTS2("JobChargesDescription")=PriorityDescription
					    RSEVENTS2("JobChargesRate")=PriorityCost
					    RSEVENTS2("JobChargesBillCode")=PriorityDescription
                        RSEVENTS2("JobChargesStatus")="c"
					    RSEVENTS2.update
					    RSEVENTS2.close			
				    set RSEVENTS2 = nothing	
                End if
                If trim(ratecharge)="" then ratecharge=0 end if
                If ratecharge>0 and CAPBOB<>"y" then
				    Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
					    RSEVENTS2.Open "JobCharges", DATABASE, 2, 2
					    RSEVENTS2.addnew
					    RSEVENTS2("fh_id")=newjobnum
					    RSEVENTS2("billtoid")=LegacyBillToID
					    RSEVENTS2("JobChargesDescription")=rtDescr
					    RSEVENTS2("JobChargesRate")=RateCharge
					    RSEVENTS2("JobChargesBillCode")=rtBillCode
                        RSEVENTS2("JobChargesStatus")="c"
					    RSEVENTS2.update
					    RSEVENTS2.close			
				    set RSEVENTS2 = nothing	
                End if
''''''''''''''''CAP BOB''''''''''''''''''''''''
                If CAPBOB="y" then
				    Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
					    RSEVENTS2.Open "JobCharges", DATABASE, 2, 2
					    RSEVENTS2.addnew
					    RSEVENTS2("fh_id")=newjobnum
					    RSEVENTS2("billtoid")=LegacyBillToID
					    RSEVENTS2("JobChargesDescription")="Cap Rate: Bobtail"
					    RSEVENTS2("JobChargesRate")="250.00"
					    RSEVENTS2("JobChargesBillCode")="CAPBOB"
                        RSEVENTS2("JobChargesStatus")="c"
					    RSEVENTS2.update
					    RSEVENTS2.close			
				    set RSEVENTS2 = nothing	
                End if
'''''''''''''''''''''''''''''''''''''''''''''''
                If trim(FuelChargeDollars)="" then FuelChargeDollars=0 end if
                If FuelCharge>0 then
				    Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
					    RSEVENTS2.Open "JobCharges", DATABASE, 2, 2
					    RSEVENTS2.addnew
					    RSEVENTS2("fh_id")=newjobnum
					    RSEVENTS2("billtoid")=LegacyBillToID
					    RSEVENTS2("JobChargesDescription")="Fuel Charge"
					    RSEVENTS2("JobChargesRate")=FuelChargeDollars
					    RSEVENTS2("JobChargesBillCode")="FE Fuel"
                        RSEVENTS2("JobChargesStatus")="c"
					    RSEVENTS2.update
					    RSEVENTS2.close			
				    set RSEVENTS2 = nothing	
                End if
                If trim(AddSkidsCost)="" then AddSkidsCost=0 end if
                If AddSkidsCost>0 and CAPBOB<>"y" then
				    Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
					    RSEVENTS2.Open "JobCharges", DATABASE, 2, 2
					    RSEVENTS2.addnew
					    RSEVENTS2("fh_id")=newjobnum
					    RSEVENTS2("billtoid")=LegacyBillToID
					    RSEVENTS2("JobChargesDescription")=AddSkids&" Additional Skids"
					    RSEVENTS2("JobChargesRate")=AddSkidsCost
					    RSEVENTS2("JobChargesBillCode")="ADDITIONAL SKIDS"
                        RSEVENTS2("JobChargesStatus")="c"
					    RSEVENTS2.update
					    RSEVENTS2.close			
				    set RSEVENTS2 = nothing	
                End if
                If trim(AddLargeSkidsCost)="" then AddLargeSkidsCost=0 end if
                If AddLargeSkidsCost>0 and CAPBOB<>"y" then
				    Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
					    RSEVENTS2.Open "JobCharges", DATABASE, 2, 2
					    RSEVENTS2.addnew
					    RSEVENTS2("fh_id")=newjobnum
					    RSEVENTS2("billtoid")=LegacyBillToID
					    RSEVENTS2("JobChargesDescription")=AddLargeSkids&" Additional Large Skids"
					    RSEVENTS2("JobChargesRate")=AddLargeSkidsCost
					    RSEVENTS2("JobChargesBillCode")="ADDTL LARGE SKIDS"
                        RSEVENTS2("JobChargesStatus")="c"
					    RSEVENTS2.update
					    RSEVENTS2.close			
				    set RSEVENTS2 = nothing	
                End if


'''''''''''''''''WIPE OUT ALL THE COOKIES!!!'''''''''''''''''
Response.Cookies ("MyCookie")("CAPBOB")=""
Response.Cookies ("MyCookie")("AddSkids")=""
Response.Cookies ("MyCookie")("AddSkidsCost")=""
Response.Cookies ("MyCookie")("AddLargeSkids")=""
Response.Cookies ("MyCookie")("AddLargeSkidsCost")=""
Response.Cookies ("MyCookie")("FuelCharge")=""
Response.Cookies ("MyCookie")("RtDescr")=""
Response.Cookies ("MyCookie")("RateCharge")=""
Response.Cookies ("MyCookie")("rtBillCode")=""
Response.Cookies ("MyCookie")("IsStandingOrder")=""
Response.Cookies ("MyCookie")("MaterialDescription")=""
Response.Cookies ("MyCookie")("PriorityDescription")=""
Response.Cookies ("MyCookie")("PriorityDescr")=""
Response.Cookies ("MyCookie")("PriorityCost")=""
Response.Cookies ("MyCookie")("PickUpDate")=""
Response.Cookies ("MyCookie")("PickUpTime")=""
Response.Cookies ("MyCookie")("ShipmentType")=""
Response.Cookies ("MyCookie")("OrigEmail")=""
Response.Cookies ("MyCookie")("DestEmail")=""
Response.Cookies ("MyCookie")("OtherEmail")=""
Response.Cookies ("MyCookie")("AdditionalEmail")=""
Response.Cookies ("MyCookie")("XSquare")=""
Response.Cookies ("MyCookie")("MarkTemp")=""
Response.Cookies ("MyCookie")("Username")=""
Response.Cookies ("MyCookie")("Password")=""
Response.Cookies ("MyCookie")("ShipmentType")=""
Response.Cookies ("MyCookie")("TimesThrough")=""
Response.Cookies ("MyCookie")("PreExistingRequestor")=""
Response.Cookies ("MyCookie")("PreExistingOrigination")=""
Response.Cookies ("MyCookie")("PreExistingDestination")=""
Response.Cookies ("MyCookie")("RequestorName")=""
Response.Cookies ("MyCookie")("RequestorPhoneNumber")=""
Response.Cookies ("MyCookie")("RequestorEmailAddress")=""
Response.Cookies ("MyCookie")("Pieces")=""
Response.Cookies ("MyCookie")("NumberOfPallets")=""
Response.Cookies ("MyCookie")("DimWeight")=""
Response.Cookies ("MyCookie")("DimLength")=""
Response.Cookies ("MyCookie")("DimWidth")=""
Response.Cookies ("MyCookie")("DimHeight")=""
Response.Cookies ("MyCookie")("IsPalletized")=""
Response.Cookies ("MyCookie")("IsStacked")=""
Response.Cookies ("MyCookie")("DimValue")=""
Response.Cookies ("MyCookie")("IsHazmat")=""
Response.Cookies ("MyCookie")("OriginationID")=""
Response.Cookies ("MyCookie")("OriginationCompany")=""
Response.Cookies ("MyCookie")("OriginationBuilding")=""
Response.Cookies ("MyCookie")("OriginationAddress")=""
Response.Cookies ("MyCookie")("OriginationSuite")=""
Response.Cookies ("MyCookie")("OriginationCity")=""
Response.Cookies ("MyCookie")("OriginationState")=""
Response.Cookies ("MyCookie")("OriginationZipCode")=""
Response.Cookies ("MyCookie")("OriginationIsCourier")=""
Response.Cookies ("MyCookie")("OriginationAliasCode")=""
Response.Cookies ("MyCookie")("OriginationContactName")=""
Response.Cookies ("MyCookie")("OriginationPhoneNumber")=""
Response.Cookies ("MyCookie")("OriginationEmail")=""
Response.Cookies ("MyCookie")("DestinationID")=""
Response.Cookies ("MyCookie")("DestinationCompany")=""
Response.Cookies ("MyCookie")("DestinationBuilding")=""
Response.Cookies ("MyCookie")("DestinationAddress")=""
Response.Cookies ("MyCookie")("DestinationSuite")=""
Response.Cookies ("MyCookie")("DestinationCity")=""
Response.Cookies ("MyCookie")("DestinationState")=""
Response.Cookies ("MyCookie")("DestinationZipCode")=""
Response.Cookies ("MyCookie")("DestinationIsCourier")=""
Response.Cookies ("MyCookie")("DestinationAliasCode")=""
Response.Cookies ("MyCookie")("DestinationContactName")=""
Response.Cookies ("MyCookie")("DestinationPhoneNumber")=""
Response.Cookies ("MyCookie")("DestinationEmail")=""
Response.Cookies ("MyCookie")("POorNWA")=""
Response.Cookies ("MyCookie")("GenericNumber")=""
Response.Cookies ("MyCookie")("bGenericNumber")=""
Response.Cookies ("MyCookie")("PONumber")=""
Response.Cookies ("MyCookie")("OriginationNotifications")=""
Response.Cookies ("MyCookie")("RequestorNotifications")=""
Response.Cookies ("MyCookie")("OriginationWaybill")=""
Response.Cookies ("MyCookie")("DestinationNotifications")=""
Response.Cookies ("MyCookie")("Comments")=""
Response.Cookies ("MyCookie")("rf_box")=""
Response.Cookies ("MyCookie")("Refrigerate")=""
Response.Cookies ("MyCookie")("Priority")=""
Response.Cookies ("MyCookie")("DeliveryDateTime")=""
Response.Cookies ("MyCookie")("Submit")=""
Response.Cookies ("MyCookie")("VarA")=""

'''''''''''''''''END - WIPE OUT ALL THE COOKIES!!!'''''''''''''''''

		        
                'Response.Write "YAY!!!! You're DONE!!!!<BR>"
		
         
            'Response.write "newjobnum="&newjobnum&"<BR>"
  ''''''''''''''''LETTER to Requestor'''''''''''''''''''
   				    Body = "Your shipment request (#"& newjobnum &") has been successfully placed online:<br><br>"  
                    Body = Body & "Your shipment will be picked up sometime between "& pickupdatetime &" and "& DeliveryDateTime &" <br><br>" 
                    Body = Body & "To print out this waybill, <a href='"&WhichSite&"/orderentry/FleetXFreightOrderConfirmation.asp?x=123&y=1&pid=view&jid="& newjobnum &"'>click here</a><br><br>"  
                    Body = Body & "***Should you need to cancel this order, please call 214-882-0620***<BR><BR>"

                    Body = Body & "REQUESTOR INFORMATION:<BR>"
                    Body = Body & "Name: "&  RequestorName &"<br>"  
                    Body = Body & "Phone Number: "&  RequestorPhoneNumber &"<br>"  
                    Body = Body & "Email Address: "&  RequestorEmailAddress &"<br>"   
                    Body = Body & "PO Number: "&  PONumber &"<br>"  
                    Body = Body & "Cost Center Number: "&  CostCenterNumber &"<br><br>" 
                    Body = Body & "COMMODITY INFORMATION:<BR>" 
                    Body = Body & "Material Description: "&  MaterialDescription &"<br>" 
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
                    Body = Body & "Suite/Cube/Dock: "&  OriginationSuite &"<br>"  
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
                    Body = Body & "Suite/Cube/Dock: "&  DestinationSuite &"<br>"  
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
                   Body = Body & "***Should you need to contact us regarding this order, please  either email FleetX@LogisticorpGroup.com or call 214-882-0620***<BR><BR>"

				    Body = Body & "Thank you,<br><br>"  
				    Body = Body & "FleetX Services<br>"  
				    Body = Body &  "FleetX@LogisticorpGroup.com<br>"  
				    Body = Body & "214-882-0620<br><br>"
				    'Recipient=FirstName&" "&LastName
			        SentToEmail4=RequestorEmailAddress
				    'Email="KWETI.Mailbox@am.kwe.com"
				    'Email="mark@maggiore.net"
				    'Set objMail = CreateObject("CDONTS.Newmail")
				    'objMail.From = "FleetX@LogisticorpGroup.com"
                    '''''If trim(OriginationNotifications)>"" then
                        '''''SentToEmail=trim(SentToEmail)&";"&trim(OriginationEmail)
                    '''''End if
				    varTo = SentToEmail4
				    varcc = "mark.maggiore@logistiCorp.us"
				    'objMail.cc = "4692269939@tmomail.net;pcurrin@ti.com"
				    varSubject = "Thank you for your FleetX shipment request"
				    'objMail.MailFormat = cdoMailFormatMIME
				   ' objMail.BodyFormat = cdoBodyFormatHTML
				    'objMail.Body = Body
                    If trim(lcase(SentToEmail4))<>"FleetX@LogisticorpGroup.com" then
				        'objMail.Send
                    'End if
				    'Set objMail = Nothing
    ''''''''''''
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
            		
				    Body = "There has been a new FleetX shipment request (#"& newjobnum &") placed online:<br><br>"   

                    Body = Body & "REQUESTOR INFORMATION:<BR>"
                    Body = Body & "Name: "&  RequestorName &"<br>"  
                    Body = Body & "Phone Number: "&  RequestorPhoneNumber &"<br>"  
                    Body = Body & "Email Address: "&  RequestorEmailAddress &"<br>"   
                    Body = Body & "PO Number: "&  PONumber &"<br>"  
                    Body = Body & "Cost Center Number: "&  CostCenterNumber &"<br><br>" 
                    Body = Body & "COMMODITY INFORMATION:<BR>"
                    Body = Body & "Material Description: "&  MaterialDescription &"<br>" 
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
                    Body = Body & "Suite/Cube/Dock: "&  OriginationSuite &"<br>"  
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
                    Body = Body & "Suite/Cube/Dock: "&  DestinationSuite &"<br>"  
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
                    'Body = Body & "<a href='http://www.logisticorp.us/intranet/dedicatedfleets/orderentry/FleetXFreightOrderConfirmation.asp?bid=86&pid=disp&jid="& newjobnum &"'>To Route or Cancel this request, click here</a><br><br>" 
				    Body = Body & "Thank you,<br><br>"  
				    Body = Body & "FleetX Services<br>"  
				    Body = Body &  "FleetX@LogisticorpGroup.com<br>"  
				    Body = Body & "214-882-0620<br><br>"
				    'Recipient=FirstName&" "&LastName
			        SentToEmail2="mark.maggiore@logisticorp.us"
                    'SentToEmail="mark.maggiore@logisticorp.us"
				    'Email="KWETI.Mailbox@am.kwe.com"
				    'Email="mark@maggiore.net"
				    'Set objMail = CreateObject("CDONTS.Newmail")
				    'objMail.From = "FleetX@LogisticorpGroup.com"
				    varTo = SentToEmail2
				    'objMail.cc = RequestorEmailAddress
				    'objMail.cc = "4692269939@tmomail.net;pcurrin@ti.com"
				    'objMail.Subject = "New Fleet Express Shipment Request"
				    If trim(Priority)="Time Critical" then
                        varSubject = "* TIME CRITICAL * New FleetX Shipment Request"
                        'objMail.Importance = 2 'High
                        else
                        varSubject = "New FleetX Shipment Request"
                    End if
				    'objMail.MailFormat = cdoMailFormatMIME
				    'objMail.BodyFormat = cdoBodyFormatHTML
				    'objMail.Body = Body
				    'objMail.Send
				    'Set objMail = Nothing
    ''''''''''''''''''''''''''''''''''''''''''
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
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    If trim(DimLength)="" then DimLength=0 end if
                    If trim(DimWidth)="" then DimWidth=0 end if
                    If trim(DimHeight)="" then DimHeight=0 end if
                    If (int(DimLength)>42 or int(DimWidth)>48 or int(DimHeight)>48) AND xyz="removethisfornow" then
				        Body = "There has been an Large FleetX shipment request (#"& newjobnum &") placed online:<br><br>"   

                        Body = Body & "REQUESTOR INFORMATION:<BR>"
                        Body = Body & "Name: "&  RequestorName &"<br>"  
                        Body = Body & "Phone Number: "&  RequestorPhoneNumber &"<br>"  
                        Body = Body & "Email Address: "&  RequestorEmailAddress &"<br>"   
                        Body = Body & "PO Number: "&  PONumber &"<br>"  
                        Body = Body & "Cost Center Number: "&  CostCenterNumber &"<br><br>" 
                        Body = Body & "COMMODITY INFORMATION:<BR>" 
                        Body = Body & "Material Description: "&  MaterialDescription &"<br>"
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
                        Body = Body & "Suite/Cube/Dock: "&  OriginationSuite &"<br>"   
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
                        Body = Body & "Suite/Cube/Dock: "&  DestinationSuite &"<br>"  
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
				        Body = Body & "214-882-0620<br><br>"
				        'Recipient=FirstName&" "&LastName
			            SentToEmail3="mark.maggiore@logisticorp.us"
                        'SentToEmail="mark.maggiore@logisticorp.us"
				        'Email="KWETI.Mailbox@am.kwe.com"
				        'Email="mark@maggiore.net"
				        'Set objMail = CreateObject("CDONTS.Newmail")
				        'objMail.From = "FleetX@LogisticorpGroup.com"
				        varTo = SentToEmail3
				        'objMail.cc = RequestorEmailAddress
				        'objMail.cc = "4692269939@tmomail.net;pcurrin@ti.com"
				        'objMail.Subject = "Large Fleet Express Shipment Request"
				        If trim(Priority)="Time Critical" then
                            varSubject = "* TIME CRITICAL * Large FleetX Shipment Request"
                            'objMail.Importance = 2 'High
                            else
                            varSubject = "Large FleetX Shipment Request"
                        End if
				        'objMail.MailFormat = cdoMailFormatMIME
				        'objMail.BodyFormat = cdoBodyFormatHTML
				        'objMail.Body = Body
				        'objMail.Send
				        'Set objMail = Nothing
    ''''''''''''''''''''''''''''''''''''''''''''''
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
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


                  Response.Redirect("FleetXFreightOrderConfirmation.asp?x=1&y=1&bid=86&pid=view&jid="& newjobnum &"&Internal="&Internal&"&XSquare="&XSquare)	
                  'Response.Redirect("TempRedirect.asp?x=1&y=1&bid=86&pid=view&jid="& newjobnum &"&Internal="&Internal&"&XSquare="&XSquare)	
		    'End if	
        End if
    End if

    ''''''''END ERROR HANDLING''''''


'Response.write "HELLO????<BR>"

%>
<link rel="stylesheet" href="../themes/classic.css" id="theme_base">
<link rel="stylesheet" href="../themes/classic.date.css" id="theme_date">
<link rel="stylesheet" href="../themes/classic.time.css" id="theme_time">
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
<!--form action="NewUser.asp" method="post" name="FindUser"-->
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
                <%If Errormessage>"" then%>   
                <%'If Errormessage="dontshowthisanymore" then%>  
                <tr><td align="center">        
                <table cellpadding="2" cellspacing="2" border="1" bordercolor="red">
                    <tr>
                        <td class="errormessage">
                            <%
                            Response.write " * * * ERROR:  "&ErrorMessage& " * * * "
                             %>
                        </td>
                    </tr>
                </table>
                </td></tr>  
                <%End if %>
    <tr><td>


   <%



    'Response.write "Line 1610 - LegacyBillToID="&LegacyBillToID&"<BR>"



   Select Case OrderStatus
   Case "1"
    %> 
   
    <table border="0" cellpadding="0" cellspacing="0" align="center">
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>
                  <%If trim(XSquare)="y" then %>
                <tr><td>&nbsp;</td></tr>
                <tr><td colspan="3"  class="FleetExpressTextBlackBold" align="center"><font color="blue">***XSquare Probe Card orders may only be placed between 7:00 AM and 12:00 PM Monday-Friday.***</font></td></tr>
               <tr><td>&nbsp;</td></tr>
                <%end if %>        
        <tr>
            <td align="center" colspan="2">
            <table border="0" cellpadding="0" cellspacing="0" width="100%">
                <tr> <td valign="top">
                <table border="0" cellpadding="3" cellspacing="0" width="100%">
                
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap>Requestor Name</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <%
                        ''''''''CHECKS TO SEE IF USER IS ADMIN
                        Set Recordset1 = Server.CreateObject("ADODB.Recordset")
                        'Response.write "Database="&Database&"<br>"
                        Recordset1.ActiveConnection = Database
                        SQL="SELECT * FROM PreExistingRequestor WHERE (RequestorID='"&Request.cookies("FleetXCookie")("UserID")&"') AND (RequestorStatus='c')"
                        Recordset1.Source = SQL
                        
                        'response.write "SQL="& SQL &"<BR>"
                        
                        Recordset1.CursorType = 0
                        Recordset1.CursorLocation = 2
                        Recordset1.LockType = 1
                        Recordset1.Open()
                        Recordset1_numRows = 0
                        'response.write "<font color='red'>SQL1="&Recordset1.Source&"<BR></font>"
	                        if NOT Recordset1.EOF then
                                Supervisor=Recordset1("RequestorType")
                                If Supervisor="A" then 
                                    Internal="y" 
                                end if
	                        End if
                    Recordset1.Close()
                    Set Recordset1 = Nothing 
                    
                    'response.write "Supervisor="& Supervisor &"<BR>"
                    'response.write "Internal="& Internal &"<BR>"
                     
                    If Internal="y" then  
                    %>
                    <form method="post">
                    <td align="left">
								<%
                                'Response.write "PreExistingRequestor="&PreExistingRequestor&"<BR>"
                                'response.write "Database="&database&"<BR>"
                                 %>
                                <select name="PreExistingRequestor" ID="Select2"  onChange="form.submit()">
								<option value="" <%if trim(PreExistingRequestor)="" then response.Write " selected" end if%>>Select an existing customer</option>
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
											<option value="<%=TempRequestorID%>" <%If trim(TempRequestorID)=trim(PreExistingRequestor) then response.write "selected" end if %>><%=TempRequestorName%></option>
											<%
										oRs.movenext
										LOOP
									Set oConn=Nothing									
									%>
								</select> 
                                </form>
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
                                'Response.write "l_cSQL="&l_cSQL&"<BR>"
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
                <td>
                    <form method="post" name="OrderForm1" action="FreightOrder.asp?Internal=<%=Internal%>">
                        <td align="right"><input id="gobutton" type="submit" name="buttonsubmit" value="Edit" />
                    </form>               
                </td>
                </tr>
                <%
                'Response.write "RequestorName="&RequestorName&"*<BR>"
                'Response.write "GenericNumber="&GenericNumber&"*<BR>"
                'If trim(RequestorName)="Jake Weber" and trim(GenericNumber)="" then
                '    'Response.write "got here!!!<BR>"
                '    GenericNumber="2883"
                'End if               
                 %>
                <!--tr>
                    <td class="FleetExpressTextBlackBold" align="left">Requestor Name</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left"><input type="hidden" name="RequestorName" value="<%=RequestorName%>" /><%=RequestorName%></td>
                </tr-->
                <form method="post" name="OrderForm1" action="FreightOrder.asp?Internal=<%=Internal%>">
                <input type="hidden" name="RequestorName" value="<%=RequestorName%>" />
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">Phone Number</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left"><input type="hidden" name="RequestorPhoneNumber" value="<%=RequestorPhoneNumber%>"/><%=RequestorPhoneNumber%></td>
                </tr>
                <%
                If trim(UCASE(RequestorEmailAddress))="SCOTT@PRIORITYLABS.COM" then
                    RequestorEmailAddress="scott@prioritylabs.com;jdrummond@prioritylabs.com"
                End if 

                 %>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">Email Address</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left"><input type="hidden" name="RequestorEmailAddress" value="<%=RequestorEmailAddress%>"/><%=RequestorEmailAddress%></td>
                </tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap>Receive Notifications</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left" class="FleetExpressTextBlack"><input type="radio" name="RequestorNotifications" value="y" <%If trim(RequestorNotifications)="" or trim(RequestorNotifications)="y" then Response.write "Checked" end if %>/>Yes&nbsp;&nbsp;&nbsp;&nbsp;<input type="radio" name="RequestorNotifications" value="n" <%If trim(RequestorNotifications)="n" then Response.write "Checked" end if %>/>No</td>
                </tr>
                <%
                'response.write "RequestorCompany="&RequestorCompany&"<BR>"
                'Response.write "sBT_ID="&sBT_ID&"<BR>"
                %>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" valign="top" nowrap>Special Instructions</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left"><textarea name="comments" rows="2" cols="30"><%=Comments%></textarea></td>
                </tr>
                <tr><td><img src="../images/pixel.gif" width="1" height="35" /></td></tr>
                <input type="hidden" name="OrderStatus" value="1b" />
                <tr><td colspan="4">&nbsp;</td><td align="right"><input id="gobutton" type="submit" name="buttonsubmit" value="Next >>>" /></td></tr>
                </form>
                </table>
           </td>
     </tr>
<%
Case "1b"
%>
    <table border="0" cellpadding="0" cellspacing="0" align="center">
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>
       
        <tr>
            <td align="center" colspan="2">
            <table border="0" cellpadding="0" cellspacing="0" width="100%">
                <tr> <td valign="top">

   <form method="post" action="FreightOrder.asp">
        <table align="center" cellpadding="0" cellspacing="0">
            <tr>
                <td colspan="3">
                Which most closely describes your shipment?
                </td>
            </tr>
            <tr><td>&nbsp;</td></tr>
             <tr><td>&nbsp;</td></tr>
            <tr>
                <td>
                <img src="../images/lightfreight.gif" height="205" width="300" />
                </td>
                <td width="80">&nbsp;</td>
                <td>
                <img src="../images/heavyfreight.jpg" height="192" width="263" />
                </td>
            </tr>
             <tr><td>&nbsp;</td></tr>
            <tr>
                <td align="center">
                <input id="gobutton" name="ShipmentType" type="submit" value="Light Package" />
                </td>
                <td>&nbsp;</td>
                <td align="center">
                <input id="gobutton" name="ShipmentType" type="submit" value="Heavy Freight" />
                </td>
            </tr>
            <input type="hidden" name="OrderStatus" value="2" />
        </table>
    </form>
    <tr><td>&nbsp;</td></tr>
    <tr><td>&nbsp;</td></tr>
    <tr><td>&nbsp;</td></tr>
    <form method="post" action="FreightOrder.asp?Internal=<%=Internal %>">
    <input type="hidden" value="1" name="OrderStatus" />
    <tr><td><input type="submit" id="gobutton" value="<<<Back" /></td></tr>
    </form>

<%
Case "2"
%>

    <table border="0" cellpadding="0" cellspacing="0" align="center">
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>
       
        <tr>
            <td align="center" colspan="2">
            <table border="0" cellpadding="0" cellspacing="0" width="100%">
                <tr> <td valign="top">

                <%
                'Response.write "ShipmentType="&ShipmentType&"<BR>"
                 %>

                <form method="post" name="OrderForm1" action="FreightOrder.asp?Internal=<%=Internal%>">
                    <table border="0" bordercolor="<%=BorderColor%>" cellpadding="3" cellspacing="0" width="<%=tablewidth%>" align="center"> 
                     <tr>
                        <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">Material Description</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left"  class="FleetExpressTextBlack" nowrap="nowrap">
                            <input type="text" name="MaterialDescription" value="<%=MaterialDescription %>" maxlength="250" /> (Ex. Control/Part Number, Documents, machine parts)
                        </td>
                    </tr>                   
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left" Nowrap>Number of Items</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
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
                                    
                            <%
                            else 
                           ' REsponse.write "ShipmentType="&ShipmentType&"<BR>"
                            %>
                            <select name="rf_box" onchange='if(this.value == "Envelopes") { this.form.submit(); }'>

                                <option value="Boxes"<%If trim(rf_box)="Boxes" then Response.write "selected" end if%>>Boxes</option>
                                 <%if lcase(shipmenttype)="light package" then%>
                                <option value="Envelopes"<%If trim(rf_box)="Envelopes" then Response.write "selected" end if%>>Envelopes</option>
                                <option value="Pieces"<%If trim(rf_box)="Pieces" then Response.write "selected" end if%>>Pieces</option>
                                <%end if %>                               
                                <%if lcase(shipmenttype)="heavy freight" then%>
                                <option value="Crates"<%If trim(rf_box)="Crates" then Response.write "selected" end if%>>Crates</option>
                                <option value="Skids"<%If trim(rf_box)="Skids" then Response.write "selected" end if%>>Skids (42x48)</option>
								<option value="Large Skids"<%If trim(rf_box)="Large Skids" then Response.write "selected" end if%>>Large Skids (83x51x62)</option>
                                <%end if %>
                            </select>
                            <%end if %>
                        </td>
                    </tr>
                    <%If trim(lcase(ShipmentType))="heavy freight" then %>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Palletization</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td nowrap class="FleetExpressTextBlackBold">
                            <select name="IsPalletized">
                                <option value=""<%If trim(IsPalletized)="" then Response.write "selected" end if%>>Select One</option>
                                <option value="y"<%If trim(IsPalletized)="y" then Response.write "selected" end if%>>Palletized</option>
                                <option value="n"<%If trim(IsPalletized)="n" then Response.write "selected" end if%>>Not Palletized</option>
                                <option value="Trailer Only Move"<%If trim(IsPalletized)="Trailer Only Move" then Response.write "selected" end if%>>Trailer Only Move</option>
                            </select>
                        </td>
                    </tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">If Palletized</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td nowrap class="FleetExpressTextBlackBold">
                            <select name="IsStacked">
                                <option value="n"<%If trim(IsStacked)="n" or trim(IsStacked)="" then Response.write "selected" end if%>>Unstacked</option>
                                <option value="y"<%If trim(IsStacked)="y" then Response.write "selected" end if%>>Double Stacked</option>
                            </select>
                        </td>
                    </tr>
                    <%End if %>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Weight</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack" nowrap="nowrap"><input type="text" name="DimWeight" value="<%=DimWeight%>" size="6" maxlength="5" /> Pounds
                        <%If trim(lcase(ShipmentType))="light package" then %>
                        &nbsp;&nbsp;(Max: 25 pounds each item)
                        <%End if %>
                        </td>
                    </tr>
                    <%'If trim(lcase(ShipmentType))="light package" and DontShowDim<>"y" then %>
					<%If DontShowDim<>"y" then %>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">Dimensions (Of largest item)</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td class="FleetExpressTextBlack" align="left" nowrap>
                            L:&nbsp;&nbsp;<input type="text" name="DimLength" value="<%=DimLength%>" size="5"  maxlength="4"/> 
                            W:&nbsp;&nbsp;<input type="text" name="DimWidth" value="<%=DimWidth%>" size="5" maxlength="4" /> 
                            H:&nbsp;&nbsp;<input type="text" name="DimHeight" value="<%=DimHeight%>" size="5"  maxlength="4"/> 
                            &nbsp;&nbsp;Inches
                         <%If trim(lcase(ShipmentType))="light package" then%>
                        &nbsp;&nbsp;(Max: 16 X 16 X 16)
						<%else%>
						&nbsp;&nbsp;(Max Height: 87 inches)
						<%end if%>
                        
                        </td>
                    </tr>
                    <%End if %>
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

                    <input type="hidden" name="IsHazmat" value="n" />
                    <input type="hidden" name="Refrigerate" value="n" />
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Service Level</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left"  class="FleetExpressTextBlackBold">
                            <select name="Priority">
<%

									Set oConn = Server.CreateObject("ADODB.Connection")
									oConn.ConnectionTimeout = 100
									oConn.Provider = "MSDASQL"
									oConn.Open DATABASE
									    l_cSQL = "Select * FROM Priorities WHERE priorityorigination<>'TI-Sherman' and  prioritystatus='c' and priority_BT_ID='"& LegacyBillToID & "' ORDER BY PriorityMinutes desc"
										SET oRs = oConn.Execute(l_cSql)
												Do While not oRs.EOF
                                                PriorityID=oRs("PriorityID")
                                                varPriority=oRs("Priority")
												PriorityDescription=oRs("PriorityDescription")
                                                PriorityCost=oRs("PriorityCost")
												PriorityMinutes=oRs("PriorityMinutes")
					                            If PriorityCost>0 then
											        %>
                                            		<option value="<%=PriorityID%>" <%if trim(PriorityID)=trim(Priority) then response.Write " selected" end if%>><%=PriorityDescription%>&nbsp;&nbsp;(Base Charge: $<%=PriorityCost %>)</option>
											        <%
                                                    Else
                                                    %>
                                                    											<option value="<%=PriorityID%>" <%if trim(PriorityID)=trim(Priority) then response.Write " selected" end if%>><%=PriorityDescription%></option>
                                                    <%
                                                End if

										oRs.movenext
										LOOP
									Set oConn=Nothing	
 %>
                                </select>
                   
                        </td>
                    </tr>
                    <%
                    'Response.write "priority="&priority&"<BR>" 
                    'Response.write "l_csql="&l_csql&"<BR>" 
                    %>
                    <tr><td><img src="../images/pixel.gif" width="1" /></td></tr>
                    <tr><td colspan="3">
                    <%'response.write "LegacyBillToID="&LegacyBillToID&"<BR>" 
                    Select case LegacyBillToID

                    Case "92"
                    %>
                    <table border="1" cellpadding="5" cellspacing="0" bordercolor="white" class="FleetXRedSection">
                        <tr>
                            <td colspan="2" align="center">SERVICE LEVEL DESCRIPTION</td>
                        </tr>
                        <tr><td nowrap valign="top">3 HOUR</td><td>SHIPMENT WILL BE PICKED UP, AT OUR DISCRETION, ANYTIME AFTER THE READY TIME,  AND DELIVERED WITHIN 3 HOURS OF THE READY TIME YOU PROVIDED</td></tr>
                        <tr><td colspan="2">***NOTE:  THE FASTEST SERVICE WE PROVIDE IS WITHIN 3 HOURS OF THE READY TIME YOU PROVIDED.  IF YOU NEED A DELIVERY FASTER THAN 3 HOURS, WE WILL NOT BE ABLE TO PROVIDE THAT SERVICE.</td></tr>
                        </table>
                    <%
                    Case "93"
                    %>
                    <table border="1" cellpadding="5" cellspacing="0" bordercolor="white" class="FleetXRedSection">
                        <tr>
                            <td colspan="2" align="center">SERVICE LEVEL DESCRIPTIONS</td>
                        </tr>
                        <tr><td nowrap valign="top">NEXT DAY</td><td>SHIPMENT WILL BE PICKED UP, AT OUR DISCRETION, ANYTIME AFTER THE READY TIME, AND DELIVERED BY MIDNIGHT THE NEXT DAY</td></tr>
                        <tr><td nowrap valign="top">8 HOUR</td><td>SHIPMENT WILL BE PICKED UP, AT OUR DISCRETION, ANYTIME AFTER THE READY TIME,  AND DELIVERED WITHIN 8 HOURS OF THE READY TIME YOU PROVIDED</td></tr>
                        <tr><td nowrap valign="top">5 HOUR</td><td>SHIPMENT WILL BE PICKED UP, AT OUR DISCRETION, ANYTIME AFTER THE READY TIME,  AND DELIVERED WITHIN 5 HOURS OF THE READY TIME YOU PROVIDED</td></tr>
                        <tr><td nowrap valign="top">3 HOUR</td><td>SHIPMENT WILL BE PICKED UP, AT OUR DISCRETION, ANYTIME AFTER THE READY TIME,  AND DELIVERED WITHIN 3 HOURS OF THE READY TIME YOU PROVIDED</td></tr>
                        <tr><td colspan="2">***NOTE:  THE FASTEST SERVICE WE PROVIDE IS WITHIN 3 HOURS OF THE READY TIME YOU PROVIDED.  IF YOU NEED A DELIVERY FASTER THAN 3 HOURS, WE WILL NOT BE ABLE TO PROVIDE THAT SERVICE.</td></tr>
                        </table>
                    <%
                    End Select
                    %>
                    </td></tr>
                    <tr><td><img src="../images/pixel.gif" width="1" height="35" /></td></tr>
                <input type="hidden" name="OrderStatus" value="3" />
                <tr><td colspan="2">

    <input type="submit" id="gobutton" name="var1b" value="<<<Back" />
    </form>               
                </td><td align="right"><input id="gobutton" type="submit" name="buttonsubmit" value="Next >>>" /></td></tr>
                <tr><td>&nbsp;</td></tr>
                <%if lcase(shipmenttype)="light package" then%>
                    <tr><td align="right" colspan="4"><img src="../images/lightpackage2.jpg" width="250" height="247" /></td></tr>
                <%else %>
                    <tr><td align="right" colspan="4"><img src="../images/heavyfreight2.jpg" width="250" height="247" /></td></tr>
                <%End if %>
                
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
 Case "3"            
         %>

    <table border="0" cellpadding="0" cellspacing="0" align="center">

       
        <tr>
            <td align="center" colspan="2">
            <table border="0" cellpadding="0" cellspacing="0" width="100%">
                <tr> <td valign="top">



                <form method="post" name="OrderForm1" action="FreightOrder.asp?Internal=<%=Internal%>">
        <table border="0" bordercolor="<%=BorderColor%>" cellpadding="3" cellspacing="0" width="<%=tablewidth%>" align="center"> 
        <tr>
            <td align="center">
                <img src="../images/boxes.jpg" height="287" width="300" />
            </td>
            <td>

                <table cellpadding="3" cellspacing="0" width="100%">
               
                    <tr><td align="left" colspan="3"><input id="gobutton" type="submit" name="buttonsubmit" value="Add/Edit Locations in your Address Book" /></td></tr>
                <tr><td>&nbsp;</td></tr>               
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap>Origination</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left">
								<%
                                'Response.write "PreExistingOrigination="&PreExistingOrigination&"<BR>"
                                 %>
                                <select name="PreExistingOrigination" ID="Select3"  onchange="form.submit()">
								<option value="" <%if trim(PreExistingOrigination)="" then response.Write " selected" end if%>>Select your origination</option>
                                <%

									Set oConn = Server.CreateObject("ADODB.Connection")
									oConn.ConnectionTimeout = 100
									oConn.Provider = "MSDASQL"
									oConn.Open DATABASE
                                        If trim(XSquare)="y" then
                                            l_cSQL = "Select * FROM PreExistingCompanies WHERE companystatus='c' and (CompanyName='EBT Probe Card Ship Room' or CompanyName='SC Building Probe Card Shop') ORDER BY Contact Name, CompanyName"
                                        else
										    l_cSQL = "Select * FROM PreExistingCompanies WHERE companystatus='c' and (CompanyOwner='"& trim(UserID) &"' or CompanyOwner='"& trim(PreExistingRequestor) &"' ) ORDER BY ContactName, CompanyName, CompanyAddress, CompanySuite"
                                        End if
										SET oRs = oConn.Execute(l_cSql)
												Do While not oRs.EOF
                                                CompanyID=oRs("CompanyID")
												CompanyName=oRs("CompanyName")
                                                CompanyBuilding=oRs("CompanyBuilding")
												CompanyAddress=oRs("CompanyAddress")
                                                CompanySuite=oRs("CompanySuite")
												CompanyCity=oRs("CompanyCity")
                                                CompanyState=oRs("CompanyState")
                                                CompanyZip=oRs("CompanyZip")
                                                ContactName=oRs("ContactName")
                                                CostCenter=oRs("CompanyCostCenter")
                                                CompanyPhone=oRs("CompanyPhone")
                                                CompanyEmail=oRs("CompanyEmail")
                                                CompanyisCourier=oRs("IsCourier")
                                                CompanyAliasCode=oRs("st_alias")
                                                If trim(PreExistingOrigination)=trim(CompanyID) and trim(PreExistingOrigination)>"" then
												    OriginationID=CompanyID
                                                    
                                                    OriginationCompany=CompanyName
                                                    OriginationBuilding=CompanyBuilding
												    OriginationAddress=CompanyAddress
                                                    OriginationSuite=CompanySuite
												    OriginationCity=CompanyCity
                                                    OriginationState=CompanyState
                                                    OriginationZipCode=CompanyZip
                                                    OriginationIsCourier=CompanyIsCourier
                                                    OriginationAliasCode=CompanyAliasCode
                                                    OriginationCostCenter=CostCenter
                                                    tempOriginationContactName=ContactName
                                                    tempOriginationPhoneNumber=CompanyPhone
                                                    tempOriginationEmail=CompanyEmail
                                                    If trim(tempOriginationContactName)>"" then OriginationContactName=tempOriginationContactName end if
                                                    If trim(tempOriginationPhoneNumber)>"" then OriginationPhoneNumber=tempOriginationPhoneNumber end if
                                                    If trim(tempOriginationEmail)>"" then OriginationEmail=tempOriginationEmail end if
                                                End if								
											%>
											<option value="<%=CompanyID%>" <%if trim(PreExistingOrigination)=trim(CompanyID) then response.Write " selected" end if%>><%=ContactName %> | <%=CompanyName%> | <%=CompanyAddress%> | <%=CompanySuite%></option>
											<%
										oRs.movenext
										LOOP
									Set oConn=Nothing									
									%>
								</select> 
                                <%
                                'Response.write "PreExistingRequestor="&PreExistingRequestor&"<br>"
                                'Response.Write "l_cSQL="&l_cSQL&"<BR>" 
                                %>
                    </td>
                </tr>
                <input type="hidden" name="orderstatus" value="3" />
                </form>
                 <form method="post" name="OrderForm111" action="FreightOrder.asp?Internal=<%=Internal%>">
                <%
                'Response.write "l_cSQL="&l_cSQL&"<BR>"
                 %>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap>Company Name</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left" nowrap class="FleetExpressTextBlack"><%=OriginationCompany %><input type="hidden" name="OriginationCompany" value="<%=OriginationCompany%>" size="45" maxlength="40" /></td>
                </tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap>Building</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left" nowrap class="FleetExpressTextBlack"><%=OriginationBuilding %><input type="hidden" name="OriginationBuilding" value="<%=OriginationBuilding%>" size="45" maxlength="40" /></td>
                </tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">Address</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left" class="FleetExpressTextBlack"><%=OriginationAddress%><input type="hidden" name="OriginationAddress" value="<%=OriginationAddress%>" size="45" maxlength="40" /></td>
                </tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">Suite/Dock</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left" class="FleetExpressTextBlack"><%=OriginationSuite%><input type="hidden" name="OriginationSuite" value="<%=OriginationSuite%>" size="45" maxlength="40" /></td>
                </tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">City/State/Zip</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td class="FleetExpressTextBlack"><%=OriginationCity%>
                       
                        <input type="hidden" name="OriginationCity" value="<%=OriginationCity%>" size="20" maxlength="30" />
                        &nbsp;TX&nbsp;&nbsp;
                        <%
                        'Response.write "OriginationID="&OriginationID&"<BR>" 
                        %>
                        <input type="hidden" name="OriginationID" value="<%=OriginationID%>" />
                        <input type="hidden" name="OriginationState" value="TX"><%=OriginationZipCode%>
                        <input type="hidden" name="OriginationAliasCode" value="<%=OriginationAliasCode%>" />
                        <input type="Hidden" name="OriginationZipCode" value="<%=OriginationZipCode%>" />
                         <input type="Hidden" name="OriginationIsCourier" value="<%=OriginationIsCourier%>" />

                    </td>
                </tr>
                
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">Cost Center</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left" class="FleetExpressTextBlack"><input type="hidden" name="OriginationCostCenter" value="<%=OriginationCostCenter%>" size="45" maxlength="50" /><%=OriginationCostCenter%></td>
                </tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">Contact Name</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left" class="FleetExpressTextBlack"><input type="hidden" name="OriginationContactName" value="<%=OriginationContactName%>" size="45" maxlength="25" /><%=OriginationContactName%></td>
                </tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">Phone Number</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left" class="FleetExpressTextBlack"><%=OriginationPhoneNumber%><input type="hidden" name="OriginationPhoneNumber" value="<%=OriginationPhoneNumber%>" size="45" maxlength="20" /></td>
                </tr> 
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">Email Address</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left" class="FleetExpressTextBlack"><%=OriginationEmail%><input type="hidden" name="OriginationEmail" value="<%=OriginationEmail%>" size="45" maxlength="100" /></td>
                </tr> 
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap>Receive Notifications</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left" class="FleetExpressTextBlack"><input type="radio" name="OriginationNotifications" value="y" <%If trim(OriginationNotifications)="" or trim(OriginationNotifications)="y" then Response.write "Checked" end if %>/>Yes&nbsp;&nbsp;&nbsp;&nbsp;<input type="radio" name="OriginationNotifications" value="n" <%If trim(OriginationNotifications)="n" then Response.write "Checked" end if %>/>No</td>
                </tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap>Receive Copy of Waybill</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left" class="FleetExpressTextBlack"><input type="radio" name="OriginationWaybill" value="y" <%If trim(OriginationWaybill)="" or trim(OriginationWaybill)="y" then Response.write "Checked" end if %>/>Yes&nbsp;&nbsp;&nbsp;&nbsp;<input type="radio" name="OriginationWaybill" value="n" <%If trim(OriginationWaybill)="n" then Response.write "Checked" end if %>/>No</td>
                </tr>
                <%
                If trim(XSquare)="y" then

                else
                If trim(PickUpDate)="" then
                    PickUpDate=date()
                End if 
                If trim(PickUpTime)="" then
                    TheHours=Hour(Time())
                    TheMinutes=Minute(Time())
                    TheMinutes=cint(TheMinutes)+10
                    If TheMinutes>60 then
                        TheMinutes=TheMinutes-60
                        TheHours=TheHours+1
                    End if
                    If len(TheMinutes)=1 then
                        TheMinutes="0"&TheMinutes
                    End if
                    'Response.write "TheMinutes="&TheMinutes&"<BR>"
                    If Int(TheHours)>12 then
                        'Response.write "Got here!<BR>"
                        TheHours=TheHours-12
                    End if
                    AMPM=Right(Time(),2)
                    TheTime=TheHours&":"&TheMinutes&" "&AMPM
                    PickUpTime=TheTime
                End if
                %>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">Ready Date</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left" nowrap class="FleetExpressTextBlack">
                        <input type="text" name="PickUpDate" id="date_1" value="<%=PickUpDate%>"/>
		            </td>
                </tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">Ready Time</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left" nowrap class="FleetExpressTextBlack">
                    <input type="text" name="PickUpTime" id="time_1" value="<%=PickUpTime%>"/>                 
                    </td>
                </tr>








                <%end if %>
                    <tr><td><img src="../images/pixel.gif" width="1" height="35" /></td></tr>
                <input type="hidden" name="OrderStatus" value="4" />
                <tr><td colspan="2">
                    <input type="submit" id="gobutton" name="var2" value="<<<Back" />
                </td><td align="right">
                <%If trim(ErrorMessage)="" then %>
                <input id="gobutton" type="submit" name="buttonsubmit" value="Next >>>" />
                <%end if %>
                </td></tr>
                 </form>
                 </table>
                
                 </td></tr></table>
<%
Case "4"
    'Response.write "PickUpDateTime="&PickUpDateTime&"<BR>"
    PUCurrentHour=Hour(PickUpDateTime)
    If ((PUCurrentHour<6 or PUCurrentHour>16) or WeekdayName(Weekday(PickUpDateTime))="Saturday" or WeekdayName(Weekday(PickUpDateTime))="Sunday") and lcase(vehicletype)="tractor" then                  
        'Response.write "VehicleType="&VehicleType&"<BR>"
        TractorAvailable="n"
        ErrorMessage="We are unable to handle an order of this size during off hours.  You will either need to break this<BR>order into smaller orders of 10 skids maximum, or place this order Monday - Friday between 6:00 AM - 4:00 PM."
    End if

 %>
    <table border="0" cellpadding="0" cellspacing="0" align="center">
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
       
        <tr>
            <td align="center">
            <table border="0" cellpadding="0" cellspacing="0" width="100%">
                <tr> <td valign="top">



                <form method="post" name="OrderForm1" action="FreightOrder.asp?Internal=<%=Internal%>">

                    <table border="0" bordercolor="<%=BorderColor%>" cellpadding="0" cellspacing="0" width="<%=tablewidth%>"> 
                    <tr> <td valign="top"> 
                    <table cellpadding="3" cellspacing="0" width="100%"> 
                    <tr><td align="left" colspan="3"><input id="gobutton" type="submit" name="buttonsubmit" value="Add/Edit Locations in your Address Book" /></td></tr>
                <tr><td>&nbsp;</td></tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap>Destination</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left">
								<%
                                'Response.write "PreExistingDestination="&PreExistingDestination&"<BR>"
                                 %>
                                <select name="PreExistingDestination" ID="Select1"  onchange="form.submit()">
								<option value="" <%if trim(PreExistingDestination)="" then response.Write " selected" end if%>>Select your destination</option>
                                <%

									Set oConn = Server.CreateObject("ADODB.Connection")
									oConn.ConnectionTimeout = 100
									oConn.Provider = "MSDASQL"
									oConn.Open DATABASE
                                        If trim(XSquare)="y" then
                                            l_cSQL = "Select * FROM PreExistingCompanies WHERE companystatus='c' and (CompanyName='EBT Probe Card Ship Room' or CompanyName='SC Building Probe Card Shop') ORDER BY ContactName, CompanyName"
                                            else
										    'l_cSQL = "Select * FROM PreExistingCompanies WHERE companystatus='c' and CompanyOwner='"& trim(UserID) & "' ORDER BY ContactName, CompanyName, CompanyAddress, CompanySuite"
                                            l_cSQL = "Select * FROM PreExistingCompanies WHERE companystatus='c' and (CompanyOwner='"& trim(UserID) &"' or CompanyOwner='"& trim(PreExistingRequestor) &"' ) ORDER BY ContactName, CompanyName, CompanyAddress, CompanySuite"
                                        End if
										SET oRs = oConn.Execute(l_cSql)
												Do While not oRs.EOF
                                                bCompanyID=oRs("CompanyID")
												bCompanyName=oRs("CompanyName")
                                                bCompanyBuilding=oRs("CompanyBuilding")
												bCompanyAddress=oRs("CompanyAddress")
                                                bCompanySuite=oRs("CompanySuite")
												bCompanyCity=oRs("CompanyCity")
                                                bCompanyState=oRs("CompanyState")
                                                bCompanyZip=oRs("CompanyZip")
                                                bCompanyIsCourier=oRS("IsCourier")
                                                bCompanyAliasCode=oRs("st_alias")
                                                bCostCenter=oRs("CompanyCostCenter")
                                                bContactName=oRs("ContactName")
                                                bCompanyPhone=oRs("CompanyPhone")
                                                bCompanyEmail=oRs("CompanyEmail")
                                                If trim(PreExistingDestination)=trim(bCompanyID) and trim(PreExistingDestination)>"" then
												    DestinationID=bCompanyID
                                                    DestinationCompany=bCompanyName
                                                    DestinationBuilding=bCompanyBuilding
												    DestinationAddress=bCompanyAddress
                                                    DestinationSuite=bCompanySuite
												    DestinationCity=bCompanyCity
                                                    DestinationState=bCompanyState
                                                    DestinationZipCode=bCompanyZip
                                                    DestinationIsCourier=bCompanyIsCourier
                                                    DestinationAliasCode=bCompanyAliasCode
                                                    tempDestinationContactName=bContactName
                                                    DestinationCostCenter=bCostCenter
                                                    tempDestinationPhoneNumber=bCompanyPhone
                                                    tempDestinationEmail=bCompanyEmail
                                                    If trim(tempDestinationContactName)>"" then DestinationContactName=tempDestinationContactName end if
                                                    If trim(tempDestinationPhoneNumber)>"" then DestinationPhoneNumber=tempDestinationPhoneNumber end if
                                                    If trim(tempDestinationEmail)>"" then DestinationEmail=tempDestinationEmail end if
                                                End if								
											%>
											<option value="<%=bCompanyID%>" <%if trim(PreExistingDestination)=trim(bCompanyID) then response.Write " selected" end if%>><%=bContactName %> | <%=bCompanyName%> | <%=bCompanyAddress%> | <%=bCompanySuite%></option>
											<%
										oRs.movenext
										LOOP
									Set oConn=Nothing									
									%>
								</select> 
                    </td>
                </tr>
                <input type="hidden" name="orderstatus" value="4" />
                </form>
                 <form method="post" name="OrderForm1112" action="FreightOrder.asp?Internal=<%=Internal%>">
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left" nowrap>Company Name</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left" nowrap class="FleetExpressTextBlack"><%=DESTINATIONCompany%><input type="hidden" name="DESTINATIONCompany" value="<%=DESTINATIONCompany%>" size="45" maxlength="40" /></td>
                    </tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Building</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=DESTINATIONBuilding%><input type="hidden" name="DESTINATIONBuilding" value="<%=DESTINATIONBuilding%>" size="45" maxlength="40" /></td>
                    </tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Address</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=DESTINATIONAddress%><input type="hidden" name="DESTINATIONAddress" value="<%=DESTINATIONAddress%>" size="45" maxlength="40" /></td>
                    </tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Suite</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=DESTINATIONSuite%><input type="hidden" name="DESTINATIONSuite" value="<%=DESTINATIONSuite%>" size="45" maxlength="40" /></td>
                    </tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">City/State/Zip</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td class="FleetExpressTextBlack">
                            <%=DESTINATIONCity%><input type="hidden" name="DESTINATIONCity" value="<%=DESTINATIONCity%>" size="20" maxlength="30" />
                            &nbsp;TX&nbsp;&nbsp;
                        <%
                        'Response.write "DestinationID="&DestinationID&"<BR>" 
                        %>
                        <input type="hidden" name="DestinationID" value="<%=DestinationID%>" />
                        <input type="hidden" name="DestinationState" value="TX">
                        <input type="hidden" name="DestinationAliasCode" value="<%=DestinationAliasCode%>">
                            <%=DESTINATIONZipCode%><input type="hidden" name="DESTINATIONZipCode" value="<%=DESTINATIONZipCode%>" />
                            <input type="hidden" name="DESTINATIONIsCourier" value="<%=DESTINATIONisCourier%>" />

                        </td>
                    </tr>
                    
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Cost Center</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=DestinationCostCenter%><input type="hidden" name="DestinationCostCenter" value="<%=DestinationCostCenter%>" size="45" maxlength="25" /></td>
                    </tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Contact Name</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=DESTINATIONContactName%><input type="hidden" name="DESTINATIONContactName" value="<%=DESTINATIONContactName%>" size="45" maxlength="25" /></td>
                    </tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Phone Number</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=DESTINATIONPhoneNumber%><input type="hidden" name="DESTINATIONPhoneNumber" value="<%=DESTINATIONPhoneNumber%>" size="45" maxlength="20" /></td>
                    </tr> 
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Email Address</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=DESTINATIONEmail%><input type="hidden" name="DESTINATIONEmail" value="<%=DESTINATIONEmail%>" size="45" maxlength="100" /></td>
                    </tr>
                 <tr>
                    <td class="FleetExpressTextBlackBold" align="left">Receive Notifications</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left" class="FleetExpressTextBlack"><input type="radio" name="DestinationNotifications" value="y" <%If trim(DestinationNotifications)="" or trim(DestinationNotifications)="y" then Response.write "Checked" end if %>/>Yes&nbsp;&nbsp;&nbsp;&nbsp;<input type="radio" name="DestinationNotifications" value="n" <%If trim(DestinationNotifications)="n" then Response.write "Checked" end if %>/>No</td>
                </tr>
                <tr><td>&nbsp;</td></tr>
                <tr><td colspan="3" class="ErrorMessage">Please Note:  Deliveries do not require a signature. If no one is available to sign at the delivery, driver will leave package at front desk or in office.</td></tr>                    
                <%If trim(XSquare)="y" then

                else %>
                <!--
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">Delivery Date/Time</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left">
                    <input type="text" name="DeliveryDateTime" id="DeliveryDateTime" value="<%=DeliveryDateTime%>" size="30" maxlength="30" />
                     <a href="javascript:NewCal('DeliveryDateTime','MMddyyyy',true,12)"><img src="../images/cal.gif" width="16" height="16" border="0" alt="Pick a date"></a>
                    </td>
                </tr>
                -->
                <%end if%>
                   <tr><td><img src="../images/pixel.gif" width="1" height="35" /></td></tr>
                <input type="hidden" name="OrderStatus" value="5" />
                <tr><td colspan="2">
                <input type="submit" id="gobutton" name="var3" value="<<<Back" />
                </td><td align="right">
                <%If trim(TractorAvailable)<>"n" then %>
                <input id="gobutton" type="submit" name="buttonsubmit" value="Next >>>" />
                <%end if %>
                </td></tr>
                 </table>
                 </td>
                 <td><img src="../images/delivery.jpg" height="200" width="271" />
                 </td></tr></table>
  <%
   Case "5"

   EstimatedCost=0
   'Response.write "EstimatedCost6="&EstimatedCost&"<BR>"
    If cDate(PickupDateTime) < Now() then
        'Response.write "Line 2042<BR>"
        PickupDateTime=Now()
    End if

    'Response.write "PickUpDateTime="&PickUpDateTime&"<BR>"
    PUCurrentHour=Hour(PickUpDateTime)
    If ((PUCurrentHour<6 or PUCurrentHour>16) or WeekdayName(Weekday(PickUpDateTime))="Saturday" or WeekdayName(Weekday(PickUpDateTime))="Sunday") and lcase(vehicletype)="tractor" then                  
        TractorAvailable="n"
        ErrorMessage="We are unable to handle an order of this size during off hours.  You will either need to break this<BR>order into smaller orders of 10 skids maximum, or place this order Monday - Friday between 6:00 AM - 4:00 PM."
    End if



	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.ConnectionTimeout = 100
	oConn.Provider = "MSDASQL"
	oConn.Open DATABASE
		l_cSQL = "Select * FROM Priorities WHERE prioritystatus='c' and PriorityID='"& Priority & "'"
		SET oRs = oConn.Execute(l_cSql)
				If not oRs.EOF then
                'PriorityID=oRs("PriorityID")
                'varPriority=oRs("Priority")
				PriorityDescription=oRs("PriorityDescription")
				PriorityMinutes=oRs("PriorityMinutes")
                MaxSkids=oRs("MaxSkids")
                EachAdditionalSkidCost=oRs("EachAdditionalSkidCost")
                MaxLargeSkids=oRs("MaxLargeSkids")
				'Response.write "2902 MaxLargeSkids="&MaxLargeSkids&"<BR>"
                EachAdditionalLargeSkidCost=oRs("EachAdditionalLargeSkidCost")
		End if
	Set oConn=Nothing	
    '''Select Case Priority
        '''Case "3 HOUR"
    'Response.write "ShipmentType="&ShipmentType&"<BR>"
    'Response.write "XXXPickUpDateTime="&PickUpDateTime&"<BR>"
    DeliveryDateTime=DateAdd("n",PriorityMinutes, PickUpDateTime)
    'Response.write "Priority="&Priority&"<BR>"
    'Response.write "PriorityDescription="&PriorityDescription&"<BR>"
    If ucase(PriorityDescription)="NEXT DAY" then
    DeliveryDay=DateAdd("d", 1, PickUpDateTime)
    'Response.write "XXXXXXDeliveryDay="&DeliveryDay&"<BR>"
    DeliveryDay=FormatDateTime(DeliveryDay, vbShortDate)
    DeliveryDateTime=DeliveryDay&" 11:59:59 PM"
    End if
    'Response.write "XXXXXXDeliveryDateTime="&DeliveryDateTime&"<BR>"
    'Response.write "DueTime="&DueTime&"<BR>"
    'Response.write "BillToID="&BillToID&"<BR>"
    %> 
   
    <table border="0" cellpadding="0" cellspacing="0" align="center">
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
        <tr>
            <td align="center" colspan="2">
            <table border="0" cellpadding="0" cellspacing="0" width="100%">
                <tr> <td valign="top">
                <table border="0" cellpadding="3" cellspacing="0" width="100%">
                <tr><td class="FleetExpressTextBlackBold" align="center" nowrap="nowrap" colspan="5">Order Summary</td></tr>
                <tr><td>&nbsp;</td></tr>
                <tr><td>&nbsp;</td></tr>
<%
									Set oConn = Server.CreateObject("ADODB.Connection")
									oConn.ConnectionTimeout = 100
									oConn.Provider = "MSDASQL"
									oConn.Open DATABASE
										    l_cSQL = "Select * FROM PreExistingCompanies WHERE companystatus='c' and Companyid='"& PreExistingOrigination & "'"
                                            'Response.write "l_cSQL="&l_cSQL&"<BR>"
										SET oRs = oConn.Execute(l_cSql)
												If not oRs.EOF then
                                                CompanyID=oRs("CompanyID")
												CompanyName=oRs("CompanyName")
                                                CompanyBuilding=oRs("CompanyBuilding")
												CompanyAddress=oRs("CompanyAddress")
                                                CompanySuite=oRs("CompanySuite")
												CompanyCity=oRs("CompanyCity")
                                                CompanyState=oRs("CompanyState")
                                                CompanyZip=oRs("CompanyZip")
                                                CompanyIsCourier=oRs("IsCourier")
                                                CompanyAliasCode=oRs("st_alias")
                                                CompanyCostCenter=oRs("CompanyCostCenter")
                                                ContactName=oRs("ContactName")
                                                CompanyPhone=oRs("CompanyPhone")
                                                CompanyEmail=oRs("CompanyEmail")
                                                End if
									Set oConn=Nothing
 %>
                <tr>
                     <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">Origination</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>               
                    <td class="FleetExpressTextBlack"><%=ContactName%> | <%=CompanyName%> | <%=CompanyAddress%> | <%=CompanySuite%></td>
                </tr>
<%
									Set oConn = Server.CreateObject("ADODB.Connection")
									oConn.ConnectionTimeout = 100
									oConn.Provider = "MSDASQL"
									oConn.Open DATABASE
										    l_cSQL = "Select * FROM PreExistingCompanies WHERE companystatus='c' and Companyid='"& PreExistingDestination & "'"
                                            'Response.write "l_cSQL="&l_cSQL&"<BR>"
										SET oRs = oConn.Execute(l_cSql)
												If not oRs.EOF then
                                                bCompanyID=oRs("CompanyID")
												bCompanyName=oRs("CompanyName")
                                                bCompanyBuilding=oRs("CompanyBuilding")
												bCompanyAddress=oRs("CompanyAddress")
                                                bCompanySuite=oRs("CompanySuite")
												bCompanyCity=oRs("CompanyCity")
                                                bCompanyState=oRs("CompanyState")
                                                bCompanyZip=oRs("CompanyZip")
                                                bCompanyIsCourier=oRs("IsCourier")
                                                bCompanyAliasCode=oRs("st_alias")
                                                bcompanyCostCenter=oRs("CompanyCostCenter")
                                                bContactName=oRs("ContactName")
                                                bCompanyPhone=oRs("CompanyPhone")
                                                bCompanyEmail=oRs("CompanyEmail")
                                                End if
									Set oConn=Nothing	
 %>
                <tr>
                     <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">Destination</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>               
                    <td class="FleetExpressTextBlack"><%=bContactName%> | <%=bCompanyName%> | <%=bCompanyAddress%> | <%=bCompanySuite%></td>
                </tr>


                <tr>
                     <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">Ready Time</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>               
                    <td class="FleetExpressTextBlack"><%=PickupDateTime %></td>
                </tr>
                <%
                StandingReadyTime=hour(PickUpDateTime)
                DayOfWeek=Weekday(PickUpDateTime)

                StandingOrigination=InStr(uCASE(CompanyBuilding),"STANDING")
                StandingDestination=InStr(uCASE(bCompanyBuilding),"STANDING")




                If DayOfWeek>1 and DayOfWeek<7 and StandingOrigination>0 and StandingDestination>0 then
                        'Standing order Type 1
                        If (StandingReadyTime=10 or StandingReadyTime=7 or StandingReadyTime=8 or StandingReadyTime=9) and left(companyaddress, 5)="13020" and (left(bcompanyaddress, 4)="3601" or left(bcompanyaddress, 3)="300") then
                            StandingOrderID=1
                            DeliveryDateTime=DateValue(PickUpDateTime)&" 10:00:00 AM"
                            'Response.write "GOT HERE!!!!!<BR>"
                            'Response.write "DeliveryDateTime="&DeliveryDateTime&"<BR>"
                        End if
                        'Standing order Type 2
                        If (StandingReadyTime=6 or StandingReadyTime=7 or StandingReadyTime=8 or StandingReadyTime=9) and left(companyaddress, 5)="13536" and left(bcompanyaddress, 5)="13601" then
                            StandingOrderID=2
                            DeliveryDateTime=DateValue(PickUpDateTime)&" 11:00:00 AM"
                            'Response.write "GOT HERE!!!!!<BR>"
                            'Response.write "DeliveryDateTime="&DeliveryDateTime&"<BR>"
                        End if
                        'Standing order Type 3
                        If (StandingReadyTime=6 or StandingReadyTime=7 or StandingReadyTime=8 or StandingReadyTime=9) and left(companyaddress, 5)="13438" and left(bcompanyaddress, 5)="13536" then
                            StandingOrderID=3
                            DeliveryDateTime=DateValue(PickUpDateTime)&" 10:00:00 AM"
                            'Response.write "GOT HERE!!!!!<BR>"
                            'Response.write "DeliveryDateTime="&DeliveryDateTime&"<BR>"
                        End if
                        'Standing order Type 4
                        If (StandingReadyTime=12 or StandingReadyTime=11 or StandingReadyTime=10 or StandingReadyTime=9) and left(companyaddress, 5)="13601" and left(bcompanyaddress, 5)="13536" then
                            StandingOrderID=4
                            DeliveryDateTime=DateValue(PickUpDateTime)&" 1:00:00 PM"
                            'Response.write "GOT HERE!!!!!<BR>"
                            'Response.write "DeliveryDateTime="&DeliveryDateTime&"<BR>"
                        End if
                        'Standing order Type 5
                        If (StandingReadyTime=12 or StandingReadyTime=11 or StandingReadyTime=13 or StandingReadyTime=14) and left(companyaddress, 5)="13532" and left(bcompanyaddress, 5)="12500" then
                            StandingOrderID=5
                            DeliveryDateTime=DateValue(PickUpDateTime)&" 2:00:00 PM"
                            'Response.write "GOT HERE!!!!!<BR>"
                            'Response.write "DeliveryDateTime="&DeliveryDateTime&"<BR>"
                        End if
                        IsCSF=0
                        'Standing order Type 6
                        'IsCSF=InStr(uCASE(DestinationBuilding),"CSF")
                        'If IsCSF<1 then
                        '    IsCSF=InStr(uCASE(DestinationBuilding),"CSSF")
                        'End if
                        'If IsCSF<1 then
                        '    IsCSF=InStr(uCASE(DestinationBuilding),"CENTRAL SHIPPING")
                        'End if
                        'Response.write "IsCSF="&IsCSF&"<BR>"
                        If (StandingReadyTime=12 or StandingReadyTime=11 or StandingReadyTime=13 or StandingReadyTime=14) and left(companyaddress, 5)="12500" and left(bcompanyaddress, 5)="13536" then
                            StandingOrderID=6
                            DeliveryDateTime=DateValue(PickUpDateTime)&" 2:30:00 PM"
                            'Response.write "GOT HERE!!!!!<BR>"
                            'Response.write "DeliveryDateTime="&DeliveryDateTime&"<BR>"
                        End if
                        'Standing order Type 7
                        If (StandingReadyTime=11 or StandingReadyTime=12 or StandingReadyTime=13 or StandingReadyTime=14) and left(companyaddress, 5)="13536" and (left(bcompanyaddress, 4)="2580") then
                            StandingOrderID=7
                            DeliveryDateTime=DateValue(PickUpDateTime)&" 03:00:00 PM"
                            'Response.write "GOT HERE!!!!!<BR>"
                            'Response.write "DeliveryDateTime="&DeliveryDateTime&"<BR>"
                        End if
                End if

                If StandingOrderID>0 then
                    Priority="12"
                    PriorityDescription="Standing Order"
                End if
                 %>
                <tr>
                     <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">Priority</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>               
                    <td class="FleetExpressTextBlack"><%=PriorityDescription %></td>
                </tr>
                <tr>
                     <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">Estimated Delivery Time</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>               
                    <td class="FleetExpressTextBlack"><%=DeliveryDateTime %></td>
                </tr>
                <input type="hidden" name="DeliveryDateTime" value="<%=DeliveryDateTime%>" />
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap" colspan="3">Estimated Charges:</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                        <%
                        'Response.write "pieces="&pieces&"<BR>"
                        'Response.write "rf_box="&rf_box&"<BR>"
                        'Response.write "MaxSkids="&MaxSkids&"<BR>"
If trim(StandingOrderID)="" then
                        If trim(ucase(rf_box))="SKIDS" then
                            'Response.write "got here line 2414<BR>"
                            If cint(pieces)>cdbl(MaxSkids) then
                                'Response.write "got here line 2416<BR>"
                                AddSkids=cInt(pieces)-cdbl(MaxSkids)
                                AddSkidsCost=AddSkids*EachAdditionalSkidCost

                            End if
                        End if
''''''''''''''''''''''''''''''''''''''''''''''''''
'Response.write "trim(ucase(rf_box))="&trim(ucase(rf_box))&"<BR>"
                        If trim(ucase(rf_box))="LARGE SKIDS" then
                            'Response.write "got here line 2414<BR>"
                            If cint(pieces)>cdbl(MaxLargeSkids) then
                                'Response.write "got here line 2416<BR>"
                                AddLargeSkids=cInt(pieces)-cdbl(MaxLargeSkids)
                                AddLargeSkidsCost=AddLargeSkids*EachAdditionalLargeSkidCost

                            End if
                        End if
						'Response.write "AddLargeSkids="&AddLargeSkids&"<BR>"
                        'Response.write "AddSkids="&AddSkids&"<BR>"
                        'Response.write "AddSkidsCost="&AddSkidsCost&"<BR>"

									Set oConn = Server.CreateObject("ADODB.Connection")
									oConn.ConnectionTimeout = 100
									oConn.Provider = "MSDASQL"
									oConn.Open DATABASE
									    l_cSQL = "Select * FROM Priorities WHERE priorityID='"&Priority&"' and prioritystatus='c' and priority_BT_ID='"& LegacyBillToID & "' ORDER BY PriorityMinutes desc"
										'Response.write "l_cSQL="&l_cSQL&"<BR>"
                                        SET oRs = oConn.Execute(l_cSql)
												If not oRs.EOF then
                                                    PriorityID=oRs("PriorityID")
                                                    varPriority=oRs("Priority")
												    PriorityDescription=oRs("PriorityDescription")
                                                    PriorityDescr=oRs("PriorityDescription")&" Service Level"
                                                    PriorityCost=oRs("PriorityCost")
                                                    'RateCharge=oRs("PriorityCost"
                                                    If PriorityCost>0 then
                                                        PriorityCost=FormatNumber((PriorityCost),2)
                                                    End if
												    PriorityMinutes=oRs("PriorityMinutes")
                                                End if
									Set oConn=Nothing
                                    If trim(PriorityCost)="" or LegacyBillToID="92" then PriorityCost=0 End if

                                    'REsponse.write "XXXXRateCharge2="&RateCharge2&"<BR>"


                                    'Response.write "EstimatedCost7="&EstimatedCost&"<BR>"
                                    If PriorityCost>0 then
                                       EstimatedCost=cDbl(EstimatedCost)+cDbl(PriorityCost)
                                    %>
                                             <tr>
                                                 <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=PriorityDescr %></td>
                                                <td width="10"><img src="../images/pixel.gif" /></td>               
                                                <td class="FleetExpressTextBlack">$<%=PriorityCost %></td>
                                            </tr>
                                        <%
                                    End if





                        Set oConn = Server.CreateObject("ADODB.Connection")
		                    oConn.ConnectionTimeout = 100
		                    oConn.Provider = "MSDASQL"
		                    oConn.Open DATABASE
                                'If trim(XSquare)="y" then
				                    '''l_cSQL = "Select * FROM PreExistingRequestor WHERE requestorstatus='c' and requestorID IN('1') ORDER BY RequestorName"
                                    ' l_cSQL = "Select * FROM PreExistingRequestor WHERE requestorstatus='c' and requestorID>'199' ORDER BY RequestorName"
                                    'else
                                    l_cSQL = "Select rtDescr, RateCharge, rtBillCode FROM RateList INNER JOIN RateType ON RateList.rtid = RateType.rtid WHERE (bt_id='"& LegacyBillToID &"') AND RateStatus='c'"
                                'End if
                                'Response.write "L_cSQL="&L_cSQL&"<BR>"
			                    SET oRs = oConn.Execute(l_cSql)
                                        If oRs.EOF then

                                        End if
					                    Do while not oRs.EOF
                                        rtDescr=trim(oRs("rtDescr"))
					                    RateCharge=trim(oRs("RateCharge"))
					                    rtBillCode=trim(oRs("rtBillCode"))
                                        EstimatedCost=cDbl(EstimatedCost)+cDbl(RateCharge)
                                        'Response.write "rtDescr="&rtDescr&"<BR>"
                                        'Response.write "RateCharge="&RateCharge&"<BR>"
                                        'Response.write "rtBillCode="&rtBillCode&"<BR>"
                                        'Response.write "EstimatedCost8="&EstimatedCost&"<BR>"
                                        %>
                                            <tr>
                                                 <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rtDescr %></td>
                                                <td width="10"><img src="../images/pixel.gif" /></td>               
                                                <td class="FleetExpressTextBlack">$<%=RateCharge %></td>
                                            </tr>
                                        <%

										oRs.movenext
										LOOP
		                    Set oConn=Nothing
                            'response.write "l_cSQL="&l_cSQL&"<BR>"
ELSE
                                    IsStandingOrder="y"
									Set oConn = Server.CreateObject("ADODB.Connection")
									oConn.ConnectionTimeout = 100
									oConn.Provider = "MSDASQL"
									oConn.Open DATABASE
									    l_cSQL = "Select StandingOrderDescription, Charge FROM StandingOrderFees WHERE StandingOrderID='"&StandingOrderID&"'"
										'Response.write "l_cSQL="&l_cSQL&"<BR>"
                                        SET oRs = oConn.Execute(l_cSql)
												If not oRs.EOF then
                                                    StandingOrderDescription=oRs("StandingOrderDescription")
                                                    StandingOrderCharge=oRs("Charge")
                                                End if
									Set oConn=Nothing
                                    StandingOrderCharge=FormatNumber((StandingOrderCharge),2)
                                    EstimatedCost=EstimatedCost+StandingOrderCharge
                                        rtDescr=StandingOrderDescription
					                    RateCharge=StandingOrderCharge
					                    rtBillCode="STND ORD"


                                        %>
                                            <tr>
                                                 <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=StandingOrderDescription %></td>
                                                <td width="10"><img src="../images/pixel.gif" /></td>               
                                                <td class="FleetExpressTextBlack">$<%=StandingOrderCharge %></td>
                                            </tr>
                                            
                                        <%
END IF
''''''''''''''''''''''''''ADD SKIDS''''''''''''''''''''''''''''''''''''''''''
                    If trim(AddSkids)>"" then
                        If AddSkids>0 then
                        %>
                            <tr>
                                    <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=AddSkids%> Additional Skids</td>
                                <td width="10"><img src="../images/pixel.gif" /></td>               
                                <td class="FleetExpressTextBlack">$<%=FormatNumber(AddSkidsCost,2) %></td>
                            </tr>
                        <%
                        EstimatedCost=EstimatedCost+AddSkidsCost
                        'Response.write "EstimatedCost4="&EstimatedCost&"<BR>"
                        End if
                    End if

'''''''''''''''''''''''''END ADD SKIDS'''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''ADD Large SKIDS''''''''''''''''''''''''''''''''''''''''''
                    If trim(AddLargeSkids)>"" then
                        If AddLargeSkids>0 then
                        %>
                            <tr>
                                    <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=AddLargeSkids%> Additional Large Skids</td>
                                <td width="10"><img src="../images/pixel.gif" /></td>               
                                <td class="FleetExpressTextBlack">$<%=FormatNumber(AddLargeSkidsCost,2) %></td>
                            </tr>
                        <%
                        EstimatedCost=EstimatedCost+AddLargeSkidsCost
                        'Response.write "EstimatedCost4="&EstimatedCost&"<BR>"
                        End if
                    End if

'''''''''''''''''''''''''END ADD SKIDS'''''''''''''''''''''''''''''''''''''''
'''''''''''''''CAPBOB''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If EstimatedCost>250 then
	Discount=EstimatedCost-250
	EstimatedCost=250
	CAPBOB="y"
	%>
                            <tr>
                                    <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap"><font color="red">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cap Rate Discount</font></td>
                                <td width="10"><img src="../images/pixel.gif" /></td>               
                                <td class="FleetExpressTextBlack"><font color="red">-$<%=FormatNumber(Discount,2) %></font></td>
                            </tr>
							
	<%
	else
	CAPBOB="n"
End if




                        Set oConn = Server.CreateObject("ADODB.Connection")
		                    oConn.ConnectionTimeout = 100
		                    oConn.Provider = "MSDASQL"
		                    oConn.Open DATABASE
                                'If trim(XSquare)="y" then
				                    '''l_cSQL = "Select * FROM PreExistingRequestor WHERE requestorstatus='c' and requestorID IN('1') ORDER BY RequestorName"
                                    ' l_cSQL = "Select * FROM PreExistingRequestor WHERE requestorstatus='c' and requestorID>'199' ORDER BY RequestorName"
                                    'else
                                    l_cSQL = "Select FuelCharge FROM FuelChargeList WHERE fuelchargeStatus='c'"
                                'End if
			                    SET oRs = oConn.Execute(l_cSql)
					                    If not oRs.EOF then
										'Response.write "EstimatedCost3="&EstimatedCost&"<BR>"
                                        FuelCharge=trim(oRs("FuelCharge"))
                                        varFuelCharge=FormatNumber((FuelCharge/100),2)
                                        FuelChargeDollars=FormatNumber((EStimatedCost*varFuelCharge),2)
                                        EstimatedCost=EstimatedCost+FuelChargeDollars

										'Response.write "FuelCharge="&FuelCharge&"<BR>"
										'Response.write "varFuelCharge="&varFuelCharge&"<BR>"
										'Response.write "FuelChargeDollars="&FuelChargeDollars&"<BR>"
                                        'Response.write "EstimatedCost3="&EstimatedCost&"<BR>"
                                        %>
                                            <tr>
                                                 <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Fuel Charge</td>
                                                <td width="10"><img src="../images/pixel.gif" /></td>               
                                                <td class="FleetExpressTextBlack">$<%=FuelChargeDollars %></td>
                                            </tr>
                                        <%

			                    End if
		                    Set oConn=Nothing

If Shouldthisberemoved="YES" then
                    If trim(AddSkids)>"" then
                        If AddSkids>0 then
                        %>
                            <tr>
                                    <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=AddSkids%> Additional Skids</td>
                                <td width="10"><img src="../images/pixel.gif" /></td>               
                                <td class="FleetExpressTextBlack">$<%=FormatNumber(AddSkidsCost,2) %></td>
                            </tr>
                        <%
                        EstimatedCost=EstimatedCost+AddSkidsCost
                        'Response.write "EstimatedCost4="&EstimatedCost&"<BR>"
                        End if
                    End if
''''''''''''''''''''''''''''''''''''''''''
                    If trim(AddLargeSkids)>"" then
                        If AddLargeSkids>0 then
                        %>
                            <tr>
                                    <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=AddLargeSkids%> Additional Large Skids</td>
                                <td width="10"><img src="../images/pixel.gif" /></td>               
                                <td class="FleetExpressTextBlack">$<%=FormatNumber(AddLargeSkidsCost,2) %></td>
                            </tr>
                        <%
                        EstimatedCost=EstimatedCost+AddLargeSkidsCost
                        'Response.write "EstimatedCost4="&EstimatedCost&"<BR>"
                        End if
                    End if
End if
                    %>


                <tr><td>&nbsp;</td></tr>
                <tr>
                        <td class="FleetExpressTextBlackBoldLarge" align="left" nowrap="nowrap">Total <font color="#d71e26"><i>Estimated</i></font> Cost</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>               
                    <td class="FleetExpressTextBlackBoldLarge">$<%=FormatNumber(EstimatedCost,2) %></td>
                </tr>                
                </tr>
                

                <tr><td><img src="../images/pixel.gif" width="1" height="35" /></td></tr>
                 <%'If sBT_ID=92 then%>
                    <form method="post" name="OrderForm1" action="FreightOrder.asp?Internal=<%=Internal%>">
					<Input type="hidden" name="CAPBOB" value="<%=CAPBOB%>">
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">
                            COST CENTER TO BE BILLED (You must select/enter one)
                            <input type="hidden" name="POorNWA" value="Cost Center #" />
                        </td>
                    </tr>
                                <!--
                                <select name="POorNWA">
                                    <option value="Cost Center #"<%If trim(POorNWA)="Cost Center #" then Response.write "selected" end if%>>Cost Center #</option>
                                    <option value="TI P/O #"<%If trim(POorNWA)="TI P/O #" then Response.write "selected" end if%>>TI P/O #</option>
                                </select>
                                -->
                    <%If trim(CompanyCostCenter)>"" then %>
                    <tr>
                        <td align="left"><input type="radio" name="GenericNumber" value="<%=CompanyCostCenter%>" <%if trim(GenericNumber)=trim(CompanyCostCenter) then response.write "checked" end if %>><%=CompanyCostCenter%> (<%=ContactName %>)</td>
                    </tr>
                    <%end if %>
                    <%If trim(bcompanyCostCenter)>"" then %>
                    <tr>
                        <td align="left"><input type="radio" name="GenericNumber" value="<%=bcompanyCostCenter%>" <%if trim(GenericNumber)=trim(bCompanyCostCenter) then response.write "checked" end if %>><%=bcompanyCostCenter%> (<%=bContactName %>)</td>
                    </tr>
                    <%end if %>
                    <tr>
                        <td align="left">
                        <input type="radio" name="GenericNumber" value="" <%if trim(GenericNumber)="" or (trim(GenericNumber)<>trim(CompanyCostCenter) AND trim(GenericNumber)<>trim(bCompanyCostCenter)) OR (trim(CompanyCostCenter)="" and trim(bCompanyCostCenter)="") then response.write "checked" end if %>>                       
                        <input type="text" name="bGenericNumber" maxlength="25" value="<%=bGenericNumber %>"</td>
                    </tr>

                    <tr><td colspan="3" class="FleetExpressTextBlackBold">(format: "C1"+3 digit division+5 digit cost center.  Ex: C112312345)</td></tr>
                    <input type="hidden" name="FuelChargeDollars" value="<%=FuelChargeDollars%>" />
                    <tr><td>&nbsp;</td></tr>
                    <tr><td colspan="3" class="FleetExpressTextBlackBold">OR, IF YOU ARE USING A PO NUMBER, ENTER IT BELOW:</td></tr>
                    <tr>
                        <td align="left">
                            <input type="text" name="PONumber" value="<%=PONumber%>"> 
                        </td>
                    </tr>                      
                    <%If trim(Priority)>"" then %>
                        <input type="hidden" name="Priority" value="<%=Priority %>" />
                    <%end if %>

                <%
                'Response.write "FuelChargeDollars="&Fuelchargedollars&"<BR>"
                If whatever="delete this part" then %>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">
                            PO NUMBER
                            <input type="hidden" name="POorNWA" value="P/O #" />
                                <!--
                                <select name="POorNWA">
                                    <option value="Cost Center #"<%If trim(POorNWA)="Cost Center #" then Response.write "selected" end if%>>Cost Center #</option>
                                    <option value="TI P/O #"<%If trim(POorNWA)="TI P/O #" then Response.write "selected" end if%>>TI P/O #</option>
                                </select>
                                -->
                        </td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left"><input type="text" name="GenericNumber" value="<%=GenericNumber%>"></td>

                        <input type="hidden" name="FuelChargeDollars" value="<%=FuelChargeDollars%>" />
                    </tr>
                <%
                
                End if %>               
                <tr><td><img src="../images/pixel.gif" width="1" height="25" /></td></tr>
                        <%
                        If TractorAvailable<>"n" then 
                        %>
                        <tr>
                            <td align="left">
                                &nbsp;
                            </td>
                    <td colspan>&nbsp;</td>
                    <td align="right">
                        <input type="hidden" name="OrderStatus" value="6" />
                        <input type="hidden" name="DeliveryDateTime" value="<%=DeliveryDateTime%>" />
                        <input type="hidden" name="AddSkids" value="<%=AddSkids%>" />
                        <input type="hidden" name="AddSkidsCost" value="<%=AddSkidsCost%>" />
                        <input type="hidden" name="AddLargeSkids" value="<%=AddLargeSkids%>" />
                        <input type="hidden" name="AddLargeSkidsCost" value="<%=AddLargeSkidsCost%>" />
                        <input type="hidden" name="FuelCharge" value="<%=FuelCharge%>" />
                        <input type="hidden" name="RtDescr" value="<%=RtDescr%>" />
                        <input type="hidden" name="RateCharge" value="<%=RateCharge%>" />

                        
                        <input type="hidden" name="rtBillCode" value="<%=rtBillCode%>" />
                        <input type="hidden" name="PriorityDescription" value="<%=PriorityDescription%>" />
                        <input type="hidden" name="PriorityCost" value="<%=PriorityCost%>" />
                        <input type="hidden" name="IsStandingOrder" value="<%=IsStandingOrder%>" />
                        <input id="gobutton" type="submit" name="buttonsubmit" value="Submit Order" />
                        <%End if %>
                        
                    </td>
                </tr>
                <tr><td>&nbsp;</td></tr>
                <tr>
                </form>
                    <td align="left">
                        <form method="post" name="OrderForm1" action="FreightOrder.asp?Internal=<%=Internal%>">
                        <input type="hidden" name="OrderStatus" value="4" />
                        <input id="gobutton" type="submit" name="buttonsubmit" value="<<< Back" />
                        </form>
                    </td>
                    <td colspan>&nbsp;</td>
                    <td align="right">
                        <form method="post" name="OrderForm1" action="FreightOrder.asp?Internal=<%=Internal%>">
                        <input type="hidden" name="OrderStatus" value="C" />
                        <input id="gobutton" type="submit" name="buttonsubmit" value="Cancel Order" />
                        </form>
                        </td>
                </tr>
                        <%
                        '1 = vbSunday - Sunday (default)
                        '2 = vbMonday - Monday
                        '3 = vbTuesday - Tuesday
                        '4 = vbWednesday - Wednesday
                        '5 = vbThursday - Thursday
                        '6 = vbFriday - Friday
                        '7 = vbSaturday - Saturday
                        'Response.write "WeekDay(Now())="&WeekDay(Now())&"<BR>"
                        'Response.write "Time()="&Time()&"<BR>"
                        'Response.write "Line 2483 VehicleType="&VehicleType&"<BR>"
                        'Response.write "Line 2484 ShipmentType="&ShipmentType&"<BR>"
                        'Response.write "Line 2484 CurrentHour="&CurrentHour&"<BR>"

            'If (CurrentHour<6 or CurrentHour>16) or WeekdayName(Weekday(now()))="Saturday" or WeekdayName(Weekday(now()))="Wednesday"  or WeekdayName(Weekday(now()))="Tuesday" then


%>
              
             
                </table>
           </td>
     </tr>

<%



   Case "C"

    %> 
   
    <table border="0" cellpadding="0" cellspacing="0" align="center">
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" height="100" width="1" /></td></tr>
        <tr>
            <td align="center" colspan="2">
            <table border="0" cellpadding="0" cellspacing="0" width="100%">
                <tr> <td valign="top">
                <table border="0" cellpadding="3" cellspacing="0" width="100%">
                
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap>Your order has been successfully cancelled</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>

                <td>
        
                </td>
                </tr>
                </table>
           </td>
     </tr>
<%



End Select                
                 %>
                     </table>
                

         
            </td>
        </tr>
        

        <input type="hidden" value="1" name="Timesthrough" />
          <tr>
            <td align="center" colspan="2">
                <%If Errormessage>"" then%>   
                <%'If Errormessage="dontshowthisanymore" then%>            
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
          <!--        
         <tr>
            <td align="left" valign="top">
            &nbsp;
           </td>
           <td align="right">
                  <input type="image" src="../images/submit_fleetexpress2.gif" alt="submit order" /> 
            </td></tr>
            -->   
        <input type="hidden" name="ColorSelect" value="<%=ColorSelect %>" />
        <input type="hidden" name="MarkTemp" value="<%=MarkTemp %>" />
        <input type="hidden" name="CaptchaSubmit" value="<%=CaptchaSubmit %>" />
        <input type="hidden" name="varCaptcha" value="<%=varCaptcha %>" />
        <input type="hidden" name="XSquare" value="<%=XSquare %>" />
        <input type="hidden" name="UserID" value="<%=UserID %>" />
        
        
    </table>
</form>
<%
'end if

'Response.write "OrderStatus="&OrderStatus&"<BR>"
%> 
    
    
    
    
    
    
    </td></tr>



 
	<tr height="50">
		<td>&nbsp;</td>
	</tr>


</table>
</td></tr>
<%
'if ErrorMessage>"" then
If ErrorMessage="dontshowthisanymore" then%>
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
</form>
<tr><td height="90%">&nbsp;</td></tr>
<tr>
    <td height="100" class="FleetXGreySection" colspan="2">
        <!-- #include file="../BottomSection.asp" -->
    </td>
</tr>
<tr><td height="15" class="FleetXRedSectionSmall" colspan="2" align="center"><%=CopywriteNotice %></td></tr>
</td></tr></table>
<%
'Response.write "BillToID="&BillToID&"<BR>"
'Response.write "ShipmentType="&ShipmentType&"<BR>" 
%>


<script src="../jquery-2.1.0.min.js"></script> 
<script src="../pickadate.js"></script> 
<script type="text/javascript">
    // PICKADATE FORMATTING
    $('#date_1').pickadate({
        format: 'mm/dd/yyyy', 	// Friendly format displayed to user
        formatSubmit: 'mm/dd/yyyy', // Actual format used by application
        hiddenName: false			// Allows two different formats
    });


    $('#time_1').pickatime({
        format: 'h:i A', 		// Displayed and application format
        interval: 10, 			// Interval between values (in minutes)
        min: '12:00 AM', 			// Starting value
        max: '11:59 PM'				// Ending value
    });

</script>
</body>
</html>
