<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title> Fleet Express Transportation Request</title>
    <link rel="stylesheet" type="text/css" href="../css/Style.css">
   <%
   '''''''''''HARDCODED STUFF
   sBT_ID=Request.QueryString("BTID")
   If trim(sBT_ID)="" then
        sBT_ID=Request.Form("BTID")
    End if
   'sBT_ID="84"
   'If sBT_ID>"" then
        sBT_ID="86"
        'BillToID=sBT_ID
   'End if
    UserID=Request.Form("UserID")
    If trim(userID)="" then
        UserID=Session("UserID")
    End if
    If trim(UserID)>"" then
        Session("UserID")=UserID
    End if
   'Response.write "UserID="&UserID&"<BR>"
   'Response.write "sbt_id="&Session("sBT_ID")&"<BR>"
   If sBT_ID<>"86" or Trim(UserID)="" then
        'Response.write "TRIED TO REDIRECT YOU!!!!<BR>"
        'Response.redirect("/intranet/default.asp")
   End if
    Comments=Request.form("Comments")
    AddedComments=Request.form("AddedComments")
    AddedComments=Replace(AddedComments, """", "`")
    AddedComments=Replace(AddedComments, "'", "`")   
    If trim(AddedComments)>"" then
         NewComments=Comments&"<br>"&AddedComments
        else
        NewComments=Comments
    End if

   'If trim(UserID)="" then response.redirect("../../default.asp") end if
   findJob=request.form("findjob")
   if trim(findjob)="" then
        FindJob=Request.QueryString("findjob")
    End if
  'Response.write "sBT_ID="&sBT_ID&"<BR>"
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
   ''''''''''''''''''''''''''
    %>
    <!-- #include file="../fleetexpress.inc" -->
    <!-- #include file="../include/checkstring.inc" -->
<script language="javascript" type="text/javascript" src="datetimepicker.js">

    //Date Time Picker script- by TengYong Ng of http://www.rainforestnet.com
    //Script featured on JavaScript Kit (http://www.javascriptkit.com)
    //For this script, visit http://www.javascriptkit.com 

</script>
    <script type="text/javascript">
        function onPageLoad() {
            if (document.form.RequestorName.value.length == 0) {
                document.form.RequestorName.focus();
            }
        }  
    </script>
    <%
  
    TableWidth="460"
    Internal=Request.Querystring("Internal")
    PageStatus=Request.form("PageStatus")
    If trim(PageStatus)="" then
        PageStatus=Request.QueryString("PageStatus")
    end if
    BillToID=sBT_ID
    'PageStatus=Request.Querystring("pid")
    JobNum=Request.Querystring("SearchJobNumber")
    If trim(JobNum)="" then
        JobNum=Request.form("SearchJobNumber")
    End if
    Submit=Request.form("Submit")
    ReferenceNumber=Request.form("ReferenceNumber")
    Courier=Request.form("Courier")
    CurrentDate=Now()
    showtime=Request.QueryString("ShowTime")

'Response.write "findjob="&findjob&"<br>"
'Response.write "JobNum="&JobNum&"<br>"
'Response.write "pagestatus="&pagestatus&"<br>"
'Response.write "sbt_id="&sbt_id&"<br>"

If findjob="y" then
	Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
		RSEVENTS.CursorLocation = 3
		RSEVENTS.CursorType = 3
        'Response.write "Database="&Database&"<br>"
		RSEVENTS.ActiveConnection = DATABASE
		'SQL = "SELECT fh_status, fh_ship_dt, fh_ready, fh_co_id, fh_co_phone, fh_co_email, fh_co_costcenter, fh_custpo, fh_priority FROM fcfgthd where (fh_id = '"& jobnum &"') AND (fh_bt_id='"& BillToID &"') and (fh_status='SCD')"
        SQL = "SELECT fh_status, fh_ship_dt, fh_ready, fh_co_id, fh_co_phone, fh_co_email, fh_co_costcenter, fh_carr_id, fh_custpo, fh_priority, fh_ref FROM fcfgthd where (fh_id = '"& jobnum &"') AND (fh_bt_id='"& BillToID &"')"
		'Response.Write "SQL="&SQL&"<BR>"
		RSEVENTS.Open SQL, DATABASE, 1, 3
       'if RSEVENTS.eof then
       '    jobnum=0
       'End if
        fh_status=RSEVENTS("fh_status")
        fh_ship_dt=RSEVENTS("fh_ship_dt")
        PickUpDateTime=RSEVENTS("fh_ready")
        RequestorName=RSEVENTS("fh_co_id")
        Requestorphonenumber=RSEVENTS("fh_co_phone")
        Requestoremailaddress=RSEVENTS("fh_co_email")
        costcenternumber=RSEVENTS("fh_co_costcenter")
        fh_carr_id=RSEVENTS("fh_carr_id")
        ponumber=RSEVENTS("fh_custpo")
        priority=RSEVENTS("fh_priority")
        fh_ref=RSEVENTS("fh_ref")
        'response.write "fh_ref="&fh_ref&"<BR>"
		RSEVENTS.close
	Set RSEVENTS = Nothing
	Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
		RSEVENTS.CursorLocation = 3
		RSEVENTS.CursorType = 3
		RSEVENTS.ActiveConnection = DATABASE
		SQL = "SELECT fl_sf_name, fl_sf_clname, fl_sf_phone, fl_sf_email, fl_sf_addr1, fl_sf_city, fl_sf_state, fl_sf_country, fl_sf_zip, fl_sf_comment, fl_st_rta, fl_st_name, fl_st_clname, fl_st_phone, fl_st_email, fl_st_addr1, fl_st_city, fl_st_state, fl_st_country, fl_st_zip, fl_t_release, fl_t_atd   FROM fclegs where (fl_fh_id = '"& jobnum &"')"
		'Response.Write "SQL="&SQL&"<BR>"
		RSEVENTS.Open SQL, DATABASE, 1, 3
        OriginationCompany=RSEVENTS("fl_sf_name")
        OriginationContactname=RSEVENTS("fl_sf_clname")
        OriginationphoneNumber=RSEVENTS("fl_sf_phone")
        Originationemail=RSEVENTS("fl_sf_email")
        Originationaddress=RSEVENTS("fl_sf_addr1")
        Originationcity=RSEVENTS("fl_sf_city")
        Originationstate=RSEVENTS("fl_sf_state")
        Originationcountry=RSEVENTS("fl_sf_country")
        Originationzip=RSEVENTS("fl_sf_zip")
        comments=RSEVENTS("fl_sf_comment")
        DeliveryDateTime=RSEVENTS("fl_st_rta")
        DestinationCompany=RSEVENTS("fl_st_name")
        DestinationContactname=RSEVENTS("fl_st_clname")
        DestinationphoneNumber=RSEVENTS("fl_st_phone")
        Destinationemail=RSEVENTS("fl_st_email")
        Destinationaddress=RSEVENTS("fl_st_addr1")
        Destinationcity=RSEVENTS("fl_st_city")
        Destinationstate=RSEVENTS("fl_st_state")
        Destinationcountry=RSEVENTS("fl_st_country")
        Destinationzipcode=RSEVENTS("fl_st_zip")
        fl_t_release=RSEVENTS("fl_t_release")
        fl_t_atd=RSEVENTS("fl_t_atd")
		RSEVENTS.close
	Set RSEVENTS = Nothing
	Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
		RSEVENTS.CursorLocation = 3
		RSEVENTS.CursorType = 3
		RSEVENTS.ActiveConnection = DATABASE
		SQL = "SELECT rf_box, POD, NumberOfPieces, IsPalletized, Weight, DimLength, DimWidth, DimHeight, Hazmat, Refrigerate FROM fcrefs where (rf_fh_id = '"& jobnum &"')"
		'Response.Write "SQL="&SQL&"<BR>"
		RSEVENTS.Open SQL, DATABASE, 1, 3
        rf_box=RSEVENTS("rf_box")
        POD=RSEVENTS("POD")
        Pieces=RSEVENTS("NumberOfPieces")
        IsPalletized=RSEVENTS("IsPalletized")
        DimWeight=RSEVENTS("Weight")
        DimLength=RSEVENTS("DimLength")
        DimWidth=RSEVENTS("DimWidth")
        DimHeight=RSEVENTS("DimHeight")
        Hazmat=RSEVENTS("Hazmat")
        Refrigerate=RSEVENTS("Refrigerate")
		RSEVENTS.close
	Set RSEVENTS = Nothing

    Select Case POorNWA
        Case " Fleet Express P/O #"
            PONumber=GenericNumber
        Case "NWA or Cost Center #"
            CostCenterNumber=GenericNumber
    End Select




    HighlightedField="RequestorName"
    CurrentDateTime=Now()
 

    ''''''''END ERROR HANDLING''''''
End if
''''''''''''HERE'S THE MOVED UPDATES'''''''''''''
   ''''''''ERROR HANDLING''''''''''
    If trim(submit)="Cancel Order" then
                         Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open INTRANET
						' 7/26/04 KK: Added canceljob functionality to be able to update the status if cancel button is pressed.
						' 11/30/04 DEC: Changed from CAN/98 TO DEL/99 to be consistent with dispatchOffice			
							l_cSQL = "UPDATE FleetRouted SET status = 'x' WHERE fh_id = '" & JobNum & "'"
							'Response.write "l_cSQL="&l_cSQL&"<BR>"
							oConn.Execute(l_cSQL)
						Set oConn=Nothing



						Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open DATABASE
						' 7/26/04 KK: Added canceljob functionality to be able to update the status if cancel button is pressed.
						' 11/30/04 DEC: Changed from CAN/98 TO DEL/99 to be consistent with dispatchOffice			
							l_cSQL = "UPDATE fcfgthd SET fh_status = 'CAN', fh_statcode = 98, fh_dispatcher = '"&UserID&"' WHERE fh_id = '" & JobNum & "'"
							'Response.write "l_cSQL="&l_cSQL&"<BR>"
							oConn.Execute(l_cSQL)
						Set oConn=Nothing
						Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open DATABASE
						' 7/26/04 KK: Added canceljob functionality to be able to update the status if cancel button is pressed.
						' 11/30/04 DEC: Changed from CAN/98 TO DEL/99 to be consistent with dispatchOffice			
							l_cSQL = "UPDATE fclegs SET fl_t_release = '"& CurrentDate &"' WHERE fl_fh_id = '" & JobNum & "'"
							'Response.write "l_cSQL="&l_cSQL&"<BR>"
							oConn.Execute(l_cSQL)
						Set oConn=Nothing
''''''''''''''''''''''EMAIL MESSAGE''''''''''''''''
				    Body = "Your order, #"& jobnum &", has been CANCELLED.<br><br>"   

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
  
                    
                    Body = Body & "Hazmat: "&  IsHazmat &"<br>"
                    Body = Body & "Refrigerate: "&  Refrigerate &"<br><br>"
                    Body = Body & "ORIGINATION:<BR>"   
                    Body = Body & "Company: "&  OriginationCompany &"<br>"   
                    Body = Body & "Address: "&  OriginationAddress &"<br>"   
                    Body = Body & "City: "&  OriginationCity &"<br>"   
                    Body = Body & "State: "&  OriginationState &"<br>"  
                    Body = Body & "Zip Code: "&  OriginationZipCode &"<br>"   
                    Body = Body & "Contact Name: "&  OriginationContactName &"<br>"   
                    Body = Body & "Phone Number: "&  OriginationPhoneNumber &"<br>"   
                    Body = Body & "Email: "&  OriginationEmail &"<br>" 
                    Body = Body & "Pick Up Date/Time: "&  PickUpDateTime &"<br><br>" 
                    Body = Body & "DESTINATION:<BR>"  
                    Body = Body & "Company: "&  DestinationCompany &"<br>"  
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

                    'Body = Body & "<a href='http://www.logisticorp.us/intranet/dedicatedfleets/orderentry/FleetExpressOrderConfirmation.asp?bid=84&pid=manage&jid="& newjobnum &"'>To Approve or Refuse this request, click here</a><br><br>" 


				    Body = Body & "Thank you,<br><br>"  
				    Body = Body & "FleetX<br>"  
				    Body = Body &  "FleetX@LogisticorpGroup.com<br>"  
				    Body = Body & "972/499-3415<br><br>"
				    'Recipient=FirstName&" "&LastName
			        SentToEmail=RequestorEmailAddress
				    'Email="KWETI.Mailbox@am.kwe.com"
				    'Email="mark@maggiore.net"
				    Set objMail = CreateObject("CDONTS.Newmail")
				    objMail.From = "FleetX@LogisticorpGroup.com"
				    objMail.To = SentToEmail
                    objMail.cc = "mark.maggiore@logisticorp.us"
				    'objMail.cc = "4692269939@tmomail.net;htmlmale@yahoo.com"
				    'objMail.cc = "4692269939@tmomail.net;pcurrin@ti.com"
				    objMail.Subject = "Cancelled FleetX Shipment Request"
				    objMail.MailFormat = cdoMailFormatMIME
				    objMail.BodyFormat = cdoBodyFormatHTML
				    objMail.Body = Body
				    objMail.Send
				    Set objMail = Nothing
                        SuccessMessage="You have Cancelled order #"& JobNum 
    End if
    If trim(submit)="Dispatch Order" then
            If trim(ReferenceNumber)="" then errormessage="You must provide a reference number" end if
                    If trim(errormessage)="" then


	                    Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
		                    RSEVENTS.CursorLocation = 3
		                    RSEVENTS.CursorType = 3
                            'Response.write "Database="&Database&"<br>"
		                    RSEVENTS.ActiveConnection = DATABASE
		                    'SQL = "SELECT fh_status, fh_ship_dt, fh_ready, fh_co_id, fh_co_phone, fh_co_email, fh_co_costcenter, fh_custpo, fh_priority FROM fcfgthd where (fh_id = '"& jobnum &"') AND (fh_bt_id='"& BillToID &"') and (fh_status='SCD')"
                            SQL = "SELECT un_id, un_dr_id FROM fcunits where (un_desc = '"& courier &"')"
		                    'Response.Write "SQL="&SQL&"<BR>"
                            'Response.Write "DATABASE="&DATABASE&"<BR>"
		                    RSEVENTS.Open SQL, DATABASE, 1, 3
                           'if RSEVENTS.eof then
                           '    jobnum=0
                           'End if
                            un_id=RSEVENTS("un_id")
                            un_dr_id=RSEVENTS("un_dr_id")
                            'response.write "fh_ref="&fh_ref&"<BR>"
		                    RSEVENTS.close
	                    Set RSEVENTS = Nothing




                    DeliveryPeriod=DateDiff("h", PickUpDateTime, DeliveryDateTime)
                    'Response.write "DeliveryPeriod="&DeliveryPeriod&"<br>"
                         Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open INTRANET
						' 7/26/04 KK: Added canceljob functionality to be able to update the status if cancel button is pressed.
						' 11/30/04 DEC: Changed from CAN/98 TO DEL/99 to be consistent with dispatchOffice			
							l_cSQL = "UPDATE FleetRouted SET status = 'x' WHERE fh_id = '" & JobNum & "'"
							'Response.write "l_cSQL="&l_cSQL&"<BR>"
							oConn.Execute(l_cSQL)
						Set oConn=Nothing
  			            Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				            RSEVENTS2.Open "FleetRouted", Intranet, 2, 2
				            RSEVENTS2.addnew
                            RSEVENTS2("fh_id")=JobNum
                            RSEVENTS2("PickDrop")="p"
                            RSEVENTS2("Courier")=Courier
                            RSEVENTS2("Location")=OriginationCompany
                            RSEVENTS2("ArrivalTime")=PickUpDateTime
                            RSEVENTS2("Pieces")=Pieces
                            RSEVENTS2("BTID")=Session("sBT_ID")
                            RSEVENTS2("DeliveryPeriod")=DeliveryPeriod
                            RSEVENTS2("Status")="c"
				            RSEVENTS2.update
				            RSEVENTS2.close			
			            set RSEVENTS2 = nothing                                       
                        
  			            Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				            RSEVENTS2.Open "FleetRouted", Intranet, 2, 2
				            RSEVENTS2.addnew
                            RSEVENTS2("fh_id")=JobNum
                            RSEVENTS2("PickDrop")="d"
                            RSEVENTS2("Courier")=Courier
                            RSEVENTS2("Location")=DestinationCompany
                            RSEVENTS2("ArrivalTime")=DeliveryDateTime
                            RSEVENTS2("Pieces")=Pieces
                            RSEVENTS2("BTID")=Session("sBT_ID")
                            RSEVENTS2("DeliveryPeriod")=DeliveryPeriod
                            RSEVENTS2("Status")="c"
				            RSEVENTS2.update
				            RSEVENTS2.close			
			            set RSEVENTS2 = nothing  



						Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open DATABASE
						' 7/26/04 KK: Added canceljob functionality to be able to update the status if cancel button is pressed.
						' 11/30/04 DEC: Changed from CAN/98 TO DEL/99 to be consistent with dispatchOffice			
							l_cSQL = "UPDATE fcfgthd SET fh_status = 'OPN', fh_statcode = 3, fh_dispatcher = '"&UserID&"', fh_ref='"& ReferenceNumber &"', fh_carr_ID='"& Courier &"' WHERE fh_id = '" & JobNum & "'"
							
                            'Response.write "l_cSQL="&l_cSQL&"<BR>"
							
                            oConn.Execute(l_cSQL)
						Set oConn=Nothing
						Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open DATABASE
						' 7/26/04 KK: Added canceljob functionality to be able to update the status if cancel button is pressed.
						' 11/30/04 DEC: Changed from CAN/98 TO DEL/99 to be consistent with dispatchOffice	
                        
                        
                        
		
							l_cSQL = "UPDATE fclegs SET fl_un_id= '"& un_id &"', fl_dr_id='"& un_dr_id &"', fl_t_disp = '"& CurrentDate &"', fl_sf_comment = '"& NewComments &"' WHERE fl_fh_id = '" & JobNum & "'"
                  
                        
                        
                        		
							'l_cSQL = "UPDATE fclegs SET fl_t_disp = '"& CurrentDate &"', fl_sf_comment = '"& NewComments &"' WHERE fl_fh_id = '" & JobNum & "'"
							'Response.write "l_cSQL="&l_cSQL&"<BR>"
							oConn.Execute(l_cSQL)
						Set oConn=Nothing
''''''''''''''''''''''EMAIL MESSAGE''''''''''''''''
				    Body = "Your order, #"& jobnum &", has been DISPATCHED!<br><br>" 
                     Body = Body & "COURIER INFORMATION:<BR>"  
                    Body = Body & "Courier: "&  Courier &"<br>"  
                    Body = Body & "Reference Number: "&  ReferenceNumber &"<br><br>" 

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
  
                    
                    Body = Body & "Hazmat: "&  IsHazmat &"<br>"
                    Body = Body & "Refrigerate: "&  Refrigerate &"<br><br>"
                    Body = Body & "ORIGINATION:<BR>"   
                    Body = Body & "Company: "&  OriginationCompany &"<br>"   
                    Body = Body & "Address: "&  OriginationAddress &"<br>"   
                    Body = Body & "City: "&  OriginationCity &"<br>"   
                    Body = Body & "State: "&  OriginationState &"<br>"  
                    Body = Body & "Zip Code: "&  OriginationZipCode &"<br>"   
                    Body = Body & "Contact Name: "&  OriginationContactName &"<br>"   
                    Body = Body & "Phone Number: "&  OriginationPhoneNumber &"<br>"   
                    Body = Body & "Email: "&  OriginationEmail &"<br>" 
                    Body = Body & "Pick Up Date/Time: "&  PickUpDateTime &"<br><br>" 
                    Body = Body & "DESTINATION:<BR>"  
                    Body = Body & "Company: "&  DestinationCompany &"<br>"  
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
                        Body = Body & ""&  newcomments &"<br><br>" 
                    End if 
                    'Body = Body & "<a href='http://www.logisticorp.us/intranet/dedicatedfleets/orderentry/FleetExpressOrderConfirmation.asp?bid=84&pid=manage&jid="& newjobnum &"'>To Approve or Refuse this request, click here</a><br><br>" 


				    Body = Body & "Thank you,<br><br>"  
				    Body = Body & "FleetX<br>"  
				    Body = Body &  "FleetX@LogisticorpGroup.com<br>"  
				    Body = Body & "972/499-3415<br><br>"
				    'Recipient=FirstName&" "&LastName
			        If lcase(RequestorEmailAddress)<>"fleetexpress@logisticorp.us"  AND lcase(RequestorEmailAddress)<>"texasinstruments@plg.cc" then
                        SentToEmail=RequestorEmailAddress
				        'Email="KWETI.Mailbox@am.kwe.com"
				        'Email="mark@maggiore.net"
				        Set objMail = CreateObject("CDONTS.Newmail")
				        objMail.From = "FleetX@LogisticorpGroup.com"
				        objMail.To = SentToEmail
                        objMail.cc = "mark.maggiore@logisticorp.us;FleetExpress@LogistiCorp.us"
				        'objMail.cc = "4692269939@tmomail.net;htmlmale@yahoo.com"
				        'objMail.cc = "4692269939@tmomail.net;pcurrin@ti.com"
				        objMail.Subject = "Dispatched FleetX Shipment Request"
				        objMail.MailFormat = cdoMailFormatMIME
				        objMail.BodyFormat = cdoBodyFormatHTML
				        objMail.Body = Body
				        objMail.Send
                    End if
				    Set objMail = Nothing


                        SuccessMessage="You have DISPATCHED order #"& JobNum 
                    End if
    End if
    If trim(submit)="Approve Order" then
                    If trim(errormessage)="" then
						Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open DATABASE
						' 7/26/04 KK: Added canceljob functionality to be able to update the status if cancel button is pressed.
						' 11/30/04 DEC: Changed from CAN/98 TO DEL/99 to be consistent with dispatchOffice			
							l_cSQL = "UPDATE fcfgthd SET fh_status = 'RAP', fh_statcode = 2, fh_dispatcher = '"&UserID&"' WHERE fh_id = '" & JobNum & "'"
							'Response.write "l_cSQL="&l_cSQL&"<BR>"
							oConn.Execute(l_cSQL)
						Set oConn=Nothing
						Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open DATABASE
						' 7/26/04 KK: Added canceljob functionality to be able to update the status if cancel button is pressed.
						' 11/30/04 DEC: Changed from CAN/98 TO DEL/99 to be consistent with dispatchOffice			
							l_cSQL = "UPDATE fclegs SET fl_t_release = '"& CurrentDate &"' WHERE fl_fh_id = '" & JobNum & "'"
							'Response.write "l_cSQL="&l_cSQL&"<BR>"
							oConn.Execute(l_cSQL)
						Set oConn=Nothing
''''''''''''''''''''''EMAIL MESSAGES''''''''''''''''
''''''''''''''''1. To REQUESTOR
				    Body = "Your order, #"& jobnum &", has been APPROVED!<br><br>"   

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
  
                    
                    Body = Body & "Hazmat: "&  IsHazmat &"<br>"
                    Body = Body & "Refrigerate: "&  Refrigerate &"<br><br>"
                    Body = Body & "ORIGINATION:<BR>"   
                    Body = Body & "Company: "&  OriginationCompany &"<br>"   
                    Body = Body & "Address: "&  OriginationAddress &"<br>"   
                    Body = Body & "City: "&  OriginationCity &"<br>"   
                    Body = Body & "State: "&  OriginationState &"<br>"  
                    Body = Body & "Zip Code: "&  OriginationZipCode &"<br>"   
                    Body = Body & "Contact Name: "&  OriginationContactName &"<br>"   
                    Body = Body & "Phone Number: "&  OriginationPhoneNumber &"<br>"   
                    Body = Body & "Email: "&  OriginationEmail &"<br>" 
                    Body = Body & "Pick Up Date/Time: "&  PickUpDateTime &"<br><br>" 
                    Body = Body & "DESTINATION:<BR>"  
                    Body = Body & "Company: "&  DestinationCompany &"<br>"  
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
			        SentToEmail=RequestorEmailAddress
				    'Email="KWETI.Mailbox@am.kwe.com"
				    'Email="mark@maggiore.net"
				    Set objMail = CreateObject("CDONTS.Newmail")
				    objMail.From = "FleetX@LogisticorpGroup.com"
				    objMail.To = SentToEmail
				    objMail.cc = "mark.maggiore@logisticorp.us"
				    'objMail.cc = "4692269939@tmomail.net;pcurrin@ti.com"
				    objMail.Subject = "Approved FleetX Shipment Request"
				    objMail.MailFormat = cdoMailFormatMIME
				    objMail.BodyFormat = cdoBodyFormatHTML
				    objMail.Body = Body
				    objMail.Send
				    Set objMail = Nothing

''''''''''''''''2. To LOGISTICORP
				    Body = "An order, #"& jobnum &", has been APPROVED!<br><br>"   

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
  
                    
                    Body = Body & "Hazmat: "&  IsHazmat &"<br>"
                    Body = Body & "Refrigerate: "&  Refrigerate &"<br><br>"
                    Body = Body & "ORIGINATION:<BR>"   
                    Body = Body & "Company: "&  OriginationCompany &"<br>"   
                    Body = Body & "Address: "&  OriginationAddress &"<br>"   
                    Body = Body & "City: "&  OriginationCity &"<br>"   
                    Body = Body & "State: "&  OriginationState &"<br>"  
                    Body = Body & "Zip Code: "&  OriginationZipCode &"<br>"   
                    Body = Body & "Contact Name: "&  OriginationContactName &"<br>"   
                    Body = Body & "Phone Number: "&  OriginationPhoneNumber &"<br>"   
                    Body = Body & "Email: "&  OriginationEmail &"<br>" 
                    Body = Body & "Pick Up Date/Time: "&  PickUpDateTime &"<br><br>" 
                    Body = Body & "DESTINATION:<BR>"  
                    Body = Body & "Company: "&  DestinationCompany &"<br>"  
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
                    Body = Body & "<a href='http://www.logisticorp.us/intranet/dedicatedfleets/orderentry/FleetExpressOrderConfirmation.asp?bid=84&pid=disp&jid="& jobnum &"'>To dispatch this request, click here</a><br><br>" 


				    Body = Body & "Thank you,<br><br>"  
				    Body = Body & "FleetX Services<br>"  
				    Body = Body &  "FleetX@LogisticorpGroup.com<br>"  
				    Body = Body & "972/499-3415<br><br>"
				    'Recipient=FirstName&" "&LastName
			        SentToEmail="xxx@LogistiCorp.us"
				    'Email="KWETI.Mailbox@am.kwe.com"
				    'Email="mark@maggiore.net"
				    Set objMail = CreateObject("CDONTS.Newmail")
				    objMail.From = "FleetX@LogisticorpGroup.com"
				    objMail.To = SentToEmail
				    objMail.cc = "mark.maggiore@logisticorp.us"
				    'objMail.cc = "4692269939@tmomail.net;pcurrin@ti.com"
				    objMail.Subject = "Approved FleetX Shipment Request"
				    objMail.MailFormat = cdoMailFormatMIME
				    objMail.BodyFormat = cdoBodyFormatHTML
				    objMail.Body = Body
				    objMail.Send
				    Set objMail = Nothing


                        SuccessMessage="You have APPROVED order #"& JobNum 
                    End if
    End if

''''''''''''END MOVED UPDATES''''''''''''''''''''







If findjob="y" then
	Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
		RSEVENTS.CursorLocation = 3
		RSEVENTS.CursorType = 3
        'Response.write "Database="&Database&"<br>"
		RSEVENTS.ActiveConnection = DATABASE
		'SQL = "SELECT fh_status, fh_ship_dt, fh_ready, fh_co_id, fh_co_phone, fh_co_email, fh_co_costcenter, fh_custpo, fh_priority FROM fcfgthd where (fh_id = '"& jobnum &"') AND (fh_bt_id='"& BillToID &"') and (fh_status='SCD')"
        SQL = "SELECT fh_status, fh_ship_dt, fh_ready, fh_co_id, fh_co_phone, fh_co_email, fh_co_costcenter, fh_carr_id, fh_custpo, fh_priority, fh_ref FROM fcfgthd where (fh_id = '"& jobnum &"') AND (fh_bt_id='"& BillToID &"')"
		'Response.Write "SQL="&SQL&"<BR>"
		RSEVENTS.Open SQL, DATABASE, 1, 3
       'if RSEVENTS.eof then
       '    jobnum=0
       'End if
        fh_status=RSEVENTS("fh_status")
        fh_ship_dt=RSEVENTS("fh_ship_dt")
        PickUpDateTime=RSEVENTS("fh_ready")
        RequestorName=RSEVENTS("fh_co_id")
        Requestorphonenumber=RSEVENTS("fh_co_phone")
        Requestoremailaddress=RSEVENTS("fh_co_email")
        costcenternumber=RSEVENTS("fh_co_costcenter")
        fh_carr_id=RSEVENTS("fh_carr_id")
        ponumber=RSEVENTS("fh_custpo")
        priority=RSEVENTS("fh_priority")
        fh_ref=RSEVENTS("fh_ref")
        'response.write "fh_ref="&fh_ref&"<BR>"
		RSEVENTS.close
	Set RSEVENTS = Nothing
	Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
		RSEVENTS.CursorLocation = 3
		RSEVENTS.CursorType = 3
		RSEVENTS.ActiveConnection = DATABASE
		SQL = "SELECT fl_sf_name, fl_sf_clname, fl_sf_phone, fl_sf_email, fl_sf_addr1, fl_sf_city, fl_sf_state, fl_sf_country, fl_sf_zip, fl_sf_comment, fl_st_rta, fl_st_name, fl_st_clname, fl_st_phone, fl_st_email, fl_st_addr1, fl_st_city, fl_st_state, fl_st_country, fl_st_zip, fl_t_release, fl_t_atd   FROM fclegs where (fl_fh_id = '"& jobnum &"')"
		'Response.Write "SQL="&SQL&"<BR>"
		RSEVENTS.Open SQL, DATABASE, 1, 3
        OriginationCompany=RSEVENTS("fl_sf_name")
        OriginationContactname=RSEVENTS("fl_sf_clname")
        OriginationphoneNumber=RSEVENTS("fl_sf_phone")
        Originationemail=RSEVENTS("fl_sf_email")
        Originationaddress=RSEVENTS("fl_sf_addr1")
        Originationcity=RSEVENTS("fl_sf_city")
        Originationstate=RSEVENTS("fl_sf_state")
        Originationcountry=RSEVENTS("fl_sf_country")
        Originationzip=RSEVENTS("fl_sf_zip")
        comments=RSEVENTS("fl_sf_comment")
        DeliveryDateTime=RSEVENTS("fl_st_rta")
        DestinationCompany=RSEVENTS("fl_st_name")
        DestinationContactname=RSEVENTS("fl_st_clname")
        DestinationphoneNumber=RSEVENTS("fl_st_phone")
        Destinationemail=RSEVENTS("fl_st_email")
        Destinationaddress=RSEVENTS("fl_st_addr1")
        Destinationcity=RSEVENTS("fl_st_city")
        Destinationstate=RSEVENTS("fl_st_state")
        Destinationcountry=RSEVENTS("fl_st_country")
        Destinationzipcode=RSEVENTS("fl_st_zip")
        fl_t_release=RSEVENTS("fl_t_release")
        fl_t_atd=RSEVENTS("fl_t_atd")
		RSEVENTS.close
	Set RSEVENTS = Nothing
	Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
		RSEVENTS.CursorLocation = 3
		RSEVENTS.CursorType = 3
		RSEVENTS.ActiveConnection = DATABASE
		SQL = "SELECT rf_box, POD, NumberOfPieces, IsPalletized, Weight, DimLength, DimWidth, DimHeight, Hazmat, Refrigerate FROM fcrefs where (rf_fh_id = '"& jobnum &"')"
		'Response.Write "SQL="&SQL&"<BR>"
		RSEVENTS.Open SQL, DATABASE, 1, 3
        rf_box=RSEVENTS("rf_box")
        POD=RSEVENTS("POD")
        Pieces=RSEVENTS("NumberOfPieces")
        IsPalletized=RSEVENTS("IsPalletized")
        DimWeight=RSEVENTS("Weight")
        DimLength=RSEVENTS("DimLength")
        DimWidth=RSEVENTS("DimWidth")
        DimHeight=RSEVENTS("DimHeight")
        Hazmat=RSEVENTS("Hazmat")
        Refrigerate=RSEVENTS("Refrigerate")
		RSEVENTS.close
	Set RSEVENTS = Nothing

    Select Case POorNWA
        Case " Fleet Express P/O #"
            PONumber=GenericNumber
        Case "NWA or Cost Center #"
            CostCenterNumber=GenericNumber
    End Select




    HighlightedField="RequestorName"
    CurrentDateTime=Now()
 

    ''''''''END ERROR HANDLING''''''
End if
     %>
</head>
<BODY leftMargin="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" OnLoad=document.OrderForm1.<%=HighlightedField%>.focus()>
<!--form method="post" name="OrderForm1"-->
    <table border="0" cellpadding="0" cellspacing="0" align="center">
        <tr><td align="left"><img src="images/pixel.gif" height="10" width="1" /></td></tr>
        <tr>
            <td align="left"><img src="images/logo_FleetExpress_space.gif" height="87" width="100" /></td>
            <td align="right" valign="bottom"><a href="mailto:mark.maggiore@logisticorp.us" class="<%=LinkClass%>">Click here to report a problem with this page</a></td>
        </tr>
        <tr><td align="left" colspan="2"><img src="images/pixel.gif" /></td></tr>
        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2">Dispatch Order #<%=JobNum %></td></tr>
        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2"><img src="images/pixel.gif" height="5" width="1" /></td></tr>
        <!--tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressBodyWhite" colspan="2">Please complete all areas below</td></tr-->
        <tr><td align="left" colspan="2"><img src="images/pixel.gif" height="10" width="1" /></td></tr>
 <%If findjob="y" then %>      
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
                <%if fh_ship_dt>"" then %>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap>Booked</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left" class="FleetExpressTextBlack"><%=fh_ship_dt%></td>
                </tr>
                <%end if %>
                <%if fl_t_release>"" then %>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap>Accepted/Cancelled</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left" class="FleetExpressTextBlack"><%=fl_t_release%></td>
                </tr>
                <%end if %>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap>Requestor Name</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left" class="FleetExpressTextBlack"><%=RequestorName%></td>
                </tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap>Phone Number</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left" class="FleetExpressTextBlack"><%=RequestorPhoneNumber%></td>
                </tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap>Email Address</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left" class="FleetExpressTextBlack"><%=RequestorEmailAddress%></td>
                </tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap>
                               <%If trim(CostCenterNumber)>"" then 
                               Response.write "NWA or Cost Center #" 
                               GenericNumber=CostCenterNumber
                               end if%>
                               <%If trim(PONumber)>"" then 
                               Response.write " Fleet Express P/O #" 
                               GenericNumber=PONumber
                               end if%>
                    </td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left" class="FleetExpressTextBlack"><%=GenericNumber%></td>
                </tr>
                <%
                If trim(NewComments)>"" then Comments=NewComments End if
                 %>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" valign="top">Special Instructions</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left" class="FleetExpressTextBlack"><%=Comments%></td>
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
                <%if fh_carr_id>"" then %>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap>Routed to</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left" class="FleetExpressTextBlack"><%=fh_carr_id%></td>
                </tr>
                <%end if %>
                <%if fh_ref>"" then %>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap>Reference Number</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left" class="FleetExpressTextBlack"><%=fh_ref%></td>
                </tr>
                <%end if %>
                <%if fl_t_atd>"" then %>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap>Delivered</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left" class="FleetExpressTextBlack"><%=fl_t_atd%></td>
                </tr>
                <%end if %>
                <%if POD>"" then %>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap>POD</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left" class="FleetExpressTextBlack"><%=POD%></td>
                </tr>
                <%end if %>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Number of Pieces</td>
                        <td width="10"><img src="images/pixel.gif" /></td>
                        <td nowrap class="FleetExpressTextBlack">
                            <%=pieces%>
                            &nbsp;&nbsp;
                            <%=rf_box%>
                            &nbsp;&nbsp;
                            <%If trim(IsPalletized)="y" then Response.write "Palletized" end if%>
                            <%If trim(IsPalletized)="n" then Response.write "Not Palletized" end if%>
                            <%If trim(IsPalletized)="Trailer Only Move" then Response.write "Trailer Only Move" end if%>
                            
                        </td>
                    </tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Weight</td>
                        <td width="10"><img src="images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=DimWeight%> Pounds</td>
                    </tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Dimensions</td>
                        <td width="10"><img src="images/pixel.gif" /></td>
                        <td class="FleetExpressTextBlack" align="left" nowrap>
                            L:&nbsp;&nbsp;<%=DimLength%>
                            W:&nbsp;&nbsp;<%=DimWidth%> 
                            H:&nbsp;&nbsp;<%=DimHeight%>
                            &nbsp;&nbsp;Inches
                        </td>
                    </tr>


                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Hazmat</td>
                        <td width="10"><img src="images/pixel.gif" /></td>
                        <td align="left"  class="FleetExpressTextBlack">
                             <%=Hazmat %>              
                        </td>
                    </tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Refrigerate</td>
                        <td width="10"><img src="images/pixel.gif" /></td>
                        <td align="left"  class="FleetExpressTextBlack">
                             <%=Refrigerate %>                    
                        </td>
                    </tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Service Level</td>
                        <td width="10"><img src="images/pixel.gif" /></td>
                        <td align="left"  class="FleetExpressTextBlack">
                             <%=Priority %>                      
                        </td>
                    </tr>
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
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap>Company Name</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left" class="FleetExpressTextBlack"><%=OriginationCompany%></td>
                </tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">Address</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left" class="FleetExpressTextBlack"><%=OriginationAddress%></td>
                </tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">City/State/Zip</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td class="FleetExpressTextBlack">
                        <%=OriginationCity%>, <%=OriginationState %>&nbsp;&nbsp;
                        <%=OriginationZipCode%>

                    </td>
                </tr>


                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">Contact Name</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left" class="FleetExpressTextBlack"><%=OriginationContactName%></td>
                </tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">Phone Number</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left" class="FleetExpressTextBlack"><%=OriginationPhoneNumber%></td>
                </tr> 
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">Email Address</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left" class="FleetExpressTextBlack"><%=OriginationEmail%></td>
                </tr> 
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">Pick Up Date/Time</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left" class="FleetExpressTextBlack"><%=PickUpDateTime%>
                    </td>
                </tr>
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
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left" nowrap>Company Name</td>
                        <td width="10"><img src="images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=DESTINATIONCompany%></td>
                    </tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Address</td>
                        <td width="10"><img src="images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=DESTINATIONAddress%></td>
                    </tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">City/State/Zip</td>
                        <td width="10"><img src="images/pixel.gif" /></td>
                        <td class="FleetExpressTextBlack">
                            <%=DESTINATIONCity%>, <%=DestinationState %>&nbsp;&nbsp; <%=DESTINATIONZipCode%>
                        </td>
                    </tr>


                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Contact Name</td>
                        <td width="10"><img src="images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=DESTINATIONContactName%></td>
                    </tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Phone Number</td>
                        <td width="10"><img src="images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=DESTINATIONPhoneNumber%></td>
                    </tr> 
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Email Address</td>
                        <td width="10"><img src="images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=DESTINATIONEmail%></td>
                    </tr> 
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left">Delivery Date/Time</td>
                    <td width="10"><img src="images/pixel.gif" /></td>
                    <td align="left" class="FleetExpressTextBlack"><%=DeliveryDateTime%>
                    </td>
                </tr>
                     </table>
                     </td></tr></table>                 
                 </td></tr>                                                                                                               
            </table>
         
            </td>
        </tr>
         <tr><td align="left"><img src="images/pixel.gif" height="10" width="1" /></td></tr>
         <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2"> Fleet Express Transportation Call Center 972-499-3415</td></tr>
         <tr><td align="left"><img src="images/pixel.gif" height="30" width="1" /></td></tr>



         <tr>
            <td colspan="3"  class="FleetExpressTextBlack" align="center">
            &nbsp;Driver:___________________ Arrival Time:___________________ Number of Pieces Picked Up:_______________ Departure Time:___________________&nbsp;
            </td>
        </tr>
        <tr><td>&nbsp;</td></tr>
         <tr>
            <td colspan="3" align="center">
                <table border="1" bordercolor="black" cellpadding="0" cellspacing="0" width="940">
                    <tr><td>
                    <table border="0" bordercolor="black" cellpadding="0" cellspacing="0" width="940">
                    <tr>
                        <td class="FleetExpressTextBlack" ><br /><img src="images/pixel.gif" height="10" width="1" /><br />
                            &nbsp;Shipper Notes/Comments:________________________________________________________________________________________________________&nbsp;<br /><img src="images/pixel.gif" height="20" width="1" /><br />
                            &nbsp;______________________________________________________________________________________________________________________________&nbsp;<br /><img src="images/pixel.gif" height="25" width="1" /><br />
                            &nbsp;Shipper Signature:__________________________________ Print Name:_____________________________________ Date:________________________&nbsp;<br /><img src="images/pixel.gif" height="20" width="1" /><br />
                        </td>
                    </tr>
                    <tr>
                        <td  class="FleetExpressTextBlackSmaller">
                            This is to certify that the above named materials are properly classified, packaged, marked, and labeled, and are in proper condition for transportation according to the applicable regulations of the DOT.
                            <br /><img src="images/pixel.gif" height="5" width="1" /><br />
                            Property described above was received by driver in good order, except as noted above.<br /><img src="images/pixel.gif" height="10" width="1" /><br />    
                        </td>
                    </tr>
                    </td></tr>
                    </table>
                    </td></tr>

                </table>
            </td>
        </tr>
        <tr><td align="left"><img src="images/pixel.gif" height="30" width="1" /></td></tr>
        <tr>
            <td colspan="3"  class="FleetExpressTextBlack" align="center">
            &nbsp;Driver:___________________ Arrival Time:___________________ Number of Pieces Delivered:_______________ Departure Time:___________________&nbsp;
            </td>
        </tr>
        <tr><td align="left"><img src="images/pixel.gif" height="30" width="1" /></td></tr>
         <tr>
            <td colspan="3" align="center">
                <table border="1" bordercolor="black" cellpadding="0" cellspacing="0" width="940">
                    <tr><td>
                    <table border="0" bordercolor="black" cellpadding="0" cellspacing="0" width="940">
                    <tr>
                        <td class="FleetExpressTextBlack" ><br /><img src="images/pixel.gif" height="10" width="1" /><br />
                            &nbsp;Consignee Notes/Comments:______________________________________________________________________________________________________&nbsp;<br /><img src="images/pixel.gif" height="20" width="1" /><br />
                            &nbsp;______________________________________________________________________________________________________________________________&nbsp;<br /><img src="images/pixel.gif" height="25" width="1" /><br />
                            &nbsp;Consignee Signature:________________________________ Print Name:_____________________________________ Date:________________________&nbsp;<br /><img src="images/pixel.gif" height="20" width="1" /><br />
                        </td>
                    </tr>
                    <tr>
                        <td  class="FleetExpressTextBlackSmaller">
                            Property described above was received by consignee in good order, except as noted above.
                            <br /><img src="images/pixel.gif" height="10" width="1" /><br />    
                        </td>
                    </tr>
                    </td></tr>
                    </table>
                    </td></tr>

                </table>
            </td>
        </tr>
        <tr><td>&nbsp;</td></tr>
        <tr><td>&nbsp;</td></tr>
                    <tr>
                        <td  class="FleetExpressTextBlackSmaller" align="right" colspan="3">
                            SWS - 9/22/2011
                            <br /><img src="images/pixel.gif" height="10" width="1" /><br />    
                        </td>
                    </tr>
        <tr><td>&nbsp;</td></tr>




         <%
         'Response.write "PID="&PID&"<BR>"
         If PageStatus="view" then
        'Response.write "internal="&Internal&"<BR>"
         %> 
          <tr><td align="center"class="FleetExpressTextBlackBold" colspan="2">Your order number is:  <%=Jobnum %>.<br /><br />If you wish to submit another request, <a href="GenericOrderPage.asp?Internal=<%=Internal%>">click here</a></td></tr>
         <%End if %>
         <%
         'Response.write "PID="&PID&"<BR>"
         If PageStatus="manage" and trim(SuccessMessage)="" and (fh_status="SCD" OR fh_status="OPN") then%> 
          <tr><td align="center"class="FleetExpressTextBlackBold" colspan="2">
          <form method="post">
              <table>
                <tr>
                    <td valign="bottom"><input type="submit" name="submit" value="Refuse Order" /></td>
                    <td width="100">&nbsp;</td>
                    <td valign="bottom" align="left">
                            <!--
                             Courier:&nbsp;&nbsp;<select name="Courier">
                                <option value="LogistiCorp 1"<%If trim(Courier)="LogistiCorp 1" then Response.write "selected" end if%>>LogistiCorp 1</option>
                                <option value="LogistiCorp 2"<%If trim(Courier)="LogistiCorp 2" then Response.write "selected" end if%>>LogistiCorp 2</option>
                                 <option value="Pronto"<%If trim(Courier)="Pronto" then Response.write "selected" end if%>>Pronto</option>
                            </select>
                            <br /><br /> 
                            Ref #:&nbsp;&nbsp;<input type="text" name="referenceNumber" size="20" maxlength="20" /><br /><br />
                            --> 
                            <input type="submit" name="submit" value="Approve Order" /></td> 
                </tr>
              </table>
              </td></tr>
          </form>
         <%End if %>
         <%
'Response.write "SuccessMessage="&SuccessMessage&"<BR>"
'Response.write "fh_status="&fh_status&"<BR>"
'Response.write "PageStatus="&PageStatus&"<BR>"
         If PageStatus="disp" and trim(SuccessMessage)="" and (fh_status="SCD" or fh_status="CAN" or fh_status="OPN" or fh_status="RAP") then

            If fh_status="CAN" then%> 
            <tr>
                <td colspan="2" align="center">
                 <table cellpadding="2" cellspacing="2" border="1" bordercolor="red">
                    <tr>
                        <td class="errormessage">
                            <%
                            Response.write "Warning...This job has been CANCELLED "
                             %>
                        </td>
                    </tr>
                </table>               
                </td>
            </tr>

            <%end if %>
          <tr><td align="center"class="FleetExpressTextBlackBold" colspan="2">
          <form method="post">
              <table>
                <tr>
                     <td valign="bottom" align="left">
                             Courier:&nbsp;&nbsp;
                            <!-------------------->
                                <select name="Courier">
                                <%
                                    'If trim(Courier)="" then Courier="Bobtail 4" End if
									Set oConn = Server.CreateObject("ADODB.Connection")
									oConn.ConnectionTimeout = 100
									oConn.Provider = "MSDASQL"
									oConn.Open INTRANET
										l_cSQL = "Select VehicleName FROM FleetVehicles WHERE vehiclestatus='c' and isFleetExpress='y' ORDER BY LogisticorpOwned desc, VehicleName"
										SET oRs = oConn.Execute(l_cSql)
												Do While not oRs.EOF
                                                VehicleName=oRs("VehicleName")
										    %>
											<option value="<%=VehicleName%>" <%if trim(VehicleName)=trim(Courier) then response.Write " selected" end if%>><%=VehicleName%></option>
											<%
										oRs.movenext
										LOOP
									Set oConn=Nothing
                                %>
                                </select>
                            <!-------------------->









                            <br /><br /> 
                            Ref #:&nbsp;&nbsp;<input type="text" name="referenceNumber" size="20" maxlength="20" /><br /><br />
                            Special Instructions:&nbsp;&nbsp;<textarea name="addedcomments" cols="40" rows="5"></textarea>
                            <input type="hidden" name="comments" value="<%=comments %>" />
                            <br /><br />
                            
                </tr>
              </table></td></tr>
              <tr><td><input type="submit" name="submit" value="Dispatch Order" /> 

          </form>
          <form method="post">
                            <input type="submit" name="submit" value="Cancel Order" />
          </form>
         <%End if %>
          </td>
       </tr>

         <tr>
            <td align="center" colspan="2">
                <%
               ' Response.write "Submit="&Submit&"<BR>"
                'Response.write "Errormessage="&Errormessage&"<BR>"
                If (Submit="Approve Order" or Submit="Dispatch Order") and Errormessage>"" then%>              
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
                <%If Trim(SuccessMessage)>"" then%>              
                <table cellpadding="2" cellspacing="2" border="1" bordercolor="blue">
                    <tr>
                        <td >
                            <%
                            Response.write "<font color='blue'><b> * * * SUCCESS:  "&SuccessMessage& " * * * </b></font>"
                             %>
                        </td>
                    </tr>
                </table>
                <%End if %>
            </td></tr>
        <input type="hidden" name="ColorSelect" value="<%=ColorSelect %>" />
        <input type="hidden" name="pagestatus" value="submit" />
    </table>
<!--/form-->
<%else %>
           <form method="post" action="FleetExpressOrderDispatch.asp">
            <tr><td>&nbsp;</td></tr>
            <tr><td>&nbsp;</td></tr>
            <tr><td>&nbsp;</td></tr>
            <tr><td align="center" colspan="2" class="FleetExpressTextBlackBold">Job Number:&nbsp;&nbsp;<input type="text" value="<%=SearchJobNumber%>" name="SearchJobNumber" />
            <input type="hidden" name="PageStatus" value="disp" />
            <input type="submit" value="Find Job" name="submit" /></td></tr>
            <input type="hidden" name="FindJob" value="y" />
            </form>
            <tr><td>&nbsp;</td></tr>
            <%if trim(ErrorMessage)>"" then %>
            <tr><td class="errormessage" align="center" colspan="2"><%=ErrorMessage %></td></tr>
            <%end if %>
            <%if trim(SuccessMessage)>"" then %>
            <tr><td class="successmessage" align="center" colspan="2"><font color="blue"><b><%=SuccessMessage %></b></font></td></tr>
            <%end if %>
             <tr>
                <td colspan="2" align="center">
                    <table>
                        <tr>
                            <td colspan="9" align="center">
                                <B>OPEN ORDERS</B>
                            </td>
                        </tr>
                        <tr><td></td></tr>
                        <tr><td>JOB NUMBER</td><td width="20">&nbsp;</td><td>STATUS</td><td width="20">&nbsp;</td><td>BOOK DATE/TIME</td><td width="20">&nbsp;</td><td>ORIGINATION</td><td width="20">&nbsp;</td><td>DESTINATION</td><td width="20">&nbsp;</td><td>ROUTED TO</td></tr>

                        <%  
                            Today=now()
                            TargetDate=Today-4
                            'Response.write "today="&today&"<BR>"
                            'Response.write "TargetDate="&TargetDate&"<BR>"
			                Set RSEVENTS2 = Server.CreateObject("ADODB.Connection")
			                RSEVENTS2.ConnectionTimeout = 100
			                RSEVENTS2.Provider = "MSDASQL"
			                RSEVENTS2.Open DATABASE
				                l_cSQL = "SELECT     fcfgthd.fh_id, fclegs.fl_sf_name, fclegs.fl_st_name, fcfgthd.fh_ship_dt, fcfgthd.fh_carr_id, fcfgthd.fh_status, fcrefs.rf_box FROM fclegs INNER JOIN fcfgthd ON fclegs.fl_fh_id = fcfgthd.fh_id INNER JOIN fcrefs ON fcfgthd.fh_id = fcrefs.rf_fh_id  " &_
						                 "WHERE (fh_status='CAN' and fh_ship_dt>'"&TargetDate&"') or fh_status='RAP' or fh_status='OPN' or fh_status='SCD'  order by fh_id desc" 
				                'Response.write "l_cSQL="&l_cSQL&"<BR>"
                                'Response.write "Database="&Database&"<BR>"
                                SET oRs = RSEVENTS2.Execute(l_cSql)
                                If oRs.eof then
                                    ErrorMessage="There are currently no open jobs"
				                End if
				                Do While not oRs.EOF	
                                    PageStatus="Edit"
                                    fh_id=Trim(oRs("fh_id"))
                                    fl_sf_name=Trim(oRs("fl_sf_name"))
                                    fl_st_name=Trim(oRs("fl_st_name"))
                                    fh_ship_dt=Trim(oRs("fh_ship_dt"))
                                    fh_carr_id=Trim(oRs("fh_carr_id"))
                                    fh_status=Trim(oRs("fh_status"))
                                    rf_box=trim(oRs("rf_box"))
                                    If left(rf_box, 1)="X" then
                                        Xfont="orange"
                                        else
                                        xfont="black"
                                    End if
                                %>
                               <tr><td nowrap><a href="FleetExpressOrderDispatch.asp?SearchJobNumber=<%=fh_ID %>&PageStatus=disp&findjob=y&btid=86"><%=fh_ID %></a></td><td width="20">&nbsp;</td><td nowrap> <font color="<%=XFont %>"><%=fh_status %></font></td><td width="20">&nbsp;</td><td nowrap> <font color="<%=XFont %>"><%=fh_ship_dt %></font></td><td width="20">&nbsp;</td><td> <font color="<%=XFont %>"><%=fl_sf_name %></font></td><td width="20">&nbsp;</td><td> <font color="<%=XFont %>"><%=fl_st_name %></font></td><td width="20">&nbsp;</td><td> <font color="<%=XFont %>"><%=fh_carr_id %></font></td></tr></font>
                                <%
								oRs.movenext
								LOOP
                                oRs.close
                                Set oRs=Nothing
                            RSEVENTS2.Close
			                Set RSEVENTS2=Nothing
                    %>
                    </table>
                </td>
            </tr>
<%end if %>
<tr><td>&nbsp;</td></tr>
<tr><td>&nbsp;</td></tr>
<tr><td>
            <table><tr><td>
           <form method="post" action="FleetExpressOrderEdit.asp?BTID=86">
                            <input type="hidden" name="SearchJobNumber" value="<%=jobnum%>" />
                            <input type="submit" name="submit" value="Edit Order" />
                            <input type="hidden" name="btid" value="86" />
           </form>
           </td>
           <td>
          <form method="post" action="../../FleetExpressDeliveries.asp?BTID=86">
                            <input type="hidden" name="SearchJobNumber" value="<%=jobnum%>" />
                            <input type="submit" name="submit" value="Close Order" />
                            <input type="hidden" name="btid" value="86" />
           </form>
           </td>
           <td>
          <form method="post" action="FleetExpressOrderInternal.asp?Internal=y">
                            <input type="submit" name="submit" value="Enter Order" />
                            <input type="hidden" name="btid" value="86" />
          </form>
          </td>
           <td>
          <form method="post" action="FleetExpressOrderDispatch.asp?BTID=86">
                            <input type="submit" name="submit" value="Return to Dispatch" />
                            <input type="hidden" name="btid" value="86" />
          </form>
          </td>
          </tr></table> 
</td></tr>
</table>
</body>
</html>
