<html>
<head>                                                                                                                              

<!-- #include file="../fleetexpress.inc" -->
<link rel="stylesheet" type="text/css" href="../css/Style.css">
<%

   '''''''''''HARDCODED STUFF
   sBT_ID=valid8(Request.QueryString("BTID"))
   If trim(sBT_ID)="" then
        sBT_ID=valid8(Request.Form("BTID"))
    End if
   'sBT_ID="84"
   'If sBT_ID>"" then
        sBT_ID="86"
        'BillToID=sBT_ID
   'End if
   sBT_ID = ""
    'UserID=Request.Form("UserID")
    'If trim(userID)="" then
        'UserID=Session("UserID")
    'End if
    'If trim(UserID)>"" then
        'Session("UserID")=UserID
    'End if
   'Response.write "UserID="&UserID&"<BR>"
   'Response.write "sbt_id="&Session("sBT_ID")&"<BR>"
   If sBT_ID<>"86" or Trim(UserID)="" then
        'Response.write "TRIED TO REDIRECT YOU!!!!<BR>"
        'Response.redirect("/intranet/default.asp")
   End if
   
   ViewType = valid8(trim(Request.Form("Viewtype")))
   if len(ViewType) < 1 then
    ViewType = valid8(trim(Request.QueryString("Viewtype")))
   end if
   if len(ViewType) < 1 then
    ViewType = "Today"
  end if
    SpecialInstructions=valid8(Request.form("SpecialInstructions"))
    'Comments=valid8(Request.form("Comments"))
    'AddedComments=valid8(Request.form("AddedComments"))
    'AddedComments=Replace(AddedComments, """", "`")
    'AddedComments=Replace(AddedComments, "'", "`")   
    'If trim(AddedComments)>"" then
    '     NewComments=Comments&"<br>"&AddedComments
    '    else
    '    NewComments=Comments
    'End if

   'If trim(UserID)="" then response.redirect("../../default.asp") end if
   findJob=valid8(request.form("findjob"))
   if trim(findjob)="" then
        FindJob=valid8(Request.QueryString("findjob"))
    End if
  'Response.write "sBT_ID="&sBT_ID&"<BR>"

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
    PageTitle="FLEETX DISPATCH ORDER"
 %>
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
    Internal=valid8(Request.Querystring("Internal"))
    PageStatus=valid8(Request.form("PageStatus"))
    If trim(PageStatus)="" then
        PageStatus=valid8(Request.QueryString("PageStatus"))
    end if
    BillToID=sBT_ID
    'PageStatus=Request.Querystring("pid")
    JobNum=valid8(Request.Querystring("SearchJobNumber"))
    If trim(JobNum)="" then
        JobNum=valid8(Request.form("SearchJobNumber"))
    End if
    Submit=Request.form("gobutton")
    ReferenceNumber=valid8(Request.form("ReferenceNumber"))
    Courier=valid8(Request.form("Courier"))
    CurrentDate=Now()
    showtime=valid8(Request.QueryString("ShowTime"))

    'response.write "110 submit=" & Submit & "<br>"


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
        'SQL = "SELECT fh_status, fh_ship_dt, fh_ready, fh_co_id, fh_co_phone, fh_co_email, fh_co_costcenter, fh_carr_id, fh_custpo, fh_priority, fh_ref FROM fcfgthd where (fh_id = '"& jobnum &"') AND (fh_bt_id='"& BillToID &"')"
        SQL = "SELECT fh_status, fh_ship_dt, fh_ready, fh_co_id, fh_co_phone, fh_co_email, fh_co_costcenter, fh_carr_id, fh_custpo, fh_priority, fh_ref FROM fcfgthd where (fh_id = '"& jobnum &"')"
		'Response.Write "123 SQL="&SQL&"<BR>"
		RSEVENTS.Open SQL, DATABASE, 1, 3
       if RSEVENTS.eof then
           jobnum=0
           findjob = "n"
       Else
        fh_status=RSEVENTS("fh_status")
        fh_ship_dt=RSEVENTS("fh_ship_dt")
        PickUpDateTime=RSEVENTS("fh_ready")
        RequestorName=RSEVENTS("fh_co_id")
        Requestorphonenumber=RSEVENTS("fh_co_phone")
        Requestoremailaddress=RSEVENTS("fh_co_email")
        costcenternumber=RSEVENTS("fh_co_costcenter")
        'IsPalletized=RSEVENTS("IsPalletized")
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
		SQL = "SELECT fl_sf_name, fl_sf_clname, fl_sf_phone, fl_sf_email, fl_sf_addr1, fl_sf_addr2, fl_sf_city, fl_sf_state, fl_sf_country, fl_sf_zip, fl_sf_comment, fl_st_rta, fl_st_name, fl_st_clname, fl_st_phone, fl_st_email, fl_st_addr1, fl_st_addr2, fl_st_city, fl_st_state, fl_st_country, fl_st_zip, fl_t_release, fl_t_atd   FROM fclegs where (fl_fh_id = '"& jobnum &"')"
		'''''Response.Write "SQL="&SQL&"<BR>"
		RSEVENTS.Open SQL, DATABASE, 1, 3
        OriginationCompany=RSEVENTS("fl_sf_name")
        OriginationContactname=RSEVENTS("fl_sf_clname")
        OriginationphoneNumber=RSEVENTS("fl_sf_phone")
        Originationemail=RSEVENTS("fl_sf_email")
        Originationaddress=RSEVENTS("fl_sf_addr1")
        Originationaddress2=RSEVENTS("fl_sf_addr2")
        Originationcity=RSEVENTS("fl_sf_city")
        Originationstate=RSEVENTS("fl_sf_state")
        Originationcountry=RSEVENTS("fl_sf_country")
        Originationzip=RSEVENTS("fl_sf_zip")
        'Pieces=RSEVENTS("NumberOfPieces")
        comments=RSEVENTS("fl_sf_comment")
        DeliveryDateTime=RSEVENTS("fl_st_rta")
        DestinationCompany=RSEVENTS("fl_st_name")
        DestinationContactname=RSEVENTS("fl_st_clname")
        DestinationphoneNumber=RSEVENTS("fl_st_phone")
        Destinationemail=RSEVENTS("fl_st_email")
        Destinationaddress=RSEVENTS("fl_st_addr1")
        Destinationaddress2=RSEVENTS("fl_st_addr2")
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
		'SQL = "SELECT rf_box, POD, NumberOfPieces, IsPalletized, Weight, DimLength, DimWidth, DimHeight, Hazmat, Refrigerate FROM fcrefs where (rf_fh_id = '"& jobnum &"')"
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
  End If
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
                    'Body = Body & "Pieces: "&  Pieces &" "& rf_box &"<br>"
                    Body = Body & "Palletized: "&  IsPalletized &"<br>"   
                    'Body = Body & "Number Of Pallets: "&  NumberOfPallets &"<br>"  
                    Body = Body & "Weight: "&  DimWeight &"LBS<br>"
                    Body = Body & " Dimensions: "&  DimLength &" X "&  DimWidth &" X "&  DimHeight &" inches<br>"       
  
                    
                    Body = Body & "Hazmat: "&  IsHazmat &"<br>"
                    Body = Body & "Refrigerate: "&  Refrigerate &"<br><br>"
                    Body = Body & "ORIGINATION:<BR>"   
                    Body = Body & "Company: "&  OriginationCompany &"<br>"   
                    Body = Body & "Address: "&  OriginationAddress &"<br>" 
                    Body = Body & "Suite/Cube/Dock: "&  OriginationAddress2 &"<br>"   
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
                    Body = Body & "Suite/Cube/Dock: "&  DestinationAddress2 &"<br>"  
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
                    Body = Body & "***Should you need to contact us regarding this order, please  either email FleetX@LogisticorpGroup.com or call 214-882-0620***<BR><BR>"
                    'Body = Body & "<a href='http://www.logisticorp.us/intranet/dedicatedfleets/orderentry/FleetExpressOrderConfirmation.asp?bid=84&pid=manage&jid="& newjobnum &"'>To Approve or Refuse this request, click here</a><br><br>" 


				    Body = Body & "Thank you,<br><br>"  
				    Body = Body & "FleetX<br>"  
				    Body = Body &  "FleetX@LogistiCorp.us<br>"  
				    Body = Body & "214-882-0620<br><br>"
				    'Recipient=FirstName&" "&LastName
			        SentToEmail=RequestorEmailAddress
				    'Email="KWETI.Mailbox@am.kwe.com"
				    'Email="mark@maggiore.net"
				    'Set objMail = CreateObject("CDONTS.Newmail")
				    'objMail.From = "FleetX@LogisticorpGroup.com"
				    varTo = SentToEmail
                    varcc = "mark.maggiore@logisticorp.us"
				    'objMail.cc = "4692269939@tmomail.net;htmlmale@yahoo.com"
				    'objMail.cc = "4692269939@tmomail.net;pcurrin@ti.com"
				    varSubject = "Cancelled FleetX Shipment Request"
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
                        SuccessMessage="You have Cancelled order #"& JobNum 
    End if
    If trim(submit)="Dispatch Order" then
        'Response.write "LOOK!  I GOT HERE!!! LINE 338<BR>"
        'JobNum=Request(JobNum)
            'If trim(ReferenceNumber)="" then errormessage="You must provide a reference number" end if
                    If trim(errormessage)="" then
                        If trim(jobnum)="" then
                            JobNum=Request.form("JobNum")
                        End if
                     	Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
                  		RSEVENTS.CursorLocation = 3
                  		RSEVENTS.CursorType = 3
                          'Response.write "Database="&Database&"<br>"
                  		RSEVENTS.ActiveConnection = DATABASE
                  		'SQL = "SELECT fh_status, fh_ship_dt, fh_ready, fh_co_id, fh_co_phone, fh_co_email, fh_co_costcenter, fh_custpo, fh_priority FROM fcfgthd where (fh_id = '"& jobnum &"') AND (fh_bt_id='"& BillToID &"') and (fh_status='SCD')"
                          'SQL = "SELECT fh_status, fh_ship_dt, fh_ready, fh_co_id, fh_co_phone, fh_co_email, fh_co_costcenter, fh_carr_id, fh_custpo, fh_priority, fh_ref FROM fcfgthd where (fh_id = '"& jobnum &"') AND (fh_bt_id='"& BillToID &"')"
                          SQL = "SELECT fh_status, fh_ship_dt, fh_ready, fh_co_id, fh_co_phone, fh_co_email, fh_co_costcenter, fh_carr_id, fh_custpo, fh_priority, fh_ref FROM fcfgthd where (fh_id = '"& jobnum &"')"
                  		'Response.Write "123 SQL="&SQL&"<BR>"
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
                            if trim(len(requestoremailaddress)) < 1 or trim(requestoremailaddress) ="" then
                              requestoremailaddress = "mark.maggiore@logisticorp.us"
                            end if

                          costcenternumber=RSEVENTS("fh_co_costcenter")
                          'IsPalletized=RSEVENTS("IsPalletized")
                          fh_carr_id=RSEVENTS("fh_carr_id")
                          ponumber=RSEVENTS("fh_custpo")
                          priority=RSEVENTS("fh_priority")
                          fh_ref=RSEVENTS("fh_ref")
                          'fh_user6=RSEVENTS("fh_user6")
                          'response.write "fh_ref="&fh_ref&"<BR>"
                  		RSEVENTS.close
                  	Set RSEVENTS = Nothing
                  	Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
                  		RSEVENTS.CursorLocation = 3
                  		RSEVENTS.CursorType = 3
                  		RSEVENTS.ActiveConnection = DATABASE
                  		SQL = "SELECT fl_sf_id, fl_sf_name, fl_sf_clname, fl_sf_phone, fl_sf_email, fl_sf_addr1, fl_sf_addr2, fl_sf_city, fl_sf_state, fl_sf_country, fl_sf_zip, fl_sf_comment, fl_st_rta, fl_st_id, fl_st_name, fl_st_clname, fl_st_phone, fl_st_email, fl_st_addr1, fl_st_addr2, fl_st_city, fl_st_state, fl_st_country, fl_st_zip, fl_t_release, fl_t_atd   FROM fclegs where (fl_fh_id = '"& jobnum &"')"
                  		'Response.Write "SQL="&SQL&"<BR>"
                  		RSEVENTS.Open SQL, DATABASE, 1, 3
                          TempFl_sf_id=RSEVENTS("fl_sf_id")
                          OriginationCompany=RSEVENTS("fl_sf_name")
                          OriginationContactname=RSEVENTS("fl_sf_clname")
                          OriginationphoneNumber=RSEVENTS("fl_sf_phone")
                          Originationemail=RSEVENTS("fl_sf_email")
                          Originationaddress=RSEVENTS("fl_sf_addr1")
                          Originationaddress2=RSEVENTS("fl_sf_addr2")
                          Originationcity=RSEVENTS("fl_sf_city")
                          Originationstate=RSEVENTS("fl_sf_state")
                          Originationcountry=RSEVENTS("fl_sf_country")
                          Originationzip=RSEVENTS("fl_sf_zip")
                          'Pieces=RSEVENTS("NumberOfPieces")
                          comments=RSEVENTS("fl_sf_comment")
                          'Response.write "Comments="&Comments&"<BR>"
                          If trim(SpecialInstructions)>"" then
                                NewComments=Comments&"<BR>"&SpecialInstructions
                                Else
                                NewComments=Comments
                          End if
                          NewComments=Replace(NewComments,"'","`")
                          NewComments=Replace(NewComments,"""","`")
                          'REsponse.write "*****************************<BR>"
                          'Response.write "Comments="&Comments&"<BR>"
                          'Response.write "SpecialInstructions="&SpecialInstructions&"<BR>"
                          'Response.write "NewComments="&NewComments&"<BR>"
                          'REsponse.write "*****************************<BR>"
                          DeliveryDateTime=RSEVENTS("fl_st_rta")
                          TempFl_ST_ID=RSEVENTS("fl_st_id")
                          DestinationCompany=RSEVENTS("fl_st_name")
                          DestinationContactname=RSEVENTS("fl_st_clname")
                          DestinationphoneNumber=RSEVENTS("fl_st_phone")
                          Destinationemail=RSEVENTS("fl_st_email")
                          Destinationaddress=RSEVENTS("fl_st_addr1")
                          Destinationaddress2=RSEVENTS("fl_st_addr2")
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
                  		'SQL = "SELECT rf_box, POD, NumberOfPieces, IsPalletized, Weight, DimLength, DimWidth, DimHeight, Hazmat, Refrigerate FROM fcrefs where (rf_fh_id = '"& jobnum &"')"
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
	                    Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
		                    RSEVENTS.CursorLocation = 3
		                    RSEVENTS.CursorType = 3
                            'Response.write "Database="&Database&"<br>"
		                    RSEVENTS.ActiveConnection = DATABASE
		                    'SQL = "SELECT fh_status, fh_ship_dt, fh_ready, fh_co_id, fh_co_phone, fh_co_email, fh_co_costcenter, fh_custpo, fh_priority FROM fcfgthd where (fh_id = '"& jobnum &"') AND (fh_bt_id='"& BillToID &"') and (fh_status='SCD')"
                            SQL = "SELECT un_id, un_dr_id FROM fcunits where (un_desc = '"& courier &"')"
		                    'Response.Write "332 SQL="&SQL&"<BR>"
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
                            'RSEVENTS2("Pieces")=Pieces
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
							
                            'Response.write "402 l_cSQL="&l_cSQL&"<BR>"
							
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
							'Response.write "XXXXXXXXl_cSQL="&l_cSQL&"<BR>"
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
                    'Body = Body & "Pieces: "&  Pieces &" "& rf_box &"<br>"
                    Body = Body & "Palletized: "&  IsPalletized &"<br>"   
                    'Body = Body & "Number Of Pallets: "&  NumberOfPallets &"<br>"  
                    Body = Body & "Weight: "&  DimWeight &"LBS<br>"
                    Body = Body & " Dimensions: "&  DimLength &" X "&  DimWidth &" X "&  DimHeight &" inches<br>"       
  
                    
                    Body = Body & "Hazmat: "&  IsHazmat &"<br>"
                    Body = Body & "Refrigerate: "&  Refrigerate &"<br><br>"
                    Body = Body & "ORIGINATION:<BR>"   
                    Body = Body & "Company: "&  OriginationCompany &"<br>"   
                    Body = Body & "Address: "&  OriginationAddress &"<br>" 
                    Body = Body & "Suite/Cube/Dock: "&  OriginationAddress2 &"<br>"   
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
                    Body = Body & "Suite/Cube/Dock: "&  DestinationAddress2 &"<br>" 
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

                    Body = Body & "***Should you need to contact us regarding this order, please  either email FleetX@LogisticorpGroup.com or call 214-882-0620***<BR><BR>"
				    Body = Body & "Thank you,<br><br>"  
				    Body = Body & "FleetX<br>"  
				    Body = Body &  "FleetX@LogisticorpGroup.com<br>"  
				    Body = Body & "214-882-0620<br><br>"
				    'Recipient=FirstName&" "&LastName
			        'If lcase(RequestorEmailAddress)<>"fleetx@logisticorp.us"  AND lcase(RequestorEmailAddress)<>"texasinstruments@plg.cc" then
                        SentToEmail=RequestorEmailAddress
				        'Email="KWETI.Mailbox@am.kwe.com"
				        'Email="mark@maggiore.net"
				        'Set objMail = CreateObject("CDONTS.Newmail")
				        'objMail.From = "FleetX@LogisticorpGroup.com"
				        'objMail.To = "wiseweblady@gmail.com"
                        If TempFl_sf_id="TISHR" or TempFl_sf_id="PHO" or TempFl_st_id="TISHR" or TempFl_st_id="PHO" then
                            varTo = SentToEmail&";kchitwood@ti.com"
                            Else
				            varTo = SentToEmail
                        End if
                        'objMail.cc = "mark.maggiore@logisticorp.us;betty.walker@logisticorp.us"
                        varcc = "mark.maggiore@logisticorp.us"
				        'objMail.cc = "4692269939@tmomail.net;htmlmale@yahoo.com"
				        'objMail.cc = "4692269939@tmomail.net;pcurrin@ti.com"
				        varSubject = "Dispatched FleetX Shipment Request"
				        'objMail.MailFormat = cdoMailFormatMIME
				        'objMail.BodyFormat = cdoBodyFormatHTML
				        'objMail.Body = Body
				        'objMail.Send
                    'End if
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
              'response.write "497 email sent: " & Body & "<br>"
              'response.write "TempFl_sf_id="&TempFl_sf_id&"<BR>"
              'response.write "TempFl_st_id="&TempFl_st_id&"<BR>"

                        SuccessMessage="You have DISPATCHED order #"& JobNum 
                    End if
    End if

    If trim(submit)="DISPATCH SELECTED ORDERS" then
      'check that bobtail or tractor not assigned to van'
       courier =  valid8(trim(Request.Form("RouteVehicle")))
       If Trim(Courier)>"" then

            'If trim(ReferenceNumber)="" then errormessage="You must provide a reference number" end if

      errormessage = ""
      'check that all origs and dests match'
          sorig = ""
          sdest = ""
          For Each selorder in Request.Form("selectedorder")
                     	Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
                    		RSEVENTS.CursorLocation = 3
                    		RSEVENTS.CursorType = 3
                    		RSEVENTS.ActiveConnection = DATABASE
                   		SQL = "SELECT fl_sf_name, fl_st_name  FROM fclegs where (fl_fh_id = '"& selorder &"')"
                    		'Response.Write "601 SQL="&SQL&"<BR>"
                    		RSEVENTS.Open SQL, DATABASE, 1, 3
                       if sorig = "" then
                        sorig = RSEVENTS("fl_sf_name")
                      end if
                      if sdest = "" then
                        sdest = RSEVENTS("fl_st_name")
                      end if
            if sorig <> RSEVENTS("fl_sf_name") or sdest <> RSEVENTS("fl_st_name") then
                errormessage = "Originations and Destinations must be the same for all jobs selected - please try again<br><br>"
            end if
                      RSEVENTS.close
                    	Set RSEVENTS = Nothing
          Next
       

       'response.write "628 - courier=" & courier & "<br>"
                     	Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
                    		RSEVENTS.CursorLocation = 3
                    		RSEVENTS.CursorType = 3
                    		RSEVENTS.ActiveConnection = DATABASE
                   		  SQL = "SELECT VehicleType FROM AvailableVehicles WHERE VehicleName = '" & courier & "' and AvailableStatus = 'c'"
                    		'Response.Write "632 SQL="&SQL&"<BR>"
                    		RSEVENTS.Open SQL, DATABASE, 1, 3
                        couriertype = trim(RSEVENTS("VehicleType"))
                        'response.write "636 couriertype=" & couriertype & "<br>"
                        RSEVENTS.close
                    	  Set RSEVENTS = Nothing
                        
                        if couriertype = "van" then
                            ' check bobtail or tractor are not selected
                            For Each selorder in Request.Form("selectedorder")
                               	  Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
                              		RSEVENTS.CursorLocation = 3
                              		RSEVENTS.CursorType = 3
                              		RSEVENTS.ActiveConnection = DATABASE
                             		  SQL = "SELECT fh_user4  FROM fcfgthd where (fh_id = '"& selorder &"')"
                              		'Response.Write "648 SQL="&SQL&"<BR>"
                              		RSEVENTS.Open SQL, DATABASE, 1, 3
                                  scourier =  trim(RSEVENTS("fh_user4"))
                                  'response.write "scourier=" & scourier & "<br>"
                                  if (scourier = "Tractor") or (scourier = "Bobtail") then
                                        errormessage = errormessage & "(" & selorder & ") Heavy Freight cannot be dispatched to a Van - please try again<br>"
                                  end if
                                  RSEVENTS.close
                                  Set RSEVENTS = Nothing
                            Next
                      end if
                      
          If courier = "SR-Material Handler" Then

            For Each selorder in Request.Form("selectedorder")
                     	Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
                    		RSEVENTS.CursorLocation = 3
                    		RSEVENTS.CursorType = 3
                    		RSEVENTS.ActiveConnection = DATABASE
                   		SQL = "SELECT fl_sf_name, fl_st_name  FROM fclegs where (fl_fh_id = '"& selorder &"')"
                    		'Response.Write "601 SQL="&SQL&"<BR>"
                    		RSEVENTS.Open SQL, DATABASE, 1, 3
                        sorig = trim(RSEVENTS("fl_sf_name"))
                        sdest = trim(RSEVENTS("fl_st_name"))
                        'response.write "sorig=" & sorig & ", sdest = " & sdest & "<br>"
              if sorig = "ESTK" and (sdest = "D6N1" or sdest = "D6N2" or sdest = "D6W3" or sdest = "DM4M" or sdest = "DM5M" or sdest = "DM5Q" = sdest = "DM6Q" or sdest = "DPI2") then
              else
                errormessage = errormessage & "(" & selorder &") Origination and Destination invalid for SR-Material Handler Vehicle - please try again<br><br>"
              end if
                      RSEVENTS.close
                    	Set RSEVENTS = Nothing
            Next
            
          End If
 
                      
           'response.write "660 error=" & errormessage & "<br>"           
          if len(errormessage) < 1 then
          
            'get new batch nbr
            Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open DATABASE
            
              lSQL = "SELECT MAX(jb_batch_id) as BID FROM JobBatches"                            
							SET oRsB = oConn.Execute(lSQL)
              BID = oRsB("BID")
              if NOT isNumeric(BID) then
                BID = 0
              end if
              BID = cLng(BID)
              BID = BID + 1
              
            Set oRsB = Nothing  
						Set oConn=Nothing

          For Each selorder in Request.Form("selectedorder")
                    If trim(errormessage)="" then
                    'response.write "509 order/job=" & selorder & "<br>"
                       JobNum = selorder
                       courier =  valid8(Request.Form("RouteVehicle"))
                       newcomments = valid8(Request.Form("specintructions"))
                    	 Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
                    		RSEVENTS.CursorLocation = 3
                    		RSEVENTS.CursorType = 3
                            'Response.write "Database="&Database&"<br>"
                    		RSEVENTS.ActiveConnection = DATABASE
                    		'SQL = "SELECT fh_status, fh_ship_dt, fh_ready, fh_co_id, fh_co_phone, fh_co_email, fh_co_costcenter, fh_custpo, fh_priority FROM fcfgthd where (fh_id = '"& jobnum &"') AND (fh_bt_id='"& BillToID &"') and (fh_status='SCD')"
                            'SQL = "SELECT fh_status, fh_ship_dt, fh_ready, fh_co_id, fh_co_phone, fh_co_email, fh_co_costcenter, fh_carr_id, fh_custpo, fh_priority, fh_ref FROM fcfgthd where (fh_id = '"& jobnum &"') AND (fh_bt_id='"& BillToID &"')"
                            SQL = "SELECT fh_status, fh_ship_dt, fh_ready, fh_co_id, fh_co_phone, fh_co_email, fh_co_costcenter, fh_carr_id, fh_custpo, fh_priority, fh_ref FROM fcfgthd where (fh_id = '"& jobnum &"')"
                    		'Response.Write "123 SQL="&SQL&"<BR>"
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
                            if trim(len(requestoremailaddress)) < 1 or trim(requestoremailaddress) ="" then
                              requestoremailaddress = "mark.maggiore@logisticorp.us"
                            end if
                            'response.write "614 requestor=" & requestoremailaddress & "<br>"
                            costcenternumber=RSEVENTS("fh_co_costcenter")
                            'IsPalletized=RSEVENTS("IsPalletized")
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
                    		SQL = "SELECT fl_sf_id, fl_sf_name, fl_sf_clname, fl_sf_phone, fl_sf_email, fl_sf_addr1, fl_sf_addr2, fl_sf_city, fl_sf_state, fl_sf_country, fl_sf_zip, fl_sf_comment, fl_st_rta, fl_st_id, fl_st_name, fl_st_clname, fl_st_phone, fl_st_email, fl_st_addr1, fl_st_addr2, fl_st_city, fl_st_state, fl_st_country, fl_st_zip, fl_t_release, fl_t_atd   FROM fclegs where (fl_fh_id = '"& jobnum &"')"
                    		'Response.Write "SQL="&SQL&"<BR>"
                    		RSEVENTS.Open SQL, DATABASE, 1, 3
                            TempFl_SF_ID=RSEVENTS("fl_sf_id")
                            OriginationCompany=RSEVENTS("fl_sf_name")
                            OriginationContactname=RSEVENTS("fl_sf_clname")
                            OriginationphoneNumber=RSEVENTS("fl_sf_phone")
                            Originationemail=RSEVENTS("fl_sf_email")
                            Originationaddress=RSEVENTS("fl_sf_addr1")
                            Originationaddress2=RSEVENTS("fl_sf_addr2")
                            Originationcity=RSEVENTS("fl_sf_city")
                            Originationstate=RSEVENTS("fl_sf_state")
                            Originationcountry=RSEVENTS("fl_sf_country")
                            Originationzip=RSEVENTS("fl_sf_zip")
                            'Pieces=RSEVENTS("NumberOfPieces")
                            comments=RSEVENTS("fl_sf_comment")
                            'Response.write "COMMENTS="&comments&"<BR>"
                            If trim(SpecialInstructions)>"" then
                                If trim(Comments)>"" then
                                    NewComments=Comments&"<BR>"&SpecialInstructions
                                    'Response.write "1NewComments="&NewComments&"<BR>"
                                    Else
                                    NewComments=SpecialInstructions
                                    'Response.write "2NewComments="&NewComments&"<BR>"
                                End if
                                Else
                                NewComments=Comments
                                'Response.write "3NewComments="&NewComments&"<BR>"
                            End if
                            NewComments=Replace(NewComments,"'","`")
                            NewComments=Replace(NewComments,"""","`")
                            DeliveryDateTime=RSEVENTS("fl_st_rta")
                            TempFl_ST_ID=RSEVENTS("fl_st_id")
                            DestinationCompany=RSEVENTS("fl_st_name")
                            DestinationContactname=RSEVENTS("fl_st_clname")
                            DestinationphoneNumber=RSEVENTS("fl_st_phone")
                            Destinationemail=RSEVENTS("fl_st_email")
                            Destinationaddress=RSEVENTS("fl_st_addr1")
                            Destinationaddress2=RSEVENTS("fl_st_addr2")
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
                    		'SQL = "SELECT rf_box, POD, NumberOfPieces, IsPalletized, Weight, DimLength, DimWidth, DimHeight, Hazmat, Refrigerate FROM fcrefs where (rf_fh_id = '"& jobnum &"')"
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
	                    Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
		                    RSEVENTS.CursorLocation = 3
		                    RSEVENTS.CursorType = 3
                            'Response.write "Database="&Database&"<br>"
		                    RSEVENTS.ActiveConnection = DATABASE
		                    'SQL = "SELECT fh_status, fh_ship_dt, fh_ready, fh_co_id, fh_co_phone, fh_co_email, fh_co_costcenter, fh_custpo, fh_priority FROM fcfgthd where (fh_id = '"& jobnum &"') AND (fh_bt_id='"& BillToID &"') and (fh_status='SCD')"
                            SQL = "SELECT un_id, un_dr_id FROM fcunits where (un_desc = '"& trim(courier) &"')"
		                    'Response.Write "line 778 SQL="&SQL&"<BR>"
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
                    'Response.write "PickUpDateTime="&PickUpDateTime&"<br>"
                    'Response.write "DeliveryDateTime="&DeliveryDateTime&"<br>"
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
                            'RSEVENTS2("Pieces")=Pieces
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
							
                            'Response.write "752 l_cSQL="&l_cSQL&"<BR>"
							
                            oConn.Execute(l_cSQL)
						Set oConn=Nothing
						Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open DATABASE
						' 7/26/04 KK: Added canceljob functionality to be able to update the status if cancel button is pressed.
						' 11/30/04 DEC: Changed from CAN/98 TO DEL/99 to be consistent with dispatchOffice	
                        
                        Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
                    		RSEVENTS2.CursorLocation = 3
                    		RSEVENTS2.CursorType = 3
                    		RSEVENTS2.ActiveConnection = DATABASE
                   		  SQL = "SELECT DriverID FROM AvailableVehicles WHERE VehicleName = '" & courier & "' and AvailableStatus = 'c'"
                    		'Response.Write "632 SQL="&SQL&"<BR>"
                    		RSEVENTS2.Open SQL, DATABASE, 1, 3
                        aDriverID = trim(RSEVENTS2("DriverID"))
                        RSEVENTS2.close
                    	  Set RSEVENTS2 = Nothing
           
                        
                        
		
							l_cSQL = "UPDATE fclegs SET fl_un_id= '"& un_id &"', fl_dr_id='"& aDriverID &"', fl_t_disp = '"& CurrentDate &"', fl_sf_comment = '"& NewComments &"' WHERE fl_fh_id = '" & JobNum & "'"
                  
                        
                        
                        		
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
                    'Body = Body & "Pieces: "&  Pieces &" "& rf_box &"<br>"
                    Body = Body & "Palletized: "&  IsPalletized &"<br>"   
                    'Body = Body & "Number Of Pallets: "&  NumberOfPallets &"<br>"  
                    Body = Body & "Weight: "&  DimWeight &"LBS<br>"
                    Body = Body & " Dimensions: "&  DimLength &" X "&  DimWidth &" X "&  DimHeight &" inches<br>"       
  
                    
                    Body = Body & "Hazmat: "&  IsHazmat &"<br>"
                    Body = Body & "Refrigerate: "&  Refrigerate &"<br><br>"
                    Body = Body & "ORIGINATION:<BR>"   
                    Body = Body & "Company: "&  OriginationCompany &"<br>"   
                    Body = Body & "Address: "&  OriginationAddress &"<br>" 
                    Body = Body & "Suite/Cube/Dock: "&  OriginationAddress2 &"<br>"  
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
                    Body = Body & "Suite/Cube/Dock: "&  DestinationAddress2 &"<br>"  
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

                    Body = Body & "***Should you need to contact us regarding this order, please  either email FleetX@LogisticorpGroup.com or call 214-882-0620***<BR><BR>"
				    Body = Body & "Thank you,<br><br>"  
				    Body = Body & "FleetX<br>"  
				    Body = Body &  "FleetX@LogisticorpGroup.com<br>"  
				    Body = Body & "214-882-0620<br><br>"
				    'Recipient=FirstName&" "&LastName
            'response.write "830 before email send<br>"
			        'If lcase(RequestorEmailAddress)<>"fleetx@logisticorp.us"  AND lcase(RequestorEmailAddress)<>"texasinstruments@plg.cc" then
                        SentToEmail=RequestorEmailAddress
                        'response.write "833 sendtoemail= " & SentToEmail & "<br>"
				        'Email="KWETI.Mailbox@am.kwe.com"
				        'Email="mark@maggiore.net"
				        'Set objMail = CreateObject("CDONTS.Newmail")
				        'objMail.From = "FleetX@LogisticorpGroup.com"
				        'objMail.To = "wiseweblady@gmail.com"
               'response.write "839 sento=" & SentToEmail & "<br>"

                        If trim(TempFl_sf_id)="TISHR" or trim(TempFl_sf_id)="PHO" or trim(TempFl_st_id)="TISHR" or trim(TempFl_st_id)="PHO" then
                            varTo = SentToEmail&";kchitwood@ti.com"
                            Else
				            varTo = SentToEmail
                        End if
                        'objMail.cc = "mark.maggiore@logisticorp.us;betty.walker@logisticorp.us"
                        varcc = "mark.maggiore@logisticorp.us"
				        'objMail.cc = "4692269939@tmomail.net;htmlmale@yahoo.com"
				        'objMail.cc = "4692269939@tmomail.net;pcurrin@ti.com"
				        varSubject = "Dispatched FleetX Shipment Request"
				        'objMail.MailFormat = cdoMailFormatMIME
				        'objMail.BodyFormat = cdoBodyFormatHTML
				        'objMail.Body = Body
				        'objMail.Send
                   'End if
				   ' Set objMail = Nothing
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
              'response.write "497 email sent: " & Body & "<br>"
              'REsponse.write "TempFl_sf_id="&TempFl_sf_id&"<BR>"
              'REsponse.write "TempFl_st_id="&TempFl_st_id&"<BR>"
                        SuccessMessage=SuccessMessage & "You have DISPATCHED order #"& JobNum & "<br>"
              End if

						'update job batches
            Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open DATABASE
            
 							l_cSQL = "INSERT INTO JobBatches (jb_batch_id,jb_fh_id) VALUES(" & BID & ",'" & JobNum & "')"
							'Response.write "l_cSQL="&l_cSQL&"<BR>"
							oConn.Execute(l_cSQL)
						Set oConn=Nothing
           
      Next

      end if
      Else
      ErrorMessage="You must select a vehicle/driver"
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
                    'Body = Body & "Pieces: "&  Pieces &" "& rf_box &"<br>"
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
                    Body = Body & "***Should you need to contact us regarding this order, please  either email FleetX@LogisticorpGroup.com or call 214-882-0620***<BR><BR>"
                    Body = Body & "Thank you,<br><br>"  
				    Body = Body & "FleetX<br>"  
				    Body = Body &  "FleetX@LogistiCorp.us<br>"  
				    Body = Body & "214-882-0620<br><br>"
				    'Recipient=FirstName&" "&LastName
			        SentToEmail=RequestorEmailAddress
				    'Email="KWETI.Mailbox@am.kwe.com"
				    'Email="mark@maggiore.net"
				    'Set objMail = CreateObject("CDONTS.Newmail")
				    'objMail.From = "FleetX@LogisticorpGroup.com"
				    varTo = SentToEmail
				    varcc = "mark.maggiore@logisticorp.us"
				    'objMail.cc = "4692269939@tmomail.net;pcurrin@ti.com"
				    varSubject = "Approved FleetX Shipment Request"
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
''''''''''''''''2. To LOGISTICORP
				    Body = "An order, #"& jobnum &", has been APPROVED!<br><br>"   

                    Body = Body & "REQUESTOR INFORMATION:<BR>"
                    Body = Body & "Name: "&  RequestorName &"<br>"  
                    Body = Body & "Phone Number: "&  RequestorPhoneNumber &"<br>"  
                    Body = Body & "Email Address: "&  RequestorEmailAddress &"<br>"   
                    Body = Body & "PO Number: "&  PONumber &"<br>"  
                    Body = Body & "Cost Center Number: "&  CostCenterNumber &"<br><br>" 
                    Body = Body & "COMMODITY INFORMATION:<BR>" 
                    'Body = Body & "Pieces: "&  Pieces &" "& rf_box &"<br>"
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
                    Body = Body & "<a href='http://www.logisticorp.us/intranet/fleetx/orderentry/FleetXOrderConfirmation.asp?bid=84&pid=disp&jid="& jobnum &"'>To dispatch this request, click here</a><br><br>" 

                    Body = Body & "***Should you need to contact us regarding this order, please  either email FleetX@LogisticorpGroup.com or call 214-882-0620***<BR><BR>"
				    Body = Body & "Thank you,<br><br>"  
				    Body = Body & "FleetX<br>"  
				    Body = Body &  "FleetX@LogistiCorp.us<br>"  
				    Body = Body & "214-882-0620<br><br>"
				    'Recipient=FirstName&" "&LastName
			        SentToEmail="xxx@LogistiCorp.us"
				    'Email="KWETI.Mailbox@am.kwe.com"
				    'Email="mark@maggiore.net"
				    'Set objMail = CreateObject("CDONTS.Newmail")
				    'objMail.From = "FleetX@LogisticorpGroup.com"
				    varTo = SentToEmail
				    varcc = "mark.maggiore@logisticorp.us"
				    'objMail.cc = "4692269939@tmomail.net;pcurrin@ti.com"
				    varSubject = "Approved FleetX Shipment Request"
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
       '' SQL = "SELECT fh_status, fh_ship_dt, fh_ready, fh_co_id, fh_co_phone, fh_co_email, fh_co_costcenter, fh_carr_id, fh_custpo, fh_priority, fh_ref FROM fcfgthd where (fh_id = '"& jobnum &"') AND (fh_bt_id='"& BillToID &"')"
        SQL = "SELECT fh_status, fh_ship_dt, fh_ready, fh_co_id, fh_co_phone, fh_co_email, fh_co_costcenter, fh_carr_id, fh_custpo, fh_priority, fh_ref FROM fcfgthd where (fh_id = '"& jobnum &"')"
		'Response.Write "SQL="&SQL&"<BR>"
		RSEVENTS.Open SQL, DATABASE, 1, 3
       if RSEVENTS.eof then
           jobnum=0
           findjob="n"
       Else
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
		SQL = "SELECT fl_sf_name, fl_sf_clname, fl_sf_phone, fl_sf_email, fl_sf_addr1, fl_sf_city, fl_sf_state, fl_sf_country, fl_sf_zip, fl_sf_comment, fl_st_comment, fl_st_rta, fl_st_name, fl_st_clname, fl_st_phone, fl_st_email, fl_st_addr1, fl_st_city, fl_st_state, fl_st_country, fl_st_zip, fl_t_release, fl_t_atd   FROM fclegs where (fl_fh_id = '"& jobnum &"')"
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
        fl_sf_comment=RSEVENTS("fl_sf_comment")
        fl_st_comment=RSEVENTS("fl_st_comment")
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

        If trim(fl_sf_comment)>"" then
            Comments=fl_sf_comment
            'Response.write "Comments1="&Comments&"<BR>"
        End if
        If trim(fl_st_comment)>"" then
        'Response.write "Comments2="&Comments&"<BR>"
            If trim(Comments)="" then
                Comments=fl_st_comment
                'Response.write "Comments3="&Comments&"<BR>"
                else
                Comments=Comments&"&nbsp;&nbsp;/&nbsp;&nbsp;"&fl_st_comment
                'Response.write "Comments4="&Comments&"<BR>"
            End if
        End if


		RSEVENTS.close
	Set RSEVENTS = Nothing
	Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
		RSEVENTS.CursorLocation = 3
		RSEVENTS.CursorType = 3
		RSEVENTS.ActiveConnection = DATABASE
		'SQL = "SELECT rf_box, POD, NumberOfPieces, IsPalletized, Weight, DimLength, DimWidth, DimHeight, Hazmat, Refrigerate FROM fcrefs where (rf_fh_id = '"& jobnum &"')"
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
  End If
End if
     %>
     	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
        <%
        'REsponse.write "FindJob="&FindJob&"<BR>"
        if FindJob<>"y" then %>

	<meta http-equiv="refresh" content="60" />
    <%end if %>


<title>FleetX - <%=PageTitle %></title>
</head>

<body onload="document.FindUser.requestorName.focus(); document.OrderForm1.<%=HighlightedField%>.focus();">
	<table align="center" border="0" bordercolor="black" cellpadding="0" cellspacing="0" ID="Table1" height="100%" width="100%">
        <tr><td align="left"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>
        <tr>
            <td align="left"><img src="../images/FleetX_Small.jpg" height="50" width="168" /></td>
            <td align="right" valign="bottom"><!-- #include file="../topnavbar.asp" --></td>
        </tr>
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" /></td></tr>	
    <tr><td colspan="2">
<!-- <form action="NewUser.asp" method="post" name="FindUser">   -->
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



    <tr><td><!-- main page stuff goes here! -->
    
    
    <%
                Set oConnA = Server.CreateObject("ADODB.Connection")
                oConnA.ConnectionTimeout = 100
                oConnA.Provider = "MSDASQL"
                oConnA.Open DATABASE
                iuSQL = "Select * FROM AutoDispatchStatus WHERE status = 'c'"
                'response.write "1702 sql=" & iuSQL & "<br>"
                SET oRsa2 = oConnA.Execute(iuSql)
                if oRsa2.eof then
                  autodispatch = "OFF"
                  adbutton = "ON"
                  adtime = Now()
                else
                  adstatus = oRsa2("status")
                  adtimeon = oRsa2("dateon")
                  adtimeoff = oRsa2("dateoff")
                  if adstatus = "c" then
                    autodispatch = "ON"
                    adbutton = "OFF"
                    adtime = adtimeon
                  else
                    autodispatch = "OFF"
                    adbutton = "ON"
                    adtime = adtimeoff
                  end if
                end if
                oRsa2.close
                Set oConnA=Nothing

%>  

<% 'if autodispatch = "ON" then  %>
                <!--AUTODISPATCH NOW PERMANENTLY ON AS OF 4/15/2018 PER LINDA
                     <table width=100% align=center><tr><td width=100% align=center>
                     <table border=1 width=100 cellspacing=1 cellpadding=5 align=center>
                     <tr><td align="center">AutoDispatch is <b><%=autodispatch%></b><br>As of <%=adtime%><br>
                     <br><form method="post" action="../Admin/AutoDispatchSwitch.asp"><input type="submit" id="gobutton" value="TURN AUTODISPATCH <%=adbutton%>" /></form></td></tr>
                     </table> 
                     &nbsp;<br>
                                   <form method="post" action="../home.asp">
                                <input type="hidden" name="SearchJobNumber" value="<%=jobnum%>" />
                                <INPUT id="gobutton" name="gobutton" TYPE="submit" name="ButtonValue" VALUE="Return to Home">-->
                                <!-- <input type="submit" name="submit" value="Return to Home" /> -->
                               
<!--<input type="hidden" name="btid" value="86" />
               </form>

                     </td></tr></table>
                     -->
<% 'else %>
    
    <table border="0" cellpadding="0" cellspacing="0" align="center">
    
   
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
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=fh_ship_dt%></td>
                    </tr>
                    <%end if %>
                    <%if fl_t_release>"" then %>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left" nowrap>Accepted/Cancelled</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=fl_t_release%></td>
                    </tr>
                    <%end if %>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left" nowrap>Requestor Name</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=RequestorName%></td>
                    </tr>
                   <tr>
                        <td class="FleetExpressTextBlackBold" align="left" nowrap>Phone Number</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=RequestorPhoneNumber%></td>
                    </tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left" nowrap>Email Address</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
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
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=GenericNumber%></td>
                    </tr>   
                    <%
                    If trim(NewComments)>"" then Comments=NewComments End if
                     %>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left" valign="top" nowrap>Special Instructions</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=Comments%></td>
                    </tr>
                    <tr><td><img src="../images/pixel.gif" width="1" height="3" /></td></tr>
                    </table>
                     </td></tr></table></td>
                     <td align="left"><img src="../images/pixel.gif" height="1" width="25" /></td>
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
                            <td width="10"><img src="../images/pixel.gif" /></td>
                            <td nowrap class="FleetExpressTextBlack">
                                Number of Pieces:&nbsp;&nbsp;<input type="text" name="Pieces" value="<%=Pieces%>" size="3" maxlength="4" />
                            </td>
                        </tr>
                        -->
                    <%if fh_carr_id>"" then %>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left" nowrap>Routed to</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=fh_carr_id%></td>
                    </tr>
                    <%end if %>
                    <%if fh_ref>"" then %>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left" nowrap>Reference Number</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=fh_ref%></td>
                    </tr>
                    <%end if %>
                    <%if fl_t_atd>"" then %>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left" nowrap>Delivered</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=fl_t_atd%></td>
                    </tr>
                    <%end if %>
                    <%if POD>"" then %>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left" nowrap>POD</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=POD%></td>
                    </tr>
                    <%end if %>
                        <tr>
                            <td class="FleetExpressTextBlackBold" align="left">Number of Pieces</td>
                            <td width="10"><img src="../images/pixel.gif" /></td>
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
                            <td width="10"><img src="../images/pixel.gif" /></td>
                            <td align="left" class="FleetExpressTextBlack"><%=DimWeight%> Pounds</td>
                        </tr>
                        <tr>
                            <td class="FleetExpressTextBlackBold" align="left">Dimensions</td>
                            <td width="10"><img src="../images/pixel.gif" /></td>
                            <td class="FleetExpressTextBlack" align="left" nowrap>
                                L:&nbsp;&nbsp;<%=DimLength%>
                                W:&nbsp;&nbsp;<%=DimWidth%> 
                                H:&nbsp;&nbsp;<%=DimHeight%>
                                &nbsp;&nbsp;Inches
                            </td>
                        </tr>
    
    
                        <tr>
                            <td class="FleetExpressTextBlackBold" align="left">Hazmat</td>
                            <td width="10"><img src="../images/pixel.gif" /></td>
                            <td align="left"  class="FleetExpressTextBlack">
                                 <%=Hazmat %>              
                            </td>
                        </tr>
                        <tr>
                            <td class="FleetExpressTextBlackBold" align="left">Refrigerate</td>
                            <td width="10"><img src="../images/pixel.gif" /></td>
                            <td align="left"  class="FleetExpressTextBlack">
                                 <%=Refrigerate %>                    
                            </td>
                        </tr>
                        <tr>
                            <td class="FleetExpressTextBlackBold" align="left">Service Level</td>
                            <td width="10"><img src="../images/pixel.gif" /></td>
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
    
            <tr><td align="left"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>
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
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=OriginationCompany%></td>
                    </tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Address</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=OriginationAddress%></td>
                    </tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">City/State/Zip</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td class="FleetExpressTextBlack">
                            <%=OriginationCity%>, <%=OriginationState %>&nbsp;&nbsp;
                            <%=OriginationZipCode%>
    
                        </td>
                    </tr>
    
    
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Contact Name</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=OriginationContactName%></td>
                    </tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Phone Number</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=OriginationPhoneNumber%></td>
                    </tr> 
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Email Address</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=OriginationEmail%></td>
                    </tr> 
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Pick Up Date/Time</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=PickUpDateTime%>
                        </td>
                    </tr>
                     </table>
                     </td></tr></table></td>
                     <td align="left"><img src="../images/pixel.gif" height="1" width="25" /></td>
                     <td align="left">
                        <table border="1" bordercolor="<%=BorderColor%>" cellpadding="0" cellspacing="0" width="<%=tablewidth%>"> 
                        <tr> <td valign="top"> <table cellpadding="3" cellspacing="0" width="100%">               <tr>
                            <td colspan="3" align="center" bgcolor="<%=BorderColor%>" class="FleetExpressBodyWhiteBold">
                                DESTINATION
                            </td>
                        </tr>
                        <tr>
                            <td class="FleetExpressTextBlackBold" align="left" nowrap>Company Name</td>
                            <td width="10"><img src="../images/pixel.gif" /></td>
                            <td align="left" class="FleetExpressTextBlack"><%=DESTINATIONCompany%></td>
                        </tr>
                        <tr>
                            <td class="FleetExpressTextBlackBold" align="left">Address</td>
                            <td width="10"><img src="../images/pixel.gif" /></td>
                            <td align="left" class="FleetExpressTextBlack"><%=DESTINATIONAddress%></td>
                        </tr>
                        <tr>
                            <td class="FleetExpressTextBlackBold" align="left">City/State/Zip</td>
                            <td width="10"><img src="../images/pixel.gif" /></td>
                            <td class="FleetExpressTextBlack">
                                <%=DESTINATIONCity%>, <%=DestinationState %>&nbsp;&nbsp; <%=DESTINATIONZipCode%>
                            </td>
                        </tr>
    
    
                        <tr>
                            <td class="FleetExpressTextBlackBold" align="left">Contact Name</td>
                            <td width="10"><img src="../images/pixel.gif" /></td>
                            <td align="left" class="FleetExpressTextBlack"><%=DESTINATIONContactName%></td>
                        </tr>
                        <tr>
                            <td class="FleetExpressTextBlackBold" align="left">Phone Number</td>
                            <td width="10"><img src="../images/pixel.gif" /></td>
                            <td align="left" class="FleetExpressTextBlack"><%=DESTINATIONPhoneNumber%></td>
                        </tr> 
                        <tr>
                            <td class="FleetExpressTextBlackBold" align="left">Email Address</td>
                            <td width="10"><img src="../images/pixel.gif" /></td>
                            <td align="left" class="FleetExpressTextBlack"><%=DESTINATIONEmail%></td>
                        </tr> 
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left">Delivery Date/Time</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack"><%=DeliveryDateTime%>
                        </td>
                    </tr>
                         </table>
                         </td></tr></table>                 
                     </td></tr>                                                                                                               
                </table>
             
                </td>
            </tr>
             <tr><td align="left"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>
             <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2"> FleetX Transportation Call Center 972-499-3415</td></tr>
             <tr><td align="left"><img src="../images/pixel.gif" height="30" width="1" /></td></tr>
    
    
    
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
                            <td class="FleetExpressTextBlack" ><br /><img src="../images/pixel.gif" height="10" width="1" /><br />
                                &nbsp;Shipper Notes/Comments:________________________________________________________________________________________________________&nbsp;<br /><img src="../images/pixel.gif" height="20" width="1" /><br />
                                &nbsp;______________________________________________________________________________________________________________________________&nbsp;<br /><img src="../images/pixel.gif" height="25" width="1" /><br />
                                &nbsp;Shipper Signature:__________________________________ Print Name:_____________________________________ Date:________________________&nbsp;<br /><img src="../images/pixel.gif" height="20" width="1" /><br />
                            </td>
                        </tr>
                        <tr>
                            <td  class="FleetExpressTextBlackSmaller">
                                This is to certify that the above named materials are properly classified, packaged, marked, and labeled, and are in proper condition for transportation according to the applicable regulations of the DOT.
                                <br /><img src="../images/pixel.gif" height="5" width="1" /><br />
                                Property described above was received by driver in good order, except as noted above.<br /><img src="../images/pixel.gif" height="10" width="1" /><br />    
                            </td>
                        </tr>
                        </td></tr>
                        </table>
                        </td></tr>
    
                    </table>
                </td>
            </tr>
            <tr><td align="left"><img src="../images/pixel.gif" height="30" width="1" /></td></tr>
            <tr>
                <td colspan="3"  class="FleetExpressTextBlack" align="center">
                &nbsp;Driver:___________________ Arrival Time:___________________ Number of Pieces Delivered:_______________ Departure Time:___________________&nbsp;
                </td>
            </tr>
            <tr><td align="left"><img src="../images/pixel.gif" height="30" width="1" /></td></tr>
             <tr>
                <td colspan="3" align="center">
                    <table border="1" bordercolor="black" cellpadding="0" cellspacing="0" width="940">
                        <tr><td>
                        <table border="0" bordercolor="black" cellpadding="0" cellspacing="0" width="940">
                        <tr>
                            <td class="FleetExpressTextBlack" ><br /><img src="../images/pixel.gif" height="10" width="1" /><br />
                                &nbsp;Consignee Notes/Comments:______________________________________________________________________________________________________&nbsp;<br /><img src="../images/pixel.gif" height="20" width="1" /><br />
                                &nbsp;______________________________________________________________________________________________________________________________&nbsp;<br /><img src="../images/pixel.gif" height="25" width="1" /><br />
                                &nbsp;Consignee Signature:________________________________ Print Name:_____________________________________ Date:________________________&nbsp;<br /><img src="../images/pixel.gif" height="20" width="1" /><br />
                            </td>
                        </tr>
                        <tr>
                            <td  class="FleetExpressTextBlackSmaller">
                                Property described above was received by consignee in good order, except as noted above.
                                <br /><img src="../images/pixel.gif" height="10" width="1" /><br />    
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
                                <br /><img src="../images/pixel.gif" height="10" width="1" /><br />    
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
                                Ref #:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="referenceNumber" size="20" maxlength="20" /><br /><br />
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
    
                <%end if 
                'Response.write "fh_carr_id="&fh_carr_id&"<BR>"
                %>
              <tr><td align="center"class="FleetExpressTextBlackBold" colspan="2">
              <form method="post">
                  <table width=95%>
                    <tr>
                         <td valign="bottom" align="left">
                                 Vehicle:&nbsp;&nbsp;
                                <!-------------------->
                                    <select name="Courier">
                                    <%
                                        'If trim(Courier)="" then Courier="Bobtail 4" End if
    									Set oConn = Server.CreateObject("ADODB.Connection")
    									oConn.ConnectionTimeout = 100
    									oConn.Provider = "MSDASQL"
    									oConn.Open DATABASE
    										l_cSQL = "Select VehicleID, VehicleName FROM AvailableVehicles WHERE AvailableStatus='c' ORDER BY VehicleName"
    										SET oRs = oConn.Execute(l_cSql)
    												Do While not oRs.EOF
                                                    VehicleID=oRs("VehicleID")
                                                    VehicleName=oRs("VehicleName")
    										    %>
    											<option value="<%=VehicleName%>" <%if trim(fh_carr_id)=trim(vehicleID) then response.Write " selected" end if%>><%=VehicleName%></option>
    											<%
    										oRs.movenext
    										LOOP
    									Set oConn=Nothing
                                        
                                    %>
                                    </select>
                                <!-------------------->

                      </td></tr>
                      <tr><td>
                                Ref #:&nbsp;&nbsp;<input type="text" name="referenceNumber" size="20" maxlength="20" /><br /><br />
                     </td><td>
                        <input type="hidden" name="jobnum" value="<%=jobnum%>"/>
                       <INPUT id="gobutton" name="gobutton" TYPE="submit" name="ButtonValue" VALUE="Dispatch Order">
                  <!-- <input type="submit" name="submit" value="Dispatch Order" />  -->
    
                      </td></tr>
                      <tr><td valign=top>
                                                      Special Instructions:&nbsp;&nbsp;<textarea name="addedcomments" cols="40" rows="5"></textarea>
                                <input type="hidden" name="comments" value="<%=comments %>" />
                      </TD>
                                    </form>
                      <td>
                                    <form method="post">
                                <INPUT id="gobutton" name="gobutton" TYPE="submit" name="ButtonValue" VALUE="Cancel Order">
                                <!-- <input type="submit" name="submit" value="Cancel Order" />  -->
              </form>
                      </td></tr>
                                <br /><br />
                                
                    </tr>
                  </table></td></tr>
                  <tr><td>
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


                <tr><td>&nbsp;</td></tr>
                <%if trim(ErrorMessage)>"" then %>
                <tr><td class="errormessage" align="center" colspan="2"><%=ErrorMessage %></td></tr>
                <%end if %>
                <%if trim(SuccessMessage)>"" then %>
                <tr><td class="successmessage" align="center" colspan="4"><font color="blue"><b><%=SuccessMessage %></b></font></td></tr>
                <%end if %>
                 <tr>
                    <td valign="top" align="left">
                        <table border="1" cellpadding="0" cellspacing="0"><tr><td>
                        <table border="0" cellpadding="0" cellspacing="0">
                            <tr>
                                <td colspan="5" align="center">
                                    <B>FLEET STATUS</B>
                                </td>
                            </tr>
                            <tr><td bgcolor="black" colspan="5"><img src="../Images/pixel.gif" height="1" width="1" /> </td></tr>
                                     <%
                                        'If trim(Courier)="" then Courier="Bobtail 4" End if
    									Set oConn = Server.CreateObject("ADODB.Connection")
    									oConn.ConnectionTimeout = 100
    									oConn.Provider = "MSDASQL"
    									oConn.Open DATABASE
    										l_cSQL = "Select un_id, un_desc FROM fcunits WHERE UnitStatus='c' and un_id<>'14' ORDER BY un_desc"
    										SET oRs = oConn.Execute(l_cSql)
    												Do While not oRs.EOF
                                                    'Response.write "got here!<BR>"
                                                    status_un_id=oRs("un_id")
                                                    Status_un_desc=oRs("un_desc")

                                                    '''''''''''''''''''''''''''''''''
    									            Set oConn77 = Server.CreateObject("ADODB.Connection")
    									            oConn77.ConnectionTimeout = 100
    									            oConn77.Provider = "MSDASQL"
    									            oConn77.Open DATABASE
    										            l_cSQL77 = "SELECT TOP 1 AvailableVehicles.VehicleID, AvailableVehicles.DriverID, lcintranet.dbo.Intranet_Users.FirstName, lcintranet.dbo.Intranet_Users.LastName FROM AvailableVehicles INNER JOIN lcintranet.dbo.Intranet_Users ON AvailableVehicles.DriverID = lcintranet.dbo.Intranet_Users.UserID WHERE AvailableStatus='c' and VehicleID='"&Status_un_id&"'"
    										            SET oRs77 = oConn77.Execute(l_cSQL77)
    												            If not oRs77.EOF then
                                                                    'Response.write "got here!<BR>"
                                                                    DriverName=" - "&oRs77("FirstName")&" "&oRs77("LastName")
                                                                    'Status_un_desc=oRs77("un_desc")
                                                                    Response.write "<tr><td><img src='..image/pixel.gif' height='1' width='5'></td><td><img src='../images/GreenLight.gif' height='20' width='20'></td><td><img src='..image/pixel.gif' height='1' width='5'></td>"
                                                                    else
                                                                    DriverName=""
                                                                    Response.write "<tr><td><img src='..image/pixel.gif' height='1' width='5'></td><td><img src='../images/RedLight.gif' height='20' width='20'></td><td><img src='..image/pixel.gif' height='1' width='5'></td>"
    										                    End if
    									            Set oConn77=Nothing
                                                    ''''''''''''''''''''''''''''''''''''''''
                                                    'response.write "l_cSQL77="&l_cSQL77&"<BR>" 



                                                    Response.write "<td>"&Status_un_desc&DriverName&"</td><td><img src='..image/pixel.gif' height='1' width='5'></td></TR><tr><td><img src='..image/pixel.gif' height='5' width='1'></td></tr>"
    										oRs.movenext
    										LOOP
    									Set oConn=Nothing
                                       'response.write "l_cSQL="&l_cSQL&"<BR>" 
                                    %>                                   
                        </table>
                        </td></tr></table>
                    </td>
                    <td valign="top"><img src="Images/pixel.gif" height="1" width="50"</td>
                    <td colspan="6" align="center">
                        <table width="95%">
                            <tr>
                                <td colspan="9" align="left">
                                    <B>OPEN ORDERS:  <%=ViewType%></B>
                                </td>
                                <form method="post" action="FleetXOrderDispatch.asp">
                                <td nowrap align="right" colspan="9" class="FleetExpressTextBlackBold">Job Number:&nbsp;&nbsp;<input type="text" value="<%=SearchJobNumber%>" name="SearchJobNumber" />
                                <input type="hidden" name="PageStatus" value="disp" />
                               <INPUT id="gobutton" name="gobutton" TYPE="submit" name="ButtonValue" VALUE="Find Job"> 
                                <input type="hidden" name="FindJob" value="y" />
                                </form>
                                <br><br><form method="post" action="FleetXOrderDispatch.asp">
                                Select View:&nbsp;&nbsp;
                                  <select name="ViewType">
                                  <option value="Today" <%if ViewType = "Today" then response.write " selected" end if %>>Jobs Due Today</option>
                                  <option value="Future" <%if ViewType = "Future" then response.write " selected" end if%>>Future Jobs</option>
                                  <option value="All" <%if ViewType = "All" then response.write " selected" end if%>>All Jobs</option>
                                </select>                      
                               <INPUT id="gobutton" name="gobutton" TYPE="submit" name="ButtonValue" VALUE="Submit"> 
                                </form>
                                </td>
                            </tr>
                          <form action="FleetXOrderDispatch.asp" method="post" name="DispatchOrders">
                           <tr><td></td></tr>
     
                            <%  
                                Today=now()
                                TargetDate=Today-4
                                'Response.write "today="&today&"<BR>"
                                'Response.write "TargetDate="&TargetDate&"<BR>"
    			                Set RSEVENTS2 = Server.CreateObject("ADODB.Connection")
    			                RSEVENTS2.ConnectionTimeout = 100
    			                RSEVENTS2.Provider = "MSDASQL"
    			                RSEVENTS2.Open DATABASE
    				                'l_cSQL = "SELECT fcfgthd.fh_id, fcfgthd.fh_user4, fcfgthd.fh_bt_id, fcfgthd.fh_ready, fclegs.fl_sf_name, fclegs.fl_st_name, fclegs.fl_st_rta, fcfgthd.fh_ship_dt, fcfgthd.fh_carr_id, fcfgthd.fh_status, fcrefs.rf_box FROM fcfgthd, fclegs, fcrefs WHERE fclegs.fl_fh_id = fcfgthd.fh_id AND fcfgthd.fh_id = fcrefs.rf_fh_id AND fh_status='RAP' or fh_status='SCD' " 

                            l_cSQL = "SELECT     fcfgthd.fh_id, fcfgthd.fh_user4, fcfgthd.fh_user6, fcfgthd.fh_bt_id, fcfgthd.fh_ready, fclegs.fl_sf_name, Fl_SF_ID, fl_sf_name, Fl_SF_Building, Fl_SF_addr1, Fl_SF_addr2, Fl_SF_City, fl_st_name, Fl_st_ID, fl_st_name, Fl_st_Building, Fl_st_addr1, Fl_st_addr2, Fl_st_City, fclegs.fl_st_name, fclegs.fl_sf_comment, fclegs.fl_st_comment, fclegs.fl_st_rta, fcfgthd.fh_ship_dt,  " &_
                                                  " fcfgthd.fh_carr_id, fcfgthd.fh_status, fcrefs.rf_box  " &_
                             "FROM         fcfgthd INNER JOIN " &_
                                                  " fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN  " &_
                                                  " fcrefs ON fcfgthd.fh_id = fcrefs.rf_fh_id " &_
                            " WHERE     (fcfgthd.fh_status = 'RAP' OR fcfgthd.fh_status = 'SCD') "  
                                                  
                                     'Response.write "fh_ready="&fh_ready&"<BR>" 
                                     'Response.write "VarX="&VarX&"<BR>"                                                                     
    				                'l_cSQL = "SELECT     fcfgthd.fh_id, fcfgthd.fh_user4, fclegs.fl_sf_name, fclegs.fl_st_name, fclegs.fl_st_rta, fcfgthd.fh_ship_dt, fcfgthd.fh_carr_id, fcfgthd.fh_status, fcrefs.rf_box FROM fclegs INNER JOIN fcfgthd ON fclegs.fl_fh_id = fcfgthd.fh_id INNER JOIN fcrefs ON fcfgthd.fh_id = fcrefs.rf_fh_id  " &_
    						                 '"WHERE (fh_status='CAN' and fh_ship_dt>'"&TargetDate&"') or fh_status='RAP' or fh_status='SCD' " 
    				                'Response.write "l_cSQL="&l_cSQL&"<BR>"
                                    'Response.write "Database="&Database&"<BR>"
          
                                    if ViewType = "All" then
                                    elseif ViewType = "Today" then
                                      'l_cSQL = l_cSQL & " AND (CONVERT(varchar, fclegs.fl_st_rta, 101) <= CONVERT(varchar, GETDATE(), 101))  AND DATEDIFF(minute,fcfgthd.fh_ready,GETDATE())>0"
                                      l_cSQL = l_cSQL & " AND (CONVERT(varchar, fclegs.fl_st_rta, 101) <= CONVERT(varchar, GETDATE(), 101))"
                                    elseif ViewType = "Future" then
                                      'l_cSQL = l_cSQL & " AND (CONVERT(varchar, fclegs.fl_st_rta, 101) > CONVERT(varchar, GETDATE(), 101))  AND DATEDIFF(minute,fcfgthd.fh_ready,GETDATE())>0" 
                                      'l_cSQL = l_cSQL & " AND (CONVERT(varchar, fclegs.fl_st_rta, 101) > CONVERT(varchar, GETDATE(), 101))  AND DATEDIFF(minute,fcfgthd.fh_ready,GETDATE())>0" 
                                      l_cSQL = l_cSQL & " AND (CONVERT(varchar, fclegs.fl_st_rta, 101) > CONVERT(varchar, GETDATE(), 101))" 
                                    end if
                                    'Response.write "l_cSQL="&l_cSQL&"<BR>"
                                    sqlorder = ""
                                    sortby = valid8(request.querystring("sortby"))
                                    psortby = valid8(request.querystring("psortby"))
                                   'if request.querystring("sortby")="job" then
                                      'l_cSQL=l_cSQL & " ORDER BY fcfgthd.fh_id"
                                    'else
                                    'if request.querystring("sortby")="cntd" then
                                      'sqlorder = " ORDER BY fclegs.fl_st_rta"
                                    'else
                                    'if request.querystring("sortby")="stat" then
                                      'sqlorder = " ORDER BY fcfgthd.fh_status"
                                    'else
                                    'if request.querystring("sortby")="bdate" then
                                      'sqlorder = " ORDER BY fcfgthd.fh_ship_dt"
                                    'else
                                    'if request.querystring("sortby")="sveh" then
                                      'sqlorder = " ORDER BY fcfgthd.fh_user4"
                                    'else
                                    'if request.querystring("sortby")="orig" then
                                      'sqlorder = " ORDER BY fclegs.fl_sf_name"
                                    'elseif request.querystring("sortby")="dest" then
                                      'sqlorder = " ORDER BY fclegs.fl_st_name"
                                    'else
                                     'sqlorder = " order by fh_id desc"
                                    'end if
                                    jobarrow = "down-arrow.png"
                                    jobsort = "job"
                                    cntdarrow = "down-arrow.png"
                                    cntdsort = "cntd"
                                    statarrow = "down-arrow.png"
                                    statsort = "stat"
                                    bdatearrow = "down-arrow.png"
                                    bdatesort = "bdate"
                                    sveharrow = "down-arrow.png"
                                    svehsort = "sveh"
                                    origarrow = "down-arrow.png"
                                    origsort = "orig"
                                    destarrow = "down-arrow.png"
                                    destsort = "dest"
                                    coarrow = "down-arrow.png"
                                    cosort = "co"
                                    Select Case sortby
                                      Case "job"
                                        if psortby = "job" then
                                          sortby = "jobu"
                                          sqlorder = " ORDER BY fcfgthd.fh_id DESC"
                                          jobarrow = "up-arrow-red.png"
                                          jobsort = "job"
                                        else
                                          sqlorder = " ORDER BY fcfgthd.fh_id"
                                          jobarrow = "down-arrow-red.png"
                                          jobsort = "jobu"
                                         end if
                                      Case "jobu"
                                        if psortby = "jobu" then
                                          sortby = "job"
                                          sqlorder = " ORDER BY fcfgthd.fh_id"
                                          jobarrow = "down-arrow-red.png"
                                         else
                                          sqlorder = " ORDER BY fcfgthd.fh_id DESC"
                                          jobarrow = "up-arrow-red.png"
                                         end if
                                      Case "cntd"
                                        if psortby = "cntd" then
                                          sortby = "cntdu"
                                          sqlorder = " ORDER BY fclegs.fl_st_rta DESC"
                                          cntdarrow = "up-arrow-red.png"
                                          cntdsort = "cntd"
                                        else
                                          sqlorder = " ORDER BY fclegs.fl_st_rta"
                                          cntdarrow = "down-arrow-red.png"
                                          cntdsort = "cntdu"
                                         end if
                                      Case "cntdu"
                                        if psortby = "cntdu" then
                                          sortby = "cntd"
                                          sqlorder = " ORDER BY fclegs.fl_st_rta"
                                          cntdarrow = "down-arrow-red.png"
                                         else
                                          sqlorder = " ORDER BY fclegs.fl_st_rta DESC"
                                          cntdarrow = "up-arrow-red.png"
                                         end if
                                      Case "stat"
                                        if psortby = "stat" then
                                          sortby = "statu"
                                          sqlorder = " ORDER BY fcfgthd.fh_status DESC"
                                          statarrow = "up-arrow-red.png"
                                          statsort = "stat"
                                        else
                                          sqlorder = " ORDER BY fcfgthd.fh_status"
                                          statarrow = "down-arrow-red.png"
                                          statsort = "statu"
                                         end if
                                      Case "statu"
                                        if psortby = "statu" then
                                          sortby = "stat"
                                          sqlorder = " ORDER BY fcfgthd.fh_status"
                                          statarrow = "down-arrow-red.png"
                                         else
                                          sqlorder = " ORDER BY fcfgthd.fh_status DESC"
                                          statarrow = "up-arrow-red.png"
                                         end if
                                      Case "bdate"
                                        if psortby = "bdate" then
                                          sortby = "bdateu"
                                          sqlorder = " ORDER BY fcfgthd.fh_ready DESC"
                                          bdatearrow = "up-arrow-red.png"
                                          bdatesort = "bdate"
                                        else
                                          sqlorder = " ORDER BY fcfgthd.fh_ready"
                                          bdatearrow = "down-arrow-red.png"
                                          bdatesort = "bdateu"
                                         end if
                                      Case "bdateu"
                                        if psortby = "bdateu" then
                                          sortby = "bdate"
                                          sqlorder = " ORDER BY fcfgthd.fh_ready"
                                          bdatearrow = "down-arrow-red.png"
                                         else
                                          sqlorder = " ORDER BY fcfgthd.fh_ready DESC"
                                          bdatearrow = "up-arrow-red.png"
                                         end if
                                      Case "sveh"
                                        if psortby = "sveh" then
                                          sortby = "svehu"
                                          sqlorder = " ORDER BY fcfgthd.fh_user4 DESC"
                                          sveharrow = "up-arrow-red.png"
                                          svehsort = "sveh"
                                        else
                                          sqlorder = " ORDER BY fcfgthd.fh_user4"
                                          sveharrow = "down-arrow-red.png"
                                          svehsort = "svehu"
                                         end if
                                      Case "svehu"
                                        if psortby = "svehu" then
                                          sortby = "sveh"
                                          sqlorder = " ORDER BY fcfgthd.fh_user4"
                                          sveharrow = "down-arrow-red.png"
                                         else
                                          sqlorder = " ORDER BY fcfgthd.fh_user4 DESC"
                                          sveharrow = "up-arrow-red.png"
                                         end if
                                      Case "orig"
                                        if psortby = "orig" then
                                          sortby = "origu"
                                          sqlorder = " ORDER BY fclegs.fl_sf_name DESC"
                                          origarrow = "up-arrow-red.png"
                                          origsort = "orig"
                                        else
                                          sqlorder = " ORDER BY fclegs.fl_sf_name"
                                          origarrow = "down-arrow-red.png"
                                          origsort = "origu"
                                         end if
                                      Case "origu"
                                        if psortby = "origu" then
                                          sortby = "orig"
                                          sqlorder = " ORDER BY fclegs.fl_sf_name"
                                          origarrow = "down-arrow-red.png"
                                         else
                                          sqlorder = " ORDER BY fclegs.fl_sf_name DESC"
                                          origarrow = "up-arrow-red.png"
                                         end if
                                      Case "dest"
                                        if psortby = "dest" then
                                          sortby = "destu"
                                          sqlorder = " ORDER BY fclegs.fl_st_name DESC"
                                          destarrow = "up-arrow-red.png"
                                          destsort = "dest"
                                        else
                                          sqlorder = " ORDER BY fclegs.fl_st_name"
                                          destarrow = "down-arrow-red.png"
                                          destsort = "destu"
                                         end if
                                      Case "destu"
                                        if psortby = "destu" then
                                          sortby = "dest"
                                          sqlorder = " ORDER BY fclegs.fl_st_name"
                                          destarrow = "down-arrow-red.png"
                                         else
                                          sqlorder = " ORDER BY fclegs.fl_st_name DESC"
                                          destarrow = "up-arrow-red.png"
                                         end if
                                      Case "co"
                                        if psortby = "co" then
                                          sortby = "cou"
                                          sqlorder = " ORDER BY fcfgthd.fh_bt_id DESC"
                                          coarrow = "up-arrow-red.png"
                                          cosort = "co"
                                        else
                                          sqlorder = " ORDER BY fcfgthd.fh_bt_id"
                                          coarrow = "down-arrow-red.png"
                                          cosort = "cou"
                                         end if
                                      Case "cou"
                                        if psortby = "cou" then
                                          sortby = "co"
                                          sqlorder = " ORDER BY fcfgthd.fh_bt_id"
                                          coarrow = "down-arrow-red.png"
                                         else
                                          sqlorder = " ORDER BY fcfgthd.fh_bt_id DESC"
                                          coarrow = "up-arrow-red.png"
                                         end if
                                    End Select
                                    psortby = sortby
                                    if autodispatch = "ON" then
                                        l_cSQL=l_cSQL&" and fh_bt_id='93' "
                                    End if
                                    l_cSQL=l_cSQL&sqlorder
         'response.write "2014 sql=" & l_cSQL & "<br>"
                                    SET oRs = RSEVENTS2.Execute(l_cSql)
                                    If oRs.eof then
                                        ErrorMessage="There are currently no open jobs"
                            End if
                           %><tr><td nowrap>&nbsp;</td><td width="20">&nbsp;</td><td nowrap><a style="text-decoration: none;" href="FleetXOrderDispatch.asp?sortby=<%=jobsort%>&ViewType=<%=ViewType%>"><font color="black"><b>JOB #</b></font>&nbsp;<img src="../images/<%=jobarrow%>" width=15 height=18></a></td><td width="20">&nbsp;</td><td nowrap><a style="text-decoration: none;" href="FleetXOrderDispatch.asp?sortby=<%=cosort%>&ViewType=<%=ViewType%>"><font color="black"><b>COMPANY</b></font>&nbsp;<img src="../images/<%=coarrow%>" width=15 height=18></a></td><td width="20">&nbsp;</td><td nowrap><a style="text-decoration: none;" href="FleetXOrderDispatch.asp?sortby=<%=cntdsort%>&ViewType=<%=ViewType%>"><font color="black"><b>COUNTDOWN</b></font>&nbsp;<img src="../images/<%=cntdarrow%>" width=15 height=18></a></td><!-- <td width="20">&nbsp;</td><td nowrap><a style="text-decoration: none;" href="FleetXOrderDispatch.asp?sortby=<%=statsort%>&ViewType=<%=ViewType%>"><font color="black"><b>STATUS</b></font>&nbsp;<img src="../images/<%=statarrow%>" width=15 height=18></a></td> --><td width="20">&nbsp;</td><td nowrap><a style="text-decoration: none;" href="FleetXOrderDispatch.asp?sortby=<%=bdatesort%>&ViewType=<%=ViewType%>"><font color="black"><b>READY DATE/TIME</b></font>&nbsp;<img src="../images/<%=bdatearrow%>" width=15 height=18></a></td><td width="20">&nbsp;</td><td nowrap><a style="text-decoration: none;" href="FleetXOrderDispatch.asp?sortby=<%=svehsort%>&ViewType=<%=ViewType%>"><font color="black"><b>VEHICLE</b></font>&nbsp;<img src="../images/<%=sveharrow%>" width=15 height=18></a></td><td width="20">&nbsp;</td><td nowrap><a style="text-decoration: none;" href="FleetXOrderDispatch.asp?sortby=<%=origsort%>&ViewType=<%=ViewType%>"><font color="black"><b>ORIGINATION</b></font>&nbsp;<img src="../images/<%=origarrow%>" width=15 height=18></a></td><td width="20">&nbsp;</td><td nowrap><a style="text-decoration: none;" href="FleetXOrderDispatch.asp?sortby=<%=destsort%>&ViewType=<%=ViewType%>"><font color="black"><b>DESTINATION</b></font>&nbsp;<img src="../images/<%=destarrow%>" width=15 height=18></a></td><!-- <td width="20">&nbsp;</td><td nowrap>ROUTED TO</td> --></tr>
                                  <tr><td colspan=20><hr></td></tr>
                            <%
    				                Do While not oRs.EOF	
                                        PageStatus="Edit"
                                        fh_id=Trim(oRs("fh_id"))
                                        fh_user4=Trim(oRs("fh_user4"))
                                        fh_user6=Trim(oRs("fh_user6"))
                                        fh_bt_id = Trim(oRs("fh_bt_id"))
                                        fh_ready = Trim(oRs("fh_ready"))
                                        fl_st_rta = Trim(oRs("fl_st_rta"))
                                        
                                       ' Tempvar=DATEDIFF(minute,fh_ready,GETDATE())
                                       ' Response.write "fl_st_rta="&fl_st_rta&"<BR>"
                                       ' Response.write "Tempvar="&Tempvar&"<BR>"
                                         'VarXXX=(CONVERT(VARCHAR, FH_READY, 109))
                                         'Response.write "VarXXX="&VarXXX&"<BR>"

                            Set oConn = Server.CreateObject("ADODB.Connection")
          									oConn.ConnectionTimeout = 100
          									oConn.Provider = "MSDASQL"
          									oConn.Open DATABASE

          									if len(fh_bt_id) > 0 then
                              l_cSQL2 = "Select bt_desc FROM fcbillto WHERE bt_id = " & fh_bt_id
          										SET oRst = oConn.Execute(l_cSql2)
                                        companyname = oRst("bt_desc")
                                    oRst.close
                                    Set oRst=Nothing
                                    Set oConn=Nothing
                            else
                              companyname = fh_bt_id & " - UNKNOWN"
                            end if

'fh_id=Trim(oRs("fh_id"))
'fh_user4=Trim(oRs("fh_user4"))
'fh_user6=Trim(oRs("fh_user6"))
'fh_bt_id=Trim(oRs("fh_bt_id"))
fh_ready=Trim(oRs("fh_ready"))
fl_sf_name=Trim(oRs("fl_sf_name"))
fl_SF_ID=Trim(oRs("fl_SF_ID"))
fl_sf_name=Trim(oRs("fl_sf_name"))
Fl_SF_Building=Trim(oRs("Fl_SF_Building"))
Fl_SF_addr1=Trim(oRs("Fl_SF_addr1"))
Fl_SF_addr2=Trim(oRs("Fl_SF_addr2"))
Fl_SF_City=Trim(oRs("Fl_SF_City"))
fl_st_name=Trim(oRs("fl_st_name"))
Fl_st_ID=Trim(oRs("Fl_st_ID"))
fl_st_name=Trim(oRs("fl_st_name"))
Fl_st_Building=Trim(oRs("Fl_st_Building"))
Fl_st_addr1=Trim(oRs("Fl_st_addr1"))
Fl_st_addr2=Trim(oRs("Fl_st_addr2"))
Fl_st_City=Trim(oRs("Fl_st_City"))
fl_st_name=Trim(oRs("fl_st_name"))
fl_sf_comment=Trim(oRs("fl_sf_comment"))
fl_st_comment=Trim(oRs("fl_st_comment"))
fl_st_rta=Trim(oRs("fl_st_rta"))
                                        duedate = fl_st_rta
                                        'response.write "2044 duedate=" & duedate &"<br>"
                                        duediff = DateDiff("n",Now(),duedate)
                                       if duediff < 0 then
                                          duediff = "<font color='red'><b>LATE</b></font>"
                                        else
                                          'duediff = duediff & " mins"
                                          duediff = datediffCNV(Now(),duedate)
                                        end if
fh_ship_dt=Trim(oRs("fh_ship_dt"))
fh_carr_id=Trim(oRs("fh_carr_id"))
fh_status=Trim(oRs("fh_status"))
rf_box=Trim(oRs("rf_box"))
                                        '''fl_sf_name=Trim(oRs("fl_sf_name"))
                                        '''Fl_SF_ID = oRs("Fl_SF_ID")
                                        '''fl_sf_name = oRs("fl_sf_name")
                                        '''Fl_SF_Building = oRs("Fl_SF_Building")
                                        '''Fl_SF_addr1 = oRs("Fl_SF_addr1")
                                        '''Fl_SF_addr2 = oRs("Fl_SF_addr2")
                                        '''Fl_SF_City = oRs("Fl_SF_City")
                                        '''fl_st_name=Trim(oRs("fl_st_name"))
                                        '''Fl_st_ID = oRs("Fl_st_ID")
                                        '''fl_st_name = oRs("fl_st_name")
                                        '''Fl_st_Building = oRs("Fl_st_Building")
                                        '''Fl_st_addr1 = oRs("Fl_st_addr1")
                                        '''Fl_st_addr2 = oRs("Fl_st_addr2")
                                        '''Fl_st_City = oRs("Fl_st_City")
                                        '''fl_sf_comment = oRs("fl_sf_comment")
                                        '''fl_st_comment = oRs("fl_st_comment")
                                        'Response.write "fl_sf_comment="&fl_sf_comment&"<BR>"
                                        'Response.write "fl_st_comment="&fl_st_comment&"<BR>"
                                        If trim(fl_sf_comment)>"" then
                                                    Comments=fl_sf_comment
                                                    'Response.write "Comments1="&Comments&"<BR>"
                                                End if
                                                If trim(fl_st_comment)>"" then
                                                'Response.write "Comments2="&Comments&"<BR>"
                                                    If trim(Comments)="" then
                                                        Comments=fl_st_comment
                                                        'Response.write "Comments3="&Comments&"<BR>"
                                                        else
                                                        Comments=Comments&"&nbsp;&nbsp;/&nbsp;&nbsp;"&fl_st_comment
                                                        'Response.write "Comments4="&Comments&"<BR>"
                                                    End if
                                                End if


                                        '''fh_ship_dt=Trim(oRs("fh_ship_dt"))
                                        '''fh_carr_id=Trim(oRs("fh_carr_id"))
                                        'Response.write "fh_carr_id="&fh_carr_id&"<BR>"
                                        '''fh_status=Trim(oRs("fh_status"))
                                        '''fh_ready=Trim(oRs("fh_ready"))
                                        'response.write "1977 ready=" & fh_ready & "<br>"
                                        '''rf_box=trim(oRs("rf_box"))
                                        If left(rf_box, 1)="X" then
                                            Xfont="orange"
                                        'elseif fh_status = "RAP" then
                                            'xfont="green"
                                        else
                                            xfont="black"
                                        End if
                                        If trim(fh_user6)="FleetX" then
                                            XFont="blue"
                                        End if
                                        Select Case fh_user4
                                        Case "Van"
                                          vehicleImage = "<img alt='Van' title='Van' width=52 height=25 src='../images/Icon_Van.jpg'>"
                                        Case "Bobtail"
                                          vehicleImage = "<img alt='Bobtail' title='Bobtail' width=66 height=35 src='../images/Icon_Bobtail.jpg'>"
                                        Case "Tractor"
                                          vehicleImage = "<img alt='Tractor Trailer' title='Tractor Trailer' width=85 height=35 src='../images/Icon_TractorTrailer.jpg'>"
                                        Case Else
                                          vehicleImage = " "
                                        End Select
                                    %>
                                   <tr><td valign="top"><input type="checkbox" name="selectedorder" value="<%=fh_ID%>"></td><td valign="top"></td><td nowrap valign="top"><!-- <a href="FleetXOrderDispatch.asp?SearchJobNumber=<%=fh_ID %>&PageStatus=disp&findjob=y&btid=86" class="FleetXRedMain"> --><%=right(fh_ID,5) %><!-- </a> --></td><td width="20" valign="top">&nbsp;</td><td nowrap valign="top"> <font color="<%=XFont %>"><%=companyname%>(<%=fh_bt_id %>)</font></td><td width="20" valign="top">&nbsp;</td><td nowrap valign="top"> <font color="<%=XFont %>"><%=duediff%></font></td><!-- <td width="20">&nbsp;</td><td nowrap> <font color="<%=XFont %>"><%=fh_status %></font></td> --><td width="20" valign="top">&nbsp;</td><td nowrap valign="top"> <font color="<%=XFont %>"><%=FormatDateTime(fh_ready,2) & " " & FormatDateTime(fh_ready,4) %></font></td><td width="20" valign="top">&nbsp;</td><td nowrap valign="top"> <font color="<%=XFont %>">&nbsp;<%=VehicleImage%></font></td><td width="20" valign="top">&nbsp;</td>
                                   <td nowrap valign="top"> <font color="<%=XFont %>">

                                   <%If fh_bt_id="91" then
                                        REsponse.write fl_sf_name
                                        else %>
                                   <%If trim(fl_sf_name)>"" then response.write fl_sf_name&"<BR>" end if%><%If trim(fl_sf_building)>"" then response.write fl_sf_Building&"<BR>" end if%><%If trim(fl_sf_addr1)>"" then response.write fl_sf_addr1&"<BR>" End if%><%If trim(fl_sf_addr2)>"" then response.write fl_sf_addr2&"<BR>"%><%If trim(fl_sf_city)>"" then response.write fl_sf_city
                                   End if
                                   %>
                                   
                                   </font></td><td width="20" valign="top">&nbsp;</td>
                                   <td valign="top"> <font color="<%=XFont %>">
                                   
                                    <%If fh_bt_id="91" then
                                        REsponse.write fl_st_name
                                        else %>
                                   <%If trim(fl_st_name)>"" then response.write fl_st_name&"<BR>" end if%><%If trim(fl_st_building)>"" then response.write fl_st_Building&"<BR>" end if%><%If trim(fl_st_addr1)>"" then response.write fl_st_addr1&"<BR>" End if%><%If trim(fl_st_addr2)>"" then response.write fl_st_addr2&"<BR>"%><%If trim(fl_st_city)>"" then response.write fl_st_city
                                   End if
                                   %>
                                   
                                   
                                   
                                   
                                   </font></td><!-- <td width="20">&nbsp;</td><td> <font color="<%=XFont %>"><%=fh_carr_id %></font></td> --></tr></font>
                                    <%If trim(Comments)>"" then %>
                                   <tr><td colspan="3">&nbsp;</td><td colspan="8"><b>Special Instructions:  </b><%=comments %></td></tr>
                                    <%
                                    End if                                  
                                   %>
                                   <tr><td colspan=20 valign="top"><hr></td></tr>
                                   <%
                                   Comments=""
    								oRs.movenext
    								LOOP
                                    oRs.close
                                    Set oRs=Nothing
                                RSEVENTS2.Close
    			                Set RSEVENTS2=Nothing
                        %>
                        <tr><td colspan=11>&nbsp;<br>
                        Route to Vehicle:<br>
                        <select name="RouteVehicle">
                            <option value="">Select a vehicle/driver</option>
                            <% 
                            Set oConn = Server.CreateObject("ADODB.Connection")
          									oConn.ConnectionTimeout = 100
          									oConn.Provider = "MSDASQL"
          									oConn.Open DATABASE
          										l_cSQL = "select VehicleName, VehicleType, DriverID, LoginTime from AvailableVehicles where DriverID IS NOT NULL and LoginTime IS NOT NULL and AvailableStatus = 'c'"
          										SET oRs = oConn.Execute(l_cSql)
                                                        If oRs.EOF then
                                                            NoVehicles="y"
                                                            'Response.write "HELLo!<BR>Line2148"
                                                        End if
          												Do While not oRs.EOF
                                         Set oConnA = Server.CreateObject("ADODB.Connection")
                        									oConnA.ConnectionTimeout = 100
                        									oConnA.Provider = "MSDASQL"
                        									oConnA.Open DATABASE
                                          'check that the vehicle is not assigned to a job
                                          iuSQL = "select fc.fh_id, fl.fl_un_id, fl.fl_st_rta from fcfgthd fc, fclegs fl where fc.fh_status <> 'CAN' and fc.fh_status <> 'CLS' and fl.fl_fh_id = fc.fh_id and fl.fl_un_id = '" & oRs("VehicleName") & "'"
                                          SET oRsa2 = oConnA.Execute(iuSql)
                       									  vavail = "n"
                                          ' if no match found, then vehicle is unassigned and available
                                          if oRsa2.EOF then
                                            vavail = "y"
                                          else
                                            'if vehicle is assigned, check to see if the due date is less than 10 minutes away, if so then mark vehicle as available
                                            mindiff = datediff("n", oRsa2("fl_st_rta"), Now())
                                            'response.write "1711 minutes = " & mindiff & "<br>" 
                                            if mindiff < 10 then
                                              vavail = "y"
                                              ''''''''''''''''''added because of problems'''''MARK
                                              else
                                              vavail = "y"
                                            end if
                                          end if
                                          oRsa2.close
                                          Set oConnA=Nothing
 
                                          Set oConn777 = Server.CreateObject("ADODB.Connection")
                        									oConn777.ConnectionTimeout = 100
                        									oConn777.Provider = "MSDASQL"
                        									oConn777.Open INTRANET
                        										SQL777 = "Select UserID, FirstName, LastName FROM intranet_users WHERE UserID = " & trim(oRs("DriverID"))
                        										'response.write "1702 sql=" & iuSQL & "<br>"
                                            SET oRs777 = oConn777.Execute(Sql777)
                                                If oRs777.eof then

                                                    else
                                                    DriverName= oRs777("FirstName") & " " & oRs777("LastName")
                                                    DriverUserID= oRs777("UserID")
                                                    Select Case DriverUserID
                                                        Case 604,82,500
                                                            DriverFontColor="red"
                                                        Case 390,212
                                                            DriverFontColor="blue"
                                                        Case 609,603,460,624,317,627
                                                            DriverFontColor="green"
                                                        Case else
                                                            DriverFontColor="black"
                                                    End Select
                                              End if
                       							
                                                oRs777.close
                                                Set oConn777=Nothing
                                               
                                  if vavail = "y" then  
                                       VehicleName=oRs("VehicleName")
                                       VehicleType=oRs("VehicleType")
          										    %>
          											<option style="color:<%=DriverFontColor%>" value="<%=VehicleName%>" <%if trim(VehicleName)=trim(Courier) then response.Write " selected" end if%>><%=VehicleName%> (<%=VehicleType%>) - <%=DriverName%></option>
          											<%
                                 end if
          										oRs.movenext
          										LOOP
          									Set oConn=Nothing
                                          %>
                        </select>
                        &nbsp;
                        <%
                        if NoVehicles="y" then
                            Response.write "There are currently no drivers logged into any vehicles"
                        End if
                        %><br><br>
                        SPECIAL INSTRUCTIONS:<br>
                        <textarea name="specialInstructions" cols=30 rows=5></textarea>
                        &nbsp;<br><br>
                        <INPUT id="gobutton" name="gobutton" TYPE="submit" name="ButtonValue" VALUE="DISPATCH SELECTED ORDERS"></td></tr>
                      </form>
                        </table>
                    </td>
                </tr>
    <%end if %>
    <tr><td>&nbsp;</td></tr>
    <tr><td>&nbsp;</td></tr>
    <tr><td nowrap>&nbsp;</td><td nowrap>&nbsp;</td><td align="center">
                <table><tr><td>
               <form method="post" action="FleetXOrderEdit.asp?BTID=86">
                                <input type="hidden" name="SearchJobNumber" value="<%=jobnum%>" />
                                <INPUT id="gobutton" name="gobutton" TYPE="submit" name="ButtonValue" VALUE="Edit Order">
                                <!-- <input type="submit" name="submit" value="Edit Order" /> -->
                                <input type="hidden" name="btid" value="86" />
               </form>
               </td>

               <td>
              <form method="post" action="FreightOrder.asp">
                                <INPUT id="gobutton" name="gobutton" TYPE="submit" name="ButtonValue" VALUE="Enter Order">
                                <!-- <input type="submit" name="submit" value="Enter Order" />   -->
                                <input type="hidden" name="btid" value="86" />
              </form>
              </td>
               <td>
              <form method="post" action="FleetXOrderDispatch.asp?BTID=86">
                                <INPUT id="gobutton" name="gobutton" TYPE="submit" name="ButtonValue" VALUE="Return to Dispatch">
                               <!-- <input type="submit" name="submit" value="Return to Dispatch" />   -->
                                <input type="hidden" name="btid" value="86" />
              </form>
             </td>
               <td>
              <form method="post" action="../home.asp">
                                <input type="hidden" name="SearchJobNumber" value="<%=jobnum%>" />
                                <INPUT id="gobutton" name="gobutton" TYPE="submit" name="ButtonValue" VALUE="Return to Home">
                                <!-- <input type="submit" name="submit" value="Return to Home" /> -->
                                <input type="hidden" name="btid" value="86" />
               </form>
              <!-- </td>  

              </tr></table> -->
   <!-- </td></tr>
    </table>   --> 
    
    
    
 
    
    
    
 <% 'end if %>  
    
    
    </td></tr>



 
	<tr Height="280">
		<td>&nbsp;</td>
	</tr>

 

</table>
</td></tr>



<%
if ErrorMessage>"" then%>
<tr><td colspan="3">
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
<!-- </form>  -->
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
