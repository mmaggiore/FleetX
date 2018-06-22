<html>
<head>
   <%
   '''''''''''HARDCODED STUFF
   'sBT_ID=Request.QueryString("BTID")
   sBT_ID="86"
   If sBT_ID>"" then
        Session("sBT_ID")=sBT_ID
   End if
   If Session("sBT_ID")<>"86" then
        'Response.redirect("https://www.logisticorp.us/intranet/")
        Session("sBT_ID")="86"
   End if
    UserID=Request.Form("UserID")
    If trim(userID)="" then
        UserID=Session("UserID")
    End if
    If trim(UserID)>"" then
        Session("UserID")=UserID
    End if
    sBT_ID = ""
  'Response.write "UserID="&UserID&"<BR>"
   ''''''''''''''''''''''''''
    %>
<!-- #include file="../fleetexpress.inc" -->
<link rel="stylesheet" type="text/css" href="../css/Style.css">
<script language="javascript" type="text/javascript" src="datetimepicker.js">

    //Date Time Picker script- by TengYong Ng of http://www.rainforestnet.com
    //Script featured on JavaScript Kit (http://www.javascriptkit.com)
    //For this script, visit http://www.javascriptkit.com
    function onPageLoad() {
        if (document.form.RequestorName.value.length == 0) {
            document.form.RequestorName.focus();
        }
    }
</script>
    <%
  
    TableWidth="460"
    PageStatus=Request.form("PageStatus")
    If trim(PageStatus)="" then
        PageStatus=request.QueryString("PageStatus")
    End if
    SearchJobNumber=Request.Form("SearchJobNumber")
    If trim(SearchJobNumber)="" then
        SearchJobNumber=request.QueryString("SearchJobNumber")
    End if
    RequestorName=Request.form("RequestorName")
    RequestorPhoneNumber=Request.form("RequestorPhoneNumber")
    RequestorEmailAddress=Request.form("RequestorEmailAddress")
    'PONumber=Request.form("PONumber")
    'CostCenterNumber=Request.form("CostCenterNumber")
    Pieces=Request.form("Pieces")
    rf_box=Request.form("rf_box")
    NumberOfPallets=Request.form("NumberOfPallets")
    DimWeight=Request.form("DimWeight")
    DimLength=Request.form("DimLength")
    DimWidth=Request.form("DimWidth")
    DimHeight=Request.form("DimHeight")
    IsPalletized=Request.form("IsPalletized")
    DimValue=Request.form("DimValue")
    IsHazmat=Request.form("IsHazmat")
    OriginationCompany=Request.form("OriginationCompany")
    OriginationAddress=Request.form("OriginationAddress")
    OriginationCity=Request.form("OriginationCity")
    OriginationState=Request.form("OriginationState")
    'Response.write "***OriginationState="&OriginationState&"<BR>"
    OriginationZipCode=Request.form("OriginationZipCode")
    OriginationContactName=Request.form("OriginationContactName")
    OriginationPhoneNumber=Request.form("OriginationPhoneNumber")
    OriginationEmail=Request.form("OriginationEmail")
    DestinationCompany=Request.form("DestinationCompany")
    DestinationAddress=Request.form("DestinationAddress")
    DestinationCity=Request.form("DestinationCity")
    DestinationState=Request.form("DestinationState")
    DestinationZipCode=Request.form("DestinationZipCode")
    DestinationContactName=Request.form("DestinationContactName")
    DestinationPhoneNumber=Request.form("DestinationPhoneNumber")
    DestinationEmail=Request.form("DestinationEmail")
    Courier=Request.form("courier")
    POorNWA=Request.form("POorNWA")
    GenericNumber=Request.form("GenericNumber")
    'Response.write "POorNWA="&POorNWA&"<BR>"
    Select Case POorNWA
        Case "TI P/O #"
            PONumber=GenericNumber
        Case "Cost Center #"
            CostCenterNumber=GenericNumber
        Case else
            CostCenterNumber=GenericNumber
    End Select
    'Response.write "PONumber="&PONumber&"<BR>"
    'Response.write "CostCenterNumber="&CostCenterNumber&"<BR>"
    Comments=Request.form("Comments")
    RequestorName=Replace(RequestorName, """", "`")
    RequestorName=Replace(RequestorName, "'", "`")
    RequestorPhoneNumber=Replace(RequestorPhoneNumber, """", "")
    RequestorPhoneNumber=Replace(RequestorPhoneNumber, "'", "")
    RequestorEmailAddress=Replace(RequestorEmailAddress, """", "")
    RequestorEmailAddress=Replace(RequestorEmailAddress, "'", "")
    Pieces=Replace(Pieces, """", "")
    Pieces=Replace(Pieces, "'", "")
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
    DeliveryDateTime=Request.form("DeliveryDateTime")
    

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
    If trim(SearchJobNumber)>"" and PageStatus<>"submit" then
			Set RSEVENTS2 = Server.CreateObject("ADODB.Connection")
			RSEVENTS2.ConnectionTimeout = 100
			RSEVENTS2.Provider = "MSDASQL"
			RSEVENTS2.Open DATABASE
				l_cSQL = "select fh_status, fh_ship_dt, fh_ready, fh_priority, fh_bt_id, fh_co_id, fh_co_phone, fh_co_email, fh_co_costcenter, fh_custpo, fh_carr_id from fcfgthd  " &_
						 "WHERE fh_id = '" & TRIM(SearchJobNumber)&"'" 
				'Response.write "l_cSQL="&l_cSQL&"<BR>"
                'Response.write "Database="&Database&"<BR>"
                SET oRs = RSEVENTS2.Execute(l_cSql)
				IF not oRs.EOF then	
                    PageStatus="Edit"
                    fh_status=Trim(oRs("fh_Status"))
                    fh_ship_dt=Trim(oRs("fh_ship_dt"))
                    PickUpDateTime=Trim(oRs("fh_ready"))
                    Priority=Trim(oRs("Fh_Priority"))
                    sBT_ID=Trim(oRs("fh_bt_ID"))
                    RequestorName=Trim(oRs("fh_co_id"))
                    RequestorPhoneNumber=Trim(oRs("fh_co_phone"))
                    RequestoremailAddress=Trim(oRs("fh_co_email"))
                    costcenterNumber=Trim(oRs("fh_co_costcenter"))
                    
                    'Response.write "costcenterNumber="&costcenterNumber&"<BR>"
                    If trim(costcenternumber)>"" then
                        genericnumber=costcenternumber
                        POorNWA="Cost Center #"
                    End if
                    PoNumber=Trim(oRs("fh_custpo"))
                    'Response.write "PoNumber="&PoNumber&"<BR>"
                    If trim(PoNumber)>"" then
                        genericnumber=PoNumber
                        POorNWA="TI P/O #"
                    End if
                    'response.write "POorNWA="&POorNWA&"<BR>"
                    courier=Trim(oRs("fh_carr_id"))
                    Else
                    ErrorMessage="That is not a valid job number"
				End if
            RSEVENTS2.Close
			Set RSEVENTS2=Nothing
            
			Set RSEVENTS2 = Server.CreateObject("ADODB.Connection")
			RSEVENTS2.ConnectionTimeout = 100
			RSEVENTS2.Provider = "MSDASQL"
			RSEVENTS2.Open DATABASE
				l_cSQL = "select fl_sf_name, fl_sf_clname, fl_sf_phone, fl_sf_email, fl_sf_addr1, fl_sf_city, fl_sf_state, fl_sf_country, fl_sf_zip, fl_st_name, fl_st_clname, fl_st_phone, fl_st_email, fl_st_addr1, fl_st_city, fl_st_state, fl_st_country, fl_st_zip, fl_sf_comment, fl_st_rta from fclegs  " &_
						 "WHERE fl_fh_id = '" & TRIM(SearchJobNumber)&"'" 
				SET oRs = RSEVENTS2.Execute(l_cSql)
				IF not oRs.EOF then	
                OriginationCompany=Trim(oRs("fl_sf_name"))
                OriginationContactName=Trim(oRs("fl_sf_clname"))
                OriginationPhoneNumber=Trim(oRs("fl_sf_phone"))
                OriginationEmail=Trim(oRs("fl_sf_email"))
                OriginationAddress=Trim(oRs("fl_sf_addr1"))
                OriginationCity=Trim(oRs("fl_sf_city"))
                OriginationState=Trim(oRs("fl_sf_state"))
                fl_sf_country=Trim(oRs("fl_sf_country"))
                OriginationZipCode=Trim(oRs("fl_sf_zip"))
                DestinationCompany=Trim(oRs("fl_st_name"))
                DestinationContactName=Trim(oRs("fl_st_clname"))
                DestinationPhoneNumber=Trim(oRs("fl_st_phone"))
                DestinationEmail=Trim(oRs("fl_st_email"))
                DestinationAddress=Trim(oRs("fl_st_addr1"))
                DestinationCity=Trim(oRs("fl_st_city"))
                DestinationState=Trim(oRs("fl_st_state"))
                fl_st_country=Trim(oRs("fl_st_country"))
                DestinationZipCode=Trim(oRs("fl_st_zip"))
                Comments=Trim(oRs("fl_sf_comment"))
                DeliveryDateTime=Trim(oRs("fl_st_rta"))
                
				End if
            RSEVENTS2.Close
			Set RSEVENTS2=Nothing    
            


			Set RSEVENTS2 = Server.CreateObject("ADODB.Connection")
			RSEVENTS2.ConnectionTimeout = 100
			RSEVENTS2.Provider = "MSDASQL"
			RSEVENTS2.Open DATABASE
				l_cSQL = "select rf_box, NumberOfPieces, IsPalletized, NumberOfPallets, Weight, DimLength, DimWidth, DimHeight, MeasurementType, Hazmat, Refrigerate from fcrefs  " &_
						 "WHERE rf_fh_id = '" & TRIM(SearchJobNumber)&"'" 
				SET oRs = RSEVENTS2.Execute(l_cSql)
				IF not oRs.EOF then
                	
                rf_box=Trim(oRs("rf_box"))
                Pieces=Trim(oRs("NumberOfPieces"))
                IsPalletized=Trim(oRs("IsPalletized"))
                NumberOfPallets=Trim(oRs("NumberOfPallets"))
                DimWeight=Trim(oRs("Weight"))
                DimLength=Trim(oRs("DimLength"))
                DimWidth=Trim(oRs("DimWidth"))
                DimHeight=Trim(oRs("DimHeight"))
                MeasurementType=Trim(oRs("MeasurementType"))
                isHazmat=Trim(oRs("Hazmat"))
                Refrigerate=Trim(oRs("Refrigerate"))
				End if
            RSEVENTS2.Close
			Set RSEVENTS2=Nothing                       	
    End if
    ''''''''ERROR HANDLING''''''''''
    If PageStatus="submit" then
        PageStatus="Edit"
        'If trim(DestinationEmail)="" then
        '    ErrorMessage="You must provide the Destination's Email"
       ' End if
       If not isdate(DeliveryDateTime) then DeliveryDateTime=now() end if
        If not isdate(PickUPDateTime) then PickUPDateTime=now() end if
        'If cdate(PickUpDateTime)<now() then PickUpDateTime=Now() end if
       ' Response.write "PickUpDateTime="&PickUpDateTime&"<BR>"
       ' Response.write "CurrentDateTime="&CurrentDateTime&"<BR>"
        'If isdate(DeliveryDateTime)  and cdate(CurrentDateTime)>=cdate(DeliveryDateTime) then
        '    ErrorMessage="The delivery date/time cannot be before the current date/time"
        'End if
        If isdate(DeliveryDateTime) and isdate(PickUPDateTime) and cdate(PickUPDateTime)>=cdate(DeliveryDateTime) then
            ErrorMessage="The delivery date/time cannot be before the pick up date/time"
        End if
        If NOT isdate(DeliveryDateTime) then
            ErrorMessage="You must provide a valid destination date/time"
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
        'If isdate(PickUpDateTime)  and cdate(CurrentDateTime)>cdate(PickUpDateTime) then
        '    ErrorMessage="The pick up time cannot be before the current date/time"
        'End if
        If NOT isdate(PickUpDateTime) then
            ErrorMessage="You must provide a valid pick up date/time"
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
        If trim(CostCenterNumber)="" AND trim(PONumber)="" then
            ErrorMessage="You must provide the NWA, Cost Center Number, or P/O Number"
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
        If trim(ErrorMessage)="" then
            'Response.write "Database="&Database&"<BR>"
		    'If trim(st_id)="DNP" or trim(st_id)="CPGPSCOT" then
            'Response.write "Database="&database&"<BR>"
            DeliveryPeriod=DateDiff("h", PickUpDateTime, DeliveryDateTime)
            'Response.write "DeliveryPeriod="&DeliveryPeriod&"<br>"

                         Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open INTRANET
						' 7/26/04 KK: Added canceljob functionality to be able to update the status if cancel button is pressed.
						' 11/30/04 DEC: Changed from CAN/98 TO DEL/99 to be consistent with dispatchOffice			
							l_cSQL = "UPDATE FleetRouted SET status = 'x' WHERE fh_id = '" & SearchJobNumber & "'"
							'Response.write "l_cSQL="&l_cSQL&"<BR>"
							oConn.Execute(l_cSQL)
						Set oConn=Nothing
  			            Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				            RSEVENTS2.Open "FleetRouted", Intranet, 2, 2
				            RSEVENTS2.addnew
                            RSEVENTS2("fh_id")=SearchJobNumber
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
                            RSEVENTS2("fh_id")=SearchJobNumber
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
				    l_cSQL = "UPDATE fcfgthd SET fh_ready='"&PickUpDateTime&"', Fh_Priority='"&Priority&"', fh_lastchg='"&now()&"', fh_bt_ID='"&Trim(sBT_ID)&"', fh_co_id='"&Trim(RequestorName)&"', fh_co_phone='"&Trim(RequestorPhoneNumber)&"', fh_co_email='"&Trim(RequestoremailAddress)&"', fh_co_costcenter='"&Trim(costcenterNumber)&"', fh_custpo='"&Trim(PoNumber)&"'" 
                    l_cSQL = l_cSQL&" WHERE (fh_id = '"& TRIM(SearchJobNumber) &"')"
				    Response.write "l_cSQL="&l_cSQL&"<BR>"
                    oConn.Execute(l_cSQL)
			    oConn.close
			    Set oConn = nothing          
  			   
               
                Set oConn = Server.CreateObject("ADODB.Connection")
			    oConn.ConnectionTimeout = 100
			    oConn.Provider = "MSDASQL"
			    oConn.Open DATABASE
				    l_cSQL = "UPDATE fclegs SET fl_sf_name='"&Trim(OriginationCompany)&"', fl_sf_clname='"&Trim(OriginationContactName)&"', fl_sf_phone='"&Trim(OriginationPhoneNumber)&"', fl_sf_email='"&Trim(OriginationEmail)&"', fl_sf_addr1='"&Trim(OriginationAddress)&"', fl_sf_city='"&Trim(OriginationCity)&"', fl_sf_state='"&Trim(OriginationState)&"', fl_sf_zip='"&Trim(OriginationZipCode)&"', fl_st_name='"&Trim(DestinationCompany)&"', fl_st_clname='"&Trim(DestinationContactName)&"', fl_st_phone='"&Trim(DestinationPhoneNumber)&"', fl_st_email='"&Trim(DestinationEmail)&"', fl_st_addr1='"&Trim(DestinationAddress)&"', fl_st_city='"&Trim(DestinationCity)&"', fl_st_state='"&Trim(DestinationState)&"', fl_st_zip='"&Trim(DestinationZipCode)&"', fl_sf_comment='"&Trim(Comments)&"', fl_st_rta='"&Trim(DeliveryDateTime)&"'" 
                    l_cSQL = l_cSQL&" WHERE (fl_fh_id = '"& TRIM(SearchJobNumber) &"')"
				    'Response.write "l_cSQL="&l_cSQL&"<BR>"
                    oConn.Execute(l_cSQL)
			    oConn.close
			    Set oConn = nothing  
                
                
                Set oConn = Server.CreateObject("ADODB.Connection")
			    oConn.ConnectionTimeout = 100
			    oConn.Provider = "MSDASQL"
			    oConn.Open DATABASE
				    l_cSQL = "UPDATE fcrefs SET rf_box='"&Trim(rf_box)&"', NumberOfPieces='"&Trim(Pieces)&"', IsPalletized='"&Trim(IsPalletized)&"', NumberOfPallets='"&Trim(NumberOfPallets)&"', Weight='"&Trim(DimWeight)&"', DimLength='"&Trim(DimLength)&"', DimWidth='"&Trim(DimWidth)&"', DimHeight='"&Trim(DimHeight)&"', MeasurementType='"&Trim(MeasurementType)&"', Hazmat='"&Trim(isHazmat)&"', Refrigerate='"&Trim(Refrigerate)&"'" 
                    l_cSQL = l_cSQL&" WHERE (rf_fh_id = '"& TRIM(SearchJobNumber) &"')"
				    'Response.write "l_cSQL="&l_cSQL&"<BR>"
                    oConn.Execute(l_cSQL)
			    oConn.close
			    Set oConn = nothing 

				Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
					RSEVENTS2.Open "JobChanges", DATABASE, 2, 2
					RSEVENTS2.addnew
					RSEVENTS2("fh_ID")=SearchJobNumber
					RSEVENTS2("SupervisorID")=UserID									
					RSEVENTS2("ChangeDate")=now()		
					RSEVENTS2("ChangeStatus") = "c"
					RSEVENTS2.update
					RSEVENTS2.close			
				set RSEVENTS2 = nothing
  
                'Response.Write "YAY!!!! You're DONE!!!!<BR>"
		
         
            'Response.write "newjobnum="&newjobnum&"<BR>"
  ''''''''''''''''LETTER to Requestor'''''''''''''''''''
   				    Body = "There have been changes made to your shipment request #"& SearchJobNumber &":<br><br>"   

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
                        Body = Body & "COMMENTS:<br>" 
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
				    objMail.cc = "mark.maggiore@logisticorp.us;fleetexpress@LogistiCorp.us"
				    'objMail.cc = "4692269939@tmomail.net;pcurrin@ti.com"
				    objMail.Subject = "Changes to your Fleet Express shipment request"
				    objMail.MailFormat = cdoMailFormatMIME
				    objMail.BodyFormat = cdoBodyFormatHTML
				    objMail.Body = Body
				    objMail.Send
				    Set objMail = Nothing         
            

                    '''Response.Redirect("GenericOrderConfirmation.asp?bid=84&pid=view&jid="& newjobnum)	
                    PageStatus="Success"
                    SuccessMessage="You have successfully updated job #"&SearchJobNumber&"<br>"
                    SearchJobNumber=""
		    'End if	
        End if
    End if
    ''''''''END ERROR HANDLING''''''

     %>
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
    PageTitle="ORDER EDIT"

%>


<title>FleetX - <%=PageTitle %></title>
</head>

<body onload="document.FindUser.requestorName.focus();document.OrderForm1.<%=HighlightedField%>.focus();">
	<table align="center" border="0" bordercolor="black" cellpadding="0" cellspacing="0" ID="Table1" height="100%" width="100%">
        <tr><td align="left"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>
        <tr>
            <td align="left"><img src="../images/FleetX_Small.jpg" height="50" width="168" /></td>
            <td align="right" valign="bottom"><!-- #include file="../topnavbar.asp" --></td>
        </tr>
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" /></td></tr>	
    <tr><td colspan="2"> 
<!-- <form action="NewUser.asp" method="post" name="FindUser"> -->
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
    
    
      <form method="post" name="OrderForm1" action="FleetXOrderEdit.asp">
          <table border="0" cellpadding="0" cellspacing="0" align="center" width="900">
              <tr><td align="left"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>
              <!--
              <tr>
                  <td align="left"><img src="../images/logo_raytheon_space.gif" height="54" width="207" /></td>
                  <td align="right" valign="bottom"><a href="mailto:mark.maggiore@logisticorp.us" class="<%=LinkClass%>">Click here to report a problem with this page</a></td>
              </tr>
              -->
              <tr><td align="left" colspan="2"><img src="../images/pixel.gif" /></td></tr>
               <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="RaytheonHeaderWhite" colspan="2">Edit A FleetX Transportation Request<%If trim(SearchJobNumber)>"" then%>-Job #<%=SearchJobNumber %><%End if %></td></tr>
              <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="RaytheonHeaderWhite" colspan="2"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>      
              <%Select Case PageStatus
              Case "Edit" %>        
              
              
              
      
              <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="RaytheonBodyWhite" colspan="2">Please complete all areas below</td></tr>
              <tr><td align="left" colspan="2"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>
              <tr>
                  <td align="center" colspan="2">
                  <table border="0" bordercolor="<%=BorderColor%>" cellpadding="0" cellspacing="0">
                  <tr> <td valign="top"><table border="1" bordercolor="<%=BorderColor%>" cellpadding="0" cellspacing="0" width="<%=tablewidth%>">
                      <tr> <td valign="top"><table cellpadding="3" cellspacing="0" width="100%">
                     
                     
                     
                     
                  
                     
                      <tr>
                          <td colspan="3" align="center" bgcolor="<%=BorderColor%>" class="RaytheonBodyWhiteBold">
                              REQUESTOR INFORMATION
                          </td>
                      </tr>
                      <tr>
                          <td class="RaytheonTextBlackBold" align="left">Requestor Name</td>
                          <td width="20"><img src="../images/pixel.gif" /></td>
                          <td align="left"><input type="text" name="RequestorName" value="<%=RequestorName%>" size="40" maxlength="30" /></td>
                      </tr>
                      <tr>
                          <td class="RaytheonTextBlackBold" align="left">Phone Number</td>
                          <td width="20"><img src="../images/pixel.gif" /></td>
                          <td align="left"><input type="text" name="RequestorPhoneNumber" value="<%=RequestorPhoneNumber%>" size="40" maxlength="20" /></td>
                      </tr>
                      <tr>
                          <td class="RaytheonTextBlackBold" align="left">Email Address</td>
                          <td width="20"><img src="../images/pixel.gif" /></td>
                          <td align="left"><input type="text" name="RequestorEmailAddress" value="<%=RequestorEmailAddress%>" size="40" maxlength="100" /></td>
                      </tr>
                      <tr>
                          <td class="RaytheonTextBlackBold" align="left">
                                  <select name="POorNWA">
                                      <option value="Cost Center #"<%If trim(POorNWA)="Cost Center #" then Response.write "selected" end if%>>Cost Center #</option>
                                      <option value="TI P/O #"<%If trim(POorNWA)="TI P/O #" then Response.write "selected" end if%>>TI P/O #</option>
                                  </select>
                          </td>
                          <td width="20"><img src="../images/pixel.gif" /></td>
                          <td align="left"><textarea name="GenericNumber" rows="2" cols="30"><%=GenericNumber%></textarea></td>
                      </tr>
                      <tr>
                          <td class="RaytheonTextBlackBold" align="left" valign="top">Comments</td>
                          <td width="20"><img src="../images/pixel.gif" /></td>
                          <td align="left"><textarea name="comments" rows="2" cols="30"><%=Comments%></textarea></td>
                      </tr>
                      <tr><td><img src="../images/pixel.gif" width="1" height="3" /></td></tr>
                      </table>
                       </td></tr></table></td>
                       <td align="left"><img src="../images/pixel.gif" height="1" width="25" /></td>
                       <td valign="top">
                          <table border="1" bordercolor="<%=BorderColor%>" cellpadding="0" cellspacing="0" width="<%=tablewidth%>"> 
                          <tr><td align="left"> <table cellpadding="3" cellspacing="0" width="100%">               <tr>
                              <td colspan="3" align="center" bgcolor="<%=BorderColor%>" class="RaytheonBodyWhiteBold">
                                  COMMODITY INFORMATION
                              </td>
                          </tr>
                          <!--
                          <tr>
                              <td class="RaytheonTextBlackBold" align="left">Pieces</td>
                              <td width="20"><img src="../images/pixel.gif" /></td>
                              <td nowrap class="RaytheonTextBlack">
                                  Number of Pieces:&nbsp;&nbsp;<input type="text" name="Pieces" value="<%=Pieces%>" size="3" maxlength="4" />
                              </td>
                          </tr>
                          -->
      
                          <tr>
                              <td class="RaytheonTextBlackBold" align="left" valign="top">Number of Pieces</td>
                              <td width="20"><img src="../images/pixel.gif" /></td>
                              <td nowrap class="RaytheonTextBlackBold">
                                  <input type="text" value="<%=pieces%>" name="pieces" size="3" maxlength="4" />
                                  &nbsp;&nbsp;
                                  <select name="rf_box">
                                      <option value="Boxes"<%If trim(rf_box)="Boxes" then Response.write "selected" end if%>>Boxes</option>
                                      <option value="Crates"<%If trim(rf_box)="Crates" then Response.write "selected" end if%>>Crates</option>
                                      <option value="Envelopes"<%If trim(rf_box)="Envelopes" then Response.write "selected" end if%>>Envelopes</option>
                                      <option value="Skids"<%If trim(rf_box)="Skids" then Response.write "selected" end if%>>Skids</option>
                                      <option value="X Square Probe Card(s)"<%If trim(rf_box)="X Square Probe Card(s)" then Response.write "selected" end if%>>X Square Probe Card(s)</option>
                                  </select>
                              </td>
                          </tr>
                          <tr>
                              <td class="RaytheonTextBlackBold" align="left">Palletized</td>
                              <td width="20"><img src="../images/pixel.gif" /></td>
                              <td align="left" class="RaytheonTextBlack">
                                  <select name="IsPalletized">
                                      <option value="y"<%If trim(IsPalletized)="y" then Response.write "selected" end if%>>Palletized</option>
                                      <option value="n"<%If trim(IsPalletized)="n" then Response.write "selected" end if%>>Not Palletized</option>
                                  </select>
                              </td>
                          </tr>
                          <tr>
                              <td class="RaytheonTextBlackBold" align="left">Weight</td>
                              <td width="20"><img src="../images/pixel.gif" /></td>
                              <td align="left" class="RaytheonTextBlack"><input type="text" name="DimWeight" value="<%=DimWeight%>" size="6" maxlength="5" /> Pounds</td>
                          </tr>
                          <tr>
                              <td class="RaytheonTextBlackBold" align="left">Dimensions</td>
                              <td width="20"><img src="../images/pixel.gif" /></td>
                              <td class="RaytheonTextBlack" align="left" nowrap>
                                  L:&nbsp;&nbsp;<input type="text" name="DimLength" value="<%=DimLength%>" size="5"  maxlength="4"/> 
                                  W:&nbsp;&nbsp;<input type="text" name="DimWidth" value="<%=DimWidth%>" size="5" maxlength="4" /> 
                                  H:&nbsp;&nbsp;<input type="text" name="DimHeight" value="<%=DimHeight%>" size="5"  maxlength="4"/> 
                                  &nbsp;&nbsp;Inches
                              </td>
                          </tr>
      
      
                          <tr>
                              <td class="RaytheonTextBlackBold" align="left">Hazmat</td>
                              <td width="20"><img src="../images/pixel.gif" /></td>
                              <td align="left"  class="RaytheonTextBlackBold">
                                   <select name="IsHazmat">
                                      <option value="n" <%If trim(IsHazmat)="n" then Response.write "selected" end if%>>No</option>
                                      <option value="y" <%If trim(IsHazmat)="y" then Response.write "selected" end if%>>Yes</option>
                                  </select>                     
                              </td>
                          </tr>
                          <tr>
                              <td class="RaytheonTextBlackBold" align="left">Refrigerate</td>
                              <td width="20"><img src="../images/pixel.gif" /></td>
                              <td align="left"  class="RaytheonTextBlackBold">
                                   <select name="Refrigerate">
                                      <option value="n" <%If trim(Refrigerate)="n" then Response.write "selected" end if%>>No</option>
                                      <option value="y" <%If trim(Refrigerate)="y" then Response.write "selected" end if%>>Yes</option>
                                  </select>                      
                              </td>
                          </tr>
                          <tr>
                              <td class="RaytheonTextBlackBold" align="left">Service Level</td>
                              <td width="20"><img src="../images/pixel.gif" /></td>
                              <td align="left"  class="RaytheonTextBlackBold">
                                   <select name="Priority">
                                      <option value="Next Day" <%If trim(Priority)="Next Day" then Response.write "selected" end if%>>Next Day</option>
                                      <option value="Same Day" <%If trim(Priority)="Same Day" then Response.write "selected" end if%>>Same Day</option>
                                      <option value="Time Critical" <%If trim(Priority)="Time Critical" then Response.write "selected" end if%>>Time Critical</option>
      
                                  </select>                      
                              </td>
                          </tr>
                          <tr>
                              <td class="RaytheonTextBlackBold" align="left">Routed To</td>
                              <td width="20"><img src="../images/pixel.gif" /></td>
                              <td align="left"  class="RaytheonTextBlack">
                                   <%=Courier %>                     
                              </td>
                          </tr>
                          <input type="hidden" name="courier" value="<%=Courier %>" />
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
                          <td colspan="3" align="center" bgcolor="<%=BorderColor%>" class="RaytheonBodyWhiteBold">
                              ORIGINATION
                          </td>
                      </tr>
                      <tr>
                          <td class="RaytheonTextBlackBold" align="left" nowrap>Company Name</td>
                          <td width="20"><img src="../images/pixel.gif" /></td>
                          <td align="left"><input type="text" name="OriginationCompany" value="<%=OriginationCompany%>" size="45" maxlength="40" /></td>
                      </tr>
                      <tr>
                          <td class="RaytheonTextBlackBold" align="left">Address</td>
                          <td width="20"><img src="../images/pixel.gif" /></td>
                          <td align="left"><input type="text" name="OriginationAddress" value="<%=OriginationAddress%>" size="45" maxlength="40" /></td>
                      </tr>
                      <tr>
                          <td class="RaytheonTextBlackBold" align="left">City/State/Zip</td>
                          <td width="20"><img src="../images/pixel.gif" /></td>
                          <td class="RaytheonTextBlackBold">
                              <input type="text" name="OriginationCity" value="<%=OriginationCity%>" size="20" maxlength="30" />
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
                              <input type="text" name="OriginationZipCode" value="<%=OriginationZipCode%>" size="11" maxlength="10" />
      
                          </td>
                      </tr>
      
      
                      <tr>
                          <td class="RaytheonTextBlackBold" align="left">Contact Name</td>
                          <td width="20"><img src="../images/pixel.gif" /></td>
                          <td align="left"><input type="text" name="OriginationContactName" value="<%=OriginationContactName%>" size="45" maxlength="25" /></td>
                      </tr>
                      <tr>
                          <td class="RaytheonTextBlackBold" align="left">Phone Number</td>
                          <td width="20"><img src="../images/pixel.gif" /></td>
                          <td align="left"><input type="text" name="OriginationPhoneNumber" value="<%=OriginationPhoneNumber%>" size="45" maxlength="20" /></td>
                      </tr> 
                      <tr>
                          <td class="RaytheonTextBlackBold" align="left">Email Address</td>
                          <td width="20"><img src="../images/pixel.gif" /></td>
                          <td align="left"><input type="text" name="OriginationEmail" value="<%=OriginationEmail%>" size="45" maxlength="100" /></td>
                      </tr> 
                      <tr>
                          <td class="RaytheonTextBlackBold" align="left">Pick Up Date/Time</td>
                          <td width="20"><img src="../images/pixel.gif" /></td>
                          <td align="left"><input type="text" id="PickupDateTime" name="PickupDateTime" value="<%=PickUpDateTime%>" size="30" maxlength="30" />
                          <a href="javascript:NewCal('PickupDateTime','MMddyyyy',true,12)"><img src="../images/cal.gif" width="16" height="16" border="0" alt="Pick a date"></a>                 
                          </td>
                      </tr>
                       </table>
                       </td></tr></table></td>
                       <td align="left"><img src="../images/pixel.gif" height="1" width="25" /></td>
                       <td align="left">
                          <table border="1" bordercolor="<%=BorderColor%>" cellpadding="0" cellspacing="0" width="<%=tablewidth%>"> 
                          <tr> <td valign="top"> <table cellpadding="3" cellspacing="0" width="100%">               <tr>
                              <td colspan="3" align="center" bgcolor="<%=BorderColor%>" class="RaytheonBodyWhiteBold">
                                  DESTINATION
                              </td>
                          </tr>
                          <tr>
                              <td class="RaytheonTextBlackBold" align="left" nowrap>Company Name</td>
                              <td width="20"><img src="../images/pixel.gif" /></td>
                              <td align="left"><input type="text" name="DESTINATIONCompany" value="<%=DESTINATIONCompany%>" size="45" maxlength="40" /></td>
                          </tr>
                          <tr>
                              <td class="RaytheonTextBlackBold" align="left">Address</td>
                              <td width="20"><img src="../images/pixel.gif" /></td>
                              <td align="left"><input type="text" name="DESTINATIONAddress" value="<%=DESTINATIONAddress%>" size="45" maxlength="40" /></td>
                          </tr>
                          <tr>
                              <td class="RaytheonTextBlackBold" align="left">City/State/Zip</td>
                              <td width="20"><img src="../images/pixel.gif" /></td>
                              <td class="RaytheonTextBlackBold">
                                  <input type="text" name="DESTINATIONCity" value="<%=DESTINATIONCity%>" size="20" maxlength="30" />
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
                                  <input type="text" name="DESTINATIONZipCode" value="<%=DESTINATIONZipCode%>" size="11" maxlength="10" />
      
                              </td>
                          </tr>
      
      
                          <tr>
                              <td class="RaytheonTextBlackBold" align="left">Contact Name</td>
                              <td width="20"><img src="../images/pixel.gif" /></td>
                              <td align="left"><input type="text" name="DESTINATIONContactName" value="<%=DESTINATIONContactName%>" size="45" maxlength="25" /></td>
                          </tr>
                          <tr>
                              <td class="RaytheonTextBlackBold" align="left">Phone Number</td>
                              <td width="20"><img src="../images/pixel.gif" /></td>
                              <td align="left"><input type="text" name="DESTINATIONPhoneNumber" value="<%=DESTINATIONPhoneNumber%>" size="45" maxlength="20" /></td>
                          </tr> 
                          <tr>
                              <td class="RaytheonTextBlackBold" align="left">Email Address</td>
                              <td width="20"><img src="../images/pixel.gif" /></td>
                              <td align="left"><input type="text" name="DESTINATIONEmail" value="<%=DESTINATIONEmail%>" size="45" maxlength="100" /></td>
                          </tr> 
                      <tr>
                          <td class="RaytheonTextBlackBold" align="left">Delivery Date/Time</td>
                          <td width="20"><img src="../images/pixel.gif" /></td>
                          <td align="left"><input type="text" name="DeliveryDateTime" id="DeliveryDateTime" value="<%=DeliveryDateTime%>" size="30" maxlength="30" />
                           <a href="javascript:NewCal('DeliveryDateTime','MMddyyyy',true,12)"><img src="../images/cal.gif" width="16" height="16" border="0" alt="Pick a date"></a>
                          </td>
                      </tr>
                           </table>
                           </td></tr></table>                 
                       </td></tr>                                                                                                               
                  </table>
               
                  </td>
              </tr>
               <tr><td align="left"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>
               <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="RaytheonHeaderWhite" colspan="2">FleetX Transportation Call Center 972-499-3415</td></tr>
               <tr><td align="left"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>
               <tr>
                  <td align="right">
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
                 <td align="right">
                                 <INPUT id="gobutton" name="gobutton" TYPE="submit" name="ButtonValue" VALUE="Submit Order">

                      <!--  <input type="image" src="../images/submit_fleetexpress2.gif" alt="submit order" />    -->
                  </td></tr> 
                  <input type="hidden" name="SearchJobNumber" value="<%=SearchJobNumber %>" />  
              <input type="hidden" name="ColorSelect" value="<%=ColorSelect %>" />
              <input type="hidden" name="pagestatus" value="submit" />
              </form>
              <%Case Else %>
                  <form method="post" action="FleetXOrderEdit.asp">
                  <tr><td>&nbsp;</td></tr>
                  <tr><td>&nbsp;</td></tr>
                  <tr><td>&nbsp;</td></tr>
                  <tr><td align="center" colspan="2" class="RaytheonTextBlackBold">Job Number:&nbsp;&nbsp;<input type="text" value="<%=SearchJobNumber%>" name="SearchJobNumber" />
                  <input type="hidden" name="PageStatus" value="FindJob" />
                  <INPUT id="gobutton" name="gobutton" TYPE="submit" name="ButtonValue" VALUE="Find Job">

                 <!-- <input type="submit" value="Find Job" name="submit" /> -->
                 </td></tr>
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
                                  <td colspan="7" align="center">
                                      <B>OPEN ORDERS</B>
                                  </td>
                              </tr>
                              <tr><td></td></tr>
                              <tr><td>JOB NUMBER</td><td width="20">&nbsp;</td><td>BOOK DATE/TIME</td><td width="20">&nbsp;</td><td>ORIGINATION</td><td width="20">&nbsp;</td><td>DESTINATION</td></tr>
      
                              <%
      			                Set RSEVENTS2 = Server.CreateObject("ADODB.Connection")
      			                RSEVENTS2.ConnectionTimeout = 100
      			                RSEVENTS2.Provider = "MSDASQL"
      			                RSEVENTS2.Open DATABASE
      				                l_cSQL = "SELECT     fcfgthd.fh_id, fclegs.fl_sf_name, fclegs.fl_st_name, fcfgthd.fh_ship_dt, fcfgthd.fh_carr_id FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id  " &_
      						                 "WHERE fh_status<> 'CAN' and fh_status<>'CLS'" 
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
                                          courier=Trim(oRs("fh_carr_id"))
      
                                      %>
                                      <tr><td><a href="FleetXOrderEdit.asp?SearchJobNumber=<%=fh_ID %>&PageStatus=FindAJob" class="FleetXRedMain"><%=fh_ID %></a></td><td width="20">&nbsp;</td><td><%=fh_ship_dt %></td><td width="20">&nbsp;</td><td><%=fl_sf_name %></td><td width="20">&nbsp;</td><td><%=fl_st_name %></td></tr>
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
              <%End Select %>
              <tr><td>&nbsp;</td></tr>
              <tr><td>&nbsp;</td></tr>
              <tr><td align="center">
                  <table><tr><td>
                <form method="post" action="FleetXOrderDispatch.asp?BTID=86">
                <INPUT id="gobutton" name="gobutton" TYPE="submit" name="ButtonValue" VALUE="Dispatch Order">

                                <!--  <input type="submit" name="submit" value="Dispatch Order" />    -->
                                  <input type="hidden" name="btid" value="86" />
                </form>
                 </td>
                <td>
      
                <form method="post" action="FreightOrder.asp">
                <INPUT id="gobutton" name="gobutton" TYPE="submit" name="ButtonValue" VALUE="Enter Order">

                                <!--  <input type="submit" name="submit" value="Enter Order" />   -->
                                  <input type="hidden" name="btid" value="86" />
                </form>
                </td><td></td></tr></table>       
              </td></tr>
          </table>    
    
    
    
    
    
    
    
    
    </td></tr>



 
	<tr Height="50">
		<td>&nbsp;</td>
	</tr>


</table>
</td></tr>
<%
if ErrorMessage>"" then%>
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
