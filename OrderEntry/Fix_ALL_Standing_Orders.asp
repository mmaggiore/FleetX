<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
    <title>Untitled Page</title>
<!-- #include file="../fleetexpress.inc" -->
</head>
<body>
    <%
    '''''''''''''''''''''''
    'ADD THE BILL TO ID !!!!!!!!!    
    '''''''''''''''''''''''
    Whatever=Request.QueryString("Whatever")
    Set RSEVENTS2 = Server.CreateObject("ADODB.Connection")
    RSEVENTS2.ConnectionTimeout = 1000
    RSEVENTS2.Provider = "MSDASQL"
    RSEVENTS2.Open DATABASE
      	l_cSQL = "SELECT distinct(fl_fh_id), fh_ready, fl_sf_building, fl_sf_addr1, fl_st_building, fl_st_addr1 FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id where fh_status='CLS' and fh_priority<>'12'" 
      	Response.write "l_cSQL="&l_cSQL&"<BR>"
        'Response.write "Database="&Database&"<BR>"
            SET oRs = RSEVENTS2.Execute(l_cSql)
            If oRs.eof then
                ErrorMessage="There are currently no open jobs"
      	End if
      	Do While not oRs.EOF
                xxx=xxx+1	
                orderID=Trim(oRs("fl_fh_id"))
                fh_ready=Trim(oRs("fh_ready"))
                'origination=Trim(oRs("origination"))
                'standingorder=Trim(oRs("standingorder"))
                fl_sf_building=Trim(oRs("fl_sf_building"))
                fl_sf_addr1=Trim(oRs("fl_sf_addr1"))
                fl_st_building=Trim(oRs("fl_st_building"))
                fl_st_addr1=Trim(oRs("fl_st_addr1"))
                companyaddress=fl_sf_addr1
                bcompanyaddress=fl_st_addr1
                fl_fh_id=orderID
                readytime=hour(fh_ready)
                'Response.write "**************<BR>"
                PickUpDateTime=fh_ready


                StandingReadyTime=hour(PickUpDateTime)
                DayOfWeek=Weekday(PickUpDateTime)
                StandingOrderID=0
                IsItCSF=0
                'Response.write "***********************<BR>"
                'Response.write "DayOfWeek="&DayOfWeek&"<BR>"
                'Response.write "StandingReadyTime="&StandingReadyTime&"<BR>"
                'Response.write "fl_sf_addr1="&fl_sf_addr1&"<BR>"
                'Response.write "fl_st_addr1="&fl_st_addr1&"<BR>"

                If DayOfWeek>1 and DayOfWeek<7 then
                        'Standing order Type 1
                        If (StandingReadyTime=10 or StandingReadyTime=7 or StandingReadyTime=8 or StandingReadyTime=9) and left(companyaddress, 7)="13020 T" and (left(bcompanyaddress,6)="3601 A" or left(bcompanyaddress,5)="300 W" or left(bcompanyaddress,5)="300 R") then
                            StandingOrderID=1
                            'DeliveryDateTime=DateValue(PickUpDateTime)&" 10:00:00 AM"
                            'Response.write "GOT HERE!!!!!<BR>"
                            'Response.write "DeliveryDateTime="&DeliveryDateTime&"<BR>"
                        End if
                        'Standing order Type 2
                        If (StandingReadyTime=6 or StandingReadyTime=7 or StandingReadyTime=8 or StandingReadyTime=9) and left(companyaddress, 7)="13536 N" and left(bcompanyaddress,7)="13601 I" then
                            StandingOrderID=2
                            'DeliveryDateTime=DateValue(PickUpDateTime)&" 11:00:00 AM"
                            'Response.write "GOT HERE!!!!!<BR>"
                            'Response.write "DeliveryDateTime="&DeliveryDateTime&"<BR>"
                        End if
                        'Standing order Type 3
                        If (StandingReadyTime=6 or StandingReadyTime=7 or StandingReadyTime=8 or StandingReadyTime=9) and left(companyaddress, 7)="13438 F" and left(bcompanyaddress,6)="13536 " then
                            StandingOrderID=3
                            'DeliveryDateTime=DateValue(PickUpDateTime)&" 10:00:00 AM"
                            'Response.write "GOT HERE!!!!!<BR>"
                            'Response.write "DeliveryDateTime="&DeliveryDateTime&"<BR>"
                        End if
                        'Standing order Type 4
                        If (StandingReadyTime=12 or StandingReadyTime=11 or StandingReadyTime=10 or StandingReadyTime=9) and left(companyaddress, 7)="13601 I" and left(bcompanyaddress,7)="13536 N" then
                            StandingOrderID=4
                            'DeliveryDateTime=DateValue(PickUpDateTime)&" 1:00:00 PM"
                            'Response.write "GOT HERE!!!!!<BR>"
                            'Response.write "DeliveryDateTime="&DeliveryDateTime&"<BR>"
                        End if
                        'Standing order Type 5
                        
                        If (StandingReadyTime=12 or StandingReadyTime=11 or StandingReadyTime=13 or StandingReadyTime=14) and left(companyaddress, 7)="13532 N" and left(bcompanyaddress,7)="12500 T" then
                            StandingOrderID=5
                            'DeliveryDateTime=DateValue(PickUpDateTime)&" 2:00:00 PM"
                            'Response.write "GOT HERE!!!!!<BR>"
                            'Response.write "DeliveryDateTime="&DeliveryDateTime&"<BR>"
                        End if
                        'Standing order Type 6
                        IsItCSF=InStr(UCASE(fl_st_building),"SOUTH BLDG CSF")
                        If (StandingReadyTime=12 or StandingReadyTime=11 or StandingReadyTime=13 or StandingReadyTime=14) and left(companyaddress, 7)="12500 T" and left(bcompanyaddress,7)="13536 N" and IsItCSF>0 then
                            StandingOrderID=6
                            'DeliveryDateTime=DateValue(PickUpDateTime)&" 2:30:00 PM"
                            'Response.write "GOT HERE!!!!!<BR>"
                            'Response.write "DeliveryDateTime="&DeliveryDateTime&"<BR>"
                        End if
                End if
                'Response.write "StandingOrderID="&StandingOrderID&"<BR>"
                If StandingOrderID>0 then
                    yyy=yyy+1
                    'Response.write "Got here #2<br>"
                    Response.write "StandingOrderID="&StandingOrderID&"/"
                    Response.write "readytime="&readytime&"/"

                    Response.write "orderID="&orderID&"/"
                    Response.write "fh_ready="&fh_ready&"/"
                    'Response.write "standingorder="&standingorder&"/"
                    'Response.write "fl_sf_building="&fl_sf_building&"/"
                    Response.write "ORIG="&fl_sf_addr1&"/"
                    'Response.write "fl_st_building="&fl_st_building&"/"
                    Response.write "DEST="&fl_st_addr1&"<BR>"
                    'Response.write "**************<BR>"
                   ' Response.write "fl_sf_id="&fl_sf_id&"<BR>"
                   ' Response.write "fl_st_id="&fl_st_id&"<BR>"
                    'Response.write "billtoid="&billtoid&"<BR>"
                    'Response.write "JobChargesDescription="&JobChargesDescription&"<BR>"
                    'Response.write "JobChargesRate="&JobChargesRate&"<BR>"
                    'Response.write "JobChargesBillCode="&JobChargesBillCode&"<BR>"
                    'Response.write "JobChargesStatus="&JobChargesStatus&"<BR>"
                    'Response.write "fh_id="&fh_id&"<BR>"
                    'Response.write "JobChargesID="&JobChargesID&"<BR>"
                    'Response.write "fl_t_disp="&fl_t_disp&"<BR>"
                    'Response.write "fh_bt_id="&fh_bt_id&"<BR>"
                    'Response.write "***********************<BR>"
                End if


                        If whatever="whatever" and StandingOrderID>0  then
                            Response.write "GOT HERE!  PROCESSING!<BR>"
                           '''If test="test" then
                                 Set oConn = Server.CreateObject("ADODB.Connection")
						        oConn.ConnectionTimeout = 100
						        oConn.Provider = "MSDASQL"
						        oConn.Open DATABASE
						        ' 7/26/04 KK: Added canceljob functionality to be able to update the status if cancel button is pressed.
						        ' 11/30/04 DEC: Changed from CAN/98 TO DEL/99 to be consistent with dispatchOffice			
							        bl_cSQL = "UPDATE jobcharges SET JobChargesStatus = 'x' WHERE fh_id = '" & fl_fh_id & "'"
							        'Response.write "bl_cSQL="&bl_cSQL&"<BR>"
							        oConn.Execute(bl_cSQL)
						        Set oConn=Nothing
                            '''End if
                            'If test="dont use anymore" then
                           '  Set oConn = Server.CreateObject("ADODB.Connection")
						   ' oConn.ConnectionTimeout = 100
						   ' oConn.Provider = "MSDASQL"
						   'oConn.Open DATABASE
						    ' 7/26/04 KK: Added canceljob functionality to be able to update the status if cancel button is pressed.
						    ' 11/30/04 DEC: Changed from CAN/98 TO DEL/99 to be consistent with dispatchOffice			
							'    bl_cSQL = "UPDATE MischargedStandingOrders SET StandingOrder = 3 WHERE orderid = '" & fl_fh_id & "'"
							    'Response.write "bl_cSQL="&bl_cSQL&"<BR>"
							'    oConn.Execute(bl_cSQL)
						   ' Set oConn=Nothing
                           ' End if
'''''''''''''''''''''''''''''''''''''''''''''

                        FleetXBillToID=93
                        'Response.write "fl_fh_id="&fl_fh_id&"<BR>"
                        'standingorder=3
                        Set oConn = Server.CreateObject("ADODB.Connection")
		                    oConn.ConnectionTimeout = 100
		                    oConn.Provider = "MSDASQL"
		                    oConn.Open DATABASE
                                'If trim(XSquare)="y" then
				                    '''l_cSQL = "Select * FROM PreExistingRequestor WHERE requestoR2789tatus='c' and requestorID IN('1') ORDER BY RequestorName"
                                    ' l_cSQL = "Select * FROM PreExistingRequestor WHERE requestoR2789tatus='c' and requestorID>'199' ORDER BY RequestorName"
                                    'else
                                    l_cSQL = "Select Charge, StandingOrderDescription FROM StandingOrderFees WHERE StandingOrderID="&standingorderID
                                'End if
			                    SET oR2789 = oConn.Execute(l_cSql)
					                    If not oR2789.EOF then
                                        PriorityCost=trim(oR2789("Charge"))
                                        PriorityDescription=trim(oR2789("StandingOrderDescription"))
                                        'varFuelCharge=FuelCharge/100
                                        'FuelChargeDollars=EStimatedCost*varFuelCharge
                                        EstimatedCost=PriorityCost
                                        End if
                                Set oR2789=Nothing
                'Response.write "Got Here***************************<BR>"
                'Response.write "fleetxNewJobNum="&fleetxNewJobNum&"<BR>"
                'Response.write "FleetXbilltoid="&FleetXbilltoid&"<BR>"
                'Response.write "PriorityDescription="&PriorityDescription&"<BR>"
               ' Response.write "PriorityCost="&PriorityCost&"<BR>"
                'Response.write "DeliveryDate="&DeliveryDate&"<BR>"

                '''If test="test" then
				Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
					RSEVENTS2.Open "JobCharges", DATABASE, 2, 2
					RSEVENTS2.addnew
					RSEVENTS2("fh_id")=fl_fh_id
					RSEVENTS2("billtoid")=FleetXBillToID
					RSEVENTS2("JobChargesDescription")=PriorityDescription
					RSEVENTS2("JobChargesRate")=PriorityCost
					RSEVENTS2("JobChargesBillCode")="N/A"
                    RSEVENTS2("JobChargesStatus")="c"
					RSEVENTS2.update
					RSEVENTS2.close			
				set RSEVENTS2 = nothing
                '''End if	
        Set oConn=Nothing
                        Set oConn = Server.CreateObject("ADODB.Connection")
		                    oConn.ConnectionTimeout = 100
		                    oConn.Provider = "MSDASQL"
		                    oConn.Open DATABASE
                                'If trim(XSquare)="y" then
				                    '''l_cSQL = "Select * FROM PreExistingRequestor WHERE requestoRs234tatus='c' and requestorID IN('1') ORDER BY RequestorName"
                                    ' l_cSQL = "Select * FROM PreExistingRequestor WHERE requestoRs234tatus='c' and requestorID>'199' ORDER BY RequestorName"
                                    'else
                                    l_cSQL = "Select FuelCharge, FuelChargeDate FROM FuelChargeList WHERE FuelChargeDate<'"&cdate(fh_ready)&"' and fuelchargestatus<>'x' order by fuelchargeID desc"
                                'End if
                                'Response.write "l_cSQL="&l_cSQL&"<BR>"
			                    SET oRs234 = oConn.Execute(l_cSql)
					                    If not oRs234.EOF then
                                            FuelCharge=trim(oRs234("FuelCharge"))
                                            FuelChargeDate=trim(oRs234("FuelChargeDate"))
                                            varFuelCharge=FuelCharge/100
                                            FuelChargeDollars=EStimatedCost*varFuelCharge
                                            Response.write "FuelCharge="&FuelCharge&"<BR>"
                                            Response.write "FuelChargeDate="&FuelChargeDate&"<BR>"
                                            Response.write "varFuelCharge="&varFuelCharge&"<BR>"
                                            Response.write "FuelChargeDollars="&FuelChargeDollars&"<BR>"
                                            Response.write "********************<BR>"
                                            'EstimatedCost=EstimatedCost+FuelChargeDollars
                                            '''if test="test" then
				                                Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
					                                RSEVENTS2.Open "JobCharges", DATABASE, 2, 2
					                                RSEVENTS2.addnew
					                                RSEVENTS2("fh_id")=fl_fh_id
					                                RSEVENTS2("billtoid")=FleetXBillToID
					                                RSEVENTS2("JobChargesDescription")="Fuel Charge"
					                                RSEVENTS2("JobChargesRate")=FuelChargeDollars
					                                RSEVENTS2("JobChargesBillCode")="Fuel Charge"
                                                    RSEVENTS2("JobChargesStatus")="c"
					                                RSEVENTS2.update
					                                RSEVENTS2.close			
				                                set RSEVENTS2 = nothing
                                            '''end if	
                                        End if
                                Set oRs234 = Nothing
'''''''''''''''''''''''''''''''''''''''''''''












                        End if
      
      	oRs.movenext
      	LOOP
            oRs.close
            Set oRs=Nothing
        'RSEVENTS2.Close
    'Set RSEVENTS2=Nothing
    Response.write "DONE!!!<BR>"
    'Response.write "l_cSQL="&l_cSQL&"<BR>"
    Response.write "xxx="&xxx&"<BR>"
    Response.write "yyy="&yyy&"<BR>"
    %>
</body>
</html>
