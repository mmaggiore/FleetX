<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
    <title>Untitled Page</title>
<!-- #include file="../fleetexpress.inc" -->
</head>
<body>
    <%
    Whatever=Request.QueryString("Whatever")
    Set RSEVENTS2 = Server.CreateObject("ADODB.Connection")
    RSEVENTS2.ConnectionTimeout = 1000
    RSEVENTS2.Provider = "MSDASQL"
    RSEVENTS2.Open DATABASE
      	l_cSQL = "SELECT distinct(fl_fh_id), fh_ready, fl_sf_building, fl_sf_addr1, fl_st_building, fl_st_addr1 FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id where fl_sf_addr1 LIKE '13438%' and fl_st_addr1 LIKE '13536%'  and fh_ready<'2/28/2017'" 
      	'Response.write "l_cSQL="&l_cSQL&"<BR>"
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

                fl_fh_id=orderID
                readytime=hour(fh_ready)
                'Response.write "**************<BR>"
                
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
                        If whatever="whatever" and (readytime=7 or readytime=8 or readytime=9) then
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
                            If test="dont use anymore" then
                             Set oConn = Server.CreateObject("ADODB.Connection")
						    oConn.ConnectionTimeout = 100
						    oConn.Provider = "MSDASQL"
						    oConn.Open DATABASE
						    ' 7/26/04 KK: Added canceljob functionality to be able to update the status if cancel button is pressed.
						    ' 11/30/04 DEC: Changed from CAN/98 TO DEL/99 to be consistent with dispatchOffice			
							    bl_cSQL = "UPDATE MischargedStandingOrders SET StandingOrder = 3 WHERE orderid = '" & fl_fh_id & "'"
							    'Response.write "bl_cSQL="&bl_cSQL&"<BR>"
							    oConn.Execute(bl_cSQL)
						    Set oConn=Nothing
                            End if
'''''''''''''''''''''''''''''''''''''''''''''

                        FleetXBillToID=93
                        'Response.write "fl_fh_id="&fl_fh_id&"<BR>"
                        standingorder=3
                        Set oConn = Server.CreateObject("ADODB.Connection")
		                    oConn.ConnectionTimeout = 100
		                    oConn.Provider = "MSDASQL"
		                    oConn.Open DATABASE
                                'If trim(XSquare)="y" then
				                    '''l_cSQL = "Select * FROM PreExistingRequestor WHERE requestoR2789tatus='c' and requestorID IN('1') ORDER BY RequestorName"
                                    ' l_cSQL = "Select * FROM PreExistingRequestor WHERE requestoR2789tatus='c' and requestorID>'199' ORDER BY RequestorName"
                                    'else
                                    l_cSQL = "Select Charge, StandingOrderDescription FROM StandingOrderFees WHERE StandingOrderID="&standingorder
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
    %>
</body>
</html>
