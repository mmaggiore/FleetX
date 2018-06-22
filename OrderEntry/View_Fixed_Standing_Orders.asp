<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
    <title>Untitled Page</title>
<!-- #include file="../fleetexpress.inc" -->
</head>
<body>
<table border="1">
    <tr><td>Job Number</td><td>Previously Billed</td><td>Should have been billed</td><td>To be billed</td></tr>
    <%
    Whatever=Request.QueryString("Whatever")
    Set RSEVENTS2 = Server.CreateObject("ADODB.Connection")
    RSEVENTS2.ConnectionTimeout = 1000
    RSEVENTS2.Provider = "MSDASQL"
    RSEVENTS2.Open DATABASE
      	l_cSQL = "SELECT distinct(orderid), deliverydate, origination, standingorder, fl_sf_building, fl_sf_addr1, fl_st_building, fl_st_addr1 FROM MischargedStandingOrders where standingorder>=1 order by fl_sf_addr1" 
      	'Response.write "l_cSQL="&l_cSQL&"<BR>"
        'Response.write "Database="&Database&"<BR>"
            SET oRs = RSEVENTS2.Execute(l_cSql)
            If oRs.eof then
                ErrorMessage="There are currently no open jobs"
      	End if
      	Do While not oRs.EOF
                xxx=xxx+1	
                orderID=Trim(oRs("orderID"))
                DeliveryDate=Trim(oRs("DeliveryDate"))
                origination=Trim(oRs("origination"))
                standingorder=Trim(oRs("standingorder"))
                fl_sf_building=Trim(oRs("fl_sf_building"))
                fl_sf_addr1=Trim(oRs("fl_sf_addr1"))
                fl_st_building=Trim(oRs("fl_st_building"))
                fl_st_addr1=Trim(oRs("fl_st_addr1"))

                fl_fh_id=orderID

                'Response.write "**************<BR>"
                'Response.write "orderID="&orderID&"/"
                'Response.write "DeliveryDate="&DeliveryDate&"/"
                'Response.write "standingorder="&standingorder&"/"
                'Response.write "fl_sf_building="&fl_sf_building&"/"
                'Response.write "fl_sf_addr1="&fl_sf_addr1&"/"
                'Response.write "fl_st_building="&fl_st_building&"/"
                'Response.write "fl_st_addr1="&fl_st_addr1&"<BR>"
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
                        'If whatever="whatever" then
'''''''''''''''''''''''''''''''''''''''''''''

                        FleetXBillToID=93
                        'Response.write "*****************<BR>"
                        
                        'Response.write "Incorrect Charges:<BR>"

                        Set oConn = Server.CreateObject("ADODB.Connection")
		                    oConn.ConnectionTimeout = 100
		                    oConn.Provider = "MSDASQL"
		                    oConn.Open DATABASE
                                'If trim(XSquare)="y" then
				                    '''l_cSQL = "Select * FROM PreExistingRequestor WHERE requestoRs234tatus='c' and requestorID IN('1') ORDER BY RequestorName"
                                    ' l_cSQL = "Select * FROM PreExistingRequestor WHERE requestoRs234tatus='c' and requestorID>'199' ORDER BY RequestorName"
                                    'else
                                    l_cSQL = "Select JobChargesDescription, JobChargesRate, JobChargesBillCode FROM JobCharges WHERE fh_id='"&fl_fh_id&"' and JobChargesStatus='x'"
                                'End if
			                    SET oRs234 = oConn.Execute(l_cSql)
					                    Do while not oRs234.EOF
                                            JobChargesDescription=trim(oRs234("JobChargesDescription"))
                                            JobChargesRate=trim(oRs234("JobChargesRate"))
                                            JobChargesBillCode=trim(oRs234("JobChargesBillCode"))
                                            'FuelChargeDollars=EStimatedCost*varFuelCharge
                                           'Response.write JobChargesDescription&"/"
                                           
                                           'Response.write JobChargesBillCode&"/"
                                           'Response.write JobChargesRate&"<BR>"
                                            'Response.write "FuelChargeDollars="&FuelChargeDollars&"<BR>"
                                           ' Response.write "********************<BR>"
                                            'EstimatedCost=EstimatedCost+FuelChargeDollars
                                            TotalIncorrectCharges=TotalIncorrectCharges+JobChargesRate
                                              	oRs234.movenext
      	                                    LOOP
                                Set oRs234 = Nothing
                                    
'''''''''''''''''''''''''''''''''''''''''''''
                                    'Response.write "Correct Charges:<BR>"
                        Set oConn = Server.CreateObject("ADODB.Connection")
		                    oConn.ConnectionTimeout = 100
		                    oConn.Provider = "MSDASQL"
		                    oConn.Open DATABASE
                                'If trim(XSquare)="y" then
				                    '''l_cSQL = "Select * FROM PreExistingRequestor WHERE requestoRs234tatus='c' and requestorID IN('1') ORDER BY RequestorName"
                                    ' l_cSQL = "Select * FROM PreExistingRequestor WHERE requestoRs234tatus='c' and requestorID>'199' ORDER BY RequestorName"
                                    'else
                                    l_cSQL = "Select JobChargesDescription, JobChargesRate, JobChargesBillCode FROM JobCharges WHERE fh_id='"&fl_fh_id&"' and JobChargesStatus='c'"
                                'End if
			                    SET oRs234 = oConn.Execute(l_cSql)
					                    Do while not oRs234.EOF
                                            JobChargesDescription=trim(oRs234("JobChargesDescription"))
                                            JobChargesRate=trim(oRs234("JobChargesRate"))
                                            JobChargesBillCode=trim(oRs234("JobChargesBillCode"))
                                            'FuelChargeDollars=EStimatedCost*varFuelCharge
                                           'Response.write JobChargesDescription&"/"
                                           
                                           'Response.write JobChargesBillCode&"/"
                                           'Response.write JobChargesRate&"<BR>"
                                           TotalcorrectCharges=TotalcorrectCharges+JobChargesRate
                                            'Response.write "FuelChargeDollars="&FuelChargeDollars&"<BR>"
                                           ' Response.write "********************<BR>"
                                            'EstimatedCost=EstimatedCost+FuelChargeDollars
                                              	oRs234.movenext
      	                                    LOOP
                                Set oRs234 = Nothing
                                NewCharge=TotalCorrectCharges-TotalIncorrectCharges
                                If NewCharge>0 then fontcolorvar="blue" end if
                                If NewCharge<0 then fontcolorvar="red" end if
                                If Newcharge=0 then fontcolorvar="black" end if
                                If trim(fontcolorvar)<>"black" then
                                    Response.write "<tr><td>"&fl_fh_id&"</td>"
                                    Response.write "<td>"&TotalIncorrectCharges&"</td>"
                                    Response.write "<td>"&TotalcorrectCharges&"</td>"

                                    

                                    Response.write "<td><font color='"&FontColorVar&"'>"& cCur(NewCharge) &"</font></b></td></tr>"
                                End if

                                NewCharge=0
                                TotalCorrectCharges=0
                                TotalIncorrectCharges=0









                        'End if
      
      	oRs.movenext
      	LOOP
            oRs.close
            Set oRs=Nothing
        'RSEVENTS2.Close
    'Set RSEVENTS2=Nothing
    'Response.write "DONE!!!<BR>"
    'Response.write "l_cSQL="&l_cSQL&"<BR>"
    'Response.write "xxx="&xxx&"<BR>"
    %>
    </table>
</body>
</html>
