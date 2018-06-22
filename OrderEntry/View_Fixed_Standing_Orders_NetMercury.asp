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
      	l_cSQL = "SELECT distinct(fl_fh_id), fl_t_atd, fl_sf_building, fl_sf_addr1, fl_st_building, fl_st_addr1, fh_ready FROM fclegs INNER JOIN fcfgthd ON fclegs.fl_fh_id = fcfgthd.fh_id where fl_sf_addr1 LIKE '13438%' and fl_st_addr1 LIKE '13536%'  and fh_ready<'2/28/2017' " 
      	'Response.write "l_cSQL="&l_cSQL&"<BR>"
        'Response.write "Database="&Database&"<BR>"
            SET oRs = RSEVENTS2.Execute(l_cSql)
            If oRs.eof then
                ErrorMessage="There are currently no open jobs"
      	End if
      	Do While not oRs.EOF
                xxx=xxx+1	
                orderID=Trim(oRs("fl_fh_id"))
                DeliveryDate=Trim(oRs("fl_t_atd"))
                'origination=Trim(oRs("origination"))
                'standingorder=Trim(oRs("standingorder"))
                fl_sf_building=Trim(oRs("fl_sf_building"))
                fl_sf_addr1=Trim(oRs("fl_sf_addr1"))
                fl_st_building=Trim(oRs("fl_st_building"))
                fl_st_addr1=Trim(oRs("fl_st_addr1"))
                fh_ready=Trim(oRs("fh_ready"))
                fl_fh_id=orderID
                readytime=hour(fh_ready)
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
                        If (readytime=7 or readytime=8 or readytime=9) then
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
                                           'Response.write "ZZZZZZZ/JobChargesDescription"&JobChargesDescription&"/"
                                           
                                           'Response.write "JobChargesBillCode="&JobChargesBillCode&"/"
                                           'Response.write "JobChargesRate="&JobChargesRate&"<BR>"
                                            'Response.write "FuelChargeDollars="&FuelChargeDollars&"<BR>"
                                           ' Response.write "********************<BR>"
                                            'EstimatedCost=EstimatedCost+FuelChargeDollars
                                            TotalIncorrectCharges=cCur(TotalIncorrectCharges)+cCur(JobChargesRate)
                                              	oRs234.movenext
      	                                    LOOP
                                Set oRs234 = Nothing
                                'Response.write "XXXXTotalIncorrectCharges="&TotalIncorrectCharges&"<BR>"
                                    
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
                                           TotalcorrectCharges=cCur(TotalcorrectCharges)+cCur(JobChargesRate)
                                            'Response.write "FuelChargeDollars="&FuelChargeDollars&"<BR>"
                                            'Response.write "********************<BR>"
                                            'EstimatedCost=EstimatedCost+FuelChargeDollars
                                              	oRs234.movenext
      	                                    LOOP
                                Set oRs234 = Nothing
                                'Response.write Fh_id&"====TotalCorrectCharges="&TotalCorrectCharges&"<BR>"
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









                        End if
      
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
