<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
    <title>Untitled Page</title>
<!-- #include file="../fleetexpress.inc" -->
</head>
<body>
Hello world!
    <%
    Whatever=Request.QueryString("Whatever")
    Set RSEVENTS2 = Server.CreateObject("ADODB.Connection")
    RSEVENTS2.ConnectionTimeout = 1000
    RSEVENTS2.Provider = "MSDASQL"
    RSEVENTS2.Open DATABASE
      	l_cSQL = "SELECT     fclegs.fl_fh_id, fclegs.fl_sf_id, fclegs.fl_st_id, JobCharges.billtoid, JobCharges.JobChargesDescription, "
        l_cSQL = l_cSQL &"JobCharges.JobChargesRate, JobCharges.JobChargesBillCode, JobCharges.JobChargesStatus, JobCharges.fh_id, JobCharges.JobChargesID, "
        l_cSQL = l_cSQL &"fclegs.fl_t_disp, fcfgthd.fh_bt_id "
        l_cSQL = l_cSQL &"FROM         fclegs INNER JOIN "
        l_cSQL = l_cSQL &"JobCharges ON fclegs.fl_fh_id = JobCharges.fh_id INNER JOIN "
        l_cSQL = l_cSQL &"fcfgthd ON JobCharges.fh_id = fcfgthd.fh_id "
        l_cSQL = l_cSQL &"WHERE     (fclegs.fl_sf_id = 'TISHR') AND (JobCharges.JobChargesDescription IS NULL) OR "
        l_cSQL = l_cSQL &"(fclegs.fl_st_id = 'TISHR') AND (JobCharges.JobChargesDescription IS NULL) "
        l_cSQL = l_cSQL &"ORDER BY fclegs.fl_fh_id DESC" 
      	Response.write "l_cSQL="&l_cSQL&"<BR>"
            'Response.write "Database="&Database&"<BR>"
            SET oRs = RSEVENTS2.Execute(l_cSql)
            If oRs.eof then
                ErrorMessage="There are currently no open jobs"
      	End if
      	Do While not oRs.EOF
                xxx=xxx+1	



                fl_fh_id=Trim(oRs("fl_fh_id"))
                fl_sf_id=Trim(oRs("fl_sf_id"))
                fl_st_id=Trim(oRs("fl_st_id"))
                billtoid=Trim(oRs("billtoid"))
                JobChargesDescription=Trim(oRs("JobChargesDescription"))
                JobChargesRate=Trim(oRs("JobChargesRate"))
                JobChargesBillCode=Trim(oRs("JobChargesBillCode"))
                JobChargesStatus=Trim(oRs("JobChargesStatus"))
                fh_id=Trim(oRs("fh_id"))
                JobChargesID=Trim(oRs("JobChargesID"))
                fl_t_disp=Trim(oRs("fl_t_disp"))
                fh_bt_id=Trim(oRs("fh_bt_id"))


                Response.write "fl_fh_id="&fl_fh_id&"<BR>"
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
                        If whatever="whatever" then
                            Response.write "GOT HERE!  PROCESSING!<BR>"
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
'''''''''''''''''''''''''''''''''''''''''''''

                        FleetXBillToID=93

                        Set oConn = Server.CreateObject("ADODB.Connection")
		                    oConn.ConnectionTimeout = 100
		                    oConn.Provider = "MSDASQL"
		                    oConn.Open DATABASE
                                'If trim(XSquare)="y" then
				                    '''l_cSQL = "Select * FROM PreExistingRequestor WHERE requestoR2789tatus='c' and requestorID IN('1') ORDER BY RequestorName"
                                    ' l_cSQL = "Select * FROM PreExistingRequestor WHERE requestoR2789tatus='c' and requestorID>'199' ORDER BY RequestorName"
                                    'else
                                    l_cSQL = "Select PriorityCost, PriorityDescription FROM Priorities WHERE PriorityDescription='4 Hour (TI-Sherman)' and PriorityStatus='c'"
                                'End if
			                    SET oR2789 = oConn.Execute(l_cSql)
					                    If not oR2789.EOF then
                                        PriorityCost=trim(oR2789("PriorityCost"))
                                        PriorityDescription=trim(oR2789("PriorityDescription"))
                                        'varFuelCharge=FuelCharge/100
                                        'FuelChargeDollars=EStimatedCost*varFuelCharge
                                        EstimatedCost=PriorityCost
                                        End if
                                Set oR2789=Nothing
                'Response.write "Got Here***************************<BR>"
                'Response.write "fleetxNewJobNum="&fleetxNewJobNum&"<BR>"
                'Response.write "billtoid="&billtoid&"<BR>"
                'Response.write "PriorityDescription="&PriorityDescription&"<BR>"
                'Response.write "PriorityCost="&PriorityCost&"<BR>"


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
        Set oConn=Nothing
                        Set oConn = Server.CreateObject("ADODB.Connection")
		                    oConn.ConnectionTimeout = 100
		                    oConn.Provider = "MSDASQL"
		                    oConn.Open DATABASE
                                'If trim(XSquare)="y" then
				                    '''l_cSQL = "Select * FROM PreExistingRequestor WHERE requestoRs234tatus='c' and requestorID IN('1') ORDER BY RequestorName"
                                    ' l_cSQL = "Select * FROM PreExistingRequestor WHERE requestoRs234tatus='c' and requestorID>'199' ORDER BY RequestorName"
                                    'else
                                    l_cSQL = "Select FuelCharge FROM FuelChargeList WHERE fuelchargeStatus='c'"
                                'End if
			                    SET oRs234 = oConn.Execute(l_cSql)
					                    If not oRs234.EOF then
                                            FuelCharge=trim(oRs234("FuelCharge"))
                                            varFuelCharge=FuelCharge/100
                                            FuelChargeDollars=EStimatedCost*varFuelCharge
                                            'EstimatedCost=EstimatedCost+FuelChargeDollars
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
    Response.write "l_cSQL="&l_cSQL&"<BR>"
    Response.write "xxx="&xxx&"<BR>"
    %>
</body>
</html>
