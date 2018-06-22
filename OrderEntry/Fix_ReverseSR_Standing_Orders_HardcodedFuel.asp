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
      	
        'l_cSQL = "SELECT distinct(fl_fh_id), fh_ready, fl_sf_building, fl_sf_addr1, fl_st_building, fl_st_addr1 FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id where fh_status='CLS' and fh_priority<>'12'" 
      	'l_cSQL = "SELECT fcfgthd.fh_id, fcfgthd.fh_ready, JobCharges.billtoid, JobCharges.JobChargesDescription, JobCharges.JobChargesRate, JobCharges.JobChargesBillCode, JobCharges.JobChargesStatus FROM fcfgthd INNER JOIN JobCharges ON fcfgthd.fh_id = JobCharges.fh_id WHERE (fcfgthd.fh_bt_id = '91') AND (JobCharges.billtoid = '93') AND (JobCharges.JobChargesDescription LIKE 'Standing Order Type%')"
        l_cSQL = "SELECT fcfgthd.fh_id, fcfgthd.fh_ready FROM fcfgthd  WHERE  (fh_id IN ('00081138', '00081137', '00081136', '00081135', '00081114', '00081129', '00081127', '00081128', '00081121', '00081122', '00081119', '00081120', '00081117', '00081118', '00081115', '00081130', '00081131', '00081132', '00081113', '00081116', '00081123'))"
        Response.write "l_cSQL="&l_cSQL&"<BR>"
        'Response.write "Database="&Database&"<BR>"
            SET oRs = RSEVENTS2.Execute(l_cSql)
            If oRs.eof then
                ErrorMessage="There are currently no open jobs"
      	End if
      	Do While not oRs.EOF
                xxx=xxx+1	
                orderID=Trim(oRs("fh_id"))
                fh_ready=Trim(oRs("fh_ready"))
                readytime=fh_ready
                'origination=Trim(oRs("origination"))
                'standingorder=Trim(oRs("standingorder"))
                'fl_sf_building=Trim(oRs("fl_sf_building"))
                'fl_sf_addr1=Trim(oRs("fl_sf_addr1"))
                'fl_st_building=Trim(oRs("fl_st_building"))
                'fl_st_addr1=Trim(oRs("fl_st_addr1"))
                'companyaddress=fl_sf_addr1
                'bcompanyaddress=fl_st_addr1
                fl_fh_id=orderID
                'readytime=hour(fh_ready)
                'Response.write "**************<BR>"
                'PickUpDateTime=fh_ready


                'StandingReadyTime=hour(PickUpDateTime)
                'DayOfWeek=Weekday(PickUpDateTime)
                StandingOrderID=0
                IsItCSF=0

                    yyy=yyy+1
                    'Response.write "Got here #2<br>"
                    'Response.write "StandingOrderID="&StandingOrderID&"/"
                    Response.write "readytime="&readytime&"/"

                    Response.write "orderID="&orderID&"<br>"

                        If whatever="whatever" then
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

                        FleetXBillToID=91


                '''If test="test" then
				Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
					RSEVENTS2.Open "JobCharges", DATABASE, 2, 2
					RSEVENTS2.addnew
					RSEVENTS2("fh_id")=fl_fh_id
					RSEVENTS2("billtoid")=FleetXBillToID
					RSEVENTS2("JobChargesDescription")="Stockroom Charge"
					RSEVENTS2("JobChargesRate")=7.25
					RSEVENTS2("JobChargesBillCode")="FE FEE"
                    RSEVENTS2("JobChargesStatus")="c"
					RSEVENTS2.update
					RSEVENTS2.close			
				set RSEVENTS2 = nothing

                EstimatedCost=7.25

                '''End if	
        Set oConn=Nothing
                                                Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
					                                RSEVENTS2.Open "JobCharges", DATABASE, 2, 2
					                                RSEVENTS2.addnew
					                                RSEVENTS2("fh_id")=fl_fh_id
					                                RSEVENTS2("billtoid")=FleetXBillToID
					                                RSEVENTS2("JobChargesDescription")="Fuel Charge"
					                                RSEVENTS2("JobChargesRate")=1.23
					                                RSEVENTS2("JobChargesBillCode")="Fuel Charge"
                                                    RSEVENTS2("JobChargesStatus")="c"
					                                RSEVENTS2.update
					                                RSEVENTS2.close			
				                                set RSEVENTS2 = nothing
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
