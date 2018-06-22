<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
    <title>Untitled Page</title>
<!-- #include file="../fleetexpress.inc" -->
</head>
<body>
Hello world!
    <%
    Set RSEVENTS2 = Server.CreateObject("ADODB.Connection")
    RSEVENTS2.ConnectionTimeout = 1000
    RSEVENTS2.Provider = "MSDASQL"
    RSEVENTS2.Open DATABASE
      	l_cSQL = "SELECT fcfgthd.fh_id AS Original_fh_id, fcfgthd.fh_ship_dt AS Original_shipdate, fcfgthd.fh_bt_id AS Original_BT_ID, JobCharges.fh_id AS Actual_fh_id, JobCharges.billtoid AS Actual_BT_ID, JobCharges.JobChargesDescription AS JobChargesDescription FROM fcfgthd INNER JOIN JobCharges ON fcfgthd.fh_id = JobCharges.fh_id WHERE fcfgthd.fh_bt_id<>JobCharges.billtoid and JobCharges.billtoid<>'38'" 
      	Response.write "l_cSQL="&l_cSQL&"<BR>"
            'Response.write "Database="&Database&"<BR>"
            SET oRs = RSEVENTS2.Execute(l_cSql)
            If oRs.eof then
                ErrorMessage="There are currently no open jobs"
      	End if
      	Do While not oRs.EOF
                xxx=xxx+1	
                Original_fh_id=Trim(oRs("Original_fh_id"))
                Original_shipdate=Trim(oRs("Original_shipdate"))
                Original_BT_ID=Trim(oRs("Original_BT_ID"))
                Actual_fh_id=Trim(oRs("Actual_fh_id"))
                Actual_BT_ID=Trim(oRs("Actual_BT_ID"))
                Response.write "Original_shipdate="&Original_shipdate&"<BR>"
                Response.write "Original_fh_id="&Original_fh_id&"<BR>"
                Response.write "Original_BT_ID="&Original_BT_ID&"<BR>"
                Response.write "Actual_fh_id="&Actual_fh_id&"<BR>"
                Response.write "Actual_BT_ID="&Actual_BT_ID&"<BR>"
                Response.write "***********************<BR>"
                         Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open DATABASE
						' 7/26/04 KK: Added canceljob functionality to be able to update the status if cancel button is pressed.
						' 11/30/04 DEC: Changed from CAN/98 TO DEL/99 to be consistent with dispatchOffice			
							bl_cSQL = "UPDATE fcfgthd SET fh_bt_id = '"& Actual_BT_ID &"' WHERE fh_id = '" & Original_fh_id & "'"
							'Response.write "bl_cSQL="&bl_cSQL&"<BR>"
							oConn.Execute(bl_cSQL)
						Set oConn=Nothing
      
      	oRs.movenext
      	LOOP
            oRs.close
            Set oRs=Nothing
        RSEVENTS2.Close
    Set RSEVENTS2=Nothing
    Response.write "l_cSQL="&l_cSQL&"<BR>"
    Response.write "xxx="&xxx&"<BR>"
    %>
</body>
</html>
