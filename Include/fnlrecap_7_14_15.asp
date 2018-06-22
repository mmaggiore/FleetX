<!-- #include file="../fleetexpress.inc" -->

<%


' FNLRECAP.asp - Allows customers to totally control
'	their final recap page!!
'
'  L O G I S T I C O R P    
'
' Modified 5/11/05
'
' NOTE: You should do a test for AIR/LTL or COURIER here and make
'	a recap appropriately!


IF l_cJobNum="" THEN ''''''1
	l_cJobNum = Request.QueryString("l_cJobNum")
	If l_cJobNum="" then
		l_cJobNum=Request.Form("txtJobNumber")
	End if
	'Response.Write("<BR><FONT COLOR=RED>Editing Job #" & l_cJobNum & "</FONT>")
	
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.ConnectionTimeout = 100
	oConn.Provider = "MSDASQL"
	oConn.Open DATABASE

	'SECURITY - First ensure that the requested jobnumber belongs to the logged user
	l_cSQL = "Select fh_id from fcfgthd " &_
			 "INNER JOIN fcbillto ON fh_bt_id = bt_id " &_
			 "INNER JOIN fcinetsc ON in_for_id = bt_id " &_
			 "WHERE fh_id = '" & l_cJobNum & "' " &_
			 "AND bt_id = '" & Session("sBT_ID") & "' " &_
			 "AND in_user = '" & Session("sUsername") & "' "

	Set oRs = oConn.Execute(l_cSQL)
	IF NOT oRs.EOF THEN
		l_lfnlrecap = TRUE
	ELSE
		l_lfnlrecap = TRUE
	END IF

	'If the requested job# doesn't belong to the logged user then display below message.  Noone should come to the
	'False part unless someone manually links to the fnlrecap.asp?l_cjobnum=xxxxxxxx.
	IF l_lfnlrecap = FALSE THEN ''''2
		
		Response.Write("<Font color=red>Job # " & l_cJobNum & " does not belong to your account" &_
			" therefore you are not authorized to view this job.</Font>")
		oConn.Close
	
	ELSE ''''2

		' Get that job's information and set all of our textboxes/listboxes
		'	so they show that job's information...
		l_cSQL = "SELECT fh_co_id, fh_status, fh_custpo, fh_priority, " & _
				"fl_sf_id, fl_sf_name, fl_sf_cfname, fl_sf_clname, fl_sf_phone, " & _
				"fl_sf_addr1, fl_sf_addr2, fl_sf_city, fl_sf_state, fl_sf_zip, " & _
				"fl_sf_rta, fh_ready, " & _
				"LTRIM(CAST(fl_sf_comment AS varchar(200))) AS fl_sf_comment, " & _
				"fl_st_id, fl_st_name, fl_st_cfname, fl_st_clname, fl_st_phone, " & _
				"fl_st_addr1, fl_st_addr2, fl_st_city, fl_st_state, fl_st_zip, " & _
				"LTRIM(CAST(fl_st_comment AS varchar(200))) AS fl_st_comment, " & _
				"fl_st_rta, fl_numboxes, fl_boxtype, fl_weight, " & _
				"fh_user4, fh_user3 " & _			
				"FROM fclegs " & _
				"JOIN fcfgthd ON fl_fh_id = fh_id " & _
				"WHERE fl_fh_id = '" & l_cJobNum & "'"
		Response.write "l_cSQL="&l_cSQL&"<BR>"		
		Set oRs = oConn.Execute(l_cSQL)

		' Now, we need to replace all of these fields that were used to seed the
		' values in the textboxes with the values for this job
		IF NOT oRs.EOF THEN
			' We have a record!  Fill our values so the screen shows the data...
			l_lSeeded = TRUE
			l_cStatus = Trim(oRs.Fields("fh_status"))
			l_cCaller = Trim(oRs.Fields("fh_co_id"))
			'l_cBillTo = Trim(oRs.Fields("bt_desc"))
			l_cContactPhone = Trim(oRs.Fields("fh_user3"))
					
			l_cPUID = Trim(oRs.Fields("fl_sf_id"))
			l_cPUCompany = Trim(oRs.Fields("fl_sf_name"))
			l_cPUContact = Trim(oRs.Fields("fl_sf_cfname")) & " " & Trim(oRs.Fields("fl_sf_clname"))
			'l_cPUEmail = Trim(oRs.Fields("sf_email"))
			l_cPUPhone = Trim(oRs.Fields("fl_sf_phone")) 
			l_cPUAddr1 = Trim(oRs.Fields("fl_sf_addr1"))
			l_cPUAddr2 = Trim(oRs.Fields("fl_sf_addr2"))
			l_cPUCity = Trim(oRs.Fields("fl_sf_city"))
			l_cPUState = Trim(oRs.Fields("fl_sf_state"))
			l_cPUZip = Trim(oRs.Fields("fl_sf_zip"))
			l_cPUComm = Trim(oRs.Fields("fl_sf_comment"))
			l_tReady = oRs.Fields("fl_sf_rta")
			l_tBookDate = oRs.Fields("fh_ready")
			l_cReadyDate = DATEPART("m",l_tReady) & "/" & _
							DATEPART("d", l_tReady) & "/" & _
							DATEPART("yyyy", l_tReady)
			'l_cSTarea = Trim(oRs.Fields("STArea")) 
			'l_cTotalRate = Trim(oRs.Fields("fl_totrate")) 					
			
			' For the drop-down this is only a 12-hour hour
			l_nReadyHour = HOUR(l_tReady)
			IF l_nReadyHour = 12 THEN
				' This is exactly 12:xx
					' This is PM (=1)
					l_cReadyHour = CSTR((l_nReadyHour-1))
					' 0=AM 1=PM
					l_cReadyAMPM = "1"
			ELSE
				IF l_nReadyHour > 12 THEN
					' This is PM (=1)
					l_cReadyHour = CSTR((l_nReadyHour-1)-12)
					' 0=AM 1=PM
					l_cReadyAMPM = "1"
				ELSE
					' This is AM
					l_cReadyHour = CSTR((l_nReadyHour)-1)
					' 0=AM 1=PM
					l_cReadyAMPM = "0"
				END IF
			END IF
		
			' For the drop-down this is only a 12-hour hour
			l_cReadyMin = CSTR(Minute(l_tReady))

			l_cDRID = Trim(oRs.Fields("fl_st_id"))
			l_cDRCompany = Trim(oRs.Fields("fl_st_name"))
			l_cDRContact = Trim(oRs.Fields("fl_st_cfname")) & " " & Trim(oRs.Fields("fl_st_clname"))
			'l_cDREmail = Trim(oRs.Fields("st_email"))
			l_cDRPhone = Trim(oRs.Fields("fl_st_phone")) 
			l_cDRAddr1 = Trim(oRs.Fields("fl_st_addr1"))
			l_cDRAddr2 = Trim(oRs.Fields("fl_st_addr2"))
			l_cDRCity = Trim(oRs.Fields("fl_st_city"))
			l_cDRState = Trim(oRs.Fields("fl_st_state"))
			l_cDRZip = Trim(oRs.Fields("fl_st_zip"))
			l_cDRComm = Trim(oRs.Fields("fl_st_comment"))

			l_cReference = Trim(oRs.Fields("fh_custpo"))
			l_cPieces = Trim(oRs.Fields("fl_numboxes"))
			l_cPieceType = Trim(oRs.Fields("fl_boxtype"))
			l_cWeight = Trim(oRs.Fields("fl_weight"))
			'l_cPriority = Trim(oRs.Fields("fp_desc"))
			l_cUser4 = Trim(oRs.Fields("fh_user4"))
			duetime = Trim(oRs.Fields("fl_st_rta"))
			'l_cSFarea = Trim(oRs.Fields("SFArea")) 
			
		END IF
	
	IF l_cStatus = "quo" THEN
		' This is only a quote
		l_cJobStr = "Quote"
	ELSE
		l_cJobStr = D_JOB
	END IF

	' Use border only for testing






    %>
    <TABLE WIDTH=100% BORDER=1 Align='left' cellpadding="0" cellspacing="0">
        <tr>
            <td colspan="2">
            <table width="100%"><tr><td align="left"><IMG SRC='../Images/LogistiCorpLogo.JPG'></td><td align="center"><b><h2><%=l_cBillTo%></h2></b></td><td align="right"><IMG SRC='../Images/LogistiCorpLogo.JPG'></td></tr></table>
            </td>
        </tr>
        <TR>
            <TD width="50%" rowspan="2" valign="top"><FONT size='+2'><b><%=l_cJobStr%></b>&nbsp;&nbsp;<%=l_cJobNum%></FONT><br /><font size='7'><b><%=l_cDRID%></b></font></TD>

            <TD width="50%"><FONT FACE='Arial'><b>Booked Date:</b> <%=l_tBookdate%></TD>
			
         </TR>
         <TR>
            <TD valign="top"><FONT FACE='Arial'><b>Ready Date/Time:</b> <%=l_tReady%><BR><b>Due Date/Time:</b> <%=duetime%></FONT></TD>
            </TR>
	        <TR>
                <TD><FONT face='Arial'><b>Special Instructions:</b> <%=l_cPUComm%></FONT></td>
		        <TD></TD>
	        </TR>
            <TR BGCOLOR=YELLOW>
                <TD ALIGN=CENTER><FONT FACE='Arial'><b>DROP ZONE</b></FONT></TD>
                <TD ALIGN=CENTER><FONT FACE='Arial'><b>ORIGINATION</b></FONT></TD>
            </TR>
            <TR><TD><%=l_cDRCompany%><BR><%=l_cDRAddr1%>
            <%IF TRIM(l_cDRAddr2) <> "" THEN%>
            	Response.Write("<BR>" & l_cDRAddr2)
	        <%END IF%>
            <BR><%=l_cDRCity%>, <%=l_cDRState%> &nbsp;&nbsp;<%=l_cDRZip%>
            </TD>
            <TD valign='top'><%=l_cPUCompany%><BR><%=l_cPUAddr1%>
            <%IF TRIM(l_cPUAddr2) <> "" THEN
		        Response.Write("<BR>" & (l_cPUAddr2))
	        END IF%>
            <BR><%=l_cPUCity%>, <%=l_cPUState%> &nbsp;&nbsp;<%=l_cPUZip%>
            </TD></TR>
            <%
		        l_cSQL = "SELECT rf_ref " & _
				"FROM fcrefs " & _
				"WHERE rf_fh_id = '" & l_cjobnum & "' " & _
				"ORDER BY rf_pkey "
		        Set oRefs = oConn.Execute(l_cSQL)
                %>
                <TR>
                    <TD VALIGN=TOP align='left'><FONT FACE='Arial'><b>Reference Number(s):</b><br></FONT></td>
		            <TD VALIGN=TOP>
                        <FONT FACE='Arial'><b>Pieces:</b> <%=l_cPieces%></FONT>
                    </TD>
                </tr>
                <tr><td colspan="2">

                    <%
'''''''''''''''''''''''''''''''''''''''''			
'Code 39 barcodes require an asterisk as the start and stop characters
			BarCodeText=l_cReference
		IF NOT oRefs.EOF THEN
			DO WHILE NOT oRefs.EOF
				BarCodeText=oRefs.Fields("rf_ref")

                Response.Write "&nbsp;&nbsp;&nbsp;"
                Response.Write BarCodeText
                Response.write "<br><br>&nbsp;&nbsp;&nbsp;<IMG SRC=""barcodes/asterisk.gif"" WIDTH=""17"" HEIGHT=""14"">"
                For x = 1 to Len(Trim(BarCodeText))
	                DisplayBarCode=mid(BarCodeText,x,1)
	                If DisplayBarCode="/" then
		                Response.write "<IMG SRC=""barcodes/!slash.gif"" WIDTH=""17"" HEIGHT=""14"">"
		                else
		                Response.Write "<IMG SRC=""barcodes/" & DisplayBarCode & _
                                 ".gif"" WIDTH=""17"" HEIGHT=""14"">"
                    End if
                Next

                'Code 39 barcodes require an asterisk as the start and stop characters
                Response.write "<IMG SRC=""barcodes/asterisk.gif"" WIDTH=""17"" HEIGHT=""14"">"
'''''''''''''''''''''''''''''''''''''''''''					
				oRefs.MoveNext
			LOOP
		END IF
		set oRefs = nothing	
        %><br /><br />
		</TD>
    </TR>
            <%
		l_nInsAmt = 0
	END IF%>
    <TR><TD align=center colspan="2"><b>Please attach this form to the material</b></TD></TR>
    <TR height><TD align=center colspan="2" valign="center"><br /><form ID="Form1"><br><input type="button" value=" Print "
			onclick="window.print();return false;" ID="Button1" NAME="Button1"/></form> <br /><br /></TD></TR>
</TABLE>
        <%'END IF
    ELSE
	%>
    Either your session timed-out due to inactivity, or this jobnumber doesn't exist
    <%
END IF '''' EndIf 1 







%>
<SCRIPT LANGUAGE="VBSCRIPT">
	sub  PrintButton_OnClick()
		window.print()
	End Sub
</SCRIPT>
