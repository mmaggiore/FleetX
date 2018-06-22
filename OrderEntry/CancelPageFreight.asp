<html>
<head>

<!-- #include file="../fleetexpress.inc" -->
<link rel="stylesheet" type="text/css" href="../css/Style.css">
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
    PageTitle="ORDER CANCELLATION"
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''BillToID=Session("Suid")
'''If BillToID="" then
'''	BillToID=Request.QueryString("BillToID")
'''End if
'''Session("sBT_ID")=BillToID
'''whatevah=Session("sBT_ID")
'''BillToName=trim(Session("sUsername"))
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''


i=0
ii=0
Response.buffer = True
changedirectory="../marketing/"
PageNameText="Cancel an Order"
Submit=Request.Form("Submit")
'ResetButton=Request.Form("ResetButton")
'Response.Write "XXXXXSubmit="&Submit&"<BR>"
If Submit="" then
	Submit=Request.QueryString("Submit")
End if
'Response.Write "YYYYYYSubmit="&Submit&"<BR>"
If Submit>"" then Submit2=Submit end if
If ResetButton<>"clear search" then
	DateSentFrom=Request.Form("DateSentFrom")
	DateSentTo=Request.Form("DateSentTo")
	DocumentNumber=trim(Request.Form("DocumentNumber"))
	If DocumentNumber="" then
		DocumentNumber=trim(Request.QueryString("DocumentNumber"))
	End if
	LotNumber=trim(Request.Form("LotNumber"))
	If LotNumber="" then
		LotNumber=Trim(Request.QueryString("LotNumber"))
	End if
	SortBy=Request.Form("SortBy")
	ToLocation=Request.Form("ToLocation")
	FromLocation=Request.Form("FromLocation")
	JobNumber=trim(Request.Form("JobNumber"))
	If JobNumber="" then
		JobNumber=trim(Request.QueryString("JobNumber"))
	End if
	ReferenceNumber=trim(Request.Form("ReferenceNumber"))
	Priority=Request.Form("Priority")
	JobStatus=Request.Form("JobStatus")
End if
'Response.Write "Priority="&Priority&"<BR>"
'Response.Write "BillToID="&BillToID&"<BR>"
Select Case BillToID
	Case 48 'KWE
		LotWord="HAWB Number"
	Case 36 'WAFER
		LotWord="Lot Number"
	Case 38, 72 'RETICLES
		LotWord="Reticle Number"
	Case 13, 14, 25 'ABBOTT ROSS
		LotWord="BOL Number"		
	Case 26 'RETICLES
		LotWord="Document Number"
	Case 75 'TI-AIMS
		LotWord="PO Number"	
	Case 76 'TOPAN
		LotWord="Reticle Number"
		'response.Write "Got here 1<BR>"				
	Case else
		LotWord="Document Number"
		'response.Write "Got here 2Based on previous delivery<BR>"			
End Select
'If DateSent="" then
	'DateSent=Date()
	'else
If Submit="" or DateSentFrom="" or DateSentTo="" then
	DateSentFrom=Date()-7
	DateSentTo=Date()
End if
If DateSentTo>"" then
	SQLDateSentTo=cDate(DateSentTo)+1
End if
	''Response.write "DateSent="&DateSent&"<BR>"
	''Response.write "DayAfter="&DayAfter&"<BR>"
'End if
AttemptedCancel=trim(Request.Form("AttemptedCancel"))
CancelJob=trim(Request.Form("CancelJob"))
CancelLot=trim(Request.Form("CancelLot"))
If CancelJob>"" then
    DisplayText=" Job #"&CancelJob
End if
If CancelLot>"" then
    DisplayText=" Lot/Reticle #"&CancelLot
End if

'Response.write "CancelJob="& CancelJob & "<BR>"
'Response.write "CancelLot="& CancelLot & "<BR>"
XID=trim(Request.Form("XID"))
Reason=Request.Form("Reason")
OtherReason=Request.Form("OtherReason")
OtherReason=Replace(OtherReason,"""","")
OtherReason=Replace(OtherReason,"'","")

DocumentNumber=Replace(DocumentNumber,"""","")
DocumentNumber=Replace(DocumentNumber,"'","")
DocumentNumber=Replace(DocumentNumber," ","")

LotNumber=Replace(LotNumber,"""","")
LotNumber=Replace(LotNumber,"'","")
LotNumber=Replace(LotNumber," ","")
JobNumber=Replace(JobNumber,"""","")
JobNumber=Replace(JobNumber,"'","")
JobNumber=Replace(JobNumber," ","")
fh_id=trim(request.form("fh_id"))
'''suid=Session("suid")
'Response.write "fh_id="&fh_id&"<BR>"
ReferenceNumber=Replace(ReferenceNumber,"""","")
ReferenceNumber=Replace(ReferenceNumber,"'","")
ReferenceNumber=Replace(ReferenceNumber," ","")
'''Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
'''	RSEVENTS2.CursorLocation = 3
'''	RSEVENTS2.CursorType = 3
'''	RSEVENTS2.ActiveConnection = Database
'''	l_csql = "SELECT bt_fhu5req FROM fcbillto WHERE (bt_id='"&BillToID&"')"
'''	Response.write("LINE 147 Query:" & l_cSQL)
'''	RSEVENTS2.Open l_cSQL, Database, 1, 3
'''	If not RSEVENTS2.EOF then
'''		UsesLots=RSEVENTS2("bt_fhu5req")
'''		Else
'''		ErrorMessage="You must log out and log back in."	
'''	End if
'''	RSEVENTS2.close
'''Set RSEVENTS2 = Nothing
'Response.write "USESLOTS="&USESLOTS&"<BR>"
'Response.write "AttemptedCancel="&AttemptedCancel&"*<BR>"
If AttemptedCancel="y" then
    If trim(XID)="" then ErrorMessage="You must provide your XID" end if
    If trim(Reason)="" then ErrorMessage="You must select a reason for cancellation" end if
    If trim(Reason)="Other" and trim(OtherReason)="" then ErrorMessage="You must provide an explanation for this cancellation" end if
    'Response.write "hello?<br>"
End if
If trim(CancelLot)>"" and trim(errormessage)="" and AttemptedCancel="y" then
            '''''''''''CHECKS TO SEE IF MORE THAN ONE LOT LEFT...IF NOT, THEN IT CANCELS THE WHOLE JOB!!!!'''''''''''''''
	        Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
		        RSEVENTS.CursorLocation = 3
		        RSEVENTS.CursorType = 3
                'Response.write "DATABASE="&DATABASE&"<BR>"
		        RSEVENTS.ActiveConnection = DATABASE
		        'SQL = "SELECT fh_status, fh_ship_dt, fh_ready, fh_co_id, fh_co_phone, fh_co_email, fh_co_costcenter, fh_custpo, fh_priority FROM fcfgthd where (fh_id = '"& jobnum &"') AND (fh_bt_id='"& BillToID &"') and (fh_status='SCD')"
                SQL = "SELECT count(rf_ref) as HowManyValidRefs FROM fcrefs where (rf_fh_id = '"& fh_id &"') AND ((ref_status<>'X') or (ref_status is NULL))"
		        'Response.Write "SQL="&SQL&"<BR>"
		        RSEVENTS.Open SQL, DATABASE, 1, 3
               'if RSEVENTS.eof then
               '    jobnum=0
               'End if
                HowManyValidRefs=RSEVENTS("HowManyValidRefs")
                'Response.write "HowManyValidRefs="&HowManyValidRefs&"<BR>"
		        RSEVENTS.close
	        Set RSEVENTS = Nothing




            If HowManyValidRefs>1 then






			    Set oConn = Server.CreateObject("ADODB.Connection")
			    oConn.ConnectionTimeout = 100
			    oConn.Provider = "MSDASQL"
			    oConn.Open DATABASE
				    l_cSQL = "UPDATE fcrefs SET ref_status = 'X'" 
                    l_cSQL = l_cSQL&" WHERE (rf_ref = '"& CancelLot &"') and rf_fh_id='"& fh_id & "'"
				    oConn.Execute(l_cSQL)
			    oConn.close
			    Set oConn = nothing
                '''''RESETTING THE ORDER STATUS TO NEW PER TI'S REQUEST 5/3/11
			    Set oConn = Server.CreateObject("ADODB.Connection")
			    oConn.ConnectionTimeout = 100
			    oConn.Provider = "MSDASQL"
			    oConn.Open DATABASE
				    l_cSQL = "UPDATE fcrefs SET ref_status = NULL" 
                    l_cSQL = l_cSQL&" WHERE (Ref_status<>'X') and rf_fh_id='"& fh_id & "'"
				    oConn.Execute(l_cSQL)
			    oConn.close
			    Set oConn = nothing
                'Response.write "l_cSQL="&l_cSQL&"<BR>"
			    Set oConn = Server.CreateObject("ADODB.Connection")
			    oConn.ConnectionTimeout = 100
			    oConn.Provider = "MSDASQL"
			    oConn.Open DATABASE
				    l_cSQL = "UPDATE Report_Refs SET rf_status = 'X'" 
                    l_cSQL = l_cSQL&" WHERE (rf_ref = '"& CancelLot &"') and rf_fh_id='"& fh_id & "'"
				    oConn.Execute(l_cSQL)
			    oConn.close
			    Set oConn = nothing
			    Set oConn = Server.CreateObject("ADODB.Connection")
			    oConn.ConnectionTimeout = 100
			    oConn.Provider = "MSDASQL"
			    oConn.Open DATABASE
				    l_cSQL = "UPDATE Report_Refs SET rf_status = NULL" 
                    l_cSQL = l_cSQL&" WHERE (Rf_status<>'X') and rf_fh_id='"& fh_id & "'"
				    oConn.Execute(l_cSQL)
			    oConn.close
			    Set oConn = nothing
                'Response.write "l_cSQL="&l_cSQL&"<BR>"
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			    Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				    RSEVENTS2.Open "CancelledOrders", DATABASE, 2, 2
				    RSEVENTS2.addnew
				    RSEVENTS2("XID")=XID
				    if trim(fh_id)>"" then
                        RSEVENTS2("fh_id")=fh_id
                        Else
                        RSEVENTS2("fh_id")=CancelJob
                    End if
                    RSEVENTS2("fh_ref")=CancelLot
				    RSEVENTS2("Reason")=Reason
 				    RSEVENTS2("OtherReason")=OtherReason
				    RSEVENTS2("CancelDate")=now()                   	
				    RSEVENTS2("CancelStatus")="c"								
				    RSEVENTS2.update
				    RSEVENTS2.close			
			    set RSEVENTS2 = nothing
                '''''''''''''''''''GATHERS THE PREVIOUS DUE TIME IN MINUTES TO RESET THE VALUE IF ORDER ISN'T COMPLETELY CANCELLED!
	            Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
		            RSEVENTS.CursorLocation = 3
		            RSEVENTS.CursorType = 3
                    'Response.write "DATABASE="&DATABASE&"<BR>"
		            RSEVENTS.ActiveConnection = DATABASE
		            'SQL = "SELECT fh_status, fh_ship_dt, fh_ready, fh_co_id, fh_co_phone, fh_co_email, fh_co_costcenter, fh_custpo, fh_priority FROM fcfgthd where (fh_id = '"& jobnum &"') AND (fh_bt_id='"& BillToID &"') and (fh_status='SCD')"
                    SQL = "SELECT fcfgthd.fh_ship_dt AS BookTime, fcfgthd.fh_status, fclegs.fl_st_rta AS OriginalDueTime FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id WHERE (fh_id='"& fh_id &"') AND (fclegs.fl_leg_status = 'c')"
		            'Response.Write "SQL="&SQL&"<BR>"
		            RSEVENTS.Open SQL, DATABASE, 1, 3
                   'if RSEVENTS.eof then
                   '    jobnum=0
                   'End if
                    BookTime=RSEVENTS("BookTime")
                    OriginalStatus=RSEVENTS("fh_status")
                    OriginalDueTime=RSEVENTS("OriginalDueTime")
                    OriginalMinutesToShip=DateDiff("n", BookTime, OriginalDueTime)
                    NewDueTime=DateAdd("n", OriginalMinutesToShip, now())

                    'Response.write "******************************************<BR>"
                    'Response.write "OriginalDueTime="&OriginalDueTime&"<BR>"
                    'Response.write "BookTime="&BookTime&"<BR>"
                    'Response.write "OriginalMinutesToShip="&OriginalMinutesToShip&"<BR>"
                    'Response.write "NewDueTime="&NewDueTime&"<BR>"
                    'Response.write "******************************************<BR>"

		            RSEVENTS.close
	            Set RSEVENTS = Nothing
''''''''''''''Continues the changing of the job status to new, per TI 5/3/11'''''''''''''''''
            If trim(OriginalStatus)<>"RAP" then
			    Set oConn = Server.CreateObject("ADODB.Connection")
			    oConn.ConnectionTimeout = 100
			    oConn.Provider = "MSDASQL"
			    oConn.Open DATABASE
				    l_cSQL = "UPDATE fcfgthd SET fh_status='OPN', fh_statcode='3'" 
                    l_cSQL = l_cSQL&" WHERE (fh_id = '"& fh_id &"')"
				    oConn.Execute(l_cSQL)
			    oConn.close
			    Set oConn = nothing

			    Set oConn = Server.CreateObject("ADODB.Connection")
			    oConn.ConnectionTimeout = 100
			    oConn.Provider = "MSDASQL"
			    oConn.Open DATABASE
				    l_cSQL = "UPDATE fclegs SET fl_st_rta='"& NewDueTime &"'" 
                    l_cSQL = l_cSQL&" WHERE (fl_fh_id = '"& fh_id &"')"
				    oConn.Execute(l_cSQL)
			    oConn.close
			    Set oConn = nothing

			    Set oConn = Server.CreateObject("ADODB.Connection")
			    oConn.ConnectionTimeout = 100
			    oConn.Provider = "MSDASQL"
			    oConn.Open DATABASE
				    l_cSQL = "UPDATE report_data SET fh_status='OPN', fl_st_rta='"& NewDueTime &"'" 
                    l_cSQL = l_cSQL&" WHERE (fh_id = '"& fh_id &"')"
				    oConn.Execute(l_cSQL)
			    oConn.close
			    Set oConn = nothing
        End if
                SuccessMessage="You have successfully cancelled lot/reticle #"&CancelLot&"<BR>"
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ELSE
			    Set oConn = Server.CreateObject("ADODB.Connection")
			    oConn.ConnectionTimeout = 100
			    oConn.Provider = "MSDASQL"
			    oConn.Open DATABASE
				    l_cSQL = "UPDATE fcfgthd SET fh_status='CAN', fh_statcode='98'" 
                    l_cSQL = l_cSQL&" WHERE (fh_id = '"& fh_id &"')"
				    oConn.Execute(l_cSQL)
			    oConn.close
			    Set oConn = nothing
			    Set oConn = Server.CreateObject("ADODB.Connection")
			    oConn.ConnectionTimeout = 100
			    oConn.Provider = "MSDASQL"
			    oConn.Open DATABASE
				    l_cSQL = "UPDATE Report_Data SET fh_status='CAN'" 
                    l_cSQL = l_cSQL&" WHERE (fh_id = '"& fh_id &"')"
				    oConn.Execute(l_cSQL)
			    oConn.close
			    Set oConn = nothing
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			    Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				    RSEVENTS2.Open "CancelledOrders", DATABASE, 2, 2
				    RSEVENTS2.addnew
				    RSEVENTS2("XID")=XID
				    RSEVENTS2("fh_id")=CancelJob
                    RSEVENTS2("fh_ref")=CancelLot
				    RSEVENTS2("Reason")=Reason
 				    RSEVENTS2("OtherReason")=OtherReason
				    RSEVENTS2("CancelDate")=now()                   	
				    RSEVENTS2("CancelStatus")="c"								
				    RSEVENTS2.update
				    RSEVENTS2.close			
			    set RSEVENTS2 = nothing
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                SuccessMessage="You have successfully cancelled job #"&CancelJob&"<BR>"
            End if
    'Response.write "Put cancel lot code here...don't forget if only one lot to cancel the whole job instead<br>"
        XID=""
        Reason=""
        OtherReason=""
        AttemptedCancel="n"
        CancelLot=""
        CancelJob=""
End if

If trim(CancelJob)>"" and trim(errormessage)="" and AttemptedCancel="y" then
    'Response.write "Put cancel job code here!<br>"
			Set oConn = Server.CreateObject("ADODB.Connection")
			oConn.ConnectionTimeout = 100
			oConn.Provider = "MSDASQL"
			oConn.Open DATABASE
				l_cSQL = "UPDATE fcfgthd SET fh_status = 'CAN', fh_statcode='98'" 
                l_cSQL = l_cSQL&" WHERE (fh_id = '"& CancelJob &"')"
				oConn.Execute(l_cSQL)
			oConn.close
			Set oConn = nothing
			Set oConn = Server.CreateObject("ADODB.Connection")
			oConn.ConnectionTimeout = 100
			oConn.Provider = "MSDASQL"
			oConn.Open DATABASE
				l_cSQL = "UPDATE Report_Data SET fh_status = 'CAN'" 
                l_cSQL = l_cSQL&" WHERE (fh_id = '"& CancelJob &"')"
				oConn.Execute(l_cSQL)
			oConn.close
			Set oConn = nothing
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				RSEVENTS2.Open "CancelledOrders", DATABASE, 2, 2
				RSEVENTS2.addnew
				RSEVENTS2("XID")=XID
				RSEVENTS2("fh_id")=CancelJob
                RSEVENTS2("fh_ref")=CancelLot
				RSEVENTS2("Reason")=Reason
 				RSEVENTS2("OtherReason")=OtherReason
				RSEVENTS2("CancelDate")=now()                   	
				RSEVENTS2("CancelStatus")="c"								
				RSEVENTS2.update
				RSEVENTS2.close			
			set RSEVENTS2 = nothing

'------------DRIVER MESSAGE ABOUT CANCELLATION

            'Response.write "GOT HERE 6!!!<BR>"

	        Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
		        RSEVENTS.CursorLocation = 3
		        RSEVENTS.CursorType = 3
                'Response.write "DATABASE="&DATABASE&"<BR>"
		        RSEVENTS.ActiveConnection = DATABASE
		        'SQL = "SELECT fh_status, fh_ship_dt, fh_ready, fh_co_id, fh_co_phone, fh_co_email, fh_co_costcenter, fh_custpo, fh_priority FROM fcfgthd where (fh_id = '"& jobnum &"') AND (fh_bt_id='"& BillToID &"') and (fh_status='SCD')"
                SQL = "SELECT fl_dr_id, fl_sf_cfname, fl_sf_building, fl_sf_addr1, fl_sf_addr2 FROM fclegs where (fl_fh_id = '"& fh_id &"')"
		        
                'Response.Write "SQL="&SQL&"<BR>"
		        
                RSEVENTS.Open SQL, DATABASE, 1, 3
               'if RSEVENTS.eof then
               '    jobnum=0
               'End if
               If not RSEVENTS.eof then

                'Response.write "GOT HERE #7<BR>"

                TheDriverID=RSEVENTS("fl_dr_id")
                fl_sf_cfname=RSEVENTS("fl_sf_cfname")
                fl_sf_building=RSEVENTS("fl_sf_building")
                fl_sf_addr1=RSEVENTS("fl_sf_addr1")
                fl_sf_addr2=RSEVENTS("fl_sf_addr2")

                'rESPONSE.WRITE "TheDriverID="&TheDriverID&"<BR>"

	            Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
		            RSEVENTS2.Open "DriverMessages", INTRANET, 2, 2
		            RSEVENTS2.addnew
		            RSEVENTS2("DriverMessage")="Job #"&fh_id&" from "& fl_sf_cfname &" "& fl_sf_building &" "& fl_sf_addr1 &" "& fl_sf_addr2 &" has been cancelled."
		            RSEVENTS2("MessageDate")=Date()
		            RSEVENTS2("MessageOriginator")="1"
                    RSEVENTS2("MessageRecipient")=TheDriverID
		            RSEVENTS2("MessageStatus")="c"								
		            RSEVENTS2.update
		            RSEVENTS2.close			
	            set RSEVENTS2 = nothing	

                'Response.write "THE END!!!!<BR>"
                
                End if
		        RSEVENTS.close
	        Set RSEVENTS = Nothing
'------END DRIVER CANCELLATION MESSAGE


            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
           SuccessMessage="You have successfully cancelled job #"&CancelJob&"<BR>"
            XID=""
            Reason=""
            OtherReason=""
            AttemptedCancel="n"
            CancelLot=""
            CancelJob=""
End if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


%>
<title>FleetX - <%=PageTitle %></title>
</head>

<body onload="document.OrderForm1.<%=HighlightedField%>.focus();">
	<table align="center" border="0" bordercolor="black" cellpadding="0" cellspacing="0" ID="Table1" width="100%" >
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
	<tr height="1">
		<td width="650">&nbsp;</td>
	</tr>
    <tr><td align="center" width="100%"><!-- main page stuff goes here! -->

    
 	<TABLE WIDTH="100%"  border="0" bordercolor="red" cellpadding="0" cellspacing="5" ID="Table2">
        <%If trim(CancelLot)>"" or trim(CancelJob)>"" then %>
        <tr><td>&nbsp;</td></tr>
            <table cellpadding="0" cellspacing="0" border="1" bordercolor="red" align="center" ID="Table4" width="650">
                <tr><td>
                    <table cellpadding="3" cellspacing="0" border="0" bordercolor="red" align="center" ID="Table6">
                        <tr><td>&nbsp;</td></tr>
                        <tr>
                            <td colspan="2"><b>Are you sure that you wish to cancel <%=DisplayText %>?</b>
                            </td>
                        </tr>
                        <form method="post">
                        <input type="hidden" name="SortBy" value="<%=SortBy%>" />
                        <input type="hidden" name="ToLocation" value="<%=ToLocation%>" />
                        <input type="hidden" name="FromLocation" value="<%=FromLocation%>" />
                        <input type="hidden" name="Priority" value="<%=Priority%>" />
                        <input type="hidden" name="JobNumber" value="<%=JobNumber%>" />
                        <input type="hidden" name="LotNumber" value="<%=LotNumber%>" />
                        <tr><td colspan="2">&nbsp;</td></tr>
                        <tr><td colspan="2">If no, click the "NO" button---->&nbsp;&nbsp;&nbsp;&nbsp;<input name="submit" type="submit" id="gobutton" value="     NO     " /></td></tr>
                        <tr><td colspan="2">&nbsp;</td></tr>
                        </form>
                        <form method="post">
                        <tr><td colspan="2">If yes, provide your XID and cancellation reason below and click the "Submit" button:</td></tr>
                        <tr><td colspan="2">&nbsp;</td></tr>
                        <tr><td align="right">XID:</td><td align="left"><input name="XID" type="text" size="20" maxlength="20" value="<%=XID %>" /></td></tr>
						<tr>
                            <tr><td align="right">Reason for cancellation:</td> 
							<td align="left">
								<select name="reason" ID="Select5">
									<option value="">Select a Reason for Cancellation</option>
									<option value="Incorrect Lot/Reticle Number" <%if reason="Incorrect Lot/Reticle Number" then response.Write " selected" end if%>>Incorrect Lot/Reticle Number</option>
									<option value="Duplicate Lot/Reticle Number" <%if reason="Duplicate Lot/Reticle Number" then response.Write " selected" end if%>>Duplicate Lot/Reticle Number</option>
									<option value="Incorrect Destination" <%if reason="Incorrect Destination" then response.Write " selected" end if%>>Incorrect Destination</option>
									<option value="Incorrect Origination" <%if reason="Incorrect Origination" then response.Write " selected" end if%>>Incorrect Origination</option>
									<option value="Incorrect Priority" <%if reason="Incorrect Priority" then response.Write " selected" end if%>>Incorrect Priority</option>
									<option value="<%=LotWord%> Not Available at Pick Up Location" <%if reason=LotWord&" Not Available at Pick Up Location" then response.Write " selected" end if%>><%=LotWord%> Not Available at Pick Up Location</option>
									<option value="Other" <%if reason="Other" then response.Write " selected" end if%>>Other</option>
                                    <option value="Developer Testing" <%if reason="Developer Testing" then response.Write " selected" end if%>>Developer Testing</option>
								</select>
							</td>
						</tr>
						<%
						If reason="Other" then
						%>
						<tr><td>&nbsp;</td></tr>
						<tr>
                            <td align="right">Explanation:</td>
							<td align="left">
								<input type="text" name="OtherReason" value="<%=OtherReason%>" size="50" maxlength="50" ID="Text1">
							</td>
						</tr>
						<%
						End if
						%>
                        <input type="hidden" name="fh_id" value="<%=fh_id%>" />
                        <input type="hidden" name="SortBy" value="<%=SortBy%>" />
                        <input type="hidden" name="ToLocation" value="<%=ToLocation%>" />
                        <input type="hidden" name="FromLocation" value="<%=FromLocation%>" />
                        <input type="hidden" name="Priority" value="<%=Priority%>" />
                        <input type="hidden" name="JobNumber" value="<%=JobNumber%>" />
                        <input type="hidden" name="LotNumber" value="<%=LotNumber%>" />
                        <Input type="hidden" name="CancelLot" value="<%=CancelLot %>" />
                        <Input type="hidden" name="CancelJob" value="<%=CancelJob %>" />
                        <Input type="hidden" name="AttemptedCancel" value="y" />
                        <%If trim(errormessage)>"" then %>
                        <tr><td>&nbsp;</td></tr>
                        <tr><td colspan="2" align="center" class="ErrorMessage">Error:  <%=ErrorMessage %></td></tr>
                        <tr><td>&nbsp;</td></tr>
                        <%end if %>
                        <tr><td colspan="2" align="center"><input name="submit" type="submit" value="Submit" id="gobutton" /></td></tr>
                        <tr><td>&nbsp;</td></tr>
                        </form>
                    </table>
            </table>
            </td>
        </tr>
        <%End if %>
		<table cellpadding="0" cellspacing="0" border="0" bordercolor="blue" align="center" ID="Table3">
			<tr><td>&nbsp;</td></tr>
            <tr><td class = 'FleetXBoldText'>Cancellation Search</td></tr>
            <tr><td class = 'FleetXBoldText'>(Provide as much, or as little information as needed/available)</td></tr>
            <tr><td class = 'FleetXBoldText'>(To see all of your jobs, leave all fields blank and click "SEARCH")</td></tr>
			<tr><td>&nbsp;</td></tr>
            <tr><td class = 'FleetXBoldText'>Note:  You can cancel any order up until the driver has on boarded it</td></tr>
            <tr><td>&nbsp;</td></tr>
			<tr>
				<td>
					<table border="0" bordercolor="red" align="center" ID="Table4">
					<form method="post" action="CancelPageFreight.asp" name="thisForm" ID="Form1">
						
						<tr>
							<td class='subheader' align="right">
							<%=LotWord%>:
							</td>
							<td>
								<input type="text" name="LotNumber" value="<%=LotNumber%>" size="20" ID="Text4">
							</td>					
						</tr>
						<tr>
							<td class='subheader' align="right">
							Job Number:
							</td>
							<td>
								<input type="text" name="JobNumber" value="<%=JobNumber%>" size="20" ID="Text2">
							</td>					
						</tr>												

						<tr>
							<td class='subheader' align="right">
								Priority:
							</td>
							<%'response.Write "ToLocation="&ToLocation&"***<BR>"%>
							<td>
								<select name="Priority" ID="Select4">
									<option value="" <%if Priority="" then response.Write " Selected " end if%>>All Priorities</option>
									<%
									Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
										RSEVENTS2.CursorLocation = 3
										RSEVENTS2.CursorType = 3
										RSEVENTS2.ActiveConnection = Database
										l_csql = "SELECT DISTINCT PriorityDescription AS PriorityDescription, PriorityMinutes AS PriorityTime, Priority AS PriorityAbbreviation FROM priorities WHERE (Priority_BT_ID IN(92,93)) order by Priority"
										'response.write("Query:" & l_cSQL)
										RSEVENTS2.Open l_cSQL, Database, 1, 3
										Do while not RSEVENTS2.EOF
											PriorityDescription=RSEVENTS2("PriorityDescription")
											PriorityAbbreviation=RSEVENTS2("PriorityAbbreviation")
										%>
											<option value="<%=PriorityAbbreviation%>" <%if PriorityAbbreviation=Priority then response.Write "selected" end if%>><%=PriorityDescription%></option>
										<%	
										RSEVENTS2.movenext
										LOOP
										RSEVENTS2.close
									Set RSEVENTS2 = Nothing							
									%>
								</select>
								<%
								'Response.Write "l_csql="&l_csql&"<BR>"
                                'response.write "userid="&UserID&"<BR>"
								%>
							</td>
						</tr>						
						<input type="hidden" name="JobStatus" value="OPEN" />
						<tr>
							<td class='subheader' align="right">
								Origination: 
							</td>
							<td>
								<select name="FromLocation" ID="Select2">
                                    <%
                                    If 	ucase(trim(temp_st_id))="PHO" or ucase(trim(temp_st_id))="TOPPAN" or ucase(trim(temp_st_id))="CPGP" or ucase(trim(temp_st_id))="DNP" or  ucase(trim(temp_st_id))="CPGPSCOT" or  ucase(trim(temp_st_id))="TOPPANSC" then
                                    ELSE
                                    %>
									<option value="" <%if FromLocation="" then response.Write " Selected " end if%>>All Locations</option>
									<%
                                    End if
                                    
									Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
										RSEVENTS2.CursorLocation = 3
										RSEVENTS2.CursorType = 3
										RSEVENTS2.ActiveConnection = Database
										l_csql = "SELECT companyname as st_name, st_id, companyaddress, companybuilding, companysuite, contactname  FROM PreExistingCompanies WHERE (companyowner='"&UserID&"')"
										If BilltoID=36 then
										    l_csql=l_csql&" OR st_id='TISHERMA' "
										End if										
										If BillToID=48 then
											l_csql=l_csql&" AND (St_Priapt='DFW')"
										End if
										If BillToID=38 then
											l_csql=l_csql&" AND (st_id<>'CPGP')"
										End if
										If BillToID=76 then
											l_csql=l_csql&" OR (st_id='TOPPAN')"
										End if	
                                        If 	ucase(trim(temp_st_id))="PHO" or ucase(trim(temp_st_id))="TOPPAN" or ucase(trim(temp_st_id))="CPGP" or ucase(trim(temp_st_id))="DNP" or  ucase(trim(temp_st_id))="CPGPSCOT" then
                                            l_csql=l_csql&" AND (st_id='"& trim(temp_st_id) &"')"
                                        End if										
										l_csql=l_csql&" ORDER BY contactname"										
										'response.write("Query2:" & l_cSQL)
										
										RSEVENTS2.Open l_cSQL, Database, 1, 3
										Do while not RSEVENTS2.EOF
										st_id=RSEVENTS2("st_id")
										st_name=RSEVENTS2("st_name")
                                        contactname=RSEVENTS2("contactname")
                                        CompanyBuilding=RSEVENTS2("CompanyBuilding")
                                        companyaddress=RSEVENTS2("companyaddress")
                                        CompanySuite=RSEVENTS2("CompanySuite")
										%>
										
											<option value="<%=st_id%>" <%if st_id=FromLocation then response.Write " Selected " end if%>><%=contactname %>/<%=st_name%>/<%=CompanyBuilding %>/<%=companyaddress %>/<%=CompanySuite %><</option>
										
										<%	
										RSEVENTS2.movenext
										LOOP
										RSEVENTS2.close
									Set RSEVENTS2 = Nothing							
									%>
								</select>
								<%
								'Response.Write "l_csql="&l_csql&"<BR>"
								%>
							</td>
						</tr>						
						<tr>
							<td class='subheader' align="right">
								Destination:
							</td>
							<%'response.Write "ToLocation="&ToLocation&"***<BR>"%>
							<td>
								<select name="ToLocation" ID="Select1">
									<option value="" <%if ToLocation="" then response.Write " Selected " end if%>>All Locations</option>
									<%
									Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
										RSEVENTS2.CursorLocation = 3
										RSEVENTS2.CursorType = 3
										RSEVENTS2.ActiveConnection = Database
										l_csql = "SELECT companyname as st_name, st_id, companyaddress, companybuilding, companysuite, contactname  FROM PreExistingCompanies WHERE (companyowner='"&UserID&"')"
										If BilltoID=36 then
										    l_csql=l_csql&" OR st_id='TISHERMA' "
										End if
										If BillToID=48 then
											l_csql=l_csql&" AND (St_Name<>'KWE')"
										End if
										If BillToID=38 then
											l_csql=l_csql&" AND (st_id<>'55')"
										End if
										If BillToID=76 then
											l_csql=l_csql&" OR (st_id='TOPPAN')"
										End if										
										l_csql=l_csql&" ORDER BY st_name"																				
										'response.write("Query:" & l_cSQL)
										
										RSEVENTS2.Open l_cSQL, Database, 1, 3
										Do while not RSEVENTS2.EOF
										st_id=RSEVENTS2("st_id")
										st_name=RSEVENTS2("st_name")
                                        contactname=RSEVENTS2("contactname")
                                        CompanyBuilding=RSEVENTS2("CompanyBuilding")
                                        companyaddress=RSEVENTS2("companyaddress")
                                        CompanySuite=RSEVENTS2("CompanySuite")
										%>
										
											<option value="<%=st_id%>" <%if st_id=ToLocation then response.Write " Selected " end if%>><%=contactname %>/<%=st_name%>/<%=CompanyBuilding %>/<%=companyaddress %>/<%=CompanySuite %><</option>
										
										<%	
										RSEVENTS2.movenext
										LOOP
										RSEVENTS2.close
									Set RSEVENTS2 = Nothing							
									%>
								</select>
								<%
								'Response.Write "l_csql="&l_csql&"<BR>"
								%>
							</td>
						</tr>

						<tr>
							<td class='subheader' align="right">
								Sort By:
							</td>
							<td>
								<select name="SortBy" ID="Select3">
								<%If BillToID=26 then%>
									<option value="fh_custpo asc" <%if SortBy="fh_custpo asc" then response.Write " Selected " end if%>><%=LotWord%> (Ascending)</option>
									<option value="fh_custpo desc" <%if SortBy="fh_custpo desc" then response.Write " Selected " end if%>><%=LotWord%> (Descending)</option>
									<option value="fl_sf_rta asc" <%if SortBy="" or SortBy="fl_sf_rta asc" or SortBy="" then response.Write " Selected " end if%>>SAP Order Time (Ascending)</option>									
									<option value="fl_sf_rta desc" <%if SortBy="fl_sf_rta desc" then response.Write " Selected " end if%>>SAP Order Time (Descending)</option>									
									
									<%
									else
									%>
									<option value="rf_ref asc" <%if SortBy="rf_ref asc" then response.Write " Selected " end if%>><%=LotWord%>  (Ascending)</option>
									<option value="rf_ref desc" <%if SortBy="rf_ref desc" then response.Write " Selected " end if%>><%=LotWord%> (Descending)</option>
									<option value="fh_ship_dt asc" <%if SortBy="fh_ship_dt asc" or SortBy="" then response.Write " Selected " end if%>>Booked Time (Ascending)</option>								
									<option value="fh_ship_dt desc" <%if SortBy="" or SortBy="fh_ship_dt desc" then response.Write " Selected " end if%>>Booked Time (Descending)</option>								

								<%end if%>
								</select>
							</td>
						</tr>
						<tr><td>&nbsp;</td></tr>						
						<tr><td><img src="../images/pixel.gif" height="1" width="1" border="0"></td></tr>
						<tr><td align="center" colspan="2"><input type="submit" name="submit" value="search" ID="gobutton"></td></tr>	
                        <tr><td>&nbsp;</td></tr>
                        <%If trim(SuccessMessage)>"" then %>
                        
                        <tr><td colspan="2" align="center"><font color="blue"><b><%=SuccessMessage %></b></font></td></tr>
                        <tr><td>&nbsp;</td></tr>
                        <%end if %>
					</form>
					</table>
				</td>
			</tr>
		</table>
		<%
		'Response.Write "got here #1<BR>"
		'Response.Write "ZZZsubmit2="&submit2&"<BR>"
		If Submit2>"" then
		'Response.Write "got here 2<BR>"
		%>
			<table cellpadding="3" cellspacing="0" border="1" align="center" ID="Table5">
				<tr>
				<%'If UsesLots=FALSE then
						'ColspanNumber="7"
				%>
                <!--
					<td class="SubHeader" nowrap>
						Document Number
					</td>	
                    -->			
					<%'else
						ColspanNumber="8"
					%>
					<td class="SubHeader" nowrap>
						Job Number
					</td>
					<td class="SubHeader" nowrap>
						<%=LotWord%> 
					</td>				
				<%'End if%>
					<!--
					<td class="MainPageTextBoldCentered" nowrap>
						Pickup
					</td>	
					<td class="MainPageTextBoldCentered" nowrap>
						Dropoff
					</td>
					-->
					<td class="SubHeader" nowrap>
						From
					</td>
					<td class="SubHeader" nowrap>
						To
					</td>										
					<td class="SubHeader" nowrap>
						Status
					</td>
					<%If UsesLots=TRUE then%>
					<td class="SubHeader" nowrap>
						Priority
					</td>
					<%End if%>					
					<td class="SubHeader" nowrap>
						Entered
					</td>	
																				
				</tr>		
		<%
			Dim colorset, i, numcolors
			'/--- This is your array of colors to use. -------------\ 
			colorset = split("#D2D9FC,White",",")
			numcolors = ubound(colorset)+1

		
		
			Server.ScriptTimeout = 1000
			optJobSel=Request.Querystring("optJobSel")
			optJobSel=Replace(optJobSel,"""","")
			optJobSel=Replace(optJobSel,"'","")
			If ReferenceNumber>"" then optJobSel="ByRef" end if
			If JobNumber>"" then optJobSel="ByJob" end if
			'Response.write "******optJobSel="&optJobSel&"<BR>"
			Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				RSEVENTS2.CursorLocation = 3
				RSEVENTS2.CursorType = 3
				RSEVENTS2.ActiveConnection = Database
				'Response.write "GOT HERE #2!<BR>"
				'Response.write "Database="&Database&"<BR>"
				''''If USESLOTS=TRUE then
				'Response.write "DateSentFrom="&DateSentFrom&"<BR>"
				'Response.write "DateSentTo="&DateSentTo&"<BR>"
				NumberofDays=datediff("d",DateSentFrom, DateSentTo)
				'Response.write "NumberofDays="&NumberofDays&"<BR>"
					l_csql = "SELECT "
					'If NumberofDays>0 then
					'	l_csql = l_csql&" * "	
					'End if		
					l_csql = l_csql&" fh_id, fh_status, fh_ship_dt, fl_sf_id, fl_sf_name, fl_sf_clname, fl_sf_addr2, fl_st_id, fl_st_name, fl_st_clname, fl_st_addr2, fl_t_acc, fl_t_atp, fl_t_atd, fh_custpo, fh_priority, RF_REF FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN fcrefs ON fclegs.fl_fh_id = fcrefs.rf_fh_id "	
					'l_csql = l_csql&" WHERE fl_st_id=fl_finaldestination "
					''''else
					''''l_csql = "SELECT Distinct(fcfgthd.fh_id), fclegs.fl_sf_rta, fcfgthd.fh_status, fcfgthd.fh_ship_dt, fclegs.fl_sf_id, fclegs.fl_sf_name, fclegs.fl_st_id, fclegs.fl_st_name, fclegs.fl_t_atp, fclegs.fl_t_atd, fclegs.fl_pod, fcfgthd.fh_custpo, fcfgthd.fh_priority FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id "	
				''''End if
				'response.write "BillToName="&BillToName&"***<BR>"
				Select Case BillToName
					Case "compsxxx"
						l_csql = l_csql&" WHERE (((fl_st_id='CPGP') OR (fl_sf_id='55')) "
					Case "Toppanxxx"
						l_csql = l_csql&" WHERE (((fl_st_id='TOPPAN') OR (fl_sf_id='TOPPAN')) "
					Case "tiretxxx"
						l_csql = l_csql&" WHERE (((fh_bt_id='"&BillToID&"') OR ((fh_bt_id<>'26') AND (fh_bt_id<>'36'))) "
					Case else
						'l_csql = l_csql&" WHERE ((fh_bt_id='"&BillToID&"') "
                        l_csql = l_csql&" WHERE ((fh_bt_id IN('92','93')) and fh_user_id ='"& UserID &"' "
				End Select

						'''If trim(ReferenceNumber)>"" then
						'''	l_csql = L_csql& "AND (rf_ref='"&ReferenceNumber&"') "
						'''End if
						'''If trim(DocumentNumber)>"" then
						'''	l_csql = L_csql& "AND (fh_custpo LIKE '%"&DocumentNumber&"') "
						'''End if
						'If trim(LotNumber)>"" then
						'	l_csql = L_csql& "AND (rf_ref LIKE '%"&LotNumber&"%') "
						'End if
						'If trim(JobNumber)>"" then
						'	l_csql = L_csql& "AND (fh_id LIKE '%"&JobNumber&"') "
						'End if	
						If trim(Priority)>"" then
						    If Priority="XP" then
						        l_csql = L_csql& "AND ((fh_priority = 'P1') OR (fh_priority = 'P0') OR (fh_priority = 'XP') ) "
						        else
							    l_csql = L_csql& "AND (fh_priority = '"&Priority&"') "
							End if
						End if							
						'''If trim(JobStatus)>"" then
							'''Select Case JobStatus
								'''Case "9"
									'''l_csql = L_csql& "AND (fh_status = 'CLS') "
								'''Case "98"
									'''l_csql = L_csql& "AND (fh_status = 'CAN') "
								'''Case "OPEN"
									l_csql = L_csql& " AND ((fh_status = 'SCD') or (fh_status = 'RAP') or (fh_status = 'OPN') or (fh_status = 'ACC')) "
							'''End Select
							
						'''End if							
																
						'If DateSentTo>"" And DateSentFrom>"" then
						'	If BillToID=26 then
						'		l_csql = L_csql& "AND (fl_sf_rta>='"&DateSentFrom&"') AND (fl_sf_rta<'"&SQLDateSentTo&"') "
						'		else
						'		l_csql = L_csql& "AND (fh_ship_dt>='"&DateSentFrom&"') AND (fh_ship_dt<'"&SQLDateSentTo&"') "
						'	End if
						'End if
						If ToLocation>"" then
							l_csql = L_csql& "AND (fl_st_id='"&ToLocation&"') "
						End if
						If FromLocation>"" then
							l_csql = L_csql& "AND (fl_sf_id='"&FromLocation&"') "
						End if
							l_csql = L_csql& ") "
						If trim(LotNumber)>"" and trim(JobNumber)="" then
							l_csql = L_csql& "AND (rf_ref LIKE '%"&LotNumber&"%') "
						End if
						If trim(JobNumber)>"" AND trim(LotNumber)="" then
							l_csql = L_csql& "AND (fh_id LIKE '%"&JobNumber&"') "
						End if	
						If trim(JobNumber)>"" AND trim(LotNumber)>"" then
							l_csql = L_csql& "AND ((rf_ref LIKE '%"&LotNumber&"%') AND (fh_id LIKE '%"&JobNumber&"')) "
						End if
                        l_csql = L_csql& "AND ((ref_status<> 'X') OR (ref_status IS NULL)) "							
						GenericSortBy="fh_ship_dt desc"
						If SortBy>"" then
							l_csql = L_csql& " ORDER BY "&Sortby
							else
							l_csql = L_csql& " ORDER BY "&GenericSortby
						End if

					
			'response.write("Query3:" & l_cSQL)
			''''''''''''''''''''''''''''''''''''''''
				RSEVENTS2.Open l_cSQL, Database, 1, 3
				If RSEVENTS2.eof then
						ErrorMessage="Based on your provided criteria, no jobs were found that are available to cancel.<br><br>Please check your criteria and try again.<br><br>Once a job has already been on boarded, then it cannot be cancelled.<br><br>If you need assistance with an order, call 214-882-0620."	
				End if				
				Do while not RSEVENTS2.EOF 
					fh_id=RSEVENTS2("fh_id")
					'fl_sf_rta=RSEVENTS2("fl_sf_rta")
					fh_status=RSEVENTS2("fh_status")
					fh_ship_dt=RSEVENTS2("fh_ship_dt")
					fl_t_acc=trim(RSEVENTS2("fl_t_acc"))
					fl_sf_id=trim(RSEVENTS2("fl_sf_id"))
                    fl_sf_clname=trim(RSEVENTS2("fl_sf_clname"))
                    fl_sf_addr2=trim(RSEVENTS2("fl_sf_addr2"))
					fl_st_id=RSEVENTS2("fl_st_id")	
                    fl_st_clname=trim(RSEVENTS2("fl_st_clname"))
                    fl_st_addr2=trim(RSEVENTS2("fl_st_addr2"))                   				
					fl_sf_name=RSEVENTS2("fl_sf_name")
					fl_st_name=RSEVENTS2("fl_st_name")
					fl_t_atp=RSEVENTS2("fl_t_atp")
					'Response.Write "i="&i&"<BR>"
					'Response.Write "fl_t_atp="&fl_t_atp&"<BR>"
					'''''If ii=0 then
					    FirstONB=fl_t_atp
					''''''End if  
					fl_t_atd=RSEVENTS2("fl_t_atd")
					'fl_pod=RSEVENTS2("fl_pod")
					fh_custpo=RSEVENTS2("fh_custpo")
					fh_priority=RSEVENTS2("fh_priority")
					'fl_sf_rta=RSEVENTS2("fl_sf_rta")
					'fl_finalDestination=RSEVENTS2("fl_finalDestination")
					If USESLOTS=TRUE then
						rf_ref=RSEVENTS2("rf_ref")
						'PODDateTime=RSEVENTS2("PODDateTime")
					End if
			Select Case fl_sf_id
				CASE "55"
					Fl_sf_id="CPGP"
				CASE "72"
					Fl_sf_id="CRI"					
			End Select
			Select Case fh_priority
				Case "WF", "CS"
					Displayfh_Priority="Standard"
				Case "XP"
					Displayfh_Priority="Expedited"					
				Case "AS"
					Displayfh_Priority="Next Day"
				Case "A0"
					Displayfh_Priority="Hot Shot"
				Case "A1"
					Displayfh_Priority="Same Day"															
				Case ELSE
					DisplayFH_Priority=FH_Priority
			End Select
			Select Case fh_status
				Case "RAP"
					Display_fh_status="Booked"			
				Case "CLS"
					Display_fh_status="Closed"
				Case "OPN"
					Display_fh_status="Open"
				Case "ACC"
					Display_fh_status="Accepted"
				Case "PUO"
					Display_fh_status="POB"					
				Case "ONB"
					Display_fh_status="On Board"
				Case "ATD"
					Display_fh_status="At Destination"
				Case "CAN"
					Display_fh_status="Cancelled"
				Case "DEL"
					Display_fh_status="Deleted"	
				Case "ARV"
					Display_fh_status="Arrived At HUB"
				Case "AC2"
					Display_fh_status="Arrived At HUB*"	
				Case "DPV"
					Display_fh_status="Departed HUB"											
				Case Else
					Display_fh_status=fh_status																			
			End Select
			if fh_ship_dt="1/1/1900" then fh_ship_dt="&nbsp;" end if
			if FirstONB="1/1/1900" then FirstONB="&nbsp;" end if
			if fl_t_atd="1/1/1900" then fl_t_atd="&nbsp;" end if
			if fl_t_acc="1/1/1900" then fl_t_acc="&nbsp;" end if
			'If ErrorMessage="" then
			'Response.Write "UsesLots="&UsesLots&"<BR>"
			'Response.Write "("&fh_id&")fl_finaldestination="&fl_finaldestination&"****<BR>"
			'If (trim(fl_st_id)=trim(fl_finaldestination)) Or (isnull(fl_finaldestination)) then
			%>
				<tr>
                    <form action="cancelpagefreight.asp" method="post">
					<td class="MainPageTextSmaller" nowrap valign="top" bgcolor="<%=colorset(i mod numcolors)%>">
						<%If tempfh_id<>fh_id then %>
                        <input type="hidden" name="SortBy" value="<%=SortBy%>" />
                        <input type="hidden" name="ToLocation" value="<%=ToLocation%>" />
                        <input type="hidden" name="FromLocation" value="<%=FromLocation%>" />
                        <input type="hidden" name="Priority" value="<%=Priority%>" />
                        <input type="hidden" name="JobNumber" value="<%=JobNumber%>" />
                        <input type="hidden" name="LotNumber" value="<%=LotNumber%>" />

                        <input type="submit" name="Submit" id="gobutton" value="Cancel Job" /> <%=trim(fh_id)%>
                        <input type="hidden" name="fh_id" value="<%=fh_id%>" />
                        <input type="hidden" name="CancelJob" value="<%=fh_id%>" />
                        <%end if %>
					</td>
                    </form>
                    <form action="cancelpagefreight.asp" method="post">
                    <input type="hidden" name="fh_id" value="<%=fh_id%>" />
                    <input type="hidden" name="rf_ref" value="<%=rf_ref%>" />
                    <input type="hidden" name="fh_custpo" value="<%=fh_custpo%>" />	
                    <input type="hidden" name="CancelLot" value="<%=rf_ref%>" />
					<%
					Select Case BillTOID
					    Case "36", "38"
					        %>
						        <td class="MainPageTextSmaller" nowrap valign="top" bgcolor="<%=colorset(i mod numcolors)%>">
							       <input type="submit" name="Submit" value="Cancel Lot" /> <%=trim(rf_ref)%>
						        </td>				
						        <%
                         Case "38"
						        'Response.Write "mark2..."
						        %>
						        <td class="MainPageTextSmaller" nowrap valign="top" bgcolor="<%=colorset(i mod numcolors)%>">
							       <input type="submit" name="Submit" value="Cancel Reticle" /> <%=trim(rf_ref)%>
						        </td>				
					        <%
					     Case Else
					        If UsesLots=FALSE then
					        %>
                                <td class="MainPageTextSmaller" nowrap valign="top" bgcolor="<%=colorset(i mod numcolors)%>">
							        <%=trim(fh_custpo)%>
						        </td>
                            <!--
						        <td class="MainPageTextSmaller" nowrap valign="top" bgcolor="<%=colorset(i mod numcolors)%>">
							        <input type="submit" name="Submit" value="Cancel Lot" /> <%=trim(fh_custpo)%>
						        </td>
                                -->				
						        <%
						        'Response.Write "mark..."
						        else
						        'Response.Write "mark2..."
						        %>
						        <td class="MainPageTextSmaller" nowrap valign="top" bgcolor="<%=colorset(i mod numcolors)%>">
							        <input type="submit" name="Submit" value="Cancel Lot" /> <%=trim(rf_ref)%>
						        </td>				
					        <%End if
					     End Select					     
					     %>
					<!--				
					<td class="MainPageTextSmaller" valign="top">
						<%=fl_sf_name%>
					</td>	
					<td class="MainPageTextSmaller" valign="top">
						<%=fl_st_name%>
					</td>
					-->
					<td class="MainPageTextSmaller" nowrap valign="top" bgcolor="<%=colorset(i mod numcolors)%>">
						<%=fl_sf_clname%><br /><%=fl_sf_addr2 %>
						<%If (trim(fl_st_id)<>trim(fl_pod) AND trim(fl_pod)>"") then response.Write "<br><font color='red'>DISCREPANCY</font>" end if%>
					</td>
					<td class="MainPageTextSmaller" nowrap valign="top" bgcolor="<%=colorset(i mod numcolors)%>">
						<%=fl_st_clname%><br /><%=fl_st_addr2 %>
					</td>					
					<td class="MainPageTextSmaller" nowrap valign="top" bgcolor="<%=colorset(i mod numcolors)%>">
						<%=Display_fh_status%>
					</td>
					<%If UsesLots=TRUE then%>
					<td class="MainPageTextSmaller" nowrap valign="top" bgcolor="<%=colorset(i mod numcolors)%>">
						<%=Displayfh_priority%>
					</td>					
					<%end if%>
					<td class="MainPageTextSmaller" nowrap valign="top" bgcolor="<%=colorset(i mod numcolors)%>">
						<%=fh_ship_dt%>
					</td>
	
                    
				    </form>
																						
				</tr>
<%
				i=i+1
                tempfh_id=fh_id
				'END IF
				ii=ii+1
				RSEVENTS2.movenext
				LOOP
				RSEVENTS2.close
			Set RSEVENTS2 = Nothing
			'Response.Write "i="&i&"<BR>"
					'Response.Write "BillToID="&BillToID&"<BR>"
					'Response.Write "fh_custpo="&fh_custpo&"<BR>"
					'Response.Write "rf_ref="&rf_ref&"<BR>"
					'Response.Write "fh_id="&fh_id&"<BR>"			
			If i>0 then
				If i>1 then
					PluralResults="s"
					else
					'Response.Write "BillToID="&BillToID&"<BR>"
					'Response.Write "fh_custpo="&fh_custpo&"<BR>"
					'Response.Write "rf_ref="&rf_ref&"<BR>"
					'Response.Write "fh_id="&fh_id&"<BR>"
					'Select Case BillToID
						'Case 26
							'Response.Write "Redirect to SR?<BR>"
							'Response.Redirect("../reporting/jobanalysis.asp?inputdocumentnumber="&fh_custpo)
						'Case 36
							'Response.Write "Redirect to Wafer?<BR>"
							'Response.Redirect("../reporting/OrderDetails.asp?inputlotnumber="&trim(rf_ref)&"&inputjobnumber="&fh_id)
						'Case 38
							'Response.Write "Redirect to Reticle?<BR>"
							'Response.Redirect("../reporting/jobanalysis.asp?inputlotnumber="&fh_custpo&"&inputjobnumber="&fh_id)
							'Response.Redirect("../reporting/OrderDetails.asp?inputjobnumber="&fh_id)																								
						'Case 48, 13, 14, 25
							'Response.Write "Redirect to KWE?<BR>"
							'Response.Redirect("../reporting/jobanalysis.asp?inputlotnumber="&rf_ref&"&inputjobnumber="&fh_id)
					'End Select
					
				End if
				Response.Write "<tr><td align='left' class='miniheader' colspan='"&ColspanNumber&"'>"&i&" Result"&PluralResults
				If (i=20 and NumberofDays>0) or i=300 then
					Response.Write " - The maximum page display is 300 results.  There may be more results, please narrow your search criteria."
				end if
				Response.Write "</td></tr>"

			End if
			'Response.Write "ColspanNumber="&ColspanNumber&"<BR>"
	%>













			<tr><td align="center" class="FleetXBoldText" colspan="<%=ColspanNumber%>"><%=ErrorMessage%>&nbsp;</td></tr>
			</table>	
			<tr><td>&nbsp;</td></tr>
			<tr><td>&nbsp;</td></tr>
			<tr><td>&nbsp;</td></tr>			
	</table>
<%end if%>   
    
 
 
 
 
 
 
 
 
 
    
    
    </td></tr>



 
</table>
<!-- </form>  -->
</td></tr>
<tr><td height="100%">&nbsp;</td></tr>

<tr>
    <td height="100" class="FleetXGreySection" colspan="2">
        <!-- #include file="../BottomSection.asp" -->
    </td>
</tr>
<tr><td height="15" class="FleetXRedSectionSmall" colspan="2" align="center"><%=CopywriteNotice %></td></tr>
</td></tr></table>


</body>
</html>
