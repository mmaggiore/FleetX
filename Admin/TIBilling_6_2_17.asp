<html>
<head>

<!-- #include file="../fleetexpress.inc" -->
<link rel="stylesheet" type="text/css" href="../css/Style.css">
<%
'''''''Shows all previous billings and allows you to export to XLS for billings
    startdate = valid8(request.form("StartDate"))
    startdate = trim(startdate)
    enddate = valid8(request.form("EndDate"))
    enddate = trim(enddate)
    OriginalEndDate=EndDate
    GoButton = valid8(request.form("GoButton"))
    'Response.write "EndDate="&EndDate&"<BR>"

If GoButton="GO" then
    BilledDate=Now()
    'Response.write "got here 14<BR>"
    If trim(startdate)="" then
        ErrorMessage="You must provide a starting date."
    End if
    If trim(enddate)="" then
        ErrorMessage="You must provide an ending date."
    End if
    If Trim(ErrorMessage)="" then
        If isdate(EndDate) then
            EndDate=DateAdd("d", 1, EndDate)
        End if
				Set oConn = Server.CreateObject("ADODB.Connection")
				oConn.ConnectionTimeout = 100
				oConn.Provider = "MSDASQL"
				oConn.Open DATABASE
					l_cSQL2 = "select fl_fh_id FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN PreExistingCompanies ON fclegs.fl_sf_id = PreExistingCompanies.ST_ID INNER JOIN PreExistingCompanies PreExistingCompanies_1 ON fclegs.fl_st_id = PreExistingCompanies_1.ST_ID WHERE fl_t_atd>'"&StartDate&"' and fl_t_atd<'"&EndDate&"' and fh_status='CLS' and (fh_billed_date IS NULL) and (PreExistingCompanies_1.CarillonID > 10) AND (PreExistingCompanies.CarillonID > 10)    " 
							'if trim(displayusername)="comps" or trim(displayusername)="Compugraphics"  then 
							'l_cSQL2 = l_cSQL2 & "  AND st_id<>'CPGP'" 
							'end if
					'Response.Write "l_cSQL2="&l_cSQL2&"<BR>"
					SET oRs = oConn.Execute(l_cSql2)

					Do While not oRs.EOF
                    fl_id=oRs("fl_fh_id")
                    QQQ=QQQ+1
                    'CarillonID=oRs("CarillonID")
        
                                    'Set oConn2 = Server.CreateObject("ADODB.Connection")
	                                'oConn2.ConnectionTimeout = 100
	                                'oConn2.Provider = "MSDASQL"
	                               ' oConn2.Open DATABASE
	                                ' 7/26/04 KK: Added canceljob functionality to be able to update the status if cancel button is pressed.
	                                ' 11/30/04 DEC: Changed from CAN/98 TO DEL/99 to be consistent with dispatchOffice			
		                            '    l_cSQL = "UPDATE fcfgthd SET fh_billed_date = '"&BilledDate&"' WHERE fh_id = '" & fl_id & "' and fh_status='CLS'"
		                            '    Response.write "l_cSQL="&l_cSQL&"<BR>"
		                            '    oConn2.Execute(l_cSQL)
	                                'Set oConn2=Nothing
					oRs.movenext
					LOOP
				Set oConn=Nothing

    					Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open DATABASE
						''''UPDATES THE WAFER
						l_cSQL = "UPDATE fcfgthd SET fh_billed_date = '"&BilledDate&"' FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN PreExistingCompanies ON fclegs.fl_sf_id = PreExistingCompanies.ST_ID INNER JOIN PreExistingCompanies PreExistingCompanies_1 ON fclegs.fl_st_id = PreExistingCompanies_1.ST_ID WHERE fl_t_atd>'"&StartDate&"' and fl_t_atd<'"&EndDate&"' and fh_status='CLS' and (fh_billed_date IS NULL) and (PreExistingCompanies_1.CarillonID > 10) AND (PreExistingCompanies.CarillonID > 10)  " 
						oConn.Execute(l_cSQL)
                        oConn.close
                        Set oConn=Nothing


If QQQ<1 then QQQ=0
	Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
		RSEVENTS2.Open "PreviousTIBillings", DATABASE, 2, 2
		RSEVENTS2.addnew
        RSEVENTS2("BillingDate")=BilledDate
        RSEVENTS2("BillingStartDate")=startdate
        RSEVENTS2("BillingEndDate")=originalenddate
        RSEVENTS2("NumberOfJobs")=QQQ
        RSEVENTS2("UserID")=UserID
        RSEVENTS2("BillingStatus")="c"
		RSEVENTS2.update
		RSEVENTS2.close			
	set RSEVENTS2 = nothing 


    SuccessMessage="You have successfully created a new billing."
    End if
End if





    ColorSelect=valid8(Request.form("ColorSelect"))
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
    PageTitle="TI BILLING"

%>
<title>FleetX - <%=PageTitle %></title>

<link rel="stylesheet" href="../themes/classic.css" id="theme_base">
<link rel="stylesheet" href="../themes/classic.date.css" id="theme_date">
<link rel="stylesheet" href="../themes/classic.time.css" id="theme_time">

</head>

<body>
	<table align="center" border="0" bordercolor="black" cellpadding="0" cellspacing="0" ID="Table1" height="100%" width="100%">
        <tr><td  nowrap="nowrap"  align="left"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>
        <tr>
            <td  nowrap="nowrap"  align="left"><img src="../images/FleetX_Small.jpg" height="50" width="168" /></td>
            <td  nowrap="nowrap"  align="right" valign="bottom"><!-- #include file="../topnavbar.asp" --></td>
        </tr>
        <tr><td  nowrap="nowrap"  align="left" colspan="2"><img src="../images/pixel.gif" /></td></tr>	
    <tr><td  nowrap="nowrap"  colspan="2">
    
<table border="0" bordercolor="green" Cellspacing="0" Cellpadding="0" align="left" width="100%">
 <tr><td  nowrap="nowrap"  align="left" colspan="2" bgcolor="<%=HeaderBorderColor%>"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>

        <tr><td  nowrap="nowrap"  align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetXRedSection" colspan="2"><%=PageTitle%></td></tr>

        <tr><td  nowrap="nowrap"  align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
        <!--tr><td  nowrap="nowrap"  align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressBodyWhite" colspan="2">In order to reach the Fleet Express Order page, please correctly type in the green verification code in the supplied text box and click "Submit."</td></tr-->
        <tr><td  nowrap="nowrap"  align="left" colspan="2"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
<tr><td  nowrap="nowrap" >
<table  border="0" bordercolor="blue" align="center" class="MainPageText" width="100%">
	<tr height="40">
		<td  nowrap="nowrap">&nbsp;</td>
	</tr>



    <tr><td  nowrap="nowrap"  align="center"><!-- main page stuff goes here! -->
    
    
 <%
 
  Set oConn = Server.CreateObject("ADODB.Connection")
oConn.ConnectionTimeout = 100
oConn.Provider = "MSDASQL"
oConn.Open DATABASE
 
''''''''DISPLAY ACCESSORIALS
Set Recordset1 = Server.CreateObject("ADODB.Recordset")
'Response.write "Database="&Database&"<br>"
Recordset1.ActiveConnection = Database
SQL = "SELECT PreviousTIBillings.BillingDate, PreviousTIBillings.BillingStartDate, PreviousTIBillings.BillingEndDate, " 
SQL = SQL&"PreviousTIBillings.NumberOfJobs, PreviousTIBillings.UserID, PreviousTIBillings.BillingStatus, PreExistingRequestor.RequestorName "
SQL = SQL&"FROM PreviousTIBillings INNER JOIN PreExistingRequestor ON PreviousTIBillings.UserID = PreExistingRequestor.RequestorID Where BillingStatus='c'"

Recordset1.Source = SQL
'response.write "SQL="& SQL &"<BR>"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()
Recordset1_numRows = 0
'response.write "<font color='red'>SQL1="&Recordset1.Source&"<BR></font>"
%>
<table align="center" border="0" bordercolor="red" cellpadding="5" width=900>
<tr><td colspan=8>

<form action="TIBilling.asp" method="post" name="DatePick">
Select Date Range to Create a Bill:&nbsp;&nbsp; FROM: <input type="text" name="StartDate" id="date_1" value="<%=startdate%>" /> &nbsp;&nbsp;TO: <input type="text" name="EndDate" id="date_2" value="<%=OriginalEndDate%>"/> &nbsp;<INPUT id="gobutton" name="gobutton" TYPE="submit" name="ButtonValue" VALUE="GO">
</form>
</td></tr>
<%if ErrorMessage>"" then%>
<tr><td  nowrap="nowrap" colspan="6">
<table width="100%" border="0" bordercolor="Yellow" cellspacing="0" cellpadding="0" align="center" class="MainPageText">
	<tr><td  nowrap="nowrap" >&nbsp;</td></tr>  
	<tr>
    <td  nowrap="nowrap"  align="center" class="Errormessage"><%=ErrorMessage%></td>
  </tr>
	<tr><td  nowrap="nowrap" >&nbsp;</td></tr>
</table>
</td></tr>
<%end if%>
<%if SuccessMessage>"" then%>
<tr><td  nowrap="nowrap" colspan="6">
<table width="100%" border="0" bordercolor="Yellow" cellspacing="0" cellpadding="0" align="center" class="MainPageText">
	<tr><td  nowrap="nowrap" >&nbsp;</td></tr>  
	<tr>
    <td  nowrap="nowrap"  align="center" class="Errormessage"><font color="blue"><%=SuccessMessage%></font></td>
  </tr>
	<tr><td  nowrap="nowrap" >&nbsp;</td></tr>
</table>
</td></tr>
<%end if%>
<tr><td>&nbsp;</td><td  nowrap="nowrap" ><b>Billing Date</b></td><td><b>From</b></td><td><b>To</b></td><td  nowrap="nowrap" ><b>Number of Jobs</b></td><td  nowrap="nowrap" ><b>Billed by</b></td></tr>
<%
	if NOT Recordset1.EOF then
    Do Until Recordset1.EOF
    
        BillingDate=Recordset1("BillingDate")
        BillingStartDate=Recordset1("BillingStartDate")
        BillingEndDate=Recordset1("BillingEndDate")
        NumberOfJobs=Recordset1("NumberOfJobs")
        RequestorName=Recordset1("RequestorName")
  

    %>
    <form method="post" action="CreateBillingSpreadsheet.asp" target="_blank">
        <tr><td><INPUT id="gobutton" name="gobutton" TYPE="submit" name="ButtonValue" VALUE="EXPORT"></td><td  nowrap="nowrap" ><%=BillingDate%></td><td><%=BillingStartDate%></td><td nowrap="nowrap"><%=BillingEndDate%></td><td  nowrap="nowrap" ><%=NumberOfJobs%></td><td  nowrap="nowrap" ><%=RequestorName%></td></tr>
        <input type="hidden" name="BilledDate" value="<%=BillingDate %>" />
        <input type="hidden" name="xyz" value="72" />
    </form>
    <%
    Recordset1.MoveNext
    Loop
	Else
      %><tr><td  nowrap="nowrap"  colspan=6>NO PREVIOUS BILLINGS FOUND</td></tr><%
  End if
Recordset1.Close()
Set Recordset1 = Nothing    
 %>   

<tr><td  nowrap="nowrap"  colspan=6>&nbsp;<br><br><br><br><a href="../home.asp" class="FleetXRedMain">CLICK HERE</a> to Return to the Home Page</a></td></tr>
</table>
 
 
   
    </td></tr>



 
	<tr Height="50">
		<td  nowrap="nowrap" >&nbsp;</td>
	</tr>


</table>
</td></tr>
<%
if ErrorMessage>"" then%>
<tr><td  nowrap="nowrap" >
<table width="100%" border="0" bordercolor="Yellow" cellspacing="0" cellpadding="0" align="center" class="MainPageText">
	<tr><td  nowrap="nowrap" >&nbsp;</td></tr>  
	<tr>
    <td  nowrap="nowrap"  align="center" class="Errormessage"><%=ErrorMessage%></td>
  </tr>
	<tr><td  nowrap="nowrap" >&nbsp;</td></tr>
</table>
</td></tr>
<%end if%>
</table>
<tr><td  nowrap="nowrap"  Height="90%">&nbsp;</td></tr>
<tr>
    <td  nowrap="nowrap"  height="100" class="FleetXGreySection" colspan="2">
        <!-- #include file="../BottomSection.asp" -->
    </td>
</tr>
<tr><td  nowrap="nowrap"  height="15" class="FleetXRedSectionSmall" colspan="2" align="center"><%=CopywriteNotice %></td></tr>
</td></tr></table>

<script src="../jquery-2.1.0.min.js"></script> 
<script src="../pickadate.js"></script> 
<script type="text/javascript">
    // PICKADATE FORMATTING
    $('#date_1').pickadate({
        format: 'mm/dd/yyyy', 	// Friendly format displayed to user
        formatSubmit: 'mm/dd/yyyy', // Actual format used by application
        hiddenName: false			// Allows two different formats
    });

    $('#date_2').pickadate({
        format: 'mm/dd/yyyy', 	// Friendly format displayed to user
        formatSubmit: 'mm/dd/yyyy', // Actual format used by application
        hiddenName: false			// Allows two different formats
    });


    $('#time_1').pickatime({
        format: 'h:i A', 		// Displayed and application format
        interval: 10, 			// Interval between values (in minutes)
        min: '12:00 AM', 			// Starting value
        max: '11:59 PM'				// Ending value
    });

</script>

</body>
</html>

