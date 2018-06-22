<html>
<head>

<!-- #include file="../fleetexpress.inc" -->
<link rel="stylesheet" type="text/css" href="../css/Style.css">
<%

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
    PageTitle="ACCESSORIAL CHARGES"

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
SQL = "SELECT a.caID, a.ca_fh_id, a.ca_bt_id, a.ca_atid, a.ca_accid, a.ca_accCharge, a.dateadded, a.addedby, a.ca_Type, a.LocationCode, b.bt_id, b.bt_desc "_
& " FROM ChargedAccessorials a "_
& " INNER JOIN fcbillto b on b.bt_id = a.ca_bt_id "

startdate = valid8(request.form("StartDate"))
startdate = trim(startdate)
enddate = valid8(request.form("EndDate"))
enddate = trim(enddate)

'response.write "91 start=" & startdate & ",end=" & enddate & "<br>"

if len(startdate) > 0 and len(enddate) > 0 then
  SQL = SQL & "WHERE a.dateadded >= '" & startdate & "' and a.dateadded < '" & dateadd("d",1,enddate) & "'"
end if
SQL = SQL & " order by a.ca_fh_id, a.dateadded"

'response.write "98 SQL = " & SQL & "<br>"

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

<form action="AccessorialCharges.asp" method="post" name="DatePick">
Date Range:&nbsp;&nbsp; FROM: <input type="text" name="StartDate" id="date_1" value="<%=startdate%>" /> &nbsp;&nbsp;TO: <input type="text" name="EndDate" id="date_2" value="<%=enddate%>"/> &nbsp;<INPUT id="gobutton" name="gobutton" TYPE="submit" name="ButtonValue" VALUE="GO">
</form>
</td></tr>
<tr><td  nowrap="nowrap" ><b>Job</b></td><td><b>Customer</b></td><td><b>Type</b></td><td  nowrap="nowrap" ><b>Date Added</b></td><td  nowrap="nowrap" ><b>Charge</b></td><td  nowrap="nowrap" ><b>ChangedBy</b></td></tr>
<%
	if NOT Recordset1.EOF then
    Do Until Recordset1.EOF
    
    accTypeID=Recordset1("ca_atid")
      SQL = "SELECT * FROM AccessorialType WHERE atid = '" & accTypeID & "'"
      SET oRsN = oConn.Execute(SQL)
      if NOT oRsN.EOF then
        accDescr = oRsN("atDescr")
        BillCode = oRsN("atBillCode")
      else
        accDescr = "UNKNOWN"
      end if
      set oRsN = Nothing
          
    accCharge = Recordset1("ca_accCharge")
    accCharge = FormatCurrency(accCharge,2)
    
  ChangedBy = Recordset1("addedby")
  'response.write "117 addedby=" & ChangedBy & "<br>"
  if isNumeric(ChangedBy) then
  
  Set Recordset11 = Server.CreateObject("ADODB.Recordset")
'Response.Write "Intranet="&Intranet&"***<BR>"
SQL777="SELECT * FROM INTRANET_USERS WHERE (USERID='"&ChangedBy&"')"
'Response.Write "SQL777="&SQL777&"***<BR>"
Recordset11.ActiveConnection = Intranet
Recordset11.Source = SQL777
Recordset11.CursorType = 0
Recordset11.CursorLocation = 2
Recordset11.LockType = 1
Recordset11.Open()
Recordset11_numRows = 0
	if NOT Recordset11.EOF then
		FirstName=Recordset11("FirstName")
		LastName=Recordset11("LastName")
    ChangedBy = FirstName & " " & LastName & "(" & ChangedBy & ") "
  else
    ChangedBy = ""
	end if
  	Recordset11.Close()
		Set Recordset11 = Nothing
  end if
  
  TheDateTime=Recordset1("dateadded")
  TheDayOnly=Day(TheDateTime)
  TheMonthOnly=Month(TheDateTime)
  TheYearOnly=Year(TheDateTime)
  DateOnly=TheMonthOnly&"/"&TheDayOnly&"/"&TheYearOnly
    %>
        <tr><td  nowrap="nowrap" ><%=Recordset1("ca_fh_id")%></td><td><%=Recordset1("ca_bt_ID")%>-<%=Recordset1("bt_desc")%></td><td  nowrap="nowrap" ><%=accDescr%></td><td  nowrap="nowrap" ><%=Recordset1("dateadded")%></td><td  nowrap="nowrap" ><%=trim(accCharge)%></td><td  nowrap="nowrap" ><%=ChangedBy%></td></tr>
    <%
    Recordset1.MoveNext
    Loop
	Else
      %><tr><td  nowrap="nowrap"  colspan=6>NO ACCESSORIALS FOUND</td></tr><%
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

