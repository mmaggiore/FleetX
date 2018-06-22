<html>
<head>

<!-- #include file="../fleetexpress.inc" -->
<link rel="stylesheet" type="text/css" href="../css/Style.css">
<%
    'Response.buffer=false
    'Response.write "UserID="&UserID&"<BR>"
    CompanyAddress = (valid8(request.form("CompanyAddress")))
    CarillonID = Trim(valid8(request.form("CarillonID")))
If not IsNumeric(CarillonID) then
    'Response.write "Look!  I got here!<BR>"
    ErrorMessage="The Carillon ID must contain only be numbers."
    CompanyAddress=""
    CarillonID=""
End if
If CarillonID>"" then
    					Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open DATABASE
						''''UPDATES THE WAFER
						l_cSQL = "UPDATE PreExistingCompanies SET CarillonID = '"&CarillonID&"', WhoSetCarillonID='"&UserID&"' WHERE (CompanyAddress = '"&CompanyAddress&"') AND (companystatus = 'c') And (CarillonID<1 or CarillonID is NULL)"
						oConn.Execute(l_cSQL)
                        oConn.close
                        Set oConn=Nothing
                        Response.write "l_cSQL="&l_cSQL&"<BR>"
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
    PageTitle="UPDATE CARILLON INFO"

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

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
'Response.write "Database="&Database&"<br>"
Recordset1.ActiveConnection = Database
SQL = "SELECT Count(ST_ID) As NumberOfJobs FROM PreExistingCompanies WHERE (CompanyStatus = 'c') AND (CarillonID < 1 OR CarillonID IS NULL) and companyID>0"
Recordset1.Source = SQL
'response.write "SQL="& SQL &"<BR>"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()
Recordset1_numRows = 0
	if NOT Recordset1.EOF then
 
        NumberOfJobs=Recordset1("NumberOfJobs")
    Recordset1.MoveNext
  End if
Recordset1.Close()
Set Recordset1 = Nothing  

'Response.write "NumberOfJobs="&NumberOfJobs&"<BR>"






 
''''''''DISPLAY ACCESSORIALS
Set Recordset1 = Server.CreateObject("ADODB.Recordset")
'Response.write "Database="&Database&"<br>"
Recordset1.ActiveConnection = Database
SQL = "SELECT Top 20 CarillonID, ST_ID, CompanyAddress, CompanyBuilding, CompanySuite FROM PreExistingCompanies WHERE (CompanyStatus = 'c') AND (CarillonID < 1 OR CarillonID IS NULL) and companyID>0 Order by CompanyAddress"
Recordset1.Source = SQL
'response.write "SQL="& SQL &"<BR>"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()
Recordset1_numRows = 0
'response.write "<font color='red'>SQL1="&Recordset1.Source&"<BR></font>"
%>
<table align="center" border="0" bordercolor="red" cellpadding="0" cellspacing="0" >
<%If trim(errormessage)>"" then %>
    <tr><td>&nbsp;</td></tr>
    <tr><td  nowrap="nowrap" colspan="9" align="center" class="errormessage"><b><%=ErrorMessage %></b></td></tr>
    <tr><td>&nbsp;</td></tr>
<%End  if%>
<tr><td  nowrap="nowrap" colspan="9" align="center"><b>Number of Jobs Missing Carillon IDs: <%=NumberOfJobs %></b></td></tr>
<tr><td>&nbsp;</td></tr>
<tr><td  nowrap="nowrap" ><b>&nbsp;</b></td><td><img src="../images/pixel.gif" height="1" width="5" /></td><td><b>Carillon ID</b></td><td><img src="../images/pixel.gif" height="1" width="5" /></td><td><b>Address</b></td><td><img src="../images/pixel.gif" height="1" width="5" /></td><td><b>Building</b></td><td><img src="../images/pixel.gif" height="1" width="5" /></td><td><b>Suite</b></td></tr>
<%
	if NOT Recordset1.EOF then
    Do Until Recordset1.EOF
    
        CompanyAddress=Recordset1("CompanyAddress")
        CompanyBuilding=Recordset1("CompanyBuilding")
        CompanySuite=Recordset1("CompanySuite")
  

    %>
    <form method="post" action="UpdateCarillonInfo.asp">
        <tr><td><INPUT id="gobutton" name="gobutton" TYPE="submit" name="ButtonValue" VALUE="UPDATE"></td><td><img src="../images/pixel.gif" height="1" width="5" /></td><td  nowrap="nowrap" ><input type="text" name="CarillonID" maxlength="10" /></td><td><img src="../images/pixel.gif" height="1" width="5" /></td><td><%=CompanyAddress%></td><td><img src="../images/pixel.gif" height="1" width="5" /></td><td  nowrap="nowrap" ><%=CompanyBuilding %></td><td><img src="../images/pixel.gif" height="1" width="5" /></td><td  nowrap="nowrap" ><%=CompanySuite %></td></tr>
        <input type="hidden" name="CompanyAddress" value="<%=CompanyAddress %>" />
    </form>
    <tr height="3"><td colspan="5"><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
    <tr height="1"><td colspan="9" bgcolor="black"><img src="../images/pixel.gif" height="1" width="1" /></td></tr>
    <tr height="3"><td colspan="5"><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
    <%
    Recordset1.MoveNext
    Loop
	Else
      %><tr><td  nowrap="nowrap"  colspan=6>NO ADDRESSES WITHOUT CARILLON IDS EXIST</td></tr><%
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

