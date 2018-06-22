<html>
<head>

<!-- #include file="../fleetexpress.inc" -->
<link rel="stylesheet" type="text/css" href="../css/Style.css">
<%
''''''''''''''THIS page scrolls through list of Carillon numbers and associates them with as many Addresses as possible
    Whatever=Request.querystring("Whatever")
    startdate = valid8(request.form("StartDate"))
    startdate = trim(startdate)
    enddate = valid8(request.form("EndDate"))
    enddate = trim(enddate)

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
	<!--table align="center" border="0" bordercolor="black" cellpadding="0" cellspacing="0" ID="Table1" height="100%" width="100%">
        <tr><td  nowrap="nowrap"  align="left"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>
        <tr>
            <td  nowrap="nowrap"  align="left"><img src="../images/FleetX_Small.jpg" height="50" width="168" /></td>
            <td  nowrap="nowrap"  align="right" valign="bottom"--><!-- #include file="../topnavbar.asp" --><!--/td>
        </tr>
        <tr><td  nowrap="nowrap"  align="left" colspan="2"><img src="../images/pixel.gif" /></td></tr>	
    <tr><td  nowrap="nowrap"  colspan="2">
    
<table border="0" bordercolor="green" Cellspacing="0" Cellpadding="0" align="left" width="100%">
 <tr><td  nowrap="nowrap"  align="left" colspan="2" bgcolor="<%=HeaderBorderColor%>"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>

        <tr><td  nowrap="nowrap"  align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetXRedSection" colspan="2"><%=PageTitle%></td></tr>

        <tr><td  nowrap="nowrap"  align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
        <!--tr><td  nowrap="nowrap"  align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressBodyWhite" colspan="2">In order to reach the Fleet Express Order page, please correctly type in the green verification code in the supplied text box and click "Submit."</td></tr>
        <tr><td  nowrap="nowrap"  align="left" colspan="2"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
<tr><td  nowrap="nowrap" >
<table  border="0" bordercolor="blue" align="center" class="MainPageText" width="100%">
	<tr height="40">
		<td  nowrap="nowrap">&nbsp;</td>
	</tr>



    <tr><td  nowrap="nowrap"  align="center"--><!-- main page stuff goes here! -->
    
    
 <%
 
  Set oConn = Server.CreateObject("ADODB.Connection")
oConn.ConnectionTimeout = 100
oConn.Provider = "MSDASQL"
oConn.Open DATABASE
 
''''''''DISPLAY ACCESSORIALS
Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Response.write "Database="&Database&"<br>"
Recordset1.ActiveConnection = Database
SQL = "SELECT CarillonID, ST_ID, CompanyAddress FROM PreExistingCompanies WHERE (CompanyStatus = 'c') AND (CarillonID < 1 OR CarillonID IS NULL) and companyID>0"
Recordset1.Source = SQL
response.write "SQL="& SQL &"<BR>"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()
Recordset1_numRows = 0
'response.write "<font color='red'>SQL1="&Recordset1.Source&"<BR></font>"
%>
<!--table align="center" border="1" bordercolor="red" cellpadding="5" width=900>
<tr><td><b>CarillonID</b></td><td  nowrap="nowrap" ><b>st_id</b></td><td  nowrap="nowrap" ><b>Address</b></td></tr-->
<%
	if NOT Recordset1.EOF then
    Do Until Recordset1.EOF
    
        CarillonID=Recordset1("CarillonID")
        ST_ID=Recordset1("ST_ID")
        CompanyAddress=Recordset1("CompanyAddress")
        xxx=xxx+1
    %>
        <!--tr><td  nowrap="nowrap" ><%=CarillonID%></td><td><%=ST_ID%></td><td nowrap="nowrap"><%=CompanyAddress%></td></tr-->
    <%

Set Recordset16 = Server.CreateObject("ADODB.Recordset")
'Response.write "Database="&Database&"<br>"
Recordset16.ActiveConnection = Database
SQL = "SELECT CarillonID from CarillonIDs where BillingStatus='c' and CompanyAddress='"&trim(CompanyAddress)&"'"

Recordset16.Source = SQL
'response.write "SQL="& SQL &"<BR>"
Recordset16.CursorType = 0
Recordset16.CursorLocation = 2
Recordset16.LockType = 1
Recordset16.Open()
Recordset16_numRows = 0
If not Recordset16.eof then

TempCarillonID=Recordset16("CarillonID")


    If whatever="whatever" then
						Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open DATABASE
						''''UPDATES THE WAFER
						l_cSQL = "UPDATE PreExistingCompanies SET CarillonID = '"&TempCarillonID&"' WHERE (st_id = '"&st_id&"') AND (companystatus = 'c')"
						oConn.Execute(l_cSQL)
                        oConn.close
                        Set oConn=Nothing
    End if


End if

Recordset16.Close()
Set Recordset16 = Nothing 
%>
<!--tr><td  colspan="3"><%=TempCarillonID%></td></tr-->
<%





    Recordset1.MoveNext
    Loop
	Else
      %><!--tr><td  nowrap="nowrap"  colspan=6>NO PREVIOUS BILLINGS FOUND</td></tr--><%
  End if
Recordset1.Close()
Set Recordset1 = Nothing 
Response.write "xxx="&xxx&"<BR>"   
 %>   

<!--tr><td  nowrap="nowrap"  colspan=6>&nbsp;<br><br><br><br><a href="../home.asp" class="FleetXRedMain">CLICK HERE</a> to Return to the Home Page</a></td></tr>
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
    <td  nowrap="nowrap"  height="100" class="FleetXGreySection" colspan="2"-->
        <!-- #include file="../BottomSection.asp" -->
    <!--/td>
</tr>
<tr><td  nowrap="nowrap"  height="15" class="FleetXRedSectionSmall" colspan="2" align="center"><%=CopywriteNotice %></td></tr>
</td></tr></table-->


</body>
</html>

