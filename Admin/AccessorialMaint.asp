<html>
<head>

<!-- #include file="../fleetexpress.inc" -->
<link rel="stylesheet" type="text/css" href="../css/Style.css">
<%
Add = valid8(trim(request.form("add")))
if Add = "Y" then
      Set oConn = Server.CreateObject("ADODB.Connection")
      oConn.ConnectionTimeout = 100
      oConn.Provider = "MSDASQL"
      oConn.Open DATABASE
      btCust = valid8(trim(request.form("btCust")))
      StartDate = valid8(trim(request.form("StartDate")))
      EndDate = valid8(trim(request.form("EndDate")))
      accTypeID = valid8(trim(request.form("atid")))
      accCharge = valid8(trim(request.form("accCharge")))
       'check to be sure the accType isn't already set for this company:
        SQL = "SELECT * from Accessorials WHERE bt_id = " & btCust & " AND atid = " & accTypeID
        SET oRs = oConn.Execute(SQL)
        if NOT oRs.EOF then
          errMessage = "Accessorial Type already exists for this Company - please try again<br><br>"
        else
        'insert new one
        SQL = "INSERT INTO Accessorials (bt_id,atid,accCharge,accdate,changedby,accStatus) values(" & btCust & "," & accTypeID & ",'" &accCharge&"','"&Now()&"','"&Request.cookies("FleetXCookie")("UserID")&"','c')"
          'SQL="INSERT INTO Accessorials values(" & btCust & "," & accTypeID & ",'" & StartDate & "','" & EndDate & "','','"&accCharge&"','"&Now()&"','"&Request.cookies("FleetXCookie")("UserID")&"','c')"
          'response.write "20 sql=" & SQL & "<br>"
          SET oRs = oConn.Execute(SQL)
        end if
      Set oRS=Nothing
      Set oConn=Nothing
end if

delID = valid8(trim(request.querystring("delID")))
if isNumeric(delID) and delID > 0 then
      Set oConn = Server.CreateObject("ADODB.Connection")
      oConn.ConnectionTimeout = 100
      oConn.Provider = "MSDASQL"
      oConn.Open DATABASE
        'delete
        SQL="UPDATE Accessorials SET accStatus = 'x' WHERE accID=" & delID
        SET oRs = oConn.Execute(SQL)
      Set oRS=Nothing
      Set oConn=Nothing
end if

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
    PageTitle="ACCESSORIALS MAINTENANCE"

%>
<title>FleetX - <%=PageTitle %></title>
<script language="javascript" type="text/javascript" src="datetimepicker.js">
    //Date Time Picker script- by TengYong Ng of http://www.rainforestnet.com
</script>
</head>

<body onload="document.FindUser.requestorName.focus();">
	<table align="center" border="0" bordercolor="black" cellpadding="0" cellspacing="0" ID="Table1" height="100%" width="100%">
        <tr><td  nowrap="nowrap"  align="left"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>
        <tr>
            <td  nowrap="nowrap"  align="left"><img src="../images/FleetX_Small.jpg" height="50" width="168" /></td>
            <td  nowrap="nowrap"  align="right" valign="bottom"><!-- #include file="../topnavbar.asp" --></td>
        </tr>
        <tr><td  nowrap="nowrap"  align="left" colspan="2"><img src="../images/pixel.gif" /></td></tr>	
    <tr><td  nowrap="nowrap"  colspan="2">
<!-- <form action="NewUser.asp" method="post" name="FindUser">   -->
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
if len(errMessage) then
  response.write "<font color='red'><b>" & errMessage & "</b></font>"
  errMessage = ""
end if

Set oConn = Server.CreateObject("ADODB.Connection")
oConn.ConnectionTimeout = 100
oConn.Provider = "MSDASQL"
oConn.Open DATABASE
 
''''''''DISPLAY ACCESSORIALS
Set Recordset1 = Server.CreateObject("ADODB.Recordset")
'Response.write "Database="&Database&"<br>"
Recordset1.ActiveConnection = Database
SQL = "SELECT a.accID, a.bt_id, a.atid, a.accCharge, a.accStatus, a.accDate, a.changedby, a.accstartdate, a.accstopdate, b.bt_id, b.bt_desc "_
& " FROM Accessorials a "_
& " INNER JOIN fcbillto b on b.bt_id = a.bt_id "_
& " WHERE (a.accStatus='c') order by b.bt_desc, a.atid, a.accstartdate"

'SQL="SELECT * FROM Accessorials WHERE (RateStatus='c')"
Recordset1.Source = SQL
'response.write "SQL="& SQL &"<BR>"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()
Recordset1_numRows = 0
'response.write "<font color='red'>SQL1="&Recordset1.Source&"<BR></font>"
%>
<table align="center" border="0" bordercolor="red" cellpadding="3">
<!-- <tr><td  nowrap="nowrap" ><b>Customer</b></td><td  nowrap="nowrap" ><b>Type</b></td><td><b>Bill Code</b></td><td  nowrap="nowrap" ><b>StartDate</b></td><td  nowrap="nowrap" ><b>EndDate</b></td><td  nowrap="nowrap" ><b>Charge</b></td><td  nowrap="nowrap" ><b>Changed</b></td><td  nowrap="nowrap" ><b>ChangedBy</b></td><td  nowrap="nowrap" ><b>action</b></td></tr>   -->
<tr><td  nowrap="nowrap" ><b>Customer</b></td><td  nowrap="nowrap" ><b>Type</b></td><td><b>Bill Code</b></td><td  nowrap="nowrap" ><b>Charge</b></td><td  nowrap="nowrap" ><b>Added</b></td><td  nowrap="nowrap" ><b>Added By</b></td><td  nowrap="nowrap" ><b>action</b></td></tr>
<%
	if NOT Recordset1.EOF then
    Do Until Recordset1.EOF
    
    accTypeID=Recordset1("atid")
      SQL = "SELECT * FROM AccessorialType WHERE atid = '" & accTypeID & "'"
      SET oRsN = oConn.Execute(SQL)
      if NOT oRsN.EOF then
        accDescr = oRsN("atDescr")
        BillCode = oRsN("atBillCode")
      else
        accDescr = "UNKNOWN"
      end if
      set oRsN = Nothing
          
    accCharge = Recordset1("accCharge")
    accCharge = FormatCurrency(accCharge,2)
    
  ChangedBy = Recordset1("ChangedBy")
  if isNumeric(ChangedBy) then
    SQL="SELECT * FROM PreExistingRequestor WHERE RequestorID=" & ChangedBy
    SET oRsN = oConn.Execute(SQL)
    if NOT oRsN.EOF then
      ChangedBy = oRsN("RequestorName") & "(" & ChangedBy & ") "
    else
      ChangedBy = ""
    end if
  else
    ChangedBy = ""
  end if
  TheDateTime=Recordset1("accDate")
  TheDayOnly=Day(TheDateTime)
  TheMonthOnly=Month(TheDateTime)
  TheYearOnly=Year(TheDateTime)
  DateOnly=TheMonthOnly&"/"&TheDayOnly&"/"&TheYearOnly
    %>
      <!--  <tr><td  nowrap="nowrap" ><%=Recordset1("bt_ID")%>-<%=Recordset1("bt_desc")%></td><td  nowrap="nowrap" ><%=accDescr%></td><td><%=BillCode%></td><td  nowrap="nowrap" ><%=Recordset1("accstartdate")%></td><td  nowrap="nowrap" ><%=Recordset1("accstopdate")%></td><td  nowrap="nowrap" ><%=trim(accCharge)%></td><td  nowrap="nowrap" ><%=DateOnly %></td><td  nowrap="nowrap" ><%=ChangedBy%></td><td  nowrap="nowrap" ><a href="AccessorialEdit.asp?id=<%=Recordset1("accID")%>" class="FleetXRedMain">edit</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="AccessorialMaint.asp?delID=<%=Recordset1("accID")%>" class="FleetXRedMain">remove</a></td></tr>    -->
        <tr><td  nowrap="nowrap" ><%=Recordset1("bt_ID")%>-<%=Recordset1("bt_desc")%></td><td  nowrap="nowrap" ><%=accDescr%></td><td><%=BillCode%></td><td  nowrap="nowrap" ><%=trim(accCharge)%></td><td  nowrap="nowrap" ><%=DateOnly %></td><td  nowrap="nowrap" ><%=ChangedBy%></td><td  nowrap="nowrap" ><a href="AccessorialEdit.asp?id=<%=Recordset1("accID")%>" class="FleetXRedMain">edit</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="AccessorialMaint.asp?delID=<%=Recordset1("accID")%>" class="FleetXRedMain">remove</a></td></tr>
    <%
    Recordset1.MoveNext
    Loop
	Else
      %><tr><td  nowrap="nowrap"  colspan=6>NO ACCESSORIALS FOUND</td></tr><%
  End if
Recordset1.Close()
Set Recordset1 = Nothing    
 %>   
 <tr><td  nowrap="nowrap"  colspan=6>&nbsp;<br><b>ADD NEW ACCESSORIAL:</b></td></tr>
<form action="AccessorialMaint.asp" method="post" name="NewAcc">
<tr><td  nowrap="nowrap" >
   <select name="btCust">
  <%
  SQL = "SELECT * FROM fcbillto WHERE bt_status = 'c'"
  SET oRsN = oConn.Execute(SQL)
  Do Until oRsN.EOF
        %><option value=<%=oRsN("bt_id")%>><%=oRsN("bt_desc")%></option><%
  oRsN.MoveNext
  Loop
  %>
  </select>
  </td><td  nowrap="nowrap" >
  <select name="atid">
  <%
  SQL = "SELECT * FROM AccessorialType ORDER BY atDescr"
  SET oRsN = oConn.Execute(SQL)
  Do Until oRsN.EOF
        %><option value=<%=oRsN("atid")%>><%=oRsN("atDescr")%></option><%
  oRsN.MoveNext
  Loop
  %>
  </select>
  </td>
    <td></td>

<!-- <td  nowrap="nowrap"  nowrap><input type="text" id="StartDate" name="StartDate">
  <a href="javascript:NewCal('StartDate','MMddyyyy',true,12)"><img src="../images/cal.gif" width="16" height="16" border="0" alt="Pick a date"></a>               
  </td>
  <td  nowrap="nowrap"  nowrap><input type="text" id="EndDate" name="EndDate">
  <a href="javascript:NewCal('EndDate','MMddyyyy',true,12)"><img src="../images/cal.gif" width="16" height="16" border="0" alt="Pick a date"></a>               
  </td> -->
  <td  nowrap="nowrap" ><input type="text" name="accCharge"></td>
  <input type="hidden" name="add" value="Y">
   <td  nowrap="nowrap"  colspan=3><INPUT id="gobutton" name="gobutton" TYPE="submit" name="ButtonValue" VALUE="ADD"></td>
 </tr>
 </form>
<tr><td  nowrap="nowrap"  colspan=6>&nbsp;<br><br><a href="AccessorialTypes.asp" class="FleetXRedMain"><b>add/edit accessorial types</b></a></td></tr>
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
<!-- </form>  -->
<tr><td  nowrap="nowrap"  Height="90%">&nbsp;</td></tr>
<tr>
    <td  nowrap="nowrap"  height="100" class="FleetXGreySection" colspan="2">
        <!-- #include file="../BottomSection.asp" -->
    </td>
</tr>
<tr><td  nowrap="nowrap"  height="15" class="FleetXRedSectionSmall" colspan="2" align="center"><%=CopywriteNotice %></td></tr>
</td></tr></table>


</body>
</html>

