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
      'StartDate = valid8(trim(request.form("StartDate")))
      'EndDate = valid8(trim(request.form("EndDate")))
      RateTypeID = valid8(trim(request.form("rtid")))
      RateCharge = valid8(trim(request.form("rateCharge")))
       'check to be sure the accType isn't already set for this company:
        SQL = "SELECT * from RateList WHERE bt_id = " & btCust & " AND rtid = " & rateTypeID
        SET oRs = oConn.Execute(SQL)
        if NOT oRs.EOF then
          errMessage = "Rate Type already exists for this Company - please try again<br><br>"
        else
        'insert new one
          SQL = "INSERT INTO RateList (bt_id,rtid,rateCharge,ratedate,changedby,RateStatus) values(" & btCust & "," & RateTypeID & ",'" &RateCharge&"','"&Now()&"','"&Request.cookies("FleetXCookie")("UserID")&"','c')"
          'SQL="INSERT INTO RateList values(" & btCust & "," & RateTypeID & ",'" & StartDate & "','" & EndDate & "','','"&RateCharge&"','c','"&Now()&"','"&Request.cookies("FleetXCookie")("UserID")&"')"
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
        SQL="UPDATE RateList SET rateStatus = 'x' WHERE rateID=" & delID
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
    PageTitle="RATE CHARGE MAINTENANCE"

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
 
''''''''DISPLAY RATES CHARGES
Set Recordset1 = Server.CreateObject("ADODB.Recordset")
'Response.write "Database="&Database&"<br>"
Recordset1.ActiveConnection = Database
SQL = "SELECT r.rateID, r.bt_id, r.rtid, r.rateCharge, r.rateStatus, r.rateDate, r.changedby, r.bstartdate, r.benddate, b.bt_id, b.bt_desc "_
& " FROM RateList r "_
& " INNER JOIN fcbillto b on b.bt_id = r.bt_id "_
& " WHERE (r.RateStatus='c') order by b.bt_desc, r.rtid, r.bstartdate"

'SQL="SELECT * FROM RateList WHERE (RateStatus='c')"
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
<!-- <tr><td  nowrap="nowrap" ><b>Customer</b></td><td  nowrap="nowrap" ><b>Rate Type</b></td><td><b>Bill Code</b></td><td  nowrap="nowrap" ><b>StartDate</b></td><td  nowrap="nowrap" ><b>EndDate</b></td><td  nowrap="nowrap" ><b>Rate Charge</b></td><td  nowrap="nowrap" ><b>Changed</b></td><td  nowrap="nowrap" ><b>ChangedBy</b></td><td  nowrap="nowrap" ><b>action</b></td></tr>  -->
<tr><td  nowrap="nowrap" ><b>Customer</b></td><td  nowrap="nowrap" ><b>Rate Type</b></td><td><b>Bill Code</b></td><td  nowrap="nowrap" ><b>Rate Charge</b></td><td  nowrap="nowrap" ><b>Added</b></td><td  nowrap="nowrap" ><b>Added By</b></td><td  nowrap="nowrap" ><b>action</b></td></tr>
<%
	if NOT Recordset1.EOF then
    Do Until Recordset1.EOF
    
    RateTypeID=Recordset1("rtid")
      SQL = "SELECT * FROM RateType WHERE rtid = '" & RateTypeID & "'"
      SET oRsN = oConn.Execute(SQL)
      if NOT oRsN.EOF then
        RateDescr = oRsN("rtDescr")
        BillCode = oRsN("rtBillCode")
      else
        RateDescr = "UNKNOWN"
      end if
      set oRsN = Nothing
          
    RateCharge = Recordset1("RateCharge")
    RateCharge = FormatCurrency(RateCharge,2)
    
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
  TheDateTime=Recordset1("RateDate")
  TheDayOnly=Day(TheDateTime)
  TheMonthOnly=Month(TheDateTime)
  TheYearOnly=Year(TheDateTime)
  DateOnly=TheMonthOnly&"/"&TheDayOnly&"/"&TheYearOnly
    %>
      <!--  <tr><td  nowrap="nowrap" ><%=Recordset1("bt_ID")%>-<%=Recordset1("bt_desc")%></td><td  nowrap="nowrap" ><%=RateDescr%></td><td><%=BillCode%></td><td  nowrap="nowrap" ><%=Recordset1("bstartdate")%></td><td  nowrap="nowrap" ><%=Recordset1("benddate")%></td><td  nowrap="nowrap" ><%=trim(RateCharge)%></td><td  nowrap="nowrap" ><%=DateOnly %></td><td  nowrap="nowrap" ><%=ChangedBy%></td><td  nowrap="nowrap" ><a href="RateEdit.asp?id=<%=Recordset1("RateID")%>" class="FleetXRedMain">edit</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="RateMaint.asp?delID=<%=Recordset1("rateID")%>" class="FleetXRedMain">remove</a></td></tr>     -->
        <tr><td  nowrap="nowrap" ><%=Recordset1("bt_ID")%>-<%=Recordset1("bt_desc")%></td><td  nowrap="nowrap" ><%=RateDescr%></td><td><%=BillCode%></td><td  nowrap="nowrap" ><%=trim(RateCharge)%></td><td  nowrap="nowrap" ><%=DateOnly %></td><td  nowrap="nowrap" ><%=ChangedBy%></td><td  nowrap="nowrap" ><a href="RateEdit.asp?id=<%=Recordset1("RateID")%>" class="FleetXRedMain">edit</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="RateMaint.asp?delID=<%=Recordset1("rateID")%>" class="FleetXRedMain">remove</a></td></tr>
    <%
    Recordset1.MoveNext
    Loop
	Else
      %><tr><td  nowrap="nowrap"  colspan=6>NO RATE CHARGES FOUND</td></tr><%
  End if
Recordset1.Close()
Set Recordset1 = Nothing    
 %>   
 <tr><td  nowrap="nowrap"  colspan=6>&nbsp;<br><b>ADD NEW RATE:</b></td></tr>
<form action="RateMaint.asp" method="post" name="NewRate">
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
  <select name="rtid">
  <%
  SQL = "SELECT * FROM RateType ORDER BY rtDescr"
  SET oRsN = oConn.Execute(SQL)
  Do Until oRsN.EOF
        %><option value=<%=oRsN("rtid")%>><%=oRsN("rtDescr")%></option><%
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
  </td>   -->
  <td  nowrap="nowrap" ><input type="text" name="rateCharge"></td>
  <input type="hidden" name="add" value="Y">
   <td  nowrap="nowrap"  colspan=3><INPUT id="gobutton" name="gobutton" TYPE="submit" name="ButtonValue" VALUE="ADD"></td>
 </tr>
 </form>
<tr><td  nowrap="nowrap"  colspan=6>&nbsp;<br><br><a href="ratetypes.asp" class="FleetXRedMain"><b>add/edit rate types</b></a></td></tr>
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

