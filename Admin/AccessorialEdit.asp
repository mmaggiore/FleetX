<html>
<head>

<!-- #include file="../fleetexpress.inc" -->
<link rel="stylesheet" type="text/css" href="../css/Style.css">
<%
ID = valid8(trim(request.form("ID")))
if len(ID) < 1 then
  ID = valid8(trim(request.querystring("ID")))
end if

ErrorMessage =""

Edit = valid8(trim(request.form("Edit")))

if Edit = "Y" then
  NewCharge = valid8(trim(request.form("NewCharge")))
  btCust = valid8(trim(request.form("btCust")))
  'StartDate = valid8(trim(request.form("StartDate")))
  'EndDate = valid8(trim(request.form("EndDate")))
  ErrorUpdate =""
  NewCharge = replace(NewCharge,"$","")
  if NOT isNumeric(NewCharge) or isNull(NewCharge) then
      ErrorUpdate="Invalid New Charge - please try again<br>"
  end if
  'if NOT isDate(StartDate) then
      'ErrorUpdate = ErrorUpdate & "You must provide a valid Start Date<br>"
  'end if
  'if NOT isDate(EndDate) then
      'ErrorUpdate = ErrorUpdate & "You must provide a valid End Date<br>"
  'end if
  'If isDate(StatDate) and isDate(EndDate) then
    'if cdate(EndDate) < cdate(StartDate) then
      'ErrorUpdate = ErrorUpdate & "End Date must be greater than Start Date<br>"
    'end if
  'end if
    if len(ErrorUpdate) < 1 then
      Set oConn = Server.CreateObject("ADODB.Connection")
      oConn.ConnectionTimeout = 100
      oConn.Provider = "MSDASQL"
      oConn.Open DATABASE
      'get current record
      SQL="SELECT * from Accessorials WHERE accID=" & ID
      SET oRs = oConn.Execute(SQL)
        cTypeID = oRs("atid")
        'update old one to history
        SQL="UPDATE Accessorials set accStatus='x' WHERE accID=" & ID
        SET oRs = oConn.Execute(SQL)
          'insert new one
          SQL = "INSERT INTO Accessorials (bt_id,atid,accCharge,accdate,changedby,accStatus) values(" & btCust & "," & cTypeID & ",'" &NewCharge&"','"&Now()&"','"&Request.cookies("FleetXCookie")("UserID")&"','c')"
          'SQL="INSERT INTO Accessorials values(" & btCust & "," & cTypeID &",'" & StartDate & "','" & EndDate & "','','"&NewCharge&"','"&Now()&"','"&Request.cookies("FleetXCookie")("UserID")&"','c')"
          'response.write "50 SQL=" & SQL & "<br>"
          SET oRs = oConn.Execute(SQL)
          SQL="SELECT MAX(accID) as accID from Accessorials WHERE atid = " & cTypeID &" and accStatus = 'c'"
          SET oRs = oConn.Execute(SQL)
          ID = oRs("accID")
        Set oConn=Nothing
        Set oRS=Nothing
        response.redirect "AccessorialMaint.asp"
    end if
end if

''''''''DISPLAY 
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.ConnectionTimeout = 100
oConn.Provider = "MSDASQL"
oConn.Open DATABASE
SQL="SELECT * from Accessorials WHERE accID=" & ID
SET oRs = oConn.Execute(SQL)
if NOT oRs.EOF then
  accType = oRs("atid")
  accCharge = oRs("accCharge")
  BillTo = oRs("bt_id")
  StartDate = oRs("accstartdate")
  EndDate = oRs("accstopdate")
else
  ErrorMessage = "Not found" 
end if 
 
'Set oConn=Nothing
Set oRS=Nothing


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
    PageTitle="ACCESSORIAL EDIT"

%>
<title>FleetX - <%=PageTitle %></title>
<script language="javascript" type="text/javascript" src="datetimepicker.js">
    //Date Time Picker script- by TengYong Ng of http://www.rainforestnet.com
</script>
</head>

<body onload="document.EditRate.RateCharge.focus();">
	<table align="center" border="0" bordercolor="black" cellpadding="0" cellspacing="0" ID="Table1" height="100%" width="100%">
        <tr><td align="left"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>
        <tr>
            <td align="left"><img src="../images/FleetX_Small.jpg" height="50" width="168" /></td>
            <td align="right" valign="bottom"><!-- #include file="../topnavbar.asp" --></td>
        </tr>
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" /></td></tr>	
    <tr><td colspan="2">
<form action="AccessorialEdit.asp" method="post" name="EditAcc">
<table border="0" bordercolor="green" Cellspacing="0" Cellpadding="0" align="left" width="100%">
 <tr><td align="left" colspan="2" bgcolor="<%=HeaderBorderColor%>"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>

        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetXRedSection" colspan="2"><%=PageTitle%></td></tr>

        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
        <!--tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressBodyWhite" colspan="2">In order to reach the Fleet Express Order page, please correctly type in the green verification code in the supplied text box and click "Submit."</td></tr-->
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
<tr><td>
<table  border="0" bordercolor="blue" align="center" class="MainPageText" width="100%">
	<tr height="40">
		<td width="650">&nbsp;</td>
	</tr>



    <tr><td align=center width="100%"><!-- main page stuff goes here! -->
    
    
 <%
%>
<table align="center" cellspacing=3 cellpadding=3 class="MainPageText">
<% if len(ErrorMessage) < 1 then %>
  <% if len(ErrorUpdate) > 0 then %>
    <tr><td colspan=6>
    <table width="100%" border="0" bordercolor="Yellow" cellspacing="0" cellpadding="0" align="center" class="MainPageText">
    	<tr><td>&nbsp;</td></tr>  
    	<tr>
        <td align="center" class="Errormessage"><%=ErrorUpdate%></td>
      </tr>
    	<tr><td>&nbsp;</td></tr>
    </table>
    </td></tr>
  <% end if %>
<!--  <tr><td><b>ID</b></td><td><%=ID%></td></tr>   -->
  <%  
      SQL = "SELECT * FROM AccessorialType WHERE atid = " & accType
      'response.write "163 sql=" & SQL & "<Br>"
      SET oRsN = oConn.Execute(SQL)
      if NOT oRsN.EOF then
        accDescr = oRsN("atDescr")
      else
        accDescr = "UNKNOWN"
      end if
      set oRsN = Nothing
  %>
  <tr><td><b>Customer</b></td><td>
  <select name="btCust">
  <%
  SQL = "SELECT * FROM fcbillto WHERE bt_status = 'c'"
  SET oRsN = oConn.Execute(SQL)
  Do Until oRsN.EOF
        selectd = ""
        'response.write "billto=" & cint(BillTo) & ",btid=" & cint(oRsN("bt_id")) & "<br>"
        if cint(BillTo) = cint(oRsN("bt_id")) then 
          selectd = " selected" 
        end if
        %><option value=<%=cint(oRsN("bt_id"))%> <%=selectd%>><%=trim(oRsN("bt_desc"))%></option><%
  oRsN.MoveNext
  Loop
  %>
  </select>
  </td></tr>
  <tr><td><b>Type</b></td><td><%=accDescr%></td></tr>
  <tr><td><b>Current Charge</b></td><td><%=FormatCurrency(accCharge)%></td></tr>
  <tr><td><b>New Charge</b></td><td><input type="text" name="NewCharge"></td></tr>
 <!-- <tr><td><b>Start Date</b></td><td><input type="text" id="StartDate" name="StartDate" value="<%=StartDate%>">
  <a href="javascript:NewCal('StartDate','MMddyyyy',true,12)"><img src="../images/cal.gif" width="16" height="16" border="0" alt="Pick a date"></a>               
  </td></tr>
  <tr><td><b>End Date</b></td><td><input type="text" id="EndDate" name="EndDate" value="<%=EndDate%>">
  <a href="javascript:NewCal('EndDate','MMddyyyy',true,12)"><img src="../images/cal.gif" width="16" height="16" border="0" alt="Pick a date"></a>               
  </td></tr> -->
  <tr><td> </td><td><INPUT id="gobutton" name="gobutton" TYPE="submit" name="ButtonValue" VALUE="SUBMIT"></td></tr> 
  <input type="hidden" name="ID" value="<%=id%>">
  <input type="hidden" name="Edit" value="Y">
<% else %>
  <tr><td colspan=6>
  <table width="100%" border="0" bordercolor="Yellow" cellspacing="0" cellpadding="0" align="center" class="MainPageText">
  	<tr><td>&nbsp;</td></tr>  
  	<tr>
      <td align="center" class="Errormessage"><%=ErrorMessage%></td>
    </tr>
  	<tr><td>&nbsp;</td></tr>
  </table>
  </td></tr>
<% end if %>


  
<tr><td colspan=6>&nbsp;<br><br><a href="AccessorialMaint.asp" class="FleetXRedMain">Click here</a> to Return to Accessorial Maintenance</td></tr>
</table>
</td></tr>

<tr><td colspan=6 align=center>
<table width="95%" align=center>
<tr><td colspan=6>&nbsp;<br><br>
<h3>Change History</h3>
<%
Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = Database
SQL="SELECT * FROM Accessorials WHERE (accStatus='h' or accStatus='x') and atid='" & accType & "' AND bt_id = " & BillTo & " order by accDate DESC"
Recordset1.Source = SQL
'response.write "221 SQL="& SQL &"<BR>"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()
Recordset1_numRows = 0
%>
<table align="center" width="95%">
<tr><td><b>Customer</b></td><td><b>Type</b></td><td><b>Charge</b></td><td><b>Added</b></td><td><b>Added By</b></td></tr>
<%
	if NOT Recordset1.EOF then
    Do Until Recordset1.EOF
 
    thisBTID = Recordset1("bt_id")
    SQL = "SELECT * FROM fcbillto WHERE bt_id = " & thisBTID
    SET oRsN = oConn.Execute(SQL)
    if NOT oRsN.EOF then
      btDESC = oRsN("bt_desc")
    else
      btDESC = "UNKNOWN"
    end if   
    RateType=Recordset1("atid")
      SQL = "SELECT * FROM AccessorialType WHERE atid = " & accType 
      SET oRsN = oConn.Execute(SQL)
      if NOT oRsN.EOF then
        accDescr = oRsN("atDescr")
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
    %>
        <tr><td><%=Recordset1("bt_ID")%>-<%=btDESC%></td><td><%=accDescr%></td><td><%=accCharge%></td><td><%=Recordset1("accDate")%></td><td><%=ChangedBy%></td></tr>
    <%
    Recordset1.MoveNext
    Loop
	Else
      %><tr><td colspan=6>NO HISTORY FOUND</td></tr><%
  End if
Recordset1.Close()
Set Recordset1 = Nothing 
Set oConn=Nothing   
 %>   
</table>
</td></tr>
</table>
   
    </td></tr>



 
	<tr Height="50">
		<td>&nbsp;</td>
	</tr>

  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
      &nbsp;
    </td>
	<td width="5">&nbsp;</td>
    <td width="136"> 
     &nbsp;
    </td>
	<td width="5">&nbsp;</td>
    <td width="725"> 
      &nbsp;
    </td>
  </tr>
</table>
</td></tr>

</table>
</form>
<tr><td Height="90%">&nbsp;</td></tr>
<tr>
    <td height="100" class="FleetXGreySection" colspan="2">
        <!-- #include file="../BottomSection.asp" -->
    </td>
</tr>
<tr><td height="15" class="FleetXRedSectionSmall" colspan="2" align="center"><%=CopywriteNotice %></td></tr>
</td></tr></table>


</body>
</html>

