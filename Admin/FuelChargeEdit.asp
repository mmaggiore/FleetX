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
  NewFuelCharge = valid8(trim(request.form("NewFuelCharge")))
  if NOT isNull(NewFuelCharge) then
    if isNumeric(NewFuelCharge) then
      Set oConn = Server.CreateObject("ADODB.Connection")
      oConn.ConnectionTimeout = 100
      oConn.Provider = "MSDASQL"
      oConn.Open DATABASE
      'get current record
      SQL="SELECT * from FuelChargeList WHERE fuelchargeID=" & ID
      SET oRs = oConn.Execute(SQL)
        cFuelChargeType = oRs("FuelChargeType")
        'update old one to history
        SQL="UPDATE FuelChargeList set FuelChargeStatus='h' WHERE fuelchargeID=" & ID
        SET oRs = oConn.Execute(SQL)
          'insert new one
          SQL="INSERT INTO FuelChargeList values('"&cFuelChargeType&"','"&NewFuelCharge&"','c','"&Now()&"','"&Request.cookies("FleetXCookie")("UserID")&"')"
          SET oRs = oConn.Execute(SQL)
          SQL="SELECT MAX(fuelchargeID) as fuelchargeID from FuelChargelist WHERE FuelChargeType = '" & cFuelChargeType &"' and FuelChargeStatus = 'c'"
          SET oRs = oConn.Execute(SQL)
          ID = oRs("fuelchargeID")
        Set oConn=Nothing
        Set oRS=Nothing
    else
      ErrorUpdate="Invalid New Fuel Charge - please try again"
    end if
  end if
end if

''''''''DISPLAY FUEL CHARGE
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.ConnectionTimeout = 100
oConn.Provider = "MSDASQL"
oConn.Open DATABASE
SQL="SELECT * from FuelChargeList WHERE FuelChargeID=" & ID
SET oRs = oConn.Execute(SQL)
if NOT oRs.EOF then
  FuelChargeType = oRs("FuelChargeType")
  FuelCharge = oRs("FuelCharge")
  FuelChargeDate = oRs("FuelChargeDate")
  ChangedBy = oRs("ChangedBy")
  if isNumeric(ChangedBy) then
    SQL="SELECT * FROM PreExistingRequestor WHERE RequestorID=" & ChangedBY
    SET oRsN = oConn.Execute(SQL)
    if NOT oRsN.EOF then
      ChangedBy = oRsN("RequestorName") & "(" & ChangedBy & ") "
    else
      ChangedBy = ""
    end if
  else
    ChangedBy = ""
  end if
    
else
  ErrorMessage = "Fuel Charge not found" 
end if 
 

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
    PageTitle="FUEL CHARGE EDIT"

%>
<title>FleetX - <%=PageTitle %></title>
</head>

<body onload="document.EditFuelCharge.FuelCharge.focus();">
	<table align="center" border="0" bordercolor="black" cellpadding="0" cellspacing="0" ID="Table1" height="100%" width="100%">
        <tr><td align="left"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>
        <tr>
            <td align="left"><img src="../images/FleetX_Small.jpg" height="50" width="168" /></td>
            <td align="right" valign="bottom"><!-- #include file="../topnavbar.asp" --></td>
        </tr>
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" /></td></tr>	
    <tr><td colspan="2">
<form action="FuelChargeEdit.asp" method="post" name="EditFuelCharge">
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
<table align="center" cellspacing=3 cellpadding=3>
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
  <!-- <tr><td><b>ID</b></td><td><%=ID%></td></tr>   -->
  <%  if FuelChargeType = "S" then
        FuelChargeDescr = "StockRoom"
      else
        FuelChargeDescr = "Fuel"
      end if
  %>
  <tr><td><b>Fuel Charge Type</b></td><td><%=FuelChargeDescr%></td></tr>
  <tr><td><b>Last Changed:</b></td><td><%=ChangedBy%> <%=FuelChargeDate%></td></tr>
  <tr><td><b>Current Fuel Charge</b></td><td><%=FuelCharge%>%</td></tr>
  <tr><td><b>New Fuel Charge</b></td><td><input type="text" name="NewFuelCharge"></td></tr>
  

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


  
<tr><td colspan=6>&nbsp;<br><br><a href="FuelChargeMaint.asp" class="FleetXRedMain">CLICK HERE</a> to Return to the Fuel Charges page</td></tr>

<tr><td colspan=6>&nbsp;<br><br>
<h3>Change History</h3>
<%
Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = Database
SQL="SELECT * FROM FuelChargeList WHERE (FuelChargeStatus='h')and FuelChargeType='" & FuelChargeType & "' order by FuelChargeDate DESC"
Recordset1.Source = SQL
'response.write "SQL="& SQL &"<BR>"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()
Recordset1_numRows = 0
%>
<table align="center" width="750"><tr><td><b>ID</b></td><td><b>Fuel Charge Type</b></td><td><b>Fuel Charge</b></td><td><b>Changed</b></td><td><b>Changed By</b></td></tr>
<%
	if NOT Recordset1.EOF then
    Do Until Recordset1.EOF
    
    FuelChargeType=Recordset1("FuelChargeType")
    if FuelChargeType = "S" then
      FuelChargeDescr = "StockRoom"
    else
      FuelChargeDescr = "Fuel"
    end if
    
    FuelCharge = Recordset1("FuelCharge")
    %>
        <tr><td><%=Recordset1("fuelchargeID")%></td><td><%=FuelChargeDescr%></td><td><%=FuelCharge%>%</td>
        <td><%=Recordset1("FuelChargeDate")%></td>
        <td> 
<%
  HistoryChange = RecordSet1("ChangedBy")
  if isNumeric(HistoryChange) then
    SQL="SELECT * FROM PreExistingRequestor WHERE RequestorID=" & HistoryChange
    SET oRsN2 = oConn.Execute(SQL)
    if NOT oRsN2.EOF then
      HistoryChange = oRsN2("RequestorName") & "(" & HistoryChange & ") "
    else
      HistoryChange = "(" & HistoryChange & ") UNKNOWN"
    end if
  else
    HistoryChange = "(" & HistoryChange & ") UNKNOWN"
  end if
%>       
    <%=HistoryChange%>
    
      </td></tr>
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

