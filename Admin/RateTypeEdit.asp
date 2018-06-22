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
      Set oConn = Server.CreateObject("ADODB.Connection")
      oConn.ConnectionTimeout = 100
      oConn.Provider = "MSDASQL"
      oConn.Open DATABASE

Edit = valid8(trim(request.form("Edit")))

if Edit = "Y" then
  ErrorUpdate = ""
  if len(valid8(trim(request.form("rtBillCode")))) < 1 then
    ErrorUpdate = "You must enter a Bill Code<br>"
  else
     SQL = "SELECT * from RateType where rtBillCode = '" & valid8(request.form("rtBillCode")) & "' and rtid <> " & ID
     set oRs = oConn.Execute(SQL)
     if NOT oRs.EOF then
      ErrorUpdate = ErrorUpdate & "Bill Code already in use - please try again<br>"
     end if
  end if
  if len(valid8(trim(request.form("rtDescr")))) < 1 then
    ErrorUpdate = ErrorUpdate & "You must enter a description<br>"
  end if 
  If len(ErrorUpdate) < 1 then
      SQL="UPDATE RateType set rtBillCode='" & valid8(request.form("rtBillCode")) & "', rtDescr = '" & valid8(request.form("rtDescr")) & "' WHERE rtID=" & ID
      'response.write "22 sql=" & SQL & "<br>"
      SET oRs = oConn.Execute(SQL)
          'insert new one
          'SQL="INSERT INTO FuelChargeList values('"&cFuelChargeType&"','"&NewFuelCharge&"','c','"&Now()&"','"&Request.cookies("FleetXCookie")("UserID")&"')"
          'SET oRs = oConn.Execute(SQL)
  end if
end if

''''''''DISPLAY RATE TYPE
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.ConnectionTimeout = 100
oConn.Provider = "MSDASQL"
oConn.Open DATABASE
SQL="SELECT * from RateType WHERE rtID=" & ID
SET oRs = oConn.Execute(SQL)
if NOT oRs.EOF then
  rtBillCode = oRs("rtBillCode")
  rtDescr = oRs("rtDescr")
else
  ErrorMessage = "Rate Type not found" 
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
    PageTitle="RATE TYPE EDIT"

%>
<title>FleetX - <%=PageTitle %></title>
</head>

<body onload="document.EditRateType.rtBillCodefocus();">
	<table align="center" border="0" bordercolor="black" cellpadding="0" cellspacing="0" ID="Table1" height="100%" width="100%">
        <tr><td align="left"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>
        <tr>
            <td align="left"><img src="../images/FleetX_Small.jpg" height="50" width="168" /></td>
            <td align="right" valign="bottom"><!-- #include file="../topnavbar.asp" --></td>
        </tr>
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" /></td></tr>	
    <tr><td colspan="2">
<form action="RateTypeEdit.asp" method="post" name="EditRateType">
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
  <tr><td><b>Rate Bill Code</b></td><td><input type=text name="rtBillCode" value="<%=rtBillCode%>"></td></tr>
  <tr><td><b>Rate Descr</b></td><td><input type="text" name="rtDescr" value="<%=rtDescr%>"></td></tr>
  <input type="hidden" name="rtType" value="<%=rtType%>">
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


  
<tr><td colspan=6>&nbsp;<br><br><a href="RateTypes.asp" class="FleetXRedMain">CLICK HERE</a> to Return to Rate Types</td></tr>


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

