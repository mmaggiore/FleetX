<html>
<head>

<!-- #include file="fleetexpress.inc" -->
<link rel="stylesheet" type="text/css" href="css/Style.css">
<%

ID = Request.cookies("FleetXCookie")("UserID")

ErrorMessage =""

Edit = valid8(trim(request.form("Edit")))

if Edit = "Y" then
      Set oConn = Server.CreateObject("ADODB.Connection")
      oConn.ConnectionTimeout = 100
      oConn.Provider = "MSDASQL"
      oConn.Open DATABASE
      NewStatus = valid8(request.form("actionStatus"))
      SQL="UPDATE PreExistingRequestor set bt_id =" & valid8(request.form("btCust")) & ", RequestorState = '" & valid8(request.form("RequestorState")) & _ 
      "', RequestorType='" & valid8(request.form("UserType")) & "', RequestorAddress = '" & valid8(request.form("RequestorAddress")) & _
      "', RequestorCity = '" & valid8(request.form("RequestorCity")) & "', RequestorZipcode = '" & valid8(request.form("RequestorZipcode")) & "', CostCenter = '" & valid8(request.form("RequestorCostCenter")) & "', RequestorPhone = '" & _
      valid8(request.form("RequestorPhone")) & "', RequestorEmail = '" & valid8(request.form("RequestorEmail")) & _
      "', RequestorPassword = '" & valid8(request.form("RequestorPassword")) & "' WHERE requestorID=" & ID
      'response.write "27 sql=" & SQL & "<br>"
      SET oRs = oConn.Execute(SQL)
      SQL = "SELECT * from PreExistingRequestor WHERE requestorID = " & ID
      SET oRs = oConn.Execute(SQL)
end if

''''''''DISPLAY NEW USER
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.ConnectionTimeout = 100
oConn.Provider = "MSDASQL"
oConn.Open DATABASE
SQL="SELECT * from PreExistingRequestor WHERE requestorID=" & ID
SET oRs = oConn.Execute(SQL)
if oRs.EOF then
  ErrorMessage = "User not found" 
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
    PageTitle="UPDATE USER INFO"

%>
<title>FleetX - <%=PageTitle %></title>
</head>

<body onload="document.EditUser.RequestorName.focus();">
	<table align="center" border="0" bordercolor="black" cellpadding="0" cellspacing="0" ID="Table1" height="100%" width="100%">
        <tr><td align="left"><img src="images/pixel.gif" height="10" width="1" /></td></tr>
        <tr>
            <td align="left"><img src="images/FleetX_Small.jpg" height="50" width="168" /></td>
            <td align="right" valign="bottom"><!-- #include file="topnavbar.asp" --></td>
        </tr>
        <tr><td align="left" colspan="2"><img src="images/pixel.gif" /></td></tr>	
    <tr><td colspan="2">
<form action="UpdateUserInfo.asp" method="post" name="EditUser">
<table border="0" bordercolor="green" Cellspacing="0" Cellpadding="0" align="left" width="100%">
 <tr><td align="left" colspan="2" bgcolor="<%=HeaderBorderColor%>"><img src="images/pixel.gif" height="5" width="1" /></td></tr>

        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetXRedSection" colspan="2"><%=PageTitle%></td></tr>

        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2"><img src="images/pixel.gif" height="5" width="1" /></td></tr>
        <!--tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressBodyWhite" colspan="2">In order to reach the Fleet Express Order page, please correctly type in the green verification code in the supplied text box and click "Submit."</td></tr-->
        <tr><td align="left" colspan="2"><img src="images/pixel.gif" height="5" width="1" /></td></tr>
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
  <tr><td><b>Name</b></td><td><input type="text" name="RequestorName" value="<%=oRs("RequestorName")%>"></td></tr>
  <tr><td><b>Company</b></td><td>
  <input type=hidden name=btCust value="<%=oRs("bt_id")%>">
  <%
  SQL = "SELECT * FROM fcbillto WHERE bt_id = " & oRs("bt_id")
  SET oRsN = oConn.Execute(SQL)
  if not oRsN.eof then
    response.write oRsN("bt_desc")
    
  else
    response.write "UNKNOWN"
  end if
  %>
  </td></tr>
  <tr><td><b>User Type</b></td><td>
  
  <% 
        selectd = ""
        thisType = oRs("RequestorType")
          if thisType = "A" then 
            response.write "ADMIN" 
          else
            response.write "USER"
          end if
        %>
     <input type=hidden name=UserType value="<%=thisType%>">

  </td></tr>
  <tr><td><b>Address</b></td><td><input type="text" name="RequestorAddress" value="<%=oRs("RequestorAddress")%>"></td></tr>
  <tr><td><b>City</b></td><td><input type="text" name="RequestorCity" value="<%=oRs("RequestorCity")%>"></td></tr>
  <tr><td><b>State</b></td><td><input type="text" name="RequestorState" value="<%=oRs("RequestorState")%>"></td></tr>
  <tr><td><b>Zip code</b></td><td><input type="text" name="RequestorZipcode" value="<%=oRs("RequestorZipcode")%>"></td></tr>
  <tr><td><b>Cost Center</b></td><td><input type="text" name="RequestorCostCenter" value="<%=oRs("CostCenter")%>"></td></tr>
  <tr><td><b>Phone</b></td><td><input type="text" name="RequestorPhone" value="<%=oRs("RequestorPhone")%>"></td></tr>
  <tr><td><b>Email</b></td><td><input type="text" name="RequestorEmail" value="<%=oRs("RequestorEmail")%>"></td></tr>
  <tr><td><b>Password</b></td><td><input type="text" name="RequestorPassword" value="<%=oRs("RequestorPassword")%>"></td></tr>
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


  
<tr><td colspan=6>&nbsp;<br><br><a href="home.asp" class="FleetXRedMain">CLICK HERE</a> to Return to Home Page</td></tr>


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
        <!-- #include file="BottomSection.asp" -->
    </td>
</tr>
<tr><td height="15" class="FleetXRedSectionSmall" colspan="2" align="center"><%=CopywriteNotice %></td></tr>
</td></tr></table>


</body>
</html>

