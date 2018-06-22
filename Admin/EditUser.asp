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
      Set oConn = Server.CreateObject("ADODB.Connection")
      oConn.ConnectionTimeout = 100
      oConn.Provider = "MSDASQL"
      oConn.Open DATABASE
      NewStatus = valid8(request.form("actionStatus"))
      SQL="UPDATE PreExistingRequestor set RequestorName = '" & valid8(request.form("RequestorName")) & "', CostCenter = '" & valid8(request.form("CostCenter")) & "', bt_id =" & valid8(request.form("btCust")) & ", RequestorState = '" & valid8(request.form("RequestorState")) & _ 
      "', RequestorType='" & valid8(request.form("UserType")) & "', RequestorAddress = '" & valid8(request.form("RequestorAddress")) & _
      "', RequestorCity = '" & valid8(request.form("RequestorCity")) & "', RequestorZipcode = '" & valid8(request.form("RequestorZipcode")) & "', RequestorPhone = '" & _
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
    PageTitle="EDIT USER"

%>
<title>FleetX - <%=PageTitle %></title>
</head>

<body onload="document.EditUser.RequestorName.focus();">
	<table align="center" border="0" bordercolor="black" cellpadding="0" cellspacing="0" ID="Table1" height="100%" width="100%">
        <tr><td align="left"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>
        <tr>
            <td align="left"><img src="../images/FleetX_Small.jpg" height="50" width="168" /></td>
            <td align="right" valign="bottom"><!-- #include file="../topnavbar.asp" --></td>
        </tr>
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" /></td></tr>	
    <tr><td colspan="2">
<form action="EditUser.asp" method="post" name="EditUser">
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
  <tr><td><b>Name</b></td><td><input type="text" name="RequestorName" value="<%=oRs("RequestorName")%>"></td></tr>
  <tr><td><b>Company</b></td><td>
  <select name="btCust">
  <%
  SQL = "SELECT * FROM fcbillto WHERE bt_status = 'c'"
  SET oRsN = oConn.Execute(SQL)
  Do Until oRsN.EOF
        selectd = ""
        BillTo = oRs("bt_id")
        if NOT isNULL(BillTo) and len(BillTo) > 1 then
          if cint(BillTo) = cint(oRsN("bt_id")) then 
            selectd = " selected" 
          end if
        end if
        %><option value=<%=cint(oRsN("bt_id"))%> <%=selectd%>><%=trim(oRsN("bt_desc"))%></option><%
  oRsN.MoveNext
  Loop
  %>
  </select>
  </td></tr>

  <tr><td><b>Cost Center</b></td><td><input type="text" maxlength=15 name="CostCenter" value="<%=oRs("CostCenter")%>"></td></tr>

  <tr><td><b>User Type</b></td><td>
  <select name="UserType">
  <% 
        selectd = ""
        thisType = oRs("RequestorType")
             if thisType = "A" then 
                selectd = " selected"
                else
                selectd2 = " selected"
          end if
        %><option value="A" <%=selectd%>>ADMIN</option><%
        %><option value="" <%=selectd2%>>USER</option><%
  %>
  </select>
  </td></tr>
  <tr><td><b>Address</b></td><td><input type="text" name="RequestorAddress" value="<%=oRs("RequestorAddress")%>"></td></tr>
  <tr><td><b>City</b></td><td><input type="text" name="RequestorCity" value="<%=oRs("RequestorCity")%>"></td></tr>
  <tr><td><b>State</b></td><td><input type="text" name="RequestorState" value="<%=oRs("RequestorState")%>"></td></tr>
  <tr><td><b>Zip code</b></td><td><input type="text" name="RequestorZipcode" value="<%=oRs("RequestorZipcode")%>"></td></tr>
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


  
<tr><td colspan=6>&nbsp;<br><br><a href="UserList.asp" class="FleetXRedMain">CLICK HERE</a> to Return to Users List Page</td></tr>


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

