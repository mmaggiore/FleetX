<html>
<head>

<!-- #include file="../fleetexpress.inc" -->
<link rel="stylesheet" type="text/css" href="../css/Style.css">
<%
ID = valid8(trim(request.form("ID")))
if len(ID) < 1 then
  ID = valid8(trim(request.querystring("ID")))
end if

action = valid8(trim(request.form("action")))
if len(action) < 1 then
  action = valid8(trim(request.querystring("action")))
end if

if action <> "add" then
  action = "edit"
end if

ErrorMessage =""

Edit = valid8(trim(request.form("Edit")))
Add = valid8(trim(request.form("Add")))

if Edit = "Y" then
      Set oConn = Server.CreateObject("ADODB.Connection")
      oConn.ConnectionTimeout = 100
      oConn.Provider = "MSDASQL"
      oConn.Open DATABASE
      SQL="UPDATE fcbillto set bt_desc ='" & valid8(request.form("btdesc")) & "', bt_addr1 ='" & _
      valid8(request.form("btaddr1")) & "', bt_addr2 = '" & valid8(request.form("btaddr2")) & _
      "', bt_city = '" & valid8(request.form("btcity")) & "', bt_state = '" & valid8(request.form("btstate")) & "', bt_zip = '" & valid8(request.form("btzip")) & "', bt_country = '" & _
      valid8(request.form("btcountry")) & "' WHERE bt_id=" & ID
      'response.write "27 sql=" & SQL & "<br>"
      SET oRs = oConn.Execute(SQL)
      SQL = "SELECT * FROM fcbillto WHERE bt_status = 'c'"
      SET oRs = oConn.Execute(SQL)
      SET oRs = Nothing
      SET oConn = NOthing
end if

if Add = "Y" then
      Set oConn = Server.CreateObject("ADODB.Connection")
      oConn.ConnectionTimeout = 100
      oConn.Provider = "MSDASQL"
      oConn.Open DATABASE
      SQL="SELECT MAX(bt_id) as btid FROM fcbillto"
      SET oRs = oConn.Execute(SQL)
      btid =  cint(oRs("btid")) + 1
      SQL="INSERT INTO fcbillto (bt_id,bt_desc,bt_addr1,bt_addr2,bt_city,bt_state,bt_zip,bt_country,bt_status) VALUES(" & btid & ",'" & _
      valid8(trim(request.form("btdesc"))) & "','" & valid8(trim(request.form("btaddr1"))) & "','" & valid8(trim(request.form("btaddr2"))) & _
      "','" & valid8(trim(request.form("btcity"))) & "','" & valid8(trim(request.form("btstate"))) & "','" & valid8(trim(request.form("btzip"))) & "','" & valid8(trim(request.form("btcountry"))) & "','c')"
      'response.write "54 sql=" & SQL & "<br>"
      SET oRs = oConn.Execute(SQL)
      response.redirect "CompanyList.asp"
      SET oRs = Nothing
      SET oConn = NOthing
end if


if action = "edit" then
  ''''''''DISPLAY COMPANY
  Set oConn = Server.CreateObject("ADODB.Connection")
  oConn.ConnectionTimeout = 100
  oConn.Provider = "MSDASQL"
  oConn.Open DATABASE
  SQL = "SELECT * FROM fcbillto WHERE bt_id = " & ID
  SET oRs = oConn.Execute(SQL)
  if oRs.EOF then
    ErrorMessage = "Company not found" 
  end if 
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
    if action = "add" then
      PageTitle = "ADD COMPANY"
    else
      PageTitle="EDIT COMPANY"
    end if

%>
<title>FleetX - <%=PageTitle %></title>
</head>

<body onload="document.EditCompany.btdesc.focus();">
	<table align="center" border="0" bordercolor="black" cellpadding="0" cellspacing="0" ID="Table1" height="100%" width="100%">
        <tr><td align="left"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>
        <tr>
            <td align="left"><img src="../images/FleetX_Small.jpg" height="50" width="168" /></td>
            <td align="right" valign="bottom"><!-- #include file="../topnavbar.asp" --></td>
        </tr>
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" /></td></tr>	
    <tr><td colspan="2">
<form action="EditCompany.asp" method="post" name="EditCompany">
<table border="0" bordercolor="green" Cellspacing="0" Cellpadding="0" align="left" width="100%">
 <tr><td align="left" colspan="2" bgcolor="<%=HeaderBorderColor%>"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>

        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetXRedSection" colspan="2"><%=PageTitle%></td></tr>

        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>
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
  <tr><td><b>Name</b></td><td><input type="text" name="btdesc" <% if action = "edit" then %>value="<%=trim(oRs("bt_desc"))%>"<%end if%>></td></tr>
  <tr><td><b>Address 1</b></td><td><input type="text" name="btaddr1" <% if action = "edit" then %>value="<%=trim(oRs("bt_addr1"))%>"<%end if%>></td></tr>
  <tr><td><b>Address 2</b></td><td><input type="text" name="btaddr2" <% if action = "edit" then %>value="<%=trim(oRs("bt_addr2"))%>"<%end if%>></td></tr>
  <tr><td><b>City</b></td><td><input type="text" name="btcity" <% if action = "edit" then %>value="<%=trim(oRs("bt_city"))%>"<%end if%>></td></tr>
  <tr><td><b>State</b></td><td><input type="text" name="btstate" <% if action = "edit" then %>value="<%=trim(oRs("bt_state"))%>"<%end if%>></td></tr>
  <tr><td><b>Zip code</b></td><td><input type="text" name="btzip" <% if action = "edit" then %>value="<%=trim(oRs("bt_zip"))%>"<%end if%>></td></tr>
  <tr><td><b>Country</b></td><td><input type="text" name="btcountry" <% if action = "edit" then %>value="<%=trim(oRs("bt_country"))%>"<%end if%>></td></tr>
  <tr><td> </td><td><INPUT id="gobutton" name="gobutton" TYPE="submit" name="ButtonValue" VALUE="SUBMIT"></td></tr> 
  <input type="hidden" name="ID" value="<%=id%>">
<% if action = "add" then %>
  <input type="hidden" name="Add" value="Y">
<% else %>
  <input type="hidden" name="Edit" value="Y">
<%end if %>
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


  
<tr><td colspan=6>&nbsp;<br><br><a href="CompanyList.asp" class="FleetXRedMain">CLICK HERE</a> to Return to Company List Page</td></tr>


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

