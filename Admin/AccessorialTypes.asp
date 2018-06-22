<html>
<head>

<!-- #include file="../fleetexpress.inc" -->
<link rel="stylesheet" type="text/css" href="../css/Style.css">

<%
Add = valid8(trim(request.form("add")))
if Add = "Y" then
  ErrorMessage = ""
  if len(valid8(trim(request.form("atBillCode")))) < 1 then
    ErrorMessage = "You must enter a Bill Code<br>"
  end if
  if len(valid8(trim(request.form("atDescr")))) < 1 then
    ErrorMessage = ErrorMessage & "You must enter a description<br>"
  end if 
  if len(ErrorMessage) < 1 then
      Set oConn = Server.CreateObject("ADODB.Connection")
      oConn.ConnectionTimeout = 100
      oConn.Provider = "MSDASQL"
      oConn.Open DATABASE
      SQL = "SELECT * FROM AccessorialType WHERE atBillCode='" & valid8(request.form("atBillCode")) & "'"
      SET oRs = oConn.Execute(SQL)
      if NOT oRs.EOF then
        ErrorMessage="Cannot add - Bill Code already exists"
      else
        'insert new one
        SQL="INSERT INTO AccessorialType (atBillCode, atDescr) values('"&valid8(request.form("atBillCode"))&"','"&valid8(request.form("atDescr"))&"')"
        SET oRs = oConn.Execute(SQL)
      end if
      Set oRS=Nothing
      Set oConn=Nothing
  end if 
end if

delID = valid8(trim(request.querystring("delID")))
if isNumeric(delID) and len(valid8(trim(request.querystring("aType")))) > 0 then
      Set oConn = Server.CreateObject("ADODB.Connection")
      oConn.ConnectionTimeout = 100
      oConn.Provider = "MSDASQL"
      oConn.Open DATABASE
      SQL = "SELECT accType FROM Accessorials WHERE atid=" & delID
      SET oRs = oConn.Execute(SQL)
      if NOT oRs.EOF then
        ErrorMessage="Cannot delete - Type is in use"
      else
        'delete
        SQL="DELETE FROM AccessorialType WHERE atID=" & delID
        SET oRs = oConn.Execute(SQL)
      end if
      Set oRS=Nothing
      Set oConn=Nothing
end if
%>
<%
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
    PageTitle="ACCESSORIAL TYPES"

%>
<title>FleetX - <%=PageTitle %></title>
</head>

<body onload="document.FindUser.requestorName.focus();">
	<table align="center" border="0" bordercolor="black" cellpadding="0" cellspacing="0" ID="Table1" height="100%" width="100%">
        <tr><td align="left"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>
        <tr>
            <td align="left"><img src="../images/FleetX_Small.jpg" height="50" width="168" /></td>
            <td align="right" valign="bottom"><!-- #include file="../topnavbar.asp" --></td>
        </tr>
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" /></td></tr>	
    <tr><td colspan="2">
<!-- <form action="NewUser.asp" method="post" name="FindUser">  -->
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
 
 Set oConn = Server.CreateObject("ADODB.Connection")
oConn.ConnectionTimeout = 100
oConn.Provider = "MSDASQL"
oConn.Open DATABASE



''''''''DISPLAY TYPES
Set Recordset1 = Server.CreateObject("ADODB.Recordset")
'Response.write "Database="&Database&"<br>"
Recordset1.ActiveConnection = Database
SQL="SELECT * FROM AccessorialType"
Recordset1.Source = SQL
'response.write "SQL="& SQL &"<BR>"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()
Recordset1_numRows = 0
'response.write "<font color='red'>SQL1="&Recordset1.Source&"<BR></font>"
%>
<table align="center" width="45%"><tr><td><b>Bill Code</b></td><td><b>Description</b></td><td><b>action</b></td></tr>
<%
	if NOT Recordset1.EOF then
    Do Until Recordset1.EOF
    
    %>
        <tr><td><%=Recordset1("atBillCode")%></td><td><%=Recordset1("atDescr")%></td><td><a href="AccessorialTypeEdit.asp?id=<%=Recordset1("atID")%>" class="FleetXRedMain">edit</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="AccessorialTypes.asp?delID=<%=Recordset1("atID")%>&rType=<%=Recordset1("atType")%>" class="FleetXRedMain">remove</a></td></tr>
    <%
    Recordset1.MoveNext
    Loop
	Else
      %><tr><td colspan=6>NO TYPES FOUND</td></tr><%
  End if
Recordset1.Close()
Set Recordset1 = Nothing    
 %>   
<tr><td colspan=6>&nbsp;<br><br><b>ADD NEW TYPE:</b></td></tr>

<form action="AccessorialTypes.asp" method="post" name="NewRateType">
<tr><td><input type=text size=13 maxlength=10 name="atBillCode"></td>
<td><input type=text name="atDescr"></td>
<input type="hidden" name="add" value="Y">
<td><INPUT id="gobutton" name="gobutton" TYPE="submit" name="ButtonValue" VALUE="ADD">
</form>
<tr><td colspan=6>&nbsp;<br><br><a href="AccessorialMaint.asp" class="FleetXRedMain">CLICK HERE</a> to Return to Accessorial Maintenance</td></tr>
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
<%
if ErrorMessage>"" then%>
<tr><td>
<table width="100%" border="0" bordercolor="Yellow" cellspacing="0" cellpadding="0" align="center" class="MainPageText">
	<tr><td>&nbsp;</td></tr>  
	<tr>
    <td align="center" class="Errormessage"><%=ErrorMessage%></td>
  </tr>
	<tr><td>&nbsp;</td></tr>
</table>
</td></tr>
<%end if%>
</table>
<!-- </form>  -->
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

