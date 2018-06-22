<%@ LANGUAGE="VBSCRIPT"%>
<%
Response.buffer = True
TheTime=time()
'Response.Write "TheTime="&TheTime&"<BR>"
'If theTime<="6:00:00 PM" then
'	Response.Write "LATE<BR>!!!!"
'End if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<SCRIPT Language="Javascript">
	function validate()
	{
			//MARK'S ADDED CODE START
			if(document.Form1.TempPODID.value=="xxx" && document.Form1.addedPOD.value=="")
			{
				alert('You must select or manually type in your POD name.');
				document.Form1.TempPODID.focus();
				return false;
			}			
	}	
	</SCRIPT> 
	<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<TITLE><% Response.Write(D_TITLEBAR) %></TITLE>
	<!-- added the include style.css-->
<!-- #include file="FleetX.inc" -->
<!-- #include file="driverinfo.inc" -->
</head>
<% 
 JobNum = request.querystring("j")
 PageStatus = request.querystring("s")
 LocationCode = request.querystring("l")
%>
<body>
<!-- #include file="LogoSection.asp" -->
	<table width="300" border="0" bordercolor="red" cellpadding="0" cellspacing="0" align="left" ID="Table1">

        <form method="post" action="DriverClose.asp" ID="Form2">
		    <tr><td align="center" colspan="3">&nbsp;<br><input type="submit" value="<<<BACK" ID="gobutton" NAME="Submit1"></td></tr>
		    </form>
        
		<tr><td align="left">
			<table cellpadding="3" cellspacing="0" width="300" border="0" align="left" ID="Table5">
                        <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>
                        <tr>
		                    <td class="FleetXRedSection" colspan="3" align="center">
			                    ADD EXCEPTION
		                    </td>
	                    </tr>
                        <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>
				
		<%
			Dim colorset, i, numcolors
			'/--- This is your array of colors to use. -------------\ 
			colorset = split("#D2D2C0,White",",")
			numcolors = ubound(colorset)+1
			Server.ScriptTimeout = 1000
   %>
				<tr>
					<td nowrap valign="top">Job #<%=JobNum%>&nbsp;<br></td></tr>
          <tr><td><hr></td></tr>
          <tr><td nowrap>
          <%
          
                    Set oConn8 = Server.CreateObject("ADODB.Connection")
                    oConn8.ConnectionTimeout = 100
                    oConn8.Provider = "MSDASQL"
                    oConn8.Open DATABASE
          
              SQL = "SELECT * from ChargedAccessorials WHERE ca_fh_id = '" & JobNum & "'"
               SET oRsN = oConn8.Execute(SQL)
               if not oRsN.EOF then
                response.write "<table><tr><td colspan=2>EXISTING EXCEPTIONS:</td></tr>"
                Do While NOT oRsN.EOF
                    accCharge = oRsN("ca_accCharge")
                    accCharge2 = FormatCurrency(accCharge,2)
                    caType = oRsN("ca_Type")
                    accid = oRsN("ca_atid")
                    LocCode = oRsN("LocationCode")
                    'response.write "accid=" & accid & "<br>"
                    SQL2 = "SELECT * from AccessorialType WHERE atid = " & accid
                    SET oRsNn = oConn8.Execute(SQL2)
                    if NOT oRsNn.EOF then
                      accType = oRsNn("atDescr")
                    else
                      accType = "UNKNOWN"
                    end if

                  response.write "<tr><td>" & accType & " - " & accCharge2 & "&nbsp;" & caType & "&nbsp;" & LocCode
                  oRsN.MoveNext
                Loop
                  response.write "</td></tr>"
               response.write "</table>"
                   response.write "<tr><td nowrap><hr></td></tr>"
               End If
                response.write "<tr><td>&nbsp;<br><b>ADD NEW EXCEPTION:</b><br><br>"
                              Set oConn89 = Server.CreateObject("ADODB.Connection")
                    oConn89.ConnectionTimeout = 100
                    oConn89.Provider = "MSDASQL"
                    oConn89.Open DATABASE
		                 SQL="SELECT fh_bt_id FROM fcfgthd where (fh_id='"& JobNum &"')"
	                    'Response.Write "LINE 998 SQL="&SQL&"<BR>"
	                    SET oRs89 = oConn89.Execute(Sql)
	                    If not oRs89.EOF then 
                            BillToID=trim(oRs89("fh_bt_id"))
                            Else
                            BillToID="9876543210"
                        End if                      
                        oRs89.Close
		                Set oRs89=Nothing

          
                      ' add exceptions for this company
                Set Recordset1e = Server.CreateObject("ADODB.Recordset")
                'Response.write "Database="&Database&"<br>"
                Recordset1e.ActiveConnection = Database
                SQL = "SELECT a.accID, a.bt_id, a.atid, a.accCharge, a.accStatus, a.accDate, a.changedby, a.accstartdate, a.accstopdate, b.bt_id, b.bt_desc "_
                & " FROM Accessorials a "_
                & " INNER JOIN fcbillto b on b.bt_id = a.bt_id "_
                & " WHERE (a.accStatus='c') and a.bt_id = " & BillToID

                '& " WHERE (a.accStatus='c') and a.bt_id = " & BillToID & " and a.accstartdate < '" & Now() & "' and a.accstopdate >= '" & Now() & "'"
                
                Recordset1e.Source = SQL
                'response.write "SQL="& SQL &"<BR>"
                Recordset1e.CursorType = 0
                Recordset1e.CursorLocation = 2
                Recordset1e.LockType = 1
                Recordset1e.Open()
                Recordset1e_numRows = 0

                	if NOT Recordset1e.EOF then
                    Do Until Recordset1e.EOF

                      Set oConn = Server.CreateObject("ADODB.Connection")
                      oConn.ConnectionTimeout = 100
                      oConn.Provider = "MSDASQL"
                      oConn.Open DATABASE
                                    
                    accTypeID=Recordset1e("atid")
                      SQL = "SELECT * FROM AccessorialType WHERE atid = '" & accTypeID & "'"
                      SET oRsN1 = oConn.Execute(SQL)
                      if NOT oRsN1.EOF then
                        accDescr = oRsN1("atDescr")
                        BillCode = oRsN1("atBillCode")
                      else
                        accDescr = "UNKNOWN"
                      end if
                      set oRsN1 = Nothing
                      Set oConn=Nothing
                          
                    accCharge = Recordset1e("accCharge")
                    accCharge2 = FormatCurrency(accCharge,2)
                    
                    'response.write accDescr & ", " & accCharge & "<br>"
                    %><form method="post" action="AddJobException.asp?t=<%=PageStatus%>&j=<%=jobnum%>&b=<%=BillToID%>&a=<%=RecordSet1e("accID")%>&c=<%=accTypeID%>&d=<%=accCharge%>&l=<%=LocationCode%>"><input type="submit" id="gobutton" value="<%=accDescr%>" /> <%=accCharge2%></form><%
                    
                    Recordset1e.MoveNext
                    Loop
                Else
                  Response.Write "NO EXCEPTIONS FOUND"
                End if
                    response.write "</td></tr>"
                Recordset1e.Close()
                Set Recordset1e = Nothing      
				

          
          
          %>									
				</td></tr>
<%
				i=i+1
							''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				Response.Write "</font>"
				If colorchanger = 1 Then
					colorchanger = 0
					color1 = "class=headerwhite"
					color2 = "class=header"
				Else
					colorchanger = 1
					color1 = "class=header"
					color2 = "class=headerwhite"	
				End If
%>


			</table>	
			<tr><td>&nbsp;</td></tr>
			<tr><td>&nbsp;</td></tr>
			<tr><td>&nbsp;</td></tr>			
	</Table>

	</td></tr>
</body>
</html>
