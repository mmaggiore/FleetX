<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<meta http-equiv="refresh" content="900" />
		<link rel="stylesheet" type="text/css" href="mainStyleSheet.css">
		<!-- #include file="../v9web/include/ifabsettings.inc" -->
		<!-- #include file="driverinfo.inc" -->	
		<title>Logisticorp Driver Message Page</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
	</head>
	<body>
	<table cellspacing="0" cellpadding="0" border="0" width="300" ID="Table1">
		<tr>
			<td class="mainpagetextboldcenter">
				Welcome&nbsp;&nbsp;<%=FirstName%>&nbsp;&nbsp;<%=LastName%><br>
			</td>
		</tr>
		<tr>
			<td class="ErrorMessageBoldCenter">
				You have messages!
			</td>			
		</tr>
		<tr><td>&nbsp;</td></tr>
		<%
			X=0
			AcknowledgeMessageID=Request.Form("AcknowledgeMessageID")
			If AcknowledgeMessageID>"" then
				Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
					RSEVENTS.Open "DriverMessageAcknowledgement", Intranet, 2, 2
					RSEVENTS.addnew	
					RSEVENTS("MessageID")=AcknowledgeMessageID		
					RSEVENTS("DriverID") = UserID
					RSEVENTS("AcknowledgeDate") = Now()
					RSEVENTS.update
					RSEVENTS.close			
				set RSEVENTS = nothing			
			End if
			Set Recordset1 = Server.CreateObject("ADODB.Recordset")
			Recordset1.ActiveConnection = Intranet
			Recordset1.Source = "SELECT DriverMessages.MessageID AS MessageID, DriverMessages.DriverMessage AS DriverMessage, DriverMessages.MessageDate AS MessageDate FROM DriverMessages"
            Recordset1.Source = Recordset1.Source&" WHERE (MessageRecipient='"&UserID&"' or MessageRecipient='-1') AND (MessageStatus='c')"
			'Response.Write "SQL="&Recordset1.Source&"<BR>"
			Recordset1.CursorType = 0
			Recordset1.CursorLocation = 2
			Recordset1.LockType = 1
			Recordset1.Open()
			Recordset1_numRows = 0

			DO WHILE NOT Recordset1.EOF 
				MessageID=Recordset1("MessageID")
				DriverMessage=Recordset1("DriverMessage")
				MessageDate=Recordset1("MessageDate")
				
				
				
				
				
			Set Recordset2 = Server.CreateObject("ADODB.Recordset")
			Recordset2.ActiveConnection = Intranet
			Recordset2.Source = "SELECT DriverID AS DriverID FROM"
            Recordset2.Source = Recordset2.Source&" DriverMessageAcknowledgement"
            Recordset2.Source = Recordset2.Source&" WHERE (DriverID='"&UserID&"') AND (MessageID='"&MessageID&"') "
			'Response.Write "SQL="&Recordset2.Source&"<BR>"
			Recordset2.CursorType = 0
			Recordset2.CursorLocation = 2
			Recordset2.LockType = 1
			Recordset2.Open()
			Recordset2_numRows = 0
			If Recordset2.EOF then
				X=1
				'Response.Redirect("DriverVehicle.asp")
				%>
				<tr>
					<td class="DriverMessage">
						<b><%=MessageDate%>-</b><%=DriverMessage%>
					</td>
				</tr>
				<form method="post">
				<tr><td><input type="submit" value="click here to acknowledge message"></td></tr>					
				<input type="hidden" name="AcknowledgeMessageID" value="<%=MessageID%>">
				</form>
				<tr><td>&nbsp;</td></tr>
				<%
			End if
			Recordset2.Close()
			Set Recordset2 = Nothing				
				
				
				
				
				
				
				
			Recordset1.MoveNext
			LOOP
			Recordset1.Close()
			Set Recordset1 = Nothing
			If X=0 then
				Response.Redirect("DriverVehicle.asp")
			End if		
		%>
	</table>
	</body>
</html>