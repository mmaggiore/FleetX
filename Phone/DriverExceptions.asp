<%@ Language=VBScript %>
<!-- 10/08/2005 Span-Renuka included this screen for CCF 2023 -->
<!-- #include file="../DedicatedFleets/include/checkstring.inc" -->
<!-- #include file="../DedicatedFleets/include/custom.inc" -->
<!-- #include file="../DedicatedFleets/include/ifabsettings.inc" -->
<!-- #include file="driverinfo.inc" -->	
<HTML>
	<HEAD>
		<!--meta http-equiv="refresh" content="240" /-->
		<link rel="stylesheet" type="text/css" href="mainStyleSheet.css">
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
		<script type="text/javascript">
		function formSubmit()
		{
		document.getElementById("thisForm").submit()
		}
		</script>		
		<%
		HAWB=Request.Form("HAWB")
		If HAWB="" then
			HAWB=Request.QueryString("HAWB")
		End if
		'If HAWB>"" then
		'REsponse.Write "HAWB=XX"&HAWB&"XX<br>"
		'End if
		'If trim(HAWB)="" then HAWB="666" end if
		LocationCode=Request.Form("LocationCode")
		FakeSubmit=Request.Form("FakeSubmit")
		If FakeSubmit="" then
			FakeSubmit=Request.QueryString("FakeSubmit")
		End if		
		PageStatus=Request.Form("PageStatus")
		txtJobNumber=Request.Form("txtJobNumber")
		Submit=Request.Form("Submit")
		BillToID=Request.Cookies("Phone")("sBT_ID")	
		Select Case BillToID
			Case "75"
				DisplayWord="BOL #"
				Email="mark.maggiore@logisticorp.us"
			Case "80"
				DisplayWord="HAWB #"
				Email="mark.maggiore@logisticorp.us;Les.Baron@Logisticorp.us"				
			Case else
				DisplayWord="HAWB #"
				Email="KWETI.Mailbox@am.kwe.com"
			End Select		
		'Response.Write "BillToID="& BillToID &"<BR>"		
		If Submit="submit" then
			ExceptionID=Request.Form("ExceptionID")
			'locationcode=Request.Form("locationcode")
			hawb=Request.Form("hawb")
			JobNumber=Request.Form("JobNumber")
			BillToID=Request.Cookies("Phone")("sBT_ID")			
			'Response.Write "GOT HERE!<BR>"
			'Response.Write "JobNumber="&JobNumber&"<BR>"
			'Response.Write "ExceptionID="&ExceptionID&"<BR>"
			'Response.Write "hawb="& hawb &"<BR>"
			'Response.Write "BillToID="& BillToID &"<BR>"
			'Response.Write "now()="&now()&"<BR>"
			FakeSubmit="fakesubmit"
			If ExceptionID>"" then
				Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
					RSEVENTS2.Open "FCJobExceptions", DATABASE, 2, 2
					RSEVENTS2.addnew
					RSEVENTS2("fh_ID")=JobNumber
					RSEVENTS2("ExceptionID")=ExceptionID									
					RSEVENTS2("Ref_Num")=hawb		
					RSEVENTS2("BillToID") = BillToID
					RSEVENTS2("ExceptionTime")=Now()		
					RSEVENTS2("Status") = "c"
					RSEVENTS2.update
					RSEVENTS2.close			
				set RSEVENTS2 = nothing	
				Set Recordset1 = Server.CreateObject("ADODB.Recordset")
				Recordset1.ActiveConnection = DATABASE
				Recordset1.Source = "SELECT ExceptionDescription FROM DriverExceptionList where (fh_bt_id='"&Request.Cookies("Phone")("sBT_ID")&"') and (Status='c') and (ExceptionID='"&ExceptionID&"')"
				Recordset1.CursorType = 0
				Recordset1.CursorLocation = 2
				Recordset1.LockType = 1
				Recordset1.Open()
				Recordset1_numRows = 0
				'Response.Write "Recordset1.Source="&Recordset1.Source&"<BR>"
				If Recordset1.eof then
					ErrorMessage="Error on Page"
				End if	
				If Not Recordset1.eof then
					ExceptionDescription=Recordset1("ExceptionDescription")
				End if	
				Recordset1.Close()
				Set Recordset1 = Nothing				
''''''''''''''''''''email notification BEGIN
					Body = "RE:&nbsp;&nbsp;" & DisplayWord & "&nbsp;&nbsp;"& hawb &"<br><br>"   & _
					"The driver has reported the following exception:<br><br>"   & _
					" "&ExceptionDescription&"<br><br>"  & _
					"If you have any questions, please do not hesitate to contact me.<br><br>"   & _
					"Thank you,<br><br>"   & _
					"Mark Maggiore<br>"  & _
					"LogistiCorp Web Developer<br>"  & _
					"mark.maggiore@LogistiCorp.us<br>"  & _ 
					"214/956-0400 xt 212<br><br>"
					Recipient=FirstName&" "&LastName

					
					'Email="mark.maggiore@logisticorp.us"
					
					'Set objMail = CreateObject("CDONTS.Newmail")
					'objMail.From = "FleetX@LogisticorpGroup.com"
					varTo = Email
					varSubject = DisplayWord & " " & hawb &" Exception"
					'objMail.MailFormat = cdoMailFormatMIME
					'objMail.BodyFormat = cdoBodyFormatHTML
					'objMail.Body = Body
					'objMail.Send
					'Set objMail = Nothing
            

                         Set iMsg = CreateObject("CDO.Message")
                        Set iConf = CreateObject("CDO.Configuration")
                        Set Flds = iConf.Fields
                        With Flds
	                        .Item("http://schemas.microsoft.com/cdo/configuration/sEndusing")				= AWS_SendUsingPort
	                        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")				= AWS_SMTPServer
	                        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl")				= AWS_SMTPUseSSL
	                        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport")			= AWS_SMTPServerPort
	                        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")		= AWS_SMTPAuthenticate
	                        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername")			= AWS_SendUserName
	                        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword")			= AWS_SendPassword
	                        .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout")	= AWS_SMTPConnectionTimeout
	                        .Update
                        End With
                        Set iMsg.Configuration = iConf

	                        iMsg.To = varTo
                            iMsg.CC = varCC
	                        iMsg.BCC = "Mark.Maggiore@Logisticorp.us"
	                        SentMail="y"
                        With iMsg
	                        Set .Configuration = iConf
	                        .From ="System.Notification@logisticorp.us"
	                        .Subject = varSubject
	                        .HTMLBody = Body
	                        .Send
                        End With 

''''''''''''''''''''email notification END				
				Response.Redirect("default.asp")				
				
				else
				ErrorMessage="You did not select an exception"
				PageStatus="loggedin"
			End if
		End if
		If FakeSubmit="fakesubmit" then
		If trim(HAWB)="" then
			Response.Redirect("default.asp")
			'Response.Write "Got here #1<br>"
		End if
			Set Recordset1 = Server.CreateObject("ADODB.Recordset")
			Recordset1.ActiveConnection = DATABASE
			Recordset1.Source = "SELECT fh_Id FROM fcfgthd INNER JOIN fcrefs ON fcfgthd.fh_id = fcrefs.rf_fh_id where (fh_bt_id='"&Request.Cookies("Phone")("sBT_ID")&"') AND (rf_ref='"& HAWB &"') AND (fh_statcode<>'9') AND (fh_statcode<>'99') AND (fh_statcode<>'98')"
			Recordset1.CursorType = 0
			Recordset1.CursorLocation = 2
			Recordset1.LockType = 1
			Recordset1.Open()
			Recordset1_numRows = 0
			'Response.Write "Recordset1.Source="&Recordset1.Source&"<BR>"
			If Recordset1.eof then
				ErrorMessage="That is not a valid " & DisplayWord
			End if			
			
			If NOT Recordset1.EOF then 
				JobNumber=Recordset1("fh_id")
			End if
			Response.Write "</font>"
			Recordset1.Close()
			Set Recordset1 = Nothing
			If ErrorMessage="" then PageStatus="loggedin" End if
		End if		
		

	
		
		%>
	</HEAD>
	<%
	'Response.Write "pagestatus="&pagestatus&"<BR>"
	if pagestatus>"" then%>
		<BODY leftMargin=0 TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<%else
		'Response.Write "THIS IS IT!!!"
		%>
		<BODY leftMargin=0 TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" OnLoad=document.thisForm.HAWB.focus()>
	<%end if%>
					<table cellpadding="0" cellspacing="0" border="0" align="left" bordercolor="red" ID="Table1">
						<tr><td align="center" colspan="9"><form method="post" action="default.asp" ID="Form5"><input type="submit" value="Return to Menu" ID="Submit7" NAME="Submit7"></form></td></tr>

			<%
			Select Case Pagestatus
				Case "loggedin"
					%>
					<form method="post" action="DriverExceptions.asp">
						<tr>
							<td align="center" class="purpleseparator" colspan="13"><b>POSSIBLE EXCEPTIONS</b></td>
						</tr>
						<%
						Set Recordset1 = Server.CreateObject("ADODB.Recordset")
						Recordset1.ActiveConnection = DATABASE
						Recordset1.Source = "SELECT ExceptionDescription, ExceptionID FROM DriverExceptionList where (fh_bt_id='"&Request.Cookies("Phone")("sBT_ID")&"') and (Status='c')"
						Recordset1.CursorType = 0
						Recordset1.CursorLocation = 2
						Recordset1.LockType = 1
						Recordset1.Open()
						Recordset1_numRows = 0
						'Response.Write "Recordset1.Source="&Recordset1.Source&"<BR>"
						If Recordset1.eof then
							ErrorMessage="There are no available suggestions"
						End if			
						
						DO WHILE NOT Recordset1.EOF 
							ExceptionDescription=Recordset1("ExceptionDescription")
							ExceptionID=Recordset1("ExceptionID")
							'Response.Write "ExceptionDescription="&ExceptionDescription&"<BR>"
						
						If X>0 then
							Response.Write "<tr><td colspan='13' bgcolor='white'><img src='images/pixel.gif' height='2' width='1' border='0'></td></tr>"
							Response.Write "<tr><td colspan='13' bgcolor='gray'><img src='images/pixel.gif' height='1' width='1' border='0'></td></tr>"
							Response.Write "<tr><td colspan='13' bgcolor='white'><img src='images/pixel.gif' height='2' width='1' border='0'></td></tr>"
							X=0
						End if
						
							%>
							<tr><td height="3"><img src="images/pixel.gif" height="3" width="1"></td></tr>
							<tr>
								<td width="20">&nbsp;</td>
								<td Class="generalcontent" width="40">
									<input type="radio" value="<%=ExceptionID%>" name="ExceptionID">
								</td>
								<td Class="generalcontent">
									<%=ExceptionDescription%>	
								</td>
							</tr>
							<tr><td height="3"><img src="images/pixel.gif" height="3" width="1"></td></tr>
							<%	
							x=x+1						
						Recordset1.Movenext
						LOOP
						Response.Write "</font>"
						Recordset1.Close()
						Set Recordset1 = Nothing						
						%>
						<tr><td colspan="3" align="center"><font color="red"><b><%=ErrorMessage%></b></font></td></tr>
						<input type="hidden" name="locationcode" value="<%=locationcode%>">
						<input type="hidden" name="hawb" value="<%=hawb%>" ID="Hidden1">
						<input type="hidden" name="JobNumber" value="<%=JobNumber%>" ID="Hidden2">
						<tr><td align="center" colspan="3"><input type="submit" name="submit" value="submit"></td></tr>
					</table>
					</form>
					<%				
				Case Else			
			%>
			<FORM ACTION="DriverExceptions.asp" method="post" name="thisForm" ID="thisForm">
					<TR> 
						<td> 
							<div class="purpleseparator"> 
								<table border="0" cellpadding="2" cellspacing="0" ID="Table4" width="300" bordercolor="blue">
									<tr> 
										<td class="mainpagetextboldright" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
									</tr>
									<tr>
										<td class='mainpagetextboldcenter' colspan="2" nowrap align="center">SCAN in <%=DisplayWord%></td>
									</tr>
									<tr>
										<td colspan='2' class='generalcontent' align="center">
											<input maxlength="20" name="HAWB" id="txtstation" type="text" size="20">
											<input maxlength='25' size='25' name='VehicleID' id='VehicleID' value='<%=VehicleID%>' type="hidden">
											<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden16">
										</td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr><td colspan="2" align="center"><input size="8" maxlength="30" type="text" value="<%=Time()%>" name="bogus" ID="Text1" onFocus="formSubmit()" readonly="readonly" class="InvisibleTextFieldPurple"></td></tr>				
			
									<%if errormessage>"" then%>
										<tr>
											<td class='generalcontenthighlight'colspan='2' align="center"><font color="red"><br><b><%=ErrorMessage%></b><br><br></font></td>
										</tr>
									<%end if%>
									<tr> 
										<td class="subheader" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
									</tr>
								</table>
							</div>
						</td>
						<!--Dummy section-->
					</TR>
					<tr><td align="center" colspan="4">&nbsp;</td></tr>					
				</TABLE>
			</FORM>	
		<%
		
		
		End select
		%>
	</BODY>
</HTML>
