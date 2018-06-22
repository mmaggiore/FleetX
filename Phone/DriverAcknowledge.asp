<%@ Language=VBScript %>

<HTML>
	<HEAD>
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<!-- #include file="FleetX.inc" -->
<!-- #include file="driverinfo.inc" -->	

<%
		FORMJOBSTATUS=TRIM(Request.Form("FORMJOBSTATUS"))
		AcknowledgeIt=Request.Form("AcknowledgeIt")
		DriverID=Request.Form("DriverID")
		LocationCode=Request.Form("LocationCode")
		Submit=Request.Form("Submit")
		PageStatus=Request.Form("PageStatus")
        MaterialType=Request.Form("MaterialType")
		PageStatus="loggedin"
		txtJobNumber=Request.Form("txtJobNumber")
		If Submit="submit" then
			If DriverID="" then
				ErrorMessage="You must provide your driver id"
			End if
		If ErrorMessage="" then PageStatus="loggedin" End if
		End if
		'Response.Write "25 userid="&userID&"<br>"
		'Response.Write "26 vehicleid="&vehicleID&"<br>"
		'Response.Write "27 unitid="&unitid&"<br>"
		'Response.Write "28 driverid="&driverid&"<br>"
		%>
	</HEAD>
	<body>
		<%
		'Response.write "DriverID="&DriverID&"<BR>"
		Select Case PageStatus
			Case "loggedin"
				If AcknowledgeIt="y" then
''''''''''''''''''''ERROR HANDLING TO PREVENT TIME FLIPPING'''''''''''''''''''''''''''''''
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				SQL = "SELECT Fh_ID, fh_user3 FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
				SQL = SQL&" WHERE (fl_un_id='"&VehicleID&"') AND (Fh_ID='"&txtJobNumber&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL))"
				''''If VehicleID=124 then
					'''''SQL = SQL&" AND ((((Fh_Status='ARV') AND (fl_st_id<>'TOPPAN')) AND (Fl_SecAcc is NULL)) "
					SQL = SQL&" AND (((Fh_Status='ARV') AND (Fl_SecAcc is NULL)) "
					'''''else
					SQL = SQL&" OR (Fh_Status='OPN')) "
				'''''End if
				'SQL = SQL&" ORDER BY Fh_Priority = '6', Fh_Priority = '9', Fh_Priority = '3', Fh_Priority = '11', Fh_Priority = '7', Fh_Priority = '8', Fh_Priority = '5', fh_id"
				
				'response.write "XXXXXXXXSQL="&SQL&"<BR>"
				'''''''''''''''''''''''''
				oRs.Open SQL, DATABASE, 1, 3
						If not oRs.EOF then
                            fh_user3=oRs("fh_user3")
							OKToChange="y"
							ELSE
							OKToChange="n"
						End if	
				oRs.Close
				Set oRs=Nothing						
''''''''''''''''''''END ERROR HANDLING'''''''''''''''''''''''''''''''				
				
					'If VehicleID<>124 then
					If OKToChange="y" then
						Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open DATABASE
								''''l_cSQL = "UPDATE fcfgthd SET fh_status = 'ACC', fh_statcode = 4 WHERE fh_id = '" & txtJobNumber & "'"
							'response.Write "PHONE_CHANGE_STATUS '" & txtJobNumber & "', '4', 'ACC', '','','',''" 
							IF FORMJOBSTATUS="ARV"   then
									oConn.Execute "PHONE_CHANGE_STATUS '" & txtJobNumber & "', '55', 'AC2','','','"& userid &"','"& vehicleID &"'" 
                                    ELSE
                            'Response.write "Line 76 PHONE_CHANGE_STATUS '" & txtJobNumber & "', '4', 'ACC', '','','',''<br>" 
							'oConn.Execute "PHONE_CHANGE_STATUS '" & txtJobNumber & "', '4', 'ACC', '"& userid &"','','',''" 
							oConn.Execute "PHONE_CHANGE_STATUS '" & txtJobNumber & "', '4', 'ACC', '"& userid &"','','',''" 
                            END IF
                            If trim(fh_user3)>"" then
						        Set oConn3 = Server.CreateObject("ADODB.Connection")
						        oConn3.ConnectionTimeout = 100
						        oConn3.Provider = "MSDASQL"
						        oConn3.Open DATABASE3
								        ''''l_cSQL = "UPDATE fcfgthd SET fh_status = 'ACC', fh_statcode = 4 WHERE fh_id = '" & txtJobNumber & "'"
							        'response.Write "PHONE_CHANGE_STATUS '" & txtJobNumber & "', '4', 'ACC', '','','',''" 
							        IF FORMJOBSTATUS="ARV"   then
									        oConn.Execute "PHONE_CHANGE_STATUS '" & fh_user3 & "', '55', 'AC2','','','"& userid &"','"& vehicleID &"'" 
                                            ELSE
                                    'Response.write "Line 76 PHONE_CHANGE_STATUS '" & txtJobNumber & "', '4', 'ACC', '','','',''<br>" 
							        'oConn3.Execute "PHONE_CHANGE_STATUS '" & fh_user3 & "', '4', 'ACC', '"& userid &"','','',''"
                                    oConn3.Execute "PHONE_CHANGE_STATUS '" & fh_user3 & "', '4', 'ACC', '"& userid &"','','',''"  
							
                                    END IF
                            End if
							'Response.Write "111txtJobNumber="&txtJobNumber&"<BR>"
							'Response.Write "111FORMJOBSTATUS="&FORMJOBSTATUS&"<BR>"
							'Response.Write "111BILLTOID="&BILLTOID&"<BR>"
							
							'response.write "UPDATE 1<BR>"
							'oConn.Execute(l_cSQL)
''''''''''''''''''''''SENDS MESSAGE TO SENDER OF ITARS
''''''''''''''''''''''FINDS EMAIL LIST
						If MaterialType = "ITAR" then
                            Set oRs = Server.CreateObject("ADODB.Recordset")
							oRs.CursorLocation = 3
							oRs.CursorType = 3
							oRs.ActiveConnection = DATABASE	
							SQL = "SELECT fcshipto.st_email AS ccAddress, fcshipto.st_id AS sf_id FROM fclegs INNER JOIN fcshipto ON fclegs.fl_sf_id = fcshipto.st_id WHERE (fclegs.fl_fh_id = '"& txtJobNumber &"')"
							
							'response.write "XXXXXXXXSQL="&SQL&"<BR>"
							'''''''''''''''''''''''''
							oRs.Open SQL, DATABASE, 1, 3
									If not oRs.EOF then
										ccAddress=oRs("ccAddress")
										AddressLocation=oRs("sf_id")
										ELSE
										'OKToChange="n"
									End if	
							oRs.Close
							Set oRs=Nothing	
							Body = "A LogistiCorp driver has just acknowledge ITAR order #"& txtJobNumber &".<br><br>Driver is now on way to pick up ITAR.<br><br>Please be at "& AddressLocation &" prepared to hand off to driver.<br><br>"& _
							"<br><br>LogistiCorp" 
							'Recipient = "mark.maggiore@logisticorp.us"
							'Set objMail = CreateObject("CDONTS.Newmail")
							'objMail.From = "FleetX@LogisticorpGroup.com"
							varTo = "itar@logisticorp.us"
							
							varcc= "mark.maggiore@logisticorp.us"

							varSubject = "ITAR Pick Up Notification"
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''							
						    Set oConn=Nothing
                        End if
					End if						
				
					
					
					
				
				End if
				
				
				
				
				
				
'------------------------------ACKNOWLEDGES ALL
			If AcknowledgeIt="all" then
				X=0
				Y=0
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				
				SQL = "SELECT Fl_SF_ID, Fh_ID, fl_sf_comment, fh_ship_dt, Fl_ST_ID, FH_Status, Fh_Priority, Fh_User3, Fh_User5 FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
				SQL = SQL&" WHERE (fl_un_id='"&VehicleID&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL))"
				SQL = SQL&" AND ((Fh_Status='OPN') OR (Fh_Status='ARV')) and fh_priority<>'6'"
				'response.write "ZZZZZZZZZZZZZZZzSQL="&SQL&"<BR>"
				oRs.Open SQL, DATABASE, 1, 3
						If not oRs.EOF then
							ELSE
						End if
						Do while not oRs.eof
						FromLocation = oRs("Fl_SF_ID")
                        'Response.write "Line 153="&FromLocation&"<BR>"
						JobNumber = oRs("Fh_ID")
						ToLocation = oRs("Fl_ST_ID")
						Fl_SF_Comment = oRs("Fl_SF_Comment")
						MaterialType = oRs("Fh_User5")
						JobStatus = oRs("fh_status")
						FORMJOBSTATUS=Trim(JobStatus)
						Priority = oRs("fh_priority")
                        fh_user3= oRs("fh_user3")
						ShipTime = oRs("fh_ship_dt")
						TimeSincePlaced=DateDiff("n",shiptime,now())
						'Response.Write "JobNumber="&JobNumber&"<BR>"
						
                        'response.write "MaterialType="&MaterialType&"***<BR>"
                        'response.write "priority="&priority&"***<BR>"

						If priority="6" or priority="P1" or priority="XP" or MaterialType="ITAR" or MaterialType="Secure Waf" or MaterialType="secret" then
							DontRedirect="y"
						End if
						
						
						
						
						'response.write "got here 177<BR>"
                        'response.write "XXXAcknowledgeIt="&AcknowledgeIt&"<BR>"
                        'response.write "XXXPriority="&Priority&"<BR>"
                        'response.write "XXXMaterialType="&MaterialType&"<BR>"
						If AcknowledgeIt="all" then
								'response.write "got here 179<BR>"
                                Set oConn = Server.CreateObject("ADODB.Connection")
								oConn.ConnectionTimeout = 100
								oConn.Provider = "MSDASQL"
								oConn.Open DATABASE
							IF FORMJOBSTATUS="ARV"  then

                                    'response.write "got here 186<BR>"
									oConn.Execute "PHONE_CHANGE_STATUS '" & JobNumber & "', '55', 'AC2','','','"& userid &"','"& vehicleID &"'" 
                                    ELSE
                                    'response.write "got here 189<BR>"								
										'oConn.Execute "PHONE_CHANGE_STATUS '" & JobNumber & "', '4', 'ACC', '','','',''" 
                                        oConn.Execute "PHONE_CHANGE_STATUS '" & JobNumber & "', '4', 'ACC', '"& userid &"','','',''" 

                                        'Response.write "Line 183 PHONE_CHANGE_STATUS '" & JobNumber & "', '4', 'ACC', '','','',''<br>" 
									END IF	
										'response.write "UPDATE 3<BR>"
										'''''l_cSQL = "UPDATE fcfgthd SET fh_status = 'ACC', fh_statcode = 4 WHERE fh_id = '" & JobNumber & "'"
									'''''oConn.Execute(l_cSQL)
								oConn.close
								Set oConn=Nothing
                                If trim(fh_user3)>"" then
						            Set oConn3 = Server.CreateObject("ADODB.Connection")
						            oConn3.ConnectionTimeout = 100
						            oConn3.Provider = "MSDASQL"
						            oConn3.Open DATABASE3
								            ''''l_cSQL = "UPDATE fcfgthd SET fh_status = 'ACC', fh_statcode = 4 WHERE fh_id = '" & txtJobNumber & "'"
							            'response.Write "PHONE_CHANGE_STATUS '" & txtJobNumber & "', '4', 'ACC', '','','',''" 
							            IF FORMJOBSTATUS="ARV"   then
									            oConn.Execute "PHONE_CHANGE_STATUS '" & fh_user3 & "', '55', 'AC2', '','','"& userid &"','"& vehicleID &"'" 
                                                ELSE
                                        'Response.write "Line 76 PHONE_CHANGE_STATUS '" & txtJobNumber & "', '4', 'ACC', '','','',''<br>" 
							            'oConn3.Execute "PHONE_CHANGE_STATUS '" & fh_user3 & "', '4', 'ACC', '','','',''" 
                                        oConn3.Execute "PHONE_CHANGE_STATUS '" & fh_user3 & "', '4', 'ACC', '"& userid &"','','',''" 
							
                                        END IF
                                End if
							'Response.Write "222txtJobNumber="&txtJobNumber&"<BR>"
							'Response.Write "222FORMJOBSTATUS="&FORMJOBSTATUS&"<BR>"
							'Response.Write "222BILLTOID="&BILLTOID&"<BR>"
						End if							
							Y=Y+1
						PreviousPriorityColor=PriorityColor
						TempJobNumber=JobNumber
						oRs.Movenext
						Loop
						oRs.Close
						Set oRs=Nothing	
						If DontRedirect<>"y"  then	
                            'Response.write "Line 200<BR>"		
							Response.Redirect("default.asp")
							'response.write "line 202 got here as well!<BR>"
						End if
				End if
'-------------------STARTS THE OTHER ORDERS IN THE PHONE
%>
<!-- #include file="LogoSection.asp" -->
					<table cellpadding="0" width="300" cellspacing="0" bordercolor="red" border="0" align="left" ID="Table5">
                        <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>
						<form method="post" action="default.asp" ID="Form2">
                        <tr><td align="center" colspan="3"><input type="submit" id="gobutton" value="Return to Menu" ID="Submit2" NAME="Submit2"></td></tr>
	                    </form>
                        <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>
                        <tr>
		                    <td class="FleetXRedSection" colspan="2" align="center">
			                    New Orders
		                    </td>
	                    </tr>
                        <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>						
						<form method="post" action="DriverAcknowledge.asp" ID="Form333">
						<tr><td valign="top" colspan="3" align="center"><input type="submit" id="gobutton" value="Acknowledge ALL" name="submit" class="<%=ButtonGrey%>" ID="Submit4"></td></tr>
						<input type="hidden" name="txtcaller" value="<%=DriverID%>" ID="Hidden1">
						<input type="hidden" name="txtstation" value="<%=ToLocation%>" ID="Hidden2">
						<input type="hidden" name="txtjobnumber" value="<%=jobnumber%>" ID="Hidden3">
						<input type="hidden" name="DriverID" value="<%=DriverID%>" ID="Hidden4">
						<input type="hidden" name="LocationCode" value="<%=LocationCode%>" ID="Hidden5">
						<input type="hidden" name="AcknowledgeIt" value="all" ID="Hidden6">
						<input type="hidden" name="PageStatus" value="loggedin" ID="Hidden7">
                        <input type="hidden" name="MaterialType" value="<%=MaterialType%>">
						</form>	
                        <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>
<%

				X=0
				Y=0
				If Request.Form("page") = "" Then
					intPage = 1	
					Else
					intPage = Request.Form("page")
				End If				
				'Response.write "Database="&Database&"<BR>"
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				SQL = "SELECT Fl_SF_ID, fl_sf_building, fl_sf_name, fl_sf_addr1, fl_sf_addr2, fl_sf_city, Fl_SF_Comment, Fh_ID, Fh_User5, fh_ship_dt, fh_bt_id, fh_user5, Fl_ST_ID, fl_st_building, fl_st_name, fl_st_addr1, fl_st_addr2, fl_st_city, Fl_St_Rta, Fl_FirstDrop, FH_Status, Fh_Priority FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
				SQL = SQL&" WHERE (Fl_un_ID='"&trim(VehicleID)&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL))"
				''''If VehicleID=124 then
					'''''SQL = SQL&" AND ((((Fh_Status='ARV') AND (fl_st_id<>'TOPPAN')) AND (Fl_SecAcc is NULL)) "
					SQL = SQL&" AND ((Fh_Status='OPN') OR (Fh_Status='ARV') )"
				'''''End if
				SQL = SQL&" ORDER BY  case fh_status when '6' then 1 when '9' then 2 when '3' then 3 when '11' then 4 when '7' then 5 when '8' then 6 else 7 end, fh_id"
				
				'response.write "Line 303 XXXXXXXXSQL="&SQL&"<BR>"
				'''''''''''''''''''''''''
				oRs.Open SQL, DATABASE, 1, 3
				
				
				
				
'RS.Open SQL, INTRANET, 1, 3
oRS.PageSize = 6
oRS.CacheSize = oRS.PageSize
intPageCount = oRS.PageCount
intRecordCount = oRS.RecordCount
If (oRS.EOF) then
	'response.write "SQL="&SQL&"<BR>"
	'response.write "got here!"
	'If vehicleid<>7 then
    'Response.write "LIne 270<BR>"
	Response.Redirect("default.asp")
	'End if
End if
If NOT (oRS.BOF AND oRS.EOF) Then

If CInt(intPage) > CInt(intPageCount) Then intPage = intPageCount
	If CInt(intPage) <= 0 Then intPage = 1
		If intRecordCount > 0 Then
			oRS.AbsolutePage = intPage
			intStart = oRS.AbsolutePosition
			If CInt(intPage) = CInt(intPageCount) Then
				intFinish = intRecordCount
			Else
				intFinish = intStart + (oRS.PageSize - 1)
			End if
		End If
	If intRecordCount > 0 Then
		For intRecord = 1 to oRS.PageSize				
				''''''''''''''''''''''''''

						'''''''''''''''''''''''''''''''
	'					If not oRs.EOF then
	'						'CloseTable="y"
	'						ELSE
	'						Response.Redirect("default.asp")
	'						'Response.Write "<tr><td colspan='3' align='center'>There are currently no unacknowledged orders.</td></tr><tr><td>&nbsp;</td></tr>"
	'					End if
	'					Do while not oRs.eof
						'''''''''''''''''''''''''''''''
						FromLocation = oRs("Fl_SF_ID")
            fh_status = oRs("fh_status") 
                        fl_sf_building = oRs("fl_sf_building")
                        fl_sf_name = oRs("fl_sf_name")
                        fl_sf_addr1 = oRs("fl_sf_addr1") 
                        fl_sf_addr2 = oRs("fl_sf_addr2")
                        fl_sf_city = oRs("fl_sf_city")
                        'Response.write "Line 301 = "&FromLocation&"<BR>"

						JobNumber = oRs("Fh_ID")
						MaterialType = oRs("Fh_User5")
						BillToID=Trim(cStr(oRs("Fh_bt_id")))
                        'Response.write "BillToID="&BillToID&"<BR>"
						ToLocation = oRs("Fl_ST_ID")

                        fl_st_building = oRs("fl_st_building")
                        fl_st_name = oRs("fl_st_name")
                        fl_st_addr1 = oRs("fl_st_addr1") 
                        fl_st_addr2 = oRs("fl_st_addr2")
                        fl_st_city = oRs("fl_st_city")

						Fl_SF_Comment = oRs("Fl_SF_Comment")
						JobStatus = oRs("fh_status")
						FORMJOBSTATUS=JobStatus
						'response.write "FORMJOBSTATUS="&FORMJOBSTATUS&"<BR>"
						Priority = oRs("fh_priority")
                        'response.write "Line Number 375 - Priority="&Priority&"<BR>"
						ShipTime = oRs("fh_ship_dt")
						'MaterialType = oRs("fh_user5")
						'Response.Write "materialtype="&MaterialType&"<BR>"
						If MaterialType="300 mm Waf" or MaterialType="Foup/Fosby" then
							MaterialSymbol="*"
							else
							MaterialSymbol=""							
						End if			
						DueTime=oRs("fl_st_rta")
						Fl_FirstDrop=oRs("Fl_FirstDrop")
						'Response.Write "fromlocation="&fromlocation&"<br>"
						If trim(fromLocation)="xx55" or trim(fromLocation)="72" then
							'Response.Write "GOT HERE????<BR>"
							If Priority="P0" OR  Priority="XP" then
								DueTime=DateAdd("n", 45, Fl_firstdrop)
								else
								DueTime=DateAdd("n", 120, Fl_firstdrop)
							End if
						End if						
						TimeSincePlaced=DateDiff("n",shiptime,now())
						
						TimeTillDue=DateDiff("n",now(),DueTime)
						'Response.Write "TimeTillDue="&TimeTillDue&"<Br>"
						If TimeTillDue<0 then
							DisplayTimeTillDue="LATE"
							Else
							HoursTillDue=TimeTillDue/60
							HoursTillDue=Int(HoursTillDue)
							MinutesTillDue=TimeTillDue-(HoursTillDue*60)
							DisplayTimeTillDue=trim(HoursTillDue&"h "&MinutesTilldue&"m")
							'Response.Write "DisplayTimeTillDue="&DisplayTimeTillDue&"<BR>"
						End if
						If AcknowledgeIt="all" and Priority<>"XP" and Priority<>"6" and Priority<>"P1" and MaterialType<>"Secure Waf" and MaterialType<>"ITAR" and MaterialType<>"secret" then
								Set oConn = Server.CreateObject("ADODB.Connection")
								oConn.ConnectionTimeout = 100
								oConn.Provider = "MSDASQL"
								oConn.Open DATABASE
							IF FORMJOBSTATUS="ARV" and (vehicleID<9) then
									oConn.Execute "PHONE_CHANGE_STATUS '" & txtJobNumber & "', '55', 'AC2','','','"& userid &"','"& vehicleID &"'" 
                                    ELSE								
								'oConn.Execute "PHONE_CHANGE_STATUS '" & JobNumber & "', '4', 'ACC', '','','',''" 
                                oConn.Execute "PHONE_CHANGE_STATUS '" & JobNumber & "', '4', 'ACC', '"& userid &"','','',''" 
							END IF			
										'''''l_cSQL = "UPDATE fcfgthd SET fh_status = 'ACC', fh_statcode = 4 WHERE fh_id = '" & JobNumber & "'"
									'''''oConn.Execute(l_cSQL)
								Set oConn=Nothing
							'Response.Write "333txtJobNumber="&txtJobNumber&"<BR>"
							'Response.Write "333FORMJOBSTATUS="&FORMJOBSTATUS&"<BR>"
							'Response.Write "333BILLTOID="&BILLTOID&"<BR>"								
						End if
                        Select Case Priority 
                            Case "6","XP" 
							PriorityColor="red"
							ButtonClass="ButtonRed"
                            Case "P1" 
							PriorityColor="blue"
							ButtonClass="ButtonBlue"
							Case Else
							ButtonClass="Button1"
							PriorityColor="black"
						End Select
						If MaterialType="Secure Waf" or MaterialType="secret" or MaterialType="ITAR" then
							PriorityColor="Orange"
						End if
						TempJobStatus=trim(JobStatus)
						Select Case JobStatus
							Case "OPN", "ARV"
								JobStatus="Open"
								ButtonText="Ack"
							Case "ACC"
								JobStatus="Acknowledged"
								ButtonText="On Board"
							Case "ONB"
								JobStatus="On Board"
								ButtonText="Close"
						End Select
							Y=Y+1
									%>
									<form method="post" action="DriverAcknowledge.asp" ID="Form1">
									<tr><td valign="top" width="20"><input type="submit" id="gobutton2" value="<%=ButtonText%>" name="submit"></td>
									<input type="hidden" name="page" value="<%=intPage%>" ID="Hidden9">
									<input type="hidden" name="txtcaller" value="<%=DriverID%>" ID="Hidden14">
									<input type="hidden" name="txtstation" value="<%=ToLocation%>" ID="Hidden15">
									<input type="hidden" name="txtjobnumber" value="<%=jobnumber%>" ID="Hidden16">
									<input type="hidden" name="DriverID" value="<%=DriverID%>" ID="Hidden17">
									<input type="hidden" name="LocationCode" value="<%=LocationCode%>" ID="Hidden31">
									<input type="hidden" name="AcknowledgeIt" value="y" ID="Hidden32">
									<input type="hidden" name="PageStatus" value="loggedin" ID="Hidden33">	
									<input type="hidden" name="MaterialType" value="<%=MaterialType%>">
                                    <input type="hidden" name="FORMJOBSTATUS" value="<%=FORMJOBSTATUS%>" ID="Hidden10">								
									<%
									If BillToID<>"26" then
										Set Recordset1 = Server.CreateObject("ADODB.Recordset")
										Recordset1.ActiveConnection = DATABASE
										Recordset1.Source = "SELECT count(rf_ref) as NumberOfLots FROM FCREFS WHERE (rf_fh_id='"&JobNumber&"') and ((ref_status<>'X') or (ref_status is NULL))"
										Recordset1.CursorType = 0
										Recordset1.CursorLocation = 2
										Recordset1.LockType = 1
										Recordset1.Open()
										Recordset1_numRows = 0
										if NOT Recordset1.EOF then
											NumberOfLots=Recordset1("NumberOfLots")
											If NumberOfLots>1 then WordLots="Items"&MaterialSymbol&"/" end if
											If NumberOfLots=1 then WordLots="Item"&MaterialSymbol&"/" end if
											If NumberOfLots=0 then WordLots="" end if
											Else
											ErrorMessage="Incorrect driver ID or password"
										End if
										NumberOfLots="/"&MaterialSymbol&NumberOfLots
										Recordset1.Close()
										Set Recordset1 = Nothing
										Priority="/"&Priority
										Else
										Priority=""
									End if	
										Set Recordset1 = Server.CreateObject("ADODB.Recordset")
										Recordset1.ActiveConnection = DATABASE
										Recordset1.Source = "SELECT NumberOfPieces, rf_box FROM FCREFS WHERE (rf_fh_id='"&JobNumber&"') and ((ref_status<>'X') or (ref_status is NULL))"
										Recordset1.CursorType = 0
										Recordset1.CursorLocation = 2
										Recordset1.LockType = 1
										Recordset1.Open()
										Recordset1_numRows = 0
										if NOT Recordset1.EOF then
                                            NumberOfPieces=Recordset1("NumberOfPieces")
                                            rf_box=Recordset1("rf_box")
											Else
											ErrorMessage="Incorrect driver ID or password"
										End if
										Recordset1.Close()
										Set Recordset1 = Nothing
							''''''''''''''''''''''''''''''''
							'Response.Write "VehicleID="&VehicleID&"<BR>"
							'Response.Write "ToLocation="&ToLocation&"<BR>"
							'Response.Write "FromLocation="&FromLocation&"<BR>"
							'Response.Write "JobStatus="&JobStatus&"<BR>"
							'>>>>>>>>>>>>>>>>>>START PO HUB FROM NEVER CHANGE - ALWAYS REPLACE WITH NON-PO HUB
							'Response.Write "JobStatus="&JobStatus&"<BR>"
							DisplayToLocation=trim(ToLocation)
							DisplayFromLocation=trim(FromLocation)
							If Trim(ToLocation)="80" then
								DisplayToLocation="LSP Warehouse"
							End if
							If Trim(FromLocation)="80" then
								DisplayFromLocation="LSP Warehouse"
							End if							
              if fh_status = "ARV" and VehicleID=912780 then
                DisplayFromLocation = "SRHUB"
               end if           
							
							Select Case DisplayToLocation
								Case "D7"
									DisplayToLocation="D1"
								Case "P1"
									DisplayToLocation="D1"
							End Select
                            'Response.write "Priority="&Priority&"<BR>"
                            If trim(Priority)="/6" then
                                PriorityColor="red"
                            End if
							'>>>>>>>>>>>>>>>>>>END-PO HUB FROM NEVER CHANGE - ALWAYS REPLACE WITH NON-PO HUB							
							'Response.Write "<td valign='top' nowrap>&nbsp;&nbsp;<font color='"&PriorityColor&"'><a href='DriverTracking.asp?JobNumber="&JobNumber&"&fh_bt_id="&fh_bt_id&"'><font color='"&PriorityColor&"'>"&Right(JobNumber,5)&"</font></a>"&Priority&NumberOfLots&" "&WordLots&"</font>"
							Response.Write "<td valign='top' nowrap>&nbsp;&nbsp;<font color='"&PriorityColor&"'><a href='DriverTracking.asp?JobNumber="&JobNumber&"&fh_bt_id="&fh_bt_id&"'><font color='"&PriorityColor&"'>"&Right(JobNumber,5)&"</font></a>/"
                            Response.Write "<font color='"&PriorityColor&"'>"&DisplayTimeTillDue&"</font><br>"
                            If BillToID<>"91" then
                                'Response.write "Line 534-PriorityColor="&PriorityColor&"<BR>"
 							    Response.Write "<font color='"&PriorityColor&"'><B>FROM:&nbsp;&nbsp;</B>"
                                If trim(fl_sf_name)>"" then
                                    Response.write fl_sf_name&"<BR>"
                                End if
                                If trim(fl_sf_building)>"" then
                                    Response.write fl_sf_building&"<BR>"
                                End if
                                If trim(fl_sf_addr1)>"" then
                                    Response.write fl_sf_addr1&"<BR>"
                                End if 
                                If trim(fl_sf_addr2)>"" then
                                    Response.write fl_sf_addr2&"<BR>"
                                End if
                                If trim(fl_sf_city)>"" then
                                    Response.write fl_sf_city&"<BR>"
                                End if
                                
                                'Response.write "<br>"
                                Response.write "<B>TO:&nbsp;&nbsp;</B>"
                                If trim(fl_st_name)>"" then
                                    Response.write fl_st_name&"<BR>"
                                End if
                                If trim(fl_st_building)>"" then
                                    Response.write fl_st_building&"<BR>"
                                End if
                                If trim(fl_st_addr1)>"" then
                                    Response.write fl_st_addr1&"<BR>"
                                End if
                                If trim(fl_st_addr2)>"" then
                                    Response.write fl_st_addr2&"<BR>"
                                End if
                                If trim(fl_st_city)>"" then
                                    Response.write fl_st_city&"<BR>"
                                End if
                                Response.write "</font></td></tr>"                               
                          
                            Else							
                                Response.Write "<font color='"&PriorityColor&"'>&nbsp;&nbsp;"&DisplayFromLocation&" - "&DisplayToLocation&"</font></td></tr>"
                            End if
                            Response.Write "<tr><td>&nbsp;</td><td colspan='2'><b><font color='"&PriorityColor&"'>"& NumberOfPieces &" "&rf_box&"</font></b></td></tr>"
							If trim(Fl_SF_Comment)>"" then
								Response.Write "<tr><td>&nbsp;</td><td colspan='2'>***"&Fl_SF_Comment&"</td></tr>"
							End if
							Response.Write "</form>"
							Response.Write "<tr><td colspan='3'><hr width='100%'></td></tr>"
						PreviousPriorityColor=PriorityColor
						TempJobNumber=JobNumber
						
						
						
						
						
						'''''''''''''''''''''''''''''''''
'						oRs.Movenext
'						Loop
'						oRs.Close
oRS.MoveNext
If colorchanger = 1 Then
	colorchanger = 0
	color1 = "class=headerwhite"
	color2 = "class=header"
Else
	colorchanger = 1
	color1 = "class=header"
	color2 = "class=headerwhite"	
End If
If oRS.EOF Then Exit for
	Next
	End if
	End if
	'End if
						''''''''''''''''''''''''''''''''''
						If CloseTable="y" then
							%>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr></table>					
							<%
							CloseTable="n"
						End if					
					Case else
						Response.Write "Error #1465<br>"
				End Select
			%>
							<tr>
								<td colspan="2">
								<table ID="Table1" width="300" align="center">
				<td width="50%" align="left" valign="top"><font face="Verdana, arial" size="1">
					<%If cInt(intPage) > 1 Then%>
						<form method="post" ID="Form3">
						<input type="submit" name="submit" value="<<Previous" ID="Submit1">
						<input type="hidden" name="page" value="<%=intPage - 1%>" ID="Hidden8"></form>
						</form>
						<!--
						<a href="?orderby=<%=orderBy%>&page=<%=intPage - 1%>&SearchFor=<%=SearchFor%>&SearchWhere=<%=SearchWhere%>&SearchVariable=<%=SearchVariable%>"><< <b>Prev</b></a>
						-->
						<%
						else
						Response.write "&nbsp;"
					End If%>
					</font>
				</td>
				<td width="50%" align="right" valign="top"><font face="Verdana, arial" size="1" >
					<%If cInt(intPage) < cInt(intPageCount) Then%>
						<form method="post" ID="Form4">
						<input type="submit" name="submit" value="Next>>" ID="Submit5">
						<input type="hidden" name="page" value="<%=intPage + 1%>" ID="Hidden11"></form>
						</form>
						<!--
						<a href="?orderby=<%=orderBy%>&page=<%=intPage + 1%>&SearchFor=<%=SearchFor%>&SearchWhere=<%=SearchWhere%>&SearchVariable=<%=SearchVariable%>"><b>Next</b> >></a>
						-->
						<%
						else
						Response.write "&nbsp;"
					End If%>
					</font>
				</td>			</table>				
								</td>
							</tr></table>		

			
	</BODY>
</HTML>
