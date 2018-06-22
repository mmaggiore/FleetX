<%@ Language=VBScript %>
<!-- #include file="FleetX.inc" -->
<!-- #include file="driverinfo.inc" -->
<html>
	<head>

<%
'Option Explicit
'Dim user_agent, mobile_browser, Regex, match, mobile_agents, mobile_ua, i, size
 
user_agent = Request.ServerVariables("HTTP_USER_AGENT")
 
mobile_browser = 0
 
Set Regex = New RegExp
With Regex
   .Pattern = "(up.browser|up.link|mmp|symbian|smartphone|midp|wap|phone|windows ce|pda|mobile|mini|palm)"
   .IgnoreCase = True
   .Global = True
End With
 
match = Regex.Test(user_agent)
 
If match Then mobile_browser = mobile_browser+1
 
If InStr(Request.ServerVariables("HTTP_ACCEPT"), "application/vnd.wap.xhtml+xml") Or Not IsEmpty(Request.ServerVariables("HTTP_X_PROFILE")) Or Not IsEmpty(Request.ServerVariables("HTTP_PROFILE")) Then
   mobile_browser = mobile_browser+1
end If
 
mobile_agents = Array("w3c ", "acs-", "alav", "alca", "amoi", "audi", "avan", "benq", "bird", "blac", "blaz", "brew", "cell", "cldc", "cmd-", "dang", "doco", "eric", "hipt", "inno", "ipaq", "java", "jigs", "kddi", "keji", "leno", "lg-c", "lg-d", "lg-g", "lge-", "maui", "maxo", "midp", "mits", "mmef", "mobi", "mot-", "moto", "mwbp", "nec-", "newt", "noki", "oper", "palm", "pana", "pant", "phil", "play", "port", "prox", "qwap", "sage", "sams", "sany", "sch-", "sec-", "send", "seri", "sgh-", "shar", "sie-", "siem", "smal", "smar", "sony", "sph-", "symb", "t-mo", "teli", "tim-", "tosh", "tsm-", "upg1", "upsi", "vk-v", "voda", "wap-", "wapa", "wapi", "wapp", "wapr", "webc", "winw", "winw", "xda", "xda-")
size = Ubound(mobile_agents)
mobile_ua = LCase(Left(user_agent, 4))
 
For i=0 To size
   If mobile_agents(i) = mobile_ua Then
      mobile_browser = mobile_browser+1
      Exit For
   End If
Next
 
 
If mobile_browser>0 Then
   'Response.Write("Mobile!")
   Android="y"
Else
   'Response.Write("Not mobile!")
   Android="n"
End If
 
%>




		<!--meta http-equiv="refresh" content="100"-->
		<link rel="stylesheet" type="text/css" href="mainStyleSheet.css">
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
		<script type="text/javascript">
		function formSubmit()
		{
		document.getElementById("thisForm").submit()
        }
        function disablecopypaste() {

            //alert("Function Not allowed");
            return false;

        }


		</script>		
		<%
        'Response.write "Android="&Android&"<BR>"
        ''''''''''TO DELETE COOKIES - UNCOMMENT BELOW''''''''''''''
        'Response.Cookies("LegalComputer").expires=#1/1/2020# 
        'Response.Cookies("LegalComputer")("ComputerID")=""
		LocationCode=Request.Form("LocationCode")
        'response.write "+++++LOCATION CODE="&LocationCode&"+++++<BR>"
		FakeSubmit=Request.Form("FakeSubmit")
		varFromLocations=" or fl_sf_id='72' "
		OtherBillToID=Request.Cookies("FleetXPhone")("sBT_ID")	
		fh_bt_id=Request.Cookies("FleetXPhone")("sBT_ID")
        ErrorMessage=Request.form("ErrorMessage")
        'response.write "FakeSubmit="&FakeSubmit&"<BR>"
        'If FakeSubmit<>"v" then
            ''response.write "NOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO!!!<BR>"
            IDCookies=Request.Cookies("LegalComputer")("ComputerID")
        'End if
        TwoFabs=Request.Form("twofabs")
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        PhoneUser=Request.ServerVariables("HTTP_UA_OS") 
        ''response.write "test="&phoneuser&"***"
        'If userid="1" then
            tempVar=Request.ServerVariables("HTTP_USER_AGENT")
            If tempVar="Mozilla/5.0 (Linux; U; Android 4.1.2; en-us; TC55 Build/JZO54K) AppleWebKit/534.30 (KHTML, like Gecko) Version/4.0 Mobile Safari/534.30" then
                OurPhone="y"
                else
                OurPhone="n"
            End if
        'End if
        
        If (trim(Android)="n" and trim(IDCookies)="") and OurPhone="n" then
            Response.redirect("SetCookie.asp")
				'Response.write "<font color='red'>DRIVERS:  If you are seeing this message right now,<br>please note on which specific phone/computer<br>it occurred, and give this information to your<br>supervisor.</font>"
                '''Body = "The users OS is<BR><BR>"& PhoneUser &"<br><br>IDCookies:<br><BR>"& IDCookies &"<br><br>Phone Cookies:"& fh_bt_id & "<br><br>Name: "& firstname & " " & LastName & "."
				'Recipient=FirstName&" "&LastName
				'''SentToEmail="mark.maggiore@logisticorp.us"
				'Email="KWETI.Mailbox@am.kwe.com"
				'Email="mark@maggiore.net"
				'''Set objMail = CreateObject("CDONTS.Newmail")
				'''objMail.From = "System.Notification@logisticorp.us"
				'''objMail.To = SentToEmail
				'objMail.cc = "4692269939@tmomail.net;htmlmale@yahoo.com"
				'objMail.cc = "4692269939@tmomail.net;pcurrin@ti.com"
				'''objMail.Subject = "User OS"
				'''objMail.MailFormat = cdoMailFormatMIME
				'''objMail.BodyFormat = cdoBodyFormatHTML
				'''objMail.Body = Body
				'''objMail.Send
				'''Set objMail = Nothing
        End if

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'response.write "***IDCookies="&IDCookies&"***<BR>"
        'response.write "***unitID="&unitID&"***<BR>"
        If Trim(IDCookies)>"" and trim(LocationCode)="" then
            'response.write "Hello, I REALLY, REALLY got here!!!!<BR>"
			Set Recordset1 = Server.CreateObject("ADODB.Recordset")
			SQL_99="SELECT st_id, sb_bt_id, st_alias FROM fcshipto INNER JOIN fcshipbt ON fcshipto.st_id = fcshipbt.sb_st_id WHERE (st_pkey='"&IDCookies&"') and (st_alias<>'T162313')"
			'response.write "MMMMSQL_99="&SQL_99&"<BR>"
			Recordset1.ActiveConnection = DATABASE
			Recordset1.Source = SQL_99
			Recordset1.CursorType = 0
			Recordset1.CursorLocation = 2
			Recordset1.LockType = 1
			Recordset1.Open()
			Recordset1_numRows = 0
			if NOT Recordset1.EOF then
                temp_st_id=Recordset1("st_id")
				AliasCode=Recordset1("st_alias")
                FakeSubmit="fakesubmit"
                ''response.write "UnitID="&UnitID&"<BR>"
                ''response.write "temp_st_id="&temp_st_id&"<BR>"
                If (UCASE(trim(UnitID))="SRB" or UCASE(trim(UnitID))="SRV") and trim(temp_st_id)="DM6WIPEDOWN" then
                    ''response.write "Got here 1<br>"
                    AliasCode="D6W3"
                End if

                If (UCASE(trim(UnitID))="NTRETICLE") and trim(temp_st_id)="DM6WIPEDOWN" then
                    ''response.write "Got here 1<br>"
                    AliasCode="57456233"
                End if
                If (UCASE(trim(UnitID))="SRVRFAB" OR UCASE(trim(UnitID))="303551" OR UCASE(trim(UnitID))="303552" OR UCASE(trim(UnitID))="303553" OR UCASE(trim(UnitID))="303554") and trim(temp_st_id)="DM6RT" then
                    ''response.write "Got here 2<br>"
                    AliasCode="54785833"
                End if





                If trim(UnitID)="NTRETICLE" and trim(temp_st_id)="RFAB-W" then
                    ''response.write "Got here 1<br>"
                    AliasCode="R6277642"
                End if
                If trim(UnitID)<>"NTRETICLE" and trim(temp_st_id)="RFAB-R" then
                    ''response.write "Got here 2<br>"
                    AliasCode="R22675615"
                End if
                ''response.write "Got here #1<br>"
				Else
                ''response.write "Got here #2<br>"
                'FakeSubmit=""
                AliasCode=""
			End if
			Recordset1.Close()
			Set Recordset1 = Nothing
        End if
		If Request.Form("page") = "" Then
			intPage = 1	
			Else
			intPage = Request.Form("page")
		End If	
		If Request.Form("page2") = "" Then
			intPage2 = 1	
			Else
			intPage2 = Request.Form("page2")
		End If			
		AcknowledgeIt=Request.Form("AcknowledgeIt")
        If trim(AliasCode)="" then
		    AliasCode=Request.Form("AliasCode")
        End if
		If trim(AliasCode)="" then
			AliasCode=Request.QueryString("AliasCode")
		End if
		If AliasCode>"" then Response.Cookies("FleetXPhone")("AliasCode")=AliasCode end if
		If aliasCode="" then aliasCode=Request.Cookies("FleetXPhone")("AliasCode") end if

		If FakeSubmit="" then
			FakeSubmit=Request.QueryString("FakeSubmit")
		End if
		If FakeSubmit>"" then
			Response.Cookies("FleetXPhone")("FakeSubmit")=FakeSubmit
		End if
		If FakeSubmit="" then
			FakeSubmit=Request.Cookies("FleetXPhone")("FakeSubmit")
		End if
		PageStatus=Request.Form("PageStatus")

		txtJobNumber=Request.Form("txtJobNumber")
        'response.write "<font color='red'>XXXfakesubmit="&fakesubmit&"</font><BR>"
        If FakeSubmit="n" then
            'response.write "locationcode="&locationcode&"***<BR>"
            'response.write "twofabs="&twofabs&"***<BR>"
            If (locationcode="SCTICS" or locationcode="SCTQC")  then
                'response.write "helloooo!!!!!<BR>"
                'ErrorMessage="twofabs"
                'PageStatus="twofabs_SCTICS_SCTQC"
                If trim(LocationCode)>"" then
                    varToLocations="fl_st_id='"&LocationCode&"'"
                    varFromLocations="fl_sf_id='"&LocationCode&"'"
                End if
            End if
        End if
		If FakeSubmit="fakesubmit" then
        'response.write "GOT HERE????<BR>"
			Set Recordset1 = Server.CreateObject("ADODB.Recordset")
			Recordset1.ActiveConnection = DATABASE
			Recordset1.Source = "SELECT st_id, CompanyName FROM PreExistingCompanies WHERE (st_alias='"&AliasCode&"')"
			
            'response.write "<font color='blue'>Line 231 XXXRecordset1.Source="&Recordset1.Source&"</font><BR>"
			
            Recordset1.CursorType = 0
			Recordset1.CursorLocation = 2
			Recordset1.LockType = 1
			Recordset1.Open()
			Recordset1_numRows = 0
			If Recordset1.eof then
				ErrorMessage="That is not a valid location"
                Body = "A driver has scanned in or entered a bogus location code!<br><br>Name: "& firstname & " " & LastName & ".<BR><BR>Bogus Code: " & AliasCode
				'Recipient=FirstName&" "&LastName
				SentToEmail="mark.maggiore@logisticorp.us"
				'Email="KWETI.Mailbox@am.kwe.com"
				'Email="mark@maggiore.net"
				Set objMail = CreateObject("CDONTS.Newmail")
				objMail.From = "FleetX@LogisticorpGroup.com"
				objMail.To = SentToEmail
				'objMail.cc = "4692269939@tmomail.net;htmlmale@yahoo.com"
				'objMail.cc = "4692269939@tmomail.net;pcurrin@ti.com"
				objMail.Subject = "FleetX Bogus Location Code"
				objMail.MailFormat = cdoMailFormatMIME
				objMail.BodyFormat = cdoBodyFormatHTML
				objMail.Body = Body
				objMail.Send
				Set objMail = Nothing

			End if			
			Do While NOT Recordset1.EOF 
			'''''''''''''CHANGED to LOOP so locations with the same alias code would work correctly''''''''''''
				SetArrivalTime="y"
				m=m+1
				LocationCode=Recordset1("st_id")
				'response.write "<font color='green'>LocationCode="&LocationCode&"</font></br>"
                '''''BillToID=Trim(cStr(Recordset1("sb_bt_id")))
				varToLocations=varToLocations&" or fl_st_id='"&trim(LocationCode)&"' "
				varFromLocations=varFromLocations&" or fl_sf_id='"&trim(LocationCode)&"' "
				If OtherBillToID="80" then
					BillToID="80"
				End if
				Recordset1.Movenext
				Loop
					Response.Write "</font>"
			Recordset1.Close()
            'Response.write "locationcode="&LocationCode&"<BR>"
			Set Recordset1 = Nothing
			LengthvarToLocations=len(varToLocations)
			LengthvarFromLocations=len(varFromLocations)
			'Response.Write "varToLocations="&varToLocations&"<BR>"
			'Response.Write "varFromLocations="&varFromLocations&"<BR>"
			'Response.Write "LengthvarToLocations="&LengthvarToLocations&"<BR>"
			'Response.Write "LengthvarFromLocations="&LengthvarFromLocations&"<BR>"
			If m>0 then
			    varToLocations="("&Right(varToLocations, (int(LengthvarToLocations)-3))&")"	
			    varFromLocations="("&Right(varFromLocations, (int(LengthvarFromLocations)-3))&")"	
			End if	
			
			AliasCode=UCASE(ALIASCODE)
			LocationCode=Trim(UCASE(LOCATIONCODE))
            'response.write "TwoFabs="&TwoFabs&"<BR>"
            'response.write "<font color='red'>XXXXXLocationCode="&LocationCode&"XXXXX</font><BR>"
            If (locationcode="SCTICS" or locationcode="SCTQC") and twofabs<>"y" then
                'response.write "hello, helloooo!!!!!<BR>"
                ErrorMessage="twofabs"
                PageStatus="twofabs_SCTICS_SCTQC"
                varToLocations="fl_sf_id='"&LocationCode&"' or fl_st_id='"&LocationCode&"'"
            End if

			DisplayLocationCode=LocationCode
            'response.write "DisplayLocationCode="&DisplayLocationcode&"<BR>"
		If ErrorMessage="" then PageStatus="loggedin" End if
		End if
        DisplayLocationCode=LocationCode
		If PageStatus>"" then 
			Response.Cookies("FleetXPhone")("PageStatus")=PageStatus 
		end if
		If PageStatus="" then 
			PageStatus=Request.Cookies("FleetXPhone")("PageStatus") 
		end if
		%>
	</HEAD>
	<%if pagestatus>"" then%>
		<BODY leftMargin=0 TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<%else%>
		<BODY leftMargin=0 TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" OnLoad="document.thisForm.AliasCode.focus()" >
       
	<%end if%>
<!-- #include file="LogoSection.asp" -->
		<%
		'response.write "PageStatus="&PageStatus&"<BR>"
        'response.write "IDCookies="&IDCookies&"<BR>"
		Select Case PageStatus
            Case "twofabs_SCTICS_SCTQC"
                ''response.write "I'M HERE AGAIN!!!<BR>"
                %>
				<TABLE WIDTH="300" cellpadding="0" cellspacing="5" ID="Table5" align="left" border="0" bordercolor="red">
					<tr><td align="center" colspan="3">
                    <%If trim(IDCookies)<>"14141414" then %>
                        <form method="post" action="default.asp" ID="Form13">
                    <%else %>
                        <form method="post" action="driverifabphoneemulator.asp" ID="Form6">
                    <%end if %>
                    <input type="submit" value="Return to Menu" ID="gobutton" NAME="Submit1">

                                <input type="hidden" value="" name="locationcode" />
                                <input type="hidden" name="twofabs" value="" />
                                <input type="hidden" name="pagestatus" value="v" />
                                <input type="hidden" name="fakesubmit" value="v" />

                    </form></td></tr>			
				    <tr><td>&nbsp;</td></tr>
                    <tr>
                        <td colspan="3" align="center">
                            <form method="post">
                                <input type="submit" value="SCTICS" name="submit" class="gobutton2" />
                                <input type="hidden" value="SCTICS" name="locationcode" />
                                <input type="hidden" name="twofabs" value="y" />
                                <input type="hidden" name="pagestatus" value="loggedin" />
                                <input type="hidden" name="fakesubmit" value="n" />
                            </form>
                        </td>
                    </tr>
				    <tr><td>&nbsp;</td></tr>
                    <tr>
                        <td colspan="3" align="center">
                            <form method="post">
                                <input type="submit" value="SCTQC" name="submit" class="gobutton2" />
                                <input type="hidden" value="SCTQC" name="locationcode" />
                                <input type="hidden" name="twofabs" value="y" />
                                <input type="hidden" name="pagestatus" value="loggedin" />
                                <input type="hidden" name="fakesubmit" value="n" />
                            </form>
                        </td>
                    </tr>
                </table>
				<br clear="all">
                <%
                
			Case "loggedin"
            'response.write "GOT HERE YO! Yo!<BR>"
'-------------------STARTS THE DROP OFF	
                varToLocations=Trim(varToLocations)			
				'Response.write "varToLocations="&varToLocations&"<BR>"
                Select Case VarToLocations
                    Case "( fl_st_id='RFAB-W' )", "( fl_st_id='RFAB-R' )"
                        varToLocations="(( fl_st_id='RFAB-W' ) OR ( fl_st_id='RFAB-R' ))"
                    Case "( fl_st_id='SCTICS' )", "( fl_st_id='SCTQC' )"
                        'Response.write "GOT HERE!!!<BR>"
                        If trim(Android)="n" then
                        'If whatever="whatever" then
                            'varToLocations="( fl_st_id='SCTICS' or fl_st_id='SCTQC')"
                            varToLocations="( fl_st_id='"&LocationCode&"' or fl_sf_id='"&LocationCode&"')"
                            else

                            End if
                    Case "( fl_st_id='EBHUB' )"
                        varToLocations="((fl_sf_id<>'EBHUB') AND ((Fl_st_ID='D6W3') OR (Fl_st_ID='D6N2') OR (Fl_st_ID='D6N1') OR (Fl_st_ID='DM4M') OR (Fl_st_ID='DM5M') OR (Fl_st_ID='DPI2') OR (Fl_st_ID='DPI3') OR (Fl_st_ID='ESTK') OR (Fl_st_ID='DM5Q') OR (Fl_st_ID='DM6Q')))"
                    Case Else

                End Select

                'If trim(varToLocations)="( fl_st_id='RFAB-W' )" or trim(varToLocations)="( fl_st_id='RFAB-R' )" then 
                '    varToLocations="(( fl_st_id='RFAB-W' ) OR ( fl_st_id='RFAB-R' ))"
                'End if
				X=0
				Y=0
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				
				SQL = "SELECT Fl_SF_ID, Fh_ID, fh_User5, Fl_ST_ID, fl_sf_comment, FH_Status, Fh_Priority FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
				SQL = SQL&" WHERE (fh_ship_dt>'"&now()-30&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND (fl_un_id='"&VehicleID&"') and ("
				SQL = SQL&varToLocations
				SQL = SQL&") AND ((fh_status='ONB') OR (fh_status='DPV'))"
				SQL = SQL&" AND (fl_sf_id<>'HFABQC')"
				SQL = SQL&" ORDER BY fh_priority, fh_id"
				'response.write "SQL="&SQL&"<BR>"
				oRs.Open SQL, DATABASE, 1, 3
				If trim(DisplayLocationCode)="55" then DisplayLocationCode="CPGP" end if
				If trim(DisplayLocationCode)="48" then DisplayLocationCode="KWEO" end if
				'Response.Write "****SQL="&SQL&"<BR>"
				%>
					<table width="300" cellpadding="0" cellspacing="0" border="0" bordercolor="green" align="left" ID="Table1">
						<tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>
                        <%
                        'response.write "tttttLocationCode="&LocationCode&"<BR>"
                        Select Case LocationCode
                            Case "SCTQC", "SCTICS"
                                %>
                                <form method="post" action="DriverIfabPhoneEmulator.asp" ID="Form12">
                                <input type="hidden" value="SCTICS" name="locationcode" />
                                <input type="hidden" name="twofabs" value="y" />
                                <input type="hidden" name="pagestatus" value="twofabs_SCTICS_SCTQC" />
                                <input type="hidden" name="errormessage" value="twofabs" />
                                <input type="hidden" name="fakesubmit" value="n" />

                        <%
                            Case else
                                %>
                                <form method="post" action="default.asp" ID="Form7">
                        <%
                        End Select

                        'If trim(Android)<>"n" then
                        %>
                        <tr><td align="center" colspan="3">
                        <input type="submit" value="Return to Menu" ID="gobutton" NAME="Submit1">
                        <%'else %>

                        <%'end if %>
                        </td></tr>
                        </form>
                        <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>
						<tr>
							<td class="mainpagetextbold" colspan="3" align="center">
								Last update: <%=Time()%>
							</td>
						</tr>						
                        <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>
                        <tr>
		                    <td class="FleetXRedSection" colspan="3" align="center">
			                    DROP OFFS AT <%=uCase(DisplayLocationCode)%>
		                    </td>
	                    </tr>
                        <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>
				<%
				If not oRs.EOF then
					%>
						<tr>
							<td align="center">&nbsp;</td>						
							<td align="left" nowrap><b>&nbsp;&nbsp;&nbsp;&nbsp;Job #</b></td>
							<td align="center" nowrap>
							<%If BillToID<>"26" then%>
								<%If BillToID="80" then%>
									<b>Items</b>
									<%else%>
									<b>Lots</b>
									<%
								End if
							End if%>
							</td>
							</tr>
						<%
						CloseTable="y"
						ELSE
						Response.Write "<tr><td colspan='4' align='center'>No orders to drop off here.</td></tr><tr><td>&nbsp;</td></tr>"
				End if
				oRS.PageSize = 4
				oRS.CacheSize = oRS.PageSize
				intPageCount = oRS.PageCount
				intRecordCount = oRS.RecordCount
				If (oRS.EOF) then
					sendback1="y"
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
				
				
				'''''''''''''''''''''''''''''''''''''''''''''''''''''''
					FromLocation = oRs("Fl_SF_ID")
					JobNumber = oRs("Fh_ID")
					MaterialType = oRs("Fh_User5")
					ToLocation = oRs("Fl_ST_ID")
					fl_sf_comment = oRs("fl_sf_comment")
					JobStatus = oRs("fh_status")
					Priority = oRs("fh_priority")
                       ' Response.write "materialtype=***"&materialtype&"***<BR>"
						If MaterialType="300 mm Waf" or MaterialType="Foup/Fosby" then
							MaterialSymbol="*"
							else
							MaterialSymbol=""							
						End if
					If Priority="P0" or Priority="XP" then 
						PriorityColor="red"
						ButtonClass="ButtonRed"
                        else
					    If Priority="P1" then 
						    PriorityColor="blue"
						    ButtonClass="ButtonBlue"
						Else
						ButtonClass="Button1"
						PriorityColor="black"
                        End if
					End if
					If MaterialType="Secure Waf" OR MaterialType="secret" OR MaterialType="ITAR" then
						PriorityColor="Orange"
					End if					
					Select Case JobStatus
						Case "OPN"
							JobStatus="Open"
							ButtonText="Acknowledge"
						Case "ACC"
							JobStatus="Acknowledged"
							ButtonText="ONB"
						Case "ONB", "DPV"
							JobStatus="ONB"
							ButtonText="CLS"
						Case "PUO"
							JobStatus="Paper on Board"
							ButtonText="CLS"							
					End Select
					'FromLocation = oRs("Fl_SF_ID")
					If JobNumber<>TempJobNumber then
						If X>0 or X=0 then
							if trim(fh_bt_id)<>"26" then
							
								Set Recordset1 = Server.CreateObject("ADODB.Recordset")
								Recordset1.ActiveConnection = DATABASE
								SQL_111="SELECT count(rf_ref) as NumberOfLots FROM FCREFS WHERE (rf_fh_id='"&JobNumber&"') AND ((ref_status<>'X') or (ref_status is NULL))"
								Recordset1.Source = SQL_111
								Recordset1.CursorType = 0
								Recordset1.CursorLocation = 2
								Recordset1.LockType = 1
								Recordset1.Open()
								Recordset1_numRows = 0
								if NOT Recordset1.EOF then
									NumberOfLots=Recordset1("NumberOfLots")
									'Response.Write NumberOfLots
									If NumberOfLots>1 then WordLots="Lots" end if
									If NumberOfLots=1 then WordLots="Lot" end if
									If NumberOfLots=0 then WordLots="" end if
									Else
									Response.Write "&nbsp;"
									ErrorMessage="Incorrect driver ID or password"
								End if
								Recordset1.Close()
								Set Recordset1 = Nothing					
							End if							
							'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
							
							Response.Write "</font></td></tr>"
							Response.Write "<tr><td colspan='4' bgcolor='white'><img src='images/pixel.gif' height='2' width='1' border='0'></td></tr>"
							Response.Write "<tr><td colspan='4' bgcolor='gray'><img src='images/pixel.gif' height='1' width='1' border='0'></td></tr>"
							If trim(fl_sf_comment)>"" then
								Response.Write "<tr><td colspan='3'>***"&fl_sf_comment&"</td></tr>"
							end if							
							Response.Write "<tr><td colspan='4' bgcolor='white'><img src='images/pixel.gif' height='2' width='1' border='0'></td></tr>"							
							X=0
							
						End if
						Y=Y+1
						If Priority="P0" or Priority="XP" then
							ButtonClass="ButtonRed"
                            else
						If Priority="P1" then
							ButtonClass="ButtonBlue"
							else
							ButtonClass="Button1"
						End if
                        End if
						'Response.Write "got here #1<BR>"
						whatever=Request.Cookies("FleetXPhone")("sBT_ID")
						'Response.Write "whatever="&whatever&"<BR>"
						'Response.Write "unitid="&unitid&"<BR>"
						'Response.Write "jobstatus="&JobStatus&"<BR>"
						Select Case JobStatus
							Case "Acknowledged","ONB", "Paper on Board"
									
                                    'If (unitid="srvrfab") OR ((Request.Cookies("FleetXPhone")("sBT_ID")<>"26" AND Request.Cookies("FleetXPhone")("sBT_ID")<>"80"  AND Request.Cookies("FleetXPhone")("sBT_ID")<>"48" AND Request.Cookies("FleetXPhone")("sBT_ID")<>"75" AND MaterialType<>"xxxSecure Waf") OR (Request.Cookies("FleetXPhone")("sBT_ID")="26" AND trim(FromLocation)="PHO")) or trim(FromLocation)="PHO" or trim(ToLocation)="PHO"  then
										%>
                                        <!--
										<form method="post" action="DriverCloseWafer.asp" ID="Form3">
                                        -->
										<%
										'Else
										%>

										<form method="post" action="DriverClose.asp" ID="Form5">
										<%
									'End if
									%>
									<tr><td valign="top"><input type="submit" value="<%=ButtonText%>" name="submit" class="<%=ButtonClass%>" ID="gobutton2"></td>
									<input type="hidden" name="txtcaller" value="<%=trim(VehicleID)%>" ID="Hidden6">
									<input type="hidden" name="txtstation" value="<%=trim(ToLocation)%>" ID="Hidden7">
									<input type="hidden" name="txtjobnumber" value="<%=trim(jobnumber)%>" ID="Hidden8">
									<input type="hidden" name="VehicleID" value="<%=trim(VehicleID)%>" ID="Hidden28">
									<input type="hidden" name="LocationCode" value="<%=trim(LocationCode)%>" ID="Hidden29">
									<input type="hidden" name="jobnumber" value="<%=trim(jobnumber)%>" ID="Hidden30">	
									<input type="hidden" name="PageStatus" value="CLS" ID="Hidden15">
									<input type="hidden" name="BillToID" value="<%=Request.Cookies("FleetXPhone")("sBT_ID")%>" ID="Hidden2">
									<input type="hidden" name="AliasCode" value="<%=trim(AliasCode)%>" ID="Hidden33">
									<input type="hidden" name="Hub2" value="<%=Hub2%>" ID="Hidden49">
									<input type="hidden" name="Hub" value="<%=Hub%>" ID="Hidden50">																		
									</form>
									<%
							Case Else
								%>
								<tr><td valign="top">&nbsp;</td>
								<%
						End Select
						Response.Write "<td valign='top' nowrap><font color='"&PriorityColor&"'>"&Y&") <a href='DriverTracking.asp?JobNumber="&JobNumber&"&fh_bt_id="&fh_bt_id&"'><font color='"&PriorityColor&"'>"&JobNumber&"</font></a></font></td>"
						Response.Write "<td valign='top' nowrap align='center'><font color='"&PriorityColor&"'>"
					End if
					x=x+1
					TempJobNumber=JobNumber
					TempX=X
					If NumberOfLots>=1 then
						TempX=NumberOfLots
					end if					

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
					
				'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
				If CloseTable="y" then
					If BillToID<>"26" then
						Response.Write MaterialSymbol&TempX&MaterialSymbol
					End if
				Response.Write "</font></td>"
					%>
					</tr><!--/table-->
										<tr>
											<td colspan="6">
											<table ID="Table6" width="300" align="center" border="0">
							<td width="50%" align="left" valign="top"><font face="Verdana, arial" size="1">
								<%If cInt(intPage) > 1 Then%>
									<form method="post" ID="Form8">
									<input type="submit" name="submit" value="<<Previous" ID="gobutton2">
									<input type="hidden" name="pagestatus" value="<%=pagestatus%>" ID="Hidden11">
									<input type="hidden" name="VehicleID" value="<%=VehicleID%>" ID="Hidden20">
									<input type="hidden" name="LocationCode" value="<%=LocationCode%>" ID="Hidden21">
									<input type="hidden" name="truckstatus" value="<%=truckstatus%>" ID="Hidden12">						
									<input type="hidden" name="page" value="<%=intPage - 1%>" ID="Hidden13">
									<input type="hidden" name="AliasCode" value="<%=Trim(AliasCode)%>" ID="Hidden40">	
									<input type="hidden" name="Hub2" value="<%=Hub2%>" ID="Hidden44">
									<input type="hidden" name="Hub" value="<%=Hub%>" ID="Hidden45">		
									</form>

									<%
									else
									Response.write "&nbsp;"
								End If%>
								</font>
							</td>
							<td width="50%" align="right" valign="top"><font face="Verdana, arial" size="1" >
								<%If cInt(intPage) < cInt(intPageCount) Then%>
									<form method="post" ID="Form9">
									
									
									<input type="hidden" name="txtcaller" value="<%=trim(VehicleID)%>" ID="Hidden48">
									<input type="hidden" name="txtstation" value="<%=trim(ToLocation)%>" ID="Hidden51">
									<input type="hidden" name="BillToID" value="<%=BillToID%>" ID="Hidden57">
									<input type="hidden" name="AliasCode" value="<%=AliasCode%>" ID="Hidden41">		
									<input type="hidden" name="Hub2" value="<%=Hub2%>" ID="Hidden46">
									<input type="hidden" name="Hub" value="<%=Hub%>" ID="Hidden47">									
									<input type="hidden" name="pagestatus" value="<%=pagestatus%>" ID="Hidden17">
									<input type="hidden" name="VehicleID" value="<%=VehicleID%>" ID="Hidden38">
									<input type="hidden" name="LocationCode" value="<%=LocationCode%>" ID="Hidden39">
									<input type="hidden" name="truckstatus" value="<%=truckstatus%>" ID="Hidden18">
									<input type="hidden" name="page" value="<%=intPage + 1%>" ID="Hidden19">										
									<input type="submit" name="submit" value="Next>>" ID="gobutton2">
								
									</form>
									<%
									else
									Response.write "&nbsp;"
								End If%>
								</font>
							</td>			</table>				
											</td>
										</tr>						
					<!------------------------------------------------------------->
					<%
					CloseTable="n"
				End if
'-------------------STARTS THE PICK UP	
				If SetArrivalTime="y" and (BillToID="48" or BillToID="80") then
						Set oRs = Server.CreateObject("ADODB.Recordset")
						oRs.CursorLocation = 3
						oRs.CursorType = 3
						oRs.ActiveConnection = DATABASE	
						''''''CHECKS TO SEE IF ITS OKAY TO CHANGE THE ORDER STATUS
						SQL = "SELECT fl_fh_id FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id "
						'SQL = SQL&" WHERE (Fl_sf_ID='"&trim(LocationCode)&"') AND (fl_t_atp = '1/1/1900') AND (fl_un_id='"&trim(VehicleID)&"') AND (fl_t_int > '1/1/1900') AND (fh_ready<='"& NOW() &"')"
						Select Case BillToID
							Case "48"
								SQL = SQL&" WHERE (Fl_sf_ID='"&trim(LocationCode)&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND (fl_t_atp = '1/1/1900') AND (fl_un_id='"&trim(VehicleID)&"') AND (fl_t_int > '1/1/1900')"
							Case else
								if trim(vehicleID)="198" then
									SQL = SQL&" WHERE "&varFromLocations&" AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND (fl_t_atp = '1/1/1900') AND (fl_un_id='"&trim(VehicleID)&"') AND (fl_t_int > '1/1/1900')"
									else
									SQL = SQL&" WHERE "&varFromLocations&" AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND (fl_t_atp = '1/1/1900') AND (fl_un_id='"&trim(VehicleID)&"')"
								End if
						End Select
						oRs.Open SQL, DATABASE, 1, 3
						'Response.Write "XXXXXSQL="&SQL&"<BR>"
						If oRs.eof then
						End if
						Do while not oRs.EOF
							AtAirlineOrder=oRs("fl_fh_id")
							Set oConn = Server.CreateObject("ADODB.Connection")
							oConn.ConnectionTimeout = 100
							oConn.Provider = "MSDASQL"
							oConn.Open DATABASE
								oConn.Execute "PHONE_ATAIRLINE_ORDERS '" & AtAirlineOrder & "'" 
							oConn.Close
							Set oConn=Nothing
							''''''''''''''''''''''''''''''''''''''''
						oRs.movenext
						Loop
						oRs.Close
						Set oRs=Nothing					
				End if					
				X=0
				Y=0
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	


				'Response.write "<tr><td colspan='5'>trim(varFromLocations)="&trim(varFromLocations)&"</td></tr>"

                'Response.write "IDCookies="&IDCookies&"<BR>"


                Select Case IDCookies
                    Case "2024213", "2024215"
                        varFromLocations="( fl_sf_id='72' or fl_sf_id='RFAB-W' or fl_sf_id='RFAB-R' )"
                    Case "2024266", "932389"
                        'varFromLocations="( fl_sf_id='SCTICS' or fl_sf_id='SCTQC')"
                    Case Else

                End Select

                'If Trim(IDCookies)="2024213" OR Trim(IDCookies)="2024215" then
                '    varFromLocations="( fl_sf_id='72' or fl_sf_id='RFAB-W' or fl_sf_id='RFAB-R' )"
                    'Response.write "GOT HERE!!!!<BR>"
                '    else
                    'Response.write "bOOOOOOOOOOOOOOOOOOOOO!<br>"
               ' End if



				SQL = "SELECT Fl_SF_ID, Fh_ID, fh_User5, Fl_ST_ID, fl_sf_comment, FH_Status, fh_bt_id, Fh_Priority FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
				SQL = SQL&" WHERE (fh_ship_dt>'"&now()-30&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND  (Fl_un_ID='"&VehicleID&"') "
				'response.write "VehicleID="&VehicleID&"<BR>"
                If VehicleID=124 or vehicleID=113 then
					SQL = SQL&" and ((Fl_sf_ID='"&LocationCode&"') "
					else
					SQL = SQL&" and ("&varFromLocations
				End if
					SQL = SQL&" ) "	
				'''End if
				If OtherBillToID="48" or trim(vehicleID)="198" then
					SQL = SQL&" AND ((fh_status='PUO') or (fh_status='AC2'))"
					Else	
					SQL = SQL&" AND ((fh_status='ACC')"
					'''''If VehicleID=124 then
						SQL = SQL&" OR ((fh_status='AC2') AND (fl_secacc is not null))"
					SQL = SQL&") "
					SQL = SQL&" AND (fl_sf_id<>'HFABQC')"
				End if
				SQL = SQL&" ORDER BY fh_priority, fh_id"
				'Response.write "vehicleID="&vehicleID&"<BR>"
				'response.write "<br><br>PICK UP SQL="&SQL&"<BR>"
				
				oRs.Open SQL, DATABASE, 1, 3
				DisplayLocationCode=LocationCode
				If Trim(LocationCode)="SBRT" then DisplayLocationCode="SB-HUB" end if
				If Trim(LocationCode)="55" then DisplayLocationCode="CPGP" end if
				If trim(LocationCode)="48" then DisplayLocationCode="KWEO" end if
				If trim(LocationCode)="80" then DisplayLocationCode="LSP Warehouse" end if
				%>
                        <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>
                        <tr>
		                    <td class="FleetXRedSection" colspan="3" align="center">
			                    PICK UPS AT <%=uCase(DisplayLocationCode)%>
		                    </td>
	                    </tr>
                        <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>
				<%
				If not oRs.EOF then
          exception_bt_id = oRs("fh_bt_id")
					%>
						<tr>
							<td>&nbsp;</td>
							<td align="left" nowrap><b>&nbsp;&nbsp;&nbsp;&nbsp;Job #</b></td>
							<td align="center" nowrap>
							<%If BillToID<>"26" and BillToID<>"80" then%>
							<b>Lots</b>
							<%else%>
							<b>Refs</b>
							<%End if%>
							</td>
							</tr>
						<%
						CloseTable="y"
						ELSE
						Response.Write "<tr><td colspan='4' align='center'>No orders to pick up here.</td></tr><tr><td>&nbsp;</td></tr>"
				End if
				'''''''''''''''''''''''''''''''''''''''''''''''''''''
				oRS.PageSize = 4
				oRS.CacheSize = oRS.PageSize
				intPageCount2 = oRS.PageCount
				intRecordCount2 = oRS.RecordCount
				If (oRS.EOF) then
					Sendback2="y"
				End if
				If NOT (oRS.BOF AND oRS.EOF) Then

				If CInt(intPage2) > CInt(intPageCount2) Then intPage2 = intPageCount2
					If CInt(intPage2) <= 0 Then intPage2 = 1
						If intRecordCount2 > 0 Then
							oRS.AbsolutePage = intPage2
							intStart = oRS.AbsolutePosition
							If CInt(intPage2) = CInt(intPageCount2) Then
								intFinish = intRecordCount
							Else
								intFinish = intStart + (oRS.PageSize - 1)
							End if
						End If
					If intRecordCount2 > 0 Then
						For intRecord2 = 1 to oRS.PageSize				
				'''''''''''''''''''''''''''''''''''''''''''''''''''''
					FromLocation = trim(oRs("Fl_SF_ID"))
					JobNumber = trim(oRs("Fh_ID"))
					MaterialType = oRs("Fh_user5")
					ToLocation = trim(oRs("Fl_ST_ID"))
					fl_sf_comment=trim(oRs("fl_sf_comment"))
					JobStatus = trim(oRs("fh_status"))
					Priority = trim(oRs("fh_priority"))
					If MaterialType="300 mm Waf" or MaterialType="Foup/Fosby" then
						MaterialSymbol="*"
						else
						MaterialSymbol=""							
					End if					
                    If Priority="P0" or Priority="XP" then 
						PriorityColor="red"
						ButtonClass="ButtonRed"
                        ELSE
                    If Priority="P1" then 
						PriorityColor="blue"
						ButtonClass="ButtonBlue"
						Else
						ButtonClass="Button1"
						PriorityColor="black"
					End if
                    End iF
					If MaterialType="Secure Waf" OR MaterialType="secret" OR MaterialType="ITAR" then
						PriorityColor="Orange"
					End if					
					Select Case JobStatus
						Case "OPN"
							JobStatus="Open"
							ButtonText="Acknowledge"
						Case "ACC", "PUO", "ARV", "AC2"
							JobStatus="Acknowledged"
							ButtonText="ONB"
						Case "ONB"
							JobStatus="ONB"
							ButtonText="CLS"
					End Select
					If JobNumber<>TempJobNumber then
						If X>=0 then
							''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''							
							if trim(fh_bt_id)<>"26" then
								Set Recordset1 = Server.CreateObject("ADODB.Recordset")
								SQL_99="SELECT count(rf_ref) as NumberOfLots FROM FCREFS WHERE (rf_fh_id='"&JobNumber&"') AND ((ref_status<>'X') OR (ref_status is NULL))"
								'Response.Write "SQL_99="&SQL_99&"<BR>"
								Recordset1.ActiveConnection = DATABASE
								Recordset1.Source = SQL_99
								Recordset1.CursorType = 0
								Recordset1.CursorLocation = 2
								Recordset1.LockType = 1
								Recordset1.Open()
								Recordset1_numRows = 0
								if NOT Recordset1.EOF then
									NumberOfLots=Recordset1("NumberOfLots")
									'Response.Write "XXXNumberofLots="&NumberOfLots&"<BR>"
									If NumberOfLots>1 then WordLots="Lots" end if
									If NumberOfLots=1 then WordLots="Lot" end if
									If NumberOfLots=0 then WordLots="" end if
									Else
									Response.Write "&nbsp;"
									ErrorMessage="Incorrect driver ID or password"
								End if
								Recordset1.Close()
								Set Recordset1 = Nothing					
							End if							
							'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
							Response.Write "</font></td></tr>"
							Response.Write "<tr><td colspan='4' bgcolor='white'><img src='images/pixel.gif' height='2' width='1' border='0'></td></tr>"
							Response.Write "<tr><td colspan='4' bgcolor='gray'><img src='images/pixel.gif' height='1' width='1' border='0'></td></tr>"
							Response.Write "<tr><td colspan='4' bgcolor='white'><img src='images/pixel.gif' height='2' width='1' border='0'></td></tr>"
							X=0
						End if
						Y=Y+1
						If Priority="P0" or Priority="XP" then
							ButtonClass="ButtonRed"
                            else
						If Priority="P1" then
							ButtonClass="ButtonBlue"
							else
							ButtonClass="Button1"
						End if
                        End if
                        SomeVar=Request.Cookies("FleetXPhone")("sBT_ID")
                        'Response.write "SomeVar="&SomeVar&"<BR>"
						Select Case JobStatus
							Case "Acknowledged","ONB", "ARV", "AC2"
									'If trim(FromLocation)="PHO" OR trim(ToLocation)="PHO" OR (unitid="srvrfab") OR (Request.Cookies("FleetXPhone")("sBT_ID")<>"26" AND Request.Cookies("FleetXPhone")("sBT_ID")<>"48" AND Request.Cookies("FleetXPhone")("sBT_ID")<>"75" AND Request.Cookies("FleetXPhone")("sBT_ID")<>"80")   then
										%>
                                        <!--
										<form method="post" action="DriverCloseWafer.asp" ID="Form4">
                                        -->
										<%
										'Else
										%>
										<form method="post" action="DriverClose.asp" ID="Form2">
										<%
									'End if
									%>
									<td valign="top"><input type="submit" value="<%=ButtonText%>" name="submit" class="<%=ButtonClass%>" ID="gobutton2"></td>
									<input type="hidden" name="txtcaller" value="<%=trim(VehicleID)%>" ID="Hidden3">
									<input type="hidden" name="txtstation" value="<%=trim(FromLocation)%>" ID="Hidden4">
									<input type="hidden" name="txtjobnumber" value="<%=trim(jobnumber)%>" ID="Hidden5">
									<input type="hidden" name="VehicleID" value="<%=trim(VehicleID)%>" ID="Hidden25">
									<input type="hidden" name="LocationCode" value="<%=Trim(LocationCode)%>" ID="Hidden26">
									<input type="hidden" name="jobnumber" value="<%=Trim(jobnumber)%>" ID="Hidden27">	
									<input type="hidden" name="AliasCode" value="<%=Trim(AliasCode)%>" ID="Hidden31">
									<input type="hidden" name="BillToID" value="<%=Trim(BillToID)%>" ID="Hidden1">
									<input type="hidden" name="PageStatus" value="ONB" ID="Hidden14">								
									</form>
									<%
							Case Else
								%>
								<td valign="top">&nbsp;</td>
								<%
						End Select
						Response.Write "<td valign='top' nowrap><font color='"&PriorityColor&"'>"&Y&") <a href='DriverTracking.asp?JobNumber="&JobNumber&"&fh_bt_id="&fh_bt_id&"'><font color='"&PriorityColor&"'>"&JobNumber&"</font></a></font></td>"
						Response.Write "<td valign='top' nowrap align='center'><font color='"&PriorityColor&"'>"
						
					End if
					x=x+1
					TempJobNumber=JobNumber
					TempX=X
					''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				Response.Write MaterialSymbol&NumberOfLots&MaterialSymbol&"</font>"
						numberoflots=0
						If trim(fl_sf_comment)>"" then
						Response.Write "<tr><td colspan='3'>***"&fl_sf_comment&"</td></tr>"
					end if			

            ' add exceptions for this company
                Set Recordset1e = Server.CreateObject("ADODB.Recordset")
                'Response.write "Database="&Database&"<br>"
                Recordset1e.ActiveConnection = Database
                SQL = "SELECT a.accID, a.bt_id, a.atid, a.accCharge, a.accStatus, a.accDate, a.changedby, a.accstartdate, a.accstopdate, b.bt_id, b.bt_desc "_
                & " FROM Accessorials a "_
                & " INNER JOIN fcbillto b on b.bt_id = a.bt_id "_
                & " WHERE (a.accStatus='c') and a.bt_id = " & exception_bt_id & " and a.accstartdate < '" & Now() & "' and a.accstopdate >= '" & Now() & "'"
                
                Recordset1e.Source = SQL
                'response.write "SQL="& SQL &"<BR>"
                Recordset1e.CursorType = 0
                Recordset1e.CursorLocation = 2
                Recordset1e.LockType = 1
                Recordset1e.Open()
                Recordset1e_numRows = 0

                	if NOT Recordset1e.EOF then
                    response.write "<tr><td colspan=3>&nbsp;<br><b>EXCEPTIONS:</b><br>"
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
                    %><form method="post" action="AddJobException.asp?t=p&j=<%=jobnumber%>&b=<%=exception_bt_id%>&a=<%=RecordSet1e("accID")%>&c=<%=accTypeID%>&d=<%=accCharge%>"><input type="submit" id="gobutton" value="<%=accDescr%>" /> <%=accCharge2%></form><%
                    
                    Recordset1e.MoveNext
                    Loop
                    response.write  "</td></tr>"
                End if
                Recordset1e.Close()
                Set Recordset1e = Nothing      
				



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
				''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

				If CloseTable="y" then
					If BillToID<>"26" then
					End if
					Response.Write "</font></td>"				
					%>
					</tr>
					<tr><td>&nbsp;</td></tr><!--/table-->
					<%
					CloseTable="n"
				End if			
	
				If CloseTable="y" then
					%>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					</table>
					<%
					CloseTable="n"
				End if	
				%>

					<!------------------------------------------------------------->
										<tr>
											<td colspan="6">
											<table ID="Table6" width="300" align="center" border="0">
							<td width="50%" align="left" valign="top"><font face="Verdana, arial" size="1">
								<%If cInt(intPage2) > 1 Then%>
									<form method="post" ID="Form10">
									<input type="submit" name="submit" value="<<Previous" ID="gobutton2">
									<input type="hidden" name="txtcaller" value="<%=trim(VehicleID)%>" ID="Hidden52">
									<input type="hidden" name="txtstation" value="<%=trim(ToLocation)%>" ID="Hidden53">
									<input type="hidden" name="BillToID" value="<%=BillToID%>" ID="Hidden54">
									<input type="hidden" name="pagestatus" value="<%=pagestatus%>" ID="Hidden9">
									<input type="hidden" name="VehicleID" value="<%=VehicleID%>" ID="Hidden10">
									<input type="hidden" name="LocationCode" value="<%=LocationCode%>" ID="Hidden22">
									<input type="hidden" name="truckstatus" value="<%=truckstatus%>" ID="Hidden23">	
									<input type="hidden" name="AliasCode" value="<%=AliasCode%>" ID="Hidden42">							
									<input type="hidden" name="page2" value="<%=intPage2 - 1%>" ID="Hidden24"></form>
									</form>
									<%
									else
									Response.write "&nbsp;"
								End If%>
								</font>
							</td>
							<td width="50%" align="right" valign="top"><font face="Verdana, arial" size="1" >
								<%If cInt(intPage2) < cInt(intPageCount2) Then%>
									<form method="post" ID="Form11">
									<input type="submit" name="submit" value="Next>>" ID="gobutton2">
									<input type="hidden" name="pagestatus" value="<%=pagestatus%>" ID="Hidden32">
									<input type="hidden" name="VehicleID" value="<%=VehicleID%>" ID="Hidden34">
									<input type="hidden" name="LocationCode" value="<%=LocationCode%>" ID="Hidden35">
									<input type="hidden" name="truckstatus" value="<%=truckstatus%>" ID="Hidden36">
									<input type="hidden" name="AliasCode" value="<%=AliasCode%>" ID="Hidden43">		
									<input type="hidden" name="page2" value="<%=intPage2 + 1%>" ID="Hidden37"></form>
									</form>
									<%
									else
									Response.write "&nbsp;"
								End If%>
								</font>
							</td>			</table>				
											</td>
										</tr>	
                                        				
					<!------------------------------------------------------------->
			<%
			Case else
			Set oConn = Server.CreateObject("ADODB.Connection")
			oConn.ConnectionTimeout = 100
			oConn.Provider = "MSDASQL"
			oConn.Open DATABASE
				l_cSQL = "select st_id, st_addr1 from fcshipto  " &_
						 "WHERE st_alias = '" & TRIM(LocationAlias)&"'" 
				SET oRs = oConn.Execute(l_cSql)
				IF not oRs.EOF then	
						XYZ=XYZ+1
						st_addr1=oRs("st_addr1")
						LocationCode=oRs("st_id")
				End if
			Set oConn=Nothing				
			%>
				<table WIDTH="300" cellpadding="0" cellspacing="0" ID="Table2" align="left" border="0" bordercolor="red">
                <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>
                <form method="post" action="default.asp" ID="Form1">
					<tr><td align="center" colspan="3"><input type="submit" value="Return to Menu" ID="gobutton" NAME="Submit1">
                    </td></tr>
                </form>    			
				<!--/table-->
			<%
            'Response.write "phoneuser="&PhoneUser&"<BR>"

            'Response.write "IDCookies="&IDCookies&"<BR>"
            'Response.write "IDCookies="&IDCookies&"<BR>"

            'response.write "///////Android="&Android&"<BR>"
            'response.write "///////IDCookies="&IDCookies&"<BR>"
            'response.write "///////ourphone="&ourphone&"<BR>"
            If trim(Android)="y" or  trim(IDCookies)="14141414" or ourphone="y"  then %>
		
				<!--TABLE WIDTH="300" cellpadding="0" cellspacing="5" ID="Table3" align="left" border="1" bordercolor="red"-->
					 <form action="DriverifabPhoneEmulator.asp" method="post" name="thisForm" id="thisForm">
                    <TR> 
						<td> 
                           
								<table border="0" cellpadding="2" cellspacing="0" ID="Table4" width="100%" bordercolor="blue">
                        <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>
                        <tr>
		                    <td class="FleetXRedSection" colspan="6" align="center">
			                    SCAN LOCATION CODE
		                    </td>
	                    </tr>
                        <tr><td><img src="images/pixel.gif" height="30" width="1" /></td></tr>
									<tr>
										<td colspan='2' class='generalcontent' align="center">
											<input maxlength="20" name="AliasCode" id="txtstation" type="password" size="15" oncopy="return disabloecopypaste();"onpaste="return disablecopypaste();"oncut="return disablecopypaste();"  />
											<!--input maxlength="20" name="AliasCode" id="Password1" type="password" size="15" /-->
                                            <input name='VehicleID' id='VehicleID' value='<%=VehicleID%>' type="hidden" />
											<input type="hidden" name="FakeSubmit" value="fakesubmit" id="Hidden16" />
                                            
										</td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr><td colspan="2" align="center"><input size="8" maxlength="30" type="text" value="<%=Time()%>" name="bogus" id="Text1" onfocus="formSubmit()" readonly="readonly" class="InvisibleTextFieldPurple" />
                                    
                                    </td></tr>				

			
									<%if errormessage>"" then%>
										<tr>
											<td class='generalcontenthighlight'colspan='2' align="center"><font color="red"><br><b><%=ErrorMessage%></b><br><br></font></td>
										</tr>
									<%end if%>
									<tr> 
										<td class="subheader" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
									</tr>
                                    
								</table>
                            </form>
						</td>
						<!--Dummy section-->
					</TR>
					<tr><td align="center" colspan="4">&nbsp;</td></tr>					
				</TABLE>
			
			
			<%
            else
            Response.write "There is an issue with this device,<br>call Mark Maggiore<br>@ 214-956-0400 xt. 212<br>to resolve."
            End if
			End Select
			%>
			
	</BODY>
</HTML>
