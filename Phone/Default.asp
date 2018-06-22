<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
	<!--	
	<bgsound src="file://\sounds\alert1.wav" >
	<bgsound src="file://\windows\ringer.wav" loop="-1">
	-->
	<!--
	<EMBED src="file://\windows\ringer.wav" width="144" height="60" autostart="true" loop="false" hidden="true">
	-->
		<META HTTP-EQUIV="PRAGMA" CONTENT="NO-CACHE">
		<meta http-equiv="refresh" content="120" />

<script language="JavaScript">
<!--

var sURL = unescape(window.location.pathname);

function doLoad()
{
    // the timeout value should be the same as in the "refresh" meta-tag
    setTimeout( "refresh()", 130*1000 );
}

function refresh()
{
    //  This version of the refresh function will cause a new
    //  entry in the visitor's history.  It is provided for
    //  those browsers that only support JavaScript 1.0.
    //
    window.location.href = sURL;
}

//-->
</script>

<script language="JavaScript1.1">
<!--
function refresh()
{
    //  This version does NOT cause an entry in the browser's
    //  page view history.  Most browsers will always retrieve
    //  the document from the web-server whether it is already
    //  in the browsers page-cache or not.
    //  
    window.location.replace( sURL );
}
//-->
</script>

<script language="JavaScript1.2">
<!--
function refresh()
{
    //  This version of the refresh function will be invoked
    //  for browsers that support JavaScript version 1.2
    //
    
    //  The argument to the location.reload function determines
    //  if the browser should retrieve the document from the
    //  web-server.  In our example all we need to do is cause
    //  the JavaScript block in the document body to be
    //  re-evaluated.  If we needed to pull the document from
    //  the web-server again (such as where the document contents
    //  change dynamically) we would pass the argument as 'true'.
    //  
    window.location.reload( false );
}
//-->
</script>
<script language="JavaScript">
<!--
    javascript: window.history.forward(1);
//-->
</script>
		
		<!-- #include file="driverinfo.inc" -->	
		<!-- #include file="FleetX.inc" -->

		<title>Logisticorp Driver Home Page</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<%
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"


''''''''''''''''''''''''''''''''''''''''''
SecureYes = Request.ServerVariables ("HTTPS")
PhoneLocations=Request.Form("PhoneLocations")
If PhoneLocations="y" then
    Response.Cookies("LegalComputer").Expires = Date() + 3500
    Response.Cookies("LegalComputer")("ComputerID")=""

    Response.Cookies("LegalComputer").Expires = Date() + 3500
    Response.Cookies("LegalComputer")("ComputerID")="14141414"

    Response.write "***This device has now been set to 'God Mode.***"
    PhoneLocations=""
End if

''''''''''''SETS VEHICLE JOB COUNTS ETC. TO ZERO''''''''''''
NumJobsOnCall=0
NumJobsOnCall_OPN=0
NumJobsOnCall_ACC=0
NumJobsOnCall_ONB=0

NumJobs440=0
NumJobs754691=0
NumJobs303553=0
NumJobsSRVRFAB=0
NumJobsSRV=0
NumJobsSRB=0
NumJobsOFB=0
NumJobsOC=0

XP440=0
XP754691=0
XP303553=0
XPSRVRFAB=0
XPSRV=0
XPSRB=0
XPOFB=0
XPOC=0

P1440=0
P1754691=0
P1303553=0
P1SRVRFAB=0
P1SRV=0
P1SRB=0
P1OFB=0
P1OC=0

LATE440=0
LATE754691=0
LATE303553=0
LATESRVRFAB=0
LATESRV=0
LATESRB=0
LATEOFB=0
LATEOC=0

'If SecureYes="off" then
If SecureYes="on" then
	''''''''''''''''''''''''''''''''''''''''''''
	''''''''THIS NEEDS TO BE CHANGED WHEN MOVED TO PRODUCTION!''''''''
	something=Request.ServerVariables("HTTP_HOST") 
	Something=left(Something,4) 
	If lcase(something)="test" then 
		'''Response.redirect("http://test.logisticorp.us/h5redo/phone/default.asp")
		'Response.Write "GOT HERE!!!" 
		MPMSendEmail="n" 
		else
		'''Response.redirect("http://www.logisticorp.us/h5redo/phone/default.asp")
	End if 
	'Response.Write "Something="&Something&"<BR>"
	'Response.Write "MPMSendEmail="&MPMSendEmail&"<BR>"

	'''''''''''''''''''''''''''''''''''''''''''
		
End if
''''''''''''''''''''''''''''''''''''''''''		
		Response.Cookies("FleetXPhone")("PageStatus")=""
		Response.Cookies("FleetXPhone")("AliasCode")=""
		Response.Cookies("FleetXPhone")("FakeSubmit")=""
		mark=request.QueryString("mark")
		if mark="y" then
			response.write "helloooo!"
			response.write "VehicleID="&VehicleID&"<BR>"
		End if
        'response.write "VehicleID="&VehicleID&"<BR>"
        '-------------FIND OUT IF DRIVER HAS ANY EMAIL MESSAGES--------------


            Set Recordset1 = Server.CreateObject("ADODB.Recordset")
			Recordset1.ActiveConnection = Intranet
			Recordset1.Source = "SELECT DriverMessages.MessageID AS MessageID FROM DriverMessages"
            Recordset1.Source = Recordset1.Source&" WHERE (MessageRecipient='"&UserID&"') AND (MessageStatus='c')"
			'Response.Write "SQL="&Recordset1.Source&"<BR>"
			Recordset1.CursorType = 0
			Recordset1.CursorLocation = 2
			Recordset1.LockType = 1
			Recordset1.Open()
			Recordset1_numRows = 0

			DO WHILE NOT Recordset1.EOF 
				MessageID=Recordset1("MessageID")
				'DriverMessage=Recordset1("DriverMessage")
				'MessageDate=Recordset1("MessageDate")
                'Response.write "MessageID="&MessageID&"<BR>"
				'Response.write "DriverMessage="&DriverMessage&"<BR>"
				
				
				
				
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
				NumberofMessages=NumberofMessages+1
			End if
			Recordset2.Close()
			Set Recordset2 = Nothing				
				
				
				
				
				
				
				
			Recordset1.MoveNext
			LOOP
			Recordset1.Close()
			Set Recordset1 = Nothing


		'------------FIND ANY JOBS THAT NEED TO BE SORTED AT HUB------------
		If trim(VehicleID)="113" then
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				'Response.Write "vehicleID="&vehicleID&"<BR>"
				SQL = "SELECT Count(Fl_SF_ID) as NumberOfSorts FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
				SQL = SQL&" WHERE (fh_status='UNS') and (fh_bt_id='36') and (fl_dr_id='"& trim(VehicleID) &"')"
				'SQL = SQL&" ORDER BY Fh_Priority, fh_status desc, fh_id"
				'Response.write "<br><font color='blue'>SQL="&SQL&"<BR></font>"
				oRs.Open SQL, DATABASE, 1, 3
				If not oRs.eof then
				    NumberOfSorts=oRs("NumberOfSorts")
					If NumberOfSorts>0 then
					    SortMessage="y"
					    'Response.Write "GOT HERE!"
					End if
					if xxx="yyy" then
					%>
                    <bgsound src="sounds/lswarn.wav">
                <audio autoplay>
                 <source src="sounds/lswarn.wav"/>
                 Your browser does not support HTML5 audio.
                 </audio>

					<!--bgsound src="file://\sounds\lswarn.wav"-->

				    <%
				    end if
				End if
				oRs.Close
				Set oRs=Nothing	
		End if		
		
		'-------------------------------------------------------------------
		'-----FIND THE JOBS THAT NEED TO BE ACKNOWLEDGED------------------------------------------------------------
			'Response.Write "database="&database&"<BR>"
            
				Tommy=request.QueryString("Tommy")
				yyy=0
				BillToID=Request.Cookies("FleetXPhone")("sBT_ID")	
				'Response.Write "BillToID="&BillToID&"<BR>"
                'REsponse.write "VehicleID="&VehicleID&"<BR>"
				if isnumeric(VehicleID) then
                'Response.write "helloooooooooooooooooo<BR>"
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				'Response.Write "vehicleID="&vehicleID&"<BR>"
				'Response.Write "UnitID="&UnitID&"<BR>"
				SQL = "SELECT fh_priority, fh_user5, fl_un_id FROM VW_Driver_FindOpenJobs"
				SQL = SQL&" WHERE (Fl_UN_ID='"&trim(UnitID)&"')"
				'response.write "<br><font color='blue'>XXXSQL="&SQL&"<BR></font>"
				'End if
				
				oRs.Open SQL, DATABASE, 1, 3
					Do while not oRs.eof
                     'Response.write "helloooooooooooooooooo<BR>"
						NewJob="y"
						yyy=yyy+1
						thepriority = oRs("fh_priority")
						MaterialType = oRs("FH_User5")
						'fl_rt_type = trim(oRs("fl_rt_type"))
						If thepriority="P0" or thepriority="XP" or MaterialType="Secure Waf" or MaterialType="ITAR" then
							PZero="y"
						End if
						'Response.Write "NumberOfOrders="&NumberOfOrders&"<BR>"

					oRs.movenext
					Loop
					oRs.Close
					Set oRs=Nothing	
					end if
						NumberOfOrders=yyy
						'Response.Write "NewJob="&NewJob&"<BR>"
						'Response.Write "VehicleID="&VehicleID&"<BR>"
						If trim(VehicleID)="xxx" then
							'response.Write "GOT HERE!!!!<BR>"
							If showtop<>"y" then
							    %>
                                <bgsound src="sounds/dinnerbell.mp3">
                                 <audio autoplay>
                                 <source src="sounds/dinnerbell.mp3"/>
                                 Your browser does not support HTML5 audio.
                                 </audio>   							
							    <!--bgsound src="sounds/dinnerbell.mp3"-->
							    <%
							End if
						End if
						'response.Write "NewJob="&NewJob&"<br>"
                        'response.Write "PZero="&PZero&"<br>"
                        'response.Write "showtop="&showtop&"<br>"
                        'response.Write "vehicleID="&vehicleID&"<br>"

						If NewJob="y" and PZero<>"y" then
							If trim(vehicleID)<>"1" and trim(vehicleID)<>"2" and trim(vehicleID)<>"145" then
								'Response.Write "vehicleID="&VehicleID&".<BR>"
								If showtop<>"y" then
                                    'Response.write "Got hereXXXX!<BR>"
								    %> 
                                    <bgsound src="sounds/newjob.wav" loop="1"> 
                                 <audio autoplay>
                                 <source src="sounds/newjob.wav"/>
                                 Your browser does not support HTML5 audio.
                                 </audio> 
                                    <!--bgsound src="sounds/newjob.wav" loop="1"--> 
								    <%
                                    mail="n"
								end if
								else%>
                                     <bgsound src="sounds/newjob.wav"> 
                                 <audio autoplay>
                                 <source src="sounds/newjob.wav"/>
                                 Your browser does not support HTML5 audio.
                                 </audio> 
                                    <!--bgsound src="sounds/newjob.wav" loop="1"--> 

								<%
							End if
						End if
						If PZero="y" then
							If trim(vehicleID)<>"1" and trim(vehicleID)<>"2" then
							    If showtop<>"y" then
                                    'Response.write "Got here 1<BR>"
								    %>

                                    <bgsound src="sounds/lswarn.wav">
                                 <audio autoplay>
                                 <source src="sounds/lswarn.wav"/>
                                 Your browser does not support HTML5 audio.
                                 </audio> 
								    <!--bgsound src="file://\sounds\lswarn.wav"-->
                                    
								    <%
                                    mail="n"
								end if
								else
                                'Response.write "Got here 2<BR>"
                                %>
                                <bgsound src="sounds/PZero.wav">
                                 <audio autoplay>
                                 <source src="sounds/PZero.wav"/>
                                 Your browser does not support HTML5 audio.
                                 </audio> 
								<!--bgsound src="sounds/PZero.wav"-->
								<%
							End if
						End if											
					BillToID=Request.Cookies("FleetXPhone")("sBT_ID")	
					'response.Write "BillToID="&BillToID&"<BR>"
					If not isnumeric(BillToID) then
                        'Response.write "Got here 375<br>"
						'Response.redirect("driverlogin.asp")
					End if
                    'Response.write "VehicleID="&VehicleID&"<BR>"
				if  (BillToID="48" or BillToID="80") then
				'response.Write "got here<br>!"
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				'Response.Write "vehicleID="&vehicleID&"<BR>"
				SQL = "SELECT Count(Fl_SF_ID) as NumberOfPaper FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
				If trim(vehicleID)="199" then
					SQL = SQL&" WHERE (Fl_dr_ID='"&cInt(VehicleID)&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL))"
					SQL = SQL&" AND (Fh_Status='ONB') and (fl_rt_type<>'out')"
					else
					SQL = SQL&" WHERE (Fl_dr_ID='"&cInt(VehicleID)&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL))"
				End if
				If VehicleID<>124 and VehicleID<>199 then
					SQL = SQL&" AND (Fh_Status='ACC')"
				End if
				'SQL = SQL&" ORDER BY Fh_Priority, fh_status desc, fh_id"
				'Response.write "<br><font color='blue'>SQL="&SQL&"<BR></font>"
				oRs.Open SQL, DATABASE, 1, 3
				If not oRs.eof then
					NumberOfPaper = oRs("NumberOfPaper")
				End if
				oRs.Close
				Set oRs=Nothing	
					end if	
	
		'-----------------------------------------------------------------
		'Response.Write "VehicleID="&VehicleID&"<br>"
		%>
        <!--bgsound src="sounds/dinnerbell.mp3" loop="1"--> 
	</head>
<body onload="doLoad()">
<!-- #include file="LogoSection.asp" -->	
	<table cellspacing="0" cellpadding="0" border="0" width="300">
	<tr>
		<td class="FleetXRedSection" colspan="2" align="center">
			Driver Home Page
		</td>
	</tr>
			<tr>
				<td class="PhonePageTextBoldRight">
					Last update: <%=Time()%>
				</td>
			</tr>
			<tr><td class="MainPageText"><font>			
			
			<%


		'Response.Write "AlertMessage="&AlertMessage&"****<BR>"
		If trim(alertmessage)>"" then
			Response.Write "<font color='red'><br>*******************<BR>"
			Response.Write alertmessage&"<BR>"
			Response.Write "*******************<BR></font>"
			If   NewJob<>"y" and PZero<>"y" then
				If trim(vehicleID)<>"1" and trim(vehicleID)<>"2" then
				    If showtop<>"y" then
					    %>
                                <bgsound src="sounds/cshk-effect.wav">
                                 <audio autoplay>
                                 <source src="sounds/cshk-effect.wav"/>
                                 Your browser does not support HTML5 audio.
                                 </audio> 
					    <!--bgsound src="file://\sounds\cshk-effect.wav"-->
					    <%
                        mail="n"
					end if    
					else%>
                                <bgsound src="sounds/PZero.wav">
                                 <audio autoplay>
                                 <source src="sounds/PZero.wav"/>
                                 Your browser does not support HTML5 audio.
                                 </audio> 
					<!--bgsound src="sounds/PZero.wav"-->
					<%
				End if
				else
				
			End if
			%>		
			<!--
				<SCRIPT type="text/javascript" language="JavaScript">	
					var AlertMessage="<%=alertmessage%>";
					alert(AlertMessage);
				</script>		
			-->
			
			<%
		end if							
		%>
		</font>
		</td></tr>				
		<%

                    UnitID=Trim(UnitID)
		''''''''''''NEW VIEW TO SHOW THE OPEN ORDERS/PO's/LATES FOR THE OTHER VEHICLES''''''''''
        Select Case UnitID
            Case "440"
                ShowOther="754691"
                ShowOther2="607"
            Case "754691"
                ShowOther="440"
                ShowOther2="x"
            Case "607"
                ShowOther="440"
                ShowOther2="x"
            Case else
                ShowOther="x"
                ShowOther2="x"
        End Select
            'Response.write "Line 421 UnitID="&UnitID&"<BR>"
				If ShowOther<>"x" then
                Response.write "<table border='0' bordercolor='#d71e26' cellspacing='0' cellpadding='0'><tr><td align='center'><b>IN UNIT "&ShowOther&"</b></td></tr><tr><td>"

				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE
                'Response.Write "DATABASE="&DATABASE&"<BR>"	
				'Response.Write "vehicleID="&vehicleID&"<BR>"
				SQL = "SELECT fh_priority, fh_user5, fl_un_id, fl_st_rta, fh_status, fl_dr_id, fl_sf_id, fl_st_id FROM VW_Driver_FindAllOpenJobsForAllWaferTrucks"
				SQL = SQL&" WHERE (Fl_UN_ID='"&ShowOther&"') order by fl_st_rta desc"
				'response.write "<br><font color='Red'>XXXSQL="&SQL&"<BR></font>"
				'End if

				oRs.Open SQL, DATABASE, 1, 3
                    If not oRs.eof then
                    Response.write "<table border='0'>"
                    CloseTable="y"
                    End if
					Do while not oRs.eof 
						NewJob="y"
						X66=X66+1
				       ' Response.Write "X66="&X66&"<BR>"
						AllPriority = oRs("fh_priority")
						AllMaterialType = oRs("FH_User5")
						AllVehicle = Trim(oRs("fl_un_id"))
						AllDueTime = oRs("fl_st_rta")
                        AllFh_status = oRs("fh_status")
                        Orig= oRs("fl_sf_id")
                        Dest = oRs("fl_st_id")
                        'Response.write "AllDueTIme="&AllDueTime&"<BR>"
                        Select Case AllFh_Status
                            Case "ONB"
                                varFontColor66="Green"
                            Case else
                                varFontColor66="Black"
                        End select
                        TimeRemaining=(DateDiff("n",now(),AllDueTime))
                        If TimeRemaining<0 then
                            TimeRemaining="LATE"
                            else
                            If TimeRemaining>60 then
                                RemainingHours=int(TimeRemaining/60)
                                'Response.write "RemainingHours="&RemainingHOurs&"<BR>"
                                RemainingMinutes=(TimeRemaining-(RemainingHours*60))
                                TimeRemaining=RemainingHours&" hrs "&RemainingMinutes&" mins"
                                else
                                TimeRemaining=TimeRemaining&" mins"
                            End if
                        End if
                        'If
                        %>
                        <tr><td><font color="<%=varFontColor66%>"><%=Orig %></font></td><td><font color="<%=varFontColor66%>">---></font></td><td><font color="<%=varFontColor66%>"><%=Dest %></font></td><td><font color="<%=varFontColor66%>">(<%=TimeRemaining %>) <%if AllFh_status="ONB" then response.write "(ONB)" end if %></font></td></tr>
                        <%
                        
						'Response.Write "****AllVehicle="&AllVehicle&"<BR>"


					oRs.movenext
					Loop
					oRs.Close
					Set oRs=Nothing	
                    
                    End if
                    If CloseTable="y" then
                    Response.write "</table></td></tr></table>"
                    End if


				If ShowOther2<>"x" then
                Response.write "<table border='0' bordercolor='#d71e26' cellspacing='0' cellpadding='0'><tr><td align='center'><b>IN UNIT "&ShowOther2&"</b></td></tr><tr><td>"

				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE
                'Response.Write "DATABASE="&DATABASE&"<BR>"	
				'Response.Write "vehicleID="&vehicleID&"<BR>"
				SQL = "SELECT fh_priority, fh_user5, fl_un_id, fl_st_rta, fl_dr_id, fl_sf_id, fl_st_id FROM VW_Driver_FindAllOpenJobsForAllWaferTrucks"
				SQL = SQL&" WHERE (Fl_UN_ID='"&ShowOther2&"') order by fl_st_rta desc"
				'response.write "<br><font color='Red'>XXXSQL="&SQL&"<BR></font>"
				'End if

				oRs.Open SQL, DATABASE, 1, 3
                    If not oRs.eof then
                    Response.write "<table border='0'>"
                    CloseTable="y"
                    End if
					Do while not oRs.eof 
						NewJob="y"
						X66=X66+1
				       ' Response.Write "X66="&X66&"<BR>"
						AllPriority = oRs("fh_priority")
						AllMaterialType = oRs("FH_User5")
						AllVehicle = Trim(oRs("fl_un_id"))
						AllDueTime = oRs("fl_st_rta")
                        Orig= oRs("fl_sf_id")
                        Dest = oRs("fl_st_id")
                        'Response.write "AllDueTIme="&AllDueTime&"<BR>"
                        TimeRemaining=(DateDiff("n",now(),AllDueTime))
                        If TimeRemaining<0 then
                            TimeRemaining="LATE"
                            else
                            If TimeRemaining>60 then
                                RemainingHours=int(TimeRemaining/60)
                                'Response.write "RemainingHours="&RemainingHOurs&"<BR>"
                                RemainingMinutes=(TimeRemaining-(RemainingHours*60))
                                TimeRemaining=RemainingHours&" hrs "&RemainingMinutes&" mins"
                                else
                                TimeRemaining=TimeRemaining&" mins"
                            End if
                        End if
                        'If
                        %>
                        <tr><td><%=Orig %></td><td>---></td><td><%=Dest %></td><td>(<%=TimeRemaining %>)</td></tr>
                        <%
                        
						'Response.Write "****AllVehicle="&AllVehicle&"<BR>"


					oRs.movenext
					Loop
					oRs.Close
					Set oRs=Nothing	
                    
                    End if
                    If CloseTable="y" then
                    Response.write "</table></td></tr></table>"
                    End if


'end if %>


<%'Response.write "UnitID="&Unitid&"<br />"%>

		<%If (ucase(UnitID)="x" OR ucase(UnitID)="x" Or ucase(UnitID)="OFB" Or ucase(UnitID)="srvrfab"  Or ucase(UnitID)="OCV" ) then 
		If NumJobsSRV>10 then JobFontSRV="red" else JobFontSRV="Green" end if 
		If NumJobsSRB>10 then JobFontSRB="red" else JobFontSRB="Green" end if 
		If NumJobsOFB>10 then JobFontOFB="red" else JobFontOFB="Green" end if 
        If NumJobsOC>10 then JobFontOC="red" else JobFontOC="Green" end if 
		'If NumJobsSRVRFAB>10 then JobFontSRVRFAB="red" else JobFontSRVRFAB="Green" end if 
		If XPSRV>0 then XPFontSRV="red" else XPFontSRV="Green" end if 
		If XPSRB>0 then XPFontSRB="red" else XPFontSRB="Green" end if 
		If XPOFB>0 then XPFontOFB="red" else XPFontOFB="Green" end if 
        If XPOC>0 then XPFontOC="red" else XPFontOC="Green" end if 

		'If NumJobsSRVRFAB>10 then JobFontSRVRFAB="red" else JobFontSRVRFAB="Green" end if 
		If P1SRV>0 then P1FontSRV="red" else P1FontSRV="Green" end if 
		If P1SRB>0 then P1FontSRB="red" else P1FontSRB="Green" end if 
		If P1OFB>0 then P1FontOFB="red" else P1FontOFB="Green" end if 
        If P1OC>0 then P1FontOC="red" else P1FontOC="Green" end if 


		'If XPSRVRFAB>0 then XPFontSRVRFAB="red" else XPFontSRVRFAB="Green" end if 
		If LateSRV>0 then LateFontSRV="red" else LateFontSRV="Green" end if 
		If LateSRB>0 then LateFontSRB="red" else LateFontSRB="Green" end if 
		If LateOFB>0 then LateFontOFB="red" else LateFontOFB="Green" end if 
        If LateOC>0 then LateFontOC="red" else LateFontOC="Green" end if 
		'If LateSRVRFAB>0 then LateFontSRVRFAB="red" else LateFontSRVRFAB="Green" end if 		
		'Response.Write "unitID="&UnitID&"<BR>"
		%>
		    <tr><td>&nbsp;hello????</td></tr>
		    <%'If Trim(UnitID)<>"srv" then %>
		        <tr><td class="mainpagetextbold">SRV: <font color="<%=JobFontSRV%>">Jobs=<%=NumJobsSRV%></font> / <font color="<%=XPFontSRV%>">P0=<%=XPSRV%></font> / <font color="<%=P1FontSRV%>">P1=<%=P1SRV%></font> / <font color="<%=LateFontSRV%>">Late=<%=LateSRV%></font></td></tr>
		    <%If trim(UnitID)<>"ofb" and trim(UnitID)<>"srv" then%>
		        <tr><td class="mainpagetextbold">OC: <font color="<%=JobFontOC%>">Jobs=<%=NumJobsOC%></font> / <font color="<%=XPFontOC%>">P0=<%=XPOC%></font> / <font color="<%=P1FontOC%>">P1=<%=P1OC%></font> / <font color="<%=LateFontOC%>">Late=<%=LateOC%></font></td></tr>
            <%End if%>
		   <%If trim(UnitID)<>"xxxofb" and trim(UnitID)<>"xxxsrv" then%>
		        <tr><td class="mainpagetextbold">SRMH: <font color="<%=JobFontSRB%>">Jobs=<%=NumJobsSRB%></font> / <font color="<%=XPFontSRB%>">P0=<%=XPSRB%></font> / <font color="<%=P1FontSRB%>">P1=<%=P1SRB%></font> / <font color="<%=LateFontSRB%>">Late=<%=LateSRB%></font></td></tr>
		    <%End if%>
		    <%'''If UnitID<>"OFB" then %>
		        <tr><td class="mainpagetextbold">OFB: <font color="<%=JobFontOFB%>">Jobs=<%=NumJobsOFB%></font> / <font color="<%=XPFontOFB%>">P0=<%=XPOFB%></font> / <font color="<%=P1FontOFB%>">P1=<%=P1OFB%></font> / <font color="<%=LateFontOFB%>">Late=<%=LateOFB%></font></td></tr>
		    <%'''End if%>
		    <%If ucase(UnitID)="SRVRFABxxx" then %>
		        <tr><td class="mainpagetextbold">W4: <font color="<%=JobFontSRVRFAB%>">Jobs=<%=NumJobsSRVRFAB%></font> / <font color="<%=XPFontSRVRFAB%>">P0=<%=XPSRVRFAB%></font> / <font color="<%=P1FontSRVRFAB%>">P1=<%=P1SRVRFAB%></font> / <font color="<%=LateFontSRVRFAB%>">Late=<%=LateSRVRFAB%></font></td></tr>
		    <%End if%>
		<%end if %>
		<%
		If NumberOfMessages>0 then
		%>
        
		<tr><td>&nbsp;</td></tr>
		<form method="post" action="DriverMessage.asp">
			<tr><td><input type="submit" id="gobutton" value="New Messages (<%=NumberOfMessages%>)"></td></tr>
		</form>	
        <%
        'Response.write "Mail="&Mail&"<BR>"
        If mail<>"n" then 
        'Response.write "Got here!!!!!<BR>"
        %>
                <bgsound src="sounds/yougotmail.wav"> 
                <audio autoplay>
                 <source src="sounds/yougotmail.wav"/>
                 Your browser does not support HTML5 audio.
                 </audio>
        <!--bgsound src="sounds/yougotmail.wav" loop="1"--> 
		<%
        End if
		End if
		%>
		<%
		If SortMessage="y" then
		%>
		<tr><td class="mainpagetextbold"><font color="blue">***There are jobs to sort at SBHUB-W***</font></td></tr>
		<%
		End if
		%>


		
		<tr><td>&nbsp;</td></tr>
        <%
 				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				'Response.Write "vehicleID="&vehicleID&"<BR>"
				SQL = "SELECT fh_carr_id AS OtherVehicles, COUNT(fh_id) AS OtherNumberOfOrders FROM fcfgthd WHERE (fh_status NOT IN ('cls', 'can')) "
                SQL = SQL&" AND ((fh_carr_id <> '"&UnitID&"') AND (fh_carr_id <> 'V"&UnitID&"')) "
				SQL = SQL&" GROUP BY fh_carr_id ORDER BY fh_carr_id"
				'SQL = SQL&" ORDER BY Fh_Priority, fh_status desc, fh_id"
				'Response.write "<br><font color='blue'>SQL="&SQL&"<BR></font>"
				oRs.Open SQL, DATABASE, 1, 3
                    If not oRs.eof then
                    ZYX="y"
                    %>
                    <tr>
                        <td align="center">
                            <table>
                                <tr><td class="PhonePageTextBoldLeft">OTHER VEHICLE(S)</td><td><img src="../images/pixel.gif" height="1" width="1" alt="blank" /></td><td class="PhonePageTextBoldLeft">JOBS</td></tr>
                    <%
                    End if
					Do while not oRs.eof
						OtherVehicles = oRs("OtherVehicles")
                        OtherNumberOfOrders = oRs("OtherNumberOfOrders")
                        %>
                    <tr><td colspan="3" bgcolor="black"><img src="../images/pixel.gif" height="1" width="1" alt="blank" /></td></tr>
                    <tr><td class="PhonePageTextBoldLeft"><%=OtherVehicles %></td><td><img src="../images/pixel.gif" height="1" width="10" alt="blank" /></td><td class="PhonePageTextBoldLeft"><%=OtherNumberOfOrders %></td></tr>
                        <%
			oRs.MoveNext
			LOOP
			oRs.Close()
			Set oRs = Nothing
            If ZYX="y" then
                %>
                            <tr><td>&nbsp;</td></tr>
                        </table>
            
                    </td>
                </tr>
                <%
            End if
      
         %>

		
		<form method="post" action="DriverAcknowledge.asp">
			<tr><td><input type="submit" id="gobutton" value="Acknowledge Orders (<%=NumberOfOrders%>)"></td></tr>
		</form>	
		<tr><td>&nbsp;</td></tr>
		<%
		If UnitID="1" or UnitID="2" or UnitID="199" or UnitID="198" or billtoID="48" Then
		'response.Write "billtoid="&billtoid&"<BR>"
		If BillToID="80" then
		%>
		<form method="post" action="DriverAIMSPOBLocations.asp" ID="Form12">
		<%
		else
		%>
		<form method="post" action="DriverPOB.asp" ID="Form6">
		<%
		end if
		%>	
			<tr><td><input type="submit" id="gobutton" value="Paper On Board  (<%=NumberOfPaper%>)" ID="Submit2" NAME="Submit1"></td></tr>
		</form>	
		<tr><td>&nbsp;</td></tr>
		<form method="post" action="DriverExceptions.asp" ID="Form7">	
			<tr><td><input type="submit" id="gobutton" value="Exceptions" ID="Submit3" NAME="Submit1"></td></tr>
		</form>	
		<tr><td>&nbsp;</td></tr>
		<%
		End if
		'Response.Write "UnitID="&UnitID&"<BR>"
		If UnitID="AIMS1" or UnitID="AIMS2" Then
		
		
		
			if isnumeric(VehicleID) and BillToID=75 then
				'response.Write "got here<br>!"
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				'Response.Write "vehicleID="&vehicleID&"<BR>"
				SQL = "SELECT Count(Distinct(Fl_FH_ID)) as BOLNeeded FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN fcrefs ON fclegs.fl_fh_id = fcrefs.rf_fh_id"
				SQL = SQL&" WHERE (Fl_dr_ID='"&cInt(VehicleID)&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL))"
				If VehicleID<>124 then
					SQL = SQL&" AND (Fh_Status='ACC') AND (fcrefs.rf_box = '')"
				End if
				'SQL = SQL&" ORDER BY Fh_Priority, fh_status desc, fh_id"
				'Response.write "<br><font color='blue'>SQL="&SQL&"<BR></font>"
				oRs.Open SQL, DATABASE, 1, 3
					If not oRs.eof then
						BOLNeeded = oRs("BOLNeeded")
					End if
					oRs.Close
					Set oRs=Nothing	
			end if
					%>						
		
		
		
		<form method="post" action="DriverBOL.asp" ID="Form8">	
			<tr><td><input type="submit" id="gobutton" value="Create a BOL (<%=BOLNeeded%>)" ID="Submit4" NAME="Submit1"></td></tr>
		</form>	
		<tr><td>&nbsp;</td></tr>
		<%End if
		
		
		'Response.Write "VehicleID="&VehicleID&"<BR>"
		if Trim(VehicleID)="613" or Trim(VehicleID)="612" then %>
		
        <form method="post" action="DriverIfabPhoneEmulator_SFABQC.asp" ID="Form17">
	        <tr><td><input type="submit" id="gobutton" value="Drop Off/Pick Up by Cart"></td></tr>
	        <input type="hidden" name="UserID" value="<%=UserID%>">
        </form>	
         <tr><td>&nbsp;</td></tr>
		<%end if %>
		<%
		VehicleID=Trim(VehicleID)
	    Select Case VehicleID
	        Case "312", "313", "212", "314", "701", "113"%>
	            
                <form method="post" action="DriverSortCart_HFABQC.asp" ID="Form18">
                    <tr><td><input id="gobutton" type="submit" value="Sort Cart at HUB"></td></tr>
                    <input type="hidden" name="UserID" value="<%=UserID%>">
                </form>	
               <tr><td>&nbsp;</td></tr>
	    <%end Select 	
		
		'response.Write "BillToID="&BillToID&"<BR>"
		If BillToID="48" then
		   ' Response.Write "KWE<BR>"
		End if
		Select Case BillToID
		    Case "75", "80", "88", "86"
		        'If BillToID=75 or BillToID=80 then
		        %>
		        <form method="post" action="DriverAIMSLocations.asp" ID="Form9">
			        <tr><td><input type="submit" id="gobutton" value="Drop Off/Pick Up" ID="Submit5" NAME="Submit5"></td></tr>
			        <input type="hidden" name="UserID" value="<%=UserID%>" ID="Hidden1">
		        </form>	
		        <%
		    Case "48"
		        %>				
		        <form method="post" action="DriverIfabPhoneEmulator_KWE.asp" ID="Form16">
			        <tr><td><input type="submit" id="gobutton" value="Drop Off/Pick Up"></td></tr>
			        <input type="hidden" name="UserID" value="<%=UserID%>">
		        </form>
		        <%		    	
		    Case Else
		        %>				
		        <form method="post" action="DriverIfabPhoneEmulator.asp" ID="Form1">
			        <tr><td><input type="submit" id="gobutton" value="Drop Off/Pick Up"></td></tr>
			        <input type="hidden" name="UserID" value="<%=UserID%>">
		        </form>
		        <%
		End Select
		%>
		
		<tr><td>&nbsp;</td></tr>
		<form method="post" action="DriverTruckLoad.asp" ID="Form2">	
			<tr><td><input type="submit" id="gobutton" value="Current Routing"></td></tr>
		</form>	
		<tr><td>&nbsp;</td></tr>
		<%
		'Response.Write "BillToID="&BillToID&"<BR>"
		'Response.Write "UnitID="&UnitID&"<BR>"
		If BillToID="36" or BillToID="38" then%>
			<form method="post" action="DriverExceptions_IFAB.asp" ID="Form19">	
				<tr><td><input type="submit" id="gobutton" value="Exceptions" ID="Submit9" NAME="Submit1"></td></tr>
			</form>	
			<tr><td>&nbsp;</td></tr>		
			<%
		End if		
		If BillToID="75" then%>
			<form method="post" action="DriverExceptions.asp" ID="Form10">	
				<tr><td><input type="submit" id="gobutton" value="Exceptions" ID="Submit6" NAME="Submit1"></td></tr>
			</form>	
			<tr><td>&nbsp;</td></tr>		
			<%
		End if
		'Response.Write "UnitID="&UnitID&"<BR>"
		'If UnitID="xxx" or UnitID="xxx" or UnitID="303553" or UnitID="303554" or UCASE(UnitID)="SRVRFAB" or UnitID="srv" or UnitId="ofb" or UnitID="SHERMAN" or UnitID="OCV" or lcase(UnitID)="srv" or lcase(UnitID)="1" or lcase(UnitID)="2" or lcase(UnitID)="3" or lcase(UnitID)="4" or lcase(UnitID)="5" or lcase(UnitID)="6" or lcase(UnitID)="7" or lcase(UnitID)="srb" Then%>
		<form method="post" action="DriverHandOff.asp" ID="Form5">	
			<tr><td><input type="submit" id="gobutton" value="Handoff a Job" ID="Submit1" NAME="Submit1"></td></tr>
		</form>	
		<tr><td>&nbsp;</td></tr>
		<%'End if%>	
		<%if UnitID="1" or UnitID="2" then%>		
			<form method="post" action="DriverInterimScan.asp" ID="Form11">	
				<tr><td><input type="submit" id="gobutton" value="Interim Shipments" ID="Submit7" NAME="Submit1"></td></tr>
			</form>	
			<tr><td>&nbsp;</td></tr>				
		<%End if%>			
		<form method="post" action="DriverVehicle.asp" ID="Form3">
			<tr><td><input type="submit" id="gobutton" value="Change Vehicle"></td></tr>
		</form>
		<tr><td>&nbsp;</td></tr>
		<form method="post" action="DriverBOL.asp" ID="Form13a">
			<tr><td><input type="submit" id="gobutton" value="Driver Bill of Lading" ID="Submit8a" NAME="Submit8a"></td></tr>
		</form>
		<tr><td>&nbsp;</td></tr>
		<form method="post" action="DriverPhoneList.asp" ID="Form13">
			<tr><td><input type="submit" id="gobutton" value="Phone List" ID="Submit8" NAME="Submit8"></td></tr>
		</form>
		<%If trim(vehicleID)="312" or trim(vehicleID)="313" or trim(vehicleID)="212" or trim(vehicleID)="314" then
	'''''''''''''FOR LUNCH/BREAK CODE...IF THEY LEAVE THE PAGE TO CIRCUMVENT THE BREAK PAGE''''''''''
			Set oConn = Server.CreateObject("ADODB.Connection")
			oConn.ConnectionTimeout = 100
			oConn.Provider = "MSDASQL"
			oConn.Open DATABASE
				l_cSQL = "UPDATE breaktable SET EndTime = '"& Now() &"' WHERE Userid = '" & UserID & "' and EndTime is NULL"
				oConn.Execute(l_cSQL)
			Set oConn=Nothing
			'Response.Write "l_cSQL="&l_cSQL&"<BR>"	
		%>
		<tr><td>&nbsp;</td></tr>			
		<form method="post" action="DriverBreak.asp?a=l&b=<%=DateAdd("n",30,now())%>&c=y" ID="Form14">
			<tr><td><input type="submit" id="gobutton" value="Take Lunch"></td></tr>
		</form>
		<tr><td>&nbsp;</td></tr>			
		<form method="post" action="DriverBreak.asp?a=b&b=<%=DateAdd("n",15,now())%>&c=y" ID="Form15">
			<tr><td><input type="submit" id="gobutton" value="Take a Break"></td></tr>
		</form>				
		<%end if%>
		<tr><td>&nbsp;</td></tr>			
		<form method="post" action="TestSoundsOnPhone.asp" ID="Form20">
			<tr><td><input type="submit" id="gobutton" value="Test Sounds on Phone"></td></tr>
		</form>
        <tr><td>&nbsp;</td></tr>
        <%If trim(UserID)="1" then %>			

        <%end if 
        'Response.write "userID="&UserID&"<BR>"%>
        <%Select Case UserID
            Case "1", "20", "60", "11", "519", "301", "445", "44", "82" %>
		    <form method="post" action="setcookie.asp" ID="Form22">
                <input type="hidden" name="validated" value="y" />
			    <tr><td><input type="submit" id="gobutton" value="Set Device ID"></td></tr>
		    </form>
            <tr><td>&nbsp;</td></tr>			
        <%
        Case Else
        end Select 
        Select Case UserID
            Case "1", "20", "60", "11", "519", "82" %>
		<form method="post" action="default.asp" ID="Form24">
            <input type="hidden" name="PhoneLocations" value="y" />
			<tr><td><input type="submit" id="gobutton" value="Re-Set Phone Locations"></td></tr>
            <tr><td>&nbsp;</td></tr>
		</form>
        <%
        Case Else
        end Select %>
		<form method="post" action="DriverLogOut.asp" ID="Form4">
			<tr><td><input type="submit" id="gobutton" value="Log Out"></td></tr>
			<input type="hidden" name="FakeSubmit" value="byebyesession" />
		</form>
	</table>
    
	</body>
</html>
