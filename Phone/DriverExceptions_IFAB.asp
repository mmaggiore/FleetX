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
        TempOrigination=Request.Form("TempOrigination")

		LocationCode=Request.Form("LocationCode")
		FakeSubmit=Request.Form("FakeSubmit")
        JobBillTo=Trim(Request.Form("JobBillTo"))
		If FakeSubmit="" then
			FakeSubmit=Request.QueryString("FakeSubmit")
		End if		
		PageStatus=Request.Form("PageStatus")
		txtJobNumber=Request.Form("txtJobNumber")
		Submit=Request.Form("Submit")
		BillToID=Request.Cookies("Phone")("sBT_ID")	
        'response.write "JobBillTo="&JobBillTo&"<BR>"
		Select Case JobBillTo
			Case "75"
				DisplayWord="BOL #"
				'Email="mark.maggiore@logisticorp.us"
			Case "80"
				DisplayWord="HAWB #"
				'Email="mark.maggiore@logisticorp.us;Les.Baron@Logisticorp.us"		
			Case "36"
				DisplayWord="Lot #"
				Email="g-dousharm1@ti.com"	
  			Case "38"
				DisplayWord="Reticle #"
				Email="j-charles2@ti.com"                              		
			Case else
				DisplayWord="Lot/Reticle #"
				'Email="mark.maggiore@logisticorp.us"       
			End Select	
           ' REsponse.write "email="&Email&"<BR>"


             If trim(TempOrigination)>"" then
				    Set Recordset1 = Server.CreateObject("ADODB.Recordset")
				    Recordset1.ActiveConnection = DATABASE
				    Recordset1.Source = "SELECT st_email FROM fcshipto where (st_id='"& TempOrigination &"')"
				    Recordset1.CursorType = 0
				    Recordset1.CursorLocation = 2
				    Recordset1.LockType = 1
				    Recordset1.Open()
				    Recordset1_numRows = 0
				    'Response.Write "Recordset1.Source="&Recordset1.Source&"<BR>"
				    If Recordset1.eof then
					    'ErrorMessage="Error on Page"
				    End if	
				    If Not Recordset1.eof then
					    tempemail=Recordset1("st_email")
                        if trim(tempemail)>"" then
                            Email=trim(Tempemail)
                        End if
				    End if	
				    Recordset1.Close()
				    Set Recordset1 = Nothing
            End if           
            	
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
            'Response.Write "JobBillTo="& JobBillTo &"<BR>"
			FakeSubmit="fakesubmit"
			If ExceptionID>"" then
				Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
					RSEVENTS2.Open "FCJobExceptions", DATABASE, 2, 2
					RSEVENTS2.addnew
					RSEVENTS2("fh_ID")=JobNumber
					RSEVENTS2("ExceptionID")=ExceptionID									
					RSEVENTS2("Ref_Num")=hawb		
					RSEVENTS2("BillToID") = trim(JobBillTo)
					RSEVENTS2("ExceptionTime")=Now()		
					RSEVENTS2("Status") = "c"
					RSEVENTS2.update
					RSEVENTS2.close			
				set RSEVENTS2 = nothing	
				Set Recordset1 = Server.CreateObject("ADODB.Recordset")
				Recordset1.ActiveConnection = DATABASE
                SQL908="SELECT ExceptionDescription FROM DriverExceptionList where (fh_bt_id='"& JobBillTo &"') and (Status='c') and (ExceptionID='"&ExceptionID&"')"
				Recordset1.Source = SQL908
				Recordset1.CursorType = 0
				Recordset1.CursorLocation = 2
				Recordset1.LockType = 1
				Recordset1.Open()
				Recordset1_numRows = 0
				'Response.Write "SQL908="&SQL908&"<BR>"
				If Recordset1.eof then
					ErrorMessage="Error on Page"
				End if	
				If Not Recordset1.eof then
                    'REsponse.write "I got here!<BR>"
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
					
					Set objMail = CreateObject("CDONTS.Newmail")
					objMail.From = "FleetX@LogisticorpGroup.com"
					objMail.To = Email
                    objMail.cc = "mark.maggiore@logisticorp.us"
					objMail.Subject = DisplayWord & " " & hawb &" Exception"
					objMail.MailFormat = cdoMailFormatMIME
					objMail.BodyFormat = cdoBodyFormatHTML
					objMail.Body = Body
					objMail.Send
					Set objMail = Nothing	
                    ErrorMessage="A message regarding this exception has been emailed out."
                    'Response.write "Email="&Email&"<BRL>"
                    'Response.write "Body="&Body&"<BRL>"
''''''''''''''''''''email notification END				
				'''''Response.Redirect("default.asp")				
				
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
            SomeDate=(now()-7)
			Set Recordset1 = Server.CreateObject("ADODB.Recordset")
			Recordset1.ActiveConnection = DATABASE
            xSQL = "SELECT fcfgthd.fh_id, fcfgthd.fh_bt_id, fclegs.fl_sf_id FROM fcfgthd INNER JOIN fcrefs ON fcfgthd.fh_id = fcrefs.rf_fh_id INNER JOIN fclegs ON fcrefs.rf_fh_id = fclegs.fl_fh_id WHERE (fcfgthd.fh_bt_id = '36' OR fcfgthd.fh_bt_id = '38') AND (fcrefs.rf_ref = '"& HAWB &"') AND (fcfgthd.fh_statcode <> '9') AND (fcfgthd.fh_statcode <> '99') AND (fcfgthd.fh_statcode <> '98') and ((fcrefs.ref_status <>'X') OR (ref_status is NULL)) AND (fcfgthd.fh_ship_dt>'"&SomeDate&"')"
            'response.write "xSQL="&xSQL&"<BR>"
			Recordset1.Source = xSQL
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
                JobBillTo=Recordset1("fh_bt_id")
                TempOrigination=Recordset1("fl_sf_id")
			End if
			Response.Write "</font>"
			Recordset1.Close()
			Set Recordset1 = Nothing
			If ErrorMessage="" then PageStatus="loggedin" End if
		End if	

 
 	
'------------------------------------------CODE TO REMOVE LOT/CANCEL JOB------------------------
XID=UserID
CancelLot=Trim(HAWB)
fh_id=JobNumber

 'response.write "GOT HERE, RIGHT BEFORE THE CODE!<BR>"  
 'response.write "CancelLot="&CancelLot&"<BR>"
 'response.write "errormessage="&errormessage&"<BR>"
 'response.write "AttemptedCancel="&AttemptedCancel&"<BR>"
 'response.write "ExceptionID="&ExceptionID&"<BR>"
 'response.write "JobNumber="&JobNumber&"<BR>"
 'response.write "fh_id="&fh_id&"<BR>"


If trim(CancelLot)>"" then AttemptedCancel="y" end if

If trim(CancelLot)>"" and trim(errormessage)>"" and AttemptedCancel="y" and trim(ExceptionID)>"" then
response.write "GOT HERE???????????????????<BR>"
 'response.write "GOT HERE, WITHIN THE CODE!<BR>" 
            '''''''''''CHECKS TO SEE IF MORE THAN ONE LOT LEFT...IF NOT, THEN IT CANCELS THE WHOLE JOB!!!!'''''''''''''''
	        Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
		        RSEVENTS.CursorLocation = 3
		        RSEVENTS.CursorType = 3
                'Response.write "DATABASE="&DATABASE&"<BR>"
		        RSEVENTS.ActiveConnection = DATABASE
		        'SQL = "SELECT fh_status, fh_ship_dt, fh_ready, fh_co_id, fh_co_phone, fh_co_email, fh_co_costcenter, fh_custpo, fh_priority FROM fcfgthd where (fh_id = '"& jobnum &"') AND (fh_bt_id='"& BillToID &"') and (fh_status='SCD')"
                SQL = "SELECT COUNT(fcrefs.rf_ref) AS HowManyValidRefs, fcfgthd.fh_status as JobStat FROM fcrefs INNER JOIN fcfgthd ON fcrefs.rf_fh_id = fcfgthd.fh_id where (rf_fh_id = '"& fh_id &"') AND ((ref_status<>'X') or (ref_status is NULL)) GROUP BY fcfgthd.fh_status"
		        'response.write "SQL111="&SQL&"<BR>"
		        RSEVENTS.Open SQL, DATABASE, 1, 3
               'if RSEVENTS.eof then
               '    jobnum=0
               'End if
                HowManyValidRefs=RSEVENTS("HowManyValidRefs")
                JobStat=RSEVENTS("JobStat")
                'Response.write "HowManyValidRefs="&HowManyValidRefs&"<BR>"
		        RSEVENTS.close
	        Set RSEVENTS = Nothing



            'response.write "JobStat="&jobstat&"<BR>"
            IF trim(JobStat)<>"ONB" then
            'response.write "GOT HERE!!! BOOOOOOOOO!<BR>"
            If HowManyValidRefs>1 then






			    Set oConn = Server.CreateObject("ADODB.Connection")
			    oConn.ConnectionTimeout = 100
			    oConn.Provider = "MSDASQL"
			    oConn.Open DATABASE
				    l_cSQL = "UPDATE fcrefs SET ref_status = 'X'" 
                    l_cSQL = l_cSQL&" WHERE  (rf_ref = '"& CancelLot &"') and rf_fh_id='"& fh_id & "'"
				    'response.write "l_cSQL222="&l_cSQL&"<BR>"
                    oConn.Execute(l_cSQL)
			    oConn.close
			    Set oConn = nothing
                '''''RESETTING THE ORDER STATUS TO NEW PER TI'S REQUEST 5/3/11
			    Set oConn = Server.CreateObject("ADODB.Connection")
			    oConn.ConnectionTimeout = 100
			    oConn.Provider = "MSDASQL"
			    oConn.Open DATABASE
				    l_cSQL = "UPDATE fcrefs SET ref_status = NULL" 
                    l_cSQL = l_cSQL&" WHERE (Ref_status<>'X') and rf_fh_id='"& fh_id & "'"
                    'response.write "l_cSQL333="&l_cSQL&"<BR>"
				    oConn.Execute(l_cSQL)
			    oConn.close
			    Set oConn = nothing
                'Response.write "l_cSQL="&l_cSQL&"<BR>"
			    Set oConn = Server.CreateObject("ADODB.Connection")
			    oConn.ConnectionTimeout = 100
			    oConn.Provider = "MSDASQL"
			    oConn.Open DATABASE
				    l_cSQL = "UPDATE Report_Refs SET rf_status = 'X'" 
                    l_cSQL = l_cSQL&" WHERE (rf_ref = '"& CancelLot &"') and rf_fh_id='"& fh_id & "'"
                    'response.write "l_cSQL444="&l_cSQL&"<BR>"
				    oConn.Execute(l_cSQL)
			    oConn.close
			    Set oConn = nothing
			    Set oConn = Server.CreateObject("ADODB.Connection")
			    oConn.ConnectionTimeout = 100
			    oConn.Provider = "MSDASQL"
			    oConn.Open DATABASE
				    l_cSQL = "UPDATE Report_Refs SET rf_status = NULL" 
                    l_cSQL = l_cSQL&" WHERE (Rf_status<>'X') and rf_fh_id='"& fh_id & "'"
                    'response.write "l_cSQL555="&l_cSQL&"<BR>"
				    oConn.Execute(l_cSQL)
			    oConn.close
			    Set oConn = nothing
                'Response.write "l_cSQL="&l_cSQL&"<BR>"
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			    Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				    RSEVENTS2.Open "CancelledOrders", DATABASE, 2, 2
				    RSEVENTS2.addnew
				    RSEVENTS2("XID")=XID
				    if trim(fh_id)>"" then
                        RSEVENTS2("fh_id")=fh_id
                        Else
                        RSEVENTS2("fh_id")=CancelJob
                    End if
                    RSEVENTS2("fh_ref")=CancelLot
				    RSEVENTS2("Reason")=Reason
 				    RSEVENTS2("OtherReason")=OtherReason
				    RSEVENTS2("CancelDate")=now()                   	
				    RSEVENTS2("CancelStatus")="c"								
				    RSEVENTS2.update
				    RSEVENTS2.close			
			    set RSEVENTS2 = nothing
                '''''''''''''''''''GATHERS THE PREVIOUS DUE TIME IN MINUTES TO RESET THE VALUE IF ORDER ISN'T COMPLETELY CANCELLED!
	            Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
		            RSEVENTS.CursorLocation = 3
		            RSEVENTS.CursorType = 3
                    'Response.write "DATABASE="&DATABASE&"<BR>"
		            RSEVENTS.ActiveConnection = DATABASE
		            'SQL = "SELECT fh_status, fh_ship_dt, fh_ready, fh_co_id, fh_co_phone, fh_co_email, fh_co_costcenter, fh_custpo, fh_priority FROM fcfgthd where (fh_id = '"& jobnum &"') AND (fh_bt_id='"& BillToID &"') and (fh_status='SCD')"
                    SQL = "SELECT fcfgthd.fh_ship_dt AS BookTime, fcfgthd.fh_status, fclegs.fl_st_rta AS OriginalDueTime FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id WHERE (fh_id='"& fh_id &"') AND (fclegs.fl_leg_status = 'c')"
		            'response.write "SQL666="&SQL&"<BR>"
		            RSEVENTS.Open SQL, DATABASE, 1, 3
                   'if RSEVENTS.eof then
                   '    jobnum=0
                   'End if
                    BookTime=RSEVENTS("BookTime")
                    OriginalStatus=RSEVENTS("fh_status")
                    OriginalDueTime=RSEVENTS("OriginalDueTime")
                    OriginalMinutesToShip=DateDiff("n", BookTime, OriginalDueTime)
                    NewDueTime=DateAdd("n", OriginalMinutesToShip, now())

                    'Response.write "******************************************<BR>"
                    'Response.write "OriginalDueTime="&OriginalDueTime&"<BR>"
                    'Response.write "BookTime="&BookTime&"<BR>"
                    'Response.write "OriginalMinutesToShip="&OriginalMinutesToShip&"<BR>"
                    'Response.write "NewDueTime="&NewDueTime&"<BR>"
                    'Response.write "******************************************<BR>"

		            RSEVENTS.close
	            Set RSEVENTS = Nothing
''''''''''''''Continues the changing of the job status to new, per TI 5/3/11'''''''''''''''''
            If (trim(OriginalStatus)="ACC" or trim(OriginalStatus)="OPN") AND SomeVar="Itwillneverequalthisvar" then
			    Set oConn = Server.CreateObject("ADODB.Connection")
			    oConn.ConnectionTimeout = 100
			    oConn.Provider = "MSDASQL"
			    oConn.Open DATABASE
				    l_cSQL = "UPDATE fcfgthd SET fh_status='ACC', fh_statcode='4'" 
                    l_cSQL = l_cSQL&" WHERE (fh_id = '"& fh_id &"')"
                    'response.write "l_cSQL777="&l_cSQL&"<BR>"
				    oConn.Execute(l_cSQL)
			    oConn.close
			    Set oConn = nothing

			    'Set oConn = Server.CreateObject("ADODB.Connection")
			    'oConn.ConnectionTimeout = 100
			    'oConn.Provider = "MSDASQL"
			    'oConn.Open DATABASE
				'    l_cSQL = "UPDATE fclegs SET fl_st_rta='"& NewDueTime &"'" 
                '    l_cSQL = l_cSQL&" WHERE (fl_fh_id = '"& fh_id &"')"
                '    Response.write "l_cSQL888="&l_cSQL&"<BR>"
				'    oConn.Execute(l_cSQL)
			    'oConn.close
			    'Set oConn = nothing

			    'Set oConn = Server.CreateObject("ADODB.Connection")
			    'oConn.ConnectionTimeout = 100
			    'oConn.Provider = "MSDASQL"
			    'oConn.Open DATABASE
				'    l_cSQL = "UPDATE report_data SET fh_status='ACC', fl_st_rta='"& NewDueTime &"'" 
                '    l_cSQL = l_cSQL&" WHERE (fh_id = '"& fh_id &"')"
                '    Response.write "l_cSQL999="&l_cSQL&"<BR>"
				'    oConn.Execute(l_cSQL)
			    'oConn.close
			    'Set oConn = nothing
        End if
                SuccessMessage="You have successfully cancelled lot/reticle #"&CancelLot&"<BR>"
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ELSE
			    Set oConn = Server.CreateObject("ADODB.Connection")
			    oConn.ConnectionTimeout = 100
			    oConn.Provider = "MSDASQL"
			    oConn.Open DATABASE
				    l_cSQL = "UPDATE fcrefs SET ref_status = 'X'" 
                    l_cSQL = l_cSQL&" WHERE (rf_ref = '"& CancelLot &"') and rf_fh_id='"& fh_id & "'"
				    'response.write "l_cSQL222="&l_cSQL&"<BR>"
                    oConn.Execute(l_cSQL)
			    oConn.close
			    Set oConn = nothing
			    Set oConn = Server.CreateObject("ADODB.Connection")
			    oConn.ConnectionTimeout = 100
			    oConn.Provider = "MSDASQL"
			    oConn.Open DATABASE
				    l_cSQL = "UPDATE Report_Refs SET rf_status = 'X'" 
                    l_cSQL = l_cSQL&" WHERE (rf_ref = '"& CancelLot &"') and rf_fh_id='"& fh_id & "'"
                    'response.write "l_cSQL444="&l_cSQL&"<BR>"
				    oConn.Execute(l_cSQL)
			    oConn.close
			    Set oConn = nothing
			    Set oConn = Server.CreateObject("ADODB.Connection")
			    oConn.ConnectionTimeout = 100
			    oConn.Provider = "MSDASQL"
			    oConn.Open DATABASE
				    l_cSQL = "UPDATE fcfgthd SET fh_status='CAN', fh_statcode='98'" 
                    l_cSQL = l_cSQL&" WHERE (fh_id = '"& fh_id &"')"
                    'response.write "l_cSQL101010="&l_cSQL&"<BR>"
				    oConn.Execute(l_cSQL)
			    oConn.close
			    Set oConn = nothing
			    Set oConn = Server.CreateObject("ADODB.Connection")
			    oConn.ConnectionTimeout = 100
			    oConn.Provider = "MSDASQL"
			    oConn.Open DATABASE
				    l_cSQL = "UPDATE Report_Data SET fh_status='CAN'" 
                    l_cSQL = l_cSQL&" WHERE (fh_id = '"& fh_id &"')"
                    'response.write "l_cSQL111111="&l_cSQL&"<BR>"
				    oConn.Execute(l_cSQL)
			    oConn.close
			    Set oConn = nothing
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			    Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				    RSEVENTS2.Open "CancelledOrders", DATABASE, 2, 2
				    RSEVENTS2.addnew
				    RSEVENTS2("XID")=XID
				    RSEVENTS2("fh_id")=CancelJob
                    RSEVENTS2("fh_ref")=CancelLot
				    RSEVENTS2("Reason")=Reason
 				    RSEVENTS2("OtherReason")=OtherReason
				    RSEVENTS2("CancelDate")=now()                   	
				    RSEVENTS2("CancelStatus")="c"								
				    RSEVENTS2.update
				    RSEVENTS2.close			
			    set RSEVENTS2 = nothing
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                SuccessMessage="You have successfully cancelled job #"&CancelJob&"<BR>"
            End if
            End if
    'Response.write "Put cancel lot code here...don't forget if only one lot to cancel the whole job instead<br>"
        XID=""
        Reason=""
        OtherReason=""
        AttemptedCancel="n"
        CancelLot=""
        CancelJob=""
End if


If trim(CancelJob)>"" and trim(errormessage)>"" and AttemptedCancel="y" and trim(ExceptionID)>"" then
    'Response.write "Put cancel job code here!<br>"
			Set oConn = Server.CreateObject("ADODB.Connection")
			oConn.ConnectionTimeout = 100
			oConn.Provider = "MSDASQL"
			oConn.Open DATABASE
				l_cSQL = "UPDATE fcfgthd SET fh_status = 'CAN', fh_statcode='98'" 
                l_cSQL = l_cSQL&" WHERE (fh_id = '"& CancelJob &"')"
                'response.write "l_cSQL="&l_cSQL&"<BR>"
				oConn.Execute(l_cSQL)
			oConn.close
			Set oConn = nothing
			Set oConn = Server.CreateObject("ADODB.Connection")
			oConn.ConnectionTimeout = 100
			oConn.Provider = "MSDASQL"
			oConn.Open DATABASE
				l_cSQL = "UPDATE Report_Data SET fh_status = 'CAN'" 
                l_cSQL = l_cSQL&" WHERE (fh_id = '"& CancelJob &"')"
                'response.write "l_cSQL="&l_cSQL&"<BR>"
				oConn.Execute(l_cSQL)
			oConn.close
			Set oConn = nothing
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				RSEVENTS2.Open "CancelledOrders", DATABASE, 2, 2
				RSEVENTS2.addnew
				RSEVENTS2("XID")=XID
				RSEVENTS2("fh_id")=CancelJob
                RSEVENTS2("fh_ref")=CancelLot
				RSEVENTS2("Reason")=Reason
 				RSEVENTS2("OtherReason")=OtherReason
				RSEVENTS2("CancelDate")=now()                   	
				RSEVENTS2("CancelStatus")="c"								
				RSEVENTS2.update
				RSEVENTS2.close			
			set RSEVENTS2 = nothing
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
           SuccessMessage="You have successfully cancelled job #"&CancelJob&"<BR>"
            XID=""
            Reason=""
            OtherReason=""
            AttemptedCancel="n"
            CancelLot=""
            CancelJob=""
End if

'-----------------------------------------END CODE TO REMOVE LOT/CANCEL JOB---------------------


		

	
		
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
					<form method="post" action="DriverExceptions_IFAB.asp">
						<tr>
							<td align="center" class="purpleseparator" colspan="13"><b>POSSIBLE EXCEPTIONS</b></td>
						</tr>
						<%
						Set Recordset1 = Server.CreateObject("ADODB.Recordset")
						Recordset1.ActiveConnection = DATABASE
                        SQL34="SELECT ExceptionDescription, ExceptionID FROM DriverExceptionList where (fh_bt_id='"&JobBillTo&"') and (Status='c')"
                        'Response.write "SQL34="&SQL34&"<BR>"
						Recordset1.Source = SQL34
						Recordset1.CursorType = 0
						Recordset1.CursorLocation = 2
						Recordset1.LockType = 1
						Recordset1.Open()
						Recordset1_numRows = 0
						'Response.Write "Recordset1.Source="&Recordset1.Source&"<BR>"
						If Recordset1.eof then
							ErrorMessage="There are no available exceptions"
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
                        <input type="hidden" name="JobBillTo" value="<%=JobBillTo%>" ID="Hidden3">
						<input type="hidden" name="JobNumber" value="<%=JobNumber%>" ID="Hidden2">
                        <input type="hidden" name="TempOrigination" value="<%=TempOrigination %>" />
						<tr><td align="center" colspan="3"><input type="submit" name="submit" value="submit"></td></tr>
					</table>
					</form>
					<%				
				Case Else			
			%>
			<FORM ACTION="DriverExceptions_IFAB.asp" method="post" name="thisForm" ID="thisForm">
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
