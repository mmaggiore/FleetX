<html>
<head>
<!-- #include file="../fleetexpress.inc" -->
<link rel="stylesheet" type="text/css" href="../css/Style.css">


<%



    ColorSelect=Request.form("ColorSelect")
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
    PageTitle="FLEETX JOB MANAGEMENT"


'response.write "******************USERID="&USERID&"<BR>"
JobManagementWafer=Session("JobManagementWafer")
'response.write "JobManagementWafer="&JobManagementWafer&"***<BR>"
JobManagementSR=Session("JobManagementSR")
'response.write "JobManagementSR="&JobManagementSR&"***<BR>"
JobManagementKWE=Session("JobManagementKWE")
'response.write "JobManagementKWE="&JobManagementKWE&"***<BR>"

'''UserID=Request.Form("UserID")
'''IF trim(UserID)="" then
'''	UserID=session("UserID")
	'Response.Write "WHOA!  USER ID="&UserID&"***********<BR>"
'''End if
'response.write "UserID="&UserID&"***<BR>"
Page=Request.form("page")
If Page="" then
	Page=Request.querystring("Page")
end if
'response.write "Page="&Page&"***<BR>"
'SearchVariable=Request.querystring("SearchVariable")
'response.write "SearchVariable="&SearchVariable&"***<BR>"
OriginalJobStatus=Request.Form("OriginalJobStatus")
'response.write "OriginalJobStatus="&OriginalJobStatus&"***<BR>"
WhichLeg=Request.Form("WhichLeg")
'response.write "WhichLeg="&WhichLeg&"***<BR>"
PageStatus=request.form("PageStatus")
'if PageStatus="" then
'	PageStatus=Request.Querystring("PageStatus")
'end if
'response.write "PageStatus="&PageStatus&"***<BR>"
'SupervisorID=Request.Form("SupervisorID")
'response.write "SupervisorID="&SupervisorID&"***<BR>"
'''''Customer=Request.Form("Customer")
lmnop=Request.Form("lmnop")
'if trim(customer)="" then
'	Customer=Session("sBT_ID")
'End if
'response.write "Customer="&Customer&"***<BR>"

'''''Select Case Customer
'''''	Case "kwe"
'''''		Session("sBT_ID")="48"
'''''	Case "tiwf"
'''''		Session("sBT_ID")="36"
'''''	Case "tisr"
'''''		Session("sBT_ID")="26"
'''''End Select
BillToID=Session("sBT_ID")
'response.write "BillToID="&BillToID&"***<BR>"
Submit=Request.Form("Submit")
'response.write "Submit="&Submit&"***<BR>"
JobNumber=trim(Request.Form("JobNumber"))
fh_id=JobNumber
'response.write "JobNumber="&JobNumber&"***<BR>"
RefNumber=trim(Request.Form("RefNumber"))
'response.write "RefNumber="&RefNumber&"***<BR>"
'response.write "fh_id="&fh_id&"***<BR>"
jobstatus=Request.Form("jobstatus")
'response.write "jobstatus="&jobstatus&"***<BR>"
'response.write "************************8JobStatus="&JobStatus&"***<BR>"
Fh_statcode=JobStatus
'response.write "Fh_statcode="&Fh_statcode&"***<BR>"
Fh_Custpo=Request.Form("Fh_Custpo")
DisplayPOD=trim(Request.Form("DisplayPOD"))
fl_t_disp=trim(Request.form("fl_t_disp"))
Leg_fl_firstdrop=Request.form("fl_firstdrop")
Leg_fl_seconb=Request.form("fl_seconb")
Leg_fl_secacc=Request.form("fl_secacc")
'response.write "Fh_Custpo="&Fh_Custpo&"***<BR>"
'al_ca_id=Request.Form("al_ca_id")
'response.write "al_ca_id="&al_ca_id&"***<BR>"
'al_trackno=Request.Form("al_trackno")
'response.write "al_trackno="&al_trackno&"***<BR>"
'al_st_ohd=Request.Form("al_st_ohd")
'response.write "al_st_ohd="&al_st_ohd&"***<BR>"
DisplayCategoryID=Request.Form("CategoryID")
'response.write "DisplayCategoryID="&DisplayCategoryID&"***<BR>"
leg_fl_t_acc_date=Request.Form("leg_fl_t_acc_date")
leg_fl_t_acc_time=Request.Form("leg_fl_t_acc_time")
leg_fl_t_acc=leg_fl_t_acc_date&" "&leg_fl_t_acc_time
'response.write "LINE 113 fl_t_acc="&fl_t_acc&"***<BR>"
leg_fl_t_int=Request.Form("leg_fl_t_int")
'response.write "fl_t_int="&fl_t_int&"***<BR>"
leg_fl_t_und=Request.Form("leg_fl_t_und")
'response.write "fl_t_und="&fl_t_und&"***<BR>"
leg_fl_t_atp_date=Request.Form("leg_fl_t_atp_date")
leg_fl_t_atp_time=Request.Form("leg_fl_t_atp_time")
leg_fl_t_atp=leg_fl_t_atp_date&" "&leg_fl_t_atp_time

'response.write "fl_t_atp="&fl_t_atp&"***<BR>"
leg_fl_t_atd_date=Request.Form("leg_fl_t_atd_date")
leg_fl_t_atd_time=Request.Form("leg_fl_t_atd_time")
leg_fl_t_atd=leg_fl_t_atd_date&" "&leg_fl_t_atd_time

'response.write "11111leg_fl_t_atd="&leg_fl_t_atd&"***<BR>"
fl_sf_id=Request.Form("fl_sf_id")
'response.write "fl_sf_id="&fl_sf_id&"***<BR>"
fl_st_id=Request.Form("fl_st_id")
'response.write "fl_st_id="&fl_st_id&"***<BR>"
Fh_Priority=Request.Form("Fh_Priority")
'response.write "Fh_Priority="&Fh_Priority&"***<BR>"
'response.write "fh_co_id="&fh_co_id&"***<BR>"
'fh_user5=Request.Form("fh_user5")
'response.write "fh_user5="&fh_user5&"***<BR>"
fl_sf_rta=Request.Form("fl_sf_rta")
'response.write "fl_sf_rta="&fl_sf_rta&"***<BR>"
fh_ship_dt=Request.Form("fh_ship_dt")
'response.write "fh_ship_dt="&fh_ship_dt&"***<BR>"
'response.write "fh_ready="&fh_ready&"***<BR>"
fl_st_rta=Request.Form("fl_st_rta")	
fh_ready=Request.Form("fh_ready")	
'response.write "fl_st_rta="&fl_st_rta&"***<BR>"
addedPOD=Request.Form("addedPOD")
fh_priority=Request.Form("fh_priority")
fh_co_id=Request.Form("fh_co_id")
fh_user5=Request.Form("fh_user5")
'response.write "addedPOD="&addedPOD&"***<BR>"	
PODID=Request.Form("TempPODID")
'response.write "PODID="&PODID&"***<BR>"
If trim(DisplayPOD)="" then
	DisplayPOD=PODID
End if
'response.write "DisplayPOD="&DisplayPOD&"***<BR>"
'response.write "************PODID="&PODID&"<BR>"
'leg_fl_t_acc=Request.Form("leg_fl_t_acc")
'response.write "LINE 152 leg_fl_t_acc="&leg_fl_t_acc&"***<BR>"
'leg_fl_t_int=Request.Form("leg_fl_t_int")
'response.write "leg_fl_t_int="&leg_fl_t_int&"***<BR>"
'leg_fl_t_und=Request.Form("leg_fl_t_und")
'response.write "leg_fl_t_und="&leg_fl_t_und&"***<BR>"
'leg_fl_t_atp=Request.Form("leg_fl_t_atp")
'response.write "leg_fl_t_atp="&leg_fl_t_atp&"***<BR>"
'leg_fl_t_atd=Request.Form("leg_fl_t_atd")
'response.write "22222leg_fl_t_atd="&leg_fl_t_atd&"***<BR>"
'response.write "leg_fl_t_atd="&leg_fl_t_atd&"***<BR>"
tempPODID=Request.Form("tempPODID")
'response.write "tempPODID="&tempPODID&"***<BR>"
AddedPOD=Request.Form("AddedPOD")
'response.write "AddedPOD="&AddedPOD&"***<BR>"
Display_Leg_FL_Leg_Status=Request.Form("Display_Leg_FL_Leg_Status")
'response.write "Display_Leg_FL_Leg_Status="&Display_Leg_FL_Leg_Status&"***<BR>"
CategoryID=Request.Form("CategoryID")
'response.write "CategoryID="&CategoryID&"***<BR>"

PODChange=Request.Form("PODChange")
reasonforchange=Request.Form("reasonforchange")
reasonforchange=Replace(reasonforchange,"""","`")
reasonforchange=Replace(reasonforchange,"'","`")
'response.write "reasonforchange="&reasonforchange&"***<BR>"
PageStatus=Request.Form("PageStatus")
'response.write "PageStatus="&PageStatus&"***<BR>"
AdminNote=Request.Form("AdminNote")
SQLExceptionID=Request.Form("SQLExceptionID")
ManagerNote=Request.form("ManagerNote")





If Jobnumber="" AND RefNumber="" AND Submit>"" then
	''response.write "GOT HERE!!!"
	ErrorMessage="You MUST provide a Job Number or Reference Number"
End if
UserFirstName=Request.Form("UserFirstName")
'response.write "UserFirstName="&UserFirstName&"***<BR>"
LastName=Request.Form("LastName")
'response.write "LastName="&LastName&"***<BR>"
Leg_fl_counter=request.Form("Leg_fl_counter")
Submitted_Leg_fl_counter=Leg_fl_counter
'response.write "Leg_fl_counter="&Leg_fl_counter&"***<BR>"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''
        Set oConn = Server.CreateObject("ADODB.Connection")
        oConn.ConnectionTimeout = 100
        oConn.Provider = "MSDASQL"
        oConn.Open DATABASE
		        SQL="SELECT fh_bt_id FROM fcfgthd  where (fh_id='"& JobNumber &"')"
	        'Response.Write "SQL="&SQL&"<BR>"
	        SET oRs = oConn.Execute(Sql)
	        Do while not oRs.EOF 
	            'Response.Write "got here...okay?" 
			    'Response.Write "FABID="&FABID&"<BR>"
			    'Response.Write "SQL="&SQL&"<BR>"
                TempvarBT_ID=trim(oRs("fh_bt_id"))
                'Response.Write "TempvarBT_ID="&TempvarBT_ID&"<BR>"
                ' Response.Write "AllRefs="&AllRefs&"<BR>"
			oRs.movenext
			LOOP                        
            oRs.Close
		    Set oRs=Nothing
''''''''''''''''''''''''''''''JOB CANCELLATION
If trim(submit)="Cancel Job" then
	If DisplayCategoryID="" then
		ErrorMessage="You must select a category for why you are cancelling this order"
	End if	
    If trim(ReasonForChange)="" then
        ErrorMessage="You must provide a reason for cancellation"
    End if



			If Trim(ErrorMessage)="" then
                'Response.write "GOT HERE!!!!<BR>"
                TempLeg_fl_st_id=Request.Form("Leg_fl_st_id")
                TempBTID=Session("sBT_ID")
				'Response.write "JobNumber="&jobnumber&"<BR>"
                'response.write "TempLeg_fl_st_id="&TempLeg_fl_st_id&"<BR>"
                'Response.write "TempBTID="&TempBTID&"<BR>"
                'Response.write "ReasonForChange="&ReasonForChange&"<BR>"
				'Response.write "GOT HERE!!!GOT HERE!!!<BR>"
                Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
					RSEVENTS2.Open "CancelledOrders", Database, 2, 2
					RSEVENTS2.addnew	
					RSEVENTS2("XID")=UserID		
					RSEVENTS2("fh_id") = JobNumber
					RSEVENTS2("fh_ref")=""
                    RSEVENTS2("Reason")=DisplayCategoryID
                    RSEVENTS2("OtherReason")=ReasonForChange
                    RSEVENTS2("CancelDate")=Now()	
					RSEVENTS2("CancelStatus") = "c"
					RSEVENTS2.update
					RSEVENTS2.close			
				set RSEVENTS2 = nothing	

			Set oConn = Server.CreateObject("ADODB.Connection")
			oConn.ConnectionTimeout = 100
			oConn.Provider = "MSDASQL"
			oConn.Open DATABASE
			''''UPDATES THE WAFER
            'Response.write "GOT HERE!!!  UPDATED THE WAFER!!!! LINE 525!!!<BR>"
			l_cSQL = "UPDATE FCFGTHD SET fh_status = 'CAN', fh_statcode='98' "
			If fh_custpo>"" then
				l_cSQL = l_cSQL & " , fh_custpo='"&fh_custpo&"' "
			End if
			l_cSQL = l_cSQL & " WHERE (fh_id = '"&JobNumber&"')"
			'response.write "UPDATE Wafers="&l_cSQL&"<BR>"
			oConn.Execute(l_cSQL)
			Set oConn=Nothing
'''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''
			Set oConn = Server.CreateObject("ADODB.Connection")
			oConn.ConnectionTimeout = 100
			oConn.Provider = "MSDASQL"
			oConn.Open DATABASE
			''''UPDATES THE WAFER
			l_cSQL = "UPDATE REPORT_DATA SET fh_status = 'CAN'"
			l_cSQL = l_cSQL & " WHERE (fh_id = '"&JobNumber&"')"
            'response.write "UPDATE Wafers="&l_cSQL&"<BR>"
			oConn.Execute(l_cSQL)
			Set oConn=Nothing			
'''''''''''''''''''''''''''''''''	

'------------DRIVER MESSAGE ABOUT CANCELLATION

            'Response.write "GOT HERE 6!!!<BR>"

	        Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
		        RSEVENTS.CursorLocation = 3
		        RSEVENTS.CursorType = 3
                'Response.write "DATABASE="&DATABASE&"<BR>"
		        RSEVENTS.ActiveConnection = DATABASE
		        'SQL = "SELECT fh_status, fh_ship_dt, fh_ready, fh_co_id, fh_co_phone, fh_co_email, fh_co_costcenter, fh_custpo, fh_priority FROM fcfgthd where (fh_id = '"& jobnum &"') AND (fh_bt_id='"& BillToID &"') and (fh_status='SCD')"
                SQL = "SELECT fl_dr_id, fl_sf_cfname, fl_sf_building, fl_sf_addr1, fl_sf_addr2 FROM fclegs where (fl_fh_id = '"& fh_id &"')"
		        
                'Response.Write "SQL="&SQL&"<BR>"
		        
                RSEVENTS.Open SQL, DATABASE, 1, 3
               'if RSEVENTS.eof then
               '    jobnum=0
               'End if
               If not RSEVENTS.eof then

                'Response.write "GOT HERE #7<BR>"

                TheDriverID=RSEVENTS("fl_dr_id")
                fl_sf_cfname=RSEVENTS("fl_sf_cfname")
                fl_sf_building=RSEVENTS("fl_sf_building")
                fl_sf_addr1=RSEVENTS("fl_sf_addr1")
                fl_sf_addr2=RSEVENTS("fl_sf_addr2")

                'rESPONSE.WRITE "TheDriverID="&TheDriverID&"<BR>"

	            Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
		            RSEVENTS2.Open "DriverMessages", INTRANET, 2, 2
		            RSEVENTS2.addnew
		            RSEVENTS2("DriverMessage")="Job #"&fh_id&" from "& fl_sf_cfname &" "& fl_sf_building &" "& fl_sf_addr1 &" "& fl_sf_addr2 &" has been cancelled."
		            RSEVENTS2("MessageDate")=Date()
		            RSEVENTS2("MessageOriginator")="1"
                    RSEVENTS2("MessageRecipient")=TheDriverID
		            RSEVENTS2("MessageStatus")="c"								
		            RSEVENTS2.update
		            RSEVENTS2.close			
	            set RSEVENTS2 = nothing	

                'Response.write "THE END!!!!<BR>"
                
                End if
		        RSEVENTS.close
	        Set RSEVENTS = Nothing
'------END DRIVER CANCELLATION MESSAGE





				Set Recordset166 = Server.CreateObject("ADODB.Recordset")
				'Response.Write "Intranet="&Intranet&"***<BR>"
				Recordset166.ActiveConnection = Database
				'TempSQL="SELECT RequestorEmail AS SendToEmail FROM PreExistingRequestor "
                'TempSQL=TempSQL & "where requestorid='"&TempLeg_fl_st_id&"'"

                TempSQL="SELECT PreExistingRequestor.RequestorEmail AS SendToEmail FROM PreExistingCompanies INNER JOIN PreExistingRequestor ON PreExistingCompanies.CompanyOwner = PreExistingRequestor.RequestorID "
                TempSQL=TempSQL & "where st_id='"&TempLeg_fl_st_id&"'"
                'response.write "TempSQL="&TempSQL&"<BR>"
                
                
                Recordset166.Source = TempSQL
				

                

                Recordset166.CursorType = 0
				Recordset166.CursorLocation = 2
				Recordset166.LockType = 1
				Recordset166.Open()
				Recordset166_numRows = 0

                

					if NOT Recordset166.EOF then
						SendToEmail=Recordset166("SendToEmail")
						'Response.write "SendToEmail="&SendToEmail&"<BR>"
                  
                        Else

					End if
					Recordset166.Close()
					Set Recordset166 = Nothing	


                    Set oConn = Server.CreateObject("ADODB.Connection")
                    oConn.ConnectionTimeout = 100
                    oConn.Provider = "MSDASQL"
                    oConn.Open DATABASE
		                 SQL="SELECT * FROM fcrefs INNER JOIN fcfgthd ON fcrefs.rf_fh_id = fcfgthd.fh_id where (fh_id='"& JobNumber &"')"
	                    'Response.Write "SQL="&SQL&"<BR>"
	                    SET oRs = oConn.Execute(Sql)
	                    Do while not oRs.EOF 
	                        'Response.Write "got here...okay?" 
			                'Response.Write "FABID="&FABID&"<BR>"
			                'Response.Write "SQL="&SQL&"<BR>"
                            varBT_ID=trim(oRs("fh_bt_id"))
                            temp_ref=trim(oRs("RF_Ref"))
                            AllRefs=AllRefs & "#" & Temp_ref & "<br>"
                           'Response.Write "varBT_ID="&varBT_ID&"<BR>"
                           ' Response.Write "AllRefs="&AllRefs&"<BR>"
				        oRs.movenext
				        LOOP                        
                        oRs.Close
		                Set oRs=Nothing
                        'SendToEmail="mark.maggiore@logisticorp.us"
                        'Response.write "SENDTOEMAIL LINE 494<BR>"
                        If trim(SendToEmail)>"" then
						    Body = "FleetX job #" & JobNumber & " has just been cancelled by a supervisor.<br><br>"& _
						    "The reason the job was cancelled is:  "&ReasonForChange&".<br><br>"& _
                            "If you still need this shipment, then you will have to re-place your order..<br><br>Sincerely<br><br><br><br>"& _
						    "FleetX" 
						    'Recipient = "mark.maggiore@logisticorp.us"
                            'response.write "body="&body&"<BR>"
						    'Set objMail = CreateObject("CDONTS.Newmail")
						    'objMail.From = "FleetX@LogisticorpGroup.com"
						    varTo = SendToEmail
                            varcc = "Mark.Maggiore@Logisticorp.us;FleetX@LogisticorpGroup.com"

						    varSubject = "FleetX Order Cancellation"
						    'objMail.MailFormat = cdoMailFormatMIME
						    'objMail.BodyFormat = cdoBodyFormatHTML
						    'objMail.Body = Body
						    'objMail.Send
						    'Set objMail = Nothing
     '''''''''''''''''''''''''''''''''''''''''''''''''''
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
                        End if
            End if


End if
If trim(submit)="Update Leg" then
	'''response.write "THIS IS THE UPDATE FUNCTION<BR>"
	If JobStatus="" then
		ErrorMessage="You must provide a job status"
	End if
    'response.write "fl_t_int="&fl_t_int&"<BR>"	
	'If not IsDate(fl_t_atd) or (fl_t_atd<>"1/1/1900" AND len(fl_t_atd)<11) Then ErrorMessage="Drop Time is not a valid date/time" end if
	'If not IsDate(fl_t_atp) or (fl_t_atp<>"1/1/1900" AND len(fl_t_atp)<11) Then ErrorMessage="On Board Time is not a valid date/time" end if
	'If not IsDate(fl_t_acc) or (fl_t_acc<>"1/1/1900" AND len(fl_t_acc)<11) Then ErrorMessage="Acknowledge Time is not a valid date/time" end if
	'If not IsDate(fl_t_und) or (fl_t_und<>"1/1/1900" AND len(fl_t_und)<11) AND BillToID="48" Then ErrorMessage="Arrived at Airline Time is not a valid date/time" end if
	'If not IsDate(fl_t_int) or (fl_t_int<>"1/1/1900" AND len(fl_t_int)<11) AND BillToID="48" Then ErrorMessage="Paperwork On Board Time is not a valid date/time" end if
	If DisplayCategoryID="" then
		ErrorMessage="You must select a category for the change of this order"
	End if	
    If trim(ReasonForChange)="" then
        ErrorMessage="You must provide a reason for change"
    End if
	If ErrorMessage="" then
    'REsponse.write "LINE:  372 Leg_fl_t_atp="&Leg_fl_t_atp&"<BR>"
    If trim(Leg_fl_t_acc)>"" and Leg_fl_t_acc<>"1/1/1900 12:00:00 AM" then
        'Response.write "Got here!  ACC  Line 374<BR>"
        fh_statcode="4"
    End if
    If trim(Leg_fl_t_atp)>"" and Leg_fl_t_atp<>"1/1/1900 12:00:00 AM" then
        'Response.write "Got here!  ONB  Line 378<BR>"
        fh_statcode="5"
    End if
    If trim(Leg_fl_t_atd)>"" and Leg_fl_t_atd<>"1/1/1900 12:00:00 AM" then
        'Response.write "Got here!  CLS  Line 382<BR>"
        fh_statcode="9"
    End if


		''response.write "*****THIS IS THE UPDATE FUNCTION!!!!!!<BR>"
		Select case fh_statcode
			Case "5"
				REF_STATUS="o"
				StatusWord="ONB"
                FH_Status="ONB"
			Case "9"
				REF_STATUS="c"
				StatusWord="CLS"
                FH_Status="CLS"
			Case "13"
				REF_STATUS="p"
				StatusWord="PUO"
                FH_Status="PUO"
			Case "2"
				REF_STATUS=NULL
				StatusWord="RAP"
                FH_Status="RAP"
			Case "3"
				REF_STATUS=NULL	
				StatusWord="OPN"
                FH_Status="OPN"
			Case "98"
				REF_STATUS="x"	
				StatusWord="CAN"
                FH_Status="CAN"
			Case "4"
				REF_STATUS=NULL	
				StatusWord="ACC"
                FH_Status="ACC"
			Case else
				REF_STATUS=NULL
				StatusWord="???"
																		
		End Select
		xxx="yes"	
		'response.write "Did I get here??????????????????<BR>"
		If OriginalJobStatus<>JobStatus then
			'response.write "GOT HERE!!!!!!!!!!!!!!!!!!!!!<BR>"
		End if
		'If xxx="yes" AND OriginalJobStatus<>JobStatus AND trim(ErrorMessage)="" then
        If xxx="yes" AND trim(ErrorMessage)="" then
            'Response.write "got here!<br>"



'Response.write "AdminNote="&AdminNote&"<BR>"
'Response.write "SQLExceptionID="&SQLExceptionID&"<BR>"
'Response.write "ManagerNote="&ManagerNote&"<BR>"
ManagerNote=AdminNote
fh_bt_id=TempvarBT_ID
InputJobNumber=jobnumber
'Response.write "SQLExceptionID="&SQLExceptionID&"<BR>"
'Response.write "UserID="&UserID&"<BR>"
'Response.write "InputJobNumber="&InputJobNumber&"<BR>"
'Response.write "fh_bt_id="&fh_bt_id&"<BR>"
'Response.write "ManagerNote="&ManagerNote&"<BR>"
'''''''''''''''''''''''''''''''''
    If trim(SQLExceptionID)>"" then 
    
        Set oConn = Server.CreateObject("ADODB.Connection")
        oConn.ConnectionTimeout = 100
        oConn.Provider = "MSDASQL"
        oConn.Open DATABASE
		        SQL="SELECT fh_co_email FROM fcfgthd  where (fh_id='"& JobNumber &"')"
	        'Response.Write "SQL="&SQL&"<BR>"
	        SET oRs = oConn.Execute(Sql)
	        Do while not oRs.EOF 
	            'Response.Write "got here...okay?" 
			    'Response.Write "FABID="&FABID&"<BR>"
			    'Response.Write "SQL="&SQL&"<BR>"
                SendToEmail=trim(oRs("fh_co_email"))
                'Response.Write "TempvarBT_ID="&TempvarBT_ID&"<BR>"
                ' Response.Write "AllRefs="&AllRefs&"<BR>"
			oRs.movenext
			LOOP                        
            oRs.Close
		    Set oRs=Nothing   
    
                  
        Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
		    RSEVENTS2.Open "FCJobExceptions", DATABASE, 2, 2
		    RSEVENTS2.addnew
 		    RSEVENTS2("ExceptionID")=SQLExceptionID	
 		    RSEVENTS2("ExceptionUserID")=UserID	
            RSEVENTS2("fh_ID")=InputJobNumber								
		    'RSEVENTS2("Ref_Num")=hawb
		    RSEVENTS2("ExceptionTime")=Now()            		
		    RSEVENTS2("BillToID") = fh_bt_id
		    RSEVENTS2("Status") = "c"
		    RSEVENTS2.update
		    RSEVENTS2.close			
	    set RSEVENTS2 = nothing

'''''''''''''''''''''''''''''''''''
						Set Recordset1 = Server.CreateObject("ADODB.Recordset")
						Recordset1.ActiveConnection = DATABASE
						Recordset1.Source = "SELECT Accessorials.accCharge, AccessorialType.atDescr, AccessorialType.atid FROM Accessorials INNER JOIN AccessorialType ON Accessorials.atid = AccessorialType.atid where (AccessorialType.atid='"&SQLExceptionID&"')"
						Recordset1.CursorType = 0
						Recordset1.CursorLocation = 2
						Recordset1.LockType = 1
						Recordset1.Open()
						Recordset1_numRows = 0
						'Response.Write "Recordset1.Source="&Recordset1.Source&"<BR>"
						If Recordset1.eof then
							'ErrorMessage="There are no available suggestions"
						End if			
						
						If NOT Recordset1.EOF then 
                            ExceptionCharge=Recordset1("accCharge")
							ExceptionDescription=Recordset1("atDescr")
						
						End if
						Recordset1.Close()
						Set Recordset1 = Nothing	
'''''''''''''''''''''''''''''''''''
        'Response.write "GOT HERE!!!<BR>"
        'Response.write "SendToEmail="&SendToEmail&"<BR>"
        'Response.write "InputJobNumber="&InputJobNumber&"<BR>"
        'Response.write "ExceptionDescription="&ExceptionDescription&"<BR>"
        'SendToEmail="mark.maggiore@logisticorp.us"
		Body = "The following exception has been entered on job #"&InputJobNumber&":<BR><BR>"&ExceptionDescription&"<br>Cost:  $"&ExceptionCharge&"<BR><BR>At this time, there are no charges associated with this exception.  However, in the future there will be.<BR><BR>If you have any questions regarding this exception, either email FleetX@Logisticorp.us or phone 214-882-0620."& _
		"<BR><BR>FleetX" 
		'Response.write "Body="&Body&"<BR>"
        'Recipient = "mark.maggiore@logisticorp.us"
		'Set objMail = CreateObject("CDONTS.Newmail")
		'objMail.From = "FleetX@LogisticorpGroup.com"
		varTo = SendToEmail
        varcc ="Mark.Maggiore@LogistiCorp.us"
		varSubject = "FleetX Exception Notice"
		'objMail.MailFormat = cdoMailFormatMIME
		'objMail.BodyFormat = cdoBodyFormatHTML
		'objMail.Body = Body
		'objMail.Send
		'Set objMail = Nothing
    '''''''''''''''''''''''''''''''''''''''''''''''''''
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
        'Response.write "EMAIL SENT!!!!<BR>"
        
    End if
    If trim(ManagerNote)>"" then
        Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
		    RSEVENTS2.Open "PrivateNotes", DATABASE, 2, 2
		    RSEVENTS2.addnew
            RSEVENTS2("PrivateNoteJobNumber")=InputJobNumber
		    RSEVENTS2("PrivateNote")=ManagerNote
		    RSEVENTS2("PrivateNoteDate")=Now()									
		    RSEVENTS2("PrivateNoteEnterer")=UserID		
		    RSEVENTS2("PrivateNoteStatus") = "c"
		    RSEVENTS2.update
		    RSEVENTS2.close			
	    set RSEVENTS2 = nothing
    End if
'''''''''''''''''''''''''''''''''


			If Trim(StatusWord)="CLS" then
                'Response.write "GOT HERE!!!!<BR>"
                TempLeg_fl_st_id=Request.Form("Leg_fl_st_id")
                TempBTID=Session("sBT_ID")
				'Response.write "JobNumber="&jobnumber&"<BR>"
                'Response.write "TempLeg_fl_st_id="&TempLeg_fl_st_id&"<BR>"
                'Response.write "TempBTID="&TempBTID&"<BR>"



				Set Recordset166 = Server.CreateObject("ADODB.Recordset")
				'Response.Write "Intranet="&Intranet&"***<BR>"
				Recordset166.ActiveConnection = Database
				TempSQL="SELECT fcshipto.st_email AS SendToEmail FROM fcshipto INNER JOIN fcshipbt ON fcshipto.st_id = fcshipbt.sb_st_id "
                TempSQL=TempSQL & "where (sb_bt_id='36' or sb_bt_id='38') and st_id='"&TempLeg_fl_st_id&"'"
                
                'response.write "TempSQL="&TempSQL&"<BR>"
                
                
                Recordset166.Source = TempSQL
				

                

                Recordset166.CursorType = 0
				Recordset166.CursorLocation = 2
				Recordset166.LockType = 1
				Recordset166.Open()
				Recordset166_numRows = 0

                

					if NOT Recordset166.EOF then
						SendToEmail=Recordset166("SendToEmail")
						'Response.write "SendToEmail="&SendToEmail&"<BR>"
                        
                        Else

					End if
					Recordset166.Close()
					Set Recordset166 = Nothing	


                    Set oConn = Server.CreateObject("ADODB.Connection")
                    oConn.ConnectionTimeout = 100
                    oConn.Provider = "MSDASQL"
                    oConn.Open DATABASE
		                 SQL="SELECT * FROM fcrefs INNER JOIN fcfgthd ON fcrefs.rf_fh_id = fcfgthd.fh_id where (fh_id='"& JobNumber &"')"
	                    'Response.Write "SQL="&SQL&"<BR>"
	                    SET oRs = oConn.Execute(Sql)
	                    Do while not oRs.EOF 
	                        'Response.Write "got here...okay?" 
			                'Response.Write "FABID="&FABID&"<BR>"
			                'Response.Write "SQL="&SQL&"<BR>"
                            varBT_ID=trim(oRs("fh_bt_id"))
                            temp_ref=trim(oRs("RF_Ref"))
                            AllRefs=AllRefs & "#" & Temp_ref & "<br>"
                           'Response.Write "varBT_ID="&varBT_ID&"<BR>"
                           ' Response.Write "AllRefs="&AllRefs&"<BR>"
				        oRs.movenext
				        LOOP                        
                        oRs.Close
		                Set oRs=Nothing
                        'SendToEmail="mark.maggiore@logisticorp.us"
                        'Response.write "SENDTOEMAIL LINE 494<BR>"
                        If trim(SendToEmail)>"" then
						    Body = "Item(s):<br><br>"& AllRefs &"<BR>has/have just been been updated by a supervisor showing that it has been delivered to "& TempLeg_fl_st_id &".<br><br>"& _
						    "It was job #" & JobNumber & "<br><br>"& _
						    "FleetX" 
						    'Recipient = "mark.maggiore@logisticorp.us"
						    'Set objMail = CreateObject("CDONTS.Newmail")
						    'objMail.From = "FleetX@LogisticorpGroup.com"
						    varTo = SendToEmail
                            varcc ="FleetX@LogisticorpGroup.com"
						    varSubject = "FleetX Delivery (Updated by Supervisor)"
						    'objMail.MailFormat = cdoMailFormatMIME
						    'objMail.BodyFormat = cdoBodyFormatHTML
						    'objMail.Body = Body
						    'objMail.Send
						    'Set objMail = Nothing
     '''''''''''''''''''''''''''''''''''''''''''''''''''
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
                        End if



            End if







			Set oConn = Server.CreateObject("ADODB.Connection")
			oConn.ConnectionTimeout = 100
			oConn.Provider = "MSDASQL"
			oConn.Open DATABASE
			''''UPDATES THE WAFER
            'Response.write "GOT HERE!!!  UPDATED THE WAFER!!!! LINE 525!!!<BR>"
			l_cSQL = "UPDATE FCFGTHD SET fh_status = '"&StatusWord&"', fh_statcode='"&fh_statcode&"' "
			If fh_custpo>"" then
				l_cSQL = l_cSQL & " , fh_custpo='"&fh_custpo&"' "
			End if
			l_cSQL = l_cSQL & " WHERE (fh_id = '"&JobNumber&"')"
			'response.write "UPDATE Wafers="&l_cSQL&"<BR>"
			oConn.Execute(l_cSQL)
			Set oConn=Nothing
'''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''
			Set oConn = Server.CreateObject("ADODB.Connection")
			oConn.ConnectionTimeout = 100
			oConn.Provider = "MSDASQL"
			oConn.Open DATABASE
			''''UPDATES THE WAFER
			l_cSQL = "UPDATE REPORT_DATA SET fh_status = '"&StatusWord&"'"
			l_cSQL = l_cSQL & " WHERE (fh_id = '"&JobNumber&"')"
            'response.write "UPDATE Wafers="&l_cSQL&"<BR>"
			oConn.Execute(l_cSQL)
			Set oConn=Nothing			
'''''''''''''''''''''''''''''''''			
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'Set oConn = Server.CreateObject("ADODB.Connection")
			'oConn.ConnectionTimeout = 100
			'oConn.Provider = "MSDASQL"
			'oConn.Open DATABASE
			'oConn.Execute "MARK_NOTIFICATION_CLOSEDJOBS '" & JobNumber & "'" 
			'Set oConn=Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''			
			
			'Set oConn = Server.CreateObject("ADODB.Connection")
			'oConn.ConnectionTimeout = 100
			'oConn.Provider = "MSDASQL"
			'oConn.Open DATABASE
			''''UPDATES THE WAFER
			'l_cSQL = "UPDATE FCAIRLEG SET Al_Ca_ID = '"&Al_Ca_ID&"', al_trackno='"&al_trackno&"' "
			'l_cSQL = l_cSQL & " , Al_ST_OHD='"&Al_ST_OHD&"' "
			'l_cSQL = l_cSQL & " WHERE (al_fh_id = '"&JobNumber&"')"
			'response.write"UPDATE Wafers="&l_cSQL&"<BR>"
			'oConn.Execute(l_cSQL)
			'Set oConn=Nothing			
			
		End if

		If xxx="yes"  AND trim(ErrorMessage)="" then
		
			'''Response.Write "****WhichLeg="&WhichLeg&"****<BR>"
			'''Response.Write "****Leg_fl_secacc="&Leg_fl_secacc&"****<BR>"
			'''Response.Write "****Leg_fl_seconb="&Leg_fl_seconb&"****<BR>"
			
			'''Response.Write "****Leg_fl_t_acc="&Leg_fl_t_acc&"****<BR>"
			'''Response.Write "****Leg_fl_t_atp="&Leg_fl_t_atp&"****<BR>"			
			
			'Select Case WhichLeg
			'	Case "intermediate"
			'		Leg_fl_t_acc=Leg_fl_secacc
			'		Leg_fl_t_atp=Leg_fl_seconb
			'End Select
			'lmnop=lmnop+1		
			
			
			
			
			
			Set oConn = Server.CreateObject("ADODB.Connection")
			oConn.ConnectionTimeout = 100
			oConn.Provider = "MSDASQL"
			oConn.Open DATABASE
			''''UPDATES THE WAFER
			
			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'''''''''''ADD THE DIFFERENT VARIABLES BASED ON THE DIFFERENT LEG TYPE!!!!!!!!'''''''''

			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'TempBTID=Session("sBT_ID")
				'Response.write "JobNumber="&jobnumber&"<BR>"
                'Response.write "TempLeg_fl_st_id="&TempLeg_fl_st_id&"<BR>"
                'Response.write "TempBTID="&TempBTID&"<BR>"

			'Response.Write "lmnop="&lmnop&"********<BR>"
			
            'Response.Write "fh_bt_id="&fh_bt_id&"********<BR><br><br>"
			
			l_cSQL = "UPDATE FCLEGS SET "
            l_cSQL87 = "UPDATE FCLEGS SET "
			if Whichleg<>"first" and lmnop>1 then
				l_cSQL = l_cSQL&"fl_seconb='"& Leg_fl_t_atp &"', "
                l_cSQL87 = l_cSQL87&"fl_seconb='"& Leg_fl_t_atp &"', "
                L87="y"
				else
				l_cSQL = l_cSQL&"fl_t_atp='"& Leg_fl_t_atp &"', "
			End if
			If Whichleg<>"last" then
				l_cSQL = l_cSQL&"fl_firstdrop='"& Leg_fl_t_atd &"', "
			    else
			    l_cSQL = l_cSQL&"fl_t_atd='"& Leg_fl_t_atd &"', "
			End if
			If Whichleg<>"first" and lmnop>1 then
				l_cSQL = l_cSQL&"fl_secacc='"& Leg_fl_t_acc &"'  "
                l_cSQL87 = l_cSQL87&"fl_secacc='"& Leg_fl_t_acc &"'  "
                L87="y"
				else
				l_cSQL = l_cSQL&"fl_t_acc='"& Leg_fl_t_acc &"'  "
			End if
			If BillToID="48" then
				l_cSQL = l_cSQL&", fl_t_int = '"& Leg_fl_t_int &"', fl_t_und='"& Leg_fl_t_und &"'"
			End if	
			l_cSQL = l_cSQL&" WHERE (FL_FH_ID = '"&JobNumber&"') and (fl_counter = '"&Leg_fl_counter&"')"
            l_cSQL87 = l_cSQL87&" WHERE (FL_FH_ID = '"&JobNumber&"')"
			'response.Write "database="&database&"<BR>"
			'response.write "XXXXXXUPDATE Wafers="&l_cSQL&"<BR>"
			oConn.Execute(l_cSQL)
			Set oConn=Nothing
            If L87="y" AND TempvarBT_ID="26" then
                'Response.write "l_cSQL87="&l_cSQL87&"<BR>"
			    Set oConn = Server.CreateObject("ADODB.Connection")
			    oConn.ConnectionTimeout = 100
			    oConn.Provider = "MSDASQL"
			    oConn.Open DATABASE

                oConn.Execute(l_cSQL87)

                Set oConn=Nothing
            End if
			
			
''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''
			Set oConn = Server.CreateObject("ADODB.Connection")
			oConn.ConnectionTimeout = 100
			oConn.Provider = "MSDASQL"
			oConn.Open DATABASE
			''''UPDATES THE WAFER
			
			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'''''''''''ADD THE DIFFERENT VARIABLES BASED ON THE DIFFERENT LEG TYPE!!!!!!!!'''''''''
			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''			
''''''''''''''''''''''''''''''''''			
			l_cSQL = "UPDATE REPORT_DATA SET "
			if Whichleg<>"first" and lmnop>1 then
				l_cSQL = l_cSQL&"fl_seconb='"& Leg_fl_t_atp &"', "
				else
				l_cSQL = l_cSQL&"fl_t_atp='"& Leg_fl_t_atp &"', "
			End if
			If Whichleg<>"last" then
				l_cSQL = l_cSQL&"fl_firstdrop='"& Leg_fl_t_atd &"', "
			    else
			    l_cSQL = l_cSQL&"fl_t_atd='"& Leg_fl_t_atd &"', "
			End if
			If Whichleg<>"first" and lmnop>1 then
				l_cSQL = l_cSQL&"fl_secacc='"& Leg_fl_t_acc &"'  "
				else
				l_cSQL = l_cSQL&"fl_t_acc='"& Leg_fl_t_acc &"'  "
			End if
			l_cSQL = l_cSQL&" WHERE (FH_ID = '"&JobNumber&"')"
			'response.Write "database="&database&"<BR>"
            'response.write "XXXXXXUPDATE Wafers="&l_cSQL&"<BR>"
			oConn.Execute(l_cSQL)
			Set oConn=Nothing
''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''
'''''''''''''''CHECKS TO SEE IF THERES AND AIRLEG

'''''''''''''''IF THERE IS, UPDATE

'''''''''''''''IF NOT, INSERT

'''''''''''''''END
		'response.write "Customer="&Customer&"<BR>"
		'response.write "MARK...fl_st_id="&fl_st_id&"<BR>"
		'response.write "MARK...BillToID="&BillToID&"<BR>"
		
		If Customer="kwe" AND xxx="yes"  AND trim(ErrorMessage)="" then
			'response.write "**********************<br>"
			'response.write "addedPOD="&addedPOD&"<BR>"
			'response.write "PODID="&PODID&"<BR>"
			'response.write "XYZ="&XYZ&"<BR>"
			'response.write "**********************<br>"
			'If trim(addedPOD)>"" and PODID="" and XYZ=0 then
			If trim(addedPOD)>"" and XYZ=0 then
				'response.write "GOT HERE TO added new POD!!!<BR>"
				Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
					RSEVENTS2.Open "PODList", Database, 2, 2
					RSEVENTS2.addnew	
					RSEVENTS2("bt_ID")=BillToID		
					RSEVENTS2("st_ID") = trim(fl_st_id)
					RSEVENTS2("Signature")=addedPOD	
					RSEVENTS2("PODStatus") = "c"
					RSEVENTS2.update
					RSEVENTS2.close			
				set RSEVENTS2 = nothing	
				
				
				Set Recordset166 = Server.CreateObject("ADODB.Recordset")
				'Response.Write "Intranet="&Intranet&"***<BR>"
				Recordset166.ActiveConnection = Database
				Recordset166.Source = "SELECT PODID FROM PODList WHERE (bt_ID='"&BillToID&"') AND (st_ID='"&trim(fl_st_id)&"') AND (Signature='"&AddedPOD&"') AND (PODStatus='c')"
				'response.write "Recordset166.Source="&Recordset166.Source&"<BR>"
				Recordset166.CursorType = 0
				Recordset166.CursorLocation = 2
				Recordset166.LockType = 1
				Recordset166.Open()
				Recordset166_numRows = 0
					if NOT Recordset166.EOF then
						PODID=Recordset166("PODID")
						'response.write "NEWLYCREATEDPODID="&PODID&"***<BR>"
						DisplayPOD=PODID
						If DisplayPOD<>"" then
							AddedPOD=""
						End if
						'Response.Redirect("DriverMessage.asp")
						Else
						ErrorMessage="No such signer exists"
					End if
					Recordset166.Close()
					Set Recordset166 = Nothing				
				
				
				'Response.Write "PODID="&PODID&"<BR>"
				
				XYZ=XYZ+1
				
				
								
			End if	
		End if




					
		End if	
		'Response.Write "PODID="&podid&"<BR>"
		'Response.Write "PODChange="&PODChange&"***<BR>"
		'Response.Write "xxx="&xxx&"***<BR>"
		'Response.Write "ErrorMessage="&ErrorMessage&"***<BR>"
		'If trim(PODChange)="" and Trim(PODID)>"" then
		'	PODChange=PODID
		'End if
		'Response.Write "PODChange="&PODChange&"***<BR>"
		If (trim(PODChange)>"") and xxx="yes"  AND trim(ErrorMessage)="" then
			'Response.Write "GOT HERE???<BR>"
				Set oConn = Server.CreateObject("ADODB.Connection")
				oConn.ConnectionTimeout = 100
				oConn.Provider = "MSDASQL"
				oConn.Open DATABASE
				''''UPDATES THE WAFER
				l_cSQL = "UPDATE FCREFS SET POD='"& PODID &"' WHERE (RF_ref = '"&PODChange&"')"
				'response.write "UPDATE POD INFO!!!!="&l_cSQL&"<BR>"
				oConn.Execute(l_cSQL)
				Set oConn=Nothing		
		End if
		If xxx="yes"  AND trim(ErrorMessage)="" then		
			If fh_statcode="5" OR fh_statcode="9" OR fh_statcode="13" then
				Set oConn = Server.CreateObject("ADODB.Connection")
				oConn.ConnectionTimeout = 100
				oConn.Provider = "MSDASQL"
				oConn.Open DATABASE
				''''UPDATES THE WAFER
				l_cSQL = "UPDATE FCREFS SET REF_STATUS = '"&REF_STATUS&"' "
				'''If podid>"" then
				'''	l_cSQL = l_cSQL & ", POD='"& PODID &"'"
				'''End if
				l_cSQL = l_cSQL & " WHERE (RF_FH_ID = '"&JobNumber&"')"
				'response.write "UPDATE POD INFO!!!!="&l_cSQL&"<BR>"
				oConn.Execute(l_cSQL)
				Set oConn=Nothing			
			End if
			If fh_statcode="2" OR fh_statcode="3" OR fh_statcode="4" then
				Set oConn = Server.CreateObject("ADODB.Connection")
				oConn.ConnectionTimeout = 100
				oConn.Provider = "MSDASQL"
				oConn.Open DATABASE
				''''UPDATES THE WAFER
				l_cSQL = "UPDATE FCREFS SET REF_STATUS = NULL WHERE (RF_FH_ID = '"&JobNumber&"')"
				'response.write "UPDATE Wafers="&l_cSQL&"<BR>"
				oConn.Execute(l_cSQL)
				Set oConn=Nothing			
			End if			
		End if
		'Response.Write "WHICHLEG="&WHICHLEG&"<BR>"
		'Response.Write "Leg_fl_counter="&Leg_fl_counter&"<BR>"
		'Response.Write "JobNumber="&JobNumber&"<BR>"
		'Response.Write "leg_fl_t_atd="&leg_fl_t_atd&"<BR>"
		If lcase(Whichleg="last") and isdate(Leg_fl_t_atd) and leg_fl_t_atd>"1/1/1900" then
				Set oConn = Server.CreateObject("ADODB.Connection")
				oConn.ConnectionTimeout = 100
				oConn.Provider = "MSDASQL"
				oConn.Open DATABASE
				''''UPDATES THE WAFER
				'''''''''''l_cSQL = "UPDATE FCLEGS SET fl_job_closed = '"& leg_fl_t_atd &"', fl_t_atp = '"& leg_fl_t_atd &"' , fl_Leg_Status='d'  WHERE (FL_FH_ID = '"&JobNumber&"')"
				l_cSQL = "UPDATE FCLEGS SET fl_job_closed = '"& leg_fl_t_atd &"',  fl_Leg_Status='d'  WHERE (FL_FH_ID = '"&JobNumber&"')"
				
				'response.write "UPDATE LEGS???="&l_cSQL&"<BR>"
				
				oConn.Execute(l_cSQL)
				Set oConn=Nothing
				else
				If lcase(Whichleg<>"last") and isdate(Leg_fl_t_atd) and leg_fl_t_atd>"1/1/1900" then
					Set oConn = Server.CreateObject("ADODB.Connection")
					oConn.ConnectionTimeout = 100
					oConn.Provider = "MSDASQL"
					oConn.Open DATABASE					
					l_cSQL = "UPDATE FCLEGS SET fl_Leg_Status='d'  WHERE (FL_FH_ID = '"&JobNumber&"') AND ( fl_counter ='"& Leg_fl_counter &"')"
					
					'response.write "UPDATE LEGS???="&l_cSQL&"<BR>"
					
					oConn.Execute(l_cSQL)
					Set oConn=Nothing
					Set oConn = Server.CreateObject("ADODB.Connection")
					oConn.ConnectionTimeout = 100
					oConn.Provider = "MSDASQL"
					oConn.Open DATABASE					
					l_cSQL = "UPDATE FCLEGS SET fl_Leg_Status='c'  WHERE (FL_FH_ID = '"&JobNumber&"') AND ( fl_counter ='"& Leg_fl_counter+1 &"') AND (fl_leg_status<>'d')"
					
					'response.write "UPDATE LEGS 2 ???="&l_cSQL&"<BR>"
					
					oConn.Execute(l_cSQL)
					Set oConn=Nothing									
				End if		
		End if
		
		If xxx="yes"  AND trim(ErrorMessage)="" then
			
			Set oConn = Server.CreateObject("ADODB.Connection")
			oConn.ConnectionTimeout = 100
			oConn.Provider = "MSDASQL"
			oConn.Open DATABASE
			l_cSQL="INSERT INTO JobChanges "_ 
			& "(Fh_ID, SupervisorID, ChangeCategory, ChangeReason, ChangeDate, ChangeStatus)"_ 
			&"VALUES('"&Fh_ID&"','"&UserID&"', '"&DisplayCategoryID&"', '"&reasonforchange&"','"&Now()&"','c')" 
			'response.write "l_cSQL="&l_cSQL&"<BR>"
			oConn.Execute(l_cSQL)
			Set oConn=Nothing			
		End if		
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''INSERT INTO THE ORDER CHANGES TABLE	
	
		
	End if	
End if
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''response.write "Intranet="&Intranet&"<BR>"
If UserFirstName="" then
	Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
		RSEVENTS.CursorLocation = 3
		RSEVENTS.CursorType = 3
		RSEVENTS.ActiveConnection = DATABASE
		SQL = "SELECT * FROM PreExistingRequestor where (RequestorID = '"&UserID&"')"
		RSEVENTS.Open SQL, DATABASE, 1, 3
		'response.write "LINE 834 - MANAGER:SQL="&SQL&"<BR>"
		'LogInName=RSEVENTS("UserName")
		UserName=RSEVENTS("RequestorEmail")
		UserFirstName=RSEVENTS("RequestorName")
		'lastname=RSEVENTS("lastname")
		password=RSEVENTS("requestorpassword")
		email=RSEVENTS("requestoremail")
		'workphone=RSEVENTS("workphone")
		'homephone=RSEVENTS("homephone")
		'cellphone=RSEVENTS("cellphone")
		'faxphone=RSEVENTS("faxphone")
		'nextelphone=RSEVENTS("nextelphone")
		'TaskPreference=RSEVENTS("TaskPreference")
		RSEVENTS.close
	Set RSEVENTS = Nothing
End if	
If (JobNumber>"" or RefNumber>"") and Submit<>"Update Leg" Then
	'response.write "Got here, looking for the job<br>"
	Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
		RSEVENTS.CursorLocation = 3
		RSEVENTS.CursorType = 3
		RSEVENTS.ActiveConnection = DATABASE
		SQL = "SELECT fl_sf_rta, fh_id, fh_status, fh_bt_id, fh_custpo, fh_statcode, fh_ship_dt, fh_ready, fh_co_id, fh_priority, fh_user5, fl_sf_id, fl_sf_name, fl_st_id, fl_st_name, fl_dr_id, fl_t_disp, fl_t_acc, fl_t_atp, fl_t_int, fl_t_atd, fl_t_und, fl_st_rta, rf_ref, POD FROM fcfgthd INNER JOIN "
        SQL = SQL&"fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN "
        SQL = SQL&"fcrefs ON fclegs.fl_fh_id = fcrefs.rf_fh_id "
        If Jobnumber>"" then
			SQL = SQL&"WHERE (fcfgthd.fh_id LIKE '%"&JobNumber&"') "
			else
			If RefNumber>"" then
				SQL = SQL&"WHERE (fcrefs.rf_ref = '"&RefNumber&"') "	
			End if
		End if
		'response.write "DATABASE="&DATABASE&"<BR>"
		'response.write "LINE 868 SQL="&SQL&"<BR>"		
		RSEVENTS.Open SQL, DATABASE, 1, 3
		If RSEVENTS.EOF then
			ErrorMessage="There is no job with those parameters."
			else
			DisplayJob="y"
			fl_sf_rta=RSEVENTS("fl_sf_rta")
			fh_id=RSEVENTS("fh_id")
			jobnumber=fh_id
			fh_bt_id=RSEVENTS("fh_bt_id")
            'Response.write "XXXFh_bt_idXXX="&fh_bt_id&"<BR>"
			fh_status=RSEVENTS("fh_status")
			fh_custpo=RSEVENTS("fh_custpo")
			fh_statcode=RSEVENTS("fh_statcode")
			JobStatus=FH_Statcode
			fh_ship_dt=RSEVENTS("fh_ship_dt")
			fh_ready=RSEVENTS("fh_ready")
			fh_co_id=RSEVENTS("fh_co_id")
			'who did job?
			fh_priority=RSEVENTS("fh_priority")
			fh_user5=RSEVENTS("fh_user5")
			'material type
			fl_sf_id=RSEVENTS("fl_sf_id")
			fl_sf_name=RSEVENTS("fl_sf_name")
			fl_st_id=RSEVENTS("fl_st_id")
			fl_st_name=RSEVENTS("fl_st_name")
			fl_dr_id=RSEVENTS("fl_dr_id")
			fl_t_disp=RSEVENTS("fl_t_disp")
			fl_t_acc=RSEVENTS("fl_t_acc")
            'Response.write "Line 897 fl_t_acc="&fl_t_acc&"<BR>"
			fl_t_atp=RSEVENTS("fl_t_atp")
			fl_t_int=RSEVENTS("fl_t_int")
			fl_t_atd=RSEVENTS("fl_t_atd")
			fl_t_und=RSEVENTS("fl_t_und")
			fl_st_rta=RSEVENTS("fl_st_rta")
			rf_ref=RSEVENTS("rf_ref")	
			DisplayPOD=RSEVENTS("POD")	
			If trim(fl_sf_id)="55" then
				Display_fl_sf_id="Compugraphics"
				else
				Display_fl_sf_id=fl_sf_id
			End if
			Set RSEVENTS666 = Server.CreateObject("ADODB.Recordset")
				RSEVENTS666.CursorLocation = 3
				RSEVENTS666.CursorType = 3
				RSEVENTS666.ActiveConnection = DATABASE
				SQL = "SELECT * FROM FCBILLTO where (bt_id = '"&fh_bt_id&"')"
				RSEVENTS666.Open SQL, DATABASE, 1, 3
				'response.write "Line 915 SQL="&SQL&"<BR>"
				If not RSEVENTS666.eof then
					'LogInName=RSEVENTS666("UserName")
					bt_desc=RSEVENTS666("bt_desc")
				End if
				RSEVENTS666.close
			Set RSEVENTS666 = Nothing			
		End if

		RSEVENTS.close
	Set RSEVENTS = Nothing
	
	
	IF trim(fl_sf_id)="55" or trim(fl_st_id)="CPGP" then
		'response.write"GOT HERE????<BR>"
		Set RS2 = Server.CreateObject("ADODB.Recordset")
			RS2.CursorLocation = 3
			RS2.CursorType = 3
			RS2.ActiveConnection = DATABASE
			SQL = "SELECT al_ca_id, al_trackno, al_st_ohd FROM FCAIRLEG"
			SQL = SQL&" WHERE (al_fh_id='"&fh_id&"') "
			'SQL = SQL&" ORDER BY Category"
			RS2.Open SQL, DATABASE, 1, 3
			'response.write SQL
			If not RS2.EOF then
				al_ca_id=RS2("al_ca_id")
				al_trackno=RS2("al_trackno")
				al_st_ohd=RS2("al_st_ohd")
			End if
			RS2.close
		Set RS2 = Nothing
	End if		
		
		

	
	
	Else
	'Response.Write "XXXXXXXXXXXXXXXfh_id="&fh_id&"XXXXXXXXXXXXXXXXx<BR>"
	If fh_id>"" then
		DisplayJob="y"
	End if
End if

%>
<link rel="stylesheet" href="../themes/classic.css" id="theme_base">
<link rel="stylesheet" href="../themes/classic.date.css" id="theme_date">
<link rel="stylesheet" href="../themes/classic.time.css" id="theme_time">


<title>FleetX - <%=PageTitle %></title>
</head>

<body onload="document.form667.JobNumber.focus();">
	<table align="center" border="0" bordercolor="black" cellpadding="0" cellspacing="0" ID="Table1" height="100%" width="100%">
        <tr><td align="left"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>
        <tr>
            <td align="left"><img src="../images/FleetX_Small.jpg" height="50" width="168" /></td>
            <td align="right" valign="bottom"><!-- #include file="../topnavbar.asp" --></td>
        </tr>
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" /></td></tr>	
    <tr><td colspan="2">

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



    <tr><td>
    
    
    
    
    
    
    
    
    
    
    
    
    
    
 	<table align="center" border="0" bordercolor="green" cellpadding="3" cellspacing="0" class="MainPageText" ID="Table2">
		<%if displayjob="" then%>
		<form method="post" name="form667" id="form667">
        <tr><td>&nbsp;</td></tr>
        <tr><td>&nbsp;</td></tr>
        <tr><td>&nbsp;</td></tr>
        <tr>
			<td align="right" nowrap>Supervisor:&nbsp;&nbsp;
				<%=UserFirstName%>
				<input type="hidden" name="Supervisorid" value="<%=SupervisorID%>">	
				<input type="hidden" name="LastName" value="<%=LastName%>" ID="Hidden1">
				<input type="hidden" name="UserFirstName" value="<%=UserFirstName%>" ID="Hidden2">	
				<input type="hidden" name="UserID" value="<%=UserID%>" ID="Hidden16">			
			</td>
		</tr>
        <!--
		<tr><td>&nbsp;</td></tr>		
		<tr>
			<td align="right" nowrap width="150"><b>xxxCustomer:</b>&nbsp;&nbsp;</td>
			<td>
				<select name="customer">
				<%
				if JobManagementkwe="yxxx" then
				%>	
					<option value="kwe" <%if customer="kwe" then response.Write "selected" end if%>>KWE</option>
				<%
				end if
				%>	
					<option value="tiret" <%if customer="tiret" then response.Write "selected" end if%>>TI-Reticle</option>
					<option value="tiwf" <%if customer="tiwf" then response.Write "selected" end if%>>TI-Wafer</option>
				</select>				
			</td>
		</tr>
        -->
		<tr><td>&nbsp;</td></tr>
		<tr>
			<td align="right" nowrap width="150"><b>Job Number:</b>&nbsp;&nbsp;</td>
			<td>
				<input type="text" name="JobNumber" value="<%=JobNumber%>">				
			</td>
		</tr>
        <!--
		<tr><td>&nbsp;</td><td><b><img src="../images/pixel.gif" height="1" width="60">OR</b></td></tr>	
		<tr>
			<td align="right" nowrap width="150"><b>SR Document Number:</b>&nbsp;&nbsp;</td>
			<td>
				<input type="text" name="RefNumber" value="<%=RefNumber%>" ID="Text1">				
			</td>
		</tr>
		<tr><td>&nbsp;</td></tr>
        -->
		<tr>
			<td class="ErrorMessage" colspan="2" align="center">
				<%
				If ErrorMessage>"" then
					Response.Write "Error Message: "&ErrorMessage
					else
					If lcase(Submit)="update leg" then response.Write "<font color='blue'><br>The job has been successfully updated<br><br></font>" End if
                    If lcase(Submit)="cancel job" then response.Write "<font color='blue'><br>The job has been successfully cancelled<br><br></font>" End if
				End if
				%>				
			</td>
		</tr>		
		<tr><td>&nbsp;</td></tr>
		<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" id="gobutton"></td></tr>
		</form>
        </table>
		<%
        else
        %>
		<form method="post" name="form666" id="form666">
		<tr>
			<td class="ErrorMessage" colspan="6" align="center">
				<%
				If ErrorMessage>"" then
					Response.Write "<BR>Error Message: "&ErrorMessage&"<BR><BR>"
					else
					Response.Write "&nbsp;"
					If lcase(Submit)="update leg" then response.Write "<font color='blue'><br>The job has been successfully updated<br><br></font>" End if
                    If lcase(Submit)="cancel job" then response.Write "<font color='blue'><br>The job has been successfully cancelled<br><br></font>" End if
				End if
				%>				
			</td>
		</tr>
		<tr>
			<td class="MainPageTextRightBold" nowrap valign="top">JOB NUMBER:&nbsp;&nbsp;<%=fh_id%><br></td>
			<td class="MainPageTextRightBold" nowrap valign="top">SUPERVISOR:&nbsp;&nbsp;<%=LastName%>, <%=UserFirstName%></td>
            <!--
			<td class="MainPageTextRightBold" nowrap valign="top"><b>CUSTOMER:</b></td>
			<td valign="top">&nbsp;&nbsp;<%=Customer%></td>	
            -->	
		</tr>
        <%
            If isnumeric(fh_priority) then
				Set Recordset166 = Server.CreateObject("ADODB.Recordset")
				Recordset166.ActiveConnection = Database
				TempSQL="SELECT PriorityDescription AS display_fh_priority FROM Priorities "
                TempSQL=TempSQL & "where priorityid='"&fh_priority&"'"
                Recordset166.Source = TempSQL
			    Recordset166.CursorType = 0
				Recordset166.CursorLocation = 2
				Recordset166.LockType = 1
				Recordset166.Open()
				Recordset166_numRows = 0
                if NOT Recordset166.EOF then
						display_fh_priority=Recordset166("display_fh_priority")
						'Response.write "SendToEmail="&SendToEmail&"<BR>"
                        Else
    			End if
				Recordset166.Close()
				Set Recordset166 = Nothing	
                else
                display_fh_priority=fh_priority
         End if      
         %>
		<tr>
			<td class="MainPageTextRightBold" nowrap valign="top">PRIORITY:&nbsp;&nbsp;<%=display_fh_priority%><br></td>
			<td class="MainPageTextRightBold" nowrap valign="top">ENTERED BY:&nbsp;&nbsp;<%=fh_co_id%><br></td>
			<!--td class="MainPageTextRightBold" nowrap valign="top">MATERIAL TYPE:&nbsp;&nbsp;<%=fh_user5%><br></td-->
		</tr>
		<tr><td>&nbsp;</td></tr>
		
	</table>
	<table align="center" border="0" bordercolor="red" cellpadding="0" cellspacing="0" class="MainPageText" ID="Table3" width="600">

		<tr><td colspan="6" align="center">
			<table border="1" cellspacing="0" cellpadding="3" ID="Table4" bordercolor="gray">
				<tr>
					<td class="MainPageTextRightBold" nowrap>Order Time</td><td><%=fh_ship_dt%></td>
				</tr>
				<tr>	
					<td class="MainPageTextRightBold" nowrap>Ready Time</td><td><%=fh_ready%></td>			
				</tr>
				<tr>	
					<td class="MainPageTextRightBold" nowrap>Dispatch Time</td><td><%=fl_t_disp%></td>			
				</tr>
				<tr>			
					<td class="MainPageTextRightBold" nowrap>Due Time</td><td><%=fl_st_rta%></td>
				</tr>
				<tr>			
					<td class="MainPageTextRightBold" nowrap>Document Number:</td><td><%=rf_ref%></td>
				</tr>
				<tr>
					<td colspan="2">POD INFO:&nbsp;&nbsp;
						<%
						Set RS2 = Server.CreateObject("ADODB.Recordset")
							RS2.CursorLocation = 3
							RS2.CursorType = 3
							RS2.ActiveConnection = DATABASE
							SQL = "SELECT rf_ref FROM FCREFS"
							SQL = SQL&" WHERE (rf_fh_id='"& fh_id &"') "
							'SQL = SQL&" ORDER BY Category"
							RS2.Open SQL, DATABASE, 1, 3
							'response.write "LIne #1127 SQL="&SQL&"<BR>"
							do while not RS2.EOF
								LotDocumentNumber=RS2("rf_ref")						
''''''''''''''''''''''''''''''''''''''''''''''''''''''	
 If deletethis="becauseitsKWE" then
                            Set RSEVENTS22 = Server.CreateObject("ADODB.Recordset")
									RSEVENTS22.CursorLocation = 3
									RSEVENTS22.CursorType = 3
									'response.Write "Liberty="&Liberty&"<BR>"
									RSEVENTS22.ActiveConnection = LIBERTY
									l_csql = "SELECT * FROM F_HAWB_DATA WHERE (SZF1='"&LotDocumentNumber&"')"
									'Response.write("Query:" & l_cSQL)
									RSEVENTS22.Open l_cSQL, LIBERTY, 1, 3
									If not RSEVENTS22.EOF then
										ULID=RSEVENTS22("ULID")
										HexULID=Hex(ULID)
										'Response.Write "HEXULID="& HEXULID &"***<BR>"
										%>
										<a href="http://document.logisticorp.us:8080/LibertyIMS::/User=WebUser;pwd=Internet42;sys=LogistiCorp/Cmd%3DGetRawDocument%3BFolder%3D%2321%3BDoc%3D<%=HexULID%>%3Bformat%3DLIC/" target="_blank"><%=DisplaySignature%></a>&nbsp;
										<%
										else
										ULID=""
										If isdate(PODDateTime) then
											%>
											<a href="../KWEPODS/<%=trim(LotDocumentNumber)%>.pdf" target="_blank"><%=DisplaySignature%></a>&nbsp;
											<%
											Else
											%>									
											N/A
											
											<%
										End if
									End if
									RSEVENTS22.close
								Set RSEVENTS22 = Nothing
END IF
							RS2.movenext
							Loop
						RS2.Close
						Set RS2=nothing

						%>					
					</td>
				</tr>
							
				<%
				''''''''''''How many legs are there???''''''''''''''''''''
				Set RSEVENTS666 = Server.CreateObject("ADODB.Recordset")
					RSEVENTS666.CursorLocation = 3
					RSEVENTS666.CursorType = 3
					RSEVENTS666.ActiveConnection = DATABASE
					SQL = "SELECT fl_counter FROM FCLEGS where (fl_fh_id = '"& Fh_id &"')"
					RSEVENTS666.Open SQL, DATABASE, 1, 3
					'response.write "LINE 1176 SQL="&SQL&"<BR>"
					do while not RSEVENTS666.eof
						LegNumbers=LegNumbers+1
						'LogInName=RSEVENTS666("UserName")
						XLeg_fl_counter=RSEVENTS666("fl_counter")
					RSEVENTS666.movenext
					LOOP
					RSEVENTS666.close
				Set RSEVENTS666 = Nothing				
				''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				Set RSEVENTS666 = Server.CreateObject("ADODB.Recordset")
					RSEVENTS666.CursorLocation = 3
					RSEVENTS666.CursorType = 3
					RSEVENTS666.ActiveConnection = DATABASE
					SQL = "SELECT * FROM FCLEGS where (fl_fh_id = '"& Fh_id &"')"
					RSEVENTS666.Open SQL, DATABASE, 1, 3
                    'response.write "LINE 1192 DATABASE="&DATABASE&"<BR>"
					'response.write "LINE 1193 XYZSQL="&SQL&"<BR>"
					do while not RSEVENTS666.eof
						abc=abc+1
                        'Response.write "Got here line 1196!!!!!<BR>"
						'LogInName=RSEVENTS666("UserName")
						'If trim(ErrorMessage)="" then
							Leg_fl_counter=trim(RSEVENTS666("fl_counter"))
							Leg_fl_Leg_Status=RSEVENTS666("fl_Leg_status")
							'Response.write "LIne 1201 Leg_fl_counter="&Leg_fl_counter&"<BR>"
                            'Response.write "LIne 1202 fl_Leg_status="&fl_Leg_status&"<BR>"
							select case Leg_fl_Leg_Status
								Case "c"
									Display_Leg_fl_Leg_Status="Current Leg"
								Case "h"
									Display_Leg_fl_Leg_Status="Future Leg"
								Case "d"
									Display_Leg_fl_Leg_Status="Completed"							
								Case else
									Display_Leg_fl_Leg_Status="Don't Know Yet"
							End select							
							'response.write "Submitted_Leg_fl_counter="&Submitted_Leg_fl_counter&"<BR>"
							'response.write "Leg_fl_counter="&Leg_fl_counter&"<BR>"
							If trim(Submitted_Leg_fl_counter)<>Leg_fl_counter or showthis="" then
								'response.write "I GOT HERE NOW!<BR>"
								If abc=1 then 
									WhichLeg="first"
									else 
									'Response.Write "LegNumbers="&LegNumbers&"<BR>"
									If abc=LegNumbers then 
										WhichLeg="last"
										else
										WhichLeg="intermediate"
									 end if
								End if
								If LegNumbers=1 then 
									WhichLeg="last"	
								End if							
								If WhichLeg="last" then
									FinalDestination=Leg_fl_st_id
								End if
								'response.write "FROM DATABASE!!!!!<BR>"
								Leg_fl_sf_id=RSEVENTS666("fl_sf_id")
								Leg_fl_st_id=RSEVENTS666("fl_st_id")

	
 								
								fl_sf_name=RSEVENTS666("fl_sf_name")
                                fl_sf_clname=RSEVENTS666("fl_sf_clname")
                                'response.write "XXXfl_sf_clname="&fl_sf_clname&"<BR>"
								
								fl_st_name=RSEVENTS666("fl_st_name")
                                fl_st_clname=RSEVENTS666("fl_st_clname")                                
                                   							
								Leg_fl_un_id=RSEVENTS666("fl_un_id")
								Leg_fl_dr_id=RSEVENTS666("fl_dr_id")
								
								Leg_fl_t_acc=RSEVENTS666("fl_t_acc")
                                'Response.write "***LINE 1239 Leg_fl_t_acc="&leg_fl_t_acc&"<BR>"
								Leg_fl_t_int=RSEVENTS666("fl_t_int")
                                
                                'Response.write "whichleg="&whichleg&"<BR>"
								
                                If lcase(whichleg)<>"last" then
                                    'Response.write "got here 1<BR>"
									Leg_fl_t_atd=RSEVENTS666("fl_firstdrop")
									else
                                    'Response.write "got here 2<BR>"
                                    Leg_fl_t_atd=RSEVENTS666("fl_t_atd")
                                    If isnull(Leg_fl_t_atd) then 
                                        Leg_fl_t_atd="1/1/1900 12:00:00 AM" 
                                    end if
								End if
								'response.write "Leg_fl_t_atd="&Leg_fl_t_atd&"***<BR>"
								Leg_fl_t_und=RSEVENTS666("fl_t_und")
								Leg_fl_t_atp=RSEVENTS666("fl_t_atp")
								Leg_fl_firstdrop=RSEVENTS666("fl_firstdrop")
								Leg_fl_seconb=RSEVENTS666("fl_seconb")
								Leg_fl_secacc=RSEVENTS666("fl_secacc")
								else
								'response.write "REQUEST FORM!!!!!<BR>"
								Leg_fl_sf_id=Request.form("Leg_fl_sf_id")
								Leg_fl_st_id=Request.form("Leg_fl_st_id")
								'Leg_fl_t_acc=Request.form("Leg_fl_t_acc")
								Leg_fl_t_int=Request.form("Leg_fl_t_int")
								'Leg_fl_t_atd=Request.form("Leg_fl_t_atd")
								'response.write "MMMMMMMMleg_fl_t_atd="&leg_fl_t_atd&"***<BR>"
								Leg_fl_t_und=Request.form("Leg_fl_t_und")
								'Leg_fl_t_atp=Request.form("Leg_fl_t_atp")
								Leg_fl_firstdrop=Request.form("Leg_fl_firstdrop")
								Leg_fl_seconb=Request.form("Leg_fl_seconb")
								Leg_fl_secacc=Request.form("Leg_fl_secacc")
								whichleg=request.Form("whichleg")
								Leg_fl_sf_id=request.Form("Leg_fl_sf_id")
								Leg_fl_st_id=request.Form("Leg_fl_st_id")
								'lmnop=lmnop+1
								'response.write "lmnop="&lmnop&"************<BR>"		
								'response.write "<font color='red'>MMMMMMMMLeg_fl_t_atd="&Leg_fl_t_atd&"***<BR></font>"
								'response.write "<font color='red'>MMMMMMMMLeg_fl_firstdrop="&Leg_fl_firstdrop&"***<BR></font>"									
							End if
							If WhichLeg<>"last" then
								''''''''''Leg_fl_t_atd=Leg_fl_firstdrop
							End if
							If WhichLeg="first" then
								'response.write "damnit...got here!<BR>"
								
								'response.write "ZZZZZleg_fl_t_atd="&leg_fl_t_atd&"***<BR>"
								else
								'Leg_fl_t_acc=Leg_fl_secacc
								
								'response.write "XXXXXWhichLeg="&WhichLeg&"<BR>"
								'response.write "XXXXXLeg_fl_t_acc="&Leg_fl_t_acc&"***<BR>"
								'response.write "XXXXXLeg_fl_secacc="&Leg_fl_secacc&"***<BR>"
								'response.write "Leg_fl_t_atp="&Leg_fl_t_atp&"***<BR>"
								If trim(Leg_fl_t_acc)="" or isnull(Leg_fl_t_acc) then
									'response.write "Got here 1<br>" 
									Leg_fl_t_acc="1/1/1900" 
								end if 
								If whichleg="intermediate" then
									'response.write "GOT HERE!!!!!!!!!!!!!!!!!!!!!!<BR>"
									''''''''''Leg_fl_t_atp=Leg_fl_seconb
									If trim(Leg_fl_t_atp)="" or isnull(Leg_fl_t_atp) then 
										'response.write "Got here 2<br>" 
										Leg_fl_t_atp="1/1/1900" 
									end if 
								End if
							End if								
								If WhichLeg="last" then
									FinalDestination=Leg_fl_st_id
								End if	
								If WhichLeg<>"first" and Leg_fl_t_acc="1/1/1900" then
									Leg_fl_t_acc=Leg_fl_secacc
								End if	
								If WhichLeg<>"first" and Leg_fl_t_atp="1/1/1900" then
									Leg_fl_t_atp=Leg_fl_seconb
								End if														
							'Response.Write "XXXLeg_fl_Leg_Status="&Leg_fl_Leg_Status&"<BR>"
								If trim(Leg_fl_t_acc)="" or isnull(Leg_fl_t_acc) then
									'response.write "Got here 1<br>" 
									Leg_fl_t_acc="1/1/1900" 
								end if
								If trim(Leg_fl_t_atp)="" or isnull(Leg_fl_t_atp) then 
									'response.write "Got here 2<br>" 
									Leg_fl_t_atp="1/1/1900" 
								end if 							
							'if abc=LegNumbers then

						
												
						'If abc	
                        
                    'response.write "Database="&Database&"<BR>"  
                     'response.write "Intranet="&Intranet&"<BR>"  
				Set RSEVENTS246 = Server.CreateObject("ADODB.Recordset")
					RSEVENTS246.CursorLocation = 3
					RSEVENTS246.CursorType = 3
					RSEVENTS246.ActiveConnection = INTRANET
					SQL = "SELECT FirstName, LastName FROM lcintranet.dbo.Intranet_Users where (UserID = '"& trim(Leg_fl_dr_id) &"')"
					RSEVENTS246.Open SQL, INTRANET, 1, 3
					'response.write "LINE 1333 SQL="&SQL&"<BR>"
					if not RSEVENTS246.eof then
						DriverFirstName=RSEVENTS246("FirstName")
						DriverLastName=RSEVENTS246("LastName")
                        Else
                        DriverFirstName="Data not available"
					End if
					RSEVENTS246.close
				Set RSEVENTS246 = Nothing                        
                        					
						%>
						<form method="post" action="JobManagement.asp">
						<tr><td nowrap colspan="2" align="left" bgcolor="#3C63C1"><%=WhichLeg%> - <%=Leg_fl_counter%><br />FROM: <%=fl_sf_clname%>/<%=fl_sf_name%><br />TO: <%=fl_st_clname%>/<%=fl_st_name%></td></tr>
						<tr><td nowrap colspan="2" align="left" bgcolor="#3C63C1">Unit ID:  <%=Leg_fl_un_id%>   Driver:  <%=DriverFirstName%>&nbsp;&nbsp;<%=DriverLastName%></td></tr>
						<tr><td nowrap class="MainPageTextRightBold">Driver Ack</td><td nowrap="nowrap">DATE:&nbsp;<input size="8" type="text" name="Leg_fl_t_acc_date" value="<%=FormatDateTime(Leg_fl_t_acc, 2)%>" ID="date_1">TIME:&nbsp;<input size="8" type="text" name="Leg_fl_t_acc_time" value="<%=FormatDateTime(Leg_fl_t_acc, 3)%>" ID="time_1"></td></tr>
						<%if abc="XYZ" then%>
							<tr><td nowrap class="MainPageTextRightBold">Paper on Board</td><td nowrap="nowrap"><input type="text" name="Leg_fl_t_int" value="<%=Leg_fl_t_int%>" ID="Text11"></td></tr>
							<tr><td nowrap class="MainPageTextRightBold">At Airline</td><td nowrap="nowrap"><input type="text" name="Leg_fl_t_und" value="<%=Leg_fl_t_und%>" ID="Text12"></td></tr>
						<%end if
                        'Response.write "Leg_fl_t_atd="&Leg_fl_t_atd&"<BR>"
                        %>
						<tr><td nowrap class="MainPageTextRightBold">On Board</td><td nowrap="nowrap">DATE:&nbsp;<input size="8"  type="text" name="Leg_fl_t_atp_date" value="<%=FormatDateTime(Leg_fl_t_atp, 2)%>" ID="date_2">TIME:&nbsp;<input size="8"  type="text" name="Leg_fl_t_atp_time" value="<%=FormatDateTime(Leg_fl_t_atp, 3)%>" ID="time_2"></td></tr>
						<tr><td nowrap class="MainPageTextRightBold">Dropped</td><td nowrap="nowrap">DATE:&nbsp;<input size="8"  type="text" name="Leg_fl_t_atd_date" value="<%=FormatDateTime(Leg_fl_t_atd, 2)%>" ID="date_3">TIME:&nbsp;<input size="8"  type="text" name="Leg_fl_t_atd_time" value="<%=FormatDateTime(Leg_fl_t_atd, 3)%>" ID="time_3"></td></tr>
						<%'end if%>
						<input type="hidden" name="fl_t_disp" value="<%=fl_t_disp %>" />
						<%if Customer="kwe" and WhichLeg="last" then%>
						<tr>			
							<td class="MainPageTextRightBold" nowrap valign="top">POD</td>
							<td valign="top">
							<%
							Set Recordset1 = Server.CreateObject("ADODB.Recordset")
							Recordset1.ActiveConnection = DATABASE							
								SQL555="SELECT fcrefs.rf_ref AS rf_ref, fcrefs.rf_fh_id AS rf_fh_id, fcrefs.POD AS POD, PODList.PODID AS PODID, PODList.Signature AS Signature, PODList.PODStatus AS PODStatus, PODList.bt_ID AS bt_ID FROM fcrefs left OUTER JOIN PODList ON fcrefs.POD = PODList.PODID WHERE rf_fh_id='" & FH_ID & "'"
							Recordset1.Source = SQL555
							Recordset1.CursorType = 0
							Recordset1.CursorLocation = 2
							Recordset1.LockType = 1
							Recordset1.Open()
							Recordset1_numRows = 0
							'Response.Write "SQL555XXXXXX="&SQL555&"<BR>"
							'If Recordset1.eof then
							'	ErrorMessage="No signers exist"
							'End if			
							DO WHILE NOT Recordset1.EOF
								PODRef=Recordset1("rf_ref")
								PODSignature=Recordset1("signature")
								%>
								<input type="radio" name="PODChange" value="<%=PODRef%>">
								<%=PODRef%>&nbsp;&nbsp;&nbsp;<%=PODSignature%><BR>
								<%
							Recordset1.Movenext
							LOOP
							Recordset1.Close()
							Set Recordset1 = Nothing							 							
							%>
							

							<select name="TempPODID" ID="Select2">	
								<option value="">Select a Signature</option>							
									<%
										''''''''''''''''''''''''''''''''''''''''''''''''''''''
										Set Recordset1 = Server.CreateObject("ADODB.Recordset")
										Recordset1.ActiveConnection = DATABASE
										SQL55="SELECT PODID, Signature FROM fcshipto INNER JOIN PODList ON fcshipto.st_id = PODList.st_ID where (PODStatus='c') AND (bt_id='"&BillToID&"') AND (fcshipto.st_id='"&FinalDestination&"') ORDER BY SIGNATURE"
										Recordset1.Source = SQL55
										Recordset1.CursorType = 0
										Recordset1.CursorLocation = 2
										Recordset1.LockType = 1
										Recordset1.Open()
										Recordset1_numRows = 0
										'Response.Write "Recordset1.Source="&Recordset1.Source&"<BR>"
										If Recordset1.eof then
											ErrorMessage="No signers exist"
										End if			
										
										DO WHILE NOT Recordset1.EOF 
											PODID=Recordset1("PODID")
											Signature=Recordset1("Signature")
											%>
												<option value="<%=PODID%>" <%if trim(PODID)=trim(DisplayPOD) then response.Write " selected" end if%>><%=Signature%>(<%=PODID%>)</option>
											<%	
										Recordset1.Movenext
										LOOP
										Recordset1.Close()
										Set Recordset1 = Nothing					
										''''''''''''''''''''''''''''''''''''''''''''''''''''''
										%>
							</select> or <br>
							<input type="text" name="addedPOD" value="<%=AddedPOD%>" maxlength="50" size="20" ID="Text10">

							
							
							</td>
						</tr>		
						<%
						'response.write "PODID="&PODID&"<BR>"
						'response.write "DisplayPOD="&DisplayPOD&"<BR>"
						'response.write "SQL55="&SQL55&"<BR>"
						else
						%>
						<input type="hidden" name="DisplayPOD" value="<%=DisplayPOD%>">
						<%
						end if%>
						<tr><td nowrap class="MainPageTextRightBold">Leg Status</td><td><%=Display_Leg_fl_Leg_Status%></td></tr>
						<input type="hidden" name="Display_Leg_Fl_Leg_Status" value="<%=Display_Leg_Fl_Leg_Status%>">
							<%if whichleg="last" then%>
							<tr>			
								<td class="MainPageTextRightBold" nowrap valign="middle">Job Status</td>
								<td valign="top"><%=FH_Status %>
                                <input type="hidden" name="JobStatus" value="<%=JobStatus%>">
                                <input type="hidden" name="fh_statcode" value="<%=JobStatus%>">
                                <input type="hidden" name="FH_Status" value="<%=FH_Status%>">




                                    <!--
									<select name="JobStatus" ID="Select3">
									<%
									Set RS2 = Server.CreateObject("ADODB.Recordset")
										RS2.CursorLocation = 3
										RS2.CursorType = 3
										RS2.ActiveConnection = DATABASE
										SQL = "SELECT * FROM FCSTATUS "
										SQL = SQL&" WHERE (ss_statcode='2') OR (ss_statcode='3') OR (ss_statcode='4') OR (ss_statcode='5') OR (ss_statcode='9') OR (ss_statcode='98') "
										If BillToID="48" then	
											SQL = SQL&" OR (ss_statcode='13') "
										End if
										SQL = SQL&" ORDER BY ss_statcode"
										RS2.Open SQL, DATABASE, 1, 3
										Do while not RS2.EOF 
											Status_Description=RS2("ss_desc")
											Status_Code=RS2("ss_statcode")
											If Status_Code="13" then
												Status_Description="Paperwork on Board"
											End if
											%>
											<option value="<%=Status_Code%>" <%if cint(JobStatus)=cint(Status_Code) then Response.Write " selected" end if%>><%=Status_Description%></option>
											<%
										RS2.Movenext
										Loop
										RS2.close
									Set RS2 = Nothing		
									%>
									</select>
                                    -->



								</td>		
							</tr>
							<%else%>
							<input type="hidden" name="JobStatus" value="<%=JobStatus%>">
						<%end if%>						
						<tr>
							<td class="MainPageTextRightBold" nowrap valign="top">Change Category</td>
							<td>
								<select name="CategoryID" ID="Select1">
									<option value="" <%If DisplayCategoryID="" then Response.Write " SELECTED" end if%>>Select a category</option>
								<%
								Set RS2 = Server.CreateObject("ADODB.Recordset")
									RS2.CursorLocation = 3
									RS2.CursorType = 3
									RS2.ActiveConnection = DATABASE
									SQL = "SELECT * FROM JOBCHANGECATEGORIES"
									SQL = SQL&" WHERE (categorystatus='c') "
									SQL = SQL&" ORDER BY Category"
									RS2.Open SQL, DATABASE, 1, 3
									'response.write SQL
									Do while not RS2.EOF 
										Category=RS2("Category")
										CategoryID=RS2("CategoryID")
										%>
										<option value="<%=trim(CategoryID)%>" <%if trim(DisplayCategoryID)=trim(CategoryID) then Response.Write " selected" end if%>><%=Category%></option>
										<%
									RS2.Movenext
									Loop
									RS2.close
								Set RS2 = Nothing		
								
								
								%>				

								</select>				
							</td>
						</tr>
						<tr>			
							<td class="MainPageTextRightBold" valign="top">Reason for Change<br /><font color="red"><b>NOTE:  Customer will see any comments that you enter!</b></font></font></td><td valign="top"><TEXTAREA name="ReasonForChange" cols="35" rows="5" ID="Textarea1"><%=ReasonForChange%></TEXTAREA></td>
						</tr>
                        <%
 						Set Recordset1 = Server.CreateObject("ADODB.Recordset")
						Recordset1.ActiveConnection = DATABASE
                        'SQL10="SELECT ExceptionTime, RequestorName, ExceptionDescription FROM FCJobExceptions INNER JOIN PreExistingRequestor ON FCJobExceptions.ExceptionUserID = PreExistingRequestor.RequestorID INNER JOIN DriverExceptionList ON FCJobExceptions.ExceptionID = DriverExceptionList.ExceptionID where (fh_id='"&InputJobNumber&"') and (FCJobExceptions.Status='c')"
						SQL10="SELECT FCJobExceptions.ExceptionTime, PreExistingRequestor.RequestorName, Accessorials.accCharge, AccessorialType.atDescr FROM FCJobExceptions INNER JOIN PreExistingRequestor ON FCJobExceptions.ExceptionUserID = PreExistingRequestor.RequestorID INNER JOIN Accessorials ON FCJobExceptions.ExceptionID = Accessorials.atID INNER JOIN AccessorialType ON Accessorials.atid = AccessorialType.atid WHERE (FCJobExceptions.fh_id ='"&fh_id&"') AND (FCJobExceptions.Status = 'c') and (Accessorials.bt_id='"&fh_bt_id&"') order by exceptiontime"
                        'SQL10="SELECT FCJobExceptions.ExceptionTime, PreExistingRequestor.RequestorName, Accessorials.accCharge, AccessorialType.atDescr FROM FCJobExceptions INNER JOIN PreExistingRequestor ON FCJobExceptions.ExceptionUserID = PreExistingRequestor.RequestorID INNER JOIN Accessorials ON FCJobExceptions.ExceptionID = Accessorials.atID INNER JOIN AccessorialType ON Accessorials.atid = AccessorialType.atid WHERE (FCJobExceptions.fh_id ='"&fh_id&"') order by exceptiontime"
                        Recordset1.Source = SQL10
						Recordset1.CursorType = 0
						Recordset1.CursorLocation = 2
						Recordset1.LockType = 1
						Recordset1.Open()
						Recordset1_numRows = 0
						'Response.Write "Recordset1.Source="&Recordset1.Source&"<BR>"
						If Recordset1.eof then
							'ErrorMessage="There are no available suggestions"
                            else
                            'Response.write"<tr><td>Private Notes:</td><td>"
                            %>
	    <tr>
		    <td class="MainPageText" valign="top">
			    <span class="MainPageTextRightBold">Exceptions:  </span></td><td nowrap>

                            <%




                            ShowExceptions="y"
						End if			
						
						DO WHILE NOT Recordset1.EOF 
							ExceptionTime=Recordset1("ExceptionTime")
							RequestorName=Recordset1("RequestorName")
                            ExceptionCharge=Recordset1("AccCharge")
                            ExceptionDescription=Recordset1("atDescr")
							'Response.Write "ExceptionDescription="&ExceptionDescription&"<BR>"
						
						If tt>0 then
							'Response.Write "<tr><td colspan='13' bgcolor='white'><img src='images/pixel.gif' height='2' width='1' border='0'></td></tr>"
							'Response.Write "<tr><td colspan='13' bgcolor='gray'><img src='images/pixel.gif' height='1' width='1' border='0'></td></tr>"
							'Response.Write "<tr><td colspan='13' bgcolor='white'><img src='images/pixel.gif' height='2' width='1' border='0'></td></tr>"
							tt=0
						End if
						
							%>

									<b><%=ExceptionTime %> - <%=RequestorName %> - </b><%=ExceptionDescription%> - $<%=ExceptionCharge %><br />	

							<%	
							tt=tt+1						
						Recordset1.Movenext
						LOOP
						Response.Write "</font>"
						Recordset1.Close()
						Set Recordset1 = Nothing
 
 
 
                          If  ShowExceptions="y" then
                            Response.write "</td></tr>"
                          End if 
                            'Response.write "fh_id="&fh_id&"<BR>" 
                            'Response.write "fh_bt_id="&fh_bt_id&"<BR>" 
                        ' Response.write "SQL10="&SQL10&"<BR>"      
         %>



						<tr>
							<td class="MainPageTextRightBold" nowrap valign="top">Add an Exception:</td>
							<td>
 
                     	<select name="SQLExceptionID" ID="Select3">
					    <option value="">Select an Exception</option>
 <%
						Set Recordset1 = Server.CreateObject("ADODB.Recordset")
						Recordset1.ActiveConnection = DATABASE
                        SQL123="SELECT AccessorialType.atDescr, AccessorialType.atid FROM Accessorials INNER JOIN AccessorialType ON Accessorials.atid = AccessorialType.atid where (bt_id='"&TempvarBT_ID&"') and (AtStatus='c') order by atDescr"
						Recordset1.Source = SQL123
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
							ExceptionDescription=Recordset1("atDescr")
							ExceptionID=Recordset1("atid")
							'Response.Write "ExceptionDescription="&ExceptionDescription&"<BR>"
						
							%>
                            <option value="<%=ExceptionID%>" <%if ExceptionID=SQLExceptionID then response.Write " selected" end if%>><%=ExceptionDescription%></option>
							<%	
							x=x+1						
						Recordset1.Movenext
						LOOP
						Recordset1.Close()
						Set Recordset1 = Nothing						
						%>
 
                        </select>
                        <%
                        'Response.write "SQL123="&SQL123&"<BR>"
                         %>			
							</td>
						</tr>




<%
 
 						Set Recordset1 = Server.CreateObject("ADODB.Recordset")
						Recordset1.ActiveConnection = DATABASE
                        SQL10="SELECT PrivateNote, PrivateNoteDate, RequestorName FROM PrivateNotes INNER JOIN PreExistingRequestor ON PrivateNotes.PrivateNoteEnterer = PreExistingRequestor.RequestorID where (PrivateNoteJobNumber='"&fh_id&"') and (PrivateNoteStatus='c')"
						Recordset1.Source = SQL10
						Recordset1.CursorType = 0
						Recordset1.CursorLocation = 2
						Recordset1.LockType = 1
						Recordset1.Open()
						Recordset1_numRows = 0
						'Response.Write "Recordset1.Source="&Recordset1.Source&"<BR>"
						If Recordset1.eof then
							'ErrorMessage="There are no available suggestions"
                            else
                            'Response.write"<tr><td>Private Notes:</td><td>"
                            %>
	    <tr>
		    <td class="MainPageText" valign="top">
			    <span class="MainPageText">Admin Notes:  </span></td><td>

                            <%




                            ShowPrivateNotes="y"
						End if			
						
						DO WHILE NOT Recordset1.EOF 
							PrivateNote=Recordset1("PrivateNote")
							PrivateNoteDate=Recordset1("PrivateNoteDate")
                            RequestorName=Recordset1("RequestorName")
							'Response.Write "ExceptionDescription="&ExceptionDescription&"<BR>"
						
						If X>0 then
							'Response.Write "<tr><td colspan='13' bgcolor='white'><img src='images/pixel.gif' height='2' width='1' border='0'></td></tr>"
							'Response.Write "<tr><td colspan='13' bgcolor='gray'><img src='images/pixel.gif' height='1' width='1' border='0'></td></tr>"
							'Response.Write "<tr><td colspan='13' bgcolor='white'><img src='images/pixel.gif' height='2' width='1' border='0'></td></tr>"
							X=0
						End if
						
							%>

									<b><%=PrivateNoteDate %> - <%=RequestorName %> - </b><%=PrivateNote%><br />	

							<%	
							x=x+1						
						Recordset1.Movenext
						LOOP
						Response.Write "</font>"
						Recordset1.Close()
						Set Recordset1 = Nothing
 
 
 
                          If  ShowPrivateNotes="y" then
                            Response.write "</td></tr>"
                          End if 
                          'Response.write "SQL10="&SQL10&"<BR>" 
%>





						<tr>			
							<td class="MainPageTextRightBold" valign="top">Add an Admin Note<br /><font color="red"><b>NOTE:  Only internal admins will see any notes that you enter!</b></font></font></td><td valign="top"><TEXTAREA name="AdminNote" cols="35" rows="5" ID="Textarea2"><%=AdminNote%></TEXTAREA></td>
						</tr>						
						<input type="hidden" name="Supervisorid" value="<%=SupervisorID%>" ID="Hidden21">	
						<input type="hidden" name="LastName" value="<%=LastName%>" ID="Hidden22">
						<input type="hidden" name="UserFirstName" value="<%=UserFirstName%>" ID="Hidden23">	
						<input type="hidden" name="UserID" value="<%=UserID%>" ID="Hidden24">
						<input type="hidden" name="Customer" value="<%=Customer%>" ID="Hidden25">							
						<input type="hidden" name="WhichLeg" value="<%=WhichLeg%>" ID="Hidden26">
						<input type="hidden" name="OriginalJobStatus" value="<%=JobStatus%>" ID="Hidden20">												
						<input type="hidden" name="JobNumber" value="<%=JobNumber%>" ID="Hidden19">						
						<input type="hidden" name="Leg_fl_counter" value="<%=Leg_fl_counter%>">
						<input type="hidden" name="fh_ship_dt" value="<%=fh_ship_dt%>" ID="Hidden27">
						<input type="hidden" name="fh_ready" value="<%=fh_ready%>" ID="Hidden28">
						<input type="hidden" name="fl_st_rta" value="<%=fl_st_rta%>" ID="Hidden29">
						<input type="hidden" name="fh_priority" value="<%=fh_priority%>" ID="Hidden30">
						<input type="hidden" name="fh_co_id" value="<%=fh_co_id%>" ID="Hidden31">
						<input type="hidden" name="fh_user5" value="<%=fh_user5%>" ID="Hidden32">
						<input type="hidden" name="Leg_fl_st_id" value="<%=Leg_fl_st_id%>" ID="Hidden36">
						<input type="hidden" name="fl_st_id" value="<%=leg_fl_st_id%>" ID="Hidden39">
						<input type="hidden" name="lmnop" value="<%=LegNumbers%>" ID="Hidden40">
						<%
						'Response.write "fl_st_id="&leg_fl_st_id&"<br>"
						%>
						<input type="hidden" name="Leg_fl_sf_id" value="<%=Leg_fl_sf_id%>" ID="Hidden37">
						<!--input type="hidden" name="DisplayPOD" value="<%=DisplayPOD%>" ID="Hidden38"-->
						
						<input type="hidden" name="Leg_fl_firstdrop" value="<%=Leg_fl_firstdrop%>" ID="Hidden33">
						<input type="hidden" name="Leg_fl_seconb" value="<%=Leg_fl_seconb%>" ID="Hidden34">
						<input type="hidden" name="Leg_fl_secacc" value="<%=Leg_fl_secacc%>" ID="Hidden35">
						<tr><td nowrap colspan="2" align="center"><input type="submit" name="submit" value="Update Leg" id="gobutton"></td></tr>
                        <tr><td nowrap colspan="2" align="Left"><input type="submit" name="submit" value="Cancel Job" id="gobutton"></td></tr>
						</form>
						<%
					RSEVENTS666.movenext
					LOOP
					RSEVENTS666.close
				Set RSEVENTS666 = Nothing	
				%>
			</table>	
		</td></tr>
		<tr><td>&nbsp;</td></tr>
		

		
		

		
		
		<%IF trim(fl_sf_id)="55" or trim(fl_st_id)="CPGP" then%>
		<tr>			
			<td class="MainPageTextRightBold" nowrap valign="top">QUICKBILL NUMBER:</td><td valign="top" colspan="5">&nbsp;&nbsp;<input type="text" name="fh_custpo" value="<%=fh_custpo%>" ID="Text9"></td>
		</tr>		
		<tr>			
			<td class="MainPageTextRightBold" nowrap valign="top">GPX NUMBER:</td><td valign="top" colspan="5">&nbsp;&nbsp;<input type="text" name="al_trackno" value="<%=al_trackno%>" ID="Text7"></td>
		</tr>		
		<tr>			
			<td class="MainPageTextRightBold" nowrap valign="top">ETA OF BUS:</td><td valign="top" colspan="5">&nbsp;&nbsp;<input type="text" name="al_st_ohd" value="<%=al_st_ohd%>" ID="Text6">&nbsp;<a href="javascript:NewCal('al_st_ohd','mmddyyyy',true,12,'dropdown',true)"><img src="../images/cal.gif" width="16" height="16" border="0" alt="Pick a date"></a></td>
		</tr>
		<tr>			
			<td class="MainPageTextRightBold" nowrap valign="top">DELIVERY METHOD:</td><td valign="top" colspan="5">
			&nbsp;
			<select name="al_ca_id">
				<option value="" <%if trim(al_ca_id)="" then response.Write " Selected"%>>Select Delivery Method</option>
				<option value="GI-DAL" <%if trim(al_ca_id)="GI-DAL" then response.Write " Selected"%>>Greyhound</option>
				<option value="SW" <%if trim(al_ca_id)="SW" then response.Write " Selected"%>>Airlines</option>
			</select>
			</td>
		</tr>
		<%end if%>				

		<tr>
			<td colspan="6">
				&nbsp;
			</td>
		</tr>
		<!--tr><td colspan="6" align="center"><input type="submit" value="update" name="submit"></td></tr-->
		
		<input type="hidden" name="LastName" value="<%=LastName%>" ID="Hidden6">
		<input type="hidden" name="UserFirstName" value="<%=UserFirstName%>" ID="Hidden7">			
		<input type="hidden" name="fh_id" value="<%=fh_id%>">
		<input type="hidden" name="fl_sf_id" value="<%=fl_sf_id%>" ID="Hidden8">
		<input type="hidden" name="fl_st_id" value="<%=fl_st_id%>" ID="Hidden9">
		<input type="hidden" name="Fh_Priority" value="<%=Fh_Priority%>" ID="Hidden10">
		<input type="hidden" name="fh_co_id" value="<%=fh_co_id%>" ID="Hidden11">
		<input type="hidden" name="fh_user5" value="<%=fh_user5%>" ID="Hidden12">
		<input type="hidden" name="fl_sf_rta" value="<%=fl_sf_rta%>" ID="Hidden13">
		<input type="hidden" name="fh_ship_dt" value="<%=fh_ship_dt%>" ID="Hidden18">
		<input type="hidden" name="fh_ready" value="<%=fh_ready%>" ID="Hidden14">
		<input type="hidden" name="fl_st_rta" value="<%=fl_st_rta%>" ID="Hidden15">		
		<input type="hidden" name="Customer" value="<%=Customer%>" ID="Hidden5">
		<input type="hidden" name="JobNumber" value="<%=fh_id%>" ID="Hidden4">
		<input type="hidden" name="UserID" value="<%=UserID%>" ID="Hidden17">
		<input type="hidden" name="lmnop" value="<%=LegNumbers%>" ID="Hidden38">
		<input type="hidden" name="PageStatus" value="update" ID="Hidden3">
		</form>	
		<form method="post">
			<tr><td>&nbsp;</td></tr>
			<tr><td colspan="6" align="center"><input type="submit" value="start over" name="FindNew" ID="gobutton"></td></tr>
		</form>
	</table>




<%


End if
'Response.Write "BillToID="&BillToID&"<BR>"
%>   
    
    
    
    
    
    
    
    
    

    
    </td></tr>



 
	<tr Height="50">
		<td>&nbsp;</td>
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

<tr><td Height="90%">&nbsp;</td></tr>
<tr>
    <td height="100" class="FleetXGreySection" colspan="2">
        <!-- #include file="../BottomSection.asp" -->
    </td>
</tr>
<tr><td height="15" class="FleetXRedSectionSmall" colspan="2" align="center"><%=CopywriteNotice %></td></tr>
</td></tr></table>

<script src="../jquery-2.1.0.min.js"></script> 
<script src="../pickadate.js"></script> 
<script type="text/javascript">
    // PICKADATE FORMATTING
    $('#date_1').pickadate({
        format: 'mm/dd/yyyy', 	// Friendly format displayed to user
        formatSubmit: 'mm/dd/yyyy', // Actual format used by application
        hiddenName: false			// Allows two different formats
    });


    $('#time_1').pickatime({
        format: 'h:i A', 		// Displayed and application format
        interval: 5, 			// Interval between values (in minutes)
        min: '12:00 AM', 			// Starting value
        max: '11:59 PM'				// Ending value
    });

    // PICKADATE FORMATTING
    $('#date_2').pickadate({
        format: 'mm/dd/yyyy', 	// Friendly format displayed to user
        formatSubmit: 'mm/dd/yyyy', // Actual format used by application
        hiddenName: false			// Allows two different formats
    });


    $('#time_2').pickatime({
        format: 'h:i A', 		// Displayed and application format
        interval: 5, 			// Interval between values (in minutes)
        min: '12:00 AM', 			// Starting value
        max: '11:59 PM'				// Ending value
    });

    // PICKADATE FORMATTING
    $('#date_3').pickadate({
        format: 'mm/dd/yyyy', 	// Friendly format displayed to user
        formatSubmit: 'mm/dd/yyyy', // Actual format used by application
        hiddenName: false			// Allows two different formats
    });


    $('#time_3').pickatime({
        format: 'h:i A', 		// Displayed and application format
        interval: 5, 			// Interval between values (in minutes)
        min: '12:00 AM', 			// Starting value
        max: '11:59 PM'				// Ending value
    });

</script>
</body>
</html>
