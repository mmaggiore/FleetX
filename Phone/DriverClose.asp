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

Dim FormBarCode(12)
Dim AllegedBarCode(12)
Dim FormJobNumber(12)
FormBarCode(1)=Request.Form("FormBarCode1")
FormBarCode(2)=Request.Form("FormBarCode2")
FormBarCode(3)=Request.Form("FormBarCode3")
FormBarCode(4)=Request.Form("FormBarCode4")
FormBarCode(5)=Request.Form("FormBarCode5")
FormBarCode(6)=Request.Form("FormBarCode6")
FormBarCode(7)=Request.Form("FormBarCode7")
FormBarCode(8)=Request.Form("FormBarCode8")
FormBarCode(9)=Request.Form("FormBarCode9")
FormBarCode(10)=Request.Form("FormBarCode10")
FormBarCode(11)=Request.Form("FormBarCode11")
FormBarCode(12)=Request.Form("FormBarCode12")
AllegedBarCode(1)=Request.Form("AllegedBarCode(1)")
AllegedBarCode(2)=Request.Form("AllegedBarCode(2)")
AllegedBarCode(3)=Request.Form("AllegedBarCode(3)")
AllegedBarCode(4)=Request.Form("AllegedBarCode(4)")
AllegedBarCode(5)=Request.Form("AllegedBarCode(5)")
AllegedBarCode(6)=Request.Form("AllegedBarCode(6)")
AllegedBarCode(7)=Request.Form("AllegedBarCode(7)")
AllegedBarCode(8)=Request.Form("AllegedBarCode(8)")
AllegedBarCode(9)=Request.Form("AllegedBarCode(9)")
AllegedBarCode(10)=Request.Form("AllegedBarCode(10)")
AllegedBarCode(11)=Request.Form("AllegedBarCode(11)")
AllegedBarCode(12)=Request.Form("AllegedBarCode(12)")
FormJobNumber(1)=Request.Form("FormJobNumber(1)")
FormJobNumber(2)=Request.Form("FormJobNumber(2)")
FormJobNumber(3)=Request.Form("FormJobNumber(3)")
FormJobNumber(4)=Request.Form("FormJobNumber(4)")
FormJobNumber(5)=Request.Form("FormJobNumber(5)")
FormJobNumber(6)=Request.Form("FormJobNumber(6)")
FormJobNumber(7)=Request.Form("FormJobNumber(7)")
FormJobNumber(8)=Request.Form("FormJobNumber(8)")
FormJobNumber(9)=Request.Form("FormJobNumber(9)")
FormJobNumber(10)=Request.Form("FormJobNumber(10)")
FormJobNumber(11)=Request.Form("FormJobNumber(11)")
FormJobNumber(12)=Request.Form("FormJobNumber(12)")
AliasCode=Request.Form("AliasCode")
LocationCode=Request.Form("LocationCode")
fh_status=Request.Form("fh_status")
BillToID=Request.Form("BillToID")
'Response.Write " 77 BillToID="&BillToID&"<BR>"
'response.write " 78 LocationCode="&LocationCode&"<BR>"
PODID=Request.Form("TempPODID")
AddedPOD=Request.Form("AddedPOD")
If AddedPOD>"" then
	AddedPOD=Replace(AddedPOD,",","")
End if
'response.Write "locationcode="&LocationCode&"<BR>"
					''''''''''''Determines if needs a POD or not''''''''''''''''''''''''
			Set Recordset1 = Server.CreateObject("ADODB.Recordset")
			Recordset1.ActiveConnection = DATABASE
			Recordset1.Source = "SELECT IsStockroom FROM PreExistingCompanies WHERE (st_id='"& LocationCode &"') and (CompanyStatus='c')"
			
			'response.write "Recordset1.Source="&Recordset1.Source&"<BR>"
			
			Recordset1.CursorType = 0
			Recordset1.CursorLocation = 2
			Recordset1.LockType = 1
			Recordset1.Open()
			Recordset1_numRows = 0
			If Recordset1.eof then
				'NeedPOD="y"
                'Response.write "Did I get here?<BR>"
			End if			
			If not Recordset1.eof then
                IsStockroom=Recordset1("IsStockroom")
                'Response.write "IsStockroom="&IsStockroom&"<BR>"
                If IsStockroom="y" then
                    NeedPOD="n"
                    else
                    NeedPOD="y"
                End if
				'Rollover="y"
				'Response.Write "GOT HERE!!  ROLLOVER=YES<BR>"
                'Response.write "NeedPOD="&NeedPOD&"<BR>"
			End if
			Set Recordset1 = Nothing					
					''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

		If UCASE(LocationCode)="EBHUB" or UCASE(LocationCode)="13536" or UCASE(LocationCode)="SBRT" then
			If UCASE(LocationCode)="EBHUB" then
				BillToID="91"
				LocationCode="EBHUB"
				Hub="y"
				else
				If UCASE(LocationCode="SBRT") then
					BillToID="36"
					LocationCode="SBRT"
					Hub="y"				
					Else
					BillToID="48"
					LocationCode=UCASE(LocationCode)
					'''''''''THIS MAY NEED TO GO AWAY!
					'''''Hub2="y"
				End if
			End if
		End if
'Response.Write "locationCode="&LocationCode&"<BR>"
If ucase(LocationCode)="R1-W" then LocationCode="R1" end if 
'''''''''''''''''''''''''''''''''''''''''
'response.write "mmmmBilltoid="&Billtoid&"<BR>"
'response.write "mmmmfh_status="&fh_status&"<BR>"
		
If trim(vehicleID)="199xxx" and fh_status="PUO" then
	'response.write "GOT HERE!!!<BR>"
	Hub="y"
End if	
'''''''''''''''''''''''''''''''''''''''''		
BarCode=Request.Form("BarCode")
BillToID=Request.Form("BillToID")
'REsponse.Write "147 Billtoid="&Billtoid&"<BR>"
If BillToID>"" then
	'Response.Write "GOT HERRE!!!!<BR>"
	Suid=BillToID
End if
													If VehicleID=666 and BarCode>"" then
														Barcode=Right(Barcode,10)
														Barcode=Left(Barcode,9)
														Barcode=0&Barcode
													End if
Submit=Request.Form("Submit")
If Submit="" and Barcode="" then
    If BilltoID=48 then
    response.Redirect("DriverifabPhoneEmulator_KWE.asp")
    else
	response.Redirect("DriverifabPhoneEmulator.asp")
	end if
End if
PageStatus=Request.Form("PageStatus")
If Submit="" then
	Submit=Request.QueryString("Submit")
End if
DateSent=Request.Form("DateSent")
If Submit="" or DateSent="" then
	DateSent=Date()
End if
If DateSent>"" then
	DayAfter=cDate(DateSent)+1
End if
DocumentNumber=Request.Form("DocumentNumber")
If DocumentNumber="" then
	DocumentNumber=Request.QueryString("DocumentNumber")
End if
DocumentNumber=Replace(DocumentNumber,"""","")
DocumentNumber=Replace(DocumentNumber,"'","")
LotNumber=Request.Form("LotNumber")
If LotNumber="" then
	LotNumber=Request.QueryString("LotNumber")
End if
LotNumber=Replace(LotNumber,"""","")
LotNumber=Replace(LotNumber,"'","")
SortBy=Request.Form("SortBy")
ToLocation=Request.Form("ToLocation")
FromLocation=Request.Form("FromLocation")
JobNumber=Request.Form("JobNumber")
If JobNumber="" then
	JobNumber=Request.QueryString("JobNumber")
End if
ReferenceNumber=Request.Form("ReferenceNumber")
JobNumber=Replace(JobNumber,"""","")
JobNumber=Replace(JobNumber,"'","")
ReferenceNumber=Replace(ReferenceNumber,"""","")
ReferenceNumber=Replace(ReferenceNumber,"'","")
'Response.Write "BillTOID="&BIllTOID&"<BR>"
'response.write "201 PageStatus="&PageStatus&"<BR>"
IF Submit="submit" THEN
	'response.write "203 get here<br>"
  	If PageStatus="ONB" then
			If (VehicleID=124) and LocationCode<>"ESTK" then
				ORDERSTATUS="S"
				else
				ORDERSTATUS="o"
			End if
			else
			ORDERSTATUS="c"
			if Hub="y" or Hub2="y" then
					ORDERSTATUS="H"
			End if
		End if
		If PageStatus="ONB" then
			'response.write "216 here<br>"

      For q=1 to 12
				If trim(FormBarCode(q))>""  then
					Set oRs = Server.CreateObject("ADODB.Recordset")
					oRs.CursorLocation = 3
					oRs.CursorType = 3
					oRs.ActiveConnection = DATABASE	
					SQL = "SELECT fcfgthd.fh_id, fcfgthd.fh_user3, fclegs.fl_sf_rta, fcfgthd.fh_status, fcfgthd.fh_ship_dt, fclegs.fl_sf_id, fclegs.fl_sf_name, fclegs.fl_st_id, fclegs.fl_st_name, fclegs.fl_t_atp, fclegs.fl_t_atd, fclegs.fl_pod, fcfgthd.fh_custpo, fcfgthd.fh_priority, fcrefs.rf_ref FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN fcrefs ON fclegs.fl_fh_id = fcrefs.rf_fh_id "	
					SQL = SQL&" WHERE (fl_un_id='"&VehicleID&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND (fh_ship_dt>'"&now()-30&"') AND (rf_ref='"&FormBarCode(q)&"')"
					If PageStatus="ONB" then
						If BillToID="48" or trim(vehicleID)="198" then
							SQL = SQL& "AND ((fh_status='PUO') OR (fh_status='AC2')) "
							SQL = SQL& "AND ((ref_status='p') OR (ref_status='a')) "
							Else
              if LocationCode = "SRHUB" then
							 SQL = SQL& "AND ((fh_status='ACC' OR fh_status = 'AC2') "
							else
               SQL = SQL& "AND ((fh_status='ACC') "
              end if
							If (VehicleID=124 or VehicleID=123 or VehicleID=113 or vehicleID=125) and LocationCode<>"ESTK" then
								SQL = SQL& " OR (fh_status='AC2')) "
								else
								SQL = SQL& " ) "
							End if
							if LocationCode <> "SRHUB" then
                SQL = SQL& "AND ((ref_status is NULL) "
              end if
							If (VehicleID=124 or VehicleID=123 or VehicleID=113) and LocationCode<>"ESTK" then
								SQL = SQL& " OR (ref_status='H')) "
								else
								if LocationCode <> "SRHUB" then
                  SQL = SQL& " ) "
                end if
							End if								
						End if
						If (VehicleID<>124 AND VehicleID<>123 and LocationCode <> "SRHUB") or LocationCode="ESTK" then
							SQL = SQL& "AND ((fl_sf_id='"&LocationCode&"') "
						End if
						''''''''''''THIS MIGHT NEED TO GO AWAY!!!!
						If HUB="yXXX" then
							SQL = SQL&" OR (Fl_sf_ID='D6W3')"
							SQL = SQL&" OR (Fl_sf_ID='D6N2')"
							SQL = SQL&" OR (Fl_sf_ID='D6N1')"
							SQL = SQL&" OR (Fl_sf_ID='DM4M')"
							SQL = SQL&" OR (Fl_sf_ID='DM5M')"
							SQL = SQL&" OR (Fl_sf_ID='DPI2')"
							SQL = SQL&" OR (Fl_sf_ID='DPI3')"
							SQL = SQL&" OR (Fl_sf_ID='ESTK')"
						End if
						If (VehicleID<>124 AND VehicleID<>123 and LocationCode <> "SRHUB") or LocationCode="ESTK" then	
							SQL = SQL& " ) "						
						End if
					end if					
					If PageStatus="CLS" then
						If (VehicleID=124 or VehicleID=123) and LocationCode<>"ESTK" then
							SQL = SQL& "AND (fh_status='ARV') or (fh_status='ONB') " 
						Else
              SQL = SQL& "AND (fh_status='ONB') "
						End if						
						if LocationCode<>"SRHUB" AND LocationCode<>"D6N1B" AND LocationCode<>"DOCK7" then
						  SQL = SQL& "AND (ref_status='o') "
              SQL = SQL& "AND ((fl_st_id='"&LocationCode&"') "
            end if
            if LocationCode = "D6N1B" then
              SQL = SQL& " AND (fl_st_id='D6W3' or fl_st_id='D6N2') "
            end if
            if LocationCode = "DOCK7" then
              SQL = SQL& " AND (fl_st_id='DM5M' or fl_st_id='DM5Q' or fl_st_id='DM5S3') "
            end if
						If HUB="y" then
							SQL = SQL&" OR ((Fl_st_ID='D6W3') AND (fl_sf_ID<>'EBHUB'))"
							SQL = SQL&" OR ((Fl_st_ID='D6N2') AND (fl_sf_ID<>'EBHUB'))"
							SQL = SQL&" OR ((Fl_st_ID='D6N1') AND (fl_sf_ID<>'EBHUB'))"
							SQL = SQL&" OR ((Fl_st_ID='DM4M') AND (fl_sf_ID<>'EBHUB'))"
							SQL = SQL&" OR ((Fl_st_ID='DM5M') AND (fl_sf_ID<>'EBHUB'))"
							SQL = SQL&" OR ((Fl_st_ID='DPI2') AND (fl_sf_ID<>'EBHUB'))"
							SQL = SQL&" OR ((Fl_st_ID='DPI3') AND (fl_sf_ID<>'EBHUB'))"
							SQL = SQL&" OR ((Fl_st_ID='ESTK') AND (fl_sf_ID<>'EBHUB'))"
							SQL = SQL&" OR ((Fl_st_ID='DM5Q') AND (fl_sf_ID<>'EBHUB'))"
							SQL = SQL&" OR ((Fl_st_ID='DM6Q') AND (fl_sf_ID<>'EBHUB'))"
							'SQL = SQL&" OR (Fl_st_ID='TISHERMA')"
						End if	
						If HUB2="yxxx" then
							SQL = SQL&" OR (Fl_st_ID='12201')"
							SQL = SQL&" OR (Fl_st_ID='12203')"
							SQL = SQL&" OR (Fl_st_ID='6430')"
							SQL = SQL&" OR (Fl_st_ID='6412')"
							SQL = SQL&" OR (Fl_st_ID='13601')"
							SQL = SQL&" OR (Fl_st_ID='12500')"
							SQL = SQL&" OR (Fl_st_ID='13020')"
							SQL = SQL&" OR (Fl_st_ID='7800')"
							SQL = SQL&" OR (Fl_st_ID='7839')"
							SQL = SQL&" OR (Fl_st_ID='13353')"
							SQL = SQL&" OR (Fl_st_ID='13536')"
							SQL = SQL&" OR (Fl_st_ID='13121')"
							SQL = SQL&" OR (Fl_st_ID='6550')"
							SQL = SQL&" OR (Fl_st_ID='13011')"
							SQL = SQL&" OR (Fl_st_ID='13570')"
						End if							
						if LocationCode<>"SRHUB" AND LocationCode<>"D6N1B" AND LocationCode<>"DOCK7" then
              SQL = SQL& " ) "
            end if								
					End if
					'Response.Write "314 SQL="&SQL&"<BR>"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''					
					oRs.Open SQL, DATABASE, 1, 3
					If oRs.EOF then
						ErrorMessage=ErrorMessage&" "&FormBarCode(q)&" is not accepted.<br>Check Paper Work/Call Supervisor (2)<br>"
					End if
					If not oRs.eof then
						TheJobNumber = oRs("fh_id")
                        fh_user3 = oRs("fh_user3")
						TheBarCode = FormBarCode(q)
						The2Address = oRs("fl_st_id")
					'''''''''AUTO DOES THE EXCEPTION OF OVERNIGHT DELIVERY*******8
						If BillToID="48" AND theTime>="6:00:00 PM" then
							Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
								RSEVENTS2.Open "FCJobExceptions", DATABASE, 2, 2
								RSEVENTS2.addnew
								RSEVENTS2("fh_ID")=TheJobNumber
								RSEVENTS2("ExceptionID")= 8									
								RSEVENTS2("Ref_Num")=TheBarCode		
								RSEVENTS2("BillToID") = BillToID
								RSEVENTS2("ExceptionTime")=Now()		
								RSEVENTS2("Status") = "c"
								RSEVENTS2.update
								RSEVENTS2.close			
							set RSEVENTS2 = nothing	
							Set Recordset1 = Server.CreateObject("ADODB.Recordset")
							Recordset1.ActiveConnection = DATABASE
							Recordset1.Source = "SELECT ExceptionDescription FROM DriverExceptionList where (fh_bt_id='"&BillToID&"') and (Status='c') and (ExceptionID='8')"
							Recordset1.CursorType = 0
							Recordset1.CursorLocation = 2
							Recordset1.LockType = 1
							Recordset1.Open()
							Recordset1_numRows = 0
							If Recordset1.eof then
								ErrorMessage="Error on Page"
							End if	
							If Not Recordset1.eof then
								ExceptionDescription=Recordset1("ExceptionDescription")
							End if	
							Recordset1.Close()
							Set Recordset1 = Nothing				
			''''''''''''''''''''email notification BEGIN
								Body = "RE:&nbsp;&nbsp; HAWB #&nbsp;&nbsp;"& TheBarCode &"<br><br>"   & _
								"The driver has reported the following exception:<br><br>"   & _
								" "& ExceptionDescription &"<br><br>"  & _
								"If you have any questions, please do not hesitate to contact me.<br><br>"   & _
								"Thank you,<br><br>"   & _
								"Mark Maggiore<br>"  & _
								"LogistiCorp Web Developer<br>"  & _
								"mark.maggiore@LogistiCorp.us<br>"  & _ 
								"214/956-0400 xt 212<br><br>"
								Recipient=FirstName&" "&LastName
								Email="KWETI.Mailbox@am.kwe.com"
								'Email="mark@maggiore.net"
								'Set objMail = CreateObject("CDONTS.Newmail")
								'objMail.From = "FleetX@LogisticorpGroup.com"
								varTo = Email
								varcc = "mark.maggiore@logisticorp.us"
								varSubject = "HAWB #"& TheBarCode &" Exception"
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
							End if
							
						'''''''''''END***************************
						Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open DATABASE
						''''UPDATES THE WAFER
						l_cSQL = "UPDATE FCREFS SET ref_status = '"&ORDERSTATUS&"' WHERE (rf_fh_id = '"&TheJobNumber&"') AND (rf_ref = '" &TheBarCode& "')"
						oConn.Execute(l_cSQL)
						''''''''''''''''''''''''''''''''''''''''
						Set oRs = Server.CreateObject("ADODB.Recordset")
						oRs.CursorLocation = 3
						oRs.CursorType = 3
						oRs.ActiveConnection = DATABASE	
						''''''CHECKS TO SEE IF ITS OKAY TO CHANGE THE ORDER STATUS
						SQL = "SELECT rf_fh_id FROM fcrefs"
						SQL = SQL&" WHERE (rf_fh_id='"&TheJobNumber&"')"
						If BillToID="48" or trim(vehicleID)="198" then
							SQL = SQL&" AND ((Ref_Status<>'o') or (Ref_Status IS NULL))"
							else
							SQL = SQL&" AND (Ref_Status IS NULL)"
						End if
						oRs.Open SQL, DATABASE, 1, 3
						If oRs.EOF then
						''''''''''''''''''''''''''''''''''''''''''''''''
						'Response.Write "aaaVehicleID="&VehicleID&"<BR>"
						'Response.Write "aaaLocationCode="&LocationCode&"<BR>"
						'Response.Write "The2Address="&The2Address&"<BR>"
						''''''''''''''''''''''''''''''''''''''''''''''''
						            Set oConn3 = Server.CreateObject("ADODB.Connection")
						            oConn3.ConnectionTimeout = 100
						            oConn3.Provider = "MSDASQL"
						            oConn3.Open DATABASE3						
							If ((trim(VehicleID)<>124 AND trim(VehicleID)<>113 ) or trim(LocationCode)="ESTK") and trim(LocationCode)<>"SRHUB" and trim(LocationCode)<>"13536" or trim(The2Address)="TISHERMA"  then
								oConn.Execute "PHONE_CHANGE_STATUS '" & TheJobNumber & "', '5', 'ONB', '', '',  '"& UserID &"', '"& UnitID &"'" 
                                If trim(fh_user3)>"" then
                                'response.write "GOT HERE!!!!<BR>"
                                oConn3.Execute "PHONE_CHANGE_STATUS '" & fh_user3 & "', '5', 'ONB', '', '',  '"& UserID &"', '"& UnitID &"'" 
                                End if
								else
								response.Write "420 PHONE_CHANGE_STATUS '" & TheJobNumber & "', '54', 'DPV', '', '',  '"& UserID &"', '"& UnitID &"'"
								oConn.Execute "PHONE_CHANGE_STATUS '" & TheJobNumber & "', '54', 'DPV', '', '',  '"& UserID &"', '"& UnitID &"'" 
                                If trim(fh_user3)>"" then
                                oConn3.Execute "PHONE_CHANGE_STATUS '" & fh_user3 & "', '54', 'DPV', '', '',  '"& UserID &"', '"& UnitID &"'" 
                                End if
							End if






						End if
						oRs.Close
						Set oRs=Nothing							
						oConn.Close	
						Set oConn=Nothing	
						TheJobNumber=""
						TheBarCode=""	
					End if
				End if
			Next
		End if
		'LocationID=fl_st_id
		'Response.Write "443 LocationCode="&LOcationCode&"<BR>"
		'Response.Write "XXXBillToID="&BillToID&"<BR>"
		'Response.Write "XXXAddedPOD="&AddedPOD&"<BR>"
		If PageStatus="CLS" then
			'response.write "435 here<br>"
      If addedPOD>"" and PODID="xxx" and XYZ=0 then
				Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
					RSEVENTS2.Open "PODList", Database, 2, 2
					RSEVENTS2.addnew	
					'RSEVENTS2("bt_ID")=BillToID		
					RSEVENTS2("st_ID") = LocationCode
					RSEVENTS2("Signature")=addedPOD	
					RSEVENTS2("PODStatus") = "c"
					RSEVENTS2.update
					RSEVENTS2.close			
				set RSEVENTS2 = nothing	
				Set Recordset166 = Server.CreateObject("ADODB.Recordset")
				Recordset166.ActiveConnection = Database
				Recordset166.Source = "SELECT PODID FROM PODList WHERE (st_ID='"&LocationCode&"') AND (Signature='"&AddedPOD&"') AND (PODStatus='c')"
				Recordset166.CursorType = 0
				Recordset166.CursorLocation = 2
				Recordset166.LockType = 1
				Recordset166.Open()
				Recordset166_numRows = 0
					if NOT Recordset166.EOF then
						PODID=Recordset166("PODID")
						Else
						ErrorMessage="No such signer exists"
					End if
					Recordset166.Close()
					Set Recordset166 = Nothing				
				XYZ=XYZ+1
			End if		
			For q=1 to 12
				If trim(FormBarCode(q))>""  then
					Set oRs = Server.CreateObject("ADODB.Recordset")
					oRs.CursorLocation = 3
					oRs.CursorType = 3
					oRs.ActiveConnection = DATABASE	
					SQL = "SELECT fcfgthd.fh_id, fcfgthd.fh_user3, fcfgthd.fh_co_email, fclegs.fl_pkey, fclegs.fl_sf_rta, fcfgthd.fh_status, fcfgthd.fh_ship_dt, fclegs.fl_sf_id, fclegs.fl_sf_name, fclegs.fl_st_id, fclegs.fl_st_name, fclegs.fl_t_atp, fclegs.fl_t_atd, fclegs.fl_pod, fclegs.fl_sf_comment, fclegs.fl_st_comment, fcfgthd.fh_custpo, fcfgthd.fh_priority, fcrefs.rf_ref FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN fcrefs ON fclegs.fl_fh_id = fcrefs.rf_fh_id "	
					SQL = SQL&" WHERE (fl_un_id='"&VehicleID&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND (fh_ship_dt>'"&now()-30&"') AND (rf_ref='"&FormBarCode(q)&"')"
					If PageStatus="ONB" then
						If BillToID="48" or trim(vehicleID)="198" then
							SQL = SQL& "AND (fh_status='PUO') "
							SQL = SQL& "AND (ref_status='p') "
							Else
							SQL = SQL& "AND ((fh_status='ACC') "
							If (VehicleID=124 OR VehicleID=123)  and LocationCode<>"ESTK" then
								SQL = SQL& " OR (fh_status='ARV')) "
								else
								SQL = SQL& " ) "
							End if
							SQL = SQL& "AND ((ref_status is NULL) "
							If (VehicleID=124 OR VehicleID=123) and LocationCode<>"ESTK" then
								SQL = SQL& " OR (ref_status='H')) "
								else
								SQL = SQL& " ) "
							End if								
						End if
						If (VehicleID<>124 AND VehicleID<>123) or LocationCode="ESTK" then
							SQL = SQL& "AND ((fl_sf_id='"&LocationCode&"') "
						End if
						'''''''''THIS MIGHT NEED TO GO AWAY
						If HUB="yXXX" then
							SQL = SQL&" OR (Fl_sf_ID='D6W3')"
							SQL = SQL&" OR (Fl_sf_ID='D6N2')"
							SQL = SQL&" OR (Fl_sf_ID='D6N1')"
							SQL = SQL&" OR (Fl_sf_ID='DM4M')"
							SQL = SQL&" OR (Fl_sf_ID='DM5M')"
							SQL = SQL&" OR (Fl_sf_ID='DPI2')"
							SQL = SQL&" OR (Fl_sf_ID='DPI3')"
							SQL = SQL&" OR (Fl_sf_ID='ESTK')"
						End if
						
						If (VehicleID<>124 AND VehicleID<>123) or LocationCode="ESTK" then	
							SQL = SQL& " ) "						
						End if
					end if					
					If PageStatus="CLS" then
							'response.write "522 here pagestatus=" & pagestatus & "<br>"
              If trim(vehicleID)="199" then
								SQL = SQL& "AND ((((fh_status='PUO')) "
								SQL = SQL& "AND ((ref_status='p'))) OR ((fh_status='ONB') AND  (fl_rt_type='out'))) "							
								else
								If VehicleID=124 or VehicleID=123 then
									SQL = SQL& "AND ((fh_status='DPV') OR (fh_status='ONB') ) "
									SQL = SQL& "AND ((ref_status='S') OR  (ref_status='o')) "
									else
									SQL = SQL& "AND (((fh_status='ONB') "
									SQL = SQL& "AND (ref_status='o')) OR fh_status='DPV' )"
								End if	
							End if
												
						If HUB2<>"y" AND LocationCode <>"SRHUB" AND LocationCode<>"D6N1B" AND LocationCode<>"DOCK7" then
							SQL = SQL&" AND ((Fl_st_ID='"&LocationCode&"')"
						End if
            if LocationCode="D6N1B" then
              SQL = SQL& " AND (Fl_st_ID='D6W3' or Fl_st_ID='D6N2') "
            end if
            if LocationCode="DOCK7" then
              SQL = SQL& " AND (Fl_st_ID='DM5M' or Fl_st_ID='DM5Q' or fl_st_id='DM5S3') "
            end if
							'Response.Write "XXXlocationcode="&locationcode&"***<BR>"
							if trim(LocationCode)="SBRT" then
								CloseThis="n"
								SQL = SQL&" OR (Fl_st_ID='TISHERMA')"
							end if							
						If HUB="y" then
							SQL = SQL&" OR (Fl_st_ID='D6W3')"
							SQL = SQL&" OR (Fl_st_ID='D6N2')"
							SQL = SQL&" OR (Fl_st_ID='D6N1')"
							SQL = SQL&" OR (Fl_st_ID='DM4M')"
							SQL = SQL&" OR (Fl_st_ID='DM5M')"
							SQL = SQL&" OR (Fl_st_ID='DPI2')"
							SQL = SQL&" OR (Fl_st_ID='DPI3')"
							SQL = SQL&" OR (Fl_st_ID='ESTK')"
							SQL = SQL&" OR (Fl_st_ID='DM5Q')"
							SQL = SQL&" OR (Fl_st_ID='DM6Q')"
							SQL = SQL&" OR (Fl_st_ID='RFAB-R')"
							'SQL = SQL&" OR (Fl_st_ID='TISHERMA')"
						End if
							If HUB2="y" then
								SQL = SQL&" OR ((Fl_st_ID='xxx')"
							End if
							'''''''''THIS MIGHT NEED TO GO AWAY!							
						If HUB2="yxxx" then
							SQL = SQL&" OR (Fl_st_ID='12201')"
							SQL = SQL&" OR (Fl_st_ID='12203')"
							SQL = SQL&" OR (Fl_st_ID='6430')"
							SQL = SQL&" OR (Fl_st_ID='6412')"
							SQL = SQL&" OR (Fl_st_ID='13601')"
							SQL = SQL&" OR (Fl_st_ID='12500')"
							SQL = SQL&" OR (Fl_st_ID='13020')"
							SQL = SQL&" OR (Fl_st_ID='7800')"
							SQL = SQL&" OR (Fl_st_ID='7839')"
							SQL = SQL&" OR (Fl_st_ID='13353')"
							SQL = SQL&" OR (Fl_st_ID='13536')"
							SQL = SQL&" OR (Fl_st_ID='13121')"
							SQL = SQL&" OR (Fl_st_ID='6550')"
							SQL = SQL&" OR (Fl_st_ID='13011')"
							SQL = SQL&" OR (Fl_st_ID='13570'))"
						End if	
						If trim(LocationCode)="D1" then
							SQL = SQL&" OR (Fl_st_ID='D7')"
							SQL = SQL&" OR (Fl_st_ID='P1')"
						End if						
						If HUB2<>"y" AND LocationCode <>"SRHUB" AND LocationCode<>"D6N1B" AND LocationCode<>"DOCK7" then
							SQL = SQL&")"
						End if													
					End if
					'Response.Write "594 SQL="&SQL&"<BR>"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''					
					oRs.Open SQL, DATABASE, 1, 3
					If oRs.EOF then
						ErrorMessage=ErrorMessage&" "&FormBarCode(q)&"  is not accepted.<br>Check Paper Work/Call Supervisor (1)<br>"
					End if
					If not oRs.eof then
						TheJobNumber = oRs("fh_id")
            D6ref = oRs("rf_ref")
            fh_user3 = oRs("fh_user3")
						fl_pkey=oRs("fl_pkey")
						TheBarCode = FormBarCode(q)
						jobcomment = oRs("fl_sf_comment")
						stcomment = oRs("fl_st_comment")
            co_email = oRs("fh_co_email")
						If Signature>"" then
						End if
						Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open DATABASE
						''''UPDATES THE WAFER
						if LocationCode <> "SRHUB" then
              l_cSQL = "UPDATE FCREFS SET ref_status = '"&ORDERSTATUS&"' "
  						If Trim(PODID)>"" then
  							l_cSQL = l_cSQL&", POD = '"&PODID&"' "
  						End if
  						l_cSQL = l_cSQL&" WHERE (rf_fh_id = '"&TheJobNumber&"') AND (rf_ref = '" &TheBarCode& "')"
  						
              'Response.Write "600 l_cSQL="&l_cSQL&"<BR>"
  
              oConn.Execute(l_cSQL)
            end if
            
            if LocationCode = "D6N1B" then
                stcomment = "**Elevator down, shipment delivered to D6N1** " & stcomment
                D6SQL = "UPDATE FCLEGS SET fl_st_comment = '" & stcomment & "' WHERE fl_fh_id='" & JobNumber & "'"
                oConn.Execute(D6SQL)
                
                Body = "The elevator was out of order so your shipment was delivered to D6N1.<br><br>"& _
    						"It was job #" & TheJobNumber & "<br>Document # " & D6ref & "<br><br>"& _
    						"You can pick it up at any time.<br><br>"& _
                "FleetX" 
    						'Set objMail = CreateObject("CDONTS.Newmail")
    						'objMail.From = "FleetX@LogisticorpGroup.com"
                varTo = co_email
                varcc = "mark.maggiore@logisticorp.us"
    
    						varSubject = D6ref&" has been delivered to D6N1"
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
            end if  
             if LocationCode = "DOCK7" then
                stcomment = "**Elevator down, shipment delivered to DOCK7** " & stcomment
                D6SQL = "UPDATE FCLEGS SET fl_st_comment = '" & stcomment & "' WHERE fl_fh_id='" & JobNumber & "'"
                oConn.Execute(D6SQL)
                
                Body = "The elevator was out of order so your shipment was delivered to DOCK7 (Coordinates: 513A1396).<br><br>"& _
    						"It was job #" & TheJobNumber & "<br>Document # " & D6ref & "<br><br>"& _
    						"You can pick it up at any time.<br><br>"& _
                "FleetX" 
    						'Set objMail = CreateObject("CDONTS.Newmail")
    						'objMail.From = "FleetX@LogisticorpGroup.com"
                varTo = co_email
                varcc = "mark.maggiore@logisticorp.us"
    
    						varSubject = D6ref&" has been delivered to DOCK7"
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
            end if            
                      
						''''''''''''Determines if needs a POD or not''''''''''''''''''''''''
						Set Recordset1 = Server.CreateObject("ADODB.Recordset")
						Recordset1.ActiveConnection = DATABASE
						Recordset1.Source = "SELECT fl_pkey, fl_st_id FROM fclegs WHERE (fl_PKey='"& fl_PKey+1 &"') and (fl_fh_id='"& TheJobNumber &"')"
						
						'response.write "XXXXXRecordset1.Source="&Recordset1.Source&"<BR>"
						
						Recordset1.CursorType = 0
						Recordset1.CursorLocation = 2
						Recordset1.LockType = 1
						Recordset1.Open()
						Recordset1_numRows = 0
						If Recordset1.eof then
							CloseThis="y"
						End if			
						If not Recordset1.eof then
							CloseThis="n"
							NextToAddress=trim(Recordset1("fl_st_id"))
							
							'Response.Write "GOT HERE!!  ROLLOVER=YES<BR>"
						End if
						Set Recordset1 = Nothing					
						''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''						
						''''''''''''''''''''''''''''''''''''''''
						Set oRs = Server.CreateObject("ADODB.Recordset")
						oRs.CursorLocation = 3
						oRs.CursorType = 3
						oRs.ActiveConnection = DATABASE	
						''''''CHECKS TO SEE IF ITS OKAY TO CHANGE THE ORDER STATUS
						If LocationCode = "SRHUB" then
              ' update just fh_status to ARV abd vehicle to 912780 for material handler
									Set oConn43 = Server.CreateObject("ADODB.Connection")
									oConn43.ConnectionTimeout = 100
									oConn43.Provider = "MSDASQL"
									oConn43.Open DATABASE
									'l_cSQL = "UPDATE fcfgthd SET fh_status = 'ARV' "
									'l_cSQL = l_cSQL&" WHERE (fh_id = '"&TheJobNumber&"')"
									'response.write "654 l_cSQL="&l_cSQL&"<BR>"
									'oConn43.Execute(l_cSQL) 
                  '' changed to use stored proc for update, per Mark - but have to reset the leg_status back to 'c' manually
								  oConn43.Execute "PHONE_CHANGE_STATUS '" & TheJobNumber & "', '53', 'ARV', '', '',  '"& UserID &"', '"& UnitID &"'" 
                  l_cSQL = "UPDATE fclegs SET fl_un_id = 912780, fl_leg_status='c' WHERE (fl_fh_id = '"&TheJobNumber&"')"                
									oConn43.Execute(l_cSQL) 
									oConn43.Close
									Set oConn43=Nothing	
            Else
            
            SQL = "SELECT rf_fh_id FROM fcrefs"
						SQL = SQL&" WHERE (rf_fh_id='"&TheJobNumber&"')"
						'SQL = SQL&" AND (Ref_Status='o')"
						If trim(locationcode)="EBHUB" or HUB="y" then
							SQL = SQL&" AND (((Ref_Status<>'H') AND (Ref_Status<>'c')) or (Ref_Status IS NULL))"
							else
							SQL = SQL&" AND ((Ref_Status<>'c') or (Ref_Status IS NULL))"
						
						End if
						'response.write "641 SQL="&SQL&"<BR>"
						'response.write "Hub="&Hub&"<BR>"
						'response.write "CloseThis="&CloseThis&"<BR>"
						
						oRs.Open SQL, DATABASE, 1, 3
						If oRs.EOF then
							 'response.write "647 EOF<br>"
              If (Hub="y" or CloseThis="n") then
								'response.write "649 - #1 PHONE_CHANGE_STATUS '" & TheJobNumber & "', '53', 'ARV', '', '',  '"& UserID &"', '"& UnitID &"'" 
								oConn.Execute "PHONE_CHANGE_STATUS '" & TheJobNumber & "', '53', 'ARV', '', '',  '"& UserID &"', '"& UnitID &"'" 
								''''''''''''''EMAILS KEITH WHEN HITS HUB AND HAS A SPECIAL INSTRUCTION!!!!''''''''''
								If BillToID="48" AND len(trim(jobcomment))>1 then
									Body = "RE:&nbsp;&nbsp; Job #&nbsp;&nbsp;"& TheJobNumber &"<br><br>"   & _
									"The driver has just dropped this job at the HUB.  It has the following comments:<br><br>"   & _
									" "& jobcomment &"<br><br>"  & _
									"This has been an automatic notification.<br><br>"   & _
									"Have a nice day.<br>"
									Recipient="Keith Chitwood"
									Email="x0035291@ti.com"
									'Email="mark@maggiore.net"
									'Set objMail = CreateObject("CDONTS.Newmail")
									'objMail.From = "FleetX@LogisticorpGroup.com"
									varTo = Email
									varcc = "mark.maggiore@logisticorp.us"
									varSubject = "Job #"& TheJobNumber &" Comment"
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
								End if								
									''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
									''''''''''Do not make this LIVE yet, until KWE is OKAY to make live'''''''''''
									''''''''''THEN FIX FOR STOCKROOM!!!!!!!!!!!!!!!!!''''''''''''''''''''''''''''
									''''''''''''''''''''''''''''''''''''''''''''''''''
								'If FixedForStockroom="yes" then
								''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
								If trim(vehicleID)="199xxx" then
									Set oConn43 = Server.CreateObject("ADODB.Connection")
									oConn43.ConnectionTimeout = 100
									oConn43.Provider = "MSDASQL"
									oConn43.Open DATABASE
											Nextfl_un_id="198"
											Nextfl_un_id="198"
									''''''''''''END ROUTING PART'''''''''''''''''''''''''''''								
									''''UPDATES THE NEXT LEG TO MAKE IT LIVE
									l_cSQL = "UPDATE FCLEGS SET fl_un_id='"& Nextfl_un_id &"', fl_un_id='"& Nextfl_un_id &"', fl_Leg_Status = 'c' "
									l_cSQL = l_cSQL&" WHERE (fl_fh_id = '"&TheJobNumber&"')"
									'response.write "l_cSQL="&l_cSQL&"<BR>"
									oConn43.Execute(l_cSQL)
									oConn43.Close
									Set oConn43=Nothing	
								End if
                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                ''''''''''''''''''''STOCKROOM HUB INFO''''''''''''''''''''''''''''''
                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
								If BillToID="91" then
                              Response.write "697 SELECT * FROM fclegs WHERE (fl_PKey='"& fl_PKey &"') and (fl_fh_id='"& TheJobNumber &"')<br><BR>"
			                        Set Recordset1 = Server.CreateObject("ADODB.Recordset")
			                        Recordset1.ActiveConnection = DATABASE
			                        Recordset1.Source = "SELECT * FROM fclegs WHERE (fl_PKey='"& fl_PKey &"') and (fl_fh_id='"& TheJobNumber &"')"
			                        Recordset1.CursorType = 0
			                        Recordset1.CursorLocation = 2
			                        Recordset1.LockType = 1
			                        Recordset1.Open()
			                        Recordset1_numRows = 0
			                        If not Recordset1.eof then
                                        varfl_fh_id=TheJobNumber
                                        varfl_joborder=0
				                        varfl_sf_id="EBHUB"
                                        varfl_sf_name="EBHUB"
                                        varfl_st_id=Recordset1("fl_st_id")
                                        varfl_st_name=Recordset1("fl_st_name")
                                        varfl_st_clname=Recordset1("fl_st_clname")
                                        varfl_st_cfname=Recordset1("fl_st_cfname")
                                        varfl_st_phone=Recordset1("fl_st_phone")
                                        varfl_st_addr1=Recordset1("fl_st_addr1")
                                        varfl_st_addr2=Recordset1("fl_st_addr2")
                                        varfl_st_city=Recordset1("fl_st_city")
                                        varfl_st_state=Recordset1("fl_st_state")
                                        varfl_st_country=Recordset1("fl_st_country")
                                        varfl_st_zip=Recordset1("fl_st_zip")
                                        varfl_estimate=0
                                        varfl_permit=0
                                        varfl_un_id=""
                                        varfl_un_id=""
                                        varfl_st_odm=0
                                        varfl_end_odm=0
                                        varfl_trf_mi=0
                                        varfl_load_mi=0
                                        varfl_empty_mi=0
                                        varfl_toll_mi=0
                                        varfl_load_ti=0
                                        varfl_unld_ti=0
                                        varfl_trip_ti=0
                                        varfl_rt_type=""
                                        varfl_totrate=0
                                        varfl_wt_xc=0
                                        varfl_wgt_xc=0
                                        varfl_pmrate=0
                                        varfl_escrate=0
                                        varfl_codconrt=0
                                        varfl_codshprt=0
                                        varfl_miscrate=0
                                        varfl_mrdesc=""
                                        varfl_pdrate=0
                                        varfl_estrate=0
                                        varfl_flatrt=0
                                        varfl_prirt=0
                                        varfl_pj_rt=0
                                        varfl_sfstrt=0
                                        varfl_codc_est=0
                                        varfl_cods_est=0
                                        varfl_sf_comment=Recordset1("fl_sf_comment")
                                        varfl_st_comment=Recordset1("fl_st_comment")
                                        varfl_sf_area=Recordset1("fl_sf_area")
                                        varfl_st_area=Recordset1("fl_st_area")
                                        varfl_t_disp="1/1/1900"
                                        varfl_t_acc="1/1/1900"
                                        varfl_t_atp="1/1/1900"
                                        varfl_t_int="1/1/1900"
                                        varfl_t_atd="1/1/1900"
                                        varfl_t_und="1/1/1900"
                                        varfl_st_rta=Recordset1("fl_st_rta")
                                        varfl_sf_rta=Recordset1("fl_sf_rta")
                                        varfl_weight=Recordset1("fl_weight")
                                        varfl_pod=Recordset1("fl_pod")
                                        varfl_wait_t=Recordset1("fl_wait_t")
                                        varfl_feesadv=Recordset1("fl_feesadv")
                                        varfl_fadesc=Recordset1("fl_fadesc")
                                        varfl_rndtrip=Recordset1("fl_rndtrip")
                                        varfl_rndt_rt=Recordset1("fl_rndt_rt")
                                        vartimestamp_column=Recordset1("timestamp_column")
                                        varfl_zipmlrt=Recordset1("fl_zipmlrt")
                                        varfl_numboxes=Recordset1("fl_numboxes")
                                        varfl_hascod=Recordset1("fl_hascod")
                                        varfl_boxrt=Recordset1("fl_boxrt")
                                        varfl_disp=Recordset1("fl_disp")
                                        varfl_sf_fullname=Recordset1("fl_sf_fullname")
                                        varfl_st_fullname=Recordset1("fl_st_fullname")
                                        varfl_user1=Recordset1("fl_user1")
                                        varfl_user2=Recordset1("fl_user2")
                                        varfl_podreq=Recordset1("fl_podreq")
                                        varfl_rentmin=Recordset1("fl_rentmin")
                                        varfl_rentrt=Recordset1("fl_rentrt")
                                        varfl_boxtype=Recordset1("fl_boxtype")
                                        varfl_permirt=Recordset1("fl_permirt")
                                        varfl_dimwgt=Recordset1("fl_dimwgt")
                                        varfl_dwfact=Recordset1("fl_dwfact")
                                        varfl_pay_on=Recordset1("fl_pay_on")
                                        varfl_ah_rt=Recordset1("fl_ah_rt")
                                        varfl_ah_code=Recordset1("fl_ah_code")
                                        varfl_pay_upd=Recordset1("fl_pay_upd")
                                        varfl_cntyrt=Recordset1("fl_cntyrt")
                                        varfl_user3=Recordset1("fl_user3")
                                        varfl_user4=Recordset1("fl_user4")
                                        varfl_billcd=Recordset1("fl_billcd")
                                        varfl_sf_apt=Recordset1("fl_sf_apt")
                                        varfl_st_apt=Recordset1("fl_st_apt")
                                        varfl_firstdrop=Recordset1("fl_firstdrop")
                                        varfl_seconb=Recordset1("fl_seconb")
                                        varfl_secacc=Recordset1("fl_secacc")
                                        varfl_pu_driver=Recordset1("fl_pu_driver")
                                        varfl_pu_vehicle=Recordset1("fl_pu_vehicle")
                                        varfl_do_driver=Recordset1("fl_do_driver")
                                        varfl_do_vehicle=Recordset1("fl_do_vehicle")
                                        varfl_job_closed=Recordset1("fl_job_closed")
                                        varfl_leg_status=Recordset1("fl_leg_status")
                                        varfl_FinalDestination=Recordset1("fl_FinalDestination")
			                        End if
			                        Set Recordset1 = Nothing



			                        Set Recordset1 = Server.CreateObject("ADODB.Recordset")
			                        Recordset1.ActiveConnection = DATABASE
			                        Recordset1.Source = "SELECT sv_val FROM fcsysval where sv_id='sy_number'"
			                        Recordset1.CursorType = 0
			                        Recordset1.CursorLocation = 2
			                        Recordset1.LockType = 1
			                        Recordset1.Open()
			                        Recordset1_numRows = 0
			                        If not Recordset1.eof then
                                        sy_number=Recordset1("sv_val")
                                        NextLegNumber=sy_number+1
		                            End if
			                        Set Recordset1 = Nothing

                                    'Response.write "*******************************<BR>"
                                    'Response.write "sy_number="&sy_number&"<BR>"
                                    'Response.write "NextLegNumber="&NextLegNumber&"<BR>"
                                    'Response.write "*******************************<BR>"
									Set oConn43 = Server.CreateObject("ADODB.Connection")
									oConn43.ConnectionTimeout = 100
									oConn43.Provider = "MSDASQL"
									oConn43.Open DATABASE
									l_cSQL = "UPDATE fcsysval SET sv_val = '"& NextLegNumber &"' where sv_id='sy_number' "
									'response.write "l_cSQL="&l_cSQL&"<BR>"
									oConn43.Execute(l_cSQL)






							        Set Recordset1 = Server.CreateObject("ADODB.Recordset")
								        Recordset1.Open "FCLegs", DATABASE, 2, 2
								        Recordset1.addnew
                                        Recordset1("fl_pkey")=NextLegNumber
                                        Recordset1("fl_fh_id")=varfl_fh_id
                                        Recordset1("fl_joborder")=varfl_joborder
				                        Recordset1("fl_sf_id")=varfl_sf_id
                                        Recordset1("fl_sf_name")=" "
                                        Recordset1("fl_sf_clname")=" "
                                        Recordset1("fl_sf_cfname")=" "
                                        Recordset1("fl_sf_phone")=" "
                                        Recordset1("fl_sf_addr1")=" "
                                        Recordset1("fl_sf_addr2")=" "
                                        Recordset1("fl_sf_city")="Dallas"
                                        Recordset1("fl_sf_state")="TX"
                                        Recordset1("fl_sf_country")="USA"
                                        Recordset1("fl_sf_zip")="75248"                           
                                        Recordset1("fl_st_id")=varfl_st_id
                                        Recordset1("fl_st_name")=varfl_st_name
                                        Recordset1("fl_st_clname")=varfl_st_clname
                                        Recordset1("fl_st_cfname")=varfl_st_cfname
                                        Recordset1("fl_st_phone")=varfl_st_phone
                                        Recordset1("fl_st_addr1")=varfl_st_addr1
                                        Recordset1("fl_st_addr2")=varfl_st_addr2
                                        Recordset1("fl_st_city")=varfl_st_city
                                        Recordset1("fl_st_state")=varfl_st_state
                                        Recordset1("fl_st_country")=varfl_st_country
                                        Recordset1("fl_st_zip")=varfl_st_zip
                                        Recordset1("fl_estimate")=varfl_estimate
                                        Recordset1("fl_permit")=varfl_permit
                                        Recordset1("fl_un_id")=varfl_un_id
                                        Recordset1("fl_un_id")=varfl_un_id
                                        Recordset1("fl_st_odm")=varfl_st_odm
                                        Recordset1("fl_end_odm")=varfl_end_odm
                                        Recordset1("fl_trf_mi")=varfl_trf_mi
                                        Recordset1("fl_load_mi")=varfl_load_mi
                                        Recordset1("fl_empty_mi")=varfl_empty_mi
                                        Recordset1("fl_toll_mi")=varfl_toll_mi
                                        Recordset1("fl_load_ti")=varfl_load_ti
                                        Recordset1("fl_unld_ti")=varfl_unld_ti
                                        Recordset1("fl_trip_ti")=varfl_trip_ti
                                        Recordset1("fl_rt_type")=varfl_rt_type
                                        Recordset1("fl_totrate")=varfl_totrate
                                        Recordset1("fl_wt_xc")=varfl_wt_xc
                                        Recordset1("fl_wgt_xc")=varfl_wgt_xc
                                        Recordset1("fl_pmrate")=varfl_pmrate
                                        Recordset1("fl_escrate")=varfl_escrate
                                        Recordset1("fl_codconrt")=varfl_codconrt
                                        Recordset1("fl_codshprt")=varfl_codshprt
                                        Recordset1("fl_miscrate")=varfl_miscrate
                                        Recordset1("fl_mrdesc")=varfl_mrdesc
                                        Recordset1("fl_pdrate")=varfl_pdrate
                                        Recordset1("fl_estrate")=varfl_estrate
                                        Recordset1("fl_flatrt")=varfl_flatrt
                                        Recordset1("fl_prirt")=varfl_prirt
                                        Recordset1("fl_pj_rt")=varfl_pj_rt
                                        Recordset1("fl_sfstrt")=varfl_sfstrt
                                        Recordset1("fl_codc_est")=varfl_codc_est
                                        Recordset1("fl_cods_est")=varfl_cods_est
                                        Recordset1("fl_sf_comment")=varfl_sf_comment
                                        Recordset1("fl_st_comment")=varfl_st_comment
                                        Recordset1("fl_sf_area")=varfl_sf_area
                                        Recordset1("fl_st_area")=varfl_st_area
                                        Recordset1("fl_t_disp")=varfl_t_disp
                                        Recordset1("fl_t_acc")=varfl_t_acc
                                        Recordset1("fl_t_atp")=varfl_t_atp
                                        Recordset1("fl_t_int")=varfl_t_int
                                        Recordset1("fl_t_atd")=varfl_t_atd
                                        Recordset1("fl_t_und")=varfl_t_und
                                        Recordset1("fl_st_rta")=varfl_st_rta
                                        Recordset1("fl_sf_rta")=varfl_sf_rta
                                        Recordset1("fl_weight")=varfl_weight
                                        Recordset1("fl_pod")=varfl_pod
                                        Recordset1("fl_wait_t")=varfl_wait_t
                                        Recordset1("fl_feesadv")=varfl_feesadv
                                        Recordset1("fl_fadesc")=varfl_fadesc
                                        Recordset1("fl_rndtrip")=varfl_rndtrip
                                        Recordset1("fl_rndt_rt")=varfl_rndt_rt
                                        Recordset1("timestamp_column")=vartimestamp_column
                                        Recordset1("fl_zipmlrt")=varfl_zipmlrt
                                        Recordset1("fl_numboxes")=varfl_numboxes
                                        Recordset1("fl_hascod")=varfl_hascod
                                        Recordset1("fl_boxrt")=varfl_boxrt
                                        Recordset1("fl_disp")=varfl_disp
                                        Recordset1("fl_sf_fullname")=varfl_sf_fullname
                                        Recordset1("fl_st_fullname")=varfl_st_fullname
                                        Recordset1("fl_user1")=varfl_user1
                                        Recordset1("fl_user2")=varfl_user2
                                        Recordset1("fl_podreq")=varfl_podreq
                                        Recordset1("fl_rentmin")=varfl_rentmin
                                        Recordset1("fl_rentrt")=varfl_rentrt
                                        Recordset1("fl_boxtype")=varfl_boxtype
                                        Recordset1("fl_permirt")=varfl_permirt
                                        Recordset1("fl_dimwgt")=varfl_dimwgt
                                        Recordset1("fl_dwfact")=varfl_dwfact
                                        Recordset1("fl_pay_on")=varfl_pay_on
                                        Recordset1("fl_ah_rt")=varfl_ah_rt
                                        Recordset1("fl_ah_code")=varfl_ah_code
                                        Recordset1("fl_pay_upd")=varfl_pay_upd
                                        Recordset1("fl_cntyrt")=varfl_cntyrt
                                        Recordset1("fl_user3")=varfl_user3
                                        Recordset1("fl_user4")=varfl_user4
                                        Recordset1("fl_billcd")=varfl_billcd
                                        Recordset1("fl_sf_apt")=varfl_sf_apt
                                        Recordset1("fl_st_apt")=varfl_st_apt
                                        Recordset1("fl_firstdrop")=varfl_firstdrop
                                        Recordset1("fl_seconb")=varfl_seconb
                                        Recordset1("fl_secacc")=varfl_secacc
                                        Recordset1("fl_pu_driver")=varfl_pu_driver
                                        Recordset1("fl_pu_vehicle")=varfl_pu_vehicle
                                        Recordset1("fl_do_driver")=varfl_do_driver
                                        Recordset1("fl_do_vehicle")=varfl_do_vehicle
                                        Recordset1("fl_job_closed")=varfl_job_closed
                                        Recordset1("fl_leg_status")=varfl_leg_status
                                        Recordset1("fl_FinalDestination")=varfl_FinalDestination
								        Recordset1.update
								        Recordset1.close			
							        set Recordset1 = nothing	





									Set oConn43 = Server.CreateObject("ADODB.Connection")
									oConn43.ConnectionTimeout = 100
									oConn43.Provider = "MSDASQL"
									oConn43.Open DATABASE
									''''UPDATES CURRENT LEG TO INDICATE DROPPED!
									l_cSQL = "UPDATE FCLEGS SET fl_Leg_Status = 'd' "
									l_cSQL = l_cSQL&" WHERE (fl_fh_id = '"&TheJobNumber&"') AND (fl_pkey = '" & fl_pkey & "')"
									'response.write "l_cSQL="&l_cSQL&"<BR>"
									oConn43.Execute(l_cSQL)
									''''''''''''ROUTING PART'''''''''''''''''''''''''''''''''
									'Response.Write "NextTOADDRESS="&NextToAddress&"<BR>"
									Select Case NextToAddress
										Case "12203", "BASIN" '''''Stafford HUB
											Nextfl_un_id="4"
											Nextfl_un_id="4"
										Case Else
											Nextfl_un_id="SRB"
											Nextfl_un_id="124"
									End Select
									''''''''''''END ROUTING PART'''''''''''''''''''''''''''''								
									''''UPDATES THE NEXT LEG TO MAKE IT LIVE
									l_cSQL = "UPDATE FCLEGS SET fl_un_id='"& Nextfl_un_id &"', fl_un_id='"& Nextfl_un_id &"', fl_Leg_Status = 'c' "
									l_cSQL = l_cSQL&" WHERE (fl_fh_id = '"&TheJobNumber&"') AND (fl_pkey <> '" & fl_pkey & "')"
									'response.write "l_cSQL="&l_cSQL&"<BR>"
									oConn43.Execute(l_cSQL)
									oConn43.Close
									Set oConn43=Nothing	
								End if
                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                '''''''''''''''''''END STOCKROOM HUB INFO'''''''''''''''''''''''''''
                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


								''''''''''''''''''''''''''''''''''''''''								
								If BillToID="48" or trim(vehicleID)="198" then
									Set oConn43 = Server.CreateObject("ADODB.Connection")
									oConn43.ConnectionTimeout = 100
									oConn43.Provider = "MSDASQL"
									oConn43.Open DATABASE
									''''UPDATES CURRENT LEG TO INDICATE DROPPED!
									l_cSQL = "UPDATE FCLEGS SET fl_Leg_Status = 'd' "
									l_cSQL = l_cSQL&" WHERE (fl_fh_id = '"&TheJobNumber&"') AND (fl_pkey = '" & fl_pkey & "')"
									'response.write "l_cSQL="&l_cSQL&"<BR>"
									oConn43.Execute(l_cSQL)
									''''''''''''ROUTING PART'''''''''''''''''''''''''''''''''
									'Response.Write "NextTOADDRESS="&NextToAddress&"<BR>"
									Select Case NextToAddress
										Case "12203", "BASIN" '''''Stafford HUB
											Nextfl_un_id="4"
											Nextfl_un_id="4"
										Case "12201" '''''Stafford final destination
											Nextfl_un_id="4"
											Nextfl_un_id="4"
										Case "6430" '''''Sherman final destination
											Nextfl_un_id="6"
											Nextfl_un_id="6"
										Case "6550", "7800", "RFAB", "13560", "13570" '''''Spring Creek final destination
											Nextfl_un_id="3"
											Nextfl_un_id="3"
										Case "12500", "13121", "13353", "13532", "13353-7", "13536F" '''''Spring Creek final destination
											Nextfl_un_id="7"
											Nextfl_un_id="7"
										Case Else
											Nextfl_un_id="8"
											Nextfl_un_id="8"
									End Select
									''''''''''''END ROUTING PART'''''''''''''''''''''''''''''								
									''''UPDATES THE NEXT LEG TO MAKE IT LIVE
									l_cSQL = "UPDATE FCLEGS SET fl_un_id='"& Nextfl_un_id &"', fl_un_id='"& Nextfl_un_id &"', fl_Leg_Status = 'c' "
									l_cSQL = l_cSQL&" WHERE (fl_fh_id = '"&TheJobNumber&"') AND (fl_pkey = '" & fl_pkey+1 & "')"
									'response.write "l_cSQL="&l_cSQL&"<BR>"
									oConn43.Execute(l_cSQL)
									oConn43.Close
									Set oConn43=Nothing	
								End if
								''''''''''''''''''''''''''''''''''''''''''''''''''	
								else
								'response.write "#2 PHONE_CHANGE_STATUS '" & TheJobNumber & "', '9', 'CLS', '', '',  '"& UserID &"', '"& UnitID &"'" 
                                 Response.write "Database="&Database&"<BR>"
						            Set oConn33 = Server.CreateObject("ADODB.Connection")
						            oConn33.ConnectionTimeout = 100
						            oConn33.Provider = "MSDASQL"
						            oConn33.Open DATABASE
								response.write "#3 PHONE_CHANGE_STATUS '" & TheJobNumber & "', '9', 'CLS', '', '',  '"& UserID &"', '"& UnitID &"'" 
                               '  Response.write "Database3="&Database3&"<BR>"
								oConn33.Execute "PHONE_CHANGE_STATUS '" & TheJobNumber & "', '9', 'CLS', '', '',  '"& UserID &"', '"& UnitID &"'" 
								'Response.Write "BillTOID="&BIllTOID&"<BR>"
                                If trim(fh_user3)>"" then
                                        
						            Set oConn3 = Server.CreateObject("ADODB.Connection")
						            oConn3.ConnectionTimeout = 100
						            oConn3.Provider = "MSDASQL"
						            oConn3.Open DATABASE3
								response.write "#4 PHONE_CHANGE_STATUS '" & fh_user3 & "', '9', 'CLS', '', '',  '"& UserID &"', '"& UnitID &"'" 
                                 Response.write "Database3="&Database3&"<BR>"
							            oConn3.Execute "PHONE_CHANGE_STATUS '" & fh_user3 & "', '9', 'CLS', '', '',  '"& UserID &"', '"& UnitID &"'" 
                                End if

'''''''''''''''''''''''''FINDS BILL TO ID FOR THE INDIVIDUAL JOB!
                    Set oConn89 = Server.CreateObject("ADODB.Connection")
                    oConn89.ConnectionTimeout = 100
                    oConn89.Provider = "MSDASQL"
                    oConn89.Open DATABASE
		                 SQL="SELECT fh_bt_id FROM fcfgthd where (fh_id='"& TheJobNumber &"')"
	                    'Response.Write "LINE 998 SQL="&SQL&"<BR>"
	                    SET oRs89 = oConn89.Execute(Sql)
	                    If not oRs89.EOF then 
                            BillToID=trim(oRs89("fh_bt_id"))
                            Else
                            BillToID="9876543210"
                        End if                      
                        oRs89.Close
		                Set oRs89=Nothing
''''''''''''''''''''''''ENDS FINDS BILL TO ID







								If trim(BillToID)="91" then
                                'Response.write "LINE 1016 GOT HERE AND NOW!!!<BR>"
                                'Response.Write "LocationCode="&LocationCode&"<BR>"
'''''''''''BEGIN PEDRO CONTRERAS REQUEST
				If (Ucase(LocationCode)="D6N1") or (Ucase(LocationCode)="D6N2") or (Ucase(LocationCode)="D6W3") then
                    Set oConn65 = Server.CreateObject("ADODB.Connection")
                    oConn65.ConnectionTimeout = 100
                    oConn65.Provider = "MSDASQL"
                    oConn65.Open DATABASE
		                 SQL="SELECT RF_Ref, fh_co_id FROM fcrefs INNER JOIN fcfgthd ON fcrefs.rf_fh_id = fcfgthd.fh_id where (fh_id='"& JobNumber &"')"
	                    'Response.Write "SQL="&SQL&"<BR>"
	                    SET oRs65 = oConn65.Execute(Sql)
	                    Do while not oRs65.EOF 
	                        'Response.Write "got here...okay?" 
			                'Response.Write "FABID="&FABID&"<BR>"
			                'Response.Write "SQL="&SQL&"<BR>"
                            temp_ref=trim(oRs65("RF_Ref"))
                            temp_fh_co_id=lcase(trim(oRs65("fh_co_id")))
                            AllRefs=AllRefs & "#" & Temp_ref & "<br>"
                           ' Response.Write "temp_ref="&temp_ref&"<BR>"
                           ' Response.Write "AllRefs="&AllRefs&"<BR>"
				        oRs65.movenext
				        LOOP                        
                        oRs65.Close
		                Set oRs65=Nothing
                        'Response.write "LocationCode="&LocationCode&"<BR>"
                        'Response.write "temp_fh_co_id="&temp_fh_co_id&"<BR>"
                        Select Case LocationCode
                            Case "D6N1", "D6N2"
                                Select Case temp_fh_co_id
                                    Case "a0272321","a0200672","a0459390","a0869103","a0201981","a0667876","a0272380","a0200166"
                                        SendD6Email="y"
                                        mailtovar="dm6impeetechtext@list.ti.com;mark.maggiore@logisticorp.us"
                                        'mailtovar="mark.maggiore@logisticorp.us"
                                        'Response.write "GOT HERE #1<BR>"
                                    Case else
                                    Sendd6Email="n"
                                End Select
                            Case "D6W3"
                                Select Case temp_fh_co_id
                                    Case "a0460940","a0225597","a0342163","a0460681","a0218992","a0215134","a0273050","a0865786"
                                        SendD6Email="y"
                                        mailtovar="dm6thermeetechtext@list.ti.com;mark.maggiore@logisticorp.us"
                                        'mailtovar="mark.maggiore@logisticorp.us"
                                        'Response.write "GOT HERE #2<BR>"
                                   Case else
                                        Sendd6Email="n"
                                End Select
                        End select

                    '''''''''''ITAR NOTIFICATION'''''''''''''''''''''''''''''''''''''''''''''''''
                        If SendD6Email="y" then
						Body = "PartNumber:<br><br>"& AllRefs &"<BR>has/have just been delivered to "& LocationCode &".<br><br>"& _
						"It was job #" & TheJobNumber & "<br><br>"& _
						"LogistiCorp" 
						'Recipient = "mark.maggiore@logisticorp.us"
						'Set objMail = CreateObject("CDONTS.Newmail")
						'objMail.From = "FleetX@LogisticorpGroup.com"
						varTo = mailtovar
                        'objMail.cc = "FleetX@LogisticorpGroup.com"

						varSubject = Temp_Ref&" has been delivered to "&LocationCode
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
                        End if
                    End if	
''''''''''END PEDRO CONTRERAS REQUEST




								''''''''FIND NOTIFICATION INFORMATION!!!!!
									Set Recordset1 = Server.CreateObject("ADODB.Recordset")
									Recordset1.ActiveConnection = DATABASE
									Recordset1.Source = "SELECT * FROM DeliveryNotifications WHERE (fh_id='"& TheJobNumber &"')"
									'response.write "LINE 1093 Recordset1.Source="&Recordset1.Source&"<BR>"
									Recordset1.CursorType = 0
									Recordset1.CursorLocation = 2
									Recordset1.LockType = 1
									Recordset1.Open()
									Recordset1_numRows = 0
									If not Recordset1.eof then
                                        
										mailref=Recordset1("Ref_ID")
										mailmaterial=Recordset1("Material")
										mailmaterialdescription=Recordset1("MaterialDescription")
										mailemailaddress=Recordset1("EmailAddress")
                                        If ucase(trim(LocationCode))="CSSF-SR" then
                                            MailEmailAddress=MailEmailAddress&";bp@ti.com"
                                        End if
										'Response.Write "GOT HERE!!  ROLLOVER=YES<BR>"
										''''''''''SEND NOTIFICATION!''''''''''''''''
										Body ="This order for P/N "& MailMaterial &" &nbsp;&nbsp;"& MailMaterialDescription &" was delivered to "   & _
										"drop zone "& LocationCode &" at "& now() & "<BR><BR>REPORT TRANSPORTATION PROBLEMS TO: FleetX@LogisticorpGroup.com<BR><BR>REPORT MATERIAL ISSUES TO: dsb_mrp@list.ti.com<br><br>"
										'Recipient=FirstName&" "&LastName
										'Email="KWETI.Mailbox@am.kwe.com"
										'Email="mark@maggiore.net"
										'Set objMail = CreateObject("CDONTS.Newmail")
										'objMail.From = "FleetX@LogisticorpGroup.com"
                                        'objMail.From = "mark@maggiore.net"
										varTo = MailEmailAddress
										'''''objMail.bcc = "j-overman@ti.com;mark.maggiore@logisticorp.us"
										'objMail.cc = "FleetX@LogisticorpGroup.com"
										'objMail.bcc = "mark.maggiore@logisticorp.us"
										varSubject = "Electronic Sales Order #"& mailref &"(FleetX Job # "& TheJobNumber &")"
										'objMail.MailFormat = cdoMailFormatMIME
										'objMail.BodyFormat = cdoBodyFormatHTML
										'objMail.Body = Body
										'objMail.Send


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
                                        'RESPONSE.WRITE "LINE 1123 SENT THAT EMAIL!!!!<BR>"
                                        'Response.write "mailref="&mailref&"<BR>"
                                        'Response.write "Body="&Body&"<BR>"
                                        'Response.write "mailemailaddress="&mailemailaddress&"<BR>"

										Set objMail = Nothing										
										''''''''''''''''''''''''''''''''''''''''''''
									End if
									Set Recordset1 = Nothing								
								''''''''END FIND NOTIFICATION INFORMATION!!!!!!!!
								end if
							End if
						End if
						oRs.Close
						Set oRs=Nothing							
						oConn.Close	
						Set oConn=Nothing	
						TheJobNumber=""
						TheBarCode=""	
                        BillToID=""
					end if
          End if
				End if
			Next
		End if
		Set oConn=Nothing	
END IF
%>
<body onload="document.Form1.FormBarCode1.focus()">
<!-- #include file="LogoSection.asp" -->
	<table width="300" border="0" bordercolor="red" cellpadding="0" cellspacing="0" align="left" ID="Table1">
		<%if billtoid="48" then %>
        <form method="post" action="DriverIfabPhoneEmulator_KWE.asp" ID="Form7">
		    <tr><td align="center" colspan="3"><input type="hidden" name="Aliascode" value="<%=AliasCode%>" ID="Hidden1"><input type="submit" value="<<<BACK" ID="gobutton" NAME="Submit1"></td></tr>
		    </form>
            <%else %>
            <form method="post" action="DriverIfabPhoneEmulator.asp" ID="Form2">
		    <tr><td align="center" colspan="3"><input type="hidden" name="Aliascode" value="<%=AliasCode%>" ID="Hidden6"><input type="submit" value="<<<BACK" ID="gobutton" NAME="Submit1"></td></tr>
		</form>
        <%end if %>
		<tr><td align="left">
		<%
		If Submit>"" then
        ColspanNumber="8"
		%>
			<table cellpadding="3" cellspacing="0" width="300" border="0" align="left" ID="Table5">
                        <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>
                        <tr>
		                    <td class="FleetXRedSection" colspan="3" align="center">
			                    <%=uCase(PageStatus)%>
		                    </td>
	                    </tr>
                        <tr><td><img src="images/pixel.gif" height="5" width="1" /></td></tr>
				<form name="Form1" id="Form1" method="post">
				<%If ErrorMessage>"" Then%>
				<tr><td class="ErrorMessageBoldCenter" colspan="2"><%=ErrorMessage%></td></tr>
				<%End if%>
					<input type="hidden" name="Scanned" value="y" ID="Hidden3">
					<%
					'Response.Write "1284 pagestatus="&pagestatus&"<BR>"
					%>
					<input type="hidden" name="PageStatus" value="<%=PageStatus%>" ID="Hidden4">
					<input type="hidden" name="txtcaller" value="<%=VehicleID%>" ID="Hidden5">
					<input type="hidden" name="txtstation" value="<%=FromLocation%>" ID="Hidden7">
					<input type="hidden" name="txtjobnumber" value="<%=jobnumber%>" ID="Hidden14">
					<input type="hidden" name="VehicleID" value="<%=VehicleID%>" ID="Hidden25">
					<input type="hidden" name="jobnumber" value="<%=jobnumber%>" ID="Hidden27">	
					<input type="hidden" name="AliasCode" value="<%=AliasCode%>" ID="Hidden15">				
		<%
			Dim colorset, i, numcolors
			'/--- This is your array of colors to use. -------------\ 
			colorset = split("#D2D2C0,White",",")
			numcolors = ubound(colorset)+1
			Server.ScriptTimeout = 1000
			optJobSel=Request.Querystring("optJobSel")
			optJobSel=Replace(optJobSel,"""","")
			optJobSel=Replace(optJobSel,"'","")
			If ReferenceNumber>"" then optJobSel="ByRef" end if
			If JobNumber>"" then optJobSel="ByJob" end if
					l_csql = "SELECT fcfgthd.fh_id, fcfgthd.fh_user5, fclegs.fl_sf_rta, fcfgthd.fh_status, fcfgthd.fh_ship_dt, fclegs.fl_pkey, fclegs.fl_sf_id, fclegs.fl_sf_name, fclegs.fl_st_id, fclegs.fl_st_name, fclegs.fl_t_atp, fclegs.fl_t_atd, fclegs.fl_pod, fclegs.fl_sf_comment, fcfgthd.fh_custpo, fcfgthd.fh_priority, fcrefs.rf_ref FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN fcrefs ON fclegs.fl_fh_id = fcrefs.rf_fh_id "	
				'l_csql = l_csql&" WHERE (ref_status<>'c') AND (fl_un_id='"&VehicleID&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND fh_ship_dt>'"&now()-30&"'"
				l_csql = l_csql&" WHERE (fl_un_id='"&VehicleID&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND fh_ship_dt>'"&now()-30&"'"
						If PageStatus="ONB" then
							If BillToID="48" or trim(vehicleID)="198" then
								l_csql = L_csql& "AND ((fh_status='PUO') OR (fh_status='AC2')) "
								l_csql = L_csql& "AND ((ref_status='p') or (ref_status='a')) "
								Else
								l_csql = L_csql& "AND ((fh_status='ACC') "
                if LocationCode = "SRHUB" then
                  l_csql = l_csql & " OR (fh_status='AC2') "
                end if
								If VehicleID=124 or VehicleID=123 or VehicleID=113 or VehicleID=125 then
									l_csql = L_csql& " OR ((fh_status='AC2') AND (fl_secacc is not null))) "
									else
									l_csql = L_csql& " ) "
								End if
								l_csql = L_csql& "AND (((ref_status<>'c') or (ref_status is NULL)) "
								If VehicleID=124 or VehicleID=123 or VehicleID=113 then
									l_csql = L_csql& " OR (ref_status='H')) "
									else
									l_csql = L_csql& " ) "
								End if								
							End if
                            'Response.write "VeehicleID="&VehicleID&"<BR>"
                            'Response.write "LocationCode="&LocationCode&"<BR>"
							If VehicleID<>124 and vehicleID<>123 and LocationCode <> "SRHUB" then
                            'Response.write "GOT HERE!!!!<BR>"
								l_csql = L_csql& "AND ((fl_sf_id='"&LocationCode&"') "
							End if
							''''''''THIS MIGHT NEED TO GO AWAY
							If HUB="yXXX" then
								l_csql = l_csql&" OR (Fl_sf_ID='D6W3')"
								l_csql = l_csql&" OR (Fl_sf_ID='D6N2')"
								l_csql = l_csql&" OR (Fl_sf_ID='D6N1')"
								l_csql = l_csql&" OR (Fl_sf_ID='DM4M')"
								l_csql = l_csql&" OR (Fl_sf_ID='DM5M')"
								l_csql = l_csql&" OR (Fl_sf_ID='DPI2')"
								l_csql = l_csql&" OR (Fl_sf_ID='DPI3')"
								l_csql = l_csql&" OR (Fl_sf_ID='ESTK')"
							End if
						
							If VehicleID<>124 and VehicleID<>123 and LocationCode <> "SRHUB" then	
								l_csql = l_csql& " ) "						
							End if
						end if
						If PageStatus="CLS" then
							If trim(vehicleID)="199" then
									l_csql = L_csql& "AND ((((fh_status='PUO'))"
									l_csql = L_csql& "AND ((ref_status='p')))OR ((fh_status='ONB') AND  (fl_rt_type='out'))) "							
								else
								If VehicleID=124 or VehicleID=123 then
									l_csql = L_csql& "AND ((fh_status='DPV') OR (fh_status='ONB'))"
									l_csql = L_csql& "AND ((ref_status='S') or (ref_status='o')) "
									else
									l_csql = L_csql& "AND (((fh_status='ONB') "
									l_csql = L_csql& "AND (ref_status='o')) OR fh_status='DPV') "
								End if
							End if
							
							If HUB2<>"y" AND LocationCode <>"SRHUB" AND LocationCode<>"D6N1B" AND LocationCode<>"DOCK7" then
							  l_csql = L_csql& "AND ("
								l_cSQL = l_cSQL&"(Fl_st_ID='"&LocationCode&"')"
							End if	
              If LocationCode = "D6N1B" then
                l_cSQL = l_cSQL & " AND (Fl_st_id='D6W3' or Fl_st_id='D6N2') "
              End If
               If LocationCode = "DOCK7" then
                l_cSQL = l_cSQL & " AND (Fl_st_id='DM5M' or Fl_st_id='DM5Q' or fl_st_id='DM5S3') "
              End If             						
							If HUB="y" then
								l_csql = l_csql&" OR (Fl_st_ID='D6W3')"
								l_csql = l_csql&" OR (Fl_st_ID='D6N2')"
								l_csql = l_csql&" OR (Fl_st_ID='D6N1')"
								l_csql = l_csql&" OR (Fl_st_ID='DM4M')"
								l_csql = l_csql&" OR (Fl_st_ID='DM5M')"
								l_csql = l_csql&" OR (Fl_st_ID='DPI2')"
								l_csql = l_csql&" OR (Fl_st_ID='DPI3')"
								l_csql = l_csql&" OR (Fl_st_ID='ESTK')"
								l_csql = l_csql&" OR (Fl_st_ID='DM5Q')"
								l_csql = l_csql&" OR (Fl_st_ID='DM6Q')"
								l_csql = l_csql&" OR (Fl_st_ID='RFAB-R')"
								'l_csql = l_csql&" OR (Fl_st_ID='TISHERMA')"
							End if
							If HUB2="y" then
								l_cSQL = l_cSQL&" (Fl_st_ID='"&LocationCode&"')"
							End if
							'''''''''''THIS WILL NEED TO GO AWAY WHEN DUAL LEGS EXIST!								
							If HUB2="yxxx" then
								l_csql = l_csql&" OR (Fl_st_ID='12201')"
								l_csql = l_csql&" OR (Fl_st_ID='12203')"
								l_csql = l_csql&" OR (Fl_st_ID='6430')"
								l_csql = l_csql&" OR (Fl_st_ID='6412')"
								l_csql = l_csql&" OR (Fl_st_ID='13601')"
								l_csql = l_csql&" OR (Fl_st_ID='12500')"
								l_csql = l_csql&" OR (Fl_st_ID='13020')"
								l_csql = l_csql&" OR (Fl_st_ID='7800')"
								l_csql = l_csql&" OR (Fl_st_ID='7839')"
								l_csql = l_csql&" OR (Fl_st_ID='13353')"
								l_csql = l_csql&" OR (Fl_st_ID='13536')"
								l_csql = l_csql&" OR (Fl_st_ID='13121')"
								l_csql = l_csql&" OR (Fl_st_ID='6550')"
								l_csql = l_csql&" OR (Fl_st_ID='13011')"
								l_csql = l_csql&" OR (Fl_st_ID='13570')"
							End if
							if LocationCode="SBRT" then
								l_csql = l_csql&" OR (Fl_st_ID='TISHERMA')"
							end if
							'response.Write "XXXXXLocationCode="&LocationCode&"****<BR>"
							If trim(LocationCode)="D1" then
								l_csql = l_csql&" OR (Fl_st_ID='D7')"
								l_csql = l_csql&" OR (Fl_st_ID='P1')"
							End if																								
							If HUB2<>"y" AND LocationCode <>"SRHUB" AND LocationCode<>"D6N1B" AND LocationCode<>"DOCK7" then
                l_csql = l_csql& " ) "
              End if								
						End if
							SortBy="fh_priority, fh_id, rf_ref"
						If SortBy>"" then
							l_csql = L_csql& " ORDER BY "&Sortby
						End if
						'response.write "VVVL_cSQL="&L_cSQL&"<BR>"
						'response.write "hub="&hub&"<BR>"
						'response.write "hub1="&hub1&"<BR>"
						'response.write "hub2="&hub2&"<BR>"
						'response.write "hub3="&hub3&"<BR>"
				'response.write "1418 L_cSQL=" & L_cSQL & "<br>"
        Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				RSEVENTS2.CursorLocation = 3
				RSEVENTS2.CursorType = 3
				RSEVENTS2.ActiveConnection = DATABASE					
				RSEVENTS2.Open L_cSQL, DATABASE, 1, 3
				If not RSEVENTS2.EOF then
						ELSE
						If BillToID="48" then
						    Response.Redirect("DriverIfabPhoneEmulator_KWE.asp?AliasCode="&AliasCode&"&FakeSubmit=fakesubmit")
						    else
						    Response.Redirect("DriverIfabPhoneEmulator.asp?AliasCode="&AliasCode&"&FakeSubmit=fakesubmit")
						End if
						'response.write "1434 Should have re-directed<BR>"
						ErrorMessage="No jobs were found that match your criteria."	
				End if				
				RSEVENTS2.PageSize = 8
				RSEVENTS2.CacheSize = RSEVENTS2.PageSize
				intPageCount2 = RSEVENTS2.PageCount
				intRecordCount2 = RSEVENTS2.RecordCount
				If (RSEVENTS2.EOF) then
					Sendback2="y"
				End if
				If NOT (RSEVENTS2.BOF AND RSEVENTS2.EOF) Then

				If CInt(intPage2) > CInt(intPageCount2) Then intPage2 = intPageCount2
					If CInt(intPage2) <= 0 Then intPage2 = 1
						If intRecordCount2 > 0 Then
							RSEVENTS2.AbsolutePage = intPage2
							intStart = RSEVENTS2.AbsolutePosition
							If CInt(intPage2) = CInt(intPageCount2) Then
								intFinish = intRecordCount
							Else
								intFinish = intStart + (RSEVENTS2.PageSize - 1)
							End if
						End If
					If intRecordCount2 > 0 Then
						For intRecord2 = 1 to RSEVENTS2.PageSize	
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''				
					fh_id=RSEVENTS2("fh_id")
					MaterialType=RSEVENTS2("fh_user5")
					fl_pkey=RSEVENTS2("fl_pkey")
					'Response.Write "fl_pkey="&fl_pkey&"<BR>"
					fl_sf_rta=RSEVENTS2("fl_sf_rta")
					fh_status=RSEVENTS2("fh_status")
					fh_ship_dt=RSEVENTS2("fh_ship_dt")
					fl_sf_id=RSEVENTS2("fl_sf_id")
					fl_st_id=RSEVENTS2("fl_st_id")					
					fl_sf_name=RSEVENTS2("fl_sf_name")
					fl_st_name=RSEVENTS2("fl_st_name")
					fl_t_atp=RSEVENTS2("fl_t_atp")
					fl_t_atd=RSEVENTS2("fl_t_atd")
					fl_pod=RSEVENTS2("fl_pod")
					fh_custpo=RSEVENTS2("fh_custpo")
					fh_priority=RSEVENTS2("fh_priority")
                    fl_sf_comment=RSEVENTS2("fl_sf_comment")
					fl_sf_rta=RSEVENTS2("fl_sf_rta")
						rf_ref=RSEVENTS2("rf_ref")
			Select Case fh_status
				Case "CLS"
					Display_fh_status="Closed"
				Case "OPN"
					Display_fh_status="Open"
				Case "ACC"
					Display_fh_status="Accepted"
				Case "ONB"
					Display_fh_status="On Board"
				Case "ATD"
					Display_fh_status="At Destination"
				Case "CAN"
					Display_fh_status="Cancelled"
				Case "DEL"
					Display_fh_status="Deleted"	
				Case Else
					Display_fh_status=fh_status																			
			End Select
			if fh_ship_dt="1/1/1900" then fh_ship_dt="&nbsp;" end if
			if fl_t_atp="1/1/1900" then fl_t_atp="&nbsp;" end if
			if fl_t_atd="1/1/1900" then fl_t_atd="&nbsp;" end if
					X=X+1

					
					%>
				<tr>
					<td class="mainpagetextboldcenter" nowrap valign="top"><input type="text" name="FormBarCode<%=X%>" ID="Text2" size="3"></td>	
					<td class="MainPageTextSmaller" nowrap valign="top" bgcolor="<%=colorset(i mod numcolors)%>">
						<!--<b><%=fh_id%></b><br>--><%=X%>)&nbsp;<%=rf_ref%>


<%
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
                                            If trim(NumberOfPieces)>"" then
                                                Response.write "<br><b>"&numberofpieces&" "&rf_box&"</b>"
                                            End if
											Else
											ErrorMessage="Incorrect driver ID or password"
										End if
										Recordset1.Close()
										Set Recordset1 = Nothing  
%>


                        <%if trim(fl_sf_comment)>"" then response.write "<br>***"&fl_sf_comment end if %>
                        </td>
          <td valign="top"><table cellpadding=0 cellpspacing=0><tr><td class="FleetXRedSection" align="center"><!--&nbsp;<a class="FleetXRedSection" href="JobException.asp?j=<%=fh_id%>&s=<%=PageStatus%>&l=<%=LocationCode%>" tabindex="-1">E</a>&nbsp;--></td></tr></table>									
				</td></tr>
<%
				i=i+1
							''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				Response.Write "</font>"
				RSEVENTS2.MoveNext
				If colorchanger = 1 Then
					colorchanger = 0
					color1 = "class=headerwhite"
					color2 = "class=header"
				Else
					colorchanger = 1
					color1 = "class=header"
					color2 = "class=headerwhite"	
				End If
				If RSEVENTS2.EOF Then Exit for
					Next
					End if
					End if				
				''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
						'response.write "XXXBillToID="&BillToID&"<BR>"	
						'response.write "XXXPageStatus="&PageStatus&"<BR>"	
						'response.write "XXXNeedPOD="&NeedPOD&"<BR>"
				''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''		
						'If PageStatus="CLS" and ((BillToID="48" and NeedPOD="y") or BillToID="75" or BillToID="36" or BillToID="80") the
                        If PageStatus="CLS" and (NeedPOD="y") then
						'Response.Write "XXXfh_status="&fh_status&"<BR>"
						%>
						
							<tr>
								<td colspan="2" align="center" class="mainpagetextboldcenter">
									<table cellpadding="0" cellspacing="0" border="0" ID="Table2">
									<tr>
									<td class="mainpagetextbold">
                                    <%
                                    'pod stuff below!
                                     %>
									POD:
									<select name="TempPODID" ID="Select1">	
									<option value="xxx">Select a Signature</option>							
										<%
											''''''''''''''''''''''''''''''''''''''''''''''''''''''
											Set Recordset1 = Server.CreateObject("ADODB.Recordset")
											Recordset1.ActiveConnection = DATABASE
											Recordset1.Source = "SELECT Signature, PODID, st_ID FROM PODList where (PODStatus='c') AND  (st_id='"& locationcode &"') ORDER BY SIGNATURE"
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
													<option value="<%=PODID%>" <%if PODID=TempPODID then response.Write " selected" end if%>><%=Signature%></option>
												<%	
											Recordset1.Movenext
											LOOP
											Recordset1.Close()
											Set Recordset1 = Nothing					
											''''''''''''''''''''''''''''''''''''''''''''''''''''''
											%>
								</select><br> or
								<input type="text" name="addedPOD" maxlength="50" size="20" ID="Text1">
								</td>
								</tr>
								</table>
								</td>
							</tr>							
							<%end if%>
			<input type="hidden" name="LocationCode" value="<%=LocationCode%>" ID="Hidden8">
			<input type="hidden" name="BillToID" value="<%=BillToID%>" ID="Hidden9">
			<input type="hidden" name="fh_status" value="<%=fh_status%>" ID="Hidden2">
			<tr>
				<td colspan="2">
					<%
					'response.Write "billtoid="&billtoid&"<BR>"
					if Pagestatus="CLS" and (NeedPOD="y" ) then%>
					<input type="submit" name="submit" value="submit" ID="gobutton" onclick="return validate()" />
					<%else%>
					<input type="submit" name="submit" value="submit" ID="gobutton">
					<%end if%>
				</td>
			</tr>
			</form>	
			</table>	
			<tr><td>&nbsp;</td></tr>
			<tr><td>&nbsp;</td></tr>
			<tr><td>&nbsp;</td></tr>			
	</Table>
<%end if
'Response.Write "BILLTOID="&BILLTOID&"!!!!!!!<BR>"
%>
	</td></tr>
</body>
</html>
