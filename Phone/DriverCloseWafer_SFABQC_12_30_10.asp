<%@ LANGUAGE="VBSCRIPT"%>
<%
Response.buffer = True
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
	<SCRIPT Language="Javascript" SRC="Script/Calendar1-902.js"></SCRIPT> 
	<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<TITLE><% Response.Write(D_TITLEBAR) %></TITLE>
	<!-- added the include style.css-->
<link rel="stylesheet" type="text/css" href="mainStyleSheet.css">
<!-- #include file="../DedicatedFleets/include/checkstring.inc" -->
<!-- #include file="../DedicatedFleets/include/custom.inc" -->
<!-- #include file="../DedicatedFleets/include/ifabsettings.inc" -->
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
WaferList=FormBarCode(1)
for xxx= 2 to 12
If Trim(FormBarCode(xxx))>"" then
    WaferList=WaferList&","&FormBarCode(xxx)
End if
next
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
LMNOP=Request.Form("LMNOP")
fh_status=Request.Form("Fh_status")
AliasCode=Request.Form("AliasCode")
LocationCode=Request.Form("LocationCode")
SecondDriver=Request.Form("SecondDriver")
SecondUserName=Request.Form("SecondUserName")
SecondPassword=Request.Form("SecondPassword")
''''''''''''''''''''''POD Stuff''''''''''''''''''''''''''''''''''''''''''''''''''
PODID=Request.Form("TempPODID")
AddedPOD=Request.Form("AddedPOD")
If AddedPOD>"" then
	AddedPOD=Replace(AddedPOD,",","")
End if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''POD2 Stuff''''''''''''''''''''''''''''''''''''''''''''''''''
PODID2=Request.Form("TempPODID2")
AddedPOD2=Request.Form("AddedPOD2")
If AddedPOD2>"" then
	AddedPOD2=Replace(AddedPOD2,",","")
End if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''Second Driver''''''''''''''''''''''''''''''''''''''''''''''
'Response.Write "UserID="&UserID&"<BR>"
SecondPassword=Replace(SecondPassword,"'","")
SecondPassword=Replace(SecondPassword,"""","")
If SecondDriver="y" then
Set Recordset1 = Server.CreateObject("ADODB.Recordset")
SQL777="SELECT * FROM INTRANET_USERS WHERE (UserID<>"&UserID&") AND (USERNAME='"&SecondUserName&"') AND (PASSWORD='"&SecondPassword&"') AND (Status='c') AND ((Rights='u') OR (Rights='a') OR (Rights='g'))"
'Response.Write "SQL777="&SQL777&"***<BR>"
Recordset1.ActiveConnection = Intranet
Recordset1.Source = SQL777
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()
Recordset1_numRows = 0
	if NOT Recordset1.EOF then
		SecondUserID=Recordset1("UserID")
		SecondFirstName=Recordset1("FirstName")
		SecondLastName=Recordset1("LastName")
		Else
		ErrorMessage="Incorrect second driver ID or password"
		End if
		Recordset1.Close()
		Set Recordset1 = Nothing
End if
'Response.Write "HERE I AM!<BR>"
If SecondDriver="y" and (SecondUserName="" or SecondPassword="") then
	ErrorMessage="Both a second driver ID and password are required<br>"
End if

BarCode=Request.Form("BarCode")
BillToID=Request.Form("BillToID")
'Response.Write "LocationCode="&LocationCode&"<BR>"
If BillToID>"" then
	Suid=BillToID
End if

Submit=Request.Form("Submit")
If Submit="" and Barcode="" then
	Response.Redirect("DriverifabPhoneEmulator_SFABQC.asp")
	'Response.Write "Response.Redirect I'm here #2!<BR>"
End if
PageStatus=Request.Form("PageStatus")
'Response.Write "XXXPageStatus="&PageStatus&"<BR>"
If Submit="" then
	Submit=Request.QueryString("Submit")
End if
'Submit="xx"
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
txtstation=trim(Request.Form("txtstation"))

'''''''''NEW CLOSING PART-6/2/2010'''''''''''''''
IF Submit="submit" THEN
    ''''''''DROPPING ORDERS'''''''''''''''
    If PageStatus="CLS" then
   
			For q=1 to LMNOP
			    
				If trim(FormBarCode(q))>""  then
                    'Response.Write "whatever="&FormBarCode(q)&"<BR>"				
		 		    Set oRs = Server.CreateObject("ADODB.Recordset")
		            oRs.CursorLocation = 3
		            oRs.CursorType = 3
		            oRs.ActiveConnection = DATABASE	
					SQL = "SELECT fcfgthd.fh_id, fclegs.fl_sf_rta, fcfgthd.fh_status, fcfgthd.fh_ship_dt, fcfgthd.fh_user5, fclegs.fl_sf_id, fclegs.fl_sf_name, fclegs.fl_st_id, fclegs.fl_st_name, fclegs.fl_t_atp, fclegs.fl_t_atd, fclegs.fl_pod, fclegs.fl_finaldestination, fcfgthd.fh_custpo, fcfgthd.fh_priority, fcrefs.rf_box FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN fcrefs ON fclegs.fl_fh_id = fcrefs.rf_fh_id "	
					SQL = SQL&" WHERE (Fl_dr_ID='"&VehicleID&"') AND (rf_box='"&FormBarCode(q)&"')"
					SQL = SQL& "AND (fh_id='"&JobNumber&"') "						
					SQL = SQL& "AND ((fh_status='ONB') OR (fh_status='DPV'))  "
					SQL = SQL& "AND (ref_status='o') AND (fl_leg_status='c') "
					oRs.Open SQL, DATABASE, 1, 3
					'REsponse.Write "SQL="&SQL&"<BR>"
					If oRs.EOF then
						ErrorMessage=ErrorMessage&" "&FormBarCode(q)&" is not accepted.<br>Check Paper Work/Call Supervisor<br>"
					End if
					If not oRs.eof and trim(ErrorMessage)="" then
						TheJobNumber = oRs("fh_id")
						JobStatus=oRs("fh_status")
						TempOrigination = trim(oRs("fl_sf_id"))
						TempDestination = trim(oRs("fl_st_id"))
						FinalDestination = trim(oRs("fl_FinalDestination"))
						TheBarCode = FormBarCode(q)
                        MaterialType = oRs("Fh_User5")
     					If ucase(MaterialType)="ITAR" AND (trim(PODID)="xxx") AND trim(addedPOD)="" then
                        	ErrorMessage="A POD 'signature' is required on all ITARs."
    					End if 							
              			If addedPOD>"" and PODID="xxx" and XYZzz=0 then
            				Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
            					RSEVENTS2.Open "PODList", Database, 2, 2
            					RSEVENTS2.addnew	
            					RSEVENTS2("bt_ID")=BillToID		
            					RSEVENTS2("st_ID") = LocationCode
            					RSEVENTS2("Signature")=addedPOD	
            					RSEVENTS2("PODStatus") = "c"
            					RSEVENTS2.update
            					RSEVENTS2.close			
            				set RSEVENTS2 = nothing	
            				Set Recordset166 = Server.CreateObject("ADODB.Recordset")
            				Recordset166.ActiveConnection = Database
            				Recordset166.Source = "SELECT PODID FROM PODList WHERE (bt_ID='"&BillToID&"') AND (st_ID='"&LocationCode&"') AND (Signature='"&AddedPOD&"') AND (PODStatus='c')"
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
            				XYZzz=XYZzz+1
            			End if	                      

						Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open DATABASE
						''''UPDATES THE WAFER
                        If trim(ErrorMessage)="" then
    						l_cSQL = "UPDATE FCREFS SET ref_status = 'c'" 
    						If PODID>"" then
							    l_cSQL = l_cSQL&", PUPOD = '"&PODID&"' "
						    End if
                            l_cSQL = l_cSQL&" WHERE (rf_fh_id = '"&TheJobNumber&"') AND (rf_box = '" &TheBarCode& "')"
    						oConn.Execute(l_cSQL)
                        End if
						oConn.Close
						Set oConn=Nothing                        
	
					End if
 			        oRs.Close
			        Set oRs=Nothing 
				End if

			Next 
		''''''CHECKS TO SEE IF ITS OKAY TO CHANGE THE ORDER STATUS
		Set oRs75 = Server.CreateObject("ADODB.Recordset")
		oRs75.CursorLocation = 3
		oRs75.CursorType = 3
		oRs75.ActiveConnection = DATABASE	
		SQL75 = "SELECT rf_fh_id FROM fcrefs"
		SQL75 = SQL75&" WHERE (rf_fh_id='"&TheJobNumber&"')"
		SQL75 = SQL75&" AND (Ref_Status='o')"
		oRs75.Open SQL75, DATABASE, 1, 3
		'Response.Write "SQL75="&SQL75&"<BR>"
		If oRs75.EOF then
		

			Set oConn = Server.CreateObject("ADODB.Connection")
			oConn.ConnectionTimeout = 100
			oConn.Provider = "MSDASQL"
			oConn.Open DATABASE		
			'Response.Write "trim(Tempdestination)="&trim(Tempdestination)&"***<BR>"
			'Response.Write "Trim(FinalDestination)="&Trim(FinalDestination)&"***<BR>"
		    If trim(Tempdestination)=Trim(FinalDestination) or trim(TempOrigination)="72" then
		        L_SQL_44="PHONE_CHANGE_STATUS '" & TheJobNumber & "', '9', 'CLS', '', '',  '"& UserID &"', '"& UnitID &"'" 
		        'Response.Write "L_SQL_44(1)="& L_SQL_44 &"<BR>"
		        oConn.Execute(L_SQL_44)	
		        
		        
 						Body = "Wafer Cart #"&TheBarCode&" has just been delivered to "& LocationCode &".<br><br>"& _
						"It was job number:  "&TheJobNumber&"<br><br>"& _
						"LogistiCorp" 
						'Recipient = "mark.maggiore@logisticorp.us"
						Set objMail = CreateObject("CDONTS.Newmail")
						objMail.From = "system.monitor@logisticorp.us"
						objMail.To = "mark.maggiore@logisticorp.us"
						Select Case LocationCode
							Case "CSSF"
							objMail.CC = "paulsingh@ti.com;tracy.johnson@ti.com;robles@ti.com;m-haralson@ti.com;xscottsp@ti.com;rb44@ti.com"
							'objMail.CC = "DM5S-TEST@joeblow.com"
							Case "SCTQCxx"
							objMail.CC = "itar@logisticorp.us"
							'objMail.CC = "SCTEST@joeblow.com"
							Case "EBTSHxx"
							objMail.CC = "itar@logisticorp.us"
							'objMail.CC = "DM5S-TEST@joeblow.com"
							Case "SCTA1xx"
							objMail.CC = "itar@logisticorp.us"
							'objMail.CC = "SCTEST@joeblow.com" 
							Case else
							'objMail.CC = "itar@logisticorp.us"
							'objMail.CC = "SCTEST@joeblow.com" 											                                          
						End Select
						objMail.Subject = "SFABQC Delivery"
						objMail.MailFormat = cdoMailFormatMIME
						objMail.BodyFormat = cdoBodyFormatHTML
						objMail.Body = Body
						objMail.Send
						Set objMail = Nothing               		        
		        
				If ucase(MaterialType)="ITAR" then
                    '''''''''''ITAR NOTIFICATION'''''''''''''''''''''''''''''''''''''''''''''''''
						Body = "ITAR #"&TheBarCode&" has just been delivered to "& LocationCode &".<br><br>"& _
						"It was job number:  "&TheJobNumber&"<br><br>"& _
						"LogistiCorp" 
						'Recipient = "mark.maggiore@logisticorp.us"
						Set objMail = CreateObject("CDONTS.Newmail")
						objMail.From = "system.monitor@logisticorp.us"
						objMail.To = "mark.maggiore@logisticorp.us"
						Select Case LocationCode
							Case "EBTW1"
							objMail.CC = "itar@logisticorp.us"
							'objMail.CC = "DM5S-TEST@joeblow.com"
							Case "SCTQC"
							objMail.CC = "itar@logisticorp.us"
							'objMail.CC = "SCTEST@joeblow.com"
							Case "EBTSH"
							objMail.CC = "itar@logisticorp.us"
							'objMail.CC = "DM5S-TEST@joeblow.com"
							Case "SCTA1"
							objMail.CC = "itar@logisticorp.us"
							'objMail.CC = "SCTEST@joeblow.com" 
							Case else
							objMail.CC = "itar@logisticorp.us"
							'objMail.CC = "SCTEST@joeblow.com" 											                                          
						End Select
						objMail.Subject = "ITAR Delivery"
						objMail.MailFormat = cdoMailFormatMIME
						objMail.BodyFormat = cdoBodyFormatHTML
						objMail.Body = Body
						objMail.Send
						Set objMail = Nothing	
                    End if			        
		        
		        
		        
		        
		        				
		        else
		        L_SQL_44="PHONE_CHANGE_STATUS '" & TheJobNumber & "', '53', 'ARV', '', '',  '"& UserID &"', '"& UnitID &"'" 
		        'Response.Write "L_SQL_44(2)="& L_SQL_44 &"<BR>"
		        oConn.Execute(L_SQL_44)
               ' Response.Write "****L_SQL_44="&L_SQL_44&"<BR>"
 				''''UPDATES THE NEXT LEG TO MAKE IT THE CURRENT LEG           
				Set oConn64 = Server.CreateObject("ADODB.Connection")
				oConn64.ConnectionTimeout = 100
				oConn64.Provider = "MSDASQL"
				oConn64.Open DATABASE
				l_cSQL64 = "UPDATE FCREFS SET ref_status = NULL" 
                l_cSQL64 = l_cSQL64&" WHERE (rf_fh_id='"& TheJobNumber &"') AND (ref_status <>'x')"
				'Response.Write "****L_SQL_64="& L_SQL_64 &"<BR>"
				oConn64.Execute(l_cSQL64)
				oConn64.Close
				Set oConn64=Nothing            
            
            
            End if
 		    oConn.Close
		    Set oConn=Nothing           
		End if
		oRs75.Close
		Set oRs75=Nothing							
		''''''''''''''''''''''''''''''''''''''''
		'End if
		TheJobNumber=""
		TheBarCode=""
    End if
    'Response.Write "PageStatus="&PageStatus&"<BR>"
	If PageStatus="ONB" then
	'Response.Write "X"
		For q=1 to LMNOP
		    'Response.Write "Q="&Q&"<BR>"
			'Response.Write "formBarCode="&FormBarCode(q)&"<BR>"
			If trim(FormBarCode(q))>""  then
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''					
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				SQL = "SELECT fcfgthd.fh_id, fclegs.fl_sf_rta, fcfgthd.fh_status, fcfgthd.fh_ship_dt, fcfgthd.fh_user5, fclegs.fl_sf_id, fclegs.fl_sf_name, fclegs.fl_st_id, fclegs.fl_st_name, fclegs.fl_t_atp, fclegs.fl_t_atd, fclegs.fl_pod, fcfgthd.fh_custpo, fcfgthd.fh_priority, fcrefs.rf_box FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN fcrefs ON fclegs.fl_fh_id = fcrefs.rf_fh_id "	
				SQL = SQL&" WHERE (Fl_dr_ID='"&VehicleID&"') AND (fh_ship_dt>'"&now()-30&"') AND (rf_box='"&FormBarCode(q)&"')"
				SQL = SQL& "AND ((fh_status='ACC') "
				SQL = SQL& "AND ((ref_status is NULL) OR (ref_status='c') OR (ref_status='H')))"
				SQL = SQL& "AND (fh_id='"&JobNumber&"') "						
				
				'response.write "<br><font color='blue'>First Select="&SQL&"<BR></font>"
				
				oRs.Open SQL, DATABASE, 1, 3
				If oRs.EOF then
					ErrorMessage=ErrorMessage&" "&FormBarCode(q)&" is not accepted.<br>Check Paper Work/Call Supervisor<br>"
				End if
				If not oRs.eof and trim(ErrorMessage)="" then
                'Response.Write "Got here #1<br>"
					TheJobNumber = oRs("fh_id")
					TempDestination = trim(oRs("fl_st_id"))
					TheBarCode = FormBarCode(q)
                    MaterialType = oRs("Fh_User5")
 					If ucase(MaterialType)="ITAR" AND (trim(PODID)="xxx") AND trim(addedPOD)="" then
                    	ErrorMessage="A POD 'signature' is required on all ITARs."
					End if 							
                    'Response.write "TheJobNumber=***"&TheJobNumber&"***<BR>"
          			If addedPOD>"" and PODID="xxx" and XYZzz=0 then
        				'Response.Write "GOT HERE!!!<BR>"
        				Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
        					RSEVENTS2.Open "PODList", Database, 2, 2
        					RSEVENTS2.addnew	
        					RSEVENTS2("bt_ID")=BillToID		
        					RSEVENTS2("st_ID") = LocationCode
        					'RSEVENTS2("LogInOut") = "o"
        					RSEVENTS2("Signature")=addedPOD	
        					RSEVENTS2("PODStatus") = "c"
        					RSEVENTS2.update
        					RSEVENTS2.close			
        				set RSEVENTS2 = nothing	
        				
        				
        				Set Recordset166 = Server.CreateObject("ADODB.Recordset")
        				'Response.Write "Intranet="&Intranet&"***<BR>"
        				Recordset166.ActiveConnection = Database
        				Recordset166.Source = "SELECT PODID FROM PODList WHERE (bt_ID='"&BillToID&"') AND (st_ID='"&LocationCode&"') AND (Signature='"&AddedPOD&"') AND (PODStatus='c')"
        				'Response.Write "Recordset166.Source="&Recordset166.Source&"<BR>"
        				Recordset166.CursorType = 0
        				Recordset166.CursorLocation = 2
        				Recordset166.LockType = 1
        				Recordset166.Open()
        				Recordset166_numRows = 0
        					if NOT Recordset166.EOF then
        						PODID=Recordset166("PODID")
        						'Response.Redirect("DriverMessage.asp")
        						Else
        						ErrorMessage="No such signer exists"
        					End if
        					Recordset166.Close()
        					Set Recordset166 = Nothing				
        				'Response.Write "PODID="&PODID&"<BR>"
        				XYZzz=XYZzz+1
        			End if	                      
					Set oConn = Server.CreateObject("ADODB.Connection")
					oConn.ConnectionTimeout = 100
					oConn.Provider = "MSDASQL"
					oConn.Open DATABASE
					''''UPDATES THE WAFER
                    If trim(ErrorMessage)="" then
                    'Response.Write "Got here #2<br>"
						l_cSQL = "UPDATE FCREFS SET ref_status = 'o'" 
						If PODID>"" then
						    l_cSQL = l_cSQL&", PUPOD = '"&PODID&"' "
					    End if
                        l_cSQL = l_cSQL&" WHERE (rf_fh_id = '"&TheJobNumber&"') AND (rf_box = '" &TheBarCode& "')"
                        'response.write "<font color='green'>UPDATE Wafers="&l_cSQL&"<BR></font>"
						oConn.Execute(l_cSQL)
                    End if
					''''''''''''''''''''''''''''''''''''''''
					Set oRs = Server.CreateObject("ADODB.Recordset")
					oRs.CursorLocation = 3
					oRs.CursorType = 3
					oRs.ActiveConnection = DATABASE	
					''''''CHECKS TO SEE IF ITS OKAY TO CHANGE THE ORDER STATUS
					SQL = "SELECT rf_fh_id FROM fcrefs"
					SQL = SQL&" WHERE (rf_fh_id='"&TheJobNumber&"')"
					'''''SQL = SQL&" AND ((Ref_Status IS NULL) or (Ref_Status='Q'))"
					SQL = SQL&" AND ((Ref_Status IS NULL) or ((Ref_Status<>'X') AND (Ref_Status<>'o')))"
				
					oRs.Open SQL, DATABASE, 1, 3
			
					If oRs.EOF then
						if trim(Hub)<>"y" and trim(Hub3)<>"y" then

							L_SQL_44144="PHONE_CHANGE_STATUS '" & TheJobNumber & "', '5', 'ONB', '', '', '"& UserID &"', '"& UnitID &"'" 
							oConn.Execute L_SQL_44144
							If trim(MaterialType)="ITAR" then
                                '''''''''''ITAR NOTIFICATION'''''''''''''''''''''''''''''''''''''''''''''''''
									Body = "A LogistiCorp driver has just On Boarded ITAR #"&TheBarCode&".  Please be at "& tempDestination &", prepared to accept the handoff of this ITAR within the next 5 minutes.<br><br>"& _
									"It is job number:  "&TheJobNumber&"<br><br>"& _
									"LogistiCorp" 
									'Recipient = "mark.maggiore@logisticorp.us"
									Set objMail = CreateObject("CDONTS.Newmail")
									objMail.From = "system.monitor@logisticorp.us"
									objMail.To = "mark.maggiore@logisticorp.us"
									Select Case LocationCode
										Case "EBTW1"
										objMail.CC = "itar@logisticorp.us"
										'objMail.CC = "DM5S-TEST@joeblow.com"
										Case "SCTQC"
										objMail.CC = "itar@logisticorp.us"
										'objMail.CC = "SCTEST@joeblow.com"
										Case "EBTSH"
										objMail.CC = "itar@logisticorp.us"
										'objMail.CC = "DM5S-TEST@joeblow.com"
										Case "SCTA1"
										objMail.CC = "itar@logisticorp.us"
										'objMail.CC = "SCTEST@joeblow.com"                                           
									End Select
									objMail.Subject = "ITAR on the way"
									objMail.MailFormat = cdoMailFormatMIME
									objMail.BodyFormat = cdoBodyFormatHTML
									objMail.Body = Body
									objMail.Send
									Set objMail = Nothing	
                                End if								
								''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''                                
							'response.write "L_SQL_44144="&L_SQL_44144&"<BR>"
							If trim(SecondUserID)>"" then
								l_cSQL = "UPDATE FCREFS SET PUDriver2 = '"&SecondUserID&"' WHERE (rf_fh_id = '"&TheJobNumber&"') AND (rf_box = '" &TheBarCode& "')"
								'response.write "<font color='green'>UPDATE Wafers="&l_cSQL&"<BR></font>"
								oConn.Execute(l_cSQL)	
							End if							
							else
							L_SQL_44="PHONE_CHANGE_STATUS '" & TheJobNumber & "', '54', 'DPV', '', '',  '"& UserID &"', '"& UnitID &"'" 
							oConn.Execute L_SQL_44
							
							'response.write "L_SQL_44="&L_SQL_44&"<BR>"
						
					
						
						End if							
						
						''''''''''''''''''''''''''''''''''''''''
					End if
					oRs.Close
					Set oRs=Nothing							
					''''''''''''''''''''''''''''''''''''''''
					oConn.Close	
					Set oConn=Nothing	
					'End if
					TheJobNumber=""
					TheBarCode=""	
				End if
				''''cookie?("TempJobNumber")=JobNumber
			End if
		Next
	End if    
End if

%>
<body leftMargin=0 TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" OnLoad=document.Form1.FormBarCode1.focus()>
	<TABLE WIDTH="300" border="0" bordercolor="red" cellpadding="0" cellspacing="5" align="left" ID="Table1">
		<tr><td align="center" colspan="3"><form method="post" action="DriverIfabPhoneEmulator_SFABQC.asp" ID="Form7"><input type="hidden" name="Aliascode" value="<%=AliasCode%>" ID="Hidden1"><input type="submit" value="Return to Drop Off/Pick Up" ID="Submit1" NAME="Submit1"></form></td></tr>
		<tr><td align="left">
		<%
		If Submit>"" then
		%>
			<table cellpadding="3" cellspacing="0" width="300" border="1" align="left" ID="Table5">
				<tr>
					<%
						ColspanNumber="8"
					%>
					<td  class="mainpagetextboldcenter" nowrap colspan="2">
						<%=PageStatus%>
					</td>				
																				
				</tr>

				<%If ErrorMessage>"" Then%>
				<tr><td class="ErrorMessageBoldCenter" colspan="2"><%=ErrorMessage%></td></tr>
				<%End if%>
				<form name="Form1" id="Form1" method="post">
					<input type="hidden" name="Scanned" value="y" ID="Hidden3">
					<input type="hidden" name="PageStatus" value="<%=PageStatus%>" ID="Hidden4">
					<input type="hidden" name="txtcaller" value="<%=VehicleID%>" ID="Hidden5">
					<input type="hidden" name="txtstation" value="<%=FromLocation%>" ID="Hidden7">
					<input type="hidden" name="txtjobnumber" value="<%=jobnumber%>" ID="Hidden14">
					<input type="hidden" name="VehicleID" value="<%=VehicleID%>" ID="Hidden25">
					<!--input type="hidden" name="LocationCode" value="<%=FromLocation%>" ID="Hidden26"-->
					<input type="hidden" name="jobnumber" value="<%=jobnumber%>" ID="Hidden27">	
					<input type="hidden" name="AliasCode" value="<%=AliasCode%>" ID="Hidden15">				
			
		<%
			Dim colorset, i, numcolors
			'/--- This is your array of colors to use. -------------\ 
			colorset = split("#D2D9FC,White",",")
			numcolors = ubound(colorset)+1

		
		    ''''''''''''RESETS THE ARRAY'''''''''''''
		    LMNOP=0
		    '''''''''''''''''''''''''''''''''''''''''
			Server.ScriptTimeout = 1000
			optJobSel=Request.Querystring("optJobSel")
			optJobSel=Replace(optJobSel,"""","")
			optJobSel=Replace(optJobSel,"'","")
			If ReferenceNumber>"" then optJobSel="ByRef" end if
			If JobNumber>"" then optJobSel="ByJob" end if
			''Response.write "******optJobSel="&optJobSel&"<BR>"
			Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				RSEVENTS2.CursorLocation = 3
				RSEVENTS2.CursorType = 3
				RSEVENTS2.ActiveConnection = Database
					l_csql = "SELECT distinct(fcfgthd.fh_id), fclegs.fl_sf_rta, fcfgthd.fh_status, fcfgthd.fh_ship_dt, fcfgthd.fh_user5, fclegs.fl_sf_id, fclegs.fl_sf_name, fclegs.fl_st_id, fclegs.fl_st_name, fclegs.fl_t_atp, fclegs.fl_t_atd, fclegs.fl_pod, fcfgthd.fh_custpo, fcfgthd.fh_priority, fcrefs.rf_box FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN fcrefs ON fclegs.fl_fh_id = fcrefs.rf_fh_id "	
				l_csql = l_csql&" WHERE (rf_box>'') AND (Fl_dr_ID='"&VehicleID&"') AND fh_ship_dt>'"&now()-30&"'"
				l_csql = L_csql& " AND (fh_id='"&JobNumber&"') AND (fl_leg_status='c') "
						If PageStatus="ONB" then
							l_csql = L_csql& "AND ((fh_status='ACC') AND ((ref_status is NULL) OR (ref_status <> 'o')))"
							'l_csql = L_csql& " OR ((fh_status='AC2') AND ((ref_status='Q') OR (ref_status='c') or (ref_status='H')))) "
						end if
						If PageStatus="CLS" then
							l_csql = L_csql& "AND ((fh_status='ONB') or (fh_status='DPV'))"
								if trim(hub)="y" or trim(hub2)="y" or trim(hub3)="y" then
									l_csql = L_csql& "AND ((ref_status='o') "
								else
									l_csql = L_csql& "AND ((ref_status='H') "
								End if
								l_csql = L_csql& " or (ref_status='o')) "							
						End if
							SortBy="rf_box, fh_priority, fh_id"
						'End if
						If SortBy>"" then
							l_csql = L_csql& " ORDER BY "&Sortby
						End if
					
					'response.write("Query3XXX:" & l_cSQL&"<BR>")
					
			''''''''''''''''''''''''''''''''''''''''
				RSEVENTS2.Open l_cSQL, Database, 1, 3
				If RSEVENTS2.eof then
						''''''''''''''''
						'''''commented out
						Response.Redirect("DriverIfabPhoneEmulator_SFABQC.asp?AliasCode="&AliasCode&"&FakeSubmit=fakesubmit")
						''''''''''''''''''''''
						'Response.write "Response.redirect...got here #8!<BR>"
						ErrorMessage="No jobs were found that match your criteria."	
				End if
				'If not RSEVENTS2.EOF THEN				
				Do while not RSEVENTS2.EOF 
				    ''''''''''''variable for number of lots''''''''''''''
				    LMNOP=LMNOP+1
				    
					fh_id=RSEVENTS2("fh_id")
					fl_sf_rta=RSEVENTS2("fl_sf_rta")
					fh_status=RSEVENTS2("fh_status")
					'response.write "fh_status="&fh_status&"<BR>"
					fh_ship_dt=RSEVENTS2("fh_ship_dt")
					fh_User5=RSEVENTS2("fh_User5")
					fl_sf_id=RSEVENTS2("fl_sf_id")
					fl_st_id=RSEVENTS2("fl_st_id")					
					fl_sf_name=RSEVENTS2("fl_sf_name")
					fl_st_name=RSEVENTS2("fl_st_name")
					fl_t_atp=RSEVENTS2("fl_t_atp")
					fl_t_atd=RSEVENTS2("fl_t_atd")
					fl_pod=RSEVENTS2("fl_pod")
					fh_custpo=RSEVENTS2("fh_custpo")
					fh_priority=RSEVENTS2("fh_priority")
					fl_sf_rta=RSEVENTS2("fl_sf_rta")
					'If USESLOTS=TRUE then
						rf_box=RSEVENTS2("rf_box")
					'End if
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
			'If ErrorMessage="" then
					If trim(rf_box)>"" then
					X=X+1
					
					%>
				<tr>
					<td class="mainpagetextboldcenter" nowrap><input type="text" name="FormBarCode<%=X%>" ID="Text2" size="3"></td>	
					<td class="MainPageTextSmaller" nowrap valign="top" bgcolor="<%=colorset(i mod numcolors)%>">
						<b><%=fh_id%></b><br><%=X%>)&nbsp;<%=rf_box%> <b>(CART ID)</b>
					</td>				
				</tr>
<%
				i=i+1
				END IF
				RSEVENTS2.movenext
				LOOP
				RSEVENTS2.close
			Set RSEVENTS2 = Nothing
			
			'Response.Write "LMNOP="&LMNOP&"<BR>"

	%>
	        <input type="hidden" name="LMNOP" value="<%=LMNOP%>" ID="Hidden6">
			<input type="hidden" name="fh_status" value="<%=fh_status%>" ID="Hidden2">
			<input type="hidden" name="LocationCode" value="<%=LocationCode%>" ID="Hidden8">
			<input type="hidden" name="BillToID" value="<%=BillToID%>" ID="Hidden9">
				<%
               'Response.Write "fh_user5="& fh_user5 &"****<BR>"
				If lcase(fh_user5)="secure waf" or ucase(fh_user5)="ITAR" then
					If ucase(fh_user5)<>"ITAR" then
                    %>
					<input type="hidden" name="SecondDriver" value="y">
					<tr>
						<td colspan="2">
							<table width="100%">
								<tr>
									<td class="mainpagetextboldcenter" nowrap>
										SECOND DRIVER<br>(Other than <%=FirstName%>&nbsp;&nbsp;<%=LastName%>)
									</td>
								</tr>							
								<tr>
									<td>
										&nbsp;&nbsp;&nbsp;Username:&nbsp;&nbsp;&nbsp;<input type="text" name="SecondUserName" value="" size="15">
									</td>
								</tr>
								<tr>
									<td>
										&nbsp;&nbsp;&nbsp;Password:&nbsp;&nbsp;&nbsp;<input type="password" name="SecondPassword" value="" size="15">
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<%
                    End if
					If PageStatus="CLS" or ucase(fh_user5)="ITAR" then
					%>
							<tr>
								<td colspan="2" align="center" class="mainpagetextboldcenter">
									<table cellpadding="0" cellspacing="0" border="0" ID="Table2">
									<tr>
									<td class="mainpagetextbold">
									POD #1:
									<select name="TempPODID" ID="Select1">	
									<option value="xxx">Select a Signature</option>	
                                    <%If ucase(fh_user5)<>"ITAR" then %>
									    <option value="168">**Secure Cabinet**</option>	
                                    <%end if %>						
										<%
											''''''''''''''''''''''''''''''''''''''''''''''''''''''
											Set Recordset1 = Server.CreateObject("ADODB.Recordset")
											Recordset1.ActiveConnection = DATABASE
											Recordset1.Source = "SELECT PODID, Signature FROM fcshipto INNER JOIN PODList ON fcshipto.st_id = PODList.st_ID COLLATE SQL_Latin1_General_CP1_CI_AS where (PODStatus='c') AND (bt_id='"&BillToID&"') AND (fcshipto.st_Alias='"&AliasCode&"') ORDER BY SIGNATURE"
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
                            <%if Ucase(fh_user5)<>"ITAR" then %>	
							<tr>
								<td colspan="2" align="center" class="mainpagetextboldcenter">
									<table cellpadding="0" cellspacing="0" border="0" ID="Table3">
									<tr>
									<td class="mainpagetextbold">
									POD #2:
									<select name="TempPODID2" ID="Select2">	
									<option value="xxx">Select a Signature</option>							
										<%
											''''''''''''''''''''''''''''''''''''''''''''''''''''''
											Set Recordset1 = Server.CreateObject("ADODB.Recordset")
											Recordset1.ActiveConnection = DATABASE
											Recordset1.Source = "SELECT PODID, Signature FROM fcshipto INNER JOIN PODList ON fcshipto.st_id = PODList.st_ID COLLATE SQL_Latin1_General_CP1_CI_AS where (PODStatus='c') AND (bt_id='"&BillToID&"') AND (fcshipto.st_Alias='"&AliasCode&"') ORDER BY SIGNATURE"
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
												PODID2=Recordset1("PODID")
												Signature2=Recordset1("Signature")
												%>
													<option value="<%=PODID2%>" <%if PODID2=TempPODID2 then response.Write " selected" end if%>><%=Signature2%></option>
												<%	
											Recordset1.Movenext
											LOOP
											Recordset1.Close()
											Set Recordset1 = Nothing					
											''''''''''''''''''''''''''''''''''''''''''''''''''''''
											%>
								</select><br> or
								<input type="text" name="addedPOD2" maxlength="50" size="20" ID="Text3">
								</td>
								</tr>
								</table>
								</td>
							</tr>																					
					<%
                            End if
						End if
					'Response.Write "hello?<BR>"
				End if
				%>			
			<tr>
				<td colspan="2">
					<input type="submit" name="submit" value="submit" ID="Submit2">
				</td>
			</tr>
			</form>	
			</table>	
			<tr><td>&nbsp;</td></tr>
			<tr><td>&nbsp;</td></tr>
			<tr><td>&nbsp;</td></tr>			

	</Table>
<%end if%>
	</td></tr>
</body>
</html>
