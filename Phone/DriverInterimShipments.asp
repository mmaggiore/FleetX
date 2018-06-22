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
PODID=Request.Form("TempPODID")
AddedPOD=Request.Form("AddedPOD")
If AddedPOD>"" then
	AddedPOD=Replace(AddedPOD,",","")
End if
		If UCASE(LocationCode)="EBHUB" or UCASE(LocationCode)="13536" then
			If UCASE(LocationCode)="EBHUB" then
				BillToID="26"
				LocationCode="EBHUB"
				Hub="y"
				Else
				BillToID="48"
				LocationCode=UCASE(LocationCode)
				'''''''''THIS MAY NEED TO GO AWAY!
				'''''Hub2="y"
			End if
		End if	
BarCode=Request.Form("BarCode")
BillToID=Request.Form("BillToID")
If BillToID>"" then
	Suid=BillToID
End if
													If VehicleID=666 and BarCode>"" then
														Barcode=Right(Barcode,10)
														Barcode=Left(Barcode,9)
														Barcode=0&Barcode
													End if
Submit=Request.Form("Submit")
If Submit="" and Barcode="" then
	response.Redirect("Default.asp")
	'Response.Write "Got here<br>"
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
IF Submit="submit" THEN
		If PageStatus="CLS" then
			For q=1 to 12
				If trim(FormBarCode(q))>""  then
					Set oRs = Server.CreateObject("ADODB.Recordset")
					oRs.CursorLocation = 3
					oRs.CursorType = 3
					oRs.ActiveConnection = DATABASE	
					SQL = "SELECT fcfgthd.fh_id, fclegs.fl_pkey, fclegs.fl_sf_rta, fcfgthd.fh_status, fcfgthd.fh_ship_dt, fclegs.fl_sf_id, fclegs.fl_sf_name, fclegs.fl_st_id, fclegs.fl_st_name, fclegs.fl_t_atp, fclegs.fl_t_atd, fclegs.fl_pod, fcfgthd.fh_custpo, fcfgthd.fh_priority, fcrefs.rf_ref FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN fcrefs ON fclegs.fl_fh_id = fcrefs.rf_fh_id "	
					SQL = SQL&" WHERE (Fl_dr_ID='"&VehicleID&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND (fh_ship_dt>'"&now()-30&"') AND (rf_ref='"&FormBarCode(q)&"') AND (fl_st_id<>'13536')"
					If PageStatus="ONB" then
						If BillToID="48" then
							SQL = SQL& "AND (fh_status='PUO') "
							SQL = SQL& "AND (ref_status='p') "
							Else
							SQL = SQL& "AND ((fh_status='ACC') "
							If VehicleID=124  and LocationCode<>"ESTK" then
								SQL = SQL& " OR (fh_status='ARV')) "
								else
								SQL = SQL& " ) "
							End if
							SQL = SQL& "AND ((ref_status is NULL) "
							If VehicleID=124 and LocationCode<>"ESTK" then
								SQL = SQL& " OR (ref_status='H')) "
								else
								SQL = SQL& " ) "
							End if								
						End if
						'If VehicleID<>124 or LocationCode="ESTK" then
						'	SQL = SQL& "AND ((fl_sf_id='"&LocationCode&"') "
						'End if
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
						If VehicleID<>124 or LocationCode="ESTK" then	
							SQL = SQL& " ) "						
						End if
					end if					
					If PageStatus="CLS" then
							If VehicleID=124 then
								SQL = SQL& "AND ((fh_status='DPV') OR (fh_status='ONB') ) "
								SQL = SQL& "AND ((ref_status='S') OR  (ref_status='o')) "
								else
								SQL = SQL& "AND (fh_status='ONB') "
								SQL = SQL& "AND (ref_status='o') "
							End if						
						'If HUB2<>"y" then
						'	SQL = SQL&" AND ((Fl_st_ID='"&LocationCode&"')"
						'End if
						If HUB="y" then
							SQL = SQL&" OR (Fl_st_ID='D6W3')"
							SQL = SQL&" OR (Fl_st_ID='D6N2')"
							SQL = SQL&" OR (Fl_st_ID='D6N1')"
							SQL = SQL&" OR (Fl_st_ID='DM4M')"
							SQL = SQL&" OR (Fl_st_ID='DM5M')"
							SQL = SQL&" OR (Fl_st_ID='DPI2')"
							SQL = SQL&" OR (Fl_st_ID='DPI3')"
							SQL = SQL&" OR (Fl_st_ID='ESTK')"
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
						'If HUB2<>"y" then
						'	SQL = SQL&")"
						'End if													
					End if
					'response.write "SQL="&SQL&"<BR>"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''					
					oRs.Open SQL, DATABASE, 1, 3
					If oRs.EOF then
						ErrorMessage=ErrorMessage&" "&FormBarCode(q)&"  is not accepted.<br>Check Paper Work/Call Supervisor<br>"
					End if
					If not oRs.eof then
						TheJobNumber = oRs("fh_id")
						fl_pkey=oRs("fl_pkey")
						TheBarCode = FormBarCode(q)
						If Signature>"" then
						End if
						Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open DATABASE
						''''UPDATES THE WAFER
						l_cSQL = "UPDATE FCREFS SET ref_status = 'H' "
						l_cSQL = l_cSQL&" WHERE (rf_fh_id = '"&TheJobNumber&"') AND (rf_ref = '" &TheBarCode& "')"
						'Response.Write "l_cSQL="&l_cSQL&"<BR>"
						oConn.Execute(l_cSQL)
						
						''''''''''''''''''''''''''''''''''''''''
						Set oRs = Server.CreateObject("ADODB.Recordset")
						oRs.CursorLocation = 3
						oRs.CursorType = 3
						oRs.ActiveConnection = DATABASE	
						''''''CHECKS TO SEE IF ITS OKAY TO CHANGE THE ORDER STATUS
						SQL = "SELECT rf_fh_id FROM fcrefs"
						SQL = SQL&" WHERE (rf_fh_id='"&TheJobNumber&"')"
						SQL = SQL&" AND (Ref_Status='o')"
						oRs.Open SQL, DATABASE, 1, 3
						If oRs.EOF then
						'''''''''''''''''''''''''''''''''GETS THE OLD LEG INFO''''''''''''''''''''
						'response.write "**********DATABASE="&Database&"*************<BR>"
						'response.write "**********TheJobNumber="&TheJobNumber&"*************<BR>"
						Set Recordset1 = Server.CreateObject("ADODB.Recordset")
						Recordset1.ActiveConnection = DATABASE
						SQL678="SELECT * FROM fclegs WHERE (fl_fh_id='"& TheJobNumber &"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND (fl_st_id<>'13536')"
						'Response.Write "SQL678="&SQL678&"<BR>"
						Recordset1.Source = SQL678
						Recordset1.CursorType = 0
						Recordset1.CursorLocation = 2
						Recordset1.LockType = 1
						Recordset1.Open()
						Recordset1_numRows = 0
						If Recordset1.eof then
						End if			
						If not Recordset1.eof then
							Fl_pkey=trim(Recordset1("Fl_pkey"))
							Fl_fh_id=trim(Recordset1("Fl_fh_id"))
							Fl_joborder=trim(Recordset1("Fl_joborder"))
							Fl_sf_id=trim(Recordset1("Fl_sf_id"))
							Fl_sf_name=trim(Recordset1("Fl_sf_name"))
							Fl_sf_clname=trim(Recordset1("Fl_sf_clname"))
							Fl_sf_cfname=trim(Recordset1("Fl_sf_cfname"))
							Fl_sf_phone=trim(Recordset1("Fl_sf_phone"))
							Fl_sf_addr1=trim(Recordset1("Fl_sf_addr1"))
							Fl_sf_addr2=trim(Recordset1("Fl_sf_addr2"))
							Fl_sf_city=trim(Recordset1("Fl_sf_city"))
							Fl_sf_state=trim(Recordset1("Fl_sf_state"))
							Fl_sf_country=trim(Recordset1("Fl_sf_country"))
							Fl_sf_zip=trim(Recordset1("Fl_sf_zip"))
							Fl_st_id=trim(Recordset1("Fl_st_id"))
							Fl_st_name=trim(Recordset1("Fl_st_name"))
							Fl_st_clname=trim(Recordset1("Fl_st_clname"))
							Fl_st_cfname=trim(Recordset1("Fl_st_cfname"))
							Fl_st_phone=trim(Recordset1("Fl_st_phone"))
							Fl_st_addr1=trim(Recordset1("Fl_st_addr1"))
							Fl_st_addr2=trim(Recordset1("Fl_st_addr2"))
							Fl_st_city=trim(Recordset1("Fl_st_city"))
							Fl_st_state=trim(Recordset1("Fl_st_state"))
							Fl_st_country=trim(Recordset1("Fl_st_country"))
							Fl_st_zip=trim(Recordset1("Fl_st_zip"))
							Fl_estimate=trim(Recordset1("Fl_estimate"))
							Fl_permit=trim(Recordset1("Fl_permit"))
							Fl_un_id=trim(Recordset1("Fl_un_id"))
							Fl_dr_id=trim(Recordset1("Fl_dr_id"))
							Fl_st_odm=trim(Recordset1("Fl_st_odm"))
							Fl_end_odm=trim(Recordset1("Fl_end_odm"))
							Fl_trf_mi=trim(Recordset1("Fl_trf_mi"))
							Fl_load_mi=trim(Recordset1("Fl_load_mi"))
							Fl_empty_mi=trim(Recordset1("Fl_empty_mi"))
							Fl_toll_mi=trim(Recordset1("Fl_toll_mi"))
							Fl_load_ti=trim(Recordset1("Fl_load_ti"))
							Fl_unld_ti=trim(Recordset1("Fl_unld_ti"))
							Fl_trip_ti=trim(Recordset1("Fl_trip_ti"))
							Fl_rt_type=trim(Recordset1("Fl_rt_type"))
							Fl_totrate=trim(Recordset1("Fl_totrate"))
							Fl_wt_xc=trim(Recordset1("Fl_wt_xc"))
							Fl_wgt_xc=trim(Recordset1("Fl_wgt_xc"))
							Fl_pmrate=trim(Recordset1("Fl_pmrate"))
							Fl_escrate=trim(Recordset1("Fl_escrate"))
							Fl_codconrt=trim(Recordset1("Fl_codconrt"))
							Fl_codshprt=trim(Recordset1("Fl_codshprt"))
							Fl_miscrate=trim(Recordset1("Fl_miscrate"))
							Fl_mrdesc=trim(Recordset1("Fl_mrdesc"))
							Fl_pdrate=trim(Recordset1("Fl_pdrate"))
							Fl_estrate=trim(Recordset1("Fl_estrate"))
							Fl_flatrt=trim(Recordset1("Fl_flatrt"))
							Fl_prirt=trim(Recordset1("Fl_prirt"))
							Fl_pj_rt=trim(Recordset1("Fl_pj_rt"))
							Fl_sfstrt=trim(Recordset1("Fl_sfstrt"))
							Fl_codc_est=trim(Recordset1("Fl_codc_est"))
							Fl_cods_est=trim(Recordset1("Fl_cods_est"))
							Fl_st_comment=trim(Recordset1("Fl_st_comment"))
							Fl_sf_comment=trim(Recordset1("Fl_sf_comment"))
							Fl_sf_area=trim(Recordset1("Fl_sf_area"))
							Fl_st_area=trim(Recordset1("Fl_st_area"))
							Fl_t_disp=trim(Recordset1("Fl_t_disp"))
							Fl_t_acc=trim(Recordset1("Fl_t_acc"))
							Fl_t_atp=trim(Recordset1("Fl_t_atp"))
							Fl_t_int=trim(Recordset1("Fl_t_int"))
							Fl_t_atd=trim(Recordset1("Fl_t_atd"))
							Fl_t_und=trim(Recordset1("Fl_t_und"))
							Fl_st_rta=trim(Recordset1("Fl_st_rta"))
							Fl_sf_rta=trim(Recordset1("Fl_sf_rta"))
							Fl_weight=trim(Recordset1("Fl_weight"))
							Fl_pod=trim(Recordset1("Fl_pod"))
							Fl_wait_t=trim(Recordset1("Fl_wait_t"))
							Fl_feesadv=trim(Recordset1("Fl_feesadv"))
							Fl_fadesc=trim(Recordset1("Fl_fadesc"))
							Fl_rndtrip=trim(Recordset1("Fl_rndtrip"))
							Fl_rndt_rt=trim(Recordset1("Fl_rndt_rt"))
							Timestamp_column=trim(Recordset1("Timestamp_column"))
							Fl_zipmlrt=trim(Recordset1("Fl_zipmlrt"))
							Fl_numboxes=trim(Recordset1("Fl_numboxes"))
							Fl_hascod=trim(Recordset1("Fl_hascod"))
							Fl_boxrt=trim(Recordset1("Fl_boxrt"))
							Fl_disp=trim(Recordset1("Fl_disp"))
							Fl_sf_fullname=trim(Recordset1("Fl_sf_fullname"))
							Fl_st_fullname=trim(Recordset1("Fl_st_fullname"))
							Fl_user1=trim(Recordset1("Fl_user1"))
							Fl_user2=trim(Recordset1("Fl_user2"))
							Fl_podreq=trim(Recordset1("Fl_podreq"))
							Fl_rentmin=trim(Recordset1("Fl_rentmin"))
							Fl_rentrt=trim(Recordset1("Fl_rentrt"))
							Fl_boxtype=trim(Recordset1("Fl_boxtype"))
							Fl_permirt=trim(Recordset1("Fl_permirt"))
							Fl_dimwgt=trim(Recordset1("Fl_dimwgt"))
							Fl_dwfact=trim(Recordset1("Fl_dwfact"))
							Fl_pay_on=trim(Recordset1("Fl_pay_on"))
							Fl_ah_rt=trim(Recordset1("Fl_ah_rt"))
							Fl_ah_code=trim(Recordset1("Fl_ah_code"))
							Fl_pay_upd=trim(Recordset1("Fl_pay_upd"))
							Fl_cntyrt=trim(Recordset1("Fl_cntyrt"))
							Fl_user3=trim(Recordset1("Fl_user3"))
							Fl_user4=trim(Recordset1("Fl_user4"))
							Fl_billcd=trim(Recordset1("Fl_billcd"))
							Fl_sf_apt=trim(Recordset1("Fl_sf_apt"))
							Fl_st_apt=trim(Recordset1("Fl_st_apt"))
							Fl_firstdrop=trim(Recordset1("Fl_firstdrop"))
							Fl_seconb=trim(Recordset1("Fl_seconb"))
							Fl_secacc=trim(Recordset1("Fl_secacc"))
							Fl_pu_driver=trim(Recordset1("Fl_pu_driver"))
							Fl_pu_vehicle=trim(Recordset1("Fl_pu_vehicle"))
							Fl_do_driver=trim(Recordset1("Fl_do_driver"))
							Fl_do_vehicle=trim(Recordset1("Fl_do_vehicle"))
							Fl_job_closed=trim(Recordset1("Fl_job_closed"))
							Fl_leg_status=trim(Recordset1("Fl_leg_status"))

							'''CloseThis="n"
							''''NextToAddress=trim(Recordset1("fl_st_id"))
							
							'Response.Write "GOT HERE!!  ROLLOVER=YES<BR>"
						End if
						Set Recordset1 = Nothing
						
									Select Case fl_st_id
										Case "12203" '''''Stafford HUB
											Nextfl_un_id="4"
											Nextfl_dr_id="4"
										Case "12201" '''''Stafford final destination
											Nextfl_un_id="4"
											Nextfl_dr_id="4"
										Case "6430" '''''Sherman final destination
											Nextfl_un_id="6"
											Nextfl_dr_id="6"
										Case "6550", "7800", "RFAB", "13560", "13570", "13536F" '''''Spring Creek final destination
											Nextfl_un_id="3"
											Nextfl_dr_id="3"
                                            ''''Remvoed from below per Keith Chitwood 2/15/12 "13532", , "13536F" 
										Case "12500", "13121", "13353", "13353-7" '''''Spring Creek final destination
											Nextfl_un_id="7"
											Nextfl_dr_id="7"
										Case Else
											Nextfl_un_id="8"
											Nextfl_dr_id="8"
									End Select											
						''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''			
						'''''''''''''''''''''''''''''''ENDS GETS THE OLD LEG INFO'''''''''''''''''
							'''''''''''''''''''''''''''CREATES A NEW LEG''''''''''''''''''''''''''
						l_sql="PR_NextID 'sy_number', '1'" 
						Set oRs20=oConn.Execute(l_sql) 
						nextnumber = oRs20.Fields("sv_val")
						Set oRs20=Nothing
							'If Hub="y" or CloseThis="n" then
								'response.write "#1 PHONE_CHANGE_STATUS '" & TheJobNumber & "', '53', 'ARV', '', '',  '"& UserID &"', '"& UnitID &"'" 
								oConn.Execute "PHONE_CHANGE_STATUS '" & TheJobNumber & "', '53', 'ARV', '', '',  '"& UserID &"', '"& UnitID &"'" 
									''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
									''''''''''Do not make this LIVE yet, until KWE is OKAY to make live'''''''''''
									''''''''''THEN FIX FOR STOCKROOM!!!!!!!!!!!!!!!!!''''''''''''''''''''''''''''
									''''''''''''''''''''''''''''''''''''''''''''''''''
								'If FixedForStockroom="yes" then
									Set oConn43 = Server.CreateObject("ADODB.Connection")
									oConn43.ConnectionTimeout = 100
									oConn43.Provider = "MSDASQL"
									oConn43.Open DATABASE
									''''UPDATES CURRENT LEG TO INDICATE DROPPED!
									l_cSQL = "UPDATE FCLEGS SET fl_Leg_Status = 'd' "
									l_cSQL = l_cSQL&" WHERE (fl_fh_id = '"&TheJobNumber&"') AND (fl_pkey = '" & fl_pkey & "')"
									'response.write "l_cSQL="&l_cSQL&"<BR>"
									oConn43.Execute(l_cSQL)						
						
							''Set oConn888 = Server.CreateObject("ADODB.Connection")
							''	oConn888.ConnectionTimeout = 100
							''	oConn888.Provider = "MSDASQL"
							''	oConn888.Open DATABASE
								''''GETS NEW JOB NUMBER''''''''''''''''
							''	l_cSql888 = "EXEC pr_NextID_SP " & _
			 				''	"@p_cField ='sy_number', " & _
							''	"@p_nNum = 1, @p_nOut = 1 " 
							''	response.Write "l_cSql888="&l_cSql888&"<BR>"
							''	Set oRs888 = oConn888.Execute(l_cSql888)
							''	response.Write "ZZZ@p_nOut="& p_nOut &"<BR>"
								'nextNumber = oRs888.Fields("p_nOut")
							''	nextNumber = oConn888.parameters(3)
							''	Set oRs888 = Nothing
							''Set oConn888 = Nothing
							'response.write "ZZZnextnumber="& nextnumber &"<BR>"	
							'''''''''''''''''''ADDING THE ACTUAL LEG''''''''''''''''''
							Set oConn999 = Server.CreateObject("ADODB.Connection")
							oConn999.ConnectionTimeout = 100
							oConn999.Provider = "MSDASQL"
							oConn999.Open DATABASE
							l_cSQL999 ="INSERT INTO FCLEGS (Fl_pkey, Fl_fh_id, Fl_joborder, Fl_sf_id, " & _
							"Fl_sf_name, Fl_sf_clname, Fl_sf_cfname, Fl_sf_phone, Fl_sf_addr1, Fl_sf_addr2, " & _
							"Fl_sf_city, Fl_sf_state, Fl_sf_country, Fl_sf_zip, Fl_st_id, Fl_st_name, " & _
							"Fl_st_clname, Fl_st_cfname, Fl_st_phone, Fl_st_addr1, Fl_st_addr2, Fl_st_city, " & _
							"Fl_st_state, Fl_st_country, Fl_st_zip, Fl_estimate, Fl_permit, fl_un_id, fl_dr_id, " & _
							"Fl_st_odm, Fl_end_odm, Fl_trf_mi, Fl_load_mi, Fl_empty_mi, Fl_toll_mi, Fl_load_ti, " & _
							"Fl_unld_ti, Fl_trip_ti, Fl_rt_type, Fl_totrate, Fl_wt_xc, Fl_wgt_xc, Fl_pmrate, " & _
							"Fl_escrate, Fl_codconrt, Fl_codshprt, Fl_miscrate, Fl_mrdesc, Fl_pdrate, Fl_estrate, " & _
							"Fl_flatrt, Fl_prirt, Fl_pj_rt, Fl_sfstrt, Fl_codc_est, Fl_cods_est, Fl_st_comment, " & _
							"Fl_sf_comment, Fl_sf_area, Fl_st_area, Fl_t_disp, Fl_t_acc, Fl_t_atp, Fl_t_int, " & _
							"Fl_t_atd, Fl_t_und, Fl_st_rta, Fl_sf_rta, Fl_weight, Fl_pod, Fl_wait_t,  " & _
							"Fl_feesadv, Fl_fadesc, Fl_rndtrip, Fl_rndt_rt, Fl_zipmlrt, " & _
							"Fl_numboxes, Fl_hascod, Fl_boxrt, Fl_disp, Fl_sf_fullname, Fl_st_fullname, " & _
							"Fl_user1, Fl_user2, Fl_podreq, Fl_rentmin, Fl_rentrt, Fl_boxtype, Fl_permirt, " & _
							"Fl_dimwgt, Fl_dwfact, Fl_pay_on, Fl_ah_rt, Fl_ah_code, Fl_pay_upd, Fl_cntyrt, " & _
							"Fl_user3, Fl_user4, Fl_billcd, Fl_sf_apt, Fl_st_apt, Fl_firstdrop, Fl_seconb, " & _
							"Fl_secacc, Fl_pu_driver, Fl_pu_vehicle, Fl_do_driver, Fl_do_vehicle, Fl_job_closed, " & _
							"Fl_leg_status) " & _
							"VALUES ('"& nextNumber &"', '"& Fl_fh_id &"', '"& Fl_joborder &"', '13536', " & _
							"'13536 HUB', '', '', '', '13536 TI BLVD', '', " & _
							"'Dallas', 'TX', 'USA', '75243', '"& Fl_st_id &"', '"& Fl_st_name &"', " & _
							"'"& Fl_st_clname &"', '"& Fl_st_cfname &"', '"& Fl_st_phone &"', '"& Fl_st_addr1 &"', '"& Fl_st_addr2 &"', '"& pFl_st_city &"', " & _
							"'"& Fl_st_state &"', '"& Fl_st_country &"', '"& Fl_st_zip &"', '0', '0', '"& NextFl_un_id &"', '"& NextFl_dr_id &"', " & _
							"'"& Fl_st_odm &"', '"& Fl_end_odm &"', '"& Fl_trf_mi &"', '"& Fl_load_mi &"', '"& Fl_empty_mi &"', '"& Fl_toll_mi &"', '"& Fl_load_ti &"', " & _
							"'"& Fl_unld_ti &"', '"& Fl_trip_ti &"', '"& Fl_rt_type &"', '"& Fl_totrate &"', '"& Fl_wt_xc &"', '"& Fl_wgt_xc &"', '"& Fl_pmrate &"', " & _
							"'"& Fl_escrate &"', '"& Fl_codconrt &"', '"& Fl_codshprt &"', '"& Fl_miscrate &"', '"& Fl_mrdesc &"', '"& Fl_pdrate &"', '"& Fl_estrate &"', " & _
							"'"& Fl_flatrt &"', '"& Fl_prirt &"', '"& Fl_pj_rt &"', '"& Fl_sfstrt &"', '"& Fl_codc_est &"', '"& Fl_cods_est &"', '"& Fl_st_comment &"', " & _
							"'"& Fl_sf_comment &"', '"& Fl_sf_area &"', '"& Fl_st_area &"', '"& now() &"', '1/1/1900', '1/1/1900', '1/1/1900', " & _
							"'1/1/1900', '1/1/1900', '"& Fl_st_rta &"', '"& Fl_sf_rta &"', '"& Fl_weight &"', '"& Fl_pod &"', '"& Fl_wait_t &"',  " & _
							"'"& Fl_feesadv &"', '"& Fl_fadesc &"', '0', '"& Fl_rndt_rt &"',  '"& Fl_zipmlrt &"', " & _
							"'"& Fl_numboxes &"', '"& Fl_hascod &"', '"& Fl_boxrt &"', '"& Fl_disp &"', '"& Fl_sf_fullname &"', '"& Fl_st_fullname &"', " & _
							"'"& Fl_user1 &"', '"& Fl_user2 &"', '"& Fl_podreq &"', '"& Fl_rentmin &"', '"& Fl_rentrt &"', '"& Fl_boxtype &"', '"& Fl_permirt &"', " & _
							"'"& Fl_dimwgt &"', '"& Fl_dwfact &"', '"& Fl_pay_on &"', '"& Fl_ah_rt &"', '"& Fl_ah_code &"', '"& Fl_pay_upd &"', '"& Fl_cntyrt &"', " & _
							"'"& Fl_user3 &"', '"& Fl_user4 &"', '"& Fl_billcd &"', '"& Fl_sf_apt &"', '"& Fl_st_apt &"', NULL, NULL, " & _
							"NULL, '', '', '', '', '1/1/1900', " & _
							"'c') "
							'response.write "<BR><BR>XXXl_cSQL999="&l_cSQL999&"<BR><BR>"
							oConn999.Execute(l_cSQL999)
							oConn999.Close	
							''''''''''''WEIRD FCFGTHD PROBLEM_THIS STOPS IT FROM BECOMING OPN AGAIN!
														''''UPDATES CURRENT LEG TO INDICATE DROPPED!
									l_cSQL = "UPDATE FCFGTHD SET fh_Status = 'ARV', fh_Statcode='53' "
									l_cSQL = l_cSQL&" WHERE (fh_id = '"&TheJobNumber&"')"
									'response.write "l_cSQL="&l_cSQL&"<BR>"
									oConn43.Execute(l_cSQL)	
							'''''''''''''RESETS THE CORRECT DRIVER DUE TO STUPID TRIGGER I DON'T WANT TO MESS WITH AT THE MOMENT!
									l_cSQL = "UPDATE FCLEGS SET fl_un_id='"& Nextfl_un_id &"', fl_dr_id='"& Nextfl_dr_id &"' "
									l_cSQL = l_cSQL&" WHERE (fl_fh_id = '"&TheJobNumber&"') AND (Fl_pkey='"& nextNumber &"')"
									'response.write "l_cSQL="&l_cSQL&"<BR>"
									oConn43.Execute(l_cSQL)			
																				
							''''''''''''''''''''''''''ENDS CREATES A NEW LEG'''''''''''''''''''''''

									''''''''''''ROUTING PART'''''''''''''''''''''''''''''''''
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
		Set oConn=Nothing	
END IF
%>
<body leftMargin=0 TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" OnLoad="document.Form1.FormBarCode1.focus()">
	<TABLE WIDTH="300" border="0" bordercolor="red" cellpadding="0" cellspacing="5" align="left" ID="Table1">
		<tr><td align="center" colspan="3"><form method="post" action="Default.asp" ID="Form7"><input type="hidden" name="Aliascode" value="<%=AliasCode%>" ID="Hidden1"><input type="submit" value="Return to Home Page" ID="Submit1" NAME="Submit1"></form></td></tr>
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
				<form name="Form1" id="Form1" method="post">
				<%If ErrorMessage>"" Then%>
				<tr><td class="ErrorMessageBoldCenter" colspan="2"><%=ErrorMessage%></td></tr>
				<%End if%>
					<input type="hidden" name="Scanned" value="y" ID="Hidden3">
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
			colorset = split("#D2D9FC,White",",")
			numcolors = ubound(colorset)+1
			Server.ScriptTimeout = 1000
			optJobSel=Request.Querystring("optJobSel")
			optJobSel=Replace(optJobSel,"""","")
			optJobSel=Replace(optJobSel,"'","")
			If ReferenceNumber>"" then optJobSel="ByRef" end if
			If JobNumber>"" then optJobSel="ByJob" end if
					l_csql = "SELECT fcfgthd.fh_id, fcfgthd.fh_user5, fclegs.fl_sf_rta, fcfgthd.fh_status, fcfgthd.fh_ship_dt, fclegs.fl_pkey, fclegs.fl_sf_id, fclegs.fl_sf_name, fclegs.fl_st_id, fclegs.fl_st_name, fclegs.fl_t_atp, fclegs.fl_t_atd, fclegs.fl_pod, fcfgthd.fh_custpo, fcfgthd.fh_priority, fcrefs.rf_ref FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN fcrefs ON fclegs.fl_fh_id = fcrefs.rf_fh_id "	
				l_csql = l_csql&" WHERE (Fl_dr_ID='"&VehicleID&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND fl_st_id<>'13536' AND fh_ship_dt>'"&now()-30&"'"
						If PageStatus="ONB" then
							If BillToID="48" then
								l_csql = L_csql& "AND ((fh_status='PUO') OR (fh_status='AC2')) "
								l_csql = L_csql& "AND ((ref_status='p') or (ref_status='a')) "
								Else
								l_csql = L_csql& "AND ((fh_status='ACC') "
								If VehicleID=124 then
									l_csql = L_csql& " OR ((fh_status='AC2') AND (fl_secacc is not null))) "
									else
									l_csql = L_csql& " ) "
								End if
								l_csql = L_csql& "AND ((ref_status is NULL) "
								If VehicleID=124 then
									l_csql = L_csql& " OR (ref_status='H')) "
									else
									l_csql = L_csql& " ) "
								End if								
							End if
							If VehicleID<>124 then
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
							If VehicleID<>124 then	
								l_csql = l_csql& " ) "						
							End if
						end if
						If PageStatus="CLS" then
							If VehicleID=124 then
								l_csql = L_csql& "AND ((fh_status='DPV') OR (fh_status='ONB'))"
								l_csql = L_csql& "AND ((ref_status='S') or (ref_status='o')) "
								else
								l_csql = L_csql& "AND (fh_status='ONB') "
								l_csql = L_csql& "AND (ref_status='o') "
							End if
						End if
							SortBy="fh_priority, fh_id, rf_ref"
						If SortBy>"" then
							l_csql = L_csql& " ORDER BY "&Sortby
						End if
						'response.write "XXXXL_cSQL="&L_cSQL&"<BR>"
				Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				RSEVENTS2.CursorLocation = 3
				RSEVENTS2.CursorType = 3
				RSEVENTS2.ActiveConnection = DATABASE					
				RSEVENTS2.Open L_cSQL, DATABASE, 1, 3
				If not RSEVENTS2.EOF then
						ELSE
						Response.Redirect("Default.asp?AliasCode="&AliasCode&"&FakeSubmit=fakesubmit")
						'response.write "Should have re-directed<BR>"
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
					''''''''''''Determines if needs a POD or not''''''''''''''''''''''''
			Set Recordset1 = Server.CreateObject("ADODB.Recordset")
			Recordset1.ActiveConnection = DATABASE
			Recordset1.Source = "SELECT fl_pkey FROM fclegs WHERE (fl_PKey='"& fl_PKey+1 &"') and (fl_fh_id='"& fh_id &"')"
			'response.write "Recordset1.Source="&Recordset1.Source&"<BR>"
			Recordset1.CursorType = 0
			Recordset1.CursorLocation = 2
			Recordset1.LockType = 1
			Recordset1.Open()
			Recordset1_numRows = 0
			If Recordset1.eof then
				NeedPOD="y"
			End if			
			If not Recordset1.eof then
				Rollover="y"
				'Response.Write "GOT HERE!!  ROLLOVER=YES<BR>"
			End if
			Set Recordset1 = Nothing					
					''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
					
					%>
				<tr>
					<td class="mainpagetextboldcenter" nowrap><input type="text" name="FormBarCode<%=X%>" ID="Text2" size="3"></td>	
					<td class="MainPageTextSmaller" nowrap valign="top" bgcolor="<%=colorset(i mod numcolors)%>">
						<!--<b><%=fh_id%></b><br>--><%=X%>)&nbsp;<%=rf_ref%>
					</td>				
				</tr>
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
				%>
			<input type="hidden" name="LocationCode" value="<%=LocationCode%>" ID="Hidden8">
			<input type="hidden" name="BillToID" value="<%=BillToID%>" ID="Hidden9">
			<tr>
				<td colspan="2">
					<%if Pagestatus="CLS" and (BillToID="48" or BillToID="75" or BillToID="36") then%>
					<input type="submit" name="submit" value="submit" ID="Submit2" onclick="return validate()" />
					<%else%>
					<input type="submit" name="submit" value="submit" ID="Submit3">
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

