<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<%
Set objWSHNetwork = Server.CreateObject("WScript.Network") 
Response.Write mid(objWSHNetwork.ComputerName,8,1)


DisplayUserName=trim(Session("sUsername"))
'Response.Write "DisplayUserName="&DisplayUserName&"<BR>"
%>
<!-- #include file="../include/custom.inc" -->
<!-- #include file="../include/checkstring.inc" -->
<link rel="stylesheet" type="text/css" href="../../phone/mainStyleSheet.css">
<script language="javascript" type="text/javascript" src="datetimepicker.js">
    //Date Time Picker script- by TengYong Ng of http://www.rainforestnet.com
    //Script featured on JavaScript Kit (http://www.javascriptkit.com)
    //For this script, visit http://www.javascriptkit.com 
</script>
<script type="text/javascript">
    /***********************************************
    * Textarea Maxlength script- © Dynamic Drive (www.dynamicdrive.com)
    * This notice must stay intact for legal use.
    * Visit http://www.dynamicdrive.com/ for full source code
    ***********************************************/
    function ismaxlength(obj) {
        var mlength = obj.getAttribute ? parseInt(obj.getAttribute("maxlength")) : ""
        if (obj.getAttribute && obj.value.length > mlength)
            obj.value = obj.value.substring(0, mlength)
    }
</script>
<SCRIPT LANGUAGE=JAVASCRIPT>
    window.parent.frames['FraTop'].location.href = '../nav/navbar.asp?Mainmenu=order';
</SCRIPT>
<!-- #include file="../include/ifabsettings.inc" -->
<title><%=DisplayUserName%>  Reticle Order Page</title>
<%
	''''''''''''''''DATA FOR THE SCHEDULES''''''''''''''''''''''''
	DMonth=month(now())
	RMonth=DMonth
	DDay=day(now())
	RDay=DDay
	DYear=year(now())
	RYear=DYear
	'TodaysDate=DDay&"/"&DMonth&"/"&DYear
	TodaysDate=Date()
	LastDate=Cdate(TodaysDate)+60
	'''''Response.Write "XXXTodaysDate="&TodaysDate&"<BR>"
	'''''Response.Write "XXXLastDate="&LastDate&"<BR>"
	Select Case DMonth
		Case "1"
			WordMonth="January"
		Case "2"
			WordMonth="February"
		Case "3"
			WordMonth="March"
		Case "4"
			WordMonth="April"
		Case "5"
			WordMonth="May"
		Case "6"
			WordMonth="June"
		Case "7"
			WordMonth="July"
		Case "8"
			WordMonth="August"
		Case "9"
			WordMonth="September"
		Case "10"
			WordMonth="October"
		Case "11"
			WordMonth="November"
		Case "12"
			WordMonth="December"
	End Select	
	''''''''''''''''END THE DATA FOR THE SCHEDULES'''''''''''''''''''

l_lRoundTrip="0"
l_cIndustry="A"
l_cStatus="RAP"
HighlightedField="UserID"
PageStatus=Request.Form("PageStatus")
UserID=Request.Form("UserID")
BillToID=Trim(Session("sBT_ID"))
If BillToID="" then
	PageStatus="TimeOut"
	Response.redirect("../../intranet/default.asp")
end if
ShipMethod=Request.Form("ShipMethod")
If ShipMethod="" then
	ShipMethod="fleet"
End if
Pieces=1
Comments=Request.Form("Comments")
Comments=replace(Comments,"'","`")
Comments=replace(Comments,"""","`")
UserID=Request.Form("UserID")
SAPDate=Request.Form("SAPDate")
SAPTime=Request.Form("SAPTime")
NotificationEmail=Request.Form("NotificationEmail")
st_id=Request.Form("st_id")
destination=Request.Form("destination")
Quantity=Request.Form("Quantity")
Material=Request.Form("Material")
MaterialDescription=Request.Form("MaterialDescription")
MaterialDescription=replace(MaterialDescription,"'","`")
MaterialDescription=replace(MaterialDescription,"""","`")
DivNote=Request.Form("DivNote")
DivItem=Request.Form("DivItem")

'st_ID="CPGP"
PreLocationAlias=Request.Form("PreLocationAlias")
If Trim(PreLocationAlias)="" then
	PreLocationAlias = Request.Cookies("Location_Logisticorp")("LocationAlias")
End if
'response.write "PreLocationAlias="&PreLocationAlias&"XXXXX<BR>"
If trim(PreLocationAlias)>"" then
	LocationAlias=PreLocationAlias
End if
Submit=Request.Form("Submit")
'''''readytime=Request.Form("readytime")


l_cStatus="OPN"
Priority="ST"
st_id=trim(st_id)
Destination=trim(Destination)
SAPDateTime=SAPDate & " " & left(SAPTime,2) & ":" & right(SAPTime, 2)
txtDRemail=Left(NotificationEmail,8) & "@TI.com"
If trim(Quantity)>"" then
    FixedQuantity=int(Quantity)
End if
FixedDivNumber=DivNote & "/" & DivItem
Pieces=FixedQuantity
MaterialType=MaterialDescription
txtfh_custpo=FixedDivNumber
ReadyTime=SAPDateTime


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
							Set oConn = Server.CreateObject("ADODB.Connection")
							oConn.ConnectionTimeout = 100
							oConn.Provider = "MSDASQL"
							oConn.Open DATABASE
								l_cSQL2 = "select st_id, st_name, st_addr1 " &_
								"FROM fcshipto INNER JOIN fcshipbt ON fcshipto.st_id = fcshipbt.sb_st_id " &_
								"WHERE (fcshipbt.sb_bt_id = '26') AND st_alias = '" & TRIM(LocationAlias)&"'" 
								'response.write "l_cSQL2="&l_cSQL2&"<BR>"
								SET oRs = oConn.Execute(l_cSql2)
								If oRs.EOF then
									PageStatus="NoLocationCode"	
								End if
								Do While not oRs.EOF
								st_addr1=oRs("st_addr1")
								st_id=oRs("st_id")
								st_name=oRs("st_name")
								Response.Cookies("Location_Logisticorp").Expires = Date() + 3500
								Response.Cookies("Location_Logisticorp")("LocationAlias")=preLocationAlias				
								XYZ=XYZ+1								
								oRs.movenext
								LOOP
							Set oConn=Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'response.write "pagestatus="&pagestatus&"!!!!!!<BR>"
'response.write "st_id="&st_id&"!!!!!!!!<BR>"
'''''''ERROR HANDLING''''''''''

If PageStatus="OrderCompleted" then
	''Response.Write "ReadyTime="&ReadyTime&"***<BR>"
	''Response.Write "DueTime="&DueTime&"***<BR>"
''''''''''for testing only''''''''''
	''DueTime=(DateAdd("n",150,ReadyTime))
	'Response.Write "DueTime="&DueTime&"***<BR>"
	'If isdate(DueTime) then response.Write "HELLO????<BR>"	
	''DueTime=cDate(DueTime)
	''ReadyTime=cDate(ReadyTime)
	''Response.Write "DueTime="&DueTime&"***<BR>"
	''Response.Write "ReadyTime="&ReadyTime&"***<BR>"
''''''''''''''''''''''''''''''''''''

	If trim(UserID)=""  then 
		ErrorMessage="You must enter your badge number as a user id" 
		HighlightedField="UserID"
	end if
	If trim(SAPDate)=""  then 
		ErrorMessage="You must enter an SAP Date" 
		HighlightedField="SAPDate"
	end if
    If isDate(SAPDate) then
        DateDifference=DateDiff("d",SAPDate,TodaysDate)
        'Response.write "DateDifference="& DateDifference & "<BR>"
    End if
	If int(DateDifference)>2  then 
		ErrorMessage="The SAP Order time cannot be longer than 2 days ago." 
		HighlightedField="SAPDate"
	end if
	If trim(SAPTime)=""  then 
		ErrorMessage="You must enter an SAP Time" 
		HighlightedField="SAPTime"
	end if
	If trim(NotificationEmail)=""  then 
		ErrorMessage="You must enter a TI notification email address" 
		HighlightedField="NotificationEmail"
	end if
	If trim(st_id)=""  then 
		ErrorMessage="You must enter an origination" 
		HighlightedField="st_id"
	end if
	If trim(destination)=""  then 
		ErrorMessage="You must enter a destination" 
		HighlightedField="destination"
	end if
	If trim(Quantity)=""  then 
		ErrorMessage="You must enter a quantity" 
		HighlightedField="Quantity"
	end if
	If trim(Material)=""  then 
		ErrorMessage="You must enter the material" 
		HighlightedField="Material"
	end if
	If trim(MaterialDescription)=""  then 
		ErrorMessage="You must enter a material description" 
		HighlightedField="MaterialDescription"
	end if
	If trim(DivNote)=""  then 
		ErrorMessage="You must enter a div note" 
		HighlightedField="DivNote"
	end if
	If trim(DivItem)=""  then 
		ErrorMessage="You must enter a div item" 
		HighlightedField="DivItem"
	end if	

	If not isdate(SAPDate)  then 
		ErrorMessage="That SAP Date is not a valid date" 
		HighlightedField="SAPDate"
	end if
	If len(SAPTime)<>4  then 
		ErrorMessage="That is not a valid SAP Time (It must be four digits)" 
		HighlightedField="SAPTime"
	end if	
	If UCASE(Right(NotificationEmail,5))<>"I.COM"  then 
		ErrorMessage="That is not a valid TI email address" 
		HighlightedField="NotificationEmail"
	end if
	Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
		RSEVENTS.CursorLocation = 3
		RSEVENTS.CursorType = 3
		RSEVENTS.ActiveConnection = DATABASE
		SQL = "SELECT st_email FROM fcshipto where (st_id = '"& st_id &"')"
		RSEVENTS.Open SQL, DATABASE, 1, 3
		if RSEVENTS.eof then
			ErrorMessage="That is not a valid LogistiCorp/TI destination" 
			HighlightedField="destination"
		end if
		DestinationEmail=RSEVENTS("st_email")
		RSEVENTS.close
	Set RSEVENTS = Nothing
	If len(DivNote)<>10  then 
		ErrorMessage="The div note must be ten characters long" 
		HighlightedField="DivNote"
	end if
	If len(DivItem)<>6  then 
		ErrorMessage="The div item must be six characters long" 
		HighlightedField="DivItem"
	end if			
	If trim(st_id)<>"DSTK" AND trim(st_id)<>"ESTK" AND trim(st_id)<>"R1"  then 
		ErrorMessage="This is not a valid location to send items from.  Only DSTK, ESTK, and R1 are currently allowed" 
		HighlightedField="st_id"
	end if




			
	If trim(st_id)>"" AND ErrorMessage>"" then
		PageStatus="OrderForm1" 
	end if		


'Response.Write "HighlightedField="&HighlightedField&"***<BR>"
''''''END ERROR HANDLING'''''''
''''''FINDS DUE TIME
'Response.Write "DATABASE="&DATABASE&"<BR>"





''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.ConnectionTimeout = 100
oConn.Provider = "MSDASQL"
oConn.Open DATABASE
	l_cSQL = "select * FROM fcshipto INNER JOIN fcshipbt ON fcshipto.st_id = fcshipbt.sb_st_id " &_
	"WHERE (fcshipto.st_id = '"& st_id &"') AND (fcshipbt.sb_bt_id = '"& BillToID &"')"
	'Response.write "l_cSQL="&l_cSQL&"<BR>"
	SET oRs = oConn.Execute(l_cSql)
	IF not oRs.EOF then	
		txtPUCompany=trim(oRs("st_Name"))
		txtPUContact=trim(oRs("st_Fullname"))
		if len(trim(txtPUContact))=1 then
			txtPUContact=""
		End if		
		txtPUPhone=trim(oRs("st_cphone"))
		fl_sf_addr1=trim(oRs("st_addr1"))
		fl_sf_addr2=trim(oRs("st_addr2"))
		fl_sf_city=trim(oRs("st_City"))
		fl_sf_state=trim(oRs("st_State"))
		txtPUZip=trim(oRs("st_Zip"))
	End if
Set oConn=Nothing
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.ConnectionTimeout = 100
oConn.Provider = "MSDASQL"
oConn.Open DATABASE
	l_cSQL = "select * from fcshipto  " &_
	"WHERE st_id = '" & TRIM(Destination)&"'" 
	'Response.write "l_cSQL="&l_cSQL&"<BR>"
	SET oRs = oConn.Execute(l_cSql)
	If oRs.eof then
		ErrorMessage=Destination & " is not a valid LogistiCorp dropzone.  Please see list of <a href='Dropzones_SR.asp' target='_blank'>valid dropzones and codes</a>" 
		HighlightedField="Destination"	
	End if	
	IF not oRs.EOF then	
		txtDRCompany=trim(oRs("st_Name"))
		txtDRContact=trim(oRs("st_Fullname"))
		if len(trim(txtDRContact))=1 then
			txtDRContact=""
		End if
		txtDRPhone=trim(oRs("st_cphone"))
		fl_st_addr1=trim(oRs("st_addr1"))
		fl_st_addr2=trim(oRs("st_addr2"))
		fl_st_city=trim(oRs("st_City"))
		fl_st_state=trim(oRs("st_State"))
		txtDRZip=trim(oRs("st_Zip"))
	End if
Set oConn=Nothing
End if

If PageStatus="OrderCompleted" and ErrorMessage="" then
	'''''DueTime=(DateAdd("n",PriorityTime,ReadyTime))&"<BR>"
'''''''FINDS CONTACT INFO FOR THE DROPZONE	
	
	'DueTime=(DateAdd("n",PriorityTime,ReadyTime))
	DueTime=(DateAdd("n",120,ReadyTime))
	'Response.Write "DueTime="&DueTime&"***<BR>"
	'If isdate(DueTime) then response.Write "HELLO????<BR>"	
	DueTime=cDate(DueTime)
	ReadyTime=cDate(ReadyTime)
	'Response.Write "DueTime="&DueTime&"***<BR>"
	'Response.Write "ReadyTime="&ReadyTime&"***<BR>"
    If lcase(trim(Destination))="cssf" then
        Destination="CSSF-SR"
    End if

	''''XXXXXXXXXXXXPLACEORDERHEREXXXXXXXXXXX
	Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.ConnectionTimeout = 100
		oConn.Provider = "MSDASQL"
		oConn.Open DATABASE
		''''GETSNEWJOBNUMBER
		l_cSql = "EXEC pr_GetJobNum"
		Set oRs = oConn.Execute(l_cSql)
		newjobnum = oRs.Fields("fh_id")	
		''''PLACESJOB
		'Response.write "newjobnum="&newjobnum&"<BR>"
		
		l_cSQL = "EXEC pr_bookjob " & _
		"@p_cJobNum ='" & left(TRIM(newjobnum),8) & "', " & _
		"@p_cbtid = '" & left(TRIM(BillToID),12) & "', " &_
		"@p_cpriority = '" & left(TRIM(Priority),2) & "', " & _
		"@p_cVertical = '" & left(TRIM(l_cIndustry),1) & "', " & _
		"@p_cstatus = '" & left(TRIM(l_cStatus),3) & "', " & _
		"@p_lroundtrip = " & CSTR(l_lRoundTrip) & ", " & _
		"@p_RetPri = '" & left(TRIM(l_cRetPri),2) & "', " & _
		"@p_cPUID = '" & left(TRIM(st_id),12) & "', " & _
		"@p_cPUCompany = '" & left(TRIM(Replace(txtPUCompany,"'", "")),40) & "', " & _
		"@p_cPUContact = '" & left(TRIM(Replace(txtPUContact,"'", "")),28) & "', " & _
		"@p_cPUPhone = '" & left(TRIM(txtPUPhone),20) & "', " & _
		"@p_cPUAddr1 = '" & left(TRIM(Replace(fl_sf_addr1,"'", "")),40) & "', " & _
		"@p_cPUAddr2 = '" & left(TRIM(Replace(fl_sf_addr2,"'", "")),40) & "', " &_
		"@p_cPUCity = '" & left(TRIM(Replace(fl_sf_city,"'", "")),30) & "', " & _
		"@p_cPUState = '" & left(TRIM(fl_sf_state),3) & "', " &_
		"@p_cPUZip = '" & left(TRIM(txtPUZip),10) & "', " & _
		"@p_cPUInstr = '" & Comments & "', " & _
		"@p_cPUEmail = '" & LEFT(TRIM(txtpuemail),40) & "', " & _
		"@p_cDRID ='" & left(TRIM(Destination),12) & "', " & _
		"@p_cDRCompany = '" & left(TRIM(Replace(txtDRCompany,"'", "")),40) & "', " & _
		"@p_cDRContact = '" & left(TRIM(Replace(txtDRContact,"'", "")),28) & "', " & _
		"@p_cDRPhone = '" & left(TRIM(txtDRPhone),20) & "', " & _
		"@p_cDRAddr1 = '" & left(TRIM(Replace(fl_st_addr1,"'", "")),40) & "', " & _
		"@p_cDRAddr2 = '" & left(TRIM(Replace(fl_st_addr2,"'", "")),40) & "', " &_
		"@p_cDRCity = '" & left(TRIM(Replace(fl_st_city,"'", "")),30) & "', " & _
		"@p_cDRState = '" & left(TRIM(fl_st_state),3) & "', " &_
		"@p_cDRZip = '" & left(TRIM(txtDRZip),10) & "', " & _
		"@p_cDRInstr = '" & left(TRIM(Replace(txtfl_st_comment,"'", "")),100) & "', " & _
		"@p_cDREmail = '" & LEFT(TRIM(nada),40) & "', " & _
		"@p_tready = '" & ReadyTime & "', " & _
		"@p_tdue = '" & DueTime & "', " & _
		"@p_cReference = '" & left(TRIM(Replace(txtfh_custpo,"'", "")),24) & "', " & _
		"@p_npieces = " & Pieces & ", " & _
		"@p_cPieceType = '" & left(TRIM(l_cPieceType),10) & "', " & _
		"@p_nweight = 0, " & _
		"@p_cTruckType = '" & left(TRIM(lstTruckType),12) & "', " & _
		"@p_lUpBT=1, " & _
		"@p_lAddOnFly='1'," &_
		"@p_cFrmAirport = '" & left(TRIM(FrmAirport),30) & "', " & _
		"@p_cAirline = '" & left(TRIM(Airline),30) & "', " & _
		"@p_cFlight = '" & left(TRIM(Flight),30) & "', " & _
		"@p_cToAirport = '" & left(TRIM(ToAirport),30) & "', " & _
		"@p_cLabelToAirport = '" & TRIM(FlightTimelbl) & "', " &_
		"@p_cNoFlightChk = '" & TRIM(NoFltChk) & "', " &_
		"@p_cPaymentType = '" & Trim(PaymentType) & "', " &_
		"@p_cCharter = '" & Trim(Charter)  & "', " &_
		"@p_cFlightTime = '" & l_cFlightTime & "' , " &_
		"@p_cPUArea = '" & l_cPUArea & "', " &_
		"@p_cDRArea = '" & l_cDRArea & "', " &_	
		"@p_cCoId = '" & UserID & "', " &_
		"@p_cfhuser3 = '" & User3 & "', " & _		
		"@p_cfhuser5 = '" & materialtype & "'"	
		'Response.Write "l_cSQL="&l_cSQL&"<BR>"
		
		Set oRs = oConn.Execute(l_cSql)
		Set oConn=Nothing
		Set oRs=Nothing
			Set oConn = Server.CreateObject("ADODB.Connection")
			oConn.ConnectionTimeout = 100
			oConn.Provider = "MSDASQL"
			oConn.Open DATABASE		
				l_nPkey = m_GetPkey(oConn, 1)
			Set oConn=Nothing
			If Trim(txtfh_custpo)>"" then	
			'NewJobNumPlusOne=NewJobNum+1
			'Response.Write "NewJobNumPlusOne="&NewJobNumPlusOne&"<BR>"
			'for lll=1 to 3
				'If Len(NewJobNumPlusOne)<8 then
					'Response.Write "Got here!<BR>"
					'NewJobNumPlusOne="0"&NewJobNumPlusOne
					'Response.Write "NewJobNumPlusOne="&NewJobNumPlusOne&"****<BR>"
				'End if
			'next																	
				Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
					RSEVENTS2.Open "FCRefs", DATABASE, 2, 2
					RSEVENTS2.addnew
					RSEVENTS2("rf_Pkey")=l_nPkey
					RSEVENTS2("rf_fh_ID")=NewJobNum
					RSEVENTS2("rf_ref")=txtfh_custpo
					If ShipMethod="bus" or ShipMethod="air" or ShipMethod="drive" then
						RSEVENTS2("ref_status")="o"
					End if									
					RSEVENTS2.update
					RSEVENTS2.close			
				set RSEVENTS2 = nothing	
				'If trim(St_id)<>"ESTK" then

                txtDREmail=lcase(txtDREmail)
                Select Case txtDREmail
                    Case "a0206910@ti.com", "a0457936@ti.com", "a0206066@ti.com", "a0215401@ti.com", "a0205657@ti.com", "a0208775@ti.com", "a0864147@ti.com", "a0864119@ti.com", "a0218446@ti.com", "a0451149@ti.com", "a0203419@ti.com", "a0460712@ti.com", "a0865911@ti.com"
                        txtDREmail="dm6pheetech_txtmsg@list.ti.com"

                End select

                Select Case userID
                    Case "a0459667", "a0206239", "a0864120", "a0865460", "a0209312", "a0321143", "a0206219", "a0865632", "a0209434", "a0218735", "a0203928", "a0206225", "a0209235", "a0865824", "a0456776", "a0865661", "a0460813"
                        txtDREmail=txtDREmail&";dm6pleetech_txtmsg@list.ti.com"
                End Select
				
                Material=Replace(Material,"/","-")
					Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
						RSEVENTS2.Open "DeliveryNotifications", DATABASE, 2, 2
						RSEVENTS2.addnew
						RSEVENTS2("fh_ID")=NewJobNum
						RSEVENTS2("ref_id")=txtfh_custpo
						RSEVENTS2("Material")=Material
						RSEVENTS2("MaterialDescription")=MaterialDescription
						RSEVENTS2("EmailAddress")=txtDRemail
						RSEVENTS2("DeliveryNotificationStatus")="c"
				
						RSEVENTS2.update
						RSEVENTS2.close			
					set RSEVENTS2 = nothing	

				'End if				
			End if
            OtherDateDifference=DateDiff("n", SAPDateTime, now())
            'Response.write "Now()="&Now()&"<BR>"
            'Response.write "SAPDateTime="&SAPDateTime&"<BR>" 
             'Response.write "OtherDateDifference="&OtherDateDifference&"<BR>"
             If OtherDateDifference>60 then
                            'Response.write "GOT HERE!!!<BR>"
		        sHTML="Dear Keith,<br><br>"
		        sHTML=sHTML&"A stockroom order has taken over an hour to pick.<br><br>"
		        sHTML=sHTML&"Below is the information that was submitted:<br><br>"
		        sHTML=sHTML&"Job Number: "&NewJobNum&"<br>"
		        sHTML=sHTML&"SAP Order Time: "&SAPDateTime&"<br>"
		        sHTML=sHTML&"Pick Time: "&Now()&"<br>"
		

		        Set objMail = CreateObject("CDONTS.Newmail")
		        objMail.From = "System.Notification@Logisticorp.us"
		        'objMail.To = "mark.maggiore@logisticorp.us"
                objMail.To = "alex.castillo@logisticorp.us;kchitwood@ti.com;mark.maggiore@logisticorp.us"
                objMail.CC= "x0019307@ti.com;RDBaker@ti.com"
		        objMail.Subject = "Long Stockroom Pick Time"
		        objMail.MailFormat = cdoMailFormatMIME
		        objMail.BodyFormat = cdoBodyFormatHTML
		        objMail.Body = sHTML
		        objMail.Send
		        Set objMail = Nothing	

             End if 



			Response.Redirect("../include/fnlrecap.asp?l_cJobNum=" & newjobnum & " ")

          
            
            	else
	If trim(st_id)>"" then
		PageStatus="OrderForm1"
	End if
End if
'Response.Write "BillToID="&BillToID&"<BR>"
'Response.Write "OtherDisplayUserName="&OtherDisplayUserName&"<BR>"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
preLocationAlias=trim(Request.Form("preLocationAlias"))
'response.write "preLocationAlias="&preLocationAlias&"***<BR>"
'response.write "st_id="&st_id&"***<BR>"

'If ErrorMessage="" then
'''''Response.Write "GOT HERE 3<BR>"
	If ResetCookie<>"y" then
		PreLocationAlias=Request.Cookies("Location_Logisticorp")("LocationAlias")
		'Response.Write "***PreLocationAlias="&PreLocationAlias&"*****<BR>"
	end if
	'Response.write "***PreLocationAlias="&PreLocationAlias&"*****<BR>"
	If PreLocationAlias<>"DSTK" AND PreLocationAlias<>"ESTK" then PreLocationAlias=""
	If Trim(PrelocationAlias)="" then PrelocationAlias="666"
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.ConnectionTimeout = 100
	oConn.Provider = "MSDASQL"
	oConn.Open DATABASE
		l_cSQL = "select st_alias, st_id, st_addr1, st_addr2 from fcshipto  " &_
		"WHERE st_alias = '" & TRIM(PreLocationAlias)&"'" 
		'Response.write "l_cSQL="&l_cSQL&"<BR>"
		SET oRs = oConn.Execute(l_cSql)
		IF not oRs.EOF then	
			LocationAlias=oRs("st_alias")
			st_addr1=oRs("st_addr1")
			st_id=trim(oRs("st_id"))
			'''''if st_id="CPGP" then
			'''''	st_addr1=oRs("st_addr2")
			'''''End if
			'Response.Cookies("Location_Logisticorp").expires=#1/1/2015# 
			'Response.Cookies("Location_Logisticorp")("LocationAlias")=preLocationAlias	
			''''else
			''''LocationAlias=""
			'''''ErrorMessage=""		
		End if
	Set oConn=Nothing
'End if	
if LocationAlias>"" then
	'''Response.Write "got here 111<br>"
	If PageStatus="" then
		'''Response.Write "got here 222111<br>"
		PageStatus="OrderForm1"
	End if
	PrelocationAlias=Trim(LocationAlias)
End if

'PageStatus="TimeOut"
'''Response.Write "LocationAlias="&LocationAlias&"****<BR>"
'response.write "PageStatus="&PageStatus&"****<BR>"
'''Response.Write "PreLocationAlias="&PreLocationAlias&"****<BR>"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>
</head>
<%
Select Case PageStatus
	Case "TimeOut"
		%>
		<script language="JavaScript">
		    open("../../intranet/default.asp", "_top")
		</script> 
		<%
	Case "OrderForm1"
	%>
<BODY leftMargin="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" OnLoad=document.OrderForm1.<%=HighlightedField%>.focus()>
<FORM method="post" name="OrderForm1" ID="OrderForm1">
<table cellpadding="0" cellspacing="0" ID="Table1" width="100%" bordercolor="green" border="0">
	<tr>
		<td width="20">&nbsp;</td>
		<td>
			<table width="100%" cellpadding="0" cellspacing="0" bordercolor="red" border="0" ID="Table2" valign="top">
				<%
				'DisplayUserName=trim(Session("sUsername"))
				'Response.Write "DisplayUserName="&DisplayUserName&"<BR>"
				'if DisplayUserName="comps" then%>
					<!--
					<tr><td valign="top"><a href="StockroomDefault.asp"><img src="../images/<%=DisplayUserName%>.jpg" height="<%=DisplayHeight%>" width="<%=DisplayWidth%>" border="0"></td><td><img src="images/pixel.gif" width="1" height="1"></a></td></tr>
					-->
				<%'end if%>
				<tr>
					<td align="left" nowrap class="pagetitle">Stockroom Order Entry</td>
					<td align="right" nowrap>To place a Window Ticket order, <a href="pmonepage.asp">click here</a></td>
				</tr>
				<tr><td height="5"><IMG SRC="images/pixel.gif" width="1" height="1"></td></tr>
				<tr>
					<td class="MainPageTextBold" nowrap>Employee ID:&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<td width="95%">
						<input maxlength="25" size="25" name="UserID" class="inputgeneral" value="<%=UserID%>" ID="Text1">
					</td>
				</tr>
				<tr><td height="5"><IMG SRC="images/pixel.gif" width="1" height="1"></td></tr>
				<tr>
					<td class="MainPageTextBold" nowrap>SAP Order Date:&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<td width="95%">
						<input maxlength="10" size="25" name="SAPDate" class="inputgeneral" value="<%=SAPDate%>" ID="Text10">
					</td>
				</tr>
				<tr><td height="5"><IMG SRC="images/pixel.gif" width="1" height="1"></td></tr>
				<tr>
					<td class="MainPageTextBold" nowrap>SAP Order Time:&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<td width="95%">
						<input maxlength="5" size="25" name="SAPTime" class="inputgeneral" value="<%=SAPTime%>" ID="Text11">
					</td>
				</tr>
				<tr><td height="5"><IMG SRC="images/pixel.gif" width="1" height="1"></td></tr>
				<tr>
					<td class="MainPageTextBold" nowrap>Notification Email:&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<td width="95%">
						<input maxlength="75" size="25" name="NotificationEmail" class="inputgeneral" value="<%=NotificationEmail%>" ID="Text12">
					</td>
				</tr>
				<tr><td height="5"><IMG SRC="images/pixel.gif" width="1" height="1"></td></tr>																
				<tr>
					<td class="MainPageTextBold" nowrap valign="top">Origination: </td>
					<td width="95%">
								<input type="hidden" name="st_id" value="<%=st_id%>" ID="Hidden4">
								<%
								Response.Write(st_name)	
								%>				
					</td>
				</tr>				
				<tr><td height="5"><IMG SRC="images/pixel.gif" width="1" height="1"></td></tr>
				<tr>
					<td class="MainPageTextBold" nowrap>Destination:&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<td width="95%">
						<input maxlength="10" size="25" name="Destination" class="inputgeneral" value="<%=Destination%>" ID="Text13">
					</td>
				</tr>
				<tr><td height="5"><IMG SRC="images/pixel.gif" width="1" height="1"></td></tr>
				<tr>
					<td class="MainPageTextBold" nowrap>Quantity:&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<td width="95%">
						<input maxlength="25" size="25" name="Quantity" class="inputgeneral" value="<%=Quantity%>" ID="Text14">
					</td>
				</tr>
				<tr><td height="5"><IMG SRC="images/pixel.gif" width="1" height="1"></td></tr>
				<tr>
					<td class="MainPageTextBold" nowrap>Material:&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<td width="95%">
						<input maxlength="25" size="25" name="Material" class="inputgeneral" value="<%=Material%>" ID="Text15">
					</td>
				</tr>
				<tr><td height="5"><IMG SRC="images/pixel.gif" width="1" height="1"></td></tr>				
				<tr>
					<td class="MainPageTextBold" nowrap>Material Description:&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<td width="95%">
						<input maxlength="75" size="75" name="MaterialDescription" class="inputgeneral" value="<%=MaterialDescription%>" ID="Text16">
					</td>
				</tr>
				<tr><td height="5"><IMG SRC="images/pixel.gif" width="1" height="1"></td></tr>
				<tr>
					<td class="MainPageTextBold" nowrap>Div Note:&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<td width="95%">
						<input maxlength="10" size="25" name="DivNote" class="inputgeneral" value="<%=DivNote%>" ID="Text17">
					</td>
				</tr>
				<tr><td height="5"><IMG SRC="images/pixel.gif" width="1" height="1"></td></tr>
				<tr>
					<td class="MainPageTextBold" nowrap>Div Item:&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<td width="95%">
						<input maxlength="6" size="25" name="DivItem" class="inputgeneral" value="<%=DivItem%>" ID="Text18">
					</td>
				</tr>
				<tr><td height="5"><IMG SRC="images/pixel.gif" width="1" height="1"></td></tr>
				<tr>
					<td class="MainPageTextBold" nowrap>Comments:&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<td width="95%">
						<textarea name="comments" cols="50" rows="2"><%=Comments%></textarea>
					</td>
				</tr>				
				<tr><td height="5"><IMG SRC="images/pixel.gif" width="1" height="1"></td></tr>														
				</table>
				</td>
				</tr>
				
				
				
				
				<tr><td height="5"><IMG SRC="images/pixel.gif" width="1" height="1"></td></tr>

				<tr><td colspan="2" align="center" class="ErrorMessage"><b><%=ErrorMessage%></b></td></tr>
				<tr><td height="5"><IMG SRC="images/pixel.gif" width="1" height="1"></td></tr>
				<tr>
					<td colspan="2" align="center">
						<input type="hidden" name="fl_sf_addr1" value="<%=fl_sf_addr1%>" ID="Hidden15">
						<input type="hidden" name="fl_sf_addr2" value="<%=fl_sf_addr2%>" ID="Hidden16">
						<input type="hidden" name="fl_sf_city" value="<%=fl_sf_city%>" ID="Hidden17">
						<input type="hidden" name="fl_sf_state" value="<%=fl_sf_state%>" ID="Hidden18">
						<input type="hidden" name="PreLocationAlias" value="<%=PreLocationAlias%>" ID="Hidden2">
						<input type="hidden" name="PageStatus" value="OrderCompleted" ID="Hidden1">
						<input type="submit" name="submit" value="Submit" ID="Submit1">	
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</form>
	<%
	Case else
	%>
	<BODY OnLoad="document.form666.preLocationAlias.focus()" leftMargin=0 TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
	<!--THIS VERIFIES/LOADS THE LOCATION COOKIE...Mark Maggiore-->
	<%
	'LocationAlias=Request.Cookies("Location_Logisticorp")("LocationAlias")	
	'If locationalias="" then
	'LocationAlias=Request.Cookies("Location_Logisticorp")("LocationAlias")
		'Response.Write "LocationAlias="&LocationAlias&"<BR>"
		'if LocationAlias="" then
			'Response.Write "here's some code!<br>"
			%>
			<table align="center" width="600" ID="Table5">
				<tr><td height="5"><IMG SRC="images/pixel.gif" width="1" height="1"></td></tr>
				<tr>
					<td class="generalcontent">If you wish to <b>TRACK AN ORDER</b> or <b>VIEW REPORTS</b>, you can do so by choosing your
					option from the navigation bar above.<br><br>
					If however, you are trying to use this computer to place an order or close an order (and
					this computer has been approved for such use by your IFAB/Transport person), then you 
					will need to follow the instructions below in order to enable this computer.</td>
				</tr>
				<tr><td height="5"><IMG SRC="images/pixel.gif" width="1" height="1"></td></tr>
			</table>			
			<form method="post" ID="form666" name="form666">
			<table width="500" border="0" align="center" ID="Table6">
				<tr>
					<td>
						<table width="500" border="0" align="center" ID="Table18">
							<tr>				
								<td class="generalcontent">
									SCAN in the location barcode of this terminal (As provided by Mark Maggiore). 
									If you are unable to SCAN in the terminal location, you must contact 
									Mark Maggiore at 214-956-0400 xt. 212 
									for assistance.
								</td>
							</tr>
							<tr><td height="5"><IMG SRC="images/pixel.gif" width="1" height="1"></td></tr>
							<tr>
								<td align="center" class="generalcontent">
									<input type="password" name="preLocationAlias" ID="Password1">
								</td>
							</tr>
							<tr><td align="center" class="generalcontent"><input type="submit" value="submit" name="submit" ID="Submit7"></td></tr>
							<%
							If ErrorMessage>"" then
								%>
								<tr><td height="5"><IMG SRC="images/pixel.gif" width="1" height="1"></td></tr>
								<tr><td><font color="red"><b><%=ErrorMessage%></b></font></td></tr>
								<tr><td height="5"><IMG SRC="images/pixel.gif" width="1" height="1"></td></tr>
								<%
							End if
							%>
						</table>
					</td>
				</tr>
			</table>
			</form>
			<%
		'End if
	'End if	
	End Select
	'response.write "pagestatus="&pagestatus&"<BR>"
	%>
</body>
</html>
