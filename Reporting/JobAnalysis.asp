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
    PageTitle="SHIPMENT DETAILS"

%>
<title>FleetX - <%=PageTitle %></title>

<%
ShowPrevious=Request.QueryString("ShowPrevious")
ShowNext=Request.QueryString("ShowNext")
InputJobNumber=trim(Request.Form("InputJobNumber"))
If InputJobNumber="" then
	InputJobNumber=trim(Request.QueryString("InputJobNumber"))
End if
InputLotNumber=trim(Request.Form("InputLotNumber"))
If InputLotNumber="" then
	InputLotNumber=trim(Request.QueryString("InputLotNumber"))
End if
InputDocumentNumber=trim(Request.Form("InputDocumentNumber"))
If InputDocumentNumber="" then
	InputDocumentNumber=trim(Request.QueryString("InputDocumentNumber"))
End if
If InputJobNumber>"" or InputLotNumber>"" or InputDocumentNumber>"" then
	hdnHaveParms="YES"
	else
	hdnHaveParms=Request.Form("hdnHaveParms")
End if
'If Trim(BillToID)="" then
'	Response.Redirect("../../intranet")
'End if
'''''''''''''QUERY STATEMENT'''''''''''''''''''''''''''''''''''''''''''
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.ConnectionTimeout = 200
oConn.Provider = "MSDASQL"
oConn.Open DATABASE
'response.write "Database="&Database&"<br>"
'response.write "billtoid="&billtoid&"<BR>"
'response.write "sbt_id="&sbt_id&"<BR>"

l_cSQL=l_cSQL&"Select * from FleetXOrderView "
if trim(sbt_id)="26" then
	l_cSQL="Select * from marksview2_sr "
End if
l_cSQL=l_cSQL&" WHERE (jobnum > '""')"
If InputJobNumber>"" then     
	'l_cSQL=l_cSQL&" AND (jobnum like '%"&InputJobNumber&"')"
	l_cSQL=l_cSQL&" AND (jobnum = '"&InputJobNumber&"')"
End if		 
If InputDocumentNumber>"" then     
	l_cSQL=l_cSQL&" AND (custpo like '%"&InputDocumentNumber&"')"
End if
If InputLotNumber>"" then     
	l_cSQL=l_cSQL&" AND (ref like '%"&InputLotNumber&"')"
End if		
l_cSQL=l_cSQL&" Order by shipdate DESC" 
'response.write "85 jobanalysis l_cSQL="&l_cSQL&"<BR>"
'response.write "!!!Database="&Database&"<BR>"
'response.write "!!!BillToID="&BillToID&"<BR>"

'response.Write "name="&Session("txt_cm_desc")&"<BR>"
Set oRs = oConn.Execute(l_cSQL)
If oRs.eof then
	ErrorMessage="There are no orders that match your criteria"
end if
If Err.Number <> 0 Then                                               
Response.Write ErrorMessage="Error Executing the query.  Error:" & Err.Description
End if
IF NOT oRs.EOF THEN
	oRs.MoveFirst
	OrderID=oRs("jobnum")
	DocumentNumber=oRs("custpo")
	ToLocation=oRs("to_id")
	
	SubmittedBy=oRs("TIUser")
	Priority=oRs("priority")
	'ShippingOrderTime=oRs("shipdate")
	BookTime=oRs("Shipdate")
	FromLocation=trim(oRs("from_id"))
	PaperworkTime=oRs("paperwork")
	DriverID=oRs("driver")
	'DriverID=0
	unit=oRs("unit")
	AtAirlineTime=oRs("atairline")
	DueTime=oRs("duetime")
	
	DispatchTime=oRs("disptime")
	DriverAcknowledgementTime=oRs("acctime")
	OnBoardTime=oRs("onbtime")
	If OnBoardTime="1/1/1900" then 
		DisplayOnBoardTime="Pending"
		else
		DisplayOnBoardTime=OnBoardTime
	End if
	DropTime=oRs("droptime")

	SAPOrderTime=oRs("readytime")
	FromComments=oRs("sfcomment")
	ToComments=oRs("stcomment")	
	
	


	
	If Trim(FromComments)>"" then 
		DisplayFromComments=FromComments
		else
		DisplayFromComments="none"
	End if	
	If Trim(ToComments)>"" then 
		DisplayToComments=ToComments
		else
		DisplayToComments="none"
	End if					
	''''''''NEW VARIABLES
	ArrivedAtHUB=ucase(trim(oRs("at_HUB")))
	
	DepartedHUB=ucase(trim(oRs("onbleg2")))
	AcknowledgedHUB=trim(oRs("accleg2"))

	'Response.Write "AcknowledgedHUB="&AcknowledgedHUB&"!!!<BR>"
		
	'Response.Write "ArrivedAtHUB="&ArrivedAtHUB&"!!!<BR>"	
	'Response.Write "DepartedHUB="&DepartedHUB&"!!!<BR>"	
	Drlname=oRs("drlname")
	Drfname=oRs("drfname")	

	''''''''END NEW VARIABLES
	DriverName=drfname&", "&drlname
	If trim(sbt_id)<>"26" then
		PODID=oRs("POD")
		ref=oRs("ref")
		'TrackingNumber=trim(oRs("trackno"))
		'Carrier=trim(oRs("carrier"))
		StatCode=oRs("statcode")
		ONBDriverID=oRs("pu_driver")
		CLSDriverID=oRs("do_driver")		
		'ETA=oRs("ETA")
		else
		StatCode=oRs("statcode")
		ONBDriverID=oRs("pu_driver")
		CLSDriverID=oRs("do_driver")
	End if

	'FromLocation=trim(oRs("from_id"))
	'ToLocation=trim(oRs("to_id"))	
	
	
	

	
	'response.write "ONBDriverID="&ONBDriverID&"<BR>"
	
	BillToID=trim(oRs("fh_bt_id"))
	MaterialType=trim(oRs("MaterialType"))
	fl_Pkey=trim(oRs("fl_Pkey"))
	fl_job_closed=trim(oRs("fl_job_closed"))
	
	If DropTime="1/1/1900" then 
		DisplayDropTime="Still In Transit"
		else
		DisplayDropTime=DropTime
	End if	
	If isdate(fl_job_closed) AND (fl_job_closed>"1/1/1900") then
		DisplayDropTime=fl_job_closed
	End if
	
	'response.write "fl_Pkey="&fl_Pkey&"<BR>"
	If trim(MaterialType)="Secure Waf" then
	'response.write "GOT HERE!!!<BR>"
		Reflist="Secure Wafer(s): "
	End if
	If trim(MaterialType)="ITAR" then
	'response.write "GOT HERE!!!<BR>"
		Reflist="ITAR(s): "
	End if    
	'response.write "MaterialType="&MaterialType&"**<BR>"
	'BillToID=Session("Suid")
	Select Case BillToID
		Case "48"
			PieceWord="HAWB #s:"
			Displaybooktime=BookTime
			DisplaybooktimeWord="Book"
			DisplayBookedWord="Booked"
		Case "72", "38", "55"
			PieceWord="Reticles:"
			Displaybooktime=BookTime
			DisplaybooktimeWord="Book"
			DisplayBookedWord="Booked"
		Case "36"
			PieceWord="Lots:"
			Displaybooktime=BookTime
			DisplaybooktimeWord="Book"
			DisplayBookedWord="Booked"
		Case "26"
			PieceWord="Documents:"
			Displaybooktime=SAPOrderTime
			DisplaybooktimeWord="SAP Order"	
			DisplayBookedWord="Booked/Picked"
		Case "75"
			PieceWord="PO Numbers:"
			Displaybooktime=BookTime
			DisplaybooktimeWord="Book"
			DisplayBookedWord="Booked"			
		Case Else
			PieceWord="Pieces:"
			Displaybooktime=BookTime
			DisplaybooktimeWord="Book"
			DisplayBookedWord="Booked"
	End select	
	'Response.Write "BookTime="&BookTime&"<BR>"
	
	'Response.Write "StatCode="&StatCode&"***<BR>"
	Select Case StatCode
		Case "0"
			StatCode="HELD"
		Case "1"
			StatCode="Scheduled"
		Case "2"
			StatCode="Booked"
		Case "3"
			StatCode="Open"
		Case "4"
			StatCode="Acknowledged by driver"
		Case "5"
			StatCode="On Board"
		Case "6"
			StatCode="Undispatched-Please report this to Mark Maggiore immediately at 214-956-0400 xt. 212"
		Case "9"
			StatCode="Closed"
		Case "10"
			StatCode="Invoiced"
		Case "13"
			StatCode="Paperwork on Board"
		Case "98"
			StatCode="<font color='red'>CANCELLED</font>"
		Case "99"
			StatCode="Deleted"
		Case "53"
			StatCode="Arrived at HUB"
		Case "54"
			StatCode="Departed HUB"
		Case "55"
			StatCode="Acknowledged by 2nd Driver"					
		Case ELSE
			StatCode="Unknown-Please report this to Mark Maggiore immediately at 214-956-0400 xt. 212"																																																																	
	End select
	
	Select Case priority
		Case "WF", "CS", "KW", "ST"
			DisplayPriority="Standard"
		Case "CE"
			DisplayPriority="Expedited"	
		Case "AS"
			DisplayPriority="Next Day"
		Case "A0"
			DisplayPriority="Hot Shot"
		Case "A1"
			DisplayPriority="Same Day"												
		Case ELSE
			DisplayPriority=Priority
	End Select
End if
Set oRs2=nothing
''''''''''''QUERY FOR DOCUMENTS/LOTS/ETC'''''''''''''''''
Set oConn2 = Server.CreateObject("ADODB.Connection")
oConn2.ConnectionTimeout = 200
oConn2.Provider = "MSDASQL"
oConn2.Open DATABASE
Err.Clear
l_cSQL2="SELECT fcrefs.rf_ref, fcrefs.pod, fcrefs.PODDateTime,  fcrefs.EDI_DateTime, fcrefs.rf_box, fcrefs.ref_Status "_ 
& " FROM  fcrefs "_  
& " WHERE (rf_fh_id= '"&OrderID&"') ORDER BY rf_ref"
'response.write "<BR><BR>****l_cSQL2="&l_cSQL2&"<BR>"					
Set oRs2 = oConn2.Execute(l_cSQL2)
	Do while not oRs2.eof
	a=a+1
	LotDocumentNumber=oRs2("rf_ref")
	PODID=oRs2("POD")
	PODDateTime=oRs2("PODDateTime")
	EDI_DateTime=oRs2("EDI_DateTime")
	Box=trim(oRs2("rf_box"))
	Ref_Status=oRs2("Ref_Status")
	Reflist=Reflist & CommaWord & LotDocumentNumber
	If trim(box)>"" then
		Reflist=Reflist & "(Box #"&Box&")&nbsp;"
	End if
	CommaWord=", "
	oRs2.movenext
	LOOP
Set oRs2=nothing
'Response.write "320 jobanalysis Reflist="&Reflist&"XX<BR>"
'If trim(sbt_id)<>"26" or trim(Reflist)>" " then
	'LengthOfReflist=Len(Reflist)-1
	'Reflist=Left(Reflist, LengthOfReflist)                                 
	'else
	Reflist=DocumentNumber
'End if

'''''''''''''QUERY FOR TO LOCATION'''''''''''''''''''''''''
Set oConn2 = Server.CreateObject("ADODB.Connection")
oConn2.ConnectionTimeout = 200
oConn2.Provider = "MSDASQL"
oConn2.Open DATABASE
Err.Clear
'response.write "334 tolocation=" & toLocation & "<br>"
l_cSQL2="SELECT CompanyName, CompanyAddress AS Address1, CompanyCity AS City, CompanyState AS State, "
l_cSQL2=l_cSQL2&" CompanyZip AS Zip FROM PreExistingCompanies " 
l_cSQL2=l_cSQL2&" WHERE (st_id= '"&ToLocation&"')"					
'response.write "338 sql=" & l_cSQL2 & "<br>" 
Set oRs2 = oConn2.Execute(l_cSQL2)
If not oRs2.eof then
ToAddress1=oRs2("Address1")
ToAddress2=""
ToCity=oRs2("City")
ToState=oRs2("State")
ToZip=oRs2("Zip")
ToCountry="USA"
ToLocation=oRs2("CompanyName")
End if
Set oRs2=nothing
'''''''''''''QUERY FOR FROM LOCATION'''''''''''''''''''''''''
Set oConn2 = Server.CreateObject("ADODB.Connection")
oConn2.ConnectionTimeout = 200
oConn2.Provider = "MSDASQL"
oConn2.Open DATABASE
Err.Clear
l_cSQL2="SELECT CompanyName, CompanyAddress AS Address1, CompanyCity AS City, CompanyState AS State, "
l_cSQL2=l_cSQL2&" CompanyZip AS Zip FROM PreExistingCompanies " 
l_cSQL2=l_cSQL2&" WHERE (st_id= '"&FromLocation&"')"					
Set oRs2 = oConn2.Execute(l_cSQL2)
If not oRs2.eof then
FromAddress1=oRs2("Address1")
FromAddress2=""
FromCity=oRs2("City")
FromState=oRs2("State")
FromZip=oRs2("Zip")
FromCountry="USA"
FromLocation=oRs2("CompanyName")
End if
Set oRs2=nothing
'''''''''''''QUERY FOR POD INFORMATION'''''''''''''''''''''''''
If BillToID="48" then

'''''''''''''QUERY FOR EXCEPTIONS INFORMATION'''''''''''''''''''''
	Set oConn2 = Server.CreateObject("ADODB.Connection")
	oConn2.ConnectionTimeout = 200
	oConn2.Provider = "MSDASQL"
	oConn2.Open DATABASE
	Err.Clear
	l_cSQL2="SELECT DriverExceptionList.ExceptionDescription, FCJobExceptions.fh_id,"
	l_cSQL2=l_cSQL2&" FCJobExceptions.Ref_num, FCJobExceptions.ExceptionTime FROM FCJobExceptions INNER JOIN"
    l_cSQL2=l_cSQL2&" DriverExceptionList ON FCJobExceptions.ExceptionID = DriverExceptionList.ExceptionID"
	l_cSQL2=l_cSQL2&" WHERE (FH_id= '"&OrderID&"')"					
	'''''Response.Write "l_cSQL2="&l_cSQL2&"<BR>"
	Set oRs2 = oConn2.Execute(l_cSQL2)
	Do while not oRs2.eof 
		Display_ExceptionDescription=oRs2("ExceptionDescription")
		Display_ERef_num=oRs2("Ref_num")
		Display_ExceptionTime=oRs2("ExceptionTime")
		'Response.Write "Display_ExceptionDescription="&Display_ExceptionDescription&"<BR>"
		'Response.Write "Display_ERef_num="&Display_ERef_num&"<BR>"
		'Response.Write "Display_ExceptionTime="&Display_ExceptionTime&"<BR>"
		Display_ExceptionList=Display_ExceptionList&"<BR>"&Display_ERef_num&" - "&Display_ExceptionTime&" - "&Display_ExceptionDescription
		NumberOfExceptions=NumberOfExceptions+1
	oRs2.movenext
	Loop
	Set oRs2=nothing	
End if
'''''''''''''QUERY FOR PICKUP DRIVER'''''''''''''''''''''''''
Set oConn2 = Server.CreateObject("ADODB.Connection")
oConn2.ConnectionTimeout = 200
oConn2.Provider = "MSDASQL"
oConn2.Open INTRANET
Err.Clear
l_cSQL2="SELECT FirstName, LastName "
l_cSQL2=l_cSQL2&" FROM Intranet_Users " 
l_cSQL2=l_cSQL2&" WHERE (Userid= '"&ONBDriverID&"')"					
Set oRs2 = oConn2.Execute(l_cSQL2)
If not oRs2.eof then
	ONBDriverName=oRs2("FirstName")&" "&oRs2("LastName")
End if
Set oRs2=nothing
'Response.Write "l_cSQL2="&l_cSQL2&"<BR>"
'''''''''''''QUERY FOR DROPOFF DRIVER'''''''''''''''''''''''''
Set oConn2 = Server.CreateObject("ADODB.Connection")
oConn2.ConnectionTimeout = 200
oConn2.Provider = "MSDASQL"
oConn2.Open INTRANET
Err.Clear
l_cSQL2="SELECT FirstName, LastName "
l_cSQL2=l_cSQL2&" FROM Intranet_Users " 
l_cSQL2=l_cSQL2&" WHERE (Userid= '"&CLSDriverID&"')"					
Set oRs2 = oConn2.Execute(l_cSQL2)
If not oRs2.eof then
	CLSDriverName=oRs2("FirstName")&" "&oRs2("LastName")
End if
Set oRs2=nothing

If trim(ONBDriverName)="" then ONBDriverName="n/a" end if
If trim(CLSDriverName)="" then CLSDriverName="n/a" end if
%>



</head>

<body onload="document.FindUser.requestorName.focus();">
	<table align="center" border="0" bordercolor="black" cellpadding="0" cellspacing="0" ID="Table1" height="100%" width="100%">
        <tr><td align="left"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>
        <tr>
            <td align="left"><img src="../images/FleetX_Small.jpg" height="50" width="168" /></td>
            <td align="right" valign="bottom"><!-- #include file="../topnavbar.asp" --></td>
        </tr>
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" /></td></tr>	
    <tr><td colspan="2">
<form action="NewUser.asp" method="post" name="FindUser">
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



    <tr><td align=center width="100%"><!-- main page stuff goes here! -->
    
    
<table width="700" cellpadding="2" cellspacing="0" border="1" align="center" ID="Table1">
	<tr>
		<td colspan="4">
			<table width="100%" ID="Table2">
				<tr>
					<td width="33%">
						<img src="../images/logisticorpdetails.jpg" height="27" width="93">
					</td>
					<td width="34%"  class="LargeHeaderCentered">
						Delivery Details
					</td>
					<td width="33%" align="right" valign="top" Class="DetailsTitlesRight"><%=Session("txt_cm_desc")%></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td colspan="4" class="MidHeaderCenteredBlack" bgcolor=#ECE9D8>
			Shipment Information
		</td>
	</tr>	
	<tr>
		<td width="25%" Class="DetailsTitles">Job Number</td>
		<td width="25%" Class="DetailsDetails"><%=OrderID%></td>
		<td width="25%" Class="DetailsTitles">Current Status</td>
		<td width="25%" Class="DetailsDetails"><%=StatCode%></td>		
	</tr>
	<tr>
		<td width="25%" Class="DetailsTitles">Submitted By</td>
		<td width="25%" Class="DetailsDetails"><%=SubmittedBy%></td>	
		<td width="25%" Class="DetailsTitles">Priority</td>
		<td width="25%" Class="DetailsDetails"><%=DisplayPriority%></td>
	</tr>
	<tr>
		<td colspan="4"><span Class="DetailsTitles"><%=PieceWord%>&nbsp;&nbsp;</span><span Class="DetailsDetails"><%=RefList%></span></td>
	</tr>
	<%
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'Response.Write "BillToID="&BillToID&"<BR>"
	if trim(BillToID)="75" then
		'Response.Write "GOT HERE!<BR>"
		Set oConn2 = Server.CreateObject("ADODB.Connection")
		oConn2.ConnectionTimeout = 200
		oConn2.Provider = "MSDASQL"
		oConn2.Open DATABASE
		Err.Clear
		l_cSQL2="SELECT * FROM FCRefs_Details "
		l_cSQL2=l_cSQL2&" WHERE (fh_id= '"&OrderID&"')"					
		Set oRs2 = oConn2.Execute(l_cSQL2)
		If not oRs2.eof then
			Response.Write "<tr><td colspan='4'><span Class='DetailsTitles'>Shipment Details</span><br>"
		end if
		Do While not oRs2.eof 
			FC_Description=oRs2("FC_Description")
			Pieces=oRs2("Pieces")
			PieceType=oRs2("PieceType")
			Skids=oRs2("Skids")
			ShowWeight=oRs2("Weight")
			DimWeight=oRs2("DimWeight")
			Dimensions=oRs2("Dimensions")
			%>
			
			
				<span Class="DetailsDetails">&nbsp;&nbsp;&nbsp;&nbsp;<%=FC_Description%>&nbsp;&nbsp;
				<%=PieceType%>:<%=Pieces%>&nbsp;&nbsp;
				Skids:<%=Skids%>&nbsp;&nbsp;
				Weight:<%=ShowWeight%>&nbsp;&nbsp;
				Dim Weight:<%=DimWeight%>&nbsp;&nbsp;
				Dimensions:<%=Dimensions%>&nbsp;&nbsp;<br>
				</span>
				
			<%
		oRs2.Movenext
		LOOP
		Response.Write "</td></tr>"
		Set oRs2=nothing			
	
	End if
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	%>
	<tr>
		<td class="MidHeaderLeftBlack" bgcolor=#ECE9D8 colspan="2">
			Pickup
		</td>
		<td class="MidHeaderLeftBlack" bgcolor=#ECE9D8 colspan="2">
			Delivery
		</td>
	</tr>
		
	<tr>
		<td class="DetailsDetails" colspan="2">
			<span class="DetailsTitles"><%=DisplayBookTimeWord%> Time:  </span><%=Displaybooktime%>
		</td>
		<td class="DetailsDetails" colspan="2">
			<span class="DetailsTitles">Due Time:  </span><%=DueTime%>
		</td>
	</tr>
	<%
	'response.Write "BillToID="&BillToID&"<BR>"
	If trim(FromLocation)="55" then FromLocation="Compugraphics" end if
	If trim(FromLocation)="72" then FromLocation="CRI" end if
	If trim(ToLocation)="GPGP" then ToLocation="Compugraphics" end if
	If trim(ToLocation)="13601" then ToLocation="TI American Regional PDC" end if
	%>
	<tr>
		<td class="DetailsDetails" colspan="2" valign="top">
			<%=FromLocation%><br>
			<%
			if trim(FromAddress1)>"" then
				Response.Write FromAddress1&"<BR>"
			End if			
			if trim(FromAddress2)>"" then
				Response.Write FromAddress2&"<BR>"
			End if
			%>
			<%=FromCity%>, <%=FromState%>&nbsp;&nbsp;<%=FromZip%><br>
			<%=FromCountry%>&nbsp;&nbsp;
		</td>
		<td class="DetailsDetails" colspan="2" valign="top">
			<%=toLocation%><br>
			<%
			if trim(toAddress1)>"" then
				Response.Write toAddress1&"<BR>"
			End if			
			if trim(toAddress2)>"" then
				Response.Write toAddress2&"<BR>"
			End if
			%>
			<%=toCity%>, <%=toState%>&nbsp;&nbsp;<%=toZip%><br>
			<%=toCountry%>&nbsp;&nbsp;
		</td>
	</tr>
	<tr>
		<td class="DetailsDetails" colspan="2" valign="top">
			<span class="DetailsTitles">Comments:  </span><%=DisplayFromComments%>
		</td>
		<td class="DetailsDetails" colspan="2" valign="top">
			<span class="DetailsTitles">Comments:  </span><%=DisplayToComments%>
		</td>
	</tr>	
	<tr>
		<td colspan="4" class="MidHeaderCenteredBlack" bgcolor=#ECE9D8>
			Delivery Information
		</td>
	</tr>
	<tr>
		<td class="DetailsDetails" colspan="2" valign="top">
			<span class="DetailsTitles">Pickup Time:</span>&nbsp;&nbsp;<%=DisplayOnBoardTime%>
		</td>
		<td class="DetailsDetails" colspan="2" valign="top">
			<span class="DetailsTitles">Delivery Time:</span>&nbsp;&nbsp;<%=DisplayDropTime%>
		</td>
	</tr>
	<tr>
		<td class="DetailsDetails" colspan="2" valign="top">
			<span class="DetailsTitles">Pickup Driver:</span>&nbsp;&nbsp;<%=ONBDriverName%>
		</td>
		<td class="DetailsDetails" colspan="2" valign="top">
			<span class="DetailsTitles">Delivery Driver:</span>&nbsp;&nbsp;<%=CLSDriverName%>
		</td>
	</tr>
	<%
	'Response.write "trackingnumber="&trackingnumber&"<BR>"
	'Response.write "FromLocation="&FromLocation&"<BR>"
	If trim(trackingnumber)>"" and (trim(FromLocation)="Compugraphics" OR trim(FromLocation)="TOPPAN") then
	%>
	<form method="post" action="http://my.shipgreyhound.com/cfw/trackOrder.login" target="_blank" ID="Form1">
	<tr>
	<input type="hidden" name="orderNumber" value="<%=trackingnumber%>" ID="Hidden1">
		<td nowrap align="right" class="DetailsDetails" colspan="2" valign="bottom"><span class="DetailsTitles">Greyhound Tracking: </span>
		<input TYPE="IMAGE" SRC="../images/btnClickHereLink.gif" ALT="click here" ID="Image1" NAME="Image1"></td>						
		<td nowrap align="right" class="DetailsDetails" colspan="2"><span class="DetailsTitles">Original Bus ETA: </span>
		<%=ETA%></td>						
	</tr>
	</form>							
	<%
	End if	
	
	If DocumentNumber>"" and (FromLocation="Compugraphics" or ToLocation="CPGP" or ToLocation="TOPPAN" or FromLocation="TOPPAN") AND trim(DocumentNumber)>"" then
		%>
		<tr>
			<td class="DetailsDetails" colspan="4" valign="top">
				<span class="DetailsTitles">Quick Tracking:</span>&nbsp;&nbsp;<a href="http://www.quickonline.com/cgi-bin/WebObjects/BOLSearch?bolNumber=<%=DocumentNumber%>" target="_blank">click here</a>
			</td>
		</tr>		
		<%
	End if		
	
	
	
		Set oConn2 = Server.CreateObject("ADODB.Connection")
		oConn2.ConnectionTimeout = 200
		oConn2.Provider = "MSDASQL"
		oConn2.Open DATABASE
		Err.Clear
		l_cSQL2="SELECT fcrefs.rf_ref, fcrefs.pupod, fcrefs.pod, fcrefs.pod2, fcrefs.PODDateTime, fcrefs.EDI_DateTime, fcrefs.ref_Status "_ 
		& " FROM  fcrefs "_  
		& " WHERE (rf_fh_id= '"&OrderID&"') ORDER BY rf_ref"					
		'response.write "l_cSQL2="&l_cSQL2&"<BR>"
		Set oRs2 = oConn2.Execute(l_cSQL2)
			Do while not oRs2.eof
			a=a+1
			LotDocumentNumber=oRs2("rf_ref")
			PUPODID=oRs2("PUPOD")
            'response.Write "PUPODID="&PUPODID&"<BR>"
            PODID=oRs2("POD")
			PODID2=oRs2("POD2")
			PODDateTime=oRs2("PODDateTime")
			EDI_DateTime=oRs2("EDI_DateTime")
			If not isdate(EDI_DateTime) then EDI_DateTime="n/a" End if
'''''''''''''''''''''''''''''''''''''''''''''''''''
			Set oConn62 = Server.CreateObject("ADODB.Connection")
			oConn62.ConnectionTimeout = 200
			oConn62.Provider = "MSDASQL"
			oConn62.Open DATABASE
			Err.Clear
			l_cSQL62="SELECT Signature "
			l_cSQL62=l_cSQL62&" FROM PODLIST " 
			l_cSQL62=l_cSQL62&" WHERE (PODid= '"&PODID&"') or (PODid='"&PODID2&"')"					
			Set oRs62 = oConn62.Execute(l_cSQL62)
			Do while not oRs62.eof
				zzzz=zzzz+1
				Signature=oRs62("Signature")

				if xzzzz>1 then
					DisplaySignature=DisplaySignature&", "&Signature
					else
					DisplaySignature=Signature
				End if
				'response.write "Signature="&Signature&"<BR>"
			oRs62.movenext
			LOOP
			Set oRs62=nothing
			'response.write "l_cSQL62="&l_cSQL62&"<BR>"
			
						'Ref_Status=oRs2("Ref_Status")
			'Reflist=Reflist & CommaWord & LotDocumentNumber
			'CommaWord=", "
				''''''''''''''TEMP SIGNATURE''DELETE''''''''''
				'Signature="TEMP SIGNATURE"
				'DisplaySignature="TEMP SIGNATURE"
				''''''''''''''''''''''''''''''''''''''''
			'If trim(signature)="" then
				'Set RSEVENTS22 = Server.CreateObject("ADODB.Recordset")
					'RSEVENTS22.CursorLocation = 3
					'RSEVENTS22.CursorType = 3
					'response.Write "Liberty="&Liberty&"<BR>"
					'RSEVENTS22.ActiveConnection = LIBERTY
					'l_csql = "SELECT * FROM F_HAWB_DATA WHERE (SZF1='"&LotDocumentNumber&"')"
					'Response.write("Query:" & l_cSQL)
					'RSEVENTS22.Open l_cSQL, LIBERTY, 1, 3
					'If not RSEVENTS22.EOF then	
					'Signature="n/a"
					'DisplaySignature="n/a"
					'End if
					'RSEVENTS22.close
				'Set RSEVENTS22 = Nothing								
			'end if					
			if trim(Signature)>"" or materialtype="ITAR" then
				%>
					<tr>
					<%if trim(BillToID)="48" then%>
						<td class="DetailsDetails" colspan="2" valign="top">
							<span class="DetailsTitles">POD EDI:</span>&nbsp;&nbsp;<%=EDI_DateTime%>
						</td>
				<%
					else
                    If trim(PUPODID)>"" then
              			Set oConn62 = Server.CreateObject("ADODB.Connection")
            			oConn62.ConnectionTimeout = 200
            			oConn62.Provider = "MSDASQL"
            			oConn62.Open DATABASE
            			Err.Clear
            			l_cSQL62="SELECT Signature "
            			l_cSQL62=l_cSQL62&" FROM PODLIST " 
            			l_cSQL62=l_cSQL62&" WHERE (PODid= '"&PUPODID&"')"					
            			Set oRs62 = oConn62.Execute(l_cSQL62)
            			Do while not oRs62.eof
            				zzzz=zzzz+1
            				PUSignature=oRs62("Signature")

            				if xzzzz>1 then
            					PUDisplaySignature=PUDisplaySignature&", "&PUSignature
            					else
            					PUDisplaySignature=PUSignature
            				End if
            				'response.write "Signature="&Signature&"<BR>"
            			oRs62.movenext
            			LOOP
            			Set oRs62=nothing                  
                        %>
 						<td class="DetailsDetails" colspan="2" valign="top">
							<span class="DetailsTitles">Proof of Pickup:</span>&nbsp;&nbsp;(<%=LotDocumentNumber%>)&nbsp;&nbsp;<%=PUDisplaySignature%>
						</td> 
                        <%                  
                    else
					%>				
						<td class="DetailsDetails" colspan="2" valign="top">
							&nbsp;&nbsp;
						</td>
					<%
					end if
                    end if
					%>
						<td class="DetailsDetails" colspan="2" valign="top">
						<%If trim(BILLTOID)="48" then
						
							Set RSEVENTS22 = Server.CreateObject("ADODB.Recordset")
								RSEVENTS22.CursorLocation = 3
								RSEVENTS22.CursorType = 3
								'response.Write "Liberty="&Liberty&"<BR>"
								RSEVENTS22.ActiveConnection = LIBERTY
								l_csql = "SELECT * FROM F_HAWB_DATA WHERE (SZF1='"&LotDocumentNumber&"')"
								'Response.write("Query:" & l_cSQL)
								RSEVENTS22.Open l_cSQL, LIBERTY, 1, 3
								If RSEVENTS22.EOF then
								   ' Response.Write "DisplaySignature="&DisplaySignature&"<BR>"
								    If trim(DisplaySignature)="" then
                                        DisplaySignature="n/a"   
                                        else
                                     End if
                                End if
                                If not RSEVENTS22.EOF then
									ULID=RSEVENTS22("ULID")
									HexULID=Hex(ULID)
									'Response.Write "HEXULID="& HEXULID &"***<BR>"
									If trim(DisplaySignature)="" then DisplaySignature="n/a" end if
									%>
									
									<span class="DetailsTitles">POD:</span>&nbsp;&nbsp;(<%=LotDocumentNumber%>)&nbsp;&nbsp;<a href="http://document.logisticorp.us:8080/LibertyIMS::/User=WebUser;pwd=Internet42;sys=LogistiCorp/Cmd%3DGetRawDocument%3BFolder%3D%2321%3BDoc%3D<%=HexULID%>%3Bformat%3DLIC/" target="_blank"><%=DisplaySignature%></a>&nbsp;
									<%
									else
									ULID=""
									If isdate(PODDateTime) then
										%>
										<span class="DetailsTitles">POD:</span>&nbsp;&nbsp;(<%=LotDocumentNumber%>)&nbsp;&nbsp;<a href="../KWEPODS/<%=trim(LotDocumentNumber)%>.pdf" target="_blank"><%=DisplaySignature%></a>&nbsp;
										<%
										Else
										%>									
									
									
									<span class="DetailsTitles">POD:</span>&nbsp;&nbsp;(<%=LotDocumentNumber%>)&nbsp;&nbsp;<%=DisplaySignature%>&nbsp;
									<%
									End if
								End if
								RSEVENTS22.close
							Set RSEVENTS22 = Nothing						
						
						%>
						<!--
							<span class="DetailsTitles">POD:</span>&nbsp;&nbsp;(<%=LotDocumentNumber%>)&nbsp;&nbsp;<a href="http://192.168.104.231:8080/LibertyIMS::/User=WebUser;pwd=Internet42;sys=LogistiCorp/Cmd%3DGetRawDocument%3BFolder%3D%2321%3BDoc%3D<%=HexULID%>%3Bformat%3DLIC/" target="_blank">xxx<%=DisplaySignature%>xxx</a>&nbsp;
						-->	
							
							<%else
							'If isdate(PODDateTime) then
								%>
								<!--
								<span class="DetailsTitles">POD:</span>&nbsp;&nbsp;(<%=LotDocumentNumber%>)&nbsp;&nbsp;<a href="../KWEPODS/<%=trim(LotDocumentNumber)%>.pdf" target="_blank"><%=DisplaySignature%></a>&nbsp;
								-->
								<%
								'Else
                                If Trim(DisplaySignature)="" then DisplaySignature="n/a" end if
								%>
								
								<span class="DetailsTitles">POD:</span>
                                <%If DisplaySignature<>"n/a" then %>
                                     &nbsp;&nbsp;(<%=LotDocumentNumber%>)
                                <%End if%>
                                &nbsp;&nbsp;<%=DisplaySignature%>&nbsp;
								<%
							'End if
						End if
						DisplaySignature=""
						%>	
						</td>
					</tr>
			<%
			end if
			oRs2.movenext
			LOOP
		Set oRs2=nothing	
	%>
	<%if trim(Display_ExceptionList)>"" then%>
		<tr>
			<td class="DetailsDetails" colspan="4" valign="top">
				<span class="DetailsTitles">Exception<%IF NumberOfExceptions>1 then response.Write "s" end if%>:</span>&nbsp;&nbsp;<%=Display_ExceptionList%>
			</td>
		</tr>
	<%end if%>	

	<tr>
		<td colspan="4" class="MidHeaderCenteredBlack" bgcolor=#ECE9D8>
			Status History
		</td>
	</tr>		
	<tr>
		<td class="DetailsTitles" colspan="2" align="left">
			Milestones
		</td>
		<td class="DetailsTitles" colspan="2" align="left">
			Times
		</td>
	</tr>
<%
	'Response.Write "BillToID="&BillToID&"<br>"
	if SAPOrdertime<>"1/1/1900" and trim(BillToID)="26" then%>	
	<tr>
		<td class="DetailsDetails" colspan="2" align="left">
			SAP Order Time
		</td>
		<td class="DetailsDetails" colspan="2" align="left">
			<%=SAPOrdertime%>
		</td>
	</tr>
	<%end if%>
  <% 'response.write "876 Booktime=" & BookTime & ",SAPOrderTime=" & SAPOrderTime & "<br>" %>
	<%if BookTime<>"1/1/1900" and not isNull(SAPOrderTime) then
		ElapsedTime=((cDate(BookTime)-cDate(SAPOrderTime))*24)*60
		Hours = Int (ElapsedTime / 60)	
		Minutes = ElapsedTime - (Hours * 60)
		If Minutes>0 then 
			Minutes=cInt(Minutes)
			else
			Minutes=0
		End if
		If BillToID<>"48" and ToLocation<>"CPGP" and ToLocation<>"TOPPAN" then
			If (hours>=0 AND minutes>=0) AND (hours>0 or minutes>0) then
				DisplaySAPOrderTime=" ("&Hours&" hrs "&Minutes&" mins)"	
				else
				DisplaySAPOrderTime=""
			End if
		End if
	End if
	If BillToID="75" or BillToID="81" then
		DisplaySAPOrderTime=""
	end if
		%>		
	<%if booktime<>"1/1/1900" then%>	
	<tr>
		<td class="DetailsDetails" colspan="2" align="left">
			<%=DisplayBookedWord%>
		</td>
		<td class="DetailsDetails" colspan="2" align="left">
			<%=booktime%><%=DisplaySAPOrderTime%>
		</td>
	</tr>
	<%end if%>
	<%
	'Response.Write "ReadyTime="&ReadyTime&"***<BR>"
	if SAPOrderTime<>"1/1/1900" and BillToID="48" then%>	
	<tr>
		<td class="DetailsDetails" colspan="2" align="left">
			Ready Time
		</td>
		<td class="DetailsDetails" colspan="2" align="left">
			<%=SAPOrderTime%>
		</td>
	</tr>
	<%end if%>	
	<%
	'Response.Write "BillToID="&BillToID&"***<BR>"
	'Response.Write "FromLocation="&FromLocation&"***<BR>"
	'Response.Write "ArrivedAtHUB="&ArrivedAtHUB&"***<BR>"
	'Response.Write "DepartedHUB="&DepartedHUB&"***<BR>"
	if DispatchTime<>"1/1/1900" And (FromLocation<>"Compugraphics" AND FromLocation<>"TOPPAN") then
		ElapsedTime=((cDate(DispatchTime)-cDate(booktime))*24)*60
		Hours = Int (ElapsedTime / 60)	
		Minutes = ElapsedTime - (Hours * 60)
		If Minutes>0 then 
			Minutes=cInt(Minutes)
			else
			Minutes=0
		End if	
		%>	
		<tr>
			<td class="DetailsDetails" colspan="2" align="left">
				Dispatched
			</td>
			<td class="DetailsDetails" colspan="2" align="left">
				<%=DispatchTime%> (<%=Hours%> hrs <%=Minutes%> mins)
			</td>
		</tr>
	<%end if%>	
	
	

	
	
	
	
	<%if DriverAcknowledgementTime<>"1/1/1900" and (FromLocation<>"TOPPAN") then
		If FromLocation="CRI" then
			DriverAcknowledgedWord="CRI Acknowledged"
			DispatchTime=BookTime
			OnBoardTime=DriverAcknowledgementTime
			else
			DriverAcknowledgedWord="Driver Acknowledged"
		End if
		If BillToID="75" then
			DriverAcknowledgedWord="Driver Acknowledged"
			DispatchTime=BookTime
		End if
		ElapsedTime=((cDate(DriverAcknowledgementTime)-cDate(DispatchTime))*24)*60
		Hours = Int (ElapsedTime / 60)	
		Minutes = ElapsedTime - (Hours * 60)
		If Minutes>0 then 
			Minutes=cInt(Minutes)
			else
			Minutes=0
		End if		
		%>	
		<tr>
			<td class="DetailsDetails" colspan="2" align="left">
				<%=DriverAcknowledgedWord%>
			</td>
			<td class="DetailsDetails" colspan="2" align="left">
				<%=DriverAcknowledgementTime%> (<%=Hours%> hrs <%=Minutes%> mins)
			</td>
		</tr>
	<%end if%>
	
	
	
	<%
	'Response.Write "FromLocation="&FromLocation&"<BR>"
	if OnBoardTime<>"1/1/1900" and ((FromLocation="Compugraphics") OR (FromLocation="TOPPAN")) then
		ElapsedTime=((cDate(OnBoardTime)-cDate(DriverAcknowledgementTime))*24)*60
		Hours = Int (ElapsedTime / 60)	
		Minutes = ElapsedTime - (Hours * 60)
		If Minutes>0 then 
			Minutes=cInt(Minutes)
			else
			Minutes=0
		End if
		Select Case FromLocation
			Case "TOPPAN", "Compugraphics"
				DisplayOnBoardSection="On Board"
			Case else
				'DisplayOnBoardSection="Acknowledged by Second Driver"
				DisplayOnBoardSection="Driver Acknowledged"
		End Select
		%>			
		<tr>
			<td class="DetailsDetails" colspan="2" align="left">
				<%=DisplayOnBoardSection%>
			</td>
			<td class="DetailsDetails" colspan="2" align="left">
				<%=OnBoardTime%> (<%=Hours%> hrs <%=Minutes%> mins)
			</td>
		</tr>
	<%end if%>	
		
	<%
	'Response.write "GOT HERE1!!!!<BR>"
	if ((FromLocation="Compugraphics") or (FromLocation="CRI") or (FromLocation="TOPPAN"))  and (trim(ArrivedAtHUB) > "")  then
		'Response.write "GOT HERE2!!!!<BR>"
		ElapsedTime=((cDate(ArrivedAtHUB)-cDate(OnBoardTime))*24)*60
		Hours = Int (ElapsedTime / 60)	
		Minutes = ElapsedTime - (Hours * 60)
		If Minutes>0 then 
			Minutes=cInt(Minutes)
			else
			Minutes=0
		End if	
		%>	
		<tr>
			<td class="DetailsDetails" colspan="2" align="left">
				Arrived at HUB
			</td>
			<td class="DetailsDetails" colspan="2" align="left">
				<%=ArrivedAtHUB%> (<%=Hours%> hrs <%=Minutes%> mins)
			</td>
		</tr>
	<%end if%>	
	<%if trim(AcknowledgedHUB)>"" and ((FromLocation="Compugraphics") or (FromLocation="CRI") or (FromLocation="TOPPAN")) then
		ElapsedTime=((cDate(AcknowledgedHUB)-cDate(ArrivedAtHUB))*24)*60
		Hours = Int (ElapsedTime / 60)	
		Minutes = ElapsedTime - (Hours * 60)
		If Minutes>0 then 
			Minutes=cInt(Minutes)
			else
			Minutes=0
		End if		
		%>	
		<tr>
			<td class="DetailsDetails" colspan="2" align="left">
				Driver Acknowledged
			</td>
			<td class="DetailsDetails" colspan="2" align="left">
				<%=AcknowledgedHUB%> (<%=Hours%> hrs <%=Minutes%> mins)
			</td>
		</tr>
	<%end if%>	
	
	
		
	<%
	'Response.write "GOT HERE1!!!!<BR>"
	if ((FromLocation="Compugraphics") or (FromLocation="CRI") or (FromLocation="TOPPAN")) and (trim(DepartedHUB) > "") and (trim(AcknowledgedHUB) > "")  then
		DontShow="y"
		'Response.write "GOT HERE2!!!!<BR>"
		'Response.Write "DepartedHUB="&DepartedHUB&"***<BR>"
		'Response.Write "AcknowledgedHUB="&AcknowledgedHUB&"***<BR>"
		ElapsedTime=((cDate(DepartedHUB)-cDate(AcknowledgedHUB))*24)*60
		Hours = Int (ElapsedTime / 60)	
		Minutes = ElapsedTime - (Hours * 60)
		If Minutes>0 then 
			Minutes=cInt(Minutes)
			else
			Minutes=0
		End if	
		%>	
		<tr>
			<td class="DetailsDetails" colspan="2" align="left">
				Departed HUB
			</td>
			<td class="DetailsDetails" colspan="2" align="left">
				<%=DepartedHUB%> (<%=Hours%> hrs <%=Minutes%> mins)
			</td>
		</tr>
	<%
	OnBoardTime=DepartedHUB
	If (FromLocation="CRI") then
		OnBoardTime=DepartedHUB
	End if
	end if%>		
	
	
	
	
	<%if PaperWorkTime<>"1/1/1900" and BillToID="48" then
		ElapsedTime=((cDate(PaperWorkTime)-cDate(DriverAcknowledgementTime))*24)*60
		Hours = Int (ElapsedTime / 60)	
		Minutes = ElapsedTime - (Hours * 60)
		If Minutes>0 then 
			Minutes=cInt(Minutes)
			else
			Minutes=0
		End if		
		%>	
		<tr>
			<td class="DetailsDetails" colspan="2" align="left">
				Paper on Board
			</td>
			<td class="DetailsDetails" colspan="2" align="left">
				<%=PaperWorkTime%> (<%=Hours%> hrs <%=Minutes%> mins)
			</td>
		</tr>
	<%end if%>	
	<%
	''Response.Write "AtAirlineTime="&AtAirlineTime&"***<BR>"
	''Response.Write "PaperWorkTime="&PaperWorkTime&"***<BR>"
	if AtAirlineTime<>"1/1/1900" and BillToID="48" then
		ElapsedTime=((cDate(AtAirlineTime)-cDate(PaperWorkTime))*24)*60
		Hours = Int (ElapsedTime / 60)	
		Minutes = ElapsedTime - (Hours * 60)
		If Minutes>0 then 
			Minutes=cInt(Minutes)
			else
			Minutes=0
		End if		
		%>	
		<tr>
			<td class="DetailsDetails" colspan="2" align="left">
				At Airline
			</td>
			<td class="DetailsDetails" colspan="2" align="left">
				<%=AtAirlineTime%> (<%=Hours%> hrs <%=Minutes%> mins)
			</td>
		</tr>
	<%
	DriverAcknowledgementTime=AtAirlineTime
	end if%>		
	
	
	
	
	
	<%if OnBoardTime<>"1/1/1900" and (FromLocation<>"Compugraphics") and (FromLocation<>"CRI") and (FromLocation<>"TOPPAN") then
		ElapsedTime=((cDate(OnBoardTime)-cDate(DriverAcknowledgementTime))*24)*60
		Hours = Int (ElapsedTime / 60)	
		Minutes = ElapsedTime - (Hours * 60)
		If Minutes>0 then 
			Minutes=cInt(Minutes)
			else
			Minutes=0
		End if
		%>			
		<tr>
			<td class="DetailsDetails" colspan="2" align="left">
				On Board
			</td>
			<td class="DetailsDetails" colspan="2" align="left">
				<%=OnBoardTime%> (<%=Hours%> hrs <%=Minutes%> mins)
			</td>
		</tr>
	<%end if

	'Response.Write "DontShow="&DontShow&"<BR>"
	'Response.write "GOT HERE1!!!!<BR>"
	'if ((ToLocation="CPGP") OR (ToLocation="TOPPAN") OR (FromLocation="DSTK"))  and (trim(ArrivedAtHUB) > "")  then
	if (trim(ArrivedAtHUB) > "") and DontShow<>"y" and (FromLocation<>"CRI")  then
		'response.write "GOT HERE2!!!!<BR>"
		ElapsedTime=((cDate(ArrivedAtHUB)-cDate(OnBoardTime))*24)*60
		Hours = Int (ElapsedTime / 60)	
		Minutes = ElapsedTime - (Hours * 60)
		If Minutes>0 then 
			Minutes=cInt(Minutes)
			else
			Minutes=0
		End if	
		%>	
		<tr>
			<td class="DetailsDetails" colspan="2" align="left">
				Arrived at HUB
			</td>
			<td class="DetailsDetails" colspan="2" align="left">
				<%=ArrivedAtHUB%> (<%=Hours%> hrs <%=Minutes%> mins)
			</td>
		</tr>
	<%
		OnBoardTime=ArrivedAtHUB
		
		
		'Response.Write "AcknowledgedHUB="&AcknowledgedHUB&"****<BR>"
	If trim(BillToID)="48" then
		'Response.Write "WHATEVER!!!<br>"
			'''''''''''''QUERY STATEMENT'''''''''''''''''''''''''''''''''''''''''''
			Set oConn = Server.CreateObject("ADODB.Connection")
			oConn.ConnectionTimeout = 200
			oConn.Provider = "MSDASQL"
			oConn.Open DATABASE
			'response.write "Database="&Database&"<br>"
			'response.write "billtoid="&billtoid&"<BR>"
			'response.write "sbt_id="&sbt_id&"<BR>"
			
			l_cSQL48=l_cSQL48&"Select * from marksview2 "
			if trim(sbt_id)="26" then
				l_cSQL48="Select * from marksview2_sr "
			End if
			l_cSQL48=l_cSQL48&" WHERE (fl_pkey>'"&fl_Pkey&"') "
			If InputJobNumber>"" then     
				'l_cSQL48=l_cSQL48&" AND (jobnum like '%"&InputJobNumber&"')"
				l_cSQL48=l_cSQL48&" AND (jobnum = '"&InputJobNumber&"')"
			End if		 
			If InputDocumentNumber>"" then     
				l_cSQL48=l_cSQL48&" AND (custpo like '%"&InputDocumentNumber&"')"
			End if
			If InputLotNumber>"" then     
				l_cSQL48=l_cSQL48&" AND (ref like '%"&InputLotNumber&"')"
			End if		
			l_cSQL48=l_cSQL48&" Order by fl_Pkey " 
			'response.write "1212 jobanalysis l_cSQL48="&l_cSQL48&"<BR>"
			'response.write "!!!Database="&Database&"<BR>"
			'response.write "!!!BillToID="&BillToID&"<BR>"

			'response.Write "name="&Session("txt_cm_desc")&"<BR>"
			Set oRs48 = oConn.Execute(l_cSQL48)
			If oRs48.eof then
				ErrorMessage="There are no orders that match your criteria"
			end if
			If Err.Number <> 0 Then                                               
			Response.Write ErrorMessage="Error Executing the query.  Error:" & Err.Description
			End if
			Do while NOT oRs48.EOF 
				
				
				'NL_Onb=oRs48("onbtime")
				FinalClose=oRs48("DropTime")
				NL_Cls=oRs48("at_hub")
				NL_ONB=oRs48("ONBLeg2")
				NL_Acc=oRs48("accleg2")
				''''ArrivedAtHUB=NL_Cls
				
				
				NL_fl_Pkey=oRs48("fl_Pkey")
				'''''''''''''MARK HERE IS WHERE YOU LEFT OFF!!!!''''''''''''''''''
				'response.write "NL_fl_Pkey="&NL_fl_Pkey&"<BR>"
				'response.write "NL_Acc="&NL_Acc&"-"
				'response.write "ArrivedAtHUB="&ArrivedAtHUB&"<BR>"
				'response.write "XXXXXNL_Cls="&NL_Cls&"***<BR>"
				
				
				'response.write "NL_fl_Pkey="&NL_fl_Pkey&"<BR>"
				'response.write "NL_Acc="& NL_Acc &"<BR>"
				'response.write "NL_Onb="&NL_Onb&"<BR>"
				'response.write "NL_Cls="&NL_Cls&"<BR>"
				If NL_Acc>" " then
						'Response.Write "GOT HERE!!!!<BR>"
						'Response.Write "****TempNL_Cls="&TempNL_Cls&"****<BR>"
						'Response.Write "****NL_Cls="&NL_Cls&"****<BR>"
						If cDate(TempNL_Cls)>cDate(ArrivedAtHUB) then
							ArrivedAtHUB=TempNL_Cls
							'Response.Write "GOT HERE!!!<BR>"
						End if
						
						ElapsedTime=((cDate(NL_Acc)-cDate(ArrivedAtHUB))*24)*60
						Hours = Int (ElapsedTime / 60)	
						Minutes = ElapsedTime - (Hours * 60)
						If Minutes>0 then 
							Minutes=cInt(Minutes)
							else
							Minutes=0
						End if
					%>
					<tr>
						<td class="DetailsDetails" colspan="2" align="left">
							Driver Acknowledged
						</td>
						<td class="DetailsDetails" colspan="2" align="left">
							<%=NL_Acc%> (<%=Hours%> hrs <%=Minutes%> mins)
						</td>
					</tr>
					<%
				End if
				If NL_Onb>"1/1/900" then
				'response.write "NL_Onb="&NL_Onb&"***<BR>"
				'response.write "NL_Acc="&NL_Acc&"***<BR>"
						ElapsedTime=((cDate(NL_Onb)-cDate(NL_Acc))*24)*60
						'response.Write "ElapsedTime="&ElapsedTime&"<BR>"
						Hours = Int (ElapsedTime / 60)	
						Minutes = ElapsedTime - (Hours * 60)
						If Minutes>0 then 
							Minutes=cInt(Minutes)
							else
							Minutes=0
						End if				
					%>
					<tr>
						<td class="DetailsDetails" colspan="2" align="left">
							Departed HUB
						</td>
						<td class="DetailsDetails" colspan="2" align="left">
							<%=NL_Onb%> (<%=Hours%> hrs <%=Minutes%> mins)
						</td>
					</tr>
					<%
				End if
				 if ors48.eof then
					Hubword="Destination"
					else
					Hubword="HUB"
				end if
				if NL_CLS>"1/1/1900" then
						ElapsedTime=((cDate(NL_CLS)-cDate(NL_Onb))*24)*60
						'response.Write "ElapsedTime="&ElapsedTime&"<BR>"
						TempNL_Cls=NL_Cls
						Hours = Int (ElapsedTime / 60)	
						Minutes = ElapsedTime - (Hours * 60)
						If Minutes>0 then 
							Minutes=cInt(Minutes)
							else
							Minutes=0
						End if					
				%>
				<tr>
					<td class="DetailsDetails" colspan="2" align="left">
						Arrived at <%=HUBWORD%>
					</td>
					<td class="DetailsDetails" colspan="2" align="left">
						<%=NL_Cls%> (<%=Hours%> hrs <%=Minutes%> mins)
					</td>
				</tr>								
				<%
				End if
			oRs48.Movenext
			LOOP
			oRs48.Close
			Set oRs48=Nothing
			'TempNL_Cls=NL_Cls
				If trim(FinalClose)>"" AND trim(NL_Onb)>"" then	
				'response.write "XXXNL_Onb="&NL_Onb&"<BR>"
						ElapsedTime=((cDate(FinalClose)-cDate(NL_Onb))*24)*60
						'response.Write "ElapsedTime="&ElapsedTime&"<BR>"
						TempNL_Cls=NL_Cls
						Hours = Int (ElapsedTime / 60)	
						Minutes = ElapsedTime - (Hours * 60)
						If Minutes>0 then 
							Minutes=cInt(Minutes)
							else
							Minutes=0
						End if
						DropTime=FinalClose	
						OnBoardTime=NL_Onb					
				%>
				<!--
				<tr>
					<td class="DetailsDetails" colspan="2" align="left">
						Arrived at Destination
					</td>
					<td class="DetailsDetails" colspan="2" align="left">
						<%=FinalClose%> (<%=Hours%> hrs <%=Minutes%> mins)
					</td>
				</tr>
				-->								
				<%		
				End if
		
		
		
		
	End if			
		
		
		
		
		
		If trim(AcknowledgedHUB)>"" then
			'Response.write "GOT HERE2!!!!<BR>"
			ElapsedTime=((cDate(AcknowledgedHUB)-cDate(ArrivedAtHUB))*24)*60
			Hours = Int (ElapsedTime / 60)	
			Minutes = ElapsedTime - (Hours * 60)
			If Minutes>0 then 
				Minutes=cInt(Minutes)
				else
				Minutes=0
			End if	
			%>	
			<tr>
				<td class="DetailsDetails" colspan="2" align="left">
					<%if trim(FromLocation)="DSTK" then
						Response.Write "Acknowledged by Material Handler"
						else
						'Response.Write "Acknowledged by Second Driver"
						Response.Write "Driver Acknowledged"
					End if%>
				</td>
				<td class="DetailsDetails" colspan="2" align="left">
					<%=AcknowledgedHUB%> (<%=Hours%> hrs <%=Minutes%> mins)
				</td>
			</tr>
			<%
			OnBoardTime=AcknowledgedHUB
		end if			
	
		
		
		
		'Response.Write "DepartedHUB="&DepartedHUB&"<BR>"
		'Response.Write "AcknowledgedHUB="&AcknowledgedHUB&"<BR>"
		If trim(DepartedHUB)>"" and trim(AcknowledgedHUB)>"" then
			'Response.write "GOT HERE2!!!!<BR>"
			ElapsedTime=((cDate(DepartedHUB)-cDate(AcknowledgedHUB))*24)*60
			Hours = Int (ElapsedTime / 60)	
			Minutes = ElapsedTime - (Hours * 60)
			If Minutes>0 then 
				Minutes=cInt(Minutes)
				else
				Minutes=0
			End if	
			%>	
			<tr>
				<td class="DetailsDetails" colspan="2" align="left">
					Departed HUB
				</td>
				<td class="DetailsDetails" colspan="2" align="left">
					<%=DepartedHUB%> (<%=Hours%> hrs <%=Minutes%> mins)
				</td>
			</tr>
			<%
			OnBoardTime=DepartedHUB
		end if
		
		
		
	
		
		
		
	end if%>
	
			
	
			
	<%if DropTime<>"1/1/1900" then
		ElapsedTime=((cDate(DropTime)-cDate(OnBoardTime))*24)*60
		Hours = Int (ElapsedTime / 60)	
		Minutes = ElapsedTime - (Hours * 60)
		If Minutes>0 then 
			Minutes=cInt(Minutes)
			else
			Minutes=0
		End if	
		%>	
		<tr>
			<td class="DetailsDetails" colspan="2" align="left">
				Delivered
			</td>
			<td class="DetailsDetails" colspan="2" align="left">
				<%=DropTime%> (<%=Hours%> hrs <%=Minutes%> mins)
			</td>
		</tr>
	<%end if%>		
	

			
	<%
	if DropTime<>"1/1/1900" then
		ElapsedTime=((cDate(DropTime)-cDate(DisplayBookTime))*24)*60
		Hours = Int (ElapsedTime / 60)	
		Minutes = ElapsedTime - (Hours * 60)
		If Minutes>0 then 
			Minutes=cInt(Minutes)
			else
			Minutes=0
		End if		
		%>	
		<tr>
			<td class="DetailsTitles" colspan="2" align="left">
				Total Delivery Time
			</td>
			<td class="DetailsTitles" colspan="2" align="left">
				<%=Hours%> hrs <%=Minutes%> mins
			</td>
		</tr>
		<%	
	End if	
	
	If trim(ToLocation)="CPGP" or trim(FromLocation)="Compugraphics" or trim(ToLocation)="TOPPAN" or trim(FromLocation)="TOPPAN" then
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
		''''''''''''QUERY FOR DOCUMENTS/LOTS/ETC'''''''''''''''''
		Set oConn2 = Server.CreateObject("ADODB.Connection")
		oConn2.ConnectionTimeout = 200
		oConn2.Provider = "MSDASQL"
		oConn2.Open DATABASE
		Err.Clear
		l_cSQL2="SELECT fcrefs.rf_ref, fcrefs.rf_fh_id, fcrefs.pod, fcrefs.ref_Status "_ 
		& " FROM  fcrefs "_  
		& " WHERE (rf_fh_id= '"&OrderID&"') ORDER BY rf_ref"					
		Set oRs2 = oConn2.Execute(l_cSQL2)
			Do while not oRs2.eof
			'a=a+1
			LotDocumentNumber=oRs2("rf_ref")
			LotJobNumber=oRs2("rf_fh_id")
			'PODID=oRs2("POD")
			Ref_Status=oRs2("Ref_Status")
			'Reflist=Reflist & CommaWord & LotDocumentNumber
			'CommaWord=", "
			'''''''''''''QUERY FOR TO LOCATION'''''''''''''''''''''''''	
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
				
			Set oConn = Server.CreateObject("ADODB.Connection")
			oConn.ConnectionTimeout = 200
			oConn.Provider = "MSDASQL"
			oConn.Open DATABASE
			l_cSQL="Select * from marksview2 "
			l_cSQL=l_cSQL&" WHERE (jobnum > '""')"
			If LotDocumentNumber>"" then     
				l_cSQL=l_cSQL&" AND (ref = '"&trim(LotDocumentNumber)&"') "
			End if
			If FromLocation="Compugraphics" OR FromLocation="TOPPAN" then
				l_cSQL=l_cSQL&" AND (jobnum < '"&LotJobNumber&"') "
				DeliveryWord="Previous Deliveries"
			End if
			If (trim(ToLocation)="CPGP" OR trim(ToLocation)="TOPPAN") then
				l_cSQL=l_cSQL&" AND (jobnum > '"&LotJobNumber&"') "
				DeliveryWord="Additional Deliveries"
			end if		
			l_cSQL=l_cSQL&" Order by shipdate DESC" 
			'Response.Write "l_cSQL="&l_cSQL&"<BR>"
			Set oRs = oConn.Execute(l_cSQL)
			
			'If oRs.eof then
			'	ErrorMessage="There are no orders that match your criteria"
			'end if
			'If Err.Number <> 0 Then                                               
			'Response.Write ErrorMessage="Error Executing the query.  Error:" & Err.Description
			'End if
			If NOT oRs.EOF then 
				If xyz<1 then
					closetable="y"
					%>
					<tr><td colspan="4" class="DetailsTitles"><%=DeliveryWord%>: 
					<%	
				end if
				xyz=xyz+1
				'response.Write "got here<br>"
				anotherjob=oRs("jobnum")
				'Response.Write "anotherjob="&anotherjob&"<BR>"
				If nop>0 then response.Write " ," end if
				%>
				<a href="jobanalysis.asp?inputjobnumber=<%=anotherjob%>"><%=LotDocumentNumber%></a>
				<%
				nop=nop+1
			End if
			Set oRS=nothing
			
			
			
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			oRs2.movenext
			LOOP
		Set oRs2=nothing
		if closetable="y" then
			Response.Write "</td></tr>"
		End if

		'LengthOfReflist=Len(Reflist)-1
		'Reflist=Left(Reflist, LengthOfReflist)
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
	End if
	
	
	%>							
	<!--	
	<tr>
		<td width="25%" Class="DetailsTitles">Job Number</td>
		<td width="25%" Class="DetailsDetails"><%=OrderID%></td>
		<td width="25%" Class="DetailsTitles">Priority</td>
		<td width="25%" Class="DetailsDetails"><%=DisplayPriority%></td>		
	</tr>
	<tr>
		<td width="25%" Class="DetailsTitles">Job Number</td>
		<td width="25%" Class="DetailsDetails"><%=OrderID%></td>
		<td width="25%" Class="DetailsTitles">Submitted By</td>
		<td width="25%" Class="DetailsDetails"><%=SubmittedBy%></td>		
	</tr>
	-->			
</table>
   
    
    </td></tr>



 
	<tr Height="50">
		<td>&nbsp;</td>
	</tr>

  <tr Height="30"> 
    <td NOWRAP valign="middle" align="right" class="MainPageText"> 
      &nbsp;
    </td>
	<td width="5">&nbsp;</td>
    <td width="136"> 
     &nbsp;
    </td>
	<td width="5">&nbsp;</td>
    <td width="725"> 
      &nbsp;
    </td>
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
</form>
<tr><td Height="90%">&nbsp;</td></tr>
<tr>
    <td height="100" class="FleetXGreySection" colspan="2">
        <!-- #include file="../BottomSection.asp" -->
    </td>
</tr>
<tr><td height="15" class="FleetXRedSectionSmall" colspan="2" align="center"><%=CopywriteNotice %></td></tr>
</td></tr></table>


</body>
</html>

