<%@ LANGUAGE="VBSCRIPT" %>
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
    PageTitle="ORDER DETAILS"

%>
<title>FleetX - <%=PageTitle %></title>

    <%
    LotPage=Request.Querystring("NewWindow")
    ''''''THIS CONVERTS ALL DATEDIFF FUNCTIONS INTO HOURS AND MINUTES....SWEET!
    function datediffToWords(d1, d2) 
        minutes = abs(datediff("n", d1, d2)) 
        if minutes <= 0 then 
            word = "0 mins" 
        else 
            word = "" 
            if minutes >= 24*60 then 
                word = word & minutes\(24*60) & " days " 
            end if 
            minutes = minutes mod (24*60) 
            if minutes >= 60 then 
                word = word & minutes\(60) & " hrs " 
            end if 
            minutes = minutes mod 60 
            word = word & minutes & " mins" 
        end if 
        datediffToWords = word 
    end function 
    
       
    
    InputJobNumber=trim(Request.Form("InputJobNumber"))
    If InputJobNumber="" then
        InputJobNumber=trim(Request.QueryString("InputJobNumber"))
    End if
 

   
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



    <tr><td width="100%" align="center"><!-- main page stuff goes here! -->
    
<table width="700" cellpadding="2" cellspacing="0" border="1" align="center" ID="Table1"> 
  
<%      
'Response.Write "Database="&Database&"<BR>"
'''''''''''''QUERY STATEMENT'''''''''''''''''''''''''''''''''''''''''''
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.ConnectionTimeout = 200
oConn.Provider = "MSDASQL"
oConn.Open DATABASE
l_cSQL=l_cSQL&"Select * "
l_cSQL=l_cSQL&"From Order_Details_FleetX "
l_cSQL=l_cSQL&"WHERE (jobnum = '"&InputJobNumber&"')"
'Response.Write "111 orderdetails l_cSQL="&l_cSQL&"<BR>"
Set oRs = oConn.Execute(l_cSQL)
If oRs.eof then
	ErrorMessage="There are no orders that match your criteria"
end if
If Err.Number <> 0 Then                                               
Response.Write ErrorMessage="Error Executing the query.  Error:" & Err.Description
End if
if not oRs.EOF then 
    xxx=xxx+1
jobnum=oRs("jobnum")
Shipdate=oRs("Shipdate")
fh_bt_id=oRs("fh_bt_id")
TIUser=oRs("TIUser")
custpo=oRs("custpo")
'Response.write "TIUser="&TIUser&"<BR>"
'Response.write "custpo="&custpo&"<BR>"
To_Id=oRs("To_Id")
Priority=oRs("Priority")
Statcode=oRs("Statcode")
MaterialType=oRs("MaterialType")
fh_user6=oRs("fh_user6")
fl_pkey=oRs("fl_pkey")
From_ID=oRs("From_ID")
FromFullName=oRs("FromFullName")
fl_sf_clname=oRs("fl_sf_clname")
fl_sf_cfname=oRs("fl_sf_cfname")
fl_sf_phone=oRs("fl_sf_phone")
fl_sf_email=oRs("fl_sf_email")
fl_sf_building=oRs("fl_sf_building")
FromAddress1=oRs("FromAddress1")
FromAddress2=oRs("FromAddress2")
FromCity=oRs("FromCity")
FromState=oRs("FromState")
FromCountry=oRs("FromCountry")
FromZipCode=oRs("FromZipCode")
To_ID=oRs("To_ID")
ToFullNAme=oRs("ToFullNAme")
fl_st_clname=oRs("fl_st_clname")
fl_st_cfname=oRs("fl_st_cfname")
fl_st_phone=oRs("fl_st_phone")
fl_st_email=oRs("fl_st_email")
fl_st_building=oRs("fl_st_building")
ToAddress1=oRs("ToAddress1")
ToAddress2=oRs("ToAddress2")
ToCity=oRs("ToCity")
ToState=oRs("ToState")
ToCountry=oRs("ToCountry")
ToZipCode=oRs("ToZipCode")
Unit=oRs("Unit")
Driver=oRs("Driver")
FromComments=oRs("SfComment")
ToComments=oRs("StComment")
Disptime=oRs("Disptime")
AccTime=oRs("AccTime")
OnbTime=oRs("OnbTime")
DropTime=oRs("DropTime")
DueTime=oRs("DueTime")
ReadyTime=oRs("ReadyTime")
At_Hub=oRs("At_Hub")
Onbleg2=oRs("Onbleg2")
Accleg2=oRs("Accleg2")
Pu_Driver=oRs("Pu_Driver")
Do_Driver=oRs("Do_Driver")
fl_pu_driver2=oRs("fl_pu_driver2")
fl_do_driver2=oRs("fl_do_driver2")
fl_job_closed=oRs("fl_job_closed")
fl_FinalDestination=oRs("fl_FinalDestination")
Ref=oRs("Ref")
MaterialDescription=oRs("MaterialDescription")
PartNumber=oRs("PartNumber")
POD=oRs("POD")
PODDateTime=oRs("PODDateTime")
rf_box=oRs("rf_box")
NumberOfPieces=oRs("NumberOfPieces")
IsPalletized=oRs("IsPalletized")
IsStacked=oRs("IsStacked")
NumberOfPallets=oRs("NumberOfPallets")
Weight=oRs("Weight")
DimLength=oRs("DimLength")
DimWidth=oRs("DimWidth")
DimHeight=oRs("DimHeight")
MeasurementType=oRs("MeasurementType")






'jobnum=oRs("jobnum")
Orderid=JobNum
'custpo=oRs("custpo")
'To_Id=oRs("To_Id")
'TIUser=oRs("TIUser")
'Priority=oRs("Priority")
'Shipdate=oRs("Shipdate")
'From_ID=oRs("From_ID")
'Driver=oRs("Driver")
'Unit=oRs("Unit")
'DueTime=oRs("DueTime")
'Disptime=oRs("Disptime")
'AccTime=oRs("AccTime")
'OnbTime=oRs("OnbTime")
'DropTime=oRs("DropTime")
'ReadyTime=oRs("ReadyTime")
'Response.write "REadyTime="& ReadyTime &"<BR>"
'SfComment=oRs("SfComment")
'StComment=oRs("StComment")
'At_Hub=oRs("At_Hub")
'Onbleg2=oRs("Onbleg2")
'Accleg2=oRs("Accleg2")
'POD=oRs("POD")
'Ref=oRs("Ref")
'Response.write "Ref="&Ref&"<BR>"
'Statcode=oRs("Statcode")
'Pu_Driver=oRs("Pu_Driver")
'Do_Driver=oRs("Do_Driver")
'fh_bt_id=oRs("fh_bt_id")
'MaterialType=oRs("MaterialType")
'fl_pkey=oRs("fl_pkey")
'fl_job_closed=oRs("fl_job_closed")
'fl_FinalDestination=oRs("fl_FinalDestination")
'FromAddress1=oRs("FromAddress1")
'FromAddress2=oRs("FromAddress2")
'FromCity=oRs("FromCity")
'FromState=oRs("FromState")
'FromCountry=oRs("FromCountry")
'FromZipCode=oRs("FromZipCode")
'ToAddress1=oRs("ToAddress1")
'ToAddress2=oRs("ToAddress2")
'ToCity=oRs("ToCity")
'ToState=oRs("ToState")
'ToCountry=oRs("ToCountry")
'ToZipCode=oRs("ToZipCode")
'fl_pu_driver2=oRs("fl_pu_driver2")
'fl_do_driver2=oRs("fl_do_driver2")
'ToFullNAme=oRs("ToFullNAme")
'FromFullName=oRs("FromFullName")
'fh_user6=oRs("fh_user6")
'fl_sf_clname=oRs("fl_sf_clname")
'fl_sf_cfname=oRs("fl_sf_cfname")
'fl_sf_phone=oRs("fl_sf_phone")
'fl_sf_email=oRs("fl_sf_email")
'fl_sf_building=oRs("fl_sf_building")
'fl_st_clname=oRs("fl_st_clname")
'fl_st_cfname=oRs("fl_st_cfname")
'fl_st_phone=oRs("fl_st_phone")
'fl_st_email=oRs("fl_st_email")
'fl_st_building=oRs("fl_st_building")
'MaterialDescription=oRs("MaterialDescription")
'PartNumber=oRs("PartNumber")
'rf_box=oRs("rf_box")
'NumberOfPieces=oRs("NumberOfPieces")
'IsPalletized=oRs("IsPalletized")
'IsStacked=oRs("IsStacked")
'NumberOfPallets=oRs("NumberOfPallets")
'Weight=oRs("Weight")
'DimLength=oRs("DimLength")
'DimWidth=oRs("DimWidth")
'DimHeight=oRs("DimHeight")
'MeasurementType=oRs("MeasurementType")

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
	If ArrivedAtHUB>"1/1/1900" then
	    DropTime=ArrivedAtHUB
	End if
	If OnBoardTime="1/1/1900" then 
		DisplayOnBoardTime="Pending"
		else
		DisplayOnBoardTime=OnBoardTime
	End if	
	If DropTime="1/1/1900" then 
		DisplayDropTime="Still In Transit"
		else
		DisplayDropTime=DropTime
	End if	
	If isdate(fl_job_closed) AND (fl_job_closed>"1/1/1900") then
		DisplayDropTime=fl_job_closed
	End if	
	If trim(MaterialType)="Secure Waf" then
		Reflist="Secure Wafer(s): "
	End if
	If trim(MaterialType)="ITAR" then
		Reflist="ITAR(s): "
	End if    
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
	'response.Write "statcode="&Statcode&"***<BR>"	
	Select Case StatCode
		Case "0", "HLD"
			StatCode="HELD"
		Case "1", "SCD"
			StatCode="Scheduled"
		Case "2", "RAP"
			StatCode="Booked"
		Case "3", "OPN"
			StatCode="Open"
		Case "4", "ACC"
			StatCode="Acknowledged by driver"
		Case "5", "ONB"
			StatCode="On Board"
		Case "6", "UND"
			StatCode="Undispatched-Please report this to Mark Maggiore immediately at 214-956-0400 xt. 212"
		Case "9", "CLS"
			StatCode="Closed"
		Case "10", "INV"
			StatCode="Invoiced"
		Case "13", "PUO"
			StatCode="Paperwork on Board"
		Case "98", "CAN"
			StatCode="<font color='red'>CANCELLED</font>"
		Case "99", "DEL"
			StatCode="Deleted"
		Case "53", "ARV"
			StatCode="Arrived at HUB"
		Case "54", "DPV"
			StatCode="Departed HUB"
		Case "55", "AC2"
			StatCode="Acknowledged by 2nd Driver"					
		Case ELSE
			StatCode="Unknown-Please report this to Mark Maggiore immediately at 214-956-0400 xt. 212"																																																																	
	End select
	Select Case priority
		Case "WF", "CS", "KW", "ST"
			DisplayPriority="Standard"
		Case "CE", "XP"
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



    '''''''''''''QUERY FOR POD INFORMATION'''''''''''''''''''''''''
    'Response.write "BillToID="&BillToID&"<BR>"
    If BillToID<>"999999" then

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
    l_cSQL2=l_cSQL2&" WHERE (Userid= '"&pu_driver&"')"					
    'Response.write "l_cSQL2="&l_cSQL2&"<BR>"	
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
    l_cSQL2=l_cSQL2&" WHERE (Userid= '"&do_driver&"')"	
    'Response.write "l_cSQL="&l_cSQL&"<BR>"				
    Set oRs2 = oConn2.Execute(l_cSQL2)
    If not oRs2.eof then
	    CLSDriverName=oRs2("FirstName")&" "&oRs2("LastName")
    End if
    Set oRs2=nothing
    
 	'fl_pu_driver2=oRs("fl_pu_driver2")
	'Response.Write "YYYY="&ONBDriverID&"YYYYYYY<BR>"
	'fl_do_driver2=oRs("fl_do_driver2")   
    
    '''''''''''''QUERY FOR PICKUP DRIVER 2'''''''''''''''''''''''''
    Set oConn2 = Server.CreateObject("ADODB.Connection")
    oConn2.ConnectionTimeout = 200
    oConn2.Provider = "MSDASQL"
    oConn2.Open INTRANET
    Err.Clear
    l_cSQL2="SELECT FirstName, LastName "
    l_cSQL2=l_cSQL2&" FROM Intranet_Users " 
    l_cSQL2=l_cSQL2&" WHERE (Userid= '"&fl_pu_driver2&"')"					
    Set oRs2 = oConn2.Execute(l_cSQL2)
    If not oRs2.eof then
	    fl_pu_driver2=oRs2("FirstName")&" "&oRs2("LastName")
    End if
    Set oRs2=nothing
    'Response.Write "l_cSQL2="&l_cSQL2&"<BR>"
    '''''''''''''QUERY FOR DROPOFF DRIVER 2'''''''''''''''''''''''''
    Set oConn2 = Server.CreateObject("ADODB.Connection")
    oConn2.ConnectionTimeout = 200
    oConn2.Provider = "MSDASQL"
    oConn2.Open INTRANET
    Err.Clear
    l_cSQL2="SELECT FirstName, LastName "
    l_cSQL2=l_cSQL2&" FROM Intranet_Users " 
    l_cSQL2=l_cSQL2&" WHERE (Userid= '"&fl_do_driver2&"')"					
    Set oRs2 = oConn2.Execute(l_cSQL2)
    If not oRs2.eof then
	    fl_do_driver2=oRs2("FirstName")&" "&oRs2("LastName")
    End if
    Set oRs2=nothing    

    ''''''''''''''''''QUERY FOR PRIORITY'''''''''''''''''''
    If isnumeric(Priority) then
        Set oConn2 = Server.CreateObject("ADODB.Connection")
        oConn2.ConnectionTimeout = 200
        oConn2.Provider = "MSDASQL"
        oConn2.Open DATABASE
        Err.Clear
        l_cSQL2="SELECT Priority "
        l_cSQL2=l_cSQL2&" FROM Priorities " 
        l_cSQL2=l_cSQL2&" WHERE (PriorityID= '"&priority&"')"					
        Set oRs2 = oConn2.Execute(l_cSQL2)
        If not oRs2.eof then
	        DisplayPriority=oRs2("Priority")
        End if
        Set oRs2=nothing   
        else
        DisplayPriority=Priority
    End if
    
    
    
    
    
    If trim(ONBDriverName)="" then ONBDriverName="n/a" end if
    If trim(CLSDriverName)="" then CLSDriverName="n/a" end if
    If xxx=1 then
    FirstLeg=fl_Pkey
    %>
    
	<tr>
		<td colspan="4">
			<table width="100%" ID="Table2">
				<tr>
					<td width="33%">
						<img src="../images/FleetX_Small.jpg" height="50" width="168" />
					</td>
					<td width="34%" class="MainPageTextBold" align="center">
						Delivery Details
					</td>
					<td width="33%" align="right" valign="top" Class="MainPageTextBoldRight"><%=Session("txt_cm_desc")%></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td colspan="4" class="FleetXRedSectionSmallWaybill" align="center">
			Shipment Information
		</td>
	</tr>	
	<tr>
		<td width="25%" Class="MainPageTextBold">Job Number</td>
		<td width="25%" Class="MainPageText"><%=OrderID%></td>
		<td width="25%" Class="MainPageTextBold">Current Status</td>
		<td width="25%" Class="MainPageText"><%=StatCode%></td>		
	</tr>
	<tr>
		<td width="25%" Class="MainPageTextBold">Submitted By</td>
		<td width="25%" Class="MainPageText"><%=TIUser%></td>	
		<td width="25%" Class="MainPageTextBold">Priority</td>
		<td width="25%" Class="MainPageText"><%=DisplayPriority%></td>
	</tr>
    <!--
	<tr>
		<td width="25%" Class="MainPageTextBold">Phone Contact</td>
		<td width="25%" Class="MainPageText"><%=TIUser%></td>	
		<td width="25%" Class="MainPageTextBold">Email Contact</td>
		<td width="25%" Class="MainPageText"><%=DisplayPriority%></td>
	</tr>	
    -->	    
    <%
    '''''''''''''QUERY FOR LOTS INFORMATION'''''''''''''''''''''
	    Set oConn2 = Server.CreateObject("ADODB.Connection")
	    oConn2.ConnectionTimeout = 200
	    oConn2.Provider = "MSDASQL"
	    oConn2.Open DATABASE
	    Err.Clear
	    l_cSQL2="SELECT rf_ref, ref_status FROM FCREFS"
	    l_cSQL2=l_cSQL2&" WHERE (RF_FH_id= '"&JobNum&"')"					
	    'Response.Write "l_cSQL2="&l_cSQL2&"<BR>"
	    Set oRs2 = oConn2.Execute(l_cSQL2)
	    Do while not oRs2.eof 
	        YYY=YYY+1
		    Refs=trim(oRs2("RF_REF"))
            Ref_Status=trim(oRs2("ReF_status"))
            ListOfRefs=ListOfRefs&Refs
            If Ref_Status="X" then
            ListOfRefs=ListOfRefs&" (Cancelled)"    
            End if
	        ListOfRefs=ListOfRefs&", "
	    oRs2.movenext
	    Loop
	    Set oRs2=nothing
	    LenRefs=Len(ListOfRefs)
	    'response.Write "lenRefs="&LenRefs&"<BR>"
        'Response.Write "ListOfRefs="&ListOfRefs&"<BR>"
	    ListOfRefs=Left(ListOfRefs,(LenRefs-2))	
    'End if
        %>

	    <tr>
		    <td class="MainPageText" colspan="2">
			    <span class="MainPageTextBold"><%=DisplayBookTimeWord%> Time:  </span><%=shipdate%>
		    </td>
		    <td class="MainPageText" colspan="2">
			    <span class="MainPageTextBold">Ready Time:  </span><%=ReadyTime%>
		    </td>
	    </tr>
 	    <tr>
		    <td class="MainPageText" colspan="2">
			    <span class="MainPageTextBold">Dispatch Time:  </span><%=Disptime%>
		    </td>
		    <td class="MainPageText" colspan="2">
			    <span class="MainPageTextBold">Due Time:  </span><%=DueTime%>
		    </td>
	    </tr>
	<tr>
		<td colspan="4" class="FleetXRedSectionSmallWaybill" align="center">
			Commodity Information
		</td>
	</tr>
    <%If trim(ListOfRefs)>"" then %>
 	<tr>
		<td width="50%" colspan="2" Class="MainPageTextBold">Document Number</td>
		<td width="50%" colspan="2" Class="MainPageText"><%=ListOfRefs%></td>
		
	</tr> 
    <%
    End if
    'Response.write "MaterialDescription="&MaterialDescription&"XXX<BR>"
    'Response.write "NumberOfPieces="&NumberOfPieces&"XXX<BR>"
    'Response.write "rf_box="&rf_box&"XXX<BR>"

    If trim(MaterialDescription)>"" or trim(NumberOfPieces)>"" or trim(rf_box)>"" then %> 
	<tr>
		<td width="25%" Class="MainPageTextBold">Description</td>
		<td width="25%" Class="MainPageText"><%=MaterialDescription%></td>
		<td width="25%" Class="MainPageTextBold">Piece(s)</td>
		<td width="25%" Class="MainPageText"><%=NumberOfPieces%>&nbsp;&nbsp;<%=rf_box %></td>		
	</tr>
    <%End if 
    If trim(IsPalletized)>"" or trim(NumberofPallets)>"" or trim(IsStacked)>"" then %>
	<tr>
		<td width="25%" Class="MainPageTextBold">Palletized?</td>
		<td width="25%" Class="MainPageText"><%=IsPalletized%>, <%=NumberofPallets %></td>	
		<td width="25%" Class="MainPageTextBold">Stacked?</td>
		<td width="25%" Class="MainPageText"><%=IsStacked%></td>
	</tr>
    <%End if
    If trim(Weight)>"" or trim(DimLength)>"" or trim(DimWidth)>""  or trim(DimHeight)>"" then %>
	<tr>
		<td width="25%" Class="MainPageTextBold">Weight</td>
		<td width="25%" Class="MainPageText"><%=Weight%> Pounds</td>	
		<td width="25%" Class="MainPageTextBold">Dimensions</td>
		<td width="25%" Class="MainPageText"><%=DimLength%>&nbsp;&nbsp;X&nbsp;&nbsp;<%=DimWidth%>&nbsp;&nbsp;X&nbsp;&nbsp;<%=DimHeight%></td>
	</tr>
    <%End if %>           
     
        
        
               	    
	    <tr>
		    <td class="FleetXRedSectionSmallWaybill" colspan="2">
			    Pickup
		    </td>
		    <td class="FleetXRedSectionSmallWaybill" colspan="2">
			    Delivery
		    </td>
	    </tr>
	    <tr>
		    <td class="MainPageText" colspan="2" valign="top">
			    <span class="MainPageTextBold">Pickup Time:</span>&nbsp;&nbsp;<%=ONBTime%>
		    </td>
		    <td class="MainPageText" colspan="2" valign="top">
			    <span class="MainPageTextBold">Delivery Time:</span>&nbsp;&nbsp;<%=DisplayDropTime%>
		    </td>
	    </tr>    		

	    <tr>
		    <td class="MainPageText" colspan="2" valign="top">
                <%=FromFullName %> <br />
			    <%
			    if trim(FromAddress1)>"" then
				    Response.Write FromAddress1&"<BR>"
			    End if			
			    if trim(FromAddress2)>"" then
				    Response.Write FromAddress2&"<BR>"
			    End if
			    %>
			    <%=FromCity%>, <%=FromState%>&nbsp;&nbsp;<%=FromZipCode%><br>
			    <%=FromCountry%>&nbsp;&nbsp;<br />
                <%=fl_sf_phone %><br />
                <a href="mailto:<%=fl_sf_email %>"><%=fl_sf_email %></a>
		    </td>
		    <td class="MainPageText" colspan="2" valign="top">
                <%=ToFullName%><br />
			    <%
			    if trim(toAddress1)>"" then
				    Response.Write toAddress1&"<BR>"
			    End if			
			    if trim(toAddress2)>"" then
				    Response.Write toAddress2&"<BR>"
			    End if
			    %>
			    <%=toCity%>, <%=toState%>&nbsp;&nbsp;<%=toZipCode%><br>
			    <%=toCountry%>&nbsp;&nbsp;<br />
                <%=fl_st_phone %><br />
                <a href="mailto:<%=fl_st_email %>"><%=fl_st_email %></a>

		    </td>
	    </tr>
	    <tr>
		    <td class="MainPageText" colspan="2" valign="top">
			    <span class="MainPageTextBold">Special Instructions:  </span><%=DisplayFromComments%>
		    </td>
		    <td class="MainPageText" colspan="2" valign="top">
			    <span class="MainPageTextBold">Special Instructions:  </span><%=DisplayToComments%>
		    </td>
	    </tr>
        <%
	    Set oConn22 = Server.CreateObject("ADODB.Connection")
	    oConn22.ConnectionTimeout = 200
	    oConn22.Provider = "MSDASQL"
	    oConn22.Open DATABASE
	    Err.Clear
	    l_cSQL22="SELECT JobChanges.fh_id, JobChanges.ChangeReason, JobChanges.ChangeDate, JobChangeCategories.Category, lcintranet.dbo.Intranet_Users.FirstName, lcintranet.dbo.Intranet_Users.LastName FROM JobChanges INNER JOIN JobChangeCategories ON JobChanges.ChangeCategory = JobChangeCategories.CategoryID INNER JOIN lcintranet.dbo.Intranet_Users ON JobChanges.SupervisorID = lcintranet.dbo.Intranet_Users.UserID   "
	    l_cSQL22=l_cSQL22&" WHERE (FH_id= '"&OrderID&"')"					
	    'Response.Write "l_cSQL22="&l_cSQL22&"<BR>"
	    Set oRs22 = oConn22.Execute(l_cSQL22)
	    Do while not oRs22.eof 
            'Response.write "GOT HERE!!!!<BR>"
        %>
 	    <tr>
		    <td class="MainPageText" colspan="4" valign="top">
			    <span class="MainPageTextBold">This job was edited by a supervisor<br /></span>
        <%

		    ChangeReason=oRs22("ChangeReason")
		    ChangeDate=oRs22("ChangeDate")
		    Category=oRs22("Category")
		    FirstName=oRs22("FirstName")
		    LastName=oRs22("LastName")
		    Category=oRs22("Category")
            %>
            Supervisor: <%=FirstName%> <%=LastName %><br />
            Change Category: <%=Category%><br />
            Comments: <%=ChangeDate%> - <%=ChangeReason %>
    		    </td>
	    </tr>
            <%

	    oRs22.movenext
	    Loop
	    Set oRs22=nothing	
    'End if



	    Set oConn22 = Server.CreateObject("ADODB.Connection")
	    oConn22.ConnectionTimeout = 200
	    oConn22.Provider = "MSDASQL"
	    oConn22.Open DATABASE
	    Err.Clear
	    l_cSQL22="SELECT XID, Reason, OtherReason, CancelDate FROM CancelledOrders "
	    l_cSQL22=l_cSQL22&" WHERE (FH_id= '"&OrderID&"')"					
	    'Response.Write "603 orderdetails l_cSQL22="&l_cSQL22&"<BR>"
	    Set oRs22 = oConn22.Execute(l_cSQL22)
	    Do while not oRs22.eof 
            'Response.write "GOT HERE!!!!<BR>"

            CancelXID=oRs22("XID")
		    CancelReason=oRs22("Reason")
            CancelReasonOther=oRs22("OtherReason")
		    CancelDate=oRs22("CancelDate")



									Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
										RSEVENTS2.CursorLocation = 3
										RSEVENTS2.CursorType = 3
										RSEVENTS2.ActiveConnection = Database
										l_csql = "SELECT RequestorName FROM PreExistingRequestor WHERE (RequestorID = '"&CancelXID&"')"
										'response.write("Query:" & l_cSQL)
										RSEVENTS2.Open l_cSQL, Database, 1, 3
										if not RSEVENTS2.EOF then
											CancelXID=RSEVENTS2("RequestorName")
                                        End if
										RSEVENTS2.close
									Set RSEVENTS2 = Nothing	

									Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
										RSEVENTS2.CursorLocation = 3
										RSEVENTS2.CursorType = 3
										RSEVENTS2.ActiveConnection = Database
										l_csql = "SELECT Category FROM JobChangeCategories WHERE (CategoryID = '"&CancelReason&"')"
										'response.write("Query:" & l_cSQL)
										RSEVENTS2.Open l_cSQL, Database, 1, 3
										if not RSEVENTS2.EOF then
											Category=RSEVENTS2("Category")
                                        End if
										RSEVENTS2.close
									Set RSEVENTS2 = Nothing	

        %>
 	     <tr>
		    <td class="MainPageText" colspan="4" valign="top">
			    <span class="MainPageTextBold">This job had a cancellation by  <%=CancelXID%><br /></span>
            Comments: <%=CancelDate%> - <%=Category %> - <%=CancelReasonOther%>
    		    </td>
	    </tr> 
            <%
	    oRs22.movenext
	    Loop
	    Set oRs22=nothing
    %>
             	
	    <tr>
		    <td colspan="4" class="FleetXRedSectionSmallWaybill" align="center">
			    Delivery History
		    </td>
	    </tr>
	    <tr>
		    <td class="MainPageTextBold" colspan="2" align="left">
			    Milestones
		    </td>
		    <td class="MainPageTextBold" colspan="2" align="left">
			    Times
		    </td>
	    </tr>
	    <%if ShipDate<>"1/1/1900" then%>	
	    <tr>
		    <td class="MainPageText" colspan="2" align="left">
			    <b><%=DisplayBookedWord%></b>
		    </td>
		    <td class="MainPageText" colspan="2" align="left">
			    <%=ShipDate%>
		    </td>
	    </tr>
	    <%'''''end if%>	    		    	    	    	    
	    <%
	    StopDisplayingLotsNow="y"
	    End if
	End if
	'If trim(fl_Pkey)<>Trim(Tempfl_Pkey) or trim(fl_pkey)="" then
	    %> 
	    <%If AccTime>"1/1/1900" then%>
	    <tr>
		    <td class="MainPageText" colspan="2" align="left">
			    <b>Acknowledged</b>
		    </td>
		    <td class="MainPageText" colspan="2" align="left">
			    <%=AccTime%>
			    <%If xxx=1 then
			        response.write "&nbsp;("&datediffToWords(ShipDate,AccTime)&")"
			        'response.Write "<br>"&Booktime&" minus "&DriverAcknowledgementTime&"<BR>"
			        else
			        response.write "&nbsp;("&datediffToWords(PreviousDropTime,AccTime)&")"
			        'response.Write "<br>"&DriverAcknowledgementTime&" minus "&PreviousDroptime&"<BR>"
			      End if
			      'Response.Write "xxx="&xxx&"<BR>"
			      'Response.Write "Booktime="&Booktime&"<BR>"
			      'Response.Write "DropTime="&DropTime&"<BR>"
			      'Response.Write "DriverAcknowledgementTime="&DriverAcknowledgementTime&"<BR>"
			      'Response.Write "***************<br>"			      
			    %>
		    </td>
	    </tr>
	    <%end if %>
	    <%
	   ' Response.Write "OnBoardTime="&OnBoardTime&"******<BR>"
	    If ONBTime>"1/1/1900" then%>
	    <tr>
		    <td class="MainPageText" colspan="2" align="left">
			    <b>On Board</b>-<%=ONBDriverName%>
		    </td>
		    <td class="MainPageText" colspan="2" align="left">
			    <%=ONBTime%> 
			    <%response.write "&nbsp;("&datediffToWords(ACCTime,ONBTime)&")"%>
		    </td>
	    </tr>
	    <%end if %>
	    <%If at_HUB>"1/1/1900" then
        If trim(DocumentNumber)>"" and onboardtime="1/1/1900" then
            onboardtime=booktime
        End if
        %>
	    <!--HERE'S THE DROP HOURS/MINUTES-->
	    <tr>
		    <td class="MainPageText" colspan="2" align="left">
			    <b>Delivered to HUB</b>-<%=CLSDriverName%> 
		    </td>
		    <td class="MainPageText" colspan="2" align="left">

			    <%=at_HUB%>
                <%
                If trim(FromLocation)="DNP" then
                    response.write "&nbsp;("&datediffToWords(BookTime,at_HUB)&")"
                    else 
			        response.write "&nbsp;("&datediffToWords(OnBoardTime,at_HUB)&")"
                End if
                %>
		    </td>
	    </tr>			 
	    <%
	    PreviousDropTime=DropTime
	    End if


'''''''''''''''''''''''''''''''''	    
'''''''''''''SECOND LEG STUFF!!!
'''''''''''''''''''''''''''''''''	    
	 'Response.Write "Droptime="&Droptime&"<BR>"  
	 'Response.Write "ACCLEG2="&ACCLEG2&"<BR>"  
	'If trim(fl_Pkey)<>Trim(Tempfl_Pkey) or trim(fl_pkey)="" then
	    %> 
	    <%If ACCLEG2>"1/1/1900" then%>
	    <tr>
		    <td class="MainPageText" colspan="2" align="left">
			    <b>Acknowledged</b>
		    </td>
		    <td class="MainPageText" colspan="2" align="left">
			    <%=ACCLEG2%>
			    <%
			        response.write "&nbsp;("&datediffToWords(at_HUB,ACCLEG2)&")"
		      
			    %>
		    </td>
	    </tr>
	    <%end if %>
	    <%
	    'Response.Write "OnBoardTime="&OnBoardTime&"******<BR>"
	    If ONBLeg2>"1/1/1900" then%>
	    <tr>
		    <td class="MainPageText" colspan="2" align="left">
			    <b>On Board</b>-<%=fl_pu_driver2%>
		    </td>
		    <td class="MainPageText" colspan="2" align="left">
			    <%=ONBLeg2%> 
			    <%response.write "&nbsp;("&datediffToWords(ACCLEG2,ONBLeg2)&")"%>
		    </td>
	    </tr>
        <%else 
        ONBLeg2=ONBTime
        end if %>
	    <%
        'Response.Write "DropTime="&DropTime&"<BR>"
        'Response.Write "ONBLeg2="&ONBLeg2&"<BR>"
        'Response.Write "ONBoardTime="&ONBoardTime&"<BR>"

	    If DropTime>"1/1/1900" then
	    If ((ONBLeg2="1/1/1900") or (isnull(ONBLeg2)) or (trim(ONBLeg2)="")) then ONBLeg2=ONBoardTime End if
        'Response.Write "xxxONBLeg2="&ONBLeg2&"<BR>"
	    %>
	    <!--HERE'S THE DROP HOURS/MINUTES-->
	    <tr>
		    <td class="MainPageText" colspan="2" align="left">
			    <b>Delivered</b>-<%=CLSDriverName%>
		    </td>
		    <td class="MainPageText" colspan="2" align="left">
			    <%=DropTime%>
			    <% 'response.write "<br>765 orderdetails ONBoardTIme = " & ONBoardTIme & ",ONBLeg2=" & ONBLeg2 & ", droptime=" & droptime & "<br>"%>
          <%response.write "&nbsp;("&datediffToWords(ONBLeg2,DropTime)&")"%>
		    </td>
	    </tr>			 
	    <%
	    'Response.Write "ONBLeg2="&ONBLeg2&"<BR>"
	    PreviousDropTime=DropTime
	    End if	    
	    
	    
	    'Tempfl_Pkey=fl_Pkey 
	'End if
End if








%>



		<%
        '''''''''''GETS HUB INFO'''''''''''''''''''''''
        Set oConn2 = Server.CreateObject("ADODB.Connection")
        oConn2.ConnectionTimeout = 200
        oConn2.Provider = "MSDASQL"
        oConn2.Open DATABASE
        Err.Clear
        l_cSQL2="SELECT fl_st_id, fl_t_atd, fl_FinalDestination "
        l_cSQL2=l_cSQL2&" FROM fclegs " 
        l_cSQL2=l_cSQL2&" WHERE (fl_fh_id= '"&InputJobNumber&"')"	
        'Response.Write "l_cSQL2="&l_cSQL2&"<BR>"				
        Set oRs2 = oConn2.Execute(l_cSQL2)
        do while not oRs2.eof 
            tempfl_st_id=oRs2("fl_st_id")
            tempfl_FinalDestination=oRs2("fl_FinalDestination")
            If trim(tempfl_st_id)=trim(tempfl_FinalDestination) then
                tempDeliveryTime=oRs2("fl_t_atd")
            End if
            'Response.Write "tempfl_st_id="&tempfl_st_id&"<BR>"
            'Response.Write "tempDeliveryTime="&tempDeliveryTime&"<BR>"
            'Response.Write "tempfl_FinalDestination="&tempfl_FinalDestination&"<BR>"
        oRs2.movenext
        loop
        Set oRs2=nothing
 	    If DropTime="1/1/1900" then 
		    DisplayDropTime="Still In Transit"
		    else
		    DisplayDropTime=DropTime
	    End if
	    If cdate(tempDeliveryTime)>cdate("1/1/1900") then
	        DisplayDropTime=cdate(tempDeliveryTime)
	    End if       		

	'If trim(trackingnumber)>"" and (trim(FromLocation)="Compugraphics" OR trim(FromLocation)="TOPPAN") then
	%>
	<!-- <form method="post" action="http://my.shipgreyhound.com/cfw/trackOrder.login" target="_blank" ID="Form1">
	<tr>
	<input type="hidden" name="orderNumber" value="<%=trackingnumber%>" ID="Hidden1">
		<td nowrap align="right" class="MainPageText" colspan="2" valign="bottom"><span class="MainPageTextBold">Greyhound Tracking: </span>
		<input TYPE="IMAGE" SRC="../images/btnClickHereLink.gif" ALT="click here" ID="Image1" NAME="Image1"></td>						
		<td nowrap align="right" class="MainPageText" colspan="2"><span class="MainPageTextBold">Original Bus ETA: </span>
		<%=ETA%></td>						
	</tr>
	</form>	-->						
	<%
	'End if	

	If trim(fh_user6)>"" then
		%>
        <form id="myform" name="myform" action="http://www.fedex.com/Tracking" target="_blank" method="post">
		<tr>
			<td class="MainPageText" colspan="4" valign="top">
				<span class="MainPageTextBold">FedEx Tracking:</span>&nbsp;&nbsp;
                
                                       
                                        <input type="hidden" name="clienttype" id="clienttype" value="dotcom">
                                        <input type="hidden" name="track" id="track" value="y">
                                        <input type="hidden" name="ascend_header" id="ascend_header" value="1">
                                        <input type="hidden" name="cntry_code" id="cntry_code" value="us">
                                        <input type="hidden" name="language" id="language" value="english">
                                        <input type="hidden" name="mi" id="mi" value="n">
                                        <input type="hidden" name="tracknumbers" id="trackNbrs" value="<%=fh_user6%>" />


                                        <%
                                        'Response.write "fh_user6="&fh_user6&"<BR>"
                                        If trim(fh_user6)>"" and ucase(trim(fh_user6))<>"FLEETX" then %>

                                           <input type="submit" value="<%=fh_user6%>" name="Submit" id="gobutton"/>

                                        <%End if %>

                                    

			</td>
		</tr>
        </form>		
		<%
	End if	
	
	If Whatever="whatever" and DocumentNumber>"" and (FromLocation="CPGP" or FromLocation="Compugraphics" or ToLocation="CPGP" or ToLocation="TOPPAN" or FromLocation="TOPPAN" or ToLocation="TOPPANSC" or FromLocation="TOPPANSC" or ToLocation="TISHR" or FromLocation="TISHR" or ToLocation="PHO" or FromLocation="PHO") then
		%>
		<tr>
			<td class="MainPageText" colspan="4" valign="top">
				<span class="MainPageTextBold">Quick Tracking:</span>&nbsp;&nbsp;<a href="http://www.quickonline.com/cgi-bin/WebObjects/BOLSearch?bolNumber=<%=DocumentNumber%>" target="_blank">click here</a>
			</td>
		</tr>		
		<%
	End if	
   ' Response.Write "CourierLink="&CourierLink&"<BR>"	
	If whatever="whatever" and Trim(CourierLink)>"" then
		%>
		<tr>
			<td class="MainPageText" colspan="4" valign="top">
				<span class="MainPageTextBold">Quick Documentation:</span>&nbsp;&nbsp;<a href="<%=CourierLink%>" target="_blank">click here</a>
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
			'Response.Write "917 orderdetails l_cSQL62 = " & l_cSQL62 & "<br>"
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
						<td class="MainPageText" colspan="2" valign="top">
							<span class="MainPageTextBold">POD EDI:</span>&nbsp;&nbsp;<%=EDI_DateTime%>
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
 						<td class="MainPageText" colspan="2" valign="top">
							<span class="MainPageTextBold">Proof of Pickup:</span>&nbsp;&nbsp;(<%=LotDocumentNumber%>)&nbsp;&nbsp;<%=PUDisplaySignature%>
						</td> 
                        <%                  
                    else
					%>	<!--			
						<td class="MainPageText" colspan="2" valign="top">
							&nbsp;&nbsp;
						</td>
                        -->
					<%
					end if
                    end if
					%>
						<td class="MainPageText" colspan="2" valign="top">
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
									
									<span class="MainPageTextBold">POD:</span>&nbsp;&nbsp;<a href="http://document.logisticorp.us:8080/LibertyIMS::/User=WebUser;pwd=Internet42;sys=LogistiCorp/Cmd%3DGetRawDocument%3BFolder%3D%2321%3BDoc%3D<%=HexULID%>%3Bformat%3DLIC/" target="_blank"><%=DisplaySignature%></a>&nbsp;
									<%
									else
									ULID=""
									If isdate(PODDateTime) then
										%>
										<span class="MainPageTextBold">POD:</span>&nbsp;&nbsp;<a href="../KWEPODS/<%=trim(LotDocumentNumber)%>.pdf" target="_blank"><%=DisplaySignature%></a>&nbsp;
										<%
										Else
										%>									
									
									
									<span class="MainPageTextBold">POD:</span>&nbsp;&nbsp;<%=DisplaySignature%>&nbsp;
									<%
									End if
								End if
								RSEVENTS22.close
							Set RSEVENTS22 = Nothing						
						
						%>
						<!--
							<span class="MainPageTextBold">POD:</span>&nbsp;&nbsp;(<%=LotDocumentNumber%>)&nbsp;&nbsp;<a href="http://192.168.104.231:8080/LibertyIMS::/User=WebUser;pwd=Internet42;sys=LogistiCorp/Cmd%3DGetRawDocument%3BFolder%3D%2321%3BDoc%3D<%=HexULID%>%3Bformat%3DLIC/" target="_blank">xxx<%=DisplaySignature%>xxx</a>&nbsp;
						-->	
							
							<%else
							'If isdate(PODDateTime) then
								%>
								<!--
								<span class="MainPageTextBold">POD:</span>&nbsp;&nbsp;(<%=LotDocumentNumber%>)&nbsp;&nbsp;<a href="../KWEPODS/<%=trim(LotDocumentNumber)%>.pdf" target="_blank"><%=DisplaySignature%></a>&nbsp;
								-->
								<%
								'Else
                                If Trim(DisplaySignature)="" then DisplaySignature="n/a" end if
								%>
								
								<span class="MainPageTextBold">POD:</span>
                                <%If DisplaySignature<>"n/a" then %>
                                <!--
                                     &nbsp;&nbsp;(<%=LotDocumentNumber%>)
                                -->
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
			<td class="MainPageText" colspan="4" valign="top">
				<span class="MainPageTextBold">Exception<%IF NumberOfExceptions>1 then response.Write "s" end if%>:</span>&nbsp;&nbsp;<%=Display_ExceptionList%>
			</td>
		</tr>
	<%end if%>	
<%
    if BookTime<>"1/1/1900" and not isnull(BookTime) and not isnull(SAPOrderTime) then
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
	
	
	
	

	

			
	<%
    'Response.Write "1120 orderdetails BookTime="&BookTime&"<BR>"
	'Response.Write "1120 orderdetails DisplayBookTime="&DisplayBookTime&"<BR>"
	'Response.Write "1121 orderdetails DropTime="&DropTime&"<BR>"
	if DropTime<>"1/1/1900" then
		ElapsedTime=((cDate(DropTime)-cDate(ShipDate))*24)*60
		Hours = Int (ElapsedTime / 60)	
		Minutes = ElapsedTime - (Hours * 60)
		If Minutes>0 then 
			Minutes=cInt(Minutes)
			else
			Minutes=0
		End if		
		%>	
		<tr>
			<td class="MainPageTextBold" colspan="2" align="left">
				Total Delivery Time
			</td>
			<td class="MainPageTextBold" colspan="2" align="left">
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
					<tr><td colspan="4" class="MainPageTextBold"><%=DeliveryWord%>: 
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
		<td width="25%" Class="MainPageTextBold">Job Number</td>
		<td width="25%" Class="MainPageText"><%=OrderID%></td>
		<td width="25%" Class="MainPageTextBold">Priority</td>
		<td width="25%" Class="MainPageText"><%=DisplayPriority%></td>		
	</tr>
	<tr>
		<td width="25%" Class="MainPageTextBold">Job Number</td>
		<td width="25%" Class="MainPageText"><%=OrderID%></td>
		<td width="25%" Class="MainPageTextBold">Submitted By</td>
		<td width="25%" Class="MainPageText"><%=SubmittedBy%></td>		
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
