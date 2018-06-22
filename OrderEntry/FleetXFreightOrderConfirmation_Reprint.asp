<html>
<head>

   <%
    FXCourieruserid=valid8(request.QueryString("VarA"))
    Supervisor=valid8(Request.QueryString("VarB"))
    tempUserID=session("tempUserID")
    session("tempUserID")=TempUserID
    'Response.write "TempUserID="&TempUserID&"<BR>"
    JID=valid8(Request.querystring("JID"))
    fleetexpresscourier="y"
    ColorSelect=valid8(Request.form("ColorSelect"))
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
            HeaderBorderColor="#41924B"  
            BorderColor="#41924B"
            LinkClass="FleetExpressGreen"
        Case else 
            HeaderBorderColor="black"  
            BorderColor="black"
            LinkClass="FleetExpressBlack"
    End Select
    HighlightedField="RequestorName"
    CurrentDateTime=Now()
   '''''''''''HARDCODED STUFF
   sBT_ID="88"
   Session("sBT_ID")=sBT_ID
   ''''''''''''''''''''''''''
   userid=valid8(Request.form("UserID"))
   LogInVerified=valid8(Request.form("LogInVerified"))
   'Response.write "UserID="&UserID&"<BR>"
    MarkTemp=valid8(Request.Form("MarkTemp"))
    
    If  trim(JID)>"" then
    %>



<!-- #include file="../fleetexpress.inc" -->
<link rel="stylesheet" type="text/css" href="../css/Style.css" />
<link rel="stylesheet" href="../css/hide.css" type="text/css" media="print" />
<%
    ColorSelect=valid8(Request.form("ColorSelect"))
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
    PageTitle="FLEETX ORDER CONFIRMATION"

%>

     <%


    timesthrough=valid8(Request.form("timesthrough"))
    
    TableWidth="460"
    Internal=valid8(Request.QueryString("Internal"))
   
   		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.ConnectionTimeout = 100
		oConn.Provider = "MSDASQL"
		oConn.Open DATABASE
			l_cSQL = "Select * FROM fcfgthd WHERE  fh_id='"& jid &"'"
            'Response.Write "l_cSql="&l_cSql&"<BR>"
			SET oRs = oConn.Execute(l_cSql)
					if NOT oRs.EOF then
                        fh_id=oRs("fh_id")
                        PickUpDateTime=oRs("fh_ready")
                        RequestorName=oRs("fh_co_id")
                        RequestorPhoneNumber=oRs("fh_co_phone")
                        RequestoremailAddress=oRs("fh_co_email")
                        costcenterNumber=oRs("fh_co_costcenter")
                        PoNumber=oRs("fh_custpo")
                        If trim(costcenterNumber)>"" then
                            CostOrPO="Cost Center"
                        End if
                        If trim(PoNumber)>"" then
                            CostOrPO="PO Number"
                        End if
                        DeliveryType=oRs("fh_co_id")
                        'BasicCharge=oRs("BasicCharge")
                    End if								
		Set oConn=Nothing
        Set oRS=Nothing
   		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.ConnectionTimeout = 100
		oConn.Provider = "MSDASQL"
		oConn.Open DATABASE
			l_cSQL = "Select * FROM fclegs WHERE  fl_fh_id='"& jid &"'"
            'Response.Write "l_cSql="&l_cSql&"<BR>"
			SET oRs = oConn.Execute(l_cSql)
					if NOT oRs.EOF then
                        OriginationCompany=oRs("fl_sf_name")
                        OriginationContactName=oRs("fl_sf_clname")
                        OriginationPhoneNumber=oRs("fl_sf_phone")
                        OriginationEmail=oRs("fl_sf_email")
                        OriginationBuilding=oRs("fl_sf_Building")
                        OriginationAddress=oRs("fl_sf_addr1")
                        OriginationSuite=oRs("fl_sf_addr2")
                        OriginationCity=oRs("fl_sf_city")
                        OriginationState=oRs("fl_sf_state")
                        OriginationCountry=oRs("fl_sf_country")
                        OriginationZipCode=oRs("fl_sf_zip")
                        OriginationAliasCode=oRs("fl_sf_alias")
                        DestinationCompany=oRs("fl_st_name")
                        DestinationContactName=oRs("fl_st_clname")
                        DestinationPhoneNumber=oRs("fl_st_phone")
                        DestinationEmail=oRs("fl_st_email")
                        DestinationBuilding=oRs("fl_st_Building")
                        DestinationAddress=oRs("fl_st_addr1")
                        DestinationSuite=oRs("fl_st_addr2")
                        DestinationCity=oRs("fl_st_city")
                        DestinationState=oRs("fl_st_state")
                        DestinationCountry=oRs("fl_st_country")
                        DestinationZipCode=oRs("fl_st_zip")
                        DestinationAliasCode=oRs("fl_st_zip")
                        DestinationAliasCode=oRs("fl_st_alias")
                        Comments=oRs("fl_sf_comment")
                        
                    End if								
		Set oConn=Nothing
        Set oRS=Nothing   
   		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.ConnectionTimeout = 100
		oConn.Provider = "MSDASQL"
		oConn.Open DATABASE
			l_cSQL = "Select * FROM fcrefs WHERE  rf_fh_id='"& jid &"'"
            'Response.Write "l_cSql="&l_cSql&"<BR>"
			SET oRs = oConn.Execute(l_cSql)
					if NOT oRs.EOF then
                        rf_box=oRs("rf_box")
                        PartNumber=oRs("rf_ref")
                        DisplayMaterialDescription=oRs("MaterialDescription")
                        DisplayPartNumber=oRs("PartNumber")
                        Pieces=oRs("NumberOfPieces")
                        NumberOfPallets=oRs("NumberOfPallets")
                        DimWeight=oRs("Weight")
                        DimLength=oRs("DimLength")
                        DimWidth=oRs("DimWidth")
                        DimHeight=oRs("DimHeight")
                        MeasurementType=oRs("MeasurementType")
                    End if								
		Set oConn=Nothing
        Set oRS=Nothing  



     %>



<title>FleetX - <%=PageTitle %></title>
</head>

<body onload="document.FindUser.requestorName.focus();document.OrderForm1.<%=HighlightedField%>.focus();">

	<table align="center" border="0" bordercolor="black" cellpadding="0" cellspacing="0" ID="Table1" height="100%" width="100%">
        
        <tr><td align="left"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>

        <tr>
            <td align="left"><div class="hide"><img src="../images/FleetX_Small.jpg" height="50" width="168" /></div></td>
            <td align="right" valign="bottom"><div class="hide"><!-- #include file="../topnavbar.asp" --></div></td>
        </tr>
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" /></td></tr>	

    <tr><td colspan="2">
<form action="NewUser.asp" method="post" name="FindUser">
<table border="0" bordercolor="green" Cellspacing="0" Cellpadding="0" align="left" width="100%">
 <tr><td align="left" colspan="2" bgcolor="<%=HeaderBorderColor%>"><div class="hide"><img src="../images/pixel.gif" height="5" width="1" /></div></td></tr>
 <div class="hide">
        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetXRedSection" colspan="2"><div class="hide"><%=PageTitle%></div></td></tr>
 </div>
        <tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressHeaderWhite" colspan="2"><div class="hide"><img src="../images/pixel.gif" height="5" width="1" /></div></td></tr>
        <!--tr><td align="center" bgcolor="<%=HeaderBorderColor%>" class="FleetExpressBodyWhite" colspan="2">In order to reach the Fleet Express Order page, please correctly type in the green verification code in the supplied text box and click "Submit."</td></tr-->
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" height="5" width="1" /></td></tr>

<tr><td>
<table  border="0" bordercolor="blue" align="center" class="MainPageText" width="100%">
	<tr height="40">
		<td width="650">&nbsp;</td>
	</tr>

    <tr><td align=center width="100%"><!-- main page stuff goes here! -->
    
    
<table border="0" cellpadding="0" cellspacing="0" align="center" width="770" bgcolor="white">
<tr>
    <td><img src="../images/pixel.gif" width="30" height="1" /></td>
    <td>
<table cellpadding="0" cellspacing="0" border="0" width="100%">
<tr>
    <td class="MainPageTextCenterLargeBlack" colspan="3">FleetX Waybill</td>
    <td rowspan="2"><img src="../images/FleetX_Small.jpg" /></td>
</tr>
<tr>
    <td class="MainPageTextBigCenterRed" colspan="3">(Please print out and attach to your shipment)</td>
</tr>
</table>

    <table border="0" cellpadding="0" cellspacing="0" align="center">
        <tr><td align="left"><img src="../images/pixel.gif" height="30" width="1" /></td></tr>
        <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
        <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
        <tr><td colspan="4" bgcolor="black"><img src="../images/pixel.gif" height="1" width="1" /></td></tr>
        <tr><td><img src="../images/pixel.gif" height="2" width="1" /></td></tr>
        <tr>
            <%somenumber=11
            If Supervisor<>"y" then
            SomeNumber=SomeNumber-1
            End if
            'Response.write "somenumber="&SomeNumber&"<BR>"
             %>
            <td valign="top" rowspan="<%=SomeNumber %>" class="OrderHeader">REQUESTOR<img src="../images/pixel.gif" height="1" width="15" /></td>
            
           <td class="FleetExpressTextBlackBold" align="left">Requestor Name</td>
            <td width="10"><img src="../images/pixel.gif" /></td>
            <td align="left"><%=RequestorName%></td>
        </tr>
        <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
        <tr>
            <td class="FleetExpressTextBlackBold" align="left">Phone Number</td>
            <td width="10"><img src="../images/pixel.gif" /></td>
            <td align="left"><%=RequestorPhoneNumber%></td>
        </tr>
        <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
        <tr>
            <td class="FleetExpressTextBlackBold" align="left">Email Address</td>
            <td width="10"><img src="../images/pixel.gif" /></td>
            <td align="left"><%=RequestorEmailAddress%></td>
        </tr>
        <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
        <tr>
            <td class="FleetExpressTextBlackBold" align="left">
                    <%=CostOrPO%>
            </td>
            <td width="10"><img src="../images/pixel.gif" /></td>
            <td align="left"><%=PONumber%><%=CostCenterNumber %></td>
        </tr>
        <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
        <tr>
            <td class="FleetExpressTextBlackBold" align="left" valign="top" nowrap="nowrap">Special Instructions&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
            <td width="10"><img src="../images/pixel.gif" /></td>
            <td align="left"><%=Comments%></td>
        </tr>

        <tr><td>&nbsp;</td></tr>
        <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
        <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
        <tr><td colspan="4" bgcolor="black"><img src="../images/pixel.gif" height="1" width="1" /></td></tr>
        <tr><td><img src="../images/pixel.gif" height="2" width="1" /></td></tr>
                 <tr>
                 <td valign="top" rowspan="7"  class="OrderHeader" nowrap="nowrap">COMMODITY<img src="../images/pixel.gif" height="1" width="15" /></td>
                        <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">Material Description</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td nowrap class="FleetExpressTextBlack">
                            <%=DisplayMaterialDescription%>
                        </td>
                    </tr> 
                    <!--
                    <tr>                      
                        <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">Control/Part #</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td nowrap class="FleetExpressTextBlack">
                            <%=DisplayPartNumber%>
                        </td>
                    </tr>
                    -->
                    <tr>                      
                        <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">Number of Pieces</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td nowrap class="FleetExpressTextBlack">
                            <%=pieces%>
                            &nbsp;&nbsp;
                            <%=rf_box%>
                        </td>
                    </tr>
                    <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                    <%'Response.write "DimWeight="&DimWeight&"<BR>" %>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">Weight</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left" class="FleetExpressTextBlack">
                           <%=DimWeight %>
                        Pounds</td>
                    </tr>
                    <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                    <%If trim(DimLength)>"" or trim(DimWidth)>"" or trim(DimHeight)>"" then %>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">Dimensions</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td class="FleetExpressTextBlack" align="left" nowrap>
                            L:&nbsp;
                            <%=DimLength %>
                            W:&nbsp;
                            <%=DimWidth %>                           
                           H:&nbsp;
                           <%=DimHeight %>
                            &nbsp;Inches
                        </td>
                    </tr>
                    <%End if %>
                    <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
            </td>
        </tr>
        <tr><td>&nbsp;</td></tr>

                <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                <tr><td colspan="4" bgcolor="black"><img src="../images/pixel.gif" height="1" width="1" /></td></tr>
                <tr><td><img src="../images/pixel.gif" height="2" width="1" /></td></tr>
                <tr>
                    <td valign="top" rowspan="2"  class="OrderHeader" nowrap="nowrap">ORIGINATION<img src="../images/pixel.gif" height="1" width="15" /></td>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap valign="top" nowrap="nowrap">Company Name</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left"><%=OriginationCompany%></td>
                </tr>
                <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                <tr>
                    <td valign="top" rowspan="15"  class="OrderHeader">
                             <%
                            '''''''''''''''''''''''''''''''''''''''''			
                            'Code 39 barcodes require an asterisk as the start and stop characters
			                            BarCodeText=OriginationAliasCode
			                            'BarCodeText="1234/567-89"
			                            'Response.Write "BarCodeText="&BarCodeText&"<BR>"
			                            If BarCodeText>"" then
				                            Response.write "<br><IMG SRC=""../images/barcodes/asterisk.gif"" WIDTH=""18"" HEIGHT=""45"">"

				                            For x = 1 to Len(Trim(BarCodeText))
					                            DisplayBarCode=mid(BarCodeText,x,1)
					                            If DisplayBarCode="/" then
						                            Response.write "<IMG SRC=""../images/barcodes/!slash.gif"" WIDTH=""18"" HEIGHT=""45"">"
						                            else
						                            Response.Write "<IMG SRC=""../images/barcodes/" & DisplayBarCode & _
								                            ".gif"" WIDTH=""18"" HEIGHT=""45"">"
					                            End if
				                            Next

				                            'Code 39 barcodes require an asterisk as the start and stop characters
				                            Response.write "<IMG SRC=""../images/barcodes/asterisk.gif"" WIDTH=""18"" HEIGHT=""45"">"
			                            End if
			
			
                            '''''''''''''''''''''''''''''''''''''''''''	
                    %>                    
                    
                    
                    </td>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">Building</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left"><%=OriginationBuilding%></td>
                </tr>
                <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">Address</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left"><%=OriginationAddress%></td>
                </tr>
                <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">Suite/Cube</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left"><%=OriginationSuite%></td>
                </tr>
                <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">City/State/Zip</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td class="FleetExpressTextBlack">
                       
                        <%=OriginationCity%>
                        &nbsp;TX&nbsp;&nbsp;
                        <%=OriginationZipCode%>

                    </td>
                </tr>
                <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>

                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">Contact Name</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left"><%=OriginationContactName%></td>
                </tr>
                <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">Phone Number</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left"><%=OriginationPhoneNumber%></td>
                </tr> 
                <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">Email Address</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left"><%=OriginationEmail%></td>
                </tr>
                 <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">Ready Date/Time</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left"><%=PickUpDateTime%>
                    </td>
                </tr>
                <tr><td>&nbsp;</td></tr>
                <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                <tr><td colspan="4" bgcolor="black"><img src="../images/pixel.gif" height="1" width="1" /></td></tr>
                <tr><td><img src="../images/pixel.gif" height="2" width="1" /></td></tr>
                <tr>
                    <td valign="top" rowspan="2" class="OrderHeader" nowrap="nowrap">DESTINATION<img src="../images/pixel.gif" height="1" width="15" />
                      
                    </td>
                    <td class="FleetExpressTextBlackBold" align="left" valign="top" nowrap="nowrap">Company Name</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left"><%=DESTINATIONCompany%></td>
                    </tr>
                    <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                    <tr>
                        <td valign="top" rowspan="15"  class="OrderHeader">
                             <%
                            '''''''''''''''''''''''''''''''''''''''''			
                            'Code 39 barcodes require an asterisk as the start and stop characters
			                            BarCodeText=DESTINATIONAliasCode
			                            'BarCodeText="1234/567-89"
			                            'Response.Write "BarCodeText="&BarCodeText&"<BR>"
			                            If BarCodeText>"" then
				                            Response.write "<br><IMG SRC=""../images/barcodes/asterisk.gif"" WIDTH=""18"" HEIGHT=""45"">"

				                            For x = 1 to Len(Trim(BarCodeText))
					                            DisplayBarCode=mid(BarCodeText,x,1)
					                            If DisplayBarCode="/" then
						                            Response.write "<IMG SRC=""../images/barcodes/!slash.gif"" WIDTH=""18"" HEIGHT=""45"">"
						                            else
						                            Response.Write "<IMG SRC=""../images/barcodes/" & DisplayBarCode & _
								                            ".gif"" WIDTH=""18"" HEIGHT=""45"">"
					                            End if
				                            Next

				                            'Code 39 barcodes require an asterisk as the start and stop characters
				                            Response.write "<IMG SRC=""../images/barcodes/asterisk.gif"" WIDTH=""18"" HEIGHT=""45"">"
			                            End if
			
			
                            '''''''''''''''''''''''''''''''''''''''''''	
                    %>                       
                        </td>
                        <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">Building</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left"><%=DESTINATIONBuilding%></td>
                    </tr>
                    <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">Address</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left"><%=DESTINATIONAddress%></td>
                    </tr>
                    <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">Suite/Cube</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left"><%=DESTINATIONSuite%></td>
                    </tr>
                    <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">City/State/Zip</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td class="FleetExpressTextBlack">
                            <%=DESTINATIONCity%>&nbsp;TX&nbsp;&nbsp;
                        <%=DESTINATIONZipCode%>

                        </td>
                    </tr>

                    <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">Contact Name</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left"><%=DESTINATIONContactName%></td>
                    </tr>
                    <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">Phone Number</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left"><%=DESTINATIONPhoneNumber%></td>
                    </tr>
                    <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr> 
                    <tr>
                        <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">Email Address</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td align="left"><%=DESTINATIONEmail%></td>
                    </tr>
                    <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr> 
                <tr>
                    <td class="FleetExpressTextBlackBold" align="left" nowrap="nowrap">Delivery Date/Time</td>
                    <td width="10"><img src="../images/pixel.gif" /></td>
                    <td align="left"><%=DeliveryDateTime%>
                    </td>
                </tr>
                <tr><td>&nbsp;</td></tr>
        <!--
        <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
        <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
        <tr><td colspan="4" bgcolor="black"><img src="../images/pixel.gif" height="1" width="1" /></td></tr>
        <tr><td><img src="../images/pixel.gif" height="2" width="1" /></td></tr>
                 <tr>
                 <td valign="top" rowspan="2"  class="OrderHeader">CHARGES<img src="../images/pixel.gif" height="1" width="15" /></td>
                        <td class="FleetExpressTextBlackBold" align="left">Estimated Costs</td>
                        <td width="10"><img src="../images/pixel.gif" /></td>
                        <td nowrap class="FleetExpressTextBlackBold">
                            $<%=BasicCharge%>
                        </td>
                    </tr>
        -->
        <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
        <tr><td><img src="../images/pixel.gif" height="3" width="1" /></td></tr>
        <tr><td colspan="4" bgcolor="black"><img src="../images/pixel.gif" height="1" width="1" /></td></tr>
        <tr><td><img src="../images/pixel.gif" height="2" width="1" /></td></tr>
 <div class="hide">                
                 <tr>
                 <td valign="top" rowspan="2"  class="OrderHeader" nowrap="nowrap">SCANNING CODE<img src="../images/pixel.gif" height="1" width="15" /></td>
                        <td class="FleetExpressTextBlackBold" align="center" colspan="3">
                            <%
                            '''''''''''''''''''''''''''''''''''''''''			
                            'Code 39 barcodes require an asterisk as the start and stop characters
			                            BarCodeText=PartNumber
			                            'BarCodeText="1234/567-89"
			                            'Response.Write "BarCodeText="&BarCodeText&"<BR>"
			                            If BarCodeText>"" then
				                            Response.write BarCodeText&"<br><br><IMG SRC=""../images/barcodes/asterisk.gif"" WIDTH=""18"" HEIGHT=""45"">"

				                            For x = 1 to Len(Trim(BarCodeText))
					                            DisplayBarCode=mid(BarCodeText,x,1)
					                            If DisplayBarCode="/" then
						                            Response.write "<IMG SRC=""../images/barcodes/!slash.gif"" WIDTH=""18"" HEIGHT=""45"">"
						                            else
						                            Response.Write "<IMG SRC=""../images/barcodes/" & DisplayBarCode & _
								                            ".gif"" WIDTH=""18"" HEIGHT=""45"">"
					                            End if
				                            Next

				                            'Code 39 barcodes require an asterisk as the start and stop characters
				                            Response.write "<IMG SRC=""../images/barcodes/asterisk.gif"" WIDTH=""18"" HEIGHT=""45"">"
			                            End if
			
			
                            '''''''''''''''''''''''''''''''''''''''''''	
                    %>

                        </td>
                    </tr>
                    <tr><td><img src="../images/pixel.gif" height="50" width="1" /></td></tr>
                    <tr><td colspan="5">
                    <table class="hide">
                        <tr>
                            <td valign=top>
                                <form ID="Form1"><input id="gobutton" type="button" value="Print Waybill" onclick="window.print();return false;" ID="Button1" NAME="Button1"/></form> 
                            </td>
                            <td>
                            <%
                            'Response.write "FXCourieruserid="&FXCourieruserid&"<BR>" 
                            'Response.write "Supervisor="&Supervisor&"<BR>" 
                            %>
                                <form method="post" action="FreightOrder.asp">
                                    <input type="hidden" name="userid" value="<%=TempUserID %>" />
                                    <input type="hidden" name="FXCourieruserid" value="<%=FXCourieruserid %>" />
                                    <input type="hidden" name="Supervisor" value="<%=Supervisor %>" />
                                    <input type="hidden" name="fleetexpresscourier" value="y" />
                                     <input type="hidden" name="LogInVerified" value="y" />
                                    <input id="gobutton" type="submit" value="Place New Order" />
                                </form>
                                </td><td>
                                <form method="post" action="../home.asp">
                                    <input type="hidden" name="userid" value="<%=TempUserID %>" />
                                    <input type="hidden" name="FXCourieruserid" value="<%=FXCourieruserid %>" />
                                    <input type="hidden" name="Supervisor" value="<%=Supervisor %>" />
                                    <input type="hidden" name="LogInVerified" value="y" />
                                    <input type="hidden" name="fleetexpresscourier" value="y" />
                                    <input id="gobutton" type="submit" value="Return Home" />
                                </form>
                            </td>
                        </tr>
                    </table>
                    </td></tr>
                    <!--
                    tempuserid=<%=tempuserid %>
                    -->
                    <tr><td><img src="../images/pixel.gif" height="45" width="1" /></td></tr>

         
        <input type="hidden" value="1" name="Timesthrough" />


        
    </table>
</td>
<td><img src="../images/pixel.gif" width="30" height="1" /></td>
</tr>
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

<%
else
Response.redirect("../home.asp")
'end if
end if
'Response.write "PageStatus="&PageStatus&"<BR>"
%>

</body>
</html>

