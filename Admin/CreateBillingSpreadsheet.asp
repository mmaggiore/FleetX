
<%
BilledDate=Request.Form("BilledDate")
xyz=Request.Form("xyz")
If BilledDate>"" and xyz="72" then
        
        DATABASE="DATABASE=FleetX;DSN=SQLConnect;UID=sa;Password=cadre;"
		INTRANET="DATABASE=lcintranet;DSN=Intranet;UID=sa;Password=cadre;"
        
        Response.ContentType = "application/vnd.ms-excel"
        Response.AddHeader "Content-Disposition", "attachment;filename=file01.xls"
        %>
        <?xml version="1.0"?>
        <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
        xmlns:o="urn:schemas-microsoft-com:office:office"
        xmlns:x="urn:schemas-microsoft-com:office:excel"
        xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
        xmlns:html="http://www.w3.org/TR/REC-html40">
        <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">
        <DownloadComponents/>
        <LocationOfComponents HRef="file:///\\"/>
        </OfficeDocumentSettings>
        <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
        <WindowHeight>12525</WindowHeight>
        <WindowWidth>15195</WindowWidth>
        <WindowTopX>480</WindowTopX>
        <WindowTopY>120</WindowTopY>
        <ActiveSheet>2</ActiveSheet>
        <ProtectStructure>False</ProtectStructure>
        <ProtectWindows>False</ProtectWindows>
        </ExcelWorkbook>
        <Styles>
        <Style ss:ID="Default" ss:Name="Normal">
        <Alignment ss:Vertical="Bottom"/>
        <Borders/>
        <Font/>
        <Interior/>
        <NumberFormat/>
        <Protection/>
        </Style>
        </Styles>
        <Worksheet ss:Name="Hdr File">
         <!--Table ss:ExpandedColumnCount="67" ss:ExpandedRowCount="2" x:FullColumns="1" x:FullRows="1">-->
        <Table>
        <%
                    r1="Invoice Number"
                    r2="Invoice Date"
                    r3="Customer ID"
                    r4="Bill To Name"
                    r5="Bill To Contact"
                    r6="Bill To Address 1"
                    r7="Bill To Address 2"
                    r8="Bill To City"
                    r9="Bill To State"
                    r10="Bill To Postal Code"
                    r11="Ship to Name"
                    r12="Ship to Contact"
                    r13="Ship to Address 1"
                    r14="Ship to Address 2"
                    r15="Ship to City"
                    r16="Ship to State"
                    r17="Ship to Postal Code"
                    r18="Salesrep ID"
                    r19="AR Account"
                    r20="Period"
                    r21="Year for Period"
                    r22="Ship to ID"
                    r23="Ship Date"
                    r24="Total Amount"
                    r25="Company ID"
                    r26="Bill to Country"
                    r26b="Ship to country"
                    r27="Invoice Reference Number"
                    r28="Invoice Adjustment Type"
                    r29="Invoice Description"
                    r30="Period Fully Paid"
                    r31="Year Fully Paid"
                    r32="Approved"
                    r33="NET DUE DATE"
                    r34="Terms Due Date"
                    r35="Terms ID"
                    r36="Branch ID"
                    r37="Carrier Name"
                    r38="FOB"
                    r39="Terms Description"
                    r40="Purchase Order Number"
                    r41="Salesrep Name"
                    r42="EDI Reference Number"
                    r43="Invoice Type"
                    r44="Amount Paid"
                    r45="Terms Taken"
                    r46="Allowed Amount"
                    r47="Paidin Full"
                    r48="Printed"
                    r49="Print Date"
                    r50="Customer ID String"
                    r51="Corporate Address ID"
                    r52="Shipping Cost"
                    r53="Memo Amount"
                    r54="Bad Debt Amount"
                    r55="Order Number"
                    r56="Order Date"
                    r57="Paid by Check Number"
                    r58="Invoice Class"
                    r59="Disputed"
                    r60="Subject to Finance Charge"
                    r61="Currency ID"
                    r62="Invoice Header User 1"
                    r63="Invoice Header User 2"
                    r64="Invoice Header User 3"
                    r65="Invoice Header User 4"
                    r66="Invoice Header User 5"
                    r67="Invoice Header User 6" 
        
         %>
         <!--
        <Row>
         <Cell><Data ss:Type="String"><%=r1 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r2 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r3 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r4 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r5 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r6 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r7 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r8 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r9 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r10 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r11 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r12 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r13 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r14 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r15 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r16 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r17 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r18 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r19 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r20 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r21 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r22 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r23 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r24 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r25 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r26 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r26b %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r27 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r28 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r29 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r30 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r31 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r32 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r33 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r34 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r35 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r36 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r37 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r38 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r39 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r40 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r41 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r42 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r43 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r44 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r45 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r46 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r47 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r48 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r49 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r50 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r51 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r52 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r53 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r54 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r55 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r56 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r57 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r58 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r59 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r60 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r61 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r62 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r63 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r64 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r65 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r66 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=r67 %></Data></Cell>       
        </Row>
        -->
        <%
     	Set oConn = Server.CreateObject("ADODB.Connection")
    	oConn.ConnectionTimeout = 100
    	oConn.Provider = "MSDASQL"
    	oConn.Open DATABASE
    		l_cSQL = "Select * FROM Mark_Hdr_file_V3 where (fh_billed_date='"&BilledDate&"')"
    		

            SET oRs = oConn.Execute(l_cSql)
    				Do While not oRs.EOF
                    cA=trim(oRs("Invoice Number"))
                    cB=oRs("Invoice Date")
                    cC=oRs("Customer ID")
                    cD=oRs("Bill To Name")
                    cE=oRs("Bill To Contact")
                    cF=oRs("Bill To Address 1")
                    cG=oRs("Bill To Address 2")
                    cH=oRs("Bill To City")
                    cI=oRs("Bill To State")
                    cJ=oRs("Bill To Postal Code")
                    cK=oRs("Ship to Name")
                    cL=oRs("Ship to Contact")
                    cM=oRs("Ship to Address 1")
                    cN=oRs("Ship to Address 2")
                    cO=oRs("Ship to City")
                    cP=oRs("Ship to State")
                    cQ=oRs("Ship to Postal Code")
                    cR=oRs("Salesrep ID")
                    cS=oRs("AR Account")
                    cT=oRs("Period")
                    cU=oRs("Year for Period")
                    cV=oRs("Ship to ID")
                    cW=oRs("Ship Date")
                    cX=oRs("Total Amount")
                    cY=oRs("Company ID")
                    cZ=oRs("Bill to Country")
                    cAA=oRs("Ship to country")
                    cAB=trim(oRs("Invoice Reference Number"))
                    cAC=oRs("Invoice Adjustment Type")
                    cAD=oRs("Invoice Description")
                    cAE=oRs("Period Fully Paid")
                    cAF=oRs("Year Fully Paid")
                    cAG=oRs("Approved")
                    cAH=oRs("NET DUE DATE")
                    cAI=oRs("Terms Due Date")
                    cAJ=oRs("Terms ID")
                    cAK=oRs("Branch ID")
                    cAL=oRs("Carrier Name")
                    cAM=oRs("FOB")
                    cAN=oRs("Terms Description")
                    cAO=oRs("Purchase Order Number")
                    cAP=oRs("Salesrep Name")
                    cAQ=oRs("EDI Reference Number")
                    cAR=oRs("Invoice Type")
                    cAS=oRs("Amount Paid")
                    cAT=oRs("Terms Taken")
                    cAU=oRs("Allowed Amount")
                    cAV=oRs("Paidin Full")
                    cAW=oRs("Printed")
                    cAX=oRs("Print Date")
                    cAY=oRs("Customer ID String")
                    cAZ=oRs("Corporate Address ID")
                    cBA=oRs("Shipping Cost")
                    cBB=oRs("Memo Amount")
                    cBC=oRs("Bad Debt Amount")
                    cBD=oRs("Order Number")
                    cBE=oRs("Order Date")
                    cBF=oRs("Paid by Check Number")
                    cBG=oRs("Invoice Class")
                    cBH=oRs("Disputed")
                    cBI=oRs("Subject to Finance Charge")
                    cBJ=oRs("Currency ID")
					cBK=FormatDateTime(CDate(Now()),2)
					cBL=FormatDateTime(CDate(Now()),2)
					cBM="AutoBill"					
                    cBN=oRs("Invoice Header User 1")
                    cBO=oRs("Invoice Header User 2")
                    ''''''c64=oRs("Invoice Header User 3")
					cBP=oRs("Invoice Header User 3")  
                    cBQ=oRs("Invoice Header User 4")
                    cBR=oRs("Invoice Header User 5")
                    cBS=oRs("Invoice Header User 6")

    
         %>
        <Row>
        <Cell><Data ss:Type="String"><%=cA %>M</Data></Cell>
        <Cell><Data ss:Type="String"><%=cB %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cC %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cD %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cE %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cF %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cG %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cH %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cI %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cJ %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cK %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cL %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cM %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cN %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cO %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cP %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cQ %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cR %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cS %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cT %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cU %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cV %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cW %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cX %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cY %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cZ %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cAA %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cAB %>M</Data></Cell>
        <Cell><Data ss:Type="String"><%=cAC %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cAD %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cAE %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cAF %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cAG %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cAH %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cAI %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cAJ %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cAK %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cAL %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cAM %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cAN %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cAO %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cAP %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cAQ %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cAR %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cAS %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cAT %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cAU %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cAV %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cAW %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cAX %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cAY %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cAZ %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cBA %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cBB %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cBC %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cBD %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cBE %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cBF %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cBG %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cBH %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cBI %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cBJ %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cBK %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cBL %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cBM %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cBN %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cBO %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cBS %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cBQ %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cBR %></Data></Cell>
        <Cell><Data ss:Type="String"></Data></Cell>
        </Row>
        <%
    		oRs.movenext
    		LOOP
    	Set oConn=Nothing           
         %>
         
        </Table>
        <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
        <ProtectObjects>False</ProtectObjects>
        <ProtectScenarios>False</ProtectScenarios>
        </WorksheetOptions>
        </Worksheet>
        <Worksheet ss:Name="Invoice Line">
         <Table>
        <!--Table ss:ExpandedColumnCount="22" ss:ExpandedRowCount="2" x:FullColumns="1" x:FullRows="1"-->
 
 
 <%
                    rr1="Invoice Number"
                    rr2="Company ID"
                    rr3="Quantity Ordered"
                    rr4="Quantity Shipped"
                    rr5="Unit of Measure"
                    rr6="Item ID"
                    rr7="Item Description"
                    rr8="Unit Price"
                    rr9="Extended Price"
                    rr10="GL Revenue Account"
                    rr11="COGS Amount"
                    rr12="Job ID"
                    rr13="GL COGS Account"
                    rr14="Pricing Quantity"
                    rr15="Line Number"
                    rr16="OE Line Number"
                    rr17="Date Created"
                    rr18="Last Date Modified"
                    rr19="Last Maintained By"
                    rr20="Origin ID"
                    rr21="Origin Name"
                    rr22="Destination ID"
                    rr23="Destination Name"
                    rr24="Customer Part Number"
                    rr25="Invoice Line User Defined 1"
                    rr26="User Defined 1"
        
         %>
         <!--
        <Row>
         <Cell><Data ss:Type="String"><%=rr1 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=rr2 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=rr3 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=rr4 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=rr5 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=rr6 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=rr7 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=rr8 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=rr9 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=rr10 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=rr11 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=rr12 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=rr13 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=rr14 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=rr15 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=rr16 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=rr17 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=rr18 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=rr19 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=rr20 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=rr21 %></Data></Cell>
        '''removed'''<Cell><Data ss:Type="String"><%=rr22 %></Data></Cell>
 
        </Row>
        -->
 
 
 
 
 
 
        
<%
     	Set oConn = Server.CreateObject("ADODB.Connection")
    	oConn.ConnectionTimeout = 100
    	oConn.Provider = "MSDASQL"
    	oConn.Open DATABASE
    		'''''l_cSQL = "Select * FROM Mark_Invoice_Line_V3 where Invoice_no='"& c1 &"'"
            l_cSQL = "Select * FROM Mark_Invoice_Line_V3 where (fh_billed_date='"&BilledDate&"')"
    		SET oRs = oConn.Execute(l_cSql)
    				Do While not oRs.EOF
                    cc1=trim(oRs("Invoice_no"))
                    cc2=oRs("Company_ID")
                    cc3=oRs("qty_requested")
                    cc4=oRs("qty_shipped")
                    cc5=oRs("unit_of_measure")
                    cc6=UCASE(oRs("item_id"))
                    cc7=oRs("item_desc")
                    cc8=oRs("Unit Price")
                    cc9=oRs("extended_price")
					If trim(cc6)="ADDITIONAL SKIDS" then
						''''unit price
						cc8=12
						''''qty requested and shipped
						cc3=(cc9/12)
						cc4=cc3
					End if
                    cc10=oRs("gl_revenue_account_no")
                    cc10=replace(cc10, "-", "")
                    cc11=oRs("cogs_amount")
                    cc12=oRs("job_id")
                    cc13=oRs("gl_cogsAccount")
                    cc14=oRs("pricing_quantity")
                    If tempcc1=cc1 then
                        line_number=line_number+1
                        else
                        line_number=1
                    End if
                    tempdate=oRS("fh_billed_date")
                    displaydate=month(tempdate)&"/"&day(tempdate)&"/"&year(tempdate)
                    cc15=line_number
                    cc16=oRs("oe_line_number")
                    cc17=displaydate
                    cc18=displaydate
                    cc19="ahilborn"
                    cc20=oRs("OriginID")
                    cc21=oRs("OriginName")
                    cc22=oRs("DestinationID")
                    cc23=oRs("DestinationName")
                    cc24=" "
                    cc25=oRs("CustomerPartNumber")
                    cc26=oRs("UserDefined 1")
                    tempcc1=cc1

        %>
        <Row>
        <Cell><Data ss:Type="String"><%=cc1 %>M</Data></Cell>
        <Cell><Data ss:Type="String"><%=cc2 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cc3 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cc4 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cc5 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cc6 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cc7 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cc8 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cc9 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cc10 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cc9 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cc12 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cc13 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cc14 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cc15 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cc16 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cc17 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cc18 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cc19 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cc20 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cc21 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cc22 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cc23 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cc25 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=cc26 %></Data></Cell>
        </Row>
        <%
    		oRs.movenext
    		LOOP
    	Set oConn=Nothing           
         %>
        
        </Table>
        <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
        <ProtectObjects>False</ProtectObjects>
        <ProtectScenarios>False</ProtectScenarios>
        </WorksheetOptions>
        </Worksheet>
        </Workbook>
        <%
        else

%>
<html>
<head>
</head>
<body>
    Hello!
    BilledDate=<%=BilledDate %>XXX
</body>
</html>
<%End if %>


