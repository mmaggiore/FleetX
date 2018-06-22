
<%
BilledDate=Request.Form("BilledDate")
xyz=Request.Form("xyz")
'If BilledDate>"" and xyz="72" then
If BilledDate="" then
        
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
    		'l_cSQL = "Select * FROM Mark_Hdr_file_V3 where (fh_billed_date='"&BilledDate&"')"
    		l_cSQL = "Select * FROM Mark_Hdr_file_V3_Temp where ([ship date]>'3/28/2018' and [ship date]<'4/1/2018' and fh_bt_id='93')"
            SET oRs = oConn.Execute(l_cSql)
    				Do While not oRs.EOF
                    c1=trim(oRs("Invoice Number"))
                    c2=oRs("Invoice Date")
                    c3=oRs("Customer ID")
                    c4=oRs("Bill To Name")
                    c5=oRs("Bill To Contact")
                    c6=oRs("Bill To Address 1")
                    c7=oRs("Bill To Address 2")
                    c8=oRs("Bill To City")
                    c9=oRs("Bill To State")
                    c10=oRs("Bill To Postal Code")
                    c11=oRs("Ship to Name")
                    c12=oRs("Ship to Contact")
                    c13=oRs("Ship to Address 1")
                    c14=oRs("Ship to Address 2")
                    c15=oRs("Ship to City")
                    c16=oRs("Ship to State")
                    c17=oRs("Ship to Postal Code")
                    c18=oRs("Salesrep ID")
                    c19=oRs("AR Account")
                    c20=oRs("Period")
                    c21=oRs("Year for Period")
                    c22=oRs("Ship to ID")
                    c23=oRs("Ship Date")
                    c24=oRs("Total Amount")
                    c25=oRs("Company ID")
                    c26=oRs("Bill to Country")
                    c26b=oRs("Ship to country")
                    c27=trim(oRs("Invoice Reference Number"))
                    c28=oRs("Invoice Adjustment Type")
                    c29=oRs("Invoice Description")
                    c30=oRs("Period Fully Paid")
                    c31=oRs("Year Fully Paid")
                    c32=oRs("Approved")
                    c33=oRs("NET DUE DATE")
                    c34=oRs("Terms Due Date")
                    c35=oRs("Terms ID")
                    c36=oRs("Branch ID")
                    c37=oRs("Carrier Name")
                    c38=oRs("FOB")
                    c39=oRs("Terms Description")
                    c40=oRs("Purchase Order Number")
                    c41=oRs("Salesrep Name")
                    c42=oRs("EDI Reference Number")
                    c43=oRs("Invoice Type")
                    c44=oRs("Amount Paid")
                    c45=oRs("Terms Taken")
                    c46=oRs("Allowed Amount")
                    c47=oRs("Paidin Full")
                    c48=oRs("Printed")
                    c49=oRs("Print Date")
                    c50=oRs("Customer ID String")
                    c51=oRs("Corporate Address ID")
                    c52=oRs("Shipping Cost")
                    c53=oRs("Memo Amount")
                    c54=oRs("Bad Debt Amount")
                    c55=oRs("Order Number")
                    c56=oRs("Order Date")
                    c57=oRs("Paid by Check Number")
                    c58=oRs("Invoice Class")
                    c59=oRs("Disputed")
                    c60=oRs("Subject to Finance Charge")
                    c61=oRs("Currency ID")
                    c62=oRs("Invoice Header User 1")
                    c63=oRs("Invoice Header User 2")
                    c64=oRs("Invoice Header User 3")
                    c65=oRs("Invoice Header User 4")
                    c66=oRs("Invoice Header User 5")
                    c67=oRs("Invoice Header User 6")  


    
         %>
        <Row>
        <Cell><Data ss:Type="String"><%=c1 %>M</Data></Cell>
        <Cell><Data ss:Type="String"><%=c2 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c3 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c4 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c5 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c6 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c7 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c8 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c9 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c10 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c11 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c12 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c13 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c14 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c15 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c16 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c17 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c18 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c19 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c20 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c21 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c22 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c23 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c24 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c25 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c26 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c26b %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c27 %>M</Data></Cell>
        <Cell><Data ss:Type="String"><%=c28 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c29 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c30 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c31 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c32 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c33 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c34 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c35 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c36 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c37 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c38 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c39 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c40 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c41 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c42 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c43 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c44 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c45 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c46 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c47 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c48 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c49 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c50 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c51 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c52 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c53 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c54 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c55 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c56 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c57 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c58 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c59 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c60 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c61 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c62 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c63 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c64 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c65 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c66 %></Data></Cell>
        <Cell><Data ss:Type="String"><%=c67 %></Data></Cell>
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
            'l_cSQL = "Select * FROM Mark_Invoice_Line_V3 where (fh_billed_date='"&BilledDate&"')"
            l_cSQL = "Select * FROM Mark_Invoice_Line_V3_Temp where (fl_t_atd>'3/28/2018' and [ship date]<'4/1/2018'  and fh_bt_id='93')"
            
    		SET oRs = oConn.Execute(l_cSql)
    				Do While not oRs.EOF
                    cc1=trim(oRs("Invoice_no"))
                    cc2=oRs("Company_ID")
                    cc3=oRs("qty_requested")
                    cc4=oRs("qty_shipped")
                    cc5=oRs("unit_of_measure")
                    cc6=oRs("item_id")
                    cc7=oRs("item_desc")
                    cc8=oRs("Unit Price")
                    cc9=oRs("extended_price")
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


