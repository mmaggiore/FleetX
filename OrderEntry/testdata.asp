
<%
Whatever=Request.querystring("Whatever")

        
        DATABASE="DATABASE=FleetX;DSN=SQLConnect;UID=sa;Password=cadre;"
		INTRANET="DATABASE=lcintranet;DSN=Intranet;UID=sa;Password=cadre;"

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
        

     	Set oConn = Server.CreateObject("ADODB.Connection")
    	oConn.ConnectionTimeout = 100
    	oConn.Provider = "MSDASQL"
    	oConn.Open DATABASE
    		l_cSQL = "Select Top 10 * FROM Mark_Hdr_file_V3"
    		SET oRs = oConn.Execute(l_cSql)
    				Do While not oRs.EOF
                    c1=oRs("Invoice Number")
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
                    c27=oRs("Invoice Reference Number")
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
Response.write c1&"<BR>"
    		oRs.movenext
    		LOOP
    	Set oConn=Nothing           

     	Set oConn = Server.CreateObject("ADODB.Connection")
    	oConn.ConnectionTimeout = 100
    	oConn.Provider = "MSDASQL"
    	oConn.Open DATABASE
    		l_cSQL = "Select * FROM Mark_Invoice_Line_V3 where Invoice_no='"& c1 &"'"
    		SET oRs = oConn.Execute(l_cSql)
    				Do While not oRs.EOF
                    cc1=oRs("Invoice_no")
                    cc2=oRs("Company_ID")
                    cc3=oRs("qty_requested")
                    cc4=oRs("qty_shipped")
                    cc5=oRs("unit_of_measure")
                    cc6=oRs("item_id")
                    cc7=oRs("item_desc")
                    cc8=oRs("Unit Price")
                    cc9=oRs("extended_price")
                    cc10=oRs("gl_revenue_account_no")
                    cc11=oRs("cogs_amount")
                    cc12=oRs("job_id")
                    cc13=oRs("gl_cogsAccount")
                    cc14=oRs("pricing_quantity")

                    line_number=line_number+1


                    cc15=line_number
                    cc16=oRs("oe_line_number")
                    cc17=oRs("OriginID")
                    cc18=oRs("OriginName")
                    cc19=oRs("DestinationID")
                    cc20=oRs("DestinationName")
                    cc21=oRs("CustomerPartNumber")
                    cc22=oRs("UserDefined 1")

                    'Response.write "************<BR>"
                    'Response.write "cc1="&cc1&"<BR>"
                    'Response.write "cc2="&cc2&"<BR>"
                    'Response.write "cc3="&cc3&"<BR>"
                    'Response.write "cc4="&cc4&"<BR>"
                    'Response.write "cc5="&cc5&"<BR>"
                    'Response.write "cc6="&cc6&"<BR>"
                    'Response.write "cc7="&cc7&"<BR>"
                    'Response.write "cc8="&cc8&"<BR>"
                    'Response.write "cc9="&cc9&"<BR>"
                    'Response.write "cc10="&cc10&"<BR>"
                    'Response.write "cc11="&cc11&"<BR>"
                    'Response.write "cc12="&cc12&"<BR>"
                    'Response.write "cc13="&cc13&"<BR>"
                    'Response.write "cc14="&cc14&"<BR>"
                    'Response.write "cc15="&cc15&"<BR>"
                    'Response.write "cc16="&cc16&"<BR>"
                    'Response.write "cc17="&cc17&"<BR>"
                    'Response.write "cc18="&cc18&"<BR>"
                    'Response.write "cc19="&cc19&"<BR>"
                    'Response.write "cc20="&cc20&"<BR>"
                   ' Response.write "cc21="&cc21&"<BR>"
                    'Response.write "cc22="&cc22&"<BR>"
                    'Response.write "line_number="&line_number&"<BR>"
    		oRs.movenext
    		LOOP
    	Set oConn=Nothing           
         %>




