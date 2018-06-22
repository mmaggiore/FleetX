<html>
<head>

<!-- #include file="../fleetexpress.inc" -->
<link rel="stylesheet" type="text/css" href="../css/Style.css"/>
<%
Session("suid")="91"


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
    PageTitle="FLEETX UNBILLED JOBS"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'==========================================================================
' CRYSTAL ENTERPRISE REPORT APPLICATION SERVER (CE EMBEDDED) 10
' Purpose:  Demonstrate how to pass parameters to subreports
'==========================================================================
SQLFromAddress=Request.Form("SQLFromAddress")
fh_bt_id=Request.Form("fh_bt_id")

If SQLFromAddress<>"ANY" then
	FromAddressSymbol="="
	else
	FromAddressSymbol="<>"
End if	
SQLToAddress=Request.Form("SQLToAddress")

If SQLToAddress<>"ANY" then
	ToAddressSymbol="="
	else
	ToAddressSymbol="<>"
End if	
    SQLDriverID=Request.Form("DriverID")
    SQLDriverName=SQLDriverID
If SQLDriverID<>"0" then
	DriverSymbol="="
	else
	'SQLDriverID=0
	DriverSymbol="<>"
End if		
   
    
' DC making this a "Multi-Call" form with a hidden object
IF Request.Form("hdnHaveParms") = "YES" Then
	' We're coming here the second time and, therefore, have parameter values

	' This is the Radio Button Selection
	l_cSel = Request.Form("STSEL")
	
	' This is the DATE the user entered
	l_cDate = Request.Form("txtDate")                                                                  
	l_cEndDate = Request.Form("strEnd")

	'If trim(SQLDriverName)="" then SQLDriverName="3" End if   	
	'Thiss line creates a string variable called reportname that we will use to pass
	'the Crystal Report filename (.rpt file) to the OpenReport method.
	reportname = "FleetX_Unbilled.rpt"



'============================================================================
' CREATE THE REPORT CLIENT DOCUMENT OBJECT AND OPEN THE REPORT
'============================================================================

' Use the Object Factory object to create other RAS objects (useful for versioning changes)
Set objFactory = CreateObject("CrystalReports.ObjectFactory")
	
' This "While/Wend" loop is used to determine the physical path (eg: C:\) to the 
' Crystal Report .rpt by translating the URL virtual path (eg: http://Domain/Dir)   
Dim path, iLen

path = Request.ServerVariables("PATH_TRANSLATED") 

While (Right(path, 1) <> "\" And Len(path) <> 0)                      
	iLen = Len(path) - 1                                                  
	path = Left(path, iLen)                                               
Wend       
                                                                                                                                  
' Create a new ReportClientDocument object
Set Session("oClientDoc") = objFactory.CreateObject("CrystalClientDoc.ReportClientDocument")

' Specify the RAS Server (computer name or IP address) to use (If SDK and RAS Service are running on seperate machines)
' 192.168.111.2
Session("oClientDoc").ReportAppServer = Session("ReportAppServerVariable")

' Open the report object to initialize the ReportClientDocument
Session("oClientDoc").Open reportName  


'==================================================================
' WORKING WITH DISCRETE PARAMETERS
'==================================================================

'This subroutine is called each time you have to pass a value to a Crystal Report parameter
Public Sub SetParamValues (ReportName, ParamName, ParamValue)

	'Creates a values collection to contain a list of parameters (discrete parameters in this case)
	Set Values = CreateObject("crystalReports.Values")
	
	'All the parameters in this example will be discrete parameters.  Therefore we create 
	'this object to hold the value of the parameter
	Set DiscreteValue = CreateObject ("CrystalReports.ParameterFieldDiscreteValue")

	'Sets the value of the discrete parameter
	DiscreteValue.Value = ParamValue
	
	'Adds the discrete parameter to the values collection
	Values.Add DiscreteValue 

	'Sets the current values for the parameter
	Session("oClientDoc").DataDefController.ParameterFieldController.SetCurrentValues ReportName, ParamName, Values

End Sub



	    
	    
	' SEcond param
SetParamValues "", "fh_bt_id", CStr(fh_bt_id)	
'SetParamValues "", "Billto_ID",  CStr(Session("sUid"))
'SetParamValues "", "ToDate", CDate(l_cEndDate)
'SetParamValues "", "FromAddress", CStr(SQLFromAddress)
'SetParamValues "", "FromAddressSymbol", CStr(FromAddressSymbol)
'SetParamValues "", "ToAddress", CStr(SQLToAddress)
'SetParamValues "", "ToAddressSymbol", CStr(ToAddressSymbol)
'SetParamValues "", "DriverID", CStr(SQLDriverName)
'SetParamValues "", "DriverSymbol", CStr(DriverSymbol)


'============================================================================
' CHOOSING THE REPORT VIEWER
'============================================================================
'
' There are four Report Viewers:
' 1.  Crystal Reports Interactive Viewer (CrystalReportsInteractiveViewer.asp)
' 2.  Crystal Reports Viewer (CrystalReportsViewer.asp)
' 3.  Crystal Reports Parts Viewer (CrystalReportsPartsViewer.asp)
' 4.  Legacy ActiveX Viewer (ActiveXViewer.asp)
'
' Note that to use this these viewers you must have the appropriate .asp file in the 
' same virtual directory as the main ASP page. Choose from one of the four viewers below,
' simply uncomment the one you want to use:

'  *** Crystal Reports Interactive Viewer ***
Response.Redirect "CrystalReportsInteractiveViewer.asp"

'  *** Crystal Reports Viewer ***
'Response.Redirect "CrystalReportsViewer.asp"

'    IMPORTANT NOTE - To use the report parts viewer successfully you are required to
'    choose and name three objects in the report to Node1, Node2 and Node3.
'    You can access an objects name by using the Format Editor dialog box.
'    For more information on the Format Editor Dialog Box and setting objects
'    names, please refer to the Help Contents (Help Menu->Crystal Reports Help)
'    or by pressing F1

'  *** Crystal Reports Parts Viewer ***
'Response.Redirect "CrystalReportsPartsViewer.asp"

' 	 IMPORTANT NOTE - The ActiveXviewer does NOT have the ability to prompt for parameters,
' 	 selection formulas, or login information.  If your report requires this information
'    to run, look at the the appropriate code samples to pass this information before
' 	 using this viewer.

'  *** Legacy ActiveX Viewer ***
'Response.Redirect "ActiveXViewer.asp"

'=============================================================================
ELSE
	' This is the first time and we need to prompt for parameters
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


%>
<title><%=PageTitle %></title>
</head>

<body onload="document.FindUser.requestorName.focus();">
	<table align="center" border="0" bordercolor="black" cellpadding="0" cellspacing="0" ID="Table4" height="100%" width="100%">
        <tr><td align="left"><img src="../images/pixel.gif" height="10" width="1" /></td></tr>
        <tr>
            <td align="left"><img src="../images/FleetX_Small.jpg" height="50" width="168" /></td>
            <td align="right" valign="bottom"><!-- #include file="../topnavbar.asp" --></td>
        </tr>
        <tr><td align="left" colspan="2"><img src="../images/pixel.gif" /></td></tr>	
    <tr><td colspan="2">

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



    <tr><td>
    <!--main page stuff goes here!-->
    
    
    
    
    
    
		<!-- 08/17/2004 Span-Renuka - Added this code for displaying calender control accordingly-->
	<% currpath=Request.ServerVariables("PATH_INFO")%>
	<% 
	'Response.Write "currpath="&currpath&"<BR>"
	IF currpath="/Reporting/MetricDSBToDropZoneNew.asp" THEN %>
		<SCRIPT LANGUAGE="JavaScript" SRC="../v9web/scripts/date-picker.js"></SCRIPT>
	<%ELSE 
	'Response.Write "GOT HERE!"
	%>
		<SCRIPT LANGUAGE="JavaScript" SRC="../scripts/date-picker.js"></SCRIPT>
	<%END IF %>
		<!-- 08/17/2004 Span-Renuka commented this code -->	
	</head>
	<%
	IF D_CBODYELEMENTS <> "" THEN
		' They have specified some custom body elements
		Response.Write(" " & D_CBODYELEMENTS & " ")
	END IF
	%>
	<CENTER>
	<%
	' Response.Write("<IMG SRC='" & Session("sLogo") & "'>")

	' Optionally show the Bold Heading of the purpose of this page
	Response.Write("<Div class='FleetXLargerBoldText'>FLEETX Unbilled Jobs<BR><BR></Div>")

	IF SHOW_BTNAME THEN
		Response.Write("<H2>"& Session("txt_cm_desc"))
		IF SHOW_BTID THEN
			Response.Write(" (Customer ID:" & Session("suid") & ")")
		END IF
		Response.Write("</H2><BR>")
	END IF
	
	'  Want to have:
	'	(*) Yesterday's Jobs
	'	(*) Today's Jobs
	'	(*) Specific Date:  [mm/dd/yy]

	%>

	<form name="GetJobParms" action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="POST" ID="Form1">
	<table border="4" cellpadding="2" cellspacing="0" ID="Table1">
	<tr>
	<td>
	<table border="0" bordercolor="red" width="350" ID="Table2">
	<!--tr><td><img src="../../images/transpixel.gif" width="1" height="10"></td></tr-->
    <TR><td class="FleetXBoldText" align="center">Please choose your report selection criteria</td></TR>
    
	<form name="GetJobParms" action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="POST" ID="Form2">

<TR> 
    <td align="center"> 
          <table border=0 cellpadding="2" cellspacing="0" ID="Table3">
            <tr> 
              <td class="subheader" colspan="4"><img src="../images/pixel.gif" height="2"></td>
            </tr>
            <!--
            <tr> 
              <td class="FleetXBoldText" width="97">From:</td>
              <td class="FleetXBoldText" nowrap="nowrap"> 
                <input type='text' size='12' name='txtDate' value='<%=Date()-1%>' maxlength="12" ID="Text1">
                &nbsp;<a href="javascript:show_calendar('GetJobParms.txtDate');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;"><img src="../../images/cal.gif" width="16" height="15" border="0" name="calendar" alt="Calendar" title="Calendar" align="ABSMIDDLE"></a>
                </td>
                
                
                
              <td class="FleetXBoldText" width="97">To:</td>
              <td class="FleetXBoldText" nowrap="nowrap"> 
                <input type='text' size='12' name='strEnd' value='<%=Date()-1%>' maxlength="12" ID="Text2">
                &nbsp;<a href="javascript:show_calendar('GetJobParms.strEnd');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;"><img src="../../images/cal.gif" width="16" height="15" border="0" name="calendar" alt="Calendar" title="Calendar" align="ABSMIDDLE"></a>
                </td>              
                
                
                
            </tr>
            <tr><td>&nbsp;</td></tr>
            <tr>
				<td nowrap class="FleetXBoldText">From:</td>
				<td colspan="3">
					<select name="SQLFromAddress">
						<option value="ANY">All Locations</option>
       					<option value="DSTK">DSTK-Dallas Support Building</option>
	                    <option value="ESTK">ESTK-East Building Stockroom</option>
                    </select>
				</td>
			</tr>
            <tr><td>&nbsp;</td></tr>
            <tr>
				<td class="FleetXBoldText" nowrap>To:</td>
				<td colspan="3">
					<select name="SQLToAddress">
						<option value="ANY">All Locations</option>
					<%    
			            
  					'Dim oRs
					Set oConn = Server.CreateObject("ADODB.Connection")
					oConn.ConnectionTimeout = 200
					oConn.Provider = "MSDASQL"
					oConn.Open DATABASE

					'On Error Resume Next                                                  
					Err.Clear
					'Response.Write "l_cSQL="&l_cSQL&"<BR>"
					l_cSQL="SELECT * FROM PreExistingCompanies WHERE (IsStockroom='y') order by CompanyName"
					Set oRs = oConn.Execute(l_cSQL)
					
					If Err.Number <> 0 Then                                               
					Response.Write "Error Executing the query.  Error:" & Err.Description
					Else
						IF NOT oRs.EOF THEN
							oRs.MoveFirst

							DO WHILE NOT oRs.EOF
							LocationID2=oRs("st_id")
							LocationDescription2=oRs("CompanyName")				

							%>          
							<option value="<%=trim(LocationDescription2)%>"><%=trim(LocationDescription2)%></option>
							<%				
							
							
							oRs.MoveNext
							LOOP
						End if
					End If
					%> 
					</select>	
				</td>
			</tr>
			<tr><td>&nbsp;</td></tr>
			<tr>
				<td class="FleetXBoldText">
					Driver:	
				</td>
				<td class="generalcontent" colspan="3">
					<select name="DriverID" ID="Select4">
					<option value="0">Any Driver</option>
					<%    
			            
  					'Dim oRs
					Set oConn = Server.CreateObject("ADODB.Connection")
					oConn.ConnectionTimeout = 200
					oConn.Provider = "MSDASQL"
					oConn.Open INTRANET

					'On Error Resume Next                                                  
					Err.Clear
					'Response.Write "l_cSQL="&l_cSQL&"<BR>"
					l_cSQL="SELECT * FROM Intranet_Users WHERE (DriverVehicle>'') AND (status='c') order by lastname, firstname"
					Set oRs = oConn.Execute(l_cSQL)
					
					If Err.Number <> 0 Then                                               
					Response.Write "Error Executing the query.  Error:" & Err.Description
					Else
						IF NOT oRs.EOF THEN
							oRs.MoveFirst

							DO WHILE NOT oRs.EOF
							UserID=oRs("UserID")
							FirstName=oRs("FirstName")
							LastName=oRs("LastName")	
							%>          
							<option value="<%=trim(UserID)%>"><%=trim(LastName)%>, <%=Trim(FirstName)%></option>
							<%				
							
							
							oRs.MoveNext
							LOOP
						End if
					End If
					%> 
					</select>           
				</td>
			</tr>
			<%
			'Response.Write "l_cSQL="&l_cSQL&"<BR>"
			%>

            -->
            <tr>
                <td>
                    <input type="radio" name="fh_Bt_id" value="92" <%If fh_bt_id="" or fh_bt_id="92" then response.write " checked" end if %> /> Courier<br />
                    <input type="radio" name="fh_Bt_id" value="93" <%If fh_bt_id="93" then response.write " checked" end if %> /> Freight<br />
                    <!--
                    <input type="radio" name="fh_Bt_id" value="91" <%If fh_bt_id="91" then response.write " checked" end if %>/> Stockroom<br />
                    -->
                </td>
            </tr>
            <tr> 
              <td class="subheader" colspan="4"><img src="../images/pixel.gif" height="2"></td>
            </tr>
            </table>
            </td></tr>
            
            </table>

      </td>
    </TR>
	</table>
	<BR>
	<input type="hidden" name="hdnHaveParms" value="YES" ID="Hidden1">
	<input type="submit" name="btnSubmit" value="View Report" ID="gobutton">
	<input type="reset" name="btnReset"  value="Reset" ID="gobutton">
	</form>
	</CENTER>
	<%
END IF
	%>    
    
    
    
    
    
    
    
    
    
    
    
    </td></tr>



 
	<tr Height="50">
		<td>&nbsp;</td>
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


	
